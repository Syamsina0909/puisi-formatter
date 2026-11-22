
import io
import re
import streamlit as st
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ==========================================
# PENGATURAN APLIKASI
# ==========================================

st.set_page_config(
    page_title="Puisi Formatter Online",
    page_icon="📖",
    layout="wide"
)

st.title("📖 Puisi Formatter Online")
st.caption("Aplikasi penataan naskah antologi puisi")

# ==========================================
# LOGIN GOOGLE DAN PEMBATASAN AKSES
# ==========================================

def check_login():
    if not st.user.is_logged_in:
        st.title("Selamat Datang!")
        st.write("Silakan masuk menggunakan akun Google.")
        st.button("Masuk dengan Google", on_click=st.login)
        st.stop()

    allowed = [
        email.strip().lower()
        for email in st.secrets.get("allowed_emails", [])
    ]

    current_email = str(st.user.get("email", "")).lower()

    if not current_email or current_email not in allowed:
        st.error("Maaf, akun Google ini belum memiliki izin akses.")
        st.button("Keluar", on_click=st.logout)
        st.stop()


check_login()

# ==========================================
# PENANDA NASKAH
# ==========================================

AUTHOR_PREFIXES = (
    "karya:",
    "oleh:",
    "penulis:",
    "ditulis oleh:"
)

BIO_MARKERS = (
    "biodata penulis",
    "biografi penulis",
    "profil penulis",
    "tentang penulis"
)


def is_title(text):
    """Deteksi awal judul berdasarkan pola teks."""
    text = text.strip()

    if len(text) < 3 or len(text) > 120:
        return False

    letters = [c for c in text if c.isalpha()]

    if not letters:
        return False

    uppercase_ratio = (
        sum(c.isupper() for c in letters) / len(letters)
    )

    return uppercase_ratio >= 0.75


def detect_document(file_bytes):
    """Membaca dan mengelompokkan karya dalam dokumen Word."""

    source = Document(io.BytesIO(file_bytes))

    paragraphs = [
        p.text.strip()
        for p in source.paragraphs
    ]

    works = []
    current = None
    pending_author = ""
    bio = []
    in_bio = False

    for text in paragraphs:

        if not text:
            if current and not in_bio:
                current["body"].append("")
            elif in_bio:
                bio.append("")
            continue

        lower = text.lower()

        # Deteksi bagian biodata
        if any(marker in lower for marker in BIO_MARKERS):
            if current:
                works.append(current)
                current = None

            in_bio = True
            bio.append(text)
            continue

        if in_bio:
            bio.append(text)
            continue

        # Deteksi nama penulis
        author_match = next(
            (
                prefix for prefix in AUTHOR_PREFIXES
                if lower.startswith(prefix)
            ),
            None
        )

        if author_match:
            author_name = text[len(author_match):].strip()

            if current:
                current["author"] = text
            else:
                pending_author = text

            continue

        # Deteksi judul baru
        if is_title(text):

            if current:
                works.append(current)

            current = {
                "title": text,
                "author": pending_author,
                "body": []
            }

            pending_author = ""
            continue

        # Masukkan teks sebagai isi puisi
        if current is None:
            current = {
                "title": "JUDUL BELUM TERDETEKSI",
                "author": pending_author,
                "body": [text]
            }
            pending_author = ""
        else:
            current["body"].append(text)

    if current:
        works.append(current)

    return works, bio


# ==========================================
# FUNGSI PEMFORMATAN WORD
# ==========================================

def set_font(run, font_name, size, bold=False):
    run.font.name = font_name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = False
    run.font.color.rgb = RGBColor(0, 0, 0)


def format_paragraph(
    paragraph,
    alignment,
    before=0,
    after=0,
    spacing=1.15
):
    paragraph.alignment = alignment
    paragraph.paragraph_format.space_before = Pt(before)
    paragraph.paragraph_format.space_after = Pt(after)
    paragraph.paragraph_format.line_spacing = spacing


def make_output(
    works,
    bio,
    include_bio,
    paper_width,
    paper_height,
    margin,
    font_name,
    font_size,
    line_spacing
):

    doc = Document()

    section = doc.sections[0]

    section.page_width = Cm(paper_width)
    section.page_height = Cm(paper_height)

    section.top_margin = Cm(margin)
    section.bottom_margin = Cm(margin)
    section.left_margin = Cm(margin)
    section.right_margin = Cm(margin)

    for index, work in enumerate(works):

        if index > 0:
            doc.add_page_break()

        # JUDUL PUISI
        title = doc.add_paragraph()

        format_paragraph(
            title,
            WD_ALIGN_PARAGRAPH.CENTER,
            after=0,
            spacing=1.0
        )

        run = title.add_run(work["title"].upper())

        set_font(run, "DengXian", 14, True)

        # NAMA PENULIS
        author = doc.add_paragraph()

        format_paragraph(
            author,
            WD_ALIGN_PARAGRAPH.CENTER,
            after=12,
            spacing=1.0
        )

        author_text = (
            work["author"]
            or "Karya: [Nama Penulis]"
        )

        run = author.add_run(author_text)

        set_font(run, font_name, font_size)

        # ISI PUISI
        for line in work["body"]:

            paragraph = doc.add_paragraph()

            format_paragraph(
                paragraph,
                WD_ALIGN_PARAGRAPH.LEFT,
                before=0,
                after=0,
                spacing=line_spacing
            )

            run = paragraph.add_run(line)

            set_font(run, font_name, font_size)

    # BAGIAN BIODATA
    if include_bio and bio:

        doc.add_page_break()

        heading = doc.add_paragraph()

        format_paragraph(
            heading,
            WD_ALIGN_PARAGRAPH.CENTER
        )

        run = heading.add_run("BIODATA PENULIS")

        set_font(run, "DengXian", 14, True)

        for line in bio:

            paragraph = doc.add_paragraph()

            format_paragraph(
                paragraph,
                WD_ALIGN_PARAGRAPH.LEFT,
                spacing=line_spacing
            )

            run = paragraph.add_run(line)

            set_font(run, font_name, font_size)

    output = io.BytesIO()

    doc.save(output)
    output.seek(0)

    return output


# ==========================================
# TAMPILAN UTAMA
# ==========================================

user_email = str(st.user.get("email", ""))

st.write(
    f"Selamat datang, {st.user.get('name', user_email)}!"
)

st.caption(f"Akun: {user_email}")

with st.sidebar:
    st.header("⚙️ Pengaturan")
    
    paper = st.selectbox(
        "Ukuran kertas",
        ["A5", "A4"]
    )

    if paper == "A5":
        paper_width, paper_height = 14.8, 21.0
    else:
        paper_width, paper_height = 21.0, 29.7

    margin = st.number_input(
        "Margin (cm)",
        min_value=1.0,
        max_value=4.0,
        value=2.0,
        step=0.25
    )

    font_name = st.selectbox(
        "Jenis huruf isi puisi",
        ["Times New Roman", "Arial", "Calibri", "Garamond"]
    )

    font_size = st.number_input(
        "Ukuran huruf",
        min_value=9,
        max_value=16,
        value=12
    )

    line_spacing = st.selectbox(
        "Jarak antarbaris",
        [1.0, 1.15, 1.5, 2.0],
        index=1
    )

    st.divider()

    if st.button("Keluar dari aplikasi"):
        st.logout()


st.subheader("1. Unggah Naskah Word")

uploaded = st.file_uploader(
    "Pilih dokumen Word (.docx)",
    type=["docx"]
)

if uploaded:

    file_bytes = uploaded.getvalue()

    if len(file_bytes) > 15 * 1024 * 1024:
        st.error("Ukuran file maksimal 15 MB.")

    else:

        try:

            works, bio = detect_document(file_bytes)

            st.success(
                f"Deteksi selesai: {len(works)} karya ditemukan."
            )

            st.info(
                "Silakan periksa dan koreksi hasil deteksi "
                "sebelum membuat dokumen."
            )

            edited = []

            st.subheader("2. Periksa Hasil Deteksi")

            for i, work in enumerate(works):

                with st.expander(
                    f"Karya {i + 1}: {work['title']}",
                    expanded=(i == 0)
                ):

                    title = st.text_input(
                        "Judul puisi",
                        value=work["title"],
                        key=f"title_{i}"
                    )

                    author = st.text_input(
                        "Nama penulis",
                        value=work["author"],
                        key=f"author_{i}"
                    )

                    body = st.text_area(
                        "Isi puisi",
                        value="\n".join(work["body"]),
                        height=220,
                        key=f"body_{i}"
                    )

                    edited.append({
                        "title": title,
                        "author": author,
                        "body": body.splitlines()
                    })

            if bio:

                st.subheader("3. Biodata Penulis")

                edited_bio = st.text_area(
                    "Periksa biodata",
                    value="\n".join(bio),
                    height=180
                )

                bio = edited_bio.splitlines()

            else:

                st.warning(
                    "Bagian biodata belum terdeteksi."
                )

            include_bio = st.checkbox(
                "Sertakan biodata penulis",
                value=bool(bio)
            )

            st.divider()

            st.subheader("4. Buat Dokumen Word")

            if st.button(
                "✨ Proses dan Buat Dokumen Word",
                type="primary",
                disabled=not edited
            ):

                result = make_output(
                    edited,
                    bio,
                    include_bio,
                    paper_width,
                    paper_height,
                    margin,
                    font_name,
                    font_size,
                    line_spacing
                )

                st.success("Dokumen Word berhasil dibuat!")

                st.download_button(
                    "⬇️ Unduh Hasil DOCX",
                    data=result,
                    file_name="Naskah_Antologi_Tertata.docx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "wordprocessingml.document"
                    )
                )

        except Exception as error:

            st.error(
                f"Dokumen tidak dapat diproses: {error}"
            )

else:

    st.info(
        "Silakan unggah naskah Word untuk memulai."
    )


st.divider()

st.caption(
    "Puisi Formatter Online | Pemrosesan sementara. "
    "Tidak ada penyimpanan naskah permanen yang disengaja."
)
