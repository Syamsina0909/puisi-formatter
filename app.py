
import io
import hmac
import streamlit as st
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# =====================================
# KONFIGURASI APLIKASI
# =====================================

st.set_page_config(
    page_title="Puisi Formatter Online",
    page_icon="📖",
    layout="wide"
)

# =====================================
# LOGIN KODE AKSES
# =====================================

def check_login():
    if st.session_state.get("authenticated", False):
        return

    st.title("📖 Puisi Formatter Online")
    st.write("Masukkan kode akses untuk menggunakan aplikasi.")

    with st.form("login"):
        password = st.text_input(
            "Kode akses",
            type="password"
        )
        submitted = st.form_submit_button("Masuk")

    if submitted:
        expected = st.secrets.get("access_password", "")

        if expected and hmac.compare_digest(password, str(expected)):
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("Kode akses salah atau belum dikonfigurasi.")

    st.stop()


check_login()

# =====================================
# TAMPILAN UTAMA
# =====================================

st.title("📖 Puisi Formatter Online")
st.caption("Rapikan naskah antologi puisi secara otomatis.")

with st.sidebar:
    st.header("⚙️ Pengaturan Format")

    paper = st.selectbox("Ukuran kertas", ["A5", "A4"])

    if paper == "A5":
        page_width, page_height = 14.8, 21.0
    else:
        page_width, page_height = 21.0, 29.7

    margin = st.number_input(
        "Margin (cm)", 1.0, 4.0, 2.0, 0.25
    )

    font_name = st.selectbox(
        "Jenis huruf",
        ["Times New Roman", "Arial", "Calibri", "Garamond"]
    )

    font_size = st.number_input(
        "Ukuran huruf",
        9, 16, 12
    )

    spacing = st.selectbox(
        "Spasi baris",
        [1.0, 1.15, 1.5, 2.0],
        index=1
    )

    st.divider()

    if st.button("Keluar"):
        st.session_state.authenticated = False
        st.rerun()


# =====================================
# DETEKSI NASKAH
# =====================================

AUTHOR_MARKERS = (
    "karya:", "oleh:", "penulis:", "ditulis oleh:"
)

BIO_MARKERS = (
    "biodata penulis",
    "biografi penulis",
    "profil penulis",
    "tentang penulis"
)


def is_title(text):
    text = text.strip()

    if len(text) < 3 or len(text) > 120:
        return False

    letters = [c for c in text if c.isalpha()]

    if not letters:
        return False

    ratio = sum(c.isupper() for c in letters) / len(letters)

    return ratio >= 0.75


def detect_document(file_bytes):
    doc = Document(io.BytesIO(file_bytes))

    paragraphs = [p.text.strip() for p in doc.paragraphs]

    works = []
    current = None
    pending_author = ""
    bio = []
    in_bio = False

    for text in paragraphs:
        lower = text.lower()

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

        if not text:
            if current:
                current["body"].append("")
            continue

        marker = next(
            (m for m in AUTHOR_MARKERS if lower.startswith(m)),
            None
        )

        if marker:
            author_text = text[len(marker):].strip()

            if current is not None and not current["author"]:
                current["author"] = text
            else:
                pending_author = text

            continue

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


# =====================================
# MEMBUAT DOKUMEN WORD
# =====================================

def set_font(run, name, size, bold=False):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor(0, 0, 0)


def make_output(
    works,
    bio,
    include_bio,
    page_width,
    page_height,
    margin,
    font_name,
    font_size,
    spacing
):
    doc = Document()

    section = doc.sections[0]
    section.page_width = Cm(page_width)
    section.page_height = Cm(page_height)

    section.top_margin = Cm(margin)
    section.bottom_margin = Cm(margin)
    section.left_margin = Cm(margin)
    section.right_margin = Cm(margin)

    for index, work in enumerate(works):

        if index > 0:
            doc.add_page_break()

        # Judul
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(0)

        r = p.add_run(work["title"].upper())
        set_font(r, "DengXian", 14, True)

        # Penulis
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(12)

        author = work["author"] or "Karya: [Nama Penulis]"
        r = p.add_run(author)
        set_font(r, font_name, font_size)

        # Isi puisi
        for line in work["body"]:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = spacing

            r = p.add_run(line)
            set_font(r, font_name, font_size)

    # Biodata
    if include_bio and bio:
        doc.add_page_break()

        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        r = p.add_run("BIODATA PENULIS")
        set_font(r, "DengXian", 14, True)

        for line in bio:
            p = doc.add_paragraph()
            p.paragraph_format.line_spacing = spacing

            r = p.add_run(line)
            set_font(r, font_name, font_size)

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)

    return output


# =====================================
# UNGGAH DAN PROSES NASKAH
# =====================================

st.subheader("1. Unggah Naskah Word")

uploaded = st.file_uploader(
    "Pilih dokumen Word (.docx)",
    type=["docx"]
)

if uploaded:

    if uploaded.size > 15 * 1024 * 1024:
        st.error("Ukuran file maksimal 15 MB.")
        st.stop()

    try:
        works, bio = detect_document(uploaded.getvalue())

        st.success(f"{len(works)} karya terdeteksi.")

        if not works:
            st.warning("Belum ada karya yang terdeteksi.")
            st.stop()

        st.subheader("2. Periksa Hasil Deteksi")

        edited = []

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
                    height=200,
                    key=f"body_{i}"
                )

                edited.append({
                    "title": title,
                    "author": author,
                    "body": body.splitlines()
                })

        if bio:
            st.subheader("3. Biodata Penulis")

            bio_text = st.text_area(
                "Periksa biodata",
                value="\n".join(bio),
                height=180
            )

            bio = bio_text.splitlines()

        include_bio = st.checkbox(
            "Sertakan biodata penulis",
            value=bool(bio)
        )

        st.divider()

        if st.button(
            "✨ Proses dan Buat Dokumen Word",
            type="primary"
        ):
            result = make_output(
                edited,
                bio,
                include_bio,
                page_width,
                page_height,
                margin,
                font_name,
                font_size,
                spacing
            )

            st.success("Dokumen berhasil dibuat!")

            st.download_button(
                "⬇️ Unduh Hasil DOCX",
                data=result,
                file_name="Naskah_Antologi_Tertata.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

    except Exception as error:
        st.error(f"Terjadi kesalahan: {error}")

else:
    st.info("Unggah dokumen Word untuk mulai merapikan naskah.")

st.divider()
st.caption("Puisi Formatter Online | Pemrosesan sementara.")
