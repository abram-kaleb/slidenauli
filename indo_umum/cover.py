# cover.py
import streamlit as st
from docx import Document
import re


def extract_cover(doc):
    paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]

    tata_ibadah = "TATA IBADAH"
    nama_minggu = ""
    topik = ""
    tanggal = ""

    re_tanggal = re.compile(
        r"((?:SENIN|SELASA|RABU|KAMIS|JUMAT|SABTU|MINGGU)\s*,?\s*)?"
        r"(\d{1,2}\s+(?:JANUARI|FEBRUARI|PEBRUARI|MARET|APRIL|MEI|JUNI|JULI|AGUSTUS|SEPTEMBER|OKTOBER|NOVEMBER|DESEMBER)\s+\d{4})",
        re.IGNORECASE
    )

    keywords_minggu = [
        "ADVENT", "NATAL", "SETELAH NATAL", "PARPUNGUAN BODARI", "TAHUN BARU",
        "SETELAH TAHUN BARU", "EPHIPANIAS", "EPIFANI", "SEPTUAGESIMA", "SEXAGESIMA",
        "ESTOMIHI", "INVOCAVIT", "REMINISCERE", "OKULI", "LETARE", "LAETARE",
        "JUDIKA", "PALMARUM", "JUMAT AGUNG", "PASKAH", "PASKA", "QUASIMODOGENITI",
        "MISERIKORDIAS DOMINI", "JUBILATE", "KANTATE", "ROGATE", "KENAIKAN",
        "EXAUDI", "PENTAKOSTA", "TRINITATIS", "SETELAH TRINITATIS", "UJUNG TAON PARHURIAON"
    ]

    for i, text in enumerate(paragraphs[:25]):
        text_upper = text.upper()

        tgl_match = re_tanggal.search(text)
        if tgl_match:
            tanggal = tgl_match.group()

        if (any(k in text_upper for k in keywords_minggu) or "SETELAH" in text_upper) and not tgl_match:
            if len(text.split()) < 12:
                clean_m = re.sub(r"^[PL]\s*[:\-]\s*", "",
                                 text, flags=re.IGNORECASE).strip()
                if "TATA IBADAH" in clean_m.upper():
                    m_frag = re.search(r"MINGGU.*", clean_m, re.IGNORECASE)
                    if m_frag:
                        clean_m = m_frag.group()
                nama_minggu = clean_m

        if "TOPIK" in text_upper:
            if ":" in text:
                topik = text.split(":", 1)[1].strip()
            elif i + 1 < len(paragraphs):
                topik = paragraphs[i+1]

        if not topik and i < 15:
            quote_match = re.search(r"[“\"].*?[”\"]", text)
            if quote_match:
                topik = quote_match.group()

    if topik:
        topik = topik.strip("“ ” \"").upper()

    if nama_minggu:
        nama_minggu = nama_minggu.upper()
        if nama_minggu.startswith("TATA IBADAH"):
            nama_minggu = nama_minggu.replace("TATA IBADAH", "").strip()

    return {
        "tata_ibadah": tata_ibadah,
        "minggu": nama_minggu,
        "topik": topik,
        "tanggal": tanggal
    }


if __name__ == "__main__":
    st.title("Ekstraksi Data Cover")
    uploaded_file = st.file_uploader("Upload file DOCX", type=["docx"])

    if uploaded_file:
        doc = Document(uploaded_file)
        data = extract_cover(doc)

        st.text(f"Tata Ibadah: {data['tata_ibadah']}")
        st.text(f"Nama Minggu: ({data['minggu']})")
        st.text(f"Topik: ({data['topik']})")
        st.text(f"Tanggal: ({data['tanggal']})")
