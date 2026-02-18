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

    pattern_tanggal = r"(\d{1,2}\s+(?:JANUARI|FEBRUARI|PEBRUARI|MARET|APRIL|MEI|JUNI|JULI|AGUSTUS|SEPTEMBER|OKTOBER|NOVEMBER|DESEMBER)\s+\d{4})"

    if paragraphs:
        first_line = paragraphs[0].upper()
        if "TATA TERTIB" in first_line or "KEBAKTIAN" in first_line:
            topik = paragraphs[0]

    for i, text in enumerate(paragraphs[:10]):
        text_upper = text.upper()

        tgl_match = re.search(pattern_tanggal, text, re.IGNORECASE)
        if tgl_match:
            tanggal = tgl_match.group()

            if "MINGGU" in text_upper:
                parts = re.split(pattern_tanggal, text, flags=re.IGNORECASE)
                m_part = parts[0].strip()
                if m_part:
                    nama_minggu = m_part

        if not nama_minggu and "MINGGU" in text_upper and len(text.split()) < 8:
            nama_minggu = text

    if topik:
        topik = topik.replace("TATA TERTIB ", "").strip().upper()

    return {
        "tata_ibadah": tata_ibadah,
        "minggu": nama_minggu.upper(),
        "topik": topik,
        "tanggal": tanggal.upper()
    }


if __name__ == "__main__":
    st.title("Ekstraksi Data Cover")
    uploaded_file = st.file_uploader("Upload file DOCX", type=["docx"])

    if uploaded_file:
        doc = Document(uploaded_file)
        data = extract_cover(doc)

        st.subheader("Hasil Ekstraksi")
        st.text(f"Tata Ibadah: {data['tata_ibadah']}")
        st.text(f"Nama Minggu: {data['minggu']}")
        st.text(f"Topik: {data['topik']}")
        st.text(f"Tanggal: {data['tanggal']}")
