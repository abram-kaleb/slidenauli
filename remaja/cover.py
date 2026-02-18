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
        "ADVENT", "NATAL", "SESUDAH NATAL", "TAHUN BARU",
        "SESUDAH TAHUN BARU", "EPIFANI", "SEPTUAGESIMA", "SEXAGESIMA",
        "ESTOMIHI", "INVOCAVIT", "REMINISCERE", "OKULI", "LETARE", "LAETARE",
        "JUDIKA", "PALMARUM", "JUMAT AGUNG", "PASKAH", "QUASIMODOGENITI",
        "MISERIKORDIAS DOMINI", "JUBILATE", "KANTATE", "ROGATE", "KENAIKAN",
        "EXAUDI", "PENTAKOSTA", "TRINITATIS", "SESUDAH TRINITATIS", "AKHIR TAHUN GEREJA"
    ]

    for i, text in enumerate(paragraphs[:25]):
        text_upper = text.upper()

        tgl_match = re_tanggal.search(text)
        if tgl_match:
            tanggal = tgl_match.group()

        if (any(k in text_upper for k in keywords_minggu) or "MINGGU" in text_upper) and not tgl_match:
            if len(text.split()) < 12:
                clean_m = re.sub(r"^[PLU]\s*[:\-]\s*", "",
                                 text, flags=re.IGNORECASE).strip()
                if "TATA IBADAH" in clean_m.upper():
                    m_frag = re.search(r"MINGGU.*", clean_m, re.IGNORECASE)
                    if m_frag:
                        clean_m = m_frag.group()
                nama_minggu = clean_m

        if "TOPIK" in text_upper:
            if ":" in text:
                res = text.split(":", 1)[1].strip()
                if len(res) > 1:
                    topik = res

            if not topik and i + 1 < len(paragraphs):
                next_text = paragraphs[i+1]
                if len(next_text) > 3 and "HURIA" not in next_text.upper():
                    topik = next_text

        if not topik and i < 15:
            # Mengutamakan teks dalam tanda petik yang bukan merupakan bagian dari ayat atau nama minggu
            quote_match = re.search(r"[“\"].*?[”\"]", text)
            if quote_match:
                candidate = quote_match.group().strip("“ ” \"")
                if len(candidate) > 10 and not any(k in candidate.upper() for k in keywords_minggu):
                    topik = candidate

    if topik:
        topik = topik.strip("“ ” \"").upper()

    if nama_minggu:
        nama_minggu = nama_minggu.upper()
        if "TATA IBADAH" in nama_minggu:
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
