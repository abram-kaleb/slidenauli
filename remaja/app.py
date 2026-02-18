# app.py
import streamlit as st
from docx import Document
from pptx import Presentation
from io import BytesIO
import re
from cover import extract_cover
from isi import extract_isi
from ppt import generate_slides

st.set_page_config(page_title="PPT Generator", layout="centered")
st.title("PPT Automation System")


@st.cache_resource
def get_document(file_bytes):
    return Document(BytesIO(file_bytes))


uploaded_file = st.file_uploader(
    "Upload File Tata Ibadah (DOCX)", type=["docx"])

if uploaded_file is not None:
    try:
        file_bytes = uploaded_file.getvalue()
        doc = get_document(file_bytes)

        data_cover = extract_cover(doc)
        data_isi = extract_isi(doc)

        st.markdown("### 📊 Meta Data")
        col1, col2 = st.columns(2)
        with col1:
            val_tata = st.text_input("Tata Ibadah", data_cover['tata_ibadah'])
            val_minggu = st.text_input("Nama Minggu", data_cover['minggu'])

        with col2:

            val_topik = st.text_area("Topik", data_cover['topik'])
            val_tanggal = st.text_input("Tanggal", data_cover['tanggal'])

        st.divider()

        # Proses generate langsung dilakukan saat tombol diklik
        prs = Presentation()
        cover_info = {
            "minggu": val_minggu,
            "topik": val_topik,
            "tanggal": val_tanggal
        }
        generate_slides(prs, cover_info, data_isi)

        ppt_io = BytesIO()
        prs.save(ppt_io)
        ppt_io.seek(0)

        # Tombol download sekarang sekaligus menjadi pemicu utama
        st.download_button(
            label="🚀 Generate & Download PPTX",
            data=ppt_io,
            file_name=f"PPT_{val_minggu.replace(' ', '_')}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
            type="primary"
        )

        st.markdown("### 📑 Preview Acara")
        for section in data_isi:
            with st.expander(f"Acara {section['nomor']}: {section['judul']}"):
                for line in section['isi']:
                    st.text(line)

    except Exception as e:
        st.error(f"Terjadi Kesalahan: {e}")
        if st.button("Clear Cache & Rerun"):
            st.cache_resource.clear()
            st.rerun()
