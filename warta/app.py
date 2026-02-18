# app.py

import streamlit as st
import io
import sys
import os

sys.path.append(os.path.abspath(os.path.dirname(__file__)))

try:
    from warta_normal import generate_full_warta as generate_normal
    from warta_wide import generate_full_warta as generate_wide
except ImportError:
    from warta.warta_normal import generate_full_warta as generate_normal
    from warta.warta_wide import generate_full_warta as generate_wide

st.set_page_config(page_title="Warta Exporter", layout="centered")

st.title("Warta Word to PPTX Converter")
st.write("Unggah file Warta Jemaat (.docx) untuk mengekstraksi seluruh teks dan gambar secara otomatis.")

uploaded_file = st.file_uploader("Upload Warta .docx", type=["docx"])

if uploaded_file:
    file_name = uploaded_file.name.upper()

    if "REMAJA" in file_name:
        mode_text = "Mode Terdeteksi: Warta Remaja (Widescreen 16:9, Font 54)"
        gen_func = generate_wide
    else:
        mode_text = "Mode Terdeteksi: Warta Normal (Standard 4:3, Font 40)"
        gen_func = generate_normal

    st.info(mode_text)

    try:
        # Proses generate dilakukan secara otomatis saat file diunggah
        with st.spinner("Memproses seluruh bagian warta..."):
            ppt_output = gen_func(uploaded_file)

        st.success("Slide Berhasil Dibuat!")

        # Tombol download muncul menggantikan tombol generate
        st.download_button(
            label="📥 Download Hasil PPTX",
            data=ppt_output,
            file_name=f"Hasil_{uploaded_file.name.split('.')[0]}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True
        )

        if st.button("🔄 Reset", use_container_width=True):
            st.rerun()

    except Exception as e:
        st.error(f"Terjadi kesalahan saat generate: {e}")
