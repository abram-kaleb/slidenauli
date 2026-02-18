# app.py
import streamlit as st
from docx import Document
import logic
import indo_umum.cover as indo_cover
import indo_umum.isi as indo_isi
import indo_umum.ppt as indo_ppt
import batak_umum.cover as batak_cover
import batak_umum.isi as batak_isi
import batak_umum.ppt as batak_ppt
import remaja.cover as remaja_cover
import remaja.isi as remaja_isi
import remaja.ppt as remaja_ppt
import sore.cover as sore_cover
import sore.isi as sore_isi
import sore.ppt as sore_ppt
import skm.cover as skm_cover
import skm.isi as skm_isi
import skm.ppt as skm_ppt
from doc_converter import ensure_docx_bytes

st.set_page_config(page_title="SLIDENAULI", layout="centered")

if "tata_bytes" not in st.session_state:
    st.session_state.tata_bytes = None
if "warta_bytes" not in st.session_state:
    st.session_state.warta_bytes = None
if "last_tata_name" not in st.session_state:
    st.session_state.last_tata_name = None
if "last_warta_name" not in st.session_state:
    st.session_state.last_warta_name = None

st.markdown("""
    <style>
    .block-container {
        padding-top: 2rem;
        max-width: 800px;
    }
    .custom-header {
        text-align: center;
        margin-bottom: 2rem;
    }
    .header-title {
        font-size: 64px;
        font-weight: 850;
        letter-spacing: -2px;
        margin-bottom: 0px;
    }
    .header-pink { color: #ff2b5e; }
    .header-subtitle {
        font-size: 16px;
        opacity: 0.7;
        margin-top: -10px;
    }
    .section-title {
        font-size: 32px;
        font-weight: 700;
        margin-top: 2rem;
        margin-bottom: 1.5rem;
    }
    .upload-label {
        font-weight: 500;
        margin-bottom: 10px;
        display: block;
    }
    hr {
        margin: 2rem 0;
        opacity: 0.1;
    }
    .stButton > button {
        width: 100%;
        border-radius: 8px;
        height: 3.5rem;
        background-color: #ff2b5e;
        color: white;
        font-weight: 600;
        border: none;
    }
    .stDownloadButton > button {
        width: 100% !important;
        background-color: #ff2b5e !important;
        color: white !important;
        border-radius: 8px;
        height: 3.5rem;
        font-weight: 600;
        border: none;
    }
    .meta-note {
        font-size: 14px;
        color: #ffffff;
        margin-top: 20px;
    }
    </style>
    
    <div class="custom-header">
        <h1 class="header-title">SLIDE<span class="header-pink">NAULI</span> 🚀</h1>
        <p class="header-subtitle">Document converter by Multimedia HKBP Perum 2 Bekasi</p>
    </div>
    <hr>
    """, unsafe_allow_html=True)


def get_document(file_bytes):
    return Document(logic.BytesIO(file_bytes))


st.markdown('<h2 class="section-title">1. Dokumen</h2>',
            unsafe_allow_html=True)
col1, col2 = st.columns(2)

with col1:
    st.markdown('<span class="upload-label">📁 TATA IBADAH \n+ Upload file .doc atau .docx</span>',
                unsafe_allow_html=True)
    uploaded_tata = st.file_uploader(
        "tata", type=["doc", "docx"], key="tata_up", label_visibility="collapsed")

with col2:
    st.markdown('<span class="upload-label">📁 WARTA \n+ Upload file .doc atau .docx</span>',
                unsafe_allow_html=True)
    uploaded_warta = st.file_uploader(
        "warta", type=["doc", "docx"], key="warta_up", label_visibility="collapsed")

if uploaded_tata:
    if uploaded_tata.name != st.session_state.last_tata_name:
        with st.spinner("Memproses file Tata Ibadah..."):
            try:
                tata_bytes, _ = ensure_docx_bytes(
                    uploaded_tata.getvalue(), uploaded_tata.name)
                st.session_state.tata_bytes = tata_bytes
                st.session_state.last_tata_name = uploaded_tata.name
            except RuntimeError as e:
                st.error(f"❌ Gagal memproses file Tata Ibadah: {e}")
                st.session_state.tata_bytes = None
                st.session_state.last_tata_name = None
else:
    st.session_state.tata_bytes = None
    st.session_state.last_tata_name = None

if uploaded_warta:
    if uploaded_warta.name != st.session_state.last_warta_name:
        with st.spinner("Memproses file Warta..."):
            try:
                warta_bytes, _ = ensure_docx_bytes(
                    uploaded_warta.getvalue(), uploaded_warta.name)
                st.session_state.warta_bytes = warta_bytes
                st.session_state.last_warta_name = uploaded_warta.name
            except RuntimeError as e:
                st.error(f"❌ Gagal memproses file Warta: {e}")
                st.session_state.warta_bytes = None
                st.session_state.last_warta_name = None
else:
    st.session_state.warta_bytes = None
    st.session_state.last_warta_name = None

det_tata = "Unknown"
det_warta = "None"
w_mode_final = "Normal"

if st.session_state.tata_bytes:
    doc_tata_check = get_document(st.session_state.tata_bytes)
    det_tata = logic.detect_format(doc_tata_check)
    if "Warta" in det_tata:
        st.error(f"❌ Terdeteksi {det_tata}. Mohon upload di kolom Warta.")
    elif det_tata == "Sekolah Minggu (SKM)":
        st.warning(
            f"⚠️ Terdeteksi: {det_tata}. Fitur ini masih dalam pengembangan.")
    else:
        st.success(f"✅ Terdeteksi: {det_tata}")

if st.session_state.warta_bytes:
    doc_warta_check = get_document(st.session_state.warta_bytes)
    det_warta = logic.detect_format(doc_warta_check)
    if "Warta" not in det_warta:
        st.error(f"❌ Terdeteksi {det_warta}. Ini bukan file Warta.")
    else:
        st.success(f"✅ Terdeteksi: {det_warta}")

if st.session_state.tata_bytes and st.session_state.warta_bytes:
    if det_tata == "Ibadah Remaja":
        if det_warta == "Warta Remaja":
            w_mode_final = "Wide"
        else:
            st.warning(
                "⚠️ Peringatan: Tata Ibadah Remaja harusnya menggunakan Warta Remaja.")
            w_mode_final = "Normal"
    else:
        if det_warta == "Warta Remaja":
            st.error(
                "❌ Kesalahan: Warta Remaja seharusnya digunakan untuk Tata Ibadah Remaja.")
        w_mode_final = "Normal"

if st.session_state.tata_bytes:
    doc_tata = get_document(st.session_state.tata_bytes)
    st.markdown('<h2 class="section-title">2. Pengaturan</h2>',
                unsafe_allow_html=True)

    c_set1, c_set2 = st.columns(2)
    options = ["Ibadah Indonesia Umum", "Ibadah Batak Umum",
               "Ibadah Remaja", "Ibadah Sore", "Sekolah Minggu (SKM)"]
    mapping = {
        "Sekolah Minggu (SKM)": (skm_cover, skm_isi, skm_ppt),
        "Ibadah Sore": (sore_cover, sore_isi, sore_ppt),
        "Ibadah Remaja": (remaja_cover, remaja_isi, remaja_ppt),
        "Ibadah Batak Umum": (batak_cover, batak_isi, batak_ppt),
        "Ibadah Indonesia Umum": (indo_cover, indo_isi, indo_ppt)
    }

    with c_set1:
        selected_fmt = st.selectbox("Format", options, index=options.index(
            det_tata) if det_tata in options else 0)
    with c_set2:
        use_bg = st.selectbox("Gunakan Background", ["Ya", "Tidak"], index=1)

    m_cover, m_isi, m_ppt = mapping[selected_fmt]
    data_cover = m_cover.extract_cover(doc_tata)
    data_isi = m_isi.extract_isi(doc_tata)

    final_warta_doc = get_document(
        st.session_state.warta_bytes) if st.session_state.warta_bytes else None

    if st.button("🚀 Proses Dokumen"):
        c_info = {"minggu": data_cover.get('minggu', ''),
                  "topik": data_cover.get('topik', ''),
                  "tanggal": data_cover.get('tanggal', ''),
                  "use_bg": True if use_bg == "Ya" else False}

        final_ppt = logic.merge_and_generate(
            final_warta_doc, c_info, data_isi, m_ppt.generate_slides, w_mode_final)

        file_name = f"ppt_{selected_fmt}_{data_cover.get('tanggal', 'slide')}.pptx".replace(
            " ", "_")

        st.download_button("📥 Download PPT", final_ppt, file_name)

    st.markdown(f"""
        <div class="meta-note">
            <b>Informasi:</b><br>
            • <b>Minggu:</b> {data_cover.get('minggu', '-')}<br>
            • <b>Tanggal:</b> {data_cover.get('tanggal', '-')}<br>
            • <b>Topik:</b> {data_cover.get('topik', '-')}
        </div>
    """, unsafe_allow_html=True)
