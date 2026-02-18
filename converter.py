import streamlit as st
import requests
import time
import os

# ─────────────────────────────────────────────
#  Konfigurasi Halaman
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="DOC → DOCX Converter",
    page_icon="📄",
    layout="centered",
)

st.title("📄 Konversi DOC ke DOCX")
st.markdown(
    "Upload file **.doc** lama Anda dan konversi ke format **.docx** modern "
    "menggunakan **CloudConvert API**."
)

# ─────────────────────────────────────────────
#  Sidebar — API Key
# ─────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Konfigurasi")
    st.markdown(
        "Dapatkan API key gratis di [cloudconvert.com](https://cloudconvert.com/api/v2).\n\n"
        "Free tier: **25 konversi/hari**."
    )
    api_key = st.text_input(
        "CloudConvert API Key",
        type="password",
        placeholder="Masukkan API key Anda...",
    )
    st.divider()
    st.markdown("**Cara mendapatkan API key:**")
    st.markdown(
        "1. Daftar di [cloudconvert.com](https://cloudconvert.com)\n"
        "2. Masuk ke **Dashboard → API Keys**\n"
        "3. Klik **Create API Key**\n"
        "4. Pilih scope: `task.read` & `task.write`"
    )

# ─────────────────────────────────────────────
#  Upload File
# ─────────────────────────────────────────────
uploaded_file = st.file_uploader(
    "Upload file DOC",
    type=["doc"],
    help="Hanya file .doc yang didukung",
)

if uploaded_file:
    st.info(
        f"📁 File: **{uploaded_file.name}** ({uploaded_file.size / 1024:.1f} KB)")

# ─────────────────────────────────────────────
#  Fungsi Konversi via CloudConvert
# ─────────────────────────────────────────────


def convert_doc_to_docx(file_bytes: bytes, filename: str, api_key: str) -> bytes:
    """Konversi DOC ke DOCX menggunakan CloudConvert API v2."""

    HEADERS = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    BASE_URL = "https://api.cloudconvert.com/v2"

    # ── 1. Buat Job (import + convert + export) ──────────────────────────────
    status_placeholder.info("🔄 Langkah 1/4 — Membuat job konversi...")
    job_payload = {
        "tasks": {
            "upload-file": {
                "operation": "import/upload"
            },
            "convert-file": {
                "operation": "convert",
                "input": "upload-file",
                "input_format": "doc",
                "output_format": "docx",
            },
            "export-file": {
                "operation": "export/url",
                "input": "convert-file",
            },
        }
    }
    resp = requests.post(f"{BASE_URL}/jobs", json=job_payload, headers=HEADERS)
    resp.raise_for_status()
    job = resp.json()["data"]
    job_id = job["id"]

    # Ambil upload task
    upload_task = next(
        t for t in job["tasks"] if t["name"] == "upload-file"
    )
    upload_url = upload_task["result"]["form"]["url"]
    upload_params = upload_task["result"]["form"]["parameters"]

    # ── 2. Upload File ────────────────────────────────────────────────────────
    status_placeholder.info("🔄 Langkah 2/4 — Mengupload file...")
    files = {"file": (filename, file_bytes, "application/msword")}
    upload_resp = requests.post(upload_url, data=upload_params, files=files)
    upload_resp.raise_for_status()

    # ── 3. Tunggu sampai selesai ──────────────────────────────────────────────
    status_placeholder.info(
        "🔄 Langkah 3/4 — Mengkonversi file (mohon tunggu)...")
    for attempt in range(60):  # max 60 detik
        time.sleep(2)
        job_resp = requests.get(f"{BASE_URL}/jobs/{job_id}", headers=HEADERS)
        job_resp.raise_for_status()
        job_data = job_resp.json()["data"]
        job_status = job_data["status"]

        if job_status == "finished":
            break
        elif job_status == "error":
            error_task = next(
                (t for t in job_data["tasks"] if t["status"] == "error"), None
            )
            error_msg = error_task["message"] if error_task else "Unknown error"
            raise RuntimeError(f"Konversi gagal: {error_msg}")
    else:
        raise TimeoutError("Konversi timeout setelah 120 detik.")

    # ── 4. Download hasil ─────────────────────────────────────────────────────
    status_placeholder.info("🔄 Langkah 4/4 — Mengunduh hasil konversi...")
    export_task = next(
        t for t in job_data["tasks"] if t["name"] == "export-file"
    )
    download_url = export_task["result"]["files"][0]["url"]

    dl_resp = requests.get(download_url)
    dl_resp.raise_for_status()
    return dl_resp.content


# ─────────────────────────────────────────────
#  Tombol Konversi
# ─────────────────────────────────────────────
st.divider()
col1, col2 = st.columns([1, 3])
with col1:
    convert_btn = st.button(
        "🚀 Konversi Sekarang",
        disabled=not (uploaded_file and api_key),
        use_container_width=True,
        type="primary",
    )

if not api_key:
    st.caption("⚠️ Masukkan API key di sidebar terlebih dahulu.")
if not uploaded_file:
    st.caption("⚠️ Upload file .doc terlebih dahulu.")

status_placeholder = st.empty()

if convert_btn and uploaded_file and api_key:
    try:
        file_bytes = uploaded_file.read()
        original_name = os.path.splitext(uploaded_file.name)[0]
        output_filename = f"{original_name}.docx"

        result_bytes = convert_doc_to_docx(
            file_bytes, uploaded_file.name, api_key)

        status_placeholder.success("✅ Konversi berhasil!")

        st.download_button(
            label=f"⬇️ Download {output_filename}",
            data=result_bytes,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

    except requests.exceptions.HTTPError as e:
        if e.response.status_code == 401:
            status_placeholder.error(
                "❌ API key tidak valid atau tidak memiliki izin.")
        elif e.response.status_code == 402:
            status_placeholder.error(
                "❌ Limit konversi habis. Upgrade plan atau tunggu besok.")
        else:
            status_placeholder.error(
                f"❌ HTTP Error: {e.response.status_code} — {e.response.text}")
    except TimeoutError as e:
        status_placeholder.error(f"❌ {e}")
    except RuntimeError as e:
        status_placeholder.error(f"❌ {e}")
    except Exception as e:
        status_placeholder.error(f"❌ Terjadi kesalahan: {e}")

# ─────────────────────────────────────────────
#  Footer Info
# ─────────────────────────────────────────────
st.divider()
with st.expander("ℹ️ Tentang Aplikasi Ini"):
    st.markdown(
        """
        **Cara Kerja:**
        1. File `.doc` Anda diupload ke **CloudConvert** melalui API yang terenkripsi.
        2. CloudConvert mengkonversi file ke format `.docx`.
        3. Hasil konversi diunduh langsung ke browser Anda.
        4. File dihapus otomatis dari server CloudConvert setelah diunduh.

        **Keamanan:**
        - File hanya tersimpan sementara di server CloudConvert (ISO 27001 certified).
        - API key disimpan hanya di session Anda, tidak dikirim ke mana pun selain CloudConvert.

        **Batasan Free Tier:** 25 konversi/hari — cukup untuk penggunaan personal.
        """
    )
