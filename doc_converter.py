# doc_converter.py
import subprocess
import tempfile
import os
import shutil


def convert_doc_to_docx(file_bytes: bytes, filename: str) -> bytes:
    """
    Konversi file .doc ke .docx menggunakan LibreOffice (headless).
    Kompatibel dengan Streamlit Cloud (Linux) dan lokal.

    Args:
        file_bytes: isi file .doc dalam bytes
        filename: nama file asli (misal: "tata_ibadah.doc")

    Returns:
        bytes dari file .docx hasil konversi

    Raises:
        RuntimeError: jika LibreOffice tidak tersedia atau konversi gagal
    """
    # Cek apakah LibreOffice tersedia
    lo_path = shutil.which("libreoffice") or shutil.which("soffice")
    if not lo_path:
        raise RuntimeError(
            "LibreOffice tidak ditemukan. "
            "Tambahkan 'libreoffice' ke packages.txt di Streamlit Cloud, "
            "atau install via: sudo apt-get install libreoffice"
        )

    with tempfile.TemporaryDirectory() as tmpdir:
        # Simpan file .doc ke temp folder
        input_path = os.path.join(tmpdir, filename)
        with open(input_path, "wb") as f:
            f.write(file_bytes)

        # Jalankan LibreOffice headless untuk konversi
        result = subprocess.run(
            [
                lo_path,
                "--headless",
                "--convert-to", "docx",
                "--outdir", tmpdir,
                input_path,
            ],
            capture_output=True,
            text=True,
            timeout=60,
        )

        if result.returncode != 0:
            raise RuntimeError(
                f"LibreOffice gagal mengkonversi file.\n"
                f"stderr: {result.stderr}\nstdout: {result.stdout}"
            )

        # Baca hasil konversi
        base_name = os.path.splitext(filename)[0]
        output_path = os.path.join(tmpdir, f"{base_name}.docx")

        if not os.path.exists(output_path):
            raise RuntimeError(
                f"File hasil konversi tidak ditemukan di: {output_path}\n"
                f"stdout: {result.stdout}"
            )

        with open(output_path, "rb") as f:
            return f.read()


def is_doc_file(filename: str) -> bool:
    """Cek apakah file adalah .doc (bukan .docx)."""
    return filename.lower().endswith(".doc") and not filename.lower().endswith(".docx")


def ensure_docx_bytes(file_bytes: bytes, filename: str) -> tuple[bytes, str]:
    """
    Pastikan file dalam format .docx.
    Jika .doc, otomatis konversi dulu.

    Returns:
        tuple (docx_bytes, docx_filename)
    """
    if is_doc_file(filename):
        docx_bytes = convert_doc_to_docx(file_bytes, filename)
        docx_filename = os.path.splitext(filename)[0] + ".docx"
        return docx_bytes, docx_filename
    return file_bytes, filename
