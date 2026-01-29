🚀 SLIDENAULI: Document Converter
Slidenauli adalah aplikasi open source berbasis Streamlit yang dirancang khusus untuk mengotomatisasi pembuatan slide presentasi (PPTX) ibadah dari dokumen Tata Ibadah dan Warta Jemaat (DOCX).

Aplikasi ini lahir dari kebutuhan efisiensi pelayanan, memungkinkan transformasi dokumen liturgi menjadi visual presentasi hanya dalam hitungan detik. Dibuat dengan ❤️ oleh Tim Multimedia HKBP Perum 2 Bekasi.

🌟 Fitur Utama
Smart Conversion: Mengubah dokumen Word (.docx) menjadi slide PowerPoint (.pptx) secara otomatis dalam hitungan detik.

Dual Mode Format:
Projector: Desain penuh dengan latar belakang dinamis untuk kebutuhan di dalam gereja.
Streaming: Desain minimalis dengan lower-third hijau (chroma key) untuk kebutuhan OBS/vMix saat Live Streaming.
Metadata Extraction: Mendeteksi secara otomatis nama Minggu, Topik, dan Bahasa (Batak/Indonesia) dari dokumen yang diunggah.
Automatic Image Handling: Mengekstrak gambar laporan keuangan atau pengumuman langsung dari file Warta Jemaat ke dalam slide.
Smart Text Splitting: Memecah teks yang panjang menjadi beberapa slide secara cerdas agar tetap nyaman dibaca.
Dynamic Backgrounds: Memilih secara acak dari 200+ latar belakang estetis untuk mode Projector.

🛠️ Teknologi yang Digunakan
Python: Bahasa pemrograman utama.
Streamlit: Framework untuk antarmuka web yang interaktif.
python-pptx: Library untuk manipulasi file PowerPoint.
python-docx: Library untuk membaca konten dari file Word.
Regex (re): Untuk pemrosesan bahasa alami dan identifikasi struktur liturgi.

🚀 Cara Menjalankan Secara Lokal
Clone Repository

Bash
git clone https://github.com/abram-kaleb/slidenauli
cd slidenauli
Instalasi Dependensi Pastikan Anda sudah menginstal Python, lalu jalankan:

Bash
pip install streamlit python-pptx python-docx lxml requests
Persiapan Folder Gambar Aplikasi ini mencari gambar latar belakang di folder pics/. Pastikan Anda memiliki file gambar dengan penamaan angka (contoh: 1.jpg, 2.jpg, dst) di dalam folder tersebut.

Jalankan Aplikasi

Bash
streamlit run app.py
