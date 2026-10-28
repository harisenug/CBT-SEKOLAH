# SISTEM SIMULASI CBT SLIDE INTERAKTIF (MULTI-SEKOLAH)

Aplikasi Computer-Based Test (CBT) berbasis presentasi Slide PPT/PDF interaktif dengan arsitektur **Zero-Knowledge Security** (kunci jawaban tidak pernah dikirim ke browser siswa), **Multi-Tenant Private Partitioning** (partisi data terisolasi per sekolah & Google Drive), **Pengacakan Paket Varian Soal A/B/C/D Terkunci (Session Lock)**, serta **Ekspor Nilai Langsung ke Excel (.xlsx)**.

---

## 🌟 Fitur Utama

1. **Keamanan Anti-Bocor Kunci (Zero-Knowledge):**
   - Kunci jawaban hanya tersimpan di memori server dan divalidasi di backend saat siswa submit.
   - Kebal intip Inspect Element / Console F12 / Network Tab browser siswa.
2. **Multi-Paket Varian Soal (A/B/C/D):**
   - Mendukung kode soal bercabang untuk bobot dan materi yang sama.
   - Server mengacak salah satu varian per siswa saat login dan **menguncinya (Session Lock)** ke nomor peserta hingga ujian selesai.
   - Siswa bersebelahan mendapatkan nomor slide dan opsi yang berbeda sehingga tidak bisa saling contek.
3. **Universal Exam Link (TKA & Mapel Pilihan):**
   - Opsi default `?mapel=all&kelas=all&token=...` memungkinkan 1 link dan 1 token dibagikan serentak untuk seluruh sekolah. Siswa dapat memilih mata pelajarannya sendiri saat login.
4. **Ekspor Nilai Langsung ke Excel (.xlsx) & CSV:**
   - Rekap nilai dapat diunduh langsung menjadi file Spreadsheet Excel (`.xlsx`) lengkap dengan kolom No, Nama Siswa, No Peserta, Kelas, Mapel, Paket Soal, Nilai Akhir (0-100), Status, dan Waktu.
5. **Fitur "Clear Cache" (Persiapan Ujian Baru Bersih):**
   - Mengosongkan bank soal lama dan sesi siswa dengan 1 kali klik agar ujian baru benar-benar bersih tanpa memanggil soal sebelumnya.
6. **Dynamic Forensic Watermark:**
   - Setiap slide soal memuat watermark transparan berputar berisi Nama Siswa, No. Peserta, dan Sekolah untuk mencegah kecurangan foto kamera HP.
7. **Integrasi Google Drive Terpisah:**
   - Setiap sekolah dapat memasukkan link folder Google Drive masing-masing dengan script Google Apps Script anti-limit kuota.

---

## 🚀 Petunjuk Upload ke GitHub

Ikuti langkah-langkah mudah berikut di terminal/command prompt komputer Anda:

### Langkah 1: Buat Repository Baru di GitHub
1. Buka [https://github.com](https://github.com) dan login ke akun Anda.
2. Klik tombol **New** (atau tanda `+` di kanan atas > **New repository**).
3. Beri nama repository, contoh: `cbt-slide-multisekolah`.
4. Pilih **Public** atau **Private** sesuai kebutuhan.
5. **PENTING:** Jangan centang *"Add a README file"* (karena repositori ini sudah memiliki README dan konfigurasi lengkap).
6. Klik **Create repository**.

### Langkah 2: Inisialisasi Git di Folder Komputer Anda
Buka Terminal / Git Bash di direktori proyek ini, lalu jalankan perintah berikut secara berurutan:

```bash
# 1. Inisialisasi git repository
git init

# 2. Tambahkan semua file yang sudah bersih ke staging
git add .

# 3. Buat commit pertama
git commit -m "feat: inisialisasi sistem CBT Slide Multi-Sekolah aman anti-hack"

# 4. Ganti nama branch utama menjadi main
git branch -M main

# 5. Hubungkan ke repository GitHub Anda (ganti USERNAME dan REPO_NAME)
git remote add origin https://github.com/USERNAME/cbt-slide-multisekolah.git

# 6. Push kode ke GitHub
git push -u origin main
```

*(Catatan: File konfigurasi rahasia dan folder `node_modules/` sudah otomatis dikecualikan melalui file `.gitignore` sehingga aman).*

---

## 💻 Cara Menjalankan Proyek di Komputer / Server Baru

Setelah clone dari GitHub:

```bash
# 1. Masuk ke folder proyek
cd cbt-slide-multisekolah

# 2. Install dependencies (termasuk express, vite, tailwindcss, xlsx)
npm install

# 3. Jalankan server pengembangan (Development Mode)
npm run dev
# Buka http://localhost:3000 di browser Anda

# 4. Untuk Production Build & Deploy
npm run build
npm start
```

---

## 📄 Format Template Upload (CSV & Excel)

Sistem sudah mendukung otomatis pendeteksian pemisah titik-koma (`;`) khas Excel Windows Indonesia maupun pemisah koma (`,`):

### 1. Template Data Siswa (`template_Data_SISWA.csv`)
```csv
No ;Nama Siswa;No Peserta;Kelas;Mapel
1;ADI;2026001;12A;INGGRIS
2;SUTRISNO;2026002;12B;INGGRIS
3;BORIS YELSIN;2026003;12C;INGGRIS
```

### 2. Template Bank Soal & Kunci Varian (`template_Soal_Varian.csv`)
```csv
NO;MAPEL;SOAL;A;B;C;D;KUNCI A;KUNCI B;KUNCI C;kunci D
1;INGGRIS;INGGRIS_1;INGGRIS_1A;INGGRIS_1B;INGGRIS_1C;INGGRIS_1D;B;C;B;E
2;INGGRIS;INGGRIS_2;INGGRIS_2A;INGGRIS_2B;INGGRIS_2C;INGGRIS_2D;B;D;C;A
3;BIOLOGI;BIOLOGI_1;;;;;BENAR;;;
```
* **Aturan D.1:** Jika kolom A, B, C, D dikosongkan, soal tersebut adalah tipe tunggal standar.
* **Aturan D.2:** Jika kolom A, B, C, D berisi kode slide soal (`INGGRIS_1A`, dst), sistem secara otomatis mengacak dan mengunci varian tersebut per siswa saat login.

---

## 🛡️ Lisensi
Apache-2.0. Bebas dikembangkan dan dimodifikasi untuk kebutuhan instansi sekolah Anda.
