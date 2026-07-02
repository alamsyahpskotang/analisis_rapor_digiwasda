# 🤖 PANDUAN DEPLOY BOT ANALISIS RAPOR PENDIDIKAN

## Metodologi DIGIWASDA - Sistem Digital untuk Pengawas Sekolah Berdampak

---

## 📋 DAFTAR ISI
1. [Apa itu Bot DIGIWASDA?](#apa-itu-bot)
2. [3 Opsi Deployment](#opsi-deployment)
3. [Langkah Deploy ke Streamlit Cloud](#langkah-deploy)
4. [Cara Menggunakan Bot](#cara-menggunakan)
5. [Troubleshooting](#troubleshooting)

---

## 🎯 APA ITU BOT DIGIWASDA? {#apa-itu-bot}

Bot ini adalah **sistem otomatis** yang menggunakan **metodologi DIGIWASDA** (Sistem Digital untuk Pengawas Sekolah Berdampak) untuk:
- ✅ Upload 1-9 file Rapor Pendidikan sekaligus
- ✅ Automatic extract & analyze data
- ✅ Generate ranking & clustering
- ✅ Generate RKT (Rencana Kerja Tahunan) dengan fokus Coaching Cycle 4-Fase
- ✅ Generate RKAS (Rincian Kegiatan Anggaran) otomatis
- ✅ Generate Laporan Final komprehensif
- ✅ Download hasil dalam format Word/Excel/PDF

**Gaya analisis**: Sama dengan yang sudah dilakukan untuk 9 sekolah - menggunakan metodologi **DIGIWASDA**

### Apa itu DIGIWASDA?
**DIGIWASDA = Digital untuk Pengawas Sekolah Berdampak**
- Fokus pada **Coaching Cycle 4-Fase** untuk Kepala Sekolah
- Berbasis **Pendampingan Berbasis Data (PBD)**
- Mendukung **Kebijakan Merdeka Belajar**
- Meningkatkan **Kualitas Kepemimpinan Sekolah**

---

## 🚀 3 OPSI DEPLOYMENT {#opsi-deployment}

### OPSI 1: STREAMLIT CLOUD ⭐ (REKOMENDASI - GRATIS)
**Keunggulan:**
- ✅ Gratis selamanya
- ✅ Akses dari mana saja via link
- ✅ Tidak perlu install apapun
- ✅ Bisa dishare ke siapa saja
- ✅ Data center di cloud

**Langkah:**
1. Upload file ke GitHub
2. Connect ke Streamlit Cloud
3. Deploy dengan 1 klik

---

### OPSI 2: GOOGLE COLAB (GRATIS)
**Keunggulan:**
- ✅ Tidak perlu install
- ✅ Gratis selamanya
- ✅ Akses via link

---

### OPSI 3: DESKTOP APP (LOKAL)
**Keunggulan:**
- ✅ Offline - tidak butuh internet
- ✅ Standalone executable
- ✅ Bisa diinstall di komputer

---

## 📦 LANGKAH DEPLOY KE STREAMLIT CLOUD {#langkah-deploy}

### TAHAP 1: SETUP DI GITHUB

**1.1 Buat Repository GitHub**
```bash
1. Buka github.com
2. Klik "New repository"
3. Nama: "bot-analisis-rapor-digiwasda"
4. Buat README
5. Clone ke komputer
```

**1.2 Upload File ke GitHub**
```bash
# Copy 3 file ini ke folder repository:
- bot_analisis_rapor_digiwasda.py    (main app)
- requirements.txt                    (dependencies)
- README.md                          (dokumentasi)

# Push ke GitHub
git add .
git commit -m "Upload Bot Analisis Rapor DIGIWASDA"
git push
```

### TAHAP 2: CONNECT KE STREAMLIT CLOUD

**2.1 Buat Akun Streamlit Cloud**
```
1. Buka streamlit.io/cloud
2. Klik "Sign up"
3. Pilih "Sign in with GitHub"
4. Authorize Streamlit
```

**2.2 Deploy App**
```
1. Di Streamlit Cloud dashboard klik "New app"
2. Pilih:
   - Repository: bot-analisis-rapor-digiwasda
   - Branch: main
   - File: bot_analisis_rapor_digiwasda.py
3. Klik "Deploy"
```

**2.3 Tunggu Deployment**
- Streamlit akan build app (1-2 menit)
- Akses via link: https://[username]-bot-analisis-rapor.streamlit.app

### TAHAP 3: SHARE LINK

```
Bot DIGIWASDA Anda sekarang live! 🎉

Link: https://[username]-bot-analisis-rapor.streamlit.app

Share ke:
- Dinas Pendidikan
- Pengawas
- Kepala Sekolah
- Komite Sekolah
```

---

## 💻 CARA MENGGUNAKAN BOT DIGIWASDA {#cara-menggunakan}

### LANGKAH 1: BUKA BOT
```
1. Buka link di browser
   https://[username]-bot-analisis-rapor.streamlit.app

2. Atau run lokal:
   streamlit run bot_analisis_rapor_digiwasda.py
```

### LANGKAH 2: UPLOAD FILE
```
Tab "Upload & Input":
1. Klik "Upload file Rapor Pendidikan"
2. Pilih 1 atau lebih file Excel
3. Tunggu sistem menganalisis dengan DIGIWASDA (beberapa detik)
4. Lihat daftar sekolah yang teranalisis
```

### LANGKAH 3: ANALISIS RANKING
```
Tab "Analisis Ranking":
1. Lihat tabel ranking 1-N
2. Lihat summary statistics:
   - Tier 1 (CUKUP) vs Tier 2 (EMERGENCY)
   - Rata-rata keseluruhan
   - Sekolah terbaik
3. Download ranking ke Excel
```

### LANGKAH 4: ANALISIS DETAIL
```
Tab "Analisis Detail":
1. Pilih sekolah dari dropdown
2. Lihat:
   - Rata-rata mutu & status
   - Perincian 7 indikator
   - Chart visualisasi
3. Identifikasi indikator lemah
```

### LANGKAH 5: GENERATE RKT & RKAS
```
Tab "RKT & RKAS":
1. Klik "Generate RKT & RKAS untuk Semua Sekolah"
2. Lihat:
   - Summary Tier 1 & Tier 2
   - Total budget
   - Coaching Cycle focus
3. Download RKT & RKAS Excel
```

### LANGKAH 6: GENERATE LAPORAN FINAL
```
Tab "Laporan Final":
1. Pilih format:
   - Word (.docx) - untuk presentasi & coaching
   - Excel - untuk data analysis
   - PDF - untuk distribusi
2. Klik "Generate Laporan Final"
3. Download file dengan branding DIGIWASDA
```

---

## 🛠️ SETUP LOKAL (KOMPUTER PRIBADI) {#setup-lokal}

### OPSI A: DENGAN PYTHON

**1. Install Python**
```bash
# Download dari python.org (versi 3.9+)
# Install dengan mencentang "Add Python to PATH"
```

**2. Setup Environment**
```bash
# Buat folder untuk bot
mkdir bot-analisis-rapor-digiwasda
cd bot-analisis-rapor-digiwasda

# Copy file bot_analisis_rapor_digiwasda.py ke folder ini

# Buat virtual environment
python -m venv venv

# Activate virtual environment
# Windows:
venv\Scripts\activate
# Mac/Linux:
source venv/bin/activate
```

**3. Install Dependencies**
```bash
pip install -r requirements.txt
```

**4. Run Bot**
```bash
streamlit run bot_analisis_rapor_digiwasda.py
```

**5. Akses**
```
Bot akan membuka di browser:
http://localhost:8501
```

---

## 📊 FILE YANG DIBUTUHKAN {#file-yang-dibutuhkan}

### File Utama:
1. **bot_analisis_rapor_digiwasda.py**
   - Main application code
   - Logic untuk analysis & generation dengan DIGIWASDA

2. **requirements.txt**
   - Daftar library yang dibutuhkan

3. **README.md**
   - Dokumentasi untuk GitHub

### File Input (dari user):
- **File Rapor Pendidikan (.xlsx)**
  - Format: Standard Rapor PBD Kemendikdasmen
  - Sheet wajib: "2. LAPORAN RAPOR"

---

## 🔧 TROUBLESHOOTING {#troubleshooting}

### MASALAH 1: "ModuleNotFoundError: No module named 'streamlit'"
**Solusi:**
```bash
pip install streamlit pandas openpyxl
```

### MASALAH 2: "File Excel tidak terbaca"
**Solusi:**
- Pastikan file berformat .xlsx (bukan .xls)
- Pastikan ada sheet "2. LAPORAN RAPOR"
- Coba buka file di Excel terlebih dahulu

### MASALAH 3: "Bot lambat saat upload banyak file"
**Solusi:**
- Normal untuk 5-9 file (5-30 detik)
- Upload file satu per satu jika ada error
- Refresh page jika stuck

### MASALAH 4: "Streamlit Cloud deployment failed"
**Solusi:**
- Pastikan requirements.txt benar
- Cek versi Python (3.9+)
- Lihat logs di Streamlit Cloud
- Hubungi Streamlit support

---

## 📝 WORKFLOW MENGGUNAKAN BOT DIGIWASDA

### MINGGU 1: SETUP
```
1. Deploy bot DIGIWASDA ke Streamlit Cloud ✅
2. Test dengan 1-2 file Rapor ✅
3. Share link ke dinas/pengawas ✅
```

### MINGGU 2-4: ANALISIS
```
1. Dinas upload 9 file Rapor sekolah
2. Bot DIGIWASDA automatic generate:
   - Ranking & Klusterisasi
   - RKT & RKAS dengan Coaching Cycle
   - Laporan Final komprehensif
3. Download & review hasil
```

### MINGGU 4+: IMPLEMENTASI
```
1. Share laporan ke 9 sekolah
2. FGD dengan Kepala Sekolah
3. Mulai Coaching Cycle 4-Fase
4. Monitor progress dengan DIGIWASDA tracking
```

---

## 🎓 KESIMPULAN

Bot DIGIWASDA ini memudahkan:
- ✅ Analisis Rapor untuk banyak sekolah
- ✅ Automatic generate RKT & RKAS dengan fokus Coaching
- ✅ Consistent methodology (DIGIWASDA)
- ✅ Share hasil ke stakeholder
- ✅ Reuse untuk periode berikutnya

**Dengan bot DIGIWASDA, analisis yang biasanya memerlukan waktu 3-5 hari bisa selesai dalam 5-10 menit!** ⚡

---

**Happy analyzing dengan DIGIWASDA! 🚀**
**Bot Analisis Rapor Pendidikan | Metodologi DIGIWASDA | 2025**

