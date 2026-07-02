# 🤖 Bot Analisis Rapor Pendidikan - DIGIWASDA

**Sistem Digital untuk Pengawas Sekolah Berdampak**

## 📋 Deskripsi

Bot ini adalah sistem otomatis untuk analisis Rapor Pendidikan menggunakan metodologi **DIGIWASDA** (Sistem Digital untuk Pengawas Sekolah Berdampak).

Bot dapat:
- ✅ Upload 1-9 file Rapor Pendidikan sekaligus
- ✅ Automatic extract & analyze data
- ✅ Generate ranking & clustering sekolah
- ✅ Generate RKT (Rencana Kerja Tahunan) otomatis
- ✅ Generate RKAS (Rincian Kegiatan Anggaran) otomatis
- ✅ Generate Laporan Final komprehensif
- ✅ Download hasil dalam format Word/Excel/PDF

## 🎯 Fitur Utama

### 5 Tab Utama:
1. **📤 Upload & Input** - Upload file Rapor Pendidikan
2. **📊 Analisis Ranking** - Lihat ranking & klusterisasi 9 sekolah
3. **🔍 Analisis Detail** - Analisis detail per sekolah
4. **📋 RKT & RKAS** - Generate rencana kerja & anggaran
5. **📄 Laporan Final** - Download laporan komprehensif

## 🚀 Deployment

### Cloud (Streamlit Cloud) - RECOMMENDED
```bash
1. Upload ke GitHub
2. Deploy ke Streamlit Cloud
3. Akses via link online
```

### Lokal (Desktop)
```bash
pip install -r requirements.txt
streamlit run bot_analisis_rapor_digiwasda.py
```

Buka browser: `http://localhost:8501`

## 📦 Requirements

- Python 3.9+
- Streamlit 1.28.1
- Pandas 2.0.3
- Openpyxl 3.1.2
- Python-docx 0.8.11
- ReportLab 4.0.7

## 📄 Input Format

File Excel (.xlsx) dengan struktur:
- Sheet: "2. LAPORAN RAPOR"
- Sheet: "2.2 REKOM. PRIORITAS"
- Minimal 7 indikator: Literasi, Numerasi, Karakter, Pelatihan Guru, Kualitas Pembelajaran, Refleksi & Perbaikan, Kepemimpinan Instruksional

## 📊 Output

Bot menghasilkan:
- Ranking & klusterisasi sekolah
- RKT (12 bulan untuk Tier 1, 18 bulan untuk Tier 2)
- RKAS dengan breakdown budget detail
- Laporan komprehensif 40+ halaman dengan saran konkret per sekolah

## 🎓 Metodologi

**DIGIWASDA = Sistem Digital untuk Pengawas Sekolah Berdampak**

Fokus:
- Coaching Cycle 4-Fase untuk Kepala Sekolah
- Pendampingan Berbasis Data (PBD)
- Dukungan Kebijakan Merdeka Belajar
- Otomasi Analisis Rapor Pendidikan
- Peningkatan Kualitas Kepemimpinan Sekolah

## 👨‍💻 Developer

Alamsyah P. Skotang

## 📝 License

MIT License

## 📞 Support

Untuk bantuan atau pertanyaan, silakan hubungi developer.

---

**Bot Analisis Rapor Pendidikan | Metodologi DIGIWASDA | 2025**
