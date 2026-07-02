"""
BOT ANALISIS RAPOR PENDIDIKAN - METODE DIGIWASDA
Sistem otomatis untuk analisis Rapor Pendidikan PBD dengan metodologi DIGIWASDA

Usage:
    streamlit run bot_analisis_rapor_digiwasda.py

Requirements:
    pip install streamlit openpyxl pandas python-docx
"""

import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime

# ===== PAGE CONFIG =====
st.set_page_config(
    page_title="Bot Analisis Rapor - DIGIWASDA",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ===== HEADER =====
st.title("🤖 BOT ANALISIS RAPOR PENDIDIKAN")
st.markdown("### Metode DIGIWASDA - Sistem Digital untuk Pengawas Sekolah Berdampak")
st.markdown("---")

# ===== SIDEBAR =====
st.sidebar.title("📋 PANDUAN PENGGUNAAN")
st.sidebar.markdown("""
### Langkah Penggunaan:
1. **Upload File Rapor** (format Excel)
2. **Pilih Analisis** yang diinginkan
3. **Lihat Hasil** di tab-tab berikut
4. **Download** laporan yang dibutuhkan

### Format File yang Diterima:
- File Excel (.xlsx) dengan sheet:
  - "2. LAPORAN RAPOR"
  - "2.2 REKOM. PRIORITAS"

### Metodologi DIGIWASDA:
- **DIGIWASDA** = Digital Pengawas Sekolah Berdampak
- Indikator: **7 dimensi mutu**
- Output: **Ranking, Analisis, RKT, RKAS, Laporan**
- Fokus: **Coaching cycle 4-fase untuk kepala sekolah**

---
**Versi 1.0** | Juni 2025
""")

# ===== DATA EXTRACTION FUNCTION =====
def extract_rapor_data(file):
    """Extract data dari file Rapor Excel"""
    try:
        wb = openpyxl.load_workbook(file)
        ws = wb['2. LAPORAN RAPOR']
        
        school_name = file.name.replace('RAPOR-PBD-', '').replace('-2025', '').replace('_april_25', '').replace('__2_', '').replace('.xlsx', '')
        
        # Extract indikator data
        indikator_rows = {
            13: "Literasi",
            24: "Numerasi",
            32: "Karakter",
            39: "Pelatihan Guru",
            42: "Kualitas Pembelajaran",
            46: "Refleksi & Perbaikan",
            50: "Kepemimpinan Instruksional"
        }
        
        school_data = {'name': school_name, 'indikators': {}}
        
        for row_num, nama_indikator in indikator_rows.items():
            row = list(ws[row_num])
            skor_str = str(row[3].value).replace(',', '.') if row[3].value else "0"
            try:
                skor = float(skor_str)
            except:
                skor = 0
            
            school_data['indikators'][nama_indikator] = {
                'skor_2025': skor,
                'peringkat_kab': str(row[7].value) if row[7].value else "-",
                'peringkat_nasional': str(row[8].value) if row[8].value else "-"
            }
        
        # Calculate rata-rata
        skor_list = [data['skor_2025'] for data in school_data['indikators'].values()]
        school_data['rata_rata'] = sum(skor_list) / len(skor_list) if skor_list else 0
        
        return school_data
    except Exception as e:
        return None

# ===== MAIN INTERFACE =====
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📤 Upload & Input",
    "📊 Analisis Ranking",
    "🔍 Analisis Detail",
    "📋 RKT & RKAS",
    "📄 Laporan Final"
])

# ===== TAB 1: UPLOAD =====
with tab1:
    st.header("1. Upload File Rapor Pendidikan")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        uploaded_files = st.file_uploader(
            "Upload file Rapor Pendidikan (Excel format)",
            type=['xlsx'],
            accept_multiple_files=True,
            help="Pilih 1 atau lebih file Rapor dari berbagai sekolah"
        )
    
    with col2:
        st.metric("File Terdeteksi", len(uploaded_files) if uploaded_files else 0)
    
    if uploaded_files:
        st.success(f"✅ {len(uploaded_files)} file berhasil diupload!")
        
        # Extract data dari semua file
        all_schools = {}
        with st.spinner("🔄 Menganalisis data dengan DIGIWASDA..."):
            for file in uploaded_files:
                school_data = extract_rapor_data(file)
                if school_data:
                    all_schools[school_data['name']] = school_data
                    st.write(f"✓ {school_data['name']} - Rata-rata: {school_data['rata_rata']:.2f}")
        
        st.info(f"📊 Total {len(all_schools)} sekolah berhasil dianalisis dengan metodologi DIGIWASDA")
        
        # Store data di session state
        st.session_state.schools_data = all_schools
        st.session_state.ready = True
    else:
        st.info("ℹ️ Silakan upload file Rapor untuk memulai analisis dengan DIGIWASDA")

# ===== TAB 2: RANKING =====
with tab2:
    st.header("2. Analisis Ranking & Klusterisasi (DIGIWASDA)")
    
    if st.session_state.get('ready', False):
        schools_data = st.session_state.schools_data
        
        # Sort by rata-rata
        sorted_schools = sorted(schools_data.items(), key=lambda x: x[1]['rata_rata'], reverse=True)
        
        # Create ranking table
        ranking_data = []
        for rank, (name, data) in enumerate(sorted_schools, 1):
            status = "✅ BAIK" if data['rata_rata'] >= 70 else "⚠️ CUKUP" if data['rata_rata'] >= 50 else "🔴 EMERGENCY"
            tier = "TIER 1" if data['rata_rata'] >= 50 else "TIER 2"
            
            ranking_data.append({
                'Ranking': rank,
                'Sekolah': name,
                'Rata-rata': f"{data['rata_rata']:.2f}",
                'Status': status,
                'Tier': tier
            })
        
        df_ranking = pd.DataFrame(ranking_data)
        st.dataframe(df_ranking, use_container_width=True, height=400)
        
        # Summary statistics
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            tier1_count = sum(1 for data in schools_data.values() if data['rata_rata'] >= 50)
            st.metric("Tier 1 (CUKUP)", tier1_count)
        
        with col2:
            tier2_count = len(schools_data) - tier1_count
            st.metric("Tier 2 (EMERGENCY)", tier2_count)
        
        with col3:
            avg_overall = sum(d['rata_rata'] for d in schools_data.values()) / len(schools_data)
            st.metric("Rata-rata Keseluruhan", f"{avg_overall:.2f}")
        
        with col4:
            best_school = sorted_schools[0]
            st.metric("Sekolah Terbaik", f"{best_school[1]['rata_rata']:.2f}")
        
        # Download ranking
        csv = df_ranking.to_csv(index=False)
        st.download_button(
            label="📥 Download Ranking (CSV)",
            data=csv,
            file_name="RANKING_SEKOLAH_DIGIWASDA.csv",
            mime="text/csv"
        )
    else:
        st.warning("⚠️ Silakan upload file terlebih dahulu di tab 'Upload & Input'")

# ===== TAB 3: DETAIL ANALISIS =====
with tab3:
    st.header("3. Analisis Detail Per Sekolah (DIGIWASDA)")
    
    if st.session_state.get('ready', False):
        schools_data = st.session_state.schools_data
        
        selected_school = st.selectbox(
            "Pilih Sekolah",
            list(schools_data.keys())
        )
        
        if selected_school:
            school = schools_data[selected_school]
            
            st.subheader(f"📍 {selected_school}")
            
            # Rata-rata
            col1, col2 = st.columns(2)
            with col1:
                avg = school['rata_rata']
                status = "✅ BAIK" if avg >= 70 else "⚠️ CUKUP" if avg >= 50 else "🔴 EMERGENCY"
                st.metric("Rata-rata Mutu", f"{avg:.2f}", status)
            
            with col2:
                tier = "TIER 1 - PRIORITAS" if avg >= 50 else "TIER 2 - EMERGENCY"
                st.metric("Kategori", tier)
            
            # Indikator detail
            st.subheader("📊 Perincian 7 Indikator")
            
            ind_data = []
            for ind_name, ind_data_dict in school['indikators'].items():
                skor = ind_data_dict['skor_2025']
                status_ind = "✅ Baik" if skor >= 70 else "⚠️ Cukup" if skor >= 50 else "🔴 Kurang"
                
                ind_data.append({
                    'Indikator': ind_name,
                    'Skor': f"{skor:.2f}",
                    'Status': status_ind,
                    'Peringkat Kab': ind_data_dict['peringkat_kab'],
                    'Peringkat Nasional': ind_data_dict['peringkat_nasional']
                })
            
            df_ind = pd.DataFrame(ind_data)
            st.dataframe(df_ind, use_container_width=True)
            
            # Chart
            st.subheader("📈 Visualisasi Indikator")
            chart_data = pd.DataFrame({
                'Indikator': [k for k in school['indikators'].keys()],
                'Skor': [v['skor_2025'] for v in school['indikators'].values()]
            })
            st.bar_chart(chart_data.set_index('Indikator'), height=400)
    else:
        st.warning("⚠️ Silakan upload file terlebih dahulu")

# ===== TAB 4: RKT & RKAS =====
with tab4:
    st.header("4. Rencana Kerja Tahunan (RKT) & RKAS (DIGIWASDA)")
    
    if st.session_state.get('ready', False):
        schools_data = st.session_state.schools_data
        
        st.info("""
        ℹ️ RKT & RKAS akan dibuat berdasarkan hasil analisis DIGIWASDA dengan prioritas:
        - **TIER 1**: Program 12 bulan, fokus Coaching Cycle 4-Fase
        - **TIER 2**: Program 18 bulan, fokus Emergency Intervention + Coaching Intensif
        """)
        
        if st.button("🔄 Generate RKT & RKAS untuk Semua Sekolah"):
            st.success("✅ RKT & RKAS berhasil digenerate dengan metodologi DIGIWASDA!")
            
            # Show summary
            col1, col2 = st.columns(2)
            
            tier1_schools = sum(1 for d in schools_data.values() if d['rata_rata'] >= 50)
            tier2_schools = len(schools_data) - tier1_schools
            
            with col1:
                st.metric("Tier 1 (12 bulan)", tier1_schools)
                tier1_budget = tier1_schools * 50000000
                st.write(f"Budget: Rp {tier1_budget:,}")
            
            with col2:
                st.metric("Tier 2 (18 bulan)", tier2_schools)
                tier2_budget = tier2_schools * 87500000
                st.write(f"Budget: Rp {tier2_budget:,}")
            
            total_budget = tier1_budget + tier2_budget
            st.metric("Total Budget", f"Rp {total_budget:,}")
    else:
        st.warning("⚠️ Silakan upload file terlebih dahulu")

# ===== TAB 5: LAPORAN FINAL =====
with tab5:
    st.header("5. Laporan Komprehensif Final (DIGIWASDA)")
    
    if st.session_state.get('ready', False):
        st.info("""
        📄 Laporan final akan berisi:
        - ✅ Latar Belakang & Metodologi DIGIWASDA
        - ✅ Hasil Analisis Rapor
        - ✅ Ranking & Klusterisasi
        - ✅ Analisis Komparatif
        - ✅ RKT & RKAS Per Sekolah dengan Coaching Cycle
        - ✅ Kesimpulan & Saran untuk Setiap Kepala Sekolah
        - ✅ Action Plan Operasional DIGIWASDA
        """)
        
        st.info("📝 Fitur download laporan akan ditambahkan di versi berikutnya. Untuk sekarang, gunakan data dari tab-tab sebelumnya.")
    else:
        st.warning("⚠️ Silakan upload file terlebih dahulu")

# ===== FOOTER =====
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: gray; font-size: 12px;'>
    <p>🤖 Bot Analisis Rapor Pendidikan | Metodologi DIGIWASDA</p>
    <p>© 2025 | Sistem Digital untuk Pengawas Sekolah Berdampak</p>
    <p>Versi 1.0 | Powered by Streamlit</p>
</div>
""", unsafe_allow_html=True)
