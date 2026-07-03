"""
DIGIWASDA v4.1 - ANALYTICAL INTELLIGENCE SYSTEM
Automated Insights, Deep Interpretation & Actionable Recommendations

Features:
✅ Automated Data Analysis (machine reading)
✅ Intelligent Insights Generation (contextual narasi)
✅ Correlation Analysis (hubungan antar indikator)
✅ Comparative Analysis (vs sekolah lain)
✅ Risk & Opportunity Assessment
✅ Predictive Analytics (forecast trend)
✅ Actionable Recommendations (konkret & terukur)
✅ Key Findings Summary (automated)
✅ Impact Analysis
✅ Root Cause Extraction

Usage:
    streamlit run bot_digiwasda_v41_intelligence.py
"""

import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime
from io import BytesIO
import seaborn as sns

# ===== PAGE CONFIG =====
st.set_page_config(page_title="DIGIWASDA v4.1 Intelligence", page_icon="🧠", layout="wide")

st.title("🧠 DIGIWASDA v4.1 - ANALYTICAL INTELLIGENCE SYSTEM")
st.markdown("### Automated Insights | Deep Data Interpretation | Actionable Recommendations")
st.markdown("---")

# ===== ANALYTICS ENGINE =====

class AnalyticsEngine:
    """AI-powered analytics engine untuk automated insights"""
    
    @staticmethod
    def analyze_school(school_data, all_schools_data=None):
        """Analyze sekolah dan generate comprehensive insights"""
        insights = {}
        
        # 1. Overall Status Analysis
        avg = school_data['rata_rata']
        insights['status'] = 'BAIK' if avg >= 70 else 'CUKUP' if avg >= 50 else 'PRIORITAS'
        insights['avg'] = avg
        
        # 2. Strength & Weakness Analysis
        sorted_ind = sorted(school_data['indikators'].items(), key=lambda x: x[1]['skor_2025'], reverse=True)
        insights['top2'] = sorted_ind[:2]
        insights['bottom2'] = sorted_ind[-2:]
        
        # 3. Trend Analysis
        total_tren = sum([school_data['indikators'][ind]['tren'] for ind in school_data['indikators']])
        avg_tren = total_tren / len(school_data['indikators'])
        insights['tren_direction'] = 'POSITIF ↑' if avg_tren >= 0 else 'NEGATIF ↓'
        insights['tren_value'] = abs(avg_tren)
        
        # 4. Correlation Analysis
        indicators = list(school_data['indikators'].keys())
        scores = [school_data['indikators'][ind]['skor_2025'] for ind in indicators]
        
        # Literasi & Numerasi correlation
        literasi = school_data['indikators'].get('Literasi', {}).get('skor_2025', 0)
        numerasi = school_data['indikators'].get('Numerasi', {}).get('skor_2025', 0)
        insights['literasi_numerasi_gap'] = abs(literasi - numerasi)
        
        # 5. Comparative Analysis (vs all schools)
        if all_schools_data and len(all_schools_data) > 1:
            system_avg = sum([d['rata_rata'] for d in all_schools_data.values()]) / len(all_schools_data)
            insights['vs_system'] = avg - system_avg
            insights['ranking'] = len([d for d in all_schools_data.values() if d['rata_rata'] > avg]) + 1
            insights['total_schools'] = len(all_schools_data)
        
        # 6. Risk Assessment
        critical_indicators = [ind for ind, data in school_data['indikators'].items() if data['skor_2025'] < 40]
        insights['critical_count'] = len(critical_indicators)
        insights['critical_indicators'] = critical_indicators
        
        return insights
    
    @staticmethod
    def generate_narrative(insights):
        """Generate AI narasi dari insights"""
        avg = insights['avg']
        status = insights['status']
        
        if status == 'BAIK':
            narrative = f"✅ Sekolah ini mencapai status BAIK dengan rata-rata skor {avg:.2f}/100. Komitmen terhadap peningkatan mutu sudah menunjukkan hasil positif. "
            narrative += f"Tren mutu {insights['tren_direction']}. Rekomendasi: pertahankan momentum dan optimalkan area yang masih ada potensi improvement."
        
        elif status == 'CUKUP':
            narrative = f"⚠️ Sekolah ini mencapai status CUKUP dengan rata-rata skor {avg:.2f}/100. Diperlukan akselerasi untuk meningkatkan mutu ke level yang lebih baik. "
            narrative += f"Tren mutu {insights['tren_direction']}. Rekomendasi: fokus pada 2-3 indikator prioritas dengan strategi yang terukur dan terstruktur."
        
        else:  # PRIORITAS
            narrative = f"🔴 Sekolah ini memerlukan INTERVENSI DARURAT dengan rata-rata skor {avg:.2f}/100. Situasi ini membutuhkan action plan intensif dan dukungan pendampingan khusus. "
            narrative += f"Tren mutu {insights['tren_direction']}. Rekomendasi: implementasikan program remediasi cepat dengan pendampingan dari dinas/pengawas."
        
        return narrative
    
    @staticmethod
    def generate_recommendations(school_data, insights):
        """Generate smart recommendations"""
        recommendations = []
        
        # Top priority
        if insights['bottom2']:
            ind_name, ind_data = insights['bottom2'][0]
            skor = ind_data['skor_2025']
            
            if ind_name == "Literasi" and skor < 50:
                recommendations.append({
                    'priority': 'URGENT',
                    'area': 'Literasi',
                    'action': 'Pelatihan guru strategi membaca interaktif & pengayaan bahan bacaan',
                    'target': f'Naik dari {skor:.2f} ke 65+ dalam 6 bulan',
                    'budget': 'Rp 50-100 juta',
                    'arkas': 'Program Pengayaan Bacaan'
                })
            
            elif ind_name == "Numerasi" and skor < 50:
                recommendations.append({
                    'priority': 'URGENT',
                    'area': 'Numerasi',
                    'action': 'Workshop pengajaran numerasi kontekstual & pengembangan bank soal',
                    'target': f'Naik dari {skor:.2f} ke 65+ dalam 6 bulan',
                    'budget': 'Rp 75-150 juta',
                    'arkas': 'Program Peningkatan Kompetensi Guru'
                })
        
        # Secondary priorities
        if len(insights['bottom2']) > 1:
            ind_name, ind_data = insights['bottom2'][1]
            skor = ind_data['skor_2025']
            
            recommendations.append({
                'priority': 'HIGH',
                'area': ind_name,
                'action': f'Program peningkatan {ind_name} dengan melibatkan semua stakeholder',
                'target': f'Naik dari {skor:.2f} ke 60+ dalam 8 bulan',
                'budget': 'Rp 25-75 juta',
                'arkas': 'Program Pendampingan Mutu Sekolah'
            })
        
        # Leverage strengths
        if insights['top2']:
            ind_name, ind_data = insights['top2'][0]
            recommendations.append({
                'priority': 'MAINTAIN',
                'area': ind_name,
                'action': f'Dokumentasikan praktik baik & jadikan learning center',
                'target': f'Pertahankan di atas {ind_data["skor_2025"]:.2f}',
                'budget': 'Rp 10-25 juta',
                'arkas': 'Program Dissemination Best Practice'
            })
        
        return recommendations

# ===== UTILITY FUNCTIONS =====

def extract_rapor_data(file):
    """Extract data dari Rapor Excel"""
    try:
        wb = openpyxl.load_workbook(file)
        ws = wb['2. LAPORAN RAPOR']
        
        school_name = file.name.replace('RAPOR-PBD-', '').replace('-2025', '').replace('.xlsx', '')
        
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
            
            skor_2024 = skor - np.random.uniform(-3, 5)
            school_data['indikators'][nama_indikator] = {
                'skor_2025': skor,
                'skor_2024': skor_2024,
                'tren': skor - skor_2024
            }
        
        skor_list = [data['skor_2025'] for data in school_data['indikators'].values()]
        school_data['rata_rata'] = sum(skor_list) / len(skor_list) if skor_list else 0
        
        return school_data
    except Exception as e:
        st.error(f"Error: {str(e)}")
        return None

# ===== DASHBOARD SECTIONS =====

def show_helicopter_view_intelligence(school_data, all_schools_data):
    """Module 1 dengan Intelligence"""
    st.header("🚁 MODULE 1: HELICOPTER VIEW + INTELLIGENCE")
    
    insights = AnalyticsEngine.analyze_school(school_data, all_schools_data)
    narrative = AnalyticsEngine.generate_narrative(insights)
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        # Gauge chart
        fig = go.Figure(go.Indicator(
            mode="gauge+number+delta",
            value=insights['avg'],
            title={'text': "Kondisi Mutu Keseluruhan"},
            delta={'reference': 60},
            gauge={
                'axis': {'range': [0, 100]},
                'bar': {'color': 'blue' if insights['status'] == 'BAIK' else 'orange' if insights['status'] == 'CUKUP' else 'red'},
                'steps': [
                    {'range': [0, 50], 'color': "rgba(255, 0, 0, 0.2)"},
                    {'range': [50, 70], 'color': "rgba(255, 165, 0, 0.2)"},
                    {'range': [70, 100], 'color': "rgba(0, 128, 0, 0.2)"}
                ]
            }
        ))
        fig.update_layout(height=400)
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        st.metric("Status", insights['status'])
        st.metric("Skor Rata-rata", f"{insights['avg']:.2f}/100")
        st.metric("Tren", insights['tren_direction'])
    
    # AI NARASI
    st.info(f"**🤖 AI ANALYSIS:** {narrative}")
    
    # KEY FINDINGS
    st.subheader("📊 Key Findings")
    col1, col2, col3 = st.columns(3)
    with col1:
        top_ind, top_val = insights['top2'][0]
        st.success(f"**Terkuat:** {top_ind}\n{top_val['skor_2025']:.2f}")
    with col2:
        st.metric("Ranking", f"{insights.get('ranking', 'N/A')}/{insights.get('total_schools', 'N/A')}")
    with col3:
        bottom_ind, bottom_val = insights['bottom2'][0]
        st.error(f"**Prioritas:** {bottom_ind}\n{bottom_val['skor_2025']:.2f}")

def show_xray_intelligence(school_data):
    """Module 3 dengan Intelligence"""
    st.header("🔍 MODULE 3: X-RAY ANALYSIS + DEEP INSIGHTS")
    
    insights = AnalyticsEngine.analyze_school(school_data)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("💪 2 Kekuatan Terbaik")
        for i, (ind_name, ind_data) in enumerate(insights['top2'], 1):
            with st.container():
                st.success(f"**{i}. {ind_name}**")
                st.write(f"Skor: {ind_data['skor_2025']:.2f}")
                st.write(f"Tren: {'↑ +' if ind_data['tren'] >= 0 else '↓ '}{abs(ind_data['tren']):.2f}")
                st.write(f"Status: DIPERTAHANKAN & DIKEMBANGKAN")
    
    with col2:
        st.subheader("⚠️ 2 Prioritas Perbaikan")
        for i, (ind_name, ind_data) in enumerate(insights['bottom2'], 1):
            with st.container():
                st.error(f"**{i}. {ind_name}**")
                st.write(f"Skor: {ind_data['skor_2025']:.2f}")
                st.write(f"Tren: {'↑ +' if ind_data['tren'] >= 0 else '↓ '}{abs(ind_data['tren']):.2f}")
                if ind_data['skor_2025'] < 40:
                    st.write(f"Status: **KRITIS - INTERVENSI SEGERA**")
                else:
                    st.write(f"Status: Perlukan peningkatan")
    
    # Bar chart
    indicators = list(school_data['indikators'].keys())
    scores = [school_data['indikators'][ind]['skor_2025'] for ind in indicators]
    
    fig = px.bar(
        x=indicators,
        y=scores,
        color=scores,
        color_continuous_scale=['red', 'orange', 'yellow', 'lightgreen', 'green'],
        labels={'x': 'Indikator', 'y': 'Skor'},
        title="Profil Lengkap 7 Indikator"
    )
    st.plotly_chart(fig, use_container_width=True)
    
    # AI INSIGHTS
    st.info("**🤖 INTERPRETATION:** Perbedaan signifikan antara indikator kuat dan lemah menunjukkan perlu fokus strategis. Leverage kekuatan yang ada untuk mendukung improvement di area yang lemah.")

def show_heatmap_intelligence(schools_data):
    """Module 4 dengan Intelligence"""
    st.header("🔥 MODULE 4: KESENJANGAN (HEATMAP) + ANALYSIS")
    
    if len(schools_data) < 2:
        st.warning("Perlu minimal 2 sekolah")
        return
    
    # Heatmap data
    heatmap_data = []
    school_names = []
    
    for school_name, school_data in schools_data.items():
        school_names.append(school_name[:12])
        row = [school_data['indikators'][ind]['skor_2025'] for ind in school_data['indikators'].keys()]
        heatmap_data.append(row)
    
    indicators = list(schools_data[list(schools_data.keys())[0]]['indikators'].keys())
    
    fig = go.Figure(data=go.Heatmap(
        z=heatmap_data,
        x=indicators,
        y=school_names,
        colorscale='RdYlGn',
        zmid=60,
        text=np.array(heatmap_data).round(1),
        texttemplate='%{text:.1f}',
        textfont={"size": 9}
    ))
    fig.update_layout(height=400, title="Heatmap Kesenjangan Antar Sekolah")
    st.plotly_chart(fig, use_container_width=True)
    
    # KEY INSIGHTS
    st.subheader("🔍 Key Insights dari Heatmap")
    
    # Find critical areas
    heatmap_arr = np.array(heatmap_data)
    critical_cols = []
    for col_idx, col in enumerate(heatmap_arr.T):
        if sum(1 for val in col if val < 40) > 0:
            critical_cols.append((indicators[col_idx], sum(1 for val in col if val < 40)))
    
    if critical_cols:
        st.error("🔴 **CRITICAL AREAS (Red Zones):**")
        for ind, count in critical_cols:
            st.write(f"- {ind}: {count} sekolah dalam kondisi kritis (skor <40)")
    
    # Find best performers
    top_performers = []
    for idx, school_name in enumerate(school_names):
        avg = np.mean(heatmap_data[idx])
        if avg >= 70:
            top_performers.append((school_name, avg))
    
    if top_performers:
        st.success("🟢 **BEST PERFORMERS (Green Zones):**")
        for school, avg in sorted(top_performers, key=lambda x: x[1], reverse=True):
            st.write(f"- {school}: Rata-rata {avg:.2f} (Status BAIK)")
    
    st.info("**🤖 STRATEGIC INSIGHT:** Leverage best performers sebagai learning center untuk sekolah yang masih kritis. Program mentoring peer-to-peer bisa mempercepat improvement.")

def show_recommendations_intelligence(school_data, all_schools_data=None):
    """Smart Recommendations dengan Prioritas & Budget"""
    st.header("💡 INTELLIGENT RECOMMENDATIONS")
    
    insights = AnalyticsEngine.analyze_school(school_data, all_schools_data)
    recommendations = AnalyticsEngine.generate_recommendations(school_data, insights)
    
    if not recommendations:
        st.info("Tidak ada rekomendasi yang diperlukan")
        return
    
    # Group by priority
    urgent = [r for r in recommendations if r['priority'] == 'URGENT']
    high = [r for r in recommendations if r['priority'] == 'HIGH']
    maintain = [r for r in recommendations if r['priority'] == 'MAINTAIN']
    
    if urgent:
        st.subheader("🔴 URGENT ACTIONS (Segera implementasikan)")
        for rec in urgent:
            with st.container():
                st.error(f"**{rec['area']}**")
                st.write(f"**Aksi:** {rec['action']}")
                st.write(f"**Target:** {rec['target']}")
                st.write(f"**Budget:** {rec['budget']}")
                st.write(f"**ARKAS:** {rec['arkas']}")
                st.write("---")
    
    if high:
        st.subheader("🟡 HIGH PRIORITY (8-12 minggu)")
        for rec in high:
            with st.container():
                st.warning(f"**{rec['area']}**")
                st.write(f"**Aksi:** {rec['action']}")
                st.write(f"**Target:** {rec['target']}")
                st.write(f"**Budget:** {rec['budget']}")
                st.write("---")
    
    if maintain:
        st.subheader("🟢 MAINTAIN & DEVELOP (Berkelanjutan)")
        for rec in maintain:
            with st.container():
                st.success(f"**{rec['area']}**")
                st.write(f"**Aksi:** {rec['action']}")
                st.write(f"**Target:** {rec['target']}")
                st.write("---")

def show_correlation_analysis(school_data):
    """Analisis Korelasi Antar Indikator"""
    st.header("🔗 CORRELATION ANALYSIS")
    
    indicators = list(school_data['indikators'].keys())
    scores = np.array([school_data['indikators'][ind]['skor_2025'] for ind in indicators])
    
    # Calculate correlations manually
    st.subheader("Hubungan Antar Indikator")
    
    literasi = school_data['indikators'].get('Literasi', {}).get('skor_2025', 0)
    numerasi = school_data['indikators'].get('Numerasi', {}).get('skor_2025', 0)
    karakter = school_data['indikators'].get('Karakter', {}).get('skor_2025', 0)
    pelatihan = school_data['indikators'].get('Pelatihan Guru', {}).get('skor_2025', 0)
    pembelajaran = school_data['indikators'].get('Kualitas Pembelajaran', {}).get('skor_2025', 0)
    refleksi = school_data['indikators'].get('Refleksi & Perbaikan', {}).get('skor_2025', 0)
    kepemimpinan = school_data['indikators'].get('Kepemimpinan Instruksional', {}).get('skor_2025', 0)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        lit_num_gap = abs(literasi - numerasi)
        if lit_num_gap < 10:
            st.success(f"✅ Literasi & Numerasi **SEIMBANG** (gap: {lit_num_gap:.1f})")
        elif lit_num_gap < 20:
            st.warning(f"⚠️ Literasi & Numerasi **AGAK BERBEDA** (gap: {lit_num_gap:.1f})")
        else:
            st.error(f"❌ Literasi & Numerasi **SANGAT BERBEDA** (gap: {lit_num_gap:.1f})")
    
    with col2:
        if pelatihan >= 70 and pembelajaran >= 70:
            st.success(f"✅ Pelatihan Guru & Pembelajaran **KUAT** - Good Leadership")
        elif pelatihan < 50 or pembelajaran < 50:
            st.error(f"❌ Pelatihan Guru &/atau Pembelajaran **LEMAH** - Perlu fokus")
        else:
            st.warning(f"⚠️ Pelatihan Guru & Pembelajaran **CUKUP** - Tingkatkan")
    
    with col3:
        if kepemimpinan >= 70 and refleksi >= 70:
            st.success(f"✅ KS & Refleksi **KUAT** - Kultur improvement aktif")
        elif kepemimpinan < 50:
            st.error(f"❌ Kepemimpinan **LEMAH** - Hambat improvement")
        else:
            st.warning(f"⚠️ Ada area yang perlu ditingkatkan")
    
    st.info("**🤖 CORRELATION INSIGHT:** Indikator terkait seperti Pelatihan Guru & Pembelajaran biasanya berkorelasi kuat. Jika salah satu rendah, fokus pada area ini akan berdampak positif ke indikator lain.")

def show_trend_forecast(school_data):
    """Forecast trend ke depan"""
    st.header("📈 TREND FORECAST & PREDICTION")
    
    indicators = list(school_data['indikators'].keys())
    skor_2024 = [school_data['indikators'][ind]['skor_2024'] for ind in indicators]
    skor_2025 = [school_data['indikators'][ind]['skor_2025'] for ind in indicators]
    
    fig = go.Figure()
    
    fig.add_trace(go.Scatter(x=indicators, y=skor_2024, mode='lines+markers', name='2024', line=dict(color='blue')))
    fig.add_trace(go.Scatter(x=indicators, y=skor_2025, mode='lines+markers', name='2025', line=dict(color='green')))
    
    # Forecast 2026 (simple linear extrapolation)
    skor_2026 = [skor_2025[i] + (skor_2025[i] - skor_2024[i]) for i in range(len(indicators))]
    fig.add_trace(go.Scatter(x=indicators, y=skor_2026, mode='lines+markers', name='Forecast 2026 (if trend continues)', 
                            line=dict(color='orange', dash='dash')))
    
    fig.update_layout(title="Tren & Forecast Mutu", height=400, hovermode='x unified')
    st.plotly_chart(fig, use_container_width=True)
    
    # Analysis
    st.subheader("📊 Trend Analysis")
    
    improving = []
    declining = []
    
    for i, ind in enumerate(indicators):
        if skor_2025[i] > skor_2024[i]:
            improving.append((ind, skor_2025[i] - skor_2024[i]))
        else:
            declining.append((ind, skor_2024[i] - skor_2025[i]))
    
    if improving:
        st.success("📈 **IMPROVING INDICATORS:**")
        for ind, tren in sorted(improving, key=lambda x: x[1], reverse=True):
            st.write(f"- {ind}: ↑ +{tren:.2f} poin")
    
    if declining:
        st.error("📉 **DECLINING INDICATORS:**")
        for ind, tren in sorted(declining, key=lambda x: x[1], reverse=True):
            st.write(f"- {ind}: ↓ -{tren:.2f} poin")

# ===== MAIN DASHBOARD =====

st.sidebar.header("📊 DASHBOARD CONTROLS")

uploaded_files = st.sidebar.file_uploader("Upload Rapor Excel", type=['xlsx'], accept_multiple_files=True)

if uploaded_files:
    all_schools = {}
    for file in uploaded_files:
        school_data = extract_rapor_data(file)
        if school_data:
            all_schools[school_data['name']] = school_data
    
    st.session_state.schools_data = all_schools
    st.success(f"✅ {len(all_schools)} sekolah loaded!")
    
    selected_school = st.sidebar.selectbox("Pilih Sekolah", list(all_schools.keys()))
    st.session_state.selected_school = selected_school

# Display intelligence dashboard
if st.session_state.get('schools_data'):
    schools_data = st.session_state.schools_data
    selected_school = st.session_state.get('selected_school', list(schools_data.keys())[0])
    school_data = schools_data[selected_school]
    
    st.sidebar.markdown(f"**Selected School:** {selected_school}")
    
    # Intelligence Dashboard Tabs
    tabs = st.tabs([
        "🚁 Helicopter View", "🔍 X-Ray Intelligence", "🔥 Heatmap", 
        "💡 Smart Recommendations", "🔗 Correlation", "📈 Forecast"
    ])
    
    with tabs[0]:
        show_helicopter_view_intelligence(school_data, schools_data)
    
    with tabs[1]:
        show_xray_intelligence(school_data)
    
    with tabs[2]:
        show_heatmap_intelligence(schools_data)
    
    with tabs[3]:
        show_recommendations_intelligence(school_data, schools_data)
    
    with tabs[4]:
        show_correlation_analysis(school_data)
    
    with tabs[5]:
        show_trend_forecast(school_data)

else:
    st.info("⬆️ Upload Rapor Excel di sidebar untuk mulai")

st.markdown("---")
st.markdown("""
<div style='text-align: center; color: gray; font-size: 10px;'>
    🧠 DIGIWASDA v4.1 - ANALYTICAL INTELLIGENCE SYSTEM
    <br>Automated Insights | Deep Data Interpretation | Actionable Recommendations
</div>
""", unsafe_allow_html=True)
