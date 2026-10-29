import streamlit as st
import pandas as pd
import plotly.express as px
import io
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import re
import json
import os
from data_parser import (
    load_data,
    generate_official_leave_excel,
    synthesize_official_summary,
    parse_official_report_data
)

# Page configuration

# --- Custom CSS for Authentic "Notebook" Stationery Style ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Caveat:wght@500;600;700&family=Kalam:wght@400;700&family=Mali:ital,wght@0,400;0,500;0,600;0,700;1,400&family=Sarabun:wght@300;400;500;600;700&family=Prompt:wght@400;500;600;700&display=swap');

    /* Global Study Desk & Graph Paper Background */
    html, body, [class*="css"], .stApp {
        font-family: 'Sarabun', 'Prompt', sans-serif !important;
        background-color: #fcfbf7 !important;
        background-image:
            linear-gradient(to right, rgba(229, 231, 235, 0.55) 1px, transparent 1px),
            linear-gradient(to bottom, rgba(229, 231, 235, 0.55) 1px, transparent 1px) !important;
        background-size: 24px 24px !important;
        color: #1e293b !important;
    }

    /* Sidebar: Lined Notebook Margin & Paper */
    [data-testid="stSidebar"] {
        background-color: #f7f4eb !important;
        background-image: repeating-linear-gradient(
            transparent,
            transparent 27px,
            #e5e7eb 27px,
            #e5e7eb 28px
        ) !important;
        border-right: 3px double #fca5a5 !important;
        box-shadow: 4px 0 15px rgba(0, 0, 0, 0.04) !important;
    }
    [data-testid="stSidebarNav"] span {
        font-family: 'Mali', 'Prompt', sans-serif !important;
        font-weight: 600 !important;
        color: #1e293b !important;
    }

    h1, h2, h3, h4, h5, h6 {
        font-family: 'Prompt', 'Sarabun', sans-serif !important;
        font-weight: 700 !important;
        color: #1e293b !important;
    }

    .stCaption, caption, small {
        font-family: 'Caveat', 'Kalam', 'Mali', cursive !important;
        font-size: 1.05rem !important;
        color: #475569 !important;
    }

    .stMarkdown, .stTable, .stDataFrame, td, th {
        font-family: 'Sarabun', 'Prompt', sans-serif !important;
        line-height: 1.6 !important;
    }

    /* Hide Streamlit branding */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}

    /* Header Banner: Ruled Notebook Sheet with Pink Margin & Washi Tape */
    .leave-header-container {
        background-color: #fcfbf7;
        background-image:
            linear-gradient(90deg, transparent 52px, #fca5a5 52px, #fca5a5 54px, transparent 54px, transparent 58px, rgba(252, 165, 165, 0.45) 58px, rgba(252, 165, 165, 0.45) 59px, transparent 59px),
            repeating-linear-gradient(transparent, transparent 30px, #e5e7eb 30px, #e5e7eb 31px);
        border: 1.5px solid #d1d5db;
        border-radius: 6px 16px 16px 6px;
        padding: 30px 30px 26px 76px;
        margin-top: 14px;
        margin-bottom: 26px;
        position: relative;
        box-shadow:
            0 1px 1px rgba(0,0,0,0.08),
            0 8px 0 -4px #f8f6f0,
            0 8px 1px -3px rgba(0,0,0,0.12),
            0 16px 0 -8px #f3f0e6,
            0 16px 1px -7px rgba(0,0,0,0.12),
            0 14px 28px rgba(15, 23, 42, 0.07);
    }

    /* Washi Tape Top Accent */
    .leave-header-container::before {
        content: "";
        position: absolute;
        top: -12px;
        left: 50%;
        transform: translateX(-50%) rotate(1.2deg);
        width: 160px;
        height: 26px;
        background: rgba(252, 231, 243, 0.9);
        border-left: 2px dotted rgba(236, 72, 153, 0.5);
        border-right: 2px dotted rgba(236, 72, 153, 0.5);
        box-shadow: 0 2px 5px rgba(0, 0, 0, 0.08);
        z-index: 2;
    }

    /* Binder Holes on Left Margin */
    .leave-header-container::after {
        content: "";
        position: absolute;
        top: 22px;
        bottom: 22px;
        left: 18px;
        width: 16px;
        background-image: radial-gradient(circle, #e5e7eb 5.5px, rgba(0,0,0,0.12) 6.5px, transparent 7.5px);
        background-size: 16px 52px;
        background-repeat: repeat-y;
        pointer-events: none;
    }

    .leave-pill-badge {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        padding: 4px 14px;
        background: #fef9c3;
        border: 1px solid #fde047;
        border-radius: 3px;
        font-family: 'Caveat', 'Kalam', 'Mali', cursive !important;
        font-size: 1.08rem;
        font-weight: 700;
        color: #854d0e;
        box-shadow: 2px 2px 6px rgba(0, 0, 0, 0.07);
        transform: rotate(-1deg);
        margin-bottom: 10px;
    }

    .leave-header-title {
        font-size: 1.85rem !important;
        font-weight: 700 !important;
        color: #1e293b !important;
        margin: 0 0 6px 0 !important;
    }

    .leave-header-sub {
        font-family: 'Mali', 'Sarabun', sans-serif !important;
        font-size: 0.96rem;
        color: #475569;
        margin: 0;
    }

    /* Official Report Header Box: Ruled Stationery Card with Washi Tape */
    .official-report-card {
        background-color: #fcfbf7;
        background-image:
            linear-gradient(90deg, transparent 36px, #fca5a5 36px, #fca5a5 38px, transparent 38px),
            repeating-linear-gradient(transparent, transparent 29px, #e5e7eb 29px, #e5e7eb 30px);
        border: 1.5px solid #cbd5e1;
        border-radius: 6px 14px 14px 6px;
        padding: 26px 26px 22px 54px;
        margin-top: 10px;
        margin-bottom: 22px;
        box-shadow: 3px 6px 16px rgba(15, 23, 42, 0.06);
        text-align: center;
        position: relative;
    }
    .official-report-card::before {
        content: "";
        position: absolute;
        top: -10px;
        right: 36px;
        width: 110px;
        height: 22px;
        background: rgba(254, 249, 195, 0.9);
        border-left: 2px dotted rgba(234, 179, 8, 0.45);
        border-right: 2px dotted rgba(234, 179, 8, 0.45);
        transform: rotate(2deg);
        box-shadow: 0 2px 4px rgba(0,0,0,0.06);
    }
    .official-org-title {
        font-size: 1.5rem;
        font-weight: 700;
        color: #1e293b;
        margin-bottom: 4px;
    }
    .official-main-title {
        font-size: 1.25rem;
        font-weight: 600;
        color: #be123c;
        margin-bottom: 4px;
    }
    .official-date-badge {
        display: inline-block;
        background: #fef9c3;
        color: #854d0e;
        border: 1px solid #fde047;
        padding: 5px 18px;
        border-radius: 3px;
        font-family: 'Mali', 'Sarabun', sans-serif !important;
        font-weight: 700;
        font-size: 0.92rem;
        margin: 8px 0 12px 0;
        box-shadow: 2px 3px 6px rgba(0,0,0,0.07);
        transform: rotate(-0.6deg);
    }
    .official-format-sub {
        font-family: 'Mali', 'Sarabun', sans-serif !important;
        font-size: 0.88rem;
        color: #475569;
        text-align: left;
        border-top: 1.5px dashed #fca5a5;
        padding-top: 8px;
        margin-top: 6px;
    }

    /* Tabs: Notebook Divider Index Tabs */
    div[data-testid="stTabs"] {
        margin-top: 8px;
        margin-bottom: 24px;
    }

    div[data-baseweb="tab-list"] {
        gap: 8px !important;
        background-color: transparent !important;
        padding: 6px 4px 0 4px !important;
        border-bottom: 2px solid #fca5a5 !important;
        border-radius: 0 !important;
        display: flex !important;
        flex-wrap: wrap !important;
    }

    button[data-baseweb="tab"] {
        border-radius: 10px 10px 0 0 !important;
        padding: 9px 18px !important;
        font-family: 'Mali', 'Prompt', sans-serif !important;
        font-weight: 600 !important;
        font-size: 0.92rem !important;
        color: #475569 !important;
        background-color: #fcfbf7 !important;
        border: 1.5px solid #d1d5db !important;
        border-bottom: none !important;
        box-shadow: 0 -2px 5px rgba(0,0,0,0.03) !important;
        transition: all 0.2s ease !important;
    }

    button[data-baseweb="tab"]:nth-child(odd) {
        background-color: #fefce8 !important;
    }

    button[data-baseweb="tab"]:nth-child(even) {
        background-color: #fdf2f8 !important;
    }

    button[data-baseweb="tab"]:hover {
        color: #e11d48 !important;
        transform: translateY(-2px) !important;
    }

    button[data-baseweb="tab"][aria-selected="true"] {
        background: #fef9c3 !important;
        color: #1e293b !important;
        font-weight: 700 !important;
        border: 2px solid #fca5a5 !important;
        border-bottom: none !important;
        box-shadow: 0 -3px 10px rgba(252, 165, 165, 0.25) !important;
        transform: translateY(-2px) !important;
    }

    div[data-baseweb="tab-border"], div[data-baseweb="tab-highlight"] {
        display: none !important;
    }

    /* Style Metric Cards: Pastel Sticky Notes with Washi Tape */
    div[data-testid="stMetric"], div[data-testid="metric-container"] {
        background: #fef9c3 !important;
        border: 1px solid #fde047 !important;
        border-radius: 3px 3px 18px 3px !important;
        padding: 20px 18px 16px 18px !important;
        box-shadow: 3px 5px 12px rgba(0, 0, 0, 0.08) !important;
        position: relative !important;
        transition: all 0.25s ease !important;
        margin-top: 8px !important;
    }
    div[data-testid="stMetric"]::before, div[data-testid="metric-container"]::before {
        content: "";
        position: absolute;
        top: -9px;
        left: 50%;
        transform: translateX(-50%) rotate(-1.5deg);
        width: 64px;
        height: 18px;
        background: rgba(255, 255, 255, 0.65);
        border: 1px solid rgba(203, 213, 225, 0.6);
        box-shadow: 0 1px 2px rgba(0,0,0,0.04);
    }
    div[data-testid="stHorizontalBlock"] > div:nth-of-type(even) div[data-testid="stMetric"],
    div[data-testid="stHorizontalBlock"] > div:nth-of-type(even) div[data-testid="metric-container"] {
        background: #fce7f3 !important;
        border-color: #f9a8d4 !important;
        transform: rotate(0.7deg) !important;
    }
    div[data-testid="stHorizontalBlock"] > div:nth-of-type(odd) div[data-testid="stMetric"],
    div[data-testid="stHorizontalBlock"] > div:nth-of-type(odd) div[data-testid="metric-container"] {
        transform: rotate(-0.6deg) !important;
    }
    div[data-testid="stMetric"]:hover, div[data-testid="metric-container"]:hover {
        transform: translateY(-5px) rotate(0deg) !important;
        box-shadow: 5px 12px 22px rgba(0, 0, 0, 0.12) !important;
    }
    div[data-testid="stMetricLabel"], div[data-testid="metric-container"] label {
        font-family: 'Mali', 'Prompt', sans-serif !important;
        font-size: 0.92rem !important;
        font-weight: 700 !important;
        color: #334155 !important;
    }
    div[data-testid="stMetricValue"], div[data-testid="metric-container"] div[data-testid="stMetricValue"] {
        font-family: 'Caveat', 'Kalam', 'Prompt', sans-serif !important;
        font-size: 2.1rem !important;
        font-weight: 700 !important;
        color: #1e293b !important;
    }

    /* Table styling improvements */
    .stDataFrame, div[data-testid="stDataFrame"] {
        border-radius: 8px !important;
        overflow: hidden !important;
        border: 1.5px solid #cbd5e1 !important;
        border-left: 4px solid #fca5a5 !important;
        background: #fcfbf7 !important;
        box-shadow: 3px 5px 14px rgba(15, 23, 42, 0.05) !important;
    }

    /* Buttons: Stationery Stamp Style */
    div.stButton > button, div.stDownloadButton > button {
        border-radius: 6px !important;
        font-family: 'Mali', 'Prompt', sans-serif !important;
        font-weight: 700 !important;
        padding: 10px 24px !important;
        background-color: #fef9c3 !important;
        color: #1e293b !important;
        border: 2px solid #1e293b !important;
        box-shadow: 3px 3px 0px #1e293b !important;
        transition: all 0.18s ease !important;
    }
    div.stButton > button:hover, div.stDownloadButton > button:hover {
        background-color: #fce7f3 !important;
        border-color: #e11d48 !important;
        color: #1e293b !important;
        transform: translate(-1px, -1px) !important;
        box-shadow: 5px 5px 0px #e11d48 !important;
    }
</style>
""", unsafe_allow_html=True)

# --- Header Section ---
st.markdown("""
<div class="leave-header-container">
    <div class="leave-pill-badge">🏖️ LEAVE TRACKER & OFFICIAL REPORTS</div>
    <div class="leave-header-title">ระบบตรวจสอบและรายงานข้อมูลการลาบุคลากร</div>
    <p class="leave-header-sub">รองรับทั้งแบบฟอร์มรายงานทางการกรมเจ้าท่า (13 คอลัมน์) และการวิเคราะห์สถิติวันลาประจำปี</p>
</div>
""", unsafe_allow_html=True)

# --- File Uploader ---
st.sidebar.markdown("### 📂 นำเข้าข้อมูล")
uploaded_file = st.sidebar.file_uploader(
    "อัปโหลดไฟล์ตารางวันลา (.xlsx, .csv)", 
    type=["xlsx", "csv"], 
    help="รองรับรายงานทางการกรมเจ้าท่า (เช่น ข้าราชการ.xlsx / .csv) และตารางวันลาประจำปี (.xlsx)"
)

@st.cache_data(show_spinner=False)
def get_data(file_source, _force_refresh=1):
    return load_data(file_source)

# Load canonical roster if available for name ordering and synthesis
roster_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "canonical_personnel.json")
CANONICAL_ROSTER = []
if os.path.exists(roster_path):
    try:
        with open(roster_path, 'r', encoding='utf-8') as rf:
            CANONICAL_ROSTER = json.load(rf)
    except Exception:
        pass

if uploaded_file is not None:
    with st.spinner('กำลังประมวลผลข้อมูล...'):
        df = get_data(uploaded_file, _force_refresh=4)
        if df is not None and not df.empty:
            st.session_state.leave_df = df
elif 'leave_df' in st.session_state and st.session_state.leave_df is not None:
    df = st.session_state.leave_df
    st.sidebar.info("💡 กำลังใช้ข้อมูลตารางวันลาที่นำเข้าจากระบบ")
else:
    df = None

if df is not None:
    if df.empty:
        st.error("❌ ไม่พบข้อมูลการลาในไฟล์ที่อัปโหลด หรือรูปแบบไฟล์ไม่ถูกต้อง")
    else:
        # Determine if official summary table exists on dataframe or synthesize it
        official_summary_df = df.attrs.get('official_summary')
        header_meta = df.attrs.get('header_meta') or {
            'org': 'กรมเจ้าท่า',
            'title': 'รายงานการมาปฏิบัติราชการของข้าราชการพลเรือนสามัญ',
            'date_range': 'ประจำเดือนปัจจุบัน',
            'format': 'รูปแบบการออกรายงาน : โครงสร้างตามกฎหมาย - ทั้งส่วนราชการ - กรมเจ้าท่า - กระทรวงคมนาคม'
        }

        # If not loaded from official report, synthesize official summary from details
        if official_summary_df is None or official_summary_df.empty:
            avail_months = df['Month'].dropna().unique().tolist() if 'Month' in df.columns else []
            sel_month = avail_months[0] if avail_months else 'ตุลาคม'
            official_summary_df = synthesize_official_summary(df, roster=CANONICAL_ROSTER, month_name=sel_month)

        # --- Sidebar Filters ---
        st.sidebar.markdown("---")
        st.sidebar.markdown("### 🎯 ตัวกรองข้อมูล")
        
        b_years = sorted(df['BudgetYear'].dropna().unique()) if 'BudgetYear' in df.columns else ['2569']
        budget_year = st.sidebar.multiselect("📆 ปีงบประมาณ", options=b_years, default=b_years)
        
        m_opts = df['Month'].dropna().unique() if 'Month' in df.columns else ['ตุลาคม']
        months = st.sidebar.multiselect("📅 เดือน", options=m_opts, default=m_opts)
        
        t_opts = df['Type'].dropna().unique() if 'Type' in df.columns else ['ลาป่วย', 'ลาพักผ่อน', 'ลากิจ', 'สาย']
        leave_types = st.sidebar.multiselect("🏷️ ประเภทการลา", options=t_opts, default=t_opts)

        # Filter data
        filtered_df = df.copy()
        if 'BudgetYear' in filtered_df.columns and budget_year:
            filtered_df = filtered_df[filtered_df['BudgetYear'].isin(budget_year)]
        if 'Month' in filtered_df.columns and months:
            filtered_df = filtered_df[filtered_df['Month'].isin(months)]
        if 'Type' in filtered_df.columns and leave_types:
            filtered_df = filtered_df[filtered_df['Type'].isin(leave_types)]
        
        filtered_df['LeaveValue'] = filtered_df['IsHalf'].apply(lambda x: 0.5 if x else 1.0) if 'IsHalf' in filtered_df.columns else 1.0

        # --- Metrics Summary ---
        st.markdown("### 📌 สรุปภาพรวมการลา")
        col1, col2, col3, col4, col5 = st.columns(5)
        
        def get_count(l_type):
            if filtered_df.empty or 'Type' not in filtered_df.columns: return 0
            val = filtered_df[filtered_df['Type'] == l_type]['LeaveValue'].sum()
            return int(val) if float(val).is_integer() else val

        col1.metric("🤒 ลาป่วย", f"{get_count('ลาป่วย')} วันทำการ")
        col2.metric("🏃 ลากิจ", f"{get_count('ลากิจ')} วันทำการ")
        col3.metric("🏖️ ลาพักผ่อน", f"{get_count('ลาพักผ่อน')} วันทำการ")
        col4.metric("⏰ มาสาย", f"{get_count('สาย')} ครั้ง")
        other_val = filtered_df[~filtered_df['Type'].isin(['ลาป่วย', 'ลากิจ', 'ลาพักผ่อน', 'สาย'])]['LeaveValue'].sum() if not filtered_df.empty else 0
        col5.metric("👶 ลาอื่นๆ", f"{int(other_val) if float(other_val).is_integer() else other_val} วันทำการ")
        
        st.markdown("<br>", unsafe_allow_html=True)

        # --- Main Content Tabs ---
        tab1, tab2, tab3, tab4 = st.tabs([
            "🏛️ รายงานตามแบบฟอร์มทางการ (กรมเจ้าท่า)",
            "📋 รายการข้อมูลเชิงลึก (รายวัน)",
            "📊 กราฟวิเคราะห์สถิติ",
            "👤 เจาะลึกรายบุคคล"
        ])

        # ----------------------------------------------------
        # TAB 1: รายงานตามแบบฟอร์มทางการ (กรมเจ้าท่า)
        # ----------------------------------------------------
        with tab1:
            org_display = header_meta.get('org') or 'กรมเจ้าท่า'
            title_display = header_meta.get('title') or 'รายงานการมาปฏิบัติราชการของข้าราชการพลเรือนสามัญ'
            range_display = header_meta.get('date_range') or 'ตั้งแต่วันที่ 01/10/2569 ถึงวันที่ 28/10/2569'
            fmt_display = header_meta.get('format') or 'รูปแบบการออกรายงาน : โครงสร้างตามกฎหมาย - ทั้งส่วนราชการ - กรมเจ้าท่า - กระทรวงคมนาคม'

            st.markdown(f"""
            <div class="official-report-card">
                <div class="official-org-title">🏛️ {org_display}</div>
                <div class="official-main-title">{title_display}</div>
                <div class="official-date-badge">📅 {range_display}</div>
                <div class="official-format-sub"><b>หมายเหตุรูปแบบ:</b> {fmt_display}</div>
            </div>
            """, unsafe_allow_html=True)

            # Filter options for official table
            col_f1, col_f2, col_f3 = st.columns([1.5, 1.5, 1.8])
            with col_f1:
                search_kw = st.text_input("🔍 ค้นหาชื่อ-สกุล ในตารางทางการ", placeholder="พิมพ์ชื่อหรือนามสกุล...")
            with col_f2:
                only_with_leave = st.checkbox("แสดงเฉพาะผู้มีประวัติการลา", value=False)
            with col_f3:
                # Export official Excel button
                excel_official_bytes = generate_official_leave_excel(official_summary_df, header_meta)
                st.download_button(
                    label="📥 ดาวน์โหลดแบบฟอร์มทางการ (Excel)",
                    data=excel_official_bytes,
                    file_name="รายงานการมาปฏิบัติราชการ_กรมเจ้าท่า.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )

            # Prepare filtered view of official summary
            display_summary = official_summary_df.copy()
            if search_kw:
                display_summary = display_summary[display_summary['ชื่อ - สกุล'].str.contains(search_kw, na=False)]
            if only_with_leave:
                has_leave_mask = (
                    display_summary['ลาป่วย_วันทำการ'].notna() |
                    display_summary['ลากิจ_วันทำการ'].notna() |
                    display_summary['ลาพักผ่อน_วันทำการ'].notna() |
                    display_summary['สาย_วันทำการ'].notna() |
                    display_summary['ลาอื่นๆ_วันทำการ'].notna() |
                    display_summary['ขาดราชการ_วันทำการ'].notna() |
                    display_summary['หมายเหตุ'].notna() & (display_summary['หมายเหตุ'] != '')
                )
                display_summary = display_summary[has_leave_mask]

            st.caption(f"👥 แสดงผล {len(display_summary)} รายการ จากทั้งหมด {len(official_summary_df)} บุคลากร")

            # Build MultiIndex table for clean official rendering
            col_tuples = [
                ('ชื่อ - สกุล', ''),
                ('ลาป่วย', 'วันทำการ'),
                ('ลาป่วย', 'จำนวนครั้ง'),
                ('ลากิจ', 'วันทำการ'),
                ('ลากิจ', 'จำนวนครั้ง'),
                ('รวมลาป่วย กิจ', 'วันทำการ'),
                ('รวมลาป่วย กิจ', 'จำนวนครั้ง'),
                ('สาย (วันทำการ)', ''),
                ('ลาพักผ่อน', 'วันทำการ'),
                ('ลาพักผ่อน', 'จำนวนครั้ง'),
                ('ลาอื่นๆ (วันทำการ)', ''),
                ('ขาดราชการ (วันทำการ)', ''),
                ('หมายเหตุ', '')
            ]

            formatted_rows = []
            for _, r in display_summary.iterrows():
                row_vals = [
                    r.get('ชื่อ - สกุล', ''),
                    r.get('ลาป่วย_วันทำการ'),
                    r.get('ลาป่วย_จำนวนครั้ง'),
                    r.get('ลากิจ_วันทำการ'),
                    r.get('ลากิจ_จำนวนครั้ง'),
                    r.get('รวมลาป่วย_กิจ_วันทำการ'),
                    r.get('รวมลาป่วย_กิจ_จำนวนครั้ง'),
                    r.get('สาย_วันทำการ'),
                    r.get('ลาพักผ่อน_วันทำการ'),
                    r.get('ลาพักผ่อน_จำนวนครั้ง'),
                    r.get('ลาอื่นๆ_วันทำการ'),
                    r.get('ขาดราชการ_วันทำการ'),
                    r.get('หมายเหตุ', '')
                ]
                clean_vals = []
                for idx, v in enumerate(row_vals):
                    if idx in [0, 12]:
                        clean_vals.append("" if (pd.isna(v) or v is None or str(v).strip() == 'None') else str(v).strip())
                    else:
                        if pd.isna(v) or v is None or v == '':
                            clean_vals.append("")
                        else:
                            try:
                                f_val = float(v)
                                clean_vals.append(str(int(f_val)) if f_val.is_integer() else str(f_val))
                            except Exception:
                                clean_vals.append(str(v))
                formatted_rows.append(clean_vals)

            if formatted_rows:
                multi_cols = pd.MultiIndex.from_tuples(col_tuples)
                df_official_table = pd.DataFrame(formatted_rows, columns=multi_cols)
                
                # Apply styles
                def highlight_cells(col):
                    styles = []
                    col_name = col.name
                    for val in col:
                        v = str(val).strip()
                        if col_name == ('ชื่อ - สกุล', ''):
                            styles.append("text-align: left; font-weight: 600; min-width: 240px;")
                        elif col_name == ('หมายเหตุ', ''):
                            styles.append("text-align: left; min-width: 300px; color: #2D3748;")
                        elif v != "":
                            if 'ป่วย' in col_name[0]:
                                styles.append("background-color: #FFF5F5; color: #C53030; font-weight: 700; text-align: center;")
                            elif 'กิจ' in col_name[0]:
                                styles.append("background-color: #F0FFF4; color: #22543D; font-weight: 700; text-align: center;")
                            elif 'พักผ่อน' in col_name[0]:
                                styles.append("background-color: #EBF8FF; color: #2B6CB0; font-weight: 700; text-align: center;")
                            elif 'สาย' in col_name[0]:
                                styles.append("background-color: #FEFCBF; color: #744210; font-weight: 700; text-align: center;")
                            else:
                                styles.append("background-color: #FAF5FF; color: #553C9A; font-weight: 700; text-align: center;")
                        else:
                            styles.append("text-align: center;")
                    return styles

                st.dataframe(
                    df_official_table.style.apply(highlight_cells, axis=0),
                    use_container_width=True,
                    height=520,
                    hide_index=True
                )
            else:
                st.info("ไม่พบรายการที่ตรงกับเงื่อนไขการค้นหา")

        # ----------------------------------------------------
        # TAB 2: รายละเอียดการลา (Daily Records)
        # ----------------------------------------------------
        with tab2:
            st.markdown("#### 📄 รายละเอียดการลาทั้งหมดตามตัวกรอง")
            if not filtered_df.empty:
                cols_to_show = [c for c in ['BudgetYear', 'Month', 'Day', 'Name', 'Type', 'IsHalf', 'Remarks'] if c in filtered_df.columns]
                display_df = filtered_df[cols_to_show].copy()
                col_rename = {'BudgetYear': 'ปีงบ', 'Month': 'เดือน', 'Day': 'วันที่', 'Name': 'ชื่อ-สกุล', 'Type': 'ประเภท', 'IsHalf': 'ครึ่งวัน', 'Remarks': 'หมายเหตุ'}
                display_df = display_df.rename(columns=col_rename)
                
                st.dataframe(
                    display_df, 
                    use_container_width=True,
                    height=450,
                    hide_index=True
                )
            else:
                st.info("ไม่พบข้อมูลรายการลาตามตัวกรองที่เลือก")

        # ----------------------------------------------------
        # TAB 3: กราฟวิเคราะห์สถิติ
        # ----------------------------------------------------
        with tab3:
            st.markdown("#### 📈 เปรียบเทียบจำนวนวัน/ครั้งการลาแยกตามบุคคล")
            summary_by_person = filtered_df.groupby(['Name', 'Type'])['LeaveValue'].sum().reset_index(name='จำนวนครั้ง') if not filtered_df.empty else pd.DataFrame()
            
            if not summary_by_person.empty:
                color_discrete_map = {
                    'ลาป่วย': '#ef553b', 
                    'ลากิจ': '#00cc96', 
                    'ลาพักผ่อน': '#636efa', 
                    'สาย': '#ab63fa',
                    'ลาช่วยภริยาคลอดบุตร': '#ffa15a'
                }
                
                fig = px.bar(
                    summary_by_person, 
                    x='Name', 
                    y='จำนวนครั้ง', 
                    color='Type',
                    barmode='group',
                    color_discrete_map=color_discrete_map,
                    template='plotly_white'
                )
                fig.update_layout(
                    xaxis_title="รายชื่อบุคลากร",
                    yaxis_title="จำนวนวัน/ครั้ง",
                    legend_title="ประเภทการลา",
                    font=dict(family="Prompt, sans-serif"),
                    hovermode="x unified",
                    margin=dict(l=20, r=20, t=20, b=20)
                )
                st.plotly_chart(fig, use_container_width=True)

                # Download graph data
                graph_pivot = summary_by_person.pivot_table(index='Name', columns='Type', values='จำนวนครั้ง', fill_value=0).reset_index()
                buf_g = io.BytesIO()
                with pd.ExcelWriter(buf_g, engine='xlsxwriter') as writer:
                    graph_pivot.to_excel(writer, index=False, sheet_name='ข้อมูลกราฟ')
                st.download_button(
                    label="📊 ดาวน์โหลดข้อมูลกราฟสถิติ (Excel)",
                    data=buf_g.getvalue(),
                    file_name="สรุปสถิติวันลา.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.info("ไม่พบข้อมูลสำหรับสร้างกราฟจากตัวกรองที่เลือก")

        # ----------------------------------------------------
        # TAB 4: เจาะลึกรายบุคคล
        # ----------------------------------------------------
        with tab4:
            st.markdown("### 🔍 ตรวจสอบประวัติการลารายบุคคล")
            all_person_names = sorted(df['Name'].dropna().unique()) if 'Name' in df.columns else []
            selected_person = st.selectbox(
                "👤 เลือกระบุรายชื่อเพื่อดูประวัติทั้งหมด", 
                options=["-- เลือกรายชื่อ --"] + all_person_names
            )
            
            if selected_person != "-- เลือกรายชื่อ --":
                person_data = filtered_df[filtered_df['Name'] == selected_person].sort_values(['BudgetYear', 'Month', 'Day']) if 'Name' in filtered_df.columns else pd.DataFrame()
                
                st.markdown(f"##### ประวัติของ: <span style='color:#E97D32; font-weight:700;'>{selected_person}</span>", unsafe_allow_html=True)
                
                if person_data.empty:
                    st.warning("ไม่พบประวัติการลาของบุคคลนี้ตามตัวกรองที่เลือก")
                else:
                    # Metrics for person
                    st.markdown("##### 📌 สรุปยอดรวมของบุคคลนี้")
                    total_summary = person_data.groupby('Type')['LeaveValue'].sum().reset_index()
                    cols_p = st.columns(len(total_summary) if not total_summary.empty else 1)
                    for i, row in total_summary.iterrows():
                        v = row['LeaveValue']
                        v_str = int(v) if float(v).is_integer() else v
                        cols_p[i].metric(row['Type'], f"{v_str} วันทำการ")
                    
                    st.markdown("<br>", unsafe_allow_html=True)
                    
                    # Detailed History Table
                    st.markdown("##### 📄 ประวัติการลาโดยละเอียด")
                    ui_display_data = []
                    for _, row in person_data.iterrows():
                        m_str = row.get('Month', '')
                        d_num = row.get('Day', '')
                        y_str = row.get('BudgetYear', '')
                        date_str = f"วันที่ {d_num} {m_str} {y_str}".strip()
                        
                        t_str = row.get('Type', '')
                        is_h = row.get('IsHalf', False)
                        val_str = '0.5' if is_h else '1'
                        
                        ui_display_data.append({
                            'วัน เดือน ปี': date_str,
                            'ประเภทการลา': t_str,
                            'จำนวนวัน': val_str,
                            'หมายเหตุ': row.get('Remarks', '')
                        })
                    
                    st.dataframe(pd.DataFrame(ui_display_data), use_container_width=True, hide_index=True)

else:
    # Empty State UI
    st.markdown("<br><br>", unsafe_allow_html=True)
    col_img1, col_img2, col_img3 = st.columns([1, 2, 1])
    with col_img2:
        st.info("👋 ยินดีต้อนรับสู่ระบบวันลา! กรุณาอัปโหลดไฟล์ตารางวันลาทางเมนูด้านซ้ายเพื่อเริ่มต้นใช้งาน")
        st.markdown(
            """
            <div style='text-align: center; opacity: 0.7; margin-top: 20px; line-height: 1.8;'>
                <p>📄 รองรับทั้งแบบฟอร์มทางการกรมเจ้าท่า (<b>ข้าราชการ.xlsx</b> หรือ <b>.csv</b>)<br>และไฟล์ตารางวันลาประจำปี (<b>1.ตารางวันลา.xlsx</b>)</p>
            </div>
            """, 
            unsafe_allow_html=True
        )
