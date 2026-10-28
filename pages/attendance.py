import os
import json
import calendar
import hmac
import streamlit as st
import pandas as pd
from datetime import datetime, time, timedelta
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import pdfplumber
import re
from io import BytesIO
import plotly.express as px
import holidays
import unicodedata
from difflib import SequenceMatcher
from data_parser import load_data

# -------------------------
# ตั้งค่าหน้าเว็บและ CSS
# -------------------------

# ฉีด CSS ดีไซน์ Playful Modern อบอุ่น สไตล์ Sonnenbloom (แก้สระลอย + ปรับแต่ง UI ครบวงจร)
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Prompt:wght@300;400;500;600;700&family=Sarabun:wght@300;400;500;600;700&display=swap');

    html, body, [class*="css"], .stApp {
        font-family: 'Prompt', 'Sarabun', sans-serif !important;
        background-color: #FDFBF7 !important;
        color: #2D2825 !important;
    }

    [data-testid="stSidebar"] {
        background-color: #F8F3EA !important;
        border-right: 1.5px solid #EAE1D3 !important;
    }
    [data-testid="stSidebarNav"] span {
        font-family: 'Prompt', sans-serif !important;
        font-weight: 500 !important;
    }

    h1, h2, h3, h4, h5, h6 {
        font-family: 'Prompt', sans-serif !important;
        font-weight: 700 !important;
        color: #2D2825 !important;
    }

    .stMarkdown, .stTable, .stDataFrame, td, th {
        font-family: 'Sarabun', 'Prompt', sans-serif !important;
        line-height: 1.6 !important;
    }

    /* Custom Header Banner */
    .att-header-container {
        background: linear-gradient(135deg, #FFF6EE 0%, #FEF1DE 45%, #F5F9F4 100%);
        border: 2px solid #F5E6D3;
        border-radius: 24px;
        padding: 26px 30px;
        margin-bottom: 22px;
        box-shadow: 0 12px 30px -10px rgba(233, 125, 50, 0.12);
    }

    .att-pill-badge {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        padding: 5px 14px;
        background: #FFFFFF;
        border: 1px solid #F0D7C2;
        border-radius: 9999px;
        font-size: 0.82rem;
        font-weight: 600;
        color: #D4661B;
        box-shadow: 0 2px 6px rgba(212, 102, 27, 0.08);
        margin-bottom: 10px;
    }

    .att-header-title {
        font-size: 1.85rem !important;
        font-weight: 700 !important;
        color: #2D2825 !important;
        margin: 0 0 6px 0 !important;
    }

    .att-header-sub {
        font-size: 0.98rem;
        color: #6E655E;
        margin: 0;
    }

    .status-badge-ready {
        display: inline-flex;
        align-items: center;
        gap: 8px;
        background: #E8F5E9;
        border: 1.5px solid #C8E6C9;
        color: #2E7D32;
        padding: 8px 18px;
        border-radius: 9999px;
        font-weight: 600;
        font-size: 0.88rem;
        box-shadow: 0 2px 8px rgba(46, 125, 50, 0.1);
    }

    .status-badge-empty {
        display: inline-flex;
        align-items: center;
        gap: 8px;
        background: #F8F5EE;
        border: 1.5px solid #EAE1D3;
        color: #7A7067;
        padding: 8px 18px;
        border-radius: 9999px;
        font-weight: 600;
        font-size: 0.88rem;
    }

    /* Custom Streamlit Tabs (Pill Segmented Control) */
    div[data-testid="stTabs"] {
        margin-top: 8px;
        margin-bottom: 24px;
    }

    div[data-baseweb="tab-list"] {
        gap: 8px !important;
        background-color: #F4EFE6 !important;
        padding: 6px !important;
        border-radius: 9999px !important;
        border: 1px solid #EAE1D3 !important;
        display: inline-flex !important;
        flex-wrap: wrap !important;
    }

    button[data-baseweb="tab"] {
        border-radius: 9999px !important;
        padding: 8px 18px !important;
        font-family: 'Prompt', sans-serif !important;
        font-weight: 500 !important;
        font-size: 0.92rem !important;
        color: #6B625B !important;
        background-color: transparent !important;
        border: none !important;
        transition: all 0.25s cubic-bezier(0.25, 0.8, 0.25, 1) !important;
    }

    button[data-baseweb="tab"]:hover {
        color: #E97D32 !important;
        background-color: rgba(255, 255, 255, 0.7) !important;
    }

    button[data-baseweb="tab"][aria-selected="true"] {
        background: linear-gradient(135deg, #FF8E42 0%, #E97D32 100%) !important;
        color: #FFFFFF !important;
        font-weight: 600 !important;
        box-shadow: 0 4px 14px rgba(233, 125, 50, 0.35) !important;
    }

    div[data-baseweb="tab-border"], div[data-baseweb="tab-highlight"] {
        display: none !important;
    }

    /* Metric Cards */
    div[data-testid="stMetric"] {
        background: #FFFFFF !important;
        border: 1.5px solid #EFE6DA !important;
        border-radius: 20px !important;
        padding: 18px 20px !important;
        box-shadow: 0 4px 14px rgba(0, 0, 0, 0.03) !important;
        transition: all 0.25s cubic-bezier(0.25, 0.8, 0.25, 1) !important;
    }

    div[data-testid="stMetric"]:hover {
        transform: translateY(-5px) !important;
        box-shadow: 0 12px 28px -4px rgba(233, 125, 50, 0.16) !important;
        border-color: #E97D32 !important;
    }

    div[data-testid="stMetricLabel"] {
        font-family: 'Prompt', sans-serif !important;
        font-size: 0.95rem !important;
        font-weight: 600 !important;
        color: #6B625B !important;
    }

    div[data-testid="stMetricValue"] {
        font-family: 'Prompt', 'Sarabun', sans-serif !important;
        font-size: 2rem !important;
        font-weight: 700 !important;
        color: #2D2825 !important;
    }

    /* Custom Buttons */
    div.stButton > button {
        border-radius: 9999px !important;
        font-family: 'Prompt', sans-serif !important;
        font-weight: 600 !important;
        font-size: 0.98rem !important;
        padding: 10px 24px !important;
        transition: all 0.25s ease !important;
    }

    div.stButton > button:hover {
        transform: translateY(-2px) scale(1.01) !important;
    }

    div.stDownloadButton > button {
        border-radius: 9999px !important;
        font-family: 'Prompt', sans-serif !important;
        font-weight: 600 !important;
        font-size: 0.95rem !important;
        padding: 10px 22px !important;
        border: 1.5px solid #EFE6DA !important;
        background: #FFFFFF !important;
        color: #2D2825 !important;
        box-shadow: 0 2px 8px rgba(0,0,0,0.04) !important;
        transition: all 0.25s ease !important;
    }

    div.stDownloadButton > button:hover {
        background: #FFF6EE !important;
        border-color: #E97D32 !important;
        color: #D4661B !important;
        transform: translateY(-2px) !important;
    }

    /* File Uploaders */
    div[data-testid="stFileUploaderDropzone"] {
        border: 2px dashed #E28743 !important;
        background: #FFFCF8 !important;
        border-radius: 20px !important;
        padding: 22px !important;
        transition: all 0.25s ease !important;
    }

    div[data-testid="stFileUploaderDropzone"]:hover {
        background: #FFF5EC !important;
        border-color: #D4661B !important;
    }

    /* Dataframe containers */
    div[data-testid="stDataFrame"] {
        border-radius: 18px !important;
        overflow: hidden !important;
        border: 1.5px solid #EFE6DA !important;
        box-shadow: 0 4px 16px rgba(0, 0, 0, 0.03) !important;
    }

    /* Score Badge */
    .score-badge {
        display: inline-block;
        padding: 6px 16px;
        border-radius: 9999px;
        font-weight: 600;
        font-size: 0.92rem;
        margin-top: 8px;
    }
    .score-high {
        background: #E8F5E9;
        color: #2E7D32;
        border: 1px solid #C8E6C9;
    }
    .score-med {
        background: #FFF8E1;
        color: #F57F17;
        border: 1px solid #FFE082;
    }
    .score-low {
        background: #FFEBEE;
        color: #C62828;
        border: 1px solid #FFCDD2;
    }
    </style>
    """, unsafe_allow_html=True)

# -------------------------
# ตั้งค่า Sidebar (Settings)
# -------------------------
with st.sidebar:
    st.header("⚙️ ตั้งค่าระบบ")
    
    # 1. ปรับเวลาเข้างานได้เอง
    late_time = st.time_input("เวลาที่ถือว่าสาย", time(8, 36))
    LATE_THRESHOLD_MINUTES = late_time.hour * 60 + late_time.minute
    
    st.subheader("📅 การจัดการวันหยุด")
    # 2. ตัวเลือกวันหยุด
    exclude_weekends = st.checkbox("ไม่นับเสาร์-อาทิตย์", value=True)
    exclude_th_holidays = st.checkbox("ไม่นับวันหยุดนักขัตฤกษ์ไทย", value=True)
    
    # 3. ระบุวันหยุดเพิ่มเติมเอง (กรณีพิเศษ)
    extra_holidays = st.date_input("เลือกวันหยุดเพิ่มเติม (กรณีพิเศษ):", value=[], help="คลิกเพื่อเลือกวันที่ต้องการให้เป็นวันหยุดเพิ่มเติมของหน่วยงาน")
    if not isinstance(extra_holidays, list):
        extra_holidays = [extra_holidays]

    st.markdown("---")
    st.markdown(
        '<p style="text-align: center; color: #6b7280; font-size: 0.9rem;">ระบบนี้จัดทำขึ้นโดย<br>นายเอกชัย ไตรโยธี</p>',
        unsafe_allow_html=True,
    )

# เตรียมข้อมูลวันหยุดไทย (รองรับปีปัจจุบันและปีถัดไป)
current_year = datetime.now().year
# ไม่ระบุ years -> ไลบรารีจะโหลดวันหยุดของปีที่ถูกค้นหาให้อัตโนมัติ (รองรับไฟล์ย้อนหลัง/ข้ามปี)
th_holidays = holidays.Thailand()

# -------------------------
# ฟังก์ชันตรวจสอบวันหยุด
# -------------------------
def is_it_holiday(dt, exclude_weekends, exclude_th_holidays, extra_list):
    if exclude_weekends and dt.weekday() >= 5:
        return True
    if exclude_th_holidays and dt in th_holidays:
        return True
    if dt.date() in extra_list:
        return True
    return False

# -------------------------
# ฟอนต์ไทยสำหรับ matplotlib
# -------------------------
def set_thai_font():
    local_font_path = "THSarabunNew.ttf" 
    if os.path.exists(local_font_path):
        try:
            fm.fontManager.addfont(local_font_path)
            prop = fm.FontProperties(fname=local_font_path)
            font_name = prop.get_name()
            plt.rcParams['font.family'] = font_name
            return font_name
        except Exception:
            pass
    return "Tahoma"

_used_font = set_thai_font()

# -------------------------
# Utilities
# -------------------------
def normalize_whitespace(s: str) -> str:
    if s is None: return ""
    
    # 1. จัดการ สระอำ ที่แตกเป็น space + วรรณยุกต์ + า (เช่น 'อ านวย' -> 'อำนวย', 'น ้าผึ้ง' -> 'น้ำผึ้ง')
    s = re.sub(r'\s+([\u0E48-\u0E4B])\s*า', r'\1ำ', s)
    s = re.sub(r'\s+า', 'ำ', s)
    s = re.sub(r'ำ([\u0E48-\u0E4B])', r'\1ำ', s)

    # 2. ซ่อมคำเฉพาะที่ฟอนต์ระบบสแกนเวลาชอบแตก
    s = re.sub(r'ข้า\s*ราชการ', 'ข้าราชการ', s)
    s = re.sub(r'พนักงาน\s*ราชการ', 'พนักงานราชการ', s)
    s = re.sub(r'ลูกจ้าง\s*ชั่วคราว', 'ลูกจ้างชั่วคราว', s)
    s = re.sub(r'ลูกจ้าง\s*ประจำ|ลูกจ้าง\s*ประจา|ประจา\b', 'ลูกจ้างประจำ', s)
    s = re.sub(r'สวนสา\s*ราญ|สวนสาราญ', 'สวนสำราญ', s)
    s = re.sub(r'วรรณคา\b', 'วรรณคำ', s)
    s = re.sub(r'\bคา\s*สมหมาย|\bคาสมหมาย', 'คำสมหมาย', s)
    
    # 3. ล้างวงกลมประ (U+25CC) และอักขระล่องหน/ควบคุมทั้งหมด (Non-printable characters)
    s = re.sub(r'[\u200b\ufeff\xa0\u25cc\u00ad]', '', s)
    s = "".join(ch for ch in s if unicodedata.category(ch)[0] != "C")

    # 4. ลบช่องว่างหรือสิ่งที่กั้นระหว่างพยัญชนะกับสระ/วรรณยุกต์
    s = re.sub(r'\s+([\u0E31\u0E34-\u0E3A\u0E47-\u0E4E])', r'\1', s)
    
    # 5. จัดลำดับ Re-ordering (สลับวรรณยุกต์ที่มาทีหลังสระข้าง เช่น า้ -> ้า)
    s = re.sub(r'([\u0E30\u0E32])([\u0E48-\u0E4B])', r'\2\1', s)
    
    # 6. จัดลำดับกรณี วรรณยุกต์สลับกับสระบน/ล่าง (สระต้องมาก่อนวรรณยุกต์)
    s = re.sub(r'([\u0E48-\u0E4B])([\u0E31\u0E34-\u0E39])', r'\2\1', s)

    # 7. จัดการ Unicode ให้เป็นรูปแบบมาตรฐานสูงสุด
    s = unicodedata.normalize('NFC', s)
    
    # 8. ลบช่องว่างซ้ำซ้อน
    s = re.sub(r'\s+', ' ', s)
    
    return s.strip()

def month_name_thai(month_num):
    months = ["มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน",
              "กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม"]
    return months[month_num-1] if 1 <= month_num <= 12 else ""

REPORT_DEPT_ORDER = ['ข้าราชการ', 'พนักงานราชการ', 'ลูกจ้างประจำ', 'ลูกจ้างชั่วคราว']
THAI_DOW = ["จ", "อ", "พ", "พฤ", "ศ", "ส", "อา"]

def get_department_options(df: pd.DataFrame):
    """ดึงรายชื่อฝ่าย/กลุ่มงานที่มีอยู่ใน DataFrame จัดเรียงตามมาตรฐาน"""
    existing = df['แผนก'].dropna().unique().tolist()
    ordered = [d for d in REPORT_DEPT_ORDER if d in existing]
    others = [d for d in existing if d and d not in DEPT_VARIANTS]
    ordered.extend(sorted(others))
    if '' in existing or df['แผนก'].isna().any():
        ordered.append("ไม่ระบุฝ่าย")
    return ordered

def filter_by_dept(df: pd.DataFrame, selected_dept: str):
    """กรอง DataFrame ตามฝ่าย/กลุ่มงาน"""
    if not selected_dept or selected_dept == "ทั้งหมด":
        return df
    if selected_dept == "ไม่ระบุฝ่าย":
        return df[df['แผนก'].isin(['', None]) | df['แผนก'].isna()]
    return df[df['แผนก'] == selected_dept]

def build_monthly_report_table(df_dept: pd.DataFrame, df_all: pd.DataFrame, exc_week: bool = True, exc_th: bool = True, extra_list: list = None):
    """สร้างตารางรายงานประจำเดือนตามรูปแบบมาตรฐานศูนย์ฝึกพาณิชย์นาวี (หัวตาราง 2 ชั้น วันในสัปดาห์ + วันที่ 1..31)"""
    if extra_list is None:
        extra_list = []
    valid_dates = df_all['วันที่'].dropna()
    if valid_dates.empty:
        return pd.DataFrame(), set(), 2026, 10
    year = int(valid_dates.dt.year.mode()[0])
    month = int(valid_dates.dt.month.mode()[0])
    days_in_month = calendar.monthrange(year, month)[1]

    df_month_all = df_all[(df_all['วันที่'].dt.year == year) & (df_all['วันที่'].dt.month == month)]
    uploaded_days = set(df_month_all['วันที่'].dt.day.unique())
    max_uploaded_day = max(uploaded_days) if uploaded_days else 0

    holiday_days = set()
    for d in range(1, days_in_month + 1):
        dt_d = pd.Timestamp(year=year, month=month, day=d)
        if is_it_holiday(dt_d, exc_week, exc_th, extra_list):
            holiday_days.add(d)

    roster = get_active_roster()
    roster_order_map = {re.sub(r'\s+', '', p['name']): i for i, p in enumerate(roster)}

    def _roster_sort_key(name: str):
        clean_k = re.sub(r'\s+', '', name)
        if clean_k in roster_order_map:
            return (0, roster_order_map[clean_k])
        core_k = clean_core_key(clean_k)
        for p_name, idx in roster_order_map.items():
            if clean_core_key(p_name) == core_k or p_name.startswith(clean_k) or clean_k.startswith(p_name):
                return (0, idx)
        return (1, name)

    emp_names = sorted(df_dept['ชื่อ-สกุล'].drop_duplicates().tolist(), key=_roster_sort_key)

    df_dept_m = df_dept[(df_dept['วันที่'].dt.year == year) & (df_dept['วันที่'].dt.month == month)]
    status_lookup = {}
    for _, r in df_dept_m.iterrows():
        d_num = int(r['วันที่'].day)
        status_lookup[(r['ชื่อ-สกุล'], d_num)] = shorten_status_for_pivot(r['สถานะ'])

    col_tuples = [('ที่', ''), ('ชื่อ - สกุล', '')]
    for d in range(1, days_in_month + 1):
        dt_d = pd.Timestamp(year=year, month=month, day=d)
        dow_str = THAI_DOW[dt_d.weekday()]
        col_tuples.append((dow_str, str(d)))

    rows = []
    for idx, name in enumerate(emp_names, start=1):
        row_data = [idx, name]
        for d in range(1, days_in_month + 1):
            if (name, d) in status_lookup:
                sym = status_lookup[(name, d)]
            else:
                if d <= max_uploaded_day and d in holiday_days:
                    sym = " -"
                elif d in uploaded_days:
                    sym = "x"
                else:
                    sym = ""
            row_data.append(sym)
        rows.append(row_data)

    columns = pd.MultiIndex.from_tuples(col_tuples)
    report_df = pd.DataFrame(rows, columns=columns)
    return report_df, holiday_days, year, month

def style_monthly_report_table(report_df: pd.DataFrame, holiday_days: set):
    """จัดสีตารางรายงานประจำเดือนให้เหมือนไฟล์ตัวอย่าง (คอลัมน์วันหยุดสีฟ้า #95B3D7 และสีตามสถานะ)"""
    def style_col(col):
        col_name = col.name
        day_str = col_name[1] if isinstance(col_name, tuple) and len(col_name) > 1 else ""
        is_hol_col = int(day_str) in holiday_days if str(day_str).isdigit() else False
        styles = []
        for val in col:
            v = str(val).strip()
            if col_name == ('ที่', ''):
                styles.append("text-align: center; font-weight: 600;")
            elif col_name == ('ชื่อ - สกุล', ''):
                styles.append("text-align: left; font-weight: 500;")
            elif is_hol_col:
                if v == "✓":
                    styles.append("background-color: #95B3D7; color: #006100; font-weight: 700; text-align: center;")
                elif v in ["สาย", "เข้า", "ออก", "x", "ขาด(บ)", "ขาด(ช)"]:
                    styles.append("background-color: #95B3D7; color: #9C0006; font-weight: 700; text-align: center;")
                else:
                    styles.append("background-color: #95B3D7; color: #1F3864; text-align: center;")
            else:
                if v == "✓":
                    styles.append("color: #1E7E34; font-weight: 700; text-align: center;")
                elif v == "สาย":
                    styles.append("background-color: #FFEB9C; color: #9C6500; font-weight: 700; text-align: center;")
                elif v in ["เข้า", "ออก"]:
                    styles.append("background-color: #FCE4D6; color: #C65911; font-weight: 700; text-align: center;")
                elif v in ["x", "ขาด(บ)", "ขาด(ช)"]:
                    styles.append("background-color: #FFC7CE; color: #9C0006; font-weight: 700; text-align: center;")
                elif v == "-":
                    styles.append("background-color: #95B3D7; color: #1F3864; text-align: center;")
                elif v and v != "nan":
                    styles.append("background-color: #D1ECF1; color: #0C5460; font-weight: 700; text-align: center;")
                else:
                    styles.append("text-align: center;")
        return styles
    return report_df.style.apply(style_col, axis=0)

def df_to_excel_bytes(pivot: pd.DataFrame, df_all: pd.DataFrame, dept_label: str = "", exc_week: bool = True, exc_th: bool = True, extra_list: list = None):
    """ส่งออกรายงานประจำเดือนเป็นไฟล์ Excel รูปแบบเดียวกับไฟล์ตัวอย่างของศูนย์ฝึกพาณิชย์นาวี (แยกชีตตามฝ่าย)"""
    if extra_list is None:
        extra_list = []
    output = BytesIO()
    valid_dates = df_all['วันที่'].dropna()
    year = int(valid_dates.dt.year.mode()[0]) if not valid_dates.empty else datetime.now().year
    month = int(valid_dates.dt.month.mode()[0]) if not valid_dates.empty else datetime.now().month
    days_in_month = calendar.monthrange(year, month)[1]
    month_th = month_name_thai(month)
    year_be = year + 543

    # กำหนดรายการชีตที่จะสร้าง
    if not dept_label or dept_label == "ทั้งหมด":
        depts_to_export = [d for d in REPORT_DEPT_ORDER if d in df_all['แผนก'].values]
        other_depts = [d for d in df_all['แผนก'].dropna().unique() if d and d not in REPORT_DEPT_ORDER]
        depts_to_export.extend(other_depts)
        if (df_all['แผนก'] == '').any() or df_all['แผนก'].isna().any():
            depts_to_export.append("ไม่ระบุฝ่าย")
        if not depts_to_export:
            depts_to_export = ["ทั้งหมด"]
    else:
        depts_to_export = [dept_label]

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        base_font = 'TH SarabunPSK'
        sym_font = 'Calibri'
        hol_bg = '#95B3D7'

        title_fmt = workbook.add_format({'bold': True, 'font_name': base_font, 'font_size': 18, 'align': 'center', 'valign': 'vcenter'})
        subtitle_fmt = workbook.add_format({'bold': True, 'font_name': base_font, 'font_size': 18, 'align': 'center', 'valign': 'vcenter', 'bottom': 1})
        hdr_fmt = workbook.add_format({'bold': True, 'font_name': base_font, 'font_size': 18, 'align': 'center', 'valign': 'vcenter', 'border': 1})
        hdr_hol_fmt = workbook.add_format({'bold': True, 'font_name': base_font, 'font_size': 18, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': hol_bg})
        idx_fmt = workbook.add_format({'font_name': base_font, 'font_size': 18, 'align': 'center', 'valign': 'vcenter', 'border': 1})
        name_fmt = workbook.add_format({'font_name': base_font, 'font_size': 14, 'align': 'left', 'valign': 'vcenter', 'border': 1})
        cell_fmt = workbook.add_format({'font_name': sym_font, 'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'border': 1})
        cell_hol_fmt = workbook.add_format({'font_name': sym_font, 'font_size': 11, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': hol_bg})
        legend_fmt = workbook.add_format({'font_name': base_font, 'font_size': 18, 'align': 'left', 'valign': 'vcenter'})

        for d_name in depts_to_export:
            df_sub = filter_by_dept(df_all, d_name)
            rep_df, holiday_days, _, _ = build_monthly_report_table(df_sub, df_all, exc_week, exc_th, extra_list)
            sheet_title = d_name[:31] if d_name else "Summary"
            ws = workbook.add_worksheet(sheet_title)

            last_day_col = 1 + days_in_month  # Col B is 1, Day 1 is 2 .. Day 31 is 32 (AG)
            dept_header_name = f"{d_name}ศูนย์ฝึกพาณิชย์นาวี" if d_name != "ทั้งหมด" else "เจ้าหน้าที่ศูนย์ฝึกพาณิชย์นาวี"

            # Row 1 & Row 2: Title and Month
            ws.set_row(0, 23.25)
            ws.set_row(1, 23.25)
            ws.merge_range(0, 1, 0, last_day_col, dept_header_name, title_fmt)
            ws.merge_range(1, 1, 1, last_day_col, f"ประจำเดือน {month_th} {year_be}", subtitle_fmt)

            # Row 3 & Row 4: Table Headers
            ws.set_row(2, 23.25)
            ws.set_row(3, 18.0)
            ws.merge_range(2, 0, 3, 0, "ที่", hdr_fmt)
            ws.merge_range(2, 1, 3, 1, "ชื่อ - สกุล", hdr_fmt)

            for d in range(1, days_in_month + 1):
                col_idx = 1 + d
                dt_d = pd.Timestamp(year=year, month=month, day=d)
                dow_str = THAI_DOW[dt_d.weekday()]
                h_fmt = hdr_hol_fmt if d in holiday_days else hdr_fmt
                ws.write(2, col_idx, dow_str, h_fmt)
                ws.write(3, col_idx, d, h_fmt)

            # Legend on right (Col AI = last_day_col + 2)
            legend_col = last_day_col + 2
            ws.write(2, legend_col, "x X (เข้า-ออก)", legend_fmt)
            ws.write(3, legend_col, " - วันหยุดราชการ", legend_fmt)

            # Data rows (Row 5+)
            for r_i, row_vals in enumerate(rep_df.values):
                excel_row = 4 + r_i
                ws.set_row(excel_row, 17.25)
                ws.write(excel_row, 0, int(row_vals[0]), idx_fmt)
                ws.write(excel_row, 1, str(row_vals[1]), name_fmt)
                for d in range(1, days_in_month + 1):
                    col_idx = 1 + d
                    val = row_vals[1 + d]
                    c_fmt = cell_hol_fmt if d in holiday_days else cell_fmt
                    ws.write(excel_row, col_idx, val if val != "" else None, c_fmt)

            # Column widths matching original template
            ws.set_column(0, 0, 4.0)
            ws.set_column(1, 1, 30.0)
            ws.set_column(2, last_day_col, 4.8)
            ws.set_column(last_day_col + 1, last_day_col + 1, 3.0)
            ws.set_column(legend_col, legend_col, 21.0)

    return output.getvalue()

def daily_to_excel_bytes(df_day: pd.DataFrame, report_date, dept_label: str = ""):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        cols_to_exp = ['แผนก', 'ชื่อ-สกุล', 'เข้างาน', 'ออกงาน', 'สถานะ']
        if 'ประเภทการลา' in df_day.columns and df_day['มีใบลา'].any():
            cols_to_exp.extend(['ประเภทการลา', 'หมายเหตุการลา'])
        export_df = df_day[cols_to_exp]
        export_df.to_excel(writer, sheet_name="Daily", startrow=2, index=False)
        workbook, worksheet = writer.book, writer.sheets["Daily"]
        base_font = 'TH SarabunPSK' 
        base_format = workbook.add_format({'font_name': base_font, 'font_size': 14, 'align': 'center', 'valign': 'bottom', 'border': 1})
        name_format = workbook.add_format({'font_name': base_font, 'font_size': 14, 'align': 'left', 'valign': 'bottom', 'border': 1})
        header_format = workbook.add_format({'font_name': base_font, 'font_size': 14, 'bold': True, 'align': 'center', 'valign': 'bottom', 'border': 1, 'bg_color': '#D9D9D9'})
        for row in range(len(export_df) + 3): worksheet.set_row(row, 30, base_format)
        for col_num, value in enumerate(export_df.columns.values): worksheet.write(2, col_num, value, header_format)
        worksheet.set_column(0, 0, 20, base_format); worksheet.set_column(1, 1, 35, name_format); worksheet.set_column(2, 4, 15, base_format)
        if len(cols_to_exp) > 5:
            worksheet.set_column(5, len(cols_to_exp)-1, 20, base_format)
        sub_dept = f" ({dept_label})" if dept_label and dept_label != "ทั้งหมด" else ""
        date_str = report_date.strftime('%d/%m/%Y') if hasattr(report_date, 'strftime') else str(report_date)
        header_text = f"รายงานการเข้า-ออกงาน{sub_dept} ประจำวันที่ {date_str}"
        merge_format = workbook.add_format({'bold': True, 'font_name': base_font, 'font_size': 18, 'align': 'center', 'valign': 'bottom'})
        worksheet.merge_range(0, 0, 0, len(export_df.columns)-1, header_text, merge_format)
    return output.getvalue()

def missing_scan_to_excel_bytes(df_all: pd.DataFrame, dept_label: str = ""):
    output = BytesIO()
    
    # 1. Prepare data
    missing_statuses = ['ลืมสแกนนิ้วเข้า', 'ลืมสแกนนิ้วออก', 'ขาดงาน', 'ขาดงานครึ่งบ่าย', 'ขาดงานครึ่งเช้า']
    df_missing = df_all[df_all['สถานะ'].isin(missing_statuses)].copy()
    if 'มีใบลา' in df_missing.columns:
        df_missing = df_missing[df_missing['มีใบลา'] != True]
    
    if df_missing.empty:
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            pd.DataFrame().to_excel(writer, sheet_name="Missing_Scans")
        return output.getvalue()

    def format_detail(row):
        d = row['วันที่']
        status = row['สถานะ']
        suffix = ""
        if status == 'ลืมสแกนนิ้วเข้า': suffix = "(เวลาเข้า)"
        elif status == 'ลืมสแกนนิ้วออก': suffix = "(เวลาออก)"
        elif status == 'ขาดงาน': suffix = "(เข้าและออก)"
        elif status == 'ขาดงานครึ่งบ่าย': suffix = f"(ขาดงานครึ่งบ่าย ออก {row['ออกงาน']})" if pd.notna(row.get('ออกงาน')) and row.get('ออกงาน') else "(ขาดงานครึ่งบ่าย)"
        elif status == 'ขาดงานครึ่งเช้า': suffix = f"(ขาดงานครึ่งเช้า เข้า {row['เข้างาน']})" if pd.notna(row.get('เข้างาน')) and row.get('เข้างาน') else "(ขาดงานครึ่งเช้า)"
        
        month_name = month_name_thai(d.month)
        year_be = d.year + 543
        return f"วันที่ {d.day} {month_name} {year_be} {suffix}"

    df_missing['detail'] = df_missing.apply(format_detail, axis=1)
    
    grouped = df_missing.groupby(['แผนก', 'ชื่อ-สกุล']).agg({
        'detail': lambda x: "\n".join(x),
        'วันที่': 'count'
    }).reset_index()
    
    grouped.columns = ['แผนก', 'ชื่อ - สกุล', 'วันที่ไม่พบข้อมูล', 'จำนวนวันที่ไม่พบข้อมูล']
    
    valid_dates = df_all['วันที่'].dropna()
    month_th = month_name_thai(int(valid_dates.dt.month.mode()[0])) if not valid_dates.empty else "ไม่ระบุ"
    year_be_header = int(valid_dates.dt.year.mode()[0]) + 543 if not valid_dates.empty else ""

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book
        worksheet = workbook.add_worksheet("Missing_Scans")
        
        base_font = 'TH SarabunPSK'
        title_format = workbook.add_format({'bold': True, 'font_name': base_font, 'font_size': 18, 'align': 'center'})
        header_format = workbook.add_format({'bold': True, 'font_name': base_font, 'font_size': 14, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#D9D9D9'})
        dept_format = workbook.add_format({'bold': True, 'font_name': base_font, 'font_size': 14, 'align': 'left', 'valign': 'vcenter', 'border': 1})
        data_center_format = workbook.add_format({'font_name': base_font, 'font_size': 14, 'align': 'center', 'valign': 'vcenter', 'border': 1})
        data_left_format = workbook.add_format({'font_name': base_font, 'font_size': 14, 'align': 'left', 'valign': 'vcenter', 'border': 1, 'text_wrap': True})
        
        sub_title = f" ({dept_label})" if dept_label and dept_label != "ทั้งหมด" else ""
        worksheet.merge_range('A1:K1', f"รายงานการสแกนนิ้วมือในการเข้า - ออกงานของเจ้าหน้าที่{sub_title}", title_format)
        worksheet.merge_range('A2:K2', "ศูนย์ฝึกพาณิชย์นาวี", title_format)
        worksheet.merge_range('A3:K3', f"ประจำเดือน {month_th} {year_be_header}", title_format)
        
        headers = ['ลำดับ', 'ชื่อ - สกุล', 'จำนวนวันที่ไม่พบข้อมูล', '', '', '', 'วันที่ไม่พบข้อมูล', '', '', '', '']
        for col_num, header in enumerate(headers):
            worksheet.write(3, col_num, header, header_format)
        worksheet.merge_range(3, 2, 3, 5, 'จำนวนวันที่ไม่พบข้อมูล', header_format)
        worksheet.merge_range(3, 6, 3, 10, 'วันที่ไม่พบข้อมูล', header_format)
        
        curr_row = 4
        global_idx = 1
        
        depts = grouped['แผนก'].unique()
        for dept in depts:
            worksheet.merge_range(curr_row, 0, curr_row, 10, dept if dept else "ไม่ระบุฝ่าย", dept_format)
            curr_row += 1
            
            dept_data = grouped[grouped['แผนก'] == dept]
            for _, row in dept_data.iterrows():
                worksheet.write(curr_row, 0, global_idx, data_center_format)
                worksheet.write(curr_row, 1, row['ชื่อ - สกุล'], data_left_format)
                worksheet.merge_range(curr_row, 2, curr_row, 5, row['จำนวนวันที่ไม่พบข้อมูล'], data_center_format)
                worksheet.merge_range(curr_row, 6, curr_row, 10, row['วันที่ไม่พบข้อมูล'], data_left_format)
                
                num_lines = row['วันที่ไม่พบข้อมูล'].count('\n') + 1
                worksheet.set_row(curr_row, 20 * num_lines)
                
                curr_row += 1
                global_idx += 1
                
        worksheet.set_column('A:A', 8)
        worksheet.set_column('B:B', 30)
        worksheet.set_column('C:F', 5)
        worksheet.set_column('G:K', 12)
        
    return output.getvalue()

# -------------------------
# PDF Parsing
# -------------------------
DEPT_VARIANTS = ['ลูกจ้างชั่วคราว', 'ลูกจ้างประจำ', 'พนักงานราชการ', 'ข้าราชการ']
# วันที่: รองรับ 1/10/2569, 01-10-2026, 1 / 10 / 2569, ปี 2 หลัก (69) และกรณีที่ปีติดกับเวลา เช่น 256908:12
DATE_RE = re.compile(r'(?<!\d)(\d{1,2})\s*[\/\-]\s*(\d{1,2})\s*[\/\-]\s*(\d{4}|\d{2}(?!\d))')
TIME_RE = re.compile(r'(?<!\d)(?:[01]?\d|2[0-3]):[0-5]\d(?!\d)')
THAI_RE = re.compile(r'[\u0E01-\u0E2E]')
TITLE_RE = re.compile(r'^(นาย|นาง|น\.ส\.|ว่าที่|ดร\.|จ\.ส\.|พ\.จ\.|ร\.ต\.|ร\.ท\.|ร\.อ\.|น\.ต\.|น\.ท\.|น\.อ\.|พ\.ต\.|พ\.ท\.|พ\.อ\.|ด\.ต\.|ส\.ต\.|ส\.อ\.|จ\.ต\.|จ\.ท\.|จ\.อ\.)')
# คำในบรรทัดหัวตาราง/หัวกระดาษ (ใช้กับบรรทัดที่ "ไม่มีวันที่" เท่านั้น)
HEADER_RE = re.compile(r'ชื่อ\s*-?\s*สกุล|เข้างาน|ออกงาน|รายงาน|หน้า\s*\d|พิมพ์เมื่อ|วันที่พิมพ์')
# คำสถานะที่อาจติดมาท้ายแถว (ต้องตัดออกก่อนนำไปเป็นชื่อคนถัดไป)
# ตัดเฉพาะคำที่อยู่ "ต้น" ข้อความ เพื่อไม่ให้ไปตัดชื่อคน เช่น "สายใจ", "ลาวัลย์"
STATUS_PREFIX_RE = re.compile(r'^\s*(?:(?:ปกติ|มาสาย|สาย|ขาดงาน|ขาด|วันหยุด|ลืมสแกน\S*|ลาป่วย|ลากิจ|ลาพักผ่อน|-)\s+)+')


def _to_date(m):
    d, mth, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if y < 100:                       # ปี 2 หลัก: >= 40 ถือเป็น พ.ศ. (25xx) ไม่งั้นเป็น ค.ศ. (20xx)
        y = 2500 + y if y >= 40 else 2000 + y
    if y > 2400:
        y -= 543
    return pd.Timestamp(year=y, month=mth, day=d)


def _split_dept(text):
    text = text.strip()
    for dv in DEPT_VARIANTS:
        if text.startswith(dv):
            return dv, text[len(dv):].strip()
    return "", text


def _classify_times(times):
    check_in, check_out = None, None
    if len(times) >= 3:
        check_in, check_out = times[0], times[1]
    elif len(times) == 2:
        h1 = int(times[0].split(':')[0])
        if times[1] == "01:00":
            if h1 >= 12: check_out, check_in = times[0], None
            else: check_in, check_out = times[0], None
        else: check_in, check_out = times[0], times[1]
    elif len(times) == 1:
        if int(times[0].split(':')[0]) >= 12: check_out = times[0]
        else: check_in = times[0]
    return check_in, check_out


def _get_status(dt_obj, check_in, check_out, late_threshold, exc_week, exc_th, extra_list):
    if not check_in and not check_out:
        return "วันหยุด" if is_it_holiday(dt_obj, exc_week, exc_th, extra_list) else "ขาดงาน"
    if not check_in:
        return "ลืมสแกนนิ้วเข้า"
    if not check_out:
        return "ลืมสแกนนิ้วออก"
    try:
        in_h, in_m = map(int, check_in.split(':'))
        out_h, out_m = map(int, check_out.split(':'))
        in_mins = in_h * 60 + in_m
        out_mins = out_h * 60 + out_m
        # สแกนออกช่วงเที่ยงหรือบ่ายก่อนเวลาเลิกงาน (ก่อน 15:30 น.) ถือว่าขาดงานครึ่งบ่าย
        if out_mins < 15 * 60 + 30:
            return "ขาดงานครึ่งบ่าย"
        # สแกนเข้าช่วงเที่ยง (11:30 - 14:00 น.) และสแกนออกช่วงเย็นตามปกติ ถือว่าขาดงานครึ่งเช้า
        if 11 * 60 + 30 <= in_mins <= 14 * 60:
            return "ขาดงานครึ่งเช้า"
        return "สาย" if in_mins >= late_threshold else "ปกติ"
    except Exception:
        return "ปกติ"


def parse_lines(lines, late_threshold, exc_week, exc_th, extra_list, source=""):
    """แปลงบรรทัดข้อความจาก PDF เป็น records

    คืนค่า (records, skipped) โดย skipped คือบรรทัดที่อ่านไม่ได้ เพื่อแสดงให้ผู้ใช้ตรวจสอบ
    """
    records, skipped = [], []
    current_dept, current_name = "", ""
    pending = []          # บรรทัดข้อความ (ไม่มีวันที่) ที่อาจเป็นชื่อ/ส่วนหนึ่งของชื่อ

    for raw in lines:
        ln = normalize_whitespace(raw)
        if not ln:
            continue
        matches = list(DATE_RE.finditer(ln))

        # ---------- บรรทัดที่ไม่มีวันที่ ----------
        if not matches:
            if HEADER_RE.search(ln):
                pending = []
                continue
            dept_only, rest = _split_dept(ln)
            if dept_only and not rest:          # บรรทัดหัวกลุ่ม เช่น "ข้าราชการ"
                current_dept = dept_only
                pending = []
                continue
            if THAI_RE.search(ln) and not TIME_RE.search(ln):
                pending.append(ln)
            elif TIME_RE.search(ln):
                skipped.append(f"{source}: {raw}")
            continue

        # ---------- บรรทัดที่มีวันที่ (อาจมีหลายรายการถ้า PDF รวมแถวติดกัน) ----------
        carry_left = ln[:matches[0].start()]
        for i, m in enumerate(matches):
            seg_end = matches[i + 1].start() if i + 1 < len(matches) else len(ln)
            seg = ln[m.end():seg_end]
            time_matches = list(TIME_RE.finditer(seg))
            times = [t.group(0) for t in time_matches]

            # ข้อความหลังเวลาตัวสุดท้าย อาจเป็นชื่อของรายการถัดไป
            tail = seg[time_matches[-1].end():] if time_matches else seg
            next_left = STATUS_PREFIX_RE.sub('', tail + ' ').strip() if i + 1 < len(matches) else ""

            left = re.sub(r'^\d+[\.\s]+', '', carry_left.strip()).strip()   # ตัดเลขลำดับหน้าแถว
            carry_left = next_left

            dept, name = _split_dept(left)
            name = normalize_whitespace(name)

            # ชื่อถูกตัดขึ้นบรรทัดก่อนหน้า เช่น "นางสาวสมหญิง" / "ใจดี 01/10/2569 ..."
            if pending:
                pend_dept, pend_name = _split_dept(" ".join(pending))
                if not name:
                    # บรรทัดที่ไม่มีคำนำหน้าชื่อ อาจเป็นนามสกุลที่ล้นมาจากแถวก่อน -> ใช้ชื่อเดิม
                    name = pend_name if (TITLE_RE.match(pend_name) or not current_name) else current_name
                elif TITLE_RE.match(pend_name) and not TITLE_RE.match(name):
                    name = f"{pend_name} {name}".strip()
                dept = dept or pend_dept
                pending = []

            # ไม่มีชื่อในแถว -> ใช้ชื่อล่าสุด (รายงานแบบพิมพ์ชื่อครั้งเดียวต่อคน)
            if not name:
                name = current_name
                dept = dept or current_dept
            dept = dept or current_dept

            if not name:
                skipped.append(f"{source}: (ไม่พบชื่อ) {raw}")
                continue

            try:
                dt_obj = _to_date(m)
            except Exception:
                skipped.append(f"{source}: (วันที่ผิดรูปแบบ) {raw}")
                continue

            current_name, current_dept = name, dept
            check_in, check_out = _classify_times(times)
            status = _get_status(dt_obj, check_in, check_out, late_threshold, exc_week, exc_th, extra_list)
            records.append({'แผนก': dept, 'ชื่อ-สกุล': name, 'วันที่': dt_obj,
                            'เข้างาน': check_in, 'ออกงาน': check_out, 'สถานะ': status})
    return records, skipped


def is_table_page(page):
    """ตรวจสอบว่าหน้านี้เป็นตารางมาตรฐานที่มีหัวคอลัมน์ แผนก/ชื่อ-สกุล/วันที่/เวลา หรือไม่"""
    header_chars = [c for c in page.chars if c['top'] < 60]
    sample = "".join(c['text'] for c in header_chars)
    return ('แผนก' in sample and ('ชื่อ' in sample or 'วัน' in sample or 'เข้างาน' in sample))

def parse_table_page(page, late_threshold, exc_week, exc_th, extra_list, source=""):
    """อ่านข้อมูลหน้าตารางโดยตรงจากพิกัดตัวอักษรเพื่อป้องกันปัญหาสระลอย/ข้อความเพี้ยนจาก pdfplumber extract_text"""
    col_bounds = [
        ('แผนก', 0, 120),
        ('ชื่อ-สกุล', 120, 288),
        ('วันที่', 288, 355),
        ('เข้างาน', 355, 415),
        ('ออกงาน', 415, 475),
        ('สาย', 475, 595),
    ]
    chars_by_top = {}
    for c in page.chars:
        top = c['top']
        found = None
        for rt in chars_by_top:
            if abs(rt - top) < 4:
                found = rt; break
        if found is None: chars_by_top[top] = [c]
        else: chars_by_top[found].append(c)
        
    records, skipped = [], []
    for rt in sorted(chars_by_top.keys()):
        row_chars = chars_by_top[rt]
        sample = "".join(c['text'] for c in row_chars)
        if 'ชื่อ-สกุล' in sample or 'Page ' in sample or ('วันที่' in sample and 'เข้างาน' in sample):
            continue
            
        raw_data = {}
        for col_name, x0, x1 in col_bounds:
            cell_chars = [c for c in row_chars if x0 <= c['x0'] < x1]
            raw_data[col_name] = normalize_whitespace("".join(c['text'] for c in cell_chars))
            
        dept = normalize_thai_chars(raw_data.get('แผนก', ''))
        name = normalize_thai_chars(raw_data.get('ชื่อ-สกุล', ''))
        date_str = raw_data.get('วันที่', '')
        cin = raw_data.get('เข้างาน', '')
        cout = raw_data.get('ออกงาน', '')
        
        date_match = re.search(r'(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4}|\d{2})', date_str)
        if not date_match or not name:
            if any(c['text'].strip() for c in row_chars):
                skipped.append(f"{source}: {sample}")
            continue
            
        try:
            d, m, y = int(date_match.group(1)), int(date_match.group(2)), int(date_match.group(3))
            if y < 100: y = 2500 + y if y >= 40 else 2000 + y
            if y > 2400: y -= 543
            dt_obj = pd.Timestamp(year=y, month=m, day=d)
        except Exception:
            skipped.append(f"{source}: (วันที่ผิดรูปแบบ) {sample}")
            continue
        
        time_cin = re.search(r'([01]?\d|2[0-3]):[0-5]\d', cin)
        time_cout = re.search(r'([01]?\d|2[0-3]):[0-5]\d', cout)
        check_in = time_cin.group(0) if time_cin else None
        check_out = time_cout.group(0) if time_cout else None
        
        status = _get_status(dt_obj, check_in, check_out, late_threshold, exc_week, exc_th, extra_list)
        records.append({
            'แผนก': dept,
            'ชื่อ-สกุล': name,
            'วันที่': dt_obj,
            'เข้างาน': check_in,
            'ออกงาน': check_out,
            'สถานะ': status
        })
    return records, skipped

def process_pdf(file, late_threshold, exc_week, exc_th, extra_list):
    """คืนค่า (DataFrame หรือ None, รายการบรรทัดที่ข้าม)"""
    fname = getattr(file, 'name', str(file))
    all_records, all_skipped = [], []
    try:
        with pdfplumber.open(file) as pdf:
            for pageno, page in enumerate(pdf.pages, start=1):
                page_src = f"{fname} หน้า {pageno}"
                if is_table_page(page):
                    recs, skipped = parse_table_page(page, late_threshold, exc_week, exc_th, extra_list, source=page_src)
                else:
                    text = page.extract_text(x_tolerance=2, y_tolerance=2)
                    if not text:
                        all_skipped.append(f"{page_src}: ไม่พบข้อความ (อาจเป็น PDF แบบสแกนภาพ)")
                        continue
                    lines = [ln.strip() for ln in text.split('\n') if ln.strip()]
                    recs, skipped = parse_lines(lines, late_threshold, exc_week, exc_th, extra_list, source=page_src)
                all_records.extend(recs)
                all_skipped.extend(skipped)
    except Exception as e:
        st.error(f"อ่านไฟล์ {fname} ไม่สำเร็จ: {e}")
    return (pd.DataFrame(all_records) if all_records else None), all_skipped

def format_df_for_display(df: pd.DataFrame) -> pd.DataFrame:
    """จัดรูปแบบคอลัมน์วันที่ให้อ่านง่ายเป็น วัน/เดือน/ปี แทน Timestamp 00:00:00"""
    df_disp = df.copy()
    if 'วันที่' in df_disp.columns:
        if pd.api.types.is_datetime64_any_dtype(df_disp['วันที่']):
            df_disp['วันที่'] = df_disp['วันที่'].dt.strftime('%d/%m/%Y')
    return df_disp

TITLES_PATTERN = r'^(ว่าที่ร้อยตรีหญิง|ว่าที่ร้อยตรี|ว่าที่ร้อยโท|ว่าที่เรือตรี|ว่าที่ร\.ต\.หญิง|ว่าที่ร\.ต\.|ว่าที่|พันตรีหญิง|พันตรี|พ\.ต\.หญิง|พ\.ต\.|จ่าเอก|จ\.อ\.|นางสาว|นาง|นาย|น\.ส\.|ดร\.|จ\.ส\.|พ\.จ\.|ร\.ต\.|ร\.ท\.|ร\.อ\.|น\.ต\.|น\.ท\.|น\.อ\.|พ\.ท\.|พ\.อ\.|ด\.ต\.|ส\.ต\.|ส\.อ\.|จ\.ต\.|จ\.ท\.)\s*'

def normalize_thai_chars(s: str) -> str:
    """จัดระเบียบตัวอักษรไทย แปลงสระอำที่แยกสระให้เป็นรูปสมบูรณ์ และจัดระเบียบช่องว่าง"""
    if not s:
        return ""
    s = str(s).strip()
    s = re.sub(r'([ก-ฮ])\s+า', r'\1ำ', s)
    s = re.sub(r'([ก-ฮ])\s+้า', r'\1้ำ', s)
    s = s.replace('ลูกจ้างประจา', 'ลูกจ้างประจำ')
    s = re.sub(r'\s+', ' ', s).strip()
    return s

def clean_core_key(name: str) -> str:
    """ตัดคำนำหน้ายศ/คำนำหน้านามและช่องว่าง เพื่อใช้เป็นคีย์หลักในการเปรียบเทียบชื่อ"""
    n = re.sub(r'\s+', '', str(name))
    n = re.sub(TITLES_PATTERN, '', n)
    return n

@st.cache_data
def get_canonical_roster_data() -> list:
    """โหลดบัญชีรายชื่อบุคลากรทางการศูนย์ฝึกพาณิชย์นาวีจาก canonical_personnel.json"""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        os.path.join(base_dir, "..", "canonical_personnel.json"),
        os.path.join(base_dir, "canonical_personnel.json"),
        os.path.join(os.getcwd(), "canonical_personnel.json"),
    ]
    for p in candidates:
        if os.path.exists(p):
            try:
                with open(p, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
    return []

def get_active_roster() -> list:
    """ดึงข้อมูลบัญชีรายชื่อบุคลากร (รวมข้อมูลตั้งต้นและข้อมูลเพิ่มเติมจากตารางวันลา)"""
    if 'active_roster' not in st.session_state:
        st.session_state.active_roster = get_canonical_roster_data().copy()
    return st.session_state.active_roster

def enrich_canonical_roster_from_leave(leave_df: pd.DataFrame):
    """เพิ่มชื่อบุคลากรจากตารางวันลาเข้าสู่ active roster หากยังไม่มี"""
    if leave_df is None or leave_df.empty or 'Name' not in leave_df.columns:
        return
    active_roster = get_active_roster()
    existing_keys = {clean_core_key(p['name']) for p in active_roster}
    added = False
    for name in leave_df['Name'].dropna().unique():
        norm_n = normalize_thai_chars(name)
        k = clean_core_key(norm_n)
        if k and k not in existing_keys:
            existing_keys.add(k)
            active_roster.append({'dept': '', 'name': norm_n})
            added = True
    if added:
        st.session_state.active_roster = active_roster

def repair_truncated_name(name: str, dept: str = "") -> tuple:
    """กู้คืนชื่อ-สกุลที่ถูกตัดทอน (Truncated) จากเครื่องสแกนนิ้วให้สมบูรณ์ โดยเทียบกับบัญชีรายชื่อทางการ"""
    norm_name = normalize_thai_chars(name)
    norm_dept = normalize_thai_chars(dept)
    if not norm_name:
        return norm_name, norm_dept

    roster = get_active_roster()
    if not roster:
        return norm_name, norm_dept

    core_pdf = clean_core_key(norm_name)
    raw_pdf = re.sub(r'\s+', '', norm_name)

    # 1. ตรงกันแบบสมบูรณ์ (ตัดช่องว่าง)
    for c in roster:
        if re.sub(r'\s+', '', c['name']) == raw_pdf:
            return c['name'], c['dept'] or norm_dept

    # 2. ตรงกันระดับชื่อหลัก (ตัดคำนำหน้าและช่องว่าง)
    for c in roster:
        if clean_core_key(c['name']) == core_pdf:
            return c['name'], c['dept'] or norm_dept

    # 3. ชื่อในฐานข้อมูลเริ่มต้นด้วยชื่อจาก PDF ที่ถูกตัดทอน (Prefix Match)
    prefix_candidates = []
    for c in roster:
        c_k = clean_core_key(c['name'])
        if c_k.startswith(core_pdf) and len(core_pdf) >= 4:
            prefix_candidates.append(c)

    if len(prefix_candidates) == 1:
        return prefix_candidates[0]['name'], prefix_candidates[0]['dept'] or norm_dept
    elif len(prefix_candidates) > 1:
        if norm_dept:
            same_dept = [c for c in prefix_candidates if norm_dept in c.get('dept', '') or c.get('dept', '') in norm_dept]
            if same_dept:
                return same_dept[0]['name'], same_dept[0]['dept'] or norm_dept
        return prefix_candidates[0]['name'], prefix_candidates[0]['dept'] or norm_dept

    # 4. Fuzzy Match ด้วยความแม่นยำสูง (>= 80%)
    best_c = None
    best_score = 0
    for c in roster:
        c_k = clean_core_key(c['name'])
        score = SequenceMatcher(None, core_pdf, c_k).ratio()
        if score > best_score:
            best_score = score
            best_c = c

    if best_score >= 0.80 and best_c:
        return best_c['name'], best_c['dept'] or norm_dept

    return norm_name, norm_dept

def unify_names(df: pd.DataFrame) -> pd.DataFrame:
    """ทำความสะอาดชื่อ-สกุล กู้คืนนามสกุลที่ถูกตัดทอน (Truncated) และจัดหมวดหมู่ฝ่ายให้ถูกต้อง"""
    if df.empty:
        return df

    df = df.copy()
    df['ชื่อ-สกุล'] = df['ชื่อ-สกุล'].apply(normalize_thai_chars)
    df['แผนก'] = df['แผนก'].apply(normalize_thai_chars)

    repaired_cache = {}
    def _repair_row(row):
        n = row['ชื่อ-สกุล']
        d = row['แผนก']
        k = (n, d)
        if k not in repaired_cache:
            rep_name, rep_dept = repair_truncated_name(n, d)
            repaired_cache[k] = (rep_name, rep_dept)
        return repaired_cache[k]

    rep_results = df.apply(_repair_row, axis=1)
    df['ชื่อ-สกุล'] = [r[0] for r in rep_results]
    df['แผนก'] = [r[1] for r in rep_results]

    # เติมแผนกที่ว่างจากแถวอื่นของคนเดียวกัน
    non_empty_dept = df[df['แผนก'] != '']
    if not non_empty_dept.empty:
        dept_map = non_empty_dept.groupby('ชื่อ-สกุล')['แผนก'].agg(lambda s: s.value_counts().index[0]).to_dict()
        df['แผนก'] = df.apply(lambda r: r['แผนก'] or dept_map.get(r['ชื่อ-สกุล'], ''), axis=1)

    return df

# -------------------------
# Helper Functions: เทียบข้อมูลวันลา
# -------------------------
def clean_name(name):
    if not name or pd.isna(name): return ""
    n = normalize_thai_chars(str(name))
    n = re.sub(r'\s+', '', n)
    n = re.sub(TITLES_PATTERN, '', n)
    return n

def split_first_last(name):
    cleaned = normalize_thai_chars(str(name))
    cleaned = re.sub(TITLES_PATTERN, '', cleaned).strip()
    parts = cleaned.split()
    first = parts[0] if parts else ''
    last = ''.join(parts[1:]) if len(parts) > 1 else ''
    return first, last

def is_name_match(n1, n2):
    c1, c2 = clean_name(n1), clean_name(n2)
    if c1 == c2: return True
    if not c1 or not c2: return False
    
    f1, l1 = split_first_last(n1)
    f2, l2 = split_first_last(n2)
    
    if f1 and f2:
        if f1 == f2 or SequenceMatcher(None, f1, f2).ratio() > 0.85:
            if l1 and l2:
                if l1.startswith(l2) or l2.startswith(l1):
                    return True
                if SequenceMatcher(None, l1, l2).ratio() > 0.8:
                    return True
            elif not l1 or not l2:
                if c1.startswith(c2) or c2.startswith(c1):
                    return True
    return False

def shorten_status_for_pivot(val):
    """แปลงสถานะเป็นสัญลักษณ์มาตรฐานของรายงานศูนย์ฝึกพาณิชย์นาวี (✓, สาย, เข้า, ออก, x, -, ป, พ, ก)"""
    if not val or pd.isna(val): return ""
    v = str(val).strip()
    if v in ["ปกติ", "✓"]:
        return "✓"
    if v in ["สาย", "สาย (มีใบขออนุญาต)", "สาย(อนุญาต)"]:
        return "สาย"
    if v in ["ลืมสแกนนิ้วเข้า", "ลืมเข้า", "เข้า"]:
        return "เข้า"
    if v in ["ลืมสแกนนิ้วออก", "ลืมออก", "ออก"]:
        return "ออก"
    if v in ["ขาดงานครึ่งบ่าย", "ขาด(บ)"]:
        return "ขาด(บ)"
    if v in ["ขาดงานครึ่งเช้า", "ขาด(ช)"]:
        return "ขาด(ช)"
    if v in ["ขาดงาน", "ขาด", "x", "X"]:
        return "x"
    if v in ["วันหยุด", "หยุด", "-", " -"]:
        return " -"
    if "ราชการ" in v:
        return "ราชการ"
    short_map = {"ลาป่วย": "ป", "ลาพักผ่อน": "พ", "ลากิจ": "ก"}
    for full_t, code in short_map.items():
        if full_t in v:
            if "ครึ่งบ่าย" in v or "(บ)" in v:
                return f"{code}(บ)"
            if "ครึ่งเช้า" in v or "(ช)" in v:
                return f"{code}(ช)"
            if "ครึ่ง" in v:
                return f"{code}(ครึ่ง)"
            return code
    return v

def cross_check_attendance_with_leave(attendance_df: pd.DataFrame, leave_df: pd.DataFrame) -> pd.DataFrame:
    df = attendance_df.copy()
    if df.empty or leave_df is None or leave_df.empty:
        if 'มีใบลา' not in df.columns:
            df['มีใบลา'] = False
            df['ประเภทการลา'] = '-'
            df['หมายเหตุการลา'] = ''
        return df
    
    leave_indexed = {}
    for idx, lrow in leave_df.iterrows():
        key = (str(lrow['Month']).strip(), int(lrow['Day']))
        if key not in leave_indexed:
            leave_indexed[key] = []
        leave_indexed[key].append(lrow)
        
    has_leave_col = []
    leave_type_col = []
    leave_remark_col = []
    updated_status_col = []
    
    for _, arow in df.iterrows():
        dt = arow['วันที่']
        m_name = month_name_thai(dt.month)
        day_num = dt.day
        att_name = arow['ชื่อ-สกุล']
        cur_status = arow['สถานะ']
        
        candidates = leave_indexed.get((m_name, day_num), [])
        matched_leave = None
        for cand in candidates:
            if is_name_match(cand['Name'], att_name):
                matched_leave = cand
                break
                
        if matched_leave is not None:
            l_type = matched_leave['Type']
            is_half = bool(matched_leave.get('IsHalf', False))
            remark = str(matched_leave.get('Remarks', '') or '')
            
            has_leave_col.append(True)
            leave_type_col.append(l_type)
            leave_remark_col.append(remark)
            
            half_tag = "ครึ่งวัน"
            if "บ่าย" in remark or "(บ)" in remark or cur_status == 'ขาดงานครึ่งบ่าย' or (cur_status == 'ลืมสแกนนิ้วออก' and is_half):
                half_tag = "ครึ่งบ่าย"
            elif "เช้า" in remark or "(ช)" in remark or cur_status == 'ขาดงานครึ่งเช้า' or (cur_status == 'ลืมสแกนนิ้วเข้า' and is_half):
                half_tag = "ครึ่งเช้า"

            if cur_status == 'ขาดงาน':
                if is_half:
                    updated_status_col.append(f"{l_type} ({half_tag})")
                else:
                    updated_status_col.append(f"{l_type} (มีใบลา)")
            elif cur_status in ['ขาดงานครึ่งบ่าย', 'ขาดงานครึ่งเช้า']:
                if l_type == 'สาย':
                    updated_status_col.append('สาย (มีใบขออนุญาต)')
                else:
                    updated_status_col.append(f"{l_type} ({half_tag})")
            elif cur_status in ['ลืมสแกนนิ้วเข้า', 'ลืมสแกนนิ้วออก']:
                if l_type == 'สาย':
                    updated_status_col.append('สาย (มีใบขออนุญาต)')
                elif is_half:
                    updated_status_col.append(f"{l_type} ({half_tag})")
                else:
                    updated_status_col.append(f"{l_type} (มีใบลา)")
            elif cur_status == 'สาย':
                if l_type == 'สาย':
                    updated_status_col.append('สาย (มีใบขออนุญาต)')
                elif is_half:
                    updated_status_col.append(f"{l_type} ({half_tag})")
                else:
                    updated_status_col.append(f"{l_type} (มีใบลา)")
            else:
                updated_status_col.append(cur_status)
        else:
            has_leave_col.append(False)
            leave_type_col.append('-')
            leave_remark_col.append('')
            updated_status_col.append(cur_status)
            
    df['มีใบลา'] = has_leave_col
    df['ประเภทการลา'] = leave_type_col
    df['หมายเหตุการลา'] = leave_remark_col
    df['สถานะ'] = updated_status_col
    return df

def cross_check_to_excel_bytes(df_all: pd.DataFrame, dept_label: str = ""):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        base_font = 'TH SarabunPSK'
        title_format = workbook.add_format({'bold': True, 'font_name': base_font, 'font_size': 18, 'align': 'center'})
        header_format = workbook.add_format({'bold': True, 'font_name': base_font, 'font_size': 14, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#D9D9D9'})
        data_center_format = workbook.add_format({'font_name': base_font, 'font_size': 14, 'align': 'center', 'valign': 'vcenter', 'border': 1})
        data_left_format = workbook.add_format({'font_name': base_font, 'font_size': 14, 'align': 'left', 'valign': 'vcenter', 'border': 1})
        
        sub_dept = f" ({dept_label})" if dept_label and dept_label != "ทั้งหมด" else ""
        
        # 1. Approved Leave Sheet
        leave_df = df_all[df_all['มีใบลา'] == True].copy() if 'มีใบลา' in df_all.columns else pd.DataFrame()
        ws_leave = workbook.add_worksheet('ผู้มีใบลาอนุมัติ')
        ws_leave.merge_range('A1:G1', f'รายงานรายชื่อผู้มีใบลาอนุมัติถูกต้อง{sub_dept}', title_format)
        headers_l = ['ลำดับ', 'วันที่', 'แผนก', 'ชื่อ - สกุล', 'ประเภทการลา', 'สถานะ', 'หมายเหตุ']
        for i, h in enumerate(headers_l):
            ws_leave.write(2, i, h, header_format)
        row_idx = 3
        for idx, r in enumerate(leave_df.itertuples(), 1):
            d_str = r.วันที่.strftime('%d/%m/%Y') if hasattr(r.วันที่, 'strftime') else str(r.วันที่)
            ws_leave.write(row_idx, 0, idx, data_center_format)
            ws_leave.write(row_idx, 1, d_str, data_center_format)
            ws_leave.write(row_idx, 2, getattr(r, 'แผนก', ''), data_center_format)
            ws_leave.write(row_idx, 3, getattr(r, '_2', getattr(r, 'ชื่อ_สกุล', '')), data_left_format)
            ws_leave.write(row_idx, 4, getattr(r, 'ประเภทการลา', '-'), data_center_format)
            ws_leave.write(row_idx, 5, getattr(r, 'สถานะ', ''), data_center_format)
            ws_leave.write(row_idx, 6, getattr(r, 'หมายเหตุการลา', ''), data_left_format)
            row_idx += 1
        ws_leave.set_column('A:A', 8)
        ws_leave.set_column('B:C', 16)
        ws_leave.set_column('D:D', 30)
        ws_leave.set_column('E:F', 18)
        ws_leave.set_column('G:G', 25)
        
        # 2. Unauthorized Absent Sheet
        absent_df = df_all[df_all['สถานะ'].isin(['ขาดงาน', 'ขาดงานครึ่งบ่าย', 'ขาดงานครึ่งเช้า'])].copy()
        ws_abs = workbook.add_worksheet('ขาดงานจริง_ไม่มีใบลา')
        ws_abs.merge_range('A1:E1', f'รายงานรายชื่อผู้ขาดงานจริง (ไม่มีใบลา){sub_dept}', title_format)
        headers_a = ['ลำดับ', 'วันที่', 'แผนก', 'ชื่อ - สกุล', 'สถานะ']
        for i, h in enumerate(headers_a):
            ws_abs.write(2, i, h, header_format)
        row_idx = 3
        for idx, r in enumerate(absent_df.itertuples(), 1):
            d_str = r.วันที่.strftime('%d/%m/%Y') if hasattr(r.วันที่, 'strftime') else str(r.วันที่)
            ws_abs.write(row_idx, 0, idx, data_center_format)
            ws_abs.write(row_idx, 1, d_str, data_center_format)
            ws_abs.write(row_idx, 2, getattr(r, 'แผนก', ''), data_center_format)
            ws_abs.write(row_idx, 3, getattr(r, '_2', getattr(r, 'ชื่อ_สกุล', '')), data_left_format)
            ws_abs.write(row_idx, 4, getattr(r, 'สถานะ', ''), data_center_format)
            row_idx += 1
        ws_abs.set_column('A:A', 8)
        ws_abs.set_column('B:C', 16)
        ws_abs.set_column('D:D', 30)
        ws_abs.set_column('E:E', 18)

    return output.getvalue()

# -------------------------
# Styling DataFrame
# -------------------------
def style_status_table(df: pd.DataFrame):
    def highlight(val):
        val_str = str(val)
        if "ปกติ" in val_str:
            return "background-color: #C6EFCE; color: #006100"
        elif "สาย (มีใบขออนุญาต)" in val_str or "สาย(อนุญาต)" in val_str:
            return "background-color: #FFF3CD; color: #856404"
        elif "สาย" in val_str:
            return "background-color: #FFEB9C; color: #9C6500"
        elif "ขาดงาน" in val_str or val_str == "ขาด":
            return "background-color: #FFC7CE; color: #9C0006"
        elif "วันหยุด" in val_str or val_str == "หยุด":
            return "background-color: #E7E6E6; color: #595959"
        elif "ลืมสแกน" in val_str or "ลืมเข้า" in val_str or "ลืมออก" in val_str:
            return "background-color: #FCE4D6; color: #9C0006"
        elif "ครึ่งวัน" in val_str or "ครึ่ง" in val_str:
            return "background-color: #E2D9F3; color: #491217"
        elif "ลา" in val_str or "พักผ่อน" in val_str:
            return "background-color: #D1ECF1; color: #0C5460"
        return ""
    subset = ['สถานะ'] if "สถานะ" in df.columns else None
    return df.style.map(highlight, subset=subset) if hasattr(df.style, "map") else df.style.applymap(highlight, subset=subset)

# -------------------------
# UI Header Banner
# -------------------------
if 'master_df' in st.session_state:
    _m = st.session_state.master_df
    status_badge_html = f'<div class="status-badge-ready">🟢 ข้อมูลพร้อม: <b>{len(_m)}</b> รายการ (<b>{_m["ชื่อ-สกุล"].nunique()}</b> คน)</div>'
else:
    status_badge_html = '<div class="status-badge-empty">⚪ รอนำเข้าข้อมูล (เริ่มที่แท็บ 1)</div>'

st.markdown(f"""
<div class="att-header-container">
    <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 12px;">
        <div>
            <div class="att-pill-badge">⏱️ ATTENDANCE & LEAVE INTELLIGENCE</div>
            <h2 class="att-header-title">ระบบบันทึกเวลาและการปฏิบัติงาน</h2>
            <p class="att-header-sub">วิเคราะห์เวลาสแกนนิ้วมือ ตรวจจับเวลาสาย ลืมสแกน และตรวจเทียบวันลาอัตโนมัติ</p>
        </div>
        <div>
            {status_badge_html}
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# -------------------------
# UI Tabs
# -------------------------
tab1, tab2, tab3, tab4, tab5 = st.tabs(["📤 นำเข้าข้อมูล", "📅 ภาพรวมรายวัน", "📈 สถิติรายเดือน", "👤 ข้อมูลรายบุคคล", "🏖️ ตรวจสอบเทียบวันลา"])

with tab1:
    st.header("📤 นำเข้าข้อมูล")
    col_u1, col_u2 = st.columns([1, 1])
    with col_u1:
        st.subheader("1. 📄 ข้อมูลสแกนนิ้ว (PDF)")
        uploaded_files = st.file_uploader("ลากไฟล์ PDF สแกนนิ้วมาวางตรงนี้", type="pdf", accept_multiple_files=True)
    with col_u2:
        st.subheader("2. 📊 ตารางวันลา (Excel)")
        uploaded_leave = st.file_uploader("อัปโหลดไฟล์ตารางวันลา เพื่อตรวจเทียบอัตโนมัติ (.xlsx, .csv)", type=["xlsx", "csv"], key="attendance_leave_uploader")
        if uploaded_leave is not None:
            try:
                leave_df_parsed = load_data(uploaded_leave)
                st.session_state.leave_df = leave_df_parsed
                enrich_canonical_roster_from_leave(leave_df_parsed)
                st.success(f"✅ โหลดข้อมูลตารางวันลาสำเร็จ: {len(leave_df_parsed)} รายการ")
                if 'master_df' in st.session_state and st.session_state.master_df is not None:
                    st.session_state.master_df = unify_names(st.session_state.master_df)
                    st.session_state.master_df = cross_check_attendance_with_leave(st.session_state.master_df, leave_df_parsed)
                    st.info("🔄 อัปเดตการตรวจเทียบวันลากับข้อมูลสแกนนิ้วเรียบร้อยแล้ว!")
            except Exception as e:
                st.error(f"❌ อ่านไฟล์ตารางวันลาไม่สำเร็จ: {e}")
        elif 'leave_df' in st.session_state and st.session_state.leave_df is not None:
            st.info(f"💡 พบข้อมูลตารางวันลาในระบบแล้ว: {len(st.session_state.leave_df)} รายการ (พร้อมตรวจเทียบ)")

    if uploaded_files and st.button("🚀 ประมวลผลข้อมูล", type="primary"):
        all_data, all_skipped = [], []
        for f in uploaded_files:
            df, skipped = process_pdf(f, LATE_THRESHOLD_MINUTES, exclude_weekends, exclude_th_holidays, extra_holidays)
            all_skipped.extend(skipped)
            if df is not None:
                all_data.append(df)
                st.write(f"📄 {f.name}: พบ {len(df)} รายการ / {df['ชื่อ-สกุล'].nunique()} คน")
        if all_data:
            master_df = unify_names(pd.concat(all_data, ignore_index=True)).drop_duplicates(subset=['ชื่อ-สกุล', 'วันที่'], keep='last')
            
            # Cross check with leave data if available
            curr_leave_df = st.session_state.get('leave_df')
            master_df = cross_check_attendance_with_leave(master_df, curr_leave_df)
            st.session_state.master_df = master_df
            
            leave_note = " และตรวจเทียบกับตารางวันลาแล้ว" if curr_leave_df is not None and not curr_leave_df.empty else ""
            st.success(f"✅ ประมวลผลสำเร็จ{leave_note}! รวม {master_df['ชื่อ-สกุล'].nunique()} คน ({len(master_df)} บันทึก)")
            
            # แสดงสรุปสถิติจำนวนคนแยกตามฝ่าย
            dept_counts = master_df.drop_duplicates(subset=['ชื่อ-สกุล']).groupby(
                master_df['แผนก'].replace({'': 'ไม่ระบุฝ่าย'})
            )['ชื่อ-สกุล'].count().to_dict()
            if dept_counts:
                st.markdown("##### 👥 จำนวนบุคลากรจำแนกตามฝ่าย/กลุ่มงาน:")
                d_cols = st.columns(max(len(dept_counts), 1))
                for i, (dept_name, count) in enumerate(dept_counts.items()):
                    d_cols[i % len(d_cols)].metric(f"🏢 {dept_name}", f"{count} คน")
        else:
            st.error("❌ ไม่พบข้อมูล")
        if all_skipped:
            with st.expander(f"⚠️ บรรทัดที่อ่านไม่ได้ {len(all_skipped)} บรรทัด (กดเพื่อตรวจสอบ)"):
                st.text("\n".join(all_skipped))

if 'master_df' in st.session_state:
    master_df = st.session_state.master_df
    dept_options = ["ทั้งหมด"] + get_department_options(master_df)

    with tab2:
        c_date, c_dept = st.columns([1, 1])
        with c_date:
            dates = sorted(master_df['วันที่'].dt.date.unique(), reverse=True)
            sel_date = st.selectbox("📅 เลือกวันที่:", dates, key="daily_date")
        with c_dept:
            sel_dept_day = st.selectbox("🏢 กรองตามฝ่าย/กลุ่มงาน:", dept_options, key="daily_dept")

        df_day_all = master_df[master_df['วันที่'].dt.date == sel_date]
        df_day = filter_by_dept(df_day_all, sel_dept_day)

        c_status = df_day['สถานะ'].value_counts().to_dict()
        leave_cnt = sum(c for s, c in c_status.items() if ('ลา' in s or 'มีใบลา' in s or 'ครึ่ง' in s) and not s.startswith('ขาดงาน'))
        late_cnt = c_status.get("สาย", 0) + c_status.get("สาย (มีใบขออนุญาต)", 0)
        forget_cnt = c_status.get("ลืมสแกนนิ้วเข้า", 0) + c_status.get("ลืมสแกนนิ้วออก", 0)
        absent_cnt = sum(c for s, c in c_status.items() if s.startswith("ขาดงาน"))

        col1, col2, col3, col4, col5, col6 = st.columns(6)
        col1.metric("👥 ทั้งหมด", len(df_day))
        col2.metric("✅ ปกติ", c_status.get("ปกติ", 0))
        col3.metric("⚠️ สาย", late_cnt)
        col4.metric("🏖️ ลา (มีใบลา)", leave_cnt)
        col5.metric("❌ ขาด (ไม่มีใบลา)", absent_cnt)
        col6.metric("🔔 ลืมสแกน", forget_cnt)
        st.markdown("---")

        c_chart, c_leader = st.columns(2)
        with c_chart:
            pie_df = pd.DataFrame(list(c_status.items()), columns=['Status', 'Count'])
            color_map = {
                "ปกติ": "#28a745",
                "สาย": "#ffc107",
                "สาย (มีใบขออนุญาต)": "#ffe082",
                "ขาดงาน": "#dc3545",
                "ขาดงานครึ่งบ่าย": "#e35d6a",
                "ขาดงานครึ่งเช้า": "#ea868f",
                "ลืมสแกนนิ้วเข้า": "#fd7e14",
                "ลืมสแกนนิ้วออก": "#e65100",
                "ลาป่วย (มีใบลา)": "#17a2b8",
                "ลากิจ (มีใบลา)": "#6f42c1",
                "ลาพักผ่อน (มีใบลา)": "#20c997",
                "ลาป่วย (ครึ่งวัน)": "#6610f2",
                "ลากิจ (ครึ่งวัน)": "#845ef7",
                "ลาพักผ่อน (ครึ่งวัน)": "#38d9a9",
                "วันหยุด": "#6c757d"
            }
            fig = px.pie(pie_df, values='Count', names='Status', hole=0.4, color='Status', color_discrete_map=color_map)
            st.plotly_chart(fig, use_container_width=True)
        with c_leader:
            leader_title = "🏆 มาเช้า 5 อันดับแรก" + (f" ({sel_dept_day})" if sel_dept_day != "ทั้งหมด" else "")
            st.subheader(leader_title)
            df_early = df_day.dropna(subset=['เข้างาน']).sort_values('เข้างาน').head(5)
            if not df_early.empty:
                for i, r in enumerate(df_early.itertuples(), 1):
                    st.markdown(f"**{'🥇' if i==1 else '🥈' if i==2 else '🥉' if i==3 else '🏃'} อันดับ {i}** | {r._2} — ⏰ **{r.เข้างาน}**")
            else:
                st.info("ไม่มีข้อมูลเวลาเข้างานในกลุ่มที่เลือก")

        show_cols_day = ['แผนก', 'ชื่อ-สกุล', 'เข้างาน', 'ออกงาน', 'สถานะ']
        if 'ประเภทการลา' in df_day.columns and df_day['มีใบลา'].any():
            show_cols_day.extend(['ประเภทการลา', 'หมายเหตุการลา'])
        st.dataframe(style_status_table(format_df_for_display(df_day[show_cols_day])), use_container_width=True)
        
        dept_file_tag = f"_{sel_dept_day}" if sel_dept_day != "ทั้งหมด" else ""
        btn_daily_label = "📥 ดาวน์โหลดรายงานวันนี้ (Excel)" + (f" - {sel_dept_day}" if sel_dept_day != "ทั้งหมด" else "")
        st.download_button(btn_daily_label, daily_to_excel_bytes(df_day, sel_date, sel_dept_day), f"daily_{sel_date}{dept_file_tag}.xlsx")

    with tab3:
        sel_dept_month = st.selectbox("🏢 กรองตามฝ่าย/กลุ่มงาน:", dept_options, key="month_dept")
        df_month = filter_by_dept(master_df, sel_dept_month)

        if sel_dept_month == "ทั้งหมด":
            st.subheader("🏢 การมาสายแยกตามฝ่าย")
            dept_stat = pd.crosstab(master_df['แผนก'].replace({'': 'ไม่ระบุฝ่าย'}), master_df['สถานะ']).fillna(0)
            late_cols = [c for c in ['สาย', 'สาย (มีใบขออนุญาต)'] if c in dept_stat.columns]
            if late_cols:
                dept_stat['รวมสาย'] = dept_stat[late_cols].sum(axis=1)
                st.plotly_chart(px.bar(dept_stat.reset_index(), x='แผนก', y='รวมสาย', color='แผนก', text_auto=True), use_container_width=True)
        else:
            st.subheader(f"🏢 สถิติการมาสายในกลุ่ม: {sel_dept_month}")
            late_in_dept = df_month[df_month['สถานะ'].isin(['สาย', 'สาย (มีใบขออนุญาต)'])].groupby('ชื่อ-สกุล')['สถานะ'].count().reset_index()
            late_in_dept.columns = ['ชื่อ-สกุล', 'จำนวนครั้งที่สาย']
            if not late_in_dept.empty:
                st.plotly_chart(px.bar(late_in_dept.sort_values('จำนวนครั้งที่สาย', ascending=False).head(10), x='ชื่อ-สกุล', y='จำนวนครั้งที่สาย', color='จำนวนครั้งที่สาย', text_auto=True, color_continuous_scale="Reds"), use_container_width=True)
            else:
                st.info(f"✨ ยอดเยี่ยม! ไม่พบประวัติการมาสายในกลุ่ม {sel_dept_month}")

        # เพิ่มส่วนสรุปภาพรวมรายเดือน
        monthly_summary_title = "### 📊 สรุปภาพรวมรายเดือน (ลืมสแกน/ขาดงาน/วันลา)" + (f" - {sel_dept_month}" if sel_dept_month != "ทั้งหมด" else "")
        st.markdown(monthly_summary_title)
        m_forget_in_count = df_month[df_month['สถานะ'] == 'ลืมสแกนนิ้วเข้า']['ชื่อ-สกุล'].nunique()
        m_forget_out_count = df_month[df_month['สถานะ'] == 'ลืมสแกนนิ้วออก']['ชื่อ-สกุล'].nunique()
        m_absent_count = df_month[df_month['สถานะ'].isin(['ขาดงาน', 'ขาดงานครึ่งบ่าย', 'ขาดงานครึ่งเช้า'])]['ชื่อ-สกุล'].nunique()
        m_leave_count = df_month[df_month['มีใบลา'] == True]['ชื่อ-สกุล'].nunique() if 'มีใบลา' in df_month.columns else 0

        c_m1, c_m2, c_m3, c_m4 = st.columns(4)
        c_m1.metric("👤 ลืมสแกนเข้า", f"{m_forget_in_count} คน")
        c_m2.metric("👤 ลืมสแกนออก", f"{m_forget_out_count} คน")
        c_m3.metric("❌ ขาดงานจริง (ไม่มีใบลา)", f"{m_absent_count} คน")
        c_m4.metric("🏖️ มีใบลาถูกต้อง", f"{m_leave_count} คน")
        st.markdown("---")

        valid_m_dates = master_df['วันที่'].dropna()
        rep_month_th = month_name_thai(int(valid_m_dates.dt.month.mode()[0])) if not valid_m_dates.empty else ""
        rep_year_be = int(valid_m_dates.dt.year.mode()[0]) + 543 if not valid_m_dates.empty else ""

        st.markdown(f"### 📋 ตารางรายงานประจำเดือน {rep_month_th} {rep_year_be} (รูปแบบมาตรฐานศูนย์ฝึกพาณิชย์นาวี)")
        st.markdown(
            '<div style="display:flex; flex-wrap:wrap; gap:10px; background:#FFFDF9; border:1px solid #EAE1D3; border-radius:14px; padding:10px 16px; margin-bottom:14px; font-size:0.88rem;">'
            '<span><b>คำอธิบายสัญลักษณ์:</b></span>'
            '<span style="color:#1E7E34; font-weight:700;">✓ = ปกติ</span> | '
            '<span style="background:#FFEB9C; color:#9C6500; padding:1px 8px; border-radius:6px; font-weight:600;">สาย = มาสาย</span> | '
            '<span style="background:#FCE4D6; color:#C65911; padding:1px 8px; border-radius:6px; font-weight:600;">เข้า = ลืมสแกนเข้า</span> | '
            '<span style="background:#FCE4D6; color:#C65911; padding:1px 8px; border-radius:6px; font-weight:600;">ออก = ลืมสแกนออก</span> | '
            '<span style="background:#FFC7CE; color:#9C0006; padding:1px 8px; border-radius:6px; font-weight:600;">x / ขาด(บ) / ขาด(ช) = ขาดงาน / ขาดครึ่งบ่าย / ขาดครึ่งเช้า</span> | '
            '<span style="background:#95B3D7; color:#1F3864; padding:1px 8px; border-radius:6px; font-weight:600;">- = วันหยุดราชการ</span> | '
            '<span style="background:#D1ECF1; color:#0C5460; padding:1px 8px; border-radius:6px; font-weight:600;">ป / พ / ก = ลาป่วย / พักผ่อน / ลากิจ</span>'
            '</div>',
            unsafe_allow_html=True
        )

        if sel_dept_month == "ทั้งหมด":
            avail_depts = [d for d in REPORT_DEPT_ORDER if d in master_df['แผนก'].values]
            if not avail_depts:
                avail_depts = ["ทั้งหมด"]
            dept_tabs = st.tabs([f"🏛️ {d} ({master_df[master_df['แผนก']==d]['ชื่อ-สกุล'].nunique()} คน)" for d in avail_depts])
            for d_tab, d_name in zip(dept_tabs, avail_depts):
                with d_tab:
                    st.markdown(f"<div style='text-align:center; font-weight:700; font-size:1.15rem; margin-bottom:8px;'>{d_name}ศูนย์ฝึกพาณิชย์นาวี — ประจำเดือน {rep_month_th} {rep_year_be}</div>", unsafe_allow_html=True)
                    df_d_sub = filter_by_dept(master_df, d_name)
                    rep_df, hol_days, _, _ = build_monthly_report_table(df_d_sub, master_df, exclude_weekends, exclude_th_holidays, extra_holidays)
                    st.dataframe(style_monthly_report_table(rep_df, hol_days), use_container_width=True, hide_index=True)
        else:
            st.markdown(f"<div style='text-align:center; font-weight:700; font-size:1.15rem; margin-bottom:8px;'>{sel_dept_month}ศูนย์ฝึกพาณิชย์นาวี — ประจำเดือน {rep_month_th} {rep_year_be}</div>", unsafe_allow_html=True)
            rep_df, hol_days, _, _ = build_monthly_report_table(df_month, master_df, exclude_weekends, exclude_th_holidays, extra_holidays)
            st.dataframe(style_monthly_report_table(rep_df, hol_days), use_container_width=True, hide_index=True)

        dept_month_tag = f"_{sel_dept_month}" if sel_dept_month != "ทั้งหมด" else ""
        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            dl_sum_label = "📥 ดาวน์โหลดรายงานประจำเดือน (Excel รูปแบบมาตรฐาน 4 ชีต)" if sel_dept_month == "ทั้งหมด" else f"📥 ดาวน์โหลดรายงานประจำเดือน (Excel) - {sel_dept_month}"
            st.download_button(
                dl_sum_label,
                df_to_excel_bytes(pd.DataFrame(), master_df, sel_dept_month, exclude_weekends, exclude_th_holidays, extra_holidays),
                f"ประจำเดือน{rep_month_th}_{rep_year_be}{dept_month_tag}.xlsx",
                use_container_width=True
            )
        with col_dl2:
            dl_missing_label = "📥 ดาวน์โหลดรายงานลืมสแกน/ขาดงานจริง (Excel)" + (f" - {sel_dept_month}" if sel_dept_month != "ทั้งหมด" else "")
            st.download_button(dl_missing_label, missing_scan_to_excel_bytes(df_month, sel_dept_month), f"missing_scans_report{dept_month_tag}.xlsx", use_container_width=True)

    with tab4:
        c_pdept, c_pname = st.columns([1, 2])
        with c_pdept:
            sel_dept_person = st.selectbox("🏢 กรองตามฝ่าย/กลุ่มงาน:", dept_options, key="person_dept")

        df_person_pool = filter_by_dept(master_df, sel_dept_person)
        available_names = sorted(df_person_pool['ชื่อ-สกุล'].unique())

        with c_pname:
            count_label = f" ({len(available_names)} คน)" if sel_dept_person != "ทั้งหมด" else ""
            search_name = st.selectbox(f"🔍 ค้นหาชื่อพนักงาน{count_label}:", ["-- เลือกชื่อ --"] + available_names, key="person_name")

        if search_name != "-- เลือกชื่อ --":
            df_p = master_df[master_df['ชื่อ-สกุล'] == search_name].sort_values('วันที่')
            emp_dept_list = df_p['แผนก'].dropna().unique()
            emp_dept_str = emp_dept_list[0] if len(emp_dept_list) > 0 and emp_dept_list[0] else "ไม่ระบุฝ่าย"
            st.markdown(f"**👤 ชื่อ-สกุล:** `{search_name}` &nbsp;&nbsp;|&nbsp;&nbsp; **🏢 ฝ่าย/กลุ่มงาน:** `{emp_dept_str}`")

            ps = df_p['สถานะ'].value_counts().to_dict()
            unexcused_late = ps.get('สาย', 0)
            missed_in = ps.get('ลืมสแกนนิ้วเข้า', 0)
            missed_out = ps.get('ลืมสแกนนิ้วออก', 0)
            unauth_absent = sum(cnt for stat, cnt in ps.items() if stat.startswith('ขาดงาน'))
            leave_p_count = sum(cnt for stat, cnt in ps.items() if ('ลา' in stat or 'ครึ่ง' in stat) and not stat.startswith('ขาดงาน'))

            score = max(100 - (unexcused_late * 2) - ((missed_in + missed_out) * 3) - (unauth_absent * 5), 0)
            if score >= 90:
                score_badge_text = "🌟 ยอดเยี่ยม (วินัยดีเด่น)"
                score_cls = "score-high"
            elif score >= 75:
                score_badge_text = "👍 ปานกลาง (ปฏิบัติตามเกณฑ์)"
                score_cls = "score-med"
            else:
                score_badge_text = "⚠️ ควรปรับปรุง (มีรายการสาย/ขาด/ลืมสแกน)"
                score_cls = "score-low"
            st.markdown(f'<div class="score-badge {score_cls}">{score_badge_text}</div><div style="height: 12px;"></div>', unsafe_allow_html=True)

            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("⭐ คะแนนวินัย", f"{score}/100")
            c2.metric("⚠️ สาย", unexcused_late + ps.get('สาย (มีใบขออนุญาต)', 0))
            c3.metric("🏖️ ลา (มีใบลา)", leave_p_count)
            c4.metric("❌ ขาด (ไม่มีใบลา)", unauth_absent)
            c5.metric("🔔 ลืมสแกน", missed_in + missed_out)

            p_cols = ['วันที่', 'เข้างาน', 'ออกงาน', 'สถานะ']
            if 'ประเภทการลา' in df_p.columns:
                p_cols.extend(['ประเภทการลา', 'หมายเหตุการลา'])
            st.dataframe(style_status_table(format_df_for_display(df_p[p_cols])), use_container_width=True)

    with tab5:
        st.subheader("🏖️ ตรวจสอบเปรียบเทียบข้อมูลสแกนนิ้วกับวันลา")
        st.markdown("ระบบวิเคราะห์เปรียบเทียบระหว่างข้อมูลการสแกนนิ้วกับตารางการลา เพื่อคัดกรองว่าผู้ที่ขาดงานหรือลืมสแกนมีใบลาที่ได้รับอนุมัติหรือไม่")

        sel_dept_cross = st.selectbox("🏢 กรองตามฝ่าย/กลุ่มงาน:", dept_options, key="cross_dept")
        df_cross = filter_by_dept(master_df, sel_dept_cross)

        total_records = len(df_cross)
        leave_records = df_cross[df_cross['มีใบลา'] == True] if 'มีใบลา' in df_cross.columns else pd.DataFrame()
        absent_records = df_cross[df_cross['สถานะ'].isin(['ขาดงาน', 'ขาดงานครึ่งบ่าย', 'ขาดงานครึ่งเช้า'])]
        forget_records = df_cross[df_cross['สถานะ'].isin(['ลืมสแกนนิ้วเข้า', 'ลืมสแกนนิ้วออก'])]

        if len(absent_records) == 0:
            st.markdown("""
            <div style="background: #E8F5E9; border: 1.5px solid #A5D6A7; border-radius: 18px; padding: 14px 20px; display: flex; align-items: center; gap: 14px; margin-bottom: 18px;">
                <span style="font-size: 1.8rem;">🎉</span>
                <div>
                    <div style="font-weight: 700; color: #2E7D32; font-size: 1rem;">ยอดเยี่ยม! ไม่มีเจ้าหน้าที่ขาดงานโดยไม่มีใบลา</div>
                    <div style="font-size: 0.88rem; color: #43A047;">เจ้าหน้าที่ทุกคนมีข้อมูลเวลาสแกนนิ้ว หรือมีบันทึกการลาที่ได้รับการอนุมัติถูกต้องครบถ้วน</div>
                </div>
            </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown(f"""
            <div style="background: #FFEBEE; border: 1.5px solid #FFCDD2; border-radius: 18px; padding: 14px 20px; display: flex; align-items: center; gap: 14px; margin-bottom: 18px;">
                <span style="font-size: 1.8rem;">🚨</span>
                <div>
                    <div style="font-weight: 700; color: #C62828; font-size: 1rem;">พบเจ้าหน้าที่ขาดงานโดยไม่มีใบลา {len(absent_records)} รายการ</div>
                    <div style="font-size: 0.88rem; color: #E53935;">โปรดตรวจสอบหรือติดตามเอกสารการลากับเจ้าหน้าที่ เพื่อดำเนินการตามระเบียบงานบุคคลต่อไป</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

        ck1, ck2, ck3, ck4 = st.columns(4)
        ck1.metric("👥 รายการทั้งหมด", f"{total_records} รายการ")
        ck2.metric("✅ มีใบลาอนุมัติถูกต้อง", f"{len(leave_records)} รายการ")
        ck3.metric("🚨 ขาดงานจริง (ไม่มีใบลา)", f"{len(absent_records)} คน")
        ck4.metric("🔔 ลืมสแกน (ไม่มีใบลา)", f"{len(forget_records)} ครั้ง")
        st.markdown("---")

        sub_c1, sub_c2, sub_c3 = st.tabs([
            f"✅ ผู้มีใบลาอนุมัติ ({len(leave_records)})",
            f"🚨 ขาดงานจริงไม่มีใบลา ({len(absent_records)})",
            f"🔔 ลืมสแกนไม่มีใบลา ({len(forget_records)})"
        ])

        with sub_c1:
            if not leave_records.empty:
                st.success(f"พบเจ้าหน้าที่ที่มีใบลาถูกต้องจำนวน {len(leave_records)} รายการ (ปรับสถานะอัตโนมัติ ไม่ถือว่าขาดงาน)")
                cols_l = ['วันที่', 'แผนก', 'ชื่อ-สกุล', 'เข้างาน', 'ออกงาน', 'ประเภทการลา', 'สถานะ', 'หมายเหตุการลา']
                cols_l = [c for c in cols_l if c in leave_records.columns]
                st.dataframe(style_status_table(format_df_for_display(leave_records[cols_l])), use_container_width=True)
            else:
                st.info("ไม่พบรายการที่มีการลาในกลุ่มที่เลือก")

        with sub_c2:
            if not absent_records.empty:
                st.error(f"⚠️ พบเจ้าหน้าที่ขาดงานโดยไม่มีใบลาจำนวน {len(absent_records)} รายการ (ควรดำเนินการตรวจสอบหรือติดตามใบลา)")
                cols_a = ['วันที่', 'แผนก', 'ชื่อ-สกุล', 'เข้างาน', 'ออกงาน', 'สถานะ']
                st.dataframe(style_status_table(format_df_for_display(absent_records[cols_a])), use_container_width=True)
            else:
                st.success("✨ ไม่มีเจ้าหน้าที่ขาดงานโดยไม่มีใบลาในกลุ่มที่เลือก")

        with sub_c3:
            if not forget_records.empty:
                st.warning(f"พบรายการลืมสแกนนิ้วเข้าหรือออกที่ไม่มีใบลา {len(forget_records)} รายการ")
                cols_f = ['วันที่', 'แผนก', 'ชื่อ-สกุล', 'เข้างาน', 'ออกงาน', 'สถานะ']
                st.dataframe(style_status_table(format_df_for_display(forget_records[cols_f])), use_container_width=True)
            else:
                st.success("✨ ไม่มีรายการลืมสแกนนิ้วในกลุ่มที่เลือก")

        dept_tag_c = f"_{sel_dept_cross}" if sel_dept_cross != "ทั้งหมด" else ""
        btn_cross_label = "📥 ดาวน์โหลดรายงานตรวจสอบวันลา (Excel)" + (f" - {sel_dept_cross}" if sel_dept_cross != "ทั้งหมด" else "")
        st.download_button(
            btn_cross_label,
            cross_check_to_excel_bytes(df_cross, sel_dept_cross),
            f"leave_cross_check_report{dept_tag_c}.xlsx",
            use_container_width=True
        )
else:
    st.info("👈 โปรดอัปโหลดไฟล์ PDF ในแท็บ 'นำเข้าข้อมูล' ก่อนครับ")
