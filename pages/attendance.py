import os
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

# -------------------------
# ตั้งค่าหน้าเว็บและ CSS
# -------------------------

# ฉีด CSS เพื่อแก้ปัญหาสระลอยและจัดฟอนต์เฉพาะส่วนเนื้อหา (ไม่ให้กระทบ UI หลัก)
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@400;700&display=swap');
    
    /* เน้นแก้เฉพาะส่วนที่เป็นตารางและข้อความบรรยาย */
    .stMarkdown, .stTable, .stDataFrame, [data-testid="stMetricValue"] {
        font-family: 'Sarabun', 'Tahoma', sans-serif !important;
    }
    
    /* ปรับแต่งตารางให้สระไม่ลอย */
    td, th {
        font-family: 'Sarabun', 'Tahoma', sans-serif !important;
        line-height: 1.6 !important;
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
    
    # 1. ล้างวงกลมประ (U+25CC) และอักขระล่องหน/ควบคุมทั้งหมด (Non-printable characters)
    # รวมถึงรหัสแปลกๆ ที่ PDF มักแอบใส่มา
    s = re.sub(r'[\u200b\ufeff\xa0\u25cc\u00ad]', '', s)
    s = "".join(ch for ch in s if unicodedata.category(ch)[0] != "C") # ล้างกลุ่ม Control characters

    # 2. **Nuclear Clean**: ลบช่องว่างหรือสิ่งที่กั้นระหว่างพยัญชนะกับสระ/วรรณยุกต์
    # ค้นหาช่องว่างที่ตามด้วยสระบน/ล่าง หรือวรรณยุกต์ แล้วลบทิ้งทันที
    s = re.sub(r'\s+([\u0E31\u0E34-\u0E3A\u0E47-\u0E4E])', r'\1', s)
    
    # 3. จัดลำดับ Re-ordering (สลับวรรณยุกต์ที่มาทีหลังสระข้าง ให้กลับมาอยู่หน้า)
    s = re.sub(r'([\u0E30\u0E32\u0E33])([\u0E48-\u0E4B])', r'\2\1', s)
    
    # 4. จัดลำดับกรณี วรรณยุกต์สลับกับสระบน/ล่าง (สระต้องมาก่อนวรรณยุกต์)
    s = re.sub(r'([\u0E48-\u0E4B])([\u0E31\u0E34-\u0E39])', r'\2\1', s)

    # 5. จัดการ Unicode ให้เป็นรูปแบบมาตรฐานสูงสุด
    s = unicodedata.normalize('NFC', s)
    
    # 6. ลบช่องว่างซ้ำซ้อน
    s = re.sub(r'\s+', ' ', s)
    
    return s.strip()

def month_name_thai(month_num):
    months = ["มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน",
              "กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม"]
    return months[month_num-1] if 1 <= month_num <= 12 else ""

def get_department_options(df: pd.DataFrame):
    """ดึงรายชื่อฝ่าย/กลุ่มงานที่มีอยู่ใน DataFrame จัดเรียงตามมาตรฐาน"""
    existing = df['แผนก'].dropna().unique().tolist()
    ordered = [d for d in DEPT_VARIANTS if d in existing]
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

def df_to_excel_bytes(pivot: pd.DataFrame, df_all: pd.DataFrame, dept_label: str = ""):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        pivot.to_excel(writer, sheet_name="Summary", startrow=2, index=True)
        workbook  = writer.book
        worksheet = writer.sheets["Summary"]
        base_font = 'TH SarabunPSK' 
        base_format = workbook.add_format({'font_name': base_font, 'font_size': 14, 'align': 'center', 'valign': 'bottom', 'border': 1})
        name_format = workbook.add_format({'font_name': base_font, 'font_size': 14, 'align': 'left', 'valign': 'bottom', 'border': 1})
        for row in range(len(pivot) + 3):
            worksheet.set_row(row, 30, base_format)
        worksheet.set_column(0, 0, 35, name_format) 
        if len(pivot.columns) > 0:
            worksheet.set_column(1, len(pivot.columns), 10, base_format)
        valid_dates = df_all['วันที่'].dropna()
        month_th = month_name_thai(int(valid_dates.dt.month.mode()[0])) if not valid_dates.empty else "ไม่ระบุ"
        year_be = int(valid_dates.dt.year.mode()[0]) + 543 if not valid_dates.empty else ""
        sub_dept = f" ({dept_label})" if dept_label and dept_label != "ทั้งหมด" else ""
        header_text = f"รายงานการเข้า-ออกงาน{sub_dept} ประจำเดือน {month_th} {year_be}".strip()
        merge_format = workbook.add_format({'bold': True, 'font_name': base_font, 'font_size': 18, 'align': 'center', 'valign': 'bottom'})
        worksheet.merge_range(0, 0, 0, len(pivot.columns), header_text, merge_format)
    return output.getvalue()

def daily_to_excel_bytes(df_day: pd.DataFrame, report_date, dept_label: str = ""):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        export_df = df_day[['แผนก', 'ชื่อ-สกุล', 'เข้างาน', 'ออกงาน', 'สถานะ']]
        export_df.to_excel(writer, sheet_name="Daily", startrow=2, index=False)
        workbook, worksheet = writer.book, writer.sheets["Daily"]
        base_font = 'TH SarabunPSK' 
        base_format = workbook.add_format({'font_name': base_font, 'font_size': 14, 'align': 'center', 'valign': 'bottom', 'border': 1})
        name_format = workbook.add_format({'font_name': base_font, 'font_size': 14, 'align': 'left', 'valign': 'bottom', 'border': 1})
        header_format = workbook.add_format({'font_name': base_font, 'font_size': 14, 'bold': True, 'align': 'center', 'valign': 'bottom', 'border': 1, 'bg_color': '#D9D9D9'})
        for row in range(len(export_df) + 3): worksheet.set_row(row, 30, base_format)
        for col_num, value in enumerate(export_df.columns.values): worksheet.write(2, col_num, value, header_format)
        worksheet.set_column(0, 0, 20, base_format); worksheet.set_column(1, 1, 35, name_format); worksheet.set_column(2, 4, 15, base_format)
        sub_dept = f" ({dept_label})" if dept_label and dept_label != "ทั้งหมด" else ""
        header_text = f"รายงานการเข้า-ออกงาน{sub_dept} ประจำวันที่ {report_date.strftime('%d/%m/%Y')}"
        merge_format = workbook.add_format({'bold': True, 'font_name': base_font, 'font_size': 18, 'align': 'center', 'valign': 'bottom'})
        worksheet.merge_range(0, 0, 0, len(export_df.columns)-1, header_text, merge_format)
    return output.getvalue()

def missing_scan_to_excel_bytes(df_all: pd.DataFrame, dept_label: str = ""):
    output = BytesIO()
    
    # 1. Prepare data
    missing_statuses = ['ลืมสแกนนิ้วเข้า', 'ลืมสแกนนิ้วออก', 'ขาดงาน']
    df_missing = df_all[df_all['สถานะ'].isin(missing_statuses)].copy()
    
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
        hh, mm = map(int, check_in.split(':'))
        return "สาย" if (hh * 60 + mm) >= late_threshold else "ปกติ"
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


def process_pdf(file, late_threshold, exc_week, exc_th, extra_list):
    """คืนค่า (DataFrame หรือ None, รายการบรรทัดที่ข้าม)"""
    fname = getattr(file, 'name', str(file))
    all_records, all_skipped = [], []
    try:
        with pdfplumber.open(file) as pdf:
            for pageno, page in enumerate(pdf.pages, start=1):
                # y_tolerance ต่ำลงเพื่อไม่ให้แถวที่อยู่ชิดกันถูกรวมเป็นบรรทัดเดียว
                text = page.extract_text(x_tolerance=2, y_tolerance=2)
                if not text:
                    all_skipped.append(f"{fname} หน้า {pageno}: ไม่พบข้อความ (อาจเป็น PDF แบบสแกนภาพ)")
                    continue
                lines = [ln.strip() for ln in text.split('\n') if ln.strip()]
                recs, skipped = parse_lines(lines, late_threshold, exc_week, exc_th, extra_list,
                                            source=f"{fname} หน้า {pageno}")
                all_records.extend(recs)
                all_skipped.extend(skipped)
    except Exception as e:
        st.error(f"อ่านไฟล์ {fname} ไม่สำเร็จ: {e}")
    return (pd.DataFrame(all_records) if all_records else None), all_skipped


def unify_names(df: pd.DataFrame) -> pd.DataFrame:
    """รวมชื่อที่ต่างกันแค่ช่องว่าง (เช่น 'นาย สมชาย' กับ 'นายสมชาย') ให้เป็นชื่อเดียวกัน"""
    key = df['ชื่อ-สกุล'].str.replace(r'\s+', '', regex=True)
    canonical = df.groupby(key)['ชื่อ-สกุล'].agg(lambda s: s.value_counts().index[0])
    df = df.copy()
    df['ชื่อ-สกุล'] = key.map(canonical)
    # เติมแผนกที่ว่างจากแถวอื่นของคนเดียวกัน
    dept_map = df[df['แผนก'] != ''].groupby('ชื่อ-สกุล')['แผนก'].agg(lambda s: s.value_counts().index[0])
    df['แผนก'] = df.apply(lambda r: r['แผนก'] or dept_map.get(r['ชื่อ-สกุล'], ''), axis=1)
    return df

# -------------------------
# Styling DataFrame
# -------------------------
def style_status_table(df: pd.DataFrame):
    def highlight(val):
        colors = {"ปกติ": "#C6EFCE; color: #006100", "สาย": "#FFEB9C; color: #9C6500", "ขาดงาน": "#FFC7CE; color: #9C0006", "วันหยุด": "#E7E6E6; color: #595959"}
        color = colors.get(val, "")
        if "ลืมสแกนนิ้ว" in str(val): color = "#FCE4D6; color: #9C0006"
        return f"background-color: {color}" if color else ""
    subset = ['สถานะ'] if "สถานะ" in df.columns else None
    return df.style.map(highlight, subset=subset) if hasattr(df.style, "map") else df.style.applymap(highlight, subset=subset)

# -------------------------
# UI Tabs
# -------------------------
tab1, tab2, tab3, tab4 = st.tabs(["📤 นำเข้าข้อมูล", "📅 ภาพรวมรายวัน", "📈 สถิติรายเดือน", "👤 ข้อมูลรายบุคคล"])

with tab1:
    st.header("📤 นำเข้าข้อมูล")
    uploaded_files = st.file_uploader("ลากไฟล์ PDF มาวางตรงนี้", type="pdf", accept_multiple_files=True)
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
            st.session_state.master_df = master_df
            st.success(f"✅ ประมวลผลสำเร็จ! รวม {master_df['ชื่อ-สกุล'].nunique()} คน ({len(master_df)} บันทึก)")
            
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
        col1, col2, col3, col4, col5 = st.columns(5)
        col1.metric("👥 ทั้งหมด", len(df_day))
        col2.metric("✅ ปกติ", c_status.get("ปกติ", 0))
        col3.metric("⚠️ สาย", c_status.get("สาย", 0))
        col4.metric("❌ ขาด", c_status.get("ขาดงาน", 0))
        col5.metric("🔔 ลืมสแกน", c_status.get("ลืมสแกนนิ้วเข้า", 0) + c_status.get("ลืมสแกนนิ้วออก", 0))
        st.markdown("---")

        c_chart, c_leader = st.columns(2)
        with c_chart:
            fig = px.pie(pd.DataFrame(list(c_status.items()), columns=['Status', 'Count']), values='Count', names='Status', hole=0.4, color='Status', color_discrete_map={"ปกติ": "#28a745", "สาย": "#ffc107", "ขาดงาน": "#dc3545", "ลืมสแกนนิ้วเข้า": "#fd7e14", "ลืมสแกนนิ้วออก": "#fd7e14", "วันหยุด": "#6c757d"})
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

        st.dataframe(style_status_table(df_day), use_container_width=True)
        
        dept_file_tag = f"_{sel_dept_day}" if sel_dept_day != "ทั้งหมด" else ""
        btn_daily_label = "📥 ดาวน์โหลดรายงานวันนี้ (Excel)" + (f" - {sel_dept_day}" if sel_dept_day != "ทั้งหมด" else "")
        st.download_button(btn_daily_label, daily_to_excel_bytes(df_day, sel_date, sel_dept_day), f"daily_{sel_date}{dept_file_tag}.xlsx")

    with tab3:
        sel_dept_month = st.selectbox("🏢 กรองตามฝ่าย/กลุ่มงาน:", dept_options, key="month_dept")
        df_month = filter_by_dept(master_df, sel_dept_month)

        if sel_dept_month == "ทั้งหมด":
            st.subheader("🏢 การมาสายแยกตามฝ่าย")
            dept_stat = pd.crosstab(master_df['แผนก'].replace({'': 'ไม่ระบุฝ่าย'}), master_df['สถานะ']).fillna(0)
            if 'สาย' in dept_stat.columns:
                st.plotly_chart(px.bar(dept_stat.reset_index(), x='แผนก', y='สาย', color='แผนก', text_auto=True), use_container_width=True)
        else:
            st.subheader(f"🏢 สถิติการมาสายในกลุ่ม: {sel_dept_month}")
            late_in_dept = df_month[df_month['สถานะ'] == 'สาย']['ชื่อ-สกุล'].value_counts().reset_index()
            late_in_dept.columns = ['ชื่อ-สกุล', 'จำนวนครั้งที่สาย']
            if not late_in_dept.empty:
                st.plotly_chart(px.bar(late_in_dept.head(10), x='ชื่อ-สกุล', y='จำนวนครั้งที่สาย', color='จำนวนครั้งที่สาย', text_auto=True, color_continuous_scale="Reds"), use_container_width=True)
            else:
                st.info(f"✨ ยอดเยี่ยม! ไม่พบประวัติการมาสายในกลุ่ม {sel_dept_month}")

        # เพิ่มส่วนสรุปภาพรวมรายเดือน
        monthly_summary_title = "### 📊 สรุปภาพรวมรายเดือน (ลืมสแกน/ขาดงาน)" + (f" - {sel_dept_month}" if sel_dept_month != "ทั้งหมด" else "")
        st.markdown(monthly_summary_title)
        m_forget_in_count = df_month[df_month['สถานะ'] == 'ลืมสแกนนิ้วเข้า']['ชื่อ-สกุล'].nunique()
        m_forget_out_count = df_month[df_month['สถานะ'] == 'ลืมสแกนนิ้วออก']['ชื่อ-สกุล'].nunique()
        m_absent_count = df_month[df_month['สถานะ'] == 'ขาดงาน']['ชื่อ-สกุล'].nunique()

        c_m1, c_m2, c_m3 = st.columns(3)
        c_m1.metric("👤 ลืมสแกนเข้า", f"{m_forget_in_count} คน")
        c_m2.metric("👤 ลืมสแกนออก", f"{m_forget_out_count} คน")
        c_m3.metric("👤 ขาดงาน", f"{m_absent_count} คน")
        st.markdown("---")

        df_month_p = df_month.copy()
        df_month_p['day'] = df_month_p['วันที่'].dt.day
        pivot = df_month_p.pivot_table(index="ชื่อ-สกุล", columns="day", values="สถานะ", aggfunc="first").fillna("")
        st.dataframe(style_status_table(pivot), use_container_width=True)

        dept_month_tag = f"_{sel_dept_month}" if sel_dept_month != "ทั้งหมด" else ""
        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            dl_sum_label = "📥 ดาวน์โหลดสรุปรวม (Excel)" + (f" - {sel_dept_month}" if sel_dept_month != "ทั้งหมด" else "")
            st.download_button(dl_sum_label, df_to_excel_bytes(pivot, df_month, sel_dept_month), f"monthly_summary{dept_month_tag}.xlsx", use_container_width=True)
        with col_dl2:
            dl_missing_label = "📥 ดาวน์โหลดรายงานลืมสแกน/ขาดงาน (Excel)" + (f" - {sel_dept_month}" if sel_dept_month != "ทั้งหมด" else "")
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
            score = max(100 - (ps.get('สาย', 0) * 2) - ((ps.get('ลืมสแกนนิ้วเข้า', 0) + ps.get('ลืมสแกนนิ้วออก', 0)) * 3) - (ps.get('ขาดงาน', 0) * 5), 0)
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("⭐ คะแนน", f"{score}/100")
            c2.metric("⚠️ สาย", ps.get('สาย', 0))
            c3.metric("❌ ขาด", ps.get('ขาดงาน', 0))
            c4.metric("🔔 ลืมสแกน", ps.get('ลืมสแกนนิ้วเข้า', 0) + ps.get('ลืมสแกนนิ้วออก', 0))
            st.dataframe(style_status_table(df_p[['วันที่', 'เข้างาน', 'ออกงาน', 'สถานะ']]), use_container_width=True)
else:
    st.info("👈 โปรดอัปโหลดไฟล์ PDF ในแท็บ 'นำเข้าข้อมูล' ก่อนครับ")
