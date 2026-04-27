import os
import streamlit as st
import pandas as pd
from datetime import datetime
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import pdfplumber
import re
from io import BytesIO

# -------------------------
# ตั้งค่าคงที่ (Constants)
# -------------------------
LATE_THRESHOLD_MINUTES = 8 * 60 + 36  # เวลาที่ถือว่ามาสาย (08:36)

# -------------------------
# ตั้งค่าหน้าเว็บ
# -------------------------
st.set_page_config(page_title="ระบบจัดการข้อมูลการเข้า-ออกงาน", page_icon="📊", layout="wide")
st.title("📊 ระบบจัดการข้อมูลการเข้า-ออกงาน")
st.markdown("---")

# -------------------------
# ฟอนต์ไทยสำหรับ matplotlib
# -------------------------
def set_thai_font():
    # แนะนำ: ถ้านำไป Deploy บน Cloud ให้นำไฟล์ฟอนต์มาวางในโฟลเดอร์เดียวกับโค้ด แล้วเปลี่ยนชื่อให้ตรงกัน
    local_font_path = "THSarabunNew.ttf" 
    if os.path.exists(local_font_path):
        try:
            fm.fontManager.addfont(local_font_path)
            prop = fm.FontProperties(fname=local_font_path)
            font_name = prop.get_name()
            plt.rcParams['font.family'] = font_name
            return font_name
        except Exception as e:
            # หากโหลดฟอนต์ไม่ได้ให้ข้ามไปใช้ฟอนต์ระบบแทน
            pass

    # กรณีหาไฟล์ฟอนต์ในโฟลเดอร์ไม่เจอ จะค้นหาจากระบบปฏิบัติการ (OS) แทน
    preferred = ["TH Sarabun New", "Sarabun", "Noto Sans Thai", "Tahoma", "DejaVu Sans"]
    available = {}
    for fpath in fm.findSystemFonts(fontpaths=None, fontext='ttf'):
        try:
            prop = fm.FontProperties(fname=fpath)
            name = prop.get_name()
            available[name] = fpath
        except Exception:
            continue
    selected = None
    for name in preferred:
        if name in available:
            fm.fontManager.addfont(available[name])
            plt.rcParams['font.family'] = name
            selected = name
            break
    if not selected:
        plt.rcParams['font.family'] = 'DejaVu Sans'
        selected = 'DejaVu Sans'
    return selected

_used_font = set_thai_font()

# -------------------------
# Utilities
# -------------------------
def normalize_whitespace(s: str) -> str:
    if s is None: return ""
    s = s.replace('\u200b', '')
    s = re.sub(r'\s+', ' ', s)
    return s.strip()

def month_name_thai(month_num):
    months = ["มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน",
              "กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม"]
    return months[month_num-1] if 1 <= month_num <= 12 else ""

def df_to_excel_bytes(pivot: pd.DataFrame, df_all: pd.DataFrame):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        pivot.to_excel(writer, sheet_name="Summary", startrow=2, index=True)

        workbook  = writer.book
        worksheet = writer.sheets["Summary"]

        # ฟอร์แมตเซลล์ปกติ
        base_format = workbook.add_format({
            'font_name': 'TH Sarabun New',
            'font_size': 14,
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })

        for row in range(len(pivot)+3):
            worksheet.set_row(row, 20, base_format)
        for col in range(len(pivot.columns)+1):
            worksheet.set_column(col, col, 15, base_format)

        # จัดการหัวตาราง (ป้องกัน IndexError กรณีไม่มีข้อมูลวันที่)
        valid_dates = df_all['วันที่'].dropna()
        if not valid_dates.empty:
            month_num = int(valid_dates.dt.month.mode()[0])
            year_be = int(valid_dates.dt.year.mode()[0]) + 543
            month_th = month_name_thai(month_num)
        else:
            month_th = "ไม่ระบุ"
            year_be = ""

        header_text = f"ข้าราชการศูนย์ฝึกพาณิชย์นาวี\nประจำเดือน {month_th} {year_be}".strip()

        merge_format = workbook.add_format({
            'bold': True,
            'font_name': 'TH Sarabun New',
            'font_size': 16,
            'align': 'center',
            'valign': 'vcenter'
        })
        worksheet.merge_range(0, 0, 0, len(pivot.columns), header_text, merge_format)

    return output.getvalue()

# -------------------------
# PDF Parsing
# -------------------------
def process_pdf(file):
    data = []
    dept_variants = ['ข้าราชการ','พนักงานราชการ','ลูกจ้างชั่วคราว','ลูกจ้างประจำ']
    try:
        with pdfplumber.open(file) as pdf:
            for pageno, page in enumerate(pdf.pages, start=1):
                text = page.extract_text()
                if not text: continue
                lines = [ln.strip() for ln in text.split('\n') if ln.strip()]
                for line in lines:
                    if re.search(r'แผนก|ชื่อ-สกุล|เข้างาน|ออกงาน|วันที่|สาย|ขาดงาน', line):
                        continue
                    line_norm = normalize_whitespace(line)
                    date_match = re.search(r'(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4})', line_norm)
                    if not date_match: continue
                    
                    # แปลงวันที่ และแก้ปัญหาปี พ.ศ. เป็น ค.ศ.
                    date_str = date_match.group(1).replace('-', '/')
                    d_parts = date_str.split('/')
                    if len(d_parts) == 3:
                        year = int(d_parts[2])
                        if year > 2400:  # ถ้าเป็นปี พ.ศ. ให้ลบ 543
                            year -= 543
                        date_str = f"{d_parts[0]}/{d_parts[1]}/{year}"

                    left = line_norm[:date_match.start()].strip()
                    right = line_norm[date_match.end():].strip()

                    # ดึงเวลาเข้า-ออก โดยใช้เวลาสุดท้ายของการสแกนเป็นเวลาออก
                    times = re.findall(r'\d{1,2}:\d{2}', right)
                    check_in = times[0] if len(times) >= 1 else None
                    check_out = times[-1] if len(times) >= 2 else None

                    # Logic กำหนดสถานะ
                    if not check_in and not check_out:
                        status = "ขาดงาน"
                    elif not check_in and check_out:
                        status = "ลืมสแกนนิ้วเข้า"
                    elif check_in and not check_out:
                        status = "ลืมสแกนนิ้วออก"
                    else:
                        # มีทั้งเข้าและออก
                        try:
                            hh, mm = map(int, check_in.split(':'))
                            if (hh * 60 + mm) >= LATE_THRESHOLD_MINUTES:
                                status = "สาย"
                            else:
                                status = "ปกติ"
                        except:
                            status = "ปกติ"

                    # แก้ปัญหา "สาย 01:00" จาก PDF (ตามโค้ดเดิมของคุณ)
                    if "01:00" in right:
                        if not check_in and check_out:
                            status = "ลืมสแกนนิ้วเข้า"
                        elif check_in and not check_out:
                            status = "ลืมสแกนนิ้วออก"

                    # แยกชื่อ/แผนก
                    dept, name = "", left
                    for dv in dept_variants:
                        if left.startswith(dv):
                            dept = dv
                            name = left[len(dv):].strip()
                            break

                    data.append({
                        'แผนก': dept,
                        'ชื่อ-สกุล': name,
                        'วันที่': date_str,
                        'เข้างาน': check_in,
                        'ออกงาน': check_out,
                        'สถานะ': status,
                        'หน้า': pageno
                    })
        if not data:
            return None
        df = pd.DataFrame(data)
        df['วันที่'] = pd.to_datetime(df['วันที่'], format="%d/%m/%Y", errors="coerce")
        return df
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการประมวลผลไฟล์: {e}")
        return None

# -------------------------
# Styling DataFrame
# -------------------------
def style_status_table(df: pd.DataFrame):
    def highlight(val):
        if val == "ปกติ":
            return "background-color: #C6EFCE; color: #006100;"
        elif val == "สาย":
            return "background-color: #FFEB9C; color: #9C6500;"
        elif "ลืมสแกนนิ้ว" in str(val):
            return "background-color: #FCE4D6; color: #9C0006;"
        elif val == "ขาดงาน":
            return "background-color: #FFC7CE; color: #9C0006;"
        return ""

    # ใช้ .map() สำหรับ Pandas เวอร์ชันใหม่ และ fallback กลับไปใช้ .applymap() สำหรับเวอร์ชันเก่า
    style_mapper = getattr(df.style, "map", df.style.applymap)
    
    if "สถานะ" in df.columns:
        return style_mapper(highlight, subset=['สถานะ'])
    else:
        return style_mapper(highlight)

# -------------------------
# UI Tabs
# -------------------------
tab1, tab2, tab3 = st.tabs(["อัปโหลดไฟล์", "สรุปรายวัน", "สรุปรายเดือน"])

with tab1:
    st.header("📤 อัปโหลดไฟล์ PDF ประจำวัน")
    uploaded_pdf = st.file_uploader("เลือกไฟล์ PDF", type="pdf")
    if uploaded_pdf:
        report_date = st.date_input("วันที่ของรายงาน", datetime.now().date())
        if st.button("ประมวลผลรายวัน"):
            with st.spinner('กำลังประมวลผล...'):
                df = process_pdf(uploaded_pdf)
                if df is not None:
                    st.session_state.df = df
                    st.session_state.report_date = report_date
                    st.success("ประมวลผลเสร็จสิ้น! → ไปที่แท็บ 'สรุปรายวัน'")

with tab2:
    st.header("📅 สรุปรายวัน")
    if 'df' in st.session_state:
        df = st.session_state.df
        report_date = pd.to_datetime(st.session_state.report_date)
        df_day = df[df['วันที่'] == report_date]

        count_status = df_day['สถานะ'].value_counts().to_dict()
        col1,col2,col3,col4,col5 = st.columns(5)
        col1.metric("มาปกติ", count_status.get("ปกติ",0))
        col2.metric("มาสาย", count_status.get("สาย",0))
        col3.metric("ขาดงาน", count_status.get("ขาดงาน",0))
        col4.metric("ลืมเข้า", count_status.get("ลืมสแกนนิ้วเข้า",0))
        col5.metric("ลืมออก", count_status.get("ลืมสแกนนิ้วออก",0))

        st.dataframe(style_status_table(df_day), use_container_width=True)
    else:
        st.info("โปรดอัปโหลดไฟล์ในแท็บ 'อัปโหลดไฟล์' ก่อน")

with tab3:
    st.header("📈 สรุปรายเดือน")
    uploaded_month_files = st.file_uploader("📤 อัปโหลดไฟล์ PDF รายวันทั้งเดือน", type="pdf", accept_multiple_files=True)
    if uploaded_month_files:
        if st.button("สร้างรายงานสรุปรายเดือน"):
            with st.spinner('กำลังรวมข้อมูล...'):
                all_data = []
                for f in uploaded_month_files:
                    df_day = process_pdf(f)
                    if df_day is not None and not df_day.empty:
                        all_data.append(df_day)
                
                if all_data:
                    df_all = pd.concat(all_data, ignore_index=True)
                    df_all['day'] = df_all['วันที่'].dt.day
                    pivot = df_all.pivot_table(index="ชื่อ-สกุล", columns="day", values="สถานะ", aggfunc="first").fillna("")

                    st.subheader("📅 ตารางสรุปทั้งเดือน")
                    st.dataframe(style_status_table(pivot), use_container_width=True)

                    excel_bytes = df_to_excel_bytes(pivot, df_all)
                    st.download_button("📥 ดาวน์โหลดสรุปรายเดือน (Excel)",
                                       data=excel_bytes,
                                       file_name="monthly_summary.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                else:
                    st.error("ไม่พบข้อมูลจากไฟล์ PDF หรือรูปแบบไฟล์ไม่ถูกต้อง")