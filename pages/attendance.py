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
th_holidays = holidays.Thailand(years=[current_year, current_year + 1])

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

def df_to_excel_bytes(pivot: pd.DataFrame, df_all: pd.DataFrame):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        pivot.to_excel(writer, sheet_name="Summary", startrow=2, index=True)
        workbook  = writer.book
        worksheet = writer.sheets["Summary"]
        base_font = 'Cordia New' 
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
        header_text = f"รายงานการเข้า-ออกงาน ประจำเดือน {month_th} {year_be}".strip()
        merge_format = workbook.add_format({'bold': True, 'font_name': base_font, 'font_size': 18, 'align': 'center', 'valign': 'bottom'})
        worksheet.merge_range(0, 0, 0, len(pivot.columns), header_text, merge_format)
    return output.getvalue()

def daily_to_excel_bytes(df_day: pd.DataFrame, report_date):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        export_df = df_day[['แผนก', 'ชื่อ-สกุล', 'เข้างาน', 'ออกงาน', 'สถานะ']]
        export_df.to_excel(writer, sheet_name="Daily", startrow=2, index=False)
        workbook, worksheet = writer.book, writer.sheets["Daily"]
        base_font = 'Cordia New' 
        base_format = workbook.add_format({'font_name': base_font, 'font_size': 14, 'align': 'center', 'valign': 'bottom', 'border': 1})
        name_format = workbook.add_format({'font_name': base_font, 'font_size': 14, 'align': 'left', 'valign': 'bottom', 'border': 1})
        header_format = workbook.add_format({'font_name': base_font, 'font_size': 14, 'bold': True, 'align': 'center', 'valign': 'bottom', 'border': 1, 'bg_color': '#D9D9D9'})
        for row in range(len(export_df) + 3): worksheet.set_row(row, 30, base_format)
        for col_num, value in enumerate(export_df.columns.values): worksheet.write(2, col_num, value, header_format)
        worksheet.set_column(0, 0, 20, base_format); worksheet.set_column(1, 1, 35, name_format); worksheet.set_column(2, 4, 15, base_format)
        header_text = f"รายงานการเข้า-ออกงาน ประจำวันที่ {report_date.strftime('%d/%m/%Y')}"
        merge_format = workbook.add_format({'bold': True, 'font_name': base_font, 'font_size': 18, 'align': 'center', 'valign': 'bottom'})
        worksheet.merge_range(0, 0, 0, len(export_df.columns)-1, header_text, merge_format)
    return output.getvalue()

# -------------------------
# PDF Parsing
# -------------------------
def process_pdf(file, late_threshold, exc_week, exc_th, extra_list):
    data = []
    dept_variants = ['ข้าราชการ','พนักงานราชการ','ลูกจ้างชั่วคราว','ลูกจ้างประจำ']
    try:
        with pdfplumber.open(file) as pdf:
            for pageno, page in enumerate(pdf.pages, start=1):
                # ใช้ x_tolerance และ y_tolerance เพื่อช่วยรวมสระ/วรรณยุกต์ไทยไม่ให้แตกกระจาย
                text = page.extract_text(x_tolerance=3, y_tolerance=3)
                if not text: continue
                lines = [ln.strip() for ln in text.split('\n') if ln.strip()]
                for line in lines:
                    if re.search(r'แผนก|ชื่อ-สกุล|เข้างาน|ออกงาน|วันที่|สาย|ขาดงาน', line): continue
                    line_norm = normalize_whitespace(line)
                    date_match = re.search(r'(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4})', line_norm)
                    if not date_match: continue
                    
                    date_str = date_match.group(1).replace('-', '/')
                    d_parts = date_str.split('/')
                    if len(d_parts) == 3:
                        year = int(d_parts[2])
                        if year > 2400: year -= 543
                        date_str = f"{d_parts[0]}/{d_parts[1]}/{year}"
                    
                    dt_obj = pd.to_datetime(date_str, format="%d/%m/%Y")
                    left = line_norm[:date_match.start()].strip()
                    right = line_norm[date_match.end():].strip()
                    times = re.findall(r'\d{1,2}:\d{2}', right)
                    
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

                    if not check_in and not check_out:
                        if is_it_holiday(dt_obj, exc_week, exc_th, extra_list): status = "วันหยุด"
                        else: status = "ขาดงาน"
                    elif not check_in and check_out: status = "ลืมสแกนนิ้วเข้า"
                    elif check_in and not check_out: status = "ลืมสแกนนิ้วออก"
                    else:
                        try:
                            hh, mm = map(int, check_in.split(':'))
                            status = "สาย" if (hh * 60 + mm) >= late_threshold else "ปกติ"
                        except: status = "ปกติ"

                    dept, name = "", left
                    for dv in dept_variants:
                        if left.startswith(dv):
                            dept, name = dv, left[len(dv):].strip(); break

                    data.append({'แผนก': dept, 'ชื่อ-สกุล': name, 'วันที่': dt_obj, 'เข้างาน': check_in, 'ออกงาน': check_out, 'สถานะ': status})
        return pd.DataFrame(data) if data else None
    except Exception as e:
        st.error(f"Error: {e}"); return None

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
        all_data = []
        for f in uploaded_files:
            df = process_pdf(f, LATE_THRESHOLD_MINUTES, exclude_weekends, exclude_th_holidays, extra_holidays)
            if df is not None: all_data.append(df)
        if all_data:
            master_df = pd.concat(all_data, ignore_index=True).drop_duplicates(subset=['ชื่อ-สกุล', 'วันที่'], keep='last')
            st.session_state.master_df = master_df
            st.success("✅ ประมวลผลสำเร็จ!")
        else: st.error("❌ ไม่พบข้อมูล")

if 'master_df' in st.session_state:
    master_df = st.session_state.master_df
    with tab2:
        dates = sorted(master_df['วันที่'].dt.date.unique(), reverse=True)
        sel_date = st.selectbox("📅 เลือกวันที่:", dates)
        df_day = master_df[master_df['วันที่'].dt.date == sel_date]
        c_status = df_day['สถานะ'].value_counts().to_dict()
        col1, col2, col3, col4, col5 = st.columns(5)
        col1.metric("👥 ทั้งหมด", len(df_day)); col2.metric("✅ ปกติ", c_status.get("ปกติ",0)); col3.metric("⚠️ สาย", c_status.get("สาย",0)); col4.metric("❌ ขาด", c_status.get("ขาดงาน",0)); col5.metric("🔔 ลืมสแกน", c_status.get("ลืมสแกนนิ้วเข้า",0)+c_status.get("ลืมสแกนนิ้วออก",0))
        st.markdown("---")
        c_chart, c_leader = st.columns(2)
        with c_chart:
            fig = px.pie(pd.DataFrame(list(c_status.items()), columns=['Status', 'Count']), values='Count', names='Status', hole=0.4, color='Status', color_discrete_map={"ปกติ": "#28a745", "สาย": "#ffc107", "ขาดงาน": "#dc3545", "ลืมสแกนนิ้วเข้า": "#fd7e14", "ลืมสแกนนิ้วออก": "#fd7e14", "วันหยุด": "#6c757d"})
            st.plotly_chart(fig, use_container_width=True)
        with c_leader:
            st.subheader("🏆 มาเช้า 5 อันดับแรก")
            df_early = df_day.dropna(subset=['เข้างาน']).sort_values('เข้างาน').head(5)
            for i, r in enumerate(df_early.itertuples(), 1): st.markdown(f"**{'🥇' if i==1 else '🥈' if i==2 else '🥉' if i==3 else '🏃'} อันดับ {i}** | {r._2} — ⏰ **{r.เข้างาน}**")
        st.dataframe(style_status_table(df_day), use_container_width=True)
        st.download_button("📥 ดาวน์โหลดรายงานวันนี้ (Excel)", daily_to_excel_bytes(df_day, sel_date), f"daily_{sel_date}.xlsx")

    with tab3:
        st.subheader("🏢 การมาสายแยกตามฝ่าย")
        dept_stat = pd.crosstab(master_df['แผนก'], master_df['สถานะ']).fillna(0)
        if 'สาย' in dept_stat.columns: st.plotly_chart(px.bar(dept_stat.reset_index(), x='แผนก', y='สาย', color='แผนก', text_auto=True), use_container_width=True)
        master_df['day'] = master_df['วันที่'].dt.day
        pivot = master_df.pivot_table(index="ชื่อ-สกุล", columns="day", values="สถานะ", aggfunc="first").fillna("")
        st.dataframe(style_status_table(pivot), use_container_width=True)
        st.download_button("📥 ดาวน์โหลดสรุปรวม (Excel)", df_to_excel_bytes(pivot, master_df), "monthly_summary.xlsx")

    with tab4:
        search_name = st.selectbox("🔍 ค้นหาชื่อพนักงาน:", ["-- เลือกชื่อ --"] + sorted(master_df['ชื่อ-สกุล'].unique()))
        if search_name != "-- เลือกชื่อ --":
            df_p = master_df[master_df['ชื่อ-สกุล'] == search_name].sort_values('วันที่')
            ps = df_p['สถานะ'].value_counts().to_dict()
            score = max(100 - (ps.get('สาย',0)*2) - ((ps.get('ลืมสแกนนิ้วเข้า',0)+ps.get('ลืมสแกนนิ้วออก',0))*3) - (ps.get('ขาดงาน',0)*5), 0)
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("⭐ คะแนน", f"{score}/100"); c2.metric("⚠️ สาย", ps.get('สาย',0)); c3.metric("❌ ขาด", ps.get('ขาดงาน',0)); c4.metric("🔔 ลืมสแกน", ps.get('ลืมสแกนนิ้วเข้า',0)+ps.get('ลืมสแกนนิ้วออก',0))
            st.dataframe(style_status_table(df_p[['วันที่', 'เข้างาน', 'ออกงาน', 'สถานะ']]), use_container_width=True)
else: st.info("👈 โปรดอัปโหลดไฟล์ PDF ในแท็บ 'นำเข้าข้อมูล' ก่อนครับ")
