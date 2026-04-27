import os
import streamlit as st
import pandas as pd
from datetime import datetime, time
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import pdfplumber
import re
from io import BytesIO
import plotly.express as px

# -------------------------
# ตั้งค่าหน้าเว็บ
# -------------------------
st.set_page_config(page_title="ระบบเข้า-ออกงาน (Pro)", page_icon="📊", layout="wide", initial_sidebar_state="expanded")

# -------------------------
# ตั้งค่า Sidebar (Settings)
# -------------------------
with st.sidebar:
    st.header("⚙️ ตั้งค่าระบบ")
    
    # 1. ปรับเวลาเข้างานได้เอง
    late_time = st.time_input("เวลาที่ถือว่าสาย", time(8, 36))
    LATE_THRESHOLD_MINUTES = late_time.hour * 60 + late_time.minute
    
    # 2. รองรับวันหยุด
    exclude_weekends = st.checkbox("ไม่นับเสาร์-อาทิตย์เป็น 'ขาดงาน'", value=True)
    
    st.markdown("---")
    st.markdown("พัฒนาโดย: ทีมงานคุณภาพ")

st.title("📊 ระบบจัดการข้อมูลการเข้า-ออกงาน (Pro Edition)")
st.markdown("---")

# -------------------------
# ฟอนต์ไทยสำหรับ matplotlib (เผื่อใช้วาดกราฟด้วย plt)
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

        base_font = 'Cordia New' 

        base_format = workbook.add_format({
            'font_name': base_font,
            'font_size': 14,
            'align': 'center',
            'valign': 'bottom',
            'border': 1
        })

        name_format = workbook.add_format({
            'font_name': base_font,
            'font_size': 14,
            'align': 'left',
            'valign': 'bottom',
            'border': 1
        })

        for row in range(len(pivot) + 3):
            worksheet.set_row(row, 30, base_format)
        
        worksheet.set_column(0, 0, 35, name_format) 
        if len(pivot.columns) > 0:
            worksheet.set_column(1, len(pivot.columns), 10, base_format)

        valid_dates = df_all['วันที่'].dropna()
        if not valid_dates.empty:
            month_num = int(valid_dates.dt.month.mode()[0])
            year_be = int(valid_dates.dt.year.mode()[0]) + 543
            month_th = month_name_thai(month_num)
        else:
            month_th = "ไม่ระบุ"
            year_be = ""

        header_text = f"รายงานการเข้า-ออกงาน ประจำเดือน {month_th} {year_be}".strip()

        merge_format = workbook.add_format({
            'bold': True,
            'font_name': base_font,
            'font_size': 18,
            'align': 'center',
            'valign': 'bottom'
        })
        worksheet.merge_range(0, 0, 0, len(pivot.columns), header_text, merge_format)

    return output.getvalue()

def daily_to_excel_bytes(df_day: pd.DataFrame, report_date):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        export_df = df_day[['แผนก', 'ชื่อ-สกุล', 'เข้างาน', 'ออกงาน', 'สถานะ']]
        export_df.to_excel(writer, sheet_name="Daily", startrow=2, index=False)

        workbook  = writer.book
        worksheet = writer.sheets["Daily"]

        base_font = 'Cordia New' 

        base_format = workbook.add_format({'font_name': base_font, 'font_size': 14, 'align': 'center', 'valign': 'bottom', 'border': 1})
        name_format = workbook.add_format({'font_name': base_font, 'font_size': 14, 'align': 'left', 'valign': 'bottom', 'border': 1})
        header_format = workbook.add_format({'font_name': base_font, 'font_size': 14, 'bold': True, 'align': 'center', 'valign': 'bottom', 'border': 1, 'bg_color': '#D9D9D9'})

        for row in range(len(export_df) + 3):
            worksheet.set_row(row, 30, base_format)
            
        for col_num, value in enumerate(export_df.columns.values):
            worksheet.write(2, col_num, value, header_format)
        
        worksheet.set_column(0, 0, 20, base_format) 
        worksheet.set_column(1, 1, 35, name_format) 
        worksheet.set_column(2, 4, 15, base_format)

        date_str = report_date.strftime("%d/%m/%Y")
        header_text = f"รายงานการเข้า-ออกงาน ประจำวันที่ {date_str}".strip()
        merge_format = workbook.add_format({'bold': True, 'font_name': base_font, 'font_size': 18, 'align': 'center', 'valign': 'bottom'})
        worksheet.merge_range(0, 0, 0, len(export_df.columns)-1, header_text, merge_format)

    return output.getvalue()

# -------------------------
# PDF Parsing
# -------------------------
def process_pdf(file, late_threshold, exclude_weekends):
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
                    
                    date_str = date_match.group(1).replace('-', '/')
                    d_parts = date_str.split('/')
                    if len(d_parts) == 3:
                        year = int(d_parts[2])
                        if year > 2400: year -= 543
                        date_str = f"{d_parts[0]}/{d_parts[1]}/{year}"

                    left = line_norm[:date_match.start()].strip()
                    right = line_norm[date_match.end():].strip()

                    times = re.findall(r'\d{1,2}:\d{2}', right)
                    
                    check_in = None
                    check_out = None
                    
                    if len(times) >= 3:
                        check_in = times[0]
                        check_out = times[1]
                    elif len(times) == 2:
                        try:
                            h1 = int(times[0].split(':')[0])
                            if times[1] == "01:00":
                                if h1 >= 12:
                                    check_out = times[0]
                                    check_in = None
                                else:
                                    check_in = times[0]
                                    check_out = None
                            else:
                                check_in = times[0]
                                check_out = times[1]
                        except:
                            check_in = times[0]
                            check_out = times[1]
                    elif len(times) == 1:
                        try:
                            h1 = int(times[0].split(':')[0])
                            if h1 >= 12:
                                check_out = times[0]
                            else:
                                check_in = times[0]
                        except:
                            check_in = times[0]

                    if not check_in and not check_out:
                        status = "ขาดงาน"
                    elif not check_in and check_out:
                        status = "ลืมสแกนนิ้วเข้า"
                    elif check_in and not check_out:
                        status = "ลืมสแกนนิ้วออก"
                    else:
                        try:
                            hh, mm = map(int, check_in.split(':'))
                            if (hh * 60 + mm) >= late_threshold:
                                status = "สาย"
                            else:
                                status = "ปกติ"
                        except:
                            status = "ปกติ"

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
        
        # จัดการวันหยุด (เสาร์-อาทิตย์)
        if exclude_weekends:
            is_weekend = df['วันที่'].dt.dayofweek >= 5
            df.loc[is_weekend & (df['สถานะ'] == 'ขาดงาน'), 'สถานะ'] = 'วันหยุด'
            
        return df
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาดในการประมวลผลไฟล์: {e}")
        return None

# -------------------------
# Styling DataFrame
# -------------------------
def style_status_table(df: pd.DataFrame):
    def highlight(val):
        if val == "ปกติ": return "background-color: #C6EFCE; color: #006100;"
        elif val == "สาย": return "background-color: #FFEB9C; color: #9C6500;"
        elif "ลืมสแกนนิ้ว" in str(val): return "background-color: #FCE4D6; color: #9C0006;"
        elif val == "ขาดงาน": return "background-color: #FFC7CE; color: #9C0006;"
        elif val == "วันหยุด": return "background-color: #E7E6E6; color: #595959;"
        return ""

    subset_cols = ['สถานะ'] if "สถานะ" in df.columns else None
    if hasattr(df.style, "map"): return df.style.map(highlight, subset=subset_cols)
    else: return df.style.applymap(highlight, subset=subset_cols)

# -------------------------
# UI Tabs
# -------------------------
tab1, tab2, tab3, tab4 = st.tabs(["📤 นำเข้าข้อมูล", "📅 ภาพรวมรายวัน", "📈 สถิติรายเดือน", "👤 ข้อมูลรายบุคคล"])

# --- TAB 1: อัปโหลด ---
with tab1:
    st.header("📥 อัปโหลดไฟล์ PDF (เลือกได้ทีละหลายไฟล์)")
    st.info("💡 นำเข้าได้ทั้งไฟล์สรุปรายวันและรายเดือน สามารถเลือกหลายไฟล์พร้อมกันเพื่อรวมข้อมูลได้เลยครับ")
    uploaded_files = st.file_uploader("ลากไฟล์ PDF มาวางตรงนี้", type="pdf", accept_multiple_files=True)
    
    if uploaded_files:
        if st.button("🚀 ประมวลผลข้อมูล", type="primary"):
            with st.spinner('กำลังอ่านและวิเคราะห์ข้อมูล...'):
                all_data = []
                for f in uploaded_files:
                    df = process_pdf(f, LATE_THRESHOLD_MINUTES, exclude_weekends)
                    if df is not None and not df.empty:
                        all_data.append(df)
                
                if all_data:
                    master_df = pd.concat(all_data, ignore_index=True)
                    # ลบข้อมูลที่ซ้ำกัน (กรณีอัปโหลดไฟล์เดิมซ้ำ)
                    master_df = master_df.drop_duplicates(subset=['ชื่อ-สกุล', 'วันที่'], keep='last')
                    st.session_state.master_df = master_df
                    st.success(f"✅ ประมวลผลสำเร็จ! ข้อมูลพร้อมใช้งานแล้ว กรุณาไปที่แท็บอื่นเพื่อดูผลลัพธ์")
                else:
                    st.error("❌ ไม่พบข้อมูลจากไฟล์ที่อัปโหลด")

# Check if data exists
if 'master_df' in st.session_state:
    master_df = st.session_state.master_df
    
    # --- TAB 2: รายวัน ---
    with tab2:
        st.header("📅 ภาพรวมรายวัน (Daily Dashboard)")
        
        available_dates = master_df['วันที่'].dropna().dt.date.unique()
        available_dates = sorted(available_dates, reverse=True)
        
        if len(available_dates) > 0:
            selected_date = st.selectbox("📅 เลือกวันที่ต้องการดูข้อมูล:", available_dates)
            df_day = master_df[master_df['วันที่'].dt.date == selected_date].copy()
            
            # KPI Metrics
            count_status = df_day['สถานะ'].value_counts().to_dict()
            total_emp = len(df_day)
            
            col1, col2, col3, col4, col5 = st.columns(5)
            col1.metric("👥 พนักงานทั้งหมด", total_emp)
            col2.metric("✅ มาปกติ", count_status.get("ปกติ",0))
            col3.metric("⚠️ มาสาย", count_status.get("สาย",0))
            col4.metric("❌ ขาดงาน", count_status.get("ขาดงาน",0))
            col5.metric("🔔 ลืมสแกน", count_status.get("ลืมสแกนนิ้วเข้า",0) + count_status.get("ลืมสแกนนิ้วออก",0))

            st.markdown("---")
            col_chart, col_leader = st.columns([1, 1])
            
            with col_chart:
                st.subheader("📊 สัดส่วนการเข้างาน")
                pie_data = pd.DataFrame(list(count_status.items()), columns=['Status', 'Count'])
                color_discrete_map = {
                    "ปกติ": "#28a745", "สาย": "#ffc107", "ขาดงาน": "#dc3545", 
                    "ลืมสแกนนิ้วเข้า": "#fd7e14", "ลืมสแกนนิ้วออก": "#fd7e14", "วันหยุด": "#6c757d"
                }
                fig = px.pie(pie_data, values='Count', names='Status', hole=0.4, color='Status', color_discrete_map=color_discrete_map)
                fig.update_layout(margin=dict(t=0, b=0, l=0, r=0))
                st.plotly_chart(fig, use_container_width=True)
                
            with col_leader:
                st.subheader("🏆 Leaderboard: มาเช้า 5 อันดับแรก")
                df_early = df_day.dropna(subset=['เข้างาน']).copy()
                if not df_early.empty:
                    df_early = df_early.sort_values('เข้างาน').head(5)
                    for i, row in enumerate(df_early.itertuples(), 1):
                        medal = "🥇" if i == 1 else "🥈" if i == 2 else "🥉" if i == 3 else "🏃"
                        st.markdown(f"**{medal} อันดับ {i}** | {row._2} — ⏰ **{row.เข้างาน}**")
                else:
                    st.info("ไม่มีข้อมูลเวลาเข้างาน")

            st.subheader("📋 รายละเอียดพนักงานประจำวัน")
            filter_status = st.multiselect("🔍 กรองตามสถานะ:", df_day['สถานะ'].unique(), default=df_day['สถานะ'].unique())
            df_day_filtered = df_day[df_day['สถานะ'].isin(filter_status)]
            st.dataframe(style_status_table(df_day_filtered), use_container_width=True)
            
            excel_bytes_daily = daily_to_excel_bytes(df_day, selected_date)
            st.download_button("📥 ดาวน์โหลดรายงานวันนี้ (Excel)", data=excel_bytes_daily, file_name=f"daily_report_{selected_date}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.warning("ไม่พบข้อมูลวันที่ที่ถูกต้องในไฟล์")

    # --- TAB 3: สถิติรายเดือน ---
    with tab3:
        st.header("📈 สถิติรายเดือน (Monthly Insights)")
        
        st.subheader("🏢 อัตราการมาสายแยกตามแผนก")
        dept_stat = pd.crosstab(master_df['แผนก'], master_df['สถานะ']).fillna(0)
        
        if 'สาย' in dept_stat.columns:
            fig_bar = px.bar(dept_stat.reset_index(), x='แผนก', y='สาย', color='แผนก', text_auto=True, title="จำนวนการมาสายของแต่ละแผนก (ครั้ง)")
            st.plotly_chart(fig_bar, use_container_width=True)
        else:
            st.info("🎉 ไม่มีใครมาสายเลยในข้อมูลชุดนี้!")

        st.markdown("---")
        st.subheader("📅 ตารางสรุปการเข้างานของทุกคน")
        master_df['day'] = master_df['วันที่'].dt.day
        pivot = master_df.pivot_table(index="ชื่อ-สกุล", columns="day", values="สถานะ", aggfunc="first").fillna("")
        st.dataframe(style_status_table(pivot), use_container_width=True)

        excel_bytes = df_to_excel_bytes(pivot, master_df)
        st.download_button("📥 ดาวน์โหลดสรุปรวม (Excel)",
                           data=excel_bytes,
                           file_name="monthly_summary_pro.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # --- TAB 4: ข้อมูลรายบุคคล ---
    with tab4:
        st.header("👤 ข้อมูลรายบุคคล (Individual Score)")
        all_names = sorted(master_df['ชื่อ-สกุล'].unique())
        search_name = st.selectbox("🔍 ค้นหาชื่อพนักงาน:", ["-- กรุณาเลือกหรือพิมพ์ชื่อ --"] + list(all_names))
        
        if search_name != "-- กรุณาเลือกหรือพิมพ์ชื่อ --":
            df_person = master_df[master_df['ชื่อ-สกุล'] == search_name].sort_values('วันที่')
            
            total_days = len(df_person)
            if total_days > 0:
                p_status = df_person['สถานะ'].value_counts().to_dict()
                late_count = p_status.get('สาย', 0)
                absent_count = p_status.get('ขาดงาน', 0)
                miss_scan = p_status.get('ลืมสแกนนิ้วเข้า', 0) + p_status.get('ลืมสแกนนิ้วออก', 0)
                
                # คิดคะแนน
                score = 100 - (late_count * 2) - (miss_scan * 3) - (absent_count * 5)
                score = max(score, 0) 
                
                st.subheader(f"📊 ประวัติการเข้างาน: {search_name}")
                
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("⭐ คะแนนความประพฤติ", f"{score} / 100")
                c2.metric("⚠️ สาย (ครั้ง)", late_count)
                c3.metric("❌ ขาด (ครั้ง)", absent_count)
                c4.metric("🔔 ลืมสแกน (ครั้ง)", miss_scan)
                
                st.markdown("---")
                # Drop day column if exists
                display_cols = ['วันที่', 'เข้างาน', 'ออกงาน', 'สถานะ']
                st.dataframe(style_status_table(df_person[display_cols]), use_container_width=True)

else:
    st.info("👈 โปรดอัปโหลดไฟล์ PDF ในแท็บ 'นำเข้าข้อมูล' ทางด้านซ้ายก่อนครับ เพื่อให้ระบบสร้าง Dashboard สรุปผลให้คุณ")