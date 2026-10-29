import streamlit as st

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Caveat:wght@500;600;700&family=Kalam:wght@400;700&family=Mali:ital,wght@0,400;0,500;0,600;0,700;1,400&family=Sarabun:wght@300;400;500;600;700&family=Prompt:wght@400;500;600;700&display=swap');

    /* Global Notebook Paper & Study Desk Atmosphere */
    html, body, [class*="css"], .stApp {
        font-family: 'Sarabun', 'Prompt', sans-serif !important;
        background-color: #fcfbf7 !important;
        background-image:
            linear-gradient(to right, rgba(229, 231, 235, 0.55) 1px, transparent 1px),
            linear-gradient(to bottom, rgba(229, 231, 235, 0.55) 1px, transparent 1px) !important;
        background-size: 24px 24px !important;
        color: #1e293b !important;
    }

    /* Sidebar: Lined Notebook Cover & Margin */
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

    .handwritten-note {
        font-family: 'Caveat', 'Kalam', 'Mali', cursive !important;
    }

    /* Hero Banner: Authentic Ruled Notebook Sheet with Pink Margin & Washi Tape */
    .hero-container {
        background-color: #fcfbf7;
        background-image:
            linear-gradient(90deg, transparent 54px, #fca5a5 54px, #fca5a5 56px, transparent 56px, transparent 60px, rgba(252, 165, 165, 0.45) 60px, rgba(252, 165, 165, 0.45) 61px, transparent 61px),
            repeating-linear-gradient(transparent, transparent 31px, #e5e7eb 31px, #e5e7eb 32px);
        border: 1.5px solid #d1d5db;
        border-radius: 6px 16px 16px 6px;
        padding: 46px 38px 40px 82px;
        text-align: left;
        box-shadow:
            0 1px 1px rgba(0,0,0,0.08),
            0 10px 0 -5px #f8f6f0,
            0 10px 1px -4px rgba(0,0,0,0.12),
            0 20px 0 -10px #f3f0e6,
            0 20px 1px -9px rgba(0,0,0,0.12),
            0 18px 32px rgba(15, 23, 42, 0.08);
        margin-top: 18px;
        margin-bottom: 38px;
        position: relative;
    }

    /* Washi Tape Top Center Accent */
    .hero-container::before {
        content: "";
        position: absolute;
        top: -14px;
        left: 50%;
        transform: translateX(-50%) rotate(-1.5deg);
        width: 170px;
        height: 30px;
        background: rgba(252, 231, 243, 0.88);
        border-left: 2px dotted rgba(244, 114, 182, 0.45);
        border-right: 2px dotted rgba(244, 114, 182, 0.45);
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.08);
        z-index: 2;
    }

    /* Notebook Binder Holes on Left Margin */
    .hero-container::after {
        content: "";
        position: absolute;
        top: 32px;
        bottom: 32px;
        left: 20px;
        width: 18px;
        background-image: radial-gradient(circle, #e5e7eb 6px, rgba(0,0,0,0.12) 7px, transparent 8px);
        background-size: 18px 64px;
        background-repeat: repeat-y;
        pointer-events: none;
    }

    .badge-pill {
        display: inline-flex;
        align-items: center;
        gap: 8px;
        padding: 6px 18px;
        background: #fef9c3;
        border: 1px solid #fde047;
        border-radius: 3px;
        font-family: 'Caveat', 'Kalam', 'Mali', cursive !important;
        font-size: 1.15rem;
        font-weight: 700;
        color: #854d0e;
        box-shadow: 2px 3px 7px rgba(0, 0, 0, 0.08);
        transform: rotate(-1.2deg);
        margin-bottom: 16px;
    }

    .hero-title {
        font-family: 'Prompt', 'Sarabun', sans-serif !important;
        font-size: 2.25rem !important;
        font-weight: 700 !important;
        color: #1e293b !important;
        margin-bottom: 10px !important;
        line-height: 1.35 !important;
    }

    .marker-highlight {
        background: linear-gradient(180deg, transparent 58%, #fef08a 58%);
        padding: 0 6px;
    }

    .hero-sub {
        font-size: 1.05rem;
        color: #334155;
        max-width: 760px;
        margin: 0 0 24px 0;
        line-height: 1.85;
    }

    .hero-stats-row {
        display: flex;
        justify-content: flex-start;
        gap: 14px;
        flex-wrap: wrap;
        margin-top: 14px;
    }

    .hero-stat-item {
        background: #fef9c3;
        border: 1px solid #fde047;
        border-radius: 3px;
        padding: 7px 16px;
        font-family: 'Mali', 'Sarabun', sans-serif !important;
        font-size: 0.9rem;
        font-weight: 600;
        color: #1e293b;
        display: flex;
        align-items: center;
        gap: 8px;
        box-shadow: 2px 3px 6px rgba(0, 0, 0, 0.07);
        transform: rotate(-0.6deg);
    }
    .hero-stat-item:nth-child(even) {
        background: #fce7f3;
        border-color: #f9a8d4;
        transform: rotate(0.8deg);
    }

    /* Main System Cards: Ruled & Graph Notebook Pages with Washi Tape */
    .system-card {
        background-color: #fcfbf7;
        background-image:
            linear-gradient(90deg, transparent 40px, #fca5a5 40px, #fca5a5 42px, transparent 42px),
            repeating-linear-gradient(transparent, transparent 29px, #e5e7eb 29px, #e5e7eb 30px);
        border: 1.5px solid #d1d5db;
        border-radius: 6px 14px 14px 6px;
        padding: 34px 28px 28px 58px;
        height: 100%;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        transition: all 0.3s ease;
        box-shadow: 3px 6px 16px rgba(15, 23, 42, 0.06);
        position: relative;
        margin-top: 12px;
    }

    /* Washi tape corner accent on System Cards */
    .system-card::before {
        content: "";
        position: absolute;
        top: -11px;
        left: 32px;
        width: 115px;
        height: 24px;
        background: rgba(254, 249, 195, 0.9);
        border-left: 2px dotted rgba(234, 179, 8, 0.45);
        border-right: 2px dotted rgba(234, 179, 8, 0.45);
        transform: rotate(-2.5deg);
        box-shadow: 0 2px 4px rgba(0,0,0,0.07);
    }

    .system-card.green::before {
        background: rgba(252, 231, 243, 0.9);
        border-left: 2px dotted rgba(236, 72, 153, 0.45);
        border-right: 2px dotted rgba(236, 72, 153, 0.45);
        transform: rotate(2deg);
        left: auto;
        right: 32px;
    }

    .system-card:hover {
        transform: translateY(-5px) rotate(-0.3deg);
        border-color: #fca5a5;
        box-shadow: 5px 14px 26px rgba(15, 23, 42, 0.11);
    }

    .system-card.green:hover {
        transform: translateY(-5px) rotate(0.3deg);
        border-color: #f472b6;
    }

    .card-icon-bubble {
        width: 64px;
        height: 64px;
        border-radius: 4px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 2rem;
        margin-bottom: 18px;
        background: #fef9c3;
        border: 1px solid #fde047;
        box-shadow: 2px 3px 8px rgba(0,0,0,0.08);
        transform: rotate(-2deg);
    }

    .card-icon-bubble.green {
        background: #fce7f3;
        border-color: #f9a8d4;
        transform: rotate(2deg);
    }

    .sticky-corner-label {
        font-family: 'Caveat', 'Kalam', 'Mali', cursive !important;
        padding: 4px 14px;
        border-radius: 3px;
        font-size: 1rem;
        font-weight: 700;
        box-shadow: 2px 2px 5px rgba(0,0,0,0.08);
        transform: rotate(2deg);
    }

    .card-title {
        font-size: 1.42rem;
        font-weight: 700;
        color: #1e293b;
        margin-bottom: 8px;
    }

    .card-desc {
        color: #334155;
        font-size: 0.96rem;
        line-height: 1.75;
        margin-bottom: 20px;
    }

    .tags-container {
        display: flex;
        flex-wrap: wrap;
        gap: 8px;
        margin-bottom: 20px;
    }

    .feature-tag {
        background: #fcfbf7;
        border: 1px dashed #94a3b8;
        border-radius: 4px;
        padding: 4px 11px;
        font-family: 'Mali', 'Sarabun', sans-serif !important;
        font-size: 0.82rem;
        font-weight: 600;
        color: #334155;
    }

    /* Feature Highlight Boxes: Authentic Pastel Sticky Notes */
    .highlight-box {
        background: #fef9c3;
        border: 1px solid #fde047;
        border-bottom-right-radius: 22px 6px;
        padding: 28px 20px 22px 20px;
        text-align: center;
        transition: transform 0.25s ease, box-shadow 0.25s ease;
        height: 100%;
        position: relative;
        box-shadow: 3px 5px 12px rgba(0, 0, 0, 0.08);
        transform: rotate(-1deg);
        margin-top: 10px;
    }

    .highlight-box.pink {
        background: #fce7f3;
        border-color: #f9a8d4;
        transform: rotate(1.2deg);
    }

    /* Washi tape strip on top of sticky notes */
    .highlight-box::before {
        content: "";
        position: absolute;
        top: -10px;
        left: 50%;
        transform: translateX(-50%) rotate(-1deg);
        width: 75px;
        height: 20px;
        background: rgba(255, 255, 255, 0.65);
        border: 1px solid rgba(203, 213, 225, 0.6);
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }

    .highlight-box:hover {
        transform: translateY(-6px) rotate(0deg) scale(1.02);
        box-shadow: 5px 12px 22px rgba(0, 0, 0, 0.12);
    }

    .highlight-icon {
        font-size: 2.1rem;
        margin-bottom: 10px;
        display: inline-block;
    }

    .highlight-title {
        font-family: 'Mali', 'Prompt', sans-serif !important;
        font-size: 1.08rem;
        font-weight: 700;
        color: #1e293b;
        margin-bottom: 6px;
    }

    .highlight-desc {
        font-size: 0.9rem;
        color: #334155;
        line-height: 1.55;
        margin: 0;
    }

    /* Step Guide: Ruled Index Cards */
    .step-card {
        background-color: #fcfbf7;
        background-image:
            linear-gradient(180deg, transparent 34px, #fca5a5 34px, #fca5a5 36px, transparent 36px),
            repeating-linear-gradient(transparent, transparent 25px, #e5e7eb 25px, #e5e7eb 26px);
        border: 1.5px solid #d1d5db;
        border-radius: 6px;
        padding: 22px 20px;
        display: flex;
        align-items: center;
        gap: 16px;
        box-shadow: 2px 4px 10px rgba(0, 0, 0, 0.06);
        position: relative;
    }

    .step-number {
        width: 46px;
        height: 46px;
        border-radius: 50%;
        background: #fef9c3;
        border: 2px dashed #e11d48;
        color: #e11d48;
        display: flex;
        align-items: center;
        justify-content: center;
        font-family: 'Caveat', 'Kalam', cursive !important;
        font-weight: 700;
        font-size: 1.65rem;
        flex-shrink: 0;
        transform: rotate(-4deg);
    }

    /* Buttons: Stationery Stamp & Tape Style */
    div.stButton > button {
        border-radius: 6px !important;
        font-family: 'Mali', 'Prompt', sans-serif !important;
        font-weight: 700 !important;
        font-size: 1rem !important;
        padding: 12px 24px !important;
        border: 2px solid #1e293b !important;
        box-shadow: 3px 3px 0px #1e293b !important;
        transition: all 0.18s ease !important;
    }

    div.stButton > button[kind="primary"] {
        background-color: #fef9c3 !important;
        color: #1e293b !important;
    }

    div.stButton > button[kind="secondary"] {
        background-color: #fce7f3 !important;
        color: #1e293b !important;
    }

    div.stButton > button:hover {
        transform: translate(-1px, -1px) !important;
        box-shadow: 5px 5px 0px #e11d48 !important;
        border-color: #e11d48 !important;
    }
</style>
""", unsafe_allow_html=True)

# -------------------------
# Hero Banner
# -------------------------
st.markdown("""
<div class="hero-container">
    <div class="badge-pill">
        ✏️ Smart HR Notebook &amp; Study Desk • ศูนย์ฝึกพาณิชย์นาวี
    </div>
    <div class="hero-title">
        <span class="marker-highlight">สมุดบันทึกเวลา</span>และตรวจสอบการปฏิบัติราชการ
    </div>
    <p class="hero-sub">
        แพลตฟอร์มวิเคราะห์เวลาสแกนนิ้ว ตรวจสอบวันลา และสถิติการปฏิบัติราชการในสไตล์สมุดโน้ตที่อ่านง่าย แม่นยำ รวดเร็ว และเป็นระเบียบเรียบร้อย
    </p>
    <div class="hero-stats-row">
        <div class="hero-stat-item">📌 อ่าน PDF สแกนนิ้วอัตโนมัติ</div>
        <div class="hero-stat-item">🗒️ ตรวจเทียบวันลารายบุคคล</div>
        <div class="hero-stat-item">📎 แยกกลุ่มงาน 4 ฝ่าย</div>
        <div class="hero-stat-item">📥 ส่งออก Excel พร้อมลงนาม</div>
    </div>
</div>
""", unsafe_allow_html=True)

# -------------------------
# Main System Cards
# -------------------------
col1, col2 = st.columns(2, gap="large")

with col1:
    st.markdown("""
    <div class="system-card">
        <div>
            <div style="display: flex; justify-content: space-between; align-items: flex-start;">
                <div class="card-icon-bubble">⏱️</div>
                <span class="sticky-corner-label" style="background: #fef9c3; color: #854d0e; border: 1px solid #fde047;">📌 Notebook #1 • ระบบหลัก</span>
            </div>
            <div class="card-title">สมุดบันทึกเวลาปฏิบัติงาน (Attendance)</div>
            <p class="card-desc">
                ประมวลผลไฟล์ PDF สแกนลายนิ้วมือ ตรวจจับเวลาเข้า-ออก ตรวจหาคนมาสาย ขาดงานครึ่งบ่าย ลืมสแกนนิ้ว และตรวจเทียบกับตารางวันลาอัตโนมัติ
            </p>
            <div class="tags-container">
                <span class="feature-tag">📄 PDF Parsing</span>
                <span class="feature-tag">🧩 ซ่อมสระลอย/ชื่อขาด</span>
                <span class="feature-tag">🏖️ เทียบวันลาอัตโนมัติ</span>
                <span class="feature-tag">📊 รายงานประจำเดือน 4 ชีต</span>
                <span class="feature-tag">📥 ส่งออก Excel</span>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)
    if st.button("✏️ เปิดสมุดบันทึกเวลาปฏิบัติงาน", use_container_width=True, type="primary"):
        st.switch_page("pages/attendance.py")

with col2:
    st.markdown("""
    <div class="system-card green">
        <div>
            <div style="display: flex; justify-content: space-between; align-items: flex-start;">
                <div class="card-icon-bubble green">🏖️</div>
                <span class="sticky-corner-label" style="background: #fce7f3; color: #9d174d; border: 1px solid #f9a8d4;">🔖 Notebook #2 • บันทึกวันลา</span>
            </div>
            <div class="card-title">สมุดตรวจสอบข้อมูลการลา (Leave Log)</div>
            <p class="card-desc">
                วิเคราะห์ตารางบันทึกการลาตามแบบฟอร์มทางการกรมเจ้าท่า (13 คอลัมน์) และตารางวันลาประจำปี สรุปสถิติโควตาวันลาป่วย ลากิจ ลาพักผ่อน และการมาสาย
            </p>
            <div class="tags-container">
                <span class="feature-tag">🏛️ แบบฟอร์มกรมเจ้าท่า</span>
                <span class="feature-tag">🤒 ลาป่วย / ลากิจ</span>
                <span class="feature-tag">🏖️ ลาพักผ่อน</span>
                <span class="feature-tag">⏰ สถิติการมาสาย</span>
                <span class="feature-tag">📈 กราฟสรุปรายบุคคล</span>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)
    if st.button("📒 เปิดสมุดตรวจสอบข้อมูลการลา", use_container_width=True, type="secondary"):
        st.switch_page("pages/leave.py")

st.markdown("<br><hr style='border: none; border-top: 2px dashed #fca5a5; margin: 30px 0;'><br>", unsafe_allow_html=True)

# -------------------------
# Feature Highlights Section
# -------------------------
st.markdown("<h3 style='text-align: center; font-weight: 700; color: #1e293b; margin-bottom: 6px;'>📌 โน้ตจุดเด่นของระบบ (Sticky Notes Summary)</h3>", unsafe_allow_html=True)
st.markdown("<p class='handwritten-note' style='text-align: center; font-size: 1.3rem; color: #e11d48; margin-bottom: 24px;'>✨ Key highlights pinned on your study desk ✨</p>", unsafe_allow_html=True)

h_col1, h_col2, h_col3, h_col4 = st.columns(4)

with h_col1:
    st.markdown("""
    <div class="highlight-box">
        <span class="highlight-icon">🧩</span>
        <div class="highlight-title">ซ่อมสระ &amp; ชื่อ-สกุลอัตโนมัติ</div>
        <p class="highlight-desc">แก้ปัญหาสระลอย วรรณยุกต์ซ้อน และเติมเต็มนามสกุลที่ถูกตัดทอนจากเครื่องสแกนให้สมบูรณ์</p>
    </div>
    """, unsafe_allow_html=True)

with h_col2:
    st.markdown("""
    <div class="highlight-box pink">
        <span class="highlight-icon">⚡</span>
        <div class="highlight-title">ตรวจเทียบวันลาทันที</div>
        <p class="highlight-desc">นำเข้าไฟล์วันลาคู่กับ PDF เพื่อแยกแยะผู้มีใบลาและตรวจจับขาดงานครึ่งบ่ายได้แม่นยำ</p>
    </div>
    """, unsafe_allow_html=True)

with h_col3:
    st.markdown("""
    <div class="highlight-box">
        <span class="highlight-icon">🏢</span>
        <div class="highlight-title">กรองตามกลุ่มงาน</div>
        <p class="highlight-desc">แยกดูข้อมูลข้าราชการ พนักงานราชการ ลูกจ้างประจำ และลูกจ้างชั่วคราวได้อิสระ</p>
    </div>
    """, unsafe_allow_html=True)

with h_col4:
    st.markdown("""
    <div class="highlight-box pink">
        <span class="highlight-icon">📥</span>
        <div class="highlight-title">รายงานพร้อมลงนาม</div>
        <p class="highlight-desc">ดาวน์โหลด Excel จัดรูปแบบมาตรฐานฟอนต์ราชการ TH SarabunPSK พร้อมใช้งาน</p>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br><br>", unsafe_allow_html=True)

# -------------------------
# Step-by-Step Guide
# -------------------------
st.markdown("<h4 style='font-weight: 700; color: #1e293b; margin-bottom: 18px;'>📝 ขั้นตอนการใช้งาน 3 สเต็ป (Quick Study Guide)</h4>", unsafe_allow_html=True)
s1, s2, s3 = st.columns(3)

with s1:
    st.markdown("""
    <div class="step-card">
        <div class="step-number">1</div>
        <div>
            <div style="font-weight: 700; color: #1e293b;">นำเข้าไฟล์ข้อมูล</div>
            <div style="font-size: 0.9rem; color: #475569;">อัปโหลดไฟล์ PDF สแกนนิ้ว หรือไฟล์ตารางวันลา Excel/CSV</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

with s2:
    st.markdown("""
    <div class="step-card">
        <div class="step-number">2</div>
        <div>
            <div style="font-weight: 700; color: #1e293b;">ประมวลผลอัตโนมัติ</div>
            <div style="font-size: 0.9rem; color: #475569;">ระบบจำแนกเวลา มาสาย ขาดครึ่งบ่าย ลืมสแกน และวันลา</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

with s3:
    st.markdown("""
    <div class="step-card">
        <div class="step-number">3</div>
        <div>
            <div style="font-weight: 700; color: #1e293b;">ตรวจสอบและดาวน์โหลด</div>
            <div style="font-size: 0.9rem; color: #475569;">ดูสถิติรายวัน รายเดือน และส่งออกรายงาน Excel ทันที</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br><br>", unsafe_allow_html=True)

# -------------------------
# Footer
# -------------------------
st.markdown("""
<div style="text-align: center; color: #64748b; font-size: 0.9rem; padding: 25px 0 10px 0; border-top: 2px dashed #cbd5e1;">
    ศูนย์ฝึกพาณิชย์นาวี • พัฒนาระบบโดย นายเอกชัย ไตรโยธี<br>
    <span class="handwritten-note" style="font-size: 1.15rem; color: #e11d48;">✏️ Notebook Stationery Edition • Ruled Paper &amp; Sticky Notes UI</span>
</div>
""", unsafe_allow_html=True)
