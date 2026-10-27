import streamlit as st

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Prompt:wght@300;400;500;600;700&family=Sarabun:wght@300;400;500;600;700&display=swap');

    html, body, [class*="css"], .stApp {
        font-family: 'Prompt', 'Sarabun', sans-serif !important;
        background-color: #FDFBF7;
        color: #2D2825;
    }

    [data-testid="stSidebar"] {
        background-color: #F8F3EA !important;
        border-right: 1.5px solid #EAE1D3 !important;
    }
    [data-testid="stSidebarNav"] span {
        font-family: 'Prompt', sans-serif !important;
        font-weight: 500 !important;
    }

    /* Hero Banner Styling */
    .hero-container {
        background: linear-gradient(135deg, #FFF5EB 0%, #FEF2E4 40%, #F3F8F5 100%);
        border: 2px solid #F5E6D3;
        border-radius: 28px;
        padding: 45px 35px;
        text-align: center;
        box-shadow: 0 16px 36px -12px rgba(233, 125, 50, 0.12), 0 4px 12px rgba(0, 0, 0, 0.03);
        margin-bottom: 35px;
        position: relative;
        overflow: hidden;
    }

    .hero-container::before {
        content: "";
        position: absolute;
        top: -60px;
        right: -60px;
        width: 180px;
        height: 180px;
        background: radial-gradient(circle, rgba(233, 125, 50, 0.15) 0%, rgba(233, 125, 50, 0) 70%);
        border-radius: 50%;
        pointer-events: none;
    }

    .hero-container::after {
        content: "";
        position: absolute;
        bottom: -50px;
        left: -50px;
        width: 160px;
        height: 160px;
        background: radial-gradient(circle, rgba(72, 167, 120, 0.15) 0%, rgba(72, 167, 120, 0) 70%);
        border-radius: 50%;
        pointer-events: none;
    }

    .badge-pill {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        padding: 6px 16px;
        background: #FFFFFF;
        border: 1px solid #F0D7C2;
        border-radius: 9999px;
        font-size: 0.88rem;
        font-weight: 600;
        color: #D4661B;
        box-shadow: 0 2px 8px rgba(212, 102, 27, 0.08);
        margin-bottom: 16px;
    }

    .hero-title {
        font-size: 2.3rem !important;
        font-weight: 700 !important;
        color: #2D2825 !important;
        margin-bottom: 12px !important;
        line-height: 1.3 !important;
    }

    .hero-sub {
        font-size: 1.1rem;
        color: #6B625B;
        max-width: 680px;
        margin: 0 auto 24px auto;
        line-height: 1.6;
    }

    .hero-stats-row {
        display: flex;
        justify-content: center;
        gap: 20px;
        flex-wrap: wrap;
        margin-top: 15px;
    }

    .hero-stat-item {
        background: rgba(255, 255, 255, 0.85);
        border: 1px solid #EFE4D6;
        border-radius: 16px;
        padding: 8px 18px;
        font-size: 0.9rem;
        font-weight: 500;
        color: #4A423D;
        display: flex;
        align-items: center;
        gap: 8px;
        backdrop-filter: blur(4px);
    }

    /* System Cards */
    .system-card {
        background: #FFFFFF;
        border: 2px solid #EFE7DC;
        border-radius: 24px;
        padding: 32px 28px;
        height: 100%;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        transition: all 0.35s cubic-bezier(0.25, 0.8, 0.25, 1);
        box-shadow: 0 8px 24px -6px rgba(0, 0, 0, 0.04);
        position: relative;
    }

    .system-card:hover {
        transform: translateY(-8px);
        border-color: #E97D32;
        box-shadow: 0 18px 36px -8px rgba(233, 125, 50, 0.18), 0 4px 12px rgba(0, 0, 0, 0.04);
    }

    .system-card.green:hover {
        border-color: #48A778;
        box-shadow: 0 18px 36px -8px rgba(72, 167, 120, 0.18), 0 4px 12px rgba(0, 0, 0, 0.04);
    }

    .card-icon-bubble {
        width: 72px;
        height: 72px;
        border-radius: 22px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 2.2rem;
        margin-bottom: 20px;
        background: #FFF1E6;
        border: 1px solid #FCDDC7;
    }

    .card-icon-bubble.green {
        background: #EAF6F0;
        border-color: #CBEAD9;
    }

    .card-title {
        font-size: 1.45rem;
        font-weight: 700;
        color: #2D2825;
        margin-bottom: 8px;
    }

    .card-desc {
        color: #6E655E;
        font-size: 0.95rem;
        line-height: 1.6;
        margin-bottom: 20px;
    }

    .tags-container {
        display: flex;
        flex-wrap: wrap;
        gap: 8px;
        margin-bottom: 24px;
    }

    .feature-tag {
        background: #F8F4EE;
        border: 1px solid #EBE2D5;
        border-radius: 12px;
        padding: 4px 12px;
        font-size: 0.8rem;
        font-weight: 500;
        color: #554D47;
    }

    /* Feature Highlight Boxes */
    .highlight-box {
        background: #FFFFFF;
        border: 1px solid #EFE7DC;
        border-radius: 20px;
        padding: 24px 20px;
        text-align: center;
        transition: transform 0.25s ease, box-shadow 0.25s ease;
        height: 100%;
    }

    .highlight-box:hover {
        transform: translateY(-4px);
        box-shadow: 0 10px 24px rgba(0, 0, 0, 0.05);
        border-color: #E2D5C3;
    }

    .highlight-icon {
        font-size: 2rem;
        margin-bottom: 12px;
        display: inline-block;
    }

    .highlight-title {
        font-size: 1.05rem;
        font-weight: 600;
        color: #2D2825;
        margin-bottom: 6px;
    }

    .highlight-desc {
        font-size: 0.88rem;
        color: #7A7068;
        line-height: 1.5;
        margin: 0;
    }

    /* Step Guide */
    .step-card {
        background: #FFFFFF;
        border: 1px solid #EFE7DC;
        border-radius: 20px;
        padding: 20px;
        display: flex;
        align-items: center;
        gap: 16px;
    }

    .step-number {
        width: 44px;
        height: 44px;
        border-radius: 14px;
        background: #E97D32;
        color: #FFFFFF;
        display: flex;
        align-items: center;
        justify-content: center;
        font-weight: 700;
        font-size: 1.2rem;
        flex-shrink: 0;
        box-shadow: 0 4px 12px rgba(233, 125, 50, 0.3);
    }

    /* Buttons */
    div.stButton > button {
        border-radius: 9999px !important;
        font-family: 'Prompt', sans-serif !important;
        font-weight: 600 !important;
        font-size: 1rem !important;
        padding: 12px 24px !important;
        transition: all 0.25s ease !important;
    }

    div.stButton > button:hover {
        transform: translateY(-2px) scale(1.01) !important;
    }
</style>
""", unsafe_allow_html=True)

# -------------------------
# Hero Banner
# -------------------------
st.markdown("""
<div class="hero-container">
    <div class="badge-pill">
        ✨ Smart HR Intelligence Portal • ศูนย์ฝึกพาณิชย์นาวี
    </div>
    <div class="hero-title">
        ระบบบริหารจัดการและตรวจนับเวลาปฏิบัติงาน
    </div>
    <p class="hero-sub">
        แพลตฟอร์มวิเคราะห์เวลาสแกนนิ้ว ตรวจสอบวันลา และสถิติการปฏิบัติราชการแบบครบวงจร แม่นยำ รวดเร็ว และเป็นระเบียบ
    </p>
    <div class="hero-stats-row">
        <div class="hero-stat-item">📄 อ่าน PDF สแกนนิ้วอัตโนมัติ</div>
        <div class="hero-stat-item">🏖️ ตรวจเทียบวันลารายบุคคล</div>
        <div class="hero-stat-item">🏢 แยกกลุ่มงาน 4 ฝ่าย</div>
        <div class="hero-stat-item">📥 ส่งออก Excel จัดรูปเล่มพร้อมลงนาม</div>
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
                <span style="background: #FFF1E6; color: #D4661B; padding: 4px 12px; border-radius: 9999px; font-size: 0.8rem; font-weight: 600;">🔥 ระบบหลัก</span>
            </div>
            <div class="card-title">ระบบบันทึกเวลาปฏิบัติงาน</div>
            <p class="card-desc">
                ประมวลผลไฟล์ PDF สแกนลายนิ้วมือ ตรวจจับเวลาเข้า-ออก ตรวจหาคนมาสาย ลืมสแกนนิ้ว และตรวจเทียบกับตารางวันลาเพื่อคัดกรองการขาดงานจริงโดยอัตโนมัติ
            </p>
            <div class="tags-container">
                <span class="feature-tag">📄 PDF Parsing</span>
                <span class="feature-tag">🧩 แก้ปัญหาสระลอย</span>
                <span class="feature-tag">🏖️ เทียบวันลาอัตโนมัติ</span>
                <span class="feature-tag">📊 สถิติรายวัน/เดือน</span>
                <span class="feature-tag">📥 ส่งออกรายงาน Excel</span>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)
    if st.button("🚀 เข้าสู่ระบบบันทึกเวลาปฏิบัติงาน", use_container_width=True, type="primary"):
        st.switch_page("pages/attendance.py")

with col2:
    st.markdown("""
    <div class="system-card green">
        <div>
            <div style="display: flex; justify-content: space-between; align-items: flex-start;">
                <div class="card-icon-bubble green">🏖️</div>
                <span style="background: #EAF6F0; color: #2E7D56; padding: 4px 12px; border-radius: 9999px; font-size: 0.8rem; font-weight: 600;">📋 บันทึกวันลา</span>
            </div>
            <div class="card-title">ระบบตรวจสอบข้อมูลการลา</div>
            <p class="card-desc">
                วิเคราะห์ตารางบันทึกการลาประจำปีจากไฟล์ Excel สรุปสถิติโควตาวันลาป่วย ลากิจ ลาพักผ่อน และการมาสาย จำแนกตามปีงบประมาณและฝ่ายงาน
            </p>
            <div class="tags-container">
                <span class="feature-tag">📑 ตารางวันลา Excel</span>
                <span class="feature-tag">🤒 ลาป่วย / ลากิจ</span>
                <span class="feature-tag">🏖️ ลาพักผ่อน</span>
                <span class="feature-tag">⏰ สถิติการมาสาย</span>
                <span class="feature-tag">📈 กราฟสรุปรายบุคคล</span>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)
    if st.button("📊 เข้าสู่ระบบตรวจสอบข้อมูลการลา", use_container_width=True, type="secondary"):
        st.switch_page("pages/leave.py")

st.markdown("<br><hr style='border: none; border-top: 1px solid #EFE7DC; margin: 30px 0;'><br>", unsafe_allow_html=True)

# -------------------------
# Feature Highlights Section
# -------------------------
st.markdown("<h3 style='text-align: center; font-weight: 700; color: #2D2825; margin-bottom: 28px;'>🌟 ฟังก์ชันและจุดเด่นของระบบ</h3>", unsafe_allow_html=True)

h_col1, h_col2, h_col3, h_col4 = st.columns(4)

with h_col1:
    st.markdown("""
    <div class="highlight-box">
        <span class="highlight-icon">🧩</span>
        <div class="highlight-title">ซ่อมสระภาษาไทย</div>
        <p class="highlight-desc">ระบบแก้ปัญหาสระลอย วรรณยุกต์ซ้อน และตัวอักษรแตกหักจากเครื่องสแกนให้อ่านง่าย</p>
    </div>
    """, unsafe_allow_html=True)

with h_col2:
    st.markdown("""
    <div class="highlight-box">
        <span class="highlight-icon">⚡</span>
        <div class="highlight-title">ตรวจเทียบวันลาทันที</div>
        <p class="highlight-desc">นำเข้า Excel วันลาพร้อม PDF เพื่อแยกแยะคนมีใบลาออกจากคนขาดงานจริง</p>
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
    <div class="highlight-box">
        <span class="highlight-icon">📥</span>
        <div class="highlight-title">รายงานพร้อมลงนาม</div>
        <p class="highlight-desc">ดาวน์โหลด Excel จัดรูปแบบมาตรฐานฟอนต์ราชการ TH SarabunPSK พร้อมใช้งาน</p>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br><br>", unsafe_allow_html=True)

# -------------------------
# Step-by-Step Guide
# -------------------------
st.markdown("<h4 style='font-weight: 700; color: #2D2825; margin-bottom: 18px;'>💡 ขั้นตอนการใช้งานง่ายๆ 3 สเต็ป</h4>", unsafe_allow_html=True)
s1, s2, s3 = st.columns(3)

with s1:
    st.markdown("""
    <div class="step-card">
        <div class="step-number">1</div>
        <div>
            <div style="font-weight: 600; color: #2D2825;">นำเข้าไฟล์ข้อมูล</div>
            <div style="font-size: 0.88rem; color: #6E655E;">อัปโหลดไฟล์ PDF สแกนนิ้ว หรือไฟล์วันลา Excel</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

with s2:
    st.markdown("""
    <div class="step-card">
        <div class="step-number">2</div>
        <div>
            <div style="font-weight: 600; color: #2D2825;">ประมวลผลอัตโนมัติ</div>
            <div style="font-size: 0.88rem; color: #6E655E;">ระบบจำแนกเวลา มาสาย ลืมสแกน ขาดงาน และวันลา</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

with s3:
    st.markdown("""
    <div class="step-card">
        <div class="step-number">3</div>
        <div>
            <div style="font-weight: 600; color: #2D2825;">ตรวจสอบและดาวน์โหลด</div>
            <div style="font-size: 0.88rem; color: #6E655E;">ดูสถิติรายวัน รายเดือน และดาวน์โหลดรายงาน Excel</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br><br>", unsafe_allow_html=True)

# -------------------------
# Footer
# -------------------------
st.markdown("""
<div style="text-align: center; color: #9C9187; font-size: 0.88rem; padding: 25px 0 10px 0; border-top: 1px solid #EFE7DC;">
    ศูนย์ฝึกพาณิชย์นาวี • พัฒนาระบบโดย นายเอกชัย ไตรโยธี<br>
    <span style="font-size: 0.8rem; color: #B3A89E;">เวอร์ชัน 2.0 (Playful Warm Edition)</span>
</div>
""", unsafe_allow_html=True)
