import streamlit as st

st.markdown("""
<style>
    .stApp {
        font-family: 'Sarabun', sans-serif;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='text-align: center; margin-top: 50px;'>🏢 ระบบสารสนเทศทรัพยากรบุคคล (HR Systems)</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #666; font-size: 1.2rem; margin-bottom: 50px;'>ยินดีต้อนรับสู่ระบบบริหารจัดการส่วนบุคคล กรุณาเลือกระบบที่ต้องการใช้งาน</p>", unsafe_allow_html=True)

col1, col2 = st.columns(2)

with col1:
    st.markdown("""
    <div style="background-color: #f8f9fa; padding: 30px; border-radius: 15px; text-align: center; border: 1px solid #eee; height: 250px; display: flex; flex-direction: column; justify-content: center;">
        <h1 style="font-size: 3rem; margin: 0;">⏱️</h1>
        <h3 style="margin-top: 10px;">ระบบเข้า-ออกงาน</h3>
        <p style="color: #666; font-size: 0.9rem;">ตรวจสอบเวลาการทำงานและเวลาเข้า-ออกรายวัน</p>
    </div>
    <br>
    """, unsafe_allow_html=True)
    if st.button("เข้าสู่ระบบเข้า-ออกงาน", use_container_width=True, type="primary"):
        st.switch_page("pages/attendance.py")

with col2:
    st.markdown("""
    <div style="background-color: #f8f9fa; padding: 30px; border-radius: 15px; text-align: center; border: 1px solid #eee; height: 250px; display: flex; flex-direction: column; justify-content: center;">
        <h1 style="font-size: 3rem; margin: 0;">🏖️</h1>
        <h3 style="margin-top: 10px;">ระบบตรวจสอบวันลา</h3>
        <p style="color: #666; font-size: 0.9rem;">ตรวจสอบประวัติการลาป่วย ลากิจ ลาพักผ่อน และสาย</p>
    </div>
    <br>
    """, unsafe_allow_html=True)
    if st.button("เข้าสู่ระบบตรวจสอบวันลา", use_container_width=True, type="primary"):
        st.switch_page("pages/leave.py")
