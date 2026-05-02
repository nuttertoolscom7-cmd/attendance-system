import streamlit as st

st.set_page_config(page_title="HR Systems Portal", page_icon="🏢", layout="wide")

pages = {
    "หน้าแรก": [
        st.Page("home.py", title="เมนูหลัก", icon="🏠", default=True),
    ],
    "ระบบงาน HR": [
        st.Page("pages/attendance.py", title="ระบบเข้า-ออกงาน", icon="⏱️"),
        st.Page("pages/leave.py", title="ระบบตรวจสอบวันลา", icon="🏖️"),
    ]
}

pg = st.navigation(pages)
pg.run()
