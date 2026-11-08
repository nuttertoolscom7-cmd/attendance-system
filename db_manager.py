import os
import sqlite3
import pandas as pd
from datetime import datetime
from io import BytesIO

DB_FILENAME = "attendance.db"

def get_db_path() -> str:
    """คืนค่าที่อยู่ของไฟล์ฐานข้อมูล SQLite ในโฟลเดอร์โปรเจกต์"""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_dir, DB_FILENAME)

def get_connection(db_path: str = None) -> sqlite3.Connection:
    """เปิดการเชื่อมต่อกับฐานข้อมูล SQLite"""
    path = db_path or get_db_path()
    conn = sqlite3.connect(path, timeout=15)
    conn.row_factory = sqlite3.Row
    return conn

def init_db(db_path: str = None):
    """สร้างตารางฐานข้อมูลและดัชนี (Index) หากยังไม่มี"""
    conn = get_connection(db_path)
    cursor = conn.cursor()

    # 1. ตารางบันทึกเวลาเข้า-ออกงานระดับรายคนรายวัน
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS attendance_records (
        date TEXT NOT NULL,
        name TEXT NOT NULL,
        department TEXT,
        check_in TEXT,
        check_out TEXT,
        status TEXT,
        has_leave INTEGER DEFAULT 0,
        leave_type TEXT DEFAULT '-',
        leave_remark TEXT DEFAULT '',
        updated_at TEXT,
        PRIMARY KEY (date, name)
    )
    """)

    # สร้างดัชนีเพื่อการค้นหาและจัดกลุ่มที่รวดเร็ว
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_att_date ON attendance_records(date)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_att_name ON attendance_records(name)")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_att_dept ON attendance_records(department)")

    # 2. ตารางประวัติการนำเข้าไฟล์ (Import Batches History)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS import_batches (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        import_time TEXT NOT NULL,
        source_label TEXT,
        records_count INTEGER,
        dates_summary TEXT,
        months_affected TEXT
    )
    """)

    conn.commit()
    conn.close()

THAI_MONTHS_NAMES = [
    "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
    "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
]

def thai_month_name(month_num: int) -> str:
    if 1 <= month_num <= 12:
        return THAI_MONTHS_NAMES[month_num - 1]
    return ""

def save_attendance_dataframe(df: pd.DataFrame, source_label: str = "อัปโหลด PDF", db_path: str = None) -> dict:
    """
    บันทึกข้อมูลเข้าฐานข้อมูลแบบ Smart Merge (Upsert)
    - หากเป็นข้อมูลวันใหม่ -> INSERT
    - หากเป็นข้อมูลวันเดิมที่มีอยู่แล้ว -> UPDATE ข้อมูลให้ทันสมัยที่สุด
    คืนค่าสรุปผลการบันทึก
    """
    if df is None or df.empty:
        return {"status": "empty", "saved_count": 0, "months_affected": []}

    init_db(db_path)
    conn = get_connection(db_path)
    cursor = conn.cursor()

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    dates_seen = set()
    months_seen = set()

    upsert_sql = """
    INSERT INTO attendance_records (
        date, name, department, check_in, check_out, status,
        has_leave, leave_type, leave_remark, updated_at
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(date, name) DO UPDATE SET
        department = excluded.department,
        check_in = excluded.check_in,
        check_out = excluded.check_out,
        status = excluded.status,
        has_leave = excluded.has_leave,
        leave_type = excluded.leave_type,
        leave_remark = excluded.leave_remark,
        updated_at = excluded.updated_at
    """

    records_to_insert = []
    for _, row in df.iterrows():
        raw_date = row.get("วันที่")
        if pd.isna(raw_date):
            continue
        if hasattr(raw_date, "strftime"):
            date_str = raw_date.strftime("%Y-%m-%d")
        else:
            date_str = str(raw_date)[:10]

        name = str(row.get("ชื่อ-สกุล", "")).strip()
        dept = str(row.get("แผนก", "")).strip() if pd.notna(row.get("แผนก")) else ""
        cin = str(row.get("เข้างาน", "")).strip() if pd.notna(row.get("เข้างาน")) and row.get("เข้างาน") else None
        cout = str(row.get("ออกงาน", "")).strip() if pd.notna(row.get("ออกงาน")) and row.get("ออกงาน") else None
        status = str(row.get("สถานะ", "ปกติ")).strip()

        has_leave_val = 1 if bool(row.get("มีใบลา", False)) else 0
        l_type = str(row.get("ประเภทการลา", "-")).strip() if pd.notna(row.get("ประเภทการลา")) else "-"
        l_remark = str(row.get("หมายเหตุการลา", "")).strip() if pd.notna(row.get("หมายเหตุการลา")) else ""

        dates_seen.add(date_str)
        months_seen.add(date_str[:7])

        records_to_insert.append((
            date_str, name, dept, cin, cout, status,
            has_leave_val, l_type, l_remark, now_str
        ))

    cursor.executemany(upsert_sql, records_to_insert)

    # บันทึกประวัติการนำเข้า
    sorted_dates = sorted(dates_seen)
    dates_summary = f"{sorted_dates[0]} ถึง {sorted_dates[-1]}" if len(sorted_dates) > 1 else (sorted_dates[0] if sorted_dates else "")
    months_str = ", ".join(sorted(months_seen))

    cursor.execute("""
    INSERT INTO import_batches (import_time, source_label, records_count, dates_summary, months_affected)
    VALUES (?, ?, ?, ?, ?)
    """, (now_str, source_label, len(records_to_insert), dates_summary, months_str))

    conn.commit()
    conn.close()

    return {
        "status": "success",
        "saved_count": len(records_to_insert),
        "days_count": len(dates_seen),
        "dates_summary": dates_summary,
        "months_affected": sorted(months_seen)
    }

def get_available_months(db_path: str = None) -> list:
    """
    ดึงรายชื่อปี-เดือนทั้งหมดที่มีในฐานข้อมูล พร้อมสถิติจำนวนวันและบุคลากร
    คืนค่าเป็นรายการ dict เรียงจากเดือนล่าสุดไปหาอดีต
    """
    init_db(db_path)
    conn = get_connection(db_path)
    cursor = conn.cursor()

    cursor.execute("""
    SELECT 
        substr(date, 1, 7) AS ym,
        count(DISTINCT date) AS days_count,
        count(DISTINCT name) AS people_count,
        count(*) AS records_count,
        min(date) AS min_date,
        max(date) AS max_date
    FROM attendance_records
    GROUP BY ym
    ORDER BY ym DESC
    """)
    rows = cursor.fetchall()
    conn.close()

    results = []
    for r in rows:
        ym = r["ym"]
        try:
            parts = ym.split("-")
            year = int(parts[0])
            month = int(parts[1])
            year_be = year + 543
            m_th = thai_month_name(month)
        except Exception:
            year, month, year_be, m_th = 0, 0, 0, ym

        days_count = r["days_count"]
        people_count = r["people_count"]
        records_count = r["records_count"]

        # วันที่ล่าสุดในเดือนนั้น
        max_d = r["max_date"]
        try:
            max_dt = datetime.strptime(max_d, "%Y-%m-%d")
            max_d_str = f"{max_dt.day} {m_th} {year_be}"
        except Exception:
            max_d_str = max_d

        label = f"📅 {m_th} {year_be} (สะสม {days_count} วัน | ล่าสุด: {max_d_str} | {people_count} คน)"

        results.append({
            "ym": ym,
            "year": year,
            "month": month,
            "year_be": year_be,
            "month_name": m_th,
            "days_count": days_count,
            "people_count": people_count,
            "records_count": records_count,
            "min_date": r["min_date"],
            "max_date": r["max_date"],
            "label": label
        })

    return results

def load_attendance_by_month(ym: str, db_path: str = None) -> pd.DataFrame:
    """
    โหลดข้อมูลประจำเดือนตามรหัสปี-เดือน (เช่น '2026-10') ออกมาเป็น DataFrame
    ในรูปแบบเดียวกับ master_df ของระบบ
    """
    init_db(db_path)
    conn = get_connection(db_path)

    query = """
    SELECT 
        department AS "แผนก",
        name AS "ชื่อ-สกุล",
        date AS "วันที่",
        check_in AS "เข้างาน",
        check_out AS "ออกงาน",
        status AS "สถานะ",
        has_leave AS "มีใบลา",
        leave_type AS "ประเภทการลา",
        leave_remark AS "หมายเหตุการลา"
    FROM attendance_records
    WHERE substr(date, 1, 7) = ?
    ORDER BY date ASC, department ASC, name ASC
    """
    df = pd.read_sql_query(query, conn, params=(ym,))
    conn.close()

    if df.empty:
        return df

    # แปลง date กลับเป็น pd.Timestamp
    df["วันที่"] = pd.to_datetime(df["วันที่"])
    df["มีใบลา"] = df["มีใบลา"].astype(bool)

    # เคลียร์ค่า None/NaN ให้เข้ากับระบบ
    df["เข้างาน"] = df["เข้างาน"].apply(lambda v: v if pd.notna(v) and str(v).strip() != "" else None)
    df["ออกงาน"] = df["ออกงาน"].apply(lambda v: v if pd.notna(v) and str(v).strip() != "" else None)

    return df

def delete_month_attendance(ym: str, db_path: str = None) -> int:
    """ลบข้อมูลของเดือนที่ระบุออกจากฐานข้อมูล"""
    init_db(db_path)
    conn = get_connection(db_path)
    cursor = conn.cursor()

    cursor.execute("DELETE FROM attendance_records WHERE substr(date, 1, 7) = ?", (ym,))
    deleted = cursor.rowcount
    conn.commit()
    conn.close()
    return deleted

def get_database_stats(db_path: str = None) -> dict:
    """คืนค่าสถิติภาพรวมของฐานข้อมูล"""
    init_db(db_path)
    path = db_path or get_db_path()
    size_bytes = os.path.getsize(path) if os.path.exists(path) else 0
    size_kb = size_bytes / 1024

    conn = get_connection(db_path)
    cursor = conn.cursor()

    cursor.execute("SELECT count(*) as total_records, count(DISTINCT name) as total_people, count(DISTINCT date) as total_days, count(DISTINCT substr(date, 1, 7)) as total_months FROM attendance_records")
    row = cursor.fetchone()
    conn.close()

    return {
        "total_records": row["total_records"] if row else 0,
        "total_people": row["total_people"] if row else 0,
        "total_days": row["total_days"] if row else 0,
        "total_months": row["total_months"] if row else 0,
        "size_kb": size_kb,
        "size_display": f"{size_kb:.1f} KB" if size_kb < 1024 else f"{size_kb/1024:.2f} MB"
    }

def get_database_bytes(db_path: str = None) -> bytes:
    """อ่านข้อมูลไฟล์ SQLite ออกมาเป็น bytes สำหรับดาวน์โหลดสำรองข้อมูล"""
    path = db_path or get_db_path()
    if os.path.exists(path):
        with open(path, "rb") as f:
            return f.read()
    return b""
