"""
Streamlit dashboard: แสดงค่าไฟฟ้า (kW) ของ Air Compressor 4 กลุ่ม แบบ real-time
จาก Google Sheet ที่ Node-RED POST ข้อมูลเข้าไป (ผ่าน Google Apps Script)

คอลัมน์ใน Google Sheet: Timestamp | AC1-3 | AC4-6 | AC7 | AC8

วิธีติดตั้ง:
    pip install streamlit pandas requests streamlit-autorefresh

วิธีรัน:
    streamlit run meter_dashboard.py
"""

import streamlit as st
import pandas as pd
import requests
from io import StringIO
from datetime import datetime
from streamlit_autorefresh import st_autorefresh

# ============== CONFIG (แก้ตรงนี้ให้ตรงกับของคุณ) ==============
# หา SHEET_ID จาก URL ของ Google Sheet:
# https://docs.google.com/spreadsheets/d/[SHEET_ID]/edit#gid=[GID]
SHEET_ID = "YOUR_SHEET_ID_HERE"
GID = "0"  # เลข gid ของแท็บ (ดูจากท้าย URL หลัง #gid=)

# ชื่อคอลัมน์ตามหัวตารางในชีต (ต้องตรงกับที่เห็นในสกรีนช็อต)
COL_TIME = "Timestamp"
AC_COLS = {
    "ac1_3": "AC1-3",
    "ac4_6": "AC4-6",
    "ac7": "AC7",
    "ac8": "AC8",
}

# threshold สำหรับแสดงสถานะแยกตามกลุ่ม (kW) ปรับตามพิกัดเครื่องจริงของแต่ละกลุ่ม
THRESHOLDS = {
    "ac1_3": {"warning": 55, "critical": 65},
    "ac4_6": {"warning": 240, "critical": 260},
    "ac7":   {"warning": 95,  "critical": 105},
    "ac8":   {"warning": 95,  "critical": 105},
}

REFRESH_SEC = 15  # ความถี่ในการอัปเดตหน้าจอ (วินาที)
# ================================================================

CSV_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID}"

st.set_page_config(page_title="AC Compressor Power Monitor", page_icon="⚡", layout="wide")
st_autorefresh(interval=REFRESH_SEC * 1000, key="refresh")


@st.cache_data(ttl=REFRESH_SEC)
def load_data():
    resp = requests.get(CSV_URL, timeout=10)
    resp.raise_for_status()
    df = pd.read_csv(StringIO(resp.text))
    df[COL_TIME] = pd.to_datetime(df[COL_TIME], errors="coerce", dayfirst=True)
    df = df.dropna(subset=[COL_TIME]).sort_values(COL_TIME)
    return df


def get_status(value, key):
    th = THRESHOLDS[key]
    if value >= th["critical"]:
        return "สูงเกินกำหนด", "inverse"
    elif value >= th["warning"]:
        return "ใกล้เต็มพิกัด", "off"
    else:
        return "ปกติ", "normal"


st.title("⚡ สถานะการใช้ไฟฟ้า Air Compressor")

try:
    df = load_data()
    if df.empty:
        st.warning("ยังไม่มีข้อมูลใน Google Sheet")
    else:
        latest = df.iloc[-1]
        last_time = latest[COL_TIME]
        age_sec = (datetime.now() - last_time.to_pydatetime().replace(tzinfo=None)).total_seconds()

        st.caption(f"อัปเดตล่าสุด: {last_time.strftime('%Y-%m-%d %H:%M:%S')} ({int(age_sec)} วินาทีที่แล้ว)")
        if age_sec > REFRESH_SEC * 4:
            st.error("⚠️ ไม่มีข้อมูลใหม่เข้ามานานผิดปกติ ตรวจสอบการเชื่อมต่อ Node-RED")

        cols = st.columns(4)
        for col, (key, sheet_col) in zip(cols, AC_COLS.items()):
            value = float(latest[sheet_col])
            status, color = get_status(value, key)
            col.metric(sheet_col, f"{value:.2f} kW", delta=status, delta_color=color)

        total_kw = sum(float(latest[c]) for c in AC_COLS.values())
        st.metric("รวมทั้งหมด (Total)", f"{total_kw:.2f} kW")

        st.subheader("แนวโน้มย้อนหลัง")
        chart_df = df.tail(120).set_index(COL_TIME)[list(AC_COLS.values())]
        st.line_chart(chart_df)

        with st.expander("ดูข้อมูลดิบล่าสุด 20 แถว"):
            st.dataframe(df.tail(20).sort_values(COL_TIME, ascending=False), use_container_width=True)

except Exception as e:
    st.error(f"โหลดข้อมูลไม่สำเร็จ: {e}")
    st.info("ตรวจสอบว่า SHEET_ID / GID ถูกต้อง และชีตเปิด public (Anyone with link can view)")
