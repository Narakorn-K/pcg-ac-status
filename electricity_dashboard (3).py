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
import altair as alt
import requests
from io import StringIO
from datetime import datetime, timedelta
from streamlit_autorefresh import st_autorefresh

# ============== CONFIG (แก้ตรงนี้ให้ตรงกับของคุณ) ==============
# หา SHEET_ID จาก URL ของ Google Sheet:
# https://docs.google.com/spreadsheets/d/[SHEET_ID]/edit#gid=[GID]
SHEET_ID = "1gFKOoTb9XnarHqawBima7fuDb82yg5KXFqL15f_plpw"
GID = "0"  # เลข gid ของแท็บ (ดูจากท้าย URL หลัง #gid=)

# ชื่อคอลัมน์ตามหัวตารางในชีต (ต้องตรงกับที่เห็นในสกรีนช็อต ห้ามแก้)
COL_TIME = "Timestamp"
SHEET_COLS = {
    "ac1_3": "AC1-3",
    "ac4_6": "AC4-6",
    "ac7": "AC7",
    "ac8": "AC8",
}

# ชื่อที่อยากให้แสดงบนหน้าจอ (แก้ได้อิสระ ไม่ต้องตรงกับชีต)
DISPLAY_NAMES = {
    "ac1_3": "AC1-3 (Production)",
    "ac4_6": "AC4-6 (Production)",
    "ac7": "AC7 (Packing)",
    "ac8": "AC8 (Packing)",
}

# threshold สถานะ Run/Stop (kW) ใช้เกณฑ์เดียวกันทุก AC
RUN_STOP_THRESHOLD = 20  # >= ค่านี้ = Run, ต่ำกว่า = Stop

REFRESH_SEC = 15  # ความถี่ในการอัปเดตหน้าจอ (วินาที)
TZ_OFFSET_HOURS = 7  # timestamp ในชีตเป็นเวลาไทย (UTC+7) แต่ server รันเป็น UTC เลยต้องชดเชยตรงนี้
# ================================================================

CSV_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID}"

st.set_page_config(page_title="AC Compressor Power Monitor", page_icon="⚡", layout="wide")
st_autorefresh(interval=REFRESH_SEC * 1000, key="refresh")

# ============== ปรับขนาดฟอนต์ตรงนี้ ==============
TITLE_FONT_SIZE = "2.2rem"    # หัวข้อบนสุด
METRIC_LABEL_SIZE = "1.1rem"  # ชื่อหัวข้อในแต่ละกล่อง (เช่น AC1-3)
METRIC_VALUE_SIZE = "2.8rem"  # ตัวเลข kW ตัวใหญ่
METRIC_DELTA_SIZE = "1rem"    # ข้อความสถานะเล็กใต้ตัวเลข
CHART_LEGEND_SIZE = 14        # ขนาดตัวอักษร legend ใต้กราฟ (หน่วย px ไม่ใช่ rem)
CHART_AXIS_SIZE = 12          # ขนาดตัวอักษรแกนกราฟ (หน่วย px)

st.markdown(f"""
<style>
h1 {{ font-size: {TITLE_FONT_SIZE} !important; }}
div[data-testid="stMetricLabel"] p {{ font-size: {METRIC_LABEL_SIZE} !important; }}
div[data-testid="stMetricValue"] {{ font-size: {METRIC_VALUE_SIZE} !important; }}
div[data-testid="stMetricDelta"] {{ font-size: {METRIC_DELTA_SIZE} !important; }}
</style>
""", unsafe_allow_html=True)
# ====================================================


@st.cache_data(ttl=REFRESH_SEC)
def load_data():
    resp = requests.get(CSV_URL, timeout=10)
    resp.raise_for_status()
    df = pd.read_csv(StringIO(resp.text))
    df[COL_TIME] = pd.to_datetime(df[COL_TIME], errors="coerce", dayfirst=True)
    df = df.dropna(subset=[COL_TIME]).sort_values(COL_TIME)
    return df


def get_status(value, key=None):
    if value >= RUN_STOP_THRESHOLD:
        return "Run", "normal"
    else:
        return "Stop", "off"


st.title("⚡ สถานะการใช้ไฟฟ้า Air Compressor")

try:
    df = load_data()
    if df.empty:
        st.warning("ยังไม่มีข้อมูลใน Google Sheet")
    else:
        latest = df.iloc[-1]
        last_time = latest[COL_TIME]
        now_th = datetime.utcnow() + timedelta(hours=TZ_OFFSET_HOURS)
        age_sec = (now_th - last_time.to_pydatetime().replace(tzinfo=None)).total_seconds()

        st.caption(f"อัปเดตล่าสุด: {last_time.strftime('%Y-%m-%d %H:%M:%S')} ({int(age_sec)} วินาทีที่แล้ว)")
        if age_sec > REFRESH_SEC * 4:
            st.error("⚠️ ไม่มีข้อมูลใหม่เข้ามานานผิดปกติ ตรวจสอบการเชื่อมต่อ Node-RED")

        cols = st.columns(4)
        for col, key in zip(cols, SHEET_COLS.keys()):
            sheet_col = SHEET_COLS[key]
            label = DISPLAY_NAMES[key]
            value = float(latest[sheet_col])
            status, color = get_status(value, key)
            col.metric(label, f"{value:.2f} kW", delta=status, delta_color=color)

        total_kw = sum(float(latest[c]) for c in SHEET_COLS.values())
        st.metric("รวมทั้งหมด (Total)", f"{total_kw:.2f} kW")

        st.subheader("แนวโน้มย้อนหลัง")
        chart_df = df.tail(120)[[COL_TIME] + list(SHEET_COLS.values())]
        chart_df = chart_df.rename(columns={SHEET_COLS[k]: DISPLAY_NAMES[k] for k in SHEET_COLS.keys()})
        chart_long = chart_df.melt(id_vars=COL_TIME, var_name="กลุ่ม", value_name="kW")

        line_chart = (
            alt.Chart(chart_long)
            .mark_line()
            .encode(
                x=alt.X(f"{COL_TIME}:T", title=None, axis=alt.Axis(labelFontSize=CHART_AXIS_SIZE, titleFontSize=CHART_AXIS_SIZE)),
                y=alt.Y("kW:Q", axis=alt.Axis(labelFontSize=CHART_AXIS_SIZE, titleFontSize=CHART_AXIS_SIZE)),
                color=alt.Color(
                    "กลุ่ม:N",
                    legend=alt.Legend(title=None, labelFontSize=CHART_LEGEND_SIZE, symbolStrokeWidth=3),
                ),
            )
            .properties(height=350)
        )
        st.altair_chart(line_chart, use_container_width=True)

        with st.expander("ดูข้อมูลดิบล่าสุด 20 แถว"):
            st.dataframe(df.tail(20).sort_values(COL_TIME, ascending=False), use_container_width=True)

except Exception as e:
    st.error(f"โหลดข้อมูลไม่สำเร็จ: {e}")
    st.info("ตรวจสอบว่า SHEET_ID / GID ถูกต้อง และชีตเปิด public (Anyone with link can view)")
