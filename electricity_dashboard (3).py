"""
Streamlit dashboard: สถานะไฟฟ้า Air Compressor
แท็บ 1: Real-time monitor จาก Google Sheet ที่ Node-RED POST เข้าไป (ผ่าน Google Apps Script)
แท็บ 2: ปริมาณการใช้ไฟฟ้ารายวัน (On Peak / Off Peak / Total) จากชีต "Daily"

วิธีติดตั้ง:
    pip install streamlit pandas requests streamlit-autorefresh altair

วิธีรัน:
    streamlit run meter_dashboard.py
"""

import re
import streamlit as st
import pandas as pd
import altair as alt
import requests
from io import StringIO
from datetime import datetime, timedelta
from streamlit_autorefresh import st_autorefresh

# ============== CONFIG: แท็บ 1 (Real-time) ==============
# หา SHEET_ID จาก URL ของ Google Sheet:
# https://docs.google.com/spreadsheets/d/[SHEET_ID]/edit#gid=[GID]
SHEET_ID = "1gFKOoTb9XnarHqawBima7fuDb82yg5KXFqL15f_plpw"
GID = "0"  # เลข gid ของแท็บ (ดูจากท้าย URL หลัง #gid=)

COL_TIME = "Timestamp"
SHEET_COLS = {  # ชื่อคอลัมน์จริงในชีต real-time (ห้ามแก้)
    "ac1_3": "AC1-3",
    "ac4_6": "AC4-6",
    "ac7": "AC7",
    "ac8": "AC8",
}
DISPLAY_NAMES = {  # ชื่อที่อยากให้แสดงบนจอ (แก้ได้อิสระ)
    "ac1_3": "AC1-3 (Production)",
    "ac4_6": "AC4-6 (Production)",
    "ac7": "AC7 (Packing)",
    "ac8": "AC8 (Packing)",
}
RUN_STOP_THRESHOLD = 20  # >= ค่านี้ = Run, ต่ำกว่า = Stop
REFRESH_SEC = 60
TZ_OFFSET_HOURS = 7  # ชีตเป็นเวลาไทย (UTC+7) แต่ server รันเป็น UTC

# ============== CONFIG: แท็บ 2 (Daily Usage) ==============
DAILY_SHEET_ID = "1Ym2yfzkLTyLTtJtLZSSgWoeew_IPWUaI_u6d45jKUnw"
DAILY_GID = "0"
DAILY_SHEET_NAME = "Daily"

# ชื่อ Meter ตามคอลัมน์ A ในชีต Daily -> map ไปยัง key เดียวกับแท็บ 1 (ใช้ DISPLAY_NAMES ร่วมกัน)
DAILY_METER_NAME_MAP = {
    "MCC5_6": "ac4_6",
    "AirComp_P7": "ac7",
    "AirComp_P8": "ac8",
    "AirComp_P1234": "ac1_3",
}
DAILY_DATA_YEAR = 2026  # ปี ค.ศ. ของข้อมูล (ในชีตมีแค่ dd/mm ไม่มีปี)

# ==========================================================

CSV_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID}"
DAILY_CSV_URL = f"https://docs.google.com/spreadsheets/d/{DAILY_SHEET_ID}/export?format=csv&gid={DAILY_GID}"

st.set_page_config(page_title="AC Compressor Power Monitor", page_icon="⚡", layout="wide")
st_autorefresh(interval=REFRESH_SEC * 1000, key="refresh")

# ============== ปรับขนาดฟอนต์ตรงนี้ ==============
TITLE_FONT_SIZE = "2.2rem"
METRIC_LABEL_SIZE = "1.1rem"
METRIC_VALUE_SIZE = "2.8rem"
METRIC_DELTA_SIZE = "1rem"
CHART_LEGEND_SIZE = 14
CHART_AXIS_SIZE = 12
PAGE_TOP_PADDING = "1.5rem"  # ระยะห่างจากขอบบนสุดของหน้า ลดตัวเลขให้ชิดขึ้น

st.markdown(f"""
<style>
h1 {{ font-size: {TITLE_FONT_SIZE} !important; }}
div[data-testid="stMetricLabel"] p {{ font-size: {METRIC_LABEL_SIZE} !important; }}
div[data-testid="stMetricValue"] {{ font-size: {METRIC_VALUE_SIZE} !important; }}
div[data-testid="stMetricDelta"] {{ font-size: {METRIC_DELTA_SIZE} !important; }}
div.block-container {{ padding-top: {PAGE_TOP_PADDING} !important; }}
</style>
""", unsafe_allow_html=True)
# ====================================================


@st.cache_data(ttl=REFRESH_SEC)
def load_realtime_data():
    resp = requests.get(CSV_URL, timeout=10)
    resp.raise_for_status()
    resp.encoding = "utf-8"
    df = pd.read_csv(StringIO(resp.text))
    df[COL_TIME] = pd.to_datetime(df[COL_TIME], errors="coerce", dayfirst=True)
    df = df.dropna(subset=[COL_TIME]).sort_values(COL_TIME)
    return df


def get_status(value):
    if value >= RUN_STOP_THRESHOLD:
        return "Run", "normal"
    else:
        return "Stop", "off"


def _parse_number(x):
    if pd.isna(x):
        return 0.0
    s = str(x).replace(",", "").strip()
    if s == "":
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


@st.cache_data(ttl=300)
def load_daily_data():
    resp = requests.get(DAILY_CSV_URL, timeout=15)
    resp.raise_for_status()
    resp.encoding = "utf-8"
    raw = pd.read_csv(StringIO(resp.text), header=None)

    date_row = raw.iloc[0]
    n_cols = raw.shape[1]

    # หาบล็อกวัน (Onpeak/Offpeak/Total) เริ่มจากคอลัมน์ E (index 4) ทีละ 3 คอลัมน์
    day_blocks = []  # (onpeak_idx, offpeak_idx, total_idx, date, weekday_th)
    idx = 4
    while idx + 2 < n_cols:
        label = str(date_row[idx])
        m = re.search(r"(\d{2})/(\d{2})\s*\(([^)]+)\)", label)
        if m:
            dd, mm, wd = m.group(1), m.group(2), m.group(3)
            try:
                date_val = datetime(DAILY_DATA_YEAR, int(mm), int(dd))
            except ValueError:
                idx += 3
                continue
            day_blocks.append((idx, idx + 1, idx + 2, date_val, wd))
        idx += 3

    records = []
    for _, row in raw.iterrows():
        meter_name = str(row[0]).strip()
        if meter_name not in DAILY_METER_NAME_MAP:
            continue
        key = DAILY_METER_NAME_MAP[meter_name]
        for onpeak_idx, offpeak_idx, total_idx, date_val, wd in day_blocks:
            records.append({
                "meter_key": key,
                "meter_name": DISPLAY_NAMES.get(key, key),
                "date": date_val,
                "weekday": wd,
                "on_peak": _parse_number(row[onpeak_idx]),
                "off_peak": _parse_number(row[offpeak_idx]),
                "total": _parse_number(row[total_idx]),
            })

    long_df = pd.DataFrame(records)
    long_df = long_df.sort_values(["date", "meter_key"])
    return long_df


st.title("⚡ สถานะการใช้ไฟฟ้า Air Compressor")

tab1, tab2 = st.tabs(["📡 Real-time Monitor", "📊 การใช้ไฟฟ้ารายวัน"])

# ================= แท็บ 1: Real-time =================
with tab1:
    try:
        df = load_realtime_data()
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
                value = float(latest[SHEET_COLS[key]])
                status, color = get_status(value)
                col.metric(DISPLAY_NAMES[key], f"{value:.2f} kW", delta=status, delta_color=color)

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
                    color=alt.Color("กลุ่ม:N", legend=alt.Legend(title=None, labelFontSize=CHART_LEGEND_SIZE, symbolStrokeWidth=3)),
                )
                .properties(height=350)
            )
            st.altair_chart(line_chart, use_container_width=True)

            with st.expander("ดูข้อมูลดิบล่าสุด 20 แถว"):
                st.dataframe(df.tail(20).sort_values(COL_TIME, ascending=False), use_container_width=True)

    except Exception as e:
        st.error(f"โหลดข้อมูลไม่สำเร็จ: {e}")
        st.info("ตรวจสอบว่า SHEET_ID / GID ถูกต้อง และชีตเปิด public (Anyone with link can view)")

# ================= แท็บ 2: Daily Usage =================
with tab2:
    try:
        daily_df = load_daily_data()
        if daily_df.empty:
            st.warning("ไม่พบข้อมูลในชีต Daily ตรวจสอบชื่อ Meter / SHEET_ID / GID")
        else:
            month_options = sorted(daily_df["date"].dt.to_period("M").unique())
            month_labels = {p: p.strftime("%B %Y") for p in month_options}
            selected_month = st.selectbox(
                "เลือกเดือน",
                options=month_options,
                format_func=lambda p: month_labels[p],
                index=len(month_options) - 1,
            )

            metric_choice = st.radio(
                "แสดงค่า", options=["total", "on_peak", "off_peak"],
                format_func=lambda x: {"total": "Total", "on_peak": "On Peak", "off_peak": "Off Peak"}[x],
                horizontal=True,
            )

            filtered = daily_df[daily_df["date"].dt.to_period("M") == selected_month].copy()
            filtered["day_label"] = filtered["date"].dt.strftime("%d/%m") + " (" + filtered["weekday"] + ")"

            st.markdown(f"**สรุปยอดใช้ไฟฟ้าเดือน {month_labels[selected_month]}**")
            summary = filtered.groupby("meter_name", sort=False)[["on_peak", "off_peak", "total"]].sum()
            summary = summary.reindex([DISPLAY_NAMES[k] for k in ["ac1_3", "ac4_6", "ac7", "ac8"]])

            card_cols = st.columns(4)
            for card_col, meter_name in zip(card_cols, summary.index):
                row = summary.loc[meter_name]
                card_col.metric(
                    meter_name,
                    f"{row['total']:,.0f} kWh",
                    delta=f"On Peak {row['on_peak']:,.0f} | Off Peak {row['off_peak']:,.0f}",
                    delta_color="off",
                )
            st.metric("รวมทั้งหมด (Total เดือนนี้)", f"{summary['total'].sum():,.0f} kWh")

            bar_chart = (
                alt.Chart(filtered)
                .mark_bar()
                .encode(
                    x=alt.X("day_label:N", sort=None, title=None,
                            axis=alt.Axis(labelFontSize=CHART_AXIS_SIZE, titleFontSize=CHART_AXIS_SIZE, labelAngle=-45)),
                    y=alt.Y(f"{metric_choice}:Q", title="kWh",
                            axis=alt.Axis(labelFontSize=CHART_AXIS_SIZE, titleFontSize=CHART_AXIS_SIZE)),
                    color=alt.Color("meter_name:N", legend=alt.Legend(title=None, labelFontSize=CHART_LEGEND_SIZE)),
                    xOffset="meter_name:N",
                    tooltip=["day_label", "meter_name", "on_peak", "off_peak", "total"],
                )
                .properties(height=420)
            )
            st.altair_chart(bar_chart, use_container_width=True)

            with st.expander("ดูข้อมูลดิบ"):
                st.dataframe(
                    filtered[["date", "weekday", "meter_name", "on_peak", "off_peak", "total"]]
                    .sort_values(["date", "meter_name"]),
                    use_container_width=True,
                )

    except Exception as e:
        st.error(f"โหลดข้อมูล Daily ไม่สำเร็จ: {e}")
        st.info("ตรวจสอบว่า DAILY_SHEET_ID / DAILY_GID ถูกต้อง และชีตเปิด public (Anyone with link can view)")
