import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import re
from datetime import datetime, timedelta
from urllib.parse import quote
import sys

def fmt_day(dt, show_year=True):
    """Cross-platform: no leading zero on day, works on Windows and Linux/Mac"""
    if sys.platform == 'win32':
        return dt.strftime('%#d %b %Y') if show_year else dt.strftime('%#d %b')
    return dt.strftime('%-d %b %Y') if show_year else dt.strftime('%-d %b')

def fmt_day_short(dt):
    return fmt_day(dt, show_year=False)

# ─── Page Config ──────────────────────────────────────────────────────────────
st.set_page_config(page_title="Electricity Dashboard", layout="wide", page_icon="💡")

# ─── Google Sheet Config ──────────────────────────────────────────────────────
SHEET_ID        = "1Ym2yfzkLTyLTtJtLZSSgWoeew_IPWUaI_u6d45jKUnw"
SHEET_NAME      = "Daily"
SHEET_TON       = "Product Ton"
EXPORT_URL      = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx&sheet={quote(SHEET_NAME)}"
EXPORT_URL_TON  = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx&sheet={quote(SHEET_TON)}"

# ─── Thai electricity tariff ──────────────────────────────────────────────────
ON_PEAK_RATE  = 4.1824
OFF_PEAK_RATE = 2.6369
FT_ADJ        = 0.1623

# ─── Thai day / month abbreviations ──────────────────────────────────────────
DAY_TH = {"อา": 6, "จ": 0, "อ": 1, "พ": 2, "พฤ": 3, "ศ": 4, "ส": 5}
MONTH_TH = {
    1: "ม.ค.", 2: "ก.พ.", 3: "มี.ค.", 4: "เม.ย.",
    5: "พ.ค.", 6: "มิ.ย.", 7: "ก.ค.", 8: "ส.ค.",
    9: "ก.ย.", 10: "ต.ค.", 11: "พ.ย.", 12: "ธ.ค.",
}

# ─── CSS ─────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-title {
        font-size:26px; font-weight:700; text-align:center;
        color:#1a237e; margin-bottom:4px;
    }
    .week-subtitle, .month-subtitle {
        font-size:16px; text-align:center; color:#546e7a;
        margin-bottom:20px; font-weight:500;
    }
    .kpi-card {
        background:#fff; border-radius:12px; padding:18px 22px;
        box-shadow:0 2px 10px rgba(0,0,0,0.08); text-align:center;
        border-top: 4px solid #1565c0;
    }
    .kpi-label {font-size:14px; color:#666; font-weight:600; margin-bottom:6px;
                text-transform:uppercase; letter-spacing:0.5px;}
    .kpi-value {font-size:26px; font-weight:800; color:#1a237e;}
    .kpi-unit  {font-size:16px; font-weight:400;}
    .kpi-sub   {font-size:13px; color:#999; margin-top:4px;}
    .kpi-on    {color:#e65100 !important;}
    .kpi-off   {color:#2e7d32 !important;}
    .kpi-cost  {color:#6a1b9a !important;}
    .kpi-ton   {color:#00838f !important;}
    .up   {color:#e53935; font-weight:700;}
    .down {color:#43a047; font-weight:700;}
    .section-header {
        font-size:18px; font-weight:700; color:#1a237e;
        border-left:4px solid #1565c0; padding-left:10px;
        margin:28px 0 14px;
    }
</style>
""", unsafe_allow_html=True)

# ─── Helpers ──────────────────────────────────────────────────────────────────
def parse_date_col(raw_val: str, fallback_year: int = 2026):
    match = re.match(r"(\d{2}/\d{2})", str(raw_val))
    if not match:
        return None, None
    date_str = match.group(1)
    d, m = map(int, date_str.split("/"))
    for yr in [fallback_year, fallback_year - 1, datetime.now().year]:
        try:
            dt = datetime(yr, m, d)
            day_match = re.search(r"\((.+?)\)", str(raw_val))
            wd = DAY_TH.get(day_match.group(1), dt.weekday()) if day_match else dt.weekday()
            return dt, wd
        except ValueError:
            continue
    return None, None


def badge(v):
    cls   = "up" if v >= 0 else "down"
    arrow = "▲" if v >= 0 else "▼"
    return f'<span class="{cls}">{arrow} {abs(v):.1f}%</span>'


# ─── Load Both Sheets in ONE request ─────────────────────────────────────────
@st.cache_data(ttl=300)
def load_all_bytes() -> bytes:
    """ดึง Excel bytes ครั้งเดียว cache ได้เลย"""
    import urllib.request
    BASE_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx"
    try:
        req  = urllib.request.Request(BASE_URL, headers={"User-Agent": "Mozilla/5.0"})
        data = urllib.request.urlopen(req, timeout=60).read()
    except Exception as e:
        st.error(f"❌ ไม่สามารถดึงข้อมูลจาก Google Sheet ได้: {e}")
        st.stop()
    return data

@st.cache_data(ttl=300)
def load_data():
    import io
    xls = pd.ExcelFile(io.BytesIO(load_all_bytes()))
    try:
        raw = xls.parse(SHEET_NAME, header=None)
    except Exception as e:
        st.error(f"❌ ไม่พบ Sheet '{SHEET_NAME}': {e}")
        st.stop()

    header_row = 0
    for r in range(min(6, raw.shape[0])):
        if raw.iloc[r, :].astype(str).str.contains(r"\d{2}/\d{2}", regex=True).any():
            header_row = r
            break

    data_start_row = header_row + 2
    fallback_year  = datetime.now().year

    date_cols = []
    for i in range(4, raw.shape[1]):
        val = raw.iloc[header_row, i]
        if pd.notna(val):
            dt, wd = parse_date_col(str(val), fallback_year)
            if dt:
                date_cols.append({"col_idx": i, "date": dt, "weekday": wd})

    if not date_cols:
        st.error("❌ ไม่พบคอลัมน์วันที่")
        st.stop()

    records = []
    for row_i in range(data_start_row, raw.shape[0]):
        meter  = raw.iloc[row_i, 0]
        group  = raw.iloc[row_i, 2]
        subgrp = raw.iloc[row_i, 3]
        if pd.isna(meter) or pd.isna(group):
            continue
        if str(group).strip() in ("nan", ""):
            continue
        for dc in date_cols:
            ci  = dc["col_idx"]
            on  = pd.to_numeric(raw.iloc[row_i, ci],     errors="coerce") if ci     < raw.shape[1] else 0
            off = pd.to_numeric(raw.iloc[row_i, ci + 1], errors="coerce") if ci + 1 < raw.shape[1] else 0
            tot = pd.to_numeric(raw.iloc[row_i, ci + 2], errors="coerce") if ci + 2 < raw.shape[1] else 0
            records.append({
                "meter":      str(meter),
                "department": str(group).strip(),
                "sub_group":  str(subgrp).strip(),
                "date":       dc["date"],
                "weekday":    dc["weekday"],
                "on_peak":    float(on)  if pd.notna(on)  else 0.0,
                "off_peak":   float(off) if pd.notna(off) else 0.0,
                "total":      float(tot) if pd.notna(tot) else 0.0,
            })

    if not records:
        st.error("❌ ไม่พบแถวข้อมูล")
        st.dataframe(raw.iloc[:5, :8])
        st.stop()

    df = pd.DataFrame(records)
    df["date"]      = pd.to_datetime(df["date"])
    df["week_num"]  = df["date"].dt.isocalendar().week.astype(int)
    df["year"]      = df["date"].dt.isocalendar().year.astype(int)
    df["year_week"] = df["year"].astype(str) + "-W" + df["week_num"].astype(str).str.zfill(2)
    df["month"]     = df["date"].dt.month
    df["ym"]        = df["year"].astype(str) + "-" + df["month"].astype(str).str.zfill(2)
    return df


# ─── Load Product Ton Data (reuse same download) ─────────────────────────────
@st.cache_data(ttl=300)
def load_ton_data():
    import io
    xls = pd.ExcelFile(io.BytesIO(load_all_bytes()))
    try:
        ton = xls.parse(SHEET_TON, header=0)
    except Exception as e:
        st.warning(f"⚠️ ไม่พบ Sheet '{SHEET_TON}': {e}")
        return pd.DataFrame(columns=["date", "ton"])

    # Column A = Date, Column B = Ton
    ton.columns = ["date", "ton"] + list(ton.columns[2:])
    ton = ton[["date", "ton"]].copy()
    ton["date"] = pd.to_datetime(ton["date"], dayfirst=True, errors="coerce")
    ton["ton"]  = pd.to_numeric(ton["ton"], errors="coerce")
    ton = ton.dropna(subset=["date", "ton"])
    ton["week_num"]  = ton["date"].dt.isocalendar().week.astype(int)
    ton["year"]      = ton["date"].dt.isocalendar().year.astype(int)
    ton["year_week"] = ton["year"].astype(str) + "-W" + ton["week_num"].astype(str).str.zfill(2)
    ton["month"]     = ton["date"].dt.month
    ton["ym"]        = ton["year"].astype(str) + "-" + ton["month"].astype(str).str.zfill(2)
    return ton


# ─── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚡ Electricity Dashboard")
    st.markdown("---")
    if st.button("🔄 Refresh ข้อมูล", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
    st.markdown("---")
    st.caption("📡 ดึงข้อมูลจาก Google Sheet อัตโนมัติ")
    st.caption("🔄 Auto-refresh ทุก 5 นาที")
    st.markdown("---")
    st.caption("**อัตราค่าไฟ MEA TOU**")
    st.caption(f"• On Peak  : {ON_PEAK_RATE + FT_ADJ:.4f} ฿/kWh")
    st.caption(f"• Off Peak : {OFF_PEAK_RATE + FT_ADJ:.4f} ฿/kWh")
    st.caption(f"• Ft Surcharge : {FT_ADJ} ฿/kWh")

# ─── Load Data ────────────────────────────────────────────────────────────────
with st.spinner("⏳ กำลังดึงข้อมูลจาก Google Sheet..."):
    df  = load_data()
    ton = load_ton_data()

# ─── Tabs ─────────────────────────────────────────────────────────────────────
tab_weekly, tab_wkton, tab_monthly, tab_monton = st.tabs(["📋 Weekly Overview", "📋 Weekly kWh/Ton", "📅 Monthly Dashboard", "📅 Monthly kWh/Ton"])


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — WEEKLY
# ══════════════════════════════════════════════════════════════════════════════
with tab_weekly:

    def week_label(year_week):
        sub = df[df["year_week"] == year_week]["date"]
        if sub.empty:
            return year_week
        start, end = sub.min(), sub.max()
        if start.month == end.month:
            return f"{start.day} – {fmt_day(end)}"
        return f"{fmt_day_short(start)} – {fmt_day(end)}"

    def get_all_weeks():
        return sorted(df["year_week"].unique().tolist())

    def get_complete_weeks():
        wk_days = df.groupby("year_week")["date"].nunique()
        return wk_days[wk_days == 7].index.tolist()

    def week_agg(year_week, dept_filter=None):
        sub = df[df["year_week"] == year_week]
        if dept_filter and dept_filter != "🏭 Factory (ทั้งหมด)":
            sub = sub[sub["department"] == dept_filter]
        return sub[["on_peak", "off_peak", "total"]].sum()

    def dept_week_agg(year_week):
        sub = df[df["year_week"] == year_week]
        return sub.groupby("department")[["on_peak", "off_peak", "total"]].sum().reset_index()

    def week_ton(year_week):
        if ton.empty:
            return 0.0
        return float(ton[ton["year_week"] == year_week]["ton"].sum())

    st.markdown('<div class="main-title">⚡ Electricity Usage Overview</div>', unsafe_allow_html=True)

    all_weeks      = get_all_weeks()
    complete_weeks = sorted(get_complete_weeks())

    if not all_weeks:
        st.error("ไม่พบข้อมูล")
        st.stop()

    # Week selector
    week_display_map = {}
    for yw in all_weeks:
        label      = week_label(yw)
        is_complete = yw in complete_weeks
        suffix     = "" if is_complete else " ⚠️ ไม่ครบ 7 วัน"
        week_display_map[yw] = f"{yw}  ·  {label}{suffix}"

    display_options = list(week_display_map.values())
    yw_keys         = list(week_display_map.keys())
    default_yw      = complete_weeks[-1] if complete_weeks else all_weeks[-1]
    default_idx     = yw_keys.index(default_yw)

    col_sel1, col_sel2, col_sel3 = st.columns([1, 2, 1])
    with col_sel2:
        selected_display = st.selectbox(
            "📅 เลือกสัปดาห์", options=display_options,
            index=default_idx, key="weekly_selector",
        )

    selected_yw  = yw_keys[display_options.index(selected_display)]
    selected_idx = yw_keys.index(selected_yw)
    prev_yw      = yw_keys[selected_idx - 1] if selected_idx > 0 else None
    sel_label    = week_label(selected_yw)
    prev_label   = week_label(prev_yw) if prev_yw else "N/A"

    st.markdown(f'<div class="week-subtitle">📆 สัปดาห์ {selected_yw} &nbsp;|&nbsp; {sel_label}</div>',
                unsafe_allow_html=True)
    st.markdown("---")

    # ── KPI Cards (5 cards: total, on, off, cost, kWh/Ton) ────────────────────
    agg       = week_agg(selected_yw)
    on_kwh    = agg["on_peak"]
    off_kwh   = agg["off_peak"]
    total_kwh = on_kwh + off_kwh
    cost      = on_kwh * (ON_PEAK_RATE + FT_ADJ) + off_kwh * (OFF_PEAK_RATE + FT_ADJ)
    on_pct    = on_kwh  / total_kwh * 100 if total_kwh else 0
    off_pct   = off_kwh / total_kwh * 100 if total_kwh else 0

    w_ton         = week_ton(selected_yw)
    kwh_per_ton   = total_kwh / w_ton if w_ton > 0 else None

    prev_w_ton      = week_ton(prev_yw) if prev_yw else 0
    prev_kwh_per_ton = (week_agg(prev_yw)["on_peak"] + week_agg(prev_yw)["off_peak"]) / prev_w_ton \
                       if (prev_yw and prev_w_ton > 0) else None

    if prev_yw:
        agg_p      = week_agg(prev_yw)
        prev_total = agg_p["on_peak"] + agg_p["off_peak"]
        chg_total  = (total_kwh - prev_total)        / prev_total        * 100 if prev_total        else 0
        chg_on     = (on_kwh    - agg_p["on_peak"])  / agg_p["on_peak"]  * 100 if agg_p["on_peak"]  else 0
        chg_off    = (off_kwh   - agg_p["off_peak"]) / agg_p["off_peak"] * 100 if agg_p["off_peak"] else 0
        chg_kpt    = (kwh_per_ton - prev_kwh_per_ton) / prev_kwh_per_ton * 100 \
                     if (kwh_per_ton and prev_kwh_per_ton) else 0
    else:
        chg_total = chg_on = chg_off = chg_kpt = 0

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">Total Energy</div>
            <div class="kpi-value">{total_kwh:,.0f} <span class="kpi-unit">kWh</span></div>
            <div class="kpi-sub">vs ก่อน &nbsp; {badge(chg_total)}</div>
        </div>""", unsafe_allow_html=True)
    with c2:
        st.markdown(f"""
        <div class="kpi-card" style="border-top-color:#e65100">
            <div class="kpi-label">On Peak</div>
            <div class="kpi-value kpi-on">{on_kwh:,.0f} <span class="kpi-unit">kWh</span></div>
            <div class="kpi-sub">({on_pct:.0f}%) &nbsp; {badge(chg_on)}</div>
        </div>""", unsafe_allow_html=True)
    with c3:
        st.markdown(f"""
        <div class="kpi-card" style="border-top-color:#2e7d32">
            <div class="kpi-label">Off Peak</div>
            <div class="kpi-value kpi-off">{off_kwh:,.0f} <span class="kpi-unit">kWh</span></div>
            <div class="kpi-sub">({off_pct:.0f}%) &nbsp; {badge(chg_off)}</div>
        </div>""", unsafe_allow_html=True)
    with c4:
        st.markdown(f"""
        <div class="kpi-card" style="border-top-color:#6a1b9a">
            <div class="kpi-label">Cost Estimate</div>
            <div class="kpi-value kpi-cost">{cost:,.0f} <span class="kpi-unit">฿</span></div>
            <div class="kpi-sub">อัตรา TOU + Ft</div>
        </div>""", unsafe_allow_html=True)
    with c5:
        if kwh_per_ton is not None:
            st.markdown(f"""
            <div class="kpi-card" style="border-top-color:#00838f">
                <div class="kpi-label">kWh / Ton</div>
                <div class="kpi-value kpi-ton">{kwh_per_ton:,.2f}</div>
                <div class="kpi-sub">Ton: {w_ton:,.1f} &nbsp; {badge(chg_kpt)}</div>
            </div>""", unsafe_allow_html=True)
        else:
            st.markdown(f"""
            <div class="kpi-card" style="border-top-color:#00838f">
                <div class="kpi-label">kWh / Ton</div>
                <div class="kpi-value kpi-ton" style="font-size:18px">ไม่มีข้อมูล Ton</div>
                <div class="kpi-sub">ตรวจสอบ Sheet "Product Ton"</div>
            </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Section 2: Weekly Comparison ──────────────────────────────────────────
    st.markdown('<div class="section-header">📊 Weekly Usage Comparison</div>', unsafe_allow_html=True)

    departments    = sorted(df["department"].unique().tolist())
    filter_options = ["🏭 Factory (ทั้งหมด)"] + departments

    col_f, col_g = st.columns([1, 3])
    with col_f:
        dept_sel = st.selectbox("🔍 เลือกแผนก", filter_options, index=0, key="weekly_dept_sel")

    agg_cur  = week_agg(selected_yw, dept_sel)
    agg_prev = week_agg(prev_yw,     dept_sel) if prev_yw else None

    label_cur  = f"สัปดาห์นี้\n({sel_label})"
    label_prev = f"สัปดาห์ก่อน\n({prev_label})" if prev_yw else "สัปดาห์ก่อน"

    on_vals  = [agg_prev["on_peak"]  if agg_prev is not None else 0, agg_cur["on_peak"]]
    off_vals = [agg_prev["off_peak"] if agg_prev is not None else 0, agg_cur["off_peak"]]
    x_labels = [label_prev, label_cur]
    tot_vals = [on_vals[i] + off_vals[i] for i in range(2)]
    on_pct_bars  = [on_vals[i]  / tot_vals[i] * 100 if tot_vals[i] else 0 for i in range(2)]
    off_pct_bars = [off_vals[i] / tot_vals[i] * 100 if tot_vals[i] else 0 for i in range(2)]

    # kWh/Ton overlay line
    ton_prev = week_ton(prev_yw) if prev_yw else 0
    ton_cur  = week_ton(selected_yw)
    kpt_prev = tot_vals[0] / ton_prev if ton_prev > 0 else None
    kpt_cur  = tot_vals[1] / ton_cur  if ton_cur  > 0 else None

    fig_w = make_subplots(specs=[[{"secondary_y": True}]])
    fig_w.add_trace(go.Bar(
        name="On Peak", x=x_labels, y=on_vals, marker_color="#e65100",
        text=[f"{v:,.0f} kWh ({on_pct_bars[i]:.0f}%)" for i, v in enumerate(on_vals)],
        textposition="inside", insidetextanchor="middle",
        textfont=dict(color="white", size=14),
    ), secondary_y=False)
    fig_w.add_trace(go.Bar(
        name="Off Peak", x=x_labels, y=off_vals, marker_color="#1565c0",
        text=[f"{v:,.0f} kWh ({off_pct_bars[i]:.0f}%)" for i, v in enumerate(off_vals)],
        textposition="inside", insidetextanchor="middle",
        textfont=dict(color="white", size=14),
    ), secondary_y=False)

    # Total annotation
    for xl, tv in zip(x_labels, tot_vals):
        fig_w.add_annotation(
            x=xl, y=tv, text=f"<b>{tv:,.0f} kWh</b>",
            showarrow=False, yshift=16, font=dict(size=14, color="#1a237e"),
        )

    # kWh/Ton line (only if data available)
    if kpt_prev is not None or kpt_cur is not None:
        kpt_vals = [kpt_prev if kpt_prev else 0, kpt_cur if kpt_cur else 0]
        fig_w.add_trace(go.Scatter(
            name="kWh/Ton", x=x_labels, y=kpt_vals,
            mode="lines+markers+text",
            line=dict(color="#02bdb3", width=3, dash="dot"),
            marker=dict(size=10, color="#02bdb3"),
            text=[f"{v:.2f}" if v else "" for v in kpt_vals],
            textposition="top center",
            textfont=dict(size=16, color="#02bdb3"),
        ), secondary_y=True)

    fig_w.update_layout(
        barmode="stack", height=420,
        title_text=f"On Peak vs Off Peak — {dept_sel}",
        title_font_size=16,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(t=60, b=20, l=20, r=60),
        plot_bgcolor="white", paper_bgcolor="white",
    )
    fig_w.update_yaxes(title_text="kWh", gridcolor="#f0f0f0", secondary_y=False)
    fig_w.update_yaxes(title_text="kWh/Ton", showgrid=False, secondary_y=True)

    with col_g:
        st.plotly_chart(fig_w, use_container_width=True)

    # Summary block
    cur_total_w  = agg_cur["on_peak"] + agg_cur["off_peak"]
    prev_total_w = (agg_prev["on_peak"] + agg_prev["off_peak"]) if agg_prev is not None else 0
    diff_w       = cur_total_w - prev_total_w
    pct_w        = diff_w / prev_total_w * 100 if prev_total_w else 0
    direction    = "เพิ่มขึ้น" if diff_w >= 0 else "ลดลง"
    dir_color    = "#e53935" if diff_w >= 0 else "#43a047"
    dir_icon     = "▲" if diff_w >= 0 else "▼"
    cur_on_pct   = agg_cur["on_peak"]  / cur_total_w * 100 if cur_total_w else 0
    cur_off_pct  = agg_cur["off_peak"] / cur_total_w * 100 if cur_total_w else 0
    kpt_txt      = f"&emsp;<b>kWh/Ton: {kwh_per_ton:.2f}</b> (Ton: {w_ton:,.1f})" if kwh_per_ton else ""

    st.markdown(f"""
<div style="background:#f8f9ff;border-left:4px solid #1565c0;border-radius:8px;
            padding:14px 22px;margin-top:4px;font-size:15px;line-height:2.0;color:#333;">
  📋 <b>สรุปสัปดาห์ {sel_label} — {dept_sel}</b><br>
  รวม <b>{cur_total_w:,.0f} kWh</b>
  &nbsp;|&nbsp; <span style="color:#e65100;font-weight:600;">On Peak</span> <b>{agg_cur["on_peak"]:,.0f} kWh</b> ({cur_on_pct:.1f}%)
  &emsp; <span style="color:#1565c0;font-weight:600;">Off Peak</span> <b>{agg_cur["off_peak"]:,.0f} kWh</b> ({cur_off_pct:.1f}%)
  {kpt_txt}<br>
  vs สัปดาห์ก่อน {prev_total_w:,.0f} kWh &nbsp;→&nbsp;
  <span style="color:{dir_color};font-weight:700;">{dir_icon} {direction} {abs(diff_w):,.0f} kWh ({abs(pct_w):.1f}%)</span>
</div>""", unsafe_allow_html=True)

    # ── Section 3: Department Breakdown ───────────────────────────────────────
    st.markdown('<div class="section-header">🏭 Department Usage Breakdown</div>', unsafe_allow_html=True)

    dept_cur  = dept_week_agg(selected_yw).set_index("department")
    dept_prev = dept_week_agg(prev_yw).set_index("department") if prev_yw else None
    dept_cur["total_combined"] = dept_cur["on_peak"] + dept_cur["off_peak"]
    all_depts = dept_cur.sort_values("total_combined", ascending=False).index.tolist()

    fig_d = make_subplots(
        rows=len(all_depts), cols=1,
        shared_xaxes=True, subplot_titles=all_depts,
        vertical_spacing=0.03,
    )
    for i, dep in enumerate(all_depts, start=1):
        cur_on  = dept_cur.loc[dep,  "on_peak"]  if dep in dept_cur.index  else 0
        cur_off = dept_cur.loc[dep,  "off_peak"] if dep in dept_cur.index  else 0
        prv_on  = dept_prev.loc[dep, "on_peak"]  if (dept_prev is not None and dep in dept_prev.index) else 0
        prv_off = dept_prev.loc[dep, "off_peak"] if (dept_prev is not None and dep in dept_prev.index) else 0
        total_cur = cur_on + cur_off
        total_prv = prv_on + prv_off
        pct       = ((total_cur - total_prv) / total_prv * 100) if total_prv else 0
        arrow_txt = f"▲{pct:.1f}%" if pct >= 0 else f"▼{abs(pct):.1f}%"
        clr       = "#e53935" if pct >= 0 else "#43a047"
        show_leg  = (i == 1)

        fig_d.add_trace(go.Bar(name="On Peak (ก่อน)",  x=[prv_on],  y=["ก่อน"], orientation="h",
            marker_color="#ffcc80", legendgroup="p_on",  showlegend=show_leg,
            hovertemplate=f"{dep} On Peak (ก่อน): %{{x:,.0f}} kWh<extra></extra>"), row=i, col=1)
        fig_d.add_trace(go.Bar(name="Off Peak (ก่อน)", x=[prv_off], y=["ก่อน"], orientation="h",
            marker_color="#90caf9", legendgroup="p_off", showlegend=show_leg,
            hovertemplate=f"{dep} Off Peak (ก่อน): %{{x:,.0f}} kWh<extra></extra>"), row=i, col=1)
        fig_d.add_trace(go.Bar(name="On Peak (นี้)",   x=[cur_on],  y=["นี้"],  orientation="h",
            marker_color="#e65100", legendgroup="c_on",  showlegend=show_leg,
            hovertemplate=f"{dep} On Peak (นี้): %{{x:,.0f}} kWh<extra></extra>"),  row=i, col=1)
        fig_d.add_trace(go.Bar(name="Off Peak (นี้)",  x=[cur_off], y=["นี้"],  orientation="h",
            marker_color="#1565c0", legendgroup="c_off", showlegend=show_leg,
            hovertemplate=f"{dep} Off Peak (นี้): %{{x:,.0f}} kWh<extra></extra>"), row=i, col=1)

        max_val = max(cur_on + cur_off, prv_on + prv_off, 1)
        fig_d.add_annotation(x=max_val * 1.02, y=0.5, yref=f"y{i}", xref=f"x{i}",
            text=f"<b>{arrow_txt}</b>", showarrow=False,
            font=dict(color=clr, size=13), xanchor="left")

    fig_d.update_layout(
        barmode="stack", height=max(130 * len(all_depts), 500),
        legend=dict(orientation="h", yanchor="bottom", y=1.005, xanchor="right", x=1),
        margin=dict(l=80, r=120, t=60, b=20),
        plot_bgcolor="white", paper_bgcolor="white",
    )
    for i in range(1, len(all_depts) + 1):
        fig_d.update_xaxes(gridcolor="#f0f0f0", row=i, col=1)
    st.plotly_chart(fig_d, use_container_width=True)

    st.markdown("---")
    last_updated = datetime.now().strftime("%d/%m/%Y %H:%M")
    num_days_sel = df[df["year_week"] == selected_yw]["date"].nunique()
    st.caption(f"📅 {selected_yw} ({sel_label}) | {num_days_sel} วัน | ก่อน: {prev_yw or 'N/A'} | 🕐 {last_updated}")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — MONTHLY
# ══════════════════════════════════════════════════════════════════════════════
with tab_monthly:

    def month_label(ym):
        y, m = ym.split("-")
        return f"{MONTH_TH[int(m)]} {y}"

    def month_agg(ym, dept_filter=None):
        sub = df[df["ym"] == ym]
        if dept_filter and dept_filter != "🏭 Factory (ทั้งหมด)":
            sub = sub[sub["department"] == dept_filter]
        return sub[["on_peak", "off_peak"]].sum()

    def month_dept_agg(ym):
        sub = df[df["ym"] == ym]
        return sub.groupby("department")[["on_peak", "off_peak"]].sum().reset_index()

    def daily_agg(ym, dept_filter=None):
        sub = df[df["ym"] == ym]
        if dept_filter and dept_filter != "🏭 Factory (ทั้งหมด)":
            sub = sub[sub["department"] == dept_filter]
        return sub.groupby("date")[["on_peak", "off_peak"]].sum().reset_index().sort_values("date")

    def month_ton(ym):
        if ton.empty:
            return 0.0
        return float(ton[ton["ym"] == ym]["ton"].sum())

    st.markdown('<div class="main-title">📅 Monthly Electricity Dashboard</div>', unsafe_allow_html=True)

    all_ym = sorted(df["ym"].unique().tolist())
    if not all_ym:
        st.error("ไม่พบข้อมูล")
        st.stop()

    ym_display = {ym: month_label(ym) for ym in all_ym}
    col_s1, col_s2, col_s3 = st.columns([1, 2, 1])
    with col_s2:
        selected_ym = st.selectbox(
            "📅 เลือกเดือน", options=all_ym,
            format_func=lambda x: ym_display[x],
            index=len(all_ym) - 1, key="monthly_selector",
        )

    sel_idx      = all_ym.index(selected_ym)
    prev_ym      = all_ym[sel_idx - 1] if sel_idx > 0 else None
    sel_label    = month_label(selected_ym)
    prev_label_m = month_label(prev_ym) if prev_ym else "N/A"

    st.markdown(f'<div class="month-subtitle">📆 เดือน {sel_label}</div>', unsafe_allow_html=True)
    st.markdown("---")

    m_departments    = sorted(df["department"].unique().tolist())
    m_filter_options = ["🏭 Factory (ทั้งหมด)"] + m_departments
    col_mf1, col_mf2, col_mf3 = st.columns([1, 2, 1])
    with col_mf2:
        m_dept_sel = st.selectbox("🔍 เลือกแผนก", m_filter_options, index=0, key="monthly_dept_sel")

    # ── KPI Cards (5 cards) ───────────────────────────────────────────────────
    agg     = month_agg(selected_ym, m_dept_sel)
    on_kwh  = agg["on_peak"]
    off_kwh = agg["off_peak"]
    total   = on_kwh + off_kwh
    cost    = on_kwh * (ON_PEAK_RATE + FT_ADJ) + off_kwh * (OFF_PEAK_RATE + FT_ADJ)
    on_pct  = on_kwh  / total * 100 if total else 0
    off_pct = off_kwh / total * 100 if total else 0

    m_ton       = month_ton(selected_ym)
    m_kpt       = total / m_ton if m_ton > 0 else None
    prev_m_ton  = month_ton(prev_ym) if prev_ym else 0
    prev_m_agg  = month_agg(prev_ym, m_dept_sel) if prev_ym else None
    prev_m_tot  = (prev_m_agg["on_peak"] + prev_m_agg["off_peak"]) if prev_m_agg is not None else 0
    prev_m_kpt  = prev_m_tot / prev_m_ton if prev_m_ton > 0 else None

    if prev_ym and prev_m_agg is not None:
        chg_tot  = (total   - prev_m_tot)              / prev_m_tot              * 100 if prev_m_tot              else 0
        chg_on   = (on_kwh  - prev_m_agg["on_peak"])   / prev_m_agg["on_peak"]  * 100 if prev_m_agg["on_peak"]  else 0
        chg_off  = (off_kwh - prev_m_agg["off_peak"])  / prev_m_agg["off_peak"] * 100 if prev_m_agg["off_peak"] else 0
        chg_kpt  = (m_kpt   - prev_m_kpt) / prev_m_kpt * 100 if (m_kpt and prev_m_kpt) else 0
    else:
        chg_tot = chg_on = chg_off = chg_kpt = 0

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">Total Energy</div>
            <div class="kpi-value">{total:,.0f} <span class="kpi-unit">kWh</span></div>
            <div class="kpi-sub">vs เดือนก่อน &nbsp; {badge(chg_tot)}</div>
        </div>""", unsafe_allow_html=True)
    with c2:
        st.markdown(f"""
        <div class="kpi-card" style="border-top-color:#e65100">
            <div class="kpi-label">On Peak</div>
            <div class="kpi-value kpi-on">{on_kwh:,.0f} <span class="kpi-unit">kWh</span></div>
            <div class="kpi-sub">({on_pct:.0f}%) &nbsp; {badge(chg_on)}</div>
        </div>""", unsafe_allow_html=True)
    with c3:
        st.markdown(f"""
        <div class="kpi-card" style="border-top-color:#2e7d32">
            <div class="kpi-label">Off Peak</div>
            <div class="kpi-value kpi-off">{off_kwh:,.0f} <span class="kpi-unit">kWh</span></div>
            <div class="kpi-sub">({off_pct:.0f}%) &nbsp; {badge(chg_off)}</div>
        </div>""", unsafe_allow_html=True)
    with c4:
        st.markdown(f"""
        <div class="kpi-card" style="border-top-color:#6a1b9a">
            <div class="kpi-label">Cost Estimate</div>
            <div class="kpi-value kpi-cost">{cost:,.0f} <span class="kpi-unit">฿</span></div>
            <div class="kpi-sub">อัตรา TOU + Ft</div>
        </div>""", unsafe_allow_html=True)
    with c5:
        if m_kpt is not None:
            st.markdown(f"""
            <div class="kpi-card" style="border-top-color:#00838f">
                <div class="kpi-label">kWh / Ton</div>
                <div class="kpi-value kpi-ton">{m_kpt:,.2f}</div>
                <div class="kpi-sub">Ton: {m_ton:,.1f} &nbsp; {badge(chg_kpt)}</div>
            </div>""", unsafe_allow_html=True)
        else:
            st.markdown(f"""
            <div class="kpi-card" style="border-top-color:#00838f">
                <div class="kpi-label">kWh / Ton</div>
                <div class="kpi-value kpi-ton" style="font-size:18px">ไม่มีข้อมูล Ton</div>
                <div class="kpi-sub">ตรวจสอบ Sheet "Product Ton"</div>
            </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Daily Trend + kWh/Ton ─────────────────────────────────────────────────
    st.markdown('<div class="section-header">📈 Daily Usage Trend</div>', unsafe_allow_html=True)

    daily = daily_agg(selected_ym, m_dept_sel)
    daily["total"] = daily["on_peak"] + daily["off_peak"]

    # Merge daily ton
    if not ton.empty:
        daily_ton = ton[ton["ym"] == selected_ym][["date", "ton"]].copy()
        daily_ton["date"] = pd.to_datetime(daily_ton["date"])
        daily = daily.merge(daily_ton, on="date", how="left")
        daily["kWh_per_ton"] = daily.apply(
            lambda r: r["total"] / r["ton"] if pd.notna(r.get("ton")) and r["ton"] > 0 else None, axis=1
        )
    else:
        daily["ton"] = None
        daily["kWh_per_ton"] = None

    fig_trend = make_subplots(specs=[[{"secondary_y": True}]])
    fig_trend.add_trace(go.Bar(
        x=daily["date"], y=daily["off_peak"],
        name="Off Peak", marker_color="#1565c0", opacity=0.85,
    ), secondary_y=False)
    fig_trend.add_trace(go.Bar(
        x=daily["date"], y=daily["on_peak"],
        name="On Peak", marker_color="#e65100", opacity=0.85,
    ), secondary_y=False)
    fig_trend.add_trace(go.Scatter(
        x=daily["date"], y=daily["total"],
        name="Total kWh", mode="lines+markers",
        line=dict(color="#1a237e", width=2),
        marker=dict(size=5),
    ), secondary_y=False)

    if "kWh_per_ton" in daily.columns and daily["kWh_per_ton"].notna().any():
        fig_trend.add_trace(go.Scatter(
            x=daily["date"], y=daily["kWh_per_ton"],
            name="kWh/Ton", mode="lines+markers",
            line=dict(color="#00838f", width=2.5, dash="dot"),
            marker=dict(size=6, color="#00838f"),
        ), secondary_y=True)

    fig_trend.update_layout(
        barmode="stack", height=400,
        title_text=f"การใช้ไฟฟ้ารายวัน — {sel_label} | {m_dept_sel}",
        title_font_size=16,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(t=60, b=20, l=20, r=60),
        plot_bgcolor="white", paper_bgcolor="white",
        hovermode="x unified",
    )
    fig_trend.update_xaxes(dtick="D1", tickformat="%d %b", tickangle=-45, gridcolor="#f0f0f0")
    fig_trend.update_yaxes(title_text="kWh", gridcolor="#f0f0f0", secondary_y=False)
    fig_trend.update_yaxes(title_text="kWh/Ton", showgrid=False, secondary_y=True)
    st.plotly_chart(fig_trend, use_container_width=True)

    # ── Month-over-Month + kWh/Ton ────────────────────────────────────────────
    st.markdown('<div class="section-header">📊 Month-over-Month Comparison</div>', unsafe_allow_html=True)

    mom_rows = []
    for ym in all_ym:
        a   = month_agg(ym, m_dept_sel)
        t   = month_ton(ym)
        tot_v = a["on_peak"] + a["off_peak"]
        mom_rows.append({
            "ym": ym, "label": month_label(ym),
            "on_peak": a["on_peak"], "off_peak": a["off_peak"],
            "total": tot_v,
            "ton": t,
            "kWh_per_ton": tot_v / t if t > 0 else None,
        })
    mom_df = pd.DataFrame(mom_rows)

    fig_mom = make_subplots(specs=[[{"secondary_y": True}]])
    fig_mom.add_trace(go.Bar(
        name="On Peak", x=mom_df["label"], y=mom_df["on_peak"],
        marker_color="#e65100", opacity=0.85,
        text=mom_df["on_peak"].apply(lambda v: f"{v:,.0f}"),
        textposition="inside", insidetextanchor="middle",
        textfont=dict(color="white", size=13),
    ), secondary_y=False)
    fig_mom.add_trace(go.Bar(
        name="Off Peak", x=mom_df["label"], y=mom_df["off_peak"],
        marker_color="#1565c0", opacity=0.85,
        text=mom_df["off_peak"].apply(lambda v: f"{v:,.0f}"),
        textposition="inside", insidetextanchor="middle",
        textfont=dict(color="white", size=13),
    ), secondary_y=False)

    for _, row in mom_df.iterrows():
        fig_mom.add_annotation(
            x=row["label"], y=row["total"],
            text=f"<b>{row['total']:,.0f}</b>",
            showarrow=False, yshift=10,
            font=dict(size=13, color="#1a237e"),
        )

    if mom_df["kWh_per_ton"].notna().any():
        fig_mom.add_trace(go.Scatter(
            name="kWh/Ton", x=mom_df["label"], y=mom_df["kWh_per_ton"],
            mode="lines+markers+text",
            line=dict(color="#00838f", width=3, dash="dot"),
            marker=dict(size=8, color="#00838f"),
            text=mom_df["kWh_per_ton"].apply(lambda v: f"{v:.2f}" if pd.notna(v) else ""),
            textposition="top center",
            textfont=dict(size=13, color="#00838f"),
        ), secondary_y=True)

    fig_mom.update_layout(
        barmode="stack", height=420,
        title_text=f"เปรียบเทียบรายเดือน — {m_dept_sel}",
        title_font_size=16,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(t=60, b=30, l=20, r=60),
        plot_bgcolor="white", paper_bgcolor="white",
    )
    fig_mom.update_yaxes(title_text="kWh", gridcolor="#f0f0f0", secondary_y=False)
    fig_mom.update_yaxes(title_text="kWh/Ton", showgrid=False, secondary_y=True)
    st.plotly_chart(fig_mom, use_container_width=True)

    # Summary
    if prev_ym:
        diff_v    = total - prev_m_tot
        pct_v     = diff_v / prev_m_tot * 100 if prev_m_tot else 0
        direction = "เพิ่มขึ้น" if diff_v >= 0 else "ลดลง"
        dir_color = "#e53935" if diff_v >= 0 else "#43a047"
        dir_icon  = "▲" if diff_v >= 0 else "▼"
        kpt_sum   = f"&emsp;<b>kWh/Ton: {m_kpt:.2f}</b> (Ton: {m_ton:,.1f})" if m_kpt else ""

        st.markdown(f"""
<div style="background:#f8f9ff;border-left:4px solid #1565c0;border-radius:8px;
            padding:14px 22px;margin-top:4px;font-size:15px;line-height:2.0;color:#333;">
  📋 <b>สรุปเดือน {sel_label}</b><br>
  รวม <b>{total:,.0f} kWh</b>
  &nbsp;|&nbsp; <span style="color:#e65100;font-weight:600;">On Peak</span> <b>{on_kwh:,.0f} kWh</b> ({on_pct:.1f}%)
  &emsp; <span style="color:#1565c0;font-weight:600;">Off Peak</span> <b>{off_kwh:,.0f} kWh</b> ({off_pct:.1f}%)
  {kpt_sum}<br>
  vs {prev_label_m}: {prev_m_tot:,.0f} kWh &nbsp;→&nbsp;
  <span style="color:{dir_color};font-weight:700;">{dir_icon} {direction} {abs(diff_v):,.0f} kWh ({abs(pct_v):.1f}%)</span>
</div>""", unsafe_allow_html=True)

    st.markdown("---")
    num_days     = df[df["ym"] == selected_ym]["date"].nunique()
    last_updated = datetime.now().strftime("%d/%m/%Y %H:%M")
    st.caption(f"📅 {sel_label} | {num_days} วัน | ก่อน: {prev_label_m} | 🕐 {last_updated}")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — WEEKLY kWh/Ton
# ══════════════════════════════════════════════════════════════════════════════
with tab_wkton:

    def wkt_week_label(year_week):
        sub = df[df["year_week"] == year_week]["date"]
        if sub.empty:
            return year_week
        start, end = sub.min(), sub.max()
        if start.month == end.month:
            return f"{start.day} – {fmt_day(end)}"
        return f"{fmt_day_short(start)} – {fmt_day(end)}"

    def wkt_week_ton(year_week):
        if ton.empty:
            return 0.0
        return float(ton[ton["year_week"] == year_week]["ton"].sum())

    def wkt_week_kwh(year_week, dept_filter=None):
        sub = df[df["year_week"] == year_week]
        if dept_filter and dept_filter != "🏭 Factory (ทั้งหมด)":
            sub = sub[sub["department"] == dept_filter]
        return float((sub["on_peak"] + sub["off_peak"]).sum())

    st.markdown('<div class="main-title">📋 Weekly kWh/Ton Overview</div>', unsafe_allow_html=True)

    all_weeks_wkt      = sorted(df["year_week"].unique().tolist())
    complete_weeks_wkt = sorted(df.groupby("year_week")["date"].nunique()[lambda x: x == 7].index.tolist())

    if not all_weeks_wkt:
        st.error("ไม่พบข้อมูล")
        st.stop()

    # Week selector
    wkt_display_map = {}
    for yw in all_weeks_wkt:
        lbl    = wkt_week_label(yw)
        suffix = "" if yw in complete_weeks_wkt else " ⚠️ ไม่ครบ 7 วัน"
        wkt_display_map[yw] = f"{yw}  ·  {lbl}{suffix}"

    wkt_display_opts = list(wkt_display_map.values())
    wkt_yw_keys      = list(wkt_display_map.keys())
    wkt_default      = complete_weeks_wkt[-1] if complete_weeks_wkt else all_weeks_wkt[-1]
    wkt_default_idx  = wkt_yw_keys.index(wkt_default)

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        wkt_sel_display = st.selectbox(
            "📅 เลือกสัปดาห์", options=wkt_display_opts,
            index=wkt_default_idx, key="wkton_selector",
        )

    wkt_sel_yw   = wkt_yw_keys[wkt_display_opts.index(wkt_sel_display)]
    wkt_sel_idx  = wkt_yw_keys.index(wkt_sel_yw)
    wkt_prev_yw  = wkt_yw_keys[wkt_sel_idx - 1] if wkt_sel_idx > 0 else None
    wkt_sel_lbl  = wkt_week_label(wkt_sel_yw)
    wkt_prev_lbl = wkt_week_label(wkt_prev_yw) if wkt_prev_yw else "N/A"

    st.markdown(f'<div class="week-subtitle">📆 สัปดาห์ {wkt_sel_yw} &nbsp;|&nbsp; {wkt_sel_lbl}</div>',
                unsafe_allow_html=True)
    st.markdown("---")

    # ── Department Filter ─────────────────────────────────────────────────────
    wkt_departments    = sorted(df["department"].unique().tolist())
    wkt_filter_options = ["🏭 Factory (ทั้งหมด)"] + wkt_departments
    col_wf1, col_wf2, col_wf3 = st.columns([1, 2, 1])
    with col_wf2:
        wkt_dept_sel = st.selectbox("🔍 เลือกแผนก", wkt_filter_options, index=0, key="wkton_dept_sel")

    # ── KPI Cards ─────────────────────────────────────────────────────────────
    cur_kwh  = wkt_week_kwh(wkt_sel_yw, wkt_dept_sel)
    cur_ton  = wkt_week_ton(wkt_sel_yw)
    cur_kpt  = cur_kwh / cur_ton if cur_ton > 0 else None

    prev_kwh = wkt_week_kwh(wkt_prev_yw, wkt_dept_sel) if wkt_prev_yw else 0
    prev_ton = wkt_week_ton(wkt_prev_yw) if wkt_prev_yw else 0
    prev_kpt = prev_kwh / prev_ton if prev_ton > 0 else None

    chg_kwh  = (cur_kwh - prev_kwh) / prev_kwh * 100 if prev_kwh else 0
    chg_ton  = (cur_ton - prev_ton) / prev_ton * 100 if prev_ton else 0
    chg_kpt  = (cur_kpt - prev_kpt) / prev_kpt * 100 if (cur_kpt and prev_kpt) else 0

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">Total Energy This Week</div>
            <div class="kpi-value">{cur_kwh:,.0f} <span class="kpi-unit">kWh</span></div>
            <div class="kpi-sub">vs ก่อน &nbsp; {badge(chg_kwh)}</div>
        </div>""", unsafe_allow_html=True)
    with c2:
        st.markdown(f"""
        <div class="kpi-card" style="border-top-color:#f57f17">
            <div class="kpi-label">Production This Week</div>
            <div class="kpi-value" style="color:#f57f17">{cur_ton:,.1f} <span class="kpi-unit">Ton</span></div>
            <div class="kpi-sub">vs ก่อน &nbsp; {badge(chg_ton)}</div>
        </div>""", unsafe_allow_html=True)
    with c3:
        if cur_kpt:
            st.markdown(f"""
            <div class="kpi-card" style="border-top-color:#00838f">
                <div class="kpi-label">Energy Intensity</div>
                <div class="kpi-value kpi-ton">{cur_kpt:,.2f} <span class="kpi-unit">kWh/Ton</span></div>
                <div class="kpi-sub">vs ก่อน &nbsp; {badge(chg_kpt)}</div>
            </div>""", unsafe_allow_html=True)
        else:
            st.markdown(f"""
            <div class="kpi-card" style="border-top-color:#00838f">
                <div class="kpi-label">Energy Intensity</div>
                <div class="kpi-value kpi-ton" style="font-size:18px">ไม่มีข้อมูล Ton</div>
                <div class="kpi-sub">ตรวจสอบ Sheet "Product Ton"</div>
            </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Weekly kWh/Ton Comparison Bar Chart ───────────────────────────────────
    st.markdown('<div class="section-header">📊 Weekly kWh/Ton Comparison</div>', unsafe_allow_html=True)

    col_f2, col_g2 = st.columns([1, 3])

    x_labels_wkt = [f"สัปดาห์ก่อน\n({wkt_prev_lbl})", f"สัปดาห์นี้\n({wkt_sel_lbl})"]
    kwh_vals     = [prev_kwh, cur_kwh]
    ton_vals_bar = [prev_ton, cur_ton]
    kpt_vals_bar = [prev_kpt if prev_kpt else 0, cur_kpt if cur_kpt else 0]

    fig_wkt = make_subplots(specs=[[{"secondary_y": True}]])
    fig_wkt.add_trace(go.Bar(
        name="kWh", x=x_labels_wkt, y=kwh_vals,
        marker_color="#1565c0", opacity=0.75,
        text=[f"{v:,.0f} kWh" for v in kwh_vals],
        textposition="inside", insidetextanchor="middle",
        textfont=dict(color="white", size=14),
    ), secondary_y=False)
    fig_wkt.add_trace(go.Bar(
        name="Ton", x=x_labels_wkt, y=ton_vals_bar,
        marker_color="#f57f17", opacity=0.75,
        text=[f"{v:,.1f} Ton" for v in ton_vals_bar],
        textposition="inside", insidetextanchor="middle",
        textfont=dict(color="white", size=14),
    ), secondary_y=False)
    if any(kpt_vals_bar):
        fig_wkt.add_trace(go.Scatter(
            name="kWh/Ton", x=x_labels_wkt, y=kpt_vals_bar,
            mode="lines+markers+text",
            line=dict(color="#00838f", width=3, dash="dot"),
            marker=dict(size=12, color="#00838f"),
            text=[f"{v:.2f}" if v else "" for v in kpt_vals_bar],
            textposition="top center",
            textfont=dict(size=14, color="#00838f"),
        ), secondary_y=True)

    fig_wkt.update_layout(
        barmode="group", height=420,
        title_text=f"kWh vs Ton vs kWh/Ton — {wkt_sel_lbl} | {wkt_dept_sel}",
        title_font_size=16,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(t=60, b=20, l=20, r=60),
        plot_bgcolor="white", paper_bgcolor="white",
    )
    fig_wkt.update_yaxes(title_text="kWh / Ton (Production)", gridcolor="#f0f0f0", secondary_y=False)
    fig_wkt.update_yaxes(title_text="kWh/Ton (Intensity)", showgrid=False, secondary_y=True)

    with col_g2:
        st.plotly_chart(fig_wkt, use_container_width=True)

    # Summary
    if cur_kpt and prev_kpt:
        diff_kpt   = cur_kpt - prev_kpt
        dir_kpt    = "เพิ่มขึ้น" if diff_kpt >= 0 else "ลดลง"
        clr_kpt    = "#e53935" if diff_kpt >= 0 else "#43a047"
        icon_kpt   = "▲" if diff_kpt >= 0 else "▼"
        st.markdown(f"""
<div style="background:#f8f9ff;border-left:4px solid #00838f;border-radius:8px;
            padding:14px 22px;margin-top:4px;font-size:15px;line-height:2.0;color:#333;">
  📋 <b>สรุป kWh/Ton สัปดาห์ {wkt_sel_lbl}</b><br>
  พลังงาน <b>{cur_kwh:,.0f} kWh</b> &nbsp;|&nbsp; การผลิต <b>{cur_ton:,.1f} Ton</b>
  &nbsp;|&nbsp; <span style="color:#00838f;font-weight:700;">Energy Intensity: {cur_kpt:.2f} kWh/Ton</span><br>
  vs สัปดาห์ก่อน {prev_kpt:.2f} kWh/Ton &nbsp;→&nbsp;
  <span style="color:{clr_kpt};font-weight:700;">{icon_kpt} {dir_kpt} {abs(diff_kpt):.2f} kWh/Ton ({abs(chg_kpt):.1f}%)</span>
</div>""", unsafe_allow_html=True)

    # ── All-weeks kWh/Ton Trend ────────────────────────────────────────────────
    st.markdown('<div class="section-header">📈 Weekly kWh/Ton Trend (ทุกสัปดาห์)</div>', unsafe_allow_html=True)

    trend_rows = []
    for yw in all_weeks_wkt:
        kwh_v = wkt_week_kwh(yw, wkt_dept_sel)
        ton_v = wkt_week_ton(yw)
        trend_rows.append({
            "year_week": yw,
            "label":     wkt_week_label(yw),
            "kwh":       kwh_v,
            "ton":       ton_v,
            "kpt":       kwh_v / ton_v if ton_v > 0 else None,
        })
    trend_df = pd.DataFrame(trend_rows)

    fig_trend_wkt = make_subplots(specs=[[{"secondary_y": True}]])
    fig_trend_wkt.add_trace(go.Bar(
        name="kWh", x=trend_df["label"], y=trend_df["kwh"],
        marker_color="#1565c0", opacity=0.6,
    ), secondary_y=False)
    fig_trend_wkt.add_trace(go.Bar(
        name="Ton", x=trend_df["label"], y=trend_df["ton"],
        marker_color="#f57f17", opacity=0.6,
    ), secondary_y=False)
    if trend_df["kpt"].notna().any():
        fig_trend_wkt.add_trace(go.Scatter(
            name="kWh/Ton", x=trend_df["label"], y=trend_df["kpt"],
            mode="lines+markers+text",
            line=dict(color="#00838f", width=2.5),
            marker=dict(size=7, color="#00838f"),
            text=trend_df["kpt"].apply(lambda v: f"{v:.1f}" if pd.notna(v) else ""),
            textposition="top center",
            textfont=dict(size=11, color="#00838f"),
        ), secondary_y=True)

    fig_trend_wkt.update_layout(
        barmode="group", height=400,
        title_text=f"แนวโน้ม kWh/Ton รายสัปดาห์ — {wkt_dept_sel}",
        title_font_size=16,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(t=60, b=40, l=20, r=60),
        plot_bgcolor="white", paper_bgcolor="white",
    )
    fig_trend_wkt.update_xaxes(tickangle=-30, gridcolor="#f0f0f0")
    fig_trend_wkt.update_yaxes(title_text="kWh / Ton (Production)", gridcolor="#f0f0f0", secondary_y=False)
    fig_trend_wkt.update_yaxes(title_text="kWh/Ton (Intensity)", showgrid=False, secondary_y=True)
    st.plotly_chart(fig_trend_wkt, use_container_width=True)

    st.markdown("---")
    last_updated = datetime.now().strftime("%d/%m/%Y %H:%M")
    st.caption(f"📅 {wkt_sel_yw} ({wkt_sel_lbl}) | ก่อน: {wkt_prev_yw or 'N/A'} | 🕐 {last_updated}")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — MONTHLY kWh/Ton
# ══════════════════════════════════════════════════════════════════════════════
with tab_monton:

    def mkt_month_label(ym):
        y, m = ym.split("-")
        return f"{MONTH_TH[int(m)]} {y}"

    def mkt_month_kwh(ym, dept_filter=None):
        sub = df[df["ym"] == ym]
        if dept_filter and dept_filter != "🏭 Factory (ทั้งหมด)":
            sub = sub[sub["department"] == dept_filter]
        return float((sub["on_peak"] + sub["off_peak"]).sum())

    def mkt_month_ton(ym):
        if ton.empty:
            return 0.0
        return float(ton[ton["ym"] == ym]["ton"].sum())

    def mkt_daily_kwh(ym, dept_filter=None):
        sub = df[df["ym"] == ym]
        if dept_filter and dept_filter != "🏭 Factory (ทั้งหมด)":
            sub = sub[sub["department"] == dept_filter]
        return sub.groupby("date")[["on_peak", "off_peak"]].sum().reset_index().sort_values("date")

    st.markdown('<div class="main-title">📅 Monthly kWh/Ton Dashboard</div>', unsafe_allow_html=True)

    all_ym_mkt = sorted(df["ym"].unique().tolist())
    if not all_ym_mkt:
        st.error("ไม่พบข้อมูล")
        st.stop()

    ym_display_mkt = {ym: mkt_month_label(ym) for ym in all_ym_mkt}

    col_s1, col_s2, col_s3 = st.columns([1, 2, 1])
    with col_s2:
        mkt_sel_ym = st.selectbox(
            "📅 เลือกเดือน", options=all_ym_mkt,
            format_func=lambda x: ym_display_mkt[x],
            index=len(all_ym_mkt) - 1, key="monton_selector",
        )

    mkt_sel_idx   = all_ym_mkt.index(mkt_sel_ym)
    mkt_prev_ym   = all_ym_mkt[mkt_sel_idx - 1] if mkt_sel_idx > 0 else None
    mkt_sel_lbl   = mkt_month_label(mkt_sel_ym)
    mkt_prev_lbl  = mkt_month_label(mkt_prev_ym) if mkt_prev_ym else "N/A"

    st.markdown(f'<div class="month-subtitle">📆 เดือน {mkt_sel_lbl}</div>', unsafe_allow_html=True)
    st.markdown("---")

    mkt_departments    = sorted(df["department"].unique().tolist())
    mkt_filter_options = ["🏭 Factory (ทั้งหมด)"] + mkt_departments
    col_mf1, col_mf2, col_mf3 = st.columns([1, 2, 1])
    with col_mf2:
        mkt_dept_sel = st.selectbox("🔍 เลือกแผนก", mkt_filter_options, index=0, key="monton_dept_sel")

    # ── KPI Cards ─────────────────────────────────────────────────────────────
    mkt_cur_kwh  = mkt_month_kwh(mkt_sel_ym, mkt_dept_sel)
    mkt_cur_ton  = mkt_month_ton(mkt_sel_ym)
    mkt_cur_kpt  = mkt_cur_kwh / mkt_cur_ton if mkt_cur_ton > 0 else None

    mkt_prev_kwh = mkt_month_kwh(mkt_prev_ym, mkt_dept_sel) if mkt_prev_ym else 0
    mkt_prev_ton = mkt_month_ton(mkt_prev_ym) if mkt_prev_ym else 0
    mkt_prev_kpt = mkt_prev_kwh / mkt_prev_ton if mkt_prev_ton > 0 else None

    mkt_chg_kwh  = (mkt_cur_kwh - mkt_prev_kwh) / mkt_prev_kwh * 100 if mkt_prev_kwh else 0
    mkt_chg_ton  = (mkt_cur_ton - mkt_prev_ton) / mkt_prev_ton * 100 if mkt_prev_ton else 0
    mkt_chg_kpt  = (mkt_cur_kpt - mkt_prev_kpt) / mkt_prev_kpt * 100 if (mkt_cur_kpt and mkt_prev_kpt) else 0

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">Total Energy This Month</div>
            <div class="kpi-value">{mkt_cur_kwh:,.0f} <span class="kpi-unit">kWh</span></div>
            <div class="kpi-sub">vs เดือนก่อน &nbsp; {badge(mkt_chg_kwh)}</div>
        </div>""", unsafe_allow_html=True)
    with c2:
        st.markdown(f"""
        <div class="kpi-card" style="border-top-color:#f57f17">
            <div class="kpi-label">Production This Month</div>
            <div class="kpi-value" style="color:#f57f17">{mkt_cur_ton:,.1f} <span class="kpi-unit">Ton</span></div>
            <div class="kpi-sub">vs เดือนก่อน &nbsp; {badge(mkt_chg_ton)}</div>
        </div>""", unsafe_allow_html=True)
    with c3:
        if mkt_cur_kpt:
            st.markdown(f"""
            <div class="kpi-card" style="border-top-color:#00838f">
                <div class="kpi-label">Energy Intensity</div>
                <div class="kpi-value kpi-ton">{mkt_cur_kpt:,.2f} <span class="kpi-unit">kWh/Ton</span></div>
                <div class="kpi-sub">vs เดือนก่อน &nbsp; {badge(mkt_chg_kpt)}</div>
            </div>""", unsafe_allow_html=True)
        else:
            st.markdown(f"""
            <div class="kpi-card" style="border-top-color:#00838f">
                <div class="kpi-label">Energy Intensity</div>
                <div class="kpi-value kpi-ton" style="font-size:18px">ไม่มีข้อมูล Ton</div>
                <div class="kpi-sub">ตรวจสอบ Sheet "Product Ton"</div>
            </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Daily kWh/Ton Trend ────────────────────────────────────────────────────
    st.markdown('<div class="section-header">📈 Daily kWh/Ton Trend</div>', unsafe_allow_html=True)

    daily_mkt = mkt_daily_kwh(mkt_sel_ym, mkt_dept_sel)
    daily_mkt["total"] = daily_mkt["on_peak"] + daily_mkt["off_peak"]

    if not ton.empty:
        daily_ton_mkt = ton[ton["ym"] == mkt_sel_ym][["date", "ton"]].copy()
        daily_ton_mkt["date"] = pd.to_datetime(daily_ton_mkt["date"])
        daily_mkt = daily_mkt.merge(daily_ton_mkt, on="date", how="left")
        daily_mkt["kpt"] = daily_mkt.apply(
            lambda r: r["total"] / r["ton"] if pd.notna(r.get("ton")) and r["ton"] > 0 else None, axis=1
        )
    else:
        daily_mkt["ton"] = None
        daily_mkt["kpt"] = None

    fig_daily_mkt = make_subplots(specs=[[{"secondary_y": True}]])
    fig_daily_mkt.add_trace(go.Bar(
        x=daily_mkt["date"], y=daily_mkt["total"],
        name="kWh", marker_color="#1565c0", opacity=0.7,
    ), secondary_y=False)
    if "ton" in daily_mkt.columns and daily_mkt["ton"].notna().any():
        fig_daily_mkt.add_trace(go.Bar(
            x=daily_mkt["date"], y=daily_mkt["ton"],
            name="Ton", marker_color="#f57f17", opacity=0.7,
        ), secondary_y=False)
    if "kpt" in daily_mkt.columns and daily_mkt["kpt"].notna().any():
        fig_daily_mkt.add_trace(go.Scatter(
            x=daily_mkt["date"], y=daily_mkt["kpt"],
            name="kWh/Ton", mode="lines+markers",
            line=dict(color="#00838f", width=2.5),
            marker=dict(size=6, color="#00838f"),
        ), secondary_y=True)

    fig_daily_mkt.update_layout(
        barmode="group", height=400,
        title_text=f"kWh / Ton รายวัน — {mkt_sel_lbl} | {mkt_dept_sel}",
        title_font_size=16,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(t=60, b=20, l=20, r=60),
        plot_bgcolor="white", paper_bgcolor="white",
        hovermode="x unified",
    )
    fig_daily_mkt.update_xaxes(dtick="D1", tickformat="%d %b", tickangle=-45, gridcolor="#f0f0f0")
    fig_daily_mkt.update_yaxes(title_text="kWh / Ton (Production)", gridcolor="#f0f0f0", secondary_y=False)
    fig_daily_mkt.update_yaxes(title_text="kWh/Ton (Intensity)", showgrid=False, secondary_y=True)
    st.plotly_chart(fig_daily_mkt, use_container_width=True)

    # ── Month-over-Month kWh/Ton ───────────────────────────────────────────────
    st.markdown('<div class="section-header">📊 Month-over-Month kWh/Ton</div>', unsafe_allow_html=True)

    mkt_mom_rows = []
    for ym in all_ym_mkt:
        kwh_v = mkt_month_kwh(ym, mkt_dept_sel)
        ton_v = mkt_month_ton(ym)
        mkt_mom_rows.append({
            "ym":    ym,
            "label": mkt_month_label(ym),
            "kwh":   kwh_v,
            "ton":   ton_v,
            "kpt":   kwh_v / ton_v if ton_v > 0 else None,
        })
    mkt_mom_df = pd.DataFrame(mkt_mom_rows)

    fig_mom_mkt = make_subplots(specs=[[{"secondary_y": True}]])
    fig_mom_mkt.add_trace(go.Bar(
        name="kWh", x=mkt_mom_df["label"], y=mkt_mom_df["kwh"],
        marker_color="#1565c0", opacity=0.75,
        text=mkt_mom_df["kwh"].apply(lambda v: f"{v:,.0f}"),
        textposition="inside", insidetextanchor="middle",
        textfont=dict(color="white", size=13),
    ), secondary_y=False)
    fig_mom_mkt.add_trace(go.Bar(
        name="Ton", x=mkt_mom_df["label"], y=mkt_mom_df["ton"],
        marker_color="#f57f17", opacity=0.75,
        text=mkt_mom_df["ton"].apply(lambda v: f"{v:,.1f}"),
        textposition="inside", insidetextanchor="middle",
        textfont=dict(color="white", size=13),
    ), secondary_y=False)
    if mkt_mom_df["kpt"].notna().any():
        fig_mom_mkt.add_trace(go.Scatter(
            name="kWh/Ton", x=mkt_mom_df["label"], y=mkt_mom_df["kpt"],
            mode="lines+markers+text",
            line=dict(color="#00838f", width=3, dash="dot"),
            marker=dict(size=9, color="#00838f"),
            text=mkt_mom_df["kpt"].apply(lambda v: f"{v:.2f}" if pd.notna(v) else ""),
            textposition="top center",
            textfont=dict(size=13, color="#00838f"),
        ), secondary_y=True)

    fig_mom_mkt.update_layout(
        barmode="group", height=420,
        title_text=f"เปรียบเทียบ kWh/Ton รายเดือน — {mkt_dept_sel}",
        title_font_size=16,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(t=60, b=30, l=20, r=60),
        plot_bgcolor="white", paper_bgcolor="white",
    )
    fig_mom_mkt.update_yaxes(title_text="kWh / Ton (Production)", gridcolor="#f0f0f0", secondary_y=False)
    fig_mom_mkt.update_yaxes(title_text="kWh/Ton (Intensity)", showgrid=False, secondary_y=True)
    st.plotly_chart(fig_mom_mkt, use_container_width=True)

    # Summary
    if mkt_cur_kpt and mkt_prev_kpt:
        diff_mkt   = mkt_cur_kpt - mkt_prev_kpt
        dir_mkt    = "เพิ่มขึ้น" if diff_mkt >= 0 else "ลดลง"
        clr_mkt    = "#e53935" if diff_mkt >= 0 else "#43a047"
        icon_mkt   = "▲" if diff_mkt >= 0 else "▼"
        st.markdown(f"""
<div style="background:#f8f9ff;border-left:4px solid #00838f;border-radius:8px;
            padding:14px 22px;margin-top:4px;font-size:15px;line-height:2.0;color:#333;">
  📋 <b>สรุป kWh/Ton เดือน {mkt_sel_lbl}</b><br>
  พลังงาน <b>{mkt_cur_kwh:,.0f} kWh</b> &nbsp;|&nbsp; การผลิต <b>{mkt_cur_ton:,.1f} Ton</b>
  &nbsp;|&nbsp; <span style="color:#00838f;font-weight:700;">Energy Intensity: {mkt_cur_kpt:.2f} kWh/Ton</span><br>
  vs {mkt_prev_lbl}: {mkt_prev_kpt:.2f} kWh/Ton &nbsp;→&nbsp;
  <span style="color:{clr_mkt};font-weight:700;">{icon_mkt} {dir_mkt} {abs(diff_mkt):.2f} kWh/Ton ({abs(mkt_chg_kpt):.1f}%)</span>
</div>""", unsafe_allow_html=True)

    st.markdown("---")
    num_days_mkt = df[df["ym"] == mkt_sel_ym]["date"].nunique()
    last_updated = datetime.now().strftime("%d/%m/%Y %H:%M")
    st.caption(f"📅 {mkt_sel_lbl} | {num_days_mkt} วัน | ก่อน: {mkt_prev_lbl} | 🕐 {last_updated}")
