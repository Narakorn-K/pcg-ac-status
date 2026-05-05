import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import re
from datetime import datetime
from urllib.parse import quote

# ─── Page Config ──────────────────────────────────────────────────────────────
st.set_page_config(page_title="Electricity Dashboard", layout="wide", page_icon="⚡")

# ─── Google Sheet Config ──────────────────────────────────────────────────────
SHEET_ID   = "1Ym2yfzkLTyLTtJtLZSSgWoeew_IPWUaI_u6d45jKUnw"
SHEET_NAME = "Daily"
EXPORT_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx&sheet={quote(SHEET_NAME)}"

# ─── Thai electricity tariff (PEA TOU rates, Baht/kWh) ───────────────────────
ON_PEAK_RATE  = 4.1824
OFF_PEAK_RATE = 2.6369
FT_ADJ        = 0.3949

# ─── Thai day / month abbreviations ──────────────────────────────────────────
DAY_TH = {"อา": 6, "จ": 0, "อ": 1, "พ": 2, "พฤ": 3, "ศ": 4, "ส": 5}
MONTH_TH = {
    1: "ม.ค.", 2: "ก.พ.", 3: "มี.ค.", 4: "เม.ย.",
    5: "พ.ค.", 6: "มิ.ย.", 7: "ก.ค.", 8: "ส.ค.",
    9: "ก.ย.", 10: "ต.ค.", 11: "พ.ย.", 12: "ธ.ค.",
}

# ─── CSS ──────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-title {
        font-size:26px; font-weight:700; text-align:center;
        color:#1a237e; margin-bottom:4px;
    }
    .week-subtitle, .month-subtitle {
        font-size:15px; text-align:center; color:#546e7a;
        margin-bottom:20px; font-weight:500;
    }
    .kpi-card {
        background:#fff; border-radius:12px; padding:18px 22px;
        box-shadow:0 2px 10px rgba(0,0,0,0.08); text-align:center;
        border-top: 4px solid #1565c0;
    }
    .kpi-label {font-size:15px; color:#666; font-weight:600; margin-bottom:6px; text-transform:uppercase; letter-spacing:0.5px;}
    .kpi-value {font-size:32px; font-weight:800; color:#1a237e;}
    .kpi-unit  {font-size:17px; font-weight:400;}
    .kpi-sub   {font-size:15px; color:#999; margin-top:4px;}
    .kpi-on    {color:#e65100 !important;}
    .kpi-off   {color:#2e7d32 !important;}
    .kpi-cost  {color:#6a1b9a !important;}
    .up   {color:#e53935; font-weight:700; font-size:15px;}
    .down {color:#43a047; font-weight:700; font-size:15px;}
    .section-header {
        font-size:19px; font-weight:700; color:#1a237e;
        border-left:4px solid #1565c0; padding-left:10px;
        margin:28px 0 14px;
    }
    div[data-testid="stSelectbox"] label {font-weight:600; font-size:16px;}
    div[data-testid="stSelectbox"] div   {font-size:16px;}
    .summary-box {font-size:16px !important; line-height:2.2;}
    p, li, .stCaption {font-size:15px !important;}
</style>
""", unsafe_allow_html=True)

# ─── Helpers ──────────────────────────────────────────────────────────────────
def parse_date_col(raw_val, fallback_year=2026):
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
    cls   = "up"  if v >= 0 else "down"
    arrow = "▲"   if v >= 0 else "▼"
    return f'<span class="{cls}">{arrow} {abs(v):.1f}%</span>'


# ─── Data Loading ─────────────────────────────────────────────────────────────
@st.cache_data(ttl=300)
def load_data():
    try:
        raw = pd.read_excel(EXPORT_URL, sheet_name=SHEET_NAME, header=None)
    except Exception as e:
        st.error(f"❌ ไม่สามารถดึงข้อมูลจาก Google Sheet ได้: {e}")
        st.stop()

    # Find header row
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
        st.error("❌ ไม่พบคอลัมน์วันที่ กรุณาตรวจสอบโครงสร้าง Sheet")
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
        st.error("❌ ดึงข้อมูลได้แต่ไม่พบแถวข้อมูล กรุณาตรวจสอบโครงสร้าง Sheet")
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


# ─── Weekly helpers ───────────────────────────────────────────────────────────
def week_label(year_week, df):
    sub = df[df["year_week"] == year_week]["date"]
    if sub.empty:
        return year_week
    start, end = sub.min(), sub.max()
    if start.month == end.month:
        return f"{start.day} – {end.strftime('%-d %b %Y')}"
    return f"{start.strftime('%-d %b')} – {end.strftime('%-d %b %Y')}"


def get_all_weeks(df):
    return sorted(df["year_week"].unique().tolist())


def get_complete_weeks(df):
    wk_days = df.groupby("year_week")["date"].nunique()
    return wk_days[wk_days == 7].index.tolist()


def week_agg(df, year_week, dept_filter=None):
    sub = df[df["year_week"] == year_week]
    if dept_filter and dept_filter != "🏭 Factory (ทั้งหมด)":
        sub = sub[sub["department"] == dept_filter]
    return sub[["on_peak", "off_peak", "total"]].sum()


def dept_week_agg(df, year_week):
    sub = df[df["year_week"] == year_week]
    return sub.groupby("department")[["on_peak", "off_peak", "total"]].sum().reset_index()


# ─── Monthly helpers ──────────────────────────────────────────────────────────
def month_label(ym):
    y, m = ym.split("-")
    return f"{MONTH_TH[int(m)]} {y}"


def month_agg(df, ym, dept_filter=None):
    sub = df[df["ym"] == ym]
    if dept_filter and dept_filter != "🏭 Factory (ทั้งหมด)":
        sub = sub[sub["department"] == dept_filter]
    return sub[["on_peak", "off_peak"]].sum()


def month_dept_agg(df, ym):
    sub = df[df["ym"] == ym]
    return sub.groupby("department")[["on_peak", "off_peak"]].sum().reset_index()


def daily_agg(df, ym, dept_filter=None):
    sub = df[df["ym"] == ym]
    if dept_filter and dept_filter != "🏭 Factory (ทั้งหมด)":
        sub = sub[sub["department"] == dept_filter]
    return sub.groupby("date")[["on_peak", "off_peak"]].sum().reset_index().sort_values("date")


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
    st.caption("**อัตราค่าไฟ PEA TOU**")
    st.caption(f"• On Peak  : {ON_PEAK_RATE + FT_ADJ:.4f} ฿/kWh")
    st.caption(f"• Off Peak : {OFF_PEAK_RATE + FT_ADJ:.4f} ฿/kWh")
    st.caption(f"• Ft Surcharge : {FT_ADJ} ฿/kWh")

# ─── Load Data ────────────────────────────────────────────────────────────────
st.markdown('<div class="main-title">⚡ Electricity Usage Dashboard</div>', unsafe_allow_html=True)

with st.spinner("⏳ กำลังดึงข้อมูลจาก Google Sheet..."):
    df = load_data()

# ─── Tabs ─────────────────────────────────────────────────────────────────────
tab_weekly, tab_monthly = st.tabs(["📋 Weekly Overview", "📅 Monthly Dashboard"])


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — WEEKLY
# ══════════════════════════════════════════════════════════════════════════════
with tab_weekly:
    all_weeks      = get_all_weeks(df)
    complete_weeks = sorted(get_complete_weeks(df))

    if not all_weeks:
        st.error("ไม่พบข้อมูล กรุณาตรวจสอบ Google Sheet")
        st.stop()

    # Week selector
    week_display_map = {}
    for yw in all_weeks:
        label      = week_label(yw, df)
        suffix     = "" if yw in complete_weeks else " ⚠️ ไม่ครบ 7 วัน"
        week_display_map[yw] = f"{yw}  ·  {label}{suffix}"

    display_options = list(week_display_map.values())
    yw_keys         = list(week_display_map.keys())
    default_yw      = complete_weeks[-1] if complete_weeks else all_weeks[-1]
    default_idx     = yw_keys.index(default_yw)

    col_sel1, col_sel2, col_sel3 = st.columns([1, 2, 1])
    with col_sel2:
        selected_display = st.selectbox(
            "📅 เลือกสัปดาห์",
            options=display_options,
            index=default_idx,
            help="เลือกสัปดาห์ที่ต้องการดู (⚠️ = ข้อมูลยังไม่ครบ 7 วัน)",
            key="week_selector",
        )

    selected_yw  = yw_keys[display_options.index(selected_display)]
    selected_idx = yw_keys.index(selected_yw)
    prev_yw      = yw_keys[selected_idx - 1] if selected_idx > 0 else None
    sel_label    = week_label(selected_yw, df)
    prev_label   = week_label(prev_yw, df) if prev_yw else "N/A"

    st.markdown(f'<div class="week-subtitle">📆 สัปดาห์ {selected_yw} &nbsp;|&nbsp; {sel_label}</div>',
                unsafe_allow_html=True)
    st.markdown("---")

    # KPI Cards
    agg       = week_agg(df, selected_yw)
    on_kwh    = agg["on_peak"]
    off_kwh   = agg["off_peak"]
    total_kwh = on_kwh + off_kwh
    cost      = on_kwh * (ON_PEAK_RATE + FT_ADJ) + off_kwh * (OFF_PEAK_RATE + FT_ADJ)
    on_pct    = on_kwh  / total_kwh * 100 if total_kwh else 0
    off_pct   = off_kwh / total_kwh * 100 if total_kwh else 0

    if prev_yw:
        agg_p      = week_agg(df, prev_yw)
        prev_total = agg_p["on_peak"] + agg_p["off_peak"]
        chg_total  = (total_kwh - prev_total)        / prev_total        * 100 if prev_total        else 0
        chg_on     = (on_kwh    - agg_p["on_peak"])  / agg_p["on_peak"]  * 100 if agg_p["on_peak"]  else 0
        chg_off    = (off_kwh   - agg_p["off_peak"]) / agg_p["off_peak"] * 100 if agg_p["off_peak"] else 0
    else:
        chg_total = chg_on = chg_off = 0

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">Total Energy This Week</div>
            <div class="kpi-value">{total_kwh:,.0f} <span class="kpi-unit">kWh</span></div>
            <div class="kpi-sub">vs สัปดาห์ก่อน &nbsp; {badge(chg_total)}</div>
        </div>""", unsafe_allow_html=True)
    with c2:
        st.markdown(f"""
        <div class="kpi-card" style="border-top-color:#e65100">
            <div class="kpi-label">On Peak Usage</div>
            <div class="kpi-value kpi-on">{on_kwh:,.0f} <span class="kpi-unit">kWh</span></div>
            <div class="kpi-sub">({on_pct:.0f}%) &nbsp; {badge(chg_on)}</div>
        </div>""", unsafe_allow_html=True)
    with c3:
        st.markdown(f"""
        <div class="kpi-card" style="border-top-color:#2e7d32">
            <div class="kpi-label">Off Peak Usage</div>
            <div class="kpi-value kpi-off">{off_kwh:,.0f} <span class="kpi-unit">kWh</span></div>
            <div class="kpi-sub">({off_pct:.0f}%) &nbsp; {badge(chg_off)}</div>
        </div>""", unsafe_allow_html=True)
    with c4:
        st.markdown(f"""
        <div class="kpi-card" style="border-top-color:#6a1b9a">
            <div class="kpi-label">Cost Estimate</div>
            <div class="kpi-value kpi-cost">{cost:,.0f} <span class="kpi-unit">฿</span></div>
            <div class="kpi-sub">คำนวณจากอัตรา TOU + Ft</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Section 2: Weekly Comparison
    st.markdown('<div class="section-header">📊 Weekly Usage Comparison</div>', unsafe_allow_html=True)

    departments    = sorted(df["department"].unique().tolist())
    filter_options = ["🏭 Factory (ทั้งหมด)"] + departments

    col_f, col_g = st.columns([1, 3])
    with col_f:
        dept_sel = st.selectbox("🔍 เลือกแผนก", filter_options, index=0, key="week_dept_sel")

    agg_cur  = week_agg(df, selected_yw, dept_sel)
    agg_prev = week_agg(df, prev_yw,     dept_sel) if prev_yw else None

    label_cur  = f"สัปดาห์นี้\n({sel_label})"
    label_prev = f"สัปดาห์ก่อน\n({prev_label})" if prev_yw else "สัปดาห์ก่อน"

    on_vals  = [agg_prev["on_peak"]  if agg_prev is not None else 0, agg_cur["on_peak"]]
    off_vals = [agg_prev["off_peak"] if agg_prev is not None else 0, agg_cur["off_peak"]]
    x_labels = [label_prev, label_cur]
    tot_vals = [on_vals[i] + off_vals[i] for i in range(2)]
    on_pct_bars  = [on_vals[i]  / tot_vals[i] * 100 if tot_vals[i] else 0 for i in range(2)]
    off_pct_bars = [off_vals[i] / tot_vals[i] * 100 if tot_vals[i] else 0 for i in range(2)]

    fig_w = go.Figure()
    fig_w.add_trace(go.Bar(
        name="On Peak", x=x_labels, y=on_vals, marker_color="#e65100",
        text=[f"{v:,.0f} kWh  ({on_pct_bars[i]:.0f}%)" for i, v in enumerate(on_vals)],
        textposition="inside", insidetextanchor="middle",
        textfont=dict(color="white", size=16),
    ))
    fig_w.add_trace(go.Bar(
        name="Off Peak", x=x_labels, y=off_vals, marker_color="#1565c0",
        text=[f"{v:,.0f} kWh  ({off_pct_bars[i]:.0f}%)" for i, v in enumerate(off_vals)],
        textposition="inside", insidetextanchor="middle",
        textfont=dict(color="white", size=16),
    ))
    for xl, tv in zip(x_labels, tot_vals):
        fig_w.add_annotation(
            x=xl, y=tv, text=f"<b>{tv:,.0f} kWh</b>",
            showarrow=False, yshift=16,
            font=dict(size=16, color="#1a237e"),
        )
    fig_w.update_layout(
        barmode="stack", height=400,
        yaxis_title="kWh",
        title_text=f"On Peak vs Off Peak — {dept_sel}",
        title_font_size=16,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(t=50, b=20, l=20, r=60),
        plot_bgcolor="white", paper_bgcolor="white",
    )
    fig_w.update_yaxes(gridcolor="#f0f0f0")

    with col_g:
        st.plotly_chart(fig_w, use_container_width=True)

    # Summary block
    cur_total_w  = agg_cur["on_peak"]  + agg_cur["off_peak"]
    prev_total_w = (agg_prev["on_peak"] + agg_prev["off_peak"]) if agg_prev is not None else 0
    diff_w       = cur_total_w - prev_total_w
    pct_w        = (diff_w / prev_total_w * 100) if prev_total_w else 0
    cur_on_pct   = agg_cur["on_peak"]  / cur_total_w * 100 if cur_total_w else 0
    cur_off_pct  = agg_cur["off_peak"] / cur_total_w * 100 if cur_total_w else 0
    direction    = "เพิ่มขึ้น" if diff_w >= 0 else "ลดลง"
    dir_color    = "#e53935"   if diff_w >= 0 else "#43a047"
    dir_icon     = "▲"         if diff_w >= 0 else "▼"
    summary_dept = dept_sel if dept_sel != "🏭 Factory (ทั้งหมด)" else "ทั้งโรงงาน"
    prev_wk_txt  = f"เทียบกับสัปดาห์ก่อน {prev_total_w:,.0f} kWh" if prev_yw else ""

    st.markdown(f"""
<div style="background:#f8f9ff;border-left:4px solid #1565c0;border-radius:8px;
            padding:14px 22px;margin-top:4px;font-size:14px;line-height:2.0;color:#333;">
  📋 <b>สรุปการใช้ไฟฟ้า &mdash; {summary_dept} &nbsp;|&nbsp; {sel_label}</b><br>
  สัปดาห์นี้ใช้ไฟฟ้ารวมทั้งสิ้น <b>{cur_total_w:,.0f} kWh</b><br>
  &nbsp;&nbsp;&nbsp;
  <span style="color:#e65100;font-weight:600;">● On Peak</span>&nbsp;&nbsp;
  <b>{agg_cur["on_peak"]:,.0f} kWh</b>
  &nbsp;<span style="color:#888;">({cur_on_pct:.1f}%)</span>
  &emsp;
  <span style="color:#1565c0;font-weight:600;">● Off Peak</span>&nbsp;
  <b>{agg_cur["off_peak"]:,.0f} kWh</b>
  &nbsp;<span style="color:#888;">({cur_off_pct:.1f}%)</span><br>
  {prev_wk_txt} &nbsp;→&nbsp;
  <span style="color:{dir_color};font-weight:700;">{dir_icon} {direction} {abs(diff_w):,.0f} kWh &nbsp;({abs(pct_w):.1f}%)</span>
</div>
""", unsafe_allow_html=True)

    # Section 3: Department Breakdown
    st.markdown('<div class="section-header">🏭 Department Usage Breakdown</div>', unsafe_allow_html=True)

    dept_cur  = dept_week_agg(df, selected_yw).set_index("department")
    dept_prev = dept_week_agg(df, prev_yw).set_index("department") if prev_yw else None
    dept_cur["total_combined"] = dept_cur["on_peak"] + dept_cur["off_peak"]
    all_depts = dept_cur.sort_values("total_combined", ascending=False).index.tolist()

    fig_d = make_subplots(
        rows=len(all_depts), cols=1,
        shared_xaxes=True,
        subplot_titles=all_depts,
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

        fig_d.add_trace(go.Bar(name="On Peak (ก่อน)",  x=[prv_on],  y=["ก่อน"], orientation="h", marker_color="#ffcc80", legendgroup="p_on",  showlegend=show_leg, hovertemplate=f"{dep} On Peak (ก่อน): %{{x:,.0f}} kWh<extra></extra>"), row=i, col=1)
        fig_d.add_trace(go.Bar(name="Off Peak (ก่อน)", x=[prv_off], y=["ก่อน"], orientation="h", marker_color="#90caf9", legendgroup="p_off", showlegend=show_leg, hovertemplate=f"{dep} Off Peak (ก่อน): %{{x:,.0f}} kWh<extra></extra>"), row=i, col=1)
        fig_d.add_trace(go.Bar(name="On Peak (นี้)",   x=[cur_on],  y=["นี้"],  orientation="h", marker_color="#e65100", legendgroup="c_on",  showlegend=show_leg, hovertemplate=f"{dep} On Peak (นี้): %{{x:,.0f}} kWh<extra></extra>"), row=i, col=1)
        fig_d.add_trace(go.Bar(name="Off Peak (นี้)",  x=[cur_off], y=["นี้"],  orientation="h", marker_color="#1565c0", legendgroup="c_off", showlegend=show_leg, hovertemplate=f"{dep} Off Peak (นี้): %{{x:,.0f}} kWh<extra></extra>"), row=i, col=1)

        max_val = max(cur_on + cur_off, prv_on + prv_off, 1)
        fig_d.add_annotation(x=max_val * 1.02, y=0.5, yref=f"y{i}", xref=f"x{i}", text=f"<b>{arrow_txt}</b>", showarrow=False, font=dict(color=clr, size=11), xanchor="left")

    fig_d.update_layout(
        barmode="stack",
        height=max(130 * len(all_depts), 500),
        legend=dict(orientation="h", yanchor="bottom", y=1.005, xanchor="right", x=1),
        margin=dict(l=80, r=120, t=60, b=20),
        plot_bgcolor="white", paper_bgcolor="white",
    )
    for i in range(1, len(all_depts) + 1):
        fig_d.update_xaxes(gridcolor="#f0f0f0", row=i, col=1)
    st.plotly_chart(fig_d, use_container_width=True)

    # Footer
    st.markdown("---")
    last_updated = datetime.now().strftime("%d/%m/%Y %H:%M")
    num_days_selected = df[df["year_week"] == selected_yw]["date"].nunique()
    st.caption(
        f"📅 สัปดาห์ที่เลือก: **{selected_yw}** ({sel_label}) | "
        f"จำนวนข้อมูล: **{num_days_selected} วัน** | "
        f"สัปดาห์ก่อน: **{prev_yw or 'N/A'}** ({prev_label}) | "
        f"🕐 โหลดล่าสุด: {last_updated}"
    )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — MONTHLY
# ══════════════════════════════════════════════════════════════════════════════
with tab_monthly:
    all_ym = sorted(df["ym"].unique().tolist())
    if not all_ym:
        st.error("ไม่พบข้อมูล")
        st.stop()

    ym_display = {ym: month_label(ym) for ym in all_ym}

    col_s1, col_s2, col_s3 = st.columns([1, 2, 1])
    with col_s2:
        selected_ym = st.selectbox(
            "📅 เลือกเดือน",
            options=all_ym,
            format_func=lambda x: ym_display[x],
            index=len(all_ym) - 1,
            key="month_selector",
        )

    sel_idx    = all_ym.index(selected_ym)
    prev_ym    = all_ym[sel_idx - 1] if sel_idx > 0 else None
    sel_label  = month_label(selected_ym)
    prev_label = month_label(prev_ym) if prev_ym else "N/A"

    st.markdown(f'<div class="month-subtitle">📆 เดือน {sel_label}</div>', unsafe_allow_html=True)
    st.markdown("---")

    # ── Department Filter (same style as Weekly) ──────────────────────────────
    m_departments    = sorted(df["department"].unique().tolist())
    m_filter_options = ["🏭 Factory (ทั้งหมด)"] + m_departments

    col_mf, col_mg = st.columns([1, 3])
    with col_mf:
        m_dept_sel = st.selectbox("🔍 เลือกแผนก", m_filter_options, index=0, key="month_dept_sel")

    # KPI Cards (filtered)
    agg     = month_agg(df, selected_ym, m_dept_sel)
    on_kwh  = agg["on_peak"]
    off_kwh = agg["off_peak"]
    total   = on_kwh + off_kwh
    cost    = on_kwh * (ON_PEAK_RATE + FT_ADJ) + off_kwh * (OFF_PEAK_RATE + FT_ADJ)
    on_pct  = on_kwh  / total * 100 if total else 0
    off_pct = off_kwh / total * 100 if total else 0

    if prev_ym:
        agg_p    = month_agg(df, prev_ym, m_dept_sel)
        prev_tot = agg_p["on_peak"] + agg_p["off_peak"]
        chg_tot  = (total   - prev_tot)          / prev_tot          * 100 if prev_tot          else 0
        chg_on   = (on_kwh  - agg_p["on_peak"])  / agg_p["on_peak"]  * 100 if agg_p["on_peak"]  else 0
        chg_off  = (off_kwh - agg_p["off_peak"]) / agg_p["off_peak"] * 100 if agg_p["off_peak"] else 0
    else:
        chg_tot = chg_on = chg_off = 0

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">Total Energy This Month</div>
            <div class="kpi-value">{total:,.0f} <span class="kpi-unit">kWh</span></div>
            <div class="kpi-sub">vs เดือนก่อน &nbsp; {badge(chg_tot)}</div>
        </div>""", unsafe_allow_html=True)
    with c2:
        st.markdown(f"""
        <div class="kpi-card" style="border-top-color:#e65100">
            <div class="kpi-label">On Peak Usage</div>
            <div class="kpi-value kpi-on">{on_kwh:,.0f} <span class="kpi-unit">kWh</span></div>
            <div class="kpi-sub">({on_pct:.0f}%) &nbsp; {badge(chg_on)}</div>
        </div>""", unsafe_allow_html=True)
    with c3:
        st.markdown(f"""
        <div class="kpi-card" style="border-top-color:#2e7d32">
            <div class="kpi-label">Off Peak Usage</div>
            <div class="kpi-value kpi-off">{off_kwh:,.0f} <span class="kpi-unit">kWh</span></div>
            <div class="kpi-sub">({off_pct:.0f}%) &nbsp; {badge(chg_off)}</div>
        </div>""", unsafe_allow_html=True)
    with c4:
        st.markdown(f"""
        <div class="kpi-card" style="border-top-color:#6a1b9a">
            <div class="kpi-label">Cost Estimate</div>
            <div class="kpi-value kpi-cost">{cost:,.0f} <span class="kpi-unit">฿</span></div>
            <div class="kpi-sub">คำนวณจากอัตรา TOU + Ft</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Section 1: Daily Trend (filtered)
    st.markdown('<div class="section-header">📈 Daily Usage Trend</div>', unsafe_allow_html=True)

    daily = daily_agg(df, selected_ym, m_dept_sel)
    daily["total"] = daily["on_peak"] + daily["off_peak"]

    m_dept_label = m_dept_sel if m_dept_sel != "🏭 Factory (ทั้งหมด)" else "ทั้งโรงงาน"

    fig_trend = go.Figure()
    fig_trend.add_trace(go.Bar(x=daily["date"], y=daily["off_peak"], name="Off Peak", marker_color="#1565c0", opacity=0.85))
    fig_trend.add_trace(go.Bar(x=daily["date"], y=daily["on_peak"],  name="On Peak",  marker_color="#e65100", opacity=0.85))
    fig_trend.add_trace(go.Scatter(
        x=daily["date"], y=daily["total"],
        name="Total", mode="lines+markers",
        line=dict(color="#1a237e", width=2.5),
        marker=dict(size=8, color="#1a237e"),
    ))
    fig_trend.update_layout(
        barmode="stack", height=420,
        yaxis_title="kWh",
        title_text=f"การใช้ไฟฟ้ารายวัน — {sel_label} | {m_dept_label}",
        title_font_size=17,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1,
                    font=dict(size=15)),
        margin=dict(t=60, b=30, l=30, r=30),
        plot_bgcolor="white", paper_bgcolor="white",
        hovermode="x unified",
        font=dict(size=14),
    )
    fig_trend.update_xaxes(dtick="D1", tickformat="%d %b", tickangle=-45, gridcolor="#f0f0f0",
                           tickfont=dict(size=13))
    fig_trend.update_yaxes(gridcolor="#f0f0f0", tickfont=dict(size=13))
    st.plotly_chart(fig_trend, use_container_width=True)

    # Section 2: Month-over-Month Comparison (filtered, uniform colors = same as Daily Trend)
    st.markdown('<div class="section-header">📊 Month-over-Month Comparison</div>', unsafe_allow_html=True)

    mom_rows = []
    for ym in all_ym:
        a = month_agg(df, ym, m_dept_sel)
        mom_rows.append({
            "ym": ym,
            "label":    month_label(ym),
            "on_peak":  a["on_peak"],
            "off_peak": a["off_peak"],
            "total":    a["on_peak"] + a["off_peak"],
        })
    mom_df = pd.DataFrame(mom_rows)

    fig_mom = go.Figure()
    fig_mom.add_trace(go.Bar(
        name="On Peak", x=mom_df["label"], y=mom_df["on_peak"],
        marker_color="#e65100", opacity=0.85,
        text=mom_df["on_peak"].apply(lambda v: f"{v:,.0f}"),
        textposition="inside", insidetextanchor="middle",
        textfont=dict(color="white", size=13),
    ))
    fig_mom.add_trace(go.Bar(
        name="Off Peak", x=mom_df["label"], y=mom_df["off_peak"],
        marker_color="#1565c0", opacity=0.85,
        text=mom_df["off_peak"].apply(lambda v: f"{v:,.0f}"),
        textposition="inside", insidetextanchor="middle",
        textfont=dict(color="white", size=13),
    ))
    for _, row in mom_df.iterrows():
        fig_mom.add_annotation(
            x=row["label"], y=row["total"],
            text=f"<b>{row['total']:,.0f}</b>",
            showarrow=False, yshift=12,
            font=dict(size=13, color="#1a237e"),
        )
    fig_mom.update_layout(
        barmode="stack", height=420,
        yaxis_title="kWh",
        title_text=f"เปรียบเทียบการใช้ไฟฟ้ารายเดือน — {m_dept_label}",
        title_font_size=17,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1,
                    font=dict(size=15)),
        margin=dict(t=60, b=40, l=30, r=30),
        plot_bgcolor="white", paper_bgcolor="white",
        font=dict(size=14),
    )
    fig_mom.update_xaxes(tickfont=dict(size=13))
    fig_mom.update_yaxes(gridcolor="#f0f0f0", tickfont=dict(size=13))
    st.plotly_chart(fig_mom, use_container_width=True)

    # Summary block
    if prev_ym:
        prev_tot_val = month_agg(df, prev_ym, m_dept_sel)["on_peak"] + month_agg(df, prev_ym, m_dept_sel)["off_peak"]
        diff_v     = total - prev_tot_val
        pct_v      = diff_v / prev_tot_val * 100 if prev_tot_val else 0
        direction  = "เพิ่มขึ้น" if diff_v >= 0 else "ลดลง"
        dir_color  = "#e53935"   if diff_v >= 0 else "#43a047"
        dir_icon   = "▲"         if diff_v >= 0 else "▼"
        prev_txt   = f"เทียบกับ {prev_label}: {prev_tot_val:,.0f} kWh"

        st.markdown(f"""
<div style="background:#f8f9ff;border-left:4px solid #1565c0;border-radius:8px;
            padding:16px 24px;margin-top:4px;font-size:16px;line-height:2.2;color:#333;">
  📋 <b>สรุปเดือน {sel_label} — {m_dept_label}</b><br>
  ใช้ไฟฟ้ารวม <b>{total:,.0f} kWh</b> &nbsp;|&nbsp;
  <span style="color:#e65100;font-weight:600;">● On Peak</span> <b>{on_kwh:,.0f} kWh</b>
  <span style="color:#888;">({on_pct:.1f}%)</span> &emsp;
  <span style="color:#1565c0;font-weight:600;">● Off Peak</span> <b>{off_kwh:,.0f} kWh</b>
  <span style="color:#888;">({off_pct:.1f}%)</span><br>
  {prev_txt} &nbsp;→&nbsp;
  <span style="color:{dir_color};font-weight:700;">{dir_icon} {direction} {abs(diff_v):,.0f} kWh ({abs(pct_v):.1f}%)</span>
</div>
""", unsafe_allow_html=True)

    # ── Section 3: Department Summary Table ──────────────────────────────────
    st.markdown('<div class="section-header">📋 Department Summary Table</div>', unsafe_allow_html=True)

    all_depts_m = sorted(df["department"].unique().tolist())
    table_rows  = []
    for dep in all_depts_m:
        # Current month
        cur = df[(df["ym"] == selected_ym) & (df["department"] == dep)][["on_peak", "off_peak"]].sum()
        cur_on  = cur["on_peak"]
        cur_off = cur["off_peak"]
        cur_tot = cur_on + cur_off
        cur_on_pct  = cur_on  / cur_tot * 100 if cur_tot else 0
        cur_off_pct = cur_off / cur_tot * 100 if cur_tot else 0
        cur_ratio   = f"{cur_on_pct:.0f} : {cur_off_pct:.0f}"

        # Previous month
        if prev_ym:
            prv = df[(df["ym"] == prev_ym) & (df["department"] == dep)][["on_peak", "off_peak"]].sum()
            prv_on  = prv["on_peak"]
            prv_off = prv["off_peak"]
            prv_tot = prv_on + prv_off
            prv_on_pct  = prv_on  / prv_tot * 100 if prv_tot else 0
            prv_off_pct = prv_off / prv_tot * 100 if prv_tot else 0
            prv_ratio   = f"{prv_on_pct:.0f} : {prv_off_pct:.0f}"
            chg         = (cur_tot - prv_tot) / prv_tot * 100 if prv_tot else 0
        else:
            prv_on = prv_off = prv_tot = 0
            prv_ratio = "—"
            chg = 0

        table_rows.append({
            "dept":    dep,
            "cur_tot": cur_tot, "cur_on": cur_on, "cur_off": cur_off,
            "cur_ratio": cur_ratio, "chg": chg,
            "prv_tot": prv_tot, "prv_on": prv_on, "prv_off": prv_off,
            "prv_ratio": prv_ratio,
        })

    # Sort by current total descending
    table_rows.sort(key=lambda r: r["cur_tot"], reverse=True)

    # ── Build HTML table ──────────────────────────────────────────────────────
    def chg_cell(v):
        if v == 0:
            return '<span style="color:#888;">—</span>'
        color  = "#e53935" if v > 0 else "#43a047"
        arrow  = "▲" if v > 0 else "▼"
        return f'<span style="color:{color};font-weight:700;">{arrow} {abs(v):.1f}%</span>'

    def num(v):
        return f"{v:,.0f}" if v else "—"

    rows_html = ""
    for r in table_rows:
        rows_html += f"""
        <tr>
          <td style="font-weight:600;white-space:nowrap">{r['dept']}</td>
          <td class="num">{num(r['cur_tot'])}</td>
          <td style="text-align:center">{chg_cell(r['chg'])}</td>
          <td class="num" style="color:#e65100">{num(r['cur_on'])}</td>
          <td class="num" style="color:#1565c0">{num(r['cur_off'])}</td>
          <td style="text-align:center">{r['cur_ratio']}</td>
          <td class="sep"></td>
          <td class="num muted">{num(r['prv_tot'])}</td>
          <td class="num muted" style="color:#e65100">{num(r['prv_on'])}</td>
          <td class="num muted" style="color:#1565c0">{num(r['prv_off'])}</td>
          <td style="text-align:center;color:#aaa">{r['prv_ratio']}</td>
        </tr>"""

    # Totals row
    gt_cur_on  = sum(r["cur_on"]  for r in table_rows)
    gt_cur_off = sum(r["cur_off"] for r in table_rows)
    gt_cur_tot = gt_cur_on + gt_cur_off
    gt_prv_on  = sum(r["prv_on"]  for r in table_rows)
    gt_prv_off = sum(r["prv_off"] for r in table_rows)
    gt_prv_tot = gt_prv_on + gt_prv_off
    gt_chg     = (gt_cur_tot - gt_prv_tot) / gt_prv_tot * 100 if gt_prv_tot else 0
    gt_cur_on_pct  = gt_cur_on  / gt_cur_tot * 100 if gt_cur_tot else 0
    gt_cur_off_pct = gt_cur_off / gt_cur_tot * 100 if gt_cur_tot else 0
    gt_prv_on_pct  = gt_prv_on  / gt_prv_tot * 100 if gt_prv_tot else 0
    gt_prv_off_pct = gt_prv_off / gt_prv_tot * 100 if gt_prv_tot else 0

    rows_html += f"""
        <tr class="total-row">
          <td style="font-weight:800">🏭 รวมทั้งหมด</td>
          <td class="num">{num(gt_cur_tot)}</td>
          <td style="text-align:center">{chg_cell(gt_chg)}</td>
          <td class="num" style="color:#e65100">{num(gt_cur_on)}</td>
          <td class="num" style="color:#1565c0">{num(gt_cur_off)}</td>
          <td style="text-align:center">{gt_cur_on_pct:.0f} : {gt_cur_off_pct:.0f}</td>
          <td class="sep"></td>
          <td class="num muted">{num(gt_prv_tot)}</td>
          <td class="num muted" style="color:#e65100">{num(gt_prv_on)}</td>
          <td class="num muted" style="color:#1565c0">{num(gt_prv_off)}</td>
          <td style="text-align:center;color:#aaa">{gt_prv_on_pct:.0f} : {gt_prv_off_pct:.0f}</td>
        </tr>"""

    prev_header = prev_label if prev_ym else "เดือนก่อน"

    table_html = f"""
<style>
  .sum-table {{
    width:100%; border-collapse:collapse; font-size:15px;
    border-radius:10px; overflow:hidden;
    box-shadow: 0 2px 12px rgba(0,0,0,0.08);
  }}
  .sum-table th {{
    background:#1a237e; color:#fff; padding:11px 14px;
    font-size:13px; font-weight:700; letter-spacing:0.3px;
    white-space:nowrap;
  }}
  .sum-table th.group-cur {{ background:#1565c0; }}
  .sum-table th.group-prv {{ background:#546e7a; }}
  .sum-table td {{ padding:10px 14px; border-bottom:1px solid #f0f0f0; vertical-align:middle; }}
  .sum-table tr:hover td {{ background:#f5f8ff; }}
  .sum-table .num {{ text-align:right; font-variant-numeric:tabular-nums; font-weight:600; }}
  .sum-table .muted {{ color:#888 !important; font-weight:400 !important; }}
  .sum-table .sep {{ width:6px; background:#e8eaf6; padding:0; }}
  .sum-table .total-row td {{
    background:#e8eaf6; font-weight:700;
    border-top:2px solid #1565c0;
  }}
  .sum-table .total-row:hover td {{ background:#dde2f0; }}
</style>
<div style="overflow-x:auto; margin-top:8px;">
<table class="sum-table">
  <thead>
    <tr>
      <th rowspan="2" style="text-align:left; background:#1a237e;">แผนก</th>
      <th colspan="5" class="group-cur" style="text-align:center; border-bottom:1px solid rgba(255,255,255,0.3);">
        📅 เดือนนี้ — {sel_label}
      </th>
      <th rowspan="2" class="sep" style="width:6px;padding:0;"></th>
      <th colspan="4" class="group-prv" style="text-align:center; border-bottom:1px solid rgba(255,255,255,0.3);">
        🕓 เดือนก่อน — {prev_header}
      </th>
    </tr>
    <tr>
      <th class="group-cur" style="text-align:right;">Total (kWh)</th>
      <th class="group-cur" style="text-align:center;">เทียบเดือนก่อน</th>
      <th class="group-cur" style="text-align:right; color:#ffcc80;">On Peak (kWh)</th>
      <th class="group-cur" style="text-align:right; color:#90caf9;">Off Peak (kWh)</th>
      <th class="group-cur" style="text-align:center;">On:Off Ratio</th>
      <th class="group-prv" style="text-align:right;">Total (kWh)</th>
      <th class="group-prv" style="text-align:right; color:#ffcc80;">On Peak (kWh)</th>
      <th class="group-prv" style="text-align:right; color:#90caf9;">Off Peak (kWh)</th>
      <th class="group-prv" style="text-align:center;">On:Off Ratio</th>
    </tr>
  </thead>
  <tbody>
    {rows_html}
  </tbody>
</table>
</div>
"""
    st.markdown(table_html, unsafe_allow_html=True)

    # Footer
    st.markdown("---")
    num_days     = df[df["ym"] == selected_ym]["date"].nunique()
    last_updated = datetime.now().strftime("%d/%m/%Y %H:%M")
    st.caption(
        f"📅 เดือนที่เลือก: **{sel_label}** | "
        f"จำนวนข้อมูล: **{num_days} วัน** | "
        f"เดือนก่อน: **{prev_label}** | "
        f"🕐 โหลดล่าสุด: {last_updated}"
    )
