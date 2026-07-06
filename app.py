
import os
from io import BytesIO
from typing import Dict, Optional, Any

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


st.set_page_config(
    page_title="SEP V4 Enterprise",
    page_icon="📊",
    layout="wide",
)

DEFAULT_FILE = os.path.join("data", "SEP_V4_Enterprise_Build_Template.xlsx")

REQUIRED_SHEETS = [
    "PARAMETERS", "CALENDAR", "TARGET_FRAMEWORK", "DIM_DEPARTMENT",
    "REVENUE_YR", "HEADCOUNT_YR", "PAYROLL_YR",
]
OPTIONAL_SHEETS = [
    "REVENUE_BY_DEPT", "CAPACITY_INDICATOR", "PROJECT_PRODUCTIVITY",
    "MONTHLY_KPI", "PROJECT_MARGIN",
]

TONE_COLOR = {
    "success": "#22C55E",
    "info": "#2563EB",
    "warning": "#F59E0B",
    "error": "#EF4444",
    "neutral": "#64748B",
}


# =========================
# BASIC FORMATTERS
# =========================
def safe_float(x, default=np.nan):
    try:
        if x is None or pd.isna(x):
            return default
        return float(x)
    except Exception:
        return default


def money(x, currency="IDR"):
    x = safe_float(x)
    if pd.isna(x):
        return "-"
    if abs(x) >= 1_000_000_000:
        return f"{currency} {x/1_000_000_000:,.1f}B".replace(",", ".")
    if abs(x) >= 1_000_000:
        return f"{currency} {x/1_000_000:,.1f}M".replace(",", ".")
    return f"{currency} {x:,.0f}".replace(",", ".")


def pct(x):
    x = safe_float(x)
    return "-" if pd.isna(x) else f"{x*100:.1f}%"


def pp(x):
    x = safe_float(x)
    return "-" if pd.isna(x) else f"{x*100:+.1f} pp"


def num(x):
    x = safe_float(x)
    return "-" if pd.isna(x) else f"{x:,.1f}".replace(",", ".")


# =========================
# PRESENTATION MODE ENGINE
# =========================
def get_presentation_scale(mode: str) -> Dict[str, Any]:
    if mode == "Smartphone":
        return dict(global_px=14.5, h1=1.70, h2=1.45, h3=1.22, metric=1.65, tab=13,
                    plot_font=12, plot_title=15, axis=12, legend=11, bubble=50, cap_bubble=42,
                    chart=380, gauge=230, padding="0.65rem", max_width="100%")
    if mode == "Compact":
        return dict(global_px=15.0, h1=2.20, h2=1.78, h3=1.34, metric=1.95, tab=14,
                    plot_font=13, plot_title=18, axis=13, legend=13, bubble=64, cap_bubble=54,
                    chart=470, gauge=280, padding="1.25rem", max_width="1540px")
    if mode == "Boardroom":
        return dict(global_px=18.0, h1=2.85, h2=2.28, h3=1.78, metric=2.75, tab=17,
                    plot_font=16, plot_title=22, axis=16, legend=16, bubble=92, cap_bubble=76,
                    chart=560, gauge=350, padding="2.4rem", max_width="1560px")
    return dict(global_px=16.5, h1=2.38, h2=1.98, h3=1.52, metric=2.18, tab=15,
                plot_font=14, plot_title=20, axis=14, legend=14, bubble=78, cap_bubble=64,
                chart=520, gauge=310, padding="2.4rem", max_width="1560px")


def inject_css(scale):
    st.markdown(f"""
    <style>
    html, body, [class*="css"] {{ font-size: {scale['global_px']}px !important; }}
    .main .block-container {{
        max-width: {scale['max_width']};
        padding-top: 1rem;
        padding-left: {scale['padding']};
        padding-right: {scale['padding']};
    }}
    h1 {{ font-size: {scale['h1']}rem !important; letter-spacing:-.04em; }}
    h2 {{ font-size: {scale['h2']}rem !important; }}
    h3 {{ font-size: {scale['h3']}rem !important; }}
    div[data-testid="stMetricValue"] {{ font-size: {scale['metric']}rem !important; font-weight: 850 !important; }}
    div[data-testid="stMetricLabel"] {{ font-weight: 800 !important; }}
    .stTabs [data-baseweb="tab"] {{
        font-size: {scale['tab']}px !important;
        font-weight: 800 !important;
        padding-left: 8px !important;
        padding-right: 8px !important;
    }}
    .stTabs [aria-selected="true"] {{ color:#EF4444 !important; border-bottom:3px solid #EF4444; }}
    [data-testid="stSidebar"] label, [data-testid="stSidebar"] p, [data-testid="stSidebar"] span {{
        font-size: {max(13, scale['tab'])}px !important;
    }}
    .sep-header {{
        display:flex; justify-content:space-between; align-items:flex-start; gap:24px; margin-bottom:14px;
    }}
    .sep-title {{ font-size:{scale['h1']}rem; font-weight:900; color:#172033; line-height:1.05; }}
    .sep-sub {{ color:#334155; font-weight:700; margin-top:6px; }}
    .sep-desc {{ color:#64748B; margin-top:8px; }}
    .sep-logo img {{ max-height:56px !important; width:auto !important; object-fit:contain; }}
    .feature-band {{
        display:grid; grid-template-columns:repeat(5,1fr); gap:14px; padding:16px 18px; margin:12px 0 18px 0;
        background:#F8FAFC; border:1px solid #E5EAF2; border-radius:16px; box-shadow:0 8px 24px rgba(15,23,42,.04);
    }}
    .feature-item {{ display:grid; grid-template-columns:38px 1fr; gap:10px; align-items:start; }}
    .feature-icon {{ width:34px; height:34px; display:flex; align-items:center; justify-content:center; border-radius:12px; background:#EEF6FF; font-size:20px; }}
    .feature-title {{ font-weight:850; color:#0F172A; font-size:{max(13, scale['tab'])}px; }}
    .feature-text {{ color:#64748B; font-size:{max(12, scale['tab']-1)}px; line-height:1.45; }}
    .how-box {{ background:linear-gradient(180deg,#F2F7FF 0%,#F8FBFF 100%); border:1px solid #DCEBFF; border-radius:16px; padding:18px 20px; }}
    .how-box b {{ color:#0B63CE; }}
    .how-box li {{ margin-bottom:8px; line-height:1.55; }}
    .insight-card {{ border:1px solid #E5EAF2; border-radius:16px; padding:18px; background:#fff; box-shadow:0 8px 22px rgba(15,23,42,.04); }}
    .memo-card {{ background:linear-gradient(90deg,#EEF6FF 0%,#F8FBFF 100%); border:1px solid #DBEAFE; border-radius:14px; padding:18px 20px; margin-top:12px; line-height:1.65; }}
    @media (max-width: 768px) {{
        .main .block-container {{ padding-left:.65rem !important; padding-right:.65rem !important; }}
        .sep-header {{ display:block !important; }}
        .feature-band {{ grid-template-columns:1fr !important; }}
        div[data-testid="column"] {{ width:100% !important; min-width:100% !important; flex:1 1 100% !important; }}
        .stTabs [data-baseweb="tab-list"] {{ overflow-x:auto !important; flex-wrap:nowrap !important; white-space:nowrap !important; }}
        .stTabs [data-baseweb="tab"] {{ min-width:max-content !important; font-size:12.5px !important; }}
        .sep-logo img {{ max-height:42px !important; }}
    }}
    </style>
    """, unsafe_allow_html=True)


def style_plot(fig, scale, height_key="chart"):
    fig.update_layout(
        height=int(scale[height_key]),
        font=dict(size=int(scale["plot_font"])),
        title_font=dict(size=int(scale["plot_title"])),
        legend=dict(font=dict(size=int(scale["legend"]))),
        margin=dict(l=10, r=10, t=50, b=10),
    )
    fig.update_xaxes(title_font=dict(size=int(scale["axis"])), tickfont=dict(size=int(scale["axis"])))
    fig.update_yaxes(title_font=dict(size=int(scale["axis"])), tickfont=dict(size=int(scale["axis"])))
    return fig


# =========================
# DATA LOADING
# =========================
@st.cache_data(show_spinner=False)
def load_workbook(file_bytes: Optional[bytes]):
    if file_bytes:
        xls = pd.ExcelFile(BytesIO(file_bytes))
    else:
        if not os.path.exists(DEFAULT_FILE):
            return {}, [f"Default file tidak ditemukan: {DEFAULT_FILE}. Upload Excel terlebih dahulu."]
        xls = pd.ExcelFile(DEFAULT_FILE)

    sheets, warnings = {}, []
    for sh in REQUIRED_SHEETS + OPTIONAL_SHEETS:
        if sh in xls.sheet_names:
            try:
                df = pd.read_excel(xls, sheet_name=sh)
                df.columns = [str(c).strip() for c in df.columns]
                sheets[sh] = df
            except Exception as e:
                warnings.append(f"Sheet {sh} gagal dibaca: {e}")
    missing = [s for s in REQUIRED_SHEETS if s not in sheets]
    if missing:
        warnings.append("Sheet wajib belum ada: " + ", ".join(missing))
    return sheets, warnings


def parameters(df):
    if df is None or df.empty or not {"Parameter", "Value"}.issubset(df.columns):
        return {}
    return {str(r["Parameter"]).strip(): str(r["Value"]).strip() for _, r in df.iterrows() if str(r.get("Parameter", "")).strip()}


def read_framework(df):
    if df is None or df.empty:
        return pd.DataFrame()
    f = df.copy()
    f.columns = [str(c).strip().replace(" ", "_") for c in f.columns]
    if "KPI" not in f.columns:
        return pd.DataFrame()
    f["KPI"] = f["KPI"].astype(str).str.strip()
    for c in ["Strong_Min","Strong_Max","Healthy_Min","Healthy_Max","Watch_Min","Watch_Max","Pressure_Min","Pressure_Max","Critical_Min","Critical_Max"]:
        if c in f.columns:
            f[c] = pd.to_numeric(f[c], errors="coerce")
    return f[f["KPI"].ne("")].copy()


def tf_row(tf, kpi):
    if tf is None or tf.empty:
        return None
    x = tf[tf["KPI"].str.lower().eq(kpi.lower())]
    return None if x.empty else x.iloc[0]


def tf_val(tf, kpi, col, default):
    r = tf_row(tf, kpi)
    if r is None or col not in r.index or pd.isna(r[col]):
        return default
    return float(r[col])


def target_text(tf, kpi, fallback):
    r = tf_row(tf, kpi)
    if r is None:
        return fallback
    if "Target_Framework" in r.index and pd.notna(r["Target_Framework"]):
        return str(r["Target_Framework"])
    return fallback


def calendar(df, years):
    if df is None or df.empty:
        return pd.DataFrame({"Year": years, "Months_Closed": [12]*len(years), "Period_Label":[f"FY {y}" for y in years], "Data_Confidence":["High"]*len(years)})
    c = df.copy()
    c["Year"] = pd.to_numeric(c["Year"], errors="coerce").astype("Int64")
    c["Months_Closed"] = pd.to_numeric(c.get("Months_Closed", 12), errors="coerce").fillna(12).clip(1,12)
    c["Period_Label"] = c.get("Period_Label", pd.Series([None]*len(c))).fillna(c["Year"].map(lambda y: f"FY {int(y)}"))
    c["Data_Confidence"] = c.get("Data_Confidence", pd.Series(["High"]*len(c))).fillna("High")
    return c.dropna(subset=["Year"]).astype({"Year": int})[["Year","Months_Closed","Period_Label","Data_Confidence"]]


def mode_factor(cal, mode):
    ytd = cal[cal["Months_Closed"] < 12]
    m = int(ytd["Months_Closed"].max()) if not ytd.empty else 12
    if mode == "Actual":
        return {int(r.Year):1.0 for _,r in cal.iterrows()}, {int(r.Year):1.0 for _,r in cal.iterrows()}, "Actual: angka sesuai input Excel.", m
    if mode == "YTD Comparable":
        return {int(r.Year): m/float(r.Months_Closed) for _,r in cal.iterrows()}, {int(r.Year): m/float(r.Months_Closed) for _,r in cal.iterrows()}, f"YTD Comparable: semua tahun dikonversi ke basis {m} bulan agar apple-to-apple.", m
    if mode == "Annual Projection":
        return {int(r.Year): 12/float(r.Months_Closed) for _,r in cal.iterrows()}, {int(r.Year): 12/float(r.Months_Closed) for _,r in cal.iterrows()}, "Annual Projection: YTD disetahunkan memakai faktor 12 / Months_Closed.", m
    # MCR Focus: annualize revenue, keep cost actual
    return {int(r.Year): 12/float(r.Months_Closed) for _,r in cal.iterrows()}, {int(r.Year): 1.0 for _,r in cal.iterrows()}, "MCR Focus: revenue diproyeksikan full-year, manpower cost tetap actual/YTD untuk stress-test MCR.", m


def compute_data(sheets, mode):
    dim = sheets["DIM_DEPARTMENT"].copy()
    rev = sheets["REVENUE_YR"].copy()
    hc = sheets["HEADCOUNT_YR"].copy()
    pay = sheets["PAYROLL_YR"].copy()
    for df in [rev, hc, pay]:
        df["Year"] = pd.to_numeric(df["Year"], errors="coerce").astype("Int64")

    all_years = sorted(set(pd.concat([rev["Year"], hc["Year"], pay["Year"]]).dropna().astype(int).tolist()))
    cal = calendar(sheets.get("CALENDAR", pd.DataFrame()), all_years)
    rev_fac, cost_fac, note, comp_m = mode_factor(cal, mode)

    for c in ["Revenue_Recognized","COGS_Direct"]:
        rev[c] = pd.to_numeric(rev.get(c), errors="coerce").fillna(0)
    rev["Revenue_Actual"] = rev["Revenue_Recognized"]
    rev["COGS_Actual"] = rev["COGS_Direct"]
    rev["Revenue_Recognized"] = rev["Revenue_Recognized"] * rev["Year"].astype(float).map(rev_fac).fillna(1)
    rev["COGS_Direct"] = rev["COGS_Direct"] * rev["Year"].astype(float).map(cost_fac).fillna(1)

    hc["Dept_ID"] = hc["Dept_ID"].astype(str).str.strip()
    hc["Avg_Headcount"] = pd.to_numeric(hc.get("Avg_Headcount"), errors="coerce").fillna(0)
    hc["New_Hires"] = pd.to_numeric(hc.get("New_Hires", 0), errors="coerce").fillna(0)
    hc["Exits"] = pd.to_numeric(hc.get("Exits", 0), errors="coerce").fillna(0)

    pay["Dept_ID"] = pay["Dept_ID"].astype(str).str.strip()
    for c in ["Payroll_Gross","Overtime","Bonus","Benefits","Employer_Tax","Total_Manpower_Cost"]:
        if c not in pay.columns:
            pay[c] = 0
        pay[c] = pd.to_numeric(pay[c], errors="coerce").fillna(0)
        pay[c] = pay[c] * pay["Year"].astype(float).map(cost_fac).fillna(1)

    dim["Dept_ID"] = dim["Dept_ID"].astype(str).str.strip()
    if "FJA_Category" not in dim.columns:
        dim["FJA_Category"] = dim.get("Function_Group", "Unmapped")
    if "Dept_Name" not in dim.columns:
        dim["Dept_Name"] = dim["Dept_ID"]
    dim_small = dim[["Dept_ID","Dept_Name","Function_Group","FJA_Category"]].copy()

    rev_tot = rev.groupby("Year", as_index=False).agg(Total_Revenue=("Revenue_Recognized","sum"), Total_COGS=("COGS_Direct","sum"))
    rev_tot["Gross_Profit"] = rev_tot["Total_Revenue"] - rev_tot["Total_COGS"]
    rev_tot["Gross_Margin_Pct"] = np.where(rev_tot["Total_Revenue"]>0, rev_tot["Gross_Profit"]/rev_tot["Total_Revenue"], np.nan)

    hc2 = hc.merge(dim_small, on="Dept_ID", how="left")
    pay2 = pay.merge(dim_small, on="Dept_ID", how="left")
    hc_tot = hc2.groupby("Year", as_index=False).agg(Total_Headcount=("Avg_Headcount","sum"), New_Hires=("New_Hires","sum"), Exits=("Exits","sum"))
    pay_tot = pay2.groupby("Year", as_index=False).agg(Total_Manpower_Cost=("Total_Manpower_Cost","sum"))
    yr = rev_tot.merge(hc_tot, on="Year", how="outer").merge(pay_tot, on="Year", how="outer").sort_values("Year")
    yr["RPE"] = np.where(yr["Total_Headcount"]>0, yr["Total_Revenue"]/yr["Total_Headcount"], np.nan)
    yr["Cost_per_HC"] = np.where(yr["Total_Headcount"]>0, yr["Total_Manpower_Cost"]/yr["Total_Headcount"], np.nan)
    yr["MCR_Pct"] = np.where(yr["Total_Revenue"]>0, yr["Total_Manpower_Cost"]/yr["Total_Revenue"], np.nan)

    for c, out in [("Total_Revenue","Revenue_YoY"),("Total_Headcount","Headcount_YoY"),("RPE","RPE_YoY"),("Total_Manpower_Cost","ManpowerCost_YoY")]:
        yr[out] = yr[c].pct_change()
    yr["GrossMargin_Delta"] = yr["Gross_Margin_Pct"].diff()
    yr["MCR_Delta"] = yr["MCR_Pct"].diff()

    fja = pay2.groupby(["Year","FJA_Category"], as_index=False).agg(Manpower_Cost=("Total_Manpower_Cost","sum")).merge(
        hc2.groupby(["Year","FJA_Category"], as_index=False).agg(Headcount=("Avg_Headcount","sum")), on=["Year","FJA_Category"], how="outer"
    )
    fja["Cost_Share"] = fja["Manpower_Cost"] / fja.groupby("Year")["Manpower_Cost"].transform("sum")
    fja["HC_Share"] = fja["Headcount"] / fja.groupby("Year")["Headcount"].transform("sum")

    dept = hc2.groupby(["Year","Dept_ID","Dept_Name","Function_Group","FJA_Category"], as_index=False).agg(Headcount=("Avg_Headcount","sum")).merge(
        pay2.groupby(["Year","Dept_ID","Dept_Name","Function_Group","FJA_Category"], as_index=False).agg(Manpower_Cost=("Total_Manpower_Cost","sum")),
        on=["Year","Dept_ID","Dept_Name","Function_Group","FJA_Category"], how="outer"
    )
    if "REVENUE_BY_DEPT" in sheets and not sheets["REVENUE_BY_DEPT"].empty:
        rd = sheets["REVENUE_BY_DEPT"].copy()
        rd["Year"] = pd.to_numeric(rd["Year"], errors="coerce").astype("Int64")
        rd["Dept_ID"] = rd["Dept_ID"].astype(str).str.strip()
        rd["Revenue_Recognized"] = pd.to_numeric(rd["Revenue_Recognized"], errors="coerce").fillna(0) * rd["Year"].astype(float).map(rev_fac).fillna(1)
        dept_rev = rd.groupby(["Year","Dept_ID"], as_index=False).agg(Dept_Revenue=("Revenue_Recognized","sum"))
        dept = dept.merge(dept_rev, on=["Year","Dept_ID"], how="left")
    else:
        dept["Dept_Revenue"] = np.nan
    dept["Revenue_per_Cost"] = np.where(dept["Manpower_Cost"]>0, dept["Dept_Revenue"]/dept["Manpower_Cost"], np.nan)
    dept["Dept_RPE"] = np.where(dept["Headcount"]>0, dept["Dept_Revenue"]/dept["Headcount"], np.nan)

    cap = sheets.get("CAPACITY_INDICATOR", pd.DataFrame()).copy()
    if not cap.empty:
        cap["Year"] = pd.to_numeric(cap["Year"], errors="coerce").astype("Int64")
        for c in ["Avg_Utilization_Pct","Overtime_Hours","Backlog_Count","SLA_Breach_Count","Turnover_Count"]:
            cap[c] = pd.to_numeric(cap.get(c, 0), errors="coerce").fillna(0)
        if "Critical_Role_Dependency" not in cap.columns:
            cap["Critical_Role_Dependency"] = "Medium"

    pp = sheets.get("PROJECT_PRODUCTIVITY", pd.DataFrame()).copy()
    if not pp.empty:
        pp["Year"] = pd.to_numeric(pp["Year"], errors="coerce").astype("Int64")
        for c in ["Revenue","COGS","Manpower_Cost","Project_Hours","HC_Allocated","Sites_or_Tickets","SLA_Breach_Count"]:
            pp[c] = pd.to_numeric(pp.get(c,0), errors="coerce").fillna(0)
        pp["Revenue"] = pp["Revenue"] * pp["Year"].astype(float).map(rev_fac).fillna(1)
        pp["COGS"] = pp["COGS"] * pp["Year"].astype(float).map(cost_fac).fillna(1)
        pp["Manpower_Cost"] = pp["Manpower_Cost"] * pp["Year"].astype(float).map(cost_fac).fillna(1)
        pp["Gross_Profit"] = pp["Revenue"] - pp["COGS"]
        pp["Gross_Margin_Pct"] = np.where(pp["Revenue"]>0, pp["Gross_Profit"]/pp["Revenue"], np.nan)
        pp["MCR_Pct"] = np.where(pp["Revenue"]>0, pp["Manpower_Cost"]/pp["Revenue"], np.nan)
        pp["Revenue_per_HC"] = np.where(pp["HC_Allocated"]>0, pp["Revenue"]/pp["HC_Allocated"], np.nan)

    return dict(yearly=yr, fja=fja, dept=dept, capacity=cap, project_productivity=pp, calendar=cal, mode_note=note, comparable_months=comp_m)


# =========================
# STATUS ENGINES
# =========================
def margin_status(gm, tf):
    strong = tf_val(tf, "Gross Margin", "Strong_Min", .30)
    healthy = tf_val(tf, "Gross Margin", "Healthy_Min", .20)
    watch = tf_val(tf, "Gross Margin", "Watch_Min", .15)
    pressure = tf_val(tf, "Gross Margin", "Pressure_Min", .10)
    if pd.isna(gm): return "UNKNOWN", "info"
    if gm >= strong: return "STRONG", "success"
    if gm >= healthy: return "HEALTHY", "success"
    if gm >= watch: return "WATCH", "warning"
    if gm >= pressure: return "PRESSURE", "warning"
    return "CRITICAL", "error"


def mcr_status(mcr, tf):
    h_min = tf_val(tf, "MCR", "Healthy_Min", .05)
    h_max = tf_val(tf, "MCR", "Healthy_Max", .07)
    watch = tf_val(tf, "MCR", "Watch_Max", .09)
    pressure = tf_val(tf, "MCR", "Pressure_Max", .14)
    if pd.isna(mcr): return "UNKNOWN", "info"
    if mcr < h_min: return "ULTRA_EFFICIENCY", "warning"
    if mcr <= h_max: return "HEALTHY", "success"
    if mcr <= watch: return "WATCH", "warning"
    if mcr <= pressure: return "COST_PRESSURE", "warning"
    return "CRITICAL", "error"


def capacity_status(row, cap):
    year = int(row["Year"])
    if cap is None or cap.empty:
        return dict(score=0, state="LOW", tone="success", details={}, recommendation="No Action")
    c = cap[cap["Year"] == year].copy()
    if c.empty:
        return dict(score=0, state="LOW", tone="success", details={}, recommendation="No Action")
    max_util = c["Avg_Utilization_Pct"].max()
    overtime = c["Overtime_Hours"].sum()
    backlog = c["Backlog_Count"].sum()
    sla = c["SLA_Breach_Count"].sum()
    turnover = c["Turnover_Count"].sum()
    score = 0
    score += 25 if max_util >= .90 else 18 if max_util >= .85 else 10 if max_util >= .75 else 0
    score += 20 if overtime > 400 else 15 if overtime > 250 else 8 if overtime > 100 else 0
    score += 20 if backlog >= 20 else 12 if backlog >= 10 else 6 if backlog > 0 else 0
    score += 15 if sla >= 5 else 8 if sla > 0 else 0
    score += 10 if turnover >= 10 else 5 if turnover > 0 else 0
    score = min(100, int(score))
    if score >= 76: state, tone, rec = "CRITICAL", "error", "Immediate Capacity Action Required"
    elif score >= 56: state, tone, rec = "HIGH", "warning", "Capacity Expansion Review"
    elif score >= 31: state, tone, rec = "MEDIUM", "warning", "Selective Capacity Action"
    else: state, tone, rec = "LOW", "success", "No Hiring Required"
    return dict(score=score, state=state, tone=tone, recommendation=rec, details=dict(max_util=max_util, overtime=overtime, backlog=backlog, sla=sla, turnover=turnover))


def business_posture(row, tf, mcr_state):
    rev_g = row.get("Revenue_YoY")
    rpe_g = row.get("RPE_YoY")
    hc_g = row.get("Headcount_YoY")
    rev_strong = tf_val(tf, "Revenue Growth", "Strong_Min", .25)
    rpe_strong = tf_val(tf, "RPE Growth", "Strong_Min", .15)
    hc_max = tf_val(tf, "Headcount Growth", "Healthy_Max", .10)
    if pd.notna(rev_g) and rev_g <= 0:
        return "DEFENSIVE", "error"
    if pd.notna(rev_g) and pd.notna(rpe_g) and pd.notna(hc_g) and rev_g >= rev_strong and rpe_g >= rpe_strong and hc_g <= hc_max and mcr_state in ["HEALTHY", "ULTRA_EFFICIENCY", "WATCH"]:
        return "HIGH_LEVERAGE", "success"
    if mcr_state in ["COST_PRESSURE", "CRITICAL"]:
        return "COST_PRESSURE", "warning"
    return "BALANCED", "info"


def board_score(row, margin_state, mcr_state, cap_state):
    def score_growth(v, strong=.25, good=.10):
        if pd.isna(v): return 50
        if v >= strong: return 100
        if v >= good: return 80
        if v >= 0: return 60
        if v >= -.10: return 35
        return 15
    margin_map = {"STRONG":100,"HEALTHY":85,"WATCH":60,"PRESSURE":35,"CRITICAL":10,"UNKNOWN":50}
    mcr_map = {"HEALTHY":95,"WATCH":75,"ULTRA_EFFICIENCY":70,"COST_PRESSURE":40,"CRITICAL":10,"UNKNOWN":50}
    cap_map = {"LOW":100,"MEDIUM":75,"HIGH":40,"CRITICAL":10}
    comps = {
        "Revenue Health": score_growth(row.get("Revenue_YoY")),
        "Margin Health": margin_map.get(margin_state,50),
        "People Health": score_growth(row.get("RPE_YoY"), .15, .08),
        "Capacity Health": cap_map.get(cap_state,50),
        "Risk Health": mcr_map.get(mcr_state,50),
    }
    score = round(comps["Revenue Health"]*.25 + comps["Margin Health"]*.2 + comps["People Health"]*.25 + comps["Capacity Health"]*.15 + comps["Risk Health"]*.15)
    if score >= 85: label, tone = "EXCELLENT", "success"
    elif score >= 70: label, tone = "HEALTHY", "success"
    elif score >= 55: label, tone = "WATCH", "warning"
    elif score >= 40: label, tone = "PRESSURE", "warning"
    else: label, tone = "DEFENSIVE MODE", "error"
    return int(score), label, tone, comps


def action_recommendation(bp, margin_state, mcr_state, cap):
    if bp == "HIGH_LEVERAGE":
        title = "Selective Scale-Up with Capacity Guardrail"
        actions = ["Tambah HC hanya pada bottleneck revenue/delivery.", "Pertahankan MCR dengan approval berbasis ROI.", "Scale-up service line recurring revenue."]
    elif bp == "DEFENSIVE":
        title = "Margin & Cash Protection"
        actions = ["Freeze hiring non-critical.", "Review pricing, scope creep, COGS, dan subcontractor.", "Fokus recovery revenue dan margin."]
    elif mcr_state in ["COST_PRESSURE", "CRITICAL"]:
        title = "People-Cost Control"
        actions = ["Audit manpower cost dan overtime.", "Tahan support non-critical.", "Naikkan revenue target sebelum penambahan HC."]
    else:
        title = "Balanced Optimization"
        actions = ["Optimalkan struktur sebelum ekspansi.", "Hiring selektif pada fungsi bottleneck.", "Monitor MCR, RPE, capacity, dan margin bulanan."]
    if margin_state in ["WATCH", "PRESSURE", "CRITICAL"]:
        actions.append("Lakukan margin quality review.")
    if cap["state"] in ["HIGH","CRITICAL"]:
        actions.append("Immediate capacity action pada fungsi overload.")
    return title, actions


# =========================
# UI HELPERS
# =========================
def logo_path():
    for p in ["assets/logo_company.png","assets/starcom_logo.png","logo_company.png","starcom_logo.png"]:
        if os.path.exists(p): return p
    return ""


def header():
    c1, c2 = st.columns([6,1.2])
    with c1:
        st.markdown('<div class="sep-title">Starcom Executive Platform (SEP) V4</div><div class="sep-sub">Enterprise Build — Excel-driven, modular, board-ready.</div><div class="sep-desc">MCR, revenue leverage, workforce productivity, capacity risk, benchmark framework, and board decision.</div>', unsafe_allow_html=True)
    with c2:
        p = logo_path()
        if p: st.image(p, use_container_width=True)
    st.markdown("""
    <div class="feature-band">
      <div class="feature-item"><div class="feature-icon">🎯</div><div><div class="feature-title">Executive Decision Support</div><div class="feature-text">Insight untuk keputusan direksi.</div></div></div>
      <div class="feature-item"><div class="feature-icon">📈</div><div><div class="feature-title">Performance & Productivity</div><div class="feature-text">Revenue, MCR, margin, dan RPE.</div></div></div>
      <div class="feature-item"><div class="feature-icon">🛡️</div><div><div class="feature-title">Risk & Capacity Control</div><div class="feature-text">Risiko kapasitas dan delivery.</div></div></div>
      <div class="feature-item"><div class="feature-icon">👥</div><div><div class="feature-title">Benchmark Framework</div><div class="feature-text">Target 100% dari Excel.</div></div></div>
      <div class="feature-item"><div class="feature-icon">📄</div><div><div class="feature-title">Board Ready Report</div><div class="feature-text">Narasi siap presentasi.</div></div></div>
    </div>
    """, unsafe_allow_html=True)


def how(title, items):
    st.markdown("<div class='how-box'><b>ⓘ " + title + "</b><ul>" + "".join(f"<li>{i}</li>" for i in items) + "</ul></div>", unsafe_allow_html=True)


# =========================
# SIDEBAR
# =========================
with st.sidebar:
    st.header("⚙️ Data Source")
    uploaded = st.file_uploader("Upload Excel SEP V4", type=["xlsx"])
    st.caption("Jika tidak upload, app membaca file default dari folder data/.")

    st.header("🧮 Analysis Mode")
    mode = st.radio("Mode", ["Actual","YTD Comparable","Annual Projection","MCR Focus"], index=1)

    st.header("🖥️ Presentation Mode")
    presentation_mode = st.radio("Display Size", ["Smartphone","Compact","Executive","Boardroom"], index=2)
    st.caption("Smartphone=HP | Compact=Laptop | Executive=Meeting | Boardroom=TV/Proyektor")

scale = get_presentation_scale(presentation_mode)
inject_css(scale)
if presentation_mode == "Smartphone":
    st.sidebar.info("Smartphone Mode aktif: layout responsive untuk HP.")

file_bytes = uploaded.getvalue() if uploaded else None
sheets, warnings = load_workbook(file_bytes)

header()
for w in warnings:
    st.warning(w)
if any(s not in sheets for s in REQUIRED_SHEETS):
    st.stop()

params = parameters(sheets["PARAMETERS"])
currency = params.get("Currency_Code", "IDR")
tf = read_framework(sheets["TARGET_FRAMEWORK"])
data = compute_data(sheets, mode)
yr = data["yearly"].dropna(subset=["Year"]).copy()
years = sorted(yr["Year"].astype(int).tolist())

with st.sidebar:
    st.header("🧭 Filters")
    year = st.selectbox("Analysis Year", years, index=len(years)-1)
    rng = st.slider("Trend Range", min_value=min(years), max_value=max(years), value=(min(years), max(years)), step=1)
    fja_all = sorted(data["fja"]["FJA_Category"].dropna().unique().tolist())
    fja_filter = st.multiselect("FJA Category", fja_all, default=fja_all)

yr_range = yr[(yr["Year"]>=rng[0]) & (yr["Year"]<=rng[1])]
latest = yr[yr["Year"] == year].iloc[0]
period_row = data["calendar"][data["calendar"]["Year"] == year]
period_label = period_row["Period_Label"].iloc[0] if not period_row.empty else f"FY {year}"
months_closed = int(period_row["Months_Closed"].iloc[0]) if not period_row.empty else 12
confidence = period_row["Data_Confidence"].iloc[0] if not period_row.empty else "High"

m_state, m_tone = margin_status(latest.get("Gross_Margin_Pct"), tf)
mcr_state, mcr_tone = mcr_status(latest.get("MCR_Pct"), tf)
cap = capacity_status(latest, data["capacity"])
bp_state, bp_tone = business_posture(latest, tf, mcr_state)
score, score_label, score_tone, score_comps = board_score(latest, m_state, mcr_state, cap["state"])
action_title, action_items = action_recommendation(bp_state, m_state, mcr_state, cap)

plot_config = {"displaylogo": False, "modeBarButtonsToRemove": ["lasso2d","select2d"]}

boards = st.tabs([
    "1. Executive Summary", "2. Revenue & Productivity", "3. Workforce Performance",
    "4. Capacity & Delivery Risk", "5. Strategic Decision Center",
    "6. KPI Target Framework", "7. CEO Closing Report", "8. Appendix"
])

# Board 1
with boards[0]:
    st.markdown(f"## KPI Snapshot ({period_label})")
    st.caption(f"Periode Data: {period_label} | Bulan Tertutup: {months_closed}/12 | Mode: {mode} | Confidence: {confidence}")
    st.info(data["mode_note"])
    left, right = st.columns([1.85,1.0])
    with left:
        c = st.columns(4)
        c[0].metric("💰 Revenue", money(latest["Total_Revenue"], currency), pct(latest.get("Revenue_YoY")))
        c[1].metric("📊 MCR", pct(latest["MCR_Pct"]), pp(latest.get("MCR_Delta")))
        c[2].metric("📈 Gross Margin", pct(latest["Gross_Margin_Pct"]), pp(latest.get("GrossMargin_Delta")))
        c[3].metric("⚡ RPE", money(latest["RPE"], currency), pct(latest.get("RPE_YoY")))
        c = st.columns(4)
        c[0].metric("👤 Headcount", num(latest["Total_Headcount"]), pct(latest.get("Headcount_YoY")))
        c[1].metric("💼 Manpower Cost", money(latest["Total_Manpower_Cost"], currency), pct(latest.get("ManpowerCost_YoY")))
        c[2].metric("🛡️ Business Posture", bp_state)
        c[3].metric("🏅 Margin Quality", m_state)
        st.markdown(f"<div class='memo-card'><b>Board Memo.</b> Pada {period_label}, SEP membaca revenue <b>{money(latest['Total_Revenue'], currency)}</b>, MCR <b>{pct(latest['MCR_Pct'])}</b>, Gross Margin <b>{pct(latest['Gross_Margin_Pct'])}</b>, dan RPE <b>{money(latest['RPE'], currency)}</b>. Status utama adalah <b>{score_label}</b>. Fokus manajemen: <b>{action_title}</b>.</div>", unsafe_allow_html=True)
    with right:
        st.markdown("### Company Health Index ⓘ")
        fig = go.Figure(go.Indicator(mode="gauge+number", value=score, number={"suffix":"/100"}, title={"text":score_label}, gauge={"axis":{"range":[0,100]}, "bar":{"color":TONE_COLOR[score_tone]}, "steps":[{"range":[0,40],"color":"#FEE2E2"},{"range":[40,70],"color":"#FEF3C7"},{"range":[70,85],"color":"#DBEAFE"},{"range":[85,100],"color":"#DCFCE7"}]}))
        st.plotly_chart(style_plot(fig, scale, "gauge"), use_container_width=True, config=plot_config)
        for k,v in score_comps.items():
            st.progress(int(v), text=f"{k}: {int(v)}/100")

# Board 2
with boards[1]:
    st.markdown("## 2. Revenue & Productivity")
    st.caption("Pertanyaan bisnis: apakah revenue tumbuh bersama produktivitas?")
    c1, c2 = st.columns(2)
    with c1:
        fig = px.line(yr_range, x="Year", y=["Total_Revenue","Total_Manpower_Cost"], markers=True, title="Revenue vs Manpower Cost")
        st.plotly_chart(style_plot(fig, scale), use_container_width=True, config=plot_config)
    with c2:
        fig = px.line(yr_range, x="Year", y=["RPE","Cost_per_HC"], markers=True, title="RPE vs Cost per HC")
        st.plotly_chart(style_plot(fig, scale), use_container_width=True, config=plot_config)

    ppdf = data["project_productivity"]
    if not ppdf.empty:
        st.markdown("### Project / Service Line Productivity")
        ppy = ppdf[ppdf["Year"] == year].copy()
        by_cat = ppy.groupby("Project_Category", as_index=False).agg(Revenue=("Revenue","sum"), Gross_Profit=("Gross_Profit","sum"), Manpower_Cost=("Manpower_Cost","sum"), HC_Allocated=("HC_Allocated","sum"), SLA_Breach_Count=("SLA_Breach_Count","sum"))
        by_cat["Revenue_Share"] = by_cat["Revenue"] / by_cat["Revenue"].sum()
        by_cat["Gross_Margin_Pct"] = np.where(by_cat["Revenue"]>0, by_cat["Gross_Profit"]/by_cat["Revenue"], np.nan)
        by_cat["MCR_Pct"] = np.where(by_cat["Revenue"]>0, by_cat["Manpower_Cost"]/by_cat["Revenue"], np.nan)
        p1, p2 = st.columns(2)
        with p1:
            fig = px.treemap(by_cat, path=["Project_Category"], values="Revenue", color="Gross_Margin_Pct", color_continuous_scale="RdYlGn", title="Revenue Contribution by Service Line")
            st.plotly_chart(style_plot(fig, scale), use_container_width=True, config=plot_config)
        with p2:
            fig = px.scatter(by_cat, x="MCR_Pct", y="Gross_Margin_Pct", size="Revenue", color="Project_Category", text="Project_Category", title="Project Productivity Matrix", size_max=scale["bubble"])
            fig.update_xaxes(tickformat=".0%")
            fig.update_yaxes(tickformat=".0%")
            st.plotly_chart(style_plot(fig, scale), use_container_width=True, config=plot_config)

# Board 3
with boards[2]:
    st.markdown("## 3. Workforce Performance")
    dept = data["dept"]
    dept = dept[(dept["Year"] == year) & (dept["FJA_Category"].isin(fja_filter))].copy()
    if not dept.empty:
        cols = st.columns([0.75, 1.8, 0.65])
        with cols[0]:
            how("How to Read", ["Setiap gelembung = department/FJA.", "Sumbu X = Headcount.", "Sumbu Y = Manpower Cost.", "Ukuran gelembung = Dept Revenue atau Headcount.", "Merah/support perlu dijaga agar tidak membebani MCR."])
        with cols[1]:
            dept["Bubble_Size"] = dept["Dept_Revenue"].fillna(0)
            if dept["Bubble_Size"].max() <= 0:
                dept["Bubble_Size"] = dept["Headcount"].fillna(1).clip(lower=1)
            colors = {"Revenue Generator":"#2563EB","Revenue Enabler":"#16A34A","Support Function":"#EF4444","Governance / Management":"#8B5CF6"}
            fig = px.scatter(dept, x="Headcount", y="Manpower_Cost", size="Bubble_Size", color="FJA_Category", text="FJA_Category", color_discrete_map=colors, hover_name="Dept_Name", size_max=scale["bubble"], title="Workforce Portfolio Matrix")
            fig.update_traces(textposition="middle center", textfont=dict(size=max(12,scale["plot_font"]), color="white"), marker=dict(opacity=.92, line=dict(width=2.5, color="white")))
            st.plotly_chart(style_plot(fig, scale), use_container_width=True, config=plot_config)
        with cols[2]:
            st.markdown("<div class='insight-card'><b>💡 Insight</b><br><br>Fokus menjaga produktivitas fungsi support dan memprioritaskan capacity pada revenue generator/enabler.</div>", unsafe_allow_html=True)

# Board 4
with boards[3]:
    st.markdown("## 4. Capacity & Delivery Risk")
    capy = data["capacity"][data["capacity"]["Year"] == year].copy() if not data["capacity"].empty else pd.DataFrame()
    c0, c1, c2 = st.columns([0.72,1.45,0.78])
    with c0:
        how("How to Read", ["Setiap gelembung = department.", "Sumbu X = Avg Utilization.", "Sumbu Y = Overtime Hours.", "Ukuran gelembung = Backlog.", "Kanan-atas = risiko overload tertinggi."])
    with c1:
        if not capy.empty:
            capy["Backlog_Visual_Size"] = capy["Backlog_Count"].clip(lower=8)
            colors = {"Low":"#EF4444","Medium":"#2563EB","High":"#16A34A","Critical":"#B91C1C"}
            fig = px.scatter(capy, x="Avg_Utilization_Pct", y="Overtime_Hours", size="Backlog_Visual_Size", color="Critical_Role_Dependency", color_discrete_map=colors, hover_data=["Dept_ID","Backlog_Count","SLA_Breach_Count","Turnover_Count"], size_max=scale["cap_bubble"], title="Capacity Risk Heatmap")
            fig.update_xaxes(tickformat=".0%")
            fig.update_traces(marker=dict(opacity=.9, line=dict(width=2.5, color="white")))
            st.plotly_chart(style_plot(fig, scale), use_container_width=True, config=plot_config)
        else:
            st.info("CAPACITY_INDICATOR belum diisi.")
    with c2:
        fig = go.Figure(go.Indicator(mode="gauge+number", value=cap["score"], number={"suffix":"/100"}, title={"text":f"Capacity Score - {cap['state']}"}, gauge={"axis":{"range":[0,100]}, "bar":{"color":TONE_COLOR[cap["tone"]]}, "steps":[{"range":[0,30],"color":"#DCFCE7"},{"range":[31,55],"color":"#FEF3C7"},{"range":[56,75],"color":"#FFEDD5"},{"range":[76,100],"color":"#FEE2E2"}]}))
        st.plotly_chart(style_plot(fig, scale, "gauge"), use_container_width=True, config=plot_config)
        d = cap["details"]
        st.write(f"**Max Utilization:** {pct(d.get('max_util'))}")
        st.write(f"**Overtime:** {num(d.get('overtime'))} | **Backlog:** {num(d.get('backlog'))}")
        st.error(cap["recommendation"]) if cap["state"] == "CRITICAL" else st.info(cap["recommendation"])

# Board 5
with boards[4]:
    st.markdown("## 5. Strategic Decision Center")
    st.caption("Pertanyaan bisnis: keputusan apa yang harus diambil?")
    c1, c2 = st.columns(2)
    priority = pd.DataFrame([
        {"Area":"Business Posture", "Status":bp_state, "Priority": 90 if bp_state in ["DEFENSIVE","COST_PRESSURE"] else 35},
        {"Area":"Margin Quality", "Status":m_state, "Priority": 90 if m_state in ["CRITICAL","PRESSURE"] else 65 if m_state=="WATCH" else 25},
        {"Area":"MCR Health", "Status":mcr_state, "Priority": 90 if mcr_state=="CRITICAL" else 65 if mcr_state in ["COST_PRESSURE","ULTRA_EFFICIENCY"] else 25},
        {"Area":"Capacity Risk", "Status":cap["state"], "Priority": 90 if cap["state"]=="CRITICAL" else 65 if cap["state"]=="HIGH" else 25},
    ])
    with c1:
        fig = px.bar(priority.sort_values("Priority"), x="Priority", y="Area", color="Status", orientation="h", text="Status", title="Management Attention Map")
        st.plotly_chart(style_plot(fig, scale), use_container_width=True, config=plot_config)
    with c2:
        st.subheader(action_title)
        for i, a in enumerate(action_items, 1):
            st.write(f"{i}. {a}")

    st.markdown("### Board Projection Simulator")
    add_hc = st.slider("Additional HC", 0, 100, 5)
    rev_growth = st.slider("Revenue Growth Target", -0.30, 1.50, 0.20, step=.01)
    gm_target = st.slider("Gross Margin Target", 0.00, .80, float(latest["Gross_Margin_Pct"]) if pd.notna(latest["Gross_Margin_Pct"]) else .25, step=.01)
    avg_cost = st.number_input("Avg Cost / New HC", min_value=0.0, value=float(latest["Cost_per_HC"]) if pd.notna(latest["Cost_per_HC"]) else 75_000_000.0, step=5_000_000.0)
    proj_rev = latest["Total_Revenue"] * (1 + rev_growth)
    proj_cost = latest["Total_Manpower_Cost"] + add_hc * avg_cost
    proj_hc = latest["Total_Headcount"] + add_hc
    proj_mcr = proj_cost / proj_rev if proj_rev else np.nan
    proj_rpe = proj_rev / proj_hc if proj_hc else np.nan
    target_mcr = tf_val(tf, "MCR", "Healthy_Max", .07)
    req_rev = proj_cost / target_mcr if target_mcr else np.nan
    c = st.columns(5)
    c[0].metric("Projected Revenue", money(proj_rev, currency))
    c[1].metric("Projected HC", num(proj_hc), f"+{add_hc}")
    c[2].metric("Projected MCR", pct(proj_mcr), pp(proj_mcr - latest["MCR_Pct"]))
    c[3].metric("Projected GM", pct(gm_target), pp(gm_target - latest["Gross_Margin_Pct"]))
    c[4].metric("Revenue Required", money(req_rev, currency))

# Board 6
with boards[5]:
    st.markdown("## 6. KPI Target Framework")
    st.markdown("### Based on Telecommunications & IT System Integrator Benchmark")
    st.caption("Target di halaman ini dibaca dari sheet TARGET_FRAMEWORK di Excel.")
    display = tf.copy()
    if not display.empty:
        cols = [c for c in ["KPI","Target_Framework","Benchmark_Basis","Why_It_Matters","Confidence"] if c in display.columns]
        st.dataframe(display[cols], use_container_width=True, hide_index=True)
    current = pd.DataFrame([
        {"KPI":"Revenue Growth", "Company":pct(latest.get("Revenue_YoY")), "Framework":target_text(tf,"Revenue Growth","15–25% healthy; >25% strong"), "Status":"Based on framework"},
        {"KPI":"RPE Growth", "Company":pct(latest.get("RPE_YoY")), "Framework":target_text(tf,"RPE Growth",">=15% strong"), "Status":"Based on framework"},
        {"KPI":"MCR", "Company":pct(latest.get("MCR_Pct")), "Framework":target_text(tf,"MCR","5–7% healthy; 7–9% watch"), "Status":mcr_state},
        {"KPI":"Gross Margin", "Company":pct(latest.get("Gross_Margin_Pct")), "Framework":target_text(tf,"Gross Margin","20–30% healthy; >30% strong"), "Status":m_state},
        {"KPI":"Capacity Score", "Company":f"{cap['score']}/100", "Framework":target_text(tf,"Capacity Score","0–30 low; 31–55 medium; 56–75 high; 76–100 critical"), "Status":cap["state"]},
    ])
    st.markdown("### Current KPI vs Framework")
    st.dataframe(current, use_container_width=True, hide_index=True)
    st.info("Untuk mengubah target, cukup edit sheet TARGET_FRAMEWORK di Excel. App.py tidak perlu diubah.")

# Board 7
with boards[6]:
    st.markdown("## 7. CEO Closing Report")
    st.info(f"SEP membaca kondisi perusahaan pada status **{score_label}** dengan Company Health Index **{score}/100**. Revenue **{money(latest['Total_Revenue'], currency)}**, MCR **{pct(latest['MCR_Pct'])}**, Gross Margin **{pct(latest['Gross_Margin_Pct'])}**, RPE **{money(latest['RPE'], currency)}**, dan Capacity Risk **{cap['state']}**.")
    st.markdown("### Top Decisions")
    for i, a in enumerate(action_items, 1):
        st.write(f"{i}. {a}")
    roadmap = pd.DataFrame([
        {"Phase":"0–3 Months", "Priority":"High", "Action":action_items[0] if action_items else "-", "Owner":"CEO / CFO / HR"},
        {"Phase":"3–6 Months", "Priority":"High", "Action":"Review service-line profitability and capacity bottleneck.", "Owner":"COO / Finance"},
        {"Phase":"6–12 Months", "Priority":"Medium", "Action":"Strengthen recurring revenue and optimize support ratio.", "Owner":"Sales / HR"},
    ])
    st.dataframe(roadmap, use_container_width=True, hide_index=True)

# Board 8
with boards[7]:
    st.markdown("## 8. Appendix")
    st.markdown("""
    ### Core Formula
    - **MCR** = Total Manpower Cost / Revenue
    - **RPE** = Revenue / Average Headcount
    - **Gross Margin** = (Revenue - COGS) / Revenue
    - **Projected MCR** = (Current Manpower Cost + Additional HC × Avg Cost per HC) / Projected Revenue

    ### Governance
    - **Owner:** Finance + HR + Operations
    - **Review cycle:** Monthly / Quarterly
    - **Threshold master:** `TARGET_FRAMEWORK` sheet
    - **Version:** SEP V4 Enterprise Build
    """)
    st.dataframe(yr, use_container_width=True, hide_index=True)

st.caption("Starcom Executive Platform (SEP) V4 Enterprise Build — Excel-driven, modular, board-ready.")
