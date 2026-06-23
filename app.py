
# app.py
# HR vs Revenue Intelligence Platform v2 — Board & CFO Edition
# Streamlit dashboard based on Excel data model:
# HR_vs_Revenue_Intelligence_Platform_v2_Board_CFO_Template_2026.xlsx

import os
from io import BytesIO
from typing import Dict, Optional, Tuple, List

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


# =========================
# APP CONFIG
# =========================
st.set_page_config(
    page_title="HR vs Revenue Intelligence Platform v2",
    page_icon="📊",
    layout="wide",
)

DEFAULT_FILE = os.path.join(
    "data",
    "HR_vs_Revenue_Intelligence_Platform_v2_Board_CFO_Template_2026.xlsx",
)

REQUIRED_SHEETS = [
    "PARAMETERS",
    "TARGETS",
    "FJA_MAPPING",
    "DIM_DEPARTMENT",
    "REVENUE_YR",
    "HEADCOUNT_YR",
    "PAYROLL_YR",
]

OPTIONAL_SHEETS = [
    "DIM_EMPLOYEE",
    "REVENUE_BY_DEPT",
    "HC_MOVEMENT",
    "PROJECT_MARGIN",
    "CAPACITY_INDICATOR",
    "EXEC_SUMMARY",
]

DEFAULT_MCR_BANDS = {
    "ultra_max": 0.05,
    "healthy_min": 0.05,
    "healthy_max": 0.07,
    "watch_max": 0.09,
    "pressure_max": 0.14,
    "critical_min": 0.14,
}

DEFAULT_GM_BANDS = {
    "strong_min": 0.35,
    "healthy_min": 0.25,
    "watch_min": 0.15,
    "pressure_min": 0.10,
    "critical_max": 0.10,
}

TONE_COLOR = {
    "success": "#22c55e",
    "info": "#3b82f6",
    "warning": "#f59e0b",
    "error": "#ef4444",
    "neutral": "#64748b",
}


# =========================
# UTILITIES
# =========================
def _money(x, currency: str = "IDR") -> str:
    try:
        if x is None or pd.isna(x):
            return "-"
        x = float(x)
        if abs(x) >= 1_000_000_000:
            return f"{currency} {x/1_000_000_000:,.1f}B".replace(",", ".")
        if abs(x) >= 1_000_000:
            return f"{currency} {x/1_000_000:,.1f}M".replace(",", ".")
        return f"{currency} {x:,.0f}".replace(",", ".")
    except Exception:
        return "-"


def _num(x) -> str:
    try:
        if x is None or pd.isna(x):
            return "-"
        return f"{float(x):,.1f}".replace(",", ".")
    except Exception:
        return "-"


def _pct(x) -> str:
    try:
        if x is None or pd.isna(x):
            return "-"
        return f"{float(x)*100:.1f}%"
    except Exception:
        return "-"


def _pp(x) -> str:
    try:
        if x is None or pd.isna(x):
            return "-"
        return f"{float(x)*100:+.1f} pp"
    except Exception:
        return "-"


def _safe_float(x):
    try:
        if x is None or pd.isna(x):
            return np.nan
        return float(x)
    except Exception:
        return np.nan


def _safe_str(x) -> str:
    if x is None or pd.isna(x):
        return ""
    return str(x).strip()


def load_targets(targets: pd.DataFrame) -> Dict[str, float]:
    out = {}
    if targets is None or targets.empty:
        return out
    if not {"KPI", "Target_Value"}.issubset(set(targets.columns)):
        return out
    for _, row in targets.iterrows():
        k = _safe_str(row.get("KPI"))
        v = row.get("Target_Value")
        try:
            out[k] = float(v)
        except Exception:
            pass
    return out


def get_target(targets: Dict[str, float], key: str, default: float) -> float:
    return float(targets.get(key, default))


def build_params(params: pd.DataFrame) -> Dict[str, str]:
    out = {}
    if params is None or params.empty:
        return out
    if not {"Parameter", "Value"}.issubset(set(params.columns)):
        return out
    for _, row in params.iterrows():
        k = _safe_str(row.get("Parameter"))
        v = _safe_str(row.get("Value"))
        if k:
            out[k] = v
    return out


@st.cache_data(show_spinner=False)
def load_workbook(file_bytes: Optional[bytes], default_path: str) -> Tuple[Dict[str, pd.DataFrame], List[str]]:
    warnings: List[str] = []
    try:
        if file_bytes:
            xls = pd.ExcelFile(BytesIO(file_bytes))
        else:
            if not os.path.exists(default_path):
                return {}, [f"File default tidak ditemukan: {default_path}. Upload file Excel terlebih dahulu."]
            xls = pd.ExcelFile(default_path)
    except Exception as e:
        return {}, [f"Excel tidak bisa dibaca: {type(e).__name__}: {e}"]

    sheets: Dict[str, pd.DataFrame] = {}
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


def clean_and_calc(sheets: Dict[str, pd.DataFrame]) -> Dict[str, pd.DataFrame]:
    dim = sheets["DIM_DEPARTMENT"].copy()
    fja_map = sheets["FJA_MAPPING"].copy()
    rev = sheets["REVENUE_YR"].copy()
    hc = sheets["HEADCOUNT_YR"].copy()
    pay = sheets["PAYROLL_YR"].copy()

    for c in ["Dept_ID", "Dept_Name", "Function_Group", "Revenue_Driver_Flag", "FJA_Category_Override"]:
        if c in dim.columns:
            dim[c] = dim[c].astype(str).str.strip()
    dim = dim[dim["Dept_ID"].astype(str).str.strip().ne("")].copy()

    if "FJA_Category_Override" not in dim.columns:
        dim["FJA_Category_Override"] = ""

    if not fja_map.empty and {"Function_Group", "FJA_Category"}.issubset(fja_map.columns):
        fja_lookup = dict(zip(fja_map["Function_Group"].astype(str).str.strip(), fja_map["FJA_Category"].astype(str).str.strip()))
    else:
        fja_lookup = {
            "Sales": "Revenue Generator",
            "Operations/Project": "Revenue Enabler",
            "Engineering": "Revenue Enabler",
            "Support": "Support Function",
            "Management": "Governance / Management",
        }

    dim["FJA_Category"] = np.where(
        dim["FJA_Category_Override"].astype(str).str.strip().ne(""),
        dim["FJA_Category_Override"].astype(str).str.strip(),
        dim["Function_Group"].map(fja_lookup).fillna("Unmapped")
    )

    rev["Year"] = pd.to_numeric(rev["Year"], errors="coerce").astype("Int64")
    for c in ["Revenue_Recognized", "COGS_Direct"]:
        if c in rev.columns:
            rev[c] = pd.to_numeric(rev[c], errors="coerce").fillna(0.0)
    if "COGS_Direct" not in rev.columns:
        rev["COGS_Direct"] = 0.0

    hc["Year"] = pd.to_numeric(hc["Year"], errors="coerce").astype("Int64")
    hc["Avg_Headcount"] = pd.to_numeric(hc.get("Avg_Headcount"), errors="coerce").fillna(0.0)
    if "Avg_FTE" in hc.columns:
        hc["Avg_FTE"] = pd.to_numeric(hc["Avg_FTE"], errors="coerce")
    if "New_Hires" in hc.columns:
        hc["New_Hires"] = pd.to_numeric(hc["New_Hires"], errors="coerce").fillna(0.0)
    if "Exits" in hc.columns:
        hc["Exits"] = pd.to_numeric(hc["Exits"], errors="coerce").fillna(0.0)

    pay["Year"] = pd.to_numeric(pay["Year"], errors="coerce").astype("Int64")
    cost_cols = ["Payroll_Gross", "Overtime", "Bonus", "Benefits", "Employer_Tax", "Total_Manpower_Cost"]
    for c in cost_cols:
        if c in pay.columns:
            pay[c] = pd.to_numeric(pay[c], errors="coerce").fillna(0.0)
    if "Total_Manpower_Cost" not in pay.columns:
        comp_cols = [c for c in ["Payroll_Gross", "Overtime", "Bonus", "Benefits", "Employer_Tax"] if c in pay.columns]
        pay["Total_Manpower_Cost"] = pay[comp_cols].sum(axis=1) if comp_cols else 0.0

    dim_small = dim[["Dept_ID", "Dept_Name", "Function_Group", "Revenue_Driver_Flag", "FJA_Category"]].copy()

    hc2 = hc.merge(dim_small, on="Dept_ID", how="left")
    pay2 = pay.merge(dim_small, on="Dept_ID", how="left")

    # Yearly KPI
    rev_tot = rev.groupby("Year", as_index=False).agg(
        Total_Revenue=("Revenue_Recognized", "sum"),
        Total_COGS=("COGS_Direct", "sum"),
    )
    rev_tot["Gross_Profit"] = rev_tot["Total_Revenue"] - rev_tot["Total_COGS"]
    rev_tot["Gross_Margin_Pct"] = np.where(rev_tot["Total_Revenue"] > 0, rev_tot["Gross_Profit"] / rev_tot["Total_Revenue"], np.nan)

    hc_tot = hc2.groupby("Year", as_index=False).agg(
        Total_Headcount=("Avg_Headcount", "sum"),
        New_Hires=("New_Hires", "sum") if "New_Hires" in hc2.columns else ("Avg_Headcount", lambda x: np.nan),
        Exits=("Exits", "sum") if "Exits" in hc2.columns else ("Avg_Headcount", lambda x: np.nan),
    )

    pay_tot = pay2.groupby("Year", as_index=False).agg(
        Total_Manpower_Cost=("Total_Manpower_Cost", "sum")
    )

    yr = rev_tot.merge(hc_tot, on="Year", how="outer").merge(pay_tot, on="Year", how="outer").sort_values("Year")
    yr["RPE"] = np.where(yr["Total_Headcount"] > 0, yr["Total_Revenue"] / yr["Total_Headcount"], np.nan)
    yr["Cost_per_HC"] = np.where(yr["Total_Headcount"] > 0, yr["Total_Manpower_Cost"] / yr["Total_Headcount"], np.nan)
    yr["MCR_Pct"] = np.where(yr["Total_Revenue"] > 0, yr["Total_Manpower_Cost"] / yr["Total_Revenue"], np.nan)
    yr["Revenue_per_Payroll"] = np.where(yr["Total_Manpower_Cost"] > 0, yr["Total_Revenue"] / yr["Total_Manpower_Cost"], np.nan)

    for col, new_col in [
        ("Total_Revenue", "Revenue_YoY"),
        ("Total_Headcount", "Headcount_YoY"),
        ("Total_Manpower_Cost", "ManpowerCost_YoY"),
        ("RPE", "RPE_YoY"),
        ("Cost_per_HC", "Cost_per_HC_YoY"),
        ("Gross_Margin_Pct", "GrossMargin_Delta"),
        ("MCR_Pct", "MCR_Delta"),
    ]:
        if col in yr.columns:
            if "Delta" in new_col:
                yr[new_col] = yr[col].diff()
            else:
                yr[new_col] = yr[col].pct_change()

    # Function/FJA breakdown
    hc_fg = hc2.groupby(["Year", "Function_Group", "FJA_Category"], dropna=False, as_index=False).agg(Headcount=("Avg_Headcount", "sum"))
    pay_fg = pay2.groupby(["Year", "Function_Group", "FJA_Category"], dropna=False, as_index=False).agg(Manpower_Cost=("Total_Manpower_Cost", "sum"))
    fg = hc_fg.merge(pay_fg, on=["Year", "Function_Group", "FJA_Category"], how="outer")
    fg["Cost_per_HC"] = np.where(fg["Headcount"] > 0, fg["Manpower_Cost"] / fg["Headcount"], np.nan)

    fja = fg.groupby(["Year", "FJA_Category"], as_index=False).agg(
        Headcount=("Headcount", "sum"),
        Manpower_Cost=("Manpower_Cost", "sum"),
    )
    fja["Cost_Share"] = fja["Manpower_Cost"] / fja.groupby("Year")["Manpower_Cost"].transform("sum")
    fja["HC_Share"] = fja["Headcount"] / fja.groupby("Year")["Headcount"].transform("sum")

    # Department view
    hcd = hc2.groupby(["Year", "Dept_ID", "Dept_Name", "Function_Group", "FJA_Category"], dropna=False, as_index=False).agg(Headcount=("Avg_Headcount", "sum"))
    payd = pay2.groupby(["Year", "Dept_ID", "Dept_Name", "Function_Group", "FJA_Category"], dropna=False, as_index=False).agg(Manpower_Cost=("Total_Manpower_Cost", "sum"))
    dept = hcd.merge(payd, on=["Year", "Dept_ID", "Dept_Name", "Function_Group", "FJA_Category"], how="outer")
    dept["Cost_per_HC"] = np.where(dept["Headcount"] > 0, dept["Manpower_Cost"] / dept["Headcount"], np.nan)

    if "REVENUE_BY_DEPT" in sheets and not sheets["REVENUE_BY_DEPT"].empty:
        rd = sheets["REVENUE_BY_DEPT"].copy()
        rd["Year"] = pd.to_numeric(rd["Year"], errors="coerce").astype("Int64")
        rd["Revenue_Recognized"] = pd.to_numeric(rd["Revenue_Recognized"], errors="coerce").fillna(0.0)
        rd_tot = rd.groupby(["Year", "Dept_ID"], as_index=False).agg(Dept_Revenue=("Revenue_Recognized", "sum"))
        dept = dept.merge(rd_tot, on=["Year", "Dept_ID"], how="left")
    else:
        dept["Dept_Revenue"] = np.nan

    dept["Dept_RPE"] = np.where(dept["Headcount"] > 0, dept["Dept_Revenue"] / dept["Headcount"], np.nan)
    dept["Revenue_per_Cost"] = np.where(dept["Manpower_Cost"] > 0, dept["Dept_Revenue"] / dept["Manpower_Cost"], np.nan)

    # Capacity indicators
    if "CAPACITY_INDICATOR" in sheets and not sheets["CAPACITY_INDICATOR"].empty:
        cap = sheets["CAPACITY_INDICATOR"].copy()
        cap["Year"] = pd.to_numeric(cap["Year"], errors="coerce").astype("Int64")
        for c in ["Avg_Utilization_Pct", "Overtime_Hours", "Backlog_Count", "SLA_Breach_Count", "Incident_Count", "Turnover_Count"]:
            if c in cap.columns:
                cap[c] = pd.to_numeric(cap[c], errors="coerce").fillna(0.0)
    else:
        cap = pd.DataFrame()

    # Project margin
    if "PROJECT_MARGIN" in sheets and not sheets["PROJECT_MARGIN"].empty:
        pm = sheets["PROJECT_MARGIN"].copy()
        pm["Year"] = pd.to_numeric(pm["Year"], errors="coerce").astype("Int64")
        for c in ["Revenue", "COGS", "Gross_Profit"]:
            if c in pm.columns:
                pm[c] = pd.to_numeric(pm[c], errors="coerce").fillna(0.0)
        if "Gross_Profit" not in pm.columns:
            pm["Gross_Profit"] = pm["Revenue"] - pm["COGS"]
        pm["Gross_Margin_Pct"] = np.where(pm["Revenue"] > 0, pm["Gross_Profit"] / pm["Revenue"], np.nan)
    else:
        pm = pd.DataFrame()

    return {
        "params": sheets.get("PARAMETERS", pd.DataFrame()),
        "targets": sheets.get("TARGETS", pd.DataFrame()),
        "dim": dim,
        "fja_map": fja_map,
        "yearly": yr,
        "fg": fg,
        "fja": fja,
        "dept": dept,
        "capacity": cap,
        "project_margin": pm,
    }


# =========================
# DECISION ENGINES
# =========================
def business_posture(rev_g, hc_g, rpe_g, mcr, targets: Dict[str, float]) -> Dict[str, str]:
    rev_min = get_target(targets, "Revenue_Growth_Min", 0.20)
    rpe_min = get_target(targets, "RPE_Growth_Min", 0.15)
    hc_max = get_target(targets, "Headcount_Growth_Max_High_Leverage", 0.10)
    mcr_watch = get_target(targets, "MCR_Watch_Max", 0.09)

    state = "BALANCED"
    tone = "info"
    hint = "Pertumbuhan, kapasitas, dan people-cost relatif seimbang."

    if pd.notna(rev_g) and rev_g <= 0:
        state, tone, hint = "DEFENSIVE", "error", "Revenue melemah; prioritas proteksi margin, cash, dan fungsi kritikal."
    elif pd.notna(rev_g) and pd.notna(rpe_g) and pd.notna(hc_g) and pd.notna(mcr) and rev_g >= rev_min and rpe_g >= rpe_min and hc_g <= hc_max and mcr <= mcr_watch:
        state, tone, hint = "HIGH_LEVERAGE", "success", "Revenue & RPE tumbuh kuat dengan HC terkendali dan MCR sehat/watch rendah."
    elif pd.notna(rev_g) and pd.notna(hc_g) and rev_g > 0.10 and hc_g > 0.15 and (pd.isna(rpe_g) or rpe_g <= 0):
        state, tone, hint = "CAPACITY_RISK", "warning", "Pertumbuhan mulai dibeli dengan tambahan kapasitas; validasi produktivitas diperlukan."
    elif pd.notna(mcr) and mcr > mcr_watch:
        state, tone, hint = "COST_PRESSURE", "warning", "People-cost mulai menekan; fokus pada disiplin biaya dan unit economics."

    return {"state": state, "tone": tone, "hint": hint}


def margin_quality(gm, targets: Dict[str, float]) -> Dict[str, str]:
    strong = get_target(targets, "Gross_Margin_Strong_Min", 0.35)
    healthy = get_target(targets, "Gross_Margin_Healthy_Min", 0.25)
    watch = get_target(targets, "Gross_Margin_Watch_Min", 0.15)
    pressure = get_target(targets, "Gross_Margin_Pressure_Min", 0.10)

    if pd.isna(gm):
        return {"state": "UNKNOWN", "tone": "info", "hint": "Gross Margin belum tersedia."}
    if gm >= strong:
        return {"state": "STRONG", "tone": "success", "hint": "Margin sangat kuat; ruang profitabilitas tinggi."}
    if gm >= healthy:
        return {"state": "HEALTHY", "tone": "success", "hint": "Margin sehat; tetap jaga pricing, delivery cost, dan project mix."}
    if gm >= watch:
        return {"state": "WATCH", "tone": "warning", "hint": "Margin perlu dipantau; validasi pricing, COGS, scope creep, dan biaya delivery."}
    if gm >= pressure:
        return {"state": "PRESSURE", "tone": "warning", "hint": "Margin mulai tertekan; perlu review profitabilitas proyek/layanan."}
    return {"state": "CRITICAL", "tone": "error", "hint": "Margin kritikal; perlu tindakan korektif segera."}


def mcr_health(mcr, targets: Dict[str, float]) -> Dict[str, str]:
    ultra = get_target(targets, "MCR_Ultra_Efficiency_Max", 0.05)
    healthy_max = get_target(targets, "MCR_Healthy_Max", 0.07)
    watch_max = get_target(targets, "MCR_Watch_Max", 0.09)
    pressure_max = get_target(targets, "MCR_Cost_Pressure_Max", 0.14)

    if pd.isna(mcr):
        return {"state": "UNKNOWN", "tone": "info", "hint": "MCR belum tersedia."}
    if mcr < ultra:
        return {"state": "ULTRA_EFFICIENCY", "tone": "warning", "hint": "Sangat efisien; perlu validasi kapasitas dan sustainability."}
    if mcr <= healthy_max:
        return {"state": "HEALTHY", "tone": "success", "hint": "MCR berada di rentang sehat 5–7%."}
    if mcr <= watch_max:
        return {"state": "WATCH", "tone": "warning", "hint": "MCR mulai perlu dipantau agar tidak bergerak lebih cepat dari revenue."}
    if mcr <= pressure_max:
        return {"state": "COST_PRESSURE", "tone": "warning", "hint": "People-cost mulai menekan profitabilitas."}
    return {"state": "CRITICAL", "tone": "error", "hint": "People-cost berada pada level kritikal."}


def capacity_risk(latest: pd.Series, cap_df: pd.DataFrame, targets: Dict[str, float]) -> Dict[str, object]:
    rev_g = latest.get("Revenue_YoY")
    hc_g = latest.get("Headcount_YoY")
    rpe_g = latest.get("RPE_YoY")
    mcr = latest.get("MCR_Pct")
    year = latest.get("Year")

    score = 0
    reasons = []

    if pd.notna(rev_g) and pd.notna(hc_g):
        gap = rev_g - hc_g
        if gap > 0.30:
            score += 2
            reasons.append("Revenue growth jauh di atas HC growth")
        elif gap > 0.15:
            score += 1
            reasons.append("Revenue growth lebih cepat dari kapasitas")
        if hc_g > 0.15:
            score += 1
            reasons.append("HC growth agresif")

    if pd.notna(rpe_g):
        if rpe_g > 0.30:
            score += 1
            reasons.append("RPE naik sangat tinggi")
        elif rpe_g < 0:
            score += 2
            reasons.append("RPE menurun")

    if pd.notna(mcr):
        if mcr < get_target(targets, "MCR_Ultra_Efficiency_Max", 0.05):
            score += 2
            reasons.append("MCR sangat rendah")
        elif mcr <= get_target(targets, "MCR_Healthy_Max", 0.07):
            score += 1
            reasons.append("MCR efisien, sustainability perlu dijaga")
        elif mcr > get_target(targets, "MCR_Cost_Pressure_Max", 0.14):
            score += 2
            reasons.append("MCR tinggi")

    if cap_df is not None and not cap_df.empty and pd.notna(year):
        c = cap_df[cap_df["Year"] == int(year)].copy()
        if not c.empty:
            util = c["Avg_Utilization_Pct"].mean() if "Avg_Utilization_Pct" in c.columns else np.nan
            sla = c["SLA_Breach_Count"].sum() if "SLA_Breach_Count" in c.columns else 0
            backlog = c["Backlog_Count"].sum() if "Backlog_Count" in c.columns else 0
            turnover = c["Turnover_Count"].sum() if "Turnover_Count" in c.columns else 0
            if pd.notna(util) and util >= 0.85:
                score += 1
                reasons.append("Utilisasi rata-rata tinggi")
            if sla > 0:
                score += 1
                reasons.append("Terdapat SLA breach")
            if backlog > 20:
                score += 1
                reasons.append("Backlog cukup tinggi")
            if turnover > 0:
                score += 1
                reasons.append("Ada turnover pada periode berjalan")

    if score >= 6:
        state, tone = "CRITICAL", "error"
    elif score >= 4:
        state, tone = "HIGH", "warning"
    elif score >= 2:
        state, tone = "MEDIUM", "warning"
    else:
        state, tone = "LOW", "success"

    return {"state": state, "tone": tone, "score": score, "hint": "; ".join(reasons) if reasons else "Tidak ada sinyal kapasitas besar."}


def management_action(business_state: str, margin_state: str, mcr_state: str, cap_state: str) -> Tuple[str, List[str]]:
    if business_state == "HIGH_LEVERAGE":
        title = "Selective Scale-Up with Capacity Guardrail"
        actions = [
            "Pertahankan struktur headcount yang disiplin; hindari ekspansi massal.",
            "Lakukan selective hiring hanya pada bottleneck revenue/delivery critical.",
            "Perkuat succession planning, knowledge transfer, dan retention role kunci.",
            "Monitor overtime, backlog, SLA/incident, dan turnover role kritikal secara bulanan.",
        ]
    elif business_state == "CAPACITY_RISK":
        title = "Capacity Stabilization"
        actions = [
            "Validasi kebutuhan kapasitas per fungsi sebelum hiring lanjutan.",
            "Evaluasi produktivitas per role/function dan perbaiki bottleneck delivery.",
            "Prioritaskan automation/process improvement sebelum menambah fixed cost.",
        ]
    elif business_state == "COST_PRESSURE":
        title = "People-Cost Control"
        actions = [
            "Perketat approval hiring berbasis ROI dan kebutuhan delivery.",
            "Audit overtime, bonus, benefit, dan struktur manpower cost.",
            "Review unit economics per project/service.",
        ]
    elif business_state == "DEFENSIVE":
        title = "Margin & Cash Protection"
        actions = [
            "Freeze hiring non-critical dan review biaya tetap.",
            "Lindungi fungsi kritikal yang menjaga revenue dan delivery.",
            "Fokus pada cash discipline, margin recovery, dan project profitability.",
        ]
    else:
        title = "Balanced Optimization"
        actions = [
            "Optimalkan struktur yang ada sebelum ekspansi baru.",
            "Tambahkan kapasitas hanya pada bottleneck yang terbukti.",
            "Jaga MCR, RPE, dan driver-support ratio secara triwulan.",
        ]

    if margin_state in ["WATCH", "PRESSURE", "CRITICAL"]:
        actions.append("Lakukan margin quality review: pricing, COGS, scope creep, rework, dan delivery cost.")
    if mcr_state in ["ULTRA_EFFICIENCY", "WATCH"]:
        actions.append("Validasi sustainability MCR melalui workload, SLA, backlog, dan turnover.")
    if cap_state in ["HIGH", "CRITICAL"]:
        actions.append("Lakukan capacity risk review bulanan sampai risiko turun ke Medium/Low.")

    return title, actions


def badge_html(label: str, state: str, tone: str, sub: str = "") -> str:
    color = TONE_COLOR.get(tone, "#64748b")
    return f"""
    <div style="
        border:1px solid rgba(148,163,184,.35);
        border-radius:16px;
        padding:14px 16px;
        background:rgba(15,23,42,.04);
        min-height:94px;">
        <div style="font-size:12px; color:#64748b; text-transform:uppercase; letter-spacing:.08em;">{label}</div>
        <div style="display:flex; align-items:center; gap:8px; margin-top:8px;">
            <span style="width:12px; height:12px; background:{color}; border-radius:99px; display:inline-block;"></span>
            <span style="font-size:20px; font-weight:700;">{state}</span>
        </div>
        <div style="font-size:12px; color:#64748b; margin-top:6px;">{sub}</div>
    </div>
    """


# =========================
# SIDEBAR
# =========================
with st.sidebar:
    st.header("⚙️ Data Source")
    uploaded = st.file_uploader("Upload Excel v2", type=["xlsx"])
    st.caption("Jika tidak upload, app akan membaca file default dari folder data/ di repo.")

file_bytes = uploaded.getvalue() if uploaded is not None else None
sheets, warnings = load_workbook(file_bytes, DEFAULT_FILE)

# =========================
# HEADER
# =========================
st.markdown(
    """
    <div style="padding:4px 0 10px 0;">
        <div style="font-size:36px; font-weight:800; line-height:1.1;">HR vs Revenue Intelligence Platform v2</div>
        <div style="font-size:15px; color:#64748b; margin-top:6px;">Board & CFO Edition — workforce productivity, cost leverage, margin quality, and capacity risk.</div>
    </div>
    """,
    unsafe_allow_html=True,
)

if warnings:
    for w in warnings:
        st.warning(w)

if any(s not in sheets for s in REQUIRED_SHEETS):
    st.error("Data belum lengkap. Pastikan Excel memiliki sheet wajib: " + ", ".join(REQUIRED_SHEETS))
    st.stop()

data = clean_and_calc(sheets)
params = build_params(data["params"])
targets = load_targets(data["targets"])
currency = params.get("Currency_Code", "IDR")

yr = data["yearly"].dropna(subset=["Year"]).copy()
if yr.empty:
    st.error("Data yearly kosong.")
    st.stop()

available_years = sorted(yr["Year"].dropna().astype(int).unique().tolist())
with st.sidebar:
    st.header("🧭 Filters")
    year_sel = st.selectbox("Analysis Year", available_years, index=len(available_years)-1)
    year_min, year_max = st.slider("Trend Range", min_value=min(available_years), max_value=max(available_years), value=(min(available_years), max(available_years)), step=1)
    fja_filter = st.multiselect("FJA Category", sorted(data["fja"]["FJA_Category"].dropna().unique().tolist()), default=sorted(data["fja"]["FJA_Category"].dropna().unique().tolist()))

yr_range = yr[(yr["Year"] >= year_min) & (yr["Year"] <= year_max)].copy()
latest = yr[yr["Year"] == year_sel].iloc[0]
prev_df = yr[yr["Year"] < year_sel].sort_values("Year")
prev = prev_df.iloc[-1] if not prev_df.empty else None

bp = business_posture(latest.get("Revenue_YoY"), latest.get("Headcount_YoY"), latest.get("RPE_YoY"), latest.get("MCR_Pct"), targets)
mq = margin_quality(latest.get("Gross_Margin_Pct"), targets)
mh = mcr_health(latest.get("MCR_Pct"), targets)
cr = capacity_risk(latest, data["capacity"], targets)

# =========================
# EXECUTIVE COCKPIT
# =========================
st.markdown("## 1. Executive Cockpit")

b1, b2, b3, b4 = st.columns(4)
with b1:
    st.markdown(badge_html("Business Posture", bp["state"], bp["tone"], bp["hint"]), unsafe_allow_html=True)
with b2:
    st.markdown(badge_html("Margin Quality", mq["state"], mq["tone"], mq["hint"]), unsafe_allow_html=True)
with b3:
    st.markdown(badge_html("MCR Health", mh["state"], mh["tone"], mh["hint"]), unsafe_allow_html=True)
with b4:
    st.markdown(badge_html("Capacity Risk", cr["state"], cr["tone"], f"Score {cr['score']} — {cr['hint']}"), unsafe_allow_html=True)

st.markdown("")

k1, k2, k3 = st.columns(3)
k4, k5, k6 = st.columns(3)
with k1:
    st.metric("💰 Revenue", _money(latest.get("Total_Revenue"), currency), _pct(latest.get("Revenue_YoY")))
with k2:
    st.metric("👥 Headcount", _num(latest.get("Total_Headcount")), _pct(latest.get("Headcount_YoY")))
with k3:
    st.metric("⚡ RPE", _money(latest.get("RPE"), currency), _pct(latest.get("RPE_YoY")))
with k4:
    st.metric("📊 MCR", _pct(latest.get("MCR_Pct")), _pp(latest.get("MCR_Delta")))
with k5:
    st.metric("📈 Gross Margin", _pct(latest.get("Gross_Margin_Pct")), _pp(latest.get("GrossMargin_Delta")))
with k6:
    st.metric("💼 Manpower Cost", _money(latest.get("Total_Manpower_Cost"), currency), _pct(latest.get("ManpowerCost_YoY")))

st.markdown("#### Board Memo")
memo = (
    f"Pada {int(year_sel)}, revenue mencapai **{_money(latest.get('Total_Revenue'), currency)}** "
    f"dengan headcount **{_num(latest.get('Total_Headcount'))}** dan RPE **{_money(latest.get('RPE'), currency)}**. "
    f"Business Posture terbaca **{bp['state']}**, sementara Margin Quality berada di **{mq['state']}** dan MCR Health **{mh['state']}**. "
    f"Fokus management: {management_action(bp['state'], mq['state'], mh['state'], cr['state'])[0]}."
)
st.info(memo)

# =========================
# WORKFORCE PRODUCTIVITY
# =========================
st.markdown("## 2. Workforce Productivity")

c1, c2 = st.columns(2)
with c1:
    fig = px.line(yr_range, x="Year", y=["Total_Revenue", "Total_Manpower_Cost"], markers=True, title="Revenue vs Manpower Cost")
    fig.update_layout(margin=dict(l=10, r=10, t=45, b=10))
    st.plotly_chart(fig, use_container_width=True)
with c2:
    fig = px.line(yr_range, x="Year", y=["RPE", "Cost_per_HC"], markers=True, title="Revenue per Employee vs Cost per HC")
    fig.update_layout(margin=dict(l=10, r=10, t=45, b=10))
    st.plotly_chart(fig, use_container_width=True)

c3, c4 = st.columns(2)
with c3:
    fig = px.line(yr_range, x="Year", y="Revenue_per_Payroll", markers=True, title="Revenue Generated per 1 Payroll Rupiah")
    fig.update_layout(margin=dict(l=10, r=10, t=45, b=10))
    st.plotly_chart(fig, use_container_width=True)
with c4:
    growth = yr_range[["Year", "Revenue_YoY", "Headcount_YoY", "RPE_YoY", "ManpowerCost_YoY"]].melt(id_vars="Year", var_name="Metric", value_name="YoY")
    fig = px.line(growth, x="Year", y="YoY", color="Metric", markers=True, title="YoY Growth Comparison")
    fig.update_yaxes(tickformat=".0%")
    fig.update_layout(margin=dict(l=10, r=10, t=45, b=10))
    st.plotly_chart(fig, use_container_width=True)

# =========================
# FJA ANALYTICS
# =========================
st.markdown("## 3. Functional Job Analysis (FJA)")

fja = data["fja"].copy()
fja = fja[(fja["Year"] >= year_min) & (fja["Year"] <= year_max)]
if fja_filter:
    fja = fja[fja["FJA_Category"].isin(fja_filter)]

fja_latest = fja[fja["Year"] == year_sel].copy()
fc1, fc2 = st.columns([1.2, 1.0])
with fc1:
    if not fja_latest.empty:
        fig = px.treemap(fja_latest, path=["FJA_Category"], values="Manpower_Cost", title=f"Functional Cost Mix — {year_sel}")
        fig.update_layout(margin=dict(l=10, r=10, t=45, b=10))
        st.plotly_chart(fig, use_container_width=True)
with fc2:
    if not fja_latest.empty:
        fja_show = fja_latest[["FJA_Category", "Manpower_Cost", "Cost_Share", "Headcount", "HC_Share"]].copy()
        fja_show["Manpower_Cost"] = fja_show["Manpower_Cost"].map(lambda x: _money(x, currency))
        fja_show["Cost_Share"] = fja_show["Cost_Share"].map(_pct)
        fja_show["HC_Share"] = fja_show["HC_Share"].map(_pct)
        st.dataframe(fja_show.sort_values("Cost_Share", ascending=False), use_container_width=True, hide_index=True)

fig = px.area(fja, x="Year", y="Manpower_Cost", color="FJA_Category", title="FJA Manpower Cost Trend")
fig.update_layout(margin=dict(l=10, r=10, t=45, b=10))
st.plotly_chart(fig, use_container_width=True)

# =========================
# DEPARTMENT PRODUCTIVITY
# =========================
st.markdown("## 4. Department Productivity")

dept = data["dept"].copy()
dept = dept[dept["Year"] == year_sel].copy()
if fja_filter:
    dept = dept[dept["FJA_Category"].isin(fja_filter)]

if not dept.empty:
    fig = px.scatter(
        dept,
        x="Headcount",
        y="Manpower_Cost",
        size="Dept_Revenue",
        color="FJA_Category",
        hover_data=["Dept_ID", "Dept_Name", "Function_Group", "Dept_RPE", "Revenue_per_Cost"],
        title="Department Cost vs Capacity vs Revenue",
    )
    fig.update_layout(margin=dict(l=10, r=10, t=45, b=10))
    st.plotly_chart(fig, use_container_width=True)

    topn = st.slider("Top N Department Table", 3, 20, 8)
    dept_show = dept[["Dept_ID", "Dept_Name", "Function_Group", "FJA_Category", "Headcount", "Manpower_Cost", "Dept_Revenue", "Dept_RPE", "Revenue_per_Cost", "Cost_per_HC"]].copy()
    st.dataframe(dept_show.sort_values("Manpower_Cost", ascending=False).head(topn), use_container_width=True, hide_index=True)

# =========================
# CAPACITY RISK CENTER
# =========================
st.markdown("## 5. Capacity Risk Center")

cap = data["capacity"].copy()
cap_year = cap[cap["Year"] == year_sel].copy() if not cap.empty else pd.DataFrame()
cc1, cc2 = st.columns([1.2, 1.0])
with cc1:
    if not cap_year.empty:
        fig = px.scatter(
            cap_year,
            x="Avg_Utilization_Pct",
            y="Overtime_Hours",
            size="Backlog_Count",
            color="Critical_Role_Dependency",
            hover_data=["Dept_ID", "SLA_Breach_Count", "Incident_Count", "Turnover_Count"],
            title=f"Capacity Risk Heatmap Proxy — {year_sel}",
        )
        fig.update_xaxes(tickformat=".0%")
        fig.update_layout(margin=dict(l=10, r=10, t=45, b=10))
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("CAPACITY_INDICATOR belum diisi.")
with cc2:
    st.markdown("#### Capacity Reading")
    st.write(f"**Risk:** {cr['state']} | **Score:** {cr['score']}")
    st.write(cr["hint"])
    st.caption("Capacity Risk adalah early warning, bukan bukti burnout. Validasi dengan overtime, utilisasi, backlog, SLA/incident, turnover, dan dependency role kritikal.")

# =========================
# PROJECT / MARGIN QUALITY
# =========================
st.markdown("## 6. Margin Quality & Project Profitability")

pm = data["project_margin"].copy()
pm_year = pm[pm["Year"] == year_sel].copy() if not pm.empty else pd.DataFrame()
mc1, mc2 = st.columns(2)
with mc1:
    fig = px.line(yr_range, x="Year", y="Gross_Margin_Pct", markers=True, title="Gross Margin Trend")
    fig.update_yaxes(tickformat=".0%")
    fig.update_layout(margin=dict(l=10, r=10, t=45, b=10))
    st.plotly_chart(fig, use_container_width=True)
with mc2:
    if not pm_year.empty:
        fig = px.bar(pm_year.sort_values("Gross_Margin_Pct"), x="Project_ID", y="Gross_Margin_Pct", color="Business_Line", title=f"Project Gross Margin — {year_sel}")
        fig.update_yaxes(tickformat=".0%")
        fig.update_layout(margin=dict(l=10, r=10, t=45, b=10))
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("PROJECT_MARGIN belum diisi.")

# =========================
# WHAT-IF SIMULATOR
# =========================
st.markdown("## 7. What-if Simulator")

s1, s2, s3 = st.columns(3)
with s1:
    add_hc = st.slider("Additional Headcount", 0, 50, 5)
with s2:
    avg_cost = st.number_input("Avg Annual Cost per New HC", min_value=0.0, value=float(latest.get("Cost_per_HC") if pd.notna(latest.get("Cost_per_HC")) else 75_000_000), step=5_000_000.0)
with s3:
    revenue_uplift = st.slider("Expected Revenue Uplift", -0.20, 1.00, 0.10, step=0.01, format="%.2f")

proj_revenue = latest.get("Total_Revenue") * (1 + revenue_uplift)
proj_cost = latest.get("Total_Manpower_Cost") + add_hc * avg_cost
proj_hc = latest.get("Total_Headcount") + add_hc
proj_rpe = proj_revenue / proj_hc if proj_hc else np.nan
proj_mcr = proj_cost / proj_revenue if proj_revenue else np.nan

wc1, wc2, wc3 = st.columns(3)
with wc1:
    st.metric("Projected Revenue", _money(proj_revenue, currency), _pct(revenue_uplift))
with wc2:
    st.metric("Projected RPE", _money(proj_rpe, currency), f"{((proj_rpe/latest.get('RPE'))-1)*100:+.1f}%" if latest.get("RPE") else "")
with wc3:
    st.metric("Projected MCR", _pct(proj_mcr), _pp(proj_mcr - latest.get("MCR_Pct")))

st.caption("Simulator ini directional. Untuk keputusan formal, validasi dengan pipeline, backlog, delivery capacity, dan margin per project.")

# =========================
# BOARD DECISION CENTER
# =========================
st.markdown("## 8. Board Decision Center")

trend = yr_range.copy()
trend["Business_Posture"] = trend.apply(lambda r: business_posture(r.get("Revenue_YoY"), r.get("Headcount_YoY"), r.get("RPE_YoY"), r.get("MCR_Pct"), targets)["state"], axis=1)
trend["Margin_Quality"] = trend["Gross_Margin_Pct"].apply(lambda x: margin_quality(x, targets)["state"])
trend["MCR_Health"] = trend["MCR_Pct"].apply(lambda x: mcr_health(x, targets)["state"])

bc1, bc2 = st.columns([1.1, 1.0])
with bc1:
    fig = px.scatter(
        trend,
        x="Year",
        y="Business_Posture",
        color="Margin_Quality",
        size=trend["Revenue_YoY"].fillna(0).abs() + 0.05,
        hover_data=["Revenue_YoY", "Headcount_YoY", "RPE_YoY", "MCR_Pct", "Gross_Margin_Pct"],
        title="Decision State Trend",
    )
    fig.update_layout(margin=dict(l=10, r=10, t=45, b=10))
    st.plotly_chart(fig, use_container_width=True)
with bc2:
    title, actions = management_action(bp["state"], mq["state"], mh["state"], cr["state"])
    st.markdown(f"#### Recommended Action: {title}")
    for i, a in enumerate(actions, start=1):
        st.write(f"{i}. {a}")

st.dataframe(
    trend[["Year", "Business_Posture", "Margin_Quality", "MCR_Health", "Revenue_YoY", "Headcount_YoY", "RPE_YoY", "MCR_Pct", "Gross_Margin_Pct"]],
    use_container_width=True,
    hide_index=True,
)

# =========================
# DATA QUALITY
# =========================
with st.expander("🔍 Data Quality & Raw Tables"):
    st.write("Required sheets loaded:", ", ".join(REQUIRED_SHEETS))
    st.write("Available sheets loaded:", ", ".join(sheets.keys()))
    st.markdown("#### Yearly KPI")
    st.dataframe(yr, use_container_width=True, hide_index=True)
    st.markdown("#### Department")
    st.dataframe(data["dept"], use_container_width=True, hide_index=True)

st.divider()
st.caption("HR vs Revenue Intelligence Platform v2 — Board & CFO Edition. Interpretasi MCR harus selalu dibaca bersama Revenue Growth, RPE, Headcount, Margin Quality, dan Capacity Risk.")
