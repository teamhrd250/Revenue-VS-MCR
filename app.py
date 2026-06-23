
# app.py
# MCR vs Revenue Platform V2 — Board & CFO Edition
# Streamlit dashboard based on Excel data model:
# HR_vs_Revenue_Intelligence_Platform_v2_Board_CFO_Template_2026_YTD_Annualized.xlsx

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
    page_title="MCR vs Revenue Platform V2",
    page_icon="📊",
    layout="wide",
)

DEFAULT_FILE = os.path.join(
    "data",
    "HR_vs_Revenue_Intelligence_Platform_v2_Board_CFO_Template_2026_YTD_Annualized.xlsx",
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
    "CALENDAR",
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


def _find_logo_path() -> str:
    """Cari logo perusahaan untuk header kanan atas.

    Simpan salah satu nama file berikut di GitHub:
    - assets/logo_company.png
    - assets/logo_company.jpg
    - assets/company_logo.png
    - logo_company.png
    """
    candidates = [
        os.path.join("assets", "logo_company.png"),
        os.path.join("assets", "logo_company.jpg"),
        os.path.join("assets", "logo_company.jpeg"),
        os.path.join("assets", "company_logo.png"),
        os.path.join("assets", "company_logo.jpg"),
        os.path.join("assets", "company_logo.jpeg"),
        "logo_company.png",
        "logo_company.jpg",
        "company_logo.png",
        "company_logo.jpg",
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    return ""


def _image_to_base64(path: str) -> Tuple[str, str]:
    import base64
    import mimetypes
    mime, _ = mimetypes.guess_type(path)
    mime = mime or "image/png"
    with open(path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode("utf-8")
    return mime, b64


def render_top_header() -> None:
    """Header dengan logo di kanan atas."""
    logo_path = _find_logo_path()
    col_title, col_logo = st.columns([6, 1.4])

    with col_title:
        st.markdown(
            """
            <div style="padding:4px 0 10px 0;">
                <div style="font-size:36px; font-weight:800; line-height:1.1;">MCR vs Revenue Platform V2</div>
                <div style="font-size:15px; color:#64748b; margin-top:6px;">Board Edition - workforce productivity, cost leverage, margin quality and capacity risk.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with col_logo:
        if logo_path:
            mime, b64 = _image_to_base64(logo_path)
            st.markdown(
                f"""
                <div style="text-align:right; padding-top:2px;">
                    <img src="data:{mime};base64,{b64}"
                         style="max-width:150px; max-height:70px; object-fit:contain;" />
                </div>
                """,
                unsafe_allow_html=True,
            )
        else:
            st.markdown(
                """
                <div style="text-align:right; padding-top:10px; color:#94a3b8; font-size:12px;">
                    Logo belum ditemukan<br>
                    <span style="font-size:11px;">assets/logo_company.png</span>
                </div>
                """,
                unsafe_allow_html=True,
            )



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


def build_calendar(calendar: Optional[pd.DataFrame], years: Optional[List[int]] = None) -> pd.DataFrame:
    """Build FY/YTD calendar.

    Required columns if provided:
    Year, Months_Closed, Period_Type, Period_Label, Annualization_Factor, Data_Confidence
    """
    if calendar is None or calendar.empty or "Year" not in calendar.columns:
        years = years or []
        return pd.DataFrame({
            "Year": years,
            "Months_Closed": [12 for _ in years],
            "Period_Type": ["FY" for _ in years],
            "Period_Label": [f"FY {y}" for y in years],
            "Annualization_Factor": [1.0 for _ in years],
            "Data_Confidence": ["High" for _ in years],
        })

    cal = calendar.copy()
    cal["Year"] = pd.to_numeric(cal["Year"], errors="coerce").astype("Int64")
    cal["Months_Closed"] = pd.to_numeric(cal.get("Months_Closed"), errors="coerce").fillna(12).clip(lower=1, upper=12)
    if "Period_Type" not in cal.columns:
        cal["Period_Type"] = np.where(cal["Months_Closed"] < 12, "YTD", "FY")
    if "Period_Label" not in cal.columns:
        cal["Period_Label"] = cal.apply(lambda r: f"{r['Period_Type']} {int(r['Year'])}", axis=1)
    if "Annualization_Factor" not in cal.columns:
        cal["Annualization_Factor"] = 12 / cal["Months_Closed"]
    else:
        cal["Annualization_Factor"] = pd.to_numeric(cal["Annualization_Factor"], errors="coerce")
        cal["Annualization_Factor"] = cal["Annualization_Factor"].fillna(12 / cal["Months_Closed"])
    if "Data_Confidence" not in cal.columns:
        cal["Data_Confidence"] = np.where(cal["Months_Closed"] >= 12, "High", np.where(cal["Months_Closed"] >= 6, "Medium", "Low"))

    cal = cal.dropna(subset=["Year"]).copy()
    cal["Year"] = cal["Year"].astype(int)
    return cal[["Year", "Months_Closed", "Period_Type", "Period_Label", "Annualization_Factor", "Data_Confidence"]]


def get_calendar_info(calendar: pd.DataFrame, year: int) -> Dict[str, object]:
    if calendar is None or calendar.empty or "Year" not in calendar.columns:
        return {"Months_Closed": 12, "Period_Type": "FY", "Period_Label": f"FY {year}", "Annualization_Factor": 1.0, "Data_Confidence": "High"}
    row = calendar[calendar["Year"] == int(year)]
    if row.empty:
        return {"Months_Closed": 12, "Period_Type": "FY", "Period_Label": f"FY {year}", "Annualization_Factor": 1.0, "Data_Confidence": "High"}
    r = row.iloc[0]
    return {
        "Months_Closed": int(r.get("Months_Closed", 12)),
        "Period_Type": str(r.get("Period_Type", "FY")),
        "Period_Label": str(r.get("Period_Label", f"FY {year}")),
        "Annualization_Factor": float(r.get("Annualization_Factor", 1.0)),
        "Data_Confidence": str(r.get("Data_Confidence", "High")),
    }


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


def clean_and_calc(sheets: Dict[str, pd.DataFrame], analysis_mode: str = "Actual") -> Dict[str, pd.DataFrame]:
    dim = sheets["DIM_DEPARTMENT"].copy()
    fja_map = sheets["FJA_MAPPING"].copy()
    rev = sheets["REVENUE_YR"].copy()
    hc = sheets["HEADCOUNT_YR"].copy()
    pay = sheets["PAYROLL_YR"].copy()

    # Calendar/YTD awareness.
    raw_years = []
    for _df in [rev, hc, pay]:
        if "Year" in _df.columns:
            raw_years.extend(pd.to_numeric(_df["Year"], errors="coerce").dropna().astype(int).tolist())
    calendar = build_calendar(sheets.get("CALENDAR", pd.DataFrame()), sorted(set(raw_years)))
    annualization_map = dict(zip(calendar["Year"].astype(int), calendar["Annualization_Factor"].astype(float)))

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

    rev["Revenue_Actual"] = rev["Revenue_Recognized"]
    rev["COGS_Actual"] = rev["COGS_Direct"]
    if analysis_mode == "Annualized":
        _rev_factor = rev["Year"].astype("float").map(annualization_map).fillna(1.0)
        rev["Revenue_Recognized"] = rev["Revenue_Recognized"] * _rev_factor
        rev["COGS_Direct"] = rev["COGS_Direct"] * _rev_factor

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

    for _c in ["Payroll_Gross", "Overtime", "Bonus", "Benefits", "Employer_Tax", "Total_Manpower_Cost"]:
        if _c in pay.columns:
            pay[f"{_c}_Actual"] = pay[_c]
    if analysis_mode == "Annualized":
        _pay_factor = pay["Year"].astype("float").map(annualization_map).fillna(1.0)
        for _c in ["Payroll_Gross", "Overtime", "Bonus", "Benefits", "Employer_Tax", "Total_Manpower_Cost"]:
            if _c in pay.columns:
                pay[_c] = pay[_c] * _pay_factor

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
        rd["Revenue_Actual"] = rd["Revenue_Recognized"]
        if analysis_mode == "Annualized":
            _rd_factor = rd["Year"].astype("float").map(annualization_map).fillna(1.0)
            rd["Revenue_Recognized"] = rd["Revenue_Recognized"] * _rd_factor
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
        if analysis_mode == "Annualized":
            _pm_factor = pm["Year"].astype("float").map(annualization_map).fillna(1.0)
            for _c in ["Revenue", "COGS", "Gross_Profit"]:
                if _c in pm.columns:
                    pm[_c] = pm[_c] * _pm_factor
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
        "calendar": calendar,
        "analysis_mode": analysis_mode,
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
    """Capacity Score Engine V2 - Board Edition.

    Score range 0-100:
    - Utilization: 25
    - Overtime: 20
    - Backlog: 20
    - SLA Breach: 15
    - Turnover: 10
    - Critical Role Dependency: 10

    Tidak berbenturan dengan TARGETS Excel karena engine ini membaca CAPACITY_INDICATOR,
    sedangkan TARGETS tetap dipakai untuk Revenue, RPE, MCR, dan Margin Quality.
    """
    year = latest.get("Year")

    score = 0
    reasons = []
    details = {
        "avg_utilization": np.nan,
        "max_utilization": np.nan,
        "total_overtime": 0.0,
        "total_backlog": 0.0,
        "total_sla_breach": 0.0,
        "total_turnover": 0.0,
        "critical_dependency": "UNKNOWN",
    }

    if cap_df is None or cap_df.empty or pd.isna(year):
        return {
            "state": "LOW",
            "tone": "success",
            "score": 0,
            "recommendation": "No Hiring Required",
            "hint": "CAPACITY_INDICATOR belum tersedia; risk default LOW.",
            "details": details,
        }

    c = cap_df[cap_df["Year"] == int(year)].copy()
    if c.empty:
        return {
            "state": "LOW",
            "tone": "success",
            "score": 0,
            "recommendation": "No Hiring Required",
            "hint": "Data capacity untuk tahun terpilih belum tersedia; risk default LOW.",
            "details": details,
        }

    util = pd.to_numeric(c.get("Avg_Utilization_Pct"), errors="coerce").fillna(0.0)
    overtime = pd.to_numeric(c.get("Overtime_Hours"), errors="coerce").fillna(0.0)
    backlog = pd.to_numeric(c.get("Backlog_Count"), errors="coerce").fillna(0.0)
    sla = pd.to_numeric(c.get("SLA_Breach_Count"), errors="coerce").fillna(0.0)
    turnover = pd.to_numeric(c.get("Turnover_Count"), errors="coerce").fillna(0.0)

    avg_util = float(util.mean()) if len(util) else 0.0
    max_util = float(util.max()) if len(util) else 0.0
    total_overtime = float(overtime.sum())
    total_backlog = float(backlog.sum())
    total_sla = float(sla.sum())
    total_turnover = float(turnover.sum())

    # Critical role dependency: take highest severity found in the selected year.
    dep_series = c.get("Critical_Role_Dependency", pd.Series(dtype=str)).astype(str).str.upper().str.strip()
    dep_rank = {"LOW": 1, "MEDIUM": 2, "HIGH": 3, "CRITICAL": 4}
    dep_ranked = dep_series.map(dep_rank).fillna(0)
    max_dep_rank = int(dep_ranked.max()) if len(dep_ranked) else 0
    dep_value = {1: "LOW", 2: "MEDIUM", 3: "HIGH", 4: "CRITICAL"}.get(max_dep_rank, "UNKNOWN")

    details.update({
        "avg_utilization": avg_util,
        "max_utilization": max_util,
        "total_overtime": total_overtime,
        "total_backlog": total_backlog,
        "total_sla_breach": total_sla,
        "total_turnover": total_turnover,
        "critical_dependency": dep_value,
    })

    # Utilization score, max 25
    # Use max utilization to detect bottleneck even if average still looks normal.
    util_ref = max_util
    if util_ref >= 0.90:
        score += 25
        reasons.append("Utilization >90% pada fungsi tertentu")
    elif util_ref >= 0.85:
        score += 18
        reasons.append("Utilization 85-90% pada fungsi tertentu")
    elif util_ref >= 0.75:
        score += 10
        reasons.append("Utilization 75-85% perlu dimonitor")

    # Overtime score, max 20
    if total_overtime > 400:
        score += 20
        reasons.append("Overtime >400 jam")
    elif total_overtime > 250:
        score += 15
        reasons.append("Overtime 250-400 jam")
    elif total_overtime > 100:
        score += 8
        reasons.append("Overtime 100-250 jam")

    # Backlog score, max 20
    if total_backlog >= 20:
        score += 20
        reasons.append("Backlog >=20")
    elif total_backlog >= 10:
        score += 12
        reasons.append("Backlog 10-19")
    elif total_backlog > 0:
        score += 6
        reasons.append("Backlog mulai muncul")

    # SLA breach score, max 15
    if total_sla >= 5:
        score += 15
        reasons.append("SLA breach >=5")
    elif total_sla > 0:
        score += 8
        reasons.append("Terdapat SLA breach")

    # Turnover score, max 10
    if total_turnover >= 10:
        score += 10
        reasons.append("Turnover >=10")
    elif total_turnover > 0:
        score += 5
        reasons.append("Ada turnover pada periode berjalan")

    # Critical Role Dependency score, max 10
    if dep_value in ["HIGH", "CRITICAL"]:
        score += 10
        reasons.append("Critical role dependency tinggi")
    elif dep_value == "MEDIUM":
        score += 5
        reasons.append("Critical role dependency medium")

    score = int(min(100, round(score)))

    if score > 75:
        state, tone, recommendation = "CRITICAL", "error", "Immediate Capacity Action Required"
    elif score >= 56:
        state, tone, recommendation = "HIGH", "warning", "Capacity Expansion Review"
    elif score >= 31:
        state, tone, recommendation = "MEDIUM", "warning", "Selective Hiring"
    else:
        state, tone, recommendation = "LOW", "success", "No Hiring Required"

    return {
        "state": state,
        "tone": tone,
        "score": score,
        "recommendation": recommendation,
        "hint": "; ".join(reasons) if reasons else "Tidak ada sinyal kapasitas besar.",
        "details": details,
    }


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



def _score_from_growth(value, strong=0.20, good=0.10, weak=0.00) -> int:
    """Convert growth metric to 0-100 score."""
    if pd.isna(value):
        return 50
    value = float(value)
    if value >= strong:
        return 100
    if value >= good:
        return 80
    if value >= weak:
        return 60
    if value >= -0.10:
        return 35
    return 15


def _score_from_margin_state(state: str) -> int:
    return {
        "STRONG": 100,
        "HEALTHY": 85,
        "WATCH": 60,
        "PRESSURE": 35,
        "CRITICAL": 10,
        "UNKNOWN": 50,
    }.get(state, 50)


def _score_from_mcr_state(state: str) -> int:
    return {
        "HEALTHY": 95,
        "WATCH": 75,
        "ULTRA_EFFICIENCY": 70,
        "COST_PRESSURE": 40,
        "CRITICAL": 10,
        "UNKNOWN": 50,
    }.get(state, 50)


def _score_from_capacity_state(state: str) -> int:
    return {
        "LOW": 100,
        "MEDIUM": 75,
        "HIGH": 40,
        "CRITICAL": 10,
    }.get(state, 50)


def board_score_engine(row: pd.Series, targets: Dict[str, float], cap_state: str = "LOW") -> Dict[str, object]:
    """Board Score Engine v3.

    Weighted score:
    - Revenue Growth 25%
    - RPE Growth 25%
    - Margin Quality 20%
    - MCR Health 15%
    - Capacity Risk 15%
    """
    rev_g = row.get("Revenue_YoY")
    rpe_g = row.get("RPE_YoY")
    gm = row.get("Gross_Margin_Pct")
    mcr = row.get("MCR_Pct")

    mq_state = margin_quality(gm, targets)["state"]
    mh_state = mcr_health(mcr, targets)["state"]

    components = {
        "Revenue Growth": _score_from_growth(rev_g, strong=get_target(targets, "Revenue_Growth_Min", 0.20), good=0.10, weak=0.00),
        "RPE Growth": _score_from_growth(rpe_g, strong=get_target(targets, "RPE_Growth_Min", 0.15), good=0.08, weak=0.00),
        "Margin Quality": _score_from_margin_state(mq_state),
        "MCR Health": _score_from_mcr_state(mh_state),
        "Capacity Risk": _score_from_capacity_state(cap_state),
    }
    weights = {
        "Revenue Growth": 0.25,
        "RPE Growth": 0.25,
        "Margin Quality": 0.20,
        "MCR Health": 0.15,
        "Capacity Risk": 0.15,
    }
    score = sum(components[k] * weights[k] for k in components)
    score = int(round(score))

    if score >= 85:
        status, tone, icon = "GROWTH_READY", "success", "🟢"
        label = "Growth Ready"
    elif score >= 70:
        status, tone, icon = "OPTIMIZE_GROWTH", "info", "🔵"
        label = "Optimize Growth"
    elif score >= 55:
        status, tone, icon = "SELECTIVE_HIRING", "warning", "🟡"
        label = "Selective Hiring"
    elif score >= 40:
        status, tone, icon = "CAPACITY_ALERT", "warning", "🟠"
        label = "Capacity Alert"
    else:
        status, tone, icon = "DEFENSIVE_MODE", "error", "🔴"
        label = "Defensive Mode"

    return {
        "score": score,
        "status": status,
        "label": label,
        "tone": tone,
        "icon": icon,
        "components": components,
        "weights": weights,
        "margin_state": mq_state,
        "mcr_state": mh_state,
        "capacity_state": cap_state,
    }


def board_status_interpretation(bs: Dict[str, object], bp_state: str, mq_state: str, mh_state: str, cap_state: str) -> str:
    label = bs.get("label", "Balanced")
    if bs.get("status") == "GROWTH_READY":
        return (
            "Perusahaan berada pada posisi siap tumbuh. Revenue dan produktivitas relatif kuat, "
            "struktur people-cost masih terkendali, dan risiko kapasitas belum menjadi hambatan utama. "
            "Arah keputusan: scale-up selektif dengan tetap menjaga margin dan kualitas delivery."
        )
    if bs.get("status") == "OPTIMIZE_GROWTH":
        return (
            "Perusahaan memiliki momentum pertumbuhan, namun masih ada area optimasi pada margin, MCR, atau kapasitas. "
            "Arah keputusan: pertahankan growth, perbaiki margin quality, dan lakukan hiring hanya pada fungsi yang terbukti menjadi bottleneck."
        )
    if bs.get("status") == "SELECTIVE_HIRING":
        return (
            "Pertumbuhan masih dapat dilanjutkan, tetapi kapasitas dan efisiensi perlu dikontrol lebih ketat. "
            "Arah keputusan: selective hiring berbasis ROI, prioritas pada revenue generator/enabler, dan review utilisasi per fungsi."
        )
    if bs.get("status") == "CAPACITY_ALERT":
        return (
            "Ada sinyal bahwa kapasitas, MCR, atau kualitas margin mulai membatasi pertumbuhan. "
            "Arah keputusan: audit workload, backlog, SLA, dan struktur biaya sebelum ekspansi lebih lanjut."
        )
    return (
        "Perusahaan perlu masuk mode proteksi. Fokus utama bukan ekspansi, melainkan menjaga margin, cash, dan fungsi kritikal. "
        "Arah keputusan: freeze hiring non-critical, review cost base, dan pulihkan profitabilitas."
    )


def board_action_v3(bs: Dict[str, object], bp_state: str, mq_state: str, mh_state: str, cap_state: str) -> List[str]:
    status = bs.get("status")
    if status == "GROWTH_READY":
        actions = [
            "Setujui scale-up selektif pada fungsi revenue generator dan revenue enabler yang terbukti mendukung pipeline/delivery.",
            "Pertahankan MCR pada zona sehat; setiap tambahan HC wajib dikaitkan dengan revenue atau kapasitas delivery yang jelas.",
            "Gunakan momentum growth untuk memperkuat role kritikal, succession plan, dan retention key talent.",
            "Jaga Margin Quality agar tidak turun akibat pricing discount, scope creep, atau delivery cost."
        ]
    elif status == "OPTIMIZE_GROWTH":
        actions = [
            "Pertahankan struktur organisasi saat ini sambil memperbaiki margin quality dan revenue per employee.",
            "Prioritaskan improvement pada project profitability, pricing discipline, dan cost-to-serve.",
            "Tambahkan kapasitas hanya pada bottleneck yang terbukti dari backlog, SLA, atau utilisasi.",
            "Monitor MCR dan RPE secara kuartalan sebagai guardrail pertumbuhan."
        ]
    elif status == "SELECTIVE_HIRING":
        actions = [
            "Lakukan hiring terbatas hanya untuk posisi kritikal yang memiliki dampak langsung pada revenue/delivery.",
            "Validasi kebutuhan tambahan HC dengan workload, backlog, SLA, dan revenue pipeline.",
            "Tahan penambahan fungsi support kecuali ada justifikasi governance/compliance yang kuat.",
            "Perbaiki produktivitas melalui automation, SOP, dan redistribusi workload sebelum menambah fixed cost."
        ]
    elif status == "CAPACITY_ALERT":
        actions = [
            "Lakukan capacity review lintas fungsi: overtime, utilisasi, backlog, SLA breach, incident, dan turnover.",
            "Evaluasi apakah efisiensi MCR berasal dari produktivitas sehat atau under-capacity.",
            "Tunda ekspansi non-critical sampai bottleneck delivery dan margin pressure terkendali.",
            "Susun recovery plan untuk fungsi dengan cost tinggi, productivity rendah, atau dependency tinggi."
        ]
    else:
        actions = [
            "Freeze hiring non-critical dan fokus pada fungsi yang menjaga revenue serta cash collection.",
            "Review project profitability dan hentikan aktivitas dengan margin negatif atau tidak strategis.",
            "Audit manpower cost, overtime, benefit, dan struktur organisasi.",
            "Susun profit recovery plan dengan target margin, cost, dan cash yang terukur."
        ]

    if mq_state in ["WATCH", "PRESSURE", "CRITICAL"]:
        actions.append("Tambahkan margin quality review khusus: pricing, COGS, scope creep, rework, dan delivery cost.")
    if cap_state in ["HIGH", "CRITICAL"]:
        actions.append("Wajibkan monthly capacity risk review sampai status turun minimal ke MEDIUM.")
    if mh_state in ["ULTRA_EFFICIENCY"]:
        actions.append("Validasi risiko under-capacity karena MCR terlalu rendah dapat menyembunyikan overload tim.")
    return actions


def board_status_card_html(bs: Dict[str, object], confidence: str, mq_state: str, mh_state: str, cap_state: str) -> str:
    color = TONE_COLOR.get(bs.get("tone", "info"), "#64748b")
    return f"""
    <div style="
        border:1px solid rgba(148,163,184,.35);
        border-radius:22px;
        padding:22px;
        background:linear-gradient(135deg, rgba(15,23,42,.05), rgba(15,23,42,.01));
        min-height:260px;">
        <div style="font-size:12px; color:#64748b; text-transform:uppercase; letter-spacing:.10em;">Board Status</div>
        <div style="font-size:34px; font-weight:900; margin-top:10px;">
            {bs.get('icon','🔵')} {bs.get('label','Balanced')}
        </div>
        <div style="font-size:54px; font-weight:900; color:{color}; margin-top:8px;">
            {bs.get('score',0)}/100
        </div>
        <div style="height:10px; background:#e5e7eb; border-radius:999px; margin-top:8px; overflow:hidden;">
            <div style="height:10px; width:{bs.get('score',0)}%; background:{color}; border-radius:999px;"></div>
        </div>
        <div style="margin-top:16px; font-size:13px; color:#475569; line-height:1.8;">
            Confidence: <b>{confidence}</b><br>
            Margin Quality: <b>{mq_state}</b><br>
            MCR Health: <b>{mh_state}</b><br>
            Capacity Risk: <b>{cap_state}</b>
        </div>
    </div>
    """

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
render_top_header()

if warnings:
    for w in warnings:
        st.warning(w)

if any(s not in sheets for s in REQUIRED_SHEETS):
    st.error("Data belum lengkap. Pastikan Excel memiliki sheet wajib: " + ", ".join(REQUIRED_SHEETS))
    st.stop()

# Calendar determines whether a year is full-year or YTD.
_pre_years = []
for _s in ["REVENUE_YR", "HEADCOUNT_YR", "PAYROLL_YR"]:
    if _s in sheets and "Year" in sheets[_s].columns:
        _pre_years.extend(pd.to_numeric(sheets[_s]["Year"], errors="coerce").dropna().astype(int).tolist())
_calendar_preview = build_calendar(sheets.get("CALENDAR", pd.DataFrame()), sorted(set(_pre_years)))

with st.sidebar:
    st.header("📅 Analysis Mode")
    analysis_mode = st.radio(
        "Compare mode",
        ["Actual", "Annualized"],
        index=1 if ((_calendar_preview["Months_Closed"] < 12).any() if not _calendar_preview.empty else False) else 0,
        help="Actual = angka sesuai data. Annualized = YTD disetahunkan untuk fair comparison vs full-year."
    )

data = clean_and_calc(sheets, analysis_mode=analysis_mode)
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
period_info = get_calendar_info(data.get("calendar", pd.DataFrame()), year_sel)

bp = business_posture(latest.get("Revenue_YoY"), latest.get("Headcount_YoY"), latest.get("RPE_YoY"), latest.get("MCR_Pct"), targets)
mq = margin_quality(latest.get("Gross_Margin_Pct"), targets)
mh = mcr_health(latest.get("MCR_Pct"), targets)
cr = capacity_risk(latest, data["capacity"], targets)

# =========================
# EXECUTIVE COCKPIT
# =========================
st.markdown("## 1. Executive Summary")
st.caption(
    f"Periode data: {period_info.get('Period_Label')} | Months closed: {period_info.get('Months_Closed')}/12 | "
    f"Mode: {analysis_mode} | Annualization factor: {period_info.get('Annualization_Factor'):.2f}x | "
    f"Data confidence: {period_info.get('Data_Confidence')}"
)
if analysis_mode == "Actual" and period_info.get("Months_Closed", 12) < 12:
    st.warning("Periode terpilih masih YTD. Komparasi Actual vs tahun full-year tidak apple-to-apple. Gunakan Annualized mode untuk pembacaan directional yang lebih fair.")
elif analysis_mode == "Annualized" and period_info.get("Months_Closed", 12) < 12:
    st.info("Annualized mode aktif: Revenue, COGS, Dept Revenue, Project Revenue/COGS, dan Manpower Cost disetahunkan. Headcount tetap sebagai average capacity.")

b1, b2, b3, b4 = st.columns(4)
with b1:
    st.markdown(badge_html("Business Posture", bp["state"], bp["tone"], bp["hint"]), unsafe_allow_html=True)
with b2:
    st.markdown(badge_html("Margin Quality", mq["state"], mq["tone"], mq["hint"]), unsafe_allow_html=True)
with b3:
    st.markdown(badge_html("MCR Health", mh["state"], mh["tone"], mh["hint"]), unsafe_allow_html=True)
with b4:
    st.markdown(badge_html("Capacity Risk", cr["state"], cr["tone"], f"Score {cr['score']}/100 — {cr.get('recommendation', '-')}"), unsafe_allow_html=True)

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
    f"Pada {period_info.get('Period_Label')} dalam mode **{analysis_mode}**, revenue terbaca **{_money(latest.get('Total_Revenue'), currency)}** "
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
    # Board-grade Department Productivity Map.
    # X = Headcount, Y = Manpower Cost, Bubble = Dept Revenue, Color = FJA Category.
    # If Dept Revenue is empty/zero, bubble uses Headcount as fallback so chart never crashes.
    dept_plot = dept.copy()
    dept_plot["Dept_Name"] = dept_plot["Dept_Name"].fillna(dept_plot["Dept_ID"]).astype(str)
    dept_plot["FJA_Category"] = dept_plot["FJA_Category"].fillna("Unmapped").astype(str)
    dept_plot["Dept_Revenue"] = pd.to_numeric(dept_plot.get("Dept_Revenue"), errors="coerce").fillna(0.0)
    dept_plot["Headcount"] = pd.to_numeric(dept_plot.get("Headcount"), errors="coerce").fillna(0.0)
    dept_plot["Manpower_Cost"] = pd.to_numeric(dept_plot.get("Manpower_Cost"), errors="coerce").fillna(0.0)
    dept_plot["Bubble_Size"] = dept_plot["Dept_Revenue"].clip(lower=0)

    if dept_plot["Bubble_Size"].max() <= 0:
        dept_plot["Bubble_Size"] = dept_plot["Headcount"].clip(lower=1)
        bubble_note = "Bubble size memakai Headcount karena Revenue by Dept belum tersedia/masih nol."
    else:
        min_positive = dept_plot.loc[dept_plot["Bubble_Size"] > 0, "Bubble_Size"].min()
        dept_plot["Bubble_Size"] = dept_plot["Bubble_Size"].replace(0, min_positive * 0.12)
        bubble_note = "Bubble size memakai Dept Revenue; nilai revenue nol diberi minimum visual agar tetap terbaca."

    fja_color_map = {
        "Revenue Generator": "#22C55E",
        "Revenue Enabler": "#3B82F6",
        "Support Function": "#F59E0B",
        "Governance / Management": "#8B5CF6",
        "Governance": "#8B5CF6",
        "Unmapped": "#64748B",
    }

    fig = px.scatter(
        dept_plot,
        x="Headcount",
        y="Manpower_Cost",
        size="Bubble_Size",
        color="FJA_Category",
        text="Dept_Name",
        color_discrete_map=fja_color_map,
        hover_name="Dept_Name",
        hover_data={
            "Dept_ID": True,
            "Function_Group": True,
            "FJA_Category": True,
            "Headcount": ":.1f",
            "Manpower_Cost": ":,.0f",
            "Dept_Revenue": ":,.0f",
            "Dept_RPE": ":,.0f",
            "Revenue_per_Cost": ":.2f",
            "Bubble_Size": False,
            "Dept_Name": False,
        },
        title="Department Productivity Map",
        labels={
            "Headcount": "Headcount",
            "Manpower_Cost": "Manpower Cost",
            "FJA_Category": "FJA Category",
        },
        size_max=56,
    )

    fig.update_traces(
        textposition="middle center",
        textfont=dict(size=11, color="white"),
        marker=dict(
            opacity=0.86,
            line=dict(width=2, color="rgba(255,255,255,0.85)")
        ),
        selector=dict(mode="markers+text"),
    )
    fig.update_layout(
        margin=dict(l=10, r=10, t=58, b=10),
        legend_title_text="FJA Category",
        height=520,
        title=dict(
            text="Department Productivity Map<br><sup>Bubble Size = Revenue | X = Headcount | Y = Manpower Cost | Color = FJA Category</sup>",
            x=0.0,
            xanchor="left",
        ),
    )
    fig.update_xaxes(showgrid=True, gridcolor="rgba(148,163,184,0.25)", zeroline=False)
    fig.update_yaxes(showgrid=True, gridcolor="rgba(148,163,184,0.25)", zeroline=False)

    st.plotly_chart(fig, use_container_width=True)
    st.caption(bubble_note + " Idealnya bubble besar berada tidak terlalu tinggi di Manpower Cost, menandakan revenue besar dengan biaya terkendali.")

    # Quick board reading cards
    highest_cost = dept_plot.sort_values("Manpower_Cost", ascending=False).head(1)
    highest_rpe = dept_plot.replace([np.inf, -np.inf], np.nan).dropna(subset=["Dept_RPE"]).sort_values("Dept_RPE", ascending=False).head(1)
    low_rev_high_cost = dept_plot.copy()
    low_rev_high_cost["Cost_Intensity"] = np.where(low_rev_high_cost["Dept_Revenue"] > 0, low_rev_high_cost["Manpower_Cost"] / low_rev_high_cost["Dept_Revenue"], np.nan)
    cost_watch = low_rev_high_cost.dropna(subset=["Cost_Intensity"]).sort_values("Cost_Intensity", ascending=False).head(1)

    rc1, rc2, rc3 = st.columns(3)
    with rc1:
        if not highest_cost.empty:
            r = highest_cost.iloc[0]
            st.metric("Highest Manpower Cost", r["Dept_Name"], _money(r["Manpower_Cost"], currency))
    with rc2:
        if not highest_rpe.empty:
            r = highest_rpe.iloc[0]
            st.metric("Highest Dept RPE", r["Dept_Name"], _money(r["Dept_RPE"], currency))
    with rc3:
        if not cost_watch.empty:
            r = cost_watch.iloc[0]
            st.metric("Cost Watch", r["Dept_Name"], f"{r['Cost_Intensity']:.1%} cost/revenue")

    with st.expander("How to read Department Productivity Map"):
        st.markdown(
            """
            - **Bubble besar** = revenue departemen besar.
            - **Semakin ke kanan** = headcount semakin banyak.
            - **Semakin ke atas** = manpower cost semakin tinggi.
            - **Hijau** = Revenue Generator, **Biru** = Revenue Enabler, **Oranye** = Support Function, **Ungu** = Governance.
            - Idealnya revenue besar tidak selalu diikuti cost yang terlalu tinggi.
            - Bubble kecil tetapi berada tinggi perlu ditinjau karena berpotensi **high cost - low revenue**.
            """
        )

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
    st.metric("Capacity Score", f"{cr['score']}/100", cr["state"])
    st.write(f"**Recommendation:** {cr.get('recommendation', '-')}")
    st.write(cr["hint"])

    d = cr.get("details", {})
    d1, d2 = st.columns(2)
    with d1:
        st.metric("Max Utilization", _pct(d.get("max_utilization")))
        st.metric("Overtime Hours", _num(d.get("total_overtime")))
        st.metric("Backlog", _num(d.get("total_backlog")))
    with d2:
        st.metric("SLA Breach", _num(d.get("total_sla_breach")))
        st.metric("Turnover", _num(d.get("total_turnover")))
        st.metric("Dependency", d.get("critical_dependency", "-"))

    if cr["state"] == "LOW":
        st.success("Kapasitas organisasi masih memadai. Tidak ada indikasi bottleneck besar.")
    elif cr["state"] == "MEDIUM":
        st.warning("Mulai ada tekanan kapasitas pada beberapa fungsi. Monitoring utilisasi, backlog, dan overtime perlu diperketat.")
    elif cr["state"] == "HIGH":
        st.warning("Beberapa fungsi menunjukkan tanda bottleneck. Perlu capacity expansion review dan selective hiring pada fungsi kritikal.")
    else:
        st.error("Risiko overload organisasi tinggi dan berpotensi mempengaruhi SLA, kualitas delivery, serta keberlanjutan operasional.")

    with st.expander("How to Read Capacity Risk Score"):
        st.markdown("""
        ### Capacity Score Classification

        | Score | Status | Board Meaning |
        |---:|---|---|
        | **0-30** | **LOW** | Kapasitas masih longgar / terkendali |
        | **31-55** | **MEDIUM** | Ada tekanan kapasitas namun masih terkendali |
        | **56-75** | **HIGH** | Bottleneck mulai muncul; perlu capacity review |
        | **76-100** | **CRITICAL** | Risiko overload organisasi; perlu tindakan segera |

        ### Scoring Weight

        | Faktor | Bobot Maksimum |
        |---|---:|
        | Utilization | 25 |
        | Overtime | 20 |
        | Backlog | 20 |
        | SLA Breach | 15 |
        | Turnover | 10 |
        | Critical Role Dependency | 10 |

        ### Bubble Chart Reading

        - **X Axis** = Average Utilization %
        - **Y Axis** = Overtime Hours
        - **Bubble Size** = Backlog Count
        - **Bubble Color** = Critical Role Dependency

        **Kanan atas** adalah area paling berisiko: utilisasi tinggi, overtime tinggi, dan backlog besar.
        """)

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

board_score = board_score_engine(latest, targets, cr["state"])
board_confidence = str(period_info.get("Data_Confidence", "High"))

bd1, bd2 = st.columns([0.95, 1.25])
with bd1:
    st.markdown(
        board_status_card_html(
            board_score,
            board_confidence,
            mq["state"],
            mh["state"],
            cr["state"],
        ),
        unsafe_allow_html=True,
    )

with bd2:
    fig = go.Figure(
        go.Indicator(
            mode="gauge+number",
            value=board_score["score"],
            number={"suffix": "/100"},
            title={"text": f"Board Score — {board_score['label']}"},
            gauge={
                "axis": {"range": [0, 100]},
                "bar": {"color": TONE_COLOR.get(board_score.get("tone", "info"), "#3b82f6")},
                "steps": [
                    {"range": [0, 40], "color": "rgba(239,68,68,.18)"},
                    {"range": [40, 55], "color": "rgba(245,158,11,.18)"},
                    {"range": [55, 70], "color": "rgba(234,179,8,.18)"},
                    {"range": [70, 85], "color": "rgba(59,130,246,.18)"},
                    {"range": [85, 100], "color": "rgba(34,197,94,.18)"},
                ],
                "threshold": {
                    "line": {"color": "#0f172a", "width": 3},
                    "thickness": 0.75,
                    "value": board_score["score"],
                },
            },
        )
    )
    fig.update_layout(height=260, margin=dict(l=10, r=10, t=50, b=10))
    st.plotly_chart(fig, use_container_width=True)

st.markdown("#### Executive Interpretation")
st.info(board_status_interpretation(board_score, bp["state"], mq["state"], mh["state"], cr["state"]))

comp = pd.DataFrame({
    "Factor": list(board_score["components"].keys()),
    "Score": list(board_score["components"].values()),
    "Weight": [board_score["weights"][k] for k in board_score["components"].keys()],
})
comp["Weighted_Score"] = comp["Score"] * comp["Weight"]

bc1, bc2 = st.columns([1.1, 1.0])
with bc1:
    fig = px.bar(
        comp,
        x="Factor",
        y="Score",
        text="Score",
        title="Board Score Components",
        hover_data=["Weight", "Weighted_Score"],
    )
    fig.update_yaxes(range=[0, 100])
    fig.update_layout(margin=dict(l=10, r=10, t=45, b=10))
    st.plotly_chart(fig, use_container_width=True)

with bc2:
    st.markdown(f"#### Recommended Action: {board_score['label']}")
    actions_v3 = board_action_v3(board_score, bp["state"], mq["state"], mh["state"], cr["state"])
    for i, a in enumerate(actions_v3, start=1):
        st.write(f"{i}. {a}")

trend = yr_range.copy()
trend["Business_Posture"] = trend.apply(lambda r: business_posture(r.get("Revenue_YoY"), r.get("Headcount_YoY"), r.get("RPE_YoY"), r.get("MCR_Pct"), targets)["state"], axis=1)
trend["Margin_Quality"] = trend["Gross_Margin_Pct"].apply(lambda x: margin_quality(x, targets)["state"])
trend["MCR_Health"] = trend["MCR_Pct"].apply(lambda x: mcr_health(x, targets)["state"])

# Build capacity state per trend year using the same engine.
_cap_states = []
_board_scores = []
_board_labels = []
for _, r in trend.iterrows():
    _cr = capacity_risk(r, data["capacity"], targets)
    _bs = board_score_engine(r, targets, _cr["state"])
    _cap_states.append(_cr["state"])
    _board_scores.append(_bs["score"])
    _board_labels.append(_bs["label"])

trend["Capacity_Risk"] = _cap_states
trend["Board_Score"] = _board_scores
trend["Board_Status"] = _board_labels

st.markdown("#### Board Status Trend")
tc1, tc2 = st.columns([1.2, 1.0])
with tc1:
    fig = px.line(
        trend,
        x="Year",
        y="Board_Score",
        markers=True,
        color="Board_Status",
        title="Board Score Trend",
        hover_data=["Business_Posture", "Margin_Quality", "MCR_Health", "Capacity_Risk", "Revenue_YoY", "RPE_YoY", "MCR_Pct", "Gross_Margin_Pct"],
    )
    fig.add_hrect(y0=85, y1=100, line_width=0, fillcolor="rgba(34,197,94,.08)")
    fig.add_hrect(y0=70, y1=85, line_width=0, fillcolor="rgba(59,130,246,.08)")
    fig.add_hrect(y0=55, y1=70, line_width=0, fillcolor="rgba(234,179,8,.08)")
    fig.add_hrect(y0=40, y1=55, line_width=0, fillcolor="rgba(245,158,11,.08)")
    fig.add_hrect(y0=0, y1=40, line_width=0, fillcolor="rgba(239,68,68,.08)")
    fig.update_yaxes(range=[0, 100])
    fig.update_layout(margin=dict(l=10, r=10, t=45, b=10))
    st.plotly_chart(fig, use_container_width=True)

with tc2:
    status_summary = trend[["Year", "Board_Score", "Board_Status", "Business_Posture", "Margin_Quality", "MCR_Health", "Capacity_Risk"]].copy()
    st.dataframe(status_summary, use_container_width=True, hide_index=True)

with st.expander("How to read Board Score"):
    st.markdown(
        """
        **Board Score** adalah ringkasan directional untuk keputusan Direksi, bukan pengganti analisis detail.

        Bobot:
        - Revenue Growth: 25%
        - RPE Growth: 25%
        - Margin Quality: 20%
        - MCR Health: 15%
        - Capacity Risk: 15%

        Status:
        - **85–100 Growth Ready**: siap scale-up terukur.
        - **70–85 Optimize Growth**: growth ada, fokus optimasi margin/kapasitas.
        - **55–70 Selective Hiring**: boleh hiring terbatas pada bottleneck.
        - **40–55 Capacity Alert**: audit kapasitas sebelum ekspansi.
        - **<40 Defensive Mode**: proteksi margin, cash, dan fungsi kritikal.
        """
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
st.caption("MCR vs Revenue Platform V2 — Board & CFO Edition. Interpretasi MCR harus selalu dibaca bersama Revenue Growth, RPE, Headcount, Margin Quality, dan Capacity Risk.")
