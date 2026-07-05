
# app.py
# Starcom Executive Platform (SEP) V4 - Executive Intelligence Edition
# Clean IA: one board = one business question.
# Focus: MCR, Revenue, Productivity, FJA, Capacity, Strategic Decision, Benchmark Framework, CEO Closing.

import os
from io import BytesIO
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


st.set_page_config(
    page_title="SEP V4 Executive Intelligence",
    page_icon="📊",
    layout="wide",
)

DEFAULT_FILE = os.path.join("data", "SEP_V4_Final_Executive_Intelligence_Template.xlsx")

REQUIRED_SHEETS = [
    "PARAMETERS", "TARGETS", "FJA_MAPPING", "DIM_DEPARTMENT",
    "REVENUE_YR", "HEADCOUNT_YR", "PAYROLL_YR"
]
OPTIONAL_SHEETS = [
    "DIM_EMPLOYEE", "REVENUE_BY_DEPT", "HC_MOVEMENT", "PROJECT_MARGIN",
    "CAPACITY_INDICATOR", "CALENDAR", "MONTHLY_KPI", "BENCHMARK_FRAMEWORK",
    "SEP_PRESENTATION_FLOW", "APPENDIX_KPI_DEFINITIONS"
]

TONE_COLOR = {
    "success": "#22c55e",
    "info": "#3b82f6",
    "warning": "#f59e0b",
    "error": "#ef4444",
    "neutral": "#64748b",
}


# =========================
# FORMATTERS
# =========================
def _safe_str(x) -> str:
    if x is None or pd.isna(x):
        return ""
    return str(x).strip()


def _safe_float(x):
    try:
        if x is None or pd.isna(x):
            return np.nan
        return float(x)
    except Exception:
        return np.nan


def _money(x, currency="IDR") -> str:
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


def _find_logo_path() -> str:
    for p in [
        os.path.join("assets", "logo_company.png"),
        os.path.join("assets", "logo_company.jpg"),
        os.path.join("assets", "starcom_logo.png"),
        "logo_company.png",
        "company_logo.png",
    ]:
        if os.path.exists(p):
            return p
    return ""


def render_header():
    c1, c2 = st.columns([6, 1.3])
    with c1:
        st.markdown("""
        <div style="padding:4px 0 12px 0;">
          <div style="font-size:36px;font-weight:900;line-height:1.1;">Starcom Executive Platform (SEP) V4</div>
          <div style="font-size:15px;color:#64748b;margin-top:6px;">Executive Intelligence Edition — MCR, revenue leverage, workforce productivity, capacity risk, benchmark framework, and board decision.</div>
        </div>
        """, unsafe_allow_html=True)
    with c2:
        logo = _find_logo_path()
        if logo:
            st.image(logo, use_container_width=True)
        else:
            st.caption("Logo: assets/logo_company.png")


def badge(label, value, tone="info", sub="") -> str:
    color = TONE_COLOR.get(tone, "#3b82f6")
    return f"""
    <div style="border:1px solid rgba(148,163,184,.35); border-radius:16px; padding:15px 16px; background:rgba(15,23,42,.04); min-height:100px;">
      <div style="font-size:12px; color:#64748b; text-transform:uppercase; letter-spacing:.08em;">{label}</div>
      <div style="font-size:22px; font-weight:800; margin-top:8px;"><span style="display:inline-block;width:11px;height:11px;border-radius:999px;background:{color};margin-right:8px;"></span>{value}</div>
      <div style="font-size:12px;color:#64748b;margin-top:6px;line-height:1.4;">{sub}</div>
    </div>
    """


# =========================
# DATA
# =========================
@st.cache_data(show_spinner=False)
def load_workbook(file_bytes: Optional[bytes], default_path: str):
    warnings = []
    try:
        if file_bytes:
            xls = pd.ExcelFile(BytesIO(file_bytes))
        else:
            if not os.path.exists(default_path):
                return {}, [f"File default tidak ditemukan: {default_path}. Upload Excel terlebih dahulu."]
            xls = pd.ExcelFile(default_path)
    except Exception as e:
        return {}, [f"Excel tidak bisa dibaca: {type(e).__name__}: {e}"]

    sheets = {}
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


def load_params(df: pd.DataFrame) -> Dict[str, str]:
    out = {}
    if df is None or df.empty or not {"Parameter", "Value"}.issubset(df.columns):
        return out
    for _, r in df.iterrows():
        k = _safe_str(r.get("Parameter"))
        if k:
            out[k] = _safe_str(r.get("Value"))
    return out


def load_targets(df: pd.DataFrame) -> Dict[str, float]:
    out = {}
    if df is None or df.empty or not {"KPI", "Target_Value"}.issubset(df.columns):
        return out
    for _, r in df.iterrows():
        k = _safe_str(r.get("KPI"))
        try:
            if k:
                out[k] = float(r.get("Target_Value"))
        except Exception:
            pass
    return out


def get_target(targets: Dict[str, float], key: str, default: float) -> float:
    return float(targets.get(key, default))


def build_calendar(calendar: Optional[pd.DataFrame], years: List[int]) -> pd.DataFrame:
    if calendar is None or calendar.empty or "Year" not in calendar.columns:
        return pd.DataFrame({
            "Year": years,
            "Months_Closed": [12] * len(years),
            "Period_Type": ["FY"] * len(years),
            "Period_Label": [f"FY {y}" for y in years],
            "Data_Confidence": ["High"] * len(years),
        })
    cal = calendar.copy()
    cal["Year"] = pd.to_numeric(cal["Year"], errors="coerce").astype("Int64")
    cal["Months_Closed"] = pd.to_numeric(cal.get("Months_Closed"), errors="coerce").fillna(12).clip(1, 12)
    if "Period_Type" not in cal.columns:
        cal["Period_Type"] = np.where(cal["Months_Closed"] < 12, "YTD", "FY")
    if "Period_Label" not in cal.columns:
        cal["Period_Label"] = cal.apply(lambda r: f"{r['Period_Type']} {int(r['Year'])}", axis=1)
    if "Data_Confidence" not in cal.columns:
        cal["Data_Confidence"] = np.where(cal["Months_Closed"] >= 12, "High", "Medium")
    cal = cal.dropna(subset=["Year"]).copy()
    cal["Year"] = cal["Year"].astype(int)
    return cal[["Year", "Months_Closed", "Period_Type", "Period_Label", "Data_Confidence"]]


def mode_maps(calendar: pd.DataFrame, mode: str):
    cal = calendar.copy()
    cal["Months_Closed"] = pd.to_numeric(cal["Months_Closed"], errors="coerce").fillna(12).clip(1, 12)
    ytd = cal[cal["Months_Closed"] < 12]
    comparable_months = int(ytd["Months_Closed"].max()) if not ytd.empty else 12
    if mode == "Actual":
        amount = {int(r["Year"]): 1.0 for _, r in cal.iterrows()}
        rev = amount.copy()
        cost = amount.copy()
        note = "Actual: angka sesuai input Excel."
    elif mode == "YTD Comparable":
        amount = {int(r["Year"]): float(comparable_months / r["Months_Closed"]) for _, r in cal.iterrows()}
        rev = amount.copy()
        cost = amount.copy()
        note = f"YTD Comparable: semua tahun dibaca pada basis {comparable_months} bulan. Jika MONTHLY_KPI terisi, dashboard memakai realisasi aktual bulan yang sama."
    elif mode == "Annual Projection":
        amount = {int(r["Year"]): float(12 / r["Months_Closed"]) for _, r in cal.iterrows()}
        rev = amount.copy()
        cost = amount.copy()
        note = "Annual Projection: YTD disetahunkan memakai faktor 12 / Months_Closed."
    elif mode == "MCR Focus":
        rev = {int(r["Year"]): float(12 / r["Months_Closed"]) for _, r in cal.iterrows()}
        cost = {int(r["Year"]): 1.0 for _, r in cal.iterrows()}
        amount = cost.copy()
        note = "MCR Focus: revenue diproyeksikan full-year, sementara COGS dan manpower cost tetap actual/YTD untuk stress-test MCR dan margin."
    else:
        amount = {int(r["Year"]): 1.0 for _, r in cal.iterrows()}
        rev = amount.copy()
        cost = amount.copy()
        note = "Fallback Actual."
    return amount, rev, cost, note, comparable_months


def get_period_info(calendar: pd.DataFrame, year: int):
    if calendar is None or calendar.empty:
        return {"Period_Label": f"FY {year}", "Months_Closed": 12, "Data_Confidence": "High"}
    x = calendar[calendar["Year"] == int(year)]
    if x.empty:
        return {"Period_Label": f"FY {year}", "Months_Closed": 12, "Data_Confidence": "High"}
    r = x.iloc[0]
    return {
        "Period_Label": str(r.get("Period_Label", f"FY {year}")),
        "Months_Closed": int(r.get("Months_Closed", 12)),
        "Data_Confidence": str(r.get("Data_Confidence", "High")),
    }


def prep_data(sheets: Dict[str, pd.DataFrame], mode: str):
    dim = sheets["DIM_DEPARTMENT"].copy()
    fja_map = sheets["FJA_MAPPING"].copy()
    rev = sheets["REVENUE_YR"].copy()
    hc = sheets["HEADCOUNT_YR"].copy()
    pay = sheets["PAYROLL_YR"].copy()

    raw_years = []
    for df in [rev, hc, pay]:
        if "Year" in df.columns:
            raw_years += pd.to_numeric(df["Year"], errors="coerce").dropna().astype(int).tolist()
    calendar = build_calendar(sheets.get("CALENDAR", pd.DataFrame()), sorted(set(raw_years)))
    amount_map, rev_map, cost_map, mode_note, comparable_months = mode_maps(calendar, mode)
    annual_rev_map = {int(r["Year"]): float(12 / r["Months_Closed"]) for _, r in calendar.iterrows()}

    # FJA mapping
    for c in ["Dept_ID", "Dept_Name", "Function_Group", "Revenue_Driver_Flag", "FJA_Category_Override"]:
        if c in dim.columns:
            dim[c] = dim[c].astype(str).str.strip()
    if "FJA_Category_Override" not in dim.columns:
        dim["FJA_Category_Override"] = ""
    if not fja_map.empty and {"Function_Group", "FJA_Category"}.issubset(fja_map.columns):
        fja_lookup = dict(zip(fja_map["Function_Group"].astype(str).str.strip(), fja_map["FJA_Category"].astype(str).str.strip()))
    else:
        fja_lookup = {"Sales": "Revenue Generator", "Operations/Project": "Revenue Enabler", "Engineering": "Revenue Enabler", "Support": "Support Function", "Management": "Governance / Management"}
    dim["FJA_Category"] = np.where(
        dim["FJA_Category_Override"].astype(str).str.strip().ne(""),
        dim["FJA_Category_Override"].astype(str).str.strip(),
        dim["Function_Group"].map(fja_lookup).fillna("Unmapped"),
    )
    dim_small = dim[["Dept_ID", "Dept_Name", "Function_Group", "Revenue_Driver_Flag", "FJA_Category"]].copy()

    # Monthly actual override for YTD comparable
    monthly_used = False
    monthly_dept_revenue = pd.DataFrame()
    if mode == "YTD Comparable" and "MONTHLY_KPI" in sheets and not sheets["MONTHLY_KPI"].empty:
        mk = sheets["MONTHLY_KPI"].copy()
        if {"Year", "Month_No"}.issubset(mk.columns):
            mk["Year"] = pd.to_numeric(mk["Year"], errors="coerce").astype("Int64")
            mk["Month_No"] = pd.to_numeric(mk["Month_No"], errors="coerce").astype("Int64")
            mk = mk[(mk["Month_No"] >= 1) & (mk["Month_No"] <= comparable_months)].copy()
            for c in ["Revenue_Recognized", "COGS_Direct", "Avg_Headcount", "New_Hires", "Exits", "Payroll_Gross", "Overtime", "Bonus", "Benefits", "Employer_Tax", "Total_Manpower_Cost"]:
                if c not in mk.columns:
                    mk[c] = 0.0
                mk[c] = pd.to_numeric(mk[c], errors="coerce").fillna(0.0)
            if "Dept_ID" not in mk.columns:
                mk["Dept_ID"] = ""
            mk["Dept_ID"] = mk["Dept_ID"].astype(str).str.strip()
            if "Business_Line" not in mk.columns:
                mk["Business_Line"] = "Total Company"

            rev_m = mk[(mk["Revenue_Recognized"] != 0) | (mk["COGS_Direct"] != 0)].copy()
            if not rev_m.empty:
                rev = rev_m.groupby(["Year", "Business_Line"], as_index=False).agg(
                    Revenue_Recognized=("Revenue_Recognized", "sum"),
                    COGS_Direct=("COGS_Direct", "sum"),
                )
                rev["Gross_Profit"] = rev["Revenue_Recognized"] - rev["COGS_Direct"]
                monthly_used = True
                drr = rev_m[rev_m["Dept_ID"].ne("") & ~rev_m["Dept_ID"].str.upper().isin(["COMPANY_TOTAL", "TOTAL", "TOTAL COMPANY"])].copy()
                if not drr.empty:
                    monthly_dept_revenue = drr.groupby(["Year", "Dept_ID"], as_index=False).agg(Dept_Revenue=("Revenue_Recognized", "sum"))

            dm = mk[mk["Dept_ID"].ne("") & ~mk["Dept_ID"].str.upper().isin(["COMPANY_TOTAL", "TOTAL", "TOTAL COMPANY"])].copy()
            if not dm.empty and dm["Avg_Headcount"].sum() > 0:
                hc = dm.groupby(["Year", "Dept_ID"], as_index=False).agg(
                    Avg_Headcount=("Avg_Headcount", "mean"),
                    Avg_FTE=("Avg_Headcount", "mean"),
                    New_Hires=("New_Hires", "sum"),
                    Exits=("Exits", "sum"),
                )
            if not dm.empty and dm["Total_Manpower_Cost"].sum() > 0:
                pay = dm.groupby(["Year", "Dept_ID"], as_index=False).agg(
                    Payroll_Gross=("Payroll_Gross", "sum"),
                    Overtime=("Overtime", "sum"),
                    Bonus=("Bonus", "sum"),
                    Benefits=("Benefits", "sum"),
                    Employer_Tax=("Employer_Tax", "sum"),
                    Total_Manpower_Cost=("Total_Manpower_Cost", "sum"),
                )

    if monthly_used:
        mode_note = f"YTD Comparable memakai MONTHLY_KPI actual Jan-{comparable_months}. KPI, Board Score, dan Closing dihitung dari realisasi YTD aktual."
    elif mode == "YTD Comparable":
        mode_note = f"YTD Comparable memakai proxy FY x {comparable_months}/12 karena MONTHLY_KPI belum aktif/terisi."

    # Revenue
    rev["Year"] = pd.to_numeric(rev["Year"], errors="coerce").astype("Int64")
    for c in ["Revenue_Recognized", "COGS_Direct"]:
        if c not in rev.columns:
            rev[c] = 0.0
        rev[c] = pd.to_numeric(rev[c], errors="coerce").fillna(0.0)
    rev["Revenue_Actual"] = rev["Revenue_Recognized"]
    rev["COGS_Actual"] = rev["COGS_Direct"]
    if not monthly_used and mode in ["YTD Comparable", "Annual Projection", "MCR Focus"]:
        rev["Revenue_Recognized"] = rev["Revenue_Recognized"] * rev["Year"].astype(float).map(rev_map).fillna(1.0)
        rev["COGS_Direct"] = rev["COGS_Direct"] * rev["Year"].astype(float).map(cost_map).fillna(1.0)

    # Headcount
    hc["Year"] = pd.to_numeric(hc["Year"], errors="coerce").astype("Int64")
    hc["Dept_ID"] = hc["Dept_ID"].astype(str).str.strip()
    hc["Avg_Headcount"] = pd.to_numeric(hc.get("Avg_Headcount"), errors="coerce").fillna(0.0)
    for c in ["New_Hires", "Exits", "Avg_FTE"]:
        if c in hc.columns:
            hc[c] = pd.to_numeric(hc[c], errors="coerce").fillna(0.0)

    # Payroll
    pay["Year"] = pd.to_numeric(pay["Year"], errors="coerce").astype("Int64")
    pay["Dept_ID"] = pay["Dept_ID"].astype(str).str.strip()
    for c in ["Payroll_Gross", "Overtime", "Bonus", "Benefits", "Employer_Tax", "Total_Manpower_Cost"]:
        if c in pay.columns:
            pay[c] = pd.to_numeric(pay[c], errors="coerce").fillna(0.0)
    if "Total_Manpower_Cost" not in pay.columns:
        cols = [c for c in ["Payroll_Gross", "Overtime", "Bonus", "Benefits", "Employer_Tax"] if c in pay.columns]
        pay["Total_Manpower_Cost"] = pay[cols].sum(axis=1) if cols else 0.0
    pay["Total_Manpower_Cost_Actual"] = pay["Total_Manpower_Cost"]
    if not monthly_used and mode in ["YTD Comparable", "Annual Projection"]:
        for c in ["Payroll_Gross", "Overtime", "Bonus", "Benefits", "Employer_Tax", "Total_Manpower_Cost"]:
            if c in pay.columns:
                pay[c] = pay[c] * pay["Year"].astype(float).map(cost_map).fillna(1.0)

    hc2 = hc.merge(dim_small, on="Dept_ID", how="left")
    pay2 = pay.merge(dim_small, on="Dept_ID", how="left")

    rev_tot = rev.groupby("Year", as_index=False).agg(
        Total_Revenue=("Revenue_Recognized", "sum"),
        Total_COGS=("COGS_Direct", "sum"),
        Revenue_Actual=("Revenue_Actual", "sum"),
        COGS_Actual=("COGS_Actual", "sum"),
    )
    rev_tot["Gross_Profit"] = rev_tot["Total_Revenue"] - rev_tot["Total_COGS"]
    rev_tot["Gross_Margin_Pct"] = np.where(rev_tot["Total_Revenue"] > 0, rev_tot["Gross_Profit"] / rev_tot["Total_Revenue"], np.nan)

    hc_tot = hc2.groupby("Year", as_index=False).agg(
        Total_Headcount=("Avg_Headcount", "sum"),
        New_Hires=("New_Hires", "sum") if "New_Hires" in hc2.columns else ("Avg_Headcount", lambda x: np.nan),
        Exits=("Exits", "sum") if "Exits" in hc2.columns else ("Avg_Headcount", lambda x: np.nan),
    )
    pay_tot = pay2.groupby("Year", as_index=False).agg(
        Total_Manpower_Cost=("Total_Manpower_Cost", "sum"),
        Total_Manpower_Cost_Actual=("Total_Manpower_Cost_Actual", "sum"),
    )

    yr = rev_tot.merge(hc_tot, on="Year", how="outer").merge(pay_tot, on="Year", how="outer").sort_values("Year")
    yr["RPE"] = np.where(yr["Total_Headcount"] > 0, yr["Total_Revenue"] / yr["Total_Headcount"], np.nan)
    yr["Cost_per_HC"] = np.where(yr["Total_Headcount"] > 0, yr["Total_Manpower_Cost"] / yr["Total_Headcount"], np.nan)
    yr["MCR_Pct"] = np.where(yr["Total_Revenue"] > 0, yr["Total_Manpower_Cost"] / yr["Total_Revenue"], np.nan)
    yr["Revenue_per_Payroll"] = np.where(yr["Total_Manpower_Cost"] > 0, yr["Total_Revenue"] / yr["Total_Manpower_Cost"], np.nan)

    for c, out in [
        ("Total_Revenue", "Revenue_YoY"),
        ("Total_Headcount", "Headcount_YoY"),
        ("Total_Manpower_Cost", "ManpowerCost_YoY"),
        ("RPE", "RPE_YoY"),
        ("Gross_Margin_Pct", "GrossMargin_Delta"),
        ("MCR_Pct", "MCR_Delta"),
    ]:
        yr[out] = yr[c].diff() if "Delta" in out else yr[c].pct_change()

    # FJA
    fja_hc = hc2.groupby(["Year", "FJA_Category"], dropna=False, as_index=False).agg(Headcount=("Avg_Headcount", "sum"))
    fja_pay = pay2.groupby(["Year", "FJA_Category"], dropna=False, as_index=False).agg(Manpower_Cost=("Total_Manpower_Cost", "sum"))
    fja = fja_hc.merge(fja_pay, on=["Year", "FJA_Category"], how="outer")
    fja["Cost_Share"] = fja["Manpower_Cost"] / fja.groupby("Year")["Manpower_Cost"].transform("sum")
    fja["HC_Share"] = fja["Headcount"] / fja.groupby("Year")["Headcount"].transform("sum")

    # Department
    hcd = hc2.groupby(["Year", "Dept_ID", "Dept_Name", "Function_Group", "FJA_Category"], dropna=False, as_index=False).agg(Headcount=("Avg_Headcount", "sum"))
    payd = pay2.groupby(["Year", "Dept_ID", "Dept_Name", "Function_Group", "FJA_Category"], dropna=False, as_index=False).agg(Manpower_Cost=("Total_Manpower_Cost", "sum"))
    dept = hcd.merge(payd, on=["Year", "Dept_ID", "Dept_Name", "Function_Group", "FJA_Category"], how="outer")
    if not monthly_dept_revenue.empty:
        dept = dept.merge(monthly_dept_revenue, on=["Year", "Dept_ID"], how="left")
    elif "REVENUE_BY_DEPT" in sheets and not sheets["REVENUE_BY_DEPT"].empty:
        rd = sheets["REVENUE_BY_DEPT"].copy()
        rd["Year"] = pd.to_numeric(rd["Year"], errors="coerce").astype("Int64")
        rd["Dept_ID"] = rd["Dept_ID"].astype(str).str.strip()
        rd["Revenue_Recognized"] = pd.to_numeric(rd.get("Revenue_Recognized"), errors="coerce").fillna(0.0)
        if mode in ["YTD Comparable", "Annual Projection", "MCR Focus"]:
            rd["Revenue_Recognized"] = rd["Revenue_Recognized"] * rd["Year"].astype(float).map(rev_map).fillna(1.0)
        dept_rev = rd.groupby(["Year", "Dept_ID"], as_index=False).agg(Dept_Revenue=("Revenue_Recognized", "sum"))
        dept = dept.merge(dept_rev, on=["Year", "Dept_ID"], how="left")
    else:
        dept["Dept_Revenue"] = np.nan
    dept["Dept_RPE"] = np.where(dept["Headcount"] > 0, dept["Dept_Revenue"] / dept["Headcount"], np.nan)
    dept["Revenue_per_Cost"] = np.where(dept["Manpower_Cost"] > 0, dept["Dept_Revenue"] / dept["Manpower_Cost"], np.nan)
    dept["Cost_per_HC"] = np.where(dept["Headcount"] > 0, dept["Manpower_Cost"] / dept["Headcount"], np.nan)

    # Capacity
    if "CAPACITY_INDICATOR" in sheets and not sheets["CAPACITY_INDICATOR"].empty:
        cap = sheets["CAPACITY_INDICATOR"].copy()
        cap["Year"] = pd.to_numeric(cap["Year"], errors="coerce").astype("Int64")
        for c in ["Avg_Utilization_Pct", "Overtime_Hours", "Backlog_Count", "SLA_Breach_Count", "Incident_Count", "Turnover_Count"]:
            if c in cap.columns:
                cap[c] = pd.to_numeric(cap[c], errors="coerce").fillna(0.0)
    else:
        cap = pd.DataFrame()

    # Project
    if "PROJECT_MARGIN" in sheets and not sheets["PROJECT_MARGIN"].empty:
        pm = sheets["PROJECT_MARGIN"].copy()
        pm["Year"] = pd.to_numeric(pm["Year"], errors="coerce").astype("Int64")
        for c in ["Revenue", "COGS", "Gross_Profit"]:
            if c in pm.columns:
                pm[c] = pd.to_numeric(pm[c], errors="coerce").fillna(0.0)
        if mode in ["YTD Comparable", "Annual Projection", "MCR Focus"] and not monthly_used:
            factor = pm["Year"].astype(float).map(rev_map).fillna(1.0)
            for c in ["Revenue", "COGS", "Gross_Profit"]:
                if c in pm.columns:
                    pm[c] = pm[c] * factor
        if "Gross_Profit" not in pm.columns:
            pm["Gross_Profit"] = pm["Revenue"] - pm["COGS"]
        pm["Gross_Margin_Pct"] = np.where(pm["Revenue"] > 0, pm["Gross_Profit"] / pm["Revenue"], np.nan)
    else:
        pm = pd.DataFrame()

    return {
        "yearly": yr, "fja": fja, "dept": dept, "capacity": cap, "project_margin": pm,
        "calendar": calendar, "mode_note": mode_note, "comparable_months": comparable_months,
        "monthly_ytd_used": monthly_used,
    }


# =========================
# ENGINES
# =========================
def business_posture(row, targets):
    rev_g = row.get("Revenue_YoY")
    hc_g = row.get("Headcount_YoY")
    rpe_g = row.get("RPE_YoY")
    mcr = row.get("MCR_Pct")
    rev_min = get_target(targets, "Revenue_Growth_Min", 0.20)
    rpe_min = get_target(targets, "RPE_Growth_Min", 0.15)
    hc_max = get_target(targets, "Headcount_Growth_Max_High_Leverage", 0.10)
    mcr_watch = get_target(targets, "MCR_Watch_Max", 0.09)
    if pd.notna(rev_g) and rev_g <= 0:
        return {"state": "DEFENSIVE", "tone": "error", "hint": "Revenue melemah; proteksi margin, cash, dan fungsi kritikal."}
    if pd.notna(rev_g) and pd.notna(rpe_g) and pd.notna(hc_g) and pd.notna(mcr) and rev_g >= rev_min and rpe_g >= rpe_min and hc_g <= hc_max and mcr <= mcr_watch:
        return {"state": "HIGH_LEVERAGE", "tone": "success", "hint": "Revenue & RPE tumbuh dengan HC terkendali dan MCR sehat/watch."}
    if pd.notna(mcr) and mcr > mcr_watch:
        return {"state": "COST_PRESSURE", "tone": "warning", "hint": "People-cost mulai menekan revenue."}
    if pd.notna(rev_g) and pd.notna(hc_g) and rev_g > 0.10 and hc_g > 0.15 and (pd.isna(rpe_g) or rpe_g <= 0):
        return {"state": "CAPACITY_RISK", "tone": "warning", "hint": "Growth mulai dibeli dengan tambahan kapasitas."}
    return {"state": "BALANCED", "tone": "info", "hint": "Revenue, capacity, dan people-cost relatif seimbang."}


def margin_quality(gm, targets):
    strong = get_target(targets, "Gross_Margin_Strong_Min", 0.30)
    healthy = get_target(targets, "Gross_Margin_Healthy_Min", 0.20)
    watch = get_target(targets, "Gross_Margin_Watch_Min", 0.15)
    pressure = get_target(targets, "Gross_Margin_Pressure_Min", 0.10)
    if pd.isna(gm):
        return {"state": "UNKNOWN", "tone": "info", "hint": "Gross Margin belum tersedia."}
    if gm >= strong:
        return {"state": "STRONG", "tone": "success", "hint": "Margin sangat kuat."}
    if gm >= healthy:
        return {"state": "HEALTHY", "tone": "success", "hint": "Margin sehat untuk telco/IT integrator project-based."}
    if gm >= watch:
        return {"state": "WATCH", "tone": "warning", "hint": "Margin perlu dipantau."}
    if gm >= pressure:
        return {"state": "PRESSURE", "tone": "warning", "hint": "Margin mulai tertekan."}
    return {"state": "CRITICAL", "tone": "error", "hint": "Margin kritikal."}


def mcr_health(mcr, targets):
    ultra = get_target(targets, "MCR_Ultra_Efficiency_Max", 0.05)
    healthy = get_target(targets, "MCR_Healthy_Max", 0.07)
    watch = get_target(targets, "MCR_Watch_Max", 0.09)
    pressure = get_target(targets, "MCR_Cost_Pressure_Max", 0.14)
    if pd.isna(mcr):
        return {"state": "UNKNOWN", "tone": "info", "hint": "MCR belum tersedia."}
    if mcr < ultra:
        return {"state": "ULTRA_EFFICIENCY", "tone": "warning", "hint": "Sangat efisien; validasi capacity sustainability."}
    if mcr <= healthy:
        return {"state": "HEALTHY", "tone": "success", "hint": "MCR sehat 5–7%."}
    if mcr <= watch:
        return {"state": "WATCH", "tone": "warning", "hint": "MCR mulai perlu dipantau."}
    if mcr <= pressure:
        return {"state": "COST_PRESSURE", "tone": "warning", "hint": "People-cost menekan revenue."}
    return {"state": "CRITICAL", "tone": "error", "hint": "People-cost kritikal."}


def capacity_risk(row, cap_df):
    year = int(row.get("Year"))
    d = {"max_utilization": np.nan, "total_overtime": 0, "total_backlog": 0, "total_sla_breach": 0, "total_turnover": 0, "critical_dependency": "UNKNOWN"}
    if cap_df is None or cap_df.empty:
        return {"state": "LOW", "tone": "success", "score": 0, "recommendation": "No Hiring Required", "hint": "CAPACITY_INDICATOR belum tersedia.", "details": d}
    c = cap_df[cap_df["Year"] == year].copy()
    if c.empty:
        return {"state": "LOW", "tone": "success", "score": 0, "recommendation": "No Hiring Required", "hint": "Data capacity tahun terpilih belum tersedia.", "details": d}
    for col in ["Avg_Utilization_Pct", "Overtime_Hours", "Backlog_Count", "SLA_Breach_Count", "Turnover_Count"]:
        if col not in c.columns:
            c[col] = 0
        c[col] = pd.to_numeric(c[col], errors="coerce").fillna(0)
    max_util = float(c["Avg_Utilization_Pct"].max())
    overtime = float(c["Overtime_Hours"].sum())
    backlog = float(c["Backlog_Count"].sum())
    sla = float(c["SLA_Breach_Count"].sum())
    turnover = float(c["Turnover_Count"].sum())
    dep = c.get("Critical_Role_Dependency", pd.Series(dtype=str)).astype(str).str.upper().str.strip()
    dep_value = {1: "LOW", 2: "MEDIUM", 3: "HIGH", 4: "CRITICAL"}.get(int(dep.map({"LOW": 1, "MEDIUM": 2, "HIGH": 3, "CRITICAL": 4}).fillna(0).max()), "UNKNOWN") if len(dep) else "UNKNOWN"
    score, reasons = 0, []
    if max_util >= 0.90: score += 25; reasons.append("Utilization >90%")
    elif max_util >= 0.85: score += 18; reasons.append("Utilization 85-90%")
    elif max_util >= 0.75: score += 10; reasons.append("Utilization 75-85%")
    if overtime > 400: score += 20; reasons.append("Overtime >400 jam")
    elif overtime > 250: score += 15; reasons.append("Overtime 250-400 jam")
    elif overtime > 100: score += 8; reasons.append("Overtime 100-250 jam")
    if backlog >= 20: score += 20; reasons.append("Backlog >=20")
    elif backlog >= 10: score += 12; reasons.append("Backlog 10-19")
    elif backlog > 0: score += 6; reasons.append("Backlog muncul")
    if sla >= 5: score += 15; reasons.append("SLA breach >=5")
    elif sla > 0: score += 8; reasons.append("Ada SLA breach")
    if turnover >= 10: score += 10; reasons.append("Turnover >=10")
    elif turnover > 0: score += 5; reasons.append("Ada turnover")
    if dep_value in ["HIGH", "CRITICAL"]: score += 10; reasons.append("Critical dependency tinggi")
    elif dep_value == "MEDIUM": score += 5; reasons.append("Critical dependency medium")
    score = int(min(100, round(score)))
    if score > 75: state, tone, rec = "CRITICAL", "error", "Immediate Capacity Action Required"
    elif score >= 56: state, tone, rec = "HIGH", "warning", "Capacity Expansion Review"
    elif score >= 31: state, tone, rec = "MEDIUM", "warning", "Selective Hiring"
    else: state, tone, rec = "LOW", "success", "No Hiring Required"
    d = {"max_utilization": max_util, "total_overtime": overtime, "total_backlog": backlog, "total_sla_breach": sla, "total_turnover": turnover, "critical_dependency": dep_value}
    return {"state": state, "tone": tone, "score": score, "recommendation": rec, "hint": "; ".join(reasons) if reasons else "Tidak ada sinyal capacity besar.", "details": d}


def score_from_state(state):
    return {"STRONG": 100, "HEALTHY": 85, "WATCH": 60, "PRESSURE": 35, "CRITICAL": 10, "UNKNOWN": 50}.get(state, 50)


def score_mcr(state):
    return {"HEALTHY": 95, "WATCH": 75, "ULTRA_EFFICIENCY": 70, "COST_PRESSURE": 40, "CRITICAL": 10, "UNKNOWN": 50}.get(state, 50)


def score_growth(v, strong=0.20, good=0.10):
    if pd.isna(v): return 50
    if v >= strong: return 100
    if v >= good: return 80
    if v >= 0: return 60
    if v >= -0.10: return 35
    return 15


def score_capacity(state):
    return {"LOW": 100, "MEDIUM": 75, "HIGH": 40, "CRITICAL": 10}.get(state, 50)


def board_score(row, targets, mq, mh, cr):
    comps = {
        "Revenue Growth": score_growth(row.get("Revenue_YoY"), get_target(targets, "Revenue_Growth_Min", 0.20), 0.10),
        "RPE Growth": score_growth(row.get("RPE_YoY"), get_target(targets, "RPE_Growth_Min", 0.15), 0.08),
        "Margin Quality": score_from_state(mq["state"]),
        "MCR Health": score_mcr(mh["state"]),
        "Capacity Risk": score_capacity(cr["state"]),
    }
    weights = {"Revenue Growth": 0.25, "RPE Growth": 0.25, "Margin Quality": 0.20, "MCR Health": 0.15, "Capacity Risk": 0.15}
    score = int(round(sum(comps[k] * weights[k] for k in comps)))
    if score >= 85: label, tone, icon = "Growth Ready", "success", "🟢"
    elif score >= 70: label, tone, icon = "Optimize Growth", "info", "🔵"
    elif score >= 55: label, tone, icon = "Selective Hiring", "warning", "🟡"
    elif score >= 40: label, tone, icon = "Capacity Alert", "warning", "🟠"
    else: label, tone, icon = "Defensive Mode", "error", "🔴"
    return {"score": score, "label": label, "tone": tone, "icon": icon, "components": comps, "weights": weights}


def management_action(bp, mq, mh, cr):
    if bp["state"] == "HIGH_LEVERAGE":
        title = "Selective Scale-Up with Capacity Guardrail"
        actions = ["Pertahankan disiplin HC.", "Selective hiring pada bottleneck revenue/delivery.", "Monitor overtime, backlog, SLA, dan role dependency."]
    elif bp["state"] == "DEFENSIVE":
        title = "Margin & Cash Protection"
        actions = ["Freeze hiring non-critical.", "Lindungi fungsi revenue/delivery kritikal.", "Fokus cash discipline dan margin recovery."]
    elif bp["state"] == "COST_PRESSURE":
        title = "People-Cost Control"
        actions = ["Perketat approval hiring berbasis ROI.", "Audit overtime, bonus, benefit, dan manpower cost.", "Review unit economics project/service."]
    else:
        title = "Balanced Optimization"
        actions = ["Optimalkan struktur sebelum ekspansi.", "Hiring hanya pada bottleneck terbukti.", "Jaga MCR, RPE, dan support ratio."]
    if mq["state"] in ["WATCH", "PRESSURE", "CRITICAL"]:
        actions.append("Jalankan margin quality review.")
    if cr["state"] in ["HIGH", "CRITICAL"]:
        actions.append("Lakukan capacity risk review bulanan.")
    return title, actions


def build_management_insights(row, bp, mq, mh, cr):
    insights = []
    rev_g, hc_g, rpe_g = row.get("Revenue_YoY"), row.get("Headcount_YoY"), row.get("RPE_YoY")
    if pd.notna(rev_g) and pd.notna(hc_g):
        if rev_g > hc_g:
            insights.append(("Revenue grows faster than Headcount", "GOOD", "Revenue tumbuh lebih cepat dari HC.", "Growth relatif produktif."))
        else:
            insights.append(("Headcount grows faster than Revenue", "WATCH", "HC tumbuh lebih cepat dari revenue.", "Tambahan kapasitas belum sepenuhnya menjadi revenue."))
    if pd.notna(rev_g) and pd.notna(rpe_g):
        if rev_g > 0 and rpe_g < 0:
            insights.append(("Revenue rises but RPE declines", "WATCH", "Revenue naik tetapi RPE turun.", "Growth kemungkinan dibeli dengan tambahan HC."))
        elif rev_g > 0 and rpe_g > 0:
            insights.append(("Revenue and RPE improve together", "GOOD", "Revenue dan RPE sama-sama tumbuh.", "Pertumbuhan berkualitas dari sisi productivity."))
    if mh["state"] == "ULTRA_EFFICIENCY" and cr["state"] in ["HIGH", "CRITICAL"]:
        insights.append(("Low MCR but high capacity risk", "CRITICAL", "MCR rendah tetapi capacity risk tinggi.", "Efisiensi bisa menyembunyikan overload."))
    if mq["state"] in ["WATCH", "PRESSURE", "CRITICAL"]:
        insights.append(("Margin quality needs attention", "WATCH", f"Gross Margin {mq['state']}.", "Revenue perlu divalidasi kualitas marginnya."))
    if cr["state"] in ["HIGH", "CRITICAL"]:
        insights.append(("Capacity bottleneck may limit growth", "CRITICAL", f"Capacity Risk {cr['state']}.", "Growth berikutnya bisa tertahan delivery capacity."))
    return insights or [("No major conflict detected", "BALANCED", "Indikator utama relatif seimbang.", "Lanjutkan monitoring bulanan.")]


def hiring_decision(dept_df, cap_df, year):
    if dept_df is None or dept_df.empty:
        return pd.DataFrame()
    d = dept_df[dept_df["Year"] == int(year)].copy()
    if d.empty:
        return pd.DataFrame()
    if cap_df is not None and not cap_df.empty and "Dept_ID" in cap_df.columns:
        cap = cap_df[cap_df["Year"] == int(year)].copy()
        for c in ["Avg_Utilization_Pct", "Overtime_Hours", "Backlog_Count", "SLA_Breach_Count", "Turnover_Count"]:
            if c not in cap.columns: cap[c] = 0
            cap[c] = pd.to_numeric(cap[c], errors="coerce").fillna(0)
        sig = cap.groupby("Dept_ID", as_index=False).agg(Avg_Utilization=("Avg_Utilization_Pct", "max"), Overtime=("Overtime_Hours", "sum"), Backlog=("Backlog_Count", "sum"), SLA_Breach=("SLA_Breach_Count", "sum"), Turnover=("Turnover_Count", "sum"))
        d = d.merge(sig, on="Dept_ID", how="left")
    else:
        for c in ["Avg_Utilization", "Overtime", "Backlog", "SLA_Breach", "Turnover"]:
            d[c] = 0
    d[["Avg_Utilization", "Overtime", "Backlog", "SLA_Breach", "Turnover"]] = d[["Avg_Utilization", "Overtime", "Backlog", "SLA_Breach", "Turnover"]].fillna(0)
    recs, reasons = [], []
    for _, r in d.iterrows():
        fja = str(r.get("FJA_Category", ""))
        rpc = r.get("Revenue_per_Cost")
        util, ot, bl, sla = r.get("Avg_Utilization", 0), r.get("Overtime", 0), r.get("Backlog", 0), r.get("SLA_Breach", 0)
        high_capacity = util >= 0.85 or ot > 250 or bl >= 10 or sla > 0
        efficient = pd.notna(rpc) and rpc >= 4
        low_eff = pd.notna(rpc) and rpc < 1.5
        if high_capacity and ("Revenue Enabler" in fja or "Revenue Generator" in fja or efficient):
            rec, reason = "HIRE", "Bottleneck/capacity pressure pada fungsi revenue/delivery."
        elif high_capacity:
            rec, reason = "SELECTIVE", "Ada tekanan kapasitas, perlu validasi workload dan kontribusi."
        elif low_eff and "Support" in fja:
            rec, reason = "FREEZE", "Support cost intensity tinggi; optimasi proses dahulu."
        elif low_eff:
            rec, reason = "HOLD", "Revenue per cost belum kuat."
        else:
            rec, reason = "HOLD", "Belum ada sinyal kuat untuk tambah HC."
        recs.append(rec); reasons.append(reason)
    d["Recommendation"] = recs
    d["Reason"] = reasons
    return d


def hc_guardrails(row, targets):
    revenue = _safe_float(row.get("Total_Revenue"))
    cost = _safe_float(row.get("Total_Manpower_Cost"))
    hc = _safe_float(row.get("Total_Headcount"))
    cph = _safe_float(row.get("Cost_per_HC"))
    target = get_target(targets, "MCR_Healthy_Max", 0.07)
    max_cost = revenue * target if pd.notna(revenue) else np.nan
    remaining = max_cost - cost if pd.notna(max_cost) and pd.notna(cost) else np.nan
    add_hc = np.floor(remaining / cph) if pd.notna(remaining) and pd.notna(cph) and cph > 0 else np.nan
    max_hc = hc + add_hc if pd.notna(hc) and pd.notna(add_hc) else np.nan
    req_rev = cost / target if pd.notna(cost) and target > 0 else np.nan
    return {"target_mcr": target, "remaining_budget": remaining, "additional_hc_capacity": add_hc, "max_hc": max_hc, "required_revenue": req_rev}


def roadmap(bp, mq, mh, cr):
    rows = []
    def add(phase, priority, action, owner, metric):
        rows.append({"Phase": phase, "Priority": priority, "Action": action, "Owner": owner, "Success Metric": metric})
    if cr["state"] in ["HIGH", "CRITICAL"]:
        add("0-3 Months", "Critical", "Capacity review: overtime, backlog, SLA, dependency.", "COO / HR", "Capacity score turun 15 poin")
    if mh["state"] in ["COST_PRESSURE", "CRITICAL"]:
        add("0-3 Months", "Critical", "Freeze non-critical hiring dan review manpower cost.", "CFO / HR", "MCR turun menuju <=9%")
    if mq["state"] in ["WATCH", "PRESSURE", "CRITICAL"]:
        add("0-3 Months", "High", "Margin quality review per project/customer.", "CFO / Sales / Ops", "Gross Margin kembali sehat")
    add("3-6 Months", "High", "Selective hiring pada fungsi revenue/delivery bottleneck.", "CEO / HR / COO", "Hiring berbasis ROI dan evidence")
    add("3-6 Months", "Medium", "Productivity program: SOP, automation, redistribusi workload.", "HR / Dept Head", "RPE naik, overtime turun")
    add("6-12 Months", "Medium", "Review ratio revenue generator, enabler, support.", "CEO / HR", "Support cost share terkendali")
    return pd.DataFrame(rows)


def closing_text(row, bs, bp, mq, mh, cr, targets, currency):
    guard = hc_guardrails(row, targets)
    return (
        f"Perusahaan berada pada status **{bs['label']}** dengan Company Health Score **{bs['score']}/100**. "
        f"Revenue terbaca **{_money(row.get('Total_Revenue'), currency)}**, MCR **{_pct(row.get('MCR_Pct'))}**, "
        f"RPE **{_money(row.get('RPE'), currency)}**, Gross Margin **{_pct(row.get('Gross_Margin_Pct'))}**, "
        f"dan Capacity Risk **{cr['state']}**. Dengan target MCR sehat **{_pct(guard['target_mcr'])}**, "
        f"revenue minimum untuk menopang manpower cost saat ini adalah **{_money(guard['required_revenue'], currency)}**. "
        f"Rekomendasi: lakukan hiring selektif pada fungsi bottleneck, tahan support non-critical, dan gunakan MCR sebagai guardrail."
    )


def make_priority_df(bp, mq, mh, cr):
    def sev(x):
        s = str(x).upper()
        if s in ["CRITICAL", "DEFENSIVE", "COST_PRESSURE", "HIGH"]: return 90
        if s in ["WATCH", "PRESSURE", "MEDIUM", "CAPACITY_RISK", "ULTRA_EFFICIENCY"]: return 65
        if s in ["HEALTHY", "BALANCED"]: return 35
        if s in ["STRONG", "LOW", "HIGH_LEVERAGE"]: return 20
        return 50
    return pd.DataFrame([
        {"Area": "Business Posture", "Status": bp["state"], "Priority": sev(bp["state"])},
        {"Area": "Margin Quality", "Status": mq["state"], "Priority": sev(mq["state"])},
        {"Area": "MCR Health", "Status": mh["state"], "Priority": sev(mh["state"])},
        {"Area": "Capacity Risk", "Status": cr["state"], "Priority": sev(cr["state"])},
    ])


def framework_from_targets(targets_df, targets):
    if targets_df is not None and not targets_df.empty and "KPI" in targets_df.columns:
        base = targets_df.copy()
    else:
        base = pd.DataFrame()
    rows = [
        ["Revenue Growth", "15–25% healthy; >25% strong", "Annual reports + management target", "Measures growth momentum", "High"],
        ["RPE Growth", ">=15% strong", "Human capital productivity benchmark", "Measures revenue productivity per employee", "Medium"],
        ["Headcount Growth", "<=10% controlled", "Organization design / workforce planning practice", "Controls fixed cost expansion", "Medium"],
        ["MCR", "5–7% healthy; 7–9% watch", "Internal benchmark for project-based telco/IT integrator", "Measures people-cost leverage", "Medium"],
        ["Gross Margin", "20–30% healthy; >30% strong", "Telecom/IT integrator project-margin benchmark", "Measures project revenue quality", "Medium"],
        ["Capacity Utilization", "75–85% optimal; >90% risk", "Capacity planning and service operations practice", "Measures overload risk", "Medium"],
    ]
    return pd.DataFrame(rows, columns=["KPI", "Target Framework", "Benchmark Basis", "Why It Matters", "Confidence"])


# =========================
# SIDEBAR & LOAD
# =========================
with st.sidebar:
    st.header("⚙️ Data Source")
    uploaded = st.file_uploader("Upload Excel SEP V4", type=["xlsx"])
    st.caption("Jika tidak upload, app membaca file default dari folder data/.")

file_bytes = uploaded.getvalue() if uploaded is not None else None
sheets, warnings = load_workbook(file_bytes, DEFAULT_FILE)

render_header()
for w in warnings:
    st.warning(w)
if any(s not in sheets for s in REQUIRED_SHEETS):
    st.error("Data belum lengkap. Sheet wajib: " + ", ".join(REQUIRED_SHEETS))
    st.stop()

# Calendar preview
_years = []
for s in ["REVENUE_YR", "HEADCOUNT_YR", "PAYROLL_YR"]:
    if s in sheets and "Year" in sheets[s].columns:
        _years += pd.to_numeric(sheets[s]["Year"], errors="coerce").dropna().astype(int).tolist()
calendar_preview = build_calendar(sheets.get("CALENDAR", pd.DataFrame()), sorted(set(_years)))

with st.sidebar:
    st.header("🧮 Analysis Mode")
    default_idx = 1 if not calendar_preview.empty and (calendar_preview["Months_Closed"] < 12).any() else 0
    mode = st.radio("Mode", ["Actual", "YTD Comparable", "Annual Projection", "MCR Focus"], index=default_idx)

data = prep_data(sheets, mode)
params = load_params(sheets.get("PARAMETERS", pd.DataFrame()))
targets = load_targets(sheets.get("TARGETS", pd.DataFrame()))
currency = params.get("Currency_Code", "IDR")

yr = data["yearly"].dropna(subset=["Year"]).copy()
available_years = sorted(yr["Year"].dropna().astype(int).unique().tolist())

with st.sidebar:
    st.header("🧭 Filters")
    year_sel = st.selectbox("Analysis Year", available_years, index=len(available_years)-1)
    year_min, year_max = st.slider("Trend Range", min_value=min(available_years), max_value=max(available_years), value=(min(available_years), max(available_years)), step=1)
    fja_all = sorted(data["fja"]["FJA_Category"].dropna().unique().tolist()) if not data["fja"].empty else []
    fja_filter = st.multiselect("FJA Category", fja_all, default=fja_all)

yr_range = yr[(yr["Year"] >= year_min) & (yr["Year"] <= year_max)].copy()
latest = yr[yr["Year"] == year_sel].iloc[0]
period = get_period_info(data["calendar"], year_sel)

bp = business_posture(latest, targets)
mq = margin_quality(latest.get("Gross_Margin_Pct"), targets)
mh = mcr_health(latest.get("MCR_Pct"), targets)
cr = capacity_risk(latest, data["capacity"])
bs = board_score(latest, targets, mq, mh, cr)
action_title, actions = management_action(bp, mq, mh, cr)


# =========================
# CLEAN BOARD NAVIGATION
# =========================
boards = st.tabs([
    "1. Executive Summary",
    "2. Revenue & Productivity",
    "3. Workforce Performance",
    "4. Capacity & Delivery Risk",
    "5. Strategic Decision Center",
    "6. Executive Benchmark & Methodology",
    "7. CEO Closing Report",
    "8. Appendix",
])


# 1 Executive Summary
with boards[0]:
    st.markdown("## 1. Executive Summary")
    st.caption(f"Periode: {period['Period_Label']} | Months closed: {period['Months_Closed']}/12 | Mode: {mode} | Confidence: {period['Data_Confidence']}")
    st.info(data["mode_note"])
    if mode == "YTD Comparable" and data.get("monthly_ytd_used"):
        st.success("MONTHLY_KPI aktif: KPI dan board narrative memakai realisasi YTD aktual.")
    elif mode == "YTD Comparable":
        st.warning("MONTHLY_KPI belum aktif/terisi. YTD Comparable memakai proxy FY x bulan YTD.")

    a, b, c, d = st.columns(4)
    with a: st.markdown(badge("Company Health", f"{bs['icon']} {bs['score']}/100", bs["tone"], bs["label"]), unsafe_allow_html=True)
    with b: st.markdown(badge("Business Posture", bp["state"], bp["tone"], bp["hint"]), unsafe_allow_html=True)
    with c: st.markdown(badge("MCR Health", mh["state"], mh["tone"], mh["hint"]), unsafe_allow_html=True)
    with d: st.markdown(badge("Capacity Risk", cr["state"], cr["tone"], f"Score {cr['score']}/100"), unsafe_allow_html=True)

    k1, k2, k3 = st.columns(3)
    k4, k5, k6 = st.columns(3)
    with k1: st.metric("💰 Revenue", _money(latest.get("Total_Revenue"), currency), _pct(latest.get("Revenue_YoY")))
    with k2: st.metric("👥 Headcount", _num(latest.get("Total_Headcount")), _pct(latest.get("Headcount_YoY")))
    with k3: st.metric("⚡ RPE", _money(latest.get("RPE"), currency), _pct(latest.get("RPE_YoY")))
    with k4: st.metric("📊 MCR", _pct(latest.get("MCR_Pct")), _pp(latest.get("MCR_Delta")))
    with k5: st.metric("📈 Gross Margin", _pct(latest.get("Gross_Margin_Pct")), _pp(latest.get("GrossMargin_Delta")))
    with k6: st.metric("💼 Manpower Cost", _money(latest.get("Total_Manpower_Cost"), currency), _pct(latest.get("ManpowerCost_YoY")))

    st.markdown("#### Board Memo")
    st.info(
        f"Pada {period['Period_Label']} dalam mode **{mode}**, SEP membaca revenue **{_money(latest.get('Total_Revenue'), currency)}**, "
        f"RPE **{_money(latest.get('RPE'), currency)}**, MCR **{_pct(latest.get('MCR_Pct'))}**, dan Gross Margin **{_pct(latest.get('Gross_Margin_Pct'))}**. "
        f"Status utama adalah **{bs['label']}** dengan fokus manajemen: **{action_title}**."
    )


# 2 Revenue & Productivity
with boards[1]:
    st.markdown("## 2. Revenue & Productivity")
    st.caption("Pertanyaan bisnis: apakah revenue tumbuh bersama produktivitas, bukan hanya karena tambahan HC?")
    c1, c2 = st.columns(2)
    with c1:
        fig = px.line(yr_range, x="Year", y=["Total_Revenue", "Total_Manpower_Cost"], markers=True, title="Revenue vs Manpower Cost")
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        fig = px.line(yr_range, x="Year", y=["RPE", "Cost_per_HC"], markers=True, title="RPE vs Cost per HC")
        st.plotly_chart(fig, use_container_width=True)
    c3, c4 = st.columns(2)
    with c3:
        growth = yr_range[["Year", "Revenue_YoY", "Headcount_YoY", "RPE_YoY", "ManpowerCost_YoY"]].melt(id_vars="Year", var_name="Metric", value_name="YoY")
        fig = px.line(growth, x="Year", y="YoY", color="Metric", markers=True, title="Growth Driver Comparison")
        fig.update_yaxes(tickformat=".0%")
        st.plotly_chart(fig, use_container_width=True)
    with c4:
        fig = px.line(yr_range, x="Year", y=["Gross_Margin_Pct", "MCR_Pct"], markers=True, title="Gross Margin vs MCR")
        fig.update_yaxes(tickformat=".0%")
        st.plotly_chart(fig, use_container_width=True)


# 3 Workforce
with boards[2]:
    st.markdown("## 3. Workforce Performance")
    st.caption("Pertanyaan bisnis: apakah struktur SDM sudah optimal?")
    fja = data["fja"].copy()
    fja = fja[(fja["Year"] >= year_min) & (fja["Year"] <= year_max)]
    if fja_filter:
        fja = fja[fja["FJA_Category"].isin(fja_filter)]
    fja_latest = fja[fja["Year"] == year_sel].copy()
    w1, w2 = st.columns([1.1, 1.0])
    with w1:
        if not fja_latest.empty:
            fig = px.treemap(fja_latest, path=["FJA_Category"], values="Manpower_Cost", title="FJA Cost Mix")
            st.plotly_chart(fig, use_container_width=True)
    with w2:
        if not fja_latest.empty:
            show = fja_latest[["FJA_Category", "Manpower_Cost", "Cost_Share", "Headcount", "HC_Share"]].copy()
            show["Manpower_Cost"] = show["Manpower_Cost"].map(lambda x: _money(x, currency))
            show["Cost_Share"] = show["Cost_Share"].map(_pct)
            show["HC_Share"] = show["HC_Share"].map(_pct)
            st.dataframe(show.sort_values("Cost_Share", ascending=False), hide_index=True, use_container_width=True)
    dept = data["dept"].copy()
    dept = dept[dept["Year"] == year_sel].copy()
    if fja_filter and not dept.empty:
        dept = dept[dept["FJA_Category"].isin(fja_filter)]
    if not dept.empty:
        dept["Bubble_Size"] = pd.to_numeric(dept.get("Dept_Revenue"), errors="coerce").fillna(0)
        if dept["Bubble_Size"].max() <= 0:
            dept["Bubble_Size"] = pd.to_numeric(dept.get("Headcount"), errors="coerce").fillna(1).clip(lower=1)
        else:
            min_pos = dept.loc[dept["Bubble_Size"] > 0, "Bubble_Size"].min()
            dept["Bubble_Size"] = dept["Bubble_Size"].replace(0, min_pos * 0.1)
        fig = px.scatter(dept, x="Headcount", y="Manpower_Cost", size="Bubble_Size", color="FJA_Category", text="Dept_Name", hover_data=["Dept_Revenue", "Dept_RPE", "Revenue_per_Cost"], title="Workforce Portfolio Matrix")
        fig.update_traces(textposition="middle center", textfont=dict(size=11, color="white"), marker=dict(opacity=0.85, line=dict(width=2, color="white")))
        fig.update_layout(height=520)
        st.plotly_chart(fig, use_container_width=True)


# 4 Capacity
with boards[3]:
    st.markdown("## 4. Capacity & Delivery Risk")
    st.caption("Pertanyaan bisnis: apakah organisasi mulai overload?")
    cap_year = data["capacity"][data["capacity"]["Year"] == year_sel].copy() if not data["capacity"].empty else pd.DataFrame()
    cc1, cc2 = st.columns([1.1, 0.9])
    with cc1:
        if not cap_year.empty:
            fig = px.scatter(cap_year, x="Avg_Utilization_Pct", y="Overtime_Hours", size="Backlog_Count", color="Critical_Role_Dependency", hover_data=["Dept_ID", "SLA_Breach_Count", "Incident_Count", "Turnover_Count"], title="Capacity Risk Heatmap")
            fig.update_xaxes(tickformat=".0%")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("CAPACITY_INDICATOR belum diisi.")
    with cc2:
        fig = go.Figure(go.Indicator(
            mode="gauge+number",
            value=cr["score"],
            title={"text": f"Capacity Score - {cr['state']}"},
            gauge={"axis": {"range": [0, 100]}, "bar": {"color": TONE_COLOR.get(cr["tone"], "#3b82f6")}}
        ))
        fig.update_layout(height=320)
        st.plotly_chart(fig, use_container_width=True)
        d = cr["details"]
        st.write(f"**Max Utilization:** {_pct(d.get('max_utilization'))}")
        st.write(f"**Overtime:** {_num(d.get('total_overtime'))} | **Backlog:** {_num(d.get('total_backlog'))} | **SLA:** {_num(d.get('total_sla_breach'))}")
        st.write(f"**Recommendation:** {cr['recommendation']}")


# 5 Strategic Decision
with boards[4]:
    st.markdown("## 5. Strategic Decision Center")
    st.caption("Pertanyaan bisnis: keputusan apa yang harus diambil?")
    s1, s2 = st.columns([1.0, 1.0])
    with s1:
        priority = make_priority_df(bp, mq, mh, cr)
        fig = px.bar(priority.sort_values("Priority"), x="Priority", y="Area", color="Status", orientation="h", text="Status", title="Management Attention Map")
        fig.update_xaxes(range=[0, 100])
        st.plotly_chart(fig, use_container_width=True)
    with s2:
        comps = pd.DataFrame({"Factor": list(bs["components"].keys()), "Score": list(bs["components"].values())})
        fig = px.bar(comps.sort_values("Score"), x="Score", y="Factor", orientation="h", text="Score", title="Company Health Components")
        fig.update_xaxes(range=[0, 100])
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("### Management Insight")
    for title, status, why, meaning in build_management_insights(latest, bp, mq, mh, cr):
        if status in ["GOOD", "BALANCED"]:
            st.success(f"**{title}** — {why}\n\n{meaning}")
        elif status in ["CRITICAL"]:
            st.error(f"**{title}** — {why}\n\n{meaning}")
        else:
            st.warning(f"**{title}** — {why}\n\n{meaning}")

    hire_df = hiring_decision(data["dept"], data["capacity"], year_sel)
    if not hire_df.empty:
        st.markdown("### Hiring Decision")
        rec_summary = hire_df["Recommendation"].value_counts().reset_index()
        rec_summary.columns = ["Recommendation", "Count"]
        h1, h2 = st.columns([0.7, 1.3])
        with h1:
            fig = px.pie(rec_summary, names="Recommendation", values="Count", title="Hiring Mix")
            st.plotly_chart(fig, use_container_width=True)
        with h2:
            show = hire_df[["Dept_ID", "Dept_Name", "FJA_Category", "Headcount", "Manpower_Cost", "Dept_Revenue", "Revenue_per_Cost", "Avg_Utilization", "Recommendation", "Reason"]].copy()
            st.dataframe(show, hide_index=True, use_container_width=True)

    st.markdown("### 90-Day Action Plan")
    for i, a in enumerate(actions[:5], start=1):
        st.write(f"{i}. {a}")


# 6 Benchmark & Methodology
with boards[5]:
    st.markdown("## 6. KPI Target Framework")
    st.markdown("### Based on Telecommunications & IT System Integrator Benchmark")
    st.caption("Pertanyaan bisnis: mengapa target ini dapat dipercaya?")

    framework = framework_from_targets(sheets.get("TARGETS", pd.DataFrame()), targets)
    st.dataframe(framework, hide_index=True, use_container_width=True)

    st.markdown("### Current KPI vs Framework")
    compare = pd.DataFrame([
        {"KPI": "Revenue Growth", "Company": _pct(latest.get("Revenue_YoY")), "Framework": "15–25% healthy; >25% strong", "Status": "Strong/Watch based on trend"},
        {"KPI": "RPE Growth", "Company": _pct(latest.get("RPE_YoY")), "Framework": ">=15% strong", "Status": "Productivity growth"},
        {"KPI": "MCR", "Company": _pct(latest.get("MCR_Pct")), "Framework": "5–7% healthy; 7–9% watch", "Status": mh["state"]},
        {"KPI": "Gross Margin", "Company": _pct(latest.get("Gross_Margin_Pct")), "Framework": "20–30% healthy; >30% strong", "Status": mq["state"]},
        {"KPI": "Capacity Score", "Company": f"{cr['score']}/100", "Framework": "0–30 low; 31–55 medium; 56–75 high; 76–100 critical", "Status": cr["state"]},
    ])
    st.dataframe(compare, hide_index=True, use_container_width=True)

    st.markdown("### Methodology")
    st.info(
        "Framework ini adalah management benchmark, bukan standar regulasi. Target disusun dari kombinasi praktik industri telekomunikasi dan IT System Integrator, analisis laporan tahunan perusahaan sejenis, praktik workforce planning, serta historical performance perusahaan. Review disarankan minimal triwulanan."
    )
    st.markdown("### Benchmark Sources to Track")
    st.write("Annual reports perusahaan telco/IT integrator, APQC, Saratoga/PwC Human Capital Benchmark, Gartner, Deloitte, McKinsey, BCG, dan benchmark internal historis perusahaan.")


# 7 CEO Closing
with boards[6]:
    st.markdown("## 7. CEO Closing Report")
    st.caption("Pertanyaan bisnis: apa kesimpulan akhir dan tindakan berikutnya?")
    x1, x2 = st.columns([0.9, 1.1])
    with x1:
        fig = go.Figure(go.Indicator(
            mode="gauge+number",
            value=bs["score"],
            number={"suffix": "/100"},
            title={"text": f"Company Health - {bs['label']}"},
            gauge={"axis": {"range": [0, 100]}, "bar": {"color": TONE_COLOR.get(bs["tone"], "#3b82f6")}},
        ))
        fig.update_layout(height=360)
        st.plotly_chart(fig, use_container_width=True)
    with x2:
        st.info(closing_text(latest, bs, bp, mq, mh, cr, targets, currency))
        st.markdown("#### Top Decisions")
        roadmap_df = roadmap(bp, mq, mh, cr)
        st.dataframe(roadmap_df, hide_index=True, use_container_width=True)
    st.success("Board-ready statement: SEP V4 merekomendasikan ekspansi HC secara selektif, bukan general hiring. Prioritas utama adalah menjaga MCR, meningkatkan RPE, dan menurunkan capacity risk pada fungsi bottleneck.")


# 8 Appendix
with boards[7]:
    st.markdown("## 8. Appendix")
    st.caption("Definisi, formula, data source, dan governance framework.")
    st.markdown("""
    ### Core Formula
    - **MCR** = Total Manpower Cost / Revenue
    - **RPE** = Revenue / Average Headcount
    - **Gross Margin** = (Revenue - COGS) / Revenue
    - **Cost per HC** = Total Manpower Cost / Average Headcount
    - **Revenue per Payroll** = Revenue / Total Manpower Cost

    ### Governance
    - **Owner:** Finance + HR + Operations
    - **Review cycle:** Quarterly
    - **Approval:** CEO / Board
    - **Framework:** KPI Target Framework based on telco & IT system integrator benchmark
    - **Version:** SEP V4 Executive Intelligence Edition

    ### Data Source
    - PARAMETERS, TARGETS, CALENDAR
    - REVENUE_YR / MONTHLY_KPI
    - HEADCOUNT_YR / PAYROLL_YR
    - DIM_DEPARTMENT / FJA_MAPPING
    - CAPACITY_INDICATOR
    - PROJECT_MARGIN
    """)
    st.markdown("### Raw Data Preview")
    st.dataframe(yr, hide_index=True, use_container_width=True)

st.divider()
st.caption("Starcom Executive Platform (SEP) V4 — clean board flow: summary, productivity, workforce, capacity, decision, benchmark, closing, appendix.")
