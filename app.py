from __future__ import annotations

import math
import os
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from plotly.subplots import make_subplots

# PostgreSQL support (optional)
try:
    from sqlalchemy import create_engine, text
    HAS_DB = True
except ImportError:
    HAS_DB = False

DB_ENGINE = None
if HAS_DB and os.getenv("DATABASE_URL") or (os.getenv("PGHOST") and os.getenv("PGUSER")):
    try:
        if os.getenv("DATABASE_URL"):
            DB_ENGINE = create_engine(os.getenv("DATABASE_URL"))
        else:
            # Railway auto-injects: PGHOST, PGPORT, PGUSER, PGPASSWORD, PGDATABASE
            pghost = os.getenv("PGHOST")
            pgport = os.getenv("PGPORT", "5432")
            pguser = os.getenv("PGUSER")
            pgpass = os.getenv("PGPASSWORD")
            pgdb = os.getenv("PGDATABASE")
            DB_ENGINE = create_engine(
                f"postgresql://{pguser}:{pgpass}@{pghost}:{pgport}/{pgdb}"
            )
        # Test connection
        with DB_ENGINE.connect() as conn:
            conn.execute(text("SELECT 1"))
    except Exception:
        DB_ENGINE = None

st.set_page_config(
    page_title="Autlán — Estrategia Financiera & Riesgos",
    layout="wide",
    initial_sidebar_state="collapsed",
)

ROOT_DIR = Path(__file__).resolve().parent
DATA_DIR = ROOT_DIR / "Bases de datos"
DATABASE_FILE = DATA_DIR / "Mineria autlán historicos.xlsx"


def find_output_dir(root: Path) -> Path:
    for candidate in root.iterdir():
        if candidate.is_dir() and candidate.name.startswith("Outputs"):
            return candidate
    return root / "Outputs analisis"


OUTPUT_DIR = find_output_dir(ROOT_DIR)

COLORS = {
    "navy":      "#101319",
    "surface":   "#344055",
    "blue":      "#4ecdc4",
    "thistle":   "#dfc2f2",
    "green":     "#c9d662",
    "red":       "#f27db2",
    "coral":     "#ff875c",
    "gold":      "#ffe66d",
    "gray":      "#888098",
    "soft_gray": "#1e253a",
    "text":      "#ffffff",
    "muted":     "#888098",
}
ACCENT_SEQ = ["#4ecdc4", "#c9d662", "#ffe66d", "#ff875c", "#f27db2"]
SCENARIO_COLORS = {"Base": COLORS["blue"], "Optimista": COLORS["green"], "Adverso": COLORS["coral"]}
MONTH_ORDER = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

# Assumption confirmed by user: approx. USD 130mm of debt is USD-denominated;
# the remainder is MXN-denominated. Used to avoid double-counting rate shocks.
USD_DEBT_ASSUMPTION_M = 130.0
TOTAL_DEBT_ASSUMPTION_M = 192.0
USD_DEBT_WEIGHT = USD_DEBT_ASSUMPTION_M / TOTAL_DEBT_ASSUMPTION_M
MXN_DEBT_WEIGHT = 1.0 - USD_DEBT_WEIGHT
GOLD_SALES_WEIGHT = 0.25
GOLD_GROSS_MARGIN_BPS_PER_1PCT = 25.0
GOLD_OPERATING_MARGIN_BPS_PER_1PCT = 50.0

st.markdown(f"""
<style>
header[data-testid="stHeader"] {{
    background-color:{COLORS["navy"]} !important;
    border-bottom:1px solid rgba(223,194,242,0.08);
}}
header[data-testid="stHeader"] * {{ color:{COLORS["muted"]} !important; }}
div[data-testid="stDecoration"] {{ display:none !important; }}

body, .stApp {{ background-color:{COLORS["navy"]} !important; }}
.main .block-container {{
    padding-top:0.6rem;padding-bottom:1rem;max-width:1600px;
    background-color:{COLORS["navy"]};
}}

section[data-testid="stSidebar"] {{
    background-color:#0d1018 !important;
    border-right:1px solid rgba(223,194,242,0.10);
}}
section[data-testid="stSidebar"] * {{ color:{COLORS["text"]} !important; }}

p, li, span, div, td, th, caption, blockquote {{ color:{COLORS["text"]}; }}
h1,h2,h3,h4,h5,h6 {{ color:{COLORS["text"]} !important; }}
.stMarkdown p, .stMarkdown li, .stMarkdown td, .stMarkdown th {{
    color:{COLORS["text"]} !important;
}}
.stMarkdown table {{ border-collapse:collapse; width:100%; }}
.stMarkdown th {{
    background:{COLORS["surface"]} !important;
    color:{COLORS["thistle"]} !important;
    border:1px solid rgba(136,128,152,0.20) !important;
    padding:6px 10px;
}}
.stMarkdown td {{
    color:{COLORS["text"]} !important;
    border:1px solid rgba(136,128,152,0.15) !important;
    padding:5px 10px;
}}
.stMarkdown tr:nth-child(even) td {{ background:rgba(52,64,85,0.30); }}

button[data-baseweb="tab"] {{
    background:transparent !important;
    color:{COLORS["muted"]} !important;
    font-size:15px;
    padding:8px 14px !important;
}}
button[data-baseweb="tab"][aria-selected="true"] {{
    color:{COLORS["thistle"]} !important;
    border-bottom-color:{COLORS["thistle"]} !important;
    font-weight:700 !important;
}}
div[data-testid="stTabs"] {{ border-bottom:1px solid rgba(136,128,152,0.20); }}

details {{ background:{COLORS["surface"]}; border-radius:8px; border:1px solid rgba(223,194,242,0.10); }}
details summary {{ color:{COLORS["thistle"]} !important; padding:8px 14px; font-weight:600; font-size:14px; }}
details > div {{ color:{COLORS["text"]} !important; }}
.streamlit-expanderHeader p {{ color:{COLORS["thistle"]} !important; }}

pre {{ background:{COLORS["navy"]} !important; border:1px solid rgba(136,128,152,0.25) !important;
       border-radius:6px !important; padding:12px 14px !important; }}
pre code {{ color:{COLORS["green"]} !important; background:transparent !important;
            font-size:12px !important; }}
code:not(pre code) {{ background:rgba(52,64,85,0.85) !important;
                      color:{COLORS["gold"]} !important;
                      padding:1px 5px !important; border-radius:3px !important; }}

.stCaption p {{ color:{COLORS["muted"]} !important; font-size:13px; }}
label {{ color:{COLORS["muted"]} !important; }}

div[data-baseweb="select"] * {{ color:{COLORS["text"]} !important; background-color:transparent; }}
div[data-baseweb="select"] > div {{
    background:{COLORS["surface"]} !important;
    border-color:rgba(223,194,242,0.20) !important;
}}
li[role="option"] {{ background:{COLORS["surface"]} !important; color:{COLORS["text"]} !important; }}
li[role="option"]:hover {{ background:rgba(78,205,196,0.15) !important; }}

div[data-testid="stMetric"] {{
    background:linear-gradient(180deg,{COLORS["surface"]} 0%,{COLORS["navy"]} 100%);
    border:1px solid rgba(223,194,242,0.18);border-radius:8px;
    padding:12px 16px;box-shadow:0 1px 4px rgba(0,0,0,0.45);}}
div[data-testid="stMetricLabel"] > div {{ color:{COLORS["muted"]} !important; font-size:13px !important; }}
div[data-testid="stMetricValue"] > div {{ color:{COLORS["text"]} !important; font-size:22px !important; }}
div[data-testid="stMetricDelta"] > div {{ color:{COLORS["green"]} !important; }}

.stDownloadButton button, .stButton button {{
    background:{COLORS["surface"]} !important;
    color:{COLORS["text"]} !important;
    border:1px solid rgba(223,194,242,0.20) !important;
    border-radius:6px;
}}

.hero {{border:1px solid rgba(223,194,242,0.18);border-radius:8px;
        padding:10px 18px;background:{COLORS["surface"]};margin-bottom:10px;}}
.hero-title {{color:{COLORS["thistle"]};font-size:22px;font-weight:760;margin-bottom:4px;}}
.hero-subtitle {{color:{COLORS["muted"]};font-size:14px;max-width:1200px;}}
.section-note {{color:{COLORS["muted"]};font-size:14px;margin-top:-4px;margin-bottom:10px;}}
.insight-box {{background:{COLORS["surface"]};border:1px solid rgba(223,194,242,0.12);
               border-left:4px solid {COLORS["blue"]};border-radius:8px;
               padding:12px 14px;margin-bottom:8px;color:{COLORS["text"]};font-size:14px;line-height:1.55;}}
.insight-box.warn {{border-left-color:{COLORS["coral"]};}}
.insight-box.gold {{border-left-color:{COLORS["gold"]};}}
.risk-card {{
    background:{COLORS["surface"]};border:1px solid rgba(136,128,152,0.18);
    border-radius:8px;padding:12px 14px;margin-bottom:8px;
}}
.risk-badge {{
    display:inline-block;font-size:11px;font-weight:700;
    padding:2px 8px;border-radius:12px;margin-bottom:6px;
}}
.kpi-label {{font-size:12px;color:{COLORS["muted"]};text-transform:uppercase;letter-spacing:.04em;}}
.kpi-value {{font-size:26px;font-weight:700;color:{COLORS["text"]};}}
.kpi-sub {{font-size:12px;color:{COLORS["muted"]};}}
</style>
""", unsafe_allow_html=True)


# ── Data loading ──────────────────────────────────────────────────────────────
def safe_read(name: str) -> pd.DataFrame:
    path = OUTPUT_DIR / name
    if not path.exists():
        return pd.DataFrame()
    df = pd.read_csv(path)
    for col in ["date", "period_end"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    return df


def _latest_value(df: pd.DataFrame, value_col: str, variable: str | None = None,
                  variable_col: str = "variable") -> float | None:
    if df.empty or value_col not in df.columns:
        return None
    work = df.copy()
    if variable is not None and variable_col in work.columns:
        work = work[work[variable_col].eq(variable)]
    if work.empty:
        return None
    date_col = "date" if "date" in work.columns else "period_end" if "period_end" in work.columns else None
    if date_col:
        work = work.sort_values(date_col)
    vals = pd.to_numeric(work[value_col], errors="coerce").dropna()
    return None if vals.empty else float(vals.iloc[-1])


def _latest_rate_from_database(column: str) -> float | None:
    """Return the latest non-zero daily rate from PostgreSQL or Excel database.

    Values are returned as decimals, e.g. 0.0355 for 3.55%.
    """
    # Try PostgreSQL first if connected
    if DB_ENGINE is not None:
        try:
            query = f"""
            SELECT "{column}" FROM precios_y_tasas
            WHERE "Fecha" IS NOT NULL AND "{column}" IS NOT NULL AND "{column}" != 0
            ORDER BY "Fecha" DESC LIMIT 1
            """
            df = pd.read_sql(query, DB_ENGINE)
            if not df.empty:
                value = float(df[column].iloc[0])
                return value / 100 if value > 1 else value
        except Exception:
            pass  # Fall through to Excel
    
    # Fall back to Excel file
    if not DATABASE_FILE.exists():
        return None
    try:
        rates = pd.read_excel(DATABASE_FILE, sheet_name="Precios y tasas")
    except Exception:
        return None
    if "Fecha" not in rates.columns or column not in rates.columns:
        return None
    work = rates[["Fecha", column]].copy()
    work["Fecha"] = pd.to_datetime(work["Fecha"], errors="coerce")
    work[column] = pd.to_numeric(work[column], errors="coerce")
    work = work.dropna(subset=["Fecha", column])
    work = work[work[column].ne(0)]
    if work.empty:
        return None
    value = float(work.sort_values("Fecha")[column].iloc[-1])
    return value / 100 if value > 1 else value


def _proj_annual_value(annual_projection: pd.DataFrame, variable: str, year: str, scenario: str) -> float | None:
    if annual_projection.empty:
        return None
    work = annual_projection.copy()
    if "year" in work.columns:
        work["year"] = work["year"].astype(str)
    row = work[(work.get("variable") == variable) & (work.get("year") == str(year))]
    if row.empty:
        return None
    scenario_col = {"Base": "base_flat_avg", "Optimista": "high_80_avg", "Adverso": "low_80_avg"}[scenario]
    if scenario_col not in row.columns:
        return None
    val = pd.to_numeric(row[scenario_col], errors="coerce").dropna()
    return None if val.empty else float(val.iloc[0])


def _prophet_rate_value(prophet: pd.DataFrame, year: str, scenario: str, variable: str) -> float | None:
    if prophet.empty or "variable" not in prophet.columns or "date" not in prophet.columns:
        return None
    work = prophet[prophet["variable"].eq(variable)].copy()
    if work.empty:
        return None
    work["date"] = pd.to_datetime(work["date"], errors="coerce")
    work = work[work["date"].dt.year.astype("Int64").astype(str).eq(str(year))]
    if work.empty:
        return None
    # Lower rates are favorable for Autlan; high_80 is the adverse rate path.
    # Base scenarios are handled separately as flat-forward at latest observed.
    col = {"Base": "base", "Optimista": "low_80", "Adverso": "high_80"}[scenario]
    vals = pd.to_numeric(work[col], errors="coerce").dropna()
    if vals.empty:
        return None
    return max(float(vals.mean()), 0.0)


def _annual_avg(df: pd.DataFrame, value_col: str, year: int) -> float | None:
    if df.empty or "period_end" not in df.columns or value_col not in df.columns:
        return None
    work = df.copy()
    work["period_end"] = pd.to_datetime(work["period_end"], errors="coerce")
    vals = pd.to_numeric(
        work.loc[work["period_end"].dt.year.eq(year), value_col],
        errors="coerce",
    ).dropna()
    return None if vals.empty else float(vals.mean())


def scenario_outputs_from_drivers(ref: pd.Series, manganese: float, gold: float, fx: float,
                                  sofr: float, tiie: float) -> dict[str, float]:
    mn_ref = float(ref["Manganeso"])
    gold_ref = float(ref["Oro (USD/oz)"])
    fx_ref = float(ref["USD/MXN"])
    sofr_ref = float(ref["SOFR (%)"]) / 100
    tiie_ref = float(ref["TIIE (%)"]) / 100

    dm = (manganese / mn_ref - 1) * 100
    dgold = (gold / gold_ref - 1) * 100
    dfx = (fx / fx_ref - 1) * 100
    dsofr = (sofr - sofr_ref) * 10_000
    dtiie = (tiie - tiie_ref) * 10_000
    weighted_rate_delta = USD_DEBT_WEIGHT * dsofr + MXN_DEBT_WEIGHT * dtiie

    gross = (
        float(ref["Mg Bruto (%)"])
        + 44.14 * dm / 100
        + GOLD_GROSS_MARGIN_BPS_PER_1PCT * dgold / 100
        + 22.05 * dfx / 100
    )
    operating = (
        float(ref["Mg Operativo (%)"])
        + 60.17 * dm / 100
        + GOLD_OPERATING_MARGIN_BPS_PER_1PCT * dgold / 100
        + 30.83 * dfx / 100
    )
    gf_sales = float(ref["Gasto Fin/Vtas (%)"]) + 160.70 * weighted_rate_delta / 10_000
    sales = float(ref["Ventas (USD M)"]) * (
        1 + 1.96 * dm / 100 + GOLD_SALES_WEIGHT * dgold / 100
    )

    gross_profit = sales * gross / 100
    operating_income = sales * operating / 100
    financial_expense = sales * gf_sales / 100
    ref_net = float(ref["Util. Neta (USD M)"])
    net_income = (
        ref_net
        + (operating_income - float(ref["Util. Op. (USD M)"]))
        - (financial_expense - float(ref["Gasto Fin (USD M)"]))
    )

    return {
        "Mg Bruto (%)": gross,
        "Mg Operativo (%)": operating,
        "Gasto Fin/Vtas (%)": gf_sales,
        "Ventas (USD M)": sales,
        "Util. Bruta (USD M)": gross_profit,
        "Util. Op. (USD M)": operating_income,
        "Gasto Fin (USD M)": financial_expense,
        "Util. Neta (USD M)": net_income,
        "Margen Neto (%)": net_income / sales * 100 if sales else np.nan,
    }


def prepare_dashboard_scenarios(
    scenarios: pd.DataFrame,
    analysis: pd.DataFrame,
    macro_q: pd.DataFrame,
    proj_actuals: pd.DataFrame,
    proj_annual: pd.DataFrame,
    gold_annual: pd.DataFrame,
    prophet: pd.DataFrame,
) -> pd.DataFrame:
    if scenarios.empty:
        return scenarios

    sc = scenarios.copy()
    sc["Año"] = sc["Año"].astype(str).replace("Ref. 2025", "Ref. Actual")

    latest_sofr = _latest_rate_from_database("SOFR")
    if latest_sofr is None:
        latest_sofr = _latest_value(macro_q, "sofr_qavg")
    if latest_sofr is None:
        latest_sofr = _latest_value(analysis, "sofr_qavg")
    latest_sofr_pct = None if latest_sofr is None else latest_sofr * 100
    latest_tiie = _latest_rate_from_database("TIIE 28")
    if latest_tiie is None:
        latest_tiie = _latest_value(macro_q, "tiie_28d_qavg")
    if latest_tiie is None:
        latest_tiie = _latest_value(analysis, "tiie_28d_qavg")
    latest_tiie_pct = None if latest_tiie is None else latest_tiie * 100
    ref_sofr = _annual_avg(macro_q, "sofr_qavg", 2025)
    if ref_sofr is None:
        ref_sofr = _annual_avg(analysis, "sofr_qavg", 2025)
    ref_sofr_pct = None if ref_sofr is None else ref_sofr * 100
    ref_tiie = _annual_avg(macro_q, "tiie_28d_qavg", 2025)
    if ref_tiie is None:
        ref_tiie = _annual_avg(analysis, "tiie_28d_qavg", 2025)
    ref_tiie_pct = None if ref_tiie is None else ref_tiie * 100
    ref_fx = _latest_value(macro_q, "usd_mxn_qavg")
    if ref_fx is None:
        ref_fx = _annual_avg(macro_q, "usd_mxn_qavg", 2025)
    ref_mn = _latest_value(macro_q, "manganese_ore_usd_dmtu_qavg")
    if ref_mn is None:
        ref_mn = _annual_avg(macro_q, "manganese_ore_usd_dmtu_qavg", 2025)
    ref_gold = _latest_value(macro_q, "gold_usd_oz_qavg")
    if ref_gold is None:
        ref_gold = _annual_avg(macro_q, "gold_usd_oz_qavg", 2025)
    ref_net_income = _annual_avg(analysis, "net_income", 2025)
    ref_net_income_m = None if ref_net_income is None else ref_net_income / 1e6

    if "SOFR (%)" not in sc.columns:
        sc["SOFR (%)"] = np.nan
    if "TIIE (%)" not in sc.columns:
        sc["TIIE (%)"] = np.nan
    if "Oro (USD/oz)" not in sc.columns:
        sc["Oro (USD/oz)"] = np.nan
    if "Util. Neta (USD M)" not in sc.columns:
        sc["Util. Neta (USD M)"] = np.nan
    if "Margen Neto (%)" not in sc.columns:
        sc["Margen Neto (%)"] = np.nan

    ref_mask = sc["Año"].eq("Ref. Actual")
    if ref_fx is not None:
        sc.loc[ref_mask, "USD/MXN"] = ref_fx
    if ref_mn is not None:
        sc.loc[ref_mask, "Manganeso"] = ref_mn
    if ref_gold is not None:
        sc.loc[ref_mask, "Oro (USD/oz)"] = ref_gold
    if latest_sofr_pct is not None:
        sc.loc[ref_mask, "SOFR (%)"] = latest_sofr_pct
    elif ref_sofr_pct is not None:
        sc.loc[ref_mask, "SOFR (%)"] = ref_sofr_pct
    if latest_tiie_pct is not None:
        sc.loc[ref_mask, "TIIE (%)"] = latest_tiie_pct
    elif ref_tiie_pct is not None:
        sc.loc[ref_mask, "TIIE (%)"] = ref_tiie_pct
    if ref_net_income_m is not None:
        sc.loc[ref_mask, "Util. Neta (USD M)"] = ref_net_income_m
        sc.loc[ref_mask, "Margen Neto (%)"] = ref_net_income_m / sc.loc[ref_mask, "Ventas (USD M)"].astype(float) * 100

    for year in ["2026", "2027"]:
        for scenario in ["Base", "Optimista", "Adverso"]:
            mask = sc["Año"].eq(year) & sc["Escenario"].eq(scenario)
            if not mask.any():
                continue

            fx = _proj_annual_value(proj_annual, "usd_mxn", year, scenario)
            mn = _proj_annual_value(proj_annual, "manganese_ore_usd_dmtu", year, scenario)
            gold = _proj_annual_value(gold_annual, "gold_usd_oz", year, scenario)
            if fx is not None:
                sc.loc[mask, "USD/MXN"] = fx
            if mn is not None:
                sc.loc[mask, "Manganeso"] = mn
            if gold is not None:
                sc.loc[mask, "Oro (USD/oz)"] = gold

            if scenario == "Base":
                if latest_sofr_pct is not None:
                    sc.loc[mask, "SOFR (%)"] = latest_sofr_pct
                if latest_tiie_pct is not None:
                    sc.loc[mask, "TIIE (%)"] = latest_tiie_pct
            else:
                sofr = _prophet_rate_value(prophet, year, scenario, "SOFR (%)")
                tiie = _prophet_rate_value(prophet, year, scenario, "TIIE 28d (%)")
                if sofr is not None:
                    sc.loc[mask, "SOFR (%)"] = sofr
                if tiie is not None:
                    sc.loc[mask, "TIIE (%)"] = tiie

    ref = sc[ref_mask].iloc[0]

    for idx, row in sc[~ref_mask].iterrows():
        outputs = scenario_outputs_from_drivers(
            ref,
            manganese=float(row["Manganeso"]),
            gold=float(row["Oro (USD/oz)"]),
            fx=float(row["USD/MXN"]),
            sofr=float(row["SOFR (%)"]) / 100,
            tiie=float(row["TIIE (%)"]) / 100,
        )
        for col, value in outputs.items():
            sc.loc[idx, col] = value

    return sc


@st.cache_data(show_spinner=False)
def load_all() -> dict[str, pd.DataFrame]:
    """Load all data with graceful fallback for missing files."""
    try:
        financials = safe_read("autlan_financials_clean.csv")
        analysis = safe_read("autlan_analysis_dataset.csv")
        macro_q = safe_read("macro_variables_quarterly.csv")
        proj_actuals = safe_read("usd_mxn_manganese_projection_actuals.csv")
        proj_annual = safe_read("usd_mxn_manganese_projection_annual_summary.csv")
        gold_annual = safe_read("gold_price_projection_annual_summary.csv")
        prophet = safe_read("prophet_projections.csv")
        scenarios = prepare_dashboard_scenarios(
            safe_read("ml_scenario_table_2026_2027.csv"),
            analysis,
            macro_q,
            proj_actuals,
            proj_annual,
            gold_annual,
            prophet,
        )

        return {
            "financials":       financials,
            "analysis":         analysis,
            "macro_q":          macro_q,
            "monthly_trends":   safe_read("monthly_macro_financial_trends_long.csv"),
            "annual_trends":    safe_read("annual_macro_financial_trends_long.csv"),
            "sensitivities":    safe_read("executive_sensitivity_summary.csv"),
            "correlations":     safe_read("micro_macro_correlations.csv"),
            "model_comparison": safe_read("model_comparison_all.csv"),
            "multi_coef":       safe_read("multivariate_model_coefficients.csv"),
            "multi_summary":    safe_read("multivariate_model_summary.csv"),
            "monthly_season":   safe_read("monthly_macro_seasonality_summary.csv"),
            "rate_regime":      safe_read("rate_regime_summary.csv"),
            "rate_insights":    safe_read("rate_regime_key_insights.csv"),
            "rolling_corr":     safe_read("rolling_correlations_by_rate_regime.csv"),
            "proj_actuals":     safe_read("usd_mxn_manganese_projection_actuals.csv"),
            "proj_forecast":    safe_read("usd_mxn_manganese_monthly_projection.csv"),
            "prophet":          prophet,
            "gold_actuals":     safe_read("gold_price_projection_actuals.csv"),
            "gold_forecast":    safe_read("gold_price_monthly_projection.csv"),
            "lstm_forecasts":   safe_read("lstm_price_forecasts.csv"),
            "shap":             safe_read("xgboost_shap_values.csv"),
            "kmeans_labels":    safe_read("kmeans_regime_labels.csv"),
            "kmeans_profiles":  safe_read("kmeans_regime_profiles.csv"),
            "proj_annual":      proj_annual,
            "gold_annual":      gold_annual,
            "scenarios":        scenarios,
        }
    except Exception as e:
        st.error(f"Error loading data: {e}")
        return {k: pd.DataFrame() for k in [
            "financials", "analysis", "macro_q", "monthly_trends", "annual_trends",
            "sensitivities", "correlations", "model_comparison", "multi_coef", "multi_summary",
            "monthly_season", "rate_regime", "rate_insights", "rolling_corr", "proj_actuals",
            "proj_forecast", "prophet", "gold_actuals", "gold_forecast", "lstm_forecasts",
            "shap", "kmeans_labels", "kmeans_profiles", "proj_annual", "gold_annual", "scenarios"
        ]}


try:
    D = load_all()
except Exception as e:
    st.error(f"Failed to initialize dashboard: {e}")
    st.stop()


# ── Helpers ───────────────────────────────────────────────────────────────────
def fmt_usd(v):
    if pd.isna(v): return "n/a"
    v = float(v)
    s = "-" if v < 0 else ""
    a = abs(v)
    if a >= 1e9: return f"{s}${a/1e9:,.2f}bn"
    if a >= 1e6: return f"{s}${a/1e6:,.1f}mm"
    return f"{s}${a:,.0f}"


def fmt_pct(v, d=1):
    if pd.isna(v): return "n/a"
    v = float(v)
    if abs(v) <= 2.5: v *= 100
    return f"{v:,.{d}f}%"


def inst_layout(fig, title="", height=380, r=40):
    fig.update_layout(
        title=dict(text=title, x=0.01, xanchor="left", font=dict(size=17, color=COLORS["thistle"])),
        height=height, margin=dict(l=52, r=r, t=60 if title else 28, b=72),
        plot_bgcolor=COLORS["surface"], paper_bgcolor=COLORS["navy"],
        font=dict(color=COLORS["text"], family="Inter, Arial, sans-serif", size=13),
        hoverlabel=dict(bgcolor=COLORS["surface"], bordercolor="rgba(223,194,242,0.35)", font_size=13),
        legend=dict(
            orientation="h", yanchor="bottom", y=-0.28, xanchor="left", x=0,
            font=dict(color=COLORS["text"], size=13),
            bgcolor="rgba(52,64,85,0.88)",
            bordercolor="rgba(223,194,242,0.25)", borderwidth=1,
        ),
    )
    fig.update_xaxes(showgrid=True, gridcolor="rgba(136,128,152,0.12)",
                     zeroline=False, tickfont=dict(size=12, color=COLORS["gray"]),
                     linecolor="rgba(136,128,152,0.20)", automargin=True)
    fig.update_yaxes(showgrid=True, gridcolor="rgba(136,128,152,0.12)",
                     zeroline=False, tickfont=dict(size=12, color=COLORS["gray"]),
                     linecolor="rgba(136,128,152,0.20)", automargin=True)
    return fig


def insight(text, variant=""):
    cls = f"insight-box {variant}".strip()
    st.markdown(f"<div class='{cls}'>{text}</div>", unsafe_allow_html=True)


def styled_table(df: pd.DataFrame, height: int = 260, row_colors=None) -> go.Figure:
    fill = row_colors if row_colors else [COLORS["surface"]] * len(df)
    fig = go.Figure(go.Table(
        header=dict(
            values=[f"<b>{c}</b>" for c in df.columns],
            fill_color=COLORS["navy"],
            font=dict(color=COLORS["thistle"], size=13, family="Arial"),
            align="left", height=36, line_color="rgba(136,128,152,0.25)",
            columnwidth="auto",
        ),
        cells=dict(
            values=[df[c].tolist() for c in df.columns],
            fill_color=[fill] * len(df.columns),
            font=dict(size=13, family="Arial", color=COLORS["text"]),
            align="left", height=32, line_color="rgba(136,128,152,0.20)",
            columnwidth="auto",
        ),
    ))
    fig.update_layout(
        height=height, paper_bgcolor=COLORS["navy"],
        font=dict(color=COLORS["text"]),
        margin=dict(t=8, b=8, l=8, r=8),
    )
    return fig


def _norm_cdf(x: float) -> float:
    return 0.5 * (1.0 + math.erf(x / math.sqrt(2)))


def _caplet_b76(F: float, K: float, sigma: float, T: float,
                P: float, alpha: float, N: float) -> float:
    if T < 1e-8 or F <= 0 or K <= 0:
        return 0.0
    d1 = (math.log(F / K) + 0.5 * sigma**2 * T) / (sigma * math.sqrt(T))
    d2 = d1 - sigma * math.sqrt(T)
    return P * alpha * N * (F * _norm_cdf(d1) - K * _norm_cdf(d2))


def price_tiie_cap(F: float, K: float, sigma: float, tenor_yr: float,
                   r: float, alpha: float, notional: float) -> float:
    n = int(round(tenor_yr / alpha))
    total = 0.0
    for i in range(1, n + 1):
        T_r = i * alpha
        T_p = (i + 1) * alpha
        P   = 1.0 / (1.0 + r * T_p)
        total += _caplet_b76(F, K, sigma, T_r, P, alpha, notional)
    return total


def latest_q(fin):
    if fin.empty or "period_end" not in fin.columns: return None
    v = fin.dropna(subset=["net_sales"]).sort_values("period_end")
    return v.iloc[-1] if not v.empty else None


def filter_dates(df, col, s, e):
    if df.empty or col not in df.columns: return df.copy()
    return df[df[col].between(s, e)].copy()


# ── Sidebar ───────────────────────────────────────────────────────────────────
all_dates = []
for key, col in [("analysis","period_end"),("monthly_trends","date"),("gold_forecast","date")]:
    df = D.get(key, pd.DataFrame())
    if col in df.columns:
        all_dates.append(df[col].dropna())
if all_dates:
    min_d = pd.concat(all_dates).min()
    max_d = pd.concat(all_dates).max()
else:
    min_d, max_d = pd.Timestamp("2015-01-01"), pd.Timestamp("2027-12-31")

default_start = max(min_d, max_d - pd.DateOffset(years=8))

st.sidebar.markdown("### Filtros globales")
dr = st.sidebar.date_input("Rango de fechas",
    value=(default_start.date(), max_d.date()),
    min_value=min_d.date(), max_value=max_d.date())
if isinstance(dr, tuple) and len(dr) == 2:
    S, E = pd.Timestamp(dr[0]), pd.Timestamp(dr[1])
else:
    S, E = default_start, max_d

st.markdown("""
<div class="hero">
  <div class="hero-title">Compañía Minera Autlán — Estrategia Financiera Internacional &amp; Gestión de Riesgos</div>
  <div class="hero-subtitle">
    Análisis histórico 2015–2025 · Modelos OLS, Prophet, LSTM, XGBoost/SHAP, K-Means · Escenarios 2026–2027 · Deuda mixta USD/MXN
  </div>
</div>
""", unsafe_allow_html=True)

TABS = [
    "1. Resumen Ejecutivo",
    "2. Contexto Autlán",
    "3. Entorno Macro",
    "4. Micro-Macro",
    "5. Sector / Commodities",
    "6. Riesgos",
    "7. Escenarios 2026–27",
    "8. Cobertura",
    "9. Anexos",
]
tabs = st.tabs(TABS)


# ═══════════════════════════════════════════════════════════════════════════════
# TAB 1 — RESUMEN EJECUTIVO
# ═══════════════════════════════════════════════════════════════════════════════
with tabs[0]:
    st.markdown("<div class='section-note'>KPIs 2025 · Drivers cuantificados · Rango de escenarios 2026–2027</div>",
                unsafe_allow_html=True)

    lq = latest_q(D["financials"])
    if lq is not None:
        c1,c2,c3,c4,c5,c6 = st.columns(6)
        c1.metric("Ventas netas",       fmt_usd(lq.get("net_sales")),                f"{lq['period_end']:%b-%Y}")
        c2.metric("Margen bruto",       fmt_pct(lq.get("gross_margin")),             "último trimestre")
        c3.metric("Margen operativo",   fmt_pct(lq.get("operating_margin")),         "último trimestre")
        c4.metric("Margen neto",        fmt_pct(lq.get("net_margin")),               "último trimestre")
        c5.metric("Gasto fin. / ventas",fmt_pct(lq.get("financial_expense_to_sales")),"de ventas")
        c6.metric("Deuda financiera",   fmt_usd(lq.get("financial_debt_approx")),    "balance")

    col_l, col_r = st.columns([1.1, 0.9])
    with col_l:
        st.markdown("#### Hallazgos críticos")
        insight(
            "<b>Modelo de negocio — asimetría estructural:</b> Autlán factura en USD, tiene costos en MXN y deuda mixta: aprox. USD 130mm en dólares y remanente en pesos. "
            "Depreciación del peso <b>mejora márgenes</b>; apreciación los comprime. "
            "El gasto financiero debe analizarse por tramo: SOFR/tasas USD para deuda en dólares y TIIE para deuda en pesos.",
        )
        insight(
            "<b>Crisis 2025:</b> Gasto financiero USD 42.5mm > Utilidad bruta USD 51.4mm. "
            "La empresa genera margen bruto (15.9%) pero la carga financiera de deuda mixta en tasas altas la lleva a pérdida neta de USD 38mm.",
            variant="warn",
        )
        insight(
            "<b>Drivers cuantificados (OLS multivariado, R²adj = 0.396):</b><br>"
            "· Manganeso: +44.1 bps Mg Bruto / +1% precio &nbsp;|&nbsp; +60.2 bps Mg Op. / +1%<br>"
            "· Tasas SOFR/TIIE: +160.7 bps GF/Vtas / +100 bps tasa ponderada; impacto debajo del EBIT<br>"
            "· USD/MXN: +22.1 bps Mg Bruto / +1% FX &nbsp;|&nbsp; +30.8 bps Mg Op. / +1%",
            variant="gold",
        )

    with col_r:
        st.markdown("#### Escenario Base 2026–2027")
        if not D["scenarios"].empty:
            sc = D["scenarios"]
            base_only = sc[sc["Escenario"] == "Base"].copy()
            _sc_show2 = base_only[["Año","Escenario","USD/MXN","SOFR (%)","TIIE (%)","Manganeso",
                                  "Oro (USD/oz)","Mg Operativo (%)","Ventas (USD M)",
                                  "Util. Neta (USD M)"]].copy()
            _sc_row_colors = ["#1d3038"] * len(_sc_show2)
            st.plotly_chart(styled_table(_sc_show2, height=260,
                            row_colors=_sc_row_colors), use_container_width=True)
        else:
            st.markdown("<div style='color:#999;font-size:13px;'>Datos de escenarios no disponibles.</div>", unsafe_allow_html=True)
# ═══════════════════════════════════════════════════════════════════════════════
# TAB 2 — CONTEXTO DE AUTLÁN
# ═══════════════════════════════════════════════════════════════════════════════
with tabs[1]:
    st.markdown("<div class='section-note'>Modelo de negocio, estructura de exposiciones y P&L 2025.</div>",
                unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    c1.markdown(f"""<div class='risk-card' style='border-left:4px solid {COLORS["blue"]};'>
<div style='color:{COLORS["blue"]};font-weight:700;font-size:13px;'>INGRESOS — USD</div>
<div style='font-size:14px;color:{COLORS["text"]};margin-top:6px;line-height:1.6;'>
Manganeso, ferrosilicio y oro cotizados internacionalmente.<br>
Depreciación MXN → <b>mejora márgenes en USD</b>.<br>
Vulnerables al precio de commodities y demanda de acero.
</div></div>""", unsafe_allow_html=True)
    c2.markdown(f"""<div class='risk-card' style='border-left:4px solid {COLORS["gold"]};'>
<div style='color:{COLORS["gold"]};font-weight:700;font-size:13px;'>COSTOS — MXN</div>
<div style='font-size:14px;color:{COLORS["text"]};margin-top:6px;line-height:1.6;'>
Mano de obra, energía, transporte y mantenimiento.<br>
Se abaratan en USD cuando MXN se deprecia.<br>
Presionados por inflación doméstica si MXN se aprecia.
</div></div>""", unsafe_allow_html=True)
    c3.markdown(f"""<div class='risk-card' style='border-left:4px solid {COLORS["red"]};'>
<div style='color:{COLORS["red"]};font-weight:700;font-size:13px;'>DEUDA — MIXTA USD/MXN</div>
<div style='font-size:14px;color:{COLORS["text"]};margin-top:6px;line-height:1.6;'>
Aprox. USD 130mm en dólares; remanente en pesos.<br>
Tramo USD con matching natural contra ingresos en USD.<br>
Riesgo de tasas separado: SOFR para USD y TIIE para MXN.
</div></div>""", unsafe_allow_html=True)

    col_l, col_r = st.columns([1.1, 0.9])
    with col_l:
        vals = [100, -84.1, 15.9, -17.2, -1.3, -10.5, -11.8]
        labels = ["Ventas","Costo ventas","Mg Bruto","SG&A+otros","Mg Operativo","Gasto fin. neto","Pérdida neta"]
        measures = ["absolute","relative","total","relative","total","relative","total"]
        fig = go.Figure(go.Waterfall(
            orientation="v", measure=measures, x=labels, y=vals,
            text=[f"{v:.1f}%" for v in vals], textposition="outside",
            connector={"line":{"color":"rgba(38,59,94,0.2)"}},
            increasing={"marker":{"color":COLORS["green"]}},
            decreasing={"marker":{"color":COLORS["red"]}},
            totals={"marker":{"color":COLORS["blue"]}},
        ))
        fig.update_yaxes(title_text="% de ventas", ticksuffix="%", range=[-22, 118])
        st.plotly_chart(inst_layout(fig, "Cascada P&L 2025 — % de ventas", height=420),
                        use_container_width=True)

    with col_r:
        st.markdown("#### Lectura del P&L 2025")
        insight(
            "<b>Costo de ventas: 84.1% de ingresos</b> — margen bruto de 15.9%, "
            "equivalente a USD 51.4mm. Nivel operativamente ajustado pero no crítico."
        )
        insight(
            "<b>Gasto financiero neto: USD 39.7mm</b> = 77% de la utilidad bruta. "
            "La TIIE 28d escaló de ~4% (2020) a ~11% (2024), triplicando la carga. "
            "Este es el golpe definitivo al P&L.",
            variant="warn",
        )
        insight(
            "<b>Pérdida neta: USD 38mm.</b> Sin reducción de deuda o alza de commodities, "
            "la empresa no es autosuficiente financieramente en el entorno actual de tasas.",
            variant="warn",
        )
        st.markdown(f"""
<div style='background:{COLORS["soft_gray"]};border-radius:6px;padding:10px 14px;
     font-size:13px;color:{COLORS["muted"]};margin-top:8px;'>
<b style='color:{COLORS["thistle"]};'>Cifras 2025 (acumulado anual, USD mm):</b><br>
Ventas 322.7 · Costo ventas 271.3 · Util. Bruta 51.4<br>
Gastos fin. 42.5 · Ingresos fin. 2.8 · Pérdida neta 38.0<br>
Deuda financiera total ~192mm · Razón corriente 1.41x
</div>""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
# TAB 3 — ENTORNO MACRO
# ═══════════════════════════════════════════════════════════════════════════════
with tabs[2]:
    st.markdown("<div class='section-note'>Las tres variables que no controla Autlán: tipo de cambio, tasas de interés y ciclo de commodities.</div>",
                unsafe_allow_html=True)

    analysis = filter_dates(D["analysis"], "period_end", S, E)
    col_l, col_r = st.columns(2)

    with col_l:
        if not analysis.empty and "usd_mxn_qavg" in analysis.columns:
            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=analysis["period_end"], y=analysis["usd_mxn_qavg"],
                mode="lines+markers", line=dict(color=COLORS["blue"], width=2.8), name="USD/MXN",
                hovertemplate="%{x|%b-%Y}: %{y:.2f}<extra></extra>"))
            fig.update_yaxes(title_text="Pesos / dólar")
            st.plotly_chart(inst_layout(fig, "Tipo de cambio USD/MXN trimestral", height=330),
                            use_container_width=True)

    with col_r:
        if not analysis.empty and "sofr_qavg" in analysis.columns:
            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=analysis["period_end"], y=analysis["sofr_qavg"]*100,
                mode="lines+markers", line=dict(color=COLORS["red"], width=2.8), name="SOFR",
                hovertemplate="%{x|%b-%Y}: %{y:.2f}%<extra></extra>"))
            if "tiie_28d_qavg" in analysis.columns:
                fig.add_trace(go.Scatter(
                    x=analysis["period_end"], y=analysis["tiie_28d_qavg"]*100,
                    mode="lines", line=dict(color=COLORS["gold"], width=2.4, dash="dash"), name="TIIE 28d",
                    hovertemplate="%{x|%b-%Y}: %{y:.2f}%<extra></extra>"))
            fig.update_yaxes(title_text="%")
            st.plotly_chart(inst_layout(fig, "SOFR y TIIE 28d — tasas de referencia", height=330),
                            use_container_width=True)

    col_a, col_b, col_c = st.columns(3)
    with col_a:
        insight(
            "<b>USD/MXN:</b> De ~13 a ~21 pesos/USD entre 2015–2025. "
            "Depreciación MXN estructural que favoreció márgenes de Autlán. "
            "Riesgo: reversión a 14–15 (escenario adverso 2027)."
        )
    with col_b:
        insight(
            "<b>TIIE 28d:</b> De 4% (2020) a 11% (2024). Esos +700 bps elevaron "
            "el gasto financiero de 4% a 13% de ventas — el detonante de la pérdida neta.",
            variant="warn",
        )
    with col_c:
        insight(
            "<b>SOFR y TIIE correlacionan altamente.</b> El ciclo de ajuste Fed + Banxico fue "
            "simultáneo. Una re-aceleración inflacionaria global recrearía el escenario de 2022–2024.",
            variant="warn",
        )


# ═══════════════════════════════════════════════════════════════════════════════
# TAB 5 — SECTOR / COMMODITIES
# ═══════════════════════════════════════════════════════════════════════════════
with tabs[4]:
    st.markdown("<div class='section-note'>Proyecciones Prophet de manganeso y oro — los supuestos que alimentan los escenarios 2026–2027.</div>",
                unsafe_allow_html=True)

    prophet = D["prophet"]
    col_l, col_r = st.columns([1.3, 0.7])

    with col_l:
        if not prophet.empty:
            prophet["date"] = pd.to_datetime(prophet["date"])
            sect_vars = ["Manganeso (USD/dmtu)", "Oro (USD/oz)", "USD/MXN (pesos/dólar)"]
            var_sel = st.selectbox("Variable",
                [v for v in sect_vars if v in prophet["variable"].unique()], key="prophet_var")
            pv = prophet[prophet["variable"] == var_sel].sort_values("date")
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=pv["date"], y=pv["base"], mode="lines",
                name="Base", line=dict(color=COLORS["blue"], width=3)))
            fig.add_trace(go.Scatter(x=pv["date"], y=pv["high_80"], mode="lines",
                name="Optimista (High 80%)", line=dict(color=COLORS["green"], width=1.8, dash="dash")))
            fig.add_trace(go.Scatter(x=pv["date"], y=pv["low_80"], mode="lines",
                name="Adverso (Low 80%)", line=dict(color=COLORS["red"], width=1.8, dash="dash"),
                fill="tonexty", fillcolor="rgba(78,205,196,0.07)"))
            cutoff = pd.Timestamp("2026-06-01")
            fig.add_vrect(x0=cutoff, x1=pv["date"].max(),
                fillcolor="rgba(223,194,242,0.04)", line_width=0, layer="below")
            fig.add_vline(x=cutoff, line_dash="dash",
                line_color="rgba(223,194,242,0.35)", line_width=1.2)
            fig.update_yaxes(title_text=var_sel)
            st.plotly_chart(inst_layout(fig, f"Proyección Prophet — {var_sel} (2015–2027)", height=380),
                            use_container_width=True)

    with col_r:
        st.markdown("#### Riesgo de commodities")
        insight(
            "<b>Manganeso — driver principal:</b> Cotiza en ~4–5 USD/dmtu (vs 6–8 en el pico 2022). "
            "Cada −1% de precio quita ~44 bps de margen bruto. "
            "Incluyendo oro, el escenario adverso 2027 implica ventas de USD 77mm vs 169mm optimista."
        )
        insight(
            "<b>Proyección Prophet:</b> Los límites del intervalo 80% definen los escenarios. "
            "Banda ancha = alta incertidumbre. "
            "Base=yhat · Optimista=High80% · Adverso=Low80%."
        )
        insight(
            "<b>Oro — diversificador, no cobertura:</b> Correlación con manganeso r=−0.05. "
            "Actúa como refugio en incertidumbre. "
            "SOFR alto puede presionar el precio del oro a la baja.",
            variant="gold",
        )
        if not prophet.empty:
            pv_mang = prophet[prophet["variable"]=="Manganeso (USD/dmtu)"].dropna(subset=["date"])
            if not pv_mang.empty:
                last = pv_mang.sort_values("date").tail(12)
                rng_low  = last["low_80"].mean()
                rng_high = last["high_80"].mean()
                _m = COLORS["muted"]; _c = COLORS["coral"]; _g = COLORS["green"]
                st.markdown(
                    f"<div style='font-size:13px;color:{_m};padding:4px 0;'>"
                    f"Rango Prophet 2027 manganeso: "
                    f"<b style='color:{_c};'>{rng_low:.2f}</b> – "
                    f"<b style='color:{_g};'>{rng_high:.2f}</b> USD/dmtu"
                    f"</div>", unsafe_allow_html=True)

    # ── Price targets por escenario ─────────────────────────────────────────
    st.markdown(
        "<hr style='border-color:rgba(255,255,255,0.10);margin:20px 0 10px;'>",
        unsafe_allow_html=True)
    st.markdown(
        "<div style='font-size:15px;font-weight:700;color:#fff;margin-bottom:4px;'>"
        "Objetivos de precio por escenario — promedios anuales proyectados</div>",
        unsafe_allow_html=True)

    if not prophet.empty:
        _vars_t4 = [
            ("Manganeso (USD/dmtu)",   "Manganeso",  "USD/dmtu",  2),
            ("Oro (USD/oz)",            "Oro",        "USD/oz",    0),
            ("USD/MXN (pesos/dólar)",  "USD/MXN",    "pesos/USD", 2),
        ]
        _scen_order = [
            ("Adverso",   "low_80",  COLORS["coral"]),
            ("Base",      "base",    COLORS["blue"]),
            ("Optimista", "high_80", COLORS["green"]),
        ]

        _col_mn2, _col_or2, _col_fx2 = st.columns(3)
        _cols_t4 = [_col_mn2, _col_or2, _col_fx2]

        for (_var_full, _var_label, _unit, _dec), _pcol in zip(_vars_t4, _cols_t4):
            with _pcol:
                _pv2 = prophet[prophet["variable"] == _var_full].copy()
                if _pv2.empty:
                    continue
                _pv2["date"] = pd.to_datetime(_pv2["date"])
                _pv2["year"] = _pv2["date"].dt.year
                _years2 = sorted([y for y in _pv2["year"].unique() if y >= 2026])
                if not _years2:
                    continue

                # Current spot (latest base value)
                _spot2 = _pv2.sort_values("date").iloc[-1]["base"]

                _bx, _bv, _bc, _bt = [], [], [], []
                for _yr2 in _years2:
                    _yd = _pv2[_pv2["year"] == _yr2]
                    for _sn, _ck, _scolor in _scen_order:
                        _val2 = _yd[_ck].mean()
                        _bx.append(f"{_yr2}<br><b>{_sn}</b>")
                        _bv.append(round(_val2, _dec))
                        _bc.append(_scolor)
                        _bt.append(f"{_val2:,.{_dec}f}")

                _fig_pt = go.Figure(go.Bar(
                    x=_bx, y=_bv,
                    marker_color=_bc,
                    text=_bt,
                    textposition="outside",
                    textfont=dict(size=10),
                    cliponaxis=False,
                ))
                _fig_pt.add_hline(
                    y=_spot2,
                    line_dash="dot",
                    line_color="rgba(255,255,255,0.40)",
                    line_width=1.2,
                    annotation_text=f"  Actual: {_spot2:,.{_dec}f}",
                    annotation_position="top left",
                    annotation_font=dict(size=9, color="rgba(255,255,255,0.50)"),
                )
                _all_v2 = _bv + [_spot2]
                _ylo2 = min(_all_v2); _yhi2 = max(_all_v2)
                _ysp2 = max(_yhi2 - _ylo2, abs(_yhi2) * 0.05, 0.5)
                _fig_pt.update_yaxes(
                    title_text=_unit,
                    range=[_ylo2 - _ysp2 * 0.15, _yhi2 + _ysp2 * 0.38],
                )
                _fig_pt.update_xaxes(tickfont=dict(size=9))
                st.plotly_chart(
                    inst_layout(_fig_pt, _var_label, height=310),
                    use_container_width=True, key=f"pt_{_var_label}",
                )


# ═══════════════════════════════════════════════════════════════════════════════
# TAB 4 — MICRO-MACRO
# ═══════════════════════════════════════════════════════════════════════════════
with tabs[3]:
    st.markdown("<div class='section-note'>Regresiones OLS que cuantifican cuánto mueve el P&L cada variable macro — evidencia estadística de los canales de riesgo.</div>",
                unsafe_allow_html=True)

    col_l, col_r = st.columns(2)

    with col_l:
        sens = D["sensitivities"]
        if not sens.empty:
            plot = sens.dropna(subset=["impact"]).copy()
            # Excluir relaciones que no se usan como canal contable directo:
            # oro -> márgenes viene de supuesto operativo, no del OLS histórico;
            # tasas -> márgenes puede ser correlación de ciclo, pero en escenarios
            # SOFR/TIIE solo afecta debajo del EBIT.
            _exclude = (plot["driver_label"].str.contains("Oro", case=False, na=False) &
                        plot["outcome_label"].str.contains("Margen", case=False, na=False))
            _exclude_rate_margin = (
                plot["driver_label"].str.contains("SOFR|TIIE", case=False, na=False)
                & plot["outcome_label"].str.contains("Margen operativo|Margen bruto", case=False, na=False)
            )
            plot = plot[~(_exclude | _exclude_rate_margin)].copy()
            plot["abs_impact"] = plot["impact"].abs()
            plot = plot.sort_values("abs_impact", ascending=True).tail(10)
            # Shorten labels for cleaner display
            _short = {
                "driver_label": {
                    "SOFR +100 bps": "SOFR +100 bps",
                    "TIIE +100 bps": "TIIE +100 bps",
                    "Manganeso +1%": "Manganeso +1%",
                    "USD/MXN +1%":   "USD/MXN +1%",
                    "Oro +1%":       "Oro +1%",
                },
                "outcome_label": {
                    "Margen operativo": "Mg Operativo",
                    "Margen bruto":     "Mg Bruto",
                    "Gasto financiero / ventas": "GF/Ventas",
                    "Ventas":           "Ventas",
                },
            }
            plot["drv_s"] = plot["driver_label"].replace(_short["driver_label"])
            plot["out_s"] = plot["outcome_label"].replace(_short["outcome_label"])
            plot["label"] = plot["drv_s"] + "  →  " + plot["out_s"]
            # GF/Ventas: más alto = peor para Autlán → invertir lógica de color
            colors_bar = [
                (COLORS["green"] if v >= 0 else COLORS["coral"])
                if "GF" not in out
                else (COLORS["coral"] if v >= 0 else COLORS["green"])
                for v, out in zip(plot["impact"], plot["out_s"])
            ]
            # Compute x-axis range with 40% padding on each side for outside labels
            _xmin = float(plot["impact"].min())
            _xmax = float(plot["impact"].max())
            _span = max(abs(_xmin), abs(_xmax), 1.0)
            _x_range = [-_span * 1.65, _span * 1.75]
            fig = go.Figure(go.Bar(
                x=plot["impact"], y=plot["label"], orientation="h",
                marker_color=colors_bar,
                text=plot["impact_label"], textposition="outside",
                textfont=dict(size=12, color=COLORS["text"]),
                cliponaxis=False,
            ))
            fig.add_vline(x=0, line_width=1, line_color="rgba(136,128,152,0.40)")
            _sens_fig = inst_layout(fig, "Sensibilidades OLS — impacto por driver (bps)", height=400)
            _sens_fig.update_layout(margin=dict(l=200, r=150, t=60, b=50))
            _sens_fig.update_xaxes(range=_x_range, tickfont=dict(size=11), automargin=False)
            _sens_fig.update_yaxes(tickfont=dict(size=12), automargin=False)
            st.plotly_chart(_sens_fig, use_container_width=True)
            st.caption("Verde = impacto positivo para Autlán · Rosa = negativo · "
                       "Para GF/Ventas el color se invierte: subir el gasto financiero es malo. "
                       "OLS histórico; tasas/márgenes y oro/márgenes excluidos como canales contables directos.")

    with col_r:
        analysis = D["analysis"]
        if not analysis.empty and "manganese_ore_usd_dmtu_qavg" in analysis.columns \
                and "operating_margin" in analysis.columns:
            pair = analysis[["period_end","manganese_ore_usd_dmtu_qavg","operating_margin"]].dropna()
            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=pair["manganese_ore_usd_dmtu_qavg"], y=pair["operating_margin"]*100,
                mode="markers",
                marker=dict(size=9, color=COLORS["blue"], opacity=0.75,
                            line=dict(width=1, color="white")),
                customdata=pair[["period_end"]],
                hovertemplate="<b>%{customdata[0]|%b-%Y}</b><br>Mang: %{x:.1f}<br>Mg Op: %{y:.1f}%<extra></extra>"))
            if len(pair) >= 3:
                b, a = np.polyfit(pair["manganese_ore_usd_dmtu_qavg"], pair["operating_margin"]*100, 1)
                x_l = np.linspace(pair["manganese_ore_usd_dmtu_qavg"].min(),
                                  pair["manganese_ore_usd_dmtu_qavg"].max(), 80)
                fig.add_trace(go.Scatter(x=x_l, y=a+b*x_l, mode="lines",
                    line=dict(color=COLORS["gold"], width=3), name="OLS", hoverinfo="skip"))
                r2 = pair["manganese_ore_usd_dmtu_qavg"].corr(pair["operating_margin"])**2
                fig.add_annotation(x=0.02, y=0.97, xref="paper", yref="paper", showarrow=False,
                    text=f"β={b:.3f}<br>R²={r2:.2f}   n={len(pair)}",
                    bgcolor="rgba(52,64,85,0.92)", font=dict(size=12, color=COLORS["text"]))
            fig.update_xaxes(title_text="Manganeso USD/dmtu")
            fig.update_yaxes(title_text="Margen operativo %")
            st.plotly_chart(inst_layout(fig, "Manganeso vs. Margen operativo (OLS)", height=420),
                            use_container_width=True)

    col_a, col_b = st.columns(2)
    with col_a:
        insight(
            "<b>Verde = driver positivo para Autlán · Rosa = driver negativo.</b> "
            "SOFR/TIIE se interpreta como riesgo de gasto financiero y utilidad neta, no como efecto directo sobre EBIT. "
            "Manganeso es el driver positivo más importante (+44.1 bps Mg Bruto / +1%)."
        )
    with col_b:
        insight(
            "<b>Scatter manganeso–Mg Op.:</b> La relación positiva es clara pero el R² univariado (~0.35) "
            "muestra que SOFR y FX también son relevantes. "
            "El modelo multivariado con 3 variables alcanza R²adj=0.396.",
            variant="gold",
        )

    # ── Football Field ───────────────────────────────────────────────────────
    st.markdown(
        "<hr style='border-color:rgba(255,255,255,0.10);margin:20px 0 10px;'>",
        unsafe_allow_html=True)
    _ff_toggle = st.toggle("Ver Football Field — rango de resultados por escenario 2026-2027", value=False)
    if _ff_toggle:
        _sc_ff = D["scenarios"].copy()
        _non_ref = _sc_ff[_sc_ff["Año"].astype(str).str.match(r"^\d{4}$")].copy()
        _non_ref["Año"] = _non_ref["Año"].astype(int)

        # Metrics to display: (col_name, display_label, unit, decimals)
        _ff_defs = [
            ("Ventas (USD M)",   "Ventas",       "USD M", 1),
            ("Mg Bruto (%)",     "Mg Bruto",     "%",     1),
            ("Mg Operativo (%)", "Mg Operativo", "%",     1),
            ("Util. Neta (USD M)","Util. Neta",  "USD M", 1),
        ]

        _ff_rows   = []   # row label
        _ff_base   = []   # base value (marker)
        _ff_lo     = []   # adverso (bar start)
        _ff_hi     = []   # optimista (bar end)
        _ff_units  = []

        for _yr_ff in [2026, 2027]:
            _yr_data = _non_ref[_non_ref["Año"] == _yr_ff]
            for _col, _lbl, _unit, _dec in _ff_defs:
                if _col not in _yr_data.columns:
                    continue
                _adv = _yr_data[_yr_data["Escenario"] == "Adverso"][_col].values
                _bas = _yr_data[_yr_data["Escenario"] == "Base"][_col].values
                _opt = _yr_data[_yr_data["Escenario"] == "Optimista"][_col].values
                if not (len(_adv) and len(_bas) and len(_opt)):
                    continue
                _ff_rows.append(f"{_lbl} {_yr_ff}")
                _ff_lo.append(float(_adv[0]))
                _ff_base.append(float(_bas[0]))
                _ff_hi.append(float(_opt[0]))
                _ff_units.append(_unit)

        if _ff_rows:
            _ff_widths = [h - l for h, l in zip(_ff_hi, _ff_lo)]

            fig_ff = go.Figure()
            # Invisible base bar (transparent) to shift bars to correct start
            fig_ff.add_trace(go.Bar(
                y=_ff_rows, x=_ff_lo, orientation="h",
                marker_color="rgba(0,0,0,0)", showlegend=False,
                hoverinfo="skip",
            ))
            # Range bar (Adverso → Optimista)
            fig_ff.add_trace(go.Bar(
                y=_ff_rows, x=_ff_widths, orientation="h",
                base=_ff_lo,
                marker_color="rgba(78,205,196,0.28)",
                marker_line=dict(color=COLORS["blue"], width=1.5),
                name="Rango Adverso–Optimista",
                hovertemplate="<b>%{y}</b><br>Adverso: %{base:.1f}<br>Optimista: %{x:.1f}<extra></extra>",
            ))
            # Base dot
            fig_ff.add_trace(go.Scatter(
                y=_ff_rows, x=_ff_base, mode="markers",
                marker=dict(symbol="diamond", size=11,
                            color=COLORS["gold"],
                            line=dict(color="white", width=1.5)),
                name="Base",
                hovertemplate="<b>%{y}</b><br>Base: %{x:.1f}<extra></extra>",
            ))
            # Adverso label (left)
            for row, lo, hi, bas in zip(_ff_rows, _ff_lo, _ff_hi, _ff_base):
                fig_ff.add_annotation(
                    x=lo, y=row, text=f"{lo:.1f}",
                    xanchor="right", xshift=-6,
                    showarrow=False,
                    font=dict(size=10, color=COLORS["coral"]),
                )
                fig_ff.add_annotation(
                    x=hi, y=row, text=f"{hi:.1f}",
                    xanchor="left", xshift=6,
                    showarrow=False,
                    font=dict(size=10, color=COLORS["green"]),
                )

            fig_ff.update_layout(barmode="stack")
            _ff_fig = inst_layout(fig_ff,
                "Football Field — Rango de resultados Adverso / Base / Optimista",
                height=420)
            _ff_fig.update_layout(
                margin=dict(l=160, r=80, t=60, b=40),
                legend=dict(orientation="h", x=0.5, xanchor="center", y=-0.08,
                            font=dict(size=11)),
            )
            _ff_fig.update_xaxes(title_text="Valor proyectado", tickfont=dict(size=11),
                                  zeroline=True, zerolinecolor="rgba(136,128,152,0.35)",
                                  zerolinewidth=1)
            _ff_fig.update_yaxes(tickfont=dict(size=11), automargin=False)
            st.plotly_chart(_ff_fig, use_container_width=True)
            st.caption(
                "Barra = rango Adverso–Optimista · Diamante dorado = escenario Base. "
                "Fuente: Prophet + OLS. Unidades mixtas: % para márgenes, USD M para valores absolutos."
            )


# ═══════════════════════════════════════════════════════════════════════════════
# TAB 6 — IDENTIFICACIÓN DE RIESGOS
# ═══════════════════════════════════════════════════════════════════════════════
with tabs[5]:
    st.markdown("<div class='section-note'>Riesgos ordenados por impacto cuantificado en el P&L de Autlán — convergencia OLS + SHAP + régimen de tasas.</div>",
                unsafe_allow_html=True)

    RISKS = [
        (COLORS["red"],  "Crítico", "Tasas de interés — SOFR/TIIE",
         "Deuda mixta: aprox. USD 130mm en dólares y remanente en pesos. "
         "El tramo USD debe mapearse a SOFR/tasas internacionales; el tramo MXN a TIIE 28d. "
         "Con gasto financiero ya en 13.1% de ventas, no hay margen de absorción."),
        (COLORS["red"],  "Crítico", "Precio de manganeso",
         "Principal driver de ingresos y margen bruto. +44.1 bps Mg Bruto / +1% precio. "
         "Escenario adverso 2027: manganeso 3.99 USD/dmtu y oro 3,954 USD/oz → ventas USD 77mm (vs 169mm optimista). "
         "Volatilidad alta, mercado poco líquido para cobertura."),
        (COLORS["gold"], "Alto", "Precio de oro",
         "Driver relevante por supuesto operativo: ~25% de ventas y ~50% de EBITDA. "
         "Se modela en escenarios con +0.25% ventas, +25 bps Mg Bruto y +50 bps Mg Op. por +1% oro. "
         "No se usa como coeficiente OLS causal porque la muestra histórica tiene multicolinealidad."),
        (COLORS["gold"], "Alto", "Cambiario — apreciación MXN",
         "El riesgo NO es depreciación (eso mejora márgenes), sino apreciación brusca del peso. "
         "−30.8 bps Mg Op. por −1% FX. Escenario adverso 2027: USD/MXN 15.6. "
         "Asimetría: ingresos en USD, costos en MXN y deuda parcialmente cubierta por ingresos USD."),
        (COLORS["gold"], "Alto", "Margen / inflación doméstica",
         "Costos laborales, energía y transporte en MXN. "
         "Inflación persistente sin traslado a precios internacionales comprime margen bruto. "
         "Amplificada si el peso se aprecia simultáneamente."),
        (COLORS["gold"], "Alto", "Liquidez y refinanciamiento",
         "Pérdida neta acumulada y GF > utilidad bruta. "
         "Razón corriente 1.41x aceptable, pero efectivo reducido. "
         "Acceso a crédito depende de condiciones de mercado y TIIE."),
        (COLORS["blue"], "Moderado", "Operativo",
         "Continuidad minera, permisos ambientales y seguridad en minas."),
        (COLORS["blue"], "Moderado", "Político / comercial",
         "T-MEC, aranceles a minerales, regulación energética y relación México–EE.UU."),
    ]

    col_l, col_r = st.columns(2)
    for i, (color, badge, title, desc) in enumerate(RISKS):
        target = col_l if i % 2 == 0 else col_r
        with target:
            badge_bg = color + "22"
            st.markdown(f"""
<div class='risk-card' style='border-left:4px solid {color};'>
  <span class='risk-badge' style='background:{badge_bg};
        border:1px solid {color}66;color:{color};'>{badge}</span>
  <div style='font-size:15px;font-weight:700;color:{COLORS["text"]};
              margin-bottom:5px;'>{title}</div>
  <div style='font-size:13px;color:{COLORS["muted"]};line-height:1.55;'>{desc}</div>
</div>""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
# TAB 7 — ESCENARIOS 2026–2027
# ═══════════════════════════════════════════════════════════════════════════════
with tabs[6]:
    st.markdown("<div class='section-note'>Prophet + OLS: ¿cuánto aporta cada driver macro al resultado estimado?</div>",
                unsafe_allow_html=True)

    sc = D["scenarios"]
    if not sc.empty:
        YEARS    = ["Ref. Actual", "2026", "2027"]
        SC_NAMES = ["Base", "Optimista", "Adverso"]
        ABBREV   = {"Base":"Base","Optimista":"Opt.","Adverso":"Adv."}

        # ── Waterfall + Scenario table ──────────────────────────────────────
        wc1, wc2, wc3 = st.columns(3)
        wf_yr  = wc1.selectbox("Año",       ["2026", "2027"], key="wf_yr")
        wf_scn = wc2.selectbox("Escenario", SC_NAMES,         key="wf_scn")
        wf_met = wc3.selectbox("Métrica",   [
            "Util. Neta (USD M)", "Ventas (USD M)", "Mg Bruto (%)", "Mg Operativo (%)", "Gasto Fin/Vtas (%)"
        ], key="wf_met")

        ref_row = sc[sc["Año"] == "Ref. Actual"].iloc[0]
        sel_row = sc[(sc["Año"] == wf_yr) & (sc["Escenario"] == wf_scn)].iloc[0]

        # Precios actuales para nota de contexto
        _mq_t7   = D["macro_q"]
        _spot_mn  = _latest_value(_mq_t7, "manganese_ore_usd_dmtu_qavg")
        _spot_gold= _latest_value(_mq_t7, "gold_usd_oz_qavg")
        _spot_fx  = _latest_value(_mq_t7, "usd_mxn_qavg")
        _spot_note = ""
        if _spot_mn and _spot_gold and _spot_fx:
            _spot_note = (
                f"Spot actual (ref. informativa): "
                f"Mn {_spot_mn:.2f} · Oro {_spot_gold:,.0f} · FX {_spot_fx:.2f}"
            )

        mn_ref = ref_row["Manganeso"]; mn_sc = sel_row["Manganeso"]
        gold_ref = ref_row["Oro (USD/oz)"]; gold_sc = sel_row["Oro (USD/oz)"]
        fx_ref = ref_row["USD/MXN"];  fx_sc = sel_row["USD/MXN"]
        sofr_ref = ref_row["SOFR (%)"] / 100
        sofr_sc  = sel_row["SOFR (%)"] / 100
        tiie_ref = ref_row["TIIE (%)"] / 100
        tiie_sc  = sel_row["TIIE (%)"] / 100
        dm  = (mn_sc / mn_ref - 1) * 100
        dgold = (gold_sc / gold_ref - 1) * 100
        dfx = (fx_sc / fx_ref - 1) * 100
        dsofr = (sofr_sc - sofr_ref) * 10_000
        dtiie = (tiie_sc - tiie_ref) * 10_000

        ref_val = float(ref_row[wf_met])
        est_val = float(sel_row[wf_met])
        is_pct  = "%" in wf_met

        def _wf_text(v, is_ref=False, is_total=False):
            s = f"{v:+.2f}" if not (is_ref or is_total) else f"{v:.2f}"
            return f"{s}{'pp' if is_pct else 'M'}"

        current = {
            "manganese": float(mn_ref),
            "gold": float(gold_ref),
            "fx": float(fx_ref),
            "sofr": float(sofr_ref),
            "tiie": float(tiie_ref),
        }
        driver_steps = [
            ("Manganeso", "manganese", float(mn_sc), f"{mn_ref:.2f} → {mn_sc:.2f} USD/dmtu ({'+' if dm>=0 else ''}{dm:.1f}%)"),
            ("Oro", "gold", float(gold_sc), f"{gold_ref:,.0f} → {gold_sc:,.0f} USD/oz ({'+' if dgold>=0 else ''}{dgold:.1f}%)"),
            ("USD/MXN", "fx", float(fx_sc), f"{fx_ref:.2f} → {fx_sc:.2f} pesos/USD ({'+' if dfx>=0 else ''}{dfx:.1f}%)"),
            ("SOFR", "sofr", float(sofr_sc), f"{sofr_ref*100:.1f}% → {sofr_sc*100:.1f}% ({'+' if dsofr>=0 else ''}{dsofr:.0f} bps; {USD_DEBT_WEIGHT:.0%} deuda USD)"),
            ("TIIE 28d", "tiie", float(tiie_sc), f"{tiie_ref*100:.1f}% → {tiie_sc*100:.1f}% ({'+' if dtiie>=0 else ''}{dtiie:.0f} bps; {MXN_DEBT_WEIGHT:.0%} deuda MXN)"),
        ]
        driver_contribs = []
        prev_val = ref_val
        for name, key, target, detail in driver_steps:
            current[key] = target
            out = scenario_outputs_from_drivers(ref_row, **current)
            new_val = float(out[wf_met])
            contrib = new_val - prev_val
            prev_val = new_val
            driver_contribs.append((name, detail, contrib))

        x_labels = ["Ref. Actual<br>último observado"] + [
            f"{name}<br>{detail.split('(')[-1].replace(')', '')}" for name, detail, _ in driver_contribs
        ] + [f"Est. {wf_yr}<br>{wf_scn}"]
        y_vals = [ref_val] + [c for _, _, c in driver_contribs] + [0]
        texts = [_wf_text(ref_val, is_ref=True)] + [_wf_text(c) for _, _, c in driver_contribs] + [
            _wf_text(est_val, is_total=True)
        ]

        col_wf, col_tbl = st.columns([1.3, 0.7])
        with col_wf:
            fig_wf = go.Figure(go.Waterfall(
                orientation="v",
                measure=["absolute"] + ["relative"] * len(driver_contribs) + ["total"],
                x=x_labels, y=y_vals, text=texts, textposition="outside",
                textfont=dict(size=13, color=COLORS["text"]),
                connector=dict(line=dict(color="rgba(136,128,152,0.35)", width=1.5, dash="dot")),
                increasing=dict(marker=dict(color=COLORS["green"],  line=dict(width=0))),
                decreasing=dict(marker=dict(color=COLORS["red"],    line=dict(width=0))),
                totals=dict(marker=dict(color=SCENARIO_COLORS[wf_scn], line=dict(width=0))),
                hovertemplate="<b>%{x}</b><br>%{y:+.3f} " + ("pp" if is_pct else "M USD") + "<extra></extra>",
            ))
            if is_pct:
                fig_wf.add_hline(y=0, line_dash="dot",
                                 line_color="rgba(136,128,152,0.45)", line_width=1.2)
            _yunit = "% de ventas" if is_pct else "USD millones"
            # Compute running totals to set y-axis range with padding for outside labels
            _running = ref_val
            _all_y = [_running]
            for _, _, _c in driver_contribs:
                _running += _c
                _all_y.append(_running)
            _all_y.append(est_val)
            _y_lo = min(_all_y); _y_hi = max(_all_y)
            _y_pad = max(abs(_y_hi - _y_lo) * 0.22, abs(_y_lo) * 0.18, 1.5)
            fig_wf.update_yaxes(title_text=_yunit, ticksuffix=("%" if is_pct else ""),
                                 range=[_y_lo - _y_pad, _y_hi + _y_pad])
            st.plotly_chart(inst_layout(fig_wf,
                title=f"Puente {wf_met} — Promedio 2025 → {wf_yr} {wf_scn}",
                height=420, r=50), use_container_width=True)

        with col_tbl:
            # ── Dynamic scenario insight ─────────────────────────────────────
            _sc_color = SCENARIO_COLORS[wf_scn]
            _sc_bg = {"Base": "#1a2830", "Optimista": "#1a2812", "Adverso": "#2a1208"}[wf_scn]
            _unit_lbl = "pp" if is_pct else "MUSD"
            _net_change = est_val - ref_val

            def _driver_line(name, delta_str, contrib):
                if abs(contrib) < 0.005:
                    return ""
                _arrow = "&#9650;" if contrib > 0 else "&#9660;"
                _col   = COLORS["green"] if contrib > 0 else COLORS["red"]
                return (f"<li style='margin-bottom:5px;'>"
                        f"<b style='color:{COLORS['text']};'>{name}:</b> "
                        f"<span style='color:{COLORS['muted']};font-size:12px;'>{delta_str}</span><br>"
                        f"<span style='color:{_col};font-weight:700;'>"
                        f"{_arrow} {abs(contrib):.2f} {_unit_lbl}</span>"
                        f" en {wf_met}</li>")

            _rows_html = "".join(
                _driver_line(name, detail, contrib)
                for name, detail, contrib in driver_contribs
            )
            _contribs_abs = {name: abs(contrib) for name, _, contrib in driver_contribs}
            _dom = max(_contribs_abs, key=lambda k: _contribs_abs[k]) if _contribs_abs else "n/a"
            _dom_val = next((contrib for name, _, contrib in driver_contribs if name == _dom), 0)
            _dom_pct = abs(_dom_val / _net_change * 100) if _net_change != 0 else 0

            _context = {
                "Base":      "Recuperación gradual siguiendo la tendencia central Prophet. "
                             "Mejora moderada vs 2025 pero con fragilidad financiera persistente.",
                "Optimista": "Condiciones favorables simultáneas: manganeso al alza, "
                             "peso depreciado y tasas cediendo. Recuperación de márgenes a terreno positivo.",
                "Adverso":   "Estrés simultáneo en drivers macro/operativos: manganeso débil, "
                             "apreciación del peso y re-aceleración de tasas. "
                             "Recrea las condiciones de 2024–2025 pero más severas.",
            }[wf_scn]

            _ref_fmt  = f"{ref_val:.1f}{'%' if is_pct else ' MUSD'}"
            _est_fmt  = f"{est_val:.1f}{'%' if is_pct else ' MUSD'}"
            _net_fmt  = f"{'+' if _net_change>=0 else ''}{_net_change:.1f} {_unit_lbl}"
            spot_html = f"<div style='font-size:11px;color:{COLORS['muted']};font-style:italic;margin-bottom:4px;'>{_spot_note}</div>" if _spot_note else ""
            sales_note = (f"<div style='margin-top:6px;font-size:11px;color:{COLORS['muted']};"
                          f"background:rgba(136,128,152,0.08);border-radius:5px;padding:5px 8px;'>"
                          f"Ventas se modela con manganeso y oro; oro pesa 25% de ventas. La tasa de referencia no afecta ingresos directamente"
                          f" &mdash; su canal macro (demanda global) no es atribuible a nivel empresa."
                          f"</div>" if wf_met == "Ventas (USD M)" else "")

            st.markdown(f"""
<div style='background:{_sc_bg};border:1px solid {_sc_color}55;
     border-left:4px solid {_sc_color};border-radius:8px;
     padding:12px 14px;margin-bottom:10px;'>
  <div style='font-size:15px;font-weight:700;color:{_sc_color};margin-bottom:5px;'>
    {wf_scn} {wf_yr} &mdash; {wf_met}
  </div>
  <div style='font-size:12px;color:{COLORS["muted"]};margin-bottom:6px;line-height:1.5;'>
    {_context}
  </div>
  <div style='font-size:12px;color:{COLORS["muted"]};margin-bottom:3px;'>
    Ref. Actual (último obs.): <b style='color:{COLORS["text"]};'>{_ref_fmt}</b>
    &nbsp;&rarr;&nbsp;
    Est. {wf_yr}: <b style='color:{_sc_color};'>{_est_fmt}</b>
    &nbsp;({_net_fmt})
  </div>
    {spot_html}
  <div style='font-size:12px;color:{COLORS["muted"]};font-weight:600;
              margin-top:8px;margin-bottom:4px;'>Contribuci&#243;n por driver:</div>
  <ul style='margin:0;padding-left:14px;font-size:13px;
             color:{COLORS["text"]};line-height:1.5;'>
    {_rows_html}
  </ul>
  <div style='margin-top:8px;font-size:12px;color:{COLORS["muted"]};
              border-top:1px solid rgba(136,128,152,0.15);padding-top:6px;'>
    Driver dominante: <b style='color:{_sc_color};'>{_dom}</b>
    &nbsp;({abs(_dom_val):.2f} {_unit_lbl} &mdash; {_dom_pct:.0f}% del cambio total)
  </div>
    {sales_note}
</div>""", unsafe_allow_html=True)

            # ── Compact scenario table ──────────────────────────────────────
            ROW_BG = {"Base":"#1d3038","Optimista":"#272f19","Adverso":"#2f1e14"}
            tbl2 = sc[["Año","Escenario","SOFR (%)","TIIE (%)","Oro (USD/oz)",
                       "Mg Operativo (%)","Ventas (USD M)","Util. Neta (USD M)"]].copy()
            col_lbls2 = ["Año","Escenario","SOFR","TIIE","Oro","Mg Op.","Ventas","Util.Neta"]
            fills2 = [[ROW_BG[s] for s in tbl2["Escenario"]] for _ in col_lbls2]
            fig_t = go.Figure(go.Table(
                columnwidth=[50,70,48,48,58,58,58,64],
                header=dict(values=[f"<b>{c}</b>" for c in col_lbls2],
                    fill_color=COLORS["navy"],
                    font=dict(color=COLORS["thistle"],size=12,family="Arial"),
                    align="center", height=30, line_color="#344055"),
                cells=dict(values=[tbl2[c].tolist() for c in tbl2.columns],
                    fill_color=fills2,
                    font=dict(size=12,family="Arial",color=COLORS["text"]),
                    align="center", height=26, line_color="rgba(136,128,152,0.25)",
                    format=[None,None,".1f",".1f",".0f",".1f",".1f",".1f"]),
            ))
            fig_t.update_layout(height=220, paper_bgcolor=COLORS["navy"],
                font=dict(color=COLORS["text"]),
                margin=dict(t=6,b=6,l=4,r=4))
            st.plotly_chart(fig_t, use_container_width=True)

    else:
        st.info("No se encontró ml_scenario_table_2026_2027.csv")


# ═══════════════════════════════════════════════════════════════════════════════
# TAB 8 — ESTRATEGIA DE COBERTURA
# ═══════════════════════════════════════════════════════════════════════════════
with tabs[7]:
    st.markdown("<div class='section-note'>Instrumentos que mitigan las exposiciones — ordenados por urgencia e impacto cuantificado.</div>",
                unsafe_allow_html=True)

    # Impact bar
    col_chart, col_ins = st.columns([1.2, 0.8])
    with col_chart:
        _impact_rows = [
            ("Tasas SOFR/TIIE",   160, "P1 Crítico",  COLORS["red"]),
            ("Manganeso",          60, "P2 Crítico",  COLORS["blue"]),
            ("Oro",                50, "P3 Alto",     COLORS["green"]),
            ("Apreciación MXN",    31, "P4 Alto",     COLORS["gold"]),
        ]
        fig_imp = go.Figure(go.Bar(
            x=[r[1] for r in _impact_rows],
            y=[r[0] for r in _impact_rows],
            orientation="h",
            marker_color=[r[3] for r in _impact_rows],
            text=[f"{r[2]} — {r[1]} bps" for r in _impact_rows],
            textposition="outside",
            textfont=dict(size=13, color=COLORS["text"]),
            cliponaxis=False,
        ))
        fig_imp.update_xaxes(range=[0, 280])
        fig_imp.update_yaxes(tickfont=dict(size=14, color=COLORS["text"]))
        st.plotly_chart(inst_layout(fig_imp,
            "Impacto cuantificado (bps sobre métrica principal / unidad de shock)",
            height=260, r=10), use_container_width=True)

    with col_ins:
        insight(
            "<b>Por qué tasas primero:</b> el gasto financiero ya absorbe una parte crítica del P&L. "
            "La cobertura debe dividirse por moneda: IRS/Cap SOFR para el tramo USD variable e IRS/Cap TIIE para el tramo MXN variable."
        )
        insight(
            "<b>Por qué commodities antes que FX:</b> Manganeso y oro mueven ventas y EBITDA. "
            "El FX también importa (+31 bps Mg Op. por +1% USD/MXN), pero parte del riesgo se compensa "
            "con la estructura natural de ingresos USD y costos MXN.",
            variant="gold",
        )

    # Cards 2×2
    HEDGES = [
        {
            "priority": "P1", "badge": "Crítico", "color": COLORS["red"],
            "title": "Tasas de interés — SOFR/TIIE",
            "exposure": "Aprox. USD 130mm en USD + remanente MXN · sensibilidad de GF/Vtas por +100 bps",
            "context": "Con carga de 13.1% de ventas, cualquier alza relevante en tasas destruye el P&L. Prioridad máxima.",
            "instruments": [
                ("IRS SOFR", "Swap pagar fija / recibir variable para deuda USD flotante"),
                ("IRS TIIE 28d", "Swap pagar fija / recibir variable para deuda MXN flotante"),
                ("Caps SOFR/TIIE", "Límite superior a tasa flotante; conserva beneficio si baja"),
            ],
            "cost": "Prima/costo depende de curva, plazo y nocional por moneda",
        },
        {
            "priority": "P2", "badge": "Crítico", "color": COLORS["blue"],
            "title": "Commodities — manganeso",
            "exposure": "Adverso 2027: 3.99 USD/dmtu + oro 3,954 USD/oz · Ventas USD 77mm vs 169mm Optimista",
            "context": "Principal driver de ingresos. Cobertura menos líquida que tasas o FX.",
            "instruments": [
                ("Forward manganeso", "Contratos forward de venta a 6–12 meses para fijar precio"),
                ("Put manganeso", "Opciones put: participan del alza, limitan la baja"),
                ("Diversificación", "Ferrosilicio y oro como cobertura operativa natural"),
            ],
            "cost": "Prima opción ~2–4% del nocional",
        },
        {
            "priority": "P3", "badge": "Alto", "color": COLORS["green"],
            "title": "Commodities — oro",
            "exposure": "~25% de ventas y ~50% de EBITDA · adverso 2027: 3,954 USD/oz",
            "context": "Driver económico relevante en escenarios. La cobertura debe evaluarse por política comercial y liquidez, no por OLS.",
            "instruments": [
                ("Forward oro", "Contratos forward de venta a 3–6 meses"),
                ("Collar oro", "Compra put + venta call para acotar el rango de precios"),
            ],
            "cost": "Prima collar ~1–2% del nocional",
        },
        {
            "priority": "P4", "badge": "Alto", "color": COLORS["gold"],
            "title": "Cambiario — apreciación MXN",
            "exposure": "USD/MXN 15.6 adverso 2027 · −30.8 bps Mg Op. / −1% FX",
            "context": "El riesgo es apreciación del peso, no depreciación. Adverso 2027 implica USD/MXN 15.6.",
            "instruments": [
                ("Forward USD/MXN", "Compra de MXN / venta de USD para fijar tipo de cambio de costos"),
                ("Put USD/MXN", "Opción put: paga si el peso se aprecia por encima del strike"),
                ("Natural hedge", "Presupuesto de costos en USD equivalente — reduce exposición neta"),
            ],
            "cost": "Spread forward ~0.5–1.5% · Prima put ~1–3%",
        },
    ]

    card_a, card_b = st.columns(2)
    for i, h in enumerate(HEDGES):
        target_col = card_a if i % 2 == 0 else card_b
        with target_col:
            instr_html = "".join(
                f"<li style='margin-bottom:5px;'>"
                f"<span style='color:{h['color']};font-weight:700;'>{nm}</span>"
                f" — {desc}</li>"
                for nm, desc in h["instruments"]
            )
            st.markdown(f"""
<div style='background:{COLORS["surface"]};border:1px solid rgba(136,128,152,0.18);
     border-left:5px solid {h["color"]};border-radius:10px;
     padding:14px 16px;margin-bottom:12px;'>
  <div style='display:flex;align-items:flex-start;justify-content:space-between;
              margin-bottom:8px;gap:10px;'>
    <span style='font-size:15px;font-weight:700;color:{COLORS["text"]};
                 line-height:1.3;'>{h["title"]}</span>
    <span style='background:{h["color"]}22;border:1px solid {h["color"]}77;
                 color:{h["color"]};font-size:11px;font-weight:700;
                 padding:2px 9px;border-radius:20px;white-space:nowrap;'>
      {h["priority"]} — {h["badge"]}
    </span>
  </div>
  <div style='font-size:12px;color:{COLORS["blue"]};
              background:rgba(78,205,196,0.08);
              padding:5px 9px;border-radius:5px;margin-bottom:8px;'>
    {h["exposure"]}
  </div>
  <p style='font-size:13px;color:{COLORS["muted"]};margin-bottom:8px;line-height:1.5;'>{h["context"]}</p>
  <ul style='margin:0;padding-left:16px;color:{COLORS["text"]};font-size:13px;line-height:1.65;'>
    {instr_html}
  </ul>
  <div style='margin-top:8px;font-size:12px;color:{COLORS["muted"]};
              border-top:1px solid rgba(136,128,152,0.15);padding-top:6px;'>
    Costo est.: {h["cost"]}
  </div>
</div>""", unsafe_allow_html=True)


    # ── Estrategia activa: Receiver Swap + Cap TIIE ──────────────────────────
    st.markdown(
        "<hr style='border:none;border-top:1px solid rgba(136,128,152,0.20);margin:22px 0;'>",
        unsafe_allow_html=True)
    st.markdown("#### Estrategia activa recomendada: Receiver Swap + Cap TIIE 9%")
    st.markdown(
        "<div class='section-note'>Salir del payer swap 8% fijo → TIIE flotante + cap al 9%. "
        "Prima estimada con Black-76 y σ histórica TIIE 28d.</div>",
        unsafe_allow_html=True)

    # Volatilidad histórica TIIE
    _mq_h  = D["macro_q"]
    _tiie_h = _mq_h["tiie_28d_qavg"].dropna() if (not _mq_h.empty and "tiie_28d_qavg" in _mq_h.columns) else pd.Series(dtype=float)
    if len(_tiie_h) > 2:
        _lr_h   = np.log(_tiie_h / _tiie_h.shift(1)).dropna()
        _sig_h  = float(_lr_h.std() * np.sqrt(4))
    else:
        _sig_h  = 0.18
    _n_obs_h = len(_tiie_h)

    # Parámetros del instrumento
    _F_h     = 0.0676    # TIIE forward flat
    _K_h     = 0.09      # strike cap
    _fix_h   = 0.08      # tasa fija swap existente
    _alp_h   = 28 / 365
    _ten_h   = 2         # años
    _ncap_h  = int(round(_ten_h / _alp_h))
    _r_h     = _F_h
    _fx_h    = 17.40
    _noc_usd_h = 62e6
    _noc_mxn_h = _noc_usd_h * _fx_h

    # Precio del cap
    _cap_mxn_h = price_tiie_cap(_F_h, _K_h, _sig_h, _ten_h, _r_h, _alp_h, _noc_mxn_h)
    _cap_usd_h = _cap_mxn_h / _fx_h
    _cap_pct_h = _cap_mxn_h / _noc_mxn_h

    # Ahorros
    _spr_h      = _fix_h - _F_h
    _ann_sav_h  = _spr_h * _noc_usd_h
    _pv_sav_h   = sum(_ann_sav_h / (1 + _r_h)**_t for _t in range(1, _ten_h + 1))
    _net_h      = _pv_sav_h - _cap_usd_h

    # Escenarios
    _esc_h = [
        ("Optimista 2027", 0.0194),
        ("Adverso 2027",   0.0560),
        ("Base / Actual",  0.0676),
        ("Rebote a 10%",   0.1000),
        ("Pico 2024 11%",  0.1100),
    ]
    _esc_lbl_h = [r[0] for r in _esc_h]
    _esc_sav_h = [(_fix_h - t + max(t - _K_h, 0)) * _noc_usd_h / 1e3 for _, t in _esc_h]
    _esc_col_h = [COLORS["green"] if v >= 0 else COLORS["coral"] for v in _esc_sav_h]

    col_bar_h, col_kpi_h = st.columns([1.3, 0.7])

    with col_bar_h:
        # Para barras negativas el texto va a la izquierda ("outside") y choca
        # con las etiquetas del eje Y → usamos "inside" solo para los negativos.
        _txt_pos_h = ["inside" if v < 0 else "outside" for v in _esc_sav_h]
        _fig_h = go.Figure(go.Bar(
            x=_esc_sav_h, y=_esc_lbl_h, orientation="h",
            marker_color=_esc_col_h,
            text=[f"{v:+,.0f}K USD/yr" for v in _esc_sav_h],
            textposition=_txt_pos_h,
            textfont=dict(size=12, color=COLORS["text"]),
            cliponaxis=False,
        ))
        _fig_h.add_vline(x=0, line_width=1.2, line_color="rgba(136,128,152,0.45)")
        _fig_h = inst_layout(_fig_h, "Ahorro neto anual vs. mantener swap 8% fijo", height=340, r=220)
        _fig_h.update_layout(
            margin=dict(l=130, r=230, t=55, b=40),
            bargap=0.35,
        )
        _fig_h.update_xaxes(ticksuffix=" K USD", automargin=False)
        _fig_h.update_yaxes(automargin=False, tickfont=dict(size=13))
        st.plotly_chart(_fig_h, use_container_width=True)
        st.markdown(
            f"<div style='font-size:12px;color:{COLORS['muted']};padding:2px 6px;'>"
            f"Break-even: <b style='color:{COLORS['text']};'>TIIE = 8.00%</b>. "
            f"El cap limita el costo efectivo en 9% — peor caso: "
            f"<b style='color:{COLORS['coral']};'>−620K USD/año</b> "
            f"vs. el swap actual (solo 100 bps adicionales)."
            f"</div>", unsafe_allow_html=True)

    with col_kpi_h:
        _kpi_rows_h = [
            ("σ TIIE hist. (anualizada)",   f"{_sig_h:.1%}",                             COLORS["muted"]),
            ("Caplets (2 años, 28d c/u)",   f"{_ncap_h}",                                COLORS["muted"]),
            ("TIIE forward flat (F)",        f"{_F_h:.2%}",                               COLORS["text"]),
            ("Strike cap (K)",               f"{_K_h:.0%}",                               COLORS["text"]),
            ("Prima cap (una sola vez)",     f"MXN {_cap_mxn_h/1e6:.2f}mm = USD {_cap_usd_h/1e3:.0f}K", COLORS["gold"]),
            ("Prima / nocional",             f"{_cap_pct_h:.3%}",                         COLORS["muted"]),
            ("Ahorro anual (base 6.76%)",    f"+USD {_ann_sav_h/1e3:.0f}K / año",         COLORS["green"]),
            ("VP neto 2 años (base)",        f"USD {_net_h/1e6:.2f}mm",                   COLORS["green"]),
        ]
        _kpi_html_h = "".join(
            f"<div style='display:flex;justify-content:space-between;align-items:baseline;"
            f"padding:5px 0;border-bottom:1px solid rgba(136,128,152,0.12);'>"
            f"<span style='font-size:11px;color:{COLORS['muted']};'>{lbl}</span>"
            f"<span style='font-size:12px;font-weight:700;color:{col};white-space:nowrap;'>{val}</span>"
            f"</div>"
            for lbl, val, col in _kpi_rows_h
        )
        _c_surf_h = COLORS["surface"]
        _c_blue_h = COLORS["blue"]
        st.markdown(
            f"<div style='background:{_c_surf_h};border:1px solid rgba(223,194,242,0.15);"
            f"border-left:4px solid {_c_blue_h};border-radius:8px;padding:12px 14px;'>"
            f"<div style='font-size:14px;font-weight:700;color:{_c_blue_h};"
            f"margin-bottom:10px;'>Receiver Swap + Cap TIIE 9%</div>"
            f"{_kpi_html_h}"
            f"</div>", unsafe_allow_html=True)

    with st.expander("¿Cómo se calculó la prima? — Explicación simple", expanded=False):
        st.markdown(f"""
**El problema que resuelve este instrumento**

Autlán tiene deuda en pesos que hoy paga una tasa fija de **8%**. La TIIE de mercado bajó a **{_F_h:.2%}**,
así que están pagando de más. Quieren pasarse a tasa variable para aprovechar las tasas bajas —
pero si las tasas suben de nuevo, quedarían expuestos.

La solución: **salir del swap fijo + comprar un "techo" (cap) al 9%**.
El cap es un seguro: si la TIIE sube por encima del 9%, el banco les paga la diferencia.
Si la TIIE se queda abajo del 9%, el cap simplemente no se usa — y Autlán pagó solo la prima.

---

**¿Qué es la prima y cómo se calculó?**

La prima es el precio de ese seguro. Para calcularlo usamos el modelo **Black-76**,
que es el estándar de la industria para valuar opciones sobre tasas de interés.

La idea central es sencilla: *¿cuánto vale hoy el derecho a que te paguen si la TIIE sube?*
Depende de tres cosas:

1. **Qué tan lejos está la tasa actual del techo** — TIIE está en {_F_h:.2%} y el techo en {_K_h:.0%}.
   Hay {(_K_h-_F_h)*10000:.0f} puntos base de distancia. Mientras más lejos, más barata la prima
   (como asegurar tu casa contra un terremoto de magnitud 10 — poco probable, prima baja).

2. **Qué tan volátil ha sido la TIIE históricamente** — calculamos la volatilidad
   con {_n_obs_h} trimestres de datos reales: σ = **{_sig_h:.1%}** anualizada.
   Más volatilidad → más probabilidad de que la tasa llegue al techo → prima más cara.

3. **Cuánto tiempo dura el seguro** — el cap cubre **{_ten_h} años**, dividido en
   **{_ncap_h} períodos de 28 días** (uno por cada reset de TIIE 28d). Cada período
   es una "mini-opción" (*caplet*). Los de los primeros meses valen casi nada
   (poco tiempo para que la TIIE suba 224 bps). Los del año 2 valen más.

---

**El resultado**

Sumando los {_ncap_h} caplets:

> Prima total = **MXN {_cap_mxn_h/1e6:.2f}mm = USD {_cap_usd_h/1e3:.0f}K**
> ({_cap_pct_h:.3%} del nocional — un seguro muy barato dado lo lejos que está el techo)

El ahorro anual en el escenario base es **USD {_ann_sav_h/1e3:.0f}K**.
La prima se recupera en menos de un mes de ahorro.

---

**Limitaciones del modelo**

— Asumimos que la TIIE se queda en {_F_h:.2%} hacia adelante (no anticipamos más recortes).
  Si Banxico recorta más, el ahorro real sería aún mayor.

— Usamos volatilidad histórica ({_sig_h:.1%}) como aproximación.
  Un banco cotizaría con implied vol, que puede ser distinta.

— Este es un modelo académico. La prima real requiere cotización bancaria.
        """)

    # ── Estrategia activa: Forward USD/MXN (CME 6M Futures) ─────────────────
    st.markdown(
        "<hr style='border:none;border-top:1px solid rgba(136,128,152,0.20);margin:22px 0;'>",
        unsafe_allow_html=True)
    st.markdown("#### Estrategia activa recomendada: Forward de Tipo de Cambio (CME 6M Futures)")
    st.markdown(
        "<div class='section-note'>60% de ventas cubiertas vía CME Mexican Peso Futures (JUN 2027). "
        "Prima = CERO — el carry del diferencial TIIE–SOFR ya está incorporado en el precio forward. "
        "Precios CME observados mayo 2026.</div>",
        unsafe_allow_html=True)

    # Parámetros base
    _S0_fx       = 17.40
    _r_mxn_fx    = 0.0676
    _r_usd_fx    = 0.0355
    _sig_fx      = 0.0944
    _fwd_12m_fx  = 17.85          # CME JUN 2027
    _T_fx        = 1.0
    _noc_usd_fx  = 71.1e6         # 60% × USD 118.5mm ventas base
    _noc_mxn_fx  = _noc_usd_fx * 17.49   # contratos CME (spot CME jun-2026 = 17.49)
    _n_cont_fx   = int(_noc_mxn_fx / 500_000)
    _K_put_fx    = 16.50
    _prem_put_fx = 9.71           # Garman-Kohlhagen prima USD mm

    # CME forward curve (datos reales mayo 2026)
    _cme_lbls  = ["JUN 2026\n(spot)", "SEP 2026\n(3m)", "DEC 2026\n(6m)",
                  "JUN 2027\n(12m)", "SEP 2027\n(15m)"]
    _cme_fwds  = [17.49, 17.63, 17.76, 17.85, 18.03]

    # Escenarios: beneficio forward = (F - S) / F × N_USD
    _esc_fx = [
        ("Opt. 2027\n(MXN deprecia)",  19.13),
        ("Opt. 2026",                  18.24),
        ("Base 2026 / 2027",           17.29),
        ("Adv. 2026",                  16.40),
        ("Adv. 2027\n(MXN aprecia)",   15.63),
    ]
    _esc_lbl_fx = [r[0] for r in _esc_fx]
    _esc_ben_fx = [(_fwd_12m_fx - s) / _fwd_12m_fx * _noc_usd_fx / 1e6 for _, s in _esc_fx]
    _esc_col_fx = [COLORS["green"] if v >= 0 else COLORS["coral"] for v in _esc_ben_fx]
    _adv27_ben  = (_fwd_12m_fx - 15.63) / _fwd_12m_fx * _noc_usd_fx / 1e6
    _put_ben_adv = max(_K_put_fx - 15.63, 0.0) * _noc_mxn_fx / 15.63 / 1e6 - _prem_put_fx

    col_fx_a, col_fx_b = st.columns([1.3, 0.7])

    with col_fx_a:
        _txt_pos_fx = ["inside" if v < 0 else "outside" for v in _esc_ben_fx]
        _fig_fx = go.Figure(go.Bar(
            x=_esc_ben_fx, y=_esc_lbl_fx, orientation="h",
            marker_color=_esc_col_fx,
            text=[f"{v:+.1f}mm USD" for v in _esc_ben_fx],
            textposition=_txt_pos_fx,
            textfont=dict(size=12, color=COLORS["text"]),
            cliponaxis=False,
        ))
        _fig_fx.add_vline(x=0, line_width=1.2, line_color="rgba(136,128,152,0.45)")
        _fig_fx = inst_layout(_fig_fx, "Beneficio neto del forward por escenario USD/MXN", height=340, r=220)
        _fig_fx.update_layout(margin=dict(l=130, r=230, t=55, b=40), bargap=0.35)
        _fig_fx.update_xaxes(ticksuffix=" mm USD", automargin=False)
        _fig_fx.update_yaxes(automargin=False, tickfont=dict(size=13))
        st.plotly_chart(_fig_fx, use_container_width=True)
        st.markdown(
            f"<div style='font-size:12px;color:{COLORS['muted']};padding:2px 6px;'>"
            f"Forward JUN 2027 = <b style='color:{COLORS['text']};'>17.85 MXN/USD</b>. "
            f"El carry TIIE–SOFR hace que el forward cotice "
            f"<b style='color:{COLORS['green']};'>+{(_fwd_12m_fx/_S0_fx - 1)*100:.2f}% sobre el spot</b> — "
            f"la cobertura bloquea una tasa ya favorable, a prima cero."
            f"</div>", unsafe_allow_html=True)

    with col_fx_b:
        _cip_fx = _S0_fx * (1 + _r_mxn_fx) / (1 + _r_usd_fx)
        _kpi_rows_fx = [
            ("σ USD/MXN hist. (anualizada)",      f"{_sig_fx:.1%}",                   COLORS["muted"]),
            ("Nocional cubierto (60% ventas base)",f"USD {_noc_usd_fx/1e6:.1f}mm",    COLORS["muted"]),
            ("Contratos CME (500K MXN c/u)",       f"{_n_cont_fx:,d}",                COLORS["text"]),
            ("Spot USD/MXN (actual)",              f"{_S0_fx:.2f}",                   COLORS["text"]),
            ("Forward CME JUN 2027",               f"17.85 MXN/USD",                  COLORS["blue"]),
            ("Prima del forward",                  "CERO",                            COLORS["green"]),
            ("Carry positivo vs. spot",            f"+{(_fwd_12m_fx/_S0_fx - 1)*100:.2f}%",  COLORS["green"]),
            ("Beneficio en adverso 2027",          f"+USD {_adv27_ben:.1f}mm",        COLORS["green"]),
            ("Put 16.50 (neto prima) adverso 2027",f"USD {_put_ben_adv:.1f}mm",       COLORS["coral"]),
        ]
        _kpi_html_fx = "".join(
            f"<div style='display:flex;justify-content:space-between;align-items:baseline;"
            f"padding:5px 0;border-bottom:1px solid rgba(136,128,152,0.12);'>"
            f"<span style='font-size:11px;color:{COLORS['muted']};'>{lbl}</span>"
            f"<span style='font-size:12px;font-weight:700;color:{col};white-space:nowrap;'>{val}</span>"
            f"</div>"
            for lbl, val, col in _kpi_rows_fx
        )
        _c_surf_fx = COLORS["surface"]
        _c_teal_fx = COLORS["blue"]
        st.markdown(
            f"<div style='background:{_c_surf_fx};border:1px solid rgba(223,194,242,0.15);"
            f"border-left:4px solid {_c_teal_fx};border-radius:8px;padding:12px 14px;'>"
            f"<div style='font-size:14px;font-weight:700;color:{_c_teal_fx};"
            f"margin-bottom:10px;'>Forward CME 6M Futures — USD/MXN</div>"
            f"{_kpi_html_fx}"
            f"</div>", unsafe_allow_html=True)

    # CME forward curve
    st.markdown(
        f"<p style='font-size:13px;font-weight:600;color:{COLORS['text']};margin:18px 0 4px;'>"
        f"Curva forward CME — Mexican Peso Futures (cotizaciones mayo 2026)</p>",
        unsafe_allow_html=True)
    _cme_bar_colors = [COLORS["muted"]] + [
        COLORS["green"] if v > _S0_fx else COLORS["coral"]
        for v in _cme_fwds[1:]
    ]
    _fig_cme = go.Figure(go.Bar(
        x=_cme_lbls, y=_cme_fwds,
        marker_color=_cme_bar_colors,
        text=[f"{v:.2f}" for v in _cme_fwds],
        textposition="outside",
        textfont=dict(size=12, color=COLORS["text"]),
        cliponaxis=False,
    ))
    _fig_cme.add_hline(
        y=_S0_fx, line_dash="dot", line_color=COLORS["muted"],
        annotation_text=f"Spot {_S0_fx:.2f}",
        annotation_position="bottom right",
    )
    _fig_cme = inst_layout(_fig_cme, "USD/MXN por vencimiento CME (todos > spot → carry positivo)", height=280)
    _fig_cme.update_layout(margin=dict(l=50, r=80, t=55, b=40))
    _fig_cme.update_yaxes(range=[15.0, 20.5])
    st.plotly_chart(_fig_cme, use_container_width=True)

    with st.expander("¿Cómo se calculó la cobertura FX? — Fórmulas y lógica", expanded=False):
        _cip_show = _S0_fx * (1 + _r_mxn_fx) / (1 + _r_usd_fx)
        _put_gross_adv = max(_K_put_fx - 15.63, 0.0) * _noc_mxn_fx / 15.63 / 1e6
        st.markdown(f"""
**El problema que resuelve este instrumento**

Autlán cobra ingresos en dólares pero paga la mayoría de sus costos operativos en pesos.
Si el peso se aprecia (USD/MXN baja), cada dólar rinde *menos* pesos — los costos locales,
medidos en USD, suben. El OLS confirmó: **−1% en USD/MXN → −30.8 bps en margen operativo**.

En el escenario adverso 2027, USD/MXN llegaría a **15.63** — apreciación ~10% vs. spot {_S0_fx:.2f}.

La cobertura natural: **vender USD a plazo** → fijar hoy el tipo de cambio al que se convertirán
los dólares de ingreso a pesos para pagar costos.

---

**¿Por qué el forward y no una opción?**

| Instrumento | Prima (USD mm) | Beneficio adverso 2027 | Resultado neto |
|-------------|---------------|------------------------|----------------|
| **Forward CME JUN 2027** | **CERO** | **+USD {_adv27_ben:.1f}mm** | **+{_adv27_ben:.1f}mm** |
| Put K=16.50 (Garman-Kohlhagen) | −USD {_prem_put_fx:.1f}mm | +USD {_put_gross_adv:.1f}mm | **{_put_ben_adv:+.1f}mm** |

El forward gana porque el diferencial TIIE ({_r_mxn_fx:.2%}) − SOFR ({_r_usd_fx:.2%}) =
**+{(_r_mxn_fx-_r_usd_fx)*100:.0f} bps** hace que el forward cotice *por encima* del spot.
Autlán bloquea **17.85** cuando el spot es {_S0_fx:.2f} — ese +{(_fwd_12m_fx/_S0_fx-1)*100:.2f}%
de carry es un beneficio gratis incorporado en el precio.

---

**Fórmula 1 — Paridad de Interés Cubierta (CIP)**

El tipo de cambio forward teórico se calcula como:

> **F = S₀ × (1 + r_MXN)ᵀ / (1 + r_USD)ᵀ**

Con los datos actuales:

> F = {_S0_fx:.2f} × (1 + {_r_mxn_fx:.4f}) / (1 + {_r_usd_fx:.4f}) = **{_cip_show:.2f}** (teórico CIP)

El CME cotiza **17.85** — muy alineado con la CIP (diferencia por liquidez y base de mercado).

---

**Fórmula 2 — Cuantificación del beneficio del forward**

Si al vencimiento el tipo de cambio de mercado es S, el beneficio de haber bloqueado F es:

> **Beneficio (USD) = (F − S) / F × N_USD**

- Si S < F (MXN se apreció más de lo esperado): beneficio positivo — la cobertura protege.
- Si S > F (MXN se depreció): "costo de oportunidad" — pero la empresa gana operativamente.

Para el adverso 2027 (S = 15.63):

> Beneficio = (17.85 − 15.63) / 17.85 × USD {_noc_usd_fx/1e6:.1f}mm = **+USD {_adv27_ben:.1f}mm**

---

**Fórmula 3 — Prima put (Garman-Kohlhagen)**

Garman-Kohlhagen extiende Black-Scholes a divisas usando *dos* tasas de descuento:

> P = K·e^(−r_USD·T)·N(−d₂) − S₀·e^(−r_MXN·T)·N(−d₁)

> d₁ = [ ln(S₀/K) + (r_USD − r_MXN + σ²/2)·T ] / (σ·√T)

> d₂ = d₁ − σ·√T

Con S₀ = {_S0_fx:.2f}, K = {_K_put_fx:.2f}, r_USD = {_r_usd_fx:.2%}, r_MXN = {_r_mxn_fx:.2%},
σ = {_sig_fx:.1%} (volatilidad histórica), T = {_T_fx:.0f} año:

> Prima ≈ **USD {_prem_put_fx:.2f}mm** para USD {_noc_usd_fx/1e6:.1f}mm nocional

Como el beneficio bruto de la put en adverso (+{_put_gross_adv:.1f}mm) es menor que la prima
({_prem_put_fx:.1f}mm), el resultado neto es **negativo** → el forward domina.

---

**Implementación práctica**

- **Instrumento:** Venta de USD/MXN forward via CME 6M Futures (ticker 6M)
- **Vencimiento:** JUN 2027 (≈ 12 meses)
- **Contratos:** {_n_cont_fx:,d} contratos × 500,000 MXN = MXN {_noc_mxn_fx/1e6:,.0f}mm ≈ USD {_noc_usd_fx/1e6:.1f}mm
- **Margen inicial estimado:** ~5–10% del nocional en CME; alternativamente, forward bancario OTC sin margen

**Limitación:** Un forward obliga a vender al precio pactado — si el MXN se deprecia (escenario
optimista), Autlán pierde la ganancia cambiaria adicional. Ese es el costo de la cobertura.
Este es un modelo académico; la estrategia real requiere validación con el área de tesorería.
        """)


# ═══════════════════════════════════════════════════════════════════════════════
# TAB 9 — ANEXOS
# ═══════════════════════════════════════════════════════════════════════════════
with tabs[8]:
    st.subheader("Anexos — Detalle metodológico y Data Explorer")

    with st.expander("Comparación de modelos OLS", expanded=False):
        models = D["model_comparison"]
        if not models.empty:
            ok  = models[models["status"].eq("ok")].copy() if "status" in models.columns else models.copy()
            top = ok.sort_values("adj_r2", ascending=False).head(15) if "adj_r2" in ok.columns else ok
            keep = [c for c in ["model_name","y_col","x_cols","n_obs","r2","adj_r2","model_family"] if c in top.columns]
            _h = max(200, 38 + 32 * len(top))
            st.plotly_chart(
                styled_table(top[keep], height=_h),
                use_container_width=True)
            fig = px.bar(top.sort_values("adj_r2") if "adj_r2" in top.columns else top,
                x="adj_r2", y="model_name", orientation="h",
                color="model_family" if "model_family" in top.columns else None,
                color_discrete_sequence=ACCENT_SEQ)
            st.plotly_chart(inst_layout(fig, "Modelos por R² ajustada", height=480),
                            use_container_width=True)

    with st.expander("Coeficientes del mejor modelo multivariado", expanded=False):
        coef = D["multi_coef"]
        if not coef.empty:
            model_names = coef["model_name"].dropna().unique().tolist() if "model_name" in coef.columns else []
            if model_names:
                sel_model = st.selectbox("Modelo", model_names, key="annex_model")
                mc = coef[(coef["model_name"].eq(sel_model)) & (~coef["term"].eq("const"))].copy() \
                     if "term" in coef.columns else coef
                if not mc.empty:
                    show = [c for c in ["term","coef","std_err","t_stat","p_value"] if c in mc.columns]
                    _h2 = max(120, 38 + 32 * len(mc))
                    st.plotly_chart(styled_table(mc[show], height=_h2), use_container_width=True)
                    fig = px.bar(mc, x="coef", y="term", orientation="h",
                        color="coef",
                        color_continuous_scale=[[0,COLORS["red"]],[0.5,COLORS["surface"]],[1,COLORS["green"]]],
                        color_continuous_midpoint=0,
                        text="p_value" if "p_value" in mc.columns else None)
                    if "p_value" in mc.columns:
                        fig.update_traces(texttemplate="p=%{text:.3f}", textposition="outside")
                    st.plotly_chart(inst_layout(fig, f"Coeficientes | {sel_model}", height=340),
                                    use_container_width=True)

    with st.expander("Perfiles K-Means por régimen", expanded=False):
        km = D["kmeans_profiles"]
        if not km.empty:
            _h3 = max(120, 38 + 32 * len(km))
            st.plotly_chart(styled_table(km, height=_h3), use_container_width=True)

    with st.expander("Metodología de escenarios", expanded=False):
        st.markdown("""
**Supuestos macro:** USD/MXN, Manganeso y Oro provienen del promedio anual de Prophet
(Base=flat-forward desde último observado, Optimista=High80%, Adverso=Low80%).
SOFR y TIIE Base se mantienen flat en el último dato diario válido de la base
`Precios y tasas`; Optimista usa Low80% y Adverso usa High80%.

**Traducción a P&L — sensibilidades OLS (R²adj = 0.396):**
- Mg Bruto: +44.1 bps / +1% manganeso · +25.0 bps / +1% oro · +22.1 bps / +1% USD/MXN
- Mg Operativo: +60.2 bps manganeso · +50.0 bps oro · +30.8 bps FX
- Gasto Fin/Ventas: +160.7 bps / +100 bps tasa ponderada SOFR/TIIE
- Ventas: +1.96% / +1% manganeso · +0.25% / +1% oro; tasas excluidas por no ser canal directo atribuible
- Utilidad neta: utilidad operativa menos gasto financiero; SOFR/TIIE afectan solo debajo del EBIT

**Ponderación de tasas:** SOFR se aplica al tramo USD (~68% de deuda, USD 130mm);
TIIE al tramo MXN remanente (~32%). Esto evita duplicar el shock de tasas.
        """)

    st.markdown("#### Data Explorer")
    available = sorted([p.name for p in OUTPUT_DIR.glob("*.csv")])
    if available:
        sel_file = st.selectbox("Dataset", available,
            index=available.index("autlan_analysis_dataset.csv")
            if "autlan_analysis_dataset.csv" in available else 0)
        explorer = safe_read(sel_file)
        if not explorer.empty:
            cols = st.multiselect("Columnas", explorer.columns.tolist(),
                default=explorer.columns.tolist()[:min(10, len(explorer.columns))])
            view = explorer[cols] if cols else explorer
            st.dataframe(view, use_container_width=True, hide_index=True, height=340)
            st.download_button("Descargar CSV",
                data=view.to_csv(index=False).encode("utf-8"),
                file_name=sel_file, mime="text/csv")
