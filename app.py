from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional, Tuple
from io import BytesIO
from datetime import datetime
import shutil

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
DATA_FILE = DATA_DIR / "Amino_Eggsactly_Data_V1.xlsx"
MATCH_FILE = DATA_DIR / "Amino_Eggsactly_Rearing_Layer_Match.xlsx"
BACKUP_DIR = DATA_DIR / "backups"

st.set_page_config(
    page_title="Amino Eggsactly Graphs",
    page_icon="📈",
    layout="wide",
)

# -----------------------------------------------------------------------------
# Styling
# -----------------------------------------------------------------------------
st.markdown(
    """
    <style>
    .block-container {padding-top: 1.7rem; padding-bottom: 2rem;}
    .hero {
        background: #4d719d;
        color: white;
        border-radius: 10px;
        padding: 18px 22px;
        margin-bottom: 28px;
        box-shadow: 0 1px 4px rgba(0,0,0,.12);
    }
    .hero-title {font-size: 24px; font-weight: 750; margin: 0;}
    .hero-subtitle {font-size: 13px; opacity: .95; margin-top: 2px;}
    .small-muted {color: #6b7280; font-size: 12px;}
    div[data-testid="stMetricValue"] {font-size: 22px;}
    </style>
    """,
    unsafe_allow_html=True,
)

# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------
def first_existing(columns: List[str], candidates: List[str]) -> Optional[str]:
    for c in candidates:
        if c in columns:
            return c
    return None


def to_dt(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce")


def to_num(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce")


def normalize_percent(series: pd.Series) -> pd.Series:
    """Return percentage values on a 0-100 scale.

    Some Eggsactly standard columns are stored as decimals, e.g. 0.96,
    while actual production is often stored as 96. This keeps both aligned.
    """
    s = to_num(series)
    max_val = s.dropna().abs().max() if s.notna().any() else np.nan
    if pd.notna(max_val) and 0 < max_val <= 1.5:
        return s * 100
    return s


def safe_str(x) -> str:
    if pd.isna(x):
        return ""
    return str(x).strip()


def clean_text_series(series: pd.Series) -> pd.Series:
    return series.astype(str).str.strip()


def add_daily_week_coverage_flags(data: pd.DataFrame) -> pd.DataFrame:
    """Add daily-row coverage to each weekly row before filtering to Weekly.

    The export often contains both Daily and Weekly rows. The latest Weekly row
    can be created from only 1-6 Daily rows, which causes false end-of-line
    drops in production and feed. The 7_Days column is not reliable in this
    workbook because it can simply say "Other", so we calculate coverage from
    the Daily rows that share the same Farm/Flock/Week_End_Date.
    """
    if data is None or data.empty:
        return data

    required = {"Farm_Name", "Flock_Name", "Reporting_Period", "Week_End_Date"}
    if not required.issubset(data.columns):
        return data

    out = data.copy()

    # Normalised keys so trailing spaces in Eggsactly farm/flock names do not
    # break the daily-to-weekly match.
    out["_DailyCoverage_FarmKey"] = out["Farm_Name"].astype(str).str.strip()
    out["_DailyCoverage_FlockKey"] = out["Flock_Name"].astype(str).str.strip()
    out["_DailyCoverage_WeekEndKey"] = pd.to_datetime(out["Week_End_Date"], errors="coerce").dt.normalize()

    period = out["Reporting_Period"].astype(str).str.strip().str.lower()
    daily_counts = (
        out.loc[
            period.eq("daily") & out["_DailyCoverage_WeekEndKey"].notna(),
            ["_DailyCoverage_FarmKey", "_DailyCoverage_FlockKey", "_DailyCoverage_WeekEndKey"],
        ]
        .groupby(["_DailyCoverage_FarmKey", "_DailyCoverage_FlockKey", "_DailyCoverage_WeekEndKey"], dropna=False)
        .size()
        .reset_index(name="_Daily_Rows_In_Week")
    )

    if daily_counts.empty:
        out["_Daily_Rows_In_Week"] = np.nan
    else:
        out = out.merge(
            daily_counts,
            on=["_DailyCoverage_FarmKey", "_DailyCoverage_FlockKey", "_DailyCoverage_WeekEndKey"],
            how="left",
        )

    out.drop(
        columns=["_DailyCoverage_FarmKey", "_DailyCoverage_FlockKey", "_DailyCoverage_WeekEndKey"],
        inplace=True,
        errors="ignore",
    )
    return out


def get_summary_row(summary: pd.DataFrame, farm_col: str, flock_col: str, farm: str, flock: str) -> pd.Series:
    if summary is None or summary.empty or farm_col not in summary.columns or flock_col not in summary.columns:
        return pd.Series(dtype=object)
    m = clean_text_series(summary[farm_col]).eq(str(farm).strip()) & clean_text_series(summary[flock_col]).eq(str(flock).strip())
    rows = summary.loc[m]
    if rows.empty:
        return pd.Series(dtype=object)
    return rows.iloc[0]


def filter_by_expected_dates(df: pd.DataFrame, expected_hatch=None, expected_transfer=None, hatch_window_days: int = 10, transfer_window_days: int = 28) -> pd.DataFrame:
    """Narrow a farm/flock filter by dates if the summary sheet has them.

    This prevents graphs from looking identical when a flock name bucket is too broad
    or when the data source contains reused/split flock identifiers. The function is
    deliberately conservative: it only applies a date filter when it leaves at least
    one row; otherwise it falls back to the farm/flock result.
    """
    if df.empty:
        return df
    narrowed = df
    if expected_hatch is not None and pd.notna(expected_hatch) and "Hatch_Date" in narrowed.columns:
        h = to_dt(narrowed["Hatch_Date"])
        mask = (h - pd.to_datetime(expected_hatch)).abs().dt.days.le(hatch_window_days)
        if mask.any():
            narrowed = narrowed.loc[mask].copy()
    if expected_transfer is not None and pd.notna(expected_transfer) and "Transfer_Date" in narrowed.columns:
        t = to_dt(narrowed["Transfer_Date"])
        # Ignore empty Excel dates such as 1900-01-01.
        valid_t = t.dt.year.gt(1950)
        mask = valid_t & (t - pd.to_datetime(expected_transfer)).abs().dt.days.le(transfer_window_days)
        if mask.any():
            narrowed = narrowed.loc[mask].copy()
    return narrowed


def file_signature(path: Path) -> Tuple[str, int, float]:
    if not path.exists():
        return (str(path), 0, 0.0)
    stat = path.stat()
    return (str(path), stat.st_size, stat.st_mtime)


def validate_data_workbook(file_bytes: bytes) -> Tuple[bool, str]:
    """Validate a refreshed Amino/Eggsactly workbook before replacing the backend file."""
    required_sheets = {"DATA", "Standards ISA Floor"}
    try:
        xl = pd.ExcelFile(BytesIO(file_bytes))
    except Exception as exc:
        return False, f"Could not open Excel workbook: {exc}"

    missing_sheets = required_sheets - set(xl.sheet_names)
    if missing_sheets:
        return False, f"Workbook is missing required sheet(s): {', '.join(sorted(missing_sheets))}"

    try:
        preview = pd.read_excel(BytesIO(file_bytes), sheet_name="DATA", nrows=5)
    except Exception as exc:
        return False, f"Could not read DATA sheet: {exc}"

    required_cols = {"Farm_Name", "Flock_Name"}
    missing_cols = required_cols - set(preview.columns)
    if missing_cols:
        return False, f"DATA sheet is missing required column(s): {', '.join(sorted(missing_cols))}"

    return True, "Workbook looks valid."


def save_uploaded_data_workbook(file_bytes: bytes) -> Path:
    """Save uploaded DATA workbook, keeping a timestamped backup of the previous file."""
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)

    if DATA_FILE.exists():
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = BACKUP_DIR / f"Amino_Eggsactly_Data_V1_backup_{stamp}.xlsx"
        shutil.copy2(DATA_FILE, backup_path)

    DATA_FILE.write_bytes(file_bytes)
    return DATA_FILE


@st.cache_data(show_spinner="Loading Excel backend...")
def load_backend(data_sig, match_sig):
    if not DATA_FILE.exists():
        raise FileNotFoundError(f"Missing data file: {DATA_FILE}")
    if not MATCH_FILE.exists():
        raise FileNotFoundError(f"Missing match file: {MATCH_FILE}")

    data = pd.read_excel(DATA_FILE, sheet_name="DATA")
    standards = pd.read_excel(DATA_FILE, sheet_name="Standards ISA Floor")
    bridge = pd.read_excel(MATCH_FILE, sheet_name="Import_Bridge")

    # Optional summary sheets from the matching workbook. These are important
    # because farm/flock names alone can be reused or can represent split cohorts.
    # The summary sheets carry Hatch Date / Transfer Date, which makes the actual
    # data filter specific to the selected flock rather than accidentally using a
    # wider farm/flock bucket.
    try:
        layer_summary = pd.read_excel(MATCH_FILE, sheet_name="Layer_Flock_Summary")
    except Exception:
        layer_summary = pd.DataFrame()
    try:
        rearing_summary = pd.read_excel(MATCH_FILE, sheet_name="Rearing_Flock_Summary")
    except Exception:
        rearing_summary = pd.DataFrame()

    # Trim text columns and parse common date/number columns.
    for df in (data, standards, bridge, layer_summary, rearing_summary):
        if df is None or df.empty:
            continue
        for col in df.select_dtypes(include=["object"]).columns:
            df[col] = df[col].map(lambda v: v.strip() if isinstance(v, str) else v)

    # IMPORTANT: calculate Daily-row coverage BEFORE filtering to weekly.
    # This lets the app remove the latest incomplete Weekly row even when the
    # workbook's 7_Days column is not populated with a true/false flag.
    data = add_daily_week_coverage_flags(data)

    # IMPORTANT: graph weekly records only. The DATA sheet can contain multiple
    # reporting-period rows, and mixing daily/monthly helper rows with weekly
    # rows can make flock graphs look duplicated or incorrect.
    if "Reporting_Period" in data.columns:
        period = data["Reporting_Period"].astype(str).str.strip().str.lower()
        data = data.loc[period.eq("weekly")].copy()

    for col in ["Trans_Date", "Week_End_Date", "Hatch_Date", "Placed_Date", "Transfer_Date"]:
        if col in data.columns:
            data[col] = to_dt(data[col])
    for df in (bridge, layer_summary, rearing_summary):
        if df is None or df.empty:
            continue
        for col in ["Hatch Date", "Placed Date", "Transfer Date", "First Week", "Last Week"]:
            if col in df.columns:
                df[col] = to_dt(df[col])

    for col in data.columns:
        if col not in ["Farm_Name", "Flock_Name", "Breed", "Egg_Type", "Flock_Status", "FeedType", "Mill", "State", "ERP_System"]:
            # Keep dates as dates; coerce numeric-looking object columns only if useful.
            if data[col].dtype == object:
                converted = pd.to_numeric(data[col], errors="ignore")
                data[col] = converted

    return data, standards, bridge, layer_summary, rearing_summary


def get_age_col(df: pd.DataFrame) -> str:
    return first_existing(list(df.columns), ["Age in__weeks", "Flock_Age", "Age", "WOL"]) or "Age in__weeks"


def standard_age_col(df: pd.DataFrame) -> str:
    """Return the standards age-in-weeks column.

    The Amino/Eggsactly standards workbook has one default standards curve.
    Do not try to match standards by breed. Use the default curve and align it
    to actual flock records by age in weeks. In this file, Age is the real
    week number; WOL is present but mostly/all zero.
    """
    if "Age" in df.columns:
        return "Age"
    return first_existing(list(df.columns), ["Age in__weeks", "Flock_Age", "WOL"]) or "Age"


def filter_data_by_flock(data: pd.DataFrame, farm: str, flock: str, expected_hatch=None, expected_transfer=None) -> pd.DataFrame:
    if not farm or not flock:
        return data.iloc[0:0].copy()
    mask = (
        clean_text_series(data["Farm_Name"]).eq(str(farm).strip())
        & clean_text_series(data["Flock_Name"]).eq(str(flock).strip())
    )
    out = data.loc[mask].copy()
    out = filter_by_expected_dates(out, expected_hatch=expected_hatch, expected_transfer=expected_transfer)
    return out


def first_numeric_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    """Return the first candidate column that exists and has at least one numeric value."""
    for c in candidates:
        if c in df.columns:
            vals = pd.to_numeric(df[c], errors="coerce")
            if vals.notna().any():
                return c
    return None


def weighted_mean(grouped: pd.core.groupby.generic.DataFrameGroupBy, value_col: str, weight_col: Optional[str]) -> pd.DataFrame:
    """Age-level weighted mean. Falls back to simple mean if weights are unusable."""
    rows = []
    for age, g in grouped:
        v = pd.to_numeric(g[value_col], errors="coerce")
        if weight_col and weight_col in g.columns:
            w = pd.to_numeric(g[weight_col], errors="coerce")
            mask = v.notna() & w.notna() & w.gt(0)
            if mask.any():
                rows.append((age, float(np.average(v[mask], weights=w[mask]))))
                continue
        rows.append((age, float(v.mean()) if v.notna().any() else np.nan))
    return pd.DataFrame(rows, columns=["Graph_Age", value_col])


def latest_partial_week_mask(df: pd.DataFrame) -> pd.Series:
    """Return True for rows that belong to the latest incomplete weekly period.

    The Eggsactly export has Reporting_Period plus a 7_Days column. In some
    exports 7_Days is a boolean/full-week flag; in others it behaves like a
    day-count. We only remove rows from the latest age/date period and only when
    that period is clearly incomplete. Older partial transfer/depletion weeks are
    left alone unless they are the latest selected point.
    """
    if df.empty:
        return pd.Series(False, index=df.index)

    if "Graph_Age" not in df.columns:
        return pd.Series(False, index=df.index)

    # Determine the latest selected period. Prefer Week_End_Date, then Trans_Date,
    # then Graph_Age so multiple source rows for the same latest week are treated
    # together.
    latest_mask = pd.Series(False, index=df.index)
    for date_col in ["Week_End_Date", "Trans_Date"]:
        if date_col in df.columns:
            dates = to_dt(df[date_col])
            if dates.notna().any():
                latest_date = dates.max()
                latest_mask = dates.eq(latest_date)
                break
    else:
        ages = to_num(df["Graph_Age"])
        if ages.notna().any():
            latest_mask = ages.eq(ages.max())

    if not latest_mask.any():
        return pd.Series(False, index=df.index)

    # Best signal: count how many Daily rows were present for the same weekly
    # period. If the latest weekly row was built from fewer than 7 daily rows,
    # remove it when the checkbox is unticked.
    if "_Daily_Rows_In_Week" in df.columns:
        daily_count = pd.to_numeric(df["_Daily_Rows_In_Week"], errors="coerce")
        latest_daily_count = daily_count.loc[latest_mask]
        if latest_daily_count.notna().any():
            incomplete_by_daily_count = latest_mask & daily_count.lt(7)
            if incomplete_by_daily_count.any():
                return incomplete_by_daily_count

    if "7_Days" not in df.columns:
        return pd.Series(False, index=df.index)

    raw = df["7_Days"]

    # Boolean style: True/full week, False/partial week.
    if pd.api.types.is_bool_dtype(raw):
        incomplete = ~raw.fillna(True)
        return latest_mask & incomplete

    text = raw.astype(str).str.strip().str.lower()
    bool_like_full = text.isin(["true", "yes", "y", "full", "complete", "completed", "7", "7.0"])
    bool_like_partial = text.isin(["false", "no", "n", "partial", "incomplete", "0", "0.0"])

    # Numeric style: either 0/1 flag or count of days in the reporting week.
    num = pd.to_numeric(raw, errors="coerce")
    non_na = num.dropna()
    if not non_na.empty:
        unique_vals = set(non_na.unique().tolist())
        if unique_vals.issubset({0, 1, 0.0, 1.0}):
            incomplete = num.eq(0) | bool_like_partial
        else:
            incomplete = num.lt(7) | bool_like_partial
    else:
        incomplete = bool_like_partial & ~bool_like_full

    return latest_mask & incomplete.fillna(False)


def remove_latest_partial_week(df: pd.DataFrame, include_latest_partial_week: bool) -> Tuple[pd.DataFrame, int]:
    """Remove the latest incomplete weekly row(s) unless the user opts in."""
    if include_latest_partial_week or df.empty:
        return df, 0
    mask = latest_partial_week_mask(df)
    removed = int(mask.sum())
    if removed:
        return df.loc[~mask].copy(), removed
    return df, 0


def aggregate_actual(df: pd.DataFrame, phase_label: str) -> pd.DataFrame:
    """Build actual graph rows for the selected flock only.

    Important: the DATA sheet contains both raw Eggsactly actual fields and
    age-standard/helper fields. The helper fields have names like Production %,
    Feed intake per day__(g), Eggs per bird cum#, Egg weight (g), and Body weight
    (g). Those helper columns can be identical for every flock at the same age,
    which is why earlier versions made every flock graph look the same.

    This function therefore prefers the raw actual fields and calculates the
    derived graph metrics from the selected raw rows:
      - Production: HD% / HH% / % Lay
      - Mortality: cumulative mortality count divided by birds placed when needed
      - Feed intake: weekly feed usage divided by birds and 7 days
      - Eggs/bird: cumulative eggs divided by birds placed
      - Egg/body weight: Average_* raw fields
    """
    if df.empty:
        return df.copy()

    age_col = get_age_col(df)
    df = df.copy()
    df["Graph_Age"] = to_num(df[age_col])
    df = df[df["Graph_Age"].notna()].copy()
    if df.empty:
        return df

    # Numeric helper columns used for calculations.
    for col in [
        "HD%", "HH%", "% Lay", "Cumulative_Mortality", "Mortality",
        "Feed_Usage", "Opening_Bird_Numbers", "Closing_Bird_Numbers",
        "Birds Placed", "Cumulative_Eggs", "Total_Eggs", "Average_Egg_Weight",
        "Average_Body_Weight",
    ]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # Bird count / placement base. Use row-level birds for weighting and latest birds.
    bird_col = first_numeric_col(df, ["Closing_Bird_Numbers", "Opening_Bird_Numbers", "Birds Placed"])
    placed_col = first_numeric_col(df, ["Birds Placed", "Opening_Bird_Numbers", "Closing_Bird_Numbers"])
    if placed_col:
        df["_birds_placed_base"] = df.groupby("Graph_Age")[placed_col].transform("max")
        # For cumulative calculations, use flock-level placed birds where present.
        flock_base = pd.to_numeric(df[placed_col], errors="coerce").dropna()
        flock_base = float(flock_base.max()) if not flock_base.empty else np.nan
    else:
        df["_birds_placed_base"] = np.nan
        flock_base = np.nan

    if bird_col:
        df["_birds_weight"] = pd.to_numeric(df[bird_col], errors="coerce")
    else:
        df["_birds_weight"] = np.nan

    # Production actual — raw actuals first. Do not use the standard/helper column "Production %".
    prod_col = first_numeric_col(df, ["HD%", "HH%", "% Lay"])
    if prod_col:
        df["_production_actual"] = normalize_percent(df[prod_col])

    # Mortality cumulative actual. Cumulative_Mortality is often a bird count;
    # convert to percentage when the values are clearly not already percent.
    mort_col = first_numeric_col(df, ["Cumulative_Mortality", "Mortality"])
    if mort_col:
        mort = pd.to_numeric(df[mort_col], errors="coerce")
        max_mort = mort.dropna().max() if mort.notna().any() else np.nan
        if pd.notna(max_mort) and max_mort > 100 and pd.notna(flock_base) and flock_base > 0:
            df["_mortality_actual"] = mort / flock_base * 100
        else:
            df["_mortality_actual"] = normalize_percent(mort)

    # Feed intake actual in grams/bird/day from weekly feed usage.
    # If Feed_Usage appears to be tonnes instead of kg, this still gives the
    # right order after multiplying by 1,000,000 below only when values are small.
    if "Feed_Usage" in df.columns:
        feed = pd.to_numeric(df["Feed_Usage"], errors="coerce")
        birds = df["_birds_weight"]
        # Most Eggsactly exports use kg/week. Small values are usually tonnes/week.
        feed_multiplier = 1_000_000 if feed.dropna().median() < 1000 else 1000
        df["_feed_intake_actual"] = np.where(birds.gt(0), feed * feed_multiplier / birds / 7, np.nan)

    # Eggs per bird cumulative actual from cumulative eggs and placed birds.
    if "Cumulative_Eggs" in df.columns and pd.notna(flock_base) and flock_base > 0:
        df["_eggs_bird_actual"] = pd.to_numeric(df["Cumulative_Eggs"], errors="coerce") / flock_base

    if "Average_Egg_Weight" in df.columns:
        df["_egg_weight_actual"] = pd.to_numeric(df["Average_Egg_Weight"], errors="coerce")
    if "Average_Body_Weight" in df.columns:
        df["_body_weight_actual"] = pd.to_numeric(df["Average_Body_Weight"], errors="coerce")

    grouped = df.groupby("Graph_Age", dropna=True)
    out = pd.DataFrame({"Graph_Age": sorted(df["Graph_Age"].dropna().unique())})

    metric_map = {
        "Production % Actual": "_production_actual",
        "Mortality % cum Actual": "_mortality_actual",
        "Feed intake/day (g) Actual": "_feed_intake_actual",
        "Eggs/bird cum. Actual": "_eggs_bird_actual",
        "Egg weight (g) Actual": "_egg_weight_actual",
        "Body weight (g) Actual": "_body_weight_actual",
    }

    for out_col, calc_col in metric_map.items():
        if calc_col in df.columns and pd.to_numeric(df[calc_col], errors="coerce").notna().any():
            tmp = weighted_mean(grouped, calc_col, "_birds_weight").rename(columns={calc_col: out_col})
            out = out.merge(tmp, on="Graph_Age", how="left")

    if bird_col:
        birds = grouped[bird_col].sum(numeric_only=True).reset_index().rename(columns={bird_col: "Birds"})
        out = out.merge(birds, on="Graph_Age", how="left")

    out["Phase"] = phase_label
    return out.sort_values("Graph_Age")

def build_combined_life(data: pd.DataFrame, bridge: pd.DataFrame, layer_summary: pd.DataFrame, rearing_summary: pd.DataFrame, layer_farm: str, layer_flock: str, include_latest_partial_week: bool = False):
    bridge_rows = bridge[
        clean_text_series(bridge["Layer Farm"]).eq(str(layer_farm).strip())
        & clean_text_series(bridge["Layer Flock"]).eq(str(layer_flock).strip())
    ].copy()

    layer_meta = get_summary_row(layer_summary, "Layer Farm", "Layer Flock", layer_farm, layer_flock)
    layer_hatch = layer_meta.get("Hatch Date") if not layer_meta.empty else None
    layer_transfer = layer_meta.get("Transfer Date") if not layer_meta.empty else None

    layer_raw = filter_data_by_flock(data, layer_farm, layer_flock, expected_hatch=layer_hatch, expected_transfer=layer_transfer)
    age_col = get_age_col(layer_raw) if not layer_raw.empty else get_age_col(data)
    if not layer_raw.empty:
        layer_raw["Graph_Age"] = to_num(layer_raw[age_col])
        layer_raw = layer_raw[layer_raw["Graph_Age"] >= 17].copy()
        layer_raw, layer_partial_rows_removed = remove_latest_partial_week(layer_raw, include_latest_partial_week)
        layer_raw["Selected Layer Farm"] = layer_farm
        layer_raw["Selected Layer Flock"] = layer_flock
    else:
        layer_partial_rows_removed = 0

    rear_parts = []
    for _, row in bridge_rows.iterrows():
        rfarm = safe_str(row.get("Rearing Farm"))
        rflock = safe_str(row.get("Rearing Flock"))
        if not rfarm or not rflock:
            continue
        rmeta = get_summary_row(rearing_summary, "Rearing Farm", "Rearing Flock", rfarm, rflock)
        rhatch = rmeta.get("Hatch Date") if not rmeta.empty else layer_hatch
        r = filter_data_by_flock(data, rfarm, rflock, expected_hatch=rhatch)
        if not r.empty:
            r_age_col = get_age_col(r)
            r["Graph_Age"] = to_num(r[r_age_col])
            r = r[r["Graph_Age"] <= 17].copy()
            r, _rear_removed = remove_latest_partial_week(r, include_latest_partial_week)
            r["Matched Rearing Farm"] = rfarm
            r["Matched Rearing Flock"] = rflock
            rear_parts.append(r)

    rear_raw = pd.concat(rear_parts, ignore_index=True) if rear_parts else data.iloc[0:0].copy()

    rear_graph = aggregate_actual(rear_raw, "Rearing")
    layer_graph = aggregate_actual(layer_raw, "Layer")
    combined = pd.concat([rear_graph, layer_graph], ignore_index=True, sort=False)
    combined = combined.drop_duplicates(subset=["Graph_Age", "Phase"], keep="last").sort_values("Graph_Age")
    return combined, rear_raw, layer_raw, bridge_rows, layer_partial_rows_removed


def build_standard_graph(standards: pd.DataFrame) -> pd.DataFrame:
    if standards.empty:
        return standards.copy()
    scol_age = standard_age_col(standards)
    std = standards.copy()
    std["Graph_Age"] = to_num(std[scol_age])

    mapping = {
        "Production % Standard": ["PercentProdTableEggWK", "HD% std", "PercentProdTEWK", "PercentProdHEWK"],
        "Mortality % cum Standard": ["HenMortACM"],
        "Feed intake/day (g) Standard": ["FeedPerHen", "FeedPerHenWK"],
        "Eggs/bird cum. Standard": ["TEPerHHACM", "TableEggPerHHACM", "HEPerHHACM"],
        "Egg weight (g) Standard": ["EggWeight"],
        "Body weight (g) Standard": ["HenWeight"],
    }
    out = pd.DataFrame({"Graph_Age": std["Graph_Age"]})
    for out_col, candidates in mapping.items():
        src = first_existing(list(std.columns), candidates)
        if src:
            out[out_col] = to_num(std[src])

    for pct_col in ["Production % Standard", "Mortality % cum Standard"]:
        if pct_col in out.columns:
            out[pct_col] = normalize_percent(out[pct_col])

    out = out.dropna(subset=["Graph_Age"]).copy()
    out["Graph_Age"] = to_num(out["Graph_Age"])

    # There is only one default standards curve. If the sheet ever contains
    # duplicate rows for an age, collapse them to one row per week so the
    # dashed standard lines join cleanly across the graph.
    numeric_cols = [c for c in out.columns if c != "Graph_Age"]
    out = out.groupby("Graph_Age", as_index=False)[numeric_cols].mean(numeric_only=True)
    return out.sort_values("Graph_Age")


def add_period_shapes(fig: go.Figure, max_age: float):
    periods = [
        (0, 17, "Rearing"),
        (17, 35, "Layer 1"),
        (35, 55, "Layer 2"),
        (55, 72, "Layer 3"),
        (72, max(84, max_age), "Layer 4"),
    ]
    fills = ["rgba(214,214,245,0.20)", "rgba(235,220,242,0.22)", "rgba(245,222,222,0.22)", "rgba(242,229,199,0.24)", "rgba(213,231,244,0.24)"]
    for idx, (x0, x1, label) in enumerate(periods):
        fig.add_vrect(x0=x0, x1=x1, fillcolor=fills[idx], line_width=0, layer="below")
        if x1 <= max_age + 5:
            fig.add_annotation(x=(x0 + x1) / 2, y=1.03, yref="paper", text=label, showarrow=False, font=dict(size=10))


def make_graph(actual: pd.DataFrame, standards: pd.DataFrame, selected_metrics: List[str], show_standards: bool, age_min: float, age_max: float) -> go.Figure:
    actual = actual[(actual["Graph_Age"] >= age_min) & (actual["Graph_Age"] <= age_max)].copy()
    std = standards[(standards["Graph_Age"] >= age_min) & (standards["Graph_Age"] <= age_max)].copy()

    def plot_series_without_zeros(frame: pd.DataFrame, col: str) -> pd.Series:
        """Return a plotting series where 0 values are treated as missing.

        Eggsactly exports often use 0 as a placeholder when a metric has not
        been populated for a weekly row. Plotting those zeros creates false
        drop-offs to the bottom of the graph. We convert zeros to NaN and use
        Plotly connectgaps=True on the traces so the line joins the real
        non-zero points instead.
        """
        y = pd.to_numeric(frame[col], errors="coerce")
        # Treat zero placeholders as missing, then round display/hover values
        # to a maximum of 2 decimal places.
        return y.mask(y.eq(0)).round(2)

    metric_config = {
        "Production %": {"actual": "Production % Actual", "standard": "Production % Standard", "axis": "y", "dash": None},
        "Mortality % cum": {"actual": "Mortality % cum Actual", "standard": "Mortality % cum Standard", "axis": "y2", "dash": None},
        "Feed intake/day (g)": {"actual": "Feed intake/day (g) Actual", "standard": "Feed intake/day (g) Standard", "axis": "y3", "dash": None},
        "Eggs/bird cum.": {"actual": "Eggs/bird cum. Actual", "standard": "Eggs/bird cum. Standard", "axis": "y4", "dash": None},
        "Egg weight (g)": {"actual": "Egg weight (g) Actual", "standard": "Egg weight (g) Standard", "axis": "y5", "dash": None},
        "Body weight (g)": {"actual": "Body weight (g) Actual", "standard": "Body weight (g) Standard", "axis": "y6", "dash": None},
    }

    fig = go.Figure()
    max_x = float(max(age_max, actual["Graph_Age"].max() if not actual.empty else age_max))
    add_period_shapes(fig, max_x)

    for metric in selected_metrics:
        cfg = metric_config[metric]
        a_col = cfg["actual"]
        s_col = cfg["standard"]
        axis = cfg["axis"]

        if show_standards and s_col in std.columns and std[s_col].notna().any():
            fig.add_trace(
                go.Scatter(
                    x=std["Graph_Age"],
                    y=plot_series_without_zeros(std, s_col),
                    mode="lines",
                    name=f"{metric} Standard",
                    line=dict(width=2, dash="dash"),
                    hovertemplate=f"{metric} Standard: %{{y:.2f}}<extra></extra>",
                    connectgaps=True,
                    yaxis=axis,
                )
            )

        if a_col in actual.columns:
            actual_y = plot_series_without_zeros(actual, a_col)
            if actual_y.notna().any():
                fig.add_trace(
                    go.Scatter(
                        x=actual["Graph_Age"],
                        y=actual_y,
                        mode="lines+markers",
                        name=f"{metric} Actual",
                        line=dict(width=3),
                        marker=dict(size=5),
                        hovertemplate=f"{metric} Actual: %{{y:.2f}}<extra></extra>",
                        connectgaps=True,
                        yaxis=axis,
                    )
                )

    # Keep the extra y-axis scales outside the plotting area.
    # Plotly positions secondary axes in paper coordinates, so we reserve the
    # right-hand side of the figure by narrowing the x-axis domain. The actual
    # graph ends at x-domain 0.72; all right-side scales start after that.
    plot_domain_right = 0.72
    layout_axes = dict(
        yaxis=dict(title="Production %", side="left", rangemode="tozero"),
        yaxis2=dict(title="Mortality % cum", overlaying="y", side="right", anchor="free", position=0.76, showgrid=False, rangemode="tozero"),
        yaxis3=dict(title="Feed intake/day (g)", overlaying="y", side="right", anchor="free", position=0.82, showgrid=False, rangemode="tozero"),
        # Cumulative eggs per bird should always use a fixed 0–450 egg scale
        # so flocks can be compared consistently.
        yaxis4=dict(title="Eggs/bird cum.", overlaying="y", side="right", anchor="free", position=0.88, showgrid=False, range=[0, 450], dtick=50),
        yaxis5=dict(title="Egg weight (g)", overlaying="y", side="right", anchor="free", position=0.94, showgrid=False, rangemode="tozero"),
        yaxis6=dict(title="Body weight (g)", overlaying="y", side="right", anchor="free", position=1.0, showgrid=False, rangemode="tozero"),
    )

    fig.update_layout(
        height=680,
        margin=dict(l=70, r=320, t=55, b=90),
        xaxis=dict(title="Age in weeks", range=[age_min, age_max], dtick=4, domain=[0.0, plot_domain_right], hoverformat=".2f"),
        legend=dict(orientation="h", yanchor="bottom", y=-0.22, xanchor="center", x=0.5),
        hovermode="x unified",
        hoverlabel=dict(align="left"),
        plot_bgcolor="white",
        paper_bgcolor="white",
        **layout_axes,
    )
    fig.update_xaxes(showgrid=True, gridcolor="rgba(0,0,0,0.08)")
    fig.update_yaxes(showgrid=True, gridcolor="rgba(0,0,0,0.08)", tickformat=".2f")
    return fig


# -----------------------------------------------------------------------------
# Header
# -----------------------------------------------------------------------------
st.markdown(
    """
    <div class="hero">
        <div class="hero-title">📈 Graphs</div>
        <div class="hero-subtitle">Review flock performance trends against standards and stitched rearing-to-layer history.</div>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    st.header("Backend")
    st.caption("Upload a refreshed Amino_Eggsactly_Data_V1.xlsx after running the Excel refresh query.")

    uploaded_data = st.file_uploader(
        "Upload Amino_Eggsactly_Data_V1.xlsx",
        type=["xlsx"],
        key="amino_data_upload",
        help="This replaces the backend DATA workbook used for the graphs. The previous file is backed up in data/backups.",
    )

    if uploaded_data is not None:
        uploaded_bytes = uploaded_data.getvalue()
        is_valid, message = validate_data_workbook(uploaded_bytes)
        if is_valid:
            st.success(message)
            if st.button("Save uploaded data workbook and refresh", use_container_width=True):
                save_uploaded_data_workbook(uploaded_bytes)
                st.cache_data.clear()
                st.success("Uploaded workbook saved. Refreshing graphs...")
                st.rerun()
        else:
            st.error(message)

    if st.button("Clear cache / refresh data", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    st.write("**Data file**")
    st.code(str(DATA_FILE.relative_to(APP_DIR)))
    st.write("**Bridge file**")
    st.code(str(MATCH_FILE.relative_to(APP_DIR)))
    st.write("**Backup folder**")
    st.code(str(BACKUP_DIR.relative_to(APP_DIR)))

try:
    data, standards_raw, bridge, layer_summary, rearing_summary = load_backend(file_signature(DATA_FILE), file_signature(MATCH_FILE))
except Exception as e:
    st.error(f"Could not load backend Excel files: {e}")
    st.stop()

required_bridge_cols = {"Layer Farm", "Layer Flock", "Rearing Farm", "Rearing Flock"}
missing_bridge = required_bridge_cols - set(bridge.columns)
if missing_bridge:
    st.error(f"Import_Bridge is missing required columns: {', '.join(sorted(missing_bridge))}")
    st.stop()

if "Farm_Name" not in data.columns or "Flock_Name" not in data.columns:
    st.error("DATA sheet must contain Farm_Name and Flock_Name columns.")
    st.stop()

standards_graph = build_standard_graph(standards_raw)

# -----------------------------------------------------------------------------
# Selectors
# -----------------------------------------------------------------------------
st.caption("Hotkeys: Shift+←/→ = Farm | ←/→ = Flock. Standalone Excel-backed version.")

# Use Layer_Flock_Summary for selectors when available. It has one row per
# layer flock, while Import_Bridge can have multiple rows per layer flock when
# a layer flock was supplied by multiple rearing flocks.
if layer_summary is not None and not layer_summary.empty and {"Layer Farm", "Layer Flock"}.issubset(layer_summary.columns):
    valid_layers = layer_summary.dropna(subset=["Layer Farm", "Layer Flock"]).copy()
else:
    valid_layers = bridge.dropna(subset=["Layer Farm", "Layer Flock"]).copy()

valid_layers["Layer Farm"] = clean_text_series(valid_layers["Layer Farm"])
valid_layers["Layer Flock"] = clean_text_series(valid_layers["Layer Flock"])

valid_bridge = bridge.dropna(subset=["Layer Farm", "Layer Flock"]).copy()
valid_bridge["Layer Farm"] = clean_text_series(valid_bridge["Layer Farm"])
valid_bridge["Layer Flock"] = clean_text_series(valid_bridge["Layer Flock"])

farm_options = sorted(valid_layers["Layer Farm"].dropna().unique().tolist())
if not farm_options:
    st.warning("No layer farms found in the bridge/summary workbook.")
    st.stop()

c1, c2, c3 = st.columns([1, 1, 1.35])
with c1:
    layer_farm = st.selectbox("Farm", farm_options, index=0)

flock_options = sorted(valid_layers.loc[valid_layers["Layer Farm"].eq(layer_farm), "Layer Flock"].dropna().unique().tolist())
with c2:
    layer_flock = st.selectbox("Flock", flock_options, index=0 if flock_options else None)

with c3:
    confidence_options = sorted(valid_bridge.loc[
        (valid_bridge["Layer Farm"].eq(layer_farm)) & (valid_bridge["Layer Flock"].eq(layer_flock)),
        "Confidence"
    ].dropna().astype(str).unique().tolist()) if "Confidence" in valid_bridge.columns else []
    st.text_input("Match confidence", value=", ".join(confidence_options) if confidence_options else "Not supplied", disabled=True)

include_latest_partial_week = st.checkbox(
    "Include latest partial week",
    value=False,
    help="Default is off. Keeps the main graph on completed weekly periods only. Turn on to see the current/incomplete week as provisional data.",
)

combined, rear_raw, layer_raw, bridge_rows, layer_partial_rows_removed = build_combined_life(
    data,
    bridge,
    layer_summary,
    rearing_summary,
    layer_farm,
    layer_flock,
    include_latest_partial_week=include_latest_partial_week,
)

if layer_partial_rows_removed:
    st.info(f"Latest incomplete weekly layer row removed from the graph: {layer_partial_rows_removed} source row(s). Tick 'Include latest partial week' to show it as provisional.")

if combined.empty:
    st.warning("No graph data found for this layer flock and its matched rearing flocks.")
    st.dataframe(bridge_rows, use_container_width=True)
    st.stop()

if layer_raw.empty:
    st.warning("No layer actual rows were found for this selected farm/flock. The graph may only show rearing and standards. Check the bridge/summary match for this flock.")

max_actual_age = float(np.nanmax(combined["Graph_Age"])) if combined["Graph_Age"].notna().any() else 84.0

r1, r2 = st.columns([3, 1])
with r1:
    graph_range = st.selectbox(
        "Graph age range",
        ["Full life default (0-current)", "Rearing default (0-17)", "Laying default (17-current)", "Custom"],
        index=0,
    )
with r2:
    show_standards = st.checkbox("Show standards", value=True)

if graph_range == "Rearing default (0-17)":
    age_min, age_max = 0.0, 17.0
elif graph_range == "Laying default (17-current)":
    age_min, age_max = 17.0, max(84.0, max_actual_age)
elif graph_range == "Custom":
    age_min, age_max = st.slider("Custom age range", 0.0, max(100.0, max_actual_age), (0.0, max(84.0, max_actual_age)), step=1.0)
else:
    age_min, age_max = 0.0, max(84.0, max_actual_age)

metric_defaults = ["Production %", "Mortality % cum", "Feed intake/day (g)", "Eggs/bird cum."]
all_metrics = ["Production %", "Mortality % cum", "Feed intake/day (g)", "Eggs/bird cum.", "Egg weight (g)", "Body weight (g)"]
selected_metrics = st.multiselect("Metrics", all_metrics, default=metric_defaults)
if not selected_metrics:
    st.info("Select at least one metric to show the graph.")
    st.stop()

# KPI row
k1, k2, k3, k4, k5, k6, k7 = st.columns(7)
with k1:
    st.metric("Matched rearing flocks", bridge_rows["Rearing Flock"].dropna().nunique() if not bridge_rows.empty else 0)
with k2:
    st.metric("Current/max age", f"{max_actual_age:.2f} wks")
with k3:
    last_prod = combined["Production % Actual"].dropna().iloc[-1] if "Production % Actual" in combined and combined["Production % Actual"].notna().any() else np.nan
    st.metric("Latest production", "-" if pd.isna(last_prod) else f"{last_prod:.2f}%")
with k4:
    latest_birds = combined["Birds"].dropna().iloc[-1] if "Birds" in combined and combined["Birds"].notna().any() else np.nan
    st.metric("Latest birds", "-" if pd.isna(latest_birds) else f"{latest_birds:,.0f}")
with k5:
    conf = bridge_rows["Confidence"].dropna().astype(str).iloc[0] if "Confidence" in bridge_rows and bridge_rows["Confidence"].notna().any() else "-"
    st.metric("Match confidence", conf)
with k6:
    st.metric("Layer rows used", f"{len(layer_raw):,}")
with k7:
    st.metric("Partial week", "Included" if include_latest_partial_week else ("Removed" if layer_partial_rows_removed else "None"))

fig = make_graph(combined, standards_graph, selected_metrics, show_standards, age_min, age_max)
st.plotly_chart(fig, use_container_width=True)

# -----------------------------------------------------------------------------
# Detail tables
# -----------------------------------------------------------------------------
with st.expander("Rearing-to-layer bridge used for this graph", expanded=True):
    display_cols = [c for c in ["Match Group ID", "Layer Farm", "Layer Flock", "Layer Birds", "Rearing Farm", "Rearing Flock", "Rearing Birds", "Confidence", "Graph Mode", "Allocation Note"] if c in bridge_rows.columns]
    st.dataframe(bridge_rows[display_cols], use_container_width=True, hide_index=True)

with st.expander("Combined graph source data"):
    st.dataframe(combined, use_container_width=True, hide_index=True)
    csv = combined.to_csv(index=False).encode("utf-8")
    st.download_button(
        "Download combined graph data CSV",
        data=csv,
        file_name=f"combined_graph_{layer_farm}_{layer_flock}.csv".replace("/", "-"),
        mime="text/csv",
    )

with st.expander("Raw matched rearing and layer rows"):
    t1, t2 = st.tabs(["Rearing raw", "Layer raw"])
    with t1:
        st.dataframe(rear_raw, use_container_width=True, hide_index=True)
    with t2:
        st.dataframe(layer_raw, use_container_width=True, hide_index=True)

with st.sidebar:
    st.divider()
    st.write("**Loaded rows**")
    st.write(f"DATA weekly rows: {len(data):,}")
    st.write(f"Standards: {len(standards_raw):,}")
    st.write(f"Bridge: {len(bridge):,}")
    st.write(f"Layer summary: {len(layer_summary):,}")
    st.write(f"Rearing summary: {len(rearing_summary):,}")
