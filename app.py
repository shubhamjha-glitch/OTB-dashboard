import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from pathlib import Path
from datetime import datetime, date
import tempfile
import os
import re

# ============================================================
# PAGE CONFIG
# ============================================================
st.set_page_config(
    page_title="BUY PLAN – OTB",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("📊 BUY PLAN – OTB")
st.caption(f"As On {datetime.now().strftime('%d-%b-%Y')}")

# ============================================================
# CONSTANTS
# ============================================================
HIERARCHY = ["DIVISION", "SECTION", "DEPARTMENT", "ART_NM", "ATTRIBUTE"]
ARTICLE_HIERARCHY = ["DIVISION", "SECTION", "DEPARTMENT", "ART_NM"]

REGULAR_DATA_FIELDS = [
    "DIVISION", "SECTION", "DEPARTMENT", "ART_NM", "ATTRIBUTE",
    "ART_STATUS", "PREFERENCE", "BP_OCT-26", "BP_NOV-26", "BP_DEC-26",
    "BP_OND", "GRC_OCT", "PPO_OCT", "PPO_ALL", "TNA_Q", "CRRT_PO",
]
WINTER_DATA_FIELDS = [
    "DIVISION", "SECTION", "DEPARTMENT", "ART_NM", "ATTRIBUTE",
    "PREFERENCE", "ART_STATUS", "BP_OCT-26", "BP_WINTER", "GRC_OCT",
    "PPO_OCT", "PPO_ALL", "TNA_Q", "CRRT_PO",
]
REGULAR_OUTPUT = REGULAR_DATA_FIELDS + [
    "FR_OCT%", "VI_OCT%", "REQ_OCT", "OTB_OCT", "PO_FR%_OCT", "OTB_OND", "PO_FR% OND",
]
WINTER_OUTPUT = WINTER_DATA_FIELDS + [
    "FR%_OCT", "VI_OCT%", "REQ_OCT", "OTB_OCT", "PO_FR%_OCT", "OTB_WINTER", "PO_FR% WINTER",
]
REGULAR_CALCULATED_FIELDS = {
    "FR_OCT%": "=MIN(IFERROR(GRC_OCT/'BP_OCT-26',0),1)",
    "VI_OCT%": "=MIN(IFERROR((GRC_OCT+TNA_Q)/'BP_OCT-26',0),1)",
    "REQ_OCT": "=IFERROR('BP_OCT-26'-(GRC_OCT+TNA_Q),0)",
    "OTB_OCT": "=IFERROR('BP_OCT-26'-(GRC_OCT+PPO_OCT),0)",
    "PO_FR%_OCT": "=MIN(IFERROR((GRC_OCT+PPO_OCT)/'BP_OCT-26',0),1)",
    "OTB_OND": "=IFERROR(BP_OND-(GRC_OCT+PPO_ALL+CRRT_PO),0)",
    "PO_FR% OND": "=MIN(IFERROR((GRC_OCT+PPO_ALL+CRRT_PO)/BP_OND,0),1)",
}
WINTER_CALCULATED_FIELDS = {
    "FR%_OCT": "=MIN(IFERROR(GRC_OCT/'BP_OCT-26',0),1)",
    "VI_OCT%": "=MIN(IFERROR((GRC_OCT+TNA_Q)/'BP_OCT-26',0),1)",
    "REQ_OCT": "=IFERROR('BP_OCT-26'-(GRC_OCT+TNA_Q),0)",
    "OTB_OCT": "=IFERROR('BP_OCT-26'-(GRC_OCT+PPO_OCT),0)",
    "PO_FR%_OCT": "=MIN(IFERROR((GRC_OCT+PPO_OCT)/'BP_OCT-26',0),1)",
    "OTB_WINTER": "=IFERROR(BP_WINTER-(GRC_OCT+PPO_ALL+CRRT_PO),0)",
    "PO_FR% WINTER": "=MIN(IFERROR((GRC_OCT+PPO_ALL+CRRT_PO)/BP_WINTER,0),1)",
}
NUMERIC_BASE = [
    "BP_OCT-26", "BP_NOV-26", "BP_DEC-26", "BP_OND", "BP_WINTER",
    "GRC_OCT", "PPO_OCT", "PPO_ALL", "TNA_Q", "CRRT_PO",
]
PERCENT_COLS = [
    "FR_OCT%", "VI_OCT%", "PO_FR%_OCT", "PO_FR% OND", "FR%_OCT", "PO_FR% WINTER",
]

# ============================================================
# HELPERS
# ============================================================
def clean_columns(df):
    df = df.copy()
    df.columns = [re.sub(r"\s+", " ", str(c).strip()).upper() for c in df.columns]
    return df


def clean_text(x):
    if pd.isna(x):
        return ""
    return str(x).strip()


def numeric_series(s):
    return pd.to_numeric(
        s.astype(str).str.replace(",", "", regex=False).str.replace("%", "", regex=False),
        errors="coerce",
    ).fillna(0)


def numeric(df, cols):
    df = df.copy()
    for c in cols:
        if c in df.columns:
            df[c] = numeric_series(df[c])
    return df


def read_excel_file(uploaded_file, sheet_name=0, header=0):
    name = uploaded_file.name.lower()
    data = uploaded_file.getvalue()
    engine = "pyxlsb" if name.endswith(".xlsb") else None
    return pd.read_excel(BytesIO(data), sheet_name=sheet_name, header=header, engine=engine)


def find_sheet(uploaded_file, candidates):
    name = uploaded_file.name.lower()
    data = uploaded_file.getvalue()
    engine = "pyxlsb" if name.endswith(".xlsb") else None
    try:
        xls = pd.ExcelFile(BytesIO(data), engine=engine)
        names = xls.sheet_names
        upper = {str(x).strip().upper(): x for x in names}
        for candidate in candidates:
            if candidate.upper() in upper:
                return upper[candidate.upper()]
        return names[0] if names else 0
    except Exception:
        return 0


def normalize_base_data(df):
    df = clean_columns(df)
    # Some GM workbooks have the real headers in the first data row.
    if len(df):
        first = df.iloc[0].astype(str).str.strip().str.upper().tolist()
        if "DIVISION" in first and ("ART_NM" in first or "ART NM" in first or "ARTICLE NAME" in first):
            df.columns = first
            df = df.iloc[1:].reset_index(drop=True)
    df = df.dropna(axis=1, how="all").dropna(axis=0, how="all")
    return clean_columns(df)


def detect_col(df, aliases):
    normalized = {re.sub(r"[^A-Z0-9]", "", str(c).upper()): c for c in df.columns}
    for alias in aliases:
        key = re.sub(r"[^A-Z0-9]", "", alias.upper())
        if key in normalized:
            return normalized[key]
    for c in df.columns:
        cc = re.sub(r"[^A-Z0-9]", "", str(c).upper())
        for alias in aliases:
            aa = re.sub(r"[^A-Z0-9]", "", alias.upper())
            if aa and (aa in cc or cc in aa):
                return c
    return None


def ensure_base_columns(df, columns):
    df = df.copy()
    for c in columns:
        if c not in df.columns:
            df[c] = 0
    return df

# ============================================================
# BASE BUY PLAN
# ============================================================
def prepare_regular(df):
    df = normalize_base_data(df)
    df = ensure_base_columns(df, HIERARCHY + ["PREFERENCE", "ART_STATUS"] + NUMERIC_BASE)
    df = numeric(df, NUMERIC_BASE)
    for c in HIERARCHY + ["PREFERENCE", "ART_STATUS"]:
        df[c] = df[c].map(clean_text)
    return df


def prepare_winter(df):
    df = normalize_base_data(df)
    df = ensure_base_columns(df, HIERARCHY + ["PREFERENCE", "ART_STATUS"] + NUMERIC_BASE)
    df = numeric(df, NUMERIC_BASE)
    for c in HIERARCHY + ["PREFERENCE", "ART_STATUS"]:
        df[c] = df[c].map(clean_text)
    return df

# ============================================================
# EXACT USER FORMULAS
# ============================================================
def calculate_regular(df):
    df = df.copy()
    for c in ["BP_OCT-26", "BP_OND", "GRC_OCT", "PPO_OCT", "PPO_ALL", "TNA_Q", "CRRT_PO"]:
        if c not in df.columns:
            df[c] = 0
        df[c] = numeric_series(df[c])

    bp_oct = df["BP_OCT-26"]
    grc = df["GRC_OCT"]
    tna = df["TNA_Q"]
    ppo_oct = df["PPO_OCT"]
    ppo_all = df["PPO_ALL"]
    crrt = df["CRRT_PO"]
    bp_ond = df["BP_OND"]

    df["VI_OCT%"] = np.minimum(np.where(bp_oct != 0, (grc + tna) / bp_oct, 0), 1)
    df["REQ_OCT"] = (bp_oct - (grc + tna)).replace([np.inf, -np.inf], 0).fillna(0)
    df["OTB_OCT"] = (bp_oct - (grc + ppo_oct)).replace([np.inf, -np.inf], 0).fillna(0)
    df["PO_FR%_OCT"] = np.minimum(np.where(bp_oct != 0, (grc + ppo_oct) / bp_oct, 0), 1)
    df["OTB_OND"] = (bp_ond - (grc + ppo_all + crrt)).replace([np.inf, -np.inf], 0).fillna(0)
    df["PO_FR% OND"] = np.minimum(
        np.where(bp_ond != 0, (grc + ppo_all + crrt) / bp_ond, 0), 1
    )
    df["FR_OCT%"] = np.minimum(np.where(bp_oct != 0, grc / bp_oct, 0), 1)
    return df


def calculate_winter(df):
    df = df.copy()
    for c in ["BP_OCT-26", "BP_WINTER", "GRC_OCT", "PPO_OCT", "PPO_ALL", "TNA_Q", "CRRT_PO"]:
        if c not in df.columns:
            df[c] = 0
        df[c] = numeric_series(df[c])

    bp_oct = df["BP_OCT-26"]
    bp_winter = df["BP_WINTER"]
    grc = df["GRC_OCT"]
    tna = df["TNA_Q"]
    ppo_oct = df["PPO_OCT"]
    ppo_all = df["PPO_ALL"]
    crrt = df["CRRT_PO"]

    df["OTB_WINTER"] = (bp_winter - (grc + ppo_all + crrt)).replace([np.inf, -np.inf], 0).fillna(0)
    df["PO_FR% WINTER"] = np.minimum(
        np.where(bp_winter != 0, (grc + ppo_all + crrt) / bp_winter, 0), 1
    )
    df["FR%_OCT"] = np.minimum(np.where(bp_oct != 0, grc / bp_oct, 0), 1)
    df["VI_OCT%"] = np.minimum(np.where(bp_oct != 0, (grc + tna) / bp_oct, 0), 1)
    df["REQ_OCT"] = (bp_oct - (grc + tna)).replace([np.inf, -np.inf], 0).fillna(0)
    df["OTB_OCT"] = (bp_oct - (grc + ppo_oct)).replace([np.inf, -np.inf], 0).fillna(0)
    df["PO_FR%_OCT"] = np.minimum(
        np.where(bp_oct != 0, (grc + ppo_oct) / bp_oct, 0), 1
    )
    return df

# ============================================================
# GRC REPORT
# ============================================================
def prepare_grc(uploaded_file):
    sheet = find_sheet(uploaded_file, ["GRC REPORT", "GRC", "DATA"])
    df = clean_columns(read_excel_file(uploaded_file, sheet_name=sheet))

    aliases = {
        "ENTRY_DATE": ["ENTRY DATE"],
        "DIVISION": ["DIVISION"],
        "SECTION": ["SECTION"],
        "DEPARTMENT": ["DEPARTMENT"],
        "ART_NM": ["ARTICLE NAME", "ART_NM"],
        "ATTRIBUTE": ["ATTRIBUTE1", "ATTRIBUTE"],
        "GRC_QTY": ["GRC QTY", "GRC_QTY"],
    }
    cols = {k: detect_col(df, v) for k, v in aliases.items()}
    missing = [k for k, v in cols.items() if v is None]
    if missing:
        raise ValueError("GRC Report is missing required columns: " + ", ".join(missing))

    df["_ENTRY_DATE"] = pd.to_datetime(df[cols["ENTRY_DATE"]], errors="coerce", dayfirst=True)
    today = date.today()
    start = pd.Timestamp(today.year, today.month, 1)
    end = start + pd.offsets.MonthBegin(1)
    df = df[(df["_ENTRY_DATE"] >= start) & (df["_ENTRY_DATE"] < end)].copy()

    out = pd.DataFrame({
        "DIVISION": df[cols["DIVISION"]].map(clean_text),
        "SECTION": df[cols["SECTION"]].map(clean_text),
        "DEPARTMENT": df[cols["DEPARTMENT"]].map(clean_text),
        "ART_NM": df[cols["ART_NM"]].map(clean_text),
        "ATTRIBUTE": df[cols["ATTRIBUTE"]].map(clean_text),
        "GRC_OCT": numeric_series(df[cols["GRC_QTY"]]),
    })
    return out.groupby(HIERARCHY, dropna=False, as_index=False)["GRC_OCT"].sum()

# ============================================================
# TNA - ARTICLE LEVEL, NO ATTRIBUTE
# ============================================================
def prepare_tna(uploaded_file):
    sheet = find_sheet(uploaded_file, ["TNA", "TNA QTY", "DATA"])
    df = normalize_base_data(read_excel_file(uploaded_file, sheet_name=sheet))

    aliases = {
        "DIVISION": ["DIVISION", "DIV"],
        "SECTION": ["SECTION", "SEC"],
        "DEPARTMENT": ["DEPARTMENT", "DEPT"],
        "ART_NM": ["ARTICLE NAME", "ART_NM", "ARTICLE", "ARTCLE NAME"],
        "TNA_Q": ["TNA Q", "TNA_Q", "TNA QTY", "TNA", "QTY"],
    }
    cols = {k: detect_col(df, v) for k, v in aliases.items()}
    missing = [k for k, v in cols.items() if v is None]
    if missing:
        raise ValueError("TNA file is missing required columns: " + ", ".join(missing))

    out = pd.DataFrame({
        "DIVISION": df[cols["DIVISION"]].map(clean_text),
        "SECTION": df[cols["SECTION"]].map(clean_text),
        "DEPARTMENT": df[cols["DEPARTMENT"]].map(clean_text),
        "ART_NM": df[cols["ART_NM"]].map(clean_text),
        "TNA_Q": numeric_series(df[cols["TNA_Q"]]),
    })
    return out.groupby(ARTICLE_HIERARCHY, dropna=False, as_index=False)["TNA_Q"].sum()

# ============================================================
# CURRENT PPO - ARTICLE LEVEL, MONTH LOGIC
# PPO_OCT = October 2026 + all delivery months before October 2026
# PPO_ALL = total PPO quantity across all supplied months
# ============================================================
def parse_month_series(s):
    raw = s.astype(str).str.strip()
    parsed = pd.to_datetime(raw, errors="coerce", dayfirst=True)

    # Try common Excel month labels if direct parsing failed.
    mask = parsed.isna()
    if mask.any():
        for fmt in ["%b-%y", "%B-%y", "%b-%Y", "%B-%Y", "%m-%Y", "%m/%Y", "%Y-%m"]:
            tmp = pd.to_datetime(raw.where(mask), format=fmt, errors="coerce")
            parsed = parsed.fillna(tmp)
            mask = parsed.isna()
            if not mask.any():
                break
    return parsed


def prepare_ppo(uploaded_file):
    """
    CURRENT PPO file is used ONLY for CRRT_PO.

    User rule:
      Current PPO = CRRT_PO.
      If Current PPO is not uploaded, CRRT_PO = 0.

    The PO Pending + Schedule Date file is used separately for PPO_OCT
    and PPO_ALL.
    """
    sheet = find_sheet(uploaded_file, ["CURRENT PPO", "PPO", "DATA"])
    df = normalize_base_data(read_excel_file(uploaded_file, sheet_name=sheet))

    aliases = {
        "DIVISION": ["DIVISION", "DIV"],
        "SECTION": ["SECTION", "SEC"],
        "DEPARTMENT": ["DEPARTMENT", "DEPT"],
        "ART_NM": ["ARTICLE NAME", "ART_NM", "ARTICLE", "ARTCLE NAME"],
        "ATTRIBUTE": ["ATTRIBUTE1", "ATTRIBUTE", "ATTR"],
        "QTY": ["MNL PO QTY", "MNL_PO_QTY", "CURRENT PPO", "PPO QTY", "PPO_QTY", "PO QTY", "QTY"],
    }
    cols = {k: detect_col(df, v) for k, v in aliases.items()}
    required = ["DIVISION", "SECTION", "DEPARTMENT", "ART_NM", "QTY"]
    missing = [k for k in required if cols[k] is None]
    if missing:
        raise ValueError("Current PPO file is missing required columns: " + ", ".join(missing))

    out = pd.DataFrame({
        "DIVISION": df[cols["DIVISION"]].map(clean_text),
        "SECTION": df[cols["SECTION"]].map(clean_text),
        "DEPARTMENT": df[cols["DEPARTMENT"]].map(clean_text),
        "ART_NM": df[cols["ART_NM"]].map(clean_text),
        "CRRT_PO": numeric_series(df[cols["QTY"]]),
    })

    # Current PPO has no Attribute in the standard report, so it is
    # aggregated at article level and then applied to matching OTB rows.
    grouped = out.groupby(ARTICLE_HIERARCHY, dropna=False, as_index=False)["CRRT_PO"].sum()
    return grouped, out


# ============================================================
# PO PENDING + SCHEDULE DATE
# CRRT_PO = current pending PO qty after schedule-date matching
# ============================================================
def prepare_po_pending(po_file, schedule_file):
    """
    PO Pending + Schedule Date supplies PPO_OCT and PPO_ALL.

    User rules:
      * PPO_OCT = PO Pending Qty for September 2026 delivery + all
        deliveries before September 2026.
      * PPO_ALL = ALL PO Pending Qty, regardless of delivery month.
      * CRRT_PO is NOT taken from PO Pending. CRRT_PO comes from the
        separate Current PPO file.
    """
    po_sheet = find_sheet(po_file, ["PO PENDING", "PO_PENDING", "DATA"])
    sch_sheet = find_sheet(schedule_file, ["SCHEDULE DATE", "SCHEDULE", "DATA"])
    po = clean_columns(read_excel_file(po_file, sheet_name=po_sheet))
    sch = clean_columns(read_excel_file(schedule_file, sheet_name=sch_sheet))

    po_aliases = {
        "DIVISION": ["DIVISION"],
        "SECTION": ["SECTION"],
        "DEPARTMENT": ["DEPARTMENT"],
        "ART_NM": ["ARTCLE NAME", "ARTICLE NAME", "ART_NM"],
        "ATTRIBUTE": ["ATTRIBUTE1", "ATTRIBUTE", "ATTR"],
        "ORDER_NO": ["ORDER NO.", "ORDER NO", "ORDER_NO"],
        "PENDING_QTY": ["PENDING QTY", "PENDING_QTY"],
    }
    po_cols = {k: detect_col(po, v) for k, v in po_aliases.items()}
    missing = [k for k, v in po_cols.items() if v is None]
    if missing:
        raise ValueError("PO Pending file is missing required columns: " + ", ".join(missing))

    sch_order = detect_col(sch, ["ORDER NO", "ORDER_NO", "ORDER NO."])
    sch_date = detect_col(sch, ["SCHEDULE DATE", "SCHEDULE_DATE"])
    if sch_order is None or sch_date is None:
        raise ValueError("Schedule Date file must contain ORDER NO and SCHEDULE DATE.")

    po["_ORDER_KEY"] = po[po_cols["ORDER_NO"]].astype(str).str.strip().str.upper()
    sch["_ORDER_KEY"] = sch[sch_order].astype(str).str.strip().str.upper()
    sch["_SCHEDULE_DATE"] = pd.to_datetime(sch[sch_date], errors="coerce", dayfirst=True)

    # One schedule date per order prevents duplicate PO quantities.
    sch_map = (
        sch.dropna(subset=["_ORDER_KEY"])
        .groupby("_ORDER_KEY", as_index=False)["_SCHEDULE_DATE"]
        .max()
    )

    merged = po.merge(sch_map, on="_ORDER_KEY", how="left")
    merged["_PENDING_QTY"] = numeric_series(merged[po_cols["PENDING_QTY"]])
    merged = merged[merged["_PENDING_QTY"] != 0].copy()

    for target in ["DIVISION", "SECTION", "DEPARTMENT", "ART_NM", "ATTRIBUTE"]:
        merged[target] = merged[po_cols[target]].map(clean_text)

    merged["DELIVERY DATE"] = merged["_SCHEDULE_DATE"]
    merged["DELIVERY MONTH"] = np.where(
        merged["_SCHEDULE_DATE"].notna(),
        merged["_SCHEDULE_DATE"].dt.strftime("%b-%y"),
        "Unknown",
    )

    # Fixed to the Buy Plan October-26 logic.
    oct_start = pd.Timestamp(2026, 10, 1)
    nov_start = pd.Timestamp(2026, 11, 1)
    merged["PPO_OCT_FLAG"] = (
        merged["_SCHEDULE_DATE"].notna()
        & (merged["_SCHEDULE_DATE"] < nov_start)
    )

    # PPO_OCT = all delivery dates before November-26, i.e. Oct-26 + all earlier delivery dates.
    # PPO_ALL = every non-zero PO Pending quantity.
    merged["PPO_OCT"] = np.where(merged["PPO_OCT_FLAG"], merged["_PENDING_QTY"], 0)
    merged["PPO_ALL"] = merged["_PENDING_QTY"]

    po_month = (
        merged.groupby("DELIVERY MONTH", dropna=False, as_index=False)["_PENDING_QTY"]
        .sum()
        .rename(columns={"_PENDING_QTY": "PO_PENDING_QTY"})
    )

    po_hierarchy = (
        merged.groupby(HIERARCHY, dropna=False, as_index=False)[["PPO_OCT", "PPO_ALL"]]
        .sum()
    )

    return merged, po_month, po_hierarchy


# ============================================================
# MERGE HELPERS
# ============================================================
def merge_article_level(df, ext, value_cols):
    if ext is None or ext.empty:
        for c in value_cols:
            df[c] = 0
        return df
    out = df.drop(columns=value_cols, errors="ignore").merge(ext, on=ARTICLE_HIERARCHY, how="left")
    for c in value_cols:
        out[c] = numeric_series(out[c]) if c in out.columns else 0
    return out


def merge_hierarchy_level(df, ext, value_cols):
    if ext is None or ext.empty:
        for c in value_cols:
            df[c] = 0
        return df
    out = df.drop(columns=value_cols, errors="ignore").merge(ext, on=HIERARCHY, how="left")
    for c in value_cols:
        out[c] = numeric_series(out[c]) if c in out.columns else 0
    return out

# ============================================================
# EXCEL FORMATTING
# ============================================================
def style_excel_sheet(ws, freeze="A2"):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    # Required user format: Aptos, size 8.
    for row in ws.iter_rows():
        for cell in row:
            cell.font = Font(name="Aptos", size=8, bold=(cell.row == 1))
            cell.alignment = Alignment(vertical="center")

    if ws.max_row >= 1:
        for cell in ws[1]:
            cell.font = Font(name="Aptos", size=8, bold=True)

    ws.freeze_panes = freeze
    ws.auto_filter.ref = ws.dimensions
    ws.sheet_view.showGridLines = False

    for col_cells in ws.columns:
        letter = get_column_letter(col_cells[0].column)
        max_len = 0
        for cell in col_cells[:1000]:
            value = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, len(value))
        ws.column_dimensions[letter].width = min(max(max_len + 2, 9), 30)


def dataframe_to_excel(df, sheet_name="DATA"):
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        ws = writer.book[sheet_name]
        style_excel_sheet(ws)
        for cell in ws[1]:
            cell.font = __import__("openpyxl").styles.Font(name="Aptos", size=8, bold=True)
        for idx, col in enumerate(df.columns, start=1):
            if "%" in str(col):
                for r in range(2, ws.max_row + 1):
                    ws.cell(r, idx).number_format = "0.0%"
    bio.seek(0)
    return bio.getvalue()

# ============================================================
# CLOUD-SAFE OTB EXCEL OUTPUT
# Works on Streamlit Cloud/Linux - no pywin32 / Excel Desktop
# ============================================================
def create_otb_excel(df, mode="regular"):
    """
    Create a cloud-safe Excel workbook containing:
      1) source/base DATA
      2) calculated OTB output

    This intentionally does not create a native Excel PivotTable because
    Streamlit Cloud runs on Linux and Microsoft Excel/COM is unavailable.
    """
    mode = mode.lower()
    if df is None or df.empty:
        raise ValueError(f"No {mode} data available.")

    if mode == "regular":
        data_fields = REGULAR_DATA_FIELDS
        output_fields = REGULAR_OUTPUT
        data_sheet = "REGULAR DATA"
        otb_sheet = "OTB-REGULAR"
        calculated = calculate_regular(df.copy())
    else:
        data_fields = WINTER_DATA_FIELDS
        output_fields = WINTER_OUTPUT
        data_sheet = "WINTER DATA"
        otb_sheet = "OTB-WINTER"
        calculated = calculate_winter(df.copy())

    source_df = df[[c for c in data_fields if c in df.columns]].copy()
    output_df = calculated[[c for c in output_fields if c in calculated.columns]].copy()

    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        source_df.to_excel(writer, index=False, sheet_name=data_sheet)
        output_df.to_excel(writer, index=False, sheet_name=otb_sheet)

        for sheet in [data_sheet, otb_sheet]:
            ws = writer.book[sheet]
            style_excel_sheet(ws)
            for idx, col in enumerate(
                source_df.columns if sheet == data_sheet else output_df.columns, start=1
            ):
                if "%" in str(col):
                    for r in range(2, ws.max_row + 1):
                        ws.cell(r, idx).number_format = "0.0%"

    bio.seek(0)
    return bio.getvalue()

# ============================================================
# SIDEBAR INPUTS
# ============================================================
st.sidebar.header("📁 Input Files")
st.sidebar.caption("Upload the daily / revised files used for BUY PLAN – OTB.")

gm_file = st.sidebar.file_uploader(
    "1. GM_PRPO / Buy Plan Excel", type=["xlsx", "xls", "xlsb"], key="gm_file"
)
po_file = st.sidebar.file_uploader(
    "2. PO Pending", type=["xlsx", "xls", "xlsb"], key="po_file"
)
schedule_file = st.sidebar.file_uploader(
    "3. Schedule Date / Delivery Date", type=["xlsx", "xls", "xlsb"], key="schedule_file"
)
grc_file = st.sidebar.file_uploader(
    "4. GRC Report – Current Month", type=["xlsx", "xls", "xlsb"], key="grc_file"
)
tna_file = st.sidebar.file_uploader(
    "5. TNA Qty – Optional", type=["xlsx", "xls", "xlsb"], key="tna_file"
)
ppo_file = st.sidebar.file_uploader(
    "6. Current PPO – Optional", type=["xlsx", "xls", "xlsb"], key="ppo_file"
)

process = st.sidebar.button("🚀 Generate BUY PLAN – OTB", type="primary", use_container_width=True)

# ============================================================
# SESSION STATE
# ============================================================
for key, default in {
    "regular": None,
    "winter": None,
    "regular_data": None,
    "winter_data": None,
    "po_pending": None,
    "po_month": None,
    "ppo_detail": None,
    "error": None,
}.items():
    if key not in st.session_state:
        st.session_state[key] = default

# ============================================================
# PROCESS
# ============================================================
if process:
    st.session_state.error = None
    if gm_file is None:
        st.session_state.error = "Please upload the GM_PRPO / Buy Plan Excel file."
    else:
        try:
            with st.spinner("Reading files and generating BUY PLAN – OTB..."):
                regular_sheet = find_sheet(gm_file, ["REGULAR DATA", "REGULAR", "OTB-REGULAR"])
                winter_sheet = find_sheet(gm_file, ["WINTER DATA", "WINTER", "OTB-WINTER"])
                regular = prepare_regular(read_excel_file(gm_file, sheet_name=regular_sheet))
                winter = prepare_winter(read_excel_file(gm_file, sheet_name=winter_sheet))

                # GRC: current month only.
                if grc_file is not None:
                    grc = prepare_grc(grc_file)
                    regular = merge_hierarchy_level(regular, grc, ["GRC_OCT"])
                    winter = merge_hierarchy_level(winter, grc, ["GRC_OCT"])
                else:
                    regular["GRC_OCT"] = 0
                    winter["GRC_OCT"] = 0

                # TNA: article level, no ATTRIBUTE. Apply once to each attribute row.
                if tna_file is not None:
                    tna = prepare_tna(tna_file)
                    regular = merge_article_level(regular, tna, ["TNA_Q"])
                    winter = merge_article_level(winter, tna, ["TNA_Q"])
                else:
                    regular["TNA_Q"] = 0
                    winter["TNA_Q"] = 0

                # PO Pending + Schedule Date -> PPO_OCT and PPO_ALL.
                # PPO_OCT = October-26 delivery + every delivery before Oct-26.
                # PPO_ALL = ALL PO Pending Qty.
                if po_file is not None and schedule_file is not None:
                    po_detail, po_month, po_hierarchy = prepare_po_pending(po_file, schedule_file)
                    regular = merge_hierarchy_level(regular, po_hierarchy, ["PPO_OCT", "PPO_ALL"])
                    winter = merge_hierarchy_level(winter, po_hierarchy, ["PPO_OCT", "PPO_ALL"])
                    st.session_state.po_pending = po_detail
                    st.session_state.po_month = po_month
                else:
                    regular["PPO_OCT"] = 0
                    regular["PPO_ALL"] = 0
                    winter["PPO_OCT"] = 0
                    winter["PPO_ALL"] = 0
                    st.session_state.po_pending = None
                    st.session_state.po_month = None

                # Current PPO = CRRT_PO. If Current PPO is not uploaded, CRRT_PO = 0.
                if ppo_file is not None:
                    ppo, ppo_detail = prepare_ppo(ppo_file)
                    regular = merge_article_level(regular, ppo, ["CRRT_PO"])
                    winter = merge_article_level(winter, ppo, ["CRRT_PO"])
                    st.session_state.ppo_detail = ppo_detail
                else:
                    regular["CRRT_PO"] = 0
                    winter["CRRT_PO"] = 0
                    st.session_state.ppo_detail = None

                # DATA sheets contain ONLY base/source fields.
                # OTB formulas are NOT stored in DATA; they are Pivot Calculated Fields.
                for c in REGULAR_DATA_FIELDS:
                    if c not in regular.columns: regular[c] = 0
                for c in WINTER_DATA_FIELDS:
                    if c not in winter.columns: winter[c] = 0

                regular_data = regular[REGULAR_DATA_FIELDS].copy()
                winter_data = winter[WINTER_DATA_FIELDS].copy()
                regular_view = calculate_regular(regular_data.copy())[REGULAR_OUTPUT]
                winter_view = calculate_winter(winter_data.copy())[WINTER_OUTPUT]

                st.session_state.regular_data = regular_data
                st.session_state.winter_data = winter_data
                st.session_state.regular = regular_view
                st.session_state.winter = winter_view

            st.success("BUY PLAN – OTB generated successfully.")
        except Exception as e:
            st.session_state.error = str(e)

if st.session_state.error:
    st.error(st.session_state.error)

regular = st.session_state.regular
winter = st.session_state.winter
po_month = st.session_state.po_month

if regular is None or winter is None:
    st.info("Upload the GM_PRPO / Buy Plan workbook and click **Generate BUY PLAN – OTB**.")
    st.markdown(
        """
        ### Input logic
        - **GM_PRPO / Buy Plan:** base BP quantities and article attributes/status.
        - **PO Pending + Schedule Date:** `PPO_OCT` and `PPO_ALL` from Pending Qty + delivery date.
        - **PPO_OCT:** October-26 delivery + all delivery months before October-26.
        - **PPO_ALL:** all PO Pending Qty.
        - **GRC Report:** current-month GRC only = `GRC_OCT`.
        - **TNA:** optional article-level TNA; missing upload = 0.
        - **Current PPO:** same as `CRRT_PO`; missing upload = 0.
        """
    )
    st.stop()

# ============================================================
# FILTERS
# ============================================================
st.sidebar.divider()
st.sidebar.header("🔎 Filters")
filter_values = {}
for col in HIERARCHY:
    values = sorted([x for x in regular[col].dropna().unique().tolist() if str(x).strip() != ""], key=lambda x: str(x))
    filter_values[col] = st.sidebar.multiselect(
        col.title().replace("_", " "), values, default=[], key=f"filter_{col}"
    )

def apply_filters(df, filters):
    out = df.copy()
    for col, values in filters.items():
        if values:
            out = out[out[col].isin(values)]
    return out

regular_f = apply_filters(regular, filter_values)
winter_f = apply_filters(winter, filter_values)

sort_metric_options = [
    c for c in [
        "BP_OCT-26", "BP_OND", "BP_WINTER", "GRC_OCT", "TNA_Q", "PPO_OCT", "PPO_ALL",
        "CRRT_PO", "REQ_OCT", "OTB_OCT", "OTB_OND", "OTB_WINTER",
        "FR_OCT%", "VI_OCT%", "PO_FR%_OCT", "PO_FR% OND", "FR%_OCT", "PO_FR% WINTER",
    ] if c in regular_f.columns or c in winter_f.columns
]
sort_metric = st.sidebar.selectbox("Sort Metric", sort_metric_options, index=0)
sort_desc = st.sidebar.checkbox("Highest first", value=True)

# ============================================================
# TABS
# ============================================================
tab_dashboard, tab_regular, tab_winter, tab_download = st.tabs(
    ["📊 Dashboard", "🟦 Regular OTB", "🟨 Winter OTB", "⬇️ Downloads"]
)

# ============================================================
# DASHBOARD
# ============================================================
with tab_dashboard:
    st.subheader("BUY PLAN – OTB Dashboard")
    bp_oct = regular_f["BP_OCT-26"].sum()
    grc = regular_f["GRC_OCT"].sum()
    tna = regular_f["TNA_Q"].sum()
    ppo = regular_f["PPO_ALL"].sum()
    crrt = regular_f["CRRT_PO"].sum()
    otb = regular_f["OTB_OCT"].sum()
    fr = grc / bp_oct if bp_oct else 0

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("BP OCT", f"{bp_oct:,.0f}")
    c2.metric("GRC OCT", f"{grc:,.0f}")
    c3.metric("TNA Qty", f"{tna:,.0f}")
    c4.metric("PPO ALL", f"{ppo:,.0f}")
    c5.metric("CRRT PO", f"{crrt:,.0f}")
    c6.metric("FR OCT", f"{fr:.1%}")

    st.divider()
    st.subheader("1. High Plan Department Qty")
    dept = (
        regular_f.groupby("DEPARTMENT", dropna=False, as_index=False)
        .agg(BP_OCT=("BP_OCT-26", "sum"), BP_OND=("BP_OND", "sum"), OTB_OCT=("OTB_OCT", "sum"), GRC_OCT=("GRC_OCT", "sum"))
        .sort_values("BP_SEP", ascending=False).head(20)
    )
    st.dataframe(dept, use_container_width=True, hide_index=True)

    st.subheader("2. High / Low Fill Rate Achievement – Department")
    fr_dept = (
        regular_f.groupby("DEPARTMENT", dropna=False)
        .agg(BP=("BP_OCT-26", "sum"), GRC=("GRC_OCT", "sum"))
        .reset_index()
    )
    fr_dept["FR_ACH%"] = np.minimum(np.where(fr_dept["BP"] != 0, fr_dept["GRC"] / fr_dept["BP"], 0), 1)
    high_fr = fr_dept.sort_values("FR_ACH%", ascending=False).head(10)
    low_fr = fr_dept.sort_values("FR_ACH%", ascending=True).head(10)
    a, b = st.columns(2)
    with a:
        st.markdown("**Highest Fill Rate**")
        st.dataframe(high_fr.style.format({"FR_ACH%": "{:.1%}"}), use_container_width=True, hide_index=True)
    with b:
        st.markdown("**Lowest Fill Rate**")
        st.dataframe(low_fr.style.format({"FR_ACH%": "{:.1%}"}), use_container_width=True, hide_index=True)

    st.subheader("3. Month-wise PO Pending Qty and GRC Qty")
    if po_month is not None and not po_month.empty:
        month_df = po_month.copy()
        # GRC is current calendar month only.
        current_label = datetime.now().strftime("%b-%y").upper()
        month_df["GRC_QTY"] = np.where(month_df["DELIVERY MONTH"].astype(str).str.upper() == current_label, grc, 0)
        st.bar_chart(month_df.set_index("DELIVERY MONTH")[["PO_PENDING_QTY", "GRC_QTY"]])
        st.dataframe(month_df, use_container_width=True, hide_index=True)
    else:
        st.info("Upload both PO Pending and Schedule Date to show month-wise pending PO.")

    st.subheader("4. TNA Qty")
    st.success(f"TNA Qty: **{tna:,.0f}**" if tna > 0 else "No TNA file uploaded / TNA quantity is considered zero.")

# ============================================================
# DATA TABLES
# ============================================================
def show_data_table(df, columns):
    view = df[[c for c in columns if c in df.columns]].copy()
    if sort_metric in view.columns:
        view = view.sort_values(sort_metric, ascending=not sort_desc)
    st.write(f"**Rows:** {len(view):,}")
    cfg = {}
    for c in view.columns:
        if "%" in c:
            cfg[c] = st.column_config.NumberColumn(c, format="%.1f%%")
        elif c not in HIERARCHY and c not in ["ART_STATUS", "PREFERENCE"]:
            cfg[c] = st.column_config.NumberColumn(c, format="%,.2f")
    st.dataframe(view, use_container_width=True, height=650, hide_index=True, column_config=cfg)

with tab_regular:
    st.subheader("🟦 Regular OTB")
    show_data_table(regular_f, REGULAR_OUTPUT)

with tab_winter:
    st.subheader("🟨 Winter OTB")
    show_data_table(winter_f, WINTER_OUTPUT)

# ============================================================
# DOWNLOADS
# ============================================================
with tab_download:
    st.subheader("⬇️ Excel Downloads")
    st.caption("Cloud-safe Excel downloads include both the base DATA sheet and the calculated OTB output sheet. No Microsoft Excel Desktop or pywin32 is required.")

    st.markdown("### 🟦 Regular")
    regular_data_f = apply_filters(st.session_state.regular_data, filter_values)
    regular_bytes = dataframe_to_excel(regular_data_f, "REGULAR DATA")
    st.download_button(
        "⬇️ Download Regular OTB Excel", regular_bytes,
        "BUY_PLAN_OTB_REGULAR.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    regular_otb_bytes = create_otb_excel(regular_data_f, "regular")
    st.download_button(
        "⬇️ Download Regular OTB + Data",
        regular_otb_bytes,
        "BUY_PLAN_OTB_REGULAR_OCT_2026.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key="download_regular_otb",
    )

    st.markdown("### 🟨 Winter")
    winter_data_f = apply_filters(st.session_state.winter_data, filter_values)
    winter_bytes = dataframe_to_excel(winter_data_f, "WINTER DATA")
    st.download_button(
        "⬇️ Download Winter OTB Excel", winter_bytes,
        "BUY_PLAN_OTB_WINTER.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    winter_otb_bytes = create_otb_excel(winter_data_f, "winter")
    st.download_button(
        "⬇️ Download Winter OTB + Data",
        winter_otb_bytes,
        "BUY_PLAN_OTB_WINTER_OCT_2026.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key="download_winter_otb",
    )

st.divider()
st.caption("BUY PLAN – OTB | October 2026 | Regular + Winter | Cloud-safe output.")
