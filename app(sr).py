import os
import sqlite3
from io import BytesIO
from datetime import datetime

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer

import streamlit as st


# ============================================================
# EATON MDC SOLUTION CONFIGURATOR
# Data source: MDC_Master_V1.xlsx
# IMPORTANT:
#   The Excel workbook must be in the same folder as this app.py.
#   The workbook contains ONE sheet:
#       1R MDC (4 Configs) BOM
# ============================================================

st.set_page_config(
    page_title="Eaton MDC Solution Configurator",
    page_icon="🏢",
    layout="wide",
)

# Compact desktop UI. This block changes presentation only; it does not
# change the configuration, BOM, pricing, download, or save logic.
st.markdown("""
<style>
    .block-container {
        padding-top: 1rem;
        padding-bottom: 1rem;
        padding-left: 2rem;
        padding-right: 2rem;
        max-width: 1500px;
    }
    div[data-testid="stVerticalBlock"] {
        gap: 0.45rem;
    }
    div[data-testid="stHorizontalBlock"] {
        gap: 0.8rem;
    }
    label, .stMarkdown p, .stCaption {
        font-size: 12px !important;
    }
    div[data-testid="stTextInput"] input,
    div[data-testid="stNumberInput"] input,
    div[data-testid="stSelectbox"] div[data-baseweb="select"],
    div[data-testid="stRadio"] label {
        font-size: 13px !important;
    }
    div[data-testid="stTextInput"],
    div[data-testid="stNumberInput"],
    div[data-testid="stSelectbox"],
    div[data-testid="stRadio"] {
        margin-bottom: 0 !important;
    }
    button[kind] {
        min-height: 34px !important;
        font-size: 13px !important;
    }
    div[data-testid="stDownloadButton"] button {
        min-height: 36px !important;
    }
    hr { margin: 0.5rem 0 !important; }
</style>
""", unsafe_allow_html=True)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MASTER_FILE = os.path.join(BASE_DIR, "MDC_Master_V1.xlsx")
TRACKING_DB = os.path.join(BASE_DIR, "MDC_Tracking.db")
DEMO_INTERNAL_PASSWORD = "MDC@123"


# ============================================================
# HELPERS
# ============================================================

def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def numeric(value):
    """Return a float for valid Excel numbers; otherwise NaN."""
    try:
        if pd.isna(value):
            return float("nan")
        if isinstance(value, str):
            value = value.strip().replace(",", "")
            if value in ("", "#N/A", "N/A", "NA", "nan", "None"):
                return float("nan")
        return float(value)
    except Exception:
        return float("nan")


def money(value):
    try:
        return f"₹ {float(value):,.2f}"
    except Exception:
        return "N/A"


def internal_password():
    try:
        return st.secrets["MDC_INTERNAL_PASSWORD"]
    except Exception:
        return DEMO_INTERNAL_PASSWORD


# ============================================================
# DATABASE
# ============================================================

def init_tracking_db():
    conn = sqlite3.connect(TRACKING_DB)
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS configurations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            configuration_id TEXT UNIQUE,
            created_at TEXT,
            customer_name TEXT,
            customer_place TEXT,
            problem TEXT,
            solution TEXT,
            mdc_type TEXT,
            configuration TEXT,
            base_cost REAL,
            optional_cost REAL,
            pdu_cost REAL,
            total_cost REAL,
            margin_pct REAL,
            freight REAL,
            installation REAL,
            warranty_pct REAL,
            margin_price REAL,
            final_selling_price REAL,
            warranty_amount REAL,
            user_code TEXT
        )
    """)

    # Upgrade an older database created by the previous app.
    cols = [r[1] for r in cur.execute(
        "PRAGMA table_info(configurations)"
    ).fetchall()]

    if "user_code" not in cols:
        cur.execute("ALTER TABLE configurations ADD COLUMN user_code TEXT")

    cur.execute("""
        CREATE TABLE IF NOT EXISTS configuration_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            configuration_id TEXT,
            component_type TEXT,
            part_code TEXT,
            description TEXT,
            quantity REAL,
            uom TEXT,
            unit_cost REAL,
            total_cost REAL,
            unit_price REAL,
            total_price REAL
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS download_counter (
            id INTEGER PRIMARY KEY,
            download_count INTEGER NOT NULL DEFAULT 0
        )
    """)

    cur.execute("""
        INSERT OR IGNORE INTO download_counter (id, download_count)
        VALUES (1, 0)
    """)

    conn.commit()
    conn.close()


def increment_download_count():
    conn = sqlite3.connect(TRACKING_DB)
    cur = conn.cursor()

    cur.execute("""
        UPDATE download_counter
        SET download_count = download_count + 1
        WHERE id = 1
    """)

    count = cur.execute("""
        SELECT download_count
        FROM download_counter
        WHERE id = 1
    """).fetchone()[0]

    conn.commit()
    conn.close()
    return count


def generate_configuration_id():
    today = datetime.now().strftime("%Y%m%d")

    conn = sqlite3.connect(TRACKING_DB)
    cur = conn.cursor()

    count = cur.execute("""
        SELECT COUNT(*)
        FROM configurations
        WHERE configuration_id LIKE ?
    """, (f"MDC-{today}-%",)).fetchone()[0] + 1

    conn.close()
    return f"MDC-{today}-{count:04d}"


def save_configuration(
    configuration_id,
    bom,
    base_cost,
    optional_cost,
    pdu_cost,
    total_cost,
    margin_pct,
    freight,
    installation,
    warranty_pct,
    margin_price,
    final_selling_price,
    warranty_amount,
):
    conn = sqlite3.connect(TRACKING_DB)
    cur = conn.cursor()

    created_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    cur.execute("""
        INSERT OR REPLACE INTO configurations (
            configuration_id,
            created_at,
            customer_name,
            customer_place,
            problem,
            solution,
            mdc_type,
            configuration,
            base_cost,
            optional_cost,
            pdu_cost,
            total_cost,
            margin_pct,
            freight,
            installation,
            warranty_pct,
            margin_price,
            final_selling_price,
            warranty_amount,
            user_code
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        configuration_id,
        created_at,
        st.session_state.customer_name,
        st.session_state.customer_place,
        st.session_state.problem,
        st.session_state.solution,
        st.session_state.mdc_type,
        st.session_state.configuration,
        float(base_cost),
        float(optional_cost),
        float(pdu_cost),
        float(total_cost),
        float(margin_pct),
        float(freight),
        float(installation),
        float(warranty_pct),
        float(margin_price),
        float(final_selling_price),
        float(warranty_amount),
        st.session_state.user_code,
    ))

    cur.execute(
        "DELETE FROM configuration_items WHERE configuration_id = ?",
        (configuration_id,),
    )

    if bom is not None and not bom.empty:
        for _, row in bom.iterrows():
            cur.execute("""
                INSERT INTO configuration_items (
                    configuration_id,
                    component_type,
                    part_code,
                    description,
                    quantity,
                    uom,
                    unit_cost,
                    total_cost,
                    unit_price,
                    total_price
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                configuration_id,
                clean_text(row.get("Component Type")),
                clean_text(row.get("Part Code")),
                clean_text(row.get("Description")),
                float(numeric(row.get("Quantity", 0))) if pd.notna(row.get("Quantity")) else 0,
                clean_text(row.get("UOM")),
                float(numeric(row.get("Unit Cost"))) if pd.notna(row.get("Unit Cost")) else 0,
                float(numeric(row.get("Total Cost"))) if pd.notna(row.get("Total Cost")) else 0,
                float(numeric(row.get("Unit Price"))) if pd.notna(row.get("Unit Price")) else None,
                float(numeric(row.get("Total Price"))) if pd.notna(row.get("Total Price")) else None,
            ))

    conn.commit()
    conn.close()


init_tracking_db()


# ============================================================
# LOAD THE ONE-SHEET EXCEL MASTER
# EVERYTHING USED BY THE UI IS READ FROM MDC_Master_V1.xlsx
# ============================================================

@st.cache_data

def load_master():
    if not os.path.exists(MASTER_FILE):
        raise FileNotFoundError(
            f"MDC_Master_V1.xlsx was not found in: {BASE_DIR}"
        )

    sheet = pd.read_excel(
        MASTER_FILE,
        sheet_name=0,
        header=None,
        engine="openpyxl",
    )

    def get(row, col):
        if col >= len(row):
            return ""
        return clean_text(row.iloc[col])

    def is_part_header(value):
        return get(pd.Series([value]), 0).upper() in {
            "PART NUMBER", "PART NO", "PART CODE"
        }

    # --------------------------------------------------------
    # SINGLE-RACK CONFIGURATIONS
    # Excel layout:
    #   Solution 1 -> rows 1-25, A:E
    #   Solution 3 -> rows 1-25, F:J
    #   Solution 2 -> rows 26-50, A:E
    #   Solution 4 -> rows 26-50, F:J
    # --------------------------------------------------------
    block_info = {
        "Configuration 1": {
            "start": 0, "end": 25,
            "part": 0, "desc": 1, "qty": 2,
            "uom": 3, "price": 4,
        },
        "Configuration 3": {
            "start": 0, "end": 25,
            "part": 5, "desc": 6, "qty": 7,
            "uom": 8, "price": 9,
        },
        "Configuration 2": {
            "start": 25, "end": 50,
            "part": 0, "desc": 1, "qty": 2,
            "uom": 3, "price": 4,
        },
        "Configuration 4": {
            "start": 25, "end": 50,
            "part": 5, "desc": 6, "qty": 7,
            "uom": 8, "price": 9,
        },
    }

    config_rows = []
    component_rows = []

    for config_name, info in block_info.items():
        block = sheet.iloc[info["start"]:info["end"]]

        # Read the actual solution title from Excel.
        solution_title = get(block.iloc[0], 0)
        if not solution_title:
            solution_title = config_name

        # Remove Excel line breaks for cleaner UI text.
        solution_title = " ".join(solution_title.split())
        solution_title = solution_title.replace("SOLUTION 1", "").replace("SOLUTION 2", "")
        solution_title = solution_title.replace("SOLUTION 3", "").replace("SOLUTION 4", "")
        solution_title = " ".join(solution_title.split()).strip(" -")

        for local_index, (_, row) in enumerate(block.iterrows()):
            # First row = solution title
            # Second row = column headings
            if local_index in (0, 1):
                continue

            part = get(row, info["part"])
            desc = get(row, info["desc"])
            qty = numeric(row.iloc[info["qty"]])
            uom = get(row, info["uom"])
            price = numeric(row.iloc[info["price"]])

            if not part and not desc:
                continue

            if part.upper() in {"PART NUMBER", "PART NO", "PART CODE"}:
                continue

            # CTO is the configuration heading, not a priced BOQ line.
            if part.upper() == "CTO3M002":
                continue

            component_rows.append({
                "MDC Type": "Single Rack",
                "Configuration": config_name,
                "Configuration Title": solution_title,
                "Part Code": part,
                "Description": desc,
                "Quantity": 0.0 if pd.isna(qty) else float(qty),
                "UOM": uom if uom else "EA",
                "Unit Cost": price,
            })

        cfg_components = pd.DataFrame([
            r for r in component_rows
            if r["Configuration"] == config_name
        ])

        if cfg_components.empty:
            base_cost = 0.0
        else:
            base_cost = float(
                pd.to_numeric(cfg_components["Unit Cost"], errors="coerce")
                .fillna(0)
                .mul(
                    pd.to_numeric(cfg_components["Quantity"], errors="coerce")
                    .fillna(0)
                )
                .sum()
            )

        config_rows.append({
            "MDC Type": "Single Rack",
            "Configuration": config_name,
            "Configuration Title": solution_title,
            "Base Cost": base_cost,
        })

    configs = pd.DataFrame(config_rows)
    components = pd.DataFrame(component_rows)

    # --------------------------------------------------------
    # MULTIRACK PLACEHOLDERS
    # --------------------------------------------------------
    for n in range(1, 10):
        configs = pd.concat([
            configs,
            pd.DataFrame([{
                "MDC Type": "Multirack",
                "Configuration": f"Configuration {n}",
                "Configuration Title": f"Multirack Configuration {n} - XXX",
                "Base Cost": 0.0,
            }])
        ], ignore_index=True)

    # --------------------------------------------------------
    # OTHER OPTIONAL ITEMS
    # Locate the section by its Excel heading instead of using
    # hardcoded row numbers.
    # --------------------------------------------------------
    accessories = []
    optional_start = None
    optional_end = None

    for i in range(len(sheet)):
        text = " ".join(get(sheet.iloc[i], 0).split()).upper()
        if "OTHER OPTIONAL ITEMS" in text:
            optional_start = i + 1
            continue
        if optional_start is not None and "SINGLE PHASE PDU" in text:
            optional_end = i
            break

    if optional_start is not None:
        if optional_end is None:
            optional_end = len(sheet)

        for _, row in sheet.iloc[optional_start:optional_end].iterrows():
            part = get(row, 0)
            desc = get(row, 1)
            qty = numeric(row.iloc[2])
            uom = get(row, 3)
            price = numeric(row.iloc[4])

            if not part or not desc:
                continue

            if part.upper() in {"PART NUMBER", "PART NO", "PART CODE"}:
                continue

            accessories.append({
                "Part Code": part,
                "Description": desc,
                "Default Quantity": 1.0 if pd.isna(qty) or qty <= 0 else float(qty),
                "UOM": uom if uom else "EA",
                "Unit Cost": price,
            })

    accessories = pd.DataFrame(accessories)

    # --------------------------------------------------------
    # SINGLE PHASE PDU'S
    # Locate section dynamically and carry merged TYPE values.
    # --------------------------------------------------------
    pdus = []
    pdu_start = None

    for i in range(len(sheet)):
        text = " ".join(get(sheet.iloc[i], 0).split()).upper()
        if "SINGLE PHASE PDU" in text:
            pdu_start = i + 1
            break

    current_pdu_type = ""

    if pdu_start is not None:
        for _, row in sheet.iloc[pdu_start:].iterrows():
            part = get(row, 0)
            desc = get(row, 1)

            if not part and not desc:
                continue

            if part.upper() in {"PART NUMBER", "PART NO", "PART CODE"}:
                continue

            c13 = numeric(row.iloc[2])
            c19 = numeric(row.iloc[3])
            excel_type = get(row, 4)
            price = numeric(row.iloc[5])

            if excel_type:
                current_pdu_type = excel_type.upper()

            if not part or not desc or not current_pdu_type:
                continue

            pdus.append({
                "Part Code": part,
                "Description": desc,
                "C13": 0.0 if pd.isna(c13) else float(c13),
                "C19": 0.0 if pd.isna(c19) else float(c19),
                "Type": current_pdu_type,
                "UOM": "EA",
                "Unit Cost": price,
            })

    pdus = pd.DataFrame(pdus)

    for df in (configs, components, accessories, pdus):
        for col in df.columns:
            if df[col].dtype == object:
                df[col] = df[col].fillna("").astype(str).str.strip()

    return configs, components, accessories, pdus


try:
    configs_df, components_df, accessories_df, pdus_df = load_master()
except Exception as exc:
    st.error("Unable to load MDC_Master_V1.xlsx.")
    st.exception(exc)
    st.stop()


# ============================================================
# SESSION STATE
# ============================================================

defaults = {
    "mode": "Sales",
    "authenticated": False,
    "customer_name": "",
    "customer_place": "",
    "problem": "",
    "solution": "",
    "mdc_type": "Single Rack",
    "configuration": "Configuration 1",
    "accessory_qty": {},
    "pdu_qty": {},
    "margin_pct": 20.0,
    "freight": 0.0,
    "installation": 0.0,
    "warranty_pct": 0.0,
    "configuration_id": None,
    "configuration_saved": False,
    "user_code": "—",
    "user_count": 0,
}

for key, value in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = value

if st.session_state.configuration_id is None:
    st.session_state.configuration_id = generate_configuration_id()

if st.session_state.user_count == 0:
    conn = sqlite3.connect(TRACKING_DB)
    count = conn.execute("""
        SELECT download_count FROM download_counter WHERE id = 1
    """).fetchone()
    st.session_state.user_count = count[0] if count else 0
    conn.close()


# ============================================================
# UI HELPERS
# ============================================================

def section_header(text):
    st.html(f"""
    <div style="
        background:#003B71;
        color:white;
        padding:7px 12px;
        border-radius:5px;
        margin:10px 0 8px 0;
        font-size:14px;
        font-weight:700;
        letter-spacing:.2px;
    ">
        {text}
    </div>
    """)


def price_box(label, value):
    st.markdown(
        f"""
        <div style="padding:2px 0 4px 0;">
            <div style="font-size:12px;color:#475569;font-weight:600;margin-bottom:2px;">
                {label}
            </div>
            <div style="font-size:18px;font-weight:700;color:#003B71;white-space:nowrap;">
                {money(value)}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ============================================================
# USER CODE
# ============================================================

def generate_user_code():
    """Generate the user code entirely from the current Excel-driven selections."""
    codes = []

    config_text = clean_text(st.session_state.configuration)
    if config_text:
        number = config_text.split()[-1]
        codes.append(f"C{number}")

    # Fire suppression is detected from Excel descriptions.
    fire_selected = []
    for part, qty in st.session_state.accessory_qty.items():
        if numeric(qty) <= 0:
            continue
        match = accessories_df[
            accessories_df["Part Code"].astype(str).str.strip() == str(part).strip()
        ]
        if not match.empty:
            desc = clean_text(match.iloc[0]["Description"]).upper()
            if "FIRE" in desc:
                fire_selected.append(desc)

    if any("EXTERNAL" in x for x in fire_selected):
        codes.append("F-EXT")
    elif any("INTERNAL" in x or "IN-RACK" in x for x in fire_selected):
        codes.append("F-INT")

    # Camera is detected from Excel descriptions, not hardcoded part numbers.
    camera_selected = False
    for part, qty in st.session_state.accessory_qty.items():
        if numeric(qty) <= 0:
            continue
        match = accessories_df[
            accessories_df["Part Code"].astype(str).str.strip() == str(part).strip()
        ]
        if not match.empty and "CAMERA" in clean_text(match.iloc[0]["Description"]).upper():
            camera_selected = True
            break

    if camera_selected:
        codes.append("CAM")

    # Other accessory codes are detected by their Excel descriptions.
    accessory_code_map = [
        ("KEYBOARD", "KT"),
        ("CABLE MANAGER", "CM"),
        ("TOP CABLE TRAY", "TCT"),
        ("BRUSH PANEL", "BP"),
    ]

    for keyword, code in accessory_code_map:
        selected = False
        for part, qty in st.session_state.accessory_qty.items():
            if numeric(qty) <= 0:
                continue
            match = accessories_df[
                accessories_df["Part Code"].astype(str).str.strip() == str(part).strip()
            ]
            if not match.empty and keyword in clean_text(match.iloc[0]["Description"]).upper():
                selected = True
                break
        if selected:
            codes.append(code)

    # PDU code is obtained from Excel TYPE.
    for part, qty in st.session_state.pdu_qty.items():
        if numeric(qty) <= 0:
            continue
        pdu_row = pdus_df[
            pdus_df["Part Code"].astype(str).str.strip() == str(part).strip()
        ]
        if not pdu_row.empty:
            pdu_type = clean_text(pdu_row.iloc[0]["Type"]).upper()
            pdu_map = {
                "BASIC": "B-PDU",
                "METERED": "M-PDU",
                "SWITCHED": "S-PDU",
                "MANAGED": "MG-PDU",
            }
            if pdu_type in pdu_map:
                codes.append(pdu_map[pdu_type])
        break

    return "-".join(codes) if codes else "—"


def handle_excel_download():
    st.session_state.user_code = generate_user_code()
    st.session_state.user_count = increment_download_count()


# ============================================================
# DATA LOOKUPS
# ============================================================

def selected_config_record():
    match = configs_df[
        (configs_df["MDC Type"] == st.session_state.mdc_type)
        & (configs_df["Configuration"] == st.session_state.configuration)
    ]
    return match.iloc[0] if not match.empty else None


def selected_components():
    return components_df[
        (components_df["MDC Type"] == st.session_state.mdc_type)
        & (components_df["Configuration"] == st.session_state.configuration)
    ].copy()


# ============================================================
# BOM
# ============================================================

def build_bom():
    rows = []

    # Base configuration
    for _, r in selected_components().iterrows():
        cost = numeric(r["Unit Cost"])
        qty = numeric(r["Quantity"])

        rows.append({
            "S.No.": len(rows) + 1,
            "Component Type": "Base (Configuration)",
            "Part Code": clean_text(r["Part Code"]),
            "Description": clean_text(r["Description"]),
            "Quantity": 0 if pd.isna(qty) else float(qty),
            "UOM": clean_text(r["UOM"]),
            "Unit Cost": cost,
            "Total Cost": cost * qty if pd.notna(cost) and pd.notna(qty) else float("nan"),
            "Source": "Configuration",
        })

    # Optional accessories
    for _, r in accessories_df.iterrows():
        part = clean_text(r["Part Code"])
        qty = numeric(st.session_state.accessory_qty.get(part, 0))

        if pd.notna(qty) and qty > 0:
            cost = numeric(r["Unit Cost"])

            rows.append({
                "S.No.": len(rows) + 1,
                "Component Type": "Optional Accessory",
                "Part Code": part,
                "Description": clean_text(r["Description"]),
                "Quantity": float(qty),
                "UOM": clean_text(r["UOM"]),
                "Unit Cost": cost,
                "Total Cost": cost * qty if pd.notna(cost) else float("nan"),
                "Source": "Optional Accessory",
            })

    # Selected PDU
    for _, r in pdus_df.iterrows():
        part = clean_text(r["Part Code"])
        qty = numeric(st.session_state.pdu_qty.get(part, 0))

        if pd.notna(qty) and qty > 0:
            cost = numeric(r["Unit Cost"])

            desc = (
                f'{clean_text(r["Description"])} | '
                f'Type: {clean_text(r["Type"])} | '
                f'C13: {numeric(r["C13"]):g} | '
                f'C19: {numeric(r["C19"]):g}'
            )

            rows.append({
                "S.No.": len(rows) + 1,
                "Component Type": "PDU",
                "Part Code": part,
                "Description": desc,
                "Quantity": float(qty),
                "UOM": clean_text(r["UOM"]),
                "Unit Cost": cost,
                "Total Cost": cost * qty if pd.notna(cost) else float("nan"),
                "Source": "PDU",
            })

    return pd.DataFrame(rows)


def cost_summary(bom):
    if bom.empty:
        return 0.0, 0.0, 0.0, 0.0

    base_cost = float(
        pd.to_numeric(
            bom.loc[bom["Source"] == "Configuration", "Total Cost"],
            errors="coerce",
        ).fillna(0).sum()
    )

    optional_cost = float(
        pd.to_numeric(
            bom.loc[bom["Source"] == "Optional Accessory", "Total Cost"],
            errors="coerce",
        ).fillna(0).sum()
    )

    pdu_cost = float(
        pd.to_numeric(
            bom.loc[bom["Source"] == "PDU", "Total Cost"],
            errors="coerce",
        ).fillna(0).sum()
    )

    return base_cost, optional_cost, pdu_cost, base_cost + optional_cost + pdu_cost


def add_selling_prices(bom, total_cost, margin_pct, freight, installation):
    """
    IMPORTANT PRICE LOGIC

    Unit Price:
        Comes directly from the Unit Cost in MDC_Master_V1.xlsx.

    Total Price:
        Excel Unit Price x selected Quantity.

    Final Selling Price:
        Calculated separately from total cost using margin/freight/
        installation. It is NOT distributed back into the Excel line
        prices.
    """
    result = bom.copy()

    if result.empty:
        margin_price = 0.0
        final_selling_price = float(freight) + float(installation)
        return result, margin_price, final_selling_price

    result["Unit Price"] = pd.to_numeric(
        result["Unit Cost"], errors="coerce"
    )

    result["Total Price"] = (
        pd.to_numeric(result["Unit Price"], errors="coerce")
        * pd.to_numeric(result["Quantity"], errors="coerce")
    )

    margin_price = (
        total_cost / (1 - margin_pct / 100)
        if margin_pct < 100
        else 0.0
    )

    final_selling_price = (
        margin_price + float(freight) + float(installation)
    )

    return result, margin_price, final_selling_price


# ============================================================
# EXCEL / PDF OUTPUT
# ============================================================

def customer_table():
    return pd.DataFrame([
        ["Customer Name", st.session_state.customer_name],
        ["Customer Place", st.session_state.customer_place],
        ["Problem Description", st.session_state.problem],
        ["Solution", st.session_state.solution],
        ["MDC Type", st.session_state.mdc_type],
        ["Configuration", st.session_state.configuration],
        ["User Code", st.session_state.user_code],
        ["Date", datetime.now().strftime("%d-%m-%Y")],
    ], columns=["Field", "Value"])


def excel_bytes(internal=False, bom=None, final_price=0.0, cost_data=None):
    """Create ONE Excel worksheet containing customer details + BOQ + totals."""
    output = BytesIO()

    if bom is None:
        bom = build_bom()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        wb = writer.book
        ws = wb.create_sheet("MDC BOQ")
        writer.sheets["MDC BOQ"] = ws

        # ---------------- CUSTOMER / CONFIGURATION ----------------
        info = customer_table()
        ws.cell(row=1, column=1, value="EATON MDC SOLUTION CONFIGURATOR")
        ws.cell(row=2, column=1, value="Customer & Configuration Details")

        row_no = 4
        for _, r in info.iterrows():
            ws.cell(row=row_no, column=1, value=r["Field"])
            ws.cell(row=row_no, column=2, value=r["Value"])
            row_no += 1

        row_no += 1
        ws.cell(row=row_no, column=1, value="FINAL BOQ")
        row_no += 1

        if internal:
            headers = [
                "S.No.", "Component Type", "Part Code", "Description",
                "Quantity", "UOM", "Unit Cost", "Total Cost",
                "Unit Price", "Total Price"
            ]
        else:
            headers = [
                "S.No.", "Component Type", "Part Code", "Description",
                "Quantity", "UOM", "Unit Price", "Total Price"
            ]

        for col_no, header in enumerate(headers, 1):
            ws.cell(row=row_no, column=col_no, value=header)

        header_row = row_no
        row_no += 1

        for _, r in bom.iterrows():
            values = []
            if internal:
                values = [
                    r.get("S.No."), r.get("Component Type"), r.get("Part Code"),
                    r.get("Description"), r.get("Quantity"), r.get("UOM"),
                    r.get("Unit Cost"), r.get("Total Cost"),
                    r.get("Unit Price"), r.get("Total Price")
                ]
            else:
                values = [
                    r.get("S.No."), r.get("Component Type"), r.get("Part Code"),
                    r.get("Description"), r.get("Quantity"), r.get("UOM"),
                    r.get("Unit Price"), r.get("Total Price")
                ]

            for col_no, value in enumerate(values, 1):
                if pd.isna(value):
                    value = None
                ws.cell(row=row_no, column=col_no, value=value)
            row_no += 1

        # ---------------- PRICE SUMMARY ON SAME SHEET ----------------
        row_no += 1
        ws.cell(row=row_no, column=1, value="PRICE SUMMARY")
        row_no += 1

        if internal:
            summary = [
                ("Base Cost", base_cost),
                ("Optional Cost", optional_cost),
                ("PDU Cost", pdu_cost),
                ("Total Cost", total_cost),
                ("Margin %", margin_pct),
                ("Margin Price", margin_price),
                ("Freight", freight),
                ("Installation", installation),
                ("Warranty %", warranty_pct),
                ("Warranty Amount", margin_price * warranty_pct / 100),
                ("Final Selling Price", final_price),
            ]
        else:
            summary = [
                ("Final Selling Price", final_price),
            ]

        for label, value in summary:
            ws.cell(row=row_no, column=1, value=label)
            ws.cell(row=row_no, column=2, value=float(value))
            row_no += 1

        # ---------------- EXCEL FORMATTING ----------------
        title_fill = "003B71"
        section_fill = "005EB8"
        header_fill = "D9EAF7"

        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(headers))

        for cell in ws[1]:
            cell.fill = __import__("openpyxl").styles.PatternFill("solid", fgColor=title_fill)
            cell.font = __import__("openpyxl").styles.Font(color="FFFFFF", bold=True, size=16)

        for cell in ws[2]:
            cell.fill = __import__("openpyxl").styles.PatternFill("solid", fgColor=section_fill)
            cell.font = __import__("openpyxl").styles.Font(color="FFFFFF", bold=True, size=11)

        for c in range(1, len(headers) + 1):
            cell = ws.cell(header_row, c)
            cell.fill = __import__("openpyxl").styles.PatternFill("solid", fgColor=header_fill)
            cell.font = __import__("openpyxl").styles.Font(bold=True)
            cell.alignment = __import__("openpyxl").styles.Alignment(horizontal="center", vertical="center")

        # Currency formats
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, (int, float)) and cell.column in range(7, len(headers) + 1):
                    cell.number_format = '₹ #,##0.00'

        # Summary values are currency except percentage fields.
        summary_start = header_row + len(bom) + 3
        for rr in range(summary_start, row_no):
            label = ws.cell(rr, 1).value
            if label in ("Margin %", "Warranty %"):
                ws.cell(rr, 2).number_format = '0.00'
            else:
                ws.cell(rr, 2).number_format = '₹ #,##0.00'

        widths = {
            1: 10, 2: 23, 3: 22, 4: 65, 5: 12,
            6: 10, 7: 17, 8: 17, 9: 17, 10: 17
        }
        for col, width in widths.items():
            if col <= len(headers):
                ws.column_dimensions[__import__("openpyxl").utils.get_column_letter(col)].width = width

        ws.freeze_panes = f"A{header_row + 1}"
        ws.auto_filter.ref = f"A{header_row}:{__import__('openpyxl').utils.get_column_letter(len(headers))}{header_row + len(bom)}"

    output.seek(0)
    return output.getvalue()


def pdf_bytes(internal=False, bom=None, final_price=0.0):
    """Create a single-page/flowing PDF report for Sales or Internal use."""
    output = BytesIO()
    doc = SimpleDocTemplate(
        output,
        pagesize=landscape(A4),
        rightMargin=10 * mm,
        leftMargin=10 * mm,
        topMargin=10 * mm,
        bottomMargin=10 * mm,
        title="Eaton MDC Solution Configurator",
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "MdcTitle", parent=styles["Title"], fontSize=17,
        leading=20, alignment=TA_CENTER, spaceAfter=4
    )
    sub_style = ParagraphStyle(
        "MdcSub", parent=styles["Normal"], fontSize=9,
        alignment=TA_CENTER, spaceAfter=8
    )
    small = ParagraphStyle(
        "MdcSmall", parent=styles["Normal"], fontSize=7,
        leading=8
    )
    right_small = ParagraphStyle(
        "MdcRight", parent=small, alignment=TA_RIGHT
    )

    story = [
        Paragraph("EATON MDC SOLUTION CONFIGURATOR", title_style),
        Paragraph("Modular Data Center Solution Configuration & Pricing", sub_style),
    ]

    info = customer_table()
    info_data = []
    for _, r in info.iterrows():
        info_data.append([
            Paragraph(f"<b>{clean_text(r['Field'])}</b>", small),
            Paragraph(clean_text(r["Value"]), small),
        ])

    info_table = Table(info_data, colWidths=[42 * mm, 90 * mm])
    info_table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.35, colors.grey),
        ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#D9EAF7")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    story += [info_table, Spacer(1, 6 * mm)]

    if internal:
        headers = [
            "S.No.", "Type", "Part Code", "Description", "Qty", "UOM",
            "Unit Cost", "Total Cost", "Unit Price", "Total Price"
        ]
    else:
        headers = [
            "S.No.", "Type", "Part Code", "Description", "Qty", "UOM",
            "Unit Price", "Total Price"
        ]

    table_data = [headers]
    for _, r in bom.iterrows():
        desc = Paragraph(clean_text(r.get("Description")), small)
        if internal:
            vals = [
                clean_text(r.get("S.No.")), clean_text(r.get("Component Type")),
                clean_text(r.get("Part Code")), desc,
                clean_text(r.get("Quantity")), clean_text(r.get("UOM")),
                money(r.get("Unit Cost")), money(r.get("Total Cost")),
                money(r.get("Unit Price")), money(r.get("Total Price")),
            ]
        else:
            vals = [
                clean_text(r.get("S.No.")), clean_text(r.get("Component Type")),
                clean_text(r.get("Part Code")), desc,
                clean_text(r.get("Quantity")), clean_text(r.get("UOM")),
                money(r.get("Unit Price")), money(r.get("Total Price")),
            ]
        table_data.append(vals)

    if internal:
        widths = [12*mm, 28*mm, 27*mm, 85*mm, 12*mm, 12*mm, 25*mm, 27*mm, 25*mm, 27*mm]
    else:
        widths = [13*mm, 30*mm, 30*mm, 105*mm, 13*mm, 13*mm, 28*mm, 30*mm]

    boq_table = Table(table_data, colWidths=widths, repeatRows=1)
    boq_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#003B71")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("GRID", (0, 0), (-1, -1), 0.3, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (2, -1), "LEFT"),
        ("ALIGN", (4, 1), (-1, -1), "RIGHT"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    story += [Paragraph("FINAL BOQ", styles["Heading3"]), boq_table, Spacer(1, 5 * mm)]

    if internal:
        summary = [
            ["Base Cost", money(base_cost), "Optional Cost", money(optional_cost)],
            ["PDU Cost", money(pdu_cost), "Total Cost", money(total_cost)],
            ["Margin %", f"{margin_pct:.2f}%", "Margin Price", money(margin_price)],
            ["Freight", money(freight), "Installation", money(installation)],
            ["Warranty %", f"{warranty_pct:.2f}%", "Warranty Amount", money(margin_price * warranty_pct / 100)],
            ["FINAL SELLING PRICE", money(final_price), "", ""],
        ]
        summary_table = Table(summary, colWidths=[38*mm, 42*mm, 45*mm, 45*mm])
    else:
        summary_table = Table(
            [["FINAL SELLING PRICE", money(final_price)]],
            colWidths=[55*mm, 45*mm]
        )

    summary_table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F4F8FC")),
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
        ("SPAN", (0, -1), (2, -1)) if internal else ("SPAN", (0, 0), (0, 0)),
        ("TEXTCOLOR", (0, -1), (-1, -1), colors.HexColor("#003B71")),
    ]))
    story += [summary_table]

    doc.build(story)
    output.seek(0)
    return output.getvalue()


# ============================================================

# CONSTANTS USED BY UI
# No product part numbers are hardcoded here.
# Product details are read from MDC_Master_V1.xlsx.


# HEADER
# ============================================================

# ============================================================
# LIGHT / COMPACT DESKTOP UI
# ============================================================
st.markdown("""
<style>
/* ---------- overall page ---------- */
[data-testid="stAppViewContainer"] {
    background: #F8FAFC;
}

[data-testid="stMainBlockContainer"] {
    max-width: 1500px;
    padding-top: 1.0rem;
    padding-bottom: 2rem;
}

/* ---------- title ---------- */
.mdc-title-card {
    background: linear-gradient(135deg, #075EA8 0%, #003B71 100%);
    border-radius: 12px;
    padding: 20px 26px 18px 26px;
    margin: 0 0 12px 0;
    box-shadow: 0 5px 18px rgba(0,59,113,.13);
}

.mdc-title {
    color: #FFFFFF !important;
    font-size: 30px !important;
    font-weight: 800 !important;
    letter-spacing: .15px;
    line-height: 1.15 !important;
    margin: 0 !important;
}

.mdc-subtitle {
    color: #DDEEFF !important;
    font-size: 13px !important;
    font-weight: 500 !important;
    margin-top: 5px !important;
    line-height: 1.25 !important;
}

/* ---------- top information strip ---------- */
.mdc-meta-card {
    background: #FFFFFF;
    border: 1px solid #DCE7F2;
    border-radius: 10px;
    padding: 9px 14px;
    margin-bottom: 12px;
    box-shadow: 0 2px 8px rgba(15,23,42,.04);
}

.mdc-meta-grid {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 10px;
    text-align: center;
}

.mdc-meta-item {
    flex: 1;
    min-width: 0;
}

.mdc-meta-label {
    color: #64748B;
    font-size: 9px;
    font-weight: 800;
    letter-spacing: .55px;
}

.mdc-meta-value {
    color: #003B71;
    font-size: 14px;
    font-weight: 800;
    margin-top: 2px;
}

/* ---------- section cards ---------- */
[data-testid="stVerticalBlockBorderWrapper"] {
    border: 1px solid #DCE7F2 !important;
    border-radius: 11px !important;
    background: #FFFFFF !important;
    box-shadow: 0 2px 9px rgba(15,23,42,.035) !important;
}

.mdc-card-heading {
    display: flex;
    align-items: center;
    gap: 9px;
    color: #173B5E;
    font-size: 14px;
    font-weight: 800;
    letter-spacing: .15px;
    margin: 0 0 8px 0;
    padding-bottom: 7px;
    border-bottom: 1px solid #E8EFF6;
}

.mdc-number {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 23px;
    height: 23px;
    border-radius: 7px;
    background: #EAF3FB;
    color: #075EA8;
    font-size: 11px;
    font-weight: 800;
}

.mdc-mini-heading {
    color: #334E68;
    font-size: 11px;
    font-weight: 800;
    margin: 7px 0 3px 0;
    text-transform: uppercase;
    letter-spacing: .3px;
}

.mdc-help {
    color: #64748B;
    font-size: 10px;
    margin-top: -3px;
    margin-bottom: 5px;
}

/* ---------- compact Streamlit controls ---------- */
[data-testid="stWidgetLabel"] p {
    font-size: 11px !important;
    font-weight: 650 !important;
    color: #475569 !important;
    margin-bottom: 2px !important;
}

[data-testid="stTextInput"] input {
    font-size: 12px !important;
    min-height: 36px !important;
}

[data-testid="stNumberInput"] input {
    font-size: 12px !important;
}

[data-baseweb="select"] {
    font-size: 12px !important;
}

[data-baseweb="select"] * {
    font-size: 12px !important;
}

[data-testid="stRadio"] label,
[data-testid="stCheckbox"] label {
    font-size: 11px !important;
}

[data-testid="stRadio"] > div {
    gap: 8px !important;
}

[data-testid="stCheckbox"] {
    margin-bottom: -4px !important;
}

.stButton > button,
.stDownloadButton > button {
    min-height: 34px !important;
    font-size: 12px !important;
    border-radius: 7px !important;
}

/* ---------- small selected-detail panel ---------- */
.mdc-detail-box {
    background: #F7FAFD;
    border: 1px solid #E1EAF2;
    border-radius: 7px;
    padding: 7px 9px;
    margin-top: 5px;
}

.mdc-detail-title {
    color: #075EA8;
    font-size: 10px;
    font-weight: 800;
    margin-bottom: 2px;
}

.mdc-detail-text {
    color: #475569;
    font-size: 10px;
    line-height: 1.35;
}

/* ---------- accessory rows ---------- */
.mdc-accessory-row {
    padding: 2px 0;
}

/* Reduce default vertical gaps without changing functionality. */
[data-testid="stVerticalBlock"] {
    gap: .45rem;
}

@media (max-width: 900px) {
    .mdc-title { font-size: 25px !important; }
    .mdc-title-card { padding: 17px 20px 15px 20px; }
}
</style>
""", unsafe_allow_html=True)

st.html("""
<div class="mdc-title-card">
    <div class="mdc-title">Eaton MDC Solution Configurator</div>
    <div class="mdc-subtitle">Modular Data Center Solution Configuration &amp; Pricing</div>
</div>
""")

current_date = datetime.now().strftime("%d-%m-%Y")

st.html(f"""
<div class="mdc-meta-card">
    <div class="mdc-meta-grid">
        <div class="mdc-meta-item">
            <div class="mdc-meta-label">USER CODE</div>
            <div class="mdc-meta-value">{st.session_state.user_code}</div>
        </div>
        <div class="mdc-meta-item">
            <div class="mdc-meta-label">USER COUNT</div>
            <div class="mdc-meta-value">{st.session_state.user_count}</div>
        </div>
        <div class="mdc-meta-item">
            <div class="mdc-meta-label">DATE</div>
            <div class="mdc-meta-value">{current_date}</div>
        </div>
    </div>
</div>
""")


# ============================================================
# SIDEBAR ACCESS
# ============================================================
with st.sidebar:
    st.header("User Access")

    mode = st.radio(
        "Select User Type",
        ["Sales", "Internal – MDC"],
        index=0 if st.session_state.mode == "Sales" else 1,
    )

    if mode != st.session_state.mode:
        st.session_state.mode = mode

        if mode == "Sales":
            st.session_state.authenticated = False

        st.rerun()

    if mode == "Internal – MDC":
        if not st.session_state.authenticated:
            st.warning("Internal MDC access requires a password.")

            pwd = st.text_input(
                "MDC Password",
                type="password",
            )

            if st.button(
                "Unlock Internal Mode",
                use_container_width=True,
            ):
                if pwd == internal_password():
                    st.session_state.authenticated = True
                    st.rerun()
                else:
                    st.error("Incorrect password.")
        else:
            st.success("Internal mode unlocked.")

            if st.button(
                "Lock Internal Mode",
                use_container_width=True,
            ):
                st.session_state.authenticated = False
                st.session_state.mode = "Sales"
                st.rerun()

is_internal = (
    st.session_state.mode == "Internal – MDC"
    and st.session_state.authenticated
)


# ============================================================
# 1 + 2 / TOP SPLIT SCREEN
# ============================================================
left_top, right_top = st.columns(2, gap="medium")

# ============================================================
# 1. CUSTOMER DETAILS
# ============================================================
with left_top:
    with st.container(border=True):
        st.markdown(
            '<div class="mdc-card-heading">'
            '<span class="mdc-number">01</span>'
            '<span>CUSTOMER DETAILS</span>'
            '</div>',
            unsafe_allow_html=True,
        )

        customer_name = st.text_input(
            "Customer Name",
            value=st.session_state.customer_name,
            key="customer_name_input",
            placeholder="Enter customer name",
        )
        st.session_state.customer_name = customer_name.strip()

# ============================================================
# 2. MDC TYPE & CONFIGURATION
# ============================================================
with right_top:
    with st.container(border=True):
        st.markdown(
            '<div class="mdc-card-heading">'
            '<span class="mdc-number">02</span>'
            '<span>MDC TYPE &amp; CONFIGURATION</span>'
            '</div>',
            unsafe_allow_html=True,
        )

        mdc_type = st.radio(
            "MDC Type",
            ["Single Rack", "Multirack"],
            horizontal=True,
            index=0 if st.session_state.mdc_type == "Single Rack" else 1,
            key="mdc_type_selection",
        )

        if mdc_type != st.session_state.mdc_type:
            st.session_state.mdc_type = mdc_type
            st.session_state.configuration = "Configuration 1"
            st.session_state.accessory_qty = {}
            st.session_state.pdu_qty = {}
            st.session_state.configuration_id = generate_configuration_id()
            st.session_state.configuration_saved = False
            st.rerun()

        available = configs_df[
            configs_df["MDC Type"] == st.session_state.mdc_type
        ].copy()

        labels = available["Configuration"].tolist()

        configuration_display_names = {
            "Configuration 1":
                "Configuration 1 - 1SR, 42U×800W×1200D, 3.5KW, W/O Dehumidifier",
            "Configuration 2":
                "Configuration 2 - 1SR, 42U×800W×1200D, 3.5KW, Dehumidifier",
            "Configuration 3":
                "Configuration 3 - 1SR, 42U×800W×1200D, 7KW, W/O Dehumidifier",
            "Configuration 4":
                "Configuration 4 - 1SR, 42U×800W×1200D, 7KW, Dehumidifier",
        }

        if labels:
            st.session_state.configuration = st.selectbox(
                "Select Configuration",
                labels,
                index=(
                    labels.index(st.session_state.configuration)
                    if st.session_state.configuration in labels
                    else 0
                ),
                format_func=lambda x: configuration_display_names.get(x, x),
                key="configuration_selection",
            )


# ============================================================
# 3 + 4 / BOTTOM SPLIT SCREEN
# ============================================================
left_bottom, right_bottom = st.columns(2, gap="medium")

# ============================================================
# 3. PDU SELECTION
# ============================================================
with left_bottom:
    with st.container(border=True):
        st.markdown(
            '<div class="mdc-card-heading">'
            '<span class="mdc-number">03</span>'
            '<span>PDU SELECTION</span>'
            '</div>',
            unsafe_allow_html=True,
        )

        # PDU types are generated from the TYPE column in Excel.
        pdu_type_display = {
            "BASIC": "Basic PDU",
            "METERED": "Metered PDU",
            "SWITCHED": "Switched PDU",
            "MANAGED": "Managed PDU",
        }

        excel_pdu_types = []
        if not pdus_df.empty:
            excel_pdu_types = [
                x for x in pdus_df["Type"].astype(str).str.strip().str.upper().unique().tolist()
                if x
            ]

        pdu_types = ["None"] + [
            pdu_type_display.get(x, f"{x.title()} PDU")
            for x in excel_pdu_types
        ]

        pdu_col1, pdu_col2 = st.columns([1.0, 1.65], gap="small")

        with pdu_col1:
            previous_pdu_type = st.session_state.get("pdu_type_selection", "None")
            if previous_pdu_type not in pdu_types:
                previous_pdu_type = "None"

            selected_pdu_type = st.selectbox(
                "PDU Type",
                pdu_types,
                index=pdu_types.index(previous_pdu_type),
                key="pdu_type_selection",
            )

        with pdu_col2:
            if selected_pdu_type != "None":
                reverse_type = {v: k for k, v in pdu_type_display.items()}
                excel_pdu_type = reverse_type.get(
                    selected_pdu_type,
                    selected_pdu_type.replace(" PDU", "").upper(),
                )

                filtered_pdus = pdus_df[
                    pdus_df["Type"].astype(str).str.strip().str.upper() == excel_pdu_type
                ].copy()

                if not filtered_pdus.empty:
                    pdu_options = [
                        f'{clean_text(r["Part Code"])} — {clean_text(r["Description"])}'
                        for _, r in filtered_pdus.iterrows()
                    ]

                    old_selection = st.session_state.get("pdu_model_selection")
                    pdu_index = pdu_options.index(old_selection) if old_selection in pdu_options else 0

                    selected_pdu = st.selectbox(
                        "Select PDU",
                        pdu_options,
                        index=pdu_index,
                        key="pdu_model_selection",
                    )

                    selected_row = filtered_pdus.iloc[pdu_options.index(selected_pdu)]
                    selected_part = clean_text(selected_row["Part Code"])

                    st.session_state.pdu_qty = {selected_part: 1}

                    st.html(f"""
                    <div class="mdc-detail-box">
                        <div class="mdc-detail-title">SELECTED PDU</div>
                        <div class="mdc-detail-text">
                            <b>{clean_text(selected_row["Part Code"])}</b><br>
                            {clean_text(selected_row["Description"])}<br>
                            Type: {clean_text(selected_row["Type"])}
                            &nbsp;•&nbsp; C13: {numeric(selected_row["C13"]):g}
                            &nbsp;•&nbsp; C19: {numeric(selected_row["C19"]):g}
                        </div>
                    </div>
                    """)
                else:
                    st.session_state.pdu_qty = {}
                    st.warning(f"No {selected_pdu_type} options found in MDC_Master_V1.xlsx.")
            else:
                st.session_state.pdu_qty = {}
                st.html("""
                <div class="mdc-detail-box">
                    <div class="mdc-detail-title">PDU STATUS</div>
                    <div class="mdc-detail-text">No PDU selected.</div>
                </div>
                """)


# ============================================================
# 4. OTHER ACCESSORIES
# ============================================================
with right_bottom:
    with st.container(border=True):
        st.markdown(
            '<div class="mdc-card-heading">'
            '<span class="mdc-number">04</span>'
            '<span>OTHER ACCESSORIES</span>'
            '</div>',
            unsafe_allow_html=True,
        )

        # Every accessory below comes from the Excel optional-items table.
        optional_lookup = {
            clean_text(r["Part Code"]): r
            for _, r in accessories_df.iterrows()
            if clean_text(r["Part Code"])
        }

        def excel_optional_rows(keyword=None):
            """Find optional items by their Excel part number/description."""
            if accessories_df.empty:
                return pd.DataFrame()

            if not keyword:
                return accessories_df.copy()

            key = str(keyword).upper()
            mask = (
                accessories_df["Part Code"].astype(str).str.upper().str.contains(key, na=False)
                | accessories_df["Description"].astype(str).str.upper().str.contains(key, na=False)
            )
            return accessories_df[mask].copy()

        def remove_rows(rows):
            for _, r in rows.iterrows():
                part = clean_text(r["Part Code"])
                if part:
                    st.session_state.accessory_qty.pop(part, None)

        def add_rows(rows, quantity=1):
            for _, r in rows.iterrows():
                part = clean_text(r["Part Code"])
                if part:
                    st.session_state.accessory_qty[part] = quantity

        # ---------------- FIRE SUPPRESSION ----------------
        st.markdown('<div class="mdc-mini-heading">Fire Suppression</div>', unsafe_allow_html=True)

        fire_rows = excel_optional_rows("FIRE")
        external_fire = fire_rows[
            fire_rows["Description"].astype(str).str.upper().str.contains("EXTERNAL", na=False)
        ].copy()
        internal_fire = fire_rows[
            fire_rows["Description"].astype(str).str.upper().str.contains("INTERNAL|IN-RACK", na=False)
        ].copy()

        fire_current = "None"
        if not external_fire.empty and any(
            numeric(st.session_state.accessory_qty.get(p, 0)) > 0
            for p in external_fire["Part Code"].astype(str).str.strip()
        ):
            fire_current = "External"
        elif not internal_fire.empty and any(
            numeric(st.session_state.accessory_qty.get(p, 0)) > 0
            for p in internal_fire["Part Code"].astype(str).str.strip()
        ):
            fire_current = "Internal"

        fire_selection = st.radio(
            "Fire Suppression",
            ["None", "External", "Internal"],
            index=["None", "External", "Internal"].index(fire_current),
            horizontal=True,
            key="fire_suppression_selection",
        )

        remove_rows(external_fire)
        remove_rows(internal_fire)

        if fire_selection == "External":
            add_rows(external_fire, 1)
        elif fire_selection == "Internal":
            add_rows(internal_fire, 1)

        # ---------------- CAMERA ----------------
        st.markdown('<div class="mdc-mini-heading">Camera</div>', unsafe_allow_html=True)

        camera_rows = excel_optional_rows("CAMERA")
        camera_parts = (
            camera_rows["Part Code"].astype(str).str.strip().tolist()
            if not camera_rows.empty else []
        )

        camera_current = "Yes" if any(
            numeric(st.session_state.accessory_qty.get(p, 0)) > 0
            for p in camera_parts
        ) else "No"

        camera_selection = st.radio(
            "Camera",
            ["Yes", "No"],
            index=["Yes", "No"].index(camera_current),
            horizontal=True,
            key="camera_system_selection",
        )

        if camera_selection == "Yes":
            add_rows(camera_rows, 1)
        else:
            remove_rows(camera_rows)

        # ---------------- OTHER OPTIONAL ACCESSORIES ----------------
        st.markdown('<div class="mdc-mini-heading">Optional Accessories</div>', unsafe_allow_html=True)

        other_accessory_keywords = [
            ("KEYBOARD", "Rotating Keyboard Tray"),
            ("CABLE MANAGER", "Cable Manager"),
            ("TOP CABLE TRAY", "Top Cable Tray"),
            ("BRUSH PANEL", "Brush Panel"),
        ]

        for keyword, fallback_label in other_accessory_keywords:
            rows = excel_optional_rows(keyword)
            if rows.empty:
                continue

            for _, r in rows.iterrows():
                part = clean_text(r["Part Code"])
                description = clean_text(r["Description"])
                if not part:
                    continue

                current_qty = int(numeric(st.session_state.accessory_qty.get(part, 0)))

                acc_col1, acc_col2 = st.columns([4.5, 1.15], gap="small", vertical_alignment="center")

                with acc_col1:
                    selected = st.checkbox(
                        description if description else fallback_label,
                        value=current_qty > 0,
                        key=f"other_acc_{part}",
                    )

                with acc_col2:
                    if selected:
                        qty = st.number_input(
                            "Qty",
                            min_value=1,
                            max_value=999,
                            step=1,
                            value=current_qty if current_qty > 0 else 1,
                            key=f"other_qty_{part}",
                            label_visibility="collapsed",
                        )
                        st.session_state.accessory_qty[part] = qty
                    else:
                        st.session_state.accessory_qty.pop(part, None)


# ============================================================
# 5. FINAL BOQ

# ============================================================
# 5. FINAL BOQ
# ============================================================

bom = build_bom()

base_cost, optional_cost, pdu_cost, total_cost = cost_summary(bom)

margin_pct = float(st.session_state.margin_pct)
freight = float(st.session_state.freight)
installation = float(st.session_state.installation)

if not bom.empty:
    bom_with_price, margin_price, final_selling_price = add_selling_prices(
        bom,
        total_cost,
        margin_pct,
        freight,
        installation,
    )
else:
    bom_with_price = bom.copy()
    margin_price = (
        total_cost / (1 - margin_pct / 100)
        if margin_pct < 100
        else 0.0
    )
    final_selling_price = margin_price + freight + installation


st.html(f"""
<div style="
    display:flex;justify-content:space-between;align-items:center;gap:15px;
    background:#003B71;color:white;padding:7px 12px;border-radius:5px;
    margin:10px 0 8px 0;
">
    <div style="font-size:14px;font-weight:700;">5. FINAL BOQ</div>
    <div style="display:flex;align-items:center;gap:8px;white-space:nowrap;">
        <span style="font-size:11px;font-weight:700;">FINAL SELLING PRICE</span>
        <span style="font-size:16px;font-weight:700;">{money(final_selling_price)}</span>
    </div>
</div>
""")


if not bom.empty:
    structure = bom_with_price[
        [
            "S.No.",
            "Part Code",
            "Description",
            "Quantity",
            "UOM",
            "Unit Price",
            "Total Price",
        ]
    ].copy()

    # --------------------------------------------------------
    # Serial numbering
    # --------------------------------------------------------

    selected_config_components = selected_components()

    cooling_part_codes = set()

    if not selected_config_components.empty:
        # The last three priced lines are the cooling system lines
        # in the supplied one-sheet workbook.
        cooling_rows = selected_config_components.tail(3)

        cooling_part_codes = set(
            cooling_rows["Part Code"]
            .dropna()
            .astype(str)
            .str.strip()
        )

    new_serial = []
    main_mdc_found = False
    mdc_sub_no = 0
    cooling_started = False
    cooling_sub_no = 0
    accessory_no = 3

    for idx, row in structure.iterrows():
        part_code = clean_text(row["Part Code"])
        description = clean_text(row["Description"])
        component_type = clean_text(
            bom.loc[row.name, "Component Type"]
        )

        if (
            not main_mdc_found
            and "SINGLE RACK MDC" in description.upper()
        ):
            new_serial.append("")
            main_mdc_found = True
            continue

        if part_code == "801029209":
            new_serial.append("1")
            continue

        if part_code in cooling_part_codes:
            cooling_started = True
            cooling_sub_no += 1
            new_serial.append(f"2.{cooling_sub_no}")
            continue

        if component_type in ("Optional Accessory", "PDU"):
            new_serial.append(str(accessory_no))
            accessory_no += 1
            continue

        if not cooling_started:
            mdc_sub_no += 1
            new_serial.append(f"1.{mdc_sub_no}")
        else:
            new_serial.append(str(accessory_no))
            accessory_no += 1

    structure["New S.No."] = new_serial

    # --------------------------------------------------------
    # HTML BOQ table
    # --------------------------------------------------------

    html = """
    <style>
    .final-structure-wrap {
        width:100%;
        overflow-x:auto;
    }

    .final-structure-table {
        width:100%;
        border-collapse:collapse;
        table-layout:fixed;
        font-family:Arial,sans-serif;
        font-size:12px;
        border:1px solid #D9E1E8;
    }

    .final-structure-table th {
        background:#F4F6F8;
        color:#555;
        font-weight:600;
        text-align:left;
        padding:7px 7px;
        border-bottom:1px solid #D9E1E8;
    }

    .final-structure-table td {
        padding:6px 7px;
        border-bottom:1px solid #E5E7EB;
        color:#333;
        vertical-align:middle;
        overflow-wrap:anywhere;
    }

    .main-mdc-row td {
        background:#003B71;
        color:white !important;
        font-weight:700;
        font-size:13px;
        text-align:center !important;
        padding:8px 7px;
    }

    .section-heading td {
        background:#005EB8;
        color:white !important;
        font-weight:700;
        font-size:12px;
        text-align:center !important;
        padding:7px 8px;
    }

    .serial {
        width:7%;
        text-align:center !important;
    }

    .part-code {
        width:13%;
    }

    .description {
        width:40%;
    }

    .quantity {
        width:8%;
        text-align:center !important;
    }

    .uom {
        width:7%;
        text-align:center !important;
    }

    .unit-price {
        width:12%;
        text-align:right !important;
        white-space:nowrap;
    }

    .total-price {
        width:13%;
        text-align:right !important;
        white-space:nowrap;
    }
    </style>

    <div class="final-structure-wrap">
    <table class="final-structure-table">
        <thead>
            <tr>
                <th class="serial">S.No.</th>
                <th class="part-code">Part Code</th>
                <th class="description">Description</th>
                <th class="quantity">Qty</th>
                <th class="uom">UOM</th>
                <th class="unit-price">Unit Price</th>
                <th class="total-price">Total Price</th>
            </tr>
        </thead>
        <tbody>
    """

    cooling_heading_added = False
    accessories_heading_added = False
    pdu_heading_added = False

    for _, row in structure.iterrows():
        part_code = clean_text(row["Part Code"])
        description = clean_text(row["Description"])
        quantity_value = numeric(row["Quantity"])
        quantity = f"{quantity_value:g}" if pd.notna(quantity_value) else ""
        uom = clean_text(row["UOM"])
        serial_no = clean_text(row["New S.No."])
        component_type = clean_text(
            bom.loc[row.name, "Component Type"]
        )

        unit_price = numeric(row["Unit Price"])
        total_price = numeric(row["Total Price"])

        unit_price_display = (
            money(unit_price) if pd.notna(unit_price) else "N/A"
        )

        total_price_display = (
            money(total_price) if pd.notna(total_price) else "N/A"
        )

        if (
            serial_no == ""
            and "SINGLE RACK MDC" in description.upper()
        ):
            html += f"""
            <tr class="main-mdc-row">
                <td colspan="7">{description}</td>
            </tr>
            """
            continue

        if (
            part_code in cooling_part_codes
            and not cooling_heading_added
        ):
            html += """
            <tr class="section-heading">
                <td colspan="7">COOLING UNIT</td>
            </tr>
            """
            cooling_heading_added = True

        if (
            component_type == "Optional Accessory"
            and not accessories_heading_added
        ):
            html += """
            <tr class="section-heading">
                <td colspan="7">OTHER ACCESSORIES</td>
            </tr>
            """
            accessories_heading_added = True

        if (
            component_type == "PDU"
            and not pdu_heading_added
        ):
            html += """
            <tr class="section-heading">
                <td colspan="7">PDU</td>
            </tr>
            """
            pdu_heading_added = True

        html += f"""
        <tr>
            <td class="serial">{serial_no}</td>
            <td class="part-code">{part_code}</td>
            <td class="description">{description}</td>
            <td class="quantity">{quantity}</td>
            <td class="uom">{uom}</td>
            <td class="unit-price">{unit_price_display}</td>
            <td class="total-price">{total_price_display}</td>
        </tr>
        """

    # This row shows the separately calculated selling price.
    html += f"""
        <tr>
            <td colspan="6" style="
                text-align:right;
                font-weight:700;
                padding:8px 7px;
                background:#F7FBFF;
                color:#003B71;
            ">
                FINAL SELLING PRICE
            </td>
            <td class="total-price" style="
                font-weight:700;
                background:#F7FBFF;
                color:#003B71;
            ">
                {money(final_selling_price)}
            </td>
        </tr>
    </tbody>
    </table>
    </div>
    """

    st.html(html)

else:
    st.info("No components selected for the current configuration.")


# ============================================================
# 6/7. INTERNAL COST & SELLING PRICE
# ============================================================

if is_internal:
    section_header("6. COST SUMMARY — INTERNAL ONLY")

    a, b, c, d = st.columns(4)

    with a:
        price_box("Base Cost", base_cost)

    with b:
        price_box("Optional Cost", optional_cost)

    with c:
        price_box("PDU Cost", pdu_cost)

    with d:
        price_box("Total Cost", total_cost)

    section_header("7. COST TO SELLING PRICE — INTERNAL ONLY")

    p1, p2, p3, p4 = st.columns(4)

    with p1:
        st.session_state.margin_pct = st.number_input(
            "Margin (%)",
            min_value=0.0,
            max_value=99.0,
            value=float(st.session_state.margin_pct),
            step=0.5,
        )

    with p2:
        st.session_state.freight = st.number_input(
            "Freight",
            min_value=0.0,
            value=float(st.session_state.freight),
            step=500.0,
        )

    with p3:
        st.session_state.installation = st.number_input(
            "Installation",
            min_value=0.0,
            value=float(st.session_state.installation),
            step=500.0,
        )

    with p4:
        st.session_state.warranty_pct = st.number_input(
            "Warranty (%)",
            min_value=0.0,
            max_value=100.0,
            value=float(st.session_state.warranty_pct),
            step=0.5,
        )

    margin_pct = float(st.session_state.margin_pct)
    freight = float(st.session_state.freight)
    installation = float(st.session_state.installation)
    warranty_pct = float(st.session_state.warranty_pct)

    margin_price = (
        total_cost / (1 - margin_pct / 100)
        if margin_pct < 100
        else 0.0
    )

    final_selling_price = margin_price + freight + installation
    warranty_amount = margin_price * warranty_pct / 100

    a, b, c, d = st.columns(4)

    with a:
        price_box("Margin Price", margin_price)

    with b:
        price_box("After Freight", margin_price + freight)

    with c:
        price_box("Final Selling Price", final_selling_price)

    with d:
        price_box("Warranty Amount", warranty_amount)


# ============================================================
# 8. DOWNLOADS — EXCEL + PDF
# ============================================================

section_header("8. DOWNLOADS")

if not bom.empty:
    internal_cost_data = [
        ["Base Cost", base_cost],
        ["Optional Cost", optional_cost],
        ["PDU Cost", pdu_cost],
        ["Total Cost", total_cost],
        ["Margin %", margin_pct],
        ["Margin Price", margin_price],
        ["Freight", freight],
        ["Installation", installation],
        ["Final Selling Price", final_selling_price],
    ]

    sales_excel = excel_bytes(
        internal=False,
        bom=bom_with_price,
        final_price=final_selling_price,
    )
    sales_pdf = pdf_bytes(
        internal=False,
        bom=bom_with_price,
        final_price=final_selling_price,
    )

    if is_internal:
        internal_excel = excel_bytes(
            internal=True,
            bom=bom_with_price,
            final_price=final_selling_price,
            cost_data=internal_cost_data,
        )
        internal_pdf = pdf_bytes(
            internal=True,
            bom=bom_with_price,
            final_price=final_selling_price,
        )

        st.markdown("**Internal – MDC**")
        i1, i2 = st.columns(2)

        with i1:
            st.download_button(
                "⬇️ Download Internal Excel",
                data=internal_excel,
                file_name="MDC_Internal_Cost.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                on_click=handle_excel_download,
            )

        with i2:
            st.download_button(
                "📄 Download Internal PDF",
                data=internal_pdf,
                file_name="MDC_Internal_Cost.pdf",
                mime="application/pdf",
                use_container_width=True,
                on_click=handle_excel_download,
            )

        st.markdown("**Sales**")

    s1, s2 = st.columns(2)

    with s1:
        st.download_button(
            "⬇️ Download Sales Excel",
            data=sales_excel,
            file_name="MDC_Sales_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            on_click=handle_excel_download,
        )

    with s2:
        st.download_button(
            "📄 Download Sales PDF",
            data=sales_pdf,
            file_name="MDC_Sales_Output.pdf",
            mime="application/pdf",
            use_container_width=True,
            on_click=handle_excel_download,
        )
else:
    st.info("Select a configuration with available BOM data before downloading.")

# ============================================================

# 9. SAVE CONFIGURATION
# ============================================================

section_header("9. SAVE CONFIGURATION")

save_col1, save_col2 = st.columns([2, 5])

with save_col1:
    if st.button(
        "💾 Save Configuration",
        use_container_width=True,
        type="primary",
    ):
        current_margin_price = (
            total_cost / (1 - margin_pct / 100)
            if margin_pct < 100
            else 0.0
        )

        current_final_price = (
            current_margin_price + freight + installation
        )

        current_warranty_amount = (
            current_margin_price * warranty_pct / 100
        )

        save_configuration(
            configuration_id=st.session_state.configuration_id,
            bom=bom_with_price,
            base_cost=base_cost,
            optional_cost=optional_cost,
            pdu_cost=pdu_cost,
            total_cost=total_cost,
            margin_pct=margin_pct,
            freight=freight,
            installation=installation,
            warranty_pct=warranty_pct,
            margin_price=current_margin_price,
            final_selling_price=current_final_price,
            warranty_amount=current_warranty_amount,
        )

        st.session_state.configuration_saved = True

        st.success(
            "Configuration saved successfully."
        )


# ============================================================
# 10. CONFIGURATION HISTORY
# INTERNAL USERS ONLY
# ============================================================

if is_internal:
    section_header("10. CONFIGURATION HISTORY")

    conn = sqlite3.connect(TRACKING_DB)

    history_df = pd.read_sql_query(
        """
        SELECT
            COALESCE(NULLIF(user_code, ''), '—') AS "User Code",
            customer_name AS "Customer Name",
            DATE(created_at) AS "Date"
        FROM configurations
        ORDER BY id DESC
        """,
        conn,
    )

    conn.close()

    if not history_df.empty:
        st.dataframe(
            history_df,
            use_container_width=True,
            hide_index=True,
            height=220,
        )
    else:
        st.info("No saved configurations available yet.")


# ============================================================
# FOOTER
# ============================================================

st.divider()
st.caption("Eaton MDC Solution Configurator")
