
import os
import sqlite3
from io import BytesIO
from datetime import datetime

import pandas as pd
import streamlit as st

# PDF
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer,
)


# ============================================================
# MDC SOLUTION CONFIGURATOR
# ============================================================

st.set_page_config(
    page_title="MDC Solution",
    page_icon="🏢",
    layout="wide",
)


# ============================================================
# PROFESSIONAL THEME
# ============================================================

PRIMARY_BLUE = "#007AC2"
DARK_BLUE = "#005080"
LIGHT_BLUE = "#DCEFF8"
PALE_BLUE = "#F4FAFD"
BORDER_BLUE = "#B8DDF0"
TEXT_DARK = "#263238"
TEXT_MUTED = "#64748B"
WHITE = "#FFFFFF"


st.markdown(
    f"""
    <style>

    /* --------------------------------------------------------
       GLOBAL
       -------------------------------------------------------- */

    .block-container {{
        padding-top: 1rem;
        padding-bottom: 2rem;
        max-width: 1500px;
    }}

    html, body, [class*="css"] {{
        font-family: Arial, sans-serif;
    }}

    /* Compact headings */

    h1 {{
        font-size: 26px !important;
    }}

    h2 {{
        font-size: 21px !important;
    }}

    h3 {{
        font-size: 17px !important;
    }}

    /* --------------------------------------------------------
       STREAMLIT HEADERS
       -------------------------------------------------------- */

    div[data-testid="stHeadingWithActionElements"] {{
        margin-top: 0.35rem;
        margin-bottom: 0.45rem;
    }}

    /* --------------------------------------------------------
       INPUTS
       -------------------------------------------------------- */

    div[data-baseweb="select"] > div {{
        min-height: 38px;
    }}

    div[data-testid="stTextInput"] input,
    div[data-testid="stNumberInput"] input {{
        min-height: 36px;
    }}

    /* --------------------------------------------------------
       BUTTON
       -------------------------------------------------------- */

    .stButton > button {{
        border-radius: 6px;
        border: 1px solid {PRIMARY_BLUE};
        font-weight: 600;
    }}

    /* --------------------------------------------------------
       SECTION RIBBON
       -------------------------------------------------------- */

    .section-ribbon {{
        background: linear-gradient(
            90deg,
            {PRIMARY_BLUE},
            #1593D0
        );
        color: white;
        padding: 7px 13px;
        border-radius: 6px;
        font-size: 15px;
        font-weight: 700;
        margin-top: 12px;
        margin-bottom: 9px;
    }}

    /* --------------------------------------------------------
       SELECTION CARD
       -------------------------------------------------------- */

    .selection-card {{
        background: {PALE_BLUE};
        border: 1px solid {BORDER_BLUE};
        border-left: 5px solid {PRIMARY_BLUE};
        border-radius: 7px;
        padding: 9px 13px;
        margin: 7px 0 10px 0;
    }}

    .selection-label {{
        font-size: 12px;
        color: {TEXT_MUTED};
        margin-bottom: 3px;
    }}

    .selection-value {{
        font-size: 15px;
        font-weight: 700;
        color: {DARK_BLUE};
    }}

    /* --------------------------------------------------------
       USER CODE
       -------------------------------------------------------- */

    .user-code-card {{
        background: #F7FBFE;
        border: 1px solid {BORDER_BLUE};
        border-left: 5px solid {DARK_BLUE};
        border-radius: 7px;
        padding: 9px 14px;
        margin-bottom: 10px;
    }}

    .user-code-label {{
        font-size: 11px;
        color: {TEXT_MUTED};
        text-transform: uppercase;
        letter-spacing: 0.4px;
    }}

    .user-code-value {{
        font-size: 17px;
        font-weight: 700;
        color: {DARK_BLUE};
        margin-top: 2px;
    }}

    /* --------------------------------------------------------
       PRICE
       -------------------------------------------------------- */

    .price-card {{
        background: white;
        border: 1px solid {BORDER_BLUE};
        border-radius: 7px;
        padding: 10px 13px;
        min-height: 65px;
    }}

    .price-label {{
        font-size: 12px;
        color: {TEXT_MUTED};
    }}

    .price-value {{
        font-size: 19px;
        font-weight: 700;
        color: {DARK_BLUE};
        margin-top: 4px;
    }}

    /* --------------------------------------------------------
       DISABLED DOWNLOAD LOOK
       -------------------------------------------------------- */

    .download-note {{
        background: #FFF8E8;
        border: 1px solid #F2D58A;
        border-radius: 6px;
        padding: 8px 12px;
        color: #765B00;
        font-size: 13px;
        margin-bottom: 8px;
    }}

    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# FILE PATHS
# ============================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

MASTER_FILE = os.path.join(
    BASE_DIR,
    "MDC_Master_V1.xlsx"
)

TRACKING_DB = os.path.join(
    BASE_DIR,
    "MDC_Tracking.db"
)

DEMO_INTERNAL_PASSWORD = "MDC@123"


# ============================================================
# DATABASE
# ============================================================

def init_tracking_db():

    conn = sqlite3.connect(TRACKING_DB)

    cursor = conn.cursor()

    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS configurations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,

            configuration_id TEXT UNIQUE,
            user_code TEXT,

            created_at TEXT,

            customer_name TEXT,
            customer_place TEXT,
            problem TEXT,
            solution TEXT,

            mdc_type TEXT,
            configuration TEXT,

            pdu_type TEXT,
            pdu_part_code TEXT,

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
            warranty_amount REAL
        )
        """
    )

    cursor.execute(
        """
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
        """
    )

    conn.commit()
    conn.close()


init_tracking_db()


# ============================================================
# USER COUNT / USER CODE
# ============================================================

def get_next_user_number():

    conn = sqlite3.connect(TRACKING_DB)

    cursor = conn.cursor()

    cursor.execute(
        """
        SELECT COUNT(*)
        FROM configurations
        """
    )

    count = cursor.fetchone()[0]

    conn.close()

    return count + 1


def generate_user_code():

    now = datetime.now()

    date_part = now.strftime("%Y%m%d")

    user_number = get_next_user_number()

    mdc_type_code = (
        "SR"
        if st.session_state.mdc_type == "Single Rack"
        else "MR"
    )

    config_number = (
        st.session_state.configuration
        .replace("Configuration ", "C")
        .replace(" ", "")
    )

    pdu_code = {
        "Basic": "PB",
        "Metered": "PM",
        "Switched": "PS",
        "None": "PN",
    }.get(
        st.session_state.pdu_type,
        "PN"
    )

    fire_code = (
        "FE"
        if st.session_state.fire_suppression == "External"
        else "FI"
        if st.session_state.fire_suppression == "Internal"
        else "FN"
    )

    camera_code = (
        "CY"
        if st.session_state.camera_enabled == "Yes"
        else "CN"
    )

    return (
        f"MDC-{date_part}-"
        f"U{user_number:04d}-"
        f"{mdc_type_code}-"
        f"{config_number}-"
        f"{pdu_code}-"
        f"{fire_code}-"
        f"{camera_code}"
    )


# ============================================================
# MASTER DATA
# ============================================================

@st.cache_data
def load_master():

    configs = pd.read_excel(
        MASTER_FILE,
        sheet_name="Configurations"
    )

    components = pd.read_excel(
        MASTER_FILE,
        sheet_name="Components"
    )

    accessories = pd.read_excel(
        MASTER_FILE,
        sheet_name="Accessories"
    )

    pdus = pd.read_excel(
        MASTER_FILE,
        sheet_name="PDUs"
    )

    return (
        configs,
        components,
        accessories,
        pdus
    )


configs_df, components_df, accessories_df, pdus_df = load_master()


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

    "pdu_type": "None",
    "selected_pdu_part": "",
    "pdu_qty": {},

    "fire_suppression": "None",
    "camera_enabled": "No",

    "accessory_qty": {},

    "rotating_keyboard": False,
    "cable_manager": False,
    "top_cable_tray": False,
    "brush_panel": False,

    "margin_pct": 20.0,
    "freight": 0.0,
    "installation": 0.0,
    "warranty_pct": 0.0,

    "configuration_saved": False,
}


for key, value in defaults.items():

    if key not in st.session_state:

        st.session_state[key] = value


# ============================================================
# COMMON FUNCTIONS
# ============================================================

def money(value):

    try:
        return f"₹ {float(value):,.2f}"
    except Exception:
        return "₹ 0.00"


def price_box(label, value):

    st.markdown(
        f"""
        <div class="price-card">
            <div class="price-label">
                {label}
            </div>
            <div class="price-value">
                {money(value)}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def internal_password():

    try:
        return st.secrets["MDC_INTERNAL_PASSWORD"]

    except Exception:

        return DEMO_INTERNAL_PASSWORD


def selected_config_record():

    match = configs_df[
        (configs_df["MDC Type"] == st.session_state.mdc_type)
        &
        (
            configs_df["Configuration"]
            == st.session_state.configuration
        )
    ]

    return (
        match.iloc[0]
        if not match.empty
        else None
    )


def selected_components():

    return components_df[
        (components_df["MDC Type"] == st.session_state.mdc_type)
        &
        (
            components_df["Configuration"]
            == st.session_state.configuration
        )
    ].copy()


def get_accessory_row(part_code):

    match = accessories_df[
        accessories_df["Part Code"]
        .astype(str)
        .str.strip()
        == str(part_code).strip()
    ]

    return (
        match.iloc[0]
        if not match.empty
        else None
    )


# ============================================================
# CONFIGURATION TITLES
# ============================================================

CONFIGURATION_DISPLAY_NAMES = {

    "Configuration 1":
        "1SR ,42U×800W×1200D, 3.5KW ,W/O Dehumidifier",

    "Configuration 2":
        "1SR ,42U×800W×1200D, 7KW, Dehumidifier",

    "Configuration 3":
        "1SR ,42U×800W×1200D, 7KW ,W/O Dehumidifier",

    "Configuration 4":
        "1SR ,42U×800W×1200D, 7KW, Dehumidifier",
}


def configuration_display_name(config):

    return CONFIGURATION_DISPLAY_NAMES.get(
        config,
        config
    )


# ============================================================
# PDU HELPERS
# ============================================================

def get_pdu_rows():

    if st.session_state.pdu_type == "None":

        return pdus_df.iloc[0:0].copy()

    if "Type" not in pdus_df.columns:

        return pdus_df.copy()

    return pdus_df[
        pdus_df["Type"]
        .astype(str)
        .str.strip()
        .str.lower()
        ==
        st.session_state.pdu_type.lower()
    ].copy()


def reset_pdu():

    st.session_state.selected_pdu_part = ""
    st.session_state.pdu_qty = {}


# ============================================================
# ACCESSORY PART DEFINITIONS
# ============================================================

FIRE_SUPPRESSION_PARTS = {

    "External": "801073203",

    "Internal": "HRD-XH1C",
}


CAMERA_PARTS = [

    "801303201",
    "801303202",
    "801303204",
    "801303206",
    "801303208",
    "801303203",
]


OTHER_OPTIONAL_PARTS = {

    "Rotating Keyboard":
        "801223664",

    "Cable Manager":
        "801075237",

    "Top Cable Tray":
        "801029022",

    "Brush Panel":
        "801075235",
}


def add_accessory(part_code, qty=1):

    row = get_accessory_row(part_code)

    if row is not None:

        st.session_state.accessory_qty[
            str(part_code)
        ] = qty


def remove_accessory(part_code):

    st.session_state.accessory_qty.pop(
        str(part_code),
        None
    )


# ============================================================
# BUILD BOQ
# ============================================================

def build_bom():

    rows = []

    # --------------------------------------------------------
    # BASE CONFIGURATION
    # --------------------------------------------------------

    for _, r in selected_components().iterrows():

        cost = r["Unit Cost"]

        try:
            qty = float(r["Quantity"])
        except Exception:
            qty = 1

        rows.append(
            {
                "S.No.": len(rows) + 1,

                "Component Type":
                    "Base (Configuration)",

                "Part Code":
                    (
                        str(r["Part Code"])
                        if pd.notna(r["Part Code"])
                        else ""
                    ),

                "Description":
                    str(r["Description"]),

                "Quantity":
                    qty,

                "UOM":
                    str(r["UOM"]),

                "Unit Cost":
                    cost,

                "Total Cost":
                    (
                        float(cost) * qty
                        if pd.notna(cost)
                        else 0
                    ),

                "Source":
                    "Configuration",
            }
        )


    # --------------------------------------------------------
    # FIRE SUPPRESSION
    # --------------------------------------------------------

    fire_selection = (
        st.session_state.fire_suppression
    )

    if fire_selection in FIRE_SUPPRESSION_PARTS:

        part = FIRE_SUPPRESSION_PARTS[
            fire_selection
        ]

        row = get_accessory_row(part)

        if row is not None:

            cost = row["Unit Cost"]

            rows.append(
                {
                    "S.No.": len(rows) + 1,
                    "Component Type":
                        "Optional Accessory",
                    "Part Code":
                        part,
                    "Description":
                        str(row["Description"]),
                    "Quantity": 1,
                    "UOM":
                        str(row["UOM"]),
                    "Unit Cost":
                        cost,
                    "Total Cost":
                        (
                            float(cost)
                            if pd.notna(cost)
                            else 0
                        ),
                    "Source":
                        "Optional Accessory",
                }
            )


    # --------------------------------------------------------
    # CAMERA SYSTEM
    # --------------------------------------------------------

    if st.session_state.camera_enabled == "Yes":

        for part in CAMERA_PARTS:

            row = get_accessory_row(part)

            if row is None:
                continue

            cost = row["Unit Cost"]

            rows.append(
                {
                    "S.No.": len(rows) + 1,
                    "Component Type":
                        "Optional Accessory",
                    "Part Code":
                        part,
                    "Description":
                        str(row["Description"]),
                    "Quantity": 1,
                    "UOM":
                        str(row["UOM"]),
                    "Unit Cost":
                        cost,
                    "Total Cost":
                        (
                            float(cost)
                            if pd.notna(cost)
                            else 0
                        ),
                    "Source":
                        "Optional Accessory",
                }
            )


    # --------------------------------------------------------
    # OTHER ACCESSORIES
    # --------------------------------------------------------

    for name, part in OTHER_OPTIONAL_PARTS.items():

        if not st.session_state.accessory_qty.get(
            part,
            0
        ):
            continue

        row = get_accessory_row(part)

        if row is None:
            continue

        qty = float(
            st.session_state.accessory_qty.get(
                part,
                1
            )
        )

        cost = row["Unit Cost"]

        rows.append(
            {
                "S.No.": len(rows) + 1,
                "Component Type":
                    "Optional Accessory",
                "Part Code":
                    part,
                "Description":
                    str(row["Description"]),
                "Quantity":
                    qty,
                "UOM":
                    str(row["UOM"]),
                "Unit Cost":
                    cost,
                "Total Cost":
                    (
                        float(cost) * qty
                        if pd.notna(cost)
                        else 0
                    ),
                "Source":
                    "Optional Accessory",
            }
        )


    # --------------------------------------------------------
    # PDU
    # --------------------------------------------------------

    if st.session_state.pdu_type != "None":

        part = (
            st.session_state.selected_pdu_part
        )

        row_match = pdus_df[
            pdus_df["Part Code"]
            .astype(str)
            .str.strip()
            == str(part).strip()
        ]

        if not row_match.empty:

            r = row_match.iloc[0]

            qty = float(
                st.session_state.pdu_qty.get(
                    part,
                    1
                )
            )

            cost = r["Unit Cost"]

            description = str(
                r["Description"]
            )

            extra = []

            for col in ["Type", "C13", "C19"]:

                if col in r.index:

                    value = r[col]

                    if pd.notna(value):

                        extra.append(
                            f"{col}: {value}"
                        )

            if extra:

                description += (
                    " | "
                    + " | ".join(extra)
                )

            rows.append(
                {
                    "S.No.": len(rows) + 1,
                    "Component Type": "PDU",
                    "Part Code": str(part),
                    "Description": description,
                    "Quantity": qty,
                    "UOM": str(r["UOM"]),
                    "Unit Cost": cost,
                    "Total Cost":
                        (
                            float(cost) * qty
                            if pd.notna(cost)
                            else 0
                        ),
                    "Source": "PDU",
                }
            )

    return pd.DataFrame(rows)


# ============================================================
# COST SUMMARY
# ============================================================

def cost_summary(bom):

    cfg = selected_config_record()

    base_cost = 0.0

    if cfg is not None:

        if pd.notna(cfg["Base Cost"]):

            base_cost = float(
                cfg["Base Cost"]
            )

    optional_cost = 0.0
    pdu_cost = 0.0

    if not bom.empty:

        optional_cost = float(
            bom.loc[
                bom["Source"]
                == "Optional Accessory",
                "Total Cost"
            ]
            .fillna(0)
            .sum()
        )

        pdu_cost = float(
            bom.loc[
                bom["Source"] == "PDU",
                "Total Cost"
            ]
            .fillna(0)
            .sum()
        )

    total_cost = (
        base_cost
        + optional_cost
        + pdu_cost
    )

    return (
        base_cost,
        optional_cost,
        pdu_cost,
        total_cost
    )


# ============================================================
# SELLING PRICE
# ============================================================

def add_selling_prices(
    bom,
    total_cost,
    margin_pct,
    freight,
    installation
):

    result = bom.copy()

    margin_price = (
        total_cost
        / (1 - margin_pct / 100)
        if margin_pct < 100
        else 0
    )

    final_selling_price = (
        margin_price
        + freight
        + installation
    )

    known_cost_total = (
        result["Total Cost"]
        .fillna(0)
        .sum()
        if not result.empty
        else 0
    )

    if known_cost_total > 0:

        result["Total Price"] = (
            result["Total Cost"]
            .fillna(0)
            / known_cost_total
            * final_selling_price
        )

        result["Unit Price"] = (
            result["Total Price"]
            / result["Quantity"]
        )

    else:

        result["Total Price"] = pd.NA
        result["Unit Price"] = pd.NA

    return (
        result,
        margin_price,
        final_selling_price
    )


# ============================================================
# CUSTOMER / CONFIGURATION DATA
# ============================================================

def customer_table():

    return pd.DataFrame(
        [
            [
                "User Code",
                st.session_state.user_code
            ],
            [
                "Date",
                datetime.now().strftime(
                    "%d-%m-%Y"
                )
            ],
            [
                "Customer Name",
                st.session_state.customer_name
            ],
            [
                "Customer Place",
                st.session_state.customer_place
            ],
            [
                "Problem Description",
                st.session_state.problem
            ],
            [
                "Solution",
                st.session_state.solution
            ],
            [
                "MDC Type",
                st.session_state.mdc_type
            ],
            [
                "Configuration",
                configuration_display_name(
                    st.session_state.configuration
                )
            ],
            [
                "PDU Type",
                st.session_state.pdu_type
            ],
        ],
        columns=[
            "Field",
            "Value"
        ]
    )


# ============================================================
# FINAL BOQ SERIAL NUMBER
# ============================================================

def prepare_final_boq(bom):

    if bom.empty:

        return bom.copy()

    structure = bom.copy()

    main_mdc_part = "801029209"

    selected_cfg_components = (
        selected_components()
    )

    cooling_part_codes = set()

    if not selected_cfg_components.empty:

        # Retain existing logic:
        # last three configuration components
        cooling_rows = (
            selected_cfg_components.tail(3)
        )

        cooling_part_codes = set(
            cooling_rows["Part Code"]
            .dropna()
            .astype(str)
            .str.strip()
        )

    new_serial = []

    main_found = False

    mdc_sub_no = 0
    cooling_sub_no = 0
    accessory_no = 3

    cooling_started = False

    for idx, row in structure.iterrows():

        part_code = str(
            row["Part Code"]
        ).strip()

        description = str(
            row["Description"]
        ).strip()

        component_type = str(
            row["Component Type"]
        ).strip()

        # ----------------------------------------------------
        # MAIN MDC TITLE
        # ----------------------------------------------------

        if (
            not main_found
            and
            "SINGLE RACK MDC"
            in description.upper()
        ):

            new_serial.append("")

            main_found = True

            continue

        # ----------------------------------------------------
        # MAIN MDC
        # ----------------------------------------------------

        if part_code == main_mdc_part:

            new_serial.append("1")

            continue

        # ----------------------------------------------------
        # COOLING
        # ----------------------------------------------------

        if part_code in cooling_part_codes:

            cooling_started = True

            cooling_sub_no += 1

            new_serial.append(
                f"2.{cooling_sub_no}"
            )

            continue

        # ----------------------------------------------------
        # OPTIONAL
        # ----------------------------------------------------

        if component_type == "Optional Accessory":

            new_serial.append(
                str(accessory_no)
            )

            accessory_no += 1

            continue

        # ----------------------------------------------------
        # PDU
        # ----------------------------------------------------

        if component_type == "PDU":

            new_serial.append(
                str(accessory_no)
            )

            accessory_no += 1

            continue

        # ----------------------------------------------------
        # BASE MDC COMPONENTS
        # ----------------------------------------------------

        if not cooling_started:

            mdc_sub_no += 1

            new_serial.append(
                f"1.{mdc_sub_no}"
            )

            continue

        # ----------------------------------------------------
        # FALLBACK
        # ----------------------------------------------------

        new_serial.append(
            str(accessory_no)
        )

        accessory_no += 1

    structure["Final S.No."] = new_serial

    return structure


# ============================================================
# FINAL BOQ HTML
# ============================================================

def final_boq_html(bom_with_price):

    if bom_with_price.empty:

        return ""

    structure = prepare_final_boq(
        bom_with_price
    )

    html = f"""
    <style>

    .boq-table {{
        width:100%;
        border-collapse:collapse;
        font-family:Arial,sans-serif;
        font-size:12px;
        border:1px solid #D5E2EA;
    }}

    .boq-table th {{
        background:#F1F5F8;
        color:#344054;
        font-weight:700;
        padding:8px 7px;
        border:1px solid #D5E2EA;
        text-align:left;
    }}

    .boq-table td {{
        padding:7px;
        border:1px solid #E2E8F0;
        vertical-align:middle;
        color:#263238;
    }}

    .boq-title td {{
        background:{DARK_BLUE};
        color:white !important;
        font-weight:700;
        font-size:14px;
        text-align:center;
        padding:10px;
    }}

    .boq-section td {{
        background:{PRIMARY_BLUE};
        color:white !important;
        font-weight:700;
        text-align:center;
        padding:7px;
    }}

    .sno {{
        width:6%;
        text-align:center !important;
    }}

    .part {{
        width:15%;
    }}

    .desc {{
        width:46%;
    }}

    .qty {{
        width:7%;
        text-align:center !important;
    }}

    .uom {{
        width:7%;
        text-align:center !important;
    }}

    .price {{
        width:10%;
        text-align:right !important;
    }}

    </style>

    <table class="boq-table">

    <thead>
        <tr>
            <th class="sno">S.No.</th>
            <th class="part">Part Code</th>
            <th class="desc">Description</th>
            <th class="qty">Qty</th>
            <th class="uom">UOM</th>
            <th class="price">Unit Price</th>
            <th class="price">Total Price</th>
        </tr>
    </thead>

    <tbody>
    """

    cooling_added = False
    accessories_added = False
    pdu_added = False

    for _, row in structure.iterrows():

        part_code = str(
            row["Part Code"]
        ).strip()

        description = str(
            row["Description"]
        ).strip()

        serial_no = str(
            row["Final S.No."]
        ).strip()

        component_type = str(
            row["Component Type"]
        ).strip()

        quantity = row["Quantity"]

        uom = str(
            row["UOM"]
        ).strip()

        unit_price = row.get(
            "Unit Price",
            None
        )

        total_price = row.get(
            "Total Price",
            None
        )

        # ----------------------------------------------------
        # MAIN MDC TITLE
        # ----------------------------------------------------

        if (
            serial_no == ""
            and
            "SINGLE RACK MDC"
            in description.upper()
        ):

            html += f"""
            <tr class="boq-title">
                <td colspan="7">
                    {description}
                </td>
            </tr>
            """

            continue

        # ----------------------------------------------------
        # COOLING HEADING
        # ----------------------------------------------------

        if (
            part_code
            and
            not cooling_added
            and
            serial_no.startswith("2.")
        ):

            html += """
            <tr class="boq-section">
                <td colspan="7">
                    COOLING UNIT
                </td>
            </tr>
            """

            cooling_added = True

        # ----------------------------------------------------
        # OTHER ACCESSORIES
        # ----------------------------------------------------

        if (
            component_type
            == "Optional Accessory"
            and
            not accessories_added
        ):

            html += """
            <tr class="boq-section">
                <td colspan="7">
                    OTHER ACCESSORIES
                </td>
            </tr>
            """

            accessories_added = True

        # ----------------------------------------------------
        # PDU
        # ----------------------------------------------------

        if (
            component_type == "PDU"
            and
            not pdu_added
        ):

            html += """
            <tr class="boq-section">
                <td colspan="7">
                    PDU
                </td>
            </tr>
            """

            pdu_added = True

        # ----------------------------------------------------
        # PRICES
        # ----------------------------------------------------

        unit_display = (
            money(unit_price)
            if pd.notna(unit_price)
            else "N/A"
        )

        total_display = (
            money(total_price)
            if pd.notna(total_price)
            else "N/A"
        )

        html += f"""
        <tr>

            <td class="sno">
                {serial_no}
            </td>

            <td class="part">
                {part_code}
            </td>

            <td class="desc">
                {description}
            </td>

            <td class="qty">
                {quantity}
            </td>

            <td class="uom">
                {uom}
            </td>

            <td class="price">
                {unit_display}
            </td>

            <td class="price">
                {total_display}
            </td>

        </tr>
        """

    html += """
    </tbody>
    </table>
    """

    return html


# ============================================================
# EXCEL EXPORT
# ============================================================

def excel_bytes(
    internal=False,
    bom=None,
    cost_data=None
):

    output = BytesIO()

    if bom is None:

        bom = build_bom()

    final_boq = prepare_final_boq(
        bom
    )

    cust = customer_table()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        # ----------------------------------------------------
        # CUSTOMER / CONFIGURATION
        # ----------------------------------------------------

        cust.to_excel(
            writer,
            index=False,
            sheet_name="Customer & Configuration"
        )

        # ----------------------------------------------------
        # FINAL BOQ
        # ----------------------------------------------------

        sales_cols = [
            "Final S.No.",
            "Part Code",
            "Description",
            "Quantity",
            "UOM",
            "Unit Price",
            "Total Price",
        ]

        sales_boq = final_boq[
            sales_cols
        ].copy()

        sales_boq.columns = [
            "S.No.",
            "Part Code",
            "Description",
            "Qty",
            "UOM",
            "Unit Price",
            "Total Price",
        ]

        sales_boq.to_excel(
            writer,
            index=False,
            sheet_name="Final BOQ"
        )

        # ----------------------------------------------------
        # INTERNAL COST
        # ----------------------------------------------------

        if internal and cost_data is not None:

            internal_cols = [
                "Final S.No.",
                "Part Code",
                "Description",
                "Quantity",
                "UOM",
                "Unit Cost",
                "Total Cost",
                "Unit Price",
                "Total Price",
            ]

            internal_boq = final_boq[
                internal_cols
            ].copy()

            internal_boq.columns = [
                "S.No.",
                "Part Code",
                "Description",
                "Qty",
                "UOM",
                "Unit Cost",
                "Total Cost",
                "Unit Price",
                "Total Price",
            ]

            internal_boq.to_excel(
                writer,
                index=False,
                sheet_name="Internal Cost BOQ"
            )

            pd.DataFrame(
                cost_data,
                columns=[
                    "Item",
                    "Value"
                ]
            ).to_excel(
                writer,
                index=False,
                sheet_name="Cost Summary"
            )

    output.seek(0)

    return output


# ============================================================
# PDF EXPORT
# ============================================================

def pdf_bytes(
    bom_with_price,
    internal=False,
    cost_data=None
):

    output = BytesIO()

    doc = SimpleDocTemplate(
        output,
        pagesize=landscape(A4),
        rightMargin=10 * mm,
        leftMargin=10 * mm,
        topMargin=10 * mm,
        bottomMargin=10 * mm,
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "MDCTitle",
        parent=styles["Title"],
        fontSize=18,
        leading=22,
        textColor=colors.HexColor(
            DARK_BLUE
        ),
        alignment=TA_CENTER,
        spaceAfter=8,
    )

    normal_style = ParagraphStyle(
        "MDCNormal",
        parent=styles["Normal"],
        fontSize=8,
        leading=10,
    )

    small_style = ParagraphStyle(
        "MDCSmall",
        parent=styles["Normal"],
        fontSize=7,
        leading=8,
    )

    story = []

    story.append(
        Paragraph(
            "Eaton MDC Solution Configurator",
            title_style
        )
    )

    story.append(
        Paragraph(
            "Modular Data Center Solution Configuration & Pricing",
            ParagraphStyle(
                "subtitle",
                parent=normal_style,
                alignment=TA_CENTER,
                fontSize=9,
                textColor=colors.HexColor(
                    TEXT_MUTED
                ),
            )
        )
    )

    story.append(
        Spacer(1, 5 * mm)
    )

    # --------------------------------------------------------
    # CUSTOMER INFORMATION
    # --------------------------------------------------------

    customer_data = [
        [
            "User Code",
            st.session_state.user_code,
            "Date",
            datetime.now().strftime(
                "%d-%m-%Y"
            ),
        ],
        [
            "Customer Name",
            st.session_state.customer_name,
            "Customer Place",
            st.session_state.customer_place,
        ],
        [
            "MDC Type",
            st.session_state.mdc_type,
            "Configuration",
            configuration_display_name(
                st.session_state.configuration
            ),
        ],
        [
            "PDU Type",
            st.session_state.pdu_type,
            "PDU Part",
            st.session_state.selected_pdu_part,
        ],
    ]

    customer_table_pdf = Table(
        customer_data,
        colWidths=[
            28 * mm,
            70 * mm,
            30 * mm,
            105 * mm,
        ],
    )

    customer_table_pdf.setStyle(
        TableStyle(
            [
                (
                    "BACKGROUND",
                    (0, 0),
                    (0, -1),
                    colors.HexColor(
                        LIGHT_BLUE
                    ),
                ),
                (
                    "BACKGROUND",
                    (2, 0),
                    (2, -1),
                    colors.HexColor(
                        LIGHT_BLUE
                    ),
                ),
                (
                    "TEXTCOLOR",
                    (0, 0),
                    (-1, -1),
                    colors.HexColor(
                        TEXT_DARK
                    ),
                ),
                (
                    "FONTNAME",
                    (0, 0),
                    (-1, -1),
                    "Helvetica",
                ),
                (
                    "FONTNAME",
                    (0, 0),
                    (0, -1),
                    "Helvetica-Bold",
                ),
                (
                    "FONTNAME",
                    (2, 0),
                    (2, -1),
                    "Helvetica-Bold",
                ),
                (
                    "FONTSIZE",
                    (0, 0),
                    (-1, -1),
                    7.5,
                ),
                (
                    "GRID",
                    (0, 0),
                    (-1, -1),
                    0.4,
                    colors.HexColor(
                        BORDER_BLUE
                    ),
                ),
                (
                    "VALIGN",
                    (0, 0),
                    (-1, -1),
                    "MIDDLE",
                ),
                (
                    "LEFTPADDING",
                    (0, 0),
                    (-1, -1),
                    5,
                ),
                (
                    "RIGHTPADDING",
                    (0, 0),
                    (-1, -1),
                    5,
                ),
            ]
        )
    )

    story.append(
        customer_table_pdf
    )

    story.append(
        Spacer(1, 5 * mm)
    )

    # --------------------------------------------------------
    # FINAL BOQ
    # --------------------------------------------------------

    final_boq = prepare_final_boq(
        bom_with_price
    )

    table_data = [
        [
            "S.No.",
            "Part Code",
            "Description",
            "Qty",
            "UOM",
            "Unit Price",
            "Total Price",
        ]
    ]

    for _, row in final_boq.iterrows():

        table_data.append(
            [
                str(
                    row["Final S.No."]
                ),
                str(
                    row["Part Code"]
                ),
                Paragraph(
                    str(
                        row["Description"]
                    ),
                    small_style
                ),
                str(
                    row["Quantity"]
                ),
                str(
                    row["UOM"]
                ),
                (
                    money(
                        row["Unit Price"]
                    )
                    if pd.notna(
                        row["Unit Price"]
                    )
                    else "N/A"
                ),
                (
                    money(
                        row["Total Price"]
                    )
                    if pd.notna(
                        row["Total Price"]
                    )
                    else "N/A"
                ),
            ]
        )

    boq_pdf_table = Table(
        table_data,
        colWidths=[
            15 * mm,
            32 * mm,
            105 * mm,
            16 * mm,
            16 * mm,
            32 * mm,
            35 * mm,
        ],
        repeatRows=1,
    )

    boq_pdf_table.setStyle(
        TableStyle(
            [
                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, 0),
                    colors.HexColor(
                        DARK_BLUE
                    ),
                ),
                (
                    "TEXTCOLOR",
                    (0, 0),
                    (-1, 0),
                    colors.white,
                ),
                (
                    "FONTNAME",
                    (0, 0),
                    (-1, 0),
                    "Helvetica-Bold",
                ),
                (
                    "FONTSIZE",
                    (0, 0),
                    (-1, -1),
                    7,
                ),
                (
                    "GRID",
                    (0, 0),
                    (-1, -1),
                    0.35,
                    colors.HexColor(
                        "#D5E2EA"
                    ),
                ),
                (
                    "VALIGN",
                    (0, 0),
                    (-1, -1),
                    "MIDDLE",
                ),
                (
                    "ALIGN",
                    (0, 0),
                    (0, -1),
                    "CENTER",
                ),
                (
                    "ALIGN",
                    (3, 1),
                    (4, -1),
                    "CENTER",
                ),
                (
                    "ALIGN",
                    (5, 1),
                    (6, -1),
                    "RIGHT",
                ),
            ]
        )
    )

    story.append(
        boq_pdf_table
    )

    story.append(
        Spacer(1, 5 * mm)
    )

    # --------------------------------------------------------
    # FINAL PRICE
    # --------------------------------------------------------

    final_price = 0.0

    if not bom_with_price.empty:

        final_price = float(
            bom_with_price[
                "Total Price"
            ]
            .fillna(0)
            .sum()
        )

    story.append(
        Paragraph(
            f"<b>Final Selling Price: "
            f"{money(final_price)}</b>",
            ParagraphStyle(
                "finalprice",
                parent=normal_style,
                fontSize=11,
                textColor=colors.HexColor(
                    DARK_BLUE
                ),
                alignment=TA_LEFT,
            )
        )
    )

    if internal and cost_data is not None:

        story.append(
            Spacer(1, 3 * mm)
        )

        story.append(
            Paragraph(
                "<b>Internal Cost Summary</b>",
                normal_style
            )
        )

        internal_data = [
            [
                str(item),
                money(value)
                if isinstance(
                    value,
                    (int, float)
                )
                else str(value)
            ]
            for item, value
            in cost_data
        ]

        internal_table = Table(
            [
                ["Item", "Value"]
            ]
            + internal_data,
            colWidths=[
                55 * mm,
                45 * mm
            ],
        )

        internal_table.setStyle(
            TableStyle(
                [
                    (
                        "BACKGROUND",
                        (0, 0),
                        (-1, 0),
                        colors.HexColor(
                            PRIMARY_BLUE
                        ),
                    ),
                    (
                        "TEXTCOLOR",
                        (0, 0),
                        (-1, 0),
                        colors.white,
                    ),
                    (
                        "GRID",
                        (0, 0),
                        (-1, -1),
                        0.3,
                        colors.HexColor(
                            BORDER_BLUE
                        ),
                    ),
                    (
                        "FONTSIZE",
                        (0, 0),
                        (-1, -1),
                        7,
                    ),
                ]
            )
        )

        story.append(
            internal_table
        )

    doc.build(story)

    output.seek(0)

    return output


# ============================================================
# SAVE CONFIGURATION
# ============================================================

def save_configuration(
    user_code,
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

    conn = sqlite3.connect(
        TRACKING_DB
    )

    cursor = conn.cursor()

    created_at = datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    cursor.execute(
        """
        INSERT OR REPLACE INTO configurations
        (
            configuration_id,
            user_code,
            created_at,

            customer_name,
            customer_place,
            problem,
            solution,

            mdc_type,
            configuration,

            pdu_type,
            pdu_part_code,

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
            warranty_amount
        )
        VALUES
        (
            ?, ?, ?,
            ?, ?, ?, ?,
            ?, ?,
            ?, ?,
            ?, ?, ?, ?,
            ?, ?, ?, ?,
            ?, ?, ?
        )
        """,
        (
            user_code,
            user_code,
            created_at,

            st.session_state.customer_name,
            st.session_state.customer_place,
            st.session_state.problem,
            st.session_state.solution,

            st.session_state.mdc_type,
            st.session_state.configuration,

            st.session_state.pdu_type,
            st.session_state.selected_pdu_part,

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
        )
    )

    cursor.execute(
        """
        DELETE FROM configuration_items
        WHERE configuration_id = ?
        """,
        (user_code,)
    )

    if bom is not None and not bom.empty:

        for _, row in bom.iterrows():

            cursor.execute(
                """
                INSERT INTO configuration_items
                (
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
                VALUES
                (
                    ?, ?, ?, ?, ?, ?,
                    ?, ?, ?, ?
                )
                """,
                (
                    user_code,

                    str(
                        row.get(
                            "Component Type",
                            ""
                        )
                    ),

                    str(
                        row.get(
                            "Part Code",
                            ""
                        )
                    ),

                    str(
                        row.get(
                            "Description",
                            ""
                        )
                    ),

                    float(
                        row.get(
                            "Quantity",
                            0
                        )
                    ),

                    str(
                        row.get(
                            "UOM",
                            ""
                        )
                    ),

                    float(
                        row.get(
                            "Unit Cost",
                            0
                        )
                    ),

                    float(
                        row.get(
                            "Total Cost",
                            0
                        )
                    ),

                    (
                        float(
                            row["Unit Price"]
                        )
                        if pd.notna(
                            row.get(
                                "Unit Price"
                            )
                        )
                        else None
                    ),

                    (
                        float(
                            row["Total Price"]
                        )
                        if pd.notna(
                            row.get(
                                "Total Price"
                            )
                        )
                        else None
                    ),
                )
            )

    conn.commit()
    conn.close()


# ============================================================
# HEADER
# ============================================================

st.markdown(
    f"""
    <div style="
        background:linear-gradient(
            135deg,
            {DARK_BLUE},
            #003B5C
        );
        padding:17px 24px;
        border-radius:8px;
        margin-bottom:12px;
        box-shadow:0 3px 10px rgba(0,80,128,0.15);
    ">

        <div style="
            color:white;
            font-size:27px;
            font-weight:700;
            line-height:1.2;
        ">
            Eaton MDC Solution Configurator
        </div>

        <div style="
            color:#DDF2FF;
            font-size:13px;
            margin-top:4px;
        ">
            Modular Data Center Solution Configuration & Pricing
        </div>

    </div>
    """,
    unsafe_allow_html=True
)


# ============================================================
# ACCESS MODE
# ============================================================

with st.sidebar:

    st.header("User Access")

    mode = st.radio(
        "Select User Type",
        [
            "Sales",
            "Internal – MDC"
        ],
        index=(
            0
            if st.session_state.mode == "Sales"
            else 1
        ),
    )

    if mode != st.session_state.mode:

        st.session_state.mode = mode

        if mode == "Sales":

            st.session_state.authenticated = False

        st.rerun()

    if mode == "Internal – MDC":

        if not st.session_state.authenticated:

            st.warning(
                "Internal MDC access requires a password."
            )

            pwd = st.text_input(
                "MDC Password",
                type="password"
            )

            if st.button(
                "Unlock Internal Mode",
                use_container_width=True
            ):

                if pwd == internal_password():

                    st.session_state.authenticated = True

                    st.rerun()

                else:

                    st.error(
                        "Incorrect password."
                    )

        else:

            st.success(
                "Internal mode unlocked."
            )

            if st.button(
                "Lock Internal Mode",
                use_container_width=True
            ):

                st.session_state.authenticated = False

                st.session_state.mode = "Sales"

                st.rerun()


is_internal = (
    st.session_state.mode
    == "Internal – MDC"
    and
    st.session_state.authenticated
)


# ============================================================
# 1. CUSTOMER DETAILS
# ============================================================

st.markdown(
    '<div class="section-ribbon">1. Customer Details</div>',
    unsafe_allow_html=True
)

c1, c2 = st.columns(2)

with c1:

    st.session_state.customer_name = st.text_input(
        "Customer Name *",
        value=st.session_state.customer_name,
        placeholder="Enter customer name"
    )

with c2:

    st.session_state.customer_place = st.text_input(
        "Customer Place",
        value=st.session_state.customer_place,
        placeholder="Enter location"
    )


c3, c4 = st.columns(2)

with c3:

    st.session_state.problem = st.text_area(
        "Problem Description",
        value=st.session_state.problem,
        height=65
    )

with c4:

    st.session_state.solution = st.text_area(
        "Solution",
        value=st.session_state.solution,
        height=65
    )


# ============================================================
# 2. MDC TYPE + CONFIGURATION
# ============================================================

st.markdown(
    '<div class="section-ribbon">2. MDC Type & Configuration</div>',
    unsafe_allow_html=True
)

config_col1, config_col2 = st.columns(
    [1, 3],
    gap="small"
)


with config_col1:

    mdc_types = [
        "Single Rack",
        "Multirack"
    ]

    current_index = (
        mdc_types.index(
            st.session_state.mdc_type
        )
        if st.session_state.mdc_type
        in mdc_types
        else 0
    )

    selected_mdc_type = st.selectbox(
        "MDC Type",
        mdc_types,
        index=current_index,
        key="mdc_type_select"
    )


with config_col2:

    available = configs_df[
        configs_df["MDC Type"]
        == selected_mdc_type
    ].copy()

    available_configs = (
        available["Configuration"]
        .astype(str)
        .tolist()
    )

    if not available_configs:

        available_configs = [
            "Configuration 1"
        ]

    if (
        st.session_state.configuration
        not in available_configs
    ):

        st.session_state.configuration = (
            available_configs[0]
        )

    selected_config = st.selectbox(
        "Configuration",
        available_configs,
        format_func=configuration_display_name,
        index=available_configs.index(
            st.session_state.configuration
        ),
        key="configuration_select"
    )


# ------------------------------------------------------------
# Detect MDC type/configuration change
# ------------------------------------------------------------

if selected_mdc_type != st.session_state.mdc_type:

    st.session_state.mdc_type = (
        selected_mdc_type
    )

    st.session_state.configuration = (
        available_configs[0]
    )

    st.session_state.pdu_type = "None"

    reset_pdu()

    st.session_state.accessory_qty = {}

    st.rerun()


if selected_config != st.session_state.configuration:

    st.session_state.configuration = (
        selected_config
    )

    st.session_state.pdu_type = "None"

    reset_pdu()

    st.session_state.accessory_qty = {}

    st.rerun()


# ------------------------------------------------------------
# Selected configuration
# ------------------------------------------------------------

cfg = selected_config_record()

if cfg is not None:

    config_title = configuration_display_name(
        st.session_state.configuration
    )

    st.markdown(
        f"""
        <div class="selection-card">

            <div class="selection-label">
                SELECTED CONFIGURATION
            </div>

            <div class="selection-value">
                {st.session_state.configuration}
                &nbsp;—&nbsp;
                {config_title}
            </div>

        </div>
        """,
        unsafe_allow_html=True
    )


# ============================================================
# USER CODE
# ============================================================

# Generate a provisional code for current selection.
# It changes automatically with selections.

st.session_state.user_code = (
    generate_user_code()
)

st.markdown(
    f"""
    <div class="user-code-card">

        <div class="user-code-label">
            User Code
        </div>

        <div class="user-code-value">
            {st.session_state.user_code}
        </div>

    </div>
    """,
    unsafe_allow_html=True
)


# ============================================================
# 3. PDU SELECTION
# ============================================================

st.markdown(
    '<div class="section-ribbon">3. PDU Selection</div>',
    unsafe_allow_html=True
)


pdu_types = [
    "Basic",
    "Metered",
    "Switched",
    "None",
]


selected_pdu_type = st.radio(
    "PDU Type",
    pdu_types,
    horizontal=True,
    index=pdu_types.index(
        st.session_state.pdu_type
    ),
    key="pdu_type_radio"
)


if selected_pdu_type != st.session_state.pdu_type:

    st.session_state.pdu_type = (
        selected_pdu_type
    )

    reset_pdu()

    st.rerun()


# ------------------------------------------------------------
# PDU component selection
# ------------------------------------------------------------

if st.session_state.pdu_type != "None":

    pdu_rows = get_pdu_rows()

    if pdu_rows.empty:

        st.warning(
            f"No {st.session_state.pdu_type} PDU "
            "components were found in the Excel master."
        )

        reset_pdu()

    else:

        pdu_options = []

        for _, r in pdu_rows.iterrows():

            pdu_options.append(
                (
                    str(r["Part Code"]),
                    f'{r["Part Code"]} — {r["Description"]}'
                )
            )

        option_labels = [
            x[1]
            for x in pdu_options
        ]

        current_pdu_index = 0

        if st.session_state.selected_pdu_part:

            current_parts = [
                x[0]
                for x in pdu_options
            ]

            if (
                st.session_state.selected_pdu_part
                in current_parts
            ):

                current_pdu_index = (
                    current_parts.index(
                        st.session_state.selected_pdu_part
                    )
                )

        pc1, pc2 = st.columns(
            [5.5, 1.2],
            gap="small"
        )

        with pc1:

            selected_pdu_label = st.selectbox(
                "Select PDU Component",
                option_labels,
                index=current_pdu_index,
                key="pdu_component_select"
            )

        with pc2:

            current_part = pdu_options[
                option_labels.index(
                    selected_pdu_label
                )
            ][0]

            current_qty = int(
                st.session_state.pdu_qty.get(
                    current_part,
                    1
                )
            )

            pdu_quantity = st.number_input(
                "Qty",
                min_value=1,
                max_value=999,
                value=current_qty,
                step=1,
                key="pdu_quantity_input"
            )

        selected_part = pdu_options[
            option_labels.index(
                selected_pdu_label
            )
        ][0]

        st.session_state.selected_pdu_part = (
            selected_part
        )

        st.session_state.pdu_qty = {
            selected_part: pdu_quantity
        }

else:

    reset_pdu()


# ============================================================
# 4. OTHER ACCESSORIES
# ============================================================

st.markdown(
    '<div class="section-ribbon">4. Other Accessories</div>',
    unsafe_allow_html=True
)


# ------------------------------------------------------------
# 4.1 Fire Suppression
# ------------------------------------------------------------

acc1, acc2 = st.columns(
    [3.8, 2.0],
    gap="small"
)

with acc1:

    fire_options = [
        "None",
        "External",
        "Internal"
    ]

    fire_index = fire_options.index(
        st.session_state.fire_suppression
    )

    fire_selection = st.selectbox(
        "3.1 Fire Suppression",
        fire_options,
        index=fire_index,
        key="fire_suppression_select"
    )

with acc2:

    st.markdown(
        """
        <div style="
            font-size:12px;
            color:#64748B;
            padding-top:30px;
        ">
            Select External or Internal
        </div>
        """,
        unsafe_allow_html=True
    )


if fire_selection != st.session_state.fire_suppression:

    st.session_state.fire_suppression = (
        fire_selection
    )


# ------------------------------------------------------------
# 3.2 Camera
# ------------------------------------------------------------

camera_options = [
    "No",
    "Yes"
]

camera_col1, camera_col2 = st.columns(
    [3.8, 2.0],
    gap="small"
)

with camera_col1:

    camera_selection = st.selectbox(
        "3.2 Camera",
        camera_options,
        index=camera_options.index(
            st.session_state.camera_enabled
        ),
        key="camera_select"
    )

with camera_col2:

    st.markdown(
        """
        <div style="
            font-size:12px;
            color:#64748B;
            padding-top:30px;
        ">
            Camera system components are
            taken automatically from Excel.
        </div>
        """,
        unsafe_allow_html=True
    )


st.session_state.camera_enabled = (
    camera_selection
)


# ------------------------------------------------------------
# 3.3 - 3.6 Other accessories
# ------------------------------------------------------------

oa1, oa2 = st.columns(2, gap="small")


with oa1:

    rotating = st.checkbox(
        "3.3 Rotating Keyboard",
        value=st.session_state.rotating_keyboard,
        key="rotating_keyboard_check"
    )

    cable_manager = st.checkbox(
        "3.4 Cable Manager",
        value=st.session_state.cable_manager,
        key="cable_manager_check"
    )


with oa2:

    top_tray = st.checkbox(
        "3.5 Top Cable Tray",
        value=st.session_state.top_cable_tray,
        key="top_cable_tray_check"
    )

    brush_panel = st.checkbox(
        "3.6 Brush Panel",
        value=st.session_state.brush_panel,
        key="brush_panel_check"
    )


st.session_state.rotating_keyboard = rotating
st.session_state.cable_manager = cable_manager
st.session_state.top_cable_tray = top_tray
st.session_state.brush_panel = brush_panel


# ------------------------------------------------------------
# Update accessory quantities
# ------------------------------------------------------------

accessory_flags = {

    "801223664":
        st.session_state.rotating_keyboard,

    "801075237":
        st.session_state.cable_manager,

    "801029022":
        st.session_state.top_cable_tray,

    "801075235":
        st.session_state.brush_panel,
}


for part, selected in accessory_flags.items():

    if selected:

        add_accessory(
            part,
            1
        )

    else:

        remove_accessory(
            part
        )


# ------------------------------------------------------------
# Fire suppression / camera quantities
# ------------------------------------------------------------

# Fire
for fire_type, part in FIRE_SUPPRESSION_PARTS.items():

    if fire_selection == fire_type:

        add_accessory(
            part,
            1
        )

    else:

        remove_accessory(
            part
        )


# Camera
for part in CAMERA_PARTS:

    if camera_selection == "Yes":

        add_accessory(
            part,
            1
        )

    else:

        remove_accessory(
            part
        )


# ============================================================
# CURRENT USER CODE - REFRESH
# ============================================================

st.session_state.user_code = (
    generate_user_code()
)

st.markdown(
    f"""
    <div class="user-code-card">

        <div class="user-code-label">
            Current User Code
        </div>

        <div class="user-code-value">
            {st.session_state.user_code}
        </div>

    </div>
    """,
    unsafe_allow_html=True
)


# ============================================================
# BUILD BOM
# ============================================================

bom = build_bom()


# ============================================================
# COST SUMMARY
# ============================================================

base_cost, optional_cost, pdu_cost, total_cost = (
    cost_summary(bom)
)


# ============================================================
# 5. COST SUMMARY
# INTERNAL ONLY
# ============================================================

if is_internal:

    st.markdown(
        '<div class="section-ribbon">5. Cost Summary — Internal Only</div>',
        unsafe_allow_html=True
    )

    a, b, c, d = st.columns(4)

    with a:
        price_box(
            "Base Cost",
            base_cost
        )

    with b:
        price_box(
            "Optional Cost",
            optional_cost
        )

    with c:
        price_box(
            "PDU Cost",
            pdu_cost
        )

    with d:
        price_box(
            "Total Cost",
            total_cost
        )


# ============================================================
# 6. COST TO PRICE CONVERSION
# INTERNAL ONLY
# ============================================================

margin_pct = st.session_state.margin_pct
freight = st.session_state.freight
installation = st.session_state.installation
warranty_pct = st.session_state.warranty_pct


if is_internal:

    st.markdown(
        '<div class="section-ribbon">6. Cost to Price Conversion — Internal Only</div>',
        unsafe_allow_html=True
    )

    p1, p2, p3, p4 = st.columns(4)

    with p1:

        margin_pct = st.number_input(
            "Margin (%)",
            min_value=0.0,
            max_value=99.0,
            value=float(
                st.session_state.margin_pct
            ),
            step=0.5
        )

        st.session_state.margin_pct = (
            margin_pct
        )

    with p2:

        freight = st.number_input(
            "Freight",
            min_value=0.0,
            value=float(
                st.session_state.freight
            ),
            step=500.0
        )

        st.session_state.freight = (
            freight
        )

    with p3:

        installation = st.number_input(
            "Installation",
            min_value=0.0,
            value=float(
                st.session_state.installation
            ),
            step=500.0
        )

        st.session_state.installation = (
            installation
        )

    with p4:

        warranty_pct = st.number_input(
            "Warranty (%)",
            min_value=0.0,
            max_value=100.0,
            value=float(
                st.session_state.warranty_pct
            ),
            step=0.5
        )

        st.session_state.warranty_pct = (
            warranty_pct
        )


margin_price = (
    total_cost
    /
    (1 - margin_pct / 100)
    if margin_pct < 100
    else 0
)

final_selling_price = (
    margin_price
    + freight
    + installation
)

warranty_amount = (
    margin_price
    * warranty_pct
    / 100
)


if is_internal:

    a, b, c, d = st.columns(4)

    with a:
        price_box(
            "Margin Price",
            margin_price
        )

    with b:
        price_box(
            "After Freight",
            margin_price + freight
        )

    with c:
        price_box(
            "Final Selling Price",
            final_selling_price
        )

    with d:
        price_box(
            "Warranty Amount",
            warranty_amount
        )


# ============================================================
# FINAL BOQ WITH SELLING PRICE
# ============================================================

bom_with_price, margin_price, final_selling_price = (
    add_selling_prices(
        bom,
        total_cost,
        margin_pct,
        freight,
        installation
    )
)


# ============================================================
# 7. FINAL BOQ
# ============================================================

st.markdown(
    '<div class="section-ribbon">7. Final BOQ</div>',
    unsafe_allow_html=True
)


if not bom_with_price.empty:

    selected_name = configuration_display_name(
        st.session_state.configuration
    )

    st.markdown(
        f"""
        <div class="selection-card">

            <div class="selection-label">
                SELECTED CONFIGURATION
            </div>

            <div class="selection-value">
                {selected_name}
            </div>

        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown(
        final_boq_html(
            bom_with_price
        ),
        unsafe_allow_html=True
    )

else:

    st.info(
        "No components available for the selected configuration."
    )


# ============================================================
# FINAL SELLING PRICE + DOWNLOADS
# ============================================================

st.markdown(
    "<br>",
    unsafe_allow_html=True
)

download_left, download_excel, download_pdf = st.columns(
    [2.2, 1.2, 1.2],
    gap="small"
)


with download_left:

    price_box(
        "Final Selling Price",
        final_selling_price
    )


customer_ready = (
    bool(
        st.session_state.customer_name
        and
        st.session_state.customer_name.strip()
    )
)


# ------------------------------------------------------------
# Export preparation
# ------------------------------------------------------------

internal_cost_data = [

    ["Base Cost", base_cost],

    ["Optional Cost", optional_cost],

    ["PDU Cost", pdu_cost],

    ["Total Cost", total_cost],

    ["Margin %", margin_pct],

    ["Margin Price", margin_price],

    ["Freight", freight],

    ["Installation", installation],

    ["Final Selling Price",
     final_selling_price],

    ["Warranty %",
     warranty_pct],

    ["Warranty Amount",
     warranty_amount],
]


sales_excel = None
internal_excel = None
sales_pdf = None
internal_pdf = None


if customer_ready and not bom_with_price.empty:

    sales_excel = excel_bytes(
        internal=False,
        bom=bom_with_price
    )

    sales_pdf = pdf_bytes(
        bom_with_price,
        internal=False
    )

    if is_internal:

        internal_excel = excel_bytes(
            internal=True,
            bom=bom_with_price,
            cost_data=internal_cost_data
        )

        internal_pdf = pdf_bytes(
            bom_with_price,
            internal=True,
            cost_data=internal_cost_data
        )


with download_excel:

    if not customer_ready:

        st.button(
            "⬇ Excel",
            disabled=True,
            use_container_width=True
        )

    elif is_internal:

        st.download_button(
            "⬇ Internal Excel",
            data=internal_excel,
            file_name=(
                f"{st.session_state.user_code}"
                "_Internal_BOQ.xlsx"
            ),
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True
        )

    else:

        st.download_button(
            "⬇ Excel",
            data=sales_excel,
            file_name=(
                f"{st.session_state.user_code}"
                "_Final_BOQ.xlsx"
            ),
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True
        )


with download_pdf:

    if not customer_ready:

        st.button(
            "⬇ PDF",
            disabled=True,
            use_container_width=True
        )

    elif is_internal:

        st.download_button(
            "⬇ Internal PDF",
            data=internal_pdf,
            file_name=(
                f"{st.session_state.user_code}"
                "_Internal_BOQ.pdf"
            ),
            mime="application/pdf",
            use_container_width=True
        )

    else:

        st.download_button(
            "⬇ PDF",
            data=sales_pdf,
            file_name=(
                f"{st.session_state.user_code}"
                "_Final_BOQ.pdf"
            ),
            mime="application/pdf",
            use_container_width=True
        )


if not customer_ready:

    st.markdown(
        """
        <div class="download-note">
            Customer Name is mandatory before downloading
            the Excel or PDF.
        </div>
        """,
        unsafe_allow_html=True
    )


# ============================================================
# 8. SAVE CONFIGURATION
# ============================================================

if is_internal:

    st.markdown(
        '<div class="section-ribbon">8. Save Configuration</div>',
        unsafe_allow_html=True
    )

    save_col1, save_col2 = st.columns(
        [1.5, 5]
    )

    with save_col1:

        if st.button(
            "💾 Save Configuration",
            use_container_width=True,
            type="primary"
        ):

            if not customer_ready:

                st.error(
                    "Customer Name is mandatory."
                )

            elif bom_with_price.empty:

                st.error(
                    "No BOQ available to save."
                )

            else:

                save_configuration(
                    user_code=
                        st.session_state.user_code,

                    bom=bom_with_price,

                    base_cost=base_cost,
                    optional_cost=optional_cost,
                    pdu_cost=pdu_cost,
                    total_cost=total_cost,

                    margin_pct=margin_pct,
                    freight=freight,
                    installation=installation,
                    warranty_pct=warranty_pct,

                    margin_price=margin_price,
                    final_selling_price=
                        final_selling_price,
                    warranty_amount=
                        warranty_amount,
                )

                st.session_state.configuration_saved = (
                    True
                )

                st.success(
                    "Configuration saved successfully."
                )


# ============================================================
# 9. CONFIGURATION HISTORY
# ============================================================

st.markdown(
    '<div class="section-ribbon">9. Configuration History</div>',
    unsafe_allow_html=True
)


if is_internal:

    conn = sqlite3.connect(
        TRACKING_DB
    )

    history_df = pd.read_sql_query(
        """
        SELECT
            user_code AS "User Code",
            substr(created_at, 1, 10) AS "Date",
            customer_name AS "Customer Name"
        FROM configurations
        ORDER BY id DESC
        """,
        conn
    )

    conn.close()

    if not history_df.empty:

        st.dataframe(
            history_df,
            use_container_width=True,
            hide_index=True
        )

    else:

        st.info(
            "No configuration history available."
        )

else:

    st.caption(
        "Configuration history is available only "
        "to authenticated Internal – MDC users."
    )


# ============================================================
# FOOTER
# ============================================================

st.divider()

st.caption(
    "MDC Solution | Eaton MDC Solution Configurator | "
    f"Current Date: {datetime.now().strftime('%d-%m-%Y')}"
)
