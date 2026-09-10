import os
import sqlite3
from io import BytesIO
from datetime import datetime

import pandas as pd
import streamlit as st

# ============================================================
# MDC SOLUTION - VERSION 1 (REAL SINGLE-RACK DATA)
# Data source: MDC_Master_V1.xlsx (same folder as app.py)
#
# Single Rack:
#   Configuration 1-4 = real data from supplied MDC BOQ
#
# Multirack:
#   Configuration 1-9 = XXX placeholders for future update
# ============================================================

st.set_page_config(
    page_title="MDC Solution",
    page_icon="🏢",
    layout="wide",
    # initial_sidebar_state="expanded",
)
# ============================================================
# EATON MDC HEADER
# ============================================================

# st.html("""
# <div style="
#     background: linear-gradient(135deg, #005EB8, #003B71);
#     padding: 22px 30px;
#     border-radius: 10px;
#     margin-bottom: 25px;
#     box-shadow: 0 4px 12px rgba(0, 59, 113, 0.18);
# ">
#     <div style="
#         color: white;
#         font-size: 32px;
#         font-weight: 700;
#         letter-spacing: 0.3px;
#         line-height: 1.2;
#     ">
#         Eaton MDC Solution Configurator
#     </div>

#     <div style="
#         color: #E6F2FF;
#         font-size: 16px;
#         font-weight: 400;
#         margin-top: 7px;
#     ">
#         Modular Data Center Solution Configuration &amp; Pricing
#     </div>
# </div>
# """)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MASTER_FILE = os.path.join(BASE_DIR, "MDC_Master_V1.xlsx")
DEMO_INTERNAL_PASSWORD = "MDC@123"  # Change before production.

# ============================================================
# MDC CONFIGURATION TRACKING DATABASE
# No login required
# ============================================================

TRACKING_DB = os.path.join(BASE_DIR, "MDC_Tracking.db")


def init_tracking_db():
    conn = sqlite3.connect(TRACKING_DB)

    cursor = conn.cursor()

    cursor.execute("""
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
            warranty_amount REAL
        )
    """)

    cursor.execute("""
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
        # --------------------------------------------------------
    # Persistent Excel download counter
    # --------------------------------------------------------
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS download_counter (
            id INTEGER PRIMARY KEY,
            download_count INTEGER NOT NULL DEFAULT 0
        )
    """)

    cursor.execute("""
        INSERT OR IGNORE INTO download_counter (id, download_count)
        VALUES (1, 0)
    """)

    conn.commit()
    conn.close()
    
def increment_download_count():
    conn = sqlite3.connect(TRACKING_DB)
    cursor = conn.cursor()

    cursor.execute("""
        UPDATE download_counter
        SET download_count = download_count + 1
        WHERE id = 1
    """)

    cursor.execute("""
        SELECT download_count
        FROM download_counter
        WHERE id = 1
    """)

    count = cursor.fetchone()[0]

    conn.commit()
    conn.close()

    return count

def generate_configuration_id():
    """
    Generates IDs like:

    MDC-20260907-0001
    MDC-20260907-0002
    MDC-20260907-0003
    """

    today = datetime.now().strftime("%Y%m%d")

    conn = sqlite3.connect(TRACKING_DB)
    cursor = conn.cursor()

    cursor.execute("""
        SELECT COUNT(*)
        FROM configurations
        WHERE configuration_id LIKE ?
    """, (f"MDC-{today}-%",))

    count = cursor.fetchone()[0] + 1

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
    cursor = conn.cursor()

    created_at = datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    # --------------------------------------------------------
    # Save configuration master record
    # --------------------------------------------------------

    cursor.execute("""
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
            warranty_amount
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
    ))

    # --------------------------------------------------------
    # Remove old items if same configuration is saved again
    # --------------------------------------------------------

    cursor.execute("""
        DELETE FROM configuration_items
        WHERE configuration_id = ?
    """, (configuration_id,))

    # --------------------------------------------------------
    # Save complete BOM
    # --------------------------------------------------------

    if bom is not None and not bom.empty:

        for _, row in bom.iterrows():

            unit_price = row.get("Unit Price", None)
            total_price = row.get("Total Price", None)

            cursor.execute("""
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
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                configuration_id,

                str(row.get("Component Type", "")),
                str(row.get("Part Code", "")),
                str(row.get("Description", "")),

                float(row.get("Quantity", 0))
                if pd.notna(row.get("Quantity"))
                else 0,

                str(row.get("UOM", "")),

                float(row.get("Unit Cost", 0))
                if pd.notna(row.get("Unit Cost"))
                else 0,

                float(row.get("Total Cost", 0))
                if pd.notna(row.get("Total Cost"))
                else 0,

                float(unit_price)
                if pd.notna(unit_price)
                else None,

                float(total_price)
                if pd.notna(total_price)
                else None,
            ))

    conn.commit()
    conn.close()


# Initialize database when application starts
init_tracking_db()

# ------------------------------------------------------------
# Load master data
# ------------------------------------------------------------
@st.cache_data
def load_master():
    configs = pd.read_excel(MASTER_FILE, sheet_name="Configurations")
    components = pd.read_excel(MASTER_FILE, sheet_name="Components")
    accessories = pd.read_excel(MASTER_FILE, sheet_name="Accessories")
    pdus = pd.read_excel(MASTER_FILE, sheet_name="PDUs")

    # Fill merged TYPE cells downward
    # Example:
    # BASIC -> BASIC -> BASIC -> BASIC
    # METERED -> METERED -> ...
    # SWITCHED -> SWITCHED -> ...
    pdus["Type"] = pdus["Type"].ffill()

    return configs, components, accessories, pdus

configs_df, components_df, accessories_df, pdus_df = load_master()

# ------------------------------------------------------------
# Session state
# ------------------------------------------------------------
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
    
# Load persistent Excel download count
if st.session_state.user_count == 0:
    conn = sqlite3.connect(TRACKING_DB)
    cursor = conn.cursor()

    cursor.execute("""
        SELECT download_count
        FROM download_counter
        WHERE id = 1
    """)

    result = cursor.fetchone()
    st.session_state.user_count = result[0] if result else 0

    conn.close()

def money(value):
    return f"₹ {value:,.2f}"


def price_box(label, value):
    st.markdown(
        f"""
        <div style="padding:4px 0 12px 0; min-height:82px; overflow:visible;">
            <div style="font-size:16px; color:#4b5563; margin-bottom:7px;">
                {label}
            </div>
            <div style="font-size:30px; font-weight:600; color:#30333d;
                        white-space:nowrap; overflow:visible;">
                {money(value)}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

def internal_password():
    # Streamlit Cloud / deployment can use st.secrets["MDC_INTERNAL_PASSWORD"].
    try:
        return st.secrets["MDC_INTERNAL_PASSWORD"]
    except Exception:
        return DEMO_INTERNAL_PASSWORD
def generate_user_code():
    codes = []

    # --------------------------------------------------------
    # MDC CONFIGURATION
    # --------------------------------------------------------
    config_text = str(st.session_state.configuration)

    if config_text:
        config_number = config_text.split()[-1]
        codes.append(f"C{config_number}")

    # --------------------------------------------------------
    # FIRE SUPPRESSION
    # --------------------------------------------------------
    fire_code = None

    for part in FIRE_SUPPRESSION_PARTS:
        if st.session_state.accessory_qty.get(part, 0) > 0:

            description = ""

            if part in optional_lookup:
                description = str(
                    optional_lookup[part]["Description"]
                ).upper()

            if "EXTERNAL" in description:
                fire_code = "F-EXT"

            elif "INTERNAL" in description:
                fire_code = "F-INT"

    if fire_code:
        codes.append(fire_code)

    # --------------------------------------------------------
    # CAMERA
    # --------------------------------------------------------
    if any(
        st.session_state.accessory_qty.get(part, 0) > 0
        for part in CAMERA_PARTS
    ):
        codes.append("CAM")

    # --------------------------------------------------------
    # OTHER ACCESSORIES
    # --------------------------------------------------------
    accessory_code_map = {
        "801223664": "KT",    # Rotating Keyboard Tray
        "801075237": "CM",    # Cable Manager
        "801029022": "TCT",   # Top Cable Tray
        "801075235": "BP",    # Brush Panel
    }

    for part, code in accessory_code_map.items():
        if st.session_state.accessory_qty.get(part, 0) > 0:
            codes.append(code)

    # --------------------------------------------------------
    # PDU
    # --------------------------------------------------------
    for part, qty in st.session_state.pdu_qty.items():

        if qty <= 0:
            continue

        pdu_rows = pdus_df[
            pdus_df["Part Code"].astype(str).str.strip() == str(part).strip()
        ]

        if not pdu_rows.empty:

            pdu_type = str(
                pdu_rows.iloc[0]["Type"]
            ).strip().upper()

            pdu_code_map = {
                "BASIC": "B-PDU",
                "METERED": "M-PDU",
                "SWITCHED": "S-PDU",
            }

            if pdu_type in pdu_code_map:
                codes.append(pdu_code_map[pdu_type])

        break

    return "-".join(codes)
def handle_excel_download():
    # Generate user code based on current selections
    st.session_state.user_code = generate_user_code()

    # Increment persistent download count
    st.session_state.user_count = increment_download_count()

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

def build_bom():
    rows = []

    # Configuration BOM
    for _, r in selected_components().iterrows():
        cost = r["Unit Cost"]
        qty = float(r["Quantity"])
        rows.append({
            "S.No.": len(rows) + 1,
            "Component Type": "Base (Configuration)",
            "Part Code": r["Part Code"] if pd.notna(r["Part Code"]) and str(r["Part Code"]).strip() and str(r["Part Code"]).lower() != "nan" else "",
            "Description": r["Description"],
            "Quantity": qty,
            "UOM": r["UOM"],
            "Unit Cost": cost,
            "Total Cost": cost * qty if pd.notna(cost) else None,
            "Source": "Configuration",
        })

    # Optional accessories
    for _, r in accessories_df.iterrows():
        part = str(r["Part Code"])
        qty = float(st.session_state.accessory_qty.get(part, 0))
        if qty > 0:
            cost = r["Unit Cost"]
            rows.append({
                "S.No.": len(rows) + 1,
                "Component Type": "Optional Accessory",
                "Part Code": part if part.strip() and part.lower() != "nan" else "",
                "Description": r["Description"],
                "Quantity": qty,
                "UOM": r["UOM"],
                "Unit Cost": cost,
                "Total Cost": cost * qty if pd.notna(cost) else None,
                "Source": "Optional Accessory",
            })

    # PDU
    for _, r in pdus_df.iterrows():
        part = str(r["Part Code"])
        qty = float(st.session_state.pdu_qty.get(part, 0))
        if qty > 0:
            cost = r["Unit Cost"]
            desc = f'{r["Description"]} | Type: {r["Type"]} | C13: {r["C13"]} | C19: {r["C19"]}'
            rows.append({
                "S.No.": len(rows) + 1,
                "Component Type": "PDU",
                "Part Code": part if part.strip() and part.lower() != "nan" else "",
                "Description": desc,
                "Quantity": qty,
                "UOM": r["UOM"],
                "Unit Cost": cost,
                "Total Cost": cost * qty if pd.notna(cost) else None,
                "Source": "PDU",
            })

    return pd.DataFrame(rows)

def cost_summary(bom):
    cfg = selected_config_record()
    base_cost = float(cfg["Base Cost"]) if cfg is not None and pd.notna(cfg["Base Cost"]) else 0.0

    optional_cost = 0.0
    pdu_cost = 0.0

    if not bom.empty:
        optional_cost = float(
            bom.loc[bom["Source"] == "Optional Accessory", "Total Cost"]
            .fillna(0).sum()
        )
        pdu_cost = float(
            bom.loc[bom["Source"] == "PDU", "Total Cost"]
            .fillna(0).sum()
        )

    total_cost = base_cost + optional_cost + pdu_cost
    return base_cost, optional_cost, pdu_cost, total_cost

def add_selling_prices(bom, total_cost, margin_pct, freight, installation):
    result = bom.copy()

    # Cost-based margin conversion.
    margin_price = total_cost / (1 - margin_pct / 100) if margin_pct < 100 else 0
    final_selling_price = margin_price + freight + installation

    # Allocate the final selling price proportionally to known-cost BOM lines.
    # This makes BOM Total Price reconcile to the final selling price.
    known_cost_total = result["Total Cost"].fillna(0).sum() if not result.empty else 0

    if known_cost_total > 0:
        result["Total Price"] = result["Total Cost"].fillna(0) / known_cost_total * final_selling_price
        result["Unit Price"] = result["Total Price"] / result["Quantity"]
    else:
        result["Total Price"] = pd.NA
        result["Unit Price"] = pd.NA

    return result, margin_price, final_selling_price

def customer_table():
    return pd.DataFrame([
        ["Customer Name", st.session_state.customer_name],
        ["Customer Place", st.session_state.customer_place],
        ["Problem Description", st.session_state.problem],
        ["Solution", st.session_state.solution],
        ["MDC Type", st.session_state.mdc_type],
        ["Configuration", st.session_state.configuration],
    ], columns=["Field", "Value"])

def excel_bytes(internal=False, bom=None, cost_data=None):
    output = BytesIO()
    cust = customer_table()

    if bom is None:
        bom = build_bom()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        cust.to_excel(writer, index=False, sheet_name="Customer & Configuration")

        sales_cols = [
            "S.No.", "Component Type", "Part Code", "Description", "Quantity",
            "UOM", "Unit Price", "Total Price"
        ]
        sales_bom = bom[sales_cols].copy()
        sales_bom.to_excel(writer, index=False, sheet_name="Final BOM")

        if internal and cost_data is not None:
            internal_bom = bom.copy()
            internal_cols = [
                "S.No.", "Component Type", "Part Code", "Description", "Quantity", "UOM",
                "Unit Cost", "Total Cost", "Unit Price", "Total Price"
            ]
            internal_bom[internal_cols].to_excel(
                writer, index=False, sheet_name="Internal Cost BOM"
            )

            pd.DataFrame(cost_data, columns=["Item", "Value"]).to_excel(
                writer, index=False, sheet_name="Cost Summary"
            )

    output.seek(0)
    return output

# ------------------------------------------------------------
# Header
# ------------------------------------------------------------
# st.markdown(
#     """
#     <div style="
#         display:flex;
#         align-items:center;
#         gap:14px;
#         padding:12px 0 16px 0;
#     ">
#         <div style="
#             width:8px;
#             height:55px;
#             background:#0167C9;
#             border-radius:4px;
#         "></div>

#         <div>
#             <div style="
#                 font-size:36px;
#                 font-weight:700;
#                 color:#004B91;
#                 line-height:1.1;
#             ">
#                 MDC Solution
#             </div>

#             <div style="
#                 font-size:16px;
#                 color:#64748B;
#                 margin-top:5px;
#             ">
#                 Rack Configuration & Solution Selection
#             </div>
#         </div>
#     </div>
#     """,
#     unsafe_allow_html=True,
# )
st.html("""
<div style="
    background: linear-gradient(135deg, #005EB8, #003B71);
    padding: 22px 30px;
    border-radius: 10px;
    margin-bottom: 25px;
    box-shadow: 0 4px 12px rgba(0, 59, 113, 0.18);
">
    <div style="
        color: white;
        font-size: 32px;
        font-weight: 700;
        letter-spacing: 0.3px;
        line-height: 1.2;
    ">
        Eaton MDC Solution Configurator
    </div>

    <div style="
        color: #E6F2FF;
        font-size: 16px;
        font-weight: 400;
        margin-top: 7px;
    ">
        Modular Data Center Solution Configuration &amp; Pricing
    </div>
</div>
""")
# ============================================================
# USER CODE / USER COUNT / DATE
# ============================================================

current_date = datetime.now().strftime("%d-%m-%Y")

st.html(f"""
<div style="
    background-color:#F7FBFF;
    border:1px solid #B8D8F5;
    border-radius:8px;
    padding:14px 18px;
    margin:0 0 20px 0;
">

    <div style="
        display:flex;
        justify-content:space-between;
        text-align:center;
        gap:20px;
    ">

        <div style="flex:1;">
            <div style="
                font-size:13px;
                color:#64748B;
                margin-bottom:5px;
            ">
                USER CODE
            </div>

            <div style="
                font-size:20px;
                font-weight:700;
                color:#003B71;
            ">
                {st.session_state.user_code}
            </div>
        </div>

        <div style="flex:1;">
            <div style="
                font-size:13px;
                color:#64748B;
                margin-bottom:5px;
            ">
                USER COUNT
            </div>

            <div style="
                font-size:20px;
                font-weight:700;
                color:#003B71;
            ">
                {st.session_state.user_count}
            </div>
        </div>

        <div style="flex:1;">
            <div style="
                font-size:13px;
                color:#64748B;
                margin-bottom:5px;
            ">
                DATE
            </div>

            <div style="
                font-size:20px;
                font-weight:700;
                color:#003B71;
            ">
                {current_date}
            </div>
        </div>

    </div>

</div>
""")
# ------------------------------------------------------------
# Access mode
# ------------------------------------------------------------
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
            pwd = st.text_input("MDC Password", type="password")
            if st.button("Unlock Internal Mode", use_container_width=True):
                if pwd == internal_password():
                    st.session_state.authenticated = True
                    st.rerun()
                else:
                    st.error("Incorrect password.")
        else:
            st.success("Internal mode unlocked.")
            if st.button("Lock Internal Mode", use_container_width=True):
                st.session_state.authenticated = False
                st.session_state.mode = "Sales"
                st.rerun()

is_internal = (
    st.session_state.mode == "Internal – MDC"
    and st.session_state.authenticated
)

# ------------------------------------------------------------
# 1 Customer details
# ------------------------------------------------------------
customer_name = st.text_input(
    "Customer Name",
    value=st.session_state.customer_name,
    key="customer_name_input",
    placeholder="Enter customer name"
)

st.session_state.customer_name = customer_name.strip()
# ============================================================
# 2. MDC Type & Configuration
# ============================================================

st.html("""
<div style="
    background: linear-gradient(135deg, #005EB8, #003B71);
    color: white;
    padding: 10px 16px;
    border-radius: 8px;
    margin: 20px 0 15px 0;
    font-size: 18px;
    font-weight: 700;
">
    2. MDC TYPE & CONFIGURATION
</div>
""")

mdc_type = st.radio(
    "MDC Type",
    ["Single Rack", "Multirack"],
    horizontal=True,
    index=0 if st.session_state.mdc_type == "Single Rack" else 1,
)

if mdc_type != st.session_state.mdc_type:

    st.session_state.mdc_type = mdc_type
    st.session_state.configuration = "Configuration 1"

    st.session_state.accessory_qty = {}
    st.session_state.pdu_qty = {}

    # New configuration ID
    st.session_state.configuration_id = generate_configuration_id()
    st.session_state.configuration_saved = False

    st.rerun()

available = configs_df[
    configs_df["MDC Type"] == st.session_state.mdc_type
].copy()

labels = available["Configuration"].tolist()

# if labels:
#     st.session_state.configuration = st.selectbox(
#         "Select Configuration",
#         labels,
#         index=(
#             labels.index(st.session_state.configuration)
#             if st.session_state.configuration in labels
#             else 0
#         ),
#     )
if labels:

    configuration_display_names = {
        "Configuration 1":
            "Configuration 1 - 1SR, 42U×800W×1200D, 3.5KW, W/O Dehumidifier",

        "Configuration 2":
            "Configuration 2 - 1SR, 42U×800W×1200D, 7KW, Dehumidifier",

        "Configuration 3":
            "Configuration 3 - 1SR, 42U×800W×1200D, 7KW, W/O Dehumidifier",

        "Configuration 4":
            "Configuration 4 - 1SR, 42U×800W×1200D, 7KW, Dehumidifier",
    }

    st.session_state.configuration = st.selectbox(
        "Select Configuration",
        labels,
        index=(
            labels.index(st.session_state.configuration)
            if st.session_state.configuration in labels
            else 0
        ),
        format_func=lambda x: configuration_display_names.get(x, x),
    )

# ------------------------------------------------------------
# Selected Configuration Display
# ------------------------------------------------------------

# cfg = selected_config_record()

# if cfg is not None:
#     st.markdown(
#         f"""
#         <div style="
#             background-color: #005EB8;
#             color: white;
#             padding: 15px 20px;
#             border-radius: 10px;
#             font-size: 18px;
#             font-weight: 600;
#             margin-top: 10px;
#             margin-bottom: 10px;
#         ">
#             Selected Configuration:
#             {cfg["Configuration"]} — {cfg["Configuration Title"]}
#         </div>
#         """,
#         unsafe_allow_html=True
#     )

# ------------------------------------------------------------
# 3 Optional accessories
# ------------------------------------------------------------
# st.header("3. Optional Accessories")

# for _, r in accessories_df.iterrows():
#     part = str(r["Part Code"])
#     key_check = f"acc_check_{part}"
#     key_qty = f"acc_qty_{part}"

#     col1, col2 = st.columns([5.5, 1.8], vertical_alignment="center")

#     with col1:
#         selected = st.checkbox(
#             f'{part} — {r["Description"]}',
#             value=st.session_state.accessory_qty.get(part, 0) > 0,
#             key=key_check
#         )

#     with col2:
#         if selected:
#             qty = st.number_input(
#                 "Quantity",
#                 min_value=1,
#                 step=1,
#                 value=st.session_state.accessory_qty.get(part, 1),
#                 key=key_qty,
#                 label_visibility="visible"
#             )
#             st.session_state.accessory_qty[part] = qty
#         else:
#             st.session_state.accessory_qty.pop(part, None)
# ------------------------------------------------------------
# 3 PDU selection
# ------------------------------------------------------------
st.html("""
<div style="
    background: linear-gradient(135deg, #005EB8, #003B71);
    color: white;
    padding: 10px 16px;
    border-radius: 8px;
    margin: 20px 0 15px 0;
    font-size: 18px;
    font-weight: 700;
">
    3. PDU SELECTION
</div>
""")

# ------------------------------------------------------------
# PDU selection - segregated by PDU TYPE
# ------------------------------------------------------------

pdu_types = [
    "None",
    "Basic PDU",
    "Metered PDU",
    "Switched PDU"
]

# Select PDU type and PDU model side-by-side
col1, col2 = st.columns([2, 5])

with col1:
    selected_pdu_type = st.selectbox(
        "PDU Type",
        pdu_types,
        index=0,
        key="pdu_type_selection"
    )

with col2:

    if selected_pdu_type != "None":

        # Map UI names to Excel TYPE values
        type_mapping = {
            "Basic PDU": "BASIC",
            "Metered PDU": "METERED",
            "Switched PDU": "SWITCHED",
        }

        excel_pdu_type = type_mapping[selected_pdu_type]

        # Fetch only the selected PDU TYPE from Excel
        filtered_pdus = pdus_df[
            pdus_df["Type"]
            .astype(str)
            .str.strip()
            .str.upper()
            == excel_pdu_type
        ].copy()

        if not filtered_pdus.empty:

            pdu_options = [
                f'{r["Part Code"]} — {r["Description"]}'
                for _, r in filtered_pdus.iterrows()
            ]

            selected_pdu = st.selectbox(
                "Select PDU",
                pdu_options,
                index=0,
                key="pdu_model_selection"
            )

            # Get selected PDU row
            selected_index = pdu_options.index(selected_pdu)
            selected_row = filtered_pdus.iloc[selected_index]

            part = str(selected_row["Part Code"])

            # PDU quantity is automatically 1
            st.session_state.pdu_qty = {
                part: 1
            }

        else:
            st.warning(
                f"No {selected_pdu_type} options found in the Excel data."
            )

    else:

        # No PDU selected
        st.session_state.pdu_qty = {}

    # else:
    #     st.warning(f"No {selected_pdu_type} options found in the Excel data.")

# ============================================================
# 4. OPTIONAL ACCESSORIES
# ============================================================

st.html("""
<div style="
    background: linear-gradient(135deg, #005EB8, #003B71);
    color: white;
    padding: 10px 16px;
    border-radius: 8px;
    margin: 20px 0 15px 0;
    font-size: 18px;
    font-weight: 700;
">
    4. OTHER ACCESSORIES
</div>
""")

# ------------------------------------------------------------
# PART CODES
# ------------------------------------------------------------

FIRE_SUPPRESSION_PARTS = [
    "801073203",   # FIRE SUPR EXT42U...
    "HRD-XH1C",    # FIRE SUPPRESS, RACK MNT...
]

CAMERA_PARTS = [
    "801303201",   # CAMERA, 4MP VANDAL
    "801303202",   # CAMERA, NVR 4 CHA INT
    "801303204",   # CAMERA, POE GB 4P
    "801303206",   # CAMERA, CAT 5 CABLE RJ45
    "801303208",   # CAMERA, SVR HDD 1TB
    "801303203",   # CAMERA, SVR HDD 4TB
]

# ------------------------------------------------------------
# CREATE PART-CODE LOOKUP FROM accessories_df
# ------------------------------------------------------------

optional_lookup = {}

for _, r in accessories_df.iterrows():

    part = str(r["Part Code"]).strip()

    if part and part.lower() != "nan":
        optional_lookup[part] = r


# ============================================================
# FIRE SUPPRESSION
# ============================================================

st.subheader("3.1 Fire Suppression")

fire_current = "None"

if st.session_state.accessory_qty.get("801073203", 0) > 0:
    fire_current = "External"
elif st.session_state.accessory_qty.get("HRD-XH1C", 0) > 0:
    fire_current = "Internal"

fire_selection = st.radio(
    "Fire Suppression",
    ["None", "External", "Internal"],
    index=["None", "External", "Internal"].index(fire_current),
    horizontal=True,
    key="fire_suppression_selection"
)

# Remove both first
st.session_state.accessory_qty.pop("801073203", None)
st.session_state.accessory_qty.pop("HRD-XH1C", None)

# Add selected one
if fire_selection == "External":
    st.session_state.accessory_qty["801073203"] = 1

elif fire_selection == "Internal":
    st.session_state.accessory_qty["HRD-XH1C"] = 1
# # ============================================================
# # CAMERA SYSTEM
# # ============================================================

# st.subheader("📷 Camera System")

# st.caption(
#     "Selecting the Camera System automatically includes "
#     "the required camera, NVR, PoE, CAT 5 cable and storage."
# )

# camera_selected = st.checkbox(
#     "Enable Camera System",
#     value=any(
#         st.session_state.accessory_qty.get(part, 0) > 0
#         for part in CAMERA_PARTS
#     ),
#     key="camera_system"
# )


# if camera_selected:

#     # --------------------------------------------------------
#     # AUTOMATICALLY ADD CAMERA COMPONENTS
#     # --------------------------------------------------------

#     for part in CAMERA_PARTS:

#         if part in optional_lookup:

#             st.session_state.accessory_qty[part] = 1

#     st.success(
#         "Camera System selected → "
#         "all required camera components automatically included."
#     )

#     # --------------------------------------------------------
#     # SHOW INCLUDED COMPONENTS
#     # --------------------------------------------------------

#     st.markdown("**Included Camera Components:**")

#     for part in CAMERA_PARTS:

#         if part in optional_lookup:

#             r = optional_lookup[part]

#             st.write(
#                 f"✓ **{part}** — {r['Description']}"
#             )

# else:

#     # Remove all camera components
#     for part in CAMERA_PARTS:

#         st.session_state.accessory_qty.pop(
#             part,
#             None
#         )
# ============================================================
# CAMERA
# ============================================================

st.subheader("3.2 Camera")

camera_current = "No"

if any(
    st.session_state.accessory_qty.get(part, 0) > 0
    for part in CAMERA_PARTS
):
    camera_current = "Yes"

camera_selection = st.radio(
    "Camera",
    ["Yes", "No"],
    index=["Yes", "No"].index(camera_current),
    horizontal=True,
    key="camera_system_selection"
)

if camera_selection == "Yes":
    for part in CAMERA_PARTS:
        if part in optional_lookup:
            st.session_state.accessory_qty[part] = 1

else:
    for part in CAMERA_PARTS:
        st.session_state.accessory_qty.pop(part, None)

# ============================================================
# OTHER OPTIONAL ACCESSORIES
# ============================================================

# st.subheader("Other Optional Accessories")

OTHER_OPTIONAL_PARTS = [
    ("801223664", "3.3 Rotating Keyboard Tray"),
    ("801075237", "3.4 Cable Manager"),
    ("801029022", "3.5 Top Cable Tray"),
    ("801075235", "3.6 Brush Panel"),
]

for part, display_name in OTHER_OPTIONAL_PARTS:

    if part not in optional_lookup:
        continue

    r = optional_lookup[part]

    col1, col2 = st.columns(
        [5.5, 1.8],
        vertical_alignment="center"
    )

    with col1:

        selected = st.checkbox(
            display_name,
            value=(
                st.session_state.accessory_qty.get(
                    part, 0
                ) > 0
            ),
            key=f"other_acc_{part}"
        )

    with col2:

        if selected:

            qty = st.number_input(
                "Quantity",
                min_value=1,
                max_value=999,
                step=1,
                value=int(
                    st.session_state.accessory_qty.get(
                        part,
                        1
                    )
                ),
                key=f"other_qty_{part}"
            )

            st.session_state.accessory_qty[part] = qty

        else:

            st.session_state.accessory_qty.pop(
                part,
                None
            )

# # ------------------------------------------------------------
# # 4 PDU selection
# # ------------------------------------------------------------
# st.header("4. PDU Selection")

# pdu_options = [
#     f'{r["Part Code"]} — {r["Description"]}'
#     for _, r in pdus_df.iterrows()
# ]

# col1, col2 = st.columns([5, 1.5])

# with col1:
#     selected_pdu = st.selectbox(
#         "Select PDU",
#         ["None"] + pdu_options,
#         index=0
#     )

# with col2:
#     qty = st.number_input(
#         "Quantity",
#         min_value=1,
#         max_value=999,
#         value=1,
#         step=1,
#         key="pdu_quantity"
#     )

# # Reset PDU quantity dictionary
# st.session_state.pdu_qty = {}

# if selected_pdu != "None":

#     selected_index = pdu_options.index(selected_pdu)
#     selected_row = pdus_df.iloc[selected_index]

#     part = str(selected_row["Part Code"])

#     st.session_state.pdu_qty[part] = qty

# ------------------------------------------------------------
# 5 Final structure - common to both users
# ------------------------------------------------------------
# st.header("5. Final Structure")

# bom = build_bom()

# if not bom.empty:
#     structure = bom[[
#         "S.No.", "Component Type", "Part Code", "Description", "Quantity", "UOM"
#     ]].copy()
#     st.dataframe(structure, use_container_width=True, hide_index=True)
# else:
#     st.info("No components selected.")
# ============================================================
# 5. FINAL STRUCTURE
# ============================================================
st.html("""
<div style="
    background: linear-gradient(135deg, #005EB8, #003B71);
    color: white;
    padding: 10px 16px;
    border-radius: 8px;
    margin: 20px 0 15px 0;
    font-size: 18px;
    font-weight: 700;
">
    5. FINAL BOQ
</div>
""")

bom = build_bom()

if not bom.empty:

    # Calculate selling prices for Final BOQ
    base_cost, optional_cost, pdu_cost, total_cost = cost_summary(bom)

    bom_with_price, _, _ = add_selling_prices(
        bom,
        total_cost,
        st.session_state.margin_pct,
        st.session_state.freight,
        st.session_state.installation
    )

    structure = bom[
        ["S.No.", "Part Code", "Description", "Quantity", "UOM", "Unit Price",
        "Total Price"]
    ].copy()

    # ========================================================
    # SPECIAL PART CODES
    # ========================================================

    MAIN_MDC_PART = "801029209"

    selected_config_components = selected_components()

    cooling_part_codes = set()

    if not selected_config_components.empty:
         cooling_rows = selected_config_components.tail(3)

         cooling_part_codes = set(
              cooling_rows["Part Code"]
              .dropna()
              .astype(str)
              .str.strip()
         )

    # ========================================================
    # CREATE NEW SERIAL NUMBERS
    # ========================================================

    new_serial = []

    main_mdc_found = False
    mdc_sub_no = 0

    cooling_started = False
    cooling_sub_no = 0

    accessories_started = False
    accessory_no = 3

    pdu_started = False
    pdu_no = 0

    for row_index, row in structure.iterrows():

        part_code = str(row["Part Code"]).strip()
        description = str(row["Description"]).strip()
        component_type = str(
            bom.loc[row.name, "Component Type"]
        ).strip()
        unit_price = row["Unit Price"]
        total_price = row["Total Price"]

        unit_price_display = (
            money(float(unit_price))
            if pd.notna(unit_price)
            else "N/A"
        )

        total_price_display = (
            money(float(total_price))
            if pd.notna(total_price)
            else "N/A"
        )

        # ----------------------------------------------------
        # MAIN MDC TITLE
        # ----------------------------------------------------

        if (
            not main_mdc_found
            and "SINGLE RACK MDC" in description.upper()
        ):
            new_serial.append("")
            main_mdc_found = True
            continue

        # ----------------------------------------------------
        # MAIN MDC
        # 801029209 -> 1
        # ----------------------------------------------------

        if part_code == MAIN_MDC_PART:
            new_serial.append("1")
            continue

        # ----------------------------------------------------
        # COOLING UNIT
        # 801401725 -> 2.1
        # 801401726 -> 2.2
        # 801401745 -> 2.3
        # ----------------------------------------------------

        if part_code in cooling_part_codes:

            cooling_started = True
            cooling_sub_no += 1

            new_serial.append(
                f"2.{cooling_sub_no}"
            )

            continue

        # ----------------------------------------------------
        # OPTIONAL ACCESSORIES
        # 3, 4, 5, 6...
        # ----------------------------------------------------

        if component_type == "Optional Accessory":

            accessories_started = True

            new_serial.append(
                str(accessory_no)
            )

            accessory_no += 1

            continue

        # ----------------------------------------------------
        # PDU
        # Continue numbering after accessories
        # ----------------------------------------------------

        if component_type == "PDU":

            pdu_started = True

            # If accessories exist:
            # continue from accessory numbering.
            #
            # Example:
            # Accessories = 3, 4
            # PDU = 5

            new_serial.append(
                str(accessory_no)
            )

            accessory_no += 1

            continue

        # ----------------------------------------------------
        # MDC COMPONENTS
        # 1.1, 1.2, 1.3 ... 1.17
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

    structure["New S.No."] = new_serial

    # ========================================================
    # HTML TABLE CSS
    # ========================================================

    html = """
    <style>

    .final-structure-table {
        width: 100%;
        border-collapse: collapse;
        font-family: Arial, sans-serif;
        font-size: 14px;
        border: 1px solid #D9E1E8;
        border-radius: 8px;
        overflow: hidden;
    }

    .final-structure-table th {
        background-color: #F4F6F8;
        color: #555555;
        font-weight: 600;
        text-align: left;
        padding: 12px 10px;
        border-bottom: 1px solid #D9E1E8;
    }

    .final-structure-table td {
        padding: 11px 10px;
        border-bottom: 1px solid #E5E7EB;
        color: #333333;
        vertical-align: middle;
    }

    /* ======================================================
       MAIN MDC TITLE
       Dark Blue + Centered
       ====================================================== */

    .main-mdc-row td {
        background-color: #003B71;
        color: white !important;
        font-weight: 700;
        font-size: 16px;
        text-align: center !important;
        padding: 15px 10px;
    }

    /* ======================================================
       SECTION HEADINGS
       Eaton Blue + Centered
       ====================================================== */

    .section-heading td {
        background-color: #005EB8;
        color: white !important;
        font-weight: 700;
        font-size: 15px;
        text-align: center !important;
        padding: 12px 14px;
    }

    /* ======================================================
       COLUMN ALIGNMENT
       ====================================================== */

    .serial {
        width: 7%;
        text-align: center !important;
    }

    .part-code {
        width: 15%;
    }

    .description {
        width: 58%;
    }

    .quantity {
        width: 10%;
        text-align: center !important;
    }

    .uom {
        width: 10%;
        text-align: center !important;
    }
    .unit-price {
        width: 12%;
        text-align: right !important;
        white-space: nowrap;
    }

    .total-price {
        width: 13%;
        text-align: right !important;
        white-space: nowrap;
    }

    </style>

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

    # ========================================================
    # SECTION FLAGS
    # ========================================================

    cooling_heading_added = False
    accessories_heading_added = False
    pdu_heading_added = False

    # ========================================================
    # ADD TABLE ROWS
    # ========================================================

    for row_index, row in structure.iterrows():

        part_code = str(row["Part Code"]).strip()
        description = str(row["Description"]).strip()
        quantity = str(row["Quantity"]).strip()
        uom = str(row["UOM"]).strip()
        serial_no = str(row["New S.No."]).strip()

        component_type = str(
            bom.loc[row.name, "Component Type"]
        ).strip()

        # ----------------------------------------------------
        # MAIN MDC TITLE
        # ----------------------------------------------------

        if (
            serial_no == ""
            and "SINGLE RACK MDC" in description.upper()
        ):

            html += f"""
            <tr class="main-mdc-row">
                <td colspan="7">
                    {description}
                </td>
            </tr>
            """

            continue

        # ----------------------------------------------------
        # COOLING UNIT HEADING
        # ----------------------------------------------------

        if (
            part_code in cooling_part_codes
            and not cooling_heading_added
        ):

            html += """
            <tr class="section-heading">
                <td colspan="5">
                    COOLING UNIT
                </td>
            </tr>
            """

            cooling_heading_added = True
        

        # ----------------------------------------------------
        # OTHER ACCESSORIES HEADING
        # ----------------------------------------------------

        if (
            component_type == "Optional Accessory"
            and not accessories_heading_added
        ):

            html += """
            <tr class="section-heading">
                <td colspan="5">
                    OTHER ACCESSORIES
                </td>
            </tr>
            """

            accessories_heading_added = True

        # ----------------------------------------------------
        # PDU HEADING
        # ----------------------------------------------------

        if (
            component_type == "PDU"
            and not pdu_heading_added
        ):

            html += """
            <tr class="section-heading">
                <td colspan="5">
                    PDU
                </td>
            </tr>
            """

            pdu_heading_added = True

        # ----------------------------------------------------
        # PART CODE
        # ----------------------------------------------------

        display_part_code = (
            ""
            if part_code.lower() == "nan"
            else part_code
        )

        # ----------------------------------------------------
        # NORMAL ROW
        # ----------------------------------------------------

        html += f"""
        <tr>
            <td class="serial">{serial_no}</td>
            <td class="part-code">{display_part_code}</td>
            <td class="description">{description}</td>
            <td class="quantity">{quantity}</td>
            <td class="uom">{uom}</td>
            <td class="unit-price">{unit_price_display}</td>
            <td class="total-price">{total_price_display}</td>
        </tr>
        """

    html += """
        </tbody>
    </table>
    """

    # ========================================================
    # DISPLAY HTML TABLE
    # ========================================================

    st.html(html)

else:

    st.info("No components selected.")

# ------------------------------------------------------------
# Cost + selling price - internal only
# ------------------------------------------------------------
base_cost, optional_cost, pdu_cost, total_cost = cost_summary(bom)

margin_pct = st.session_state.margin_pct
freight = st.session_state.freight
installation = st.session_state.installation
warranty_pct = st.session_state.warranty_pct

if is_internal:
    st.header("6. Cost Summary — Internal Only")

    a, b, c, d = st.columns(4)
    with a:
        price_box("Base Cost", base_cost)
    with b:
        price_box("Optional Cost", optional_cost)
    with c:
        price_box("PDU Cost", pdu_cost)
    with d:
        price_box("Total Cost", total_cost)

    st.header("7. Cost to Selling Price — Internal Only")

    p1, p2, p3, p4 = st.columns(4)
    with p1:
        margin_pct = st.number_input(
            "Margin (%)", 0.0, 99.0,
            st.session_state.margin_pct, 0.5
        )
        st.session_state.margin_pct = margin_pct
    with p2:
        freight = st.number_input(
            "Freight", 0.0,
            value=st.session_state.freight, step=500.0
        )
        st.session_state.freight = freight
    with p3:
        installation = st.number_input(
            "Installation", 0.0,
            value=st.session_state.installation, step=500.0
        )
        st.session_state.installation = installation
    with p4:
        warranty_pct = st.number_input(
            "Warranty (%)", 0.0, 100.0,
            st.session_state.warranty_pct, 0.5
        )
        st.session_state.warranty_pct = warranty_pct

    margin_price = total_cost / (1 - margin_pct / 100) if margin_pct < 100 else 0
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

# ------------------------------------------------------------
# 8 Final BOM
# ------------------------------------------------------------
# st.header("8. Final BOM")

# if not bom.empty:
#     bom_with_price, margin_price, final_selling_price = add_selling_prices(
#         bom, total_cost, margin_pct, freight, installation
#     )

#     display = bom_with_price[[
#         "S.No.", "Component Type", "Part Code", "Description", "Quantity",
#         "UOM", "Unit Price", "Total Price"
#     ]].copy()

#     display["Unit Price"] = display["Unit Price"].apply(
#         lambda x: money(float(x)) if pd.notna(x) else "N/A"
#     )
#     display["Total Price"] = display["Total Price"].apply(
#         lambda x: money(float(x)) if pd.notna(x) else "N/A"
#     )

#     st.dataframe(display, use_container_width=True, hide_index=True)

#     known = bom_with_price["Total Price"].dropna().sum()
#     price_box("BOM Selling Value", float(known))

#     if not is_internal:
#         st.caption("Sales view contains selling prices only. Internal unit cost and total cost are not displayed.")
# else:
#     bom_with_price = bom
#     st.info("No BOM available.")

# ------------------------------------------------------------
# 9 Excel
# ------------------------------------------------------------
st.header("9. Excel Download")

if not bom.empty:
    internal_cost_data = [
        ["Base Cost", base_cost],
        ["Optional Cost", optional_cost],
        ["PDU Cost", pdu_cost],
        ["Total Cost", total_cost],
        ["Margin %", margin_pct],
        ["Margin Price", margin_price if is_internal else 0],
        ["Freight", freight if is_internal else 0],
        ["Installation", installation if is_internal else 0],
        ["Final Selling Price", final_selling_price if is_internal else 0],
        ["Warranty %", warranty_pct if is_internal else 0],
        ["Warranty Amount", (margin_price * warranty_pct / 100) if is_internal else 0],
    ]

    # Sales Excel is always safe for both roles.
    sales_file = excel_bytes(
        internal=False,
        bom=bom_with_price,
        cost_data=None,
    )
    customer_name_valid = bool(
    st.session_state.customer_name.strip()
)

    if is_internal and customer_name_valid:
        st.success("Internal MDC user: both Excel versions are available.")

        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                "⬇️ Download Internal Cost Excel",
                data=excel_bytes(
                    internal=True,
                    bom=bom_with_price,
                    cost_data=internal_cost_data,
                ),
                file_name="MDC_Internal_Cost.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                on_click=handle_excel_download,
            )
        with col2:
            st.download_button(
                "⬇️ Download Sales Excel",
                data=sales_file,
                file_name="MDC_Sales_Output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                on_click=handle_excel_download,
            )
    else:
        st.download_button(
            "⬇️ Download Sales Excel",
            data=sales_file,
            file_name="MDC_Sales_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            on_click=handle_excel_download,
        )
        
# ============================================================
# 10. SAVE CONFIGURATION
# ============================================================

st.header("10. Save Configuration")

st.caption(
    "Save the current MDC configuration for future tracking and reference."
)

save_col1, save_col2 = st.columns([2, 5])

with save_col1:

    if st.button(
        "💾 Save Configuration",
        use_container_width=True,
        type="primary"
    ):

        # Calculate selling price
        current_margin_price = (
            total_cost / (1 - margin_pct / 100)
            if margin_pct < 100
            else 0
        )

        current_final_price = (
            current_margin_price
            + freight
            + installation
        )

        current_warranty_amount = (
            current_margin_price
            * warranty_pct
            / 100
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
            f"Configuration {st.session_state.configuration_id} saved successfully."
        )
st.divider()
st.caption(
    "MDC Solution V1 | Single Rack data loaded from the supplied 01.09.2026 BOQ | "
    "Multirack configurations are XXX placeholders for future updates."
)
# ============================================================
# 11. CONFIGURATION HISTORY
# INTERNAL USERS ONLY
# ============================================================

if is_internal:

    st.header("11. Configuration History")

    conn = sqlite3.connect(TRACKING_DB)

    history_df = pd.read_sql_query(
        """
        SELECT
            configuration_id AS "Configuration ID",
            created_at AS "Created At",
            customer_name AS "Customer",
            customer_place AS "Place",
            mdc_type AS "MDC Type",
            configuration AS "Configuration",
            final_selling_price AS "Final Selling Price"
        FROM configurations
        ORDER BY id DESC
        """,
        conn
    )

    conn.close()

    if not history_df.empty:

        history_df["Final Selling Price"] = (
            history_df["Final Selling Price"]
            .apply(money)
        )

        st.dataframe(
            history_df,
            use_container_width=True,
            hide_index=True
        )

    else:

        st.info(
            "No saved configurations available yet."
        )
