import streamlit as st
import pandas as pd
from PIL import Image
import plotly.express as px
from datetime import datetime

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(page_title="RSC Sales Dashboard", layout="wide")

# ─────────────────────────────────────────────
# LOAD DATA
# ─────────────────────────────────────────────
@st.cache_data(show_spinner="Loading data...")
def load_data():
    return pd.read_excel(
        "MOM RSC Performance_Jan'24 To Dec'25- North  South_Region V1.xlsb",
        sheet_name="RAW data",
        skiprows=1,
        engine="pyxlsb"
    )

df = load_data()

# ─────────────────────────────────────────────
# CLEAN COLUMN NAMES
# ─────────────────────────────────────────────
df.columns = df.columns.astype(str).str.strip()

# ─────────────────────────────────────────────
# SAFE COLUMN FINDER
# ─────────────────────────────────────────────
def find_col(possible_names):
    for col in df.columns:
        for name in possible_names:
            if col.lower().strip() == name.lower().strip():
                return col
    return None

REGION_COL = find_col(["Region"])
RM_COL = find_col(["RM's Territory", "RM Territory", "RMs Territory"])
FOM_COL = find_col([
    "Field Op Manager",
    "Field Operation Manager",
    "Field Ops Manager",
    "FOM"
])

FTD_PIXMA_COL = find_col(["FTD PIXMA Zone"])
FTD_MBO_COL = find_col(["FTD MBO"])
MTD_PIXMA_COL = find_col(["MTD PIXMA Zone"])
MTD_MBO_COL = find_col(["MTD MBO"])

# ─────────────────────────────────────────────
# FAIL FAST IF REQUIRED COLS MISSING
# ─────────────────────────────────────────────
required_cols = {
    "Region": REGION_COL,
    "RM Territory": RM_COL,
    "Field Op Manager": FOM_COL,
    "FTD PIXMA": FTD_PIXMA_COL,
    "FTD MBO": FTD_MBO_COL,
    "MTD PIXMA": MTD_PIXMA_COL,
    "MTD MBO": MTD_MBO_COL,
}

missing = [k for k, v in required_cols.items() if v is None]

if missing:
    st.error(f"❌ Missing required columns: {', '.join(missing)}")
    st.stop()

# ─────────────────────────────────────────────
# DATE HANDLING
# ─────────────────────────────────────────────
possible_date_cols = [
    "Refer Date", "ReferDate", "Reference Date",
    "Ref Date", "Invoice Date", "Date"
]

DATE_COL = next((c for c in possible_date_cols if c in df.columns), None)

if DATE_COL is None:
    st.error("❌ Date column not found")
    st.stop()

df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors="coerce")

df["Year"] = df[DATE_COL].dt.year
df["Month_No"] = df[DATE_COL].dt.month
df["Month_Name"] = df[DATE_COL].dt.strftime("%b")

df = df[df["Year"].between(2024, 2025)]

# ─────────────────────────────────────────────
# SIDEBAR FILTERS
# ─────────────────────────────────────────────
st.sidebar.title("🔍 Filters")

region_list = sorted(df[REGION_COL].dropna().unique())
selected_region = st.sidebar.multiselect(
    "Region", region_list, default=region_list
)

# ─────────────────────────────────────────────
# APPLY FILTERS
# ─────────────────────────────────────────────
df_filtered = df[df["Status"] == "Passed"]
df_filtered = df_filtered[df_filtered[REGION_COL].isin(selected_region)]

# ─────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────
@st.cache_resource
def load_logo():
    return Image.open("canon-press-centre-canon-logo.png")

col1, col2 = st.columns([0.15, 0.85])

with col1:
    st.image(load_logo(), width=140)

with col2:
    st.markdown(
        "<h1 style='margin-bottom:0'>RSC Sales Performance Dashboard</h1>",
        unsafe_allow_html=True
    )

st.markdown(f"**Last Updated:** {datetime.now().strftime('%d %B %Y')}")
st.markdown(f"**Selected Region:** {', '.join(selected_region)}")

# ─────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────
tab1, tab2 = st.tabs(["📊 Dashboard", "📄 Region Summary"])

# =====================================================
# TAB 1 – DASHBOARD
# =====================================================
with tab1:
    month_qty = (
        df_filtered
        .groupby(["Month_No", "Month_Name"], as_index=False)["Sales Quantity"]
        .sum()
        .sort_values("Month_No")
    )

    st.plotly_chart(
        px.bar(
            month_qty,
            x="Month_Name",
            y="Sales Quantity",
            text="Sales Quantity",
            title="Month-wise Sales Trend"
        ).update_layout(xaxis_tickangle=-30),
        use_container_width=True
    )

# =====================================================
# TAB 2 – REGION SUMMARY (IMAGE STYLE)
# =====================================================
with tab2:
    st.subheader("Region-wise Performance Summary")

    summary = (
        df_filtered
        .groupby([REGION_COL, RM_COL, FOM_COL], as_index=False)
        .agg(
            **{
                "Retail Sales Consultant Count": ("Name", "nunique"),
                "FTD PIXMA Zone": (FTD_PIXMA_COL, "sum"),
                "FTD MBO": (FTD_MBO_COL, "sum"),
                "MTD PIXMA Zone": (MTD_PIXMA_COL, "sum"),
                "MTD MBO": (MTD_MBO_COL, "sum"),
            }
        )
    )

    summary["FTD Total"] = summary["FTD PIXMA Zone"] + summary["FTD MBO"]
    summary["MTD Total"] = summary["MTD PIXMA Zone"] + summary["MTD MBO"]

    summary = summary.rename(columns={
        REGION_COL: "Region",
        RM_COL: "RM's Territory",
        FOM_COL: "Field Op Manager"
    })

    # REGION TOTAL
    region_total = summary.groupby("Region", as_index=False).sum(numeric_only=True)
    region_total["RM's Territory"] = ""
    region_total["Field Op Manager"] = "Total"

    # GRAND TOTAL
    grand_total = summary.sum(numeric_only=True).to_frame().T
    grand_total["Region"] = ""
    grand_total["RM's Territory"] = ""
    grand_total["Field Op Manager"] = "Grand Total"

    final_df = pd.concat([summary, region_total, grand_total], ignore_index=True)

    final_df = final_df[
        [
            "Region",
            "RM's Territory",
            "Field Op Manager",
            "Retail Sales Consultant Count",
            "FTD PIXMA Zone",
            "FTD MBO",
            "FTD Total",
            "MTD PIXMA Zone",
            "MTD MBO",
            "MTD Total",
        ]
    ]

    st.dataframe(final_df, use_container_width=True)
