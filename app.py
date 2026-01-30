import streamlit as st
import pandas as pd
from PIL import Image
import plotly.express as px
from datetime import datetime

# ─────────────────────────────────────────────
# Page configuration
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="RSC Sales Dashboard",
    layout="wide"
)

# ─────────────────────────────────────────────
# FAST DATA LOADING (FILE IN ROOT)
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
# AUTO-DETECT REFER DATE COLUMN
# ─────────────────────────────────────────────
possible_date_cols = [
    "Refer Date",
    "ReferDate",
    "Reference Date",
    "Ref Date",
    "Invoice Date",
    "Date"
]

DATE_COL = next((c for c in possible_date_cols if c in df.columns), None)

if DATE_COL is None:
    st.error("❌ Refer Date column not found")
    st.write("Available columns:", df.columns.tolist())
    st.stop()

# ─────────────────────────────────────────────
# EXCEL SAFE DATE CONVERSION
# ─────────────────────────────────────────────
if pd.api.types.is_numeric_dtype(df[DATE_COL]):
    df[DATE_COL] = pd.to_datetime(
        df[DATE_COL],
        unit="D",
        origin="1899-12-30",
        errors="coerce"
    )
else:
    df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors="coerce")

# ─────────────────────────────────────────────
# YEAR & MONTH CREATION
# ─────────────────────────────────────────────
df["Year"] = df[DATE_COL].dt.year
df["Month_No"] = df[DATE_COL].dt.month
df["Month_Name"] = df[DATE_COL].dt.strftime("%b")

df = df[df["Year"].between(2024, 2025)]

# ─────────────────────────────────────────────
# Styling
# ─────────────────────────────────────────────
st.markdown(
    "<style>div.block-container{padding-top:1rem;}</style>",
    unsafe_allow_html=True
)

# ─────────────────────────────────────────────
# SIDEBAR FILTERS
# ─────────────────────────────────────────────
st.sidebar.title("🔍 Filters")

selected_year = st.sidebar.multiselect(
    "Year",
    sorted(df["Year"].dropna().unique()),
    default=sorted(df["Year"].dropna().unique())
)

selected_city = st.sidebar.multiselect(
    "City",
    sorted(df["City"].dropna().unique()),
    default=sorted(df["City"].dropna().unique())
)

selected_store = st.sidebar.multiselect(
    "Store Name",
    sorted(df["Storename"].dropna().unique()),
    default=sorted(df["Storename"].dropna().unique())
)

selected_name = st.sidebar.multiselect(
    "Name",
    sorted(df["Name"].dropna().unique()),
    default=sorted(df["Name"].dropna().unique())
)

# ─────────────────────────────────────────────
# APPLY FILTERS
# ─────────────────────────────────────────────
df_filtered = df[df["Status"] == "Passed"]

if selected_year:
    df_filtered = df_filtered[df_filtered["Year"].isin(selected_year)]

if selected_city:
    df_filtered = df_filtered[df_filtered["City"].isin(selected_city)]

if selected_store:
    df_filtered = df_filtered[df_filtered["Storename"].isin(selected_store)]

if selected_name:
    df_filtered = df_filtered[df_filtered["Name"].isin(selected_name)]

# ─────────────────────────────────────────────
# LOAD LOGO
# ─────────────────────────────────────────────
@st.cache_resource
def load_logo():
    return Image.open("canon-press-centre-canon-logo.png")

logo = load_logo()

# ─────────────────────────────────────────────
# Header
# ─────────────────────────────────────────────
col1, col2 = st.columns([0.15, 0.85])

with col1:
    st.image(logo, width=140)

with col2:
    st.markdown(
        """
        <style>
        .title-test {
            font-weight: bold;
            font-size: 34px;
            padding-top: 15px;
        }
        </style>
        <div class="title-test">RSC Sales Performance Dashboard</div>
        """,
        unsafe_allow_html=True
    )

# ─────────────────────────────────────────────
# Last Updated
# ─────────────────────────────────────────────
st.markdown(f"**Last Updated:** {datetime.now().strftime('%d %B %Y')}")

# ─────────────────────────────────────────────
# MONTH-WISE SALES TREND
# ─────────────────────────────────────────────
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
        title="Month-wise Sales Trend (Quantity – Passed Only)"
    ).update_traces(textposition="inside")
     .update_layout(xaxis_tickangle=-30),
    use_container_width=True
)

# ─────────────────────────────────────────────
# TOP 5 PRODUCT CATEGORIES
# ─────────────────────────────────────────────
colA, colB = st.columns(2)

with colA:
    cat_qty = (
        df_filtered.groupby("Product Category", as_index=False)["Sales Quantity"]
        .sum().sort_values("Sales Quantity", ascending=False).head(5)
    )
    st.plotly_chart(
        px.bar(cat_qty, x="Product Category", y="Sales Quantity",
               text="Sales Quantity",
               title="Top 5 Product Categories – Quantity"),
        use_container_width=True
    )

with colB:
    cat_val = (
        df_filtered.groupby("Product Category", as_index=False)["Sales Value"]
        .sum().sort_values("Sales Value", ascending=False).head(5)
    )
    st.plotly_chart(
        px.bar(cat_val, x="Product Category", y="Sales Value",
               text="Sales Value",
               title="Top 5 Product Categories – Value"),
        use_container_width=True
    )

# ─────────────────────────────────────────────
# TOP PRODUCTS & STORES
# ─────────────────────────────────────────────
colC, colD = st.columns(2)

with colC:
    top_products = (
        df_filtered.groupby("Model Name", as_index=False)["Sales Quantity"]
        .sum().sort_values("Sales Quantity", ascending=False).head(5)
    )
    st.plotly_chart(
        px.bar(top_products, x="Model Name", y="Sales Quantity",
               text="Sales Quantity",
               title="Top 5 Best Seller Products"),
        use_container_width=True
    )

with colD:
    top_stores = (
        df_filtered.groupby("Storename", as_index=False)["Sales Quantity"]
        .sum().sort_values("Sales Quantity", ascending=False).head(5)
    )
    st.plotly_chart(
        px.bar(top_stores, x="Storename", y="Sales Quantity",
               text="Sales Quantity",
               title="Top 5 Stores"),
        use_container_width=True
    )

# ─────────────────────────────────────────────
# LEADERSHIP BOARD – TOP 10 SELLERS
# ─────────────────────────────────────────────
leaderboard = (
    df_filtered
    .groupby("Name", as_index=False)["Sales Quantity"]
    .sum()
    .sort_values("Sales Quantity", ascending=False)
    .head(10)
)

leaderboard["Rank"] = range(1, len(leaderboard) + 1)

st.plotly_chart(
    px.bar(
        leaderboard,
        x="Sales Quantity",
        y="Name",
        orientation="h",
        text="Sales Quantity",
        title="🏆 Top 10 Sellers – Leadership Board"
    ).update_layout(
        yaxis=dict(autorange="reversed"),
        xaxis_title="Sales Quantity",
        yaxis_title="Seller Name"
    ),
    use_container_width=True
)
