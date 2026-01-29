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
# FAST DATA LOADING
# ─────────────────────────────────────────────
@st.cache_data(show_spinner="Loading data...")
def load_data():
    return pd.read_excel(
        r"D:\Main Working Files\MOM RSC Performance_Jan'24 To Dec'25- North  South_Region V1.xlsb",
        sheet_name="RAW data",
        skiprows=1
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
# YEAR & MONTH CREATION (ORDER SAFE)
# ─────────────────────────────────────────────
df["Year"] = df[DATE_COL].dt.year
df["Month_No"] = df[DATE_COL].dt.month
df["Month_Name"] = df[DATE_COL].dt.strftime("%b")  # Jan, Feb, Mar...

# Keep only valid data years
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

# Year filter
year_list = sorted(df["Year"].dropna().unique())
selected_year = st.sidebar.multiselect(
    "Year",
    year_list,
    default=year_list
)

# City filter
city_list = sorted(df["City"].dropna().unique())
selected_city = st.sidebar.multiselect(
    "City",
    city_list,
    default=city_list
)

# Store filter
store_list = sorted(df["Storename"].dropna().unique())
selected_store = st.sidebar.multiselect(
    "Store Name",
    store_list,
    default=store_list
)

# ─────────────────────────────────────────────
# APPLY FILTERS (PASSED ONLY)
# ─────────────────────────────────────────────
df_filtered = df.copy()

df_filtered = df_filtered[df_filtered["Status"] == "Passed"]

if selected_year:
    df_filtered = df_filtered[df_filtered["Year"].isin(selected_year)]

if selected_city:
    df_filtered = df_filtered[df_filtered["City"].isin(selected_city)]

if selected_store:
    df_filtered = df_filtered[df_filtered["Storename"].isin(selected_store)]

# ─────────────────────────────────────────────
# Load Logo
# ─────────────────────────────────────────────
@st.cache_resource
def load_logo():
    return Image.open(
        r"C:\Users\ce-vipin.kp\Downloads\canon-press-centre-canon-logo.png"
    )

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
# MONTH-WISE SALES TREND (ORDERED)
# ─────────────────────────────────────────────
month_qty = (
    df_filtered
    .groupby(["Month_No", "Month_Name"], as_index=False)["Sales Quantity"]
    .sum()
    .sort_values("Month_No")
)

fig_month = px.bar(
    month_qty,
    x="Month_Name",
    y="Sales Quantity",
    text="Sales Quantity",
    title="Month-wise Sales Trend (Quantity – Passed Only)"
)

fig_month.update_traces(textposition="inside")
fig_month.update_layout(xaxis_tickangle=-30)

st.plotly_chart(fig_month, use_container_width=True)

# ─────────────────────────────────────────────
# TOP 5 PRODUCT CATEGORIES – QTY & VALUE
# ─────────────────────────────────────────────
colA, colB = st.columns(2)

with colA:
    cat_qty = (
        df_filtered.groupby("Product Category", as_index=False)["Sales Quantity"]
        .sum()
        .sort_values("Sales Quantity", ascending=False)
        .head(5)
    )

    fig_cat_qty = px.bar(
        cat_qty,
        x="Product Category",
        y="Sales Quantity",
        text="Sales Quantity",
        title="Top 5 Product Categories – Quantity"
    )
    fig_cat_qty.update_traces(textposition="inside")
    fig_cat_qty.update_layout(xaxis_tickangle=-30)
    st.plotly_chart(fig_cat_qty, use_container_width=True)

with colB:
    cat_val = (
        df_filtered.groupby("Product Category", as_index=False)["Sales Value"]
        .sum()
        .sort_values("Sales Value", ascending=False)
        .head(5)
    )

    fig_cat_val = px.bar(
        cat_val,
        x="Product Category",
        y="Sales Value",
        text="Sales Value",
        title="Top 5 Product Categories – Value"
    )
    fig_cat_val.update_traces(textposition="inside")
    fig_cat_val.update_layout(xaxis_tickangle=-30)
    st.plotly_chart(fig_cat_val, use_container_width=True)

# ─────────────────────────────────────────────
# TOP 5 PRODUCTS & STORES
# ─────────────────────────────────────────────
st.markdown("### 🏆 Top Performers")

colC, colD = st.columns(2)

with colC:
    top5_products = (
        df_filtered.groupby("Model Name", as_index=False)["Sales Quantity"]
        .sum()
        .sort_values("Sales Quantity", ascending=False)
        .head(5)
    )

    fig_prod = px.bar(
        top5_products,
        x="Model Name",
        y="Sales Quantity",
        text="Sales Quantity",
        title="Top 5 Best Seller Products"
    )
    fig_prod.update_traces(textposition="inside")
    fig_prod.update_layout(xaxis_tickangle=-30)
    st.plotly_chart(fig_prod, use_container_width=True)

with colD:
    top5_stores = (
        df_filtered.groupby("Storename", as_index=False)["Sales Quantity"]
        .sum()
        .sort_values("Sales Quantity", ascending=False)
        .head(5)
    )

    fig_store = px.bar(
        top5_stores,
        x="Storename",
        y="Sales Quantity",
        text="Sales Quantity",
        title="Top 5 Stores"
    )
    fig_store.update_traces(textposition="inside")
    fig_store.update_layout(xaxis_tickangle=-30)
    st.plotly_chart(fig_store, use_container_width=True)

# ─────────────────────────────────────────────
# RUNNING (CUMULATIVE) LINE – TOP 10 SELLERS
# ─────────────────────────────────────────────
st.markdown("### 📈 Running Sales Contribution – Top 10 Sellers")

top10_sellers = (
    df_filtered
    .groupby("Name", as_index=False)["Sales Quantity"]
    .sum()
    .sort_values("Sales Quantity", ascending=False)
    .head(10)
)

top10_sellers = top10_sellers.sort_values("Sales Quantity")
top10_sellers["Running Quantity"] = top10_sellers["Sales Quantity"].cumsum()

fig_running = px.line(
    top10_sellers,
    x="Name",
    y="Running Quantity",
    markers=True,
    title="Running (Cumulative) Sales Quantity – Top 10 Sellers"
)

fig_running.update_layout(xaxis_tickangle=-30)

st.plotly_chart(fig_running, use_container_width=True)
