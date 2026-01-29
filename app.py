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
# FILE UPLOAD (REQUIRED FOR CLOUD)
# ─────────────────────────────────────────────
uploaded_file = st.file_uploader(
    "📂 Upload MOM RSC Performance File (.xlsb)",
    type=["xlsb", "xlsx"]
)

if uploaded_file is None:
    st.warning("👆 Please upload the sales file to continue")
    st.stop()

# ─────────────────────────────────────────────
# FAST DATA LOADING
# ─────────────────────────────────────────────
@st.cache_data(show_spinner="Loading data...")
def load_data(file):
    return pd.read_excel(
        file,
        sheet_name="RAW data",
        skiprows=1,
        engine="pyxlsb" if file.name.endswith(".xlsb") else None
    )

df = load_data(uploaded_file)

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

# ─────────────────────────────────────────────
# APPLY FILTERS (PASSED ONLY)
# ─────────────────────────────────────────────
df_filtered = df[df["Status"] == "Passed"]

if selected_year:
    df_filtered = df_filtered[df_filtered["Year"].isin(selected_year)]

if selected_city:
    df_filtered = df_filtered[df_filtered["City"].isin(selected_city)]

if selected_store:
    df_filtered = df_filtered[df_filtered["Storename"].isin(selected_store)]

# ─────────────────────────────────────────────
# LOGO UPLOAD (OPTIONAL)
# ─────────────────────────────────────────────
logo_file = st.sidebar.file_uploader(
    "Upload Logo (optional)",
    type=["png", "jpg", "jpeg"]
)

logo = Image.open(logo_file) if logo_file else None

# ─────────────────────────────────────────────
# Header
# ─────────────────────────────────────────────
col1, col2 = st.columns([0.15, 0.85])

with col1:
    if logo:
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
# TOP 5 PRODUCT CATEGORIES
# ─────────────────────────────────────────────
colA, colB = st.columns(2)

with colA:
    cat_qty = (
        df_filtered.groupby("Product Category", as_index=False)["Sales Quantity"]
        .sum()
        .sort_values("Sales Quantity", ascending=False)
        .head(5)
    )

    st.plotly_chart(
        px.bar(cat_qty, x="Product Category", y="Sales Quantity",
               text="Sales Quantity", title="Top 5 Product Categories – Quantity"),
        use_container_width=True
    )

with colB:
    cat_val = (
        df_filtered.groupby("Product Category", as_index=False)["Sales Value"]
        .sum()
        .sort_values("Sales Value", ascending=False)
        .head(5)
    )

    st.plotly_chart(
        px.bar(cat_val, x="Product Category", y="Sales Value",
               text="Sales Value", title="Top 5 Product Categories – Value"),
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
               text="Sales Quantity", title="Top 5 Best Seller Products"),
        use_container_width=True
    )

with colD:
    top_stores = (
        df_filtered.groupby("Storename", as_index=False)["Sales Quantity"]
        .sum().sort_values("Sales Quantity", ascending=False).head(5)
    )

    st.plotly_chart(
        px.bar(top_stores, x="Storename", y="Sales Quantity",
               text="Sales Quantity", title="Top 5 Stores"),
        use_container_width=True
    )

# ─────────────────────────────────────────────
# RUNNING LINE – TOP 10 SELLERS
# ─────────────────────────────────────────────
top10 = (
    df_filtered.groupby("Name", as_index=False)["Sales Quantity"]
    .sum().sort_values("Sales Quantity", ascending=False).head(10)
)

top10 = top10.sort_values("Sales Quantity")
top10["Running Quantity"] = top10["Sales Quantity"].cumsum()

st.plotly_chart(
    px.line(top10, x="Name", y="Running Quantity",
            markers=True,
            title="Running (Cumulative) Sales Quantity – Top 10 Sellers"),
    use_container_width=True
)
