import streamlit as st
import pandas as pd
from PIL import Image
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import os

# ─────────────────────────────────────────────
# Page Configuration
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Retail Sales Performance Dashboard",
    page_icon="🛍️",
    layout="wide"
)

# Custom CSS
st.markdown("""
<style>
.main {background-color: #f8f9fa;}
.metric-card {background-color: white; padding: 15px; border-radius: 10px;}
h1 {color: #c00000 !important;}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# DEBUG
# ─────────────────────────────────────────────
st.sidebar.header("Debug Info")
st.sidebar.write(os.listdir("."))

# ─────────────────────────────────────────────
# DATA LOADING
# ─────────────────────────────────────────────
@st.cache_data(show_spinner="Loading data...")
def load_data():
    file_path = "MOM RSC Performance_Jan'24 To Dec'25- North South_Region V1.xlsb"
    return pd.read_excel(
        file_path,
        sheet_name="RAW data",
        skiprows=1,
        engine="pyxlsb"
    )

df = load_data()
df.columns = df.columns.astype(str).str.strip()

# ─────────────────────────────────────────────
# DATE HANDLING
# ─────────────────────────────────────────────
possible_date_cols = ["Refer Date", "ReferDate", "Reference Date", "Ref Date", "Invoice Date", "Date"]
DATE_COL = next((c for c in possible_date_cols if c in df.columns), None)

if DATE_COL is None:
    st.error("❌ Date column not found")
    st.stop()

if pd.api.types.is_numeric_dtype(df[DATE_COL]):
    df[DATE_COL] = pd.to_datetime(df[DATE_COL], unit="D", origin="1899-12-30", errors="coerce")
else:
    df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors="coerce")

df = df.dropna(subset=[DATE_COL])
df = df[df[DATE_COL].dt.year.between(2024, 2025)]

# ─────────────────────────────────────────────
# SIDEBAR FILTERS
# ─────────────────────────────────────────────
st.sidebar.header("🔍 Dashboard Filters")

min_date = df[DATE_COL].min().date()
max_date = df[DATE_COL].max().date()

date_range = st.sidebar.date_input(
    "Select Date Range",
    [min_date, max_date],
    min_value=min_date,
    max_value=max_date
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

selected_category = st.sidebar.multiselect(
    "Product Category",
    sorted(df["Product Category"].dropna().unique()),
    default=sorted(df["Product Category"].dropna().unique())
)

selected_name = st.sidebar.multiselect(
    "Sales Person",
    sorted(df["Name"].dropna().unique()),
    default=sorted(df["Name"].dropna().unique())
)

# ─────────────────────────────────────────────
# APPLY FILTERS
# ─────────────────────────────────────────────
df_filtered = df[df["Status"] == "Passed"].copy()

if len(date_range) == 2:
    start_date, end_date = date_range
    df_filtered = df_filtered[
        (df_filtered[DATE_COL] >= pd.to_datetime(start_date)) &
        (df_filtered[DATE_COL] <= pd.to_datetime(end_date))
    ]

if selected_city:
    df_filtered = df_filtered[df_filtered["City"].isin(selected_city)]
if selected_store:
    df_filtered = df_filtered[df_filtered["Storename"].isin(selected_store)]
if selected_category:
    df_filtered = df_filtered[df_filtered["Product Category"].isin(selected_category)]
if selected_name:
    df_filtered = df_filtered[df_filtered["Name"].isin(selected_name)]

if df_filtered.empty:
    st.error("No data found for selected filters.")
    st.stop()

# ─────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────
col1, col2 = st.columns([0.15, 0.85])

with col1:
    try:
        st.image(Image.open("canon-press-centre-canon-logo.png"), width=150)
    except:
        st.markdown("**Canon**")

with col2:
    st.markdown("<h1>Retail Sales Performance Dashboard</h1>", unsafe_allow_html=True)
    st.caption(f"Last Updated: {datetime.now().strftime('%d %B %Y, %I:%M %p')}")

st.divider()

# ─────────────────────────────────────────────
# KPI CARDS
# ─────────────────────────────────────────────
total_qty = df_filtered["Sales Quantity"].sum()
total_value = df_filtered["Sales Value"].sum()
total_orders = len(df_filtered)
avg_order_value = total_value / total_orders if total_orders else 0

k1, k2, k3, k4 = st.columns(4)
k1.metric("Total Sales Value", f"₹{total_value:,.0f}")
k2.metric("Total Quantity Sold", f"{total_qty:,.0f}")
k3.metric("Total Orders", f"{total_orders:,}")
k4.metric("Avg Order Value", f"₹{avg_order_value:,.0f}")

st.divider()

# ─────────────────────────────────────────────
# MONTHLY TREND (FIXED SORTING)
# ─────────────────────────────────────────────
df_filtered["Month_Year"] = df_filtered[DATE_COL].dt.to_period("M").astype(str)

monthly = df_filtered.groupby("Month_Year").agg({
    "Sales Quantity": "sum",
    "Sales Value": "sum"
}).reset_index()

fig_trend = go.Figure()
fig_trend.add_trace(go.Scatter(x=monthly["Month_Year"], y=monthly["Sales Quantity"],
                               mode="lines+markers", name="Quantity"))
fig_trend.add_trace(go.Scatter(x=monthly["Month_Year"], y=monthly["Sales Value"],
                               mode="lines+markers", name="Sales Value", yaxis="y2"))

fig_trend.update_layout(
    yaxis2=dict(overlaying="y", side="right"),
    template="plotly_white",
    height=420
)

st.plotly_chart(fig_trend, use_container_width=True)

# ─────────────────────────────────────────────
# SOURCE OF LEAD – DONUT CHART (SAFE)
# ─────────────────────────────────────────────
st.subheader("📌 Source of Lead Distribution")

lead_source = (
    df_filtered.groupby("Source Of Lead")["Sales Quantity"]
    .sum()
    .reset_index()
)

fig_pie = px.pie(
    lead_source,
    names="Source Of Lead",
    values="Sales Quantity",
    hole=0.45,
    color_discrete_sequence=px.colors.sequential.Reds_r
)

fig_pie.update_traces(textinfo="percent+label")
fig_pie.update_layout(template="plotly_white")

st.plotly_chart(fig_pie, use_container_width=True)
