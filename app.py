import streamlit as st
import pandas as pd
from PIL import Image
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime

# ─────────────────────────────────────────────
# Page Configuration
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Retail Sales Performance Dashboard",
    page_icon="🛍️",
    layout="wide"
)

# Custom CSS for better aesthetics
st.markdown("""
    <style>
    .main {background-color: #f8f9fa;}
    .metric-card {background-color: white; padding: 15px; border-radius: 10px; box-shadow: 0 2px 5px rgba(0,0,0,0.1);}
    h1 {color: #c00000 !important;}
    </style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# FILE UPLOAD
# ─────────────────────────────────────────────
st.sidebar.header("📂 Data Upload")
uploaded_file = st.sidebar.file_uploader("Upload your Excel file (.xlsb)", type=["xlsb"])

if uploaded_file is None:
    st.warning("Please upload the Excel file to proceed.")
    st.stop()

# ─────────────────────────────────────────────
# DATA LOADING & PREPROCESSING
# ─────────────────────────────────────────────
@st.cache_data(show_spinner="Loading data...")
def load_data(file):
    df = pd.read_excel(
        file,
        sheet_name="RAW data",
        skiprows=1,
        engine="pyxlsb"
    )
    return df

df = load_data(uploaded_file)
df.columns = df.columns.astype(str).str.strip()

# Date Handling
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

# Date Range
min_date = df[DATE_COL].min().date()
max_date = df[DATE_COL].max().date()
date_range = st.sidebar.date_input("Select Date Range", 
                                   [min_date, max_date], 
                                   min_value=min_date, 
                                   max_value=max_date)

# Other Filters
selected_city = st.sidebar.multiselect("City", sorted(df["City"].dropna().unique()), 
                                       default=sorted(df["City"].dropna().unique()))

selected_store = st.sidebar.multiselect("Store Name", sorted(df["Storename"].dropna().unique()), 
                                        default=sorted(df["Storename"].dropna().unique()))

selected_category = st.sidebar.multiselect("Product Category", sorted(df["Product Category"].dropna().unique()), 
                                           default=sorted(df["Product Category"].dropna().unique()))

selected_name = st.sidebar.multiselect("Sales Person", sorted(df["Name"].dropna().unique()), 
                                       default=sorted(df["Name"].dropna().unique()))

# ─────────────────────────────────────────────
# APPLY FILTERS
# ─────────────────────────────────────────────
df_filtered = df[df["Status"] == "Passed"].copy()

if len(date_range) == 2:
    start_date, end_date = date_range
    df_filtered = df_filtered[(df_filtered[DATE_COL] >= pd.to_datetime(start_date)) & 
                              (df_filtered[DATE_COL] <= pd.to_datetime(end_date))]

if selected_city:
    df_filtered = df_filtered[df_filtered["City"].isin(selected_city)]
if selected_store:
    df_filtered = df_filtered[df_filtered["Storename"].isin(selected_store)]
if selected_category:
    df_filtered = df_filtered[df_filtered["Product Category"].isin(selected_category)]
if selected_name:
    df_filtered = df_filtered[df_filtered["Name"].isin(selected_name)]

if df_filtered.empty:
    st.error("No data found for the selected filters.")
    st.stop()

# ─────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────
col1, col2 = st.columns([0.15, 0.85])
with col1:
    try:
        logo = Image.open("canon-press-centre-canon-logo.png")
        st.image(logo, width=150)
    except:
        st.markdown("**Canon**")

with col2:
    st.markdown("<h1 style='margin-bottom:0;'>Retail Sales Performance Dashboard</h1>", unsafe_allow_html=True)
    st.caption(f"Last Updated: {datetime.now().strftime('%d %B %Y, %I:%M %p')}")

st.divider()

# ─────────────────────────────────────────────
# KPI CARDS
# ─────────────────────────────────────────────
st.subheader("📊 Key Performance Indicators")

total_qty = df_filtered["Sales Quantity"].sum()
total_value = df_filtered["Sales Value"].sum()
total_orders = len(df_filtered)
avg_order_value = total_value / total_orders if total_orders > 0 else 0

# Simple YoY Growth (Optional - Can be enhanced further)
current_year = df_filtered[DATE_COL].dt.year.max()
prev_year_data = df[(df[DATE_COL].dt.year == current_year - 1) & 
                    (df["Status"] == "Passed")]
prev_value = prev_year_data["Sales Value"].sum()

yoy_growth = ((total_value - prev_value) / prev_value * 100) if prev_value > 0 else 0

kpi1, kpi2, kpi3, kpi4 = st.columns(4)

with kpi1:
    st.metric("Total Sales Value", f"₹{total_value:,.0f}", delta=f"{yoy_growth:.1f}% YoY")

with kpi2:
    st.metric("Total Quantity Sold", f"{total_qty:,.0f}")

with kpi3:
    st.metric("Total Orders", f"{total_orders:,}")

with kpi4:
    st.metric("Avg Order Value", f"₹{avg_order_value:,.0f}")

st.divider()

# ─────────────────────────────────────────────
# MONTHLY TREND - Dual Axis Line Chart (Best Practice)
# ─────────────────────────────────────────────
st.subheader("📈 Monthly Sales Trend")

df_filtered["Month"] = df_filtered[DATE_COL].dt.strftime("%b %Y")

monthly = df_filtered.groupby("Month").agg({
    "Sales Quantity": "sum",
    "Sales Value": "sum"
}).reset_index()

fig_trend = go.Figure()

fig_trend.add_trace(go.Scatter(
    x=monthly["Month"], y=monthly["Sales Quantity"],
    mode='lines+markers', name='Quantity', line=dict(color='#c00000', width=3)
))

fig_trend.add_trace(go.Scatter(
    x=monthly["Month"], y=monthly["Sales Value"],
    mode='lines+markers', name='Sales Value (₹)', yaxis="y2", line=dict(color='#1f77b4', width=3)
))

fig_trend.update_layout(
    title="Monthly Quantity vs Sales Value Trend",
    xaxis_title="Month",
    yaxis=dict(title=dict(text="Sales Quantity", font=dict(color='#c00000'))),
    yaxis2=dict(title=dict(text="Sales Value (₹)", font=dict(color='#1f77b4')), overlaying='y', side='right'),
    template="plotly_white",
    legend=dict(orientation="h", yanchor="bottom", y=1.02),
    height=420
)

st.plotly_chart(fig_trend, use_container_width=True)

# ─────────────────────────────────────────────
# TOP CATEGORIES + TOP PRODUCTS
# ─────────────────────────────────────────────
col1, col2 = st.columns(2)

with col1:
    st.subheader("🏷️ Top 5 Product Categories")
    cat_qty = df_filtered.groupby("Product Category")["Sales Quantity"].sum().nlargest(5).reset_index()
    fig_cat = px.bar(cat_qty, y="Product Category", x="Sales Quantity", orientation='h',
                     text="Sales Quantity", color_discrete_sequence=["#c00000"])
    fig_cat.update_traces(textposition='inside')
    fig_cat.update_layout(template="plotly_white", height=380)
    st.plotly_chart(fig_cat, use_container_width=True)

with col2:
    st.subheader("🔥 Top 5 Best Selling Products")
    top_prod = df_filtered.groupby("Model Name")["Sales Quantity"].sum().nlargest(5).reset_index()
    fig_prod = px.bar(top_prod, y="Model Name", x="Sales Quantity", orientation='h',
                      text="Sales Quantity", color_discrete_sequence=["#c00000"])
    fig_prod.update_traces(textposition='inside')
    fig_prod.update_layout(template="plotly_white", height=380)
    st.plotly_chart(fig_prod, use_container_width=True)

# ─────────────────────────────────────────────
# TOP STORES + LEADERSHIP BOARD
# ─────────────────────────────────────────────
col3, col4 = st.columns(2)

with col3:
    st.subheader("🏪 Top 5 Stores")
    top_stores = df_filtered.groupby("Storename")["Sales Quantity"].sum().nlargest(5).reset_index()
    fig_store = px.bar(top_stores, y="Storename", x="Sales Quantity", orientation='h',
                       text="Sales Quantity", color_discrete_sequence=["#c00000"])
    fig_store.update_traces(textposition='inside')
    fig_store.update_layout(template="plotly_white", height=380)
    st.plotly_chart(fig_store, use_container_width=True)

with col4:
    st.subheader("🏆 Top 10 Sellers - Leadership Board")
    leaderboard = df_filtered.groupby("Name")["Sales Quantity"].sum().nlargest(10).reset_index()
    fig_leader = px.bar(leaderboard, y="Name", x="Sales Quantity", orientation='h',
                        text="Sales Quantity", color_discrete_sequence=["#c00000"])
    fig_leader.update_traces(textposition='inside')
    fig_leader.update_layout(template="plotly_white", height=380)
    st.plotly_chart(fig_leader, use_container_width=True)

# ─────────────────────────────────────────────
# SOURCE OF LEAD
# ─────────────────────────────────────────────
st.subheader("📌 Source of Lead Distribution")
lead_source = df_filtered.groupby("Source Of Lead")["Sales Quantity"].sum().reset_index()

fig_pie = px.pie(lead_source, names="Source Of Lead", values="Sales Quantity",
                 hole=0.45, color_discrete_sequence=px.colors.sequential.Reds_r)

fig_pie.update_traces(textinfo="percent+label", textfont_size=13)
fig_pie.update_layout(template="plotly_white")

st.plotly_chart(fig_pie, use_container_width=True)
