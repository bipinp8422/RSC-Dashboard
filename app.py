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
    page_title="RSC Sales Performance Dashboard",
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
# FILE UPLOAD (to ensure functionality like original but avoid FileNotFound)
# ─────────────────────────────────────────────
st.sidebar.header("📂 Data Upload")
uploaded_file = st.sidebar.file_uploader("Upload your Excel file (.xlsb)", type=["xlsb"], help="Upload 'MOM RSC Performance_Jan'24 To Dec'25- North South_Region V1.xlsb' or similar")

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
df["Year"] = df[DATE_COL].dt.year
df["Month_No"] = df[DATE_COL].dt.month
df["Month_Name"] = df[DATE_COL].dt.strftime("%b")
df = df[df["Year"].between(2024, 2025)]

# ─────────────────────────────────────────────
# SIDEBAR FILTERS
# ─────────────────────────────────────────────
st.sidebar.header("🔍 Filters")
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

if df_filtered.empty:
    st.error("No data found for the selected filters.")
    st.stop()

# ─────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────
@st.cache_resource
def load_logo():
    return Image.open("canon-press-centre-canon-logo.png")

col1, col2 = st.columns([0.15, 0.85])
with col1:
    try:
        st.image(load_logo(), width=140)
    except:
        st.markdown("**Canon**")

with col2:
    st.markdown(
        "<h1 style='margin-bottom:0'>RSC Sales Performance Dashboard</h1>",
        unsafe_allow_html=True
    )
st.markdown(f"**Last Updated:** {datetime.now().strftime('%d %B %Y')}")

st.divider()

# ─────────────────────────────────────────────
# KPI CARDS (Added for attractiveness)
# ─────────────────────────────────────────────
st.subheader("📊 Key Performance Indicators")

total_qty = df_filtered["Sales Quantity"].sum()
total_value = df_filtered["Sales Value"].sum() if "Sales Value" in df_filtered.columns else 0
total_orders = len(df_filtered)
avg_order_value = total_value / total_orders if total_orders > 0 else 0

current_year = df_filtered["Year"].max()
prev_year_data = df[(df["Year"] == current_year - 1) & (df["Status"] == "Passed")]
prev_value = prev_year_data["Sales Value"].sum() if "Sales Value" in prev_year_data.columns else 0

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
# MONTH-WISE SALES TREND (Enhanced to dual axis for attractiveness)
# ─────────────────────────────────────────────
st.subheader("📈 Month-wise Sales Trend (Quantity – Passed Only)")

month_qty = (
    df_filtered
    .groupby(["Month_No", "Month_Name"], as_index=False)["Sales Quantity"]
    .sum()
    .sort_values("Month_No")
)

fig_trend = go.Figure()

fig_trend.add_trace(go.Bar(
    x=month_qty["Month_Name"], y=month_qty["Sales Quantity"],
    name='Quantity', marker_color='#c00000'
))

if "Sales Value" in df_filtered.columns:
    month_val = (
        df_filtered
        .groupby(["Month_No", "Month_Name"], as_index=False)["Sales Value"]
        .sum()
        .sort_values("Month_No")
    )
    fig_trend.add_trace(go.Scatter(
        x=month_val["Month_Name"], y=month_val["Sales Value"],
        mode='lines+markers', name='Sales Value (₹)', yaxis="y2", line=dict(color='#1f77b4', width=3)
    ))

fig_trend.update_layout(
    title="Month-wise Sales Trend (Quantity and Value – Passed Only)",
    xaxis_title="Month",
    yaxis=dict(title="Sales Quantity", titlefont=dict(color='#c00000')),
    yaxis2=dict(title="Sales Value (₹)", overlaying='y', side='right', titlefont=dict(color='#1f77b4')),
    template="plotly_white",
    legend=dict(orientation="h", yanchor="bottom", y=1.02),
    height=420,
    xaxis_tickangle=-30
)

st.plotly_chart(fig_trend, use_container_width=True)

# ─────────────────────────────────────────────
# TOP 5 PRODUCT CATEGORIES – QTY & VALUE
# ─────────────────────────────────────────────
colA, colB = st.columns(2)
with colA:
    cat_qty = (
        df_filtered.groupby("Product Category", as_index=False)["Sales Quantity"]
        .sum().sort_values("Sales Quantity", ascending=False).head(5)
    )
    st.plotly_chart(
        px.bar(
            cat_qty,
            y="Product Category",
            x="Sales Quantity",
            orientation="h",
            text="Sales Quantity",
            title="Top 5 Product Categories – Quantity",
            color_discrete_sequence=["#c00000"]
        ).update_traces(textposition="inside")
        .update_layout(template="plotly_white"),
        use_container_width=True
    )
with colB:
    if "Sales Value" in df_filtered.columns:
        cat_val = (
            df_filtered.groupby("Product Category", as_index=False)["Sales Value"]
            .sum().sort_values("Sales Value", ascending=False).head(5)
        )
        st.plotly_chart(
            px.bar(
                cat_val,
                y="Product Category",
                x="Sales Value",
                orientation="h",
                text="Sales Value",
                title="Top 5 Product Categories – Value",
                color_discrete_sequence=["#c00000"]
            ).update_traces(textposition="inside")
            .update_layout(template="plotly_white"),
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
        px.bar(
            top_products,
            y="Model Name",
            x="Sales Quantity",
            orientation="h",
            text="Sales Quantity",
            title="Top 5 Best Seller Products",
            color_discrete_sequence=["#c00000"]
        ).update_traces(textposition="inside")
        .update_layout(template="plotly_white"),
        use_container_width=True
    )
with colD:
    top_stores = (
        df_filtered.groupby("Storename", as_index=False)["Sales Quantity"]
        .sum().sort_values("Sales Quantity", ascending=False).head(5)
    )
    st.plotly_chart(
        px.bar(
            top_stores,
            y="Storename",
            x="Sales Quantity",
            orientation="h",
            text="Sales Quantity",
            title="Top 5 Stores",
            color_discrete_sequence=["#c00000"]
        ).update_traces(textposition="inside")
        .update_layout(template="plotly_white"),
        use_container_width=True
    )

# ─────────────────────────────────────────────
# TOP 10 SELLERS – LEADERSHIP BOARD
# ─────────────────────────────────────────────
leaderboard = (
    df_filtered
    .groupby("Name", as_index=False)["Sales Quantity"]
    .sum()
    .sort_values("Sales Quantity", ascending=False)
    .head(10)
)
st.plotly_chart(
    px.bar(
        leaderboard,
        x="Sales Quantity",
        y="Name",
        orientation="h",
        text="Sales Quantity",
        title="🏆 Top 10 Sellers – Leadership Board",
        color_discrete_sequence=["#c00000"]
    ).update_layout(yaxis=dict(autorange="reversed"), template="plotly_white"),
    use_container_width=True
)

# ─────────────────────────────────────────────
# SOURCE OF LEAD – CIRCULAR (DONUT) CHART
# ─────────────────────────────────────────────
lead_source_perf = (
    df_filtered
    .groupby("Source Of Lead", as_index=False)["Sales Quantity"]
    .sum()
)
st.plotly_chart(
    px.pie(
        lead_source_perf,
        names="Source Of Lead",
        values="Sales Quantity",
        hole=0.45,
        title="📌 Source Of Lead Contribution (%)",
        color_discrete_sequence=px.colors.sequential.Reds_r
    ).update_traces(textinfo="percent+label", textfont_size=13)
    .update_layout(template="plotly_white"),
    use_container_width=True
)
