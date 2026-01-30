import streamlit as st
import pandas as pd
from PIL import Image
import plotly.express as px
from datetime import datetime

# ─────────────────────────────────────────────
# Page configuration
# ─────────────────────────────────────────────
st.set_page_config(page_title="RSC Sales Dashboard", layout="wide")

# ─────────────────────────────────────────────
# DATA LOADING
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
df.columns = df.columns.astype(str).str.strip()

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

if pd.api.types.is_numeric_dtype(df[DATE_COL]):
    df[DATE_COL] = pd.to_datetime(df[DATE_COL], unit="D", origin="1899-12-30", errors="coerce")
else:
    df[DATE_COL] = pd.to_datetime(df[DATE_COL], errors="coerce")

df["Year"] = df[DATE_COL].dt.year
df["Month_No"] = df[DATE_COL].dt.month
df["Month_Name"] = df[DATE_COL].dt.strftime("%b")
df = df[df["Year"].between(2024, 2025)]

# ─────────────────────────────────────────────
# SIDEBAR FILTERS (DATA ONLY)
# ─────────────────────────────────────────────
st.sidebar.title("🔍 Filters")

selected_year = st.sidebar.multiselect(
    "Year", sorted(df["Year"].dropna().unique()),
    default=sorted(df["Year"].dropna().unique())
)

selected_city = st.sidebar.multiselect(
    "City", sorted(df["City"].dropna().unique()),
    default=sorted(df["City"].dropna().unique())
)

selected_store = st.sidebar.multiselect(
    "Store Name", sorted(df["Storename"].dropna().unique()),
    default=sorted(df["Storename"].dropna().unique())
)

selected_name = st.sidebar.multiselect(
    "Name", sorted(df["Name"].dropna().unique()),
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
# HEADER
# ─────────────────────────────────────────────
@st.cache_resource
def load_logo():
    return Image.open("canon-press-centre-canon-logo.png")

col1, col2 = st.columns([0.15, 0.85])
with col1:
    st.image(load_logo(), width=140)
with col2:
    st.markdown("<h1>RSC Sales Performance Dashboard</h1>", unsafe_allow_html=True)

st.markdown(f"**Last Updated:** {datetime.now().strftime('%d %B %Y')}")

# ─────────────────────────────────────────────
# MONTH-WISE SALES TREND
# ─────────────────────────────────────────────
st.subheader("📅 Month-wise Sales Trend")
chart_type_month = st.selectbox("Chart Type", ["Bar", "Line", "Area"], key="month")

month_qty = (
    df_filtered.groupby(["Month_No", "Month_Name"], as_index=False)["Sales Quantity"]
    .sum().sort_values("Month_No")
)

if chart_type_month == "Bar":
    fig = px.bar(month_qty, x="Month_Name", y="Sales Quantity", text="Sales Quantity")
elif chart_type_month == "Line":
    fig = px.line(month_qty, x="Month_Name", y="Sales Quantity", markers=True)
else:
    fig = px.area(month_qty, x="Month_Name", y="Sales Quantity")

st.plotly_chart(fig, use_container_width=True)

# ─────────────────────────────────────────────
# TOP 10 SELLERS – LEADERSHIP
# ─────────────────────────────────────────────
st.subheader("🏆 Top 10 Sellers – Leadership Board")
chart_type_leader = st.selectbox("Chart Type", ["Bar", "Pie"], key="leader")

leaderboard = (
    df_filtered.groupby("Name", as_index=False)["Sales Quantity"]
    .sum().sort_values("Sales Quantity", ascending=False).head(10)
)

if chart_type_leader == "Bar":
    fig = px.bar(
        leaderboard, x="Sales Quantity", y="Name",
        orientation="h", text="Sales Quantity"
    ).update_layout(yaxis=dict(autorange="reversed"))
else:
    fig = px.pie(leaderboard, names="Name", values="Sales Quantity")

st.plotly_chart(fig, use_container_width=True)

# ─────────────────────────────────────────────
# SOURCE OF LEAD PERFORMANCE (NO FILTER)
# ─────────────────────────────────────────────
st.subheader("📌 Source Of Lead Performance")
chart_type_source = st.selectbox("Chart Type", ["Bar", "Line", "Area", "Pie"], key="source")

lead_perf = (
    df_filtered.groupby("Source Of Lead", as_index=False)["Sales Quantity"]
    .sum().sort_values("Sales Quantity", ascending=False)
)

if chart_type_source == "Bar":
    fig = px.bar(
        lead_perf, x="Sales Quantity", y="Source Of Lead",
        orientation="h", text="Sales Quantity"
    ).update_layout(yaxis=dict(autorange="reversed"))
elif chart_type_source == "Line":
    fig = px.line(lead_perf, x="Source Of Lead", y="Sales Quantity", markers=True)
elif chart_type_source == "Area":
    fig = px.area(lead_perf, x="Source Of Lead", y="Sales Quantity")
else:
    fig = px.pie(lead_perf, names="Source Of Lead", values="Sales Quantity")

st.plotly_chart(fig, use_container_width=True)
