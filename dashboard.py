import streamlit as st
import pandas as pd
import plotly.express as px

# ---------------- Load Data ----------------
st.set_page_config(page_title="KPI Dashboard", layout="wide")
st.title("📊 KPI Dashboard")

# File uploader
uploaded_file = st.file_uploader("Upload KPI Excel File", type=["xlsx"])
if uploaded_file is None:
    st.info("Please upload an Excel file to continue.")
    st.stop()

df = pd.read_excel(uploaded_file)

# ---------------- Data Cleaning ----------------
# Standardize column names
df.columns = df.columns.str.strip()

# Ensure expected columns exist
expected_cols = [
    "Member", "Month", "Quality Score", "Revision Rate",
    "On-time Delivery", "Target Work Hours", "Actual Work Hours", "Task Name"
]
missing = [c for c in expected_cols if c not in df.columns]
if missing:
    st.error(f"Missing columns in Excel: {missing}")
    st.stop()

# Convert Month to datetime
df["month"] = pd.to_datetime(df["Month"])

# Convert to fractions
df["QS_frac"] = df["Quality Score"] / 100
df["Rev_frac"] = df["Revision Rate"] / 100
df["OnTime"] = df["On-time Delivery"] / 100

# Recompute Efficiency: Target / Actual
df["Eff_frac"] = (df["Target Work Hours"] / df["Actual Work Hours"]).replace([float("inf"), -float("inf")], pd.NA)

# ---------------- Sidebar Filters ----------------
st.sidebar.header("Filters")
members = st.sidebar.multiselect(
    "Select team members:", 
    options=df["Member"].unique(), 
    default=df["Member"].unique()
)
flt = df[df["Member"].isin(members)]

# ---------------- Summary KPIs ----------------
st.header("📌 Summary KPIs")

if not flt.empty:
    total_tasks = len(flt)
    avg_qs = flt["QS_frac"].mean()
    avg_rev = flt["Rev_frac"].mean()
    avg_eff = flt["Eff_frac"].mean()
    avg_on = flt["OnTime"].mean()
    total_manhours = flt["Actual Work Hours"].sum()

    kpis = {
        "Total Tasks": f"{total_tasks:,}",
        "Avg Quality Score": f"{avg_qs:.1%}" if pd.notna(avg_qs) else "N/A",
        "Avg Revision Rate": f"{avg_rev:.1%}" if pd.notna(avg_rev) else "N/A",
        "Avg Efficiency": f"{avg_eff:.1%}" if pd.notna(avg_eff) else "N/A",
        "On-time Delivery": f"{avg_on:.1%}" if pd.notna(avg_on) else "N/A",
        "Total Man-hours": f"{int(total_manhours):,}",
    }

    cols = st.columns(len(kpis))
    for i, (label, val) in enumerate(kpis.items()):
        with cols[i]:
            st.metric(label, val)
else:
    st.info("No data available to compute summary KPIs.")

# ---------------- Individual KPI Tracking ----------------
st.header("📈 Individual KPI Tracking")

def plot_line(df_plot, x_col, y_col, title, is_pct=False):
    if df_plot.empty:
        st.warning(f"No data for {title}")
        return
    fig = px.line(
        df_plot, x=x_col, y=y_col, color="Member", markers=True, title=title
    )
    if is_pct:
        fig.update_yaxes(tickformat=".0%")
    st.plotly_chart(fig, use_container_width=True)

plot_line(flt, "month", "QS_frac", "Average Quality Score (%)", is_pct=True)
plot_line(flt, "month", "Rev_frac", "Average Revision Rate (%)", is_pct=True)
plot_line(flt, "month", "Task Name", "Total Completed Tasks")
plot_line(flt, "month", "OnTime", "On-time Delivery (%)", is_pct=True)
plot_line(flt, "month", "Eff_frac", "Actual Work Efficiency (%)", is_pct=True)
plot_line(flt, "month", "Actual Work Hours", "Man-hours Spent (Total)")

# ---------------- Team KPI Tracking ----------------
st.header("👥 Team KPI Tracking (Averaged)")

team = (
    flt.groupby("month")
    .agg({
        "QS_frac": "mean",
        "Rev_frac": "mean",
        "Task Name": "count",
        "OnTime": "mean",
        "Eff_frac": "mean",
        "Actual Work Hours": "sum",
    })
    .reset_index()
)

plot_line(team, "month", "QS_frac", "Team Average Quality Score (%)", is_pct=True)
plot_line(team, "month", "Rev_frac", "Team Average Revision Rate (%)", is_pct=True)
plot_line(team, "month", "Task Name", "Team Total Completed Tasks")
plot_line(team, "month", "OnTime", "Team On-time Delivery (%)", is_pct=True)
plot_line(team, "month", "Eff_frac", "Team Actual Work Efficiency (%)", is_pct=True)
plot_line(team, "month", "Actual Work Hours", "Team Man-hours Spent (Total)")

# ---------------- Leaderboards ----------------
st.header("🏆 KPI Leaderboards")

leaderboard = (
    flt.groupby("Member")
    .agg({
        "QS_frac": "mean",
        "Rev_frac": "mean",
        "OnTime": "mean",
        "Eff_frac": "mean",
        "Task Name": "count",
        "Actual Work Hours": "sum",
    })
    .reset_index()
)

cols = st.columns(6)
metrics = {
    "Avg Quality Score (%)": ("QS_frac", True),
    "Avg Revision Rate (%)": ("Rev_frac", True),
    "On-time Delivery (%)": ("OnTime", True),
    "Avg Efficiency (%)": ("Eff_frac", True),
    "Total Tasks": ("Task Name", False),
    "Total Man-hours": ("Actual Work Hours", False),
}

for i, (label, (colname, is_pct)) in enumerate(metrics.items()):
    with cols[i]:
        st.subheader(label)
        tmp = leaderboard.sort_values(colname, ascending=False)
        for _, row in tmp.iterrows():
            val = f"{row[colname]:.1%}" if is_pct else f"{int(row[colname]):,}"
            st.write(f"**{row['Member']}**: {val}")
