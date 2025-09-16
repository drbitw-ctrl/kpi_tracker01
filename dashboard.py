# dashboard.py
"""
Streamlit KPI Dashboard — Line Charts + Leaderboards
- Efficiency always recomputed as Target / Actual (100% = perfect)
- QS, Revision, On-time, Efficiency shown as %
- Line charts for individuals + team
- Leaderboards always visible
"""

import re
import pandas as pd
import plotly.express as px
import streamlit as st

# ---------------- Config ----------------
st.set_page_config(page_title="KPI Dashboard", layout="wide")
st.title("📊 KPI Dashboard — Line Charts")

# ---------------- Helpers ----------------
@st.cache_data
def load_excel_anysheet(path_or_buffer):
    xls = pd.ExcelFile(path_or_buffer)
    sheet = xls.sheet_names[0]
    return pd.read_excel(xls, sheet_name=sheet)

def parse_numeric_yyyymmdd(x):
    if pd.isna(x):
        return pd.NaT
    s = str(x).strip()
    if re.match(r"^\d{8}$", s):
        return pd.to_datetime(s, format="%Y%m%d", errors="coerce")
    return pd.to_datetime(s, errors="coerce")

def parse_work_duration_column(df, col="Work Duration"):
    starts, ends = [], []
    for v in df.get(col, []):
        if pd.isna(v):
            starts.append(pd.NaT); ends.append(pd.NaT); continue
        s = str(v)
        parts = re.split(r"\s*[-–—]\s*|\s+to\s+", s)
        st = parse_numeric_yyyymmdd(parts[0])
        en = parse_numeric_yyyymmdd(parts[1]) if len(parts) > 1 else pd.NaT
        starts.append(st); ends.append(en)
    df = df.copy()
    df["start_date"] = pd.to_datetime(starts, errors="coerce")
    df["end_date"] = pd.to_datetime(ends, errors="coerce")
    return df

def clean_and_prepare(df):
    df = df.copy()
    df.columns = [c.strip() for c in df.columns]

    # parse dates
    if "Date Completed" in df.columns:
        df["Date Completed"] = df["Date Completed"].apply(parse_numeric_yyyymmdd)
    if "Work Duration" in df.columns:
        df = parse_work_duration_column(df)

    # fallback for end_date
    if "end_date" in df.columns and "Date Completed" in df.columns:
        df["end_date"] = df["end_date"].fillna(df["Date Completed"])

    # timeline column
    if "end_date" in df.columns:
        fallback = df["end_date"]
    elif "Date Completed" in df.columns:
        fallback = df["Date Completed"]
    else:
        fallback = pd.NaT
    df["month_dt"] = pd.to_datetime(fallback, errors="coerce")
    df["month"] = df["month_dt"].dt.to_period("M").dt.to_timestamp()

    # numeric conversions
    for c in ["Target Work Hours", "Actual Work Hours", "QS%", "Revision/s"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    # recompute efficiency (Target / Actual)
    if "Target Work Hours" in df.columns and "Actual Work Hours" in df.columns:
        df["Eff_frac"] = df["Target Work Hours"] / df["Actual Work Hours"]
    else:
        df["Eff_frac"] = pd.NA

    # on-time delivery
    if "Date Completed" in df.columns and "end_date" in df.columns:
        df["OnTime"] = (df["Date Completed"] <= df["end_date"]).astype("Int64")
    else:
        df["OnTime"] = pd.NA

    # normalize QS
    if "QS%" in df.columns:
        df["QS_frac"] = df["QS%"].copy()
        if df["QS_frac"].max(skipna=True) > 1.5:
            df["QS_frac"] = df["QS_frac"] / 100.0

    # normalize Revision
    if "Revision/s" in df.columns:
        df["Rev_frac"] = df["Revision/s"].copy()
        if df["Rev_frac"].max(skipna=True) > 1.5:
            df["Rev_frac"] = df["Rev_frac"] / 100.0

    # safe manhours
    if "Actual Work Hours" in df.columns:
        df["Actual Work Hours"] = df["Actual Work Hours"].fillna(0)

    # task id
    df["_task_id"] = range(len(df))

    return df

def plot_line(df_plot, x_col, y_col, title, color_col=None, is_pct=False):
    if df_plot.empty: 
        return
    if color_col:
        fig = px.line(df_plot, x=x_col, y=y_col, color=color_col, markers=True, title=title)
    else:
        fig = px.line(df_plot, x=x_col, y=y_col, markers=True, title=title)
    if is_pct:
        fig.update_yaxes(tickformat=".0%")
    st.plotly_chart(fig, use_container_width=True)

# ---------------- Data load ----------------
st.sidebar.header("Upload Data")
uploaded = st.sidebar.file_uploader("Upload Excel", type=["xlsx","xls"])

if uploaded:
    raw = load_excel_anysheet(uploaded)
else:
    raw = load_excel_anysheet("/mnt/data/KPI METRICS 2.xlsx")

df = clean_and_prepare(raw)

members = sorted(df["Name"].dropna().unique().tolist()) if "Name" in df.columns else []
selected = st.sidebar.multiselect("Members", options=members, default=members)

flt = df[df["Name"].isin(selected)] if selected else df

# ---------------- Aggregations ----------------
group_ind = flt.groupby(["month","Name"], as_index=False).agg(
    QS_mean=("QS_frac","mean"),
    Rev_mean=("Rev_frac","mean"),
    OnTime_pct=("OnTime","mean"),
    Eff_mean=("Eff_frac","mean"),
    Manhours=("Actual Work Hours","sum"),
    Tasks=("_task_id","count")
)

group_team = flt.groupby("month", as_index=False).agg(
    QS_mean=("QS_frac","mean"),
    Rev_mean=("Rev_frac","mean"),
    OnTime_pct=("OnTime","mean"),
    Eff_mean=("Eff_frac","mean"),
    Manhours=("Actual Work Hours","sum"),
    Tasks=("_task_id","count")
)

# ---------------- Header ----------------
latest = df["month_dt"].max()
if pd.notna(latest):
    st.subheader(f"Team KPI Dashboard — Updated as of {latest.strftime('%Y-%m-%d')}")

# ---------------- Individual Charts ----------------
st.header("Individual KPI Tracking")
plot_line(group_ind, "month","QS_mean","Quality Score (%)", color_col="Name", is_pct=True)
plot_line(group_ind, "month","Rev_mean","Revision Rate (%)", color_col="Name", is_pct=True)
plot_line(group_ind, "month","Tasks","Completed Tasks", color_col="Name")
plot_line(group_ind, "month","OnTime_pct","On-time Delivery (%)", color_col="Name", is_pct=True)
plot_line(group_ind, "month","Eff_mean","Efficiency (%)", color_col="Name", is_pct=True)
plot_line(group_ind, "month","Manhours","Man-hours Spent", color_col="Name")

# ---------------- Team Charts ----------------
st.header("Team KPI Tracking")
plot_line(group_team, "month","QS_mean","Team Avg Quality Score (%)", is_pct=True)
plot_line(group_team, "month","Rev_mean","Team Avg Revision Rate (%)", is_pct=True)
plot_line(group_team, "month","Tasks","Team Completed Tasks")
plot_line(group_team, "month","OnTime_pct","Team On-time Delivery (%)", is_pct=True)
plot_line(group_team, "month","Eff_mean","Team Efficiency (%)", is_pct=True)
plot_line(group_team, "month","Manhours","Team Man-hours Spent")

# ---------------- Leaderboards ----------------
st.header("Leaderboards (Latest Month)")
if not group_ind.empty:
    latest_month = group_ind["month"].max()
    latest_df = group_ind[group_ind["month"]==latest_month]

    boards = {
        "Efficiency (%)":"Eff_mean",
        "Quality Score (%)":"QS_mean",
        "On-time Delivery (%)":"OnTime_pct",
        "Tasks Completed":"Tasks",
        "Man-hours":"Manhours"
    }

    cols = st.columns(len(boards))
    for i,(title,colname) in enumerate(boards.items()):
        tmp = latest_df[["Name",colname]].sort_values(colname,ascending=False)
        if "%" in title:
            tmp[colname] = tmp[colname].apply(lambda v: f"{v:.1%}" if pd.notna(v) else "N/A")
        else:
            tmp[colname] = tmp[colname].apply(lambda v: f"{int(v):,}" if pd.notna(v) else "N/A")
        cols[i].subheader(title)
        cols[i].dataframe(tmp.reset_index(drop=True), use_container_width=True)
# ---------------- Streamlit App ----------------
st.title("📊 KPI Dashboard")

# Sidebar filters
st.sidebar.header("Filters")
members = st.sidebar.multiselect("Select team members:", df["Member"].unique(), default=df["Member"].unique())
flt = df[df["Member"].isin(members)]

# --- Summary KPIs ---
st.header("📌 Summary KPIs")

if not flt.empty:
    latest_month = flt["month"].max()
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
