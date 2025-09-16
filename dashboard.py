# dashboard.py
"""
Streamlit KPI Dashboard — Presentation-ready (Line charts + Leaderboards)
- Uses full timeline (no date range filter)
- Sidebar: member multi-select
- Individual & Team line charts
- Leaderboards always visible (side-by-side)
- Proper percentage formatting for QS, Revision, Efficiency, On-time
"""

import re
from datetime import datetime
import streamlit as st
import pandas as pd
import plotly.express as px

# ---------- Page config ----------
st.set_page_config(page_title="KPI Dashboard — Line Charts", layout="wide")
st.title("📊 KPI Dashboard — Line Charts")

# ---------- Helpers ----------
@st.cache_data
def load_excel_anysheet(path_or_buffer):
    """
    Load first sensible sheet: prefer '5','1','Sheet1' otherwise first sheet.
    """
    try:
        xls = pd.ExcelFile(path_or_buffer)
        preferred = None
        for s in ['5', '1', 'Sheet1', 'Sheet 1']:
            if s in xls.sheet_names:
                preferred = s
                break
        sheet = preferred if preferred is not None else xls.sheet_names[0]
        df = pd.read_excel(xls, sheet_name=sheet)
    except Exception:
        df = pd.read_excel(path_or_buffer)
    return df

def parse_numeric_yyyymmdd(x):
    if pd.isna(x):
        return pd.NaT
    s = str(x).strip()
    # numeric 8-digit YYYYMMDD
    m = re.match(r"^(\d{8})(?:\.0)?$", s)
    if m:
        try:
            return pd.to_datetime(m.group(1), format="%Y%m%d")
        except Exception:
            pass
    # try several formats and pandas fallback
    fmts = ["%Y%m%d","%Y-%m-%d","%Y/%m/%d","%Y.%m.%d","%d/%m/%Y","%m/%d/%Y"]
    for f in fmts:
        try:
            return pd.to_datetime(s, format=f)
        except Exception:
            continue
    return pd.to_datetime(s, errors='coerce')

def parse_work_duration_column(df, col="Work Duration"):
    starts, ends = [], []
    for v in df.get(col, pd.Series([pd.NA]*len(df))):
        if pd.isna(v):
            starts.append(pd.NaT); ends.append(pd.NaT); continue
        s = str(v).strip()
        # split by hyphen or 'to'
        parts = re.split(r"\s*[-–—]\s*|\s+to\s+", s)
        if len(parts) >= 2:
            st = parse_numeric_yyyymmdd(parts[0])
            en = parse_numeric_yyyymmdd(parts[1])
        else:
            st = parse_numeric_yyyymmdd(parts[0])
            en = pd.NaT
        starts.append(st); ends.append(en)
    df = df.copy()
    df['start_date'] = pd.to_datetime(starts, errors='coerce')
    df['end_date'] = pd.to_datetime(ends, errors='coerce')
    return df

def clean_and_prepare(df):
    df = df.copy()
    # normalize column names
    df.columns = [c.strip() for c in df.columns]

    # parse date columns
    if 'Date Completed' in df.columns:
        df['Date Completed'] = df['Date Completed'].apply(parse_numeric_yyyymmdd)

    if 'Work Duration' in df.columns:
        df = parse_work_duration_column(df, 'Work Duration')

    # if end_date missing, use Date Completed if available
    if 'end_date' in df.columns and 'Date Completed' in df.columns:
        df['end_date'] = df['end_date'].fillna(df['Date Completed'])

    # decide month timestamp column (for full timeline grouping)
    fallback = None
    if 'end_date' in df.columns:
        fallback = df['end_date']
    elif 'start_date' in df.columns:
        fallback = df['start_date']
    elif 'Date Completed' in df.columns:
        fallback = df['Date Completed']
    else:
        fallback = pd.NaT

    df['month_dt'] = pd.to_datetime(fallback, errors='coerce')
    df['month'] = df['month_dt'].dt.to_period('M').dt.to_timestamp()

    # numeric conversion for relevant columns
    for c in ['Target Work Hours','Actual Work Hours','Efficiency','QS%','Revision/s']:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce')

    # compute OnTime (1 if Date Completed <= end_date)
    if 'Date Completed' in df.columns and 'end_date' in df.columns:
        df['OnTime'] = (df['Date Completed'] <= df['end_date']).astype('Int64')
    else:
        df['OnTime'] = pd.NA

    # normalize percent-like columns to fractions (0..1)
    if 'QS%' in df.columns:
        df['QS_frac'] = df['QS%'].copy()
        if df['QS_frac'].max(skipna=True) is not None and df['QS_frac'].max(skipna=True) > 1.5:
            df['QS_frac'] = df['QS_frac'] / 100.0

    if 'Revision/s' in df.columns:
        df['Rev_frac'] = df['Revision/s'].copy()
        # if values look like percent points (e.g., 20) convert to fraction
        if df['Rev_frac'].max(skipna=True) is not None and df['Rev_frac'].max(skipna=True) > 1.5:
            df['Rev_frac'] = df['Rev_frac'] / 100.0

    if 'Efficiency' in df.columns:
        df['Eff_frac'] = df['Efficiency'].copy()
        if df['Eff_frac'].max(skipna=True) is not None and df['Eff_frac'].max(skipna=True) > 1.5:
            df['Eff_frac'] = df['Eff_frac'] / 100.0

    # Actual Work Hours safe numeric
    if 'Actual Work Hours' in df.columns:
        df['Actual Work Hours'] = pd.to_numeric(df['Actual Work Hours'], errors='coerce').fillna(0)

    # task id for counting
    if 'Ref. number' in df.columns:
        df['_task_id'] = df['Ref. number']
    else:
        df['_task_id'] = range(len(df))
    if "Target Work Hours" in df.columns and "Actual Work Hours" in df.columns:
        df["Eff_frac"] = df["Target Work Hours"] / df["Actual Work Hours"]
    else:
        df["Eff_frac"] = pd.NA


    return df

def plot_line_generic(df_plot, x_col, y_col, title, color_col=None, is_pct=False):
    if df_plot.empty:
        st.write(f"No data for {title}")
        return
    if color_col and color_col in df_plot.columns:
        fig = px.line(df_plot, x=x_col, y=y_col, color=color_col, markers=True, title=title)
    else:
        fig = px.line(df_plot, x=x_col, y=y_col, markers=True, title=title)
    if is_pct:
        fig.update_yaxes(tickformat=".0%")
    fig.update_layout(margin=dict(l=30, r=20, t=50, b=30))
    st.plotly_chart(fig, use_container_width=True)

# ---------- UI: Sidebar ----------
st.sidebar.header("Data & Filters")
uploaded = st.sidebar.file_uploader("Upload KPI Excel (.xlsx)", type=['xlsx','xls'])
st.sidebar.markdown("Select member(s) to compare (default = all).")
# no date filter — full timeline

# ---------- Load data ----------
try:
    if uploaded is not None:
        raw = load_excel_anysheet(uploaded)
    else:
        raw = load_excel_anysheet("/mnt/data/KPI METRICS 2.xlsx")
except Exception as e:
    st.error(f"Failed to load Excel file: {e}")
    st.stop()

if raw is None or raw.empty:
    st.error("No data found. Upload a valid KPI Excel file.")
    st.stop()

df = clean_and_prepare(raw)

# ---------- Header summary ----------
latest_dt = None
if df['month_dt'].dropna().any():
    latest_dt = df['month_dt'].max()
    st.subheader(f"Team KPI Dashboard — Updated as of {latest_dt.strftime('%Y-%m-%d')}")
else:
    st.subheader("Team KPI Dashboard")

# ---------- Sidebar selects ----------
members = sorted(df['Name'].dropna().unique().tolist()) if 'Name' in df.columns else []
selected_members = st.sidebar.multiselect("Team member(s)", options=members, default=members)

# apply member selection filter to individual charts (team charts still computed from full filtered set)
flt = df.copy()
if selected_members:
    flt = flt[flt['Name'].isin(selected_members)]

if flt.empty:
    st.error("No data after member selection. Please choose different members or upload a different file.")
    st.stop()

# ---------- Aggregations ----------
# monthly per-member aggregation
group_ind = flt.groupby(['month','Name'], as_index=False).agg(
    QS_mean = ('QS_frac','mean') if 'QS_frac' in flt.columns else ('QS%','mean'),
    Rev_mean = ('Rev_frac','mean') if 'Rev_frac' in flt.columns else ('Revision/s','mean'),
    OnTime_pct = ('OnTime','mean'),
    Eff_mean = ('Eff_frac','mean') if 'Eff_frac' in flt.columns else ('Efficiency','mean'),
    Manhours = ('Actual Work Hours','sum'),
    Tasks = ('_task_id','count')
).sort_values('month')

# team-level monthly aggregation (use full dataset, but filtered by selected members? 
# as requested earlier: show team averaged across members in chosen selection; we will use flt)
group_team = flt.groupby(['month'], as_index=False).agg(
    QS_mean = ('QS_frac','mean') if 'QS_frac' in flt.columns else ('QS%','mean'),
    Rev_mean = ('Rev_frac','mean') if 'Rev_frac' in flt.columns else ('Revision/s','mean'),
    OnTime_pct = ('OnTime','mean'),
    Eff_mean = ('Eff_frac','mean') if 'Eff_frac' in flt.columns else ('Efficiency','mean'),
    Manhours = ('Actual Work Hours','sum'),
    Tasks = ('_task_id','count')
).sort_values('month')

# ---------- Top KPI quick metrics ----------
st.markdown("---")
st.subheader("Quick team metrics (filtered selection)")
col1, col2, col3, col4, col5 = st.columns(5)
total_tasks = int(flt['_task_id'].count())
completed_tasks = int(flt[flt['Status'].str.lower() == 'completed'].shape[0]) if 'Status' in flt.columns else "N/A"
avg_eff_team = group_team['Eff_mean'].mean() if not group_team.empty else None
avg_qs_team = group_team['QS_mean'].mean() if not group_team.empty else None
total_manhours = int(flt['Actual Work Hours'].sum()) if 'Actual Work Hours' in flt.columns else 0

col1.metric("Total tasks (filtered)", total_tasks)
col2.metric("Completed tasks", completed_tasks)
col3.metric("Avg Efficiency (team)", f"{avg_eff_team:.1%}" if avg_eff_team is not None and pd.notna(avg_eff_team) else "N/A")
col4.metric("Avg Quality Score (team)", f"{avg_qs_team:.1%}" if avg_qs_team is not None and pd.notna(avg_qs_team) else "N/A")
col5.metric("Total Man-hours", f"{total_manhours:,}")

st.markdown("---")

# ---------- Individual KPI line charts ----------
st.header("Individual KPI Tracking — Full Timeline (per member)")
plot_line_generic(group_ind, 'month', 'QS_mean', "Average Quality Score (per member)", color_col='Name', is_pct=True)
plot_line_generic(group_ind, 'month', 'Rev_mean', "Average Revision Rate (per member)", color_col='Name', is_pct=True)
plot_line_generic(group_ind, 'month', 'Tasks', "Total Completed Tasks (per member)", color_col='Name', is_pct=False)
plot_line_generic(group_ind, 'month', 'OnTime_pct', "On-time Delivery (per member)", color_col='Name', is_pct=True)
plot_line_generic(group_ind, 'month', 'Eff_mean', "Actual Work Efficiency (per member)", color_col='Name', is_pct=True)
plot_line_generic(group_ind, 'month', 'Manhours', "Man-hours Spent (per member)", color_col='Name', is_pct=False)

st.markdown("---")

# ---------- Team KPI line charts ----------
st.header("Team KPI Tracking — Full Timeline (averaged)")
plot_line_generic(group_team, 'month', 'QS_mean', "Team Average Quality Score", is_pct=True)
plot_line_generic(group_team, 'month', 'Rev_mean', "Team Average Revision Rate", is_pct=True)
plot_line_generic(group_team, 'month', 'Tasks', "Team Total Completed Tasks", is_pct=False)
plot_line_generic(group_team, 'month', 'OnTime_pct', "Team On-time Delivery", is_pct=True)
plot_line_generic(group_team, 'month', 'Eff_mean', "Team Actual Work Efficiency", is_pct=True)
plot_line_generic(group_team, 'month', 'Manhours', "Team Man-hours Spent", is_pct=False)

st.markdown("---")

# ---------- Leaderboards (always visible) ----------
st.header("Leaderboards — Latest Month (filtered selection)")
if not group_ind.empty:
    latest_month = group_ind['month'].max()
    latest_df = group_ind[group_ind['month'] == latest_month].copy()
    if latest_df.empty:
        st.write("No data available for latest month after filtering.")
    else:
        lb_info = {
            "Efficiency (%)": ("Eff_mean", True),
            "Quality Score (%)": ("QS_mean", True),
            "On-time Delivery (%)": ("OnTime_pct", True),
            "Tasks Completed": ("Tasks", False),
            "Man-hours Spent": ("Manhours", False)
        }
        cols = st.columns(len(lb_info))
        for i, (title, (colname, is_pct)) in enumerate(lb_info.items()):
            df_lb = latest_df[['Name', colname]].sort_values(by=colname, ascending=False).reset_index(drop=True)
            display = df_lb.copy()
            if is_pct:
                display[colname] = display[colname].apply(lambda v: f"{v:.1%}" if pd.notna(v) else "N/A")
            else:
                display[colname] = display[colname].apply(lambda v: f"{int(v):,}" if pd.notna(v) else "N/A")
            cols[i].subheader(title)
            cols[i].dataframe(display, use_container_width=True)
else:
    st.write("Not enough data for leaderboards.")

st.markdown("---")

# ---------- Data inspector & download ----------
with st.expander("Show filtered raw data (expand)"):
    st.dataframe(flt.reset_index(drop=True))

csv = flt.to_csv(index=False).encode('utf-8')
st.download_button("Download filtered CSV", data=csv, file_name="filtered_kpis.csv", mime="text/csv")

st.markdown("Notes: This app attempts to automatically normalize percentage columns (e.g., `QS%`, `Revision/s`, `Efficiency`) whether they are entered as `0.92` or `92`. If a column looks wrong, tell me and I will add a manual toggle to force interpretation.")
