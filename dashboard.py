# dashboard.py
"""
KPI Dashboard - Line chart version (fixed)
"""

import re
from datetime import datetime
import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="KPI Dashboard (Line charts)", layout="wide")
st.title("KPI Dashboard — Line Charts")

# ---------------------
# Helpers
# ---------------------
@st.cache_data
def load_excel_anysheet(path_or_buffer):
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
    m = re.match(r"^(\d{8})$", s)
    if m:
        try:
            return pd.to_datetime(m.group(1), format="%Y%m%d")
        except Exception:
            pass
    fmts = ["%Y%m%d", "%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%d/%m/%Y", "%m/%d/%Y"]
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
    df.columns = [c.strip() for c in df.columns]

    if 'Date Completed' in df.columns:
        df['Date Completed'] = df['Date Completed'].apply(parse_numeric_yyyymmdd)

    if 'Work Duration' in df.columns:
        df = parse_work_duration_column(df, 'Work Duration')

    if 'end_date' in df.columns and 'Date Completed' in df.columns:
        df['end_date'] = df['end_date'].fillna(df['Date Completed'])

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

    numcols = ['Target Work Hours', 'Actual Work Hours', 'Efficiency', 'QS%', 'Revision/s']
    for c in numcols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce')

    if 'Date Completed' in df.columns and 'end_date' in df.columns:
        df['OnTime'] = (df['Date Completed'] <= df['end_date']).astype('Int64')
    else:
        df['OnTime'] = pd.NA

    if 'QS%' in df.columns:
        df['QS_frac'] = df['QS%'].copy()
        if df['QS_frac'].max(skipna=True) is not None and df['QS_frac'].max(skipna=True) > 1.5:
            df['QS_frac'] = df['QS_frac'] / 100.0
    if 'Revision/s' in df.columns:
        df['Rev_frac'] = df['Revision/s'].copy()
        if df['Rev_frac'].max(skipna=True) is not None and df['Rev_frac'].max(skipna=True) > 1.5:
            df['Rev_frac'] = df['Rev_frac'] / 100.0
    if 'Efficiency' in df.columns:
        df['Eff_frac'] = df['Efficiency'].copy()
        if df['Eff_frac'].max(skipna=True) is not None and df['Eff_frac'].max(skipna=True) > 1.5:
            df['Eff_frac'] = df['Eff_frac'] / 100.0

    if 'Actual Work Hours' in df.columns:
        df['Actual Work Hours'] = pd.to_numeric(df['Actual Work Hours'], errors='coerce').fillna(0)

    if 'Ref. number' in df.columns:
        df['_task_id'] = df['Ref. number']
    else:
        df['_task_id'] = range(len(df))
    return df

# ---------------------
# Load
# ---------------------
with st.sidebar:
    uploaded = st.file_uploader("Upload KPI Excel", type=['xlsx', 'xls'])

try:
    if uploaded is not None:
        raw = load_excel_anysheet(uploaded)
    else:
        raw = load_excel_anysheet("/mnt/data/KPI METRICS 2.xlsx")
except Exception as e:
    st.error(f"Failed to load Excel file: {e}")
    st.stop()

if raw is None or raw.empty:
    st.error("No data found in the file.")
    st.stop()

df = clean_and_prepare(raw)

# Filters
members = sorted(df['Name'].dropna().unique().tolist()) if 'Name' in df.columns else []
sel_members = st.sidebar.multiselect("Members", options=members, default=members)
months = sorted(df['month'].dropna().unique().tolist())
sel_months = st.sidebar.multiselect("Months", options=months, default=months)

flt = df.copy()
if sel_members:
    flt = flt[flt['Name'].isin(sel_members)]
if sel_months:
    flt = flt[flt['month'].isin(sel_months)]

if flt.empty:
    st.warning("No data after filtering.")
    st.stop()

# Aggregations
group_ind = flt.groupby(['month', 'Name'], as_index=False).agg(
    QS_mean=('QS_frac', 'mean'),
    Rev_mean=('Rev_frac', 'mean'),
    OnTime_pct=('OnTime', 'mean'),
    Eff_mean=('Eff_frac', 'mean'),
    Manhours=('Actual Work Hours', 'sum'),
    Tasks=('_task_id', 'count')
)
group_team = flt.groupby(['month'], as_index=False).agg(
    QS_mean=('QS_frac', 'mean'),
    Rev_mean=('Rev_frac', 'mean'),
    OnTime_pct=('OnTime', 'mean'),
    Eff_mean=('Eff_frac', 'mean'),
    Manhours=('Actual Work Hours', 'sum'),
    Tasks=('_task_id', 'count')
)

# ---------------------
# Plotting
# ---------------------
def plot_line(df_plot, x_col, y_col, title, is_pct=False):
    if df_plot.empty:
        st.write(f"No data for {title}")
        return
    if "Name" in df_plot.columns:
        fig = px.line(df_plot, x=x_col, y=y_col, color='Name', markers=True, title=title)
    else:
        fig = px.line(df_plot, x=x_col, y=y_col, markers=True, title=title)
    if is_pct:
        fig.update_yaxes(tickformat=".0%")
    st.plotly_chart(fig, use_container_width=True)

# Individual KPI line charts
st.header("Individual KPI Tracking (per member)")
plot_line(group_ind, 'month', 'QS_mean', 'Quality Score', True)
plot_line(group_ind, 'month', 'Rev_mean', 'Revision Rate', True)
plot_line(group_ind, 'month', 'Tasks', 'Tasks Completed')
plot_line(group_ind, 'month', 'OnTime_pct', 'On-time Delivery', True)
plot_line(group_ind, 'month', 'Eff_mean', 'Efficiency', True)
plot_line(group_ind, 'month', 'Manhours', 'Man-hours Spent')

# Team KPI line charts
st.header("Team KPI Tracking (averaged)")
plot_line(group_team, 'month', 'QS_mean', 'Team Quality Score', True)
plot_line(group_team, 'month', 'Rev_mean', 'Team Revision Rate', True)
plot_line(group_team, 'month', 'Tasks', 'Team Tasks Completed')
plot_line(group_team, 'month', 'OnTime_pct', 'Team On-time Delivery', True)
plot_line(group_team, 'month', 'Eff_mean', 'Team Efficiency', True)
plot_line(group_team, 'month', 'Manhours', 'Team Man-hours Spent')
