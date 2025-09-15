# dashboard.py
"""
KPI Dashboard - Reverted-style app but with LINE CHARTS.
Reads an Excel with the same structure you've been using.
- Handles Work Duration parsing (start/end)
- Computes On-time (Date Completed <= Work Duration end)
- Normalizes QS%, Revision/s, Efficiency to fraction (0-1) when appropriate
- Shows line charts per member and team-aggregate line charts
- Shows leaderboards side-by-side for the latest month
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
    # Try reading sensible sheet names: '5', '1', else first sheet
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
        # fallback: try direct read (may raise)
        df = pd.read_excel(path_or_buffer)
    return df

def parse_numeric_yyyymmdd(x):
    if pd.isna(x):
        return pd.NaT
    s = str(x).strip()
    # If numeric like 20250703 or float like 20250703.0
    m = re.match(r"^(\d{8})$", s)
    if m:
        try:
            return pd.to_datetime(m.group(1), format="%Y%m%d")
        except Exception:
            pass
    # try common formats
    fmts = ["%Y%m%d", "%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%d/%m/%Y", "%m/%d/%Y"]
    for f in fmts:
        try:
            return pd.to_datetime(s, format=f)
        except Exception:
            continue
    # pandas fallback
    return pd.to_datetime(s, errors='coerce')

def parse_work_duration_column(df, col="Work Duration"):
    # parse strings like "20250623-20250704" or "2025/06/23 - 2025/07/04"
    starts = []
    ends = []
    for v in df.get(col, pd.Series([pd.NA]*len(df))):
        if pd.isna(v):
            starts.append(pd.NaT); ends.append(pd.NaT); continue
        s = str(v).strip()
        # common delimiters
        parts = re.split(r"\s*[-–—]\s*|\s+to\s+|\s*/\s*", s)
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
    # standardize column names whitespace
    df.columns = [c.strip() for c in df.columns]

    # parse Date Completed column (handles numeric yyyymmdd)
    if 'Date Completed' in df.columns:
        df['Date Completed'] = df['Date Completed'].apply(parse_numeric_yyyymmdd)

    # parse Work Duration into start_date/end_date
    if 'Work Duration' in df.columns:
        df = parse_work_duration_column(df, 'Work Duration')

    # If end_date missing but Date Completed present, use Date Completed as end_date
    if 'end_date' in df.columns and 'Date Completed' in df.columns:
        df['end_date'] = df['end_date'].fillna(df['Date Completed'])

    # month bucket for grouping: prefer end_date then start_date then Date Completed
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

    # ensure numeric columns
    numcols = ['Target Work Hours', 'Actual Work Hours', 'Efficiency', 'QS%', 'Revision/s']
    for c in numcols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors='coerce')

    # compute OnTime (1 if Date Completed <= end_date)
    if 'Date Completed' in df.columns and 'end_date' in df.columns:
        df['OnTime'] = (df['Date Completed'] <= df['end_date']).astype('Int64')
    else:
        df['OnTime'] = pd.NA

    # Create normalized percentage columns (as 0..1)
    # QS%
    if 'QS%' in df.columns:
        df['QS_frac'] = df['QS%'].copy()
        # if values look like whole percentages (>1.5 likely means percent points), divide by 100
        if df['QS_frac'].max(skipna=True) is not None and df['QS_frac'].max(skipna=True) > 1.5:
            df['QS_frac'] = df['QS_frac'] / 100.0

    # Revision/s -> treat as revision rate if in 0..1; if large numbers but <=100, try convert /100
    if 'Revision/s' in df.columns:
        df['Rev_frac'] = df['Revision/s'].copy()
        if df['Rev_frac'].max(skipna=True) is not None and df['Rev_frac'].max(skipna=True) > 1.5:
            df['Rev_frac'] = df['Rev_frac'] / 100.0

    # Efficiency -> if values appear as percentages >1.5 (e.g., 98), convert by /100
    if 'Efficiency' in df.columns:
        df['Eff_frac'] = df['Efficiency'].copy()
        if df['Eff_frac'].max(skipna=True) is not None and df['Eff_frac'].max(skipna=True) > 1.5:
            df['Eff_frac'] = df['Eff_frac'] / 100.0

    # ensure Actual Work Hours numeric
    if 'Actual Work Hours' in df.columns:
        df['Actual Work Hours'] = pd.to_numeric(df['Actual Work Hours'], errors='coerce').fillna(0)

    # create a consistent task identifier column for counting
    if 'Ref. number' in df.columns:
        df['_task_id'] = df['Ref. number']
    else:
        df['_task_id'] = range(len(df))

    return df

# ---------------------
# UI & Data load
# ---------------------
with st.sidebar:
    st.write("Upload Excel (optional). If none, app will try default /mnt/data/KPI METRICS 2.xlsx on server.")
    uploaded = st.file_uploader("KPI Excel", type=['xlsx', 'xls'])

# load
try:
    if uploaded is not None:
        raw = load_excel_anysheet(uploaded)
    else:
        # default path used in your environment earlier
        raw = load_excel_anysheet("/mnt/data/KPI METRICS 2.xlsx")
except Exception as e:
    st.error(f"Failed to load Excel file: {e}")
    st.stop()

if raw is None or raw.empty:
    st.error("No data found in the file.")
    st.stop()

df = clean_and_prepare(raw)

# show top columns and quick info
st.sidebar.write("Rows loaded:", len(df))
st.sidebar.write("Columns:", list(df.columns))

# member filter
members = sorted(df['Name'].dropna().unique().tolist()) if 'Name' in df.columns else []
selected_members = st.sidebar.multiselect("Team member(s)", options=members, default=members)

# month filter (optional)
months = sorted(df['month'].dropna().unique().tolist())
selected_months = st.sidebar.multiselect("Month(s)", options=months, default=months)

# project filter
projects = sorted(df['Project Involvement'].dropna().unique().tolist()) if 'Project Involvement' in df.columns else []
selected_projects = st.sidebar.multiselect("Project(s)", options=projects, default=projects)

# apply filters
filtered = df.copy()
if selected_members:
    filtered = filtered[filtered['Name'].isin(selected_members)]
if selected_months:
    filtered = filtered[filtered['month'].isin(selected_months)]
if selected_projects:
    filtered = filtered[filtered['Project Involvement'].isin(selected_projects)]

if filtered.empty:
    st.warning("No rows after applying filters. Adjust filters or upload a different file.")
    st.stop()

# ---------------------
# Aggregations (per month & per member)
# ---------------------
agg_cols = {}

# Individual grouping: month + Name
group_ind = filtered.groupby(['month', 'Name'], as_index=False).agg(
    QS_mean = ('QS_frac', 'mean') if 'QS_frac' in filtered.columns else ('QS%', 'mean'),
    Rev_mean = ('Rev_frac', 'mean') if 'Rev_frac' in filtered.columns else ('Revision/s', 'mean'),
    OnTime_pct = ('OnTime', 'mean'),
    Eff_mean = ('Eff_frac', 'mean') if 'Eff_frac' in filtered.columns else ('Efficiency', 'mean'),
    Manhours = ('Actual Work Hours', 'sum'),
    Tasks = ('_task_id', 'count')
)

# Team grouping: month only (averaged across all members)
group_team = filtered.groupby(['month'], as_index=False).agg(
    QS_mean = ('QS_frac', 'mean') if 'QS_frac' in filtered.columns else ('QS%', 'mean'),
    Rev_mean = ('Rev_frac', 'mean') if 'Rev_frac' in filtered.columns else ('Revision/s', 'mean'),
    OnTime_pct = ('OnTime', 'mean'),
    Eff_mean = ('Eff_frac', 'mean') if 'Eff_frac' in filtered.columns else ('Efficiency', 'mean'),
    Manhours = ('Actual Work Hours', 'sum'),
    Tasks = ('_task_id', 'count')
)

# convert month to proper x-axis dtype
if 'month' in group_ind.columns:
    group_ind = group_ind.sort_values('month')
if 'month' in group_team.columns:
    group_team = group_team.sort_values('month')

# ---------------------
# Top KPI metrics (wide)
# ---------------------
st.subheader("Top
