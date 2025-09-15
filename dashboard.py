# dashboard.py
"""
KPI Dashboard for team task metrics (Streamlit app)

Features:
- Upload or use default Excel file (looks for sheet '1' or first sheet)
- Cleans columns, parses Work Duration and Date Completed
- Filters: team member(s), month/year, project, status
- Charts:
    - Overview KPIs (total tasks, avg efficiency, avg quality)
    - Target vs Actual hours (grouped by person or task)
    - Efficiency trend over time
    - Tasks completed per month
    - Quality score vs Revisions scatter
- Table view and CSV download of filtered data
- Interactive (Plotly)
"""

import io
from datetime import datetime
from typing import Optional

import pandas as pd
import plotly.express as px
import streamlit as st

# ---------- Helpers ----------

@st.cache_data
def load_excel(path_or_buffer) -> pd.DataFrame:
    # Try reading sheet '1' else first sheet
    try:
        xls = pd.ExcelFile(path_or_buffer)
        sheet = '1' if '1' in xls.sheet_names else xls.sheet_names[0]
        df = pd.read_excel(xls, sheet_name=sheet)
    except Exception as e:
        # If given buffer that is not a path, try direct read
        df = pd.read_excel(path_or_buffer)
    return df

def parse_work_duration(df: pd.DataFrame, col='Work Duration') -> pd.DataFrame:
    # Expected format examples: "20250623-20250704" or "2025/06/23 - 2025/07/04"
    def parse_range(s):
        if pd.isna(s):
            return (pd.NaT, pd.NaT)
        s = str(s).strip()
        # common delimiters
        for delim in [' - ', '-', '—', '–']:
            if delim in s:
                parts = [p.strip() for p in s.split(delim)]
                if len(parts) >= 2:
                    return to_date(parts[0]), to_date(parts[1])
        # fallback single date
        return to_date(s), pd.NaT

    def to_date(x):
        if pd.isna(x):
            return pd.NaT
        x = str(x).strip()
        # try few formats
        fmts = ['%Y%m%d', '%Y-%m-%d', '%Y/%m/%d', '%Y.%m.%d', '%d/%m/%Y', '%m/%d/%Y']
        for f in fmts:
            try:
                return datetime.strptime(x, f).date()
            except Exception:
                continue
        # try parsing with pandas
        try:
            return pd.to_datetime(x, errors='coerce').date()
        except Exception:
            return pd.NaT

    start_dates = []
    end_dates = []
    for v in df.get(col, pd.Series([])):
        s, e = parse_range(v)
        start_dates.append(s)
        end_dates.append(e)
    df = df.copy()
    df['start_date'] = pd.to_datetime(start_dates)
    df['end_date'] = pd.to_datetime(end_dates)
    # If no end_date, but Date Completed present, use that as end_date
    if 'Date Completed' in df.columns:
        # Date Completed may be numeric "20250703" -> convert
        df['Date Completed'] = df['Date Completed'].apply(parse_numeric_yyyymmdd)
        df['Date Completed'] = pd.to_datetime(df['Date Completed'], errors='coerce')
        df['end_date'] = df['end_date'].fillna(df['Date Completed'])
    # For grouping convenience:
    df['year_month'] = df['start_date'].dt.to_period('M').astype(str)
    return df

def parse_numeric_yyyymmdd(x):
    if pd.isna(x):
        return pd.NaT
    s = str(x).strip()
    if s.isnumeric() and len(s) == 8:
        try:
            return datetime.strptime(s, '%Y%m%d').date()
        except Exception:
            pass
    # fallback try to parse with pandas
    try:
        return pd.to_datetime(s, errors='coerce').date()
    except Exception:
        return pd.NaT

def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    # normalize column names
    df = df.copy()
    df.columns = [c.strip() for c in df.columns]
    # convert numeric-looking strings
    numeric_cols = ['Target Work Hours', 'Actual Work Hours', 'Efficiency', 'Quality Score', 'QS%']
    for c in numeric_cols:
        if c in df.columns:
            # remove percent sign or weird chars
            df[c] = pd.to_numeric(df[c].astype(str).str.replace('%', '').str.replace(',', '').str.strip(), errors='coerce')
    # % Accomplishment may be decimal or 1/0 or percent
    if '% Accomplishment' in df.columns:
        df['% Accomplishment'] = pd.to_numeric(df['% Accomplishment'].astype(str).str.replace('%','').str.strip(), errors='coerce')
    # Revisions often numeric
    if 'Revision/s' in df.columns:
        df['Revision/s'] = pd.to_numeric(df['Revision/s'], errors='coerce').fillna(0).astype(int)
    return df

def compute_summary(df: pd.DataFrame):
    total_tasks = len(df)
    avg_eff = df['Efficiency'].dropna().mean() if 'Efficiency' in df.columns else None
    avg_qs = df['Quality Score'].dropna().mean() if 'Quality Score' in df.columns else None
    completed_tasks = df[df['Status'].str.lower().eq('completed')].shape[0] if 'Status' in df.columns else None
    return {
        'total_tasks': total_tasks,
        'completed_tasks': completed_tasks,
        'avg_eff': avg_eff,
        'avg_qs': avg_qs
    }

def download_link(df: pd.DataFrame, filename='filtered_kpis.csv'):
    csv = df.to_csv(index=False).encode('utf-8')
    return csv

# ---------- Streamlit UI ----------

st.set_page_config(page_title="KPI Dashboard", layout="wide", initial_sidebar_state="expanded")

st.title("KPI Dashboard — Team Task Tracking")
st.markdown("Interactive dashboard for KPI metrics. Upload an Excel file or use the default provided.")

# Sidebar: file uploader and filters
with st.sidebar:
    st.header("Data Input & Filters")
    uploaded_file = st.file_uploader("Upload Excel (same structure)", type=['xlsx', 'xls'])
    use_default = False
    if uploaded_file is None:
        st.info("No file uploaded — using default dataset if available on /mnt/data/KPI METRICS 1.xlsx")
        use_default = True

    # load data (cached)
    try:
        if uploaded_file is not None:
            df_raw = load_excel(uploaded_file)
        else:
            # default path - update if necessary
            default_path = "/mnt/data/KPI METRICS 1.xlsx"
            df_raw = load_excel(default_path)
    except Exception as e:
        st.error(f"Failed to load Excel: {e}")
        st.stop()

    st.write("Loaded rows:", len(df_raw))

    # Clean and parse
    df_clean = clean_columns(df_raw)
    df_clean = parse_work_duration(df_clean)

    # Provide available filter values
    names = sorted(df_clean['Name'].dropna().unique().tolist()) if 'Name' in df_clean.columns else []
    projects = sorted(df_clean['Project Involvement'].dropna().unique().tolist()) if 'Project Involvement' in df_clean.columns else []
    months = sorted(df_clean['year_month'].dropna().unique().tolist())
    status_vals = sorted(df_clean['Status'].dropna().unique().tolist()) if 'Status' in df_clean.columns else []

    selected_names = st.multiselect("Team member(s)", options=names, default=names)
    selected_projects = st.multiselect("Project(s)", options=projects, default=projects)
    selected_months = st.multiselect("Month(s) (YYYY-MM)", options=months, default=months)
    selected_status = st.multiselect("Status", options=status_vals, default=status_vals)

    # Apply filters
    df_filtered = df_clean.copy()
    if selected_names:
        df_filtered = df_filtered[df_filtered['Name'].isin(selected_names)]
    if selected_projects:
        df_filtered = df_filtered[df_filtered['Project Involvement'].isin(selected_projects)]
    if selected_months:
        df_filtered = df_filtered[df_filtered['year_month'].isin(selected_months)]
    if selected_status and 'Status' in df_filtered.columns:
        df_filtered = df_filtered[df_filtered['Status'].isin(selected_status)]

    st.markdown("---")
    st.write("Filtered rows:", len(df_filtered))

    # Option to download filtered CSV
    csv_bytes = download_link(df_filtered)
    st.download_button("Download filtered CSV", data=csv_bytes, file_name="filtered_kpis.csv", mime="text/csv")

# ---------- Main layout ----------

# Top-level KPIs
summary = compute_summary(df_filtered)
col1, col2, col3, col4 = st.columns(4)
col1.metric("Total tasks (shown)", summary['total_tasks'])
col2.metric("Completed tasks", summary['completed_tasks'] if summary['completed_tasks'] is not None else "N/A")
col3.metric("Avg Efficiency", f"{summary['avg_eff']:.2f}" if summary['avg_eff'] is not None and not pd.isna(summary['avg_eff']) else "N/A")
col4.metric("Avg Quality Score", f"{summary['avg_qs']:.2f}" if summary['avg_qs'] is not None and not pd.isna(summary['avg_qs']) else "N/A")

st.markdown("----")

# Row: Target vs Actual hours (grouped by person)
st.subheader("Target vs Actual Work Hours")
if 'Target Work Hours' in df_filtered.columns and 'Actual Work Hours' in df_filtered.columns:
    agg = df_filtered.groupby('Name', as_index=False)[['Target Work Hours', 'Actual Work Hours']].sum().reset_index(drop=True)
    if agg.shape[0] == 0:
        st.write("No data to plot.")
    else:
        fig1 = px.bar(agg.melt(id_vars=['Name'], value_vars=['Target Work Hours','Actual Work Hours'],
                               var_name='Type', value_name='Hours'),
                      x='Name', y='Hours', color='Type', barmode='group',
                      title='Target vs Actual Hours by Person')
        st.plotly_chart(fig1, use_container_width=True)
else:
    st.write("Target Work Hours and/or Actual Work Hours columns not found in data.")

# Row: Efficiency trend
st.subheader("Efficiency Trend Over Time")
if 'start_date' in df_filtered.columns and 'Efficiency' in df_filtered.columns:
    # Use monthly average of Efficiency
    temp = df_filtered.dropna(subset=['start_date', 'Efficiency']).copy()
    if temp.empty:
        st.write("No efficiency/date data to plot.")
    else:
        temp['month'] = temp['start_date'].dt.to_period('M').dt.to_timestamp()
        eff_ts = temp.groupby(['month','Name'], as_index=False)['Efficiency'].mean()
        fig2 = px.line(eff_ts, x='month', y='Efficiency', color='Name', markers=True, title='Efficiency by Month (avg)')
        fig2.update_layout(xaxis_title='Month', yaxis_title='Efficiency')
        st.plotly_chart(fig2, use_container_width=True)
else:
    st.write("Missing Efficiency or start_date for trend.")

# Row: Tasks completed per month
st.subheader("Tasks Completed per Month")
if 'Status' in df_filtered.columns and 'end_date' in df_filtered.columns:
    completed = df_filtered[df_filtered['Status'].str.lower() == 'completed'].copy()
    if completed.empty:
        st.write("No completed tasks in filtered dataset.")
    else:
        completed['completed_month'] = completed['end_date'].dt.to_period('M').astype(str)
        monthly = completed.groupby(['completed_month','Name'], as_index=False).size().rename(columns={'size':'count'})
        fig3 = px.bar(monthly, x='completed_month', y='count', color='Name', title='Completed Tasks per Month')
        fig3.update_layout(xaxis_title='Month (YYYY-MM)', yaxis_title='Completed tasks')
        st.plotly_chart(fig3, use_container_width=True)
else:
    st.write("Missing Status and/or end_date for completed task chart.")

# Row: Quality Score vs Revisions
st.subheader("Quality Score vs Revision/s")
if 'Quality Score' in df_filtered.columns and 'Revision/s' in df_filtered.columns:
    qdf = df_filtered.dropna(subset=['Quality Score']).copy()
    if qdf.empty:
        st.write("No quality score data to plot.")
    else:
        fig4 = px.scatter(qdf, x='Revision/s', y='Quality Score', color='Name',
                          hover_data=['Assigned Tasks (Summary)', 'Project Involvement', 'Ref. number'],
                          title='Quality Score vs Number of Revisions')
        st.plotly_chart(fig4, use_container_width=True)
else:
    st.write("Quality Score and/or Revision/s columns not available.")

st.markdown("----")

# Optional Gantt-like view (timeline)
st.subheader("Task Timeline (Gantt-like)")
if 'start_date' in df_filtered.columns and 'end_date' in df_filtered.columns:
    timeline_df = df_filtered.dropna(subset=['start_date']).copy()
    # If end_date empty, use start_date + 1 day for plotting
    timeline_df['end_for_plot'] = timeline_df['end_date'].fillna(timeline_df['start_date'] + pd.Timedelta(days=1))
    # Plot only some columns
    try:
        fig5 = px.timeline(timeline_df.sort_values('start_date'), x_start='start_date', x_end='end_for_plot',
                           y='Name', color='Project Involvement', hover_data=['Assigned Tasks (Summary)', 'Status'])
        fig5.update_yaxes(autorange="reversed")
        fig5.update_layout(title='Task timeline by person')
        st.plotly_chart(fig5, use_container_width=True)
    except Exception as e:
        st.write("Timeline plot failed:", e)
else:
    st.write("Start and end date not available to plot timeline.")

st.markdown("----")

# Data explorer
st.subheader("Filtered Data (table)")
st.dataframe(df_filtered.reset_index(drop=True))

# Allow user to download the current filtered dataset again
csv2 = df_filtered.to_csv(index=False).encode('utf-8')
st.download_button("Download filtered dataset (CSV)", data=csv2, file_name="filtered_kpis.csv", mime="text/csv")

st.markdown("---")
st.markdown("**Notes & tips**:\n"
            "- If your Excel has the same column names used in this app, everything will work smoothly.\n"
            "- `Work Duration` values like `20250623-20250704` will be parsed into start/end dates.\n"
            "- `Date Completed` values like `20250703` (YYYYMMDD numeric) will be parsed.\n"
            "- You can modify the code to add more charts or export a PowerPoint.\n")

