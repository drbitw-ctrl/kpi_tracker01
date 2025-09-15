import streamlit as st
import pandas as pd
import plotly.express as px

# =========================
# Load Data
# =========================
@st.cache_data
def load_data(file):
    df = pd.read_excel(file, sheet_name="5")
    # Ensure date columns are parsed
    df['Date Completed'] = pd.to_datetime(df['Date Completed'], errors='coerce')
    df['Work Duration'] = pd.to_datetime(df['Work Duration'], errors='coerce')
    return df

st.title("📊 KPI Dashboard")

uploaded_file = st.file_uploader("Upload KPI Excel file", type=["xlsx"])
if uploaded_file:
    df = load_data(uploaded_file)

    # =========================
    # Data Preprocessing
    # =========================
    df['Month'] = df['Date Completed'].dt.to_period("M").astype(str)

    # On-time Delivery (1 = on time, 0 = late)
    df['OnTime'] = (df['Date Completed'] <= df['Work Duration']).astype(int)

    # Convert metrics
    df['QS%'] = pd.to_numeric(df['QS%'], errors='coerce')
    df['Efficiency'] = pd.to_numeric(df['Efficiency'], errors='coerce')
    df['RevisionRate'] = pd.to_numeric(df['Revision/s'], errors='coerce')

    # =========================
    # Sidebar Filters
    # =========================
    members = df['Name'].dropna().unique()
    selected_members = st.sidebar.multiselect("Select Team Member(s)", members, default=members)

    filtered_df = df[df['Name'].isin(selected_members)]

    # =========================
    # Individual KPI Tracking
    # =========================
    st.header("📈 Individual KPI Tracking (per member)")

    # Aggregate metrics
    ind_group = (
        filtered_df.groupby(["Month", "Name"])
        .agg({
            "QS%": "mean",
            "RevisionRate": "mean",
            "OnTime": "mean",
            "Efficiency": "mean",
            "Actual Work Hours": "sum"
        })
        .reset_index()
    )

    # Add task counts
    task_counts = (
        filtered_df.groupby(["Month", "Name"])
        .size()
        .reset_index(name="Task Count")
    )
    ind_group = ind_group.merge(task_counts, on=["Month", "Name"])

    kpi_mapping = {
        "Average Quality Score (%)": "QS%",
        "Average Revision Rate (%)": "RevisionRate",
        "Total Completed Tasks": "Task Count",
        "On-time Delivery (%)": "OnTime",
        "Actual Work Efficiency (%)": "Efficiency",
        "Man-hours Spent (total)": "Actual Work Hours"
    }

    # Plot individual KPIs
    for title, col in kpi_mapping.items():
        fig = px.line(
            ind_group,
            x="Month",
            y=col,
            color="Name",
            markers=True,
            title=title
        )
        if "%" in title:
            fig.update_layout(yaxis_tickformat=".0%")
        st.plotly_chart(fig, use_container_width=True)

    # =========================
    # Team KPI Tracking (averages)
    # =========================
    st.header("📊 Team KPI Tracking (averaged across members)")

    team_group = (
        filtered_df.groupby("Month")
        .agg({
            "QS%": "mean",
            "RevisionRate": "mean",
            "OnTime": "mean",
            "Efficiency": "mean",
            "Actual Work Hours": "sum"
        })
        .reset_index()
    )

    # Add task counts
    task_counts_team = (
        filtered_df.groupby("Month")
        .size()
        .reset_index(name="Task Count")
    )
    team_group = team_group.merge(task_counts_team, on="Month")

    for title, col in kpi_mapping.items():
        fig = px.line(
            team_group,
            x="Month",
            y=col,
            markers=True,
            title=f"Team {title}"
        )
        if "%" in title:
            fig.update_layout(yaxis_tickformat=".0%")
        st.plotly_chart(fig, use_container_width=True)

    # =========================
    # Leaderboards
    # =========================
    st.header("🏆 Leaderboards (Latest Month)")

    latest_month = ind_group['Month'].max()
    latest_data = ind_group[ind_group['Month'] == latest_month].copy()

    leaderboard_cols = {
        "Efficiency (%)": "Efficiency",
        "Quality Score (%)": "QS%",
        "On-time Delivery (%)": "OnTime",
        "Tasks Completed": "Task Count",
        "Man-hours Spent": "Actual Work Hours"
    }

    cols = st.columns(len(leaderboard_cols))
    for i, (title, col) in enumerate(leaderboard_cols.items()):
        lb = latest_data[["Name", col]].sort_values(by=col, ascending=False)
        if "%" in title:
            lb[col] = lb[col].apply(lambda x: f"{x:.1%}")
        else:
            lb[col] = lb[col].apply(lambda x: f"{x:.0f}")
        cols[i].subheader(title)
        cols[i].dataframe(lb.reset_index(drop=True), use_container_width=True)
