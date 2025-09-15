import streamlit as st
import pandas as pd
import plotly.express as px

# App title
st.title("📊 KPI Dashboard")

# File uploader
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file:
    # Read Excel
    df = pd.read_excel(uploaded_file)

    # Convert date column if available
    if "Date" in df.columns:
        df["Date"] = pd.to_datetime(df["Date"])

    # Sidebar for member selection
    members = df["Member"].unique()
    selected_member = st.sidebar.selectbox("Select a member", members)

    # Filter for selected member
    member_df = df[df["Member"] == selected_member]

    st.header(f"👤 {selected_member} - Individual KPI Tracking")

    # Line graphs for Individual
    metrics = {
        "Average Quality Score (%)": "QS%",
        "Average Revision Rate (%)": "RR%",
        "Total Completed Task": "Task Name",
        "On-time Delivery (%)": "OTD%",
        "Actual Work Efficiency (%)": "AWE%",
        "Man-hours Spent (Total)": "Mhrs"
    }

    for title, col in metrics.items():
        if col in member_df.columns:
            if col == "Task Name":  # count tasks
                chart_data = member_df.groupby("Date")[col].count().reset_index()
                y_label = "Total Tasks"
            elif col == "Mhrs":  # sum man-hours
                chart_data = member_df.groupby("Date")[col].sum().reset_index()
                y_label = "Man-hours"
            else:  # mean percentage metrics
                chart_data = member_df.groupby("Date")[col].mean().reset_index()
                y_label = title

            fig = px.line(chart_data, x="Date", y=col if col != "Task Name" else chart_data.columns[1],
                          markers=True, title=title)
            fig.update_layout(yaxis_title=y_label, xaxis_title="Date")
            st.plotly_chart(fig, use_container_width=True)

    # Team KPI Tracking
    st.header("👥 Team KPI Tracking")

    for title, col in metrics.items():
        if col in df.columns:
            if col == "Task Name":
                chart_data = df.groupby("Date")[col].count().reset_index()
                y_label = "Total Tasks"
            elif col == "Mhrs":
                chart_data = df.groupby("Date")[col].sum().reset_index()
                y_label = "Man-hours"
            else:
                chart_data = df.groupby("Date")[col].mean().reset_index()
                y_label = title

            fig = px.line(chart_data, x="Date", y=col if col != "Task Name" else chart_data.columns[1],
                          markers=True, title=title)
            fig.update_layout(yaxis_title=y_label, xaxis_title="Date")
            st.plotly_chart(fig, use_container_width=True)
