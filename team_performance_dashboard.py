import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.backends.backend_pdf import PdfPages
import io
import re

# App configuration
st.set_page_config(layout="wide")
st.title("📊 Team Performance Dashboard")
st.write("Upload the performance Excel sheet to generate visual analysis per team.")

# File upload
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    # Load and reshape data
    df = pd.read_excel(uploaded_file, engine="openpyxl")
    df = df.set_index("Row Labels").T
    df.index.name = "Sprint"

    # Extract team names and metric mapping using regex
    pattern = r"^(.*?)\s{1,2}([^\(]+(?:\(.*\))?)$"
    team_metric_map = {}

    for col in df.columns:
        match = re.match(pattern, col.strip())
        if match:
            team = match.group(1).strip()
            metric = match.group(2).strip()
            team_metric_map.setdefault(team, {})[metric] = col

    # User selects team
    selected_team = st.selectbox("Select Team to Visualize", sorted(team_metric_map.keys()))

    if selected_team:
        # Define standard metric mapping
        rename_map = {
            "Velocity (SP)": "Velocity",
            "Billable TS": "Billable TS",
            "Investment TS": "Non-Billable TS",
            "Non-Billable TS": "Non-Billable TS",
            "Bugs created": "Bugs Created",
            "Bugs closed": "Bugs Closed",
            "# Bugs created": "Bugs Created",
            "# Bugs closed": "Bugs Closed",
            "# Releases in Prod": "Releases",
            "SP to Hour Ratio": "SP/Hour"
        }

        # Get original column names for selected team
        team_metrics = team_metric_map[selected_team]

        # Filter for metrics that match rename_map
        selected_cols = {rename_map[k]: v for k, v in team_metrics.items() if k.strip() in rename_map}
        team_df = df[selected_cols.values()].copy()
        team_df.columns = selected_cols.keys()
        team_df = team_df.apply(pd.to_numeric, errors='coerce')

        # Add calculated columns
        if "Billable TS" in team_df.columns and "Non-Billable TS" in team_df.columns:
            team_df["Total Time"] = team_df["Billable TS"] + team_df["Non-Billable TS"]
            team_df["Efficiency"] = team_df["Velocity"] / (team_df["Billable TS"] / 100)

        # Display data
        st.subheader(f"{selected_team} - Metrics Overview")
        st.dataframe(team_df.style.format("{:.2f}"))

        # Create charts
        sns.set(style="whitegrid")
        fig, axs = plt.subplots(4, 1, figsize=(15, 20))
        fig.suptitle(f"{selected_team} - Performance Dashboard", fontsize=18, weight='bold')

        # Plot 1 - Velocity
        if "Velocity" in team_df.columns:
            team_df["Velocity"].plot(ax=axs[0], marker='o', color='blue', title="Velocity Trend")
            axs[0].set_ylabel("Story Points")

        # Plot 2 - Releases
        if "Releases" in team_df.columns:
            team_df["Releases"].plot(kind='bar', ax=axs[1], color='orange', title="Releases per Sprint")
            axs[1].set_ylabel("Releases")
            axs[1].set_xlabel("Sprint")
            axs[1].tick_params(axis='x', rotation=45)

        # Plot 3 - Bugs Created/Closed
        if "Bugs Created" in team_df.columns and "Bugs Closed" in team_df.columns:
            team_df[["Bugs Created", "Bugs Closed"]].plot(kind='bar', ax=axs[2], stacked=False, title="Bugs Created vs Closed")
            axs[2].set_ylabel("Bug Count")
            axs[2].set_xlabel("Sprint")
            axs[2].tick_params(axis='x', rotation=45)

        # Plot 4 - SP/Hour
        if "SP/Hour" in team_df.columns:
            team_df["SP/Hour"].plot(ax=axs[3], marker='D', color='purple', title="SP to Hour Ratio")
            axs[3].set_ylabel("SP/Hour")
            axs[3].set_xlabel("Sprint")
            axs[3].tick_params(axis='x', rotation=45)

        plt.tight_layout(rect=[0, 0.03, 1, 0.95])
        st.pyplot(fig)

        # Export to PDF
        pdf_buffer = io.BytesIO()
        pdf = PdfPages(pdf_buffer)
        pdf.savefig(fig)
        plt.close()
        pdf.close()

        # PDF Download
        st.download_button(
            label="📥 Download PDF Report",
            data=pdf_buffer.getvalue(),
            file_name=f"{selected_team.replace(' ', '_')}_Dashboard.pdf",
            mime="application/pdf"
        )
