import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.backends.backend_pdf import PdfPages
import io

# App configuration
st.set_page_config(layout="wide")
st.title("📊 Team Performance Dashboard")
st.write("Upload the performance Excel sheet to generate visual analysis per team.")

# Upload Excel file
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    # Load and reshape data
    df = pd.read_excel(uploaded_file, engine="openpyxl")
    df = df.set_index("Row Labels").T
    df.index.name = "Sprint"
    st.write("### 🔍 Available Columns in Uploaded Data")
    st.write(list(df.columns))  # Debug: Show actual column headers

    # Updated team metric mappings based on actual Excel sheet
    teams = {
        "Agency (Team Ephesus)": [
            "Agency (Team Ephesus) Velocity (SP)", "Agency (Team Ephesus) Billable hours", "Agency (Team Ephesus) Investment hours",
            "Agency (Team Ephesus) # Bugs created", "Agency (Team Ephesus) # Bugs closed", "Agency (Team Ephesus) # Releases in Prod",
            "Agency (Team Ephesus) SP to Hour Ratio"
        ],
        "Agency (Team Cyclopes)": [
            "Agency (Team Cyclopes) Velocity (SP)", "Agency (Team Cyclopes) Billable hours", "Agency (Team Cyclopes) Investment hours",
            "Agency (Team Cyclopes) # Bugs created", "Agency (Team Cyclopes) # Bugs closed", "Agency (Team Cyclopes) # Releases in Prod",
            "Agency (Team Cyclopes) SP to Hour Ratio"
        ],
        "Production Systems": [
            "Production Systems Velocity (SP)", "Production Systems Billable hours", "Production Systems Investment hours",
            "Production Systems # Bugs created", "Production Systems # Bugs closed", "Production Systems # Releases in Prod",
            "Production Systems SP to Hour Ratio"
        ],
        "TPS & DP": [
            "TPS & DP Velocity (SP)", "TPS & DP Billable hours", "TPS & DP Investment hours",
            "TPS & DP # Bugs created", "TPS & DP # Bugs closed", "TPS & DP # Releases in Prod",
            "TPS & DP SP to Hour Ratio"
        ]
    }

    pdf_buffer = io.BytesIO()
    pdf = PdfPages(pdf_buffer)

    # Team selection
    selected_team = st.selectbox("Select Team to Visualize", list(teams.keys()))

    if selected_team:
        # Prepare data
        cols = teams[selected_team]
        team_df = df[cols].copy()
        team_df.columns = ["Velocity", "Billable TS", "Non-Billable TS", "Bugs Created", "Bugs Closed", "Releases", "SP/Hour"]
        team_df = team_df.apply(pd.to_numeric, errors='coerce')

        # Optional metrics
        team_df["Total Time"] = team_df["Billable TS"] + team_df["Non-Billable TS"]
        team_df["Efficiency"] = team_df["Velocity"] / (team_df["Billable TS"] / 100)

        # Display table
        st.subheader(f"{selected_team} - Metrics Overview")
        st.dataframe(team_df.style.format("{:.2f}"))

        # Create charts
        sns.set(style="whitegrid")
        fig, axs = plt.subplots(4, 1, figsize=(15, 20))
        fig.suptitle(f"{selected_team} - Performance Dashboard", fontsize=18, weight='bold')

        # 1. Velocity
        team_df["Velocity"].plot(ax=axs[0], marker='o', color='blue', title=f"{selected_team} - Velocity Trend")
        axs[0].set_ylabel("Story Points")

        # 2. Bugs created/closed
        team_df[["Bugs Created", "Bugs Closed"]].plot(kind='bar', ax=axs[2], title=f"{selected_team} - Bugs Trend", stacked=False)
        axs[2].set_ylabel("Count")
        axs[2].set_xlabel("Sprint")
        axs[2].tick_params(axis='x', rotation=45)

        # 3. SP to Hour Ratio
        team_df["SP/Hour"].plot(ax=axs[3], marker='D', color='purple', title=f"{selected_team} - SP to Hour Ratio")
        axs[3].set_ylabel("SP/Hour")
        axs[3].set_xlabel("Sprint")
        axs[3].tick_params(axis='x', rotation=45)

        # 4. Releases per sprint
        team_df["Releases"].plot(kind='bar', ax=axs[1], color='orange', title=f"{selected_team} - Number of Releases per Sprint")
        axs[1].set_ylabel("Number of Releases")
        axs[1].set_xlabel("Sprint")
        axs[1].tick_params(axis='x', rotation=45)

        plt.tight_layout(rect=[0, 0.03, 1, 0.95])
        st.pyplot(fig)

        # Save to PDF
        pdf.savefig(fig)
        plt.close()
        pdf.close()

        st.download_button(
            label="📥 Download PDF Report",
            data=pdf_buffer.getvalue(),
            file_name=f"{selected_team.replace(' ', '_')}_Dashboard.pdf",
            mime="application/pdf"
        )
