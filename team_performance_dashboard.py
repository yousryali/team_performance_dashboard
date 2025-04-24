import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.backends.backend_pdf import PdfPages
import io

# —————————————————————————
# App configuration
# —————————————————————————
st.set_page_config(layout="wide")
st.title("📊 Team Performance Dashboard")
st.write("Upload the performance Excel sheet to generate visual analysis per team.")

# —————————————————————————
# Upload
# —————————————————————————
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])
if not uploaded_file:
    st.stop()

# —————————————————————————
# Load & reshape
# —————————————————————————
df = pd.read_excel(uploaded_file, engine="openpyxl")
df = df.set_index("Row Labels").T
df.index.name = "Sprint"

# —————————————————————————
# Teams
# —————————————————————————
teams = [
    "Agency (Team Ephesus)",
    "Agency (Team Cyclopes)",
    "Production Systems",
    "TPS & DP"
]

# —————————————————————————
# Metric keyword → display name
# —————————————————————————
metric_map = [
    ("Velocity", ["Velocity"]),
    ("Billable TS", ["Billable"]),
    ("Non-Billable TS", ["Investment"]),
    ("Bugs Created", ["# Bugs created"]),
    ("Bugs Closed", ["# Bugs closed"]),
    ("Releases", ["# Releases in Prod"]),
    ("SP/Hour", ["Efficiency - Ratio"]),
]

# —————————————————————————
# Team picker
# —————————————————————————
selected_team = st.selectbox("Select Team to Visualize", teams)

# —————————————————————————
# Gather that team’s columns
# —————————————————————————
team_cols = [c for c in df.columns if c.startswith(selected_team)]
if not team_cols:
    st.error(f"No columns found for `{selected_team}`. Check your Excel headers.")
    st.stop()

# —————————————————————————
# Map to our standard names
# —————————————————————————
cols, names = [], []
for disp, keys in metric_map:
    # find a column that contains any of our keywords
    match = next((c for c in team_cols if any(k.lower() in c.lower() for k in keys)), None)
    if match:
        cols.append(match)
        names.append(disp)

# final DataFrame
team_df = df[cols].copy()
team_df.columns = names
team_df = team_df.apply(pd.to_numeric, errors="coerce")

# —————————————————————————
# Derived metrics
# —————————————————————————
if {"Billable TS", "Non-Billable TS"}.issubset(team_df.columns):
    team_df["Total Time"] = team_df["Billable TS"] + team_df["Non-Billable TS"]
if "Velocity" in team_df.columns and "Billable TS" in team_df.columns:
    team_df["Efficiency"] = team_df["Velocity"] / (team_df["Billable TS"] / 100)

# —————————————————————————
# Show table
# —————————————————————————
st.subheader(f"{selected_team} — Metrics Overview")
st.dataframe(team_df.style.format("{:.2f}"))

# —————————————————————————
# Plotting
# —————————————————————————
sns.set(style="whitegrid")
fig, axs = plt.subplots(4, 1, figsize=(15, 20))
fig.suptitle(f"{selected_team} — Performance Dashboard", fontsize=18, weight="bold")

# 1) Velocity
if "Velocity" in team_df:
    team_df["Velocity"].plot(ax=axs[0], marker="o", title="Velocity Trend")
    axs[0].set_ylabel("Story Points")
else:
    axs[0].text(0.5, 0.5, "Velocity missing", ha="center", va="center")
    axs[0].axis("off")

# 2) Releases
if "Releases" in team_df:
    team_df["Releases"].plot(kind="bar", ax=axs[1], title="Releases per Sprint")
    axs[1].set_ylabel("Releases")
    axs[1].tick_params(axis="x", rotation=45)
else:
    axs[1].text(0.5, 0.5, "Releases missing", ha="center", va="center")
    axs[1].axis("off")

# 3) Bugs Created vs Closed
if {"Bugs Created", "Bugs Closed"}.issubset(team_df.columns):
    team_df[["Bugs Created", "Bugs Closed"]].plot(
        kind="bar", ax=axs[2], title="Bugs Created vs Closed"
    )
    axs[2].set_ylabel("Count")
    axs[2].tick_params(axis="x", rotation=45)
else:
    axs[2].text(0.5, 0.5, "Bug data missing", ha="center", va="center")
    axs[2].axis("off")

# 4) SP/Hour
if "SP/Hour" in team_df:
    team_df["SP/Hour"].plot(ax=axs[3], marker="D", title="SP to Hour Ratio")
    axs[3].set_ylabel("SP/Hour")
    axs[3].tick_params(axis="x", rotation=45)
else:
    axs[3].text(0.5, 0.5, "SP/Hour missing", ha="center", va="center")
    axs[3].axis("off")

plt.tight_layout(rect=[0, 0.03, 1, 0.95])
st.pyplot(fig)

# —————————————————————————
# PDF Export
# —————————————————————————
buffer = io.BytesIO()
pp = PdfPages(buffer)
pp.savefig(fig)
pp.close()

st.download_button(
    "📥 Download PDF Report",
    data=buffer.getvalue(),
    file_name=f"{selected_team.replace(' ', '_')}_Dashboard.pdf",
    mime="application/pdf"
)
