import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime, timedelta
from collections import defaultdict
import requests
from io import BytesIO

st.set_page_config(page_title="Resource Allocation Assistant", layout="wide")
st.markdown("""
    <style>
        .block-container {
            padding-top: 4rem !important;
        }
        section[data-testid="stSidebar"] {
            padding-top: 0rem !important;
        }
        [data-testid="stSidebar"] .css-ng1t4o {
            padding-top: 0rem !important;
        }
        h1 {
            margin-top: 3rem !important;
        }
    </style>
""", unsafe_allow_html=True)

st.title("\n\n🔍 Resource Allocation Assistant")

upload_method = st.radio("Select Input Method", ["📂 Upload File", "🔗 SharePoint Link"], horizontal=True)

uploaded_file = None
sharepoint_url = None
if upload_method == "📂 Upload File":
    uploaded_file = st.file_uploader("Upload the DC-v2 Excel File", type=["xlsx"])
elif upload_method == "🔗 SharePoint Link":
    sharepoint_url = st.text_input("Paste SharePoint file link")

def get_excel_from_sharepoint(url):
    try:
        response = requests.get(url)
        if response.status_code == 200:
            return BytesIO(response.content)
        else:
            st.error("Failed to fetch file from SharePoint. Please ensure the link is accessible.")
            return None
    except Exception as e:
        st.error(f"Error accessing SharePoint file: {e}")
        return None

@st.cache_data
def load_data(file):
    people_df = pd.read_excel(file, sheet_name="All Active Team Members - Consu")
    agg_df = pd.read_excel(file, sheet_name="People Aggregated", header=4)

    people_df = people_df[["Resource", "Primary Role"]].rename(columns={"Resource": "Name"})
    excluded_names = ["Consultant", "Senior Consultant", "Associate", "Senior Associate", "Engagement Manager"]
    agg_df = agg_df[~agg_df[agg_df.columns[1]].isin(excluded_names)]

    agg_df.rename(columns={
        agg_df.columns[0]: "Cluster",
        agg_df.columns[1]: "Name"
    }, inplace=True)

    if "Role" in agg_df.columns:
        agg_df.drop(columns=["Role"], inplace=True)

    datetime_cols = [col for col in agg_df.columns if isinstance(col, datetime)]

    merged_df = pd.merge(agg_df, people_df, on="Name", how="left")
    merged_df.rename(columns={"Primary Role": "Role"}, inplace=True)

    return merged_df, datetime_cols

# Determine input file source
input_file = None
if uploaded_file:
    input_file = uploaded_file
elif sharepoint_url:
    input_file = get_excel_from_sharepoint(sharepoint_url)

if input_file:
    df, datetime_cols = load_data(input_file)
    df = df[df["Cluster"] != "Others"]
    clusters = sorted(df["Cluster"].dropna().unique().tolist())

    with st.sidebar:
        st.header("📋 Project Details")
        project_name = st.text_input("Project Name")
        start_date = st.date_input("Start Date")
        end_date = st.date_input("End Date")

        start_date = datetime.combine(start_date, datetime.min.time())
        start_date = start_date - timedelta(days=start_date.weekday())
        end_date = datetime.combine(end_date, datetime.min.time())
        end_date = end_date + timedelta(days=(6 - end_date.weekday()))

        filtered_weeks = [col for col in datetime_cols if start_date <= col <= end_date]
        st.markdown(f"🗓️ **Number of weeks considered:** {len(filtered_weeks)}")

        min_fit_percent = st.slider("Minimum Fit %", min_value=0, max_value=100, value=80)
        selected_cluster = st.selectbox("Select Cluster", options=clusters)

        available_roles = sorted(df["Role"].dropna().astype(str).unique(), key=lambda x: (
            0 if "Execution Owner" in x else 1 if "Senior" in x else 2
        ))

        st.subheader("💡 Effort Required by Role")
        role_efforts = {}
        for role in available_roles:
            effort = st.number_input(f"{role}", min_value=0, step=1, value=0)
            if effort > 0:
                role_efforts[role] = effort

    df["Assigned Hours"] = df[filtered_weeks].sum(axis=1)
    df["Weeks"] = len(filtered_weeks)
    df["Capacity"] = df["Weeks"] * 40
    df["Free Hours"] = df["Capacity"] - df["Assigned Hours"]
    df["Utilization %"] = df["Assigned Hours"] / df["Capacity"] * 100

    if "suggest" not in st.session_state:
        st.session_state.suggest = False
    if st.sidebar.button("Suggest Resources"):
        st.session_state.suggest = True

    if st.session_state.suggest and role_efforts:
        st.session_state.final_output = []

        st.markdown("### 🔍 Compare with Specific People")
        compare_role = st.selectbox("Select Role", options=available_roles)
        filtered_names = df[df['Role'] == compare_role]['Name'].dropna().unique()
        selected_name = st.selectbox("Select Name", options=sorted(filtered_names))
        if selected_name:
            compare_df = df[df["Name"] == selected_name].copy()
            for role, effort in role_efforts.items():
                compare_df["Anticipated Utilization %"] = (compare_df["Assigned Hours"] + effort) / compare_df["Capacity"] * 100
            st.dataframe(compare_df[["Name", "Cluster", "Free Hours", "Utilization %", "Anticipated Utilization %"]])

        for role, effort in role_efforts.items():
            st.markdown(f"### 🔹 {role}")

            cluster_df = df[(df["Cluster"] == selected_cluster) & (df["Role"] == role)].copy()
            cluster_df["Fit %"] = (cluster_df["Free Hours"] / effort) * 100
            cluster_df = cluster_df[cluster_df["Fit %"] >= min_fit_percent]
            cluster_df["Anticipated Utilization %"] = (cluster_df["Assigned Hours"] + effort) / cluster_df["Capacity"] * 100
            cluster_top = cluster_df.sort_values(by="Fit %", ascending=False).head(3)

            other_df = df[(df["Role"] == role)].copy()
            other_df["Fit %"] = (other_df["Free Hours"] / effort) * 100
            other_df = other_df[other_df["Fit %"] >= min_fit_percent]
            other_df = other_df[~other_df["Name"].isin(cluster_top["Name"])]
            other_df["Anticipated Utilization %"] = (other_df["Assigned Hours"] + effort) / other_df["Capacity"] * 100
            overall_top = other_df.sort_values(by="Fit %", ascending=False).head(3)

            if not cluster_top.empty:
                st.markdown(f"**Top 3 in {selected_cluster}**")
                st.dataframe(cluster_top[["Name", "Cluster", "Free Hours", "Fit %", "Utilization %", "Anticipated Utilization %"]])
                for r in cluster_top.itertuples(index=False):
                    st.session_state.final_output.append(r._asdict())
            else:
                st.warning(f"No suitable candidates found for role: {role} in {selected_cluster}.")

            if not overall_top.empty:
                st.markdown("**Top 3 Overall (All Clusters)**")
                st.dataframe(overall_top[["Name", "Cluster", "Free Hours", "Fit %", "Utilization %", "Anticipated Utilization %"]])
                for r in overall_top.itertuples(index=False):
                    st.session_state.final_output.append(r._asdict())

        if st.session_state.final_output:
            df_final = pd.DataFrame(st.session_state.final_output)
            csv = df_final.to_csv(index=False).encode("utf-8")
            st.download_button("📥 Download Selected Suggestions", csv, f"{project_name}_{selected_cluster}_suggestions.csv")
