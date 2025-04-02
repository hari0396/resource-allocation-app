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

input_file = uploaded_file if uploaded_file else get_excel_from_sharepoint(sharepoint_url)

if input_file:
    df, datetime_cols = load_data(input_file)
    df = df[df["Cluster"] != "Others"]
    clusters = sorted(df["Cluster"].dropna().unique().tolist())

    with st.sidebar:
        st.header("📋 Project Details")
        project_count = st.number_input("Number of Projects", min_value=1, max_value=3, value=1)
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

        compare_enabled = st.checkbox("Enable Comparison")
        if compare_enabled:
            st.markdown("### 🔍 Compare with Specific People")
            compare_role = st.selectbox("Select Role", options=available_roles)
            filtered_names = df[df['Role'] == compare_role]['Name'].dropna().unique()
            selected_name = st.selectbox("Select Name", options=sorted(filtered_names))
            if selected_name:
                compare_df = df[df["Name"] == selected_name].copy()
                compare_df["Assigned Hours"] = compare_df[filtered_weeks].sum(axis=1)
                compare_df["Weeks"] = len(filtered_weeks)
                compare_df["Capacity"] = compare_df["Weeks"] * 40
                compare_df["Free Hours"] = compare_df["Capacity"] - compare_df["Assigned Hours"]
                compare_df["Utilization %"] = compare_df["Assigned Hours"] / compare_df["Capacity"] * 100
                for role in available_roles:
                    effort = st.number_input(f"{role} (For Comparison Only)", min_value=0, step=1, value=0)
                    if effort > 0:
                        compare_df["Anticipated Utilization %"] = (compare_df["Assigned Hours"] + effort) / compare_df["Capacity"] * 100
                st.dataframe(compare_df[["Name", "Cluster", "Free Hours", "Utilization %", "Anticipated Utilization %"]])

    if "project_allocations" not in st.session_state:
        st.session_state.project_allocations = []

    current_len = len(st.session_state.project_allocations)
    if current_len < project_count:
        st.session_state.project_allocations.extend([{}] * (project_count - current_len))
    elif current_len > project_count:
        st.session_state.project_allocations = st.session_state.project_allocations[:project_count]

    for i in range(project_count):
        with st.container():
            st.header(f"🚀 Project {i+1} Resource Suggestions")
            st.subheader("💡 Effort Required by Role")

            role_efforts = {}
            for role in available_roles:
                effort = st.sidebar.number_input(f"Project {i+1} - {role}", min_value=0, step=1, key=f"effort_{i}_{role}")
                if effort > 0:
                    role_efforts[role] = effort

            df_copy = df.copy()

            for past_project in st.session_state.project_allocations[:i]:
                for selected_name, added_effort in past_project.items():
                    df_copy.loc[df_copy["Name"] == selected_name, filtered_weeks] += added_effort / len(filtered_weeks)

            df_copy["Assigned Hours"] = df_copy[filtered_weeks].sum(axis=1)
            df_copy["Weeks"] = len(filtered_weeks)
            df_copy["Capacity"] = df_copy["Weeks"] * 40
            df_copy["Free Hours"] = df_copy["Capacity"] - df_copy["Assigned Hours"]
            df_copy["Utilization %"] = df_copy["Assigned Hours"] / df_copy["Capacity"] * 100

            project_allocated = {}

            for role, effort in role_efforts.items():
                st.markdown(f"### 🔹 {role}")
                cluster_df = df_copy[(df_copy["Cluster"] == selected_cluster) & (df_copy["Role"] == role)].copy()
                cluster_df["Fit %"] = (cluster_df["Free Hours"] / effort) * 100
                cluster_df = cluster_df[cluster_df["Fit %"] >= min_fit_percent]
                cluster_df["Anticipated Utilization %"] = (cluster_df["Assigned Hours"] + effort) / cluster_df["Capacity"] * 100
                cluster_top = cluster_df.sort_values(by="Fit %", ascending=False).head(3)

                if not cluster_top.empty:
                    selected_candidates = st.multiselect(
                        f"Select resources for {role} (Project {i+1})",
                        options=cluster_top["Name"].tolist(),
                        default=[],
                        key=f"select_{i}_{role}"
                    )
                    for name in selected_candidates:
                        project_allocated[name] = project_allocated.get(name, 0) + effort
                    st.dataframe(cluster_top[["Name", "Cluster", "Free Hours", "Fit %", "Utilization %", "Anticipated Utilization %"]])
                else:
                    st.warning(f"No suitable candidates found for role: {role} in {selected_cluster}.")

            st.session_state.project_allocations[i] = project_allocated