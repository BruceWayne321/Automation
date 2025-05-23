import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(layout="wide", page_title="Team Timesheet Dashboard", page_icon="📊")

st.title("📊 Team Timesheet Dashboard")

# Sidebar for inputs
with st.sidebar:
    st.header("Upload & Filters")
    uploaded_projects_file = st.file_uploader("Upload ZOHO Timesheet", type=["xlsx"], key="projects")
    uploaded_bandwidth_file = st.file_uploader("Upload Microsoft Form Sheet", type=["xlsx"], key="bandwidth")
    project_selector_placeholder = st.empty()

if uploaded_projects_file and uploaded_bandwidth_file:
    # Load 'All Projects' Excel file
    xls_projects = pd.ExcelFile(uploaded_projects_file)
    df_projects_raw = xls_projects.parse('All Projects')
    df_projects_raw.columns = [col.strip() for col in df_projects_raw.columns]

    # Clean and parse date
    df_projects_raw["Date"] = pd.to_datetime(df_projects_raw["Date"], dayfirst=True, errors="coerce")
    df_projects_raw["Date"] = df_projects_raw["Date"].dt.strftime("%d/%m/%Y")

    # Calculate summary from 'All Projects' sheet
    billable = df_projects_raw[df_projects_raw["Billing Type"] == "Billable"]
    nonbillable = df_projects_raw[df_projects_raw["Billing Type"] == "Non Billable"]

    df_billable = billable.groupby("User")["Hours(For Calculation)"].sum().reset_index().rename(columns={"Hours(For Calculation)": "Billable Hours"})
    df_nonbillable = nonbillable.groupby("User")["Hours(For Calculation)"].sum().reset_index().rename(columns={"Hours(For Calculation)": "Non-Billable Hours"})

    df_summary = pd.merge(df_billable, df_nonbillable, on="User", how="outer").fillna(0)
    df_summary["Total Hours Worked"] = df_summary["Billable Hours"] + df_summary["Non-Billable Hours"]
    df_summary = df_summary.rename(columns={"User": "Name"})
    
    
    
   # Load bandwidth & compliance Excel file
    xls_bandwidth = pd.ExcelFile(uploaded_bandwidth_file)

    # Assume planned hours & compliance are in the last sheet of this second file
    last_sheet_name = xls_bandwidth.sheet_names[-1]
    df_bandwidth = xls_bandwidth.parse(last_sheet_name)
    df_bandwidth.columns = [col.strip() for col in df_bandwidth.columns]
 
    
    # Adjust for leaves before calculating utilization %
# Load leaves and calculate Total Hours for Week
    df_leaves = df_bandwidth[["Name", "Number of leaves past week"]].dropna()
    df_leaves["Number of leaves past week"] = pd.to_numeric(df_leaves["Number of leaves past week"], errors="coerce").fillna(0)
    df_leaves = df_leaves.drop_duplicates(subset="Name", keep="last")

# Calculate max weekly hours per person (fallback = 45)
    df_leaves["Base Max Hours"] = 45  # You can make this dynamic if needed

# Adjust based on number of leaves (assume 9 hours per leave)
    df_leaves["Total Hours for Week"] = df_leaves["Base Max Hours"] - (df_leaves["Number of leaves past week"] * 9)

# Ensure no negative values
    df_leaves["Total Hours for Week"] = df_leaves["Total Hours for Week"].clip(lower=0)


    # Standardize names in both datasets for reliable merging
    df_summary["Name_clean"] = df_summary["Name"].str.strip().str.lower()
    df_leaves["Name_clean"] = df_leaves["Name"].str.strip().str.lower()

# Merge on the cleaned name
    df_summary = df_summary.merge(
        df_leaves[["Name_clean", "Number of leaves past week", "Total Hours for Week"]],
        on="Name_clean",
        how="left"
    )

# Drop helper column and fill missing values
    df_summary["Total Hours for Week"] = df_summary["Total Hours for Week"].fillna(45)
    df_summary["Number of leaves past week"] = df_summary["Number of leaves past week"].fillna(0)

# Compute utilization again with updated adjusted hours
    df_summary["Utilization (%)"] = (df_summary["Total Hours Worked"] / df_summary["Total Hours for Week"] * 100).round(2)

# Remove the helper column
    df_summary.drop(columns=["Name_clean"], inplace=True)

    df_summary["Total Hours for Week"] = df_summary["Total Hours for Week"].fillna(45)
    df_summary["Utilization (%)"] = (df_summary["Total Hours Worked"] / df_summary["Total Hours for Week"] * 100).round(2)




    # Add Total row and calculate total utilization properly
    df_actual = df_summary.copy()
    df_actual = df_actual[df_actual["Name"] != "Total"]

    total_row = pd.Series({
        "Name": "Total",
        "Billable Hours": df_actual["Billable Hours"].sum(),
        "Non-Billable Hours": df_actual["Non-Billable Hours"].sum(),
        "Total Hours Worked": df_actual["Total Hours Worked"].sum(),
        "Total Hours for Week": df_actual["Total Hours for Week"].sum(),
        "Utilization (%)": (df_actual["Total Hours Worked"].sum() / df_actual["Total Hours for Week"].sum() * 100)

    })

    df_summary = pd.concat([df_actual, pd.DataFrame([total_row])], ignore_index=True)

    st.subheader("🔹 Team Member Summary")
    df_display = df_summary.copy()
    df_display.index = df_display.index + 1
    st.dataframe(df_display)
    

    # Utilization bar chart
    st.plotly_chart(
        px.bar(
            df_actual,
            x="Name", y="Total Hours Worked", color="Total Hours Worked", text="Total Hours Worked",
            title="Utilization by Team Member"
        ).update_traces(textposition="outside")
    )

    # Billable vs Non-billable bar chart
    bill_chart = df_actual.melt(
        id_vars="Name", value_vars=["Billable Hours", "Non-Billable Hours"]
    )
    st.plotly_chart(
        px.bar(
            bill_chart, x="Name", y="value", color="variable", barmode="group", text="value",
            title="Billable vs Non-Billable Hours"
        ).update_traces(textposition="outside")
    )

    # ---- Team Utilization Overview ----
    st.subheader("🧮 Team Power Utilization Overview")

    total_billable = df_actual["Billable Hours"].sum()
    total_nonbillable = df_actual["Non-Billable Hours"].sum()
    total_utilized = total_billable + total_nonbillable

    total_members = df_actual.shape[0]
    total_capacity = df_actual["Total Hours for Week"].sum()
    total_available = max(0, total_capacity - total_utilized)

    team_utilization_percent = (total_utilized / total_capacity) if total_capacity > 0 else 0

    # Display utilization % and hours
    col1, col2 = st.columns([1, 2])
    with col1:
        st.metric("✅ Total Team Utilization (%)", f"{team_utilization_percent * 100:.2f}%")
    with col2:
        st.markdown(
            f"**Utilized Hours:** {total_utilized:.2f} hrs &nbsp;&nbsp;&nbsp; | &nbsp;&nbsp;&nbsp; "
            f"**Not Utilized Hours:** {total_available:.2f} hrs", unsafe_allow_html=True
        )

    # Pie chart for overall team utilization
    team_util_df = pd.DataFrame({
        "Category": ["Utilized", "Not Utilized"],
        "Hours": [total_utilized, total_available]
    })

    fig_team_util = px.pie(
        team_util_df,
        names="Category",
        values="Hours",
        title="Team Utilization for past week",
        width=650,
        height=500
    )
    fig_team_util.update_traces(textinfo="percent+label", textposition="inside")
    st.plotly_chart(fig_team_util)

    # 📂 Project-wise Task Breakdown
    st.subheader("📂 Project-wise Task Breakdown")

    projects = df_projects_raw["Project Name"].dropna().unique()
    selected_project = st.selectbox("Select Project", projects)

    if selected_project:
        df_filtered = df_projects_raw[df_projects_raw["Project Name"] == selected_project]
        tasks = df_filtered["Task/General/Issue"].dropna().unique()

        for task in tasks:
            df_task = df_filtered[df_filtered["Task/General/Issue"] == task]
            st.markdown(f"### 🔸 Task: {task}")

            df_task_display = df_task[["User", "Hours(For Calculation)", "Billing Type", "Date"]].rename(
                columns={
                    "User": "Team Member",
                    "Hours(For Calculation)": "Hours",
                    "Billing Type": "Type"
                }
            ).reset_index(drop=True)  # Reset index to 0,1,2...
            df_task_display.index = df_task_display.index + 1  # Start from 1
            st.dataframe(df_task_display)


        # 🔢 Summary: Total Hours for Selected Project
        total_hours = df_filtered["Hours(For Calculation)"].sum()
        st.markdown(f"### 🧾 Total Hours Logged for 📁 **{selected_project}**")
        st.metric(label="⏱️ Total Hours", value=f"{round(total_hours, 2)} 🕒")


        # Project hours bar chart with data labels
        st.plotly_chart(
            px.bar(
                df_filtered.groupby("User")["Hours(For Calculation)"].sum().reset_index(),
                x="User", y="Hours(For Calculation)", text="Hours(For Calculation)",
                title="Total Hours per Team Member"
            ).update_traces(textposition="outside")
        )

 
    # Process Planned vs Available Bandwidth
    st.subheader("📊 Planned vs Available Bandwidth (Next Week)")

    df_planned = df_bandwidth[["Name", "Planned hours for the coming week"]].dropna()
    df_planned["Planned hours for the coming week"] = pd.to_numeric(
        df_planned["Planned hours for the coming week"], errors="coerce"
    )
    df_planned = df_planned.dropna(subset=["Planned hours for the coming week"])
    df_planned = df_planned.drop_duplicates(subset="Name", keep="last")
    df_planned["Available hours"] = 45 - df_planned["Planned hours for the coming week"]

    df_melt = df_planned.melt(
        id_vars="Name",
        value_vars=["Planned hours for the coming week", "Available hours"],
        var_name="Hour Type",
        value_name="Hours"
    )

    fig = px.bar(
        df_melt,
        x="Name",
        y="Hours",
        color="Hour Type",
        barmode="group",
        text="Hours",
        title="🧑‍💼 Planned vs Available Hours per Team Member"
    )
    fig.update_layout(xaxis_title="", yaxis_title="Hours")
    fig.update_traces(textposition="outside")

    st.plotly_chart(fig, use_container_width=True)

    # ------------------ Compliance Overview Pie Charts ------------------
    st.subheader("✅ Compliance Overview")

    df_last_unique = df_bandwidth.dropna(subset=["Name"]).drop_duplicates(subset=["Name"], keep="last")

    compliance_columns = {
        "Is your timesheet submitted?": "Timesheet Submission",
        "All tasks access requested for and created?": "Task Access Requested",
        "All checkin and checkout times accurate for the week? Regularized where inaccurate?": "Check-in/Checkout Accuracy"
    }

    for col, title in compliance_columns.items():
        if col in df_last_unique.columns:
            yes_names = df_last_unique[df_last_unique[col].str.strip().str.lower() == "yes"]["Name"].tolist()
            no_names = df_last_unique[df_last_unique[col].str.strip().str.lower() != "yes"]["Name"].tolist()

            data = []
            if len(yes_names) > 0:
                data.append({"Response": "Yes", "Count": len(yes_names), "Names": ", ".join(yes_names)})

            if len(no_names) > 0:
                data.append({"Response": "No", "Count": len(no_names), "Names": ", ".join(no_names)})

            if not data:
                data.append({"Response": "None", "Count": 1, "Names": "None"})

            df_plot = pd.DataFrame(data)

            fig = px.pie(
                df_plot,
                names="Response",
                values="Count",
                title=f"{title} - Compliance Overview",
                width=550,
                height=500
            )

            fig.update_traces(
                textinfo="percent+label",
                textposition="inside",
                customdata=df_plot["Names"],
                hovertemplate="<b>%{label}</b><br>Count: %{value}<br>Names: %{customdata}<extra></extra>"
            )

            st.plotly_chart(fig)
else:
    st.info("Please upload both Excel files: 'ZOHO Timesheet' and 'Microsoft Forms Sheet' to see the dashboard.")
