import streamlit as st
import sys
import os
from pathlib import Path
import tempfile
import io
import csv
import calendar
from datetime import datetime
import openpyxl

# Import the converter module
import oncall_converter_Radiology_demo_v2 as rad_converter

# Configure page
st.set_page_config(
    page_title="Radiology - Schedule Management",
    page_icon="🔬",
    layout="centered",
    initial_sidebar_state="expanded"
)

# Sidebar with About section ONLY (no custom navigation)
with st.sidebar:
    st.markdown("### About")
    st.info("**Project by JA RAD**")

# Main page title
st.title("🔬 Radiology Schedule Management")
st.markdown("---")

# Section 1: Create Rad Work Schedule (Placeholder)
st.header("📝 Create Rad Work Schedule")
st.markdown("Upload a blank or partially filled Work Schedule template to generate a completed schedule.")
st.warning("⚠️ **In Development** - This feature is currently under development.")

work_template = st.file_uploader(
    "Upload Work Schedule Template (Excel)", 
    type=['xlsx'],
    key='work_template',
    help="Upload the Work Schedule template Excel file"
)

if work_template:
    st.info("📋 Work Schedule creation feature coming soon...")

st.markdown("---")

# Section 2: Create Rad On-Call Schedule (Placeholder)
st.header("📅 Create Rad On-Call Schedule")
st.markdown("Upload a blank or partially filled On-Call Schedule template to generate a completed schedule.")
st.warning("⚠️ **In Development** - This feature is currently under development.")

oncall_template = st.file_uploader(
    "Upload On-Call Schedule Template (Excel)", 
    type=['xlsx'],
    key='oncall_template',
    help="Upload the On-Call Schedule template Excel file"
)

if oncall_template:
    st.info("📋 On-Call Schedule creation feature coming soon...")

st.markdown("---")

# Section 3: Convert Rad Schedules for Import (Functional)
st.header("🔄 Convert Rad Schedules for Import")
st.markdown("Upload completed Work Schedule and On-Call Schedule files to generate the import CSV file.")

# File uploaders
st.subheader("Upload Schedule Files (in order):")

col1, col2 = st.columns(2)

with col1:
    work_file = st.file_uploader(
        "1. Work Schedule (Excel)", 
        type=['xlsx'],
        key='work_schedule',
        help="Upload the completed Work Schedule Excel file"
    )

with col2:
    oncall_file = st.file_uploader(
        "2. On-Call Schedule (Excel)", 
        type=['xlsx'],
        key='oncall_schedule',
        help="Upload the completed On-Call Schedule Excel file"
    )

# Process button and conversion logic
if work_file and oncall_file:
    st.markdown("---")
    
    # Add month/year selection
    st.subheader("📅 Select Processing Month")
    
    col1, col2 = st.columns(2)
    with col1:
        selected_month = st.selectbox(
            "Month",
            options=list(range(1, 13)),
            format_func=lambda x: calendar.month_name[x],
            index=10  # Default to November (index 10 for month 11)
        )
    
    with col2:
        selected_year = st.number_input(
            "Year",
            min_value=2020,
            max_value=2030,
            value=2025,
            step=1
        )
    
    st.info(f"📅 Will process: **{calendar.month_name[selected_month]} {selected_year}**")
    
    if st.button("🔄 Convert to Import Format", type="primary"):
        try:
            with st.spinner("Processing schedules..."):
                # Save uploaded files temporarily
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_work:
                    tmp_work.write(work_file.getvalue())
                    work_path = tmp_work.name
                
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_oncall:
                    tmp_oncall.write(oncall_file.getvalue())
                    oncall_path = tmp_oncall.name
                
                # Load workbooks
                wb_work = openpyxl.load_workbook(work_path)
                ws_work = wb_work['WORK SCHEDULE']
                
                wb_oncall = openpyxl.load_workbook(oncall_path, data_only=True)
                ws_oncall = wb_oncall['Sheet1']
                
                # Process schedules with EXPLICIT month/year
                output_data = rad_converter.process_schedules(
                    ws_work, 
                    ws_oncall, 
                    selected_year,
                    selected_month
                )
                
                # Show success message
                st.success(f"✅ Generated {len(output_data)} schedule entries")
                
                # Show detailed statistics
                with st.expander("📊 View Detailed Statistics"):
                    days_in_month = calendar.monthrange(selected_year, selected_month)[1]
                    
                    # Calculate expected entries
                    weekday_count = 0
                    weekend_count = 0
                    for day in range(1, days_in_month + 1):
                        date = datetime(selected_year, selected_month, day)
                        if date.weekday() in [6, 0, 1, 2, 3]:  # Sun-Thu
                            weekday_count += 1
                        else:  # Fri-Sat
                            weekend_count += 1
                    
                    expected_entries = weekday_count * 11 + weekend_count * 5
                    
                    st.write(f"**Month:** {calendar.month_name[selected_month]} {selected_year}")
                    st.write(f"**Days in month:** {days_in_month}")
                    st.write(f"**Weekdays (Sun-Thu):** {weekday_count} days")
                    st.write(f"**Weekends (Fri-Sat):** {weekend_count} days")
                    st.write(f"**Expected entries:** {expected_entries}")
                    st.write(f"**Actual entries:** {len(output_data)}")
                    
                    if len(output_data) == expected_entries:
                        st.success("✅ Entry count is correct!")
                    else:
                        st.warning(f"⚠️ {abs(expected_entries - len(output_data))} entries difference")
                    
                    # Team breakdown
                    team_counts = {}
                    for entry in output_data:
                        team = entry['TEAM']
                        team_counts[team] = team_counts.get(team, 0) + 1
                    
                    st.write("\n**Entries per team:**")
                    team_names = {
                        '114': 'Gen_CT',
                        '115': 'IRA',
                        '116': 'MRI',
                        '126': 'US',
                        '127': 'Fluoro'
                    }
                    for team_id in sorted(team_counts.keys()):
                        st.write(f"- {team_names.get(team_id, team_id)}: {team_counts[team_id]} entries")
                
                # Create CSV output
                output = io.StringIO()
                fieldnames = ['EMPLOYEE', 'TEAM', 'STARTDATE', 'STARTTIME', 
                             'ENDDATE', 'ENDTIME', 'ROLE', 'NOTES', 'ORDER', 'TEAMCOMMENT']
                writer = csv.DictWriter(output, fieldnames=fieldnames, delimiter='^')
                writer.writeheader()
                writer.writerows(output_data)
                
                csv_data = output.getvalue()
                
                # Provide download button
                st.download_button(
                    label="📥 Download Import_OnCall_Radiology.csv",
                    data=csv_data,
                    file_name="Import_OnCall_Radiology.csv",
                    mime="text/csv",
                    type="primary"
                )
                
                st.success("✅ Conversion complete! Click the button above to download your file.")
                
                # Clean up temp files
                os.unlink(work_path)
                os.unlink(oncall_path)
                
        except Exception as e:
            st.error(f"❌ Error during conversion: {str(e)}")
            st.error("Please check that your Excel files have the correct structure and try again.")
            import traceback
            with st.expander("🔍 View Error Details"):
                st.code(traceback.format_exc())

else:
    st.info("👆 Please upload both Excel files to begin conversion")

# Footer
st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: gray; font-size: 0.8em;'>"
    "Schedule Management System | Powered by Streamlit"
    "</div>",
    unsafe_allow_html=True
)
