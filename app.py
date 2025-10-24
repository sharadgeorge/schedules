import streamlit as st
import sys
import os
from pathlib import Path
import tempfile
import shutil

# Configure page
st.set_page_config(
    page_title="Schedule Management",
    page_icon="🏥",
    layout="centered",
    initial_sidebar_state="expanded"
)

# Sidebar with About section
with st.sidebar:
    st.markdown("### About")
    st.info("**Project by JA RAD**")
    st.markdown("---")
    st.markdown("### Navigation")
    st.page_link("app.py", label="🔬 Radiology", icon="🏠")
    st.page_link("pages/cardiology.py", label="❤️ Cardiology")

# Main page title
st.title("🔬 Radiology Schedule Management")
st.markdown("---")

# Section 1: Create Rad Work Schedule
st.header("📅 Create Rad Work Schedule")
st.markdown("Upload a blank or partially filled Work Schedule template to generate a completed schedule.")

work_schedule_file = st.file_uploader(
    "Upload Work Schedule Template (Excel)", 
    type=['xlsx'],
    key='work_schedule_upload',
    help="Upload the Excel file containing the Work Schedule template"
)

if work_schedule_file:
    if st.button("Generate Work Schedule", key='gen_work'):
        with st.spinner("Processing Work Schedule..."):
            try:
                # Create temporary directory for processing
                with tempfile.TemporaryDirectory() as temp_dir:
                    # Save uploaded file
                    input_path = Path(temp_dir) / work_schedule_file.name
                    with open(input_path, 'wb') as f:
                        f.write(work_schedule_file.getbuffer())
                    
                    # Import and run the script
                    sys.path.insert(0, str(Path(__file__).parent))
                    from create_Rad_Work_Schedule import main as work_schedule_main
                    
                    # Note: This is a placeholder - the actual script needs to be implemented
                    st.info("⚠️ Work Schedule generation script is under development.")
                    st.info("This section will be functional once the script is completed.")
                    
            except Exception as e:
                st.error(f"An error occurred: {str(e)}")
                st.info("This feature is currently under development.")

st.markdown("---")

# Section 2: Create Rad On-Call Schedule
st.header("🌙 Create Rad On-Call Schedule")
st.markdown("Upload a blank or partially filled On-Call Schedule template to generate a completed schedule.")

oncall_schedule_file = st.file_uploader(
    "Upload On-Call Schedule Template (Excel)", 
    type=['xlsx'],
    key='oncall_schedule_upload',
    help="Upload the Excel file containing the On-Call Schedule template"
)

if oncall_schedule_file:
    if st.button("Generate On-Call Schedule", key='gen_oncall'):
        with st.spinner("Processing On-Call Schedule..."):
            try:
                # Create temporary directory for processing
                with tempfile.TemporaryDirectory() as temp_dir:
                    # Save uploaded file
                    input_path = Path(temp_dir) / oncall_schedule_file.name
                    with open(input_path, 'wb') as f:
                        f.write(oncall_schedule_file.getbuffer())
                    
                    # Import and run the script
                    sys.path.insert(0, str(Path(__file__).parent))
                    from Create_Rad_OnCall_Schedule import main as oncall_schedule_main
                    
                    # Note: This is a placeholder - the actual script needs to be implemented
                    st.info("⚠️ On-Call Schedule generation script is under development.")
                    st.info("This section will be functional once the script is completed.")
                    
            except Exception as e:
                st.error(f"An error occurred: {str(e)}")
                st.info("This feature is currently under development.")

st.markdown("---")

# Section 3: Convert Rad Schedules for Import
st.header("🔄 Convert Rad Schedules for Import")
st.markdown("Upload **two** Excel files (Work Schedule and On-Call Schedule) to generate the import CSV file.")

col1, col2 = st.columns(2)

with col1:
    rad_work_file = st.file_uploader(
        "1. Work Schedule (Excel)", 
        type=['xlsx'],
        key='rad_work_convert',
        help="Upload the completed Work Schedule Excel file"
    )

with col2:
    rad_oncall_file = st.file_uploader(
        "2. On-Call Schedule (Excel)", 
        type=['xlsx'],
        key='rad_oncall_convert',
        help="Upload the completed On-Call Schedule Excel file"
    )

if rad_work_file and rad_oncall_file:
    if st.button("Convert to Import Format", key='convert_rad'):
        with st.spinner("Converting schedules to import format..."):
            try:
                # Create temporary directory for processing
                with tempfile.TemporaryDirectory() as temp_dir:
                    # Save uploaded files
                    work_path = Path(temp_dir) / "work_schedule.xlsx"
                    oncall_path = Path(temp_dir) / "oncall_schedule.xlsx"
                    
                    with open(work_path, 'wb') as f:
                        f.write(rad_work_file.getbuffer())
                    
                    with open(oncall_path, 'wb') as f:
                        f.write(rad_oncall_file.getbuffer())
                    
                    # Import necessary modules
                    sys.path.insert(0, str(Path(__file__).parent))
                    import oncall_converter_Radiology_demo_v2 as rad_converter
                    
                    # Load workbooks
                    import openpyxl
                    wb_work = openpyxl.load_workbook(work_path)
                    ws_work = wb_work['WORK SCHEDULE']
                    
                    wb_oncall = openpyxl.load_workbook(oncall_path, data_only=True)
                    ws_oncall = wb_oncall['Sheet1']
                    
                    # Extract month and year
                    import calendar
                    from datetime import datetime
                    filename = oncall_path.stem
                    current_month, current_year = rad_converter.extract_month_year_from_filename(filename)
                    
                    if current_month is None:
                        current_month = datetime.now().month
                    if current_year is None:
                        current_year = datetime.now().year
                    
                    # Process schedules
                    output_data = rad_converter.process_schedules(ws_work, ws_oncall, current_year, current_month)
                    
                    # Write CSV output
                    import csv
                    csv_output = Path(temp_dir) / "Import_OnCall_Radiology.csv"
                    with open(csv_output, 'w', newline='') as csvfile:
                        fieldnames = ['EMPLOYEE', 'TEAM', 'STARTDATE', 'STARTTIME', 
                                     'ENDDATE', 'ENDTIME', 'ROLE', 'NOTES', 'ORDER', 'TEAMCOMMENT']
                        writer = csv.DictWriter(csvfile, fieldnames=fieldnames, delimiter='^')
                        writer.writeheader()
                        writer.writerows(output_data)
                    
                    # Read the CSV file for download
                    with open(csv_output, 'rb') as f:
                        csv_data = f.read()
                    
                    # Success message
                    st.success(f"✅ Successfully generated {len(output_data)} schedule entries for {calendar.month_name[current_month]} {current_year}")
                    
                    # Download button
                    st.download_button(
                        label="📥 Download Import_OnCall_Radiology.csv",
                        data=csv_data,
                        file_name="Import_OnCall_Radiology.csv",
                        mime="text/csv",
                        key='download_rad_csv'
                    )
                    
            except Exception as e:
                st.error(f"❌ An error occurred: {str(e)}")
                st.error("Please ensure you've uploaded valid Excel files with the correct format.")

# Footer
st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: gray; font-size: 0.8em;'>"
    "Schedule Management System | Powered by Streamlit"
    "</div>",
    unsafe_allow_html=True
)
