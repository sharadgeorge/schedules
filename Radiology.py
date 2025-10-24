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
    page_icon="🩻",
    layout="centered",
    initial_sidebar_state="expanded"
)

# Sidebar with About section ONLY (no custom navigation)
with st.sidebar:
    #st.markdown("### About")
    st.info("**Project by JA RAD**")

# Main page title
st.title("🩻 Radiology Schedule Management")
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

# Section 2: Create Rad On-Call Schedule
st.header("📅 Create Rad On-Call Schedule")
st.markdown("Upload a blank schedule template to generate an optimized on-call schedule.")

oncall_template = st.file_uploader(
    "Upload Blank Schedule Template (Excel)", 
    type=['xlsx'],
    key='oncall_template',
    help="Upload the blank oncall schedule template Excel file"
)

if oncall_template:
    st.success("✅ Template uploaded successfully")
    
    # Create expandable section for preferences
    with st.expander("⚙️ Scheduling Preferences (Optional)", expanded=False):
        st.markdown("### Prior Month Information")
        st.markdown("Provide information about on-call assignments from the previous month to optimize scheduling.")
        
        # Section 1: GEN rads on last weekends of prior month
        st.markdown("#### 1. GEN Rads on Last Weekend of Prior Month")
        st.caption("These rads will NOT be assigned the first weekend of current month")
        
        col1, col2 = st.columns(2)
        with col1:
            gen_last_weekend_1 = st.selectbox(
                "GEN Rad #1",
                [""] + ['NN', 'MB', 'LK', 'PR', 'AT', 'AK', 'MC', 'AO', 'MM', 'IG', 'MF', 'AS'],
                key='gen_last_weekend_1'
            )
        with col2:
            gen_last_weekend_2 = st.selectbox(
                "GEN Rad #2",
                [""] + ['NN', 'MB', 'LK', 'PR', 'AT', 'AK', 'MC', 'AO', 'MM', 'IG', 'MF', 'AS'],
                key='gen_last_weekend_2'
            )
        
        # Section 2: GEN rad on last day of prior month
        st.markdown("#### 2. GEN Rad on Last Day of Prior Month")
        st.caption("This rad will NOT be assigned on days 1-2 of current month")
        
        gen_last_day = st.selectbox(
            "GEN Rad",
            [""] + ['NN', 'MB', 'LK', 'PR', 'AT', 'AK', 'MC', 'AO', 'MM', 'IG', 'MF', 'AS'],
            key='gen_last_day'
        )
        
        # Section 3: IRA rad on last weekend of prior month
        st.markdown("#### 3. IRA Rad on Last Weekend of Prior Month")
        st.caption("This rad is discouraged but can be assigned during first week")
        
        ira_last_weekend = st.selectbox(
            "IRA Rad",
            [""] + ['IG', 'MF', 'AS'],
            key='ira_last_weekend'
        )
        
        # Section 4: Additional requests
        st.markdown("#### 4. Additional Preference Requests (OFF Days)")
        st.caption("⚠️ Add custom OFF requests (days rads should NOT work)")
        st.caption("📌 Note: Pre-filled X marks in your Excel template are automatically preserved as locked ON assignments")
        
        if 'additional_requests' not in st.session_state:
            st.session_state.additional_requests = []
        
        with st.form("add_request_form"):
            col1, col2, col3, col4 = st.columns([2, 2, 2, 1])
            
            with col1:
                request_section = st.selectbox(
                    "Section",
                    ["GEN", "IRA", "MRI"],
                    key='req_section'
                )
            
            with col2:
                if request_section == "GEN":
                    rad_options = [""] + ['NN', 'MB', 'LK', 'PR', 'AT', 'AK', 'MC', 'AO', 'MM', 'IG', 'MF', 'AS']
                elif request_section == "IRA":
                    rad_options = [""] + ['IG', 'MF', 'AS']
                else:  # MRI
                    rad_options = [""] + ['PR', 'AT', 'AK', 'MC', 'AO', 'MM', 'MF', 'AS']
                
                request_rad = st.selectbox("Rad", rad_options, key='req_rad')
            
            with col3:
                request_day = st.number_input("Day", min_value=1, max_value=31, value=1, key='req_day')
            
            with col4:
                st.write("")  # Spacer
                st.write("")  # Spacer
                request_hard = st.checkbox("Hard", key='req_hard', help="Hard constraint = cannot assign")
            
            add_button = st.form_submit_button("➕ Add Request")
            
            if add_button:
                if not request_rad or request_rad == "":
                    st.warning("⚠️ Please select a Rad before adding request")
                else:
                    constraint_type = "HARD" if request_hard else "SOFT"
                    request_str = f"{request_section}/{request_rad}/Day{request_day}/OFF ({constraint_type})"
                    st.session_state.additional_requests.append({
                        'section': request_section,
                        'rad': request_rad,
                        'day': request_day,
                        'hard': request_hard,
                        'display': request_str
                    })
                    st.rerun()
        
        # Display added requests
        if st.session_state.additional_requests:
            st.markdown("**Added Requests:**")
            for idx, req in enumerate(st.session_state.additional_requests):
                col1, col2 = st.columns([4, 1])
                with col1:
                    st.caption(f"• {req['display']}")
                with col2:
                    if st.button("🗑️", key=f"del_{idx}"):
                        st.session_state.additional_requests.pop(idx)
                        st.rerun()
    
    st.markdown("---")
    
    # Generate button
    if st.button("🚀 Generate Schedule", type="primary", use_container_width=True):
        try:
            with st.spinner("Generating optimized schedule..."):
                # Import the scheduler
                import create_oncall_schedule_v3 as scheduler_module
                from collections import defaultdict
                from datetime import timedelta
                
                # Save uploaded file temporarily
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
                    tmp.write(oncall_template.getvalue())
                    temp_path = tmp.name
                
                # Create scheduler instance
                scheduler = scheduler_module.OnCallScheduler(temp_path)
                
                # Apply user preferences programmatically
                scheduler.special_requests_off = defaultdict(set)
                scheduler.soft_constraints_off = defaultdict(set)
                
                # Section 1: GEN last weekend
                last_weekend_gen = [r for r in [gen_last_weekend_1, gen_last_weekend_2] if r]
                
                # Find first weekend of current month
                first_weekend_days = set()
                for day in range(1, min(15, scheduler.days_in_month + 1)):
                    date = datetime(scheduler.year, scheduler.month, day)
                    if date.weekday() == 4:  # Friday
                        first_weekend_days.add(day)
                        if day + 1 <= scheduler.days_in_month:
                            first_weekend_days.add(day + 1)  # Saturday
                        break
                
                # Calculate prior month info
                first_day_of_month = datetime(scheduler.year, scheduler.month, 1)
                last_day_of_prior_month = first_day_of_month - timedelta(days=1)
                
                for rad in last_weekend_gen:
                    scheduler.special_requests_off[rad].update(first_weekend_days)
                    if last_day_of_prior_month.weekday() in [4, 5]:
                        if 1 <= scheduler.days_in_month:
                            scheduler.special_requests_off[rad].add(1)
                
                # Section 2: GEN last day
                if gen_last_day and gen_last_day not in last_weekend_gen:
                    scheduler.special_requests_off[gen_last_day].update([1, 2])
                
                # Section 3: IRA last weekend
                if ira_last_weekend:
                    first_week_days = set(range(1, min(8, scheduler.days_in_month + 1)))
                    scheduler.soft_constraints_off[ira_last_weekend].update(first_week_days)
                
                # Section 4: Additional requests
                for req in st.session_state.additional_requests:
                    if req['hard']:
                        scheduler.special_requests_off[req['rad']].add(req['day'])
                    else:
                        scheduler.soft_constraints_off[req['rad']].add(req['day'])
                
                # Generate schedule
                scheduler.assign_gen_thursday_saturday()
                scheduler.assign_ira_triplets()
                scheduler.assign_remaining_days('GEN', scheduler_module.GEN_RADS_WITH_IRA)
                scheduler.assign_ira_remaining_weekdays()
                scheduler.assign_mri_optimized()
                
                # Write to Excel
                output_path = scheduler.write_schedule_to_excel()
                
                # Read the generated file
                with open(output_path, 'rb') as f:
                    excel_data = f.read()
                
                # Clean up temp files
                os.unlink(temp_path)
                os.unlink(output_path)
                
                # Calculate statistics
                month_name = calendar.month_name[scheduler.month]
                gen_count = len(scheduler.assignments['GEN'])
                ira_count = len(scheduler.assignments['IRA'])
                mri_count = len(scheduler.assignments['MRI'])
                
                st.success(f"✅ **{month_name} {scheduler.year}**: Schedule generated successfully!")
                
                # Show statistics
                with st.expander("📊 View Details"):
                    st.write(f"**Month:** {month_name} {scheduler.year}")
                    st.write(f"**Days in month:** {scheduler.days_in_month}")
                    st.write("")
                    st.write("**Assignment Coverage:**")
                    st.write(f"- GEN: {gen_count}/{scheduler.days_in_month} days")
                    st.write(f"- IRA: {ira_count}/{scheduler.days_in_month} days")
                    st.write(f"- MRI: {mri_count}/{scheduler.days_in_month} days")
                    
                    # Check for missing days
                    if gen_count < scheduler.days_in_month:
                        missing_gen = [d for d in range(1, scheduler.days_in_month + 1) if d not in scheduler.assignments['GEN']]
                        st.warning(f"⚠️ GEN missing days: {missing_gen}")
                    if ira_count < scheduler.days_in_month:
                        missing_ira = [d for d in range(1, scheduler.days_in_month + 1) if d not in scheduler.assignments['IRA']]
                        st.warning(f"⚠️ IRA missing days: {missing_ira}")
                    if mri_count < scheduler.days_in_month:
                        missing_mri = [d for d in range(1, scheduler.days_in_month + 1) if d not in scheduler.assignments['MRI']]
                        st.warning(f"⚠️ MRI missing days: {missing_mri}")
                    
                    if gen_count == scheduler.days_in_month and ira_count == scheduler.days_in_month and mri_count == scheduler.days_in_month:
                        st.success("✅ All days fully assigned!")

                
                # Download button
                filename = f"OnCall_Schedule_{month_name}_{scheduler.year}_GENERATED.xlsx"
                st.download_button(
                    label="📥 Download Generated Schedule",
                    data=excel_data,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
                
        except Exception as e:
            st.error(f"❌ Error: {str(e)}")
            with st.expander("🔍 View Error Details"):
                import traceback
                st.code(traceback.format_exc())

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
                    file_name="Epic_Import_OnCall_Radiology.csv",
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
