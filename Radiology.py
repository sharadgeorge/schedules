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
    st.markdown("### About")
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
            # Create debug container that starts collapsed
            debug_expander = st.expander("🔍 Debug Information (Click to expand if needed)", expanded=False)
            
            with st.spinner("Creating scheduler..."):
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
                for day in range(1, scheduler.days_in_month + 1):
                    if scheduler.get_day_type(day) in ['thu', 'fri', 'sat']:
                        first_weekend_days.add(day)
                        if len(first_weekend_days) >= 3:
                            break
                
                for rad in last_weekend_gen:
                    for day in first_weekend_days:
                        scheduler.special_requests_off['GEN'].add((rad, day))
                
                # Section 2: GEN last day
                if gen_last_day:
                    scheduler.special_requests_off['GEN'].add((gen_last_day, 1))
                    scheduler.special_requests_off['GEN'].add((gen_last_day, 2))
                
                # Section 3: IRA last weekend
                if ira_last_weekend:
                    for day in range(1, min(8, scheduler.days_in_month + 1)):
                        scheduler.soft_constraints_off['IRA'].add((ira_last_weekend, day))
                
                # Section 4: Additional requests
                for req in st.session_state.additional_requests:
                    if req['day'] <= scheduler.days_in_month:
                        if req['hard']:
                            scheduler.special_requests_off[req['section']].add((req['rad'], req['day']))
                        else:
                            scheduler.soft_constraints_off[req['section']].add((req['rad'], req['day']))
            
            # Generate schedule
            with st.spinner("Generating schedule..."):
                output_path = scheduler.create_schedule()
            
            # CHECK AND SHOW DEBUG INFO
            with debug_expander:
                st.write("### Quality Metrics Check")
                has_quality = hasattr(scheduler, 'mri_quality_metrics') and scheduler.mri_quality_metrics
                st.write(f"- Has mri_quality_metrics: {has_quality}")
                if has_quality:
                    st.json(scheduler.mri_quality_metrics)
                else:
                    st.error("❌ Quality metrics not found - check your scheduler code")
                
                st.write("### Variance Results Check")
                has_variance = hasattr(scheduler, 'ytd_variance_results') and scheduler.ytd_variance_results
                st.write(f"- Has ytd_variance_results: {has_variance}")
                if has_variance:
                    st.write("- Keys:", list(scheduler.ytd_variance_results.keys()))
                else:
                    st.error("❌ Variance results not found - check your scheduler code")
            
            # DISPLAY QUALITY METRICS
            st.markdown("---")
            st.subheader("📊 MRI Assignment Quality Assessment")
            
            quality_metrics = getattr(scheduler, 'mri_quality_metrics', None)
            
            if quality_metrics and isinstance(quality_metrics, dict):
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.metric("2-Rad Days", quality_metrics.get('two_rad_days', 0),
                             help="Days where GEN or IRA can do MRI")
                
                with col2:
                    three_rad_count = quality_metrics.get('three_rad_days', 0)
                    st.metric("3-Rad Days", three_rad_count,
                             help="Days needing Python assignment")
                
                with col3:
                    opt_level = quality_metrics.get('optimization_level', 'Unknown')
                    st.metric("Optimization", opt_level.replace('✓', '').replace('⚠', '').strip())
                
                if three_rad_count > 0:
                    col1, col2 = st.columns(2)
                    with col1:
                        st.info(f"🔹 Weekends: {quality_metrics.get('three_rad_weekends', 0)} triplet(s)")
                    with col2:
                        st.info(f"🔹 Weekdays: {quality_metrics.get('three_rad_weekdays', 0)} day(s)")
                
                if three_rad_count == 0:
                    st.success("🎉 **PERFECTLY OPTIMIZED**: Zero 3-rad days!")
                elif three_rad_count <= 3:
                    st.success("✅ **WELL OPTIMIZED**: Minimal 3-rad days")
                elif three_rad_count <= 6:
                    st.warning("⚠️ **MODERATELY OPTIMIZED**: Could be improved")
                else:
                    st.error("⚠️ **POORLY OPTIMIZED**: Needs improvement")
            else:
                st.error("❌ Quality metrics not available - expand Debug Information above")
            
            # DISPLAY VARIANCE ANALYSIS
            st.markdown("---")
            st.subheader("📈 YTD Variance Analysis")
            
            variance_results = getattr(scheduler, 'ytd_variance_results', None)
            
            if variance_results and isinstance(variance_results, dict):
                summary = variance_results.get('summary', {})
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    st.metric("Weekend RMSE", f"{summary.get('weekend', {}).get('rmse', 0):.2f}")
                with col2:
                    st.metric("Thursday RMSE", f"{summary.get('thu', {}).get('rmse', 0):.2f}")
                with col3:
                    st.metric("Weekday RMSE", f"{summary.get('weekday', {}).get('rmse', 0):.2f}")
                with col4:
                    st.metric("Overall Score", f"{variance_results.get('overall_score', 0):.2f}")
                
                overall_score = variance_results.get('overall_score', 0)
                if overall_score < 1.5:
                    st.success("✅ **EXCELLENT** balance!")
                elif overall_score < 2.5:
                    st.success("✅ **GOOD** balance")
                elif overall_score < 3.5:
                    st.warning("⚠️ **FAIR** balance")
                else:
                    st.error("⚠️ **POOR** balance")
            else:
                st.error("❌ Variance results not available - expand Debug Information above")
            
            # Assignment Summary
            st.markdown("---")
            st.subheader("📋 Assignment Summary")
            
            gen_count = len(scheduler.assignments.get('GEN', {}))
            ira_count = len(scheduler.assignments.get('IRA', {}))
            mri_count = len(scheduler.assignments.get('MRI', {}))
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("GEN", gen_count)
            with col2:
                st.metric("IRA", ira_count)
            with col3:
                st.metric("MRI", f"{mri_count} (3-rad)")
            
            if gen_count == scheduler.days_in_month and ira_count == scheduler.days_in_month:
                st.success("✅ All days fully assigned!")
            
            # Download button
            st.markdown("---")
            with open(output_path, 'rb') as f:
                excel_data = f.read()
            
            month_name = calendar.month_name[scheduler.month]
            st.download_button(
                label="📥 Download Generated Schedule",
                data=excel_data,
                file_name=f"OnCall_Schedule_{month_name}_{scheduler.year}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
            
        except Exception as e:
            st.error(f"❌ Error: {str(e)}")
            with st.expander("🔍 View Error Details"):
                import traceback
                st.code(traceback.format_exc())

st.markdown("---")

# Section 3: Convert Rad Schedules for Import
st.header("🔄 Convert Rad Schedules for Import")
st.markdown("Upload completed Work Schedule and On-Call Schedule files to generate the import CSV file.")

col1, col2 = st.columns(2)

with col1:
    work_file = st.file_uploader("1. Work Schedule (Excel)", type=['xlsx'], key='work_schedule')

with col2:
    oncall_file = st.file_uploader("2. On-Call Schedule (Excel)", type=['xlsx'], key='oncall_schedule')

if work_file and oncall_file:
    st.markdown("---")
    st.subheader("📅 Select Processing Month")
    
    col1, col2 = st.columns(2)
    with col1:
        selected_month = st.selectbox("Month", options=list(range(1, 13)),
                                     format_func=lambda x: calendar.month_name[x], index=10)
    with col2:
        selected_year = st.number_input("Year", min_value=2020, max_value=2030, value=2025, step=1)
    
    st.info(f"📅 Will process: **{calendar.month_name[selected_month]} {selected_year}**")
    
    if st.button("🔄 Convert to Import Format", type="primary"):
        try:
            with st.spinner("Processing schedules..."):
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_work:
                    tmp_work.write(work_file.getvalue())
                    work_path = tmp_work.name
                
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_oncall:
                    tmp_oncall.write(oncall_file.getvalue())
                    oncall_path = tmp_oncall.name
                
                wb_work = openpyxl.load_workbook(work_path)
                ws_work = wb_work['WORK SCHEDULE']
                
                wb_oncall = openpyxl.load_workbook(oncall_path, data_only=True)
                ws_oncall = wb_oncall['Sheet1']
                
                output_data = rad_converter.process_schedules(ws_work, ws_oncall, selected_year, selected_month)
                
                st.success(f"✅ Generated {len(output_data)} schedule entries")
                
                output = io.StringIO()
                fieldnames = ['EMPLOYEE', 'TEAM', 'STARTDATE', 'STARTTIME', 
                             'ENDDATE', 'ENDTIME', 'ROLE', 'NOTES', 'ORDER', 'TEAMCOMMENT']
                writer = csv.DictWriter(output, fieldnames=fieldnames, delimiter='^')
                writer.writeheader()
                writer.writerows(output_data)
                
                csv_data = output.getvalue()
                
                st.download_button(
                    label="📥 Download Import_OnCall_Radiology.csv",
                    data=csv_data,
                    file_name="Epic_Import_OnCall_Radiology.csv",
                    mime="text/csv",
                    type="primary"
                )
                
                os.unlink(work_path)
                os.unlink(oncall_path)
                
        except Exception as e:
            st.error(f"❌ Error: {str(e)}")
            import traceback
            with st.expander("🔍 View Error Details"):
                st.code(traceback.format_exc())
else:
    st.info("👆 Please upload both Excel files to begin conversion")

st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: gray; font-size: 0.8em;'>"
    "Schedule Management System | Powered by Streamlit"
    "</div>",
    unsafe_allow_html=True
)
