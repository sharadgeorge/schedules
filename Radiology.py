"""
Streamlit App for Enhanced OnCall Scheduler
Updated to integrate with oncall_scheduler_enhanced_streamlit.py
"""

import streamlit as st
import sys
import os
from pathlib import Path
import tempfile
import io
import calendar
from datetime import datetime
from collections import defaultdict

# Import the enhanced scheduler
import oncall_scheduler_enhanced_streamlit as scheduler_module

# Configure page
st.set_page_config(
    page_title="Radiology - Schedule Management",
    page_icon="🩻",
    layout="centered",
    initial_sidebar_state="expanded"
)

# Sidebar
with st.sidebar:
    st.info("**Project by JA RAD**")
    st.markdown("---")
    st.markdown("### ✨ Enhanced Features")
    st.success("✅ Dynamic YTD Targets")
    st.success("✅ Partial Weekend Consolidation")
    st.success("✅ Intelligent Load Balancing")

# Main page title
st.title("🩻 Radiology Schedule Management")
st.markdown("### Enhanced OnCall Scheduler")
st.markdown("---")

# Section: Create Rad On-Call Schedule
st.header("📅 Create Rad On-Call Schedule")
st.markdown("Upload a blank schedule template to generate an optimized on-call schedule with **dynamic YTD target calculations**.")

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
    
    st.markdown("---")
    
    # Generate button
    if st.button("🚀 Generate Schedule", type="primary", use_container_width=True):
        try:
            with st.spinner("Generating optimized schedule with dynamic YTD targets..."):
                # Save uploaded file temporarily
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
                    tmp.write(oncall_template.getvalue())
                    temp_path = tmp.name
                
                # Create enhanced scheduler instance
                scheduler = scheduler_module.OnCallScheduler(temp_path)
                
                # Apply user preferences programmatically
                scheduler.special_requests_off = defaultdict(set)
                scheduler.soft_constraints_off = defaultdict(set)
                
                # Section 1: GEN last weekend
                last_weekend_gen = [r for r in [gen_last_weekend_1, gen_last_weekend_2] if r]
                
                # Find first weekend of current month (Friday + Saturday)
                first_weekend_days = set()
                for day in range(1, min(15, scheduler.days_in_month + 1)):
                    date = datetime(scheduler.year, scheduler.month, day)
                    if date.weekday() == 4:  # Friday
                        first_weekend_days.add(day)
                        if day + 1 <= scheduler.days_in_month:
                            first_weekend_days.add(day + 1)  # Saturday
                        break
                
                for rad in last_weekend_gen:
                    scheduler.special_requests_off[rad] = first_weekend_days
                
                # Section 2: GEN last day
                if gen_last_day:
                    if gen_last_day not in last_weekend_gen:
                        scheduler.special_requests_off[gen_last_day] = {1, 2}
                
                # Section 3: IRA last weekend (soft constraint)
                if ira_last_weekend:
                    scheduler.soft_constraints_off[ira_last_weekend] = set(range(1, min(8, scheduler.days_in_month + 1)))
                
                # Capture console output
                import io
                import sys
                old_stdout = sys.stdout
                sys.stdout = captured_output = io.StringIO()
                
                # Generate the schedule
                output_path = scheduler.generate_schedule()
                
                # Restore stdout
                sys.stdout = old_stdout
                console_output = captured_output.getvalue()
                
                # Success message
                st.success("✅ Schedule generated successfully!")
                
                # Display key metrics
                st.markdown("---")
                st.subheader("📊 Schedule Summary")
                
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    gen_total = len(scheduler.assignments['GEN'])
                    st.metric("GEN Days Assigned", gen_total)
                
                with col2:
                    ira_total = len(scheduler.assignments['IRA'])
                    st.metric("IRA Days Assigned", ira_total)
                
                with col3:
                    mri_total = len(scheduler.assignments['MRI'])
                    st.metric("MRI 3-Rad Days", mri_total)
                
                # Display YTD targets info
                st.markdown("---")
                st.subheader("🎯 YTD Targets Calculated")
                st.info("✅ Dynamic YTD targets calculated and written to Excel columns AQ, AR, AS")
                
                # Show sample targets
                with st.expander("View Sample YTD Targets"):
                    st.markdown("**GEN Section Targets:**")
                    for rad in ['NN', 'MB', 'PR', 'AT']:
                        if ('GEN', rad) in scheduler.ytd_targets:
                            targets = scheduler.ytd_targets[('GEN', rad)]
                            st.caption(f"• {rad}: WD={targets['weekday']:.2f}, Thu={targets['thu']:.2f}, WE={targets['weekend']:.2f}")
                    
                    st.markdown("**IRA Section Targets:**")
                    for rad in ['IG', 'MF', 'AS']:
                        if ('IRA', rad) in scheduler.ytd_targets:
                            targets = scheduler.ytd_targets[('IRA', rad)]
                            st.caption(f"• {rad}: WD={targets['weekday']:.2f}, Thu={targets['thu']:.2f}, WE={targets['weekend']:.2f}")
                
                # Display partial weekend consolidations
                st.markdown("---")
                st.subheader("🔄 Weekend Consolidations")
                
                # Check for partial weekend consolidations
                consolidations_found = "partial weekend consolidation" in console_output.lower()
                
                if consolidations_found:
                    st.success("✅ Partial weekend consolidations detected!")
                    # Extract consolidation info from console output
                    consolidation_lines = [line for line in console_output.split('\n') 
                                          if 'Partial weekend' in line or 'Consolidating' in line]
                    for line in consolidation_lines:
                        st.caption(f"• {line.strip()}")
                else:
                    st.info("ℹ️ No partial weekend consolidations in this month")
                
                # Display monthly assignment summary
                st.markdown("---")
                st.subheader("📋 Monthly Assignment Summary")
                
                summary_data = []
                for rad in sorted(set(scheduler_module.GEN_RADS_WITH_IRA + scheduler_module.IRA_RADS + scheduler_module.MRI_RADS)):
                    gen_count = scheduler.gen_monthly_total[rad]
                    ira_count = scheduler.ira_monthly_total[rad]
                    mri_count = scheduler.mri_monthly_total[rad]
                    
                    if gen_count > 0 or ira_count > 0 or mri_count > 0:
                        counts = scheduler.monthly_counts[rad]
                        summary_data.append({
                            'Rad': rad,
                            'GEN': gen_count,
                            'GEN_WD': counts['weekday'],
                            'GEN_Thu': counts['thu'],
                            'GEN_WE': counts['weekend'],
                            'IRA': ira_count,
                            'MRI': mri_count
                        })
                
                if summary_data:
                    import pandas as pd
                    df = pd.DataFrame(summary_data)
                    st.dataframe(df, use_container_width=True)
                
                # Download button
                st.markdown("---")
                
                # Read the generated file
                with open(output_path, 'rb') as f:
                    file_data = f.read()
                
                # Get filename
                filename = Path(output_path).name
                
                st.download_button(
                    label="📥 Download Completed Schedule",
                    data=file_data,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
                # Show console output in expander
                with st.expander("📝 View Detailed Console Output"):
                    st.code(console_output, language='text')
                
                # Clean up temp file
                try:
                    os.unlink(temp_path)
                    os.unlink(output_path)
                except:
                    pass
        
        except Exception as e:
            st.error(f"❌ Error generating schedule: {str(e)}")
            st.exception(e)

# Enhanced features info
st.markdown("---")
st.markdown("### ✨ Enhanced Features")

col1, col2 = st.columns(2)

with col1:
    st.markdown("**🎯 Dynamic YTD Targets**")
    st.caption("• Calculated based on start dates")
    st.caption("• Adjusts for mid-year hires")
    st.caption("• Excludes holidays")
    st.caption("• Written to Excel columns AQ, AR, AS")
    
    st.markdown("**🧮 Intelligent Balancing**")
    st.caption("• High YTD → fewer assignments")
    st.caption("• Low YTD → more assignments")
    st.caption("• Natural convergence to targets")

with col2:
    st.markdown("**🔄 Weekend Consolidation**")
    st.caption("• Detects partial weekends")
    st.caption("• Consolidates MRI-capable rads")
    st.caption("• Reduces fragmentation")
    st.caption("• More efficient coverage")
    
    st.markdown("**⚖️ Monthly Limits**")
    st.caption("• GEN: Max 5 days/month")
    st.caption("• IRA: Max 12 days/month")
    st.caption("• MRI: Max 5 days/month")
    st.caption("• Prevents overload")

st.markdown("---")
st.caption("Enhanced OnCall Scheduler v1.0 | JA RAD Project")
