import streamlit as st
import sys
import os
from pathlib import Path
import tempfile
import calendar
from datetime import datetime, timedelta

# Configure page
st.set_page_config(
    page_title="Cardiology - Schedule Management",
    page_icon="❤",
    layout="centered",
    initial_sidebar_state="expanded"
)

# Sidebar with About section ONLY (no custom navigation)
with st.sidebar:
    st.markdown("### About")
    st.info("**Project by JA RAD**")

# Main page title - using HTML to force red color
st.markdown("# <span style='color: #e74c3c;'>❤</span> Cardiology Schedule Management", unsafe_allow_html=True)
st.markdown("---")

# Section: Convert Cardiology Schedules for Import
st.header("🔄 Convert Cardiology Schedules for Import")
st.markdown("Upload the required Cardiology Schedule Excel files to generate the import CSV file.")
st.info("ℹ️ Currently supports 2 input files. This will be expandable in the future for additional teams.")

# File uploaders
st.subheader("Upload Schedule Files (in order):")

col1, col2 = st.columns(2)

with col1:
    cardio_file1 = st.file_uploader(
        "1. Cardiovascular Schedule (Excel)", 
        type=['xlsx'],
        key='cardio_file1',
        help="Upload the Cardiovascular schedule Excel file"
    )

with col2:
    cardio_file2 = st.file_uploader(
        "2. Interventional Cardiologist Schedule (Excel)", 
        type=['xlsx'],
        key='cardio_file2',
        help="Upload the Interventional Cardiologist schedule Excel file"
    )

# Future expandability note
with st.expander("ℹ️ About Future Expansion"):
    st.markdown("""
    This converter currently processes:
    - **Team 8**: Cardiovascular (Echo Tech Adult & Pediatric)
    - **Team 94**: Interventional Cardiologist
    
    In the future, additional file upload slots can be added here for new teams.
    The converter script will need to be updated accordingly to handle the new team configurations.
    """)

# Process button and conversion logic
if cardio_file1 and cardio_file2:
    st.markdown("---")
    
    # Add month/year selection
    st.subheader("📅 Select Processing Month")
    
    col1, col2 = st.columns(2)
    with col1:
        selected_month = st.selectbox(
            "Month",
            options=list(range(1, 13)),
            format_func=lambda x: calendar.month_name[x],
            index=10,  # Default to November (index 10 for month 11)
            key='cardio_month'
        )
    
    with col2:
        selected_year = st.number_input(
            "Year",
            min_value=2020,
            max_value=2030,
            value=2025,
            step=1,
            key='cardio_year'
        )
    
    # Validate month is in correct range
    if not isinstance(selected_month, int) or selected_month < 1 or selected_month > 12:
        st.error(f"⚠️ Invalid month value: {selected_month}. Must be between 1 and 12.")
        st.stop()
    
    if st.button("🔄 Convert to Import Format", type="primary"):
        try:
            with st.spinner("Processing Cardiology schedules..."):
                # Import the converter
                import oncall_converter_Cardiology_demo_v3 as cardio_converter
                import openpyxl
                import csv
                import io
                
                # Save uploaded files temporarily
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp1:
                    tmp1.write(cardio_file1.getvalue())
                    cardio_path1 = tmp1.name
                
                with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp2:
                    tmp2.write(cardio_file2.getvalue())
                    cardio_path2 = tmp2.name
                
                # Load workbooks
                wb_cardio1 = openpyxl.load_workbook(cardio_path1, data_only=False)
                wb_cardio2 = openpyxl.load_workbook(cardio_path2, data_only=False)
                
                # Set B4 to month name for converter to read
                month_name = calendar.month_name[int(selected_month)]
                
                for sheet_name in wb_cardio1.sheetnames:
                    if month_name.lower() in sheet_name.lower() or 'nov' in sheet_name.lower():
                        ws = wb_cardio1[sheet_name]
                        ws['B4'] = month_name
                        break
                
                for sheet_name in wb_cardio2.sheetnames:
                    if month_name.lower() in sheet_name.lower() or 'nov' in sheet_name.lower():
                        ws = wb_cardio2[sheet_name]
                        ws['B4'] = month_name  
                        break
                
                # Convert schedules
                month_int = int(selected_month)
                year_int = int(selected_year)
                
                try:
                    cardio_data = cardio_converter.read_cardiovascular_data(
                        wb_cardio1,
                        month_int,
                        year_int
                    )
                except IndexError as e:
                    st.error(f"❌ IndexError: {str(e)}")
                    st.error("Month value is out of range. Please check cell B4 in your Excel file contains a valid month number (1-12).")
                    raise
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")
                    import traceback
                    with st.expander("🔍 View Error Details"):
                        st.code(traceback.format_exc())
                    raise
                
                intv_data = cardio_converter.read_interventional_data(
                    wb_cardio2,
                    month_int,
                    year_int
                )
                
                # Combine data
                output_data = cardio_converter.create_output_data(
                    cardio_data,
                    intv_data,
                    year_int,
                    month_int
                )
                
                # Show result
                if len(output_data) == 0:
                    st.error("⚠️ No entries were generated!")
                    # Still allow download of empty file for inspection
                else:
                    st.success(f"✅ **{calendar.month_name[selected_month]} {selected_year}**: Generated {len(output_data)} schedule entries")
                
                # Show detailed statistics (collapsed by default)
                with st.expander("📊 View Details"):
                    st.write(f"**Month:** {calendar.month_name[selected_month]} {selected_year}")
                    st.write(f"**Total entries:** {len(output_data)}")
                    
                    # Team breakdown
                    team_counts = {}
                    for entry in output_data:
                        team = entry['TEAM']
                        team_counts[team] = team_counts.get(team, 0) + 1
                    
                    st.write("\n**Entries per team:**")
                    team_names = {
                        '8': 'Cardiovascular',
                        '94': 'Interventional Cardiologist'
                    }
                    for team_id in sorted(team_counts.keys()):
                        team_name = team_names.get(team_id, f"Team {team_id}")
                        st.write(f"- {team_name}: {team_counts[team_id]} entries")
                
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
                    label="📥 Download Import_OnCall_Cardiology.csv",
                    data=csv_data,
                    file_name="Import_OnCall_Cardiology.csv",
                    mime="text/csv",
                    type="primary"
                )
                
                # Clean up temp files
                os.unlink(cardio_path1)
                os.unlink(cardio_path2)
                
        except Exception as e:
            st.error(f"❌ Error: {str(e)}")
            with st.expander("🔍 View Error Details"):
                import traceback
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
