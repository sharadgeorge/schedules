import streamlit as st
import sys
import os
from pathlib import Path
import tempfile

# Configure page
st.set_page_config(
    page_title="Cardiology - Schedule Management",
    page_icon="❤",
    layout="centered",
    initial_sidebar_state="expanded"
)

# Sidebar with About section
with st.sidebar:
    st.markdown("### About")
    st.info("**Project by JA RAD**")

# Main page title - using simple heart to avoid encoding issues
st.title("❤ Cardiology Schedule Management")
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
                wb_cardio1 = openpyxl.load_workbook(cardio_path1, data_only=True)
                wb_cardio2 = openpyxl.load_workbook(cardio_path2, data_only=True)
                
                # Get worksheets (assuming Sheet1)
                ws_cardio1 = wb_cardio1['Sheet1']
                ws_cardio2 = wb_cardio2['Sheet1']
                
                # Extract month and year from first filename
                filename = cardio_file1.name
                current_month, current_year = cardio_converter.extract_month_year_from_filename(filename)
                
                if current_month is None:
                    from datetime import datetime
                    current_month = datetime.now().month
                    st.warning(f"⚠️ Could not detect month from filename, using current month")
                
                # Process schedules
                output_data = cardio_converter.process_schedules(
                    ws_cardio1,
                    ws_cardio2,
                    current_year,
                    current_month
                )
                
                st.success(f"✅ Generated {len(output_data)} schedule entries")
                
                # Show some statistics
                team_counts = {}
                for entry in output_data:
                    team = entry['TEAM']
                    team_counts[team] = team_counts.get(team, 0) + 1
                
                with st.expander("📊 View Statistics"):
                    st.write(f"**Total entries:** {len(output_data)}")
                    st.write("**Entries per team:**")
                    for team_id in sorted(team_counts.keys()):
                        team_name = {
                            '8': 'Cardiovascular',
                            '94': 'Interventional Cardiologist'
                        }.get(team_id, f"Team {team_id}")
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
                
                st.success("✅ Conversion complete! Click the button above to download your file.")
                
                # Clean up temp files
                os.unlink(cardio_path1)
                os.unlink(cardio_path2)
                
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
