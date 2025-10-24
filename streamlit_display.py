# UPDATED STREAMLIT SECTION - Replace the schedule generation section in Radiology.py
# This goes in the "if st.button('🚀 Generate Schedule')" block (around line 174)

# Inside the try block, after scheduler.create_schedule() is called:

try:
    with st.spinner("Generating optimized schedule..."):
        # Import the scheduler
        import create_oncall_schedule_v3 as scheduler_module
        from collections import defaultdict
        from datetime import timedelta
        
        # [Previous code for saving file and creating scheduler instance...]
        # ...
        
        # Generate the schedule
        scheduler.create_schedule()
        
        # Get quality metrics from MRI assignment
        # NOTE: The assign_mri_3rad_days_only method should return quality_metrics dict
        quality_metrics = getattr(scheduler, 'mri_quality_metrics', None)
        
        # Display MRI Assignment Quality
        st.markdown("---")
        st.subheader("📊 MRI Assignment Quality Assessment")
        
        if quality_metrics:
            # Create columns for metrics
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.metric(
                    "2-Rad Days", 
                    quality_metrics['two_rad_days'],
                    help="Days handled by Excel formulas (GEN or IRA can do MRI)"
                )
            
            with col2:
                st.metric(
                    "3-Rad Days", 
                    quality_metrics['three_rad_days'],
                    delta="Lower is better",
                    delta_color="inverse"
                )
            
            with col3:
                st.metric(
                    "Optimization Level",
                    quality_metrics['optimization_level'].replace('✓', '').replace('⚠', '').strip()
                )
            
            # Show breakdown
            if quality_metrics['three_rad_days'] > 0:
                st.markdown("**3-Rad Day Breakdown:**")
                col1, col2 = st.columns(2)
                with col1:
                    st.info(f"🔹 Weekends: {quality_metrics['three_rad_weekends']} triplets")
                with col2:
                    st.info(f"🔹 Weekdays: {quality_metrics['three_rad_weekdays']} days")
                
                # Show MRI-only distribution if available
                if quality_metrics.get('mri_only_distribution'):
                    with st.expander("📋 View MRI-Only Assignment Distribution"):
                        st.markdown("**Rads with MRI-only assignments (not on GEN/IRA):**")
                        
                        for rad, counts in quality_metrics['mri_only_distribution'].items():
                            weekend_count = counts['weekend_triplets']
                            weekday_count = counts['weekdays']
                            
                            weekend_status = "✓" if weekend_count <= 1 else "⚠ Over limit"
                            weekday_status = "✓" if weekday_count <= 2 else "⚠ Over limit"
                            
                            st.markdown(
                                f"- **{rad}**: {weekend_count} weekend triplets {weekend_status}, "
                                f"{weekday_count} weekdays {weekday_status}"
                            )
            else:
                st.success("✅ Perfect optimization! All MRI assignments handled by Excel formulas.")
            
            # Quality interpretation
            three_rad_count = quality_metrics['three_rad_days']
            if three_rad_count == 0:
                st.success("🎉 **PERFECTLY OPTIMIZED**: Zero 3-rad days!")
            elif three_rad_count <= 3:
                st.success("✅ **WELL OPTIMIZED**: Minimal 3-rad days")
            elif three_rad_count <= 6:
                st.warning("⚠️ **MODERATELY OPTIMIZED**: Could be improved")
            else:
                st.error("⚠️ **POORLY OPTIMIZED**: Needs significant improvement")
        
        # Calculate and display YTD Variance
        st.markdown("---")
        st.subheader("📈 YTD Variance Analysis")
        
        with st.spinner("Calculating YTD variance..."):
            # Read the generated file to get YTD values
            output_wb = openpyxl.load_workbook(output_path, data_only=True)
            output_ws = output_wb['Sheet1']
            
            # Calculate variance
            variance_results = scheduler.calculate_ytd_variance(output_ws)
            
            # Display summary metrics
            col1, col2, col3, col4 = st.columns(4)
            
            summary = variance_results['summary']
            
            with col1:
                weekend_rmse = summary['weekend']['rmse']
                weekend_color = "normal" if weekend_rmse < 1.5 else "inverse"
                st.metric(
                    "Weekend RMSE",
                    f"{weekend_rmse:.2f}",
                    delta="Lower is better",
                    delta_color=weekend_color
                )
            
            with col2:
                thu_rmse = summary['thu']['rmse']
                thu_color = "normal" if thu_rmse < 1.5 else "inverse"
                st.metric(
                    "Thursday RMSE",
                    f"{thu_rmse:.2f}",
                    delta="Lower is better",
                    delta_color=thu_color
                )
            
            with col3:
                weekday_rmse = summary['weekday']['rmse']
                weekday_color = "normal" if weekday_rmse < 1.5 else "inverse"
                st.metric(
                    "Weekday RMSE",
                    f"{weekday_rmse:.2f}",
                    delta="Lower is better",
                    delta_color=weekday_color
                )
            
            with col4:
                overall_score = variance_results['overall_score']
                overall_color = "normal" if overall_score < 2.0 else "inverse"
                st.metric(
                    "Overall Score",
                    f"{overall_score:.2f}",
                    delta="Weighted RMSE",
                    delta_color=overall_color
                )
            
            # Quality interpretation
            if overall_score < 1.5:
                st.success("✅ **EXCELLENT** balance between actual and target YTD!")
            elif overall_score < 2.5:
                st.success("✅ **GOOD** balance")
            elif overall_score < 3.5:
                st.warning("⚠️ **FAIR** balance - could be improved")
            else:
                st.error("⚠️ **POOR** balance - needs improvement")
            
            # Detailed variance table
            with st.expander("📊 View Detailed Variance Data"):
                st.markdown("**Aggregate Variance by Day Type:**")
                st.markdown("(Total Absolute Variance across all rads)")
                
                import pandas as pd
                
                variance_df = pd.DataFrame([
                    {
                        'Day Type': 'Weekend',
                        'Total Abs Variance': f"{summary['weekend']['total_abs_variance']:.2f}",
                        'Avg Abs Variance': f"{summary['weekend']['avg_abs_variance']:.3f}",
                        'RMSE': f"{summary['weekend']['rmse']:.3f}",
                        'Status': '✓' if summary['weekend']['rmse'] < 1.5 else '⚠'
                    },
                    {
                        'Day Type': 'Thursday',
                        'Total Abs Variance': f"{summary['thu']['total_abs_variance']:.2f}",
                        'Avg Abs Variance': f"{summary['thu']['avg_abs_variance']:.3f}",
                        'RMSE': f"{summary['thu']['rmse']:.3f}",
                        'Status': '✓' if summary['thu']['rmse'] < 1.5 else '⚠'
                    },
                    {
                        'Day Type': 'Weekday',
                        'Total Abs Variance': f"{summary['weekday']['total_abs_variance']:.2f}",
                        'Avg Abs Variance': f"{summary['weekday']['avg_abs_variance']:.3f}",
                        'RMSE': f"{summary['weekday']['rmse']:.3f}",
                        'Status': '✓' if summary['weekday']['rmse'] < 1.5 else '⚠'
                    }
                ])
                
                st.dataframe(variance_df, use_container_width=True, hide_index=True)
                
                st.caption("**RMSE**: Root Mean Square Error - measures average deviation from target")
                st.caption("**Overall Score**: Weighted RMSE (weekend×3 + thursday×2 + weekday×1) ÷ 6")
        
        # Display assignment summary (existing code)
        st.markdown("---")
        st.subheader("📋 Assignment Summary")
        
        # [Rest of existing code...]
        
except Exception as e:
    st.error(f"❌ Error: {str(e)}")
    with st.expander("🔍 View Error Details"):
        import traceback
        st.code(traceback.format_exc())
