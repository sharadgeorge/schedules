# UPDATED write_schedule_to_excel METHOD
# Replace the existing method around line 1093

def write_schedule_to_excel(self):
    """Write all assignments to Excel - preserving MRI formulas for 2-rad days"""
    print("\n=== Writing schedule to Excel ===")
    
    wb_write = openpyxl.load_workbook(self.excel_path)
    ws_write = wb_write['Sheet1']
    
    # Verify assignment counts before writing
    gen_days = len(self.assignments['GEN'])
    ira_days = len(self.assignments['IRA'])
    mri_days = len(self.assignments['MRI'])
    
    print(f"Assignment counts: GEN={gen_days}, IRA={ira_days}, MRI={mri_days} (3-rad days only)")
    print(f"Expected: {self.days_in_month} days for GEN and IRA")
    print(f"Note: MRI count shows only 3-rad days (2-rad days handled by formulas)")
    
    if gen_days != self.days_in_month:
        print(f"  WARNING: GEN has {gen_days} assignments (expected {self.days_in_month})")
    if ira_days != self.days_in_month:
        print(f"  WARNING: IRA has {ira_days} assignments (expected {self.days_in_month})")
    
    # Clear existing X marks for unlocked cells (GEN and IRA only, NOT MRI)
    print("\n=== Clearing existing assignments (GEN and IRA) ===")
    
    for day in range(1, self.days_in_month + 1):
        col = day + 3
        
        # Clear GEN section
        for rad, row in GEN_ROWS.items():
            if rad in ['TA']:
                continue
            if not (day in self.locked_assignments['GEN'] and self.locked_assignments['GEN'][day] == rad):
                ws_write.cell(row, col).value = None
        
        # Clear IRA section
        for rad, row in IRA_ROWS.items():
            if not (day in self.locked_assignments['IRA'] and self.locked_assignments['IRA'][day] == rad):
                ws_write.cell(row, col).value = None
        
        # DO NOT CLEAR MRI SECTION - preserve formulas for 2-rad days
        # We'll only write to specific MRI cells for 3-rad days
    
    print("GEN and IRA sections cleared (MRI formulas preserved)")
    
    # Write GEN assignments
    print("\n=== Writing GEN assignments ===")
    for day, rad in self.assignments['GEN'].items():
        if day > self.days_in_month:
            print(f"  ERROR: Skipping GEN day {day} (beyond month end)")
            continue
        row = GEN_ROWS[rad]
        col = day + 3
        ws_write.cell(row, col, 'X')
    print(f"Wrote {len(self.assignments['GEN'])} GEN assignments")
    
    # Write IRA assignments
    print("\n=== Writing IRA assignments ===")
    for day, rad in self.assignments['IRA'].items():
        if day > self.days_in_month:
            print(f"  ERROR: Skipping IRA day {day} (beyond month end)")
            continue
        row = IRA_ROWS[rad]
        col = day + 3
        ws_write.cell(row, col, 'X')
    print(f"Wrote {len(self.assignments['IRA'])} IRA assignments")
    
    # Write MRI assignments ONLY for 3-rad days
    # Excel formulas will handle 2-rad days automatically
    print("\n=== Writing MRI assignments (3-rad days only) ===")
    if len(self.assignments['MRI']) == 0:
        print("✓ No 3-rad days! All MRI assignments handled by formulas")
    else:
        # First, clear MRI cells for 3-rad days only (to override formulas)
        for day in self.assignments['MRI'].keys():
            col = day + 3
            for rad, row in MRI_ROWS.items():
                # Clear this specific day's MRI cells
                if not (day in self.locked_assignments['MRI'] and self.locked_assignments['MRI'][day] == rad):
                    ws_write.cell(row, col).value = None
        
        # Now write the 3-rad day assignments
        for day, rad in self.assignments['MRI'].items():
            if day > self.days_in_month:
                print(f"  ERROR: Skipping MRI day {day} (beyond month end)")
                continue
            row = MRI_ROWS[rad]
            col = day + 3
            ws_write.cell(row, col, 'X')
            print(f"  Day {day} -> {rad} (3-rad day, overriding formula)")
        print(f"Wrote {len(self.assignments['MRI'])} MRI assignments for 3-rad days")
    
    # Update YTD totals (columns AM, AN, AO remain as formulas)
    # These are calculated by Excel and should not be modified by Python
    
    # Save output file
    base_name = Path(self.excel_path).stem
    month_name = calendar.month_name[self.month]
    output_path = Path(self.excel_path).parent / f"{base_name}_COMPLETED_{month_name}_{self.year}.xlsx"
    
    wb_write.save(output_path)
    print(f"\n✓ Schedule saved to: {output_path}")
    
    return str(output_path)
