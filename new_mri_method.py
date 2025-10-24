# NEW MRI ASSIGNMENT METHOD - To replace assign_mri_optimized() in create_oncall_schedule_v3_FIXED.py
# This should be inserted around line 964

def assign_mri_3rad_days_only(self):
    """
    NEW APPROACH: Only assign MRI for 3-rad days (where neither GEN nor IRA can do MRI)
    Let Excel formulas handle 2-rad days automatically
    
    Strategy 3: For 3-rad weekend days, if a GEN rad with MRI capability is already 
                on GEN for ANY day of that weekend, assign them to MRI for the 
                ENTIRE weekend (counts as 2-rad for MRI-only distribution)
    Strategy 4: For remaining 3-rad days, assign a different rad with balanced 
                distribution (max 1 full weekend triplet, max 2 weekday-only per month)
    """
    print("\n=== Identifying 3-Rad Days (where MRI needs assignment) ===")
    print("Note: 2-rad days will be handled by Excel formulas")
    
    # Step 1: Identify which days are 3-rad days
    three_rad_days = []
    two_rad_days = []
    
    for day in range(1, self.days_in_month + 1):
        gen_rad = self.assignments['GEN'].get(day)
        ira_rad = self.assignments['IRA'].get(day)
        
        # Check if this is a 2-rad day (either GEN or IRA can do MRI)
        gen_can_do_mri = gen_rad and gen_rad in MRI_RADS and gen_rad != ira_rad
        ira_can_do_mri = ira_rad and ira_rad in MRI_RADS
        
        if gen_can_do_mri or ira_can_do_mri:
            two_rad_days.append(day)
            formula_handler = gen_rad if gen_can_do_mri else ira_rad
            print(f"  Day {day}: 2-rad day (formula will assign {formula_handler})")
        else:
            three_rad_days.append(day)
            print(f"  Day {day}: 3-rad day (needs Python assignment)")
    
    print(f"\nSummary: {len(two_rad_days)} 2-rad days, {len(three_rad_days)} 3-rad days")
    
    if len(three_rad_days) == 0:
        print("✓ No 3-rad days! Perfect optimization!")
        return
    
    # Step 2: Group 3-rad days into weekends (Thu-Fri-Sat triplets) and weekdays
    three_rad_weekends = []  # List of (thu, fri, sat) tuples
    three_rad_weekdays = []  # Individual days
    
    print("\n=== Grouping 3-Rad Days ===")
    
    for day in range(1, self.days_in_month + 1):
        if self.get_day_type(day) == 'thu':
            fri_day = day + 1
            sat_day = day + 2
            if sat_day <= self.days_in_month:
                # Check if all three days are 3-rad days
                if day in three_rad_days and fri_day in three_rad_days and sat_day in three_rad_days:
                    three_rad_weekends.append((day, fri_day, sat_day))
                    print(f"  3-rad weekend: Days {day}-{sat_day} (Thu-Fri-Sat)")
    
    # Weekdays are 3-rad days not part of weekends
    weekend_days_covered = set()
    for weekend in three_rad_weekends:
        weekend_days_covered.update(weekend)
    
    three_rad_weekdays = [d for d in three_rad_days if d not in weekend_days_covered]
    print(f"  3-rad weekdays: {three_rad_weekdays}")
    
    print(f"\nTotal: {len(three_rad_weekends)} 3-rad weekends, {len(three_rad_weekdays)} 3-rad weekdays")
    
    # Step 3: Track MRI-only assignments for balancing
    # Counts as MRI-only ONLY if rad doesn't have GEN on that day/weekend
    mri_only_weekend_triplets = defaultdict(int)  # Full weekend triplets
    mri_only_weekdays = defaultdict(int)  # Individual weekdays
    
    # Step 4: Assign Strategy 3 - For 3-rad weekends, use GEN rad with MRI if possible
    print("\n=== Strategy 3: Assigning 3-Rad Weekends ===")
    
    for thu_day, fri_day, sat_day in three_rad_weekends:
        # Check if any GEN rad with MRI capability is on GEN this weekend
        gen_thu = self.assignments['GEN'].get(thu_day)
        gen_fri = self.assignments['GEN'].get(fri_day)
        gen_sat = self.assignments['GEN'].get(sat_day)
        
        # Find if any of these GEN rads can do MRI
        candidate_gen_rad = None
        for gen_rad in [gen_thu, gen_fri, gen_sat]:
            if gen_rad and gen_rad in MRI_RADS:
                candidate_gen_rad = gen_rad
                break
        
        if candidate_gen_rad:
            # Strategy 3: Assign this GEN rad to MRI for entire weekend
            self.assignments['MRI'][thu_day] = candidate_gen_rad
            self.assignments['MRI'][fri_day] = candidate_gen_rad
            self.assignments['MRI'][sat_day] = candidate_gen_rad
            
            # Update counts
            self.monthly_counts[candidate_gen_rad]['thu'] += 1
            self.monthly_counts[candidate_gen_rad]['weekend'] += 2
            self.mri_monthly_total[candidate_gen_rad] += 3
            
            # This counts as 2-rad for MRI-only distribution (they're already local)
            print(f"  Days {thu_day}-{sat_day} -> {candidate_gen_rad} (Strategy 3: already on GEN, counts as 2-rad)")
        else:
            # Strategy 4: Need to assign a different rad (true 3-rad day)
            # Find available MRI rads who aren't on GEN or IRA this weekend
            available_rads = []
            for rad in MRI_RADS:
                # Check if rad is available for all 3 days
                available_all_three = True
                for day in [thu_day, fri_day, sat_day]:
                    if not self.is_available(rad, day, 'MRI'):
                        available_all_three = False
                        break
                    # Also check they're not on GEN/IRA these days
                    if self.assignments['GEN'].get(day) == rad or self.assignments['IRA'].get(day) == rad:
                        available_all_three = False
                        break
                
                if available_all_three:
                    available_rads.append(rad)
            
            if not available_rads:
                print(f"  WARNING: No available MRI rads for days {thu_day}-{sat_day}")
                available_rads = MRI_RADS  # Fallback
            
            # Balance: Prefer rads with fewer MRI-only weekend triplets
            # Max 1 full weekend triplet per rad per month
            best_rad = None
            min_count = float('inf')
            for rad in available_rads:
                if mri_only_weekend_triplets[rad] < min_count:
                    min_count = mri_only_weekend_triplets[rad]
                    best_rad = rad
            
            if mri_only_weekend_triplets[best_rad] >= 1:
                print(f"  WARNING: {best_rad} already has {mri_only_weekend_triplets[best_rad]} MRI-only weekend(s)")
            
            # Assign
            self.assignments['MRI'][thu_day] = best_rad
            self.assignments['MRI'][fri_day] = best_rad
            self.assignments['MRI'][sat_day] = best_rad
            
            # Update counts
            self.monthly_counts[best_rad]['thu'] += 1
            self.monthly_counts[best_rad]['weekend'] += 2
            self.mri_monthly_total[best_rad] += 3
            mri_only_weekend_triplets[best_rad] += 1
            
            print(f"  Days {thu_day}-{sat_day} -> {best_rad} (Strategy 4: true 3-rad, MRI-only weekend #{mri_only_weekend_triplets[best_rad]})")
    
    # Step 5: Assign Strategy 4 - For 3-rad weekdays
    print("\n=== Strategy 4: Assigning 3-Rad Weekdays ===")
    
    for day in three_rad_weekdays:
        day_type = self.get_day_type(day)
        
        # Find available MRI rads
        available_rads = []
        for rad in MRI_RADS:
            if not self.is_available(rad, day, 'MRI'):
                continue
            # Check they're not on GEN/IRA this day
            if self.assignments['GEN'].get(day) == rad or self.assignments['IRA'].get(day) == rad:
                continue
            available_rads.append(rad)
        
        if not available_rads:
            print(f"  WARNING: No available MRI rads for day {day}")
            available_rads = MRI_RADS  # Fallback
        
        # Balance: Max 2 MRI-only weekdays per rad per month
        best_rad = None
        min_count = float('inf')
        for rad in available_rads:
            if mri_only_weekdays[rad] < min_count:
                min_count = mri_only_weekdays[rad]
                best_rad = rad
        
        if mri_only_weekdays[best_rad] >= 2:
            print(f"  WARNING: {best_rad} already has {mri_only_weekdays[best_rad]} MRI-only weekday(s)")
        
        # Assign
        self.assignments['MRI'][day] = best_rad
        self.monthly_counts[best_rad][day_type] += 1
        self.mri_monthly_total[best_rad] += 1
        mri_only_weekdays[best_rad] += 1
        
        print(f"  Day {day} ({day_type}) -> {best_rad} (MRI-only weekday #{mri_only_weekdays[best_rad]})")
    
    # Step 6: Final Summary
    print("\n=== MRI Assignment Summary ===")
    print(f"2-rad days (handled by Excel formulas): {len(two_rad_days)}")
    print(f"3-rad days (assigned by Python): {len(three_rad_days)}")
    print(f"  - 3-rad weekends: {len(three_rad_weekends)} triplets")
    print(f"  - 3-rad weekdays: {len(three_rad_weekdays)} days")
    
    # Display MRI-only distribution
    print("\nMRI-Only Assignment Distribution:")
    print("-" * 60)
    print(f"{'Rad':<6} {'Weekend Triplets':<20} {'Weekdays':<15}")
    print("-" * 60)
    for rad in sorted(MRI_RADS):
        weekend_count = mri_only_weekend_triplets.get(rad, 0)
        weekday_count = mri_only_weekdays.get(rad, 0)
        if weekend_count > 0 or weekday_count > 0:
            weekend_status = f"{weekend_count} (max: 1)" if weekend_count <= 1 else f"{weekend_count} (⚠ over limit)"
            weekday_status = f"{weekday_count} (max: 2)" if weekday_count <= 2 else f"{weekday_count} (⚠ over limit)"
            print(f"{rad:<6} {weekend_status:<20} {weekday_status:<15}")
    
    # Check if well-optimized
    print(f"\n{'='*60}")
    if len(three_rad_days) == 0:
        print("✓ PERFECTLY OPTIMIZED: 0 3-rad days!")
    elif len(three_rad_days) <= 3:
        print(f"✓ WELL OPTIMIZED: Only {len(three_rad_days)} 3-rad days")
    elif len(three_rad_days) <= 6:
        print(f"⚠ MODERATELY OPTIMIZED: {len(three_rad_days)} 3-rad days (could be improved)")
    else:
        print(f"⚠ POORLY OPTIMIZED: {len(three_rad_days)} 3-rad days (needs improvement)")
    print(f"{'='*60}")
