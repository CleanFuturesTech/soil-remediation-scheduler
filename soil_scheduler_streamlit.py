"""
Soil Remediation Scheduler - Streamlit Interactive App
Interactive web app for multi-phase soil remediation with capacity pooling
"""

APP_VERSION = "2.00"

import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter
import math

# ============================================================================
# Helper Functions
# ============================================================================

def is_valid_day(date, phase_name, phases_df):
    """Check if a date is valid for a given phase based on weekend rules"""
    phase_settings = phases_df[phases_df['Phase'] == phase_name].iloc[0]
    day_of_week = date.weekday()  # Monday=0, Sunday=6
    
    # Weekdays (Mon-Fri) are always valid
    if day_of_week < 5:
        return True
    
    # Saturday
    if day_of_week == 5:
        saturday_val = str(phase_settings['Saturday']).strip().lower()
        return saturday_val == 'yes'
    
    # Sunday
    if day_of_week == 6:
        sunday_val = str(phase_settings['Sunday']).strip().lower()
        return sunday_val == 'yes'
    
    return False


# ============================================================================
# Cell Flip Class
# ============================================================================

class CellFlip:
    """Tracks the state of a single cell flip"""
    def __init__(self, flip_num, cell_num, soil_needed):
        self.flip_num = flip_num
        self.cell_num = cell_num
        self.soil_needed = soil_needed
        self.soil_loaded = 0
        self.load_complete = False
        self.load_complete_date = None
        self.current_phase = 'Loading'
        
    def add_soil(self, amount):
        """Add soil to this cell"""
        self.soil_loaded += amount
        if self.soil_loaded >= self.soil_needed:
            self.load_complete = True
            
    def remaining_capacity(self):
        """How much more soil can this cell accept?"""
        return max(0, self.soil_needed - self.soil_loaded)
    
    def is_loading_complete(self):
        """Check if loading is done"""
        return self.load_complete


# ============================================================================
# Main Simulation Function
# ============================================================================

def simulate_remediation(params, phases_df):
    """Run the soil remediation simulation with capacity pooling"""
    
    total_flips = math.ceil(params['TotalSoil_CY'] / params['CellSize_CY'])
    num_cells = int(params['NumCells'])
    daily_load_capacity = params['DailyLoad_CY']
    daily_unload_capacity = params['DailyUnload_CY']
    start_date = params['StartDate']
    
    all_activities = []
    
    # Track all flips
    flips_in_progress = []
    flips_completed_loading = []
    total_soil_processed = 0
    
    current_date = start_date
    flip_counter = 1
    
    # Phase durations
    rip_days = int(phases_df[phases_df['Phase'] == 'Rip']['Duration_Days'].iloc[0])
    treat_days = int(phases_df[phases_df['Phase'] == 'Treat']['Duration_Days'].iloc[0])
    dry_days = int(phases_df[phases_df['Phase'] == 'Dry']['Duration_Days'].iloc[0])
    
    # Track which cells are in use and when they'll be free
    cell_availability = {i: start_date for i in range(1, num_cells + 1)}
    
    # Track total soil loaded and unloaded
    total_soil_loaded = 0
    total_soil_unloaded = 0
    
    # Continue until all soil is unloaded
    max_days = 1000
    day_count = 0
    
    while total_soil_unloaded < params['TotalSoil_CY'] and day_count < max_days:
        day_count += 1
        
        # ================== UNLOADING PHASE (PRIORITY 1) ==================
        # Unloading happens FIRST to free up cells and get priority on loader capacity
        if is_valid_day(current_date, 'Unload', phases_df):
            # Check how much capacity was already used by loading today (should be 0 since unload is first)
            todays_load = sum(a['SoilIn'] for a in all_activities if a['Date'] == current_date)
            remaining_unload_capacity = max(0, daily_unload_capacity - todays_load)
            
            # Find flips ready to unload (Dry completed on a previous day)
            ready_to_unload = [f for f in flips_completed_loading if 
                        (f.current_phase == 'DryComplete' and hasattr(f, 'dry_complete_date') and current_date > f.dry_complete_date) or
                        (hasattr(f, 'unload_started') and f.unload_started and not hasattr(f, 'unload_complete'))]
            
            # Sort by flip number (earlier flips get priority)
            ready_to_unload.sort(key=lambda f: f.flip_num)
            
            for flip in ready_to_unload:
                
                # Mark this flip as ready to unload now
                if flip.current_phase == 'DryComplete':
                    flip.current_phase = 'ReadyToUnload'
                
                if not hasattr(flip, 'unload_started'):
                    flip.unload_started = True
                    flip.soil_unloaded = 0
                
                # Don't unload more than total needed
                max_we_can_unload = params['TotalSoil_CY'] - total_soil_unloaded
                if max_we_can_unload <= 0:
                    break
                
                # How much can we unload today?
                remaining_in_cell = flip.soil_needed - flip.soil_unloaded
                amount_to_unload = min(remaining_in_cell, remaining_unload_capacity, max_we_can_unload)
                
                if amount_to_unload > 0:
                    flip.soil_unloaded += amount_to_unload
                    remaining_unload_capacity -= amount_to_unload
                    total_soil_unloaded += amount_to_unload
                    
                    all_activities.append({
                        'Date': current_date,
                        'Phase': f'Unload ({int(amount_to_unload)})',
                        'CellNum': flip.cell_num,
                        'FlipNum': flip.flip_num,
                        'SoilIn': 0,
                        'SoilOut': amount_to_unload
                    })
                    
                    # Check if unloading is complete
                    if flip.soil_unloaded >= flip.soil_needed:
                        flip.unload_complete = True
                        flips_completed_loading.remove(flip)
                        # Cell is now available for next flip
                        cell_availability[flip.cell_num] = current_date + timedelta(days=1)
                
                if remaining_unload_capacity <= 0 or total_soil_unloaded >= params['TotalSoil_CY']:
                    break
        
        # ================== LOADING PHASE (PRIORITY 2) ==================
        if is_valid_day(current_date, 'Load', phases_df) and total_soil_loaded < params['TotalSoil_CY']:
            # Check how much capacity was already used by unloading today
            todays_unload = sum(a['SoilOut'] for a in all_activities if a['Date'] == current_date)
            remaining_daily_capacity = max(0, daily_load_capacity - todays_unload)
            
            # Sort flips in progress by flip number (earlier flips get priority)
            flips_in_progress.sort(key=lambda f: f.flip_num)
            
            # Process loading for existing flips in progress
            for flip in flips_in_progress[:]:
                if flip.is_loading_complete():
                    continue
                
                # Don't load more than total needed
                max_we_can_load = params['TotalSoil_CY'] - total_soil_loaded
                if max_we_can_load <= 0:
                    break
                    
                # How much can we load today into this flip?
                space_in_cell = flip.remaining_capacity()
                amount_to_load = min(space_in_cell, remaining_daily_capacity, max_we_can_load)
                
                if amount_to_load > 0:
                    flip.add_soil(amount_to_load)
                    remaining_daily_capacity -= amount_to_load
                    total_soil_loaded += amount_to_load
                    
                    # Record this activity
                    all_activities.append({
                        'Date': current_date,
                        'Phase': f'Load ({int(amount_to_load)})',
                        'CellNum': flip.cell_num,
                        'FlipNum': flip.flip_num,
                        'SoilIn': amount_to_load,
                        'SoilOut': 0
                    })
                    
                    # Check if this flip completed loading today
                    if flip.is_loading_complete():
                        flip.load_complete_date = current_date
                        flips_completed_loading.append(flip)
                        flips_in_progress.remove(flip)
                
                if remaining_daily_capacity <= 0 or total_soil_loaded >= params['TotalSoil_CY']:
                    break
            
            # Start new flips if we have capacity and haven't loaded all soil
            while remaining_daily_capacity > 0 and total_soil_loaded < params['TotalSoil_CY']:
                # Determine next cell to use (cycle through cells in strict sequence)
                next_cell = ((flip_counter - 1) % num_cells) + 1
                
                # Check if this cell is available (not in use by any active flip)
                cell_in_use = False
                
                # Check flips still loading
                for f in flips_in_progress:
                    if f.cell_num == next_cell:
                        cell_in_use = True
                        break
                
                # Check flips that completed loading but haven't finished unloading
                for f in flips_completed_loading:
                    if f.cell_num == next_cell:
                        cell_in_use = True
                        break
                
                # Also check the cell_availability date
                if current_date < cell_availability[next_cell] or cell_in_use:
                    # Cell not available - WAIT for it (don't skip to next cell)
                    # Break out of the loop and try again tomorrow
                    break
                
                # Calculate soil needed for this flip
                remaining_soil_to_load = params['TotalSoil_CY'] - total_soil_loaded
                soil_for_flip = min(params['CellSize_CY'], remaining_soil_to_load)
                
                # Create new flip
                new_flip = CellFlip(flip_counter, next_cell, soil_for_flip)
                
                # Load what we can today
                amount_to_load = min(soil_for_flip, remaining_daily_capacity, remaining_soil_to_load)
                new_flip.add_soil(amount_to_load)
                remaining_daily_capacity -= amount_to_load
                total_soil_loaded += amount_to_load
                
                # Record activity
                all_activities.append({
                    'Date': current_date,
                    'Phase': f'Load ({int(amount_to_load)})',
                    'CellNum': new_flip.cell_num,
                    'FlipNum': new_flip.flip_num,
                    'SoilIn': amount_to_load,
                    'SoilOut': 0
                })
                
                total_soil_processed += amount_to_load
                
                # Check if flip completed loading today
                if new_flip.is_loading_complete():
                    new_flip.load_complete_date = current_date
                    flips_completed_loading.append(new_flip)
                else:
                    flips_in_progress.append(new_flip)
                
                flip_counter += 1
                
                if total_soil_loaded >= params['TotalSoil_CY']:
                    break
        
        # ================== RIP, TREAT, DRY PHASES ==================
        # Process flips that completed loading
        for flip in flips_completed_loading[:]:
            days_since_load_complete = (current_date - flip.load_complete_date).days
            
            # Rip starts day after load completes
            if days_since_load_complete == 1:
                flip.current_phase = 'Rip'
                flip.rip_start_date = current_date
            
            # Check if we're in Rip phase
            if flip.current_phase == 'Rip' and hasattr(flip, 'rip_start_date'):
                days_in_rip = (current_date - flip.rip_start_date).days
                
                # Count valid rip days
                valid_rip_days_count = 0
                for i in range(days_in_rip + 1):
                    check_date = flip.rip_start_date + timedelta(days=i)
                    if is_valid_day(check_date, 'Rip', phases_df):
                        valid_rip_days_count += 1
                
                if is_valid_day(current_date, 'Rip', phases_df):
                    all_activities.append({
                        'Date': current_date,
                        'Phase': 'Rip',
                        'CellNum': flip.cell_num,
                        'FlipNum': flip.flip_num,
                        'SoilIn': 0,
                        'SoilOut': 0
                    })
                
                # Check if Rip is complete
                if valid_rip_days_count >= rip_days:
                    flip.current_phase = 'Treat'
                    flip.treat_start_date = current_date + timedelta(days=1)
            
            # Treat phase
            if flip.current_phase == 'Treat' and hasattr(flip, 'treat_start_date'):
                if current_date >= flip.treat_start_date:
                    days_in_treat = (current_date - flip.treat_start_date).days
                    
                    valid_treat_days_count = 0
                    for i in range(days_in_treat + 1):
                        check_date = flip.treat_start_date + timedelta(days=i)
                        if is_valid_day(check_date, 'Treat', phases_df):
                            valid_treat_days_count += 1
                    
                    if is_valid_day(current_date, 'Treat', phases_df):
                        all_activities.append({
                            'Date': current_date,
                            'Phase': 'Treat',
                            'CellNum': flip.cell_num,
                            'FlipNum': flip.flip_num,
                            'SoilIn': 0,
                            'SoilOut': 0
                        })
                    
                    if valid_treat_days_count >= treat_days:
                        flip.current_phase = 'Dry'
                        flip.dry_start_date = current_date + timedelta(days=1)
            
            # Dry phase
            if flip.current_phase == 'Dry' and hasattr(flip, 'dry_start_date'):
                if current_date >= flip.dry_start_date:
                    days_in_dry = (current_date - flip.dry_start_date).days
                    
                    valid_dry_days_count = 0
                    for i in range(days_in_dry + 1):
                        check_date = flip.dry_start_date + timedelta(days=i)
                        if is_valid_day(check_date, 'Dry', phases_df):
                            valid_dry_days_count += 1
                    
                    if is_valid_day(current_date, 'Dry', phases_df):
                        all_activities.append({
                            'Date': current_date,
                            'Phase': 'Dry',
                            'CellNum': flip.cell_num,
                            'FlipNum': flip.flip_num,
                            'SoilIn': 0,
                            'SoilOut': 0
                        })
                    
                    if valid_dry_days_count >= dry_days:
                        flip.current_phase = 'DryComplete'
                        flip.dry_complete_date = current_date
        
        current_date += timedelta(days=1)
    
    return all_activities


# ============================================================================
# Build Schedule from Activities
# ============================================================================

def build_schedule(all_activities, params, phases_df, num_days=1000):
    """Build the schedule DataFrame from activities and detect idle days"""
    
    dates = [params['StartDate'] + timedelta(days=i) for i in range(num_days)]
    
    # Calculate month (4-week periods) and week numbers
    months = [(i // 28) + 1 for i in range(num_days)]
    weeks = [(i // 7) + 1 for i in range(num_days)]
    
    num_cells = int(params['NumCells'])
    
    # Create base columns
    schedule_data = {
        'Month': months,
        'Week': weeks,
        'Day Count': range(1, num_days + 1),
        'Date': dates,
        'DayName': [d.strftime('%A') for d in dates],
        'SoilIn': [0] * num_days,
        'SoilOut': [0] * num_days
    }
    
    # Add cell phase columns dynamically
    for i in range(1, num_cells + 1):
        schedule_data[f'Cell{i}Phase'] = [''] * num_days
    
    schedule = pd.DataFrame(schedule_data)
    
    # Merge activities into schedule
    for activity in all_activities:
        matching_rows = schedule['Date'] == activity['Date']
        cell_col = f"Cell{activity['CellNum']}Phase"
        
        # Get existing phase value
        existing_phase = schedule.loc[matching_rows, cell_col].values[0] if len(schedule.loc[matching_rows]) > 0 else ''
        
        # If there's already a phase on this day for this cell, append
        if existing_phase and existing_phase != '':
            new_phase = existing_phase + ' + ' + activity['Phase']
        else:
            new_phase = activity['Phase']
        
        schedule.loc[matching_rows, cell_col] = new_phase
        schedule.loc[matching_rows, 'SoilIn'] += activity['SoilIn']
        schedule.loc[matching_rows, 'SoilOut'] += activity['SoilOut']
    
    # Calculate cumulative totals
    schedule['CumSoilIn'] = schedule['SoilIn'].cumsum()
    schedule['CumSoilOut'] = schedule['SoilOut'].cumsum()
    
    # DETECT IDLE DAYS ON FULL SCHEDULE BEFORE FILTERING
    idle_days_count = detect_idle_capacity_days(schedule, params, phases_df)
    
    # Filter to relevant rows
    last_day_idx = None
    for idx, row in schedule.iterrows():
        if row['CumSoilOut'] >= params['TotalSoil_CY']:
            last_day_idx = idx
            break
    
    if last_day_idx is not None:
        schedule_filtered = schedule.iloc[:last_day_idx + 6].copy()
    else:
        schedule_filtered = schedule[
            (schedule['SoilIn'] > 0) | 
            (schedule['SoilOut'] > 0) | 
            (schedule['CumSoilIn'] > 0) | 
            (schedule['CumSoilOut'] > 0)
        ].copy()
    
    return schedule_filtered, idle_days_count


# ============================================================================
# Detect Idle Capacity Days
# ============================================================================

def detect_idle_capacity_days(schedule, params, phases_df):
    """
    Detect days where loading/unloading could occur but didn't due to no cells being ready.
    Returns idle days count.
    
    An idle day is when:
    - At least one of loading or unloading is allowed (valid work day)
    - BOTH loading AND unloading are zero (nothing happened)
    - There's still work to be done
    """
    idle_days_count = 0
    
    for idx, row in schedule.iterrows():
        # Check if this is a valid work day for load or unload
        is_valid_load_day = is_valid_day(row['Date'], 'Load', phases_df)
        is_valid_unload_day = is_valid_day(row['Date'], 'Unload', phases_df)
        
        # Check if any work happened
        no_loading = row['SoilIn'] == 0
        no_unloading = row['SoilOut'] == 0
        
        # Check if there's still work to be done
        still_soil_to_load = row['CumSoilIn'] < params['TotalSoil_CY']
        still_soil_to_unload = row['CumSoilOut'] < params['TotalSoil_CY']
        loading_has_started = row['CumSoilIn'] > 0
        
        # A day is idle if:
        # 1. It's a valid work day (for load or unload)
        # 2. Nothing happened (both load and unload are zero)
        # 3. There's still work that could be done
        could_have_worked = is_valid_load_day or is_valid_unload_day
        did_nothing = no_loading and no_unloading
        work_remains = (still_soil_to_load or (still_soil_to_unload and loading_has_started))
        
        if could_have_worked and did_nothing and work_remains:
            idle_days_count += 1
    
    return idle_days_count


# ============================================================================
# Cost Calculations
# ============================================================================

def calculate_costs(activities, schedule, params, cost_params, phases_df):
    """
    Calculate daily and total costs for the remediation project.
    
    Cost categories:
    - Equipment: daily fleet rental (excavator, loader, bulldozer, skidsteer)
                 charged for every day of project duration
    - Water purchase: BBL/day × $/BBL, where daily BBL is driven by cells in Treat today
    - Water trucking: truck-hours × $/hr, based on BBLs hauled
    - Leachate disposal: $/BBL × daily BBL, where daily BBL = pct × water_(today - lag_days)
    - Leachate trucking: truck-hours × $/hr, based on BBLs hauled
    - Amendments: $/CY, charged on first Treat day per flip (lump per flip)
    
    Water/leachate model mirrors the spreadsheet:
      water_BBL_today = water_BBL_per_CY × cell_size × N_cells_in_Treat / treat_duration
      leachate_BBL_today = leachate_pct × water_BBL_(today - lag_days)
    
    Returns:
        daily_costs_df: per-day cost breakdown
        summary: dict of total costs by category
    """
    # ---------- Per-flip soil volumes & treat schedules ----------
    flip_cy = {}            # flip_num -> total CY loaded
    flip_first_treat = {}   # flip_num -> first Treat date (for amendments)
    
    for act in activities:
        fn = act['FlipNum']
        if 'Load' in act['Phase']:
            flip_cy[fn] = flip_cy.get(fn, 0) + act['SoilIn']
        elif act['Phase'] == 'Treat':
            if fn not in flip_first_treat:
                flip_first_treat[fn] = act['Date']
    
    # ---------- Treat duration (days) ----------
    try:
        treat_duration_days = int(phases_df[phases_df['Phase'] == 'Treat']['Duration_Days'].iloc[0])
    except Exception:
        treat_duration_days = 3
    if treat_duration_days <= 0:
        treat_duration_days = 1
    
    # ---------- Cell-size for daily water calc ----------
    cell_size_cy = params.get('CellSize_CY', 0)
    
    # ---------- Initialize daily cost dataframe ----------
    daily_costs = pd.DataFrame({
        'Date': schedule['Date'].values,
        'ExcavatorCost': 0.0,
        'LoaderCost': 0.0,
        'BulldozerCost': 0.0,
        'SkidsteerCost': 0.0,
        'EquipmentCost': 0.0,
        'CellsInTreat': 0,
        'WaterBBL': 0.0,
        'WaterPurchaseCost': 0.0,
        'WaterTruckingCost': 0.0,
        'WaterCost': 0.0,
        'LeachateBBL': 0.0,
        'LeachateDisposalCost': 0.0,
        'LeachateTruckingCost': 0.0,
        'LeachateCost': 0.0,
        'AmendmentCost': 0.0,
    })
    daily_costs = daily_costs.set_index('Date')
    
    # ---------- Equipment costs (daily fleet rental for every day in schedule) ----------
    daily_excavator = cost_params.get('n_excavator', 0) * cost_params.get('excavator_daily', 0.0)
    daily_loader    = cost_params.get('n_loader', 0)    * cost_params.get('loader_daily', 0.0)
    daily_bulldozer = cost_params.get('n_bulldozer', 0) * cost_params.get('bulldozer_daily', 0.0)
    daily_skidsteer = cost_params.get('n_skidsteer', 0) * cost_params.get('skidsteer_daily', 0.0)
    daily_fleet_cost = daily_excavator + daily_loader + daily_bulldozer + daily_skidsteer
    
    for date in daily_costs.index:
        daily_costs.at[date, 'ExcavatorCost'] = daily_excavator
        daily_costs.at[date, 'LoaderCost']    = daily_loader
        daily_costs.at[date, 'BulldozerCost'] = daily_bulldozer
        daily_costs.at[date, 'SkidsteerCost'] = daily_skidsteer
        daily_costs.at[date, 'EquipmentCost'] = daily_fleet_cost
    
    # ---------- Water (per-day from cells in Treat) ----------
    water_bbl_per_cy   = cost_params.get('water_bbl_per_cy', 0.0)
    water_cost_per_bbl = cost_params.get('water_cost_per_bbl', 0.0)
    water_truck_cap    = cost_params.get('water_truck_capacity', 120)
    water_trip_hr      = cost_params.get('water_truck_trip_hr', 1.5)
    water_truck_rate   = cost_params.get('water_truck_hourly', 95.0)
    water_truck_cost_per_bbl = (water_trip_hr * water_truck_rate / water_truck_cap) if water_truck_cap > 0 else 0.0
    
    # BBL per cell-day in Treat = total_water_per_CY × cell_size / treat_duration_days
    water_bbl_per_treat_cell_day = (water_bbl_per_cy * cell_size_cy / treat_duration_days) if treat_duration_days > 0 else 0.0
    
    # Count cells in Treat each day from the schedule
    num_cells = int(params['NumCells'])
    cell_cols = [f'Cell{i}Phase' for i in range(1, num_cells + 1)]
    
    # Build a list of (date, n_treat) so we can apply leachate lag by index
    schedule_dates = list(schedule['Date'].values)
    n_treat_by_date_idx = []
    
    for idx, row in schedule.iterrows():
        date = row['Date']
        if date not in daily_costs.index:
            n_treat_by_date_idx.append(0)
            continue
        
        n_treat = 0
        for col in cell_cols:
            phase_str = str(row[col]) if pd.notna(row[col]) else ''
            if 'Treat' in phase_str:
                n_treat += 1
        
        n_treat_by_date_idx.append(n_treat)
        
        water_bbl = n_treat * water_bbl_per_treat_cell_day
        daily_costs.at[date, 'CellsInTreat'] = n_treat
        daily_costs.at[date, 'WaterBBL'] = water_bbl
        daily_costs.at[date, 'WaterPurchaseCost'] = water_bbl * water_cost_per_bbl
        daily_costs.at[date, 'WaterTruckingCost'] = water_bbl * water_truck_cost_per_bbl
        daily_costs.at[date, 'WaterCost'] = daily_costs.at[date, 'WaterPurchaseCost'] + daily_costs.at[date, 'WaterTruckingCost']
    
    # ---------- Leachate (3-day lag from water by default) ----------
    leachate_pct       = cost_params.get('leachate_pct_of_water', 80) / 100.0
    leachate_lag       = int(cost_params.get('leachate_lag_days', 3))
    leachate_cost_per_bbl = cost_params.get('leachate_cost_per_bbl', 0.0)
    leachate_truck_cap = cost_params.get('leachate_truck_capacity', 120)
    leachate_trip_hr   = cost_params.get('leachate_truck_trip_hr', 1.0)
    leachate_truck_rate = cost_params.get('leachate_truck_hourly', 95.0)
    leachate_truck_cost_per_bbl = (leachate_trip_hr * leachate_truck_rate / leachate_truck_cap) if leachate_truck_cap > 0 else 0.0
    
    n_dates = len(schedule_dates)
    for src_idx in range(n_dates):
        water_bbl_that_day = n_treat_by_date_idx[src_idx] * water_bbl_per_treat_cell_day
        if water_bbl_that_day <= 0:
            continue
        
        # Target day for leachate
        tgt_idx = min(src_idx + leachate_lag, n_dates - 1)  # clamp to last day if lag exceeds buffer
        tgt_date = schedule_dates[tgt_idx]
        if tgt_date not in daily_costs.index:
            continue
        
        leachate_bbl = water_bbl_that_day * leachate_pct
        daily_costs.at[tgt_date, 'LeachateBBL'] += leachate_bbl
        daily_costs.at[tgt_date, 'LeachateDisposalCost'] += leachate_bbl * leachate_cost_per_bbl
        daily_costs.at[tgt_date, 'LeachateTruckingCost'] += leachate_bbl * leachate_truck_cost_per_bbl
    
    daily_costs['LeachateCost'] = daily_costs['LeachateDisposalCost'] + daily_costs['LeachateTruckingCost']
    
    # ---------- Amendments (lump sum on first Treat day per flip) ----------
    amendment_cost_per_cy = cost_params.get('amendment_cost_per_cy', 0.0)
    
    for fn, cy in flip_cy.items():
        if fn in flip_first_treat:
            first_date = flip_first_treat[fn]
            if first_date in daily_costs.index:
                daily_costs.at[first_date, 'AmendmentCost'] += cy * amendment_cost_per_cy
    
    # ---------- Totals & cumulative ----------
    daily_costs['TotalCost'] = (
        daily_costs['EquipmentCost'] +
        daily_costs['WaterCost'] +
        daily_costs['LeachateCost'] +
        daily_costs['AmendmentCost']
    )
    daily_costs['CumTotalCost'] = daily_costs['TotalCost'].cumsum()
    daily_costs = daily_costs.reset_index()
    
    # ---------- Summary ----------
    total_cy = params['TotalSoil_CY']
    project_days = len(daily_costs)
    
    summary = {
        'Equipment': daily_costs['EquipmentCost'].sum(),
        'WaterPurchase':    daily_costs['WaterPurchaseCost'].sum(),
        'WaterTrucking':    daily_costs['WaterTruckingCost'].sum(),
        'Water':            daily_costs['WaterCost'].sum(),
        'LeachateDisposal': daily_costs['LeachateDisposalCost'].sum(),
        'LeachateTrucking': daily_costs['LeachateTruckingCost'].sum(),
        'Leachate':         daily_costs['LeachateCost'].sum(),
        'Amendments':       daily_costs['AmendmentCost'].sum(),
    }
    summary['Total'] = (
        summary['Equipment'] + summary['Water'] + summary['Leachate'] + summary['Amendments']
    )
    summary['CostPerCY'] = summary['Total'] / total_cy if total_cy > 0 else 0.0
    summary['ProjectDays'] = project_days
    
    # Equipment breakdown by type
    summary['EquipmentByType'] = {
        'Excavator': daily_costs['ExcavatorCost'].sum(),
        'Loader':    daily_costs['LoaderCost'].sum(),
        'Bulldozer': daily_costs['BulldozerCost'].sum(),
        'Skidsteer': daily_costs['SkidsteerCost'].sum(),
    }
    
    # Material volumes
    summary['TotalWaterBBL']    = daily_costs['WaterBBL'].sum()
    summary['TotalLeachateBBL'] = daily_costs['LeachateBBL'].sum()
    
    # Peak day metrics (useful for fleet sizing)
    summary['PeakWaterBBL_Day']    = daily_costs['WaterBBL'].max()
    summary['PeakLeachateBBL_Day'] = daily_costs['LeachateBBL'].max()
    summary['PeakCellsInTreat']    = int(daily_costs['CellsInTreat'].max())
    
    return daily_costs, summary


# ============================================================================
# Streamlit UI
# ============================================================================

def main():
    st.set_page_config(page_title=f"Soil Remediation Scheduler v{APP_VERSION}", layout="wide")
    
    # Display company logo
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.image("Clean_Futures_2.png", use_container_width=True)
    
    st.title("🏗️ Soil Remediation Scheduler")
    st.markdown(f"**Interactive multi-phase soil remediation simulator with capacity pooling** &nbsp; · &nbsp; `v{APP_VERSION}`")
    
    # Sidebar for parameters
    st.sidebar.header("📋 Project Parameters")
    
    total_soil = st.sidebar.number_input("Total Soil (CY)", min_value=100, max_value=100000, value=36000, step=100)
    cell_size = st.sidebar.number_input("Cell Size (CY)", min_value=100, max_value=10000, value=4500, step=100)
    num_cells = st.sidebar.number_input("Number of Cells", min_value=1, max_value=20, value=4)
    start_date = st.sidebar.date_input("Start Date", value=datetime(2025, 12, 1))
    
    st.sidebar.caption("ℹ️ Daily soil-movement capacity is derived from the equipment fleet below.")
    
    st.sidebar.markdown("---")
    st.sidebar.header("⚙️ Phase Settings")
    
    # Phase settings table
    phase_data = {
        'Phase': ['Load', 'Rip', 'Treat', 'Dry', 'Unload'],
        'Duration_Days': [0, 1, 3, 5, 0],
        'Saturday': ['no', 'yes', 'yes', 'yes', 'no'],
        'Sunday': ['no', 'yes', 'yes', 'yes', 'no']
    }
    
    phases_df = pd.DataFrame(phase_data)
    
    # Editable phase settings
    st.sidebar.markdown("**Load Phase**")
    load_sat = st.sidebar.checkbox("Load on Saturday", value=False, key='load_sat')
    load_sun = st.sidebar.checkbox("Load on Sunday", value=False, key='load_sun')
    
    st.sidebar.markdown("**Rip Phase**")
    rip_duration = st.sidebar.number_input("Rip Duration (days)", min_value=0, max_value=30, value=1)
    rip_sat = st.sidebar.checkbox("Rip on Saturday", value=True, key='rip_sat')
    rip_sun = st.sidebar.checkbox("Rip on Sunday", value=True, key='rip_sun')
    
    st.sidebar.markdown("**Treat Phase**")
    treat_duration = st.sidebar.number_input("Treat Duration (days)", min_value=0, max_value=90, value=3)
    treat_sat = st.sidebar.checkbox("Treat on Saturday", value=True, key='treat_sat')
    treat_sun = st.sidebar.checkbox("Treat on Sunday", value=True, key='treat_sun')
    
    st.sidebar.markdown("**Dry Phase**")
    dry_duration = st.sidebar.number_input("Dry Duration (days)", min_value=0, max_value=30, value=5)
    dry_sat = st.sidebar.checkbox("Dry on Saturday", value=True, key='dry_sat')
    dry_sun = st.sidebar.checkbox("Dry on Sunday", value=True, key='dry_sun')
    
    st.sidebar.markdown("**Unload Phase**")
    unload_sat = st.sidebar.checkbox("Unload on Saturday", value=False, key='unload_sat')
    unload_sun = st.sidebar.checkbox("Unload on Sunday", value=False, key='unload_sun')
    
    # Update phases_df with user inputs
    phases_df.loc[phases_df['Phase'] == 'Load', 'Saturday'] = 'yes' if load_sat else 'no'
    phases_df.loc[phases_df['Phase'] == 'Load', 'Sunday'] = 'yes' if load_sun else 'no'
    
    phases_df.loc[phases_df['Phase'] == 'Rip', 'Duration_Days'] = rip_duration
    phases_df.loc[phases_df['Phase'] == 'Rip', 'Saturday'] = 'yes' if rip_sat else 'no'
    phases_df.loc[phases_df['Phase'] == 'Rip', 'Sunday'] = 'yes' if rip_sun else 'no'
    
    phases_df.loc[phases_df['Phase'] == 'Treat', 'Duration_Days'] = treat_duration
    phases_df.loc[phases_df['Phase'] == 'Treat', 'Saturday'] = 'yes' if treat_sat else 'no'
    phases_df.loc[phases_df['Phase'] == 'Treat', 'Sunday'] = 'yes' if treat_sun else 'no'
    
    phases_df.loc[phases_df['Phase'] == 'Dry', 'Duration_Days'] = dry_duration
    phases_df.loc[phases_df['Phase'] == 'Dry', 'Saturday'] = 'yes' if dry_sat else 'no'
    phases_df.loc[phases_df['Phase'] == 'Dry', 'Sunday'] = 'yes' if dry_sun else 'no'
    
    phases_df.loc[phases_df['Phase'] == 'Unload', 'Saturday'] = 'yes' if unload_sat else 'no'
    phases_df.loc[phases_df['Phase'] == 'Unload', 'Sunday'] = 'yes' if unload_sun else 'no'
    
    # ============================================================
    # Equipment Fleet & Cost Inputs
    # ============================================================
    st.sidebar.markdown("---")
    st.sidebar.header("🚜 Equipment Fleet")
    
    with st.sidebar.expander("Excavator (excavates soil, gates load/unload)", expanded=True):
        n_excavator = st.number_input("# Excavators", min_value=1, max_value=10, value=1, step=1, key='n_excavator')
        excavator_capacity = st.number_input("CY/day per excavator", min_value=50, max_value=5000, value=750, step=50, key='excavator_capacity')
        excavator_daily = st.number_input("$/day per excavator", min_value=0.0, value=800.0, step=50.0, key='excavator_daily')
    
    with st.sidebar.expander("Rubber Tire Loader (moves piles)", expanded=True):
        n_loader = st.number_input("# Loaders", min_value=1, max_value=10, value=1, step=1, key='n_loader')
        loader_capacity = st.number_input("CY/day per loader", min_value=50, max_value=5000, value=1000, step=50, key='loader_capacity')
        loader_daily = st.number_input("$/day per loader", min_value=0.0, value=600.0, step=50.0, key='loader_daily')
    
    with st.sidebar.expander("Bulldozer (levels for treat, rips cells)", expanded=False):
        n_bulldozer = st.number_input("# Bulldozers", min_value=0, max_value=10, value=1, step=1, key='n_bulldozer')
        bulldozer_daily = st.number_input("$/day per bulldozer", min_value=0.0, value=900.0, step=50.0, key='bulldozer_daily')
        st.caption("Bulldozer is on site daily; doesn't gate soil throughput.")
    
    with st.sidebar.expander("Skidsteer (on site, not in soil processing)", expanded=False):
        n_skidsteer = st.number_input("# Skidsteers", min_value=0, max_value=10, value=1, step=1, key='n_skidsteer')
        skidsteer_daily = st.number_input("$/day per skidsteer", min_value=0.0, value=300.0, step=25.0, key='skidsteer_daily')
    
    # Derived daily soil-movement capacity (bottleneck of excavator vs loader fleet)
    excavator_fleet_capacity = n_excavator * excavator_capacity
    loader_fleet_capacity = n_loader * loader_capacity
    effective_daily_capacity = min(excavator_fleet_capacity, loader_fleet_capacity)
    bottleneck = "Excavator" if excavator_fleet_capacity <= loader_fleet_capacity else "Loader"
    
    st.sidebar.success(
        f"**Daily Soil Capacity: {effective_daily_capacity:,} CY**  \n"
        f"Bottleneck: {bottleneck} fleet  \n"
        f"Excavator fleet: {excavator_fleet_capacity:,} CY/day  \n"
        f"Loader fleet: {loader_fleet_capacity:,} CY/day  \n"
        f"_(Shared pool for load + unload)_"
    )
    
    # ============================================================
    # Material Costs
    # ============================================================
    st.sidebar.markdown("---")
    st.sidebar.header("💰 Material Costs")
    
    with st.sidebar.expander("💧 Water", expanded=False):
        water_bbl_per_cy = st.number_input("Water usage (BBL/CY total per treat cycle)", min_value=0.0, max_value=10.0, value=1.75, step=0.1, key='water_bbl_per_cy',
                                            help="Total water per CY across the full treat phase. Typical: 0.5 – 5 BBL/CY")
        water_cost_per_bbl = st.number_input("Water cost ($/BBL)", min_value=0.0, max_value=20.0, value=0.50, step=0.25, format="%.2f", key='water_cost_per_bbl',
                                              help="Typical: $0.25 – $5.00/BBL")
        st.caption(f"Effective: ${water_bbl_per_cy * water_cost_per_bbl:.2f}/CY • Distributed across Treat days")
        
        st.markdown("**Water Trucking**")
        water_truck_capacity = st.number_input("Water truck capacity (BBL/truck)", min_value=10, max_value=500, value=120, step=10, key='water_truck_capacity')
        water_truck_trip_hr = st.number_input("Water truck round-trip time (hr)", min_value=0.1, max_value=12.0, value=1.5, step=0.1, key='water_truck_trip_hr')
        water_truck_hourly = st.number_input("Water truck $/hr", min_value=0.0, value=95.0, step=5.0, key='water_truck_hourly')
        water_trucking_per_bbl = (water_truck_trip_hr * water_truck_hourly) / water_truck_capacity if water_truck_capacity > 0 else 0
        st.caption(f"Effective trucking: ${water_trucking_per_bbl:.3f}/BBL = ${water_trucking_per_bbl * water_bbl_per_cy:.2f}/CY")
    
    with st.sidebar.expander("🛢️ Leachate Disposal", expanded=False):
        leachate_pct_of_water = st.slider("Leachate collected (% of water used)", min_value=0, max_value=100, value=80, step=5, key='leachate_pct_of_water',
                                            help="Typical: 75% – 100% of water becomes leachate")
        leachate_lag_days = st.number_input("Leachate timing lag (days after water)", min_value=0, max_value=10, value=3, step=1, key='leachate_lag_days',
                                              help="Days between water addition and leachate emergence. Default 3 matches percolation through treated soil.")
        leachate_cost_per_bbl = st.number_input("Leachate disposal ($/BBL)", min_value=0.0, max_value=20.0, value=0.25, step=0.25, format="%.2f", key='leachate_cost_per_bbl',
                                                  help="Typical: $0.25 – $5.00/BBL")
        leachate_bbl_per_cy = water_bbl_per_cy * (leachate_pct_of_water / 100.0)
        st.caption(f"Leachate: {leachate_bbl_per_cy:.2f} BBL/CY • Disposal: ${leachate_bbl_per_cy * leachate_cost_per_bbl:.2f}/CY • Emerges {leachate_lag_days}d after water added")
        
        st.markdown("**Leachate Trucking**")
        leachate_truck_capacity = st.number_input("Leachate truck capacity (BBL/truck)", min_value=10, max_value=500, value=120, step=10, key='leachate_truck_capacity')
        leachate_truck_trip_hr = st.number_input("Leachate truck round-trip time (hr)", min_value=0.1, max_value=12.0, value=1.0, step=0.1, key='leachate_truck_trip_hr')
        leachate_truck_hourly = st.number_input("Leachate truck $/hr", min_value=0.0, value=95.0, step=5.0, key='leachate_truck_hourly')
        leachate_trucking_per_bbl = (leachate_truck_trip_hr * leachate_truck_hourly) / leachate_truck_capacity if leachate_truck_capacity > 0 else 0
        st.caption(f"Effective trucking: ${leachate_trucking_per_bbl:.3f}/BBL = ${leachate_trucking_per_bbl * leachate_bbl_per_cy:.2f}/CY")
    
    with st.sidebar.expander("⚗️ Amendments", expanded=False):
        amendment_cost_per_cy = st.number_input("Amendment cost ($/CY)", min_value=0.0, value=12.0, step=0.5, key='amendment_cost_per_cy',
                                                  help="Lump cost per CY; mixture breakdown coming later")
        st.caption("Charged on first Treat day of each flip")
    
    # Workday hours (used for trucks-required display only)
    workday_hours = 8.0  # informational; doesn't affect cost
    
    # Build cost params dict
    cost_params = {
        # Equipment fleet
        'n_excavator': n_excavator,
        'excavator_capacity': excavator_capacity,
        'excavator_daily': excavator_daily,
        'n_loader': n_loader,
        'loader_capacity': loader_capacity,
        'loader_daily': loader_daily,
        'n_bulldozer': n_bulldozer,
        'bulldozer_daily': bulldozer_daily,
        'n_skidsteer': n_skidsteer,
        'skidsteer_daily': skidsteer_daily,
        'effective_daily_capacity': effective_daily_capacity,
        'bottleneck': bottleneck,
        # Water
        'water_bbl_per_cy': water_bbl_per_cy,
        'water_cost_per_bbl': water_cost_per_bbl,
        'water_truck_capacity': water_truck_capacity,
        'water_truck_trip_hr': water_truck_trip_hr,
        'water_truck_hourly': water_truck_hourly,
        # Leachate
        'leachate_pct_of_water': leachate_pct_of_water,
        'leachate_bbl_per_cy': leachate_bbl_per_cy,
        'leachate_cost_per_bbl': leachate_cost_per_bbl,
        'leachate_lag_days': leachate_lag_days,
        'leachate_truck_capacity': leachate_truck_capacity,
        'leachate_truck_trip_hr': leachate_truck_trip_hr,
        'leachate_truck_hourly': leachate_truck_hourly,
        # Amendments
        'amendment_cost_per_cy': amendment_cost_per_cy,
        # Misc
        'workday_hours': workday_hours,
    }
    
    # Run button
    if st.sidebar.button("▶️ Run Simulation", type="primary", use_container_width=True):
        
        # Prepare parameters - daily soil capacity is the equipment bottleneck (shared load+unload pool)
        params = {
            'TotalSoil_CY': total_soil,
            'CellSize_CY': cell_size,
            'NumCells': num_cells,
            'DailyLoad_CY': effective_daily_capacity,
            'DailyUnload_CY': effective_daily_capacity,
            'StartDate': datetime.combine(start_date, datetime.min.time())
        }
        
        # Run simulation
        with st.spinner("Running simulation..."):
            all_activities = simulate_remediation(params, phases_df)
            schedule, idle_days_count = build_schedule(all_activities, params, phases_df)
            
            # Calculate costs
            daily_costs, cost_summary = calculate_costs(all_activities, schedule, params, cost_params, phases_df)
            
            # Store in session state
            st.session_state.activities = all_activities
            st.session_state.schedule = schedule
            st.session_state.params = params
            st.session_state.phases_df = phases_df
            st.session_state.idle_days_count = idle_days_count
            st.session_state.cost_params = cost_params
            st.session_state.daily_costs = daily_costs
            st.session_state.cost_summary = cost_summary
    
    # Display results if available
    if 'schedule' in st.session_state:
        schedule = st.session_state.schedule
        activities = st.session_state.activities
        params = st.session_state.params
        
        # Summary metrics
        st.header("📊 Summary Statistics")
        
        col1, col2, col3, col4, col5, col6 = st.columns(6)
        
        total_flips = math.ceil(params['TotalSoil_CY'] / params['CellSize_CY'])
        idle_days_count = st.session_state.get('idle_days_count', 0)
        cost_summary = st.session_state.get('cost_summary', {})
        
        with col1:
            st.metric("Total Soil", f"{params['TotalSoil_CY']:,} CY")
        
        with col2:
            st.metric("Total Flips", total_flips)
        
        with col3:
            completion_date = schedule[schedule['SoilOut'] > 0]['Date'].max() if len(schedule[schedule['SoilOut'] > 0]) > 0 else None
            if completion_date:
                st.metric("Completion Date", completion_date.strftime('%Y-%m-%d'))
            else:
                st.metric("Completion Date", "N/A")
        
        with col4:
            # Calculate actual project duration from start to completion
            if completion_date:
                total_days = (completion_date - params['StartDate']).days + 1
            else:
                total_days = 0
            st.metric("Total Days", total_days)
        
        with col5:
            final_cum = schedule['CumSoilOut'].max()
            st.metric("Soil Processed", f"{final_cum:,.0f} CY")
        
        with col6:
            st.metric("Idle Days", idle_days_count, help="Days when loader capacity was available but no cells were ready")
        
        # Second row: cost metrics
        if cost_summary:
            cc1, cc2, cc3, cc4, cc5, cc6 = st.columns(6)
            with cc1:
                st.metric("💰 Total Cost", f"${cost_summary.get('Total', 0):,.0f}")
            with cc2:
                st.metric("$ / CY", f"${cost_summary.get('CostPerCY', 0):,.2f}")
            with cc3:
                st.metric("Equipment", f"${cost_summary.get('Equipment', 0):,.0f}")
            with cc4:
                st.metric("Water (all-in)", f"${cost_summary.get('Water', 0):,.0f}",
                          help=f"Purchase ${cost_summary.get('WaterPurchase', 0):,.0f} + Trucking ${cost_summary.get('WaterTrucking', 0):,.0f}")
            with cc5:
                st.metric("Leachate (all-in)", f"${cost_summary.get('Leachate', 0):,.0f}",
                          help=f"Disposal ${cost_summary.get('LeachateDisposal', 0):,.0f} + Trucking ${cost_summary.get('LeachateTrucking', 0):,.0f}")
            with cc6:
                st.metric("Amendments", f"${cost_summary.get('Amendments', 0):,.0f}")
        
        # Tabs for different views
        tab1, tab2, tab3, tab4, tab5 = st.tabs(["📅 Schedule", "📋 Activities", "📈 Charts", "💰 Costs", "💾 Export"])
        
        with tab1:
            st.subheader("Daily Schedule")
            st.dataframe(schedule, use_container_width=True, height=600)
        
        with tab2:
            st.subheader("Cell Activities Log")
            activities_df = pd.DataFrame(activities)
            activities_df['Date'] = pd.to_datetime(activities_df['Date'])
            activities_df['Date'] = activities_df['Date'].dt.strftime('%Y-%m-%d')
            st.dataframe(activities_df, use_container_width=True, height=600)
        
        with tab3:
            st.subheader("Cumulative Soil In/Out")
            
            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=schedule['Date'],
                y=schedule['CumSoilIn'],
                name='Cumulative Soil In',
                line=dict(color='#8ED973', width=2)
            ))
            fig.add_trace(go.Scatter(
                x=schedule['Date'],
                y=schedule['CumSoilOut'],
                name='Cumulative Soil Out',
                line=dict(color='#00B0F0', width=2)
            ))
            fig.update_layout(
                xaxis_title="Date",
                yaxis_title="Soil (CY)",
                hovermode='x unified',
                height=400
            )
            st.plotly_chart(fig, use_container_width=True)
            
            # Daily soil movement
            st.subheader("Daily Soil Movement")
            fig2 = go.Figure()
            fig2.add_trace(go.Bar(
                x=schedule['Date'],
                y=schedule['SoilIn'],
                name='Soil In',
                marker_color='#8ED973'
            ))
            fig2.add_trace(go.Bar(
                x=schedule['Date'],
                y=schedule['SoilOut'],
                name='Soil Out',
                marker_color='#00B0F0'
            ))
            fig2.update_layout(
                xaxis_title="Date",
                yaxis_title="Soil (CY)",
                barmode='group',
                height=400
            )
            st.plotly_chart(fig2, use_container_width=True)
        
        with tab4:
            st.subheader("💰 Cost Breakdown")
            
            daily_costs = st.session_state.get('daily_costs')
            cost_summary = st.session_state.get('cost_summary', {})
            
            if daily_costs is None or daily_costs.empty:
                st.info("No cost data available. Re-run the simulation.")
            else:
                # Summary table
                st.markdown("### Project Totals")
                
                summary_rows = [
                    {'Category': 'Equipment',          'Cost': cost_summary.get('Equipment', 0)},
                    {'Category': 'Water — Purchase',   'Cost': cost_summary.get('WaterPurchase', 0)},
                    {'Category': 'Water — Trucking',   'Cost': cost_summary.get('WaterTrucking', 0)},
                    {'Category': 'Leachate — Disposal','Cost': cost_summary.get('LeachateDisposal', 0)},
                    {'Category': 'Leachate — Trucking','Cost': cost_summary.get('LeachateTrucking', 0)},
                    {'Category': 'Amendments',         'Cost': cost_summary.get('Amendments', 0)},
                ]
                total_cost = cost_summary.get('Total', 0)
                summary_df_cost = pd.DataFrame(summary_rows)
                summary_df_cost['% of Total'] = summary_df_cost['Cost'].apply(
                    lambda x: f"{(x / total_cost * 100):.1f}%" if total_cost > 0 else "0.0%"
                )
                summary_df_cost['Cost'] = summary_df_cost['Cost'].apply(lambda x: f"${x:,.2f}")
                
                col_a, col_b = st.columns([2, 1])
                with col_a:
                    st.dataframe(summary_df_cost, use_container_width=True, hide_index=True)
                with col_b:
                    st.metric("Total Project Cost", f"${total_cost:,.2f}")
                    st.metric("Cost per CY", f"${cost_summary.get('CostPerCY', 0):,.2f}")
                    st.metric("Project Days", f"{cost_summary.get('ProjectDays', 0)}")
                
                # Equipment breakdown by type
                st.markdown("### Equipment Cost by Type")
                equip_by_type = cost_summary.get('EquipmentByType', {})
                if equip_by_type:
                    equip_rows = [{'Equipment': p, 'Cost': v} for p, v in equip_by_type.items() if v > 0]
                    if equip_rows:
                        equip_df = pd.DataFrame(equip_rows)
                        equip_total = equip_df['Cost'].sum()
                        equip_df['% of Equipment'] = equip_df['Cost'].apply(
                            lambda x: f"{(x / equip_total * 100):.1f}%" if equip_total > 0 else "0.0%"
                        )
                        equip_df['Cost'] = equip_df['Cost'].apply(lambda x: f"${x:,.2f}")
                        st.dataframe(equip_df, use_container_width=True, hide_index=True)
                
                # Material volumes
                st.markdown("### Material Volumes & Peak Day")
                total_water_bbl = cost_summary.get('TotalWaterBBL', 0)
                total_leachate_bbl = cost_summary.get('TotalLeachateBBL', 0)
                peak_water = cost_summary.get('PeakWaterBBL_Day', 0)
                peak_leach = cost_summary.get('PeakLeachateBBL_Day', 0)
                peak_cells = cost_summary.get('PeakCellsInTreat', 0)
                
                mv1, mv2, mv3 = st.columns(3)
                with mv1:
                    st.metric("Total Water Used", f"{total_water_bbl:,.0f} BBL", help=f"{total_water_bbl * 42:,.0f} gallons")
                    st.caption(f"Peak day: {peak_water:,.0f} BBL")
                with mv2:
                    st.metric("Total Leachate Disposed", f"{total_leachate_bbl:,.0f} BBL", help=f"{total_leachate_bbl * 42:,.0f} gallons")
                    st.caption(f"Peak day: {peak_leach:,.0f} BBL")
                with mv3:
                    st.metric("Peak Cells in Treat", f"{peak_cells}", help="Maximum concurrent cells in Treat phase — used for sizing water trucking capacity")
                
                # Pie chart of cost breakdown
                st.markdown("### Cost Distribution")
                pie_data = pd.DataFrame([
                    {'Category': 'Equipment',           'Cost': cost_summary.get('Equipment', 0)},
                    {'Category': 'Water Purchase',      'Cost': cost_summary.get('WaterPurchase', 0)},
                    {'Category': 'Water Trucking',      'Cost': cost_summary.get('WaterTrucking', 0)},
                    {'Category': 'Leachate Disposal',   'Cost': cost_summary.get('LeachateDisposal', 0)},
                    {'Category': 'Leachate Trucking',   'Cost': cost_summary.get('LeachateTrucking', 0)},
                    {'Category': 'Amendments',          'Cost': cost_summary.get('Amendments', 0)},
                ])
                pie_data = pie_data[pie_data['Cost'] > 0]
                if not pie_data.empty:
                    fig_pie = px.pie(
                        pie_data,
                        values='Cost',
                        names='Category',
                        color_discrete_sequence=['#8ED973', '#83CCEB', '#5B9BD5', '#FFC000', '#ED7D31', '#F2CEEF']
                    )
                    fig_pie.update_traces(textposition='inside', textinfo='percent+label')
                    fig_pie.update_layout(height=400)
                    st.plotly_chart(fig_pie, use_container_width=True)
                
                # Daily cost chart
                st.markdown("### Daily Costs Over Time")
                fig_cost = go.Figure()
                fig_cost.add_trace(go.Bar(
                    x=daily_costs['Date'], y=daily_costs['ExcavatorCost'],
                    name='Excavator', marker_color='#8ED973'
                ))
                fig_cost.add_trace(go.Bar(
                    x=daily_costs['Date'], y=daily_costs['LoaderCost'],
                    name='Loader', marker_color='#5BBF4F'
                ))
                fig_cost.add_trace(go.Bar(
                    x=daily_costs['Date'], y=daily_costs['BulldozerCost'],
                    name='Bulldozer', marker_color='#2E8B57'
                ))
                fig_cost.add_trace(go.Bar(
                    x=daily_costs['Date'], y=daily_costs['SkidsteerCost'],
                    name='Skidsteer', marker_color='#90EE90'
                ))
                fig_cost.add_trace(go.Bar(
                    x=daily_costs['Date'], y=daily_costs['WaterPurchaseCost'],
                    name='Water Purchase', marker_color='#83CCEB'
                ))
                fig_cost.add_trace(go.Bar(
                    x=daily_costs['Date'], y=daily_costs['WaterTruckingCost'],
                    name='Water Trucking', marker_color='#5B9BD5'
                ))
                fig_cost.add_trace(go.Bar(
                    x=daily_costs['Date'], y=daily_costs['LeachateDisposalCost'],
                    name='Leachate Disposal', marker_color='#FFC000'
                ))
                fig_cost.add_trace(go.Bar(
                    x=daily_costs['Date'], y=daily_costs['LeachateTruckingCost'],
                    name='Leachate Trucking', marker_color='#ED7D31'
                ))
                fig_cost.add_trace(go.Bar(
                    x=daily_costs['Date'], y=daily_costs['AmendmentCost'],
                    name='Amendments', marker_color='#F2CEEF'
                ))
                fig_cost.update_layout(
                    xaxis_title="Date",
                    yaxis_title="Daily Cost ($)",
                    barmode='stack',
                    height=420,
                    hovermode='x unified',
                    legend=dict(orientation="h", yanchor="bottom", y=-0.4)
                )
                st.plotly_chart(fig_cost, use_container_width=True)
                
                # Cumulative cost chart
                st.markdown("### Cumulative Project Cost")
                fig_cum = go.Figure()
                fig_cum.add_trace(go.Scatter(
                    x=daily_costs['Date'],
                    y=daily_costs['CumTotalCost'],
                    name='Cumulative Cost',
                    line=dict(color='#00B0F0', width=3),
                    fill='tozeroy',
                    fillcolor='rgba(0, 176, 240, 0.15)'
                ))
                fig_cum.update_layout(
                    xaxis_title="Date",
                    yaxis_title="Cumulative Cost ($)",
                    height=400,
                    hovermode='x unified',
                    yaxis_tickformat='$,.0f'
                )
                st.plotly_chart(fig_cum, use_container_width=True)
                
                # Daily water & leachate volumes
                st.markdown("### Daily Water & Leachate Volumes")
                fig_vol = go.Figure()
                fig_vol.add_trace(go.Bar(
                    x=daily_costs['Date'], y=daily_costs['WaterBBL'],
                    name='Water Added (BBL)', marker_color='#5B9BD5'
                ))
                fig_vol.add_trace(go.Bar(
                    x=daily_costs['Date'], y=daily_costs['LeachateBBL'],
                    name='Leachate Collected (BBL)', marker_color='#ED7D31'
                ))
                fig_vol.update_layout(
                    xaxis_title="Date",
                    yaxis_title="Volume (BBL)",
                    barmode='group',
                    height=380,
                    hovermode='x unified'
                )
                st.plotly_chart(fig_vol, use_container_width=True)
                
                # Daily costs table
                st.markdown("### Daily Cost Detail")
                display_costs = daily_costs.copy()
                display_costs['Date'] = pd.to_datetime(display_costs['Date']).dt.strftime('%Y-%m-%d')
                money_cols = ['ExcavatorCost', 'LoaderCost', 'BulldozerCost', 'SkidsteerCost',
                              'EquipmentCost', 'WaterPurchaseCost', 'WaterTruckingCost', 'WaterCost',
                              'LeachateDisposalCost', 'LeachateTruckingCost', 'LeachateCost',
                              'AmendmentCost', 'TotalCost', 'CumTotalCost']
                for col in money_cols:
                    if col in display_costs.columns:
                        display_costs[col] = display_costs[col].apply(lambda x: f"${x:,.2f}")
                vol_cols = ['WaterBBL', 'LeachateBBL']
                for col in vol_cols:
                    if col in display_costs.columns:
                        display_costs[col] = display_costs[col].apply(lambda x: f"{x:,.1f}")
                st.dataframe(display_costs, use_container_width=True, height=400, hide_index=True)
        
        with tab5:
            st.subheader("Export Data")
            
            # Rebuild full unfiltered schedule to detect idle days for highlighting
            num_days = 1000
            dates = [params['StartDate'] + timedelta(days=i) for i in range(num_days)]
            months = [(i // 28) + 1 for i in range(num_days)]
            weeks = [(i // 7) + 1 for i in range(num_days)]
            num_cells_int = int(params['NumCells'])
            
            full_schedule_data = {
                'Month': months,
                'Week': weeks,
                'Day Count': range(1, num_days + 1),
                'Date': dates,
                'DayName': [d.strftime('%A') for d in dates],
                'SoilIn': [0] * num_days,
                'SoilOut': [0] * num_days
            }
            
            for i in range(1, num_cells_int + 1):
                full_schedule_data[f'Cell{i}Phase'] = [''] * num_days
            
            full_schedule = pd.DataFrame(full_schedule_data)
            
            # Merge activities
            for activity in activities:
                matching_rows = full_schedule['Date'] == activity['Date']
                cell_col = f"Cell{activity['CellNum']}Phase"
                existing_phase = full_schedule.loc[matching_rows, cell_col].values[0] if len(full_schedule.loc[matching_rows]) > 0 else ''
                if existing_phase and existing_phase != '':
                    new_phase = existing_phase + ' + ' + activity['Phase']
                else:
                    new_phase = activity['Phase']
                full_schedule.loc[matching_rows, cell_col] = new_phase
                full_schedule.loc[matching_rows, 'SoilIn'] += activity['SoilIn']
                full_schedule.loc[matching_rows, 'SoilOut'] += activity['SoilOut']
            
            full_schedule['CumSoilIn'] = full_schedule['SoilIn'].cumsum()
            full_schedule['CumSoilOut'] = full_schedule['SoilOut'].cumsum()
            
            # Detect idle days on full schedule
            idle_day_indices = []
            phases_df_local = st.session_state.phases_df
            for idx, row in full_schedule.iterrows():
                is_valid_load_day = is_valid_day(row['Date'], 'Load', phases_df_local)
                is_valid_unload_day = is_valid_day(row['Date'], 'Unload', phases_df_local)
                no_loading = row['SoilIn'] == 0
                no_unloading = row['SoilOut'] == 0
                still_soil_to_load = row['CumSoilIn'] < params['TotalSoil_CY']
                still_soil_to_unload = row['CumSoilOut'] < params['TotalSoil_CY']
                loading_has_started = row['CumSoilIn'] > 0
                could_have_worked = is_valid_load_day or is_valid_unload_day
                did_nothing = no_loading and no_unloading
                work_remains = (still_soil_to_load or (still_soil_to_unload and loading_has_started))
                if could_have_worked and did_nothing and work_remains:
                    idle_day_indices.append(idx)
            
            # Filter to match displayed schedule
            last_day_idx = None
            for idx, row in full_schedule.iterrows():
                if row['CumSoilOut'] >= params['TotalSoil_CY']:
                    last_day_idx = idx
                    break
            
            if last_day_idx is not None:
                schedule_for_export = full_schedule.iloc[:last_day_idx + 6].copy()
            else:
                schedule_for_export = full_schedule[
                    (full_schedule['SoilIn'] > 0) | 
                    (full_schedule['SoilOut'] > 0) | 
                    (full_schedule['CumSoilIn'] > 0) | 
                    (full_schedule['CumSoilOut'] > 0)
                ].copy()
            
            # Filter idle days to exported schedule
            filtered_idle_days = [idx for idx in idle_day_indices if idx in schedule_for_export.index]
            
            # Get cost data from session state
            daily_costs_export = st.session_state.get('daily_costs')
            cost_summary_export = st.session_state.get('cost_summary', {})
            cost_params_export = st.session_state.get('cost_params', {})
            
            # Create Excel file with formatting
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                schedule_for_export.to_excel(writer, sheet_name='Schedule', index=False)
                activities_df.to_excel(writer, sheet_name='Cell_Activities', index=False)
                
                # Costs sheet (daily breakdown)
                if daily_costs_export is not None and not daily_costs_export.empty:
                    costs_for_export = daily_costs_export.copy()
                    costs_for_export.to_excel(writer, sheet_name='Daily_Costs', index=False)
                
                # Summary sheet
                equip_by_type_exp = cost_summary_export.get('EquipmentByType', {})
                summary_metrics = [
                    'Total Soil (CY)',
                    'Cell Size (CY)',
                    'Number of Cells',
                    'Total Flips',
                    'Daily Soil Capacity (CY) [derived]',
                    'Bottleneck',
                    'Start Date',
                    'Completion Date',
                    'Total Days',
                    'Final CumSoilOut (CY)',
                    'Idle Capacity Days',
                    '',
                    '--- MATERIAL VOLUMES ---',
                    'Total Water Used (BBL)',
                    'Total Leachate Disposed (BBL)',
                    'Peak Water Day (BBL)',
                    'Peak Leachate Day (BBL)',
                    'Peak Cells in Treat',
                    '',
                    '--- COST SUMMARY ---',
                    'Equipment Cost ($)',
                    'Water Purchase Cost ($)',
                    'Water Trucking Cost ($)',
                    'Water — Total ($)',
                    'Leachate Disposal Cost ($)',
                    'Leachate Trucking Cost ($)',
                    'Leachate — Total ($)',
                    'Amendment Cost ($)',
                    'TOTAL PROJECT COST ($)',
                    'Cost per CY ($/CY)',
                    '',
                    '--- EQUIPMENT BY TYPE ---',
                    'Excavator ($)',
                    'Loader ($)',
                    'Bulldozer ($)',
                    'Skidsteer ($)',
                ]
                summary_values = [
                    params['TotalSoil_CY'],
                    params['CellSize_CY'],
                    params['NumCells'],
                    total_flips,
                    params['DailyLoad_CY'],
                    cost_params_export.get('bottleneck', 'N/A'),
                    params['StartDate'].strftime('%Y-%m-%d'),
                    completion_date.strftime('%Y-%m-%d') if completion_date else 'N/A',
                    total_days,
                    schedule_for_export['CumSoilOut'].max(),
                    len(filtered_idle_days),
                    '',
                    '',
                    round(cost_summary_export.get('TotalWaterBBL', 0), 2),
                    round(cost_summary_export.get('TotalLeachateBBL', 0), 2),
                    round(cost_summary_export.get('PeakWaterBBL_Day', 0), 2),
                    round(cost_summary_export.get('PeakLeachateBBL_Day', 0), 2),
                    cost_summary_export.get('PeakCellsInTreat', 0),
                    '',
                    '',
                    round(cost_summary_export.get('Equipment', 0), 2),
                    round(cost_summary_export.get('WaterPurchase', 0), 2),
                    round(cost_summary_export.get('WaterTrucking', 0), 2),
                    round(cost_summary_export.get('Water', 0), 2),
                    round(cost_summary_export.get('LeachateDisposal', 0), 2),
                    round(cost_summary_export.get('LeachateTrucking', 0), 2),
                    round(cost_summary_export.get('Leachate', 0), 2),
                    round(cost_summary_export.get('Amendments', 0), 2),
                    round(cost_summary_export.get('Total', 0), 2),
                    round(cost_summary_export.get('CostPerCY', 0), 2),
                    '',
                    '',
                    round(equip_by_type_exp.get('Excavator', 0), 2),
                    round(equip_by_type_exp.get('Loader', 0), 2),
                    round(equip_by_type_exp.get('Bulldozer', 0), 2),
                    round(equip_by_type_exp.get('Skidsteer', 0), 2),
                ]
                summary_data = {'Metric': summary_metrics, 'Value': summary_values}
                summary_df = pd.DataFrame(summary_data)
                summary_df.to_excel(writer, sheet_name='Summary', index=False)
                
                # Cost Inputs sheet (record the rates used)
                cost_inputs_data = {
                    'Parameter': [
                        '--- EQUIPMENT FLEET ---',
                        '# Excavators',
                        'Excavator capacity (CY/day each)',
                        'Excavator $/day each',
                        '# Loaders',
                        'Loader capacity (CY/day each)',
                        'Loader $/day each',
                        '# Bulldozers',
                        'Bulldozer $/day each',
                        '# Skidsteers',
                        'Skidsteer $/day each',
                        'Derived daily soil capacity (CY)',
                        'Bottleneck equipment',
                        '',
                        '--- WATER ---',
                        'Water usage (BBL/CY total)',
                        'Water cost ($/BBL)',
                        'Water truck capacity (BBL)',
                        'Water truck trip time (hr)',
                        'Water truck $/hr',
                        '',
                        '--- LEACHATE ---',
                        'Leachate collection (% of water)',
                        'Leachate volume (BBL/CY) [derived]',
                        'Leachate lag (days after water)',
                        'Leachate disposal ($/BBL)',
                        'Leachate truck capacity (BBL)',
                        'Leachate truck trip time (hr)',
                        'Leachate truck $/hr',
                        '',
                        '--- AMENDMENTS ---',
                        'Amendment cost ($/CY)',
                    ],
                    'Value': [
                        '',
                        cost_params_export.get('n_excavator', 0),
                        cost_params_export.get('excavator_capacity', 0),
                        cost_params_export.get('excavator_daily', 0),
                        cost_params_export.get('n_loader', 0),
                        cost_params_export.get('loader_capacity', 0),
                        cost_params_export.get('loader_daily', 0),
                        cost_params_export.get('n_bulldozer', 0),
                        cost_params_export.get('bulldozer_daily', 0),
                        cost_params_export.get('n_skidsteer', 0),
                        cost_params_export.get('skidsteer_daily', 0),
                        cost_params_export.get('effective_daily_capacity', 0),
                        cost_params_export.get('bottleneck', 'N/A'),
                        '',
                        '',
                        cost_params_export.get('water_bbl_per_cy', 0),
                        cost_params_export.get('water_cost_per_bbl', 0),
                        cost_params_export.get('water_truck_capacity', 0),
                        cost_params_export.get('water_truck_trip_hr', 0),
                        cost_params_export.get('water_truck_hourly', 0),
                        '',
                        '',
                        cost_params_export.get('leachate_pct_of_water', 0),
                        round(cost_params_export.get('leachate_bbl_per_cy', 0), 4),
                        cost_params_export.get('leachate_lag_days', 0),
                        cost_params_export.get('leachate_cost_per_bbl', 0),
                        cost_params_export.get('leachate_truck_capacity', 0),
                        cost_params_export.get('leachate_truck_trip_hr', 0),
                        cost_params_export.get('leachate_truck_hourly', 0),
                        '',
                        '',
                        cost_params_export.get('amendment_cost_per_cy', 0),
                    ]
                }
                cost_inputs_df = pd.DataFrame(cost_inputs_data)
                cost_inputs_df.to_excel(writer, sheet_name='Cost_Inputs', index=False)
                
                # Apply formatting to Schedule sheet
                workbook = writer.book
                worksheet = writer.sheets['Schedule']
                
                # Define colors
                phase_colors = {
                    'Load': '#8ED973',
                    'Rip': '#83CCEB',
                    'Treat': '#FFC000',
                    'Dry': '#F2CEEF',
                    'Unload': '#00B0F0'
                }
                
                sunday_fill = PatternFill(start_color='FFFFFF00', end_color='FFFFFF00', fill_type='solid')
                idle_fill = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')
                
                thin_border = Border(
                    left=Side(style='thin', color='000000'),
                    right=Side(style='thin', color='000000'),
                    top=Side(style='thin', color='000000'),
                    bottom=Side(style='thin', color='000000')
                )
                aptos_font = Font(name='Aptos Narrow', size=10)
                aptos_font_bold = Font(name='Aptos Narrow', size=10, bold=True)
                center_aligned = Alignment(horizontal='center', vertical='center')
                
                # Find columns
                cell_columns = []
                date_column_idx = None
                dayname_column_idx = None
                last_data_column_idx = len(schedule_for_export.columns)
                
                for col_idx, col_name in enumerate(schedule_for_export.columns, start=1):
                    if 'Phase' in col_name:
                        cell_columns.append((col_idx, col_name))
                    if col_name == 'Date':
                        date_column_idx = col_idx
                    if col_name == 'DayName':
                        dayname_column_idx = col_idx
                
                # Apply formatting
                for row_idx in range(1, len(schedule_for_export) + 2):
                    is_sunday = False
                    is_idle_day = False
                    
                    if row_idx > 1:
                        df_row_idx = schedule_for_export.index[row_idx - 2]
                        is_idle_day = df_row_idx in filtered_idle_days
                        
                        if dayname_column_idx:
                            day_name_cell = worksheet.cell(row=row_idx, column=dayname_column_idx)
                            is_sunday = str(day_name_cell.value) == 'Sunday'
                    
                    for col_idx in range(1, last_data_column_idx + 1):
                        cell = worksheet.cell(row=row_idx, column=col_idx)
                        is_phase_column = any(col_idx == phase_col_idx for phase_col_idx, _ in cell_columns)
                        
                        cell.border = thin_border
                        cell.font = aptos_font_bold if row_idx == 1 else aptos_font
                        
                        if col_idx == date_column_idx and row_idx > 1:
                            cell.number_format = 'M/D/YYYY'
                            if is_idle_day:
                                cell.fill = idle_fill
                        
                        if is_sunday and not is_phase_column and row_idx > 1 and col_idx != date_column_idx:
                            cell.fill = sunday_fill
                        
                        if row_idx > 1:
                            for phase_col_idx, phase_col_name in cell_columns:
                                if col_idx == phase_col_idx:
                                    cell.alignment = center_aligned
                                    phase_value = str(cell.value) if cell.value else ''
                                    
                                    for phase_name, color in phase_colors.items():
                                        if phase_name in phase_value:
                                            hex_color = color.lstrip('#')
                                            openpyxl_color = 'FF' + hex_color
                                            cell.fill = PatternFill(start_color=openpyxl_color,
                                                                  end_color=openpyxl_color,
                                                                  fill_type='solid')
                                            break
                
                # Auto-adjust column widths
                for column_cells in worksheet.columns:
                    length = max(len(str(cell.value) if cell.value else "") for cell in column_cells)
                    worksheet.column_dimensions[get_column_letter(column_cells[0].column)].width = min(length + 2, 50)
            
            output.seek(0)
            
            # Generate filename with new format
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            num_cells = int(params['NumCells'])
            cell_size = int(params['CellSize_CY'])
            capacity = int(params['DailyLoad_CY'])
            excel_filename = f"v{APP_VERSION}_{timestamp}_{num_cells}_{cell_size}_{capacity}.xlsx"
            
            st.download_button(
                label="📥 Download Excel Report",
                data=output,
                file_name=excel_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            # CSV downloads
            col1, col2, col3 = st.columns(3)
            
            with col1:
                csv_schedule = schedule.to_csv(index=False)
                st.download_button(
                    label="📄 Download Schedule (CSV)",
                    data=csv_schedule,
                    file_name=f"v{APP_VERSION}_schedule_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv"
                )
            
            with col2:
                csv_activities = activities_df.to_csv(index=False)
                st.download_button(
                    label="📄 Download Activities (CSV)",
                    data=csv_activities,
                    file_name=f"v{APP_VERSION}_activities_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv"
                )
            
            with col3:
                daily_costs_csv_data = st.session_state.get('daily_costs')
                if daily_costs_csv_data is not None and not daily_costs_csv_data.empty:
                    csv_costs = daily_costs_csv_data.to_csv(index=False)
                    st.download_button(
                        label="📄 Download Costs (CSV)",
                        data=csv_costs,
                        file_name=f"v{APP_VERSION}_daily_costs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                        mime="text/csv"
                    )
    
    else:
        st.info("👈 Configure parameters in the sidebar and click **Run Simulation** to start")


if __name__ == "__main__":
    main()
