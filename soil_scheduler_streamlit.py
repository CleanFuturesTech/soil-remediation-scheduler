"""
Soil Remediation Scheduler - Streamlit Interactive App
Interactive web app for multi-phase soil remediation with capacity pooling
"""

APP_VERSION = "2.04"

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

def enrich_schedule_with_costs(schedule, daily_costs):
    """
    Merge daily cost and volume columns from daily_costs into the schedule.
    Adds columns for water/leachate volumes and per-day cost breakdown,
    placed in a logical order alongside the existing soil columns.
    
    Returns a new DataFrame; does not modify the input.
    """
    if daily_costs is None or daily_costs.empty:
        return schedule.copy()
    
    # Columns from daily_costs we want to surface in the schedule
    cost_cols_to_merge = [
        'WaterBBL', 'LeachateBBL',
        'WaterPurchaseCost', 'WaterTruckingCost',
        'LeachateDisposalCost', 'LeachateTruckingCost',
        'AmendmentCost',
    ]
    available_cols = [c for c in cost_cols_to_merge if c in daily_costs.columns]
    
    merge_df = daily_costs[['Date'] + available_cols].copy()
    
    enriched = schedule.merge(merge_df, on='Date', how='left')
    
    # Fill any NaNs (days outside daily_costs window) with 0
    for c in available_cols:
        if c in enriched.columns:
            enriched[c] = enriched[c].fillna(0)
    
    # Compute cumulative water in and leachate out for the schedule view
    if 'WaterBBL' in enriched.columns:
        enriched['CumWaterBBL'] = enriched['WaterBBL'].cumsum()
    if 'LeachateBBL' in enriched.columns:
        enriched['CumLeachateBBL'] = enriched['LeachateBBL'].cumsum()
    
    # Friendly aliases that match the user's mental model
    rename_map = {
        'WaterBBL':              'WaterIn_BBL',
        'LeachateBBL':           'LeachateOut_BBL',
        'CumWaterBBL':           'CumWaterIn_BBL',
        'CumLeachateBBL':        'CumLeachateOut_BBL',
        'WaterPurchaseCost':     'WaterCost_$',
        'WaterTruckingCost':     'WaterTrucking_$',
        'LeachateDisposalCost':  'LeachateDisposal_$',
        'LeachateTruckingCost':  'LeachateTrucking_$',
        'AmendmentCost':         'Amendments_$',
    }
    enriched = enriched.rename(columns=rename_map)
    
    # Reorder columns: identifiers first, soil volumes, water/leachate volumes,
    # daily costs, then cell phases
    id_cols  = [c for c in ['Month', 'Week', 'Day Count', 'Date', 'DayName'] if c in enriched.columns]
    soil_cols = [c for c in ['SoilIn', 'CumSoilIn', 'SoilOut', 'CumSoilOut'] if c in enriched.columns]
    vol_cols = [c for c in ['WaterIn_BBL', 'CumWaterIn_BBL', 'LeachateOut_BBL', 'CumLeachateOut_BBL'] if c in enriched.columns]
    cost_cols_ordered = [c for c in ['WaterCost_$', 'WaterTrucking_$', 'LeachateDisposal_$', 'LeachateTrucking_$', 'Amendments_$'] if c in enriched.columns]
    # Everything else (cell phase cols, etc.) at the end
    placed = set(id_cols + soil_cols + vol_cols + cost_cols_ordered)
    other_cols = [c for c in enriched.columns if c not in placed]
    
    enriched = enriched[id_cols + soil_cols + vol_cols + cost_cols_ordered + other_cols]
    
    return enriched


def calculate_costs(activities, schedule, params, cost_params, phases_df):
    """
    Calculate daily and total costs for the remediation project.
    
    Cost categories:
    - Equipment: daily fleet rental (excavator, loader, bulldozer, skidsteer)
                 charged for every day of project duration
    - Water purchase: BBL/day × $/BBL, where daily BBL is driven by cells in Treat today
    - Water trucking: truck-hours × $/hr, based on BBLs hauled
    - Leachate disposal: $/BBL × daily BBL, distributed evenly across each flip's Dry days
    - Leachate trucking: truck-hours × $/hr, based on BBLs hauled
    - Amendments: $/CY total per flip, distributed evenly across each flip's Treat days
    
    Per-day volume model:
      water_BBL_today = water_BBL_per_CY × cell_size × N_cells_in_Treat / treat_duration
      leachate_BBL_today = (water_BBL × leachate_pct × treat_duration / dry_duration) × N_cells_in_Dry
      amendment_$_today = (amendment_$_per_CY × cell_size / treat_duration) × N_cells_in_Treat
    
    Returns:
        daily_costs_df: per-day cost breakdown
        summary: dict of total costs by category
    """
    # ---------- Per-flip soil volumes ----------
    flip_cy = {}            # flip_num -> total CY loaded
    
    for act in activities:
        fn = act['FlipNum']
        if 'Load' in act['Phase']:
            flip_cy[fn] = flip_cy.get(fn, 0) + act['SoilIn']
    
    # ---------- Treat and Dry duration (days) ----------
    try:
        treat_duration_days = int(phases_df[phases_df['Phase'] == 'Treat']['Duration_Days'].iloc[0])
    except Exception:
        treat_duration_days = 3
    if treat_duration_days <= 0:
        treat_duration_days = 1
    
    try:
        dry_duration_days = int(phases_df[phases_df['Phase'] == 'Dry']['Duration_Days'].iloc[0])
    except Exception:
        dry_duration_days = 5
    if dry_duration_days <= 0:
        dry_duration_days = 1
    
    # ---------- Cell-size ----------
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
        'CellsInDry': 0,
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
    
    # ---------- Rate constants ----------
    water_bbl_per_cy   = cost_params.get('water_bbl_per_cy', 0.0)
    water_cost_per_bbl = cost_params.get('water_cost_per_bbl', 0.0)
    water_truck_cap    = cost_params.get('water_truck_capacity', 120)
    water_trip_hr      = cost_params.get('water_truck_trip_hr', 1.5)
    water_truck_rate   = cost_params.get('water_truck_hourly', 95.0)
    water_truck_cost_per_bbl = (water_trip_hr * water_truck_rate / water_truck_cap) if water_truck_cap > 0 else 0.0
    
    leachate_pct          = cost_params.get('leachate_pct_of_water', 80) / 100.0
    leachate_cost_per_bbl = cost_params.get('leachate_cost_per_bbl', 0.0)
    leachate_truck_cap    = cost_params.get('leachate_truck_capacity', 120)
    leachate_trip_hr      = cost_params.get('leachate_truck_trip_hr', 1.0)
    leachate_truck_rate   = cost_params.get('leachate_truck_hourly', 95.0)
    leachate_truck_cost_per_bbl = (leachate_trip_hr * leachate_truck_rate / leachate_truck_cap) if leachate_truck_cap > 0 else 0.0
    if cost_params.get('onsite_evap_pond', False):
        leachate_truck_cost_per_bbl = 0.0  # Pumped to onsite pond — no trucking
    
    amendment_cost_per_cy = cost_params.get('amendment_cost_per_cy', 0.0)
    
    # ---------- Per-cell-day rates ----------
    # Water: total water per CY × cell_size, spread across treat_duration days
    water_bbl_per_treat_cell_day = (water_bbl_per_cy * cell_size_cy / treat_duration_days) if treat_duration_days > 0 else 0.0
    # Leachate: total water per cell × leachate_pct, spread across dry_duration days
    leachate_bbl_per_dry_cell_day = (water_bbl_per_cy * cell_size_cy * leachate_pct / dry_duration_days) if dry_duration_days > 0 else 0.0
    # Amendments: total amendment cost per cell, spread across treat_duration days
    amendment_per_treat_cell_day = (amendment_cost_per_cy * cell_size_cy / treat_duration_days) if treat_duration_days > 0 else 0.0
    
    # ---------- Walk the schedule once, counting cells in Treat and Dry each day ----------
    num_cells = int(params['NumCells'])
    cell_cols = [f'Cell{i}Phase' for i in range(1, num_cells + 1)]
    
    for idx, row in schedule.iterrows():
        date = row['Date']
        if date not in daily_costs.index:
            continue
        
        n_treat = 0
        n_dry   = 0
        for col in cell_cols:
            phase_str = str(row[col]) if pd.notna(row[col]) else ''
            if 'Treat' in phase_str:
                n_treat += 1
            if 'Dry' in phase_str:
                n_dry += 1
        
        # Water (driven by cells in Treat)
        water_bbl = n_treat * water_bbl_per_treat_cell_day
        daily_costs.at[date, 'CellsInTreat'] = n_treat
        daily_costs.at[date, 'WaterBBL'] = water_bbl
        daily_costs.at[date, 'WaterPurchaseCost'] = water_bbl * water_cost_per_bbl
        daily_costs.at[date, 'WaterTruckingCost'] = water_bbl * water_truck_cost_per_bbl
        daily_costs.at[date, 'WaterCost'] = daily_costs.at[date, 'WaterPurchaseCost'] + daily_costs.at[date, 'WaterTruckingCost']
        
        # Leachate (driven by cells in Dry — spread evenly across each flip's dry days)
        leachate_bbl = n_dry * leachate_bbl_per_dry_cell_day
        daily_costs.at[date, 'CellsInDry'] = n_dry
        daily_costs.at[date, 'LeachateBBL'] = leachate_bbl
        daily_costs.at[date, 'LeachateDisposalCost'] = leachate_bbl * leachate_cost_per_bbl
        daily_costs.at[date, 'LeachateTruckingCost'] = leachate_bbl * leachate_truck_cost_per_bbl
        daily_costs.at[date, 'LeachateCost'] = daily_costs.at[date, 'LeachateDisposalCost'] + daily_costs.at[date, 'LeachateTruckingCost']
        
        # Amendments (driven by cells in Treat — spread evenly across each flip's treat days)
        daily_costs.at[date, 'AmendmentCost'] = n_treat * amendment_per_treat_cell_day
    
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
# Formula-driven Excel export
# ============================================================================

def build_formula_excel(schedule_for_export, activities_df, params, cost_params,
                       cost_summary, completion_date, total_days,
                       filtered_idle_days, total_flips, phases_df):
    """
    Build the Excel workbook with formula-driven calculations.
    
    Sheets produced:
      - Inputs: editable rate constants (source of truth)
      - Schedule: daily schedule; cost/volume columns are formulas referencing Inputs
      - Cell_Activities: raw activity log (no formulas)
      - Summary: aggregated totals using SUM() formulas referencing Schedule
    
    Returns:
      BytesIO of the Excel file
    """
    # ---------- Phase durations from phases_df ----------
    def _phase_dur(name, default):
        try:
            v = int(phases_df[phases_df['Phase'] == name]['Duration_Days'].iloc[0])
            return v if v > 0 else default
        except Exception:
            return default
    treat_dur = _phase_dur('Treat', 3)
    dry_dur   = _phase_dur('Dry', 5)
    
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        workbook = writer.book
        
        # ============================================================
        # 1. INPUTS sheet — editable rate constants
        # ============================================================
        inputs_rows = []   # list of (label, value_or_formula, note)
        INPUTS = {}        # logical name -> 'Inputs!$B$N'
        
        def add_input(name, label, value, note=''):
            inputs_rows.append((label, value, note))
            INPUTS[name] = f'Inputs!$B${len(inputs_rows) + 1}'
        
        def add_section(title):
            inputs_rows.append((title, '', ''))
        
        add_section('PROJECT')
        add_input('total_soil',  'Total Soil (CY)',         int(params['TotalSoil_CY']))
        add_input('cell_size',   'Cell Size (CY)',          int(params['CellSize_CY']))
        add_input('num_cells',   'Number of Cells',         int(params['NumCells']))
        add_input('treat_days',  'Treat Duration (days)',   treat_dur)
        add_input('dry_days',    'Dry Duration (days)',     dry_dur)
        
        add_section('')
        add_section('WATER')
        add_input('water_bbl_per_cy',  'Water usage (BBL/CY)',           cost_params.get('water_bbl_per_cy', 0))
        add_input('water_cost',        'Water cost ($/BBL)',             cost_params.get('water_cost_per_bbl', 0), 'Set to 0 if onsite well')
        add_input('water_truck_cap',   'Water truck capacity (BBL)',     cost_params.get('water_truck_capacity', 0))
        add_input('water_trip_hr',     'Water truck trip time (hr)',     cost_params.get('water_truck_trip_hr', 0))
        add_input('water_truck_hourly','Water truck $/hr',               cost_params.get('water_truck_hourly', 0))
        add_input('water_truck_per_bbl','Water trucking ($/BBL) [derived]',
                  f"={INPUTS['water_trip_hr']}*{INPUTS['water_truck_hourly']}/{INPUTS['water_truck_cap']}")
        add_input('water_per_cell_day','Water BBL per cell-day Treat [derived]',
                  f"={INPUTS['water_bbl_per_cy']}*{INPUTS['cell_size']}/{INPUTS['treat_days']}")
        
        add_section('')
        add_section('LEACHATE')
        add_input('leach_pct',         'Leachate % of water',            cost_params.get('leachate_pct_of_water', 0) / 100.0)
        add_input('leach_cost',        'Leachate disposal ($/BBL)',      cost_params.get('leachate_cost_per_bbl', 0), 'Set to 0 if onsite evap pond')
        add_input('leach_truck_cap',   'Leachate truck capacity (BBL)',  cost_params.get('leachate_truck_capacity', 0))
        add_input('leach_trip_hr',     'Leachate truck trip time (hr)',  cost_params.get('leachate_truck_trip_hr', 0))
        add_input('leach_truck_hourly','Leachate truck $/hr',            cost_params.get('leachate_truck_hourly', 0))
        # If onsite evap pond, force leachate trucking to 0
        if cost_params.get('onsite_evap_pond', False):
            add_input('leach_truck_per_bbl','Leachate trucking ($/BBL) [derived]', 0, 'Zeroed: onsite evap pond')
        else:
            add_input('leach_truck_per_bbl','Leachate trucking ($/BBL) [derived]',
                      f"={INPUTS['leach_trip_hr']}*{INPUTS['leach_truck_hourly']}/{INPUTS['leach_truck_cap']}")
        add_input('leach_per_cell_day','Leachate BBL per cell-day Dry [derived]',
                  f"={INPUTS['water_bbl_per_cy']}*{INPUTS['cell_size']}*{INPUTS['leach_pct']}/{INPUTS['dry_days']}")
        
        add_section('')
        add_section('AMENDMENTS')
        add_input('amend_cost',        'Amendment cost ($/CY)',          cost_params.get('amendment_cost_per_cy', 0))
        add_input('amend_per_cell_day','Amendment $ per cell-day Treat [derived]',
                  f"={INPUTS['amend_cost']}*{INPUTS['cell_size']}/{INPUTS['treat_days']}")
        
        add_section('')
        add_section('EQUIPMENT')
        add_input('n_excavator', '# Excavators',       cost_params.get('n_excavator', 0))
        add_input('exc_daily',   'Excavator $/day',    cost_params.get('excavator_daily', 0))
        add_input('n_loader',    '# Loaders',          cost_params.get('n_loader', 0))
        add_input('load_daily',  'Loader $/day',       cost_params.get('loader_daily', 0))
        add_input('n_bulldozer', '# Bulldozers',       cost_params.get('n_bulldozer', 0))
        add_input('bull_daily',  'Bulldozer $/day',    cost_params.get('bulldozer_daily', 0))
        add_input('n_skidsteer', '# Skidsteers',       cost_params.get('n_skidsteer', 0))
        add_input('skid_daily',  'Skidsteer $/day',    cost_params.get('skidsteer_daily', 0))
        add_input('daily_equip_total', 'Daily fleet cost ($) [derived]',
                  f"={INPUTS['n_excavator']}*{INPUTS['exc_daily']}+{INPUTS['n_loader']}*{INPUTS['load_daily']}+{INPUTS['n_bulldozer']}*{INPUTS['bull_daily']}+{INPUTS['n_skidsteer']}*{INPUTS['skid_daily']}")
        
        # Create the Inputs sheet at position 0 (first tab)
        inputs_ws = workbook.create_sheet('Inputs', 0)
        # Remove any default empty sheet that openpyxl created
        for sheet_name in list(workbook.sheetnames):
            if sheet_name == 'Sheet':
                del workbook[sheet_name]
        
        inputs_ws['A1'] = 'Parameter'
        inputs_ws['B1'] = 'Value'
        inputs_ws['C1'] = 'Notes'
        inputs_ws['A1'].font = Font(name='Aptos Narrow', size=10, bold=True)
        inputs_ws['B1'].font = Font(name='Aptos Narrow', size=10, bold=True)
        inputs_ws['C1'].font = Font(name='Aptos Narrow', size=10, bold=True)
        
        section_fill = PatternFill(start_color='FFE7E6E6', end_color='FFE7E6E6', fill_type='solid')
        for i, (label, val, note) in enumerate(inputs_rows, start=2):
            a_cell = inputs_ws.cell(row=i, column=1, value=label)
            b_cell = inputs_ws.cell(row=i, column=2, value=val)
            c_cell = inputs_ws.cell(row=i, column=3, value=note)
            a_cell.font = Font(name='Aptos Narrow', size=10)
            b_cell.font = Font(name='Aptos Narrow', size=10)
            c_cell.font = Font(name='Aptos Narrow', size=10, italic=True)
            if label in ('PROJECT', 'WATER', 'LEACHATE', 'AMENDMENTS', 'EQUIPMENT') or label == '':
                a_cell.fill = section_fill
                b_cell.fill = section_fill
                c_cell.fill = section_fill
                a_cell.font = Font(name='Aptos Narrow', size=10, bold=True)
        
        inputs_ws.column_dimensions['A'].width = 42
        inputs_ws.column_dimensions['B'].width = 14
        inputs_ws.column_dimensions['C'].width = 32
        
        # ============================================================
        # 2. SCHEDULE sheet — data first, then overwrite with formulas
        # ============================================================
        schedule_for_export.to_excel(writer, sheet_name='Schedule', index=False)
        sched_ws = writer.sheets['Schedule']
        
        # Column letter map
        sched_cols = {col: get_column_letter(i + 1)
                      for i, col in enumerate(schedule_for_export.columns)}
        
        cell_phase_columns = [c for c in schedule_for_export.columns if 'Phase' in c]
        if cell_phase_columns:
            first_phase_col = sched_cols[cell_phase_columns[0]]
            last_phase_col  = sched_cols[cell_phase_columns[-1]]
        else:
            first_phase_col = last_phase_col = 'A'
        
        n_rows = len(schedule_for_export)
        for row in range(2, n_rows + 2):
            phase_range = f"${first_phase_col}{row}:${last_phase_col}{row}"
            
            def set_formula(col_name, formula):
                if col_name in sched_cols:
                    sched_ws[f"{sched_cols[col_name]}{row}"] = formula
            
            # Volume formulas
            set_formula('WaterIn_BBL',
                        f'=COUNTIF({phase_range},"*Treat*")*{INPUTS["water_per_cell_day"]}')
            set_formula('LeachateOut_BBL',
                        f'=COUNTIF({phase_range},"*Dry*")*{INPUTS["leach_per_cell_day"]}')
            
            # Cumulative volumes
            if 'WaterIn_BBL' in sched_cols:
                wcol = sched_cols['WaterIn_BBL']
                set_formula('CumWaterIn_BBL', f'=SUM(${wcol}$2:{wcol}{row})')
            if 'LeachateOut_BBL' in sched_cols:
                lcol = sched_cols['LeachateOut_BBL']
                set_formula('CumLeachateOut_BBL', f'=SUM(${lcol}$2:{lcol}{row})')
            
            # Cumulative soil
            if 'SoilIn' in sched_cols:
                si = sched_cols['SoilIn']
                set_formula('CumSoilIn', f'=SUM(${si}$2:{si}{row})')
            if 'SoilOut' in sched_cols:
                so = sched_cols['SoilOut']
                set_formula('CumSoilOut', f'=SUM(${so}$2:{so}{row})')
            
            # Cost formulas (referenced from volume cells × rates from Inputs)
            if 'WaterIn_BBL' in sched_cols:
                wb_cell = f'{sched_cols["WaterIn_BBL"]}{row}'
                set_formula('WaterCost_$',     f'={wb_cell}*{INPUTS["water_cost"]}')
                set_formula('WaterTrucking_$', f'={wb_cell}*{INPUTS["water_truck_per_bbl"]}')
            if 'LeachateOut_BBL' in sched_cols:
                lb_cell = f'{sched_cols["LeachateOut_BBL"]}{row}'
                set_formula('LeachateDisposal_$', f'={lb_cell}*{INPUTS["leach_cost"]}')
                set_formula('LeachateTrucking_$', f'={lb_cell}*{INPUTS["leach_truck_per_bbl"]}')
            set_formula('Amendments_$',
                        f'=COUNTIF({phase_range},"*Treat*")*{INPUTS["amend_per_cell_day"]}')
        
        # ============================================================
        # 3. CELL ACTIVITIES sheet — raw, no formulas
        # ============================================================
        activities_df.to_excel(writer, sheet_name='Cell_Activities', index=False)
        
        # ============================================================
        # 4. SUMMARY sheet — formulas referencing Schedule + Inputs
        # ============================================================
        summary_ws = workbook.create_sheet('Summary')
        summary_ws['A1'] = 'Metric'
        summary_ws['B1'] = 'Value'
        summary_ws['A1'].font = Font(name='Aptos Narrow', size=10, bold=True)
        summary_ws['B1'].font = Font(name='Aptos Narrow', size=10, bold=True)
        
        def sched_col_range(col_name):
            if col_name not in sched_cols:
                return None
            c = sched_cols[col_name]
            return f"Schedule!${c}$2:${c}${n_rows + 1}"
        
        # Date column reference for project-day count
        date_col_letter = sched_cols.get('Date', 'D')
        project_days_formula = f"COUNTA(Schedule!${date_col_letter}$2:${date_col_letter}${n_rows + 1})"
        
        # Define summary rows. Use list so we can compute row offsets for inter-row references.
        sched_water_bbl_range  = sched_col_range('WaterIn_BBL')
        sched_leach_bbl_range  = sched_col_range('LeachateOut_BBL')
        sched_water_cost_range = sched_col_range('WaterCost_$')
        sched_water_truck_range= sched_col_range('WaterTrucking_$')
        sched_leach_disp_range = sched_col_range('LeachateDisposal_$')
        sched_leach_truck_range= sched_col_range('LeachateTrucking_$')
        sched_amend_range      = sched_col_range('Amendments_$')
        
        summary_rows = [
            ('Total Soil (CY)',              f"={INPUTS['total_soil']}"),
            ('Cell Size (CY)',               f"={INPUTS['cell_size']}"),
            ('Number of Cells',              f"={INPUTS['num_cells']}"),
            ('Total Flips',                  total_flips),
            ('Daily Soil Capacity (CY)',     int(params.get('DailyLoad_CY', 0))),
            ('Bottleneck',                   cost_params.get('bottleneck', 'N/A')),
            ('Start Date',                   params['StartDate'].strftime('%Y-%m-%d')),
            ('Completion Date',              completion_date.strftime('%Y-%m-%d') if completion_date else 'N/A'),
            ('Project Days',                 f"={project_days_formula}"),
            ('Idle Capacity Days',           len(filtered_idle_days)),
            ('', ''),
            ('MATERIAL VOLUMES',     ''),
            ('Total Water Used (BBL)',       f"=SUM({sched_water_bbl_range})"  if sched_water_bbl_range  else 0),
            ('Total Leachate Disposed (BBL)',f"=SUM({sched_leach_bbl_range})" if sched_leach_bbl_range else 0),
            ('', ''),
            ('COST SUMMARY',         ''),
            ('Equipment Cost ($)',           f"={INPUTS['daily_equip_total']}*{project_days_formula}"),
            ('Water Purchase Cost ($)',      f"=SUM({sched_water_cost_range})"  if sched_water_cost_range  else 0),
            ('Water Trucking Cost ($)',      f"=SUM({sched_water_truck_range})" if sched_water_truck_range else 0),
            ('Water — Total ($)',            None),   # filled below
            ('Leachate Disposal Cost ($)',   f"=SUM({sched_leach_disp_range})"  if sched_leach_disp_range  else 0),
            ('Leachate Trucking Cost ($)',   f"=SUM({sched_leach_truck_range})" if sched_leach_truck_range else 0),
            ('Leachate — Total ($)',         None),   # filled below
            ('Amendment Cost ($)',           f"=SUM({sched_amend_range})" if sched_amend_range else 0),
            ('TOTAL PROJECT COST ($)',       None),   # filled below
            ('Cost per CY ($/CY)',           None),   # filled below
            ('', ''),
            ('EQUIPMENT BY TYPE',    ''),
            ('Excavator ($)',                f"={INPUTS['n_excavator']}*{INPUTS['exc_daily']}*{project_days_formula}"),
            ('Loader ($)',                   f"={INPUTS['n_loader']}*{INPUTS['load_daily']}*{project_days_formula}"),
            ('Bulldozer ($)',                f"={INPUTS['n_bulldozer']}*{INPUTS['bull_daily']}*{project_days_formula}"),
            ('Skidsteer ($)',                f"={INPUTS['n_skidsteer']}*{INPUTS['skid_daily']}*{project_days_formula}"),
        ]
        
        # First pass: write everything except None placeholders
        for i, (label, val) in enumerate(summary_rows, start=2):
            a_cell = summary_ws.cell(row=i, column=1, value=label)
            a_cell.font = Font(name='Aptos Narrow', size=10, bold=label in ('MATERIAL VOLUMES', 'COST SUMMARY', 'EQUIPMENT BY TYPE', 'TOTAL PROJECT COST ($)'))
            if val is not None:
                b_cell = summary_ws.cell(row=i, column=2, value=val)
                b_cell.font = Font(name='Aptos Narrow', size=10)
                # Number format for currency
                if '($)' in label or '$/' in label:
                    b_cell.number_format = '"$"#,##0.00'
                elif '(CY)' in label or '(BBL)' in label or 'Days' in label or 'Cells' in label or 'Flips' in label:
                    b_cell.number_format = '#,##0'
        
        # Second pass: fix up rows that reference other summary cells
        label_to_row = {label: i + 2 for i, (label, _) in enumerate(summary_rows)}
        
        def set_summary(label, formula, fmt='"$"#,##0.00'):
            r = label_to_row[label]
            cell = summary_ws.cell(row=r, column=2, value=formula)
            cell.font = Font(name='Aptos Narrow', size=10, bold=(label == 'TOTAL PROJECT COST ($)'))
            cell.number_format = fmt
        
        set_summary('Water — Total ($)',
                    f"=B{label_to_row['Water Purchase Cost ($)']}+B{label_to_row['Water Trucking Cost ($)']}")
        set_summary('Leachate — Total ($)',
                    f"=B{label_to_row['Leachate Disposal Cost ($)']}+B{label_to_row['Leachate Trucking Cost ($)']}")
        set_summary('TOTAL PROJECT COST ($)',
                    f"=B{label_to_row['Equipment Cost ($)']}+B{label_to_row['Water — Total ($)']}+B{label_to_row['Leachate — Total ($)']}+B{label_to_row['Amendment Cost ($)']}")
        set_summary('Cost per CY ($/CY)',
                    f"=B{label_to_row['TOTAL PROJECT COST ($)']}/B{label_to_row['Total Soil (CY)']}",
                    fmt='"$"#,##0.00')
        
        summary_ws.column_dimensions['A'].width = 36
        summary_ws.column_dimensions['B'].width = 18
        
        # ============================================================
        # 5. SCHEDULE FORMATTING (colors, borders, number formats)
        # ============================================================
        phase_colors = {
            'Load': '#8ED973',
            'Rip': '#83CCEB',
            'Treat': '#FFC000',
            'Dry': '#F2CEEF',
            'Unload': '#00B0F0'
        }
        
        sunday_fill = PatternFill(start_color='FFFFFF00', end_color='FFFFFF00', fill_type='solid')
        idle_fill   = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')
        thin_border = Border(
            left=Side(style='thin', color='000000'),
            right=Side(style='thin', color='000000'),
            top=Side(style='thin', color='000000'),
            bottom=Side(style='thin', color='000000')
        )
        aptos_font      = Font(name='Aptos Narrow', size=10)
        aptos_font_bold = Font(name='Aptos Narrow', size=10, bold=True)
        center_aligned  = Alignment(horizontal='center', vertical='center')
        
        money_col_names  = {'WaterCost_$', 'WaterTrucking_$', 'LeachateDisposal_$', 'LeachateTrucking_$', 'Amendments_$'}
        volume_col_names = {'WaterIn_BBL', 'CumWaterIn_BBL', 'LeachateOut_BBL', 'CumLeachateOut_BBL'}
        soil_col_names   = {'SoilIn', 'SoilOut', 'CumSoilIn', 'CumSoilOut'}
        
        cell_columns = []
        date_column_idx = None
        dayname_column_idx = None
        money_col_idxs  = set()
        volume_col_idxs = set()
        soil_col_idxs   = set()
        
        for col_idx, col_name in enumerate(schedule_for_export.columns, start=1):
            if 'Phase' in col_name:
                cell_columns.append((col_idx, col_name))
            if col_name == 'Date':       date_column_idx    = col_idx
            if col_name == 'DayName':    dayname_column_idx = col_idx
            if col_name in money_col_names:  money_col_idxs.add(col_idx)
            if col_name in volume_col_names: volume_col_idxs.add(col_idx)
            if col_name in soil_col_names:   soil_col_idxs.add(col_idx)
        
        last_data_column_idx = len(schedule_for_export.columns)
        
        for row_idx in range(1, n_rows + 2):
            is_sunday  = False
            is_idle_day = False
            
            if row_idx > 1:
                df_row_idx = schedule_for_export.index[row_idx - 2]
                is_idle_day = df_row_idx in filtered_idle_days
                if dayname_column_idx:
                    day_name_cell = sched_ws.cell(row=row_idx, column=dayname_column_idx)
                    is_sunday = str(day_name_cell.value) == 'Sunday'
            
            for col_idx in range(1, last_data_column_idx + 1):
                cell = sched_ws.cell(row=row_idx, column=col_idx)
                is_phase_column = any(col_idx == p_idx for p_idx, _ in cell_columns)
                
                cell.border = thin_border
                cell.font = aptos_font_bold if row_idx == 1 else aptos_font
                
                if col_idx == date_column_idx and row_idx > 1:
                    cell.number_format = 'M/D/YYYY'
                    if is_idle_day:
                        cell.fill = idle_fill
                
                if row_idx > 1:
                    if col_idx in money_col_idxs:
                        cell.number_format = '"$"#,##0.00'
                    elif col_idx in volume_col_idxs:
                        cell.number_format = '#,##0.0'
                    elif col_idx in soil_col_idxs:
                        cell.number_format = '#,##0'
                
                if is_sunday and not is_phase_column and row_idx > 1 and col_idx != date_column_idx:
                    cell.fill = sunday_fill
                
                if row_idx > 1 and is_phase_column:
                    cell.alignment = center_aligned
                    phase_value = str(cell.value) if cell.value else ''
                    for phase_name, color in phase_colors.items():
                        if phase_name in phase_value:
                            hex_color = color.lstrip('#')
                            cell.fill = PatternFill(start_color='FF' + hex_color,
                                                    end_color='FF' + hex_color,
                                                    fill_type='solid')
                            break
        
        # Auto-adjust schedule column widths
        for column_cells in sched_ws.columns:
            length = max(len(str(cell.value) if cell.value else "") for cell in column_cells)
            sched_ws.column_dimensions[get_column_letter(column_cells[0].column)].width = min(length + 2, 50)
    
    output.seek(0)
    return output


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
    
    total_soil = st.sidebar.number_input("Total Soil (CY)", min_value=100, max_value=100000, value=16000, step=100)
    cell_size = st.sidebar.number_input("Cell Size (CY)", min_value=100, max_value=10000, value=1500, step=100)
    num_cells = st.sidebar.number_input("Number of Cells", min_value=1, max_value=20, value=4)
    start_date = st.sidebar.date_input("Start Date", value=datetime(2026, 6, 1))
    
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
        loader_capacity = st.number_input("CY/day per loader", min_value=50, max_value=5000, value=750, step=50, key='loader_capacity')
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
        water_bbl_per_cy = st.number_input("Water usage (BBL/CY total per treat cycle)", min_value=0.0, max_value=10.0, value=1.5, step=0.1, key='water_bbl_per_cy',
                                            help="Total water per CY across the full treat phase. Typical: 0.5 – 5 BBL/CY")
        
        onsite_water_well = st.toggle("Onsite freshwater well (water is free)", value=False, key='onsite_water_well',
                                       help="If enabled, water purchase cost is zero. Volume is still tracked for reporting.")
        
        if onsite_water_well:
            water_cost_per_bbl = 0.0
            st.caption("✅ Onsite well: $0.00/BBL water purchase. Volume still tracked.")
        else:
            water_cost_per_bbl = st.number_input("Water cost ($/BBL)", min_value=0.0, max_value=20.0, value=1.00, step=0.25, format="%.2f", key='water_cost_per_bbl',
                                                  help="Typical: $0.25 – $5.00/BBL")
            st.caption(f"Effective: ${water_bbl_per_cy * water_cost_per_bbl:.2f}/CY • Distributed across Treat days")
        
        st.markdown("**Water Trucking**")
        water_truck_capacity = st.number_input("Water truck capacity (BBL/truck)", min_value=10, max_value=500, value=120, step=10, key='water_truck_capacity')
        water_truck_trip_hr = st.number_input("Water truck round-trip time (hr)", min_value=0.1, max_value=12.0, value=1.5, step=0.1, key='water_truck_trip_hr')
        water_truck_hourly = st.number_input("Water truck $/hr", min_value=0.0, value=95.0, step=5.0, key='water_truck_hourly')
        water_trucking_per_bbl = (water_truck_trip_hr * water_truck_hourly) / water_truck_capacity if water_truck_capacity > 0 else 0
        if onsite_water_well:
            st.caption(f"⚠️ Trucking still applies if well is offsite or capacity insufficient. Effective: ${water_trucking_per_bbl:.3f}/BBL")
        else:
            st.caption(f"Effective trucking: ${water_trucking_per_bbl:.3f}/BBL = ${water_trucking_per_bbl * water_bbl_per_cy:.2f}/CY")
    
    with st.sidebar.expander("🛢️ Leachate Disposal", expanded=False):
        leachate_pct_of_water = st.slider("Leachate collected (% of water used)", min_value=0, max_value=100, value=80, step=5, key='leachate_pct_of_water',
                                            help="Typical: 75% – 100% of water becomes leachate")
        leachate_bbl_per_cy = water_bbl_per_cy * (leachate_pct_of_water / 100.0)
        
        onsite_evap_pond = st.toggle("Onsite evaporation pond (disposal is free)", value=False, key='onsite_evap_pond',
                                      help="If enabled, leachate disposal and trucking costs are zero. Volume is still tracked.")
        
        if onsite_evap_pond:
            leachate_cost_per_bbl = 0.0
            st.caption(f"✅ Onsite evap pond: $0.00/BBL disposal. Leachate volume: {leachate_bbl_per_cy:.2f} BBL/CY (still tracked).")
        else:
            leachate_cost_per_bbl = st.number_input("Leachate disposal ($/BBL)", min_value=0.0, max_value=20.0, value=0.25, step=0.25, format="%.2f", key='leachate_cost_per_bbl',
                                                      help="Typical: $0.25 – $5.00/BBL")
            st.caption(f"Leachate: {leachate_bbl_per_cy:.2f} BBL/CY • Disposal: ${leachate_bbl_per_cy * leachate_cost_per_bbl:.2f}/CY • Spread evenly across each flip's Dry days")
        
        st.markdown("**Leachate Trucking**")
        leachate_truck_capacity = st.number_input("Leachate truck capacity (BBL/truck)", min_value=10, max_value=500, value=120, step=10, key='leachate_truck_capacity')
        leachate_truck_trip_hr = st.number_input("Leachate truck round-trip time (hr)", min_value=0.1, max_value=12.0, value=1.0, step=0.1, key='leachate_truck_trip_hr')
        leachate_truck_hourly = st.number_input("Leachate truck $/hr", min_value=0.0, value=95.0, step=5.0, key='leachate_truck_hourly')
        leachate_trucking_per_bbl = (leachate_truck_trip_hr * leachate_truck_hourly) / leachate_truck_capacity if leachate_truck_capacity > 0 else 0
        if onsite_evap_pond:
            st.caption("✅ Onsite pond: no leachate trucking cost (pumped to pond on-site).")
        else:
            st.caption(f"Effective trucking: ${leachate_trucking_per_bbl:.3f}/BBL = ${leachate_trucking_per_bbl * leachate_bbl_per_cy:.2f}/CY")
    
    with st.sidebar.expander("⚗️ Amendments", expanded=False):
        amendment_cost_per_cy = st.number_input("Amendment cost ($/CY)", min_value=0.0, value=4.0, step=0.5, key='amendment_cost_per_cy',
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
        'onsite_water_well': onsite_water_well,
        'water_truck_capacity': water_truck_capacity,
        'water_truck_trip_hr': water_truck_trip_hr,
        'water_truck_hourly': water_truck_hourly,
        # Leachate
        'leachate_pct_of_water': leachate_pct_of_water,
        'leachate_bbl_per_cy': leachate_bbl_per_cy,
        'leachate_cost_per_bbl': leachate_cost_per_bbl,
        'onsite_evap_pond': onsite_evap_pond,
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
            
            # Enrich the schedule with cost and volume columns (for display + export)
            enriched_schedule = enrich_schedule_with_costs(schedule, daily_costs)
            
            # Store in session state
            st.session_state.activities = all_activities
            st.session_state.schedule = enriched_schedule
            st.session_state.schedule_raw = schedule   # original soil-only schedule (kept for any internal use)
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
            
            # Build column-level formatting for cost & volume columns
            schedule_col_config = {}
            money_cols_in_sched = ['WaterCost_$', 'WaterTrucking_$', 'LeachateDisposal_$', 'LeachateTrucking_$', 'Amendments_$']
            for c in money_cols_in_sched:
                if c in schedule.columns:
                    schedule_col_config[c] = st.column_config.NumberColumn(c, format="$%.2f")
            vol_cols_in_sched = ['WaterIn_BBL', 'CumWaterIn_BBL', 'LeachateOut_BBL', 'CumLeachateOut_BBL']
            for c in vol_cols_in_sched:
                if c in schedule.columns:
                    schedule_col_config[c] = st.column_config.NumberColumn(c, format="%.1f")
            soil_cols_in_sched = ['SoilIn', 'CumSoilIn', 'SoilOut', 'CumSoilOut']
            for c in soil_cols_in_sched:
                if c in schedule.columns:
                    schedule_col_config[c] = st.column_config.NumberColumn(c, format="%d")
            
            st.dataframe(
                schedule,
                use_container_width=True,
                height=600,
                column_config=schedule_col_config,
                hide_index=True,
            )
        
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
            
            # Enrich the exported schedule with the same cost & volume columns
            # used in the online display
            schedule_for_export = enrich_schedule_with_costs(schedule_for_export, daily_costs_export)
            
            # Create Excel file with formula-driven calculations
            output = build_formula_excel(
                schedule_for_export=schedule_for_export,
                activities_df=activities_df,
                params=params,
                cost_params=cost_params_export,
                cost_summary=cost_summary_export,
                completion_date=completion_date,
                total_days=total_days,
                filtered_idle_days=filtered_idle_days,
                total_flips=total_flips,
                phases_df=phases_df_local,
            )
            
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
