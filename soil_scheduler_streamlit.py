"""
Soil Remediation Scheduler - Streamlit Interactive App
Interactive web app for multi-phase soil remediation with capacity pooling
"""

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
# Streamlit UI
# ============================================================================

def main():
    st.set_page_config(page_title="Soil Remediation Scheduler", layout="wide")
    
    # Display company logo
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.image("Clean_Futures_2.png", use_container_width=True)
    
    st.title("🏗️ Soil Remediation Scheduler")
    st.markdown("**Interactive multi-phase soil remediation simulator with capacity pooling**")
    
    # Sidebar for parameters
    st.sidebar.header("📋 Project Parameters")
    
    total_soil = st.sidebar.number_input("Total Soil (CY)", min_value=100, max_value=100000, value=36000, step=100)
    cell_size = st.sidebar.number_input("Cell Size (CY)", min_value=100, max_value=10000, value=4500, step=100)
    num_cells = st.sidebar.number_input("Number of Cells", min_value=1, max_value=20, value=4)
    daily_load = st.sidebar.number_input("Daily Load Capacity (CY)", min_value=100, max_value=10000, value=1500, step=100)
    daily_unload = st.sidebar.number_input("Daily Unload Capacity (CY)", min_value=100, max_value=10000, value=1500, step=100)
    start_date = st.sidebar.date_input("Start Date", value=datetime(2025, 12, 1))
    
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
    
    # Run button
    if st.sidebar.button("▶️ Run Simulation", type="primary", use_container_width=True):
        
        # Prepare parameters
        params = {
            'TotalSoil_CY': total_soil,
            'CellSize_CY': cell_size,
            'NumCells': num_cells,
            'DailyLoad_CY': daily_load,
            'DailyUnload_CY': daily_unload,
            'StartDate': datetime.combine(start_date, datetime.min.time())
        }
        
        # Run simulation
        with st.spinner("Running simulation..."):
            all_activities = simulate_remediation(params, phases_df)
            schedule, idle_days_count = build_schedule(all_activities, params, phases_df)
            
            # Store in session state
            st.session_state.activities = all_activities
            st.session_state.schedule = schedule
            st.session_state.params = params
            st.session_state.phases_df = phases_df
            st.session_state.idle_days_count = idle_days_count
    
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
        
        # Tabs for different views
        tab1, tab2, tab3, tab4 = st.tabs(["📅 Schedule", "📋 Activities", "📈 Charts", "💾 Export"])
        
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
            
            # Create Excel file with formatting
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                schedule_for_export.to_excel(writer, sheet_name='Schedule', index=False)
                activities_df.to_excel(writer, sheet_name='Cell_Activities', index=False)
                
                # Summary sheet
                summary_data = {
                    'Metric': [
                        'Total Soil (CY)',
                        'Cell Size (CY)',
                        'Number of Cells',
                        'Total Flips',
                        'Daily Load Capacity (CY)',
                        'Daily Unload Capacity (CY)',
                        'Start Date',
                        'Completion Date',
                        'Total Days',
                        'Final CumSoilOut (CY)',
                        'Idle Capacity Days'
                    ],
                    'Value': [
                        params['TotalSoil_CY'],
                        params['CellSize_CY'],
                        params['NumCells'],
                        total_flips,
                        params['DailyLoad_CY'],
                        params['DailyUnload_CY'],
                        params['StartDate'].strftime('%Y-%m-%d'),
                        completion_date.strftime('%Y-%m-%d') if completion_date else 'N/A',
                        total_days,
                        schedule_for_export['CumSoilOut'].max(),
                        len(filtered_idle_days)
                    ]
                }
                summary_df = pd.DataFrame(summary_data)
                summary_df.to_excel(writer, sheet_name='Summary', index=False)
                
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
            excel_filename = f"{timestamp}_{num_cells}_{cell_size}_{capacity}.xlsx"
            
            st.download_button(
                label="📥 Download Excel Report",
                data=output,
                file_name=excel_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            # CSV downloads
            col1, col2 = st.columns(2)
            
            with col1:
                csv_schedule = schedule.to_csv(index=False)
                st.download_button(
                    label="📄 Download Schedule (CSV)",
                    data=csv_schedule,
                    file_name=f"schedule_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv"
                )
            
            with col2:
                csv_activities = activities_df.to_csv(index=False)
                st.download_button(
                    label="📄 Download Activities (CSV)",
                    data=csv_activities,
                    file_name=f"activities_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv"
                )
    
    else:
        st.info("👈 Configure parameters in the sidebar and click **Run Simulation** to start")


if __name__ == "__main__":
    main()
