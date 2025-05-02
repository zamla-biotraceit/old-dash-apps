#!/usr/bin/env python
# coding: utf-8

# In[1]:


### MANUAL AND PYTHON BASES RWIN PEAK DETECTION WITH RESULTS
from dash import Dash, html, dcc, dash_table
from dash.dependencies import Input, Output, State
from dash.exceptions import PreventUpdate
import plotly.graph_objects as go
from openpyxl import load_workbook, Workbook
from openpyxl.drawing.image import Image
import plotly.io as pio
import kaleido
import pandas as pd
import numpy as np
import scipy
from scipy.signal import find_peaks
import traceback
import os
from datetime import datetime
import dash
import glob
import re
from scipy import stats
import plotly.express as px
import plotly.subplots as sp

VAS_CONFIG = {
    'show_vas': True,  # Always show VAS
    'value': True,     # Default VAS value
    'color': 'red',    # Default VAS color
    'opacity': 0.3     # Default VAS opacity
}

# Initialize the app
app = Dash(__name__, suppress_callback_exceptions=True, title="Peak Detection Analysis")

# Default word filters
DEFAULT_INCLUDE_WORDS = ['']
DEFAULT_EXCLUDE_WORDS = ['Pre']

# Default event filters
DEFAULT_INCLUDE_EVENTS = ['']
DEFAULT_EXCLUDE_EVENTS = ['Pre']

# Color scheme for graphs
COLORS = {
    'line': 'rgb(0, 140, 255)',
    'detected_base': 'rgb(255, 165, 0)',
    'detected_peak': 'rgb(255, 0, 0)',
    'mbase': 'rgb(0, 128, 0)',
    'mpeak': 'rgb(128, 0, 128)'
}

def detect_peaks_rolling_window(data, sheet_name, window_size=5):
    """
    Detect peaks and valleys using a rolling window comparison method with edge case handling.
    """
    peaks_data = []
    n = len(data)
    # window_size = 5
    
    if n < (2 * window_size + 1):
        print(f"Warning: Not enough points in data for sheet {sheet_name}.")
        return pd.DataFrame()
    
    try:
        for i in range(n):
            current_value = data.iloc[i]['Value']
            current_offset = data.iloc[i]['Offset']
            
            start_idx = max(0, i - window_size)
            end_idx = min(n, i + window_size + 1)
            
            if i > start_idx and i < end_idx - 1:
                prev_values = data.iloc[start_idx:i]['Value'].values
                next_values = data.iloc[i+1:end_idx]['Value'].values
                
                if len(prev_values) > 0 and len(next_values) > 0:
                    # Check for maxima (base)
                    if all(current_value > prev_values) and all(current_value > next_values):
                        current_base = {
                            'sheet_name': sheet_name,
                            'Type': 'Rolling Base',
                            'Offset': float(current_offset/1000),
                            'Value': float(current_value*10),
                            'Annotation': data.iloc[i].get('Annotation', ''),
                            'Event': data.iloc[i].get('Event', ''),
                            'mbase': None,
                            'mbase_time': None,
                            'mpeak': None,
                            'mpeak_time': None,
                            'pos_delta': None,
                            'neg_delta': None,
                            'selected': True,  # Default to selected
                            'VAS': None,
                            'Notes': f'Detected using {len(prev_values)}/{len(next_values)}-point window'
                        }
                        peaks_data.append(current_base)
                        
                    # Check for minima (peak)
                    if all(current_value < prev_values) and all(current_value < next_values):
                        current_peak = {
                            'sheet_name': sheet_name,
                            'Type': 'Rolling Peak',
                            'Offset': float(current_offset/1000),
                            'Value': float(current_value*10),
                            'Annotation': data.iloc[i].get('Annotation', ''),
                            'Event': data.iloc[i].get('Event', ''),
                            'mbase': None,
                            'mbase_time': None,
                            'mpeak': None,
                            'mpeak_time': None,
                            'pos_delta': None,
                            'neg_delta': None,
                            'selected': True,  # Default to selected
                            'VAS': None,
                            'Notes': f'Detected using {len(prev_values)}/{len(next_values)}-point window'
                        }
                        peaks_data.append(current_peak)
        
        df = pd.DataFrame(peaks_data)
        if not df.empty:
            numeric_columns = ['Offset', 'Value', 'pos_delta', 'neg_delta']
            for col in numeric_columns:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce')
            
            df = df.sort_values('Offset')
            
            # Calculate deltas after sorting
            df = calculate_deltas_with_selection(df)
        
        return df
    
    except Exception as e:
        print(f"Error processing sheet {sheet_name}: {str(e)}")
        traceback.print_exc()
        return pd.DataFrame()

# Calculate selected deltas
def calculate_deltas_with_selection(data):
    """Calculate deltas only using selected rows"""
    if not isinstance(data, pd.DataFrame) or data.empty:
        return data
        
    # Ensure data is sorted by Offset
    data = data.sort_values('Offset')
    
    # Reset all deltas to None
    data['pos_delta'] = None
    data['neg_delta'] = None
    
    last_base = None
    last_peak = None
    
    # Process each row in chronological order
    for idx in data.index:
        if not data.loc[idx, 'selected']:
            continue
            
        if data.loc[idx, 'Type'] == 'Rolling Base':
            # Find last selected peak before this base
            peak_mask = (data['Type'] == 'Rolling Peak') & \
                       (data['Offset'] < data.loc[idx, 'Offset']) & \
                       (data['selected'])
            if peak_mask.any():
                last_peak = data[peak_mask].iloc[-1]
                data.loc[idx, 'pos_delta'] = float(data.loc[idx, 'Value'] - last_peak['Value'])
            
        elif data.loc[idx, 'Type'] == 'Rolling Peak':
            # Find last selected base before this peak
            base_mask = (data['Type'] == 'Rolling Base') & \
                       (data['Offset'] < data.loc[idx, 'Offset']) & \
                       (data['selected'])
            if base_mask.any():
                last_base = data[base_mask].iloc[-1]
                data.loc[idx, 'neg_delta'] = float(data.loc[idx, 'Value'] - last_base['Value'])
    
    return data

def update_table_with_selection(new_table, current_table):
    """Preserve selection status when updating table data"""
    if not isinstance(new_table, pd.DataFrame):
        new_table = pd.DataFrame(new_table)
    if not isinstance(current_table, pd.DataFrame):
        current_table = pd.DataFrame(current_table)
        
    # Ensure we have the delta columns
    if 'pos_delta' not in new_table.columns:
        new_table['pos_delta'] = None
    if 'neg_delta' not in new_table.columns:
        new_table['neg_delta'] = None
        
    current_selected = {}
    for _, row in current_table.iterrows():
        key = (row['Offset'], row['Value'], row['Type'])
        current_selected[key] = row.get('selected', True)
        
    new_table['selected'] = True
    for idx, row in new_table.iterrows():
        key = (row['Offset'], row['Value'], row['Type'])
        if key in current_selected:
            new_table.at[idx, 'selected'] = current_selected[key]
    
    # Recalculate all deltas based on current selection state
    new_table = calculate_deltas_with_selection(new_table)
            
    return new_table
    
def analyze_multiple_graphs_statistics(sheet_data, selected_sheets, visible_range=None):
    """
    Calculate statistics for multiple graphs, both combined and individual.
    
    Parameters:
    -----------
    sheet_data : dict
        Dictionary containing data for each sheet
    selected_sheets : list
        List of selected sheet names to analyze
    visible_range : tuple, optional
        (min_x, max_x) range to calculate statistics for
        
    Returns:
    --------
    tuple
        (combined_stats, individual_stats)
    """
    try:
        all_values = []
        individual_stats = {}
        
        # Process each selected sheet
        for sheet_name in selected_sheets:
            if sheet_name not in sheet_data:
                continue
                
            data = sheet_data[sheet_name]['data']
            if data.empty:
                continue
            
            # Apply visible range filter if provided
            if visible_range:
                min_x, max_x = visible_range
                mask = (data['Offset']/1000 >= min_x) & (data['Offset']/1000 <= max_x)
                data = data[mask]
            
            values = data['Value'].values
            
            if len(values) > 0:
                # Calculate individual sheet statistics
                stats = {
                    'mean': np.mean(values),
                    'std': np.std(values),
                    'max': np.max(values),
                    'min': np.min(values),
                    'range': np.max(values) - np.min(values),
                    'median': np.median(values),
                    'q1': np.percentile(values, 25),
                    'q3': np.percentile(values, 75),
                    'iqr': np.percentile(values, 75) - np.percentile(values, 25)
                }
                individual_stats[sheet_name] = stats
                all_values.extend(values)
        
        # Calculate combined statistics if we have any values
        if all_values:
            combined_stats = {
                'mean': np.mean(all_values),
                'std': np.std(all_values),
                'max': np.max(all_values),
                'min': np.min(all_values),
                'range': np.max(all_values) - np.min(all_values),
                'median': np.median(all_values),
                'q1': np.percentile(all_values, 25),
                'q3': np.percentile(all_values, 75),
                'iqr': np.percentile(all_values, 75) - np.percentile(all_values, 25)
            }
        else:
            combined_stats = None
            
        return combined_stats, individual_stats
        
    except Exception as e:
        print(f"Error in analyze_multiple_graphs_statistics: {str(e)}")
        return None, {}

def create_multi_stats_table(combined_stats, individual_stats):
    """Create a formatted DataTable to display statistics for multiple graphs"""
    if not combined_stats and not individual_stats:
        return dash_table.DataTable(
            id='stats-table',
            columns=[],
            data=[]
        )
    
    columns = [
        {'name': 'Statistic', 'id': 'Statistic'},
        {'name': 'Combined', 'id': 'Combined'}
    ]
    
    for sheet_name in individual_stats.keys():
        columns.append({'name': sheet_name, 'id': sheet_name})
    
    stats_keys = ['mean', 'std', 'max', 'min', 'range', 'median', 'q1', 'q3', 'iqr']
    table_data = []
    
    for stat in stats_keys:
        row = {'Statistic': stat.upper()}
        
        if combined_stats:
            row['Combined'] = f"{combined_stats[stat]:.3f}"
        
        for sheet_name, stats in individual_stats.items():
            row[sheet_name] = f"{stats[stat]:.3f}"
        
        table_data.append(row)
    
    return dash_table.DataTable(
        id='stats-table',
        columns=columns,
        data=table_data,
        style_table={
            'overflowX': 'auto',
            'maxHeight': '400px',
            'overflowY': 'auto'
        },
        style_cell={
            'textAlign': 'left',
            'minWidth': '100px',
            'maxWidth': '180px',
            'whiteSpace': 'normal'
        },
        style_header={
            'backgroundColor': 'rgb(230, 230, 230)',
            'fontWeight': 'bold'
        },
        style_data_conditional=[
            {
                'if': {'column_id': 'Statistic'},
                'fontWeight': 'bold',
                'backgroundColor': 'rgb(248, 248, 248)'
            }
        ]
    )

def filter_by_words(data, include_words, exclude_words):
    """Filter data based on include and exclude word lists"""
    if not isinstance(data, pd.DataFrame) or data.empty:
        return data
        
    filtered_data = data.copy()
    
    if include_words:
        mask = filtered_data['Annotation'].str.contains('|'.join(include_words), case=False, na=False)
        filtered_data = filtered_data[mask]
        print(f"After including words, {len(filtered_data)} rows remain")
    
    if exclude_words:
        mask = ~filtered_data['Annotation'].str.contains('|'.join(exclude_words), case=False, na=False)
        filtered_data = filtered_data[mask]
        print(f"After excluding words, {len(filtered_data)} rows remain")
    
    return filtered_data

def filter_by_events(data, include_events, exclude_events):
    """Filter data based on include and exclude event lists"""
    if not isinstance(data, pd.DataFrame) or data.empty:
        return data
        
    filtered_data = data.copy()
    
    if include_events:
        mask = filtered_data['Event'].str.contains('|'.join(include_events), case=False, na=False)
        filtered_data = filtered_data[mask]
        print(f"After including events, {len(filtered_data)} rows remain")
    
    if exclude_events:
        mask = ~filtered_data['Event'].str.contains('|'.join(exclude_events), case=False, na=False)
        filtered_data = filtered_data[mask]
    
    return filtered_data

def find_available_y_position(x_pos, y_pos, used_positions, min_gap=1.0):
    """
    Find an available y position for an annotation that doesn't overlap with existing ones.
    
    Parameters:
    -----------
    x_pos : float
        X position of the new annotation
    y_pos : float
        Base Y position of the new annotation
    used_positions : list of tuples
        List of (x, y) positions of existing annotations
    min_gap : float
        Minimum vertical gap between annotations
        
    Returns:
    --------
    float
        Adjusted y position that avoids overlap
    """
    # Initialize the proposed y position
    proposed_y = y_pos
    
    # Define the x-range window where we check for conflicts
    x_window = 2.0  # Time window in seconds to check for nearby annotations
    
    while True:
        # Check for any nearby annotations
        conflicts = any(
            abs(x - x_pos) < x_window and abs(y - proposed_y) < min_gap
            for x, y in used_positions
        )
        
        if not conflicts:
            break
            
        # If there's a conflict, move up by the minimum gap
        proposed_y += min_gap
    
    return proposed_y

def read_excel_sheets(file_path):
    """Read worksheets from Excel file that contain offset, value, annotation, and event data"""
    sheet_data = {}
    with pd.ExcelFile(file_path) as xls:
        for sheet_name in xls.sheet_names:
            if any(phase in sheet_name for phase in ['P1', 'P3', 'P4']):
                try:
                    df = pd.read_excel(xls, sheet_name=sheet_name, usecols=[0, 2, 3, 4, 5, 7], 
                                     names=['Offset', 'Value', 'Annotations', 'Annotation', 'Event', 'VAS'])
                    
                    actions = df['Annotations'].dropna().unique().tolist()
                    
                    df = df.dropna(subset=['Offset', 'Value', 'Annotation', 'Event'])
                    
                    df['VAS'] = df['VAS'].fillna(0)
                    
                    sheet_data[sheet_name] = {'data': df, 'actions': actions}
                    # print(f"Loaded sheet: {sheet_name} with {len(df)} rows")
                except Exception as e:
                    print(f"Error reading sheet '{sheet_name}': {str(e)}")
    return sheet_data

def read_measurement_data(file_path):
    """Read manual analysis data from all worksheets"""
    # print(f"\nReading measurement data from {file_path}")
    data = []
    valid_ids = set()

    xlsx = pd.ExcelFile(file_path)
    for sheet_name in xlsx.sheet_names:
        try:
            df = pd.read_excel(xlsx, sheet_name=sheet_name, header=None)
            headers = df.iloc[0]
            # print(f"Headers found: {headers.tolist()}")

            event_row = df.iloc[31]  # 31 for 0-based indexing to get row 32
            #print(f"Event row values: {event_row.values}")
            
            for idx, row in df.iloc[1:30].iterrows():
                id_value = str(row[0])
                if pd.isna(id_value) or id_value == "Event":
                    continue
                    
                # print(f"Processing ID: {id_value}")
                
                for col in range(1, len(row), 4):
                    if pd.notna(row[col]):
                        try:
                            base_value = float(row[col])
                            peak_value = float(row[col + 1])
                            base_time = float(row[col + 2])
                            peak_time = float(row[col + 3])

                            event_value = str(event_row[col]) if event_row is not None and pd.notna(event_row[col]) else None
                            
                            pos_delta = peak_value - base_value if peak_value > base_value else None
                            neg_delta = peak_value - base_value if peak_value < base_value else None
                            
                            base_name = str(headers[col]).strip()
                            peak_name = str(headers[col + 1]).strip()
                            
                            # Add base measurement
                            data.append({
                                'sheet_name': id_value,
                                'Type': 'Manual Base',
                                'Annotation': base_name,
                                'Event': event_value,
                                'Offset': base_time,
                                'Value': base_value,
                                'mbase': base_value,
                                'mbase_time': base_time,
                                'mpeak': None,
                                'mpeak_time': None,
                                'pos_delta': pos_delta,
                                'neg_delta': neg_delta,
                                'VAS': None,
                                'Notes': ''
                            })
                            
                            # Add peak measurement
                            data.append({
                                'sheet_name': id_value,
                                'Type': 'Manual Peak',
                                'Annotation': peak_name,
                                'Event': event_value,
                                'Offset': peak_time,
                                'Value': peak_value,
                                'mbase': None,
                                'mbase_time': None,
                                'mpeak': peak_value,
                                'mpeak_time': peak_time,
                                'pos_delta': pos_delta,
                                'neg_delta': neg_delta,
                                'VAS': None,
                                'Notes': ''
                            })
                            
                        except Exception as e:
                            print(f"Error processing measurements at column {col} for ID {id_value}: {str(e)}")
                            continue
                
                valid_ids.add(id_value)
        except Exception as e:
            print(f"Error processing sheet {sheet_name}: {str(e)}")
            continue    
    
    result_df = pd.DataFrame(data)
    return valid_ids, result_df

def calculate_manual_deltas_with_selection(data):
    """Calculate deltas only for manual peaks/bases using selected rows"""
    if not isinstance(data, pd.DataFrame) or data.empty:
        return data
        
    if isinstance(data, list):
        data = pd.DataFrame(data)
    
    # Ensure 'selected' column exists
    if 'selected' not in data.columns:
        data['selected'] = True
        
    # Sort by offset to maintain chronological order
    data = data.sort_values('Offset')
    
    # Only process selected rows
    selected_data = data[data['selected']].copy()
    
    last_manual_base = None
    last_manual_peak = None
    
    # Process each row in chronological order
    for idx in data.index:
        if not data.loc[idx, 'selected']:
            continue
            
        if data.loc[idx, 'Type'] == 'Manual Base':
            # Find last selected manual peak before this base
            peak_mask = (data['Type'] == 'Manual Peak') & \
                       (data['Offset'] < data.loc[idx, 'Offset']) & \
                       (data['selected'])
            if peak_mask.any():
                last_manual_peak = data[peak_mask].iloc[-1]
                data.loc[idx, 'pos_delta'] = float(data.loc[idx, 'Value'] - last_manual_peak['Value'])
            
        elif data.loc[idx, 'Type'] == 'Manual Peak':
            # Find last selected manual base before this peak
            base_mask = (data['Type'] == 'Manual Base') & \
                       (data['Offset'] < data.loc[idx, 'Offset']) & \
                       (data['selected'])
            if base_mask.any():
                last_manual_base = data[base_mask].iloc[-1]
                data.loc[idx, 'neg_delta'] = float(data.loc[idx, 'Value'] - last_manual_base['Value'])
    
    return data
    
def calculate_event_statistics(data, selected_sheets):
    """Calculate statistics for events using only selected points"""
    stats_data = []
    
    for sheet_name in selected_sheets:
        if sheet_name not in data:
            continue
            
        df = data[sheet_name]['data']
        if df.empty:
            continue
            
        # Group by Event
        for event_name, event_group in df.groupby('Event'):
            if pd.isna(event_name) or event_name == '':
                continue
                
            # Filter for selected points if the column exists
            if 'selected' in event_group.columns:
                event_group = event_group[event_group['selected']]
            
            if event_group.empty:
                continue
                
            # Sort by Offset
            event_group = event_group.sort_values('Offset')
            
            # Calculate statistics
            stats = {
                'sheet_name': sheet_name,
                'Event': event_name,
                'Start': float(event_group['Offset'].iloc[0]/1000),
                'End': float(event_group['Offset'].iloc[-1]/1000),
                'Avg': float(event_group['Value'].mean()*10),
                'Std_Dev': float(event_group['Value'].std()*10),
                'Max': float(event_group['Value'].max()*10),
                'Min': float(event_group['Value'].min()*10),
                'Delta': float((event_group['Value'].max() - event_group['Value'].min())*10),
                'Median': float(event_group['Value'].median()*10),
                'Q1-25%': float(event_group['Value'].quantile(0.25)*10),
                'Q3-75%': float(event_group['Value'].quantile(0.75)*10),
                'Q1-Q3': float((event_group['Value'].quantile(0.75) - 
                              event_group['Value'].quantile(0.25))*10),
                'Notes': ''
            }
            
            # Calculate Diff
            start_avg = event_group['Value'].head(10).mean()
            end_avg = event_group['Value'].tail(10).mean()
            stats['Diff'] = float((start_avg - end_avg)*10)
            
            stats_data.append(stats)
    
    return stats_data

def calculate_phase_statistics(df):
    """Calculate statistics for a phase dataset"""
    stats = {
        'overall_avg': df['Value'].mean(),
        'overall_max': df['Value'].max(),
        'overall_min': df['Value'].min(),
        'avg_pos_delta': df['pos_delta'].mean(),
        'avg_neg_delta': df['neg_delta'].mean(),
        'peak_count': len(df[df['Type'].str.contains('Peak', na=False)]),
        'base_count': len(df[df['Type'].str.contains('Base', na=False)])
    }
    return pd.Series(stats)

def create_comparison_graphs(data, phase):
    """Create improved comparison graphs for each metric"""
    # Box plot for overall statistics
    overall_fig = px.box(
        data,
        x='number',
        y='overall_avg',
        color='phase',
        title=f'P{phase} Overall Value Distribution',
        labels={'overall_avg': 'Overall Average', 'number': 'Sample'},
        points='all'  # Show all points
    )
    overall_fig.update_traces(boxpoints='all', jitter=0.3)
    
    # Grouped bar chart with error bars for deltas
    delta_fig = go.Figure()
    for delta_type, color in [('avg_pos_delta', 'green'), ('avg_neg_delta', 'red')]:
        mean = data[delta_type].mean()
        std = data[delta_type].std()
        delta_fig.add_trace(go.Bar(
            name=f'{"Positive" if "pos" in delta_type else "Negative"} Deltas',
            x=data['number'],
            y=data[delta_type],
            error_y=dict(type='data', array=[std]*len(data)),
            marker_color=color
        ))
    delta_fig.update_layout(
        title=f'P{phase} Delta Analysis',
        barmode='group',
        xaxis_title='Sample',
        yaxis_title='Delta Value'
    )
    
    # Stacked bar chart for peak counts
    counts_fig = go.Figure(data=[
        go.Bar(name='Peaks', x=data['number'], y=data['peak_count'], marker_color='red'),
        go.Bar(name='Bases', x=data['number'], y=data['base_count'], marker_color='blue')
    ])
    counts_fig.update_layout(
        title=f'P{phase} Peak and Base Counts',
        barmode='stack',
        xaxis_title='Sample',
        yaxis_title='Count'
    )
    
    return overall_fig, delta_fig, counts_fig

def parse_filename(filename):
    """Extract number and phase from filename"""
    match = re.search(r'mnl_py_results_(.+)_p([134])\.xlsx', filename)
    if match:
        return match.group(1), match.group(2)
    return None, None

def load_peaks_data(filename, target_phase):
    """
    Load peaks data from specified worksheet only if file phase matches target phase
    
    Parameters:
    -----------
    filename : str
        The Excel file to load from
    target_phase : str
        The phase to look for ('1', '3', or '4')
        
    Returns:
    --------
    pd.DataFrame
        DataFrame containing the peaks data, or empty DataFrame if no match
    """
    try:
        # Extract file phase
        number, file_phase = parse_filename(filename)
        
        # Only proceed if file phase matches target phase
        if file_phase == target_phase:
            sheet_name = f"P{target_phase}_mnl_py_peaks"
            try:
                df = pd.read_excel(filename, sheet_name=sheet_name)
                df['number'] = number
                df['phase'] = f'P{target_phase}'
                return df
            except Exception as e:
                print(f"Error loading {filename}, {sheet_name}: {str(e)}")
                return pd.DataFrame()
        else:
            # Skip files that don't match the target phase
            return pd.DataFrame()
            
    except Exception as e:
        print(f"Error processing file {filename}: {str(e)}")
        return pd.DataFrame()

def load_and_process_events(filename, target_phase):
    """
    Load and process event statistics from the events worksheet only if file phase matches target phase
    
    Parameters:
    -----------
    filename : str
        The Excel file to load from
    target_phase : str
        The phase to look for ('1', '3', or '4')
        
    Returns:
    --------
    pd.DataFrame
        DataFrame containing the event data, or empty DataFrame if no match
    """
    try:
        # Extract file phase
        number, file_phase = parse_filename(filename)
        
        # Only proceed if file phase matches target phase
        if file_phase == target_phase:
            sheet_name = f"P{target_phase}_mnl_py_events"
            try:
                df = pd.read_excel(filename, sheet_name=sheet_name)
                df['number'] = number
                df['phase'] = f'P{target_phase}'
                
                # Round numeric columns to 3 decimal places
                numeric_cols = ['Avg', 'Delta', 'Diff', 'Max', 'Min', 'Q1-25%', 'Q3-75%', 'Q1-Q3']
                for col in numeric_cols:
                    if col in df.columns:
                        df[col] = df[col].round(3)
                return df
            except Exception as e:
                print(f"Error loading {filename}, {sheet_name}: {str(e)}")
                return pd.DataFrame()
        else:
            # Skip files that don't match the target phase
            return pd.DataFrame()
            
    except Exception as e:
        print(f"Error processing file {filename}: {str(e)}")
        return pd.DataFrame()

def get_results_files():
    """Find all results Excel files matching the pattern"""
    files = []
    for phase in ['1', '3', '4']:
        phase_files = glob.glob(f"mnl_py_results_*_p{phase}.xlsx")
        files.extend(phase_files)
    return sorted(files)

def create_event_comparison_graphs(event_data, phase):
    """Create comparison graphs for event statistics"""
    # Box plot for event averages
    avg_fig = px.box(
        event_data,
        x='Event',
        y='Avg',
        color='number',
        title=f'P{phase} Event Averages Distribution',
        points='all'
    )
    
    # Bar chart for event deltas
    delta_fig = px.bar(
        event_data,
        x='Event',
        y=['Delta', 'Diff'],
        barmode='group',
        title=f'P{phase} Event Deltas Comparison',
        color='number'
    )
    
    return avg_fig, delta_fig
    
def perform_statistical_analysis(data, metric):
    """Perform statistical analysis on metric across phases"""
    if data.empty or metric not in data.columns:
        return pd.DataFrame()
        
    try:
        phases = data['phase'].unique()
        results = []
        
        # Within phase analysis
        for phase in phases:
            phase_data = data[data['phase'] == phase][metric]
            if not phase_data.empty:
                results.append({
                    'comparison': f'{phase} Statistics',
                    'mean': float(phase_data.mean()),
                    'std': float(phase_data.std()),
                    'min': float(phase_data.min()),
                    'max': float(phase_data.max()),
                    't_statistic': None,  # No t-stat for single phase
                    'p_value': None       # No p-value for single phase
                })
        
        # Between phase analysis
        for i in range(len(phases)):
            for j in range(i + 1, len(phases)):
                phase1_data = data[data['phase'] == phases[i]][metric]
                phase2_data = data[data['phase'] == phases[j]][metric]
                
                if not phase1_data.empty and not phase2_data.empty:
                    t_stat, p_value = stats.ttest_ind(
                        phase1_data.dropna(),
                        phase2_data.dropna()
                    )
                    
                    results.append({
                        'comparison': f'{phases[i]} vs {phases[j]}',
                        'mean': None,     # No mean for comparison
                        'std': None,      # No std for comparison
                        'min': None,      # No min for comparison
                        'max': None,      # No max for comparison
                        't_statistic': float(t_stat),
                        'p_value': float(p_value)
                    })
        
        return pd.DataFrame(results)
        
    except Exception as e:
        print(f"Error in perform_statistical_analysis: {str(e)}")
        return pd.DataFrame()

def create_phase_graph_and_table(sheet_data, measurement_data, phase, selected_sheets, 
                               include_words=None, exclude_words=None, include_events=None, 
                               exclude_events=None, show_actions=False, window_size=5):
    """Create graph and table for a phase with selection support and VAS visualization"""
    global VAS_CONFIG  # Access global config
    print(f"\nCreating {phase} graph and table")
    print(f"Selected sheets: {selected_sheets}")
    print(f"Action labels are {'ON' if show_actions else 'OFF'}")
    
    # Create subplot figure with secondary y-axis for VAS bars
    fig = sp.make_subplots(specs=[[{"secondary_y": True}]])
    all_data = []
    
    # Filter sheets for the specified phase
    phase_sheets = {name: data for name, data in sheet_data.items() if phase in name}
    if selected_sheets:
        phase_sheets = {name: data for name, data in phase_sheets.items() if name in selected_sheets}
    
    print(f"Processing sheets: {list(phase_sheets.keys())}")
    
    for sheet_name, sheet_info in phase_sheets.items():
        print(f"\nProcessing sheet: {sheet_name}")
        data = sheet_info['data']
        
        # Filter data based on words and events
        filtered_data = filter_by_words(data.copy(), include_words, exclude_words)
        filtered_data = filter_by_events(filtered_data, include_events, exclude_events)

        if not filtered_data.empty:
            # Add main line trace
            fig.add_trace(
                go.Scatter(
                    name=f'{sheet_name}',
                    x=filtered_data['Offset']/1000,
                    y=filtered_data['Value']*10,
                    mode='lines',
                    line=dict(color=COLORS['line'], width=2),
                    showlegend=True
                ),
                secondary_y=False
            )

        # Add event span lines
        if not filtered_data.empty:
            events = filtered_data.groupby('Event')
            used_y_positions = []
            
            for event_name, event_group in events:
                if pd.notna(event_name) and event_name != '':
                    # Get min and max offset for this event
                    min_offset = event_group['Offset'].min() / 1000
                    max_offset = event_group['Offset'].max() / 1000
                    #label_x = (min_offset + max_offset) / 2
                    
                    # Add vertical lines
                    fig.add_trace(
                        go.Scatter(
                            x=[min_offset, min_offset],
                            y=[filtered_data['Value'].min() * 10, filtered_data['Value'].max() * 10],
                            mode='lines',
                            line=dict(color='rgba(128, 128, 128, 0.5)', width=1, dash='dash'),
                            name=f'Event {event_name} Start',
                            showlegend=False
                        ),
                        secondary_y=False
                    )
                    
                    fig.add_trace(
                        go.Scatter(
                            x=[max_offset, max_offset],
                            y=[filtered_data['Value'].min() * 10, filtered_data['Value'].max() * 10],
                            mode='lines',
                            line=dict(color='rgba(128, 128, 128, 0.5)', width=1, dash='dash'),
                            name=f'Event {event_name} End',
                            showlegend=False
                        ),
                        secondary_y=False
                    )
                    
                    # Find available y position for label
                    base_y = filtered_data['Value'].max() * 10
                    label_x = (min_offset + max_offset) / 2  # Place label in middle of event span
                    
                    # Find available y position
                    y_pos = find_available_y_position(
                        label_x, 
                        base_y, 
                        used_positions=used_y_positions,
                        min_gap=2.0
                    )
                    used_y_positions.append((label_x, y_pos))
                    
                    # Add event label
                    fig.add_annotation(
                        x=min_offset,
                        y=filtered_data['Value'].max() * 10,
                        text=f'Event {event_name}',
                        showarrow=False,
                        yshift=10,
                        textangle=45,
                        font=dict(size=10)
                    )
            
            rolling_points = detect_peaks_rolling_window(filtered_data, sheet_name)

            if not rolling_points.empty:
                rolling_points['selected'] = True
                # Add VAS values for rolling points
                rolling_points['VAS'] = None
                for idx in rolling_points.index:
                    offset = rolling_points.loc[idx, 'Offset']
                    matching_rows = filtered_data[filtered_data['Offset']/1000 == offset]
                    if not matching_rows.empty:
                        vas_value = matching_rows.iloc[0]['VAS']
                        rolling_points.at[idx, 'VAS'] = vas_value
                
                # Add rolling bases to graph
                rolling_bases = rolling_points[rolling_points['Type'] == 'Rolling Base']
                if not rolling_bases.empty:
                    fig.add_trace(
                        go.Scatter(
                            name=f'{sheet_name} (Rolling Base)',
                            x=rolling_bases['Offset'],
                            y=rolling_bases['Value'],
                            mode='markers',
                            marker=dict(
                                color='rgb(102, 255, 51)',
                                size=12,
                                symbol='circle',
                                line=dict(color='white', width=1)
                            ),
                            showlegend=True
                        ),
                        secondary_y=False
                    )
                
                # Add rolling peaks to graph
                rolling_peaks = rolling_points[rolling_points['Type'] == 'Rolling Peak']
                if not rolling_peaks.empty:
                    fig.add_trace(
                        go.Scatter(
                            name=f'{sheet_name} (Rolling Peak)',
                            x=rolling_peaks['Offset'],
                            y=rolling_peaks['Value'],
                            mode='markers',
                            marker=dict(
                                color='rgb(0, 176, 240)',
                                size=12,
                                symbol='circle',
                                line=dict(color='white', width=1)
                            ),
                            showlegend=True
                        ),
                        secondary_y=False
                    )

                    # Add negative delta labels for rolling peaks
                    roll_peak_labels = rolling_peaks[rolling_peaks['neg_delta'].notna()]
                    if not roll_peak_labels.empty:
                        fig.add_trace(
                            go.Scatter(
                                name=f'{sheet_name} (Rolling Peak Deltas)',
                                x=roll_peak_labels['Offset'],
                                y=[val - 0.3 for val in roll_peak_labels['Value']],
                                mode='text',
                                text=[f'Δ: {delta:.1f}' for delta in roll_peak_labels['neg_delta']],
                                textposition='bottom right',
                                textfont=dict(size=10, color='purple'),
                                showlegend=False,
                                hoverinfo='skip'
                            ),
                            secondary_y=False
                        )
                
                all_data.append(rolling_points)

            # Add VAS bar graph
            vas_data = filtered_data[filtered_data['VAS'] >= 0]
            if not vas_data.empty:
                # Add VAS bars
                fig.add_trace(
                    go.Bar(
                        name=f'{sheet_name} VAS',
                        x=vas_data['Offset']/1000,
                        y=vas_data['VAS'],
                        marker_color='rgba(255, 0, 0, 0.3)',
                        showlegend=True
                    ),
                    secondary_y=True
                )

                fig.update_yaxes(
                    title_text="VAS Score",
                    showgrid=False,
                    range=[-1, 10],  # Updated range to start at -1
                    secondary_y=True
                )
                
                # Process VAS data to find spans of same values
                vas_spans = []
                current_vas = None
                start_offset = None
                
                # Sort by offset to ensure proper order
                vas_data_sorted = vas_data.sort_values('Offset')
                
                for _, row in vas_data_sorted.iterrows():
                    if current_vas != row['VAS']:
                        if current_vas is not None and current_vas >= 0:
                            vas_spans.append({
                                'value': current_vas,
                                'offset': (start_offset + last_offset) / 2  # Middle of the span
                            })
                        current_vas = row['VAS']
                        start_offset = row['Offset']
                    last_offset = row['Offset']
                
                # Add the last span
                if current_vas is not None:
                    vas_spans.append({
                        'value': current_vas,
                        'offset': (start_offset + last_offset) / 2
                    })
                
                # Add VAS labels at the middle of each span
                if vas_spans:
                    fig.add_trace(
                        go.Scatter(
                            name=f'{sheet_name} VAS Labels',
                            x=[span['offset']/1000 for span in vas_spans],
                            y=[span['value'] for span in vas_spans],
                            mode='text',
                            text=[f'VAS: {span["value"]:.1f}' for span in vas_spans],
                            textposition='top center',
                            textfont=dict(size=10, color='red'),
                            showlegend=False,
                            hoverinfo='skip'
                        ),
                        secondary_y=True
                    )

        # Filter and add manual measurements
        manual_data = measurement_data[measurement_data['sheet_name'] == sheet_name].copy()
        if include_words or exclude_words:
            manual_data = filter_by_words(manual_data, include_words, exclude_words)
        if include_events or exclude_events:
            manual_data = filter_by_events(manual_data, include_events, exclude_events)

        if not manual_data.empty:
            manual_data['selected'] = True  # Default to selected
            # Add manual bases
            manual_bases = manual_data[manual_data['Type'] == 'Manual Base']
            if not manual_bases.empty:
                fig.add_trace(
                    go.Scatter(
                        name=f'{sheet_name} (Manual Base)',
                        x=manual_bases['Offset'],
                        y=manual_bases['Value'],
                        mode='markers',
                        marker=dict(
                            color=COLORS['mbase'],
                            size=8,
                            symbol='diamond',
                            line=dict(color='white', width=1)
                        ),
                        showlegend=True
                    )
                )
            
            # Add manual peaks
            manual_peaks = manual_data[manual_data['Type'] == 'Manual Peak']
            if not manual_peaks.empty:
                fig.add_trace(
                    go.Scatter(
                        name=f'{sheet_name} (Manual Peak)',
                        x=manual_peaks['Offset'],
                        y=manual_peaks['Value'],
                        mode='markers',
                        marker=dict(
                            color=COLORS['mpeak'],
                            size=8,
                            symbol='diamond',
                            line=dict(color='white', width=1)
                        ),
                        showlegend=True
                    )
                )
            
            all_data.append(manual_data)

            # Calculate event averages
            event_stats = {}
            event_starts = {}
            for event_name, event_group in filtered_data.groupby('Event'):
                if pd.notna(event_name) and event_name != '':
                    avg_value = event_group['Value'].mean() * 10  # Convert to same scale
                    event_stats[event_name] = avg_value
                    # Find the first occurrence (minimum offset) for this event
                    event_starts[event_name] = event_group['Offset'].min()

            # Add action labels and event averages if enabled
            if show_actions:
                # Track used positions for annotation placement
                used_positions = []
                            
                # Calculate event averages
                event_stats = {}
                event_starts = {}
                for event_name, event_group in filtered_data.groupby('Event'):
                    if pd.notna(event_name) and event_name != '':
                        avg_value = event_group['Value'].mean() * 10
                        event_stats[event_name] = avg_value
                        event_starts[event_name] = event_group['Offset'].min()
                
                non_empty_annotations = filtered_data[filtered_data['Annotations'].notna()]
                if not non_empty_annotations.empty:
                    # Keep track of which events we've already labeled
                    labeled_events = set()

                    # Sort annotations by x position to process them in order
                    sorted_annotations = non_empty_annotations.sort_values('Offset')
                    
                    for _, row in sorted_annotations.iterrows():
                        event_name = row['Event']
                        x_pos = row['Offset']/1000
                        base_y = (row['Value']*10) + 0.75
                        
                        # Find an available y position that doesn't overlap
                        adjusted_y = find_available_y_position(
                            x_pos, 
                            base_y, 
                            used_positions,
                            min_gap=1.0
                        )
                        
                        # Only add the average if this is the first occurrence of this event
                        if event_name in event_stats and event_name not in labeled_events and \
                           row['Offset'] == event_starts[event_name]:
                            avg_text = f"\nAvg: {event_stats[event_name]:.2f}"
                            labeled_events.add(event_name)
                        else:
                            avg_text = ""
                            
                        # Add the trace with adjusted y position
                        fig.add_trace(
                            go.Scatter(
                                name=f'{sheet_name} (Actions)',
                                x=[x_pos],
                                y=[adjusted_y],
                                mode='text',
                                text=[f"{row['Annotations']}{avg_text}"],
                                textposition='top center',
                                textfont=dict(size=12, color='black'),
                                showlegend=False,
                                hoverinfo='skip'
                            ),
                            secondary_y=False
                        )
                                    
                        # Record the used position
                        used_positions.append((x_pos, adjusted_y))

    # Update layout
    fig.update_layout(
        title=dict(
            text=f'{phase} Series Analysis',
            x=0.5,
            font=dict(size=24)
        ),
        xaxis_title='Time (seconds)',
        legend=dict(
            x=1.05,
            y=1,
            bordercolor='Black',
            borderwidth=1
        ),
        showlegend=True,
        height=600,
        margin=dict(l=50, r=150, t=50, b=50),
        plot_bgcolor='white',
        paper_bgcolor='white'
    )
    
    # Update axes
    fig.update_yaxes(
        title_text="Value",
        showgrid=True,
        gridwidth=1,
        gridcolor='LightGray',
        zeroline=True,
        zerolinewidth=1,
        zerolinecolor='LightGray',
        secondary_y=False
    )
    
    fig.update_yaxes(
        title_text="VAS Score",
        showgrid=False,
        range=[-1, 10],  # VAS scale is typically -1-10
        secondary_y=True
    )

    # Combine all data for the table
    if all_data:
        combined_data = pd.concat(all_data, ignore_index=True)
        combined_data = combined_data.sort_values(['sheet_name', 'Offset'])
        combined_data['selected'] = True
        print(f"\nTable summary for {phase}:")
        print(f"Number of rows: {len(combined_data)}")
        print(f"Columns: {combined_data.columns.tolist()}")
        print(f"Data types: {combined_data.dtypes}")
        print(f"Sample of first few rows:")
        print(combined_data.head())
        return fig, combined_data
    else:
        empty_df = pd.DataFrame(columns=['sheet_name', 'Type', 'Annotation', 'Event', 'Offset', 'Value', 
                                       'mbase', 'mbase_time', 'mpeak', 'mpeak_time', 
                                       'pos_delta', 'neg_delta', 'VAS', 'Notes', 'selected'])
        empty_df['selected'] = True
        print(f"\nNo data found for {phase}, returning empty DataFrame")
        return fig, empty_df

def save_single_phase_to_excel(phase, fig_data, table_data, event_data, base_filename='mnl_py_results'):
    """
    Save a single phase's graph, peaks table, and event statistics to Excel workbook.
    
    Parameters:
    -----------
    phase : str
        Phase identifier (P1, P3, or P4)
    fig_data : dict
        Plotly figure data dictionary from dcc.Graph
    table_data : list or pd.DataFrame
        Peak detection data
    event_data : list or pd.DataFrame
        Event statistics data
    base_filename : str
        Base filename without extension
    """
    filename = f"{base_filename}.xlsx"
    temp_image_dir = 'temp_images'
    
    # Create temp directory if it doesn't exist
    os.makedirs(temp_image_dir, exist_ok=True)
        
    # Use os.path.join for proper path handling
    img_path = os.path.join(temp_image_dir, f"{phase}_chart.png")
    
    try:
        # Convert the figure data to a plotly figure object
        fig = go.Figure(fig_data)
        
        # Save current graph state as image with specific settings
        pio.write_image(
            fig,
            img_path,
            width=1000, 
            height=600,
            scale=2,  # Increase resolution
            engine='kaleido'
        )
        
        # Load existing workbook if it exists, otherwise create new
        try:
            wb = load_workbook(filename)
        except FileNotFoundError:
            wb = Workbook()
            if 'Sheet' in wb.sheetnames:
                wb.remove(wb.active)
        
        # Sheet names
        chart_sheet_name = f"{phase}_mnl_py_chart"
        peaks_sheet_name = f"{phase}_mnl_py_peaks"
        events_sheet_name = f"{phase}_mnl_py_events"
        
        # Remove existing sheets if they exist
        for sheet_name in [chart_sheet_name, peaks_sheet_name, events_sheet_name]:
            if sheet_name in wb.sheetnames:
                wb.remove(wb[sheet_name])
        
        # Create new sheets
        ws_chart = wb.create_sheet(chart_sheet_name)
        ws_peaks = wb.create_sheet(peaks_sheet_name)
        ws_events = wb.create_sheet(events_sheet_name)
        
        # Add image to chart sheet with error handling
        if os.path.exists(img_path) and os.path.getsize(img_path) > 0:
            try:
                img = Image(img_path)
                ws_chart.add_image(img, 'A1')
            except Exception as e:
                print(f"Warning: Could not add image to worksheet: {str(e)}")
                ws_chart.cell(row=1, column=1, value="Graph image could not be loaded")
        else:
            print(f"Warning: Image file not found or empty at {img_path}")
            ws_chart.cell(row=1, column=1, value="Graph image could not be generated")

        def write_dataframe_to_sheet(df, worksheet):
            if df.empty:
                worksheet.cell(row=1, column=1, value="No data available")
                return
                
            # Write headers
            for col, header in enumerate(df.columns, 1):
                worksheet.cell(row=1, column=col, value=str(header))
            
            # Write data
            for row_idx, row in enumerate(df.values, 2):
                for col_idx, value in enumerate(row, 1):
                    if pd.isna(value):
                        cell_value = ''
                    elif isinstance(value, (float, np.float64)):
                        cell_value = float(value)
                    elif isinstance(value, (int, np.int64)):
                        cell_value = int(value)
                    else:
                        cell_value = str(value)
                    worksheet.cell(row=row_idx, column=col_idx, value=cell_value)

        # Convert and write table data
        if isinstance(table_data, list):
            table_data = pd.DataFrame(table_data)
        write_dataframe_to_sheet(table_data, ws_peaks)

        # Convert and write event data
        if isinstance(event_data, list):
            event_data = pd.DataFrame(event_data)
        write_dataframe_to_sheet(event_data, ws_events)
        
        # Save workbook
        wb.save(filename)
        print(f"Successfully saved {phase} results to {filename}")
        return True, f"Saved {phase} results to {filename}"
        
    except Exception as e:
        print(f"Error saving {phase} results: {str(e)}")
        print("Traceback:")
        traceback.print_exc()
        return False, f"Error saving {phase} results: {str(e)}"
        
    finally:
        # Clean up temporary files
        try:
            if os.path.exists(img_path):
                os.remove(img_path)
            if os.path.exists(temp_image_dir) and not os.listdir(temp_image_dir):
                os.rmdir(temp_image_dir)
        except Exception as e:
            print(f"Error cleaning up temporary files: {str(e)}")
            
# Load data
try:
    # First read measurement data to get valid IDs
    valid_ids, measurement_data = read_measurement_data('Mnl_Analysis.xlsx')

    # Then read the phase data
    sheet_data = read_excel_sheets('All_Phase_Sheets.xlsx')

    # Add validation message
    print("\nData Validation:")
    if not measurement_data.empty and sheet_data:
        print("Successfully loaded measurement and phase data")
    else:
        print("Warning: Some data may be missing or empty")

except FileNotFoundError as e:
    print(f"Error loading files: {str(e)}")
    sheet_data = {}
    measurement_data = pd.DataFrame()
    print("No data loaded")

def get_rgba_values(color_name):
    """Convert color name to RGB values"""
    color_map = {
        'red': '255,0,0',
        'blue': '0,0,255',
        'green': '0,255,0',                    
        'purple': '128,0,128',
        'orange': '255,165,0'
    }
    return color_map.get(color_name, '255,0,0')


# def normalize_time_series(df):
#     """Normalize the time series to start at 0"""
#     if not overlay_df.empty:
#         min_offset = overlay_df['Offset'].min()
#         overlay_df = overlay_df.copy()
#         overlay_df['Offset'] = overlay_df['Offset'] - min_offset
#     return overlay_df

def create_phase_overlay_graph(p1_overlay_sheets, p3_overlay_sheets, p4_overlay_sheets, sheet_data):
    """
    Create an overlay graph combining data from selected sheets across phases
    with time-normalized offsets.
    """
    print("\nCreating phase overlay graph")
    fig = go.Figure()
    
    colors = {
        'P1': 'rgb(31, 119, 180)',  # Blue
        'P3': 'rgb(255, 127, 14)',  # Orange
        'P4': 'rgb(44, 160, 44)'    # Green
    }
    
    # Process each phase's data
    for phase_name, selected_sheets, color in [
        ('P1', p1_overlay_sheets, colors['P1']),
        ('P3', p3_overlay_sheets, colors['P3']),
        ('P4', p4_overlay_sheets, colors['P4'])
    ]:
        if not selected_sheets:
            print(f"No sheets selected for {phase_name}, skipping")
            continue
            
        print(f"\nProcessing {phase_name} data:")
        # Process each selected sheet
        for sheet_name in selected_sheets:
            if sheet_name not in sheet_data:
                print(f"Sheet {sheet_name} not found in data, skipping")
                continue
                
            overlay_df = sheet_data[sheet_name]['data']
            if overlay_df.empty:
                print(f"Empty dataframe for {sheet_name}, skipping")
                continue
            
            print(f"Processing sheet: {sheet_name}")
            # Normalize time series to start at 0
            min_offset = overlay_df['Offset'].min()
            normalized_offset = (overlay_df['Offset'] - min_offset) / 1000  # Convert to seconds
            
            # Add trace for this sheet
            fig.add_trace(
                go.Scatter(
                    name=f'{phase_name} - {sheet_name}',
                    x=normalized_offset,
                    y=overlay_df['Value'] * 10,  # Fixed: Using overlay_df instead of df
                    mode='lines',
                    line=dict(
                        color=color,
                        width=2
                    ),
                    opacity=0.7
                )
            )
            print(f"Added trace for {sheet_name}")
    
    # Update layout
    fig.update_layout(
        title='Phase Overlay Comparison (Time Normalized)',
        xaxis_title='Time (seconds from start)',
        yaxis_title='Value',
        hovermode='closest',
        showlegend=True,
        legend=dict(
            yanchor="top",
            y=0.99,
            xanchor="left",
            x=1.05
        ),
        plot_bgcolor='white',
        paper_bgcolor='white',
        margin=dict(r=150),
        height=600
    )
    
    # Update axes
    fig.update_xaxis(showgrid=True, gridwidth=1, gridcolor='LightGray')
    fig.update_yaxis(showgrid=True, gridwidth=1, gridcolor='LightGray')
    
    print("Completed creating overlay graph")
    return fig

# Define app layout
app.layout = html.Div([
    dcc.Tabs([
        dcc.Tab(label='Peak Detection', children=[
            html.H1('Peak Detection Analysis Dashboard'),
                    
            # Display Controls
            html.Div([
                html.H3('Display Controls', style={'marginBottom': '10px'}),
                html.Div([
                    # Action toggle button
                    html.Button(
                        'Toggle Action Labels',
                        id='action-toggle',
                        n_clicks=0,
                        style={
                            'marginRight': '10px',
                            'padding': '5px 10px',
                            'backgroundColor': '#f8f9fa',
                            'border': '1px solid #ddd',
                            'borderRadius': '4px',
                            'cursor': 'pointer'
                        }
                    ),
                    html.Div(
                        "Action Labels OFF",
                        id='action-toggle-status',
                        style={'marginLeft': '10px', 'color': 'red'}
                    ),
                    # Window size input
                    html.Div([
                        html.Label('Window Size:', style={'marginLeft': '20px', 'marginRight': '10px'}),
                        dcc.Input(
                            id='window-size-input',
                            type='number',
                            value=5,
                            min=2,
                            max=20,
                            step=1,
                            style={'width': '60px'}
                        )
                    ], style={'display': 'inline-block'})
                ], style={'marginBottom': '20px', 'display': 'flex', 'alignItems': 'center'})
            ], style={'padding': '10px', 'backgroundColor': '#f8f9fa', 'borderRadius': '5px', 'marginBottom': '20px'}),
            
            # Word and Event Filters
            html.Div([
                html.Label('Include Words (comma-separated):'),
                dcc.Input(
                    id='include-words', 
                    type='text', 
                    value=','.join(DEFAULT_INCLUDE_WORDS),
                    style={'marginRight': '20px', 'width': '200px'}
                ),
                html.Label('Exclude Words (comma-separated):'),
                dcc.Input(
                    id='exclude-words', 
                    type='text', 
                    value=','.join(DEFAULT_EXCLUDE_WORDS),
                    style={'width': '200px'}
                )
            ], style={'marginBottom': '20px', 'display': 'flex', 'gap': '10px', 'alignItems': 'center'}),
            
            html.Div([
                html.Label('Include Events (comma-separated):'),
                dcc.Input(
                    id='include-events', 
                    type='text', 
                    value=','.join(DEFAULT_INCLUDE_EVENTS),
                    style={'marginRight': '20px', 'width': '200px'}
                ),
                html.Label('Exclude Events (comma-separated):'),
                dcc.Input(
                    id='exclude-events', 
                    type='text', 
                    value=','.join(DEFAULT_EXCLUDE_EVENTS),
                    style={'width': '200px'}
                )
            ], style={'marginBottom': '20px', 'display': 'flex', 'gap': '10px', 'alignItems': 'center'}),
            
            # P1 Section
            html.Div([
                html.H2('P1 Analysis'),
                html.Div([
                    html.Div([
                        html.Label('Select P1 Worksheets:', style={'marginBottom': '5px'}),
                        dcc.Dropdown(
                            id='p1-worksheet-selector',
                            options=[
                                {'label': name, 'value': name} 
                                for name in sheet_data.keys() 
                                if 'P1' in name
                            ],
                            value=[],
                            multi=True,
                            placeholder="Select P1 worksheets (leave empty for all)",
                            style={'width': '500px'}
                        ),
                    ], style={'flex': '1'}),
                    html.Div([
                        html.Label('Filename: '),
                        dcc.Input(
                            id='p1-filename-input',
                            type='text',
                            placeholder='Enter filename (without .xlsx)',
                            value='mnl_py_results',
                            style={'width': '200px', 'marginRight': '10px'}
                        ),
                        html.Button(
                            'Save P1 Results',
                            id='save-p1-button',
                            n_clicks=0,
                            style={'marginLeft': '10px', 'height': '40px'}
                        ),
                    ], style={'display': 'flex', 'alignItems': 'center'}),
                    html.Div(id='p1-save-status')
                ], style={'display': 'flex', 'alignItems': 'flex-end', 'marginBottom': '10px'}),
                
                html.Div([
                    html.Button(
                        'Select All', 
                        id='p1-select-all', 
                        n_clicks=0,
                        style={
                            'marginRight': '10px',
                            'padding': '5px 10px',
                            'backgroundColor': '#f8f9fa',
                            'border': '1px solid #ddd',
                            'borderRadius': '4px',
                            'cursor': 'pointer'
                        }
                    ),
                    html.Button(
                        'Deselect All', 
                        id='p1-deselect-all', 
                        n_clicks=0,
                        style={
                            'padding': '5px 10px',
                            'backgroundColor': '#f8f9fa',
                            'border': '1px solid #ddd',
                            'borderRadius': '4px',
                            'cursor': 'pointer'
                        }
                    ),
                    dcc.ConfirmDialog(
                        id='p1-confirm-deselect',
                        message='Are you sure you want to deselect all rows?',
                    ),
                ], style={'marginBottom': '10px'}),
                
                dcc.Graph(id='p1-graph'),
                html.H3('P1 Inflection Points', style={'marginTop': '20px', 'marginBottom': '10px'}),
                
                # P1 Table with VAS column
                dash_table.DataTable(
                    id='p1-table',
                    columns=[
                        {"name": "Include", "id": "selected", "type": "any"},
                        {"name": "Offset", "id": "Offset", "type": "numeric", "format": {"specifier": ".3f"}},
                        {"name": "Value", "id": "Value", "type": "numeric", "format": {"specifier": ".2f"}},
                        {"name": "Type", "id": "Type"},
                        {"name": "Annotation", "id": "Annotation"},
                        {"name": "Event", "id": "Event"},
                        {"name": "sheet_name", "id": "sheet_name"},
                        {"name": "mbase", "id": "mbase", "type": "numeric", "format": {"specifier": ".2f"}},
                        {"name": "mbase time", "id": "mbase_time", "type": "numeric", "format": {"specifier": ".0f"}},
                        {"name": "mpeak", "id": "mpeak", "type": "numeric", "format": {"specifier": ".2f"}},
                        {"name": "mpeak time", "id": "mpeak_time", "type": "numeric", "format": {"specifier": ".0f"}},
                        {"name": "pos_delta", "id": "pos_delta", "type": "numeric", "format": {"specifier": ".2f"}},
                        {"name": "neg_delta", "id": "neg_delta", "type": "numeric", "format": {"specifier": ".2f"}},
                        {"name": "VAS", "id": "VAS", "type": "numeric", "format": {"specifier": ".1f"}},
                        {"name": "Notes", "id": "Notes", "presentation": "input"}
                    ],
                    data=[],
                    persistence=True,
                    persistence_type='session',
                    editable=True,
                    row_selectable='multi',
                    selected_rows=[],
                    style_table={'height': '300px', 'overflowY': 'auto'},
                    style_cell={
                        'minWidth': 95,
                        'maxWidth': 200,
                        'width': 95,
                        'textAlign': 'left'
                    },
                    style_cell_conditional=[
                        {
                            'if': {'column_id': 'Notes'},
                            'width': '200px',
                            'textAlign': 'left'
                        },
                        {
                            'if': {'column_id': 'selected'},
                            'textAlign': 'center'
                        }
                    ],
                    style_data_conditional=[
                        {
                            'if': {'row_index': 'odd'},
                            'backgroundColor': 'rgb(248, 248, 248)'
                        },
                        {
                            'if': {'column_id': 'pos_delta', 'filter_query': '{pos_delta} > 0'},
                            'backgroundColor': '#e6ffe6'
                        },
                        {
                            'if': {'column_id': 'neg_delta', 'filter_query': '{neg_delta} < 0'},
                            'backgroundColor': '#ffe6e6'
                        }
                    ]
                ),
                
                html.H3('P1 Event Statistics', style={'marginTop': '20px', 'marginBottom': '10px'}),
                dash_table.DataTable(
                    id='p1-stats-table',
                    columns=[
                        {"name": "sheet_name", "id": "sheet_name"},
                        {"name": "Event", "id": "Event"},
                        {"name": "Start", "id": "Start", "type": "numeric", "format": {"specifier": ".3f"}},
                        {"name": "End", "id": "End", "type": "numeric", "format": {"specifier": ".3f"}},
                        {"name": "Avg", "id": "Avg", "type": "numeric", "format": {"specifier": ".3f"}},
                        {"name": "Std_Dev", "id": "Std_Dev", "type": "numeric", "format": {"specifier": ".3f"}},
                        {"name": "Max", "id": "Max", "type": "numeric", "format": {"specifier": ".3f"}},
                        {"name": "Min", "id": "Min", "type": "numeric", "format": {"specifier": ".3f"}},
                        {"name": "Delta", "id": "Delta", "type": "numeric", "format": {"specifier": ".3f"}},
                        {"name": "Median", "id": "Median", "type": "numeric", "format": {"specifier": ".3f"}},
                        {"name": "Q1-25%", "id": "Q1-25%", "type": "numeric", "format": {"specifier": ".3f"}},
                        {"name": "Q3-75%", "id": "Q3-75%", "type": "numeric", "format": {"specifier": ".3f"}},
                        {"name": "Q1-Q3", "id": "Q1-Q3", "type": "numeric", "format": {"specifier": ".3f"}},
                        {"name": "Notes", "id": "Notes", "presentation": "input"}
                    ],
                    data=[],
                    editable=True,
                    style_table={'height': '300px', 'overflowY': 'auto'},
                    style_cell={
                        'minWidth': 95,
                        'maxWidth': 200,
                        'width': 95,
                        'textAlign': 'left'
                    },
                    style_cell_conditional=[
                        {
                            'if': {'column_id': 'Notes'},
                            'width': '200px',
                            'textAlign': 'left'
                        }
                    ],
                    style_data_conditional=[
                        {
                            'if': {'row_index': 'odd'},
                            'backgroundColor': 'rgb(248, 248, 248)'
                        }
                    ]
                )
            ], style={'marginBottom': '30px', 'padding': '20px', 'backgroundColor': '#ffffff', 'borderRadius': '5px', 'boxShadow': '0 2px 4px rgba(0,0,0,0.1)'}),
            
            # P3 Section
            html.Div([
                html.H2('P3 Analysis'),
                html.Div([
                    html.Div([
                        html.Label('Select P3 Worksheets:', style={'marginBottom': '5px'}),
                        dcc.Dropdown(
                            id='p3-worksheet-selector',
                            options=[
                                {'label': name, 'value': name} 
                                for name in sheet_data.keys() 
                                if 'P3' in name
                            ],
                            value=[],
                            multi=True,
                            placeholder="Select P3 worksheets (leave empty for all)",
                            style={'width': '500px'}
                        ),
                    ], style={'flex': '1'}),
                    html.Div([
                        html.Label('Filename: '),
                        dcc.Input(
                            id='p3-filename-input',
                            type='text',
                            placeholder='Enter filename (without .xlsx)',
                            value='mnl_py_results',
                            style={'width': '200px', 'marginRight': '10px'}
                        ),
                        html.Button(
                            'Save P3 Results',
                            id='save-p3-button',
                            n_clicks=0,
                            style={'marginLeft': '10px', 'height': '40px'}
                        ),
                    ], style={'display': 'flex', 'alignItems': 'center'}),
                    html.Div(id='p3-save-status')
                ], style={'display': 'flex', 'alignItems': 'flex-end', 'marginBottom': '10px'}),

                html.Div([
                    html.Button(
                        'Select All', 
                        id='p3-select-all', 
                        n_clicks=0,
                        style={
                            'marginRight': '10px',
                            'padding': '5px 10px',
                            'backgroundColor': '#f8f9fa',
                            'border': '1px solid #ddd',
                            'borderRadius': '4px',
                            'cursor': 'pointer'
                        }
                    ),
                    html.Button(
                        'Deselect All', 
                        id='p3-deselect-all', 
                        n_clicks=0,
                        style={
                            'padding': '5px 10px',
                            'backgroundColor': '#f8f9fa',
                            'border': '1px solid #ddd',
                            'borderRadius': '4px',
                            'cursor': 'pointer'
                        }
                    ),
                    dcc.ConfirmDialog(
                        id='p3-confirm-deselect',
                        message='Are you sure you want to deselect all rows?',
                    ),
                ], style={'marginBottom': '10px'}),
                
                dcc.Graph(id='p3-graph'),
                html.H3('P3 Inflection Points', style={'marginTop': '20px', 'marginBottom': '10px'}),
                
                # P3 Table with VAS column
                dash_table.DataTable(
                    id='p3-table',
                    columns=[
                        {"name": "Include", "id": "selected", "type": "any"},
                        {"name": "Offset", "id": "Offset", "type": "numeric", "format": {"specifier": ".3f"}},
                        {"name": "Value", "id": "Value", "type": "numeric", "format": {"specifier": ".2f"}},
                        {"name": "Type", "id": "Type"},
                        {"name": "Annotation", "id": "Annotation"},
                        {"name": "Event", "id": "Event"},
                        {"name": "sheet_name", "id": "sheet_name"},
                        {"name": "mbase", "id": "mbase", "type": "numeric", "format": {"specifier": ".2f"}},
                        {"name": "mbase time", "id": "mbase_time", "type": "numeric", "format": {"specifier": ".0f"}},
                        {"name": "mpeak", "id": "mpeak", "type": "numeric", "format": {"specifier": ".2f"}},
                        {"name": "mpeak time", "id": "mpeak_time", "type": "numeric", "format": {"specifier": ".0f"}},
                        {"name": "pos_delta", "id": "pos_delta", "type": "numeric", "format": {"specifier": ".2f"}},
                        {"name": "neg_delta", "id": "neg_delta", "type": "numeric", "format": {"specifier": ".2f"}},
                        {"name": "VAS", "id": "VAS", "type": "numeric", "format": {"specifier": ".1f"}},
                        {"name": "Notes", "id": "Notes", "presentation": "input"}
                    ],
                    data=[],
                    persistence=True,
                    persistence_type='session',
                    editable=True,
                    row_selectable='multi',
                    selected_rows=[],
                    style_table={'height': '300px', 'overflowY': 'auto'},
                    style_cell={
                        'minWidth': 95,
                        'maxWidth': 200,
                        'width': 95,
                        'textAlign': 'left'
                    },
                    style_cell_conditional=[
                        {
                            'if': {'column_id': 'Notes'},
                            'width': '200px',
                            'textAlign': 'left'
                        },
                        {
                            'if': {'column_id': 'selected'},
                            'textAlign': 'center'
                        }
                    ],
                    style_data_conditional=[
                        {
                            'if': {'row_index': 'odd'},
                            'backgroundColor': 'rgb(248, 248, 248)'
                        },
                        {
                            'if': {'column_id': 'pos_delta', 'filter_query': '{pos_delta} > 0'},
                            'backgroundColor': '#e6ffe6'
                        },
                        {
                            'if': {'column_id': 'neg_delta', 'filter_query': '{neg_delta} < 0'},
                            'backgroundColor': '#ffe6e6'
                        }
                    ]
                ),

                html.H3('P3 Event Statistics', style={'marginTop': '20px', 'marginBottom': '10px'}),
                dash_table.DataTable(
                    id='p3-stats-table',
                    columns=[
                        {"name": "sheet_name", "id": "sheet_name"},
                        {"name": "Event", "id": "Event"},
                        {"name": "Start", "id": "Start", "type": "numeric", "format": {"specifier": ".3f"}},
                        {"name": "End", "id": "End", "type": "numeric", "format": {"specifier": ".3f"}},
                        {"name": "Avg", "id": "Avg", "type": "numeric", "format": {"specifier": ".3f"}},
                        {"name": "Std_Dev", "id": "Std_Dev", "type": "numeric", "format": {"specifier": ".3f"}},
                        {"name": "Max", "id": "Max", "type": "numeric", "format": {"specifier": ".3f"}},
                        {"name": "Min", "id": "Min", "type": "numeric", "format": {"specifier": ".3f"}},
                        {"name": "Delta", "id": "Delta", "type": "numeric", "format": {"specifier": ".3f"}},
                        {"name": "Median", "id": "Median", "type": "numeric", "format": {"specifier": ".3f"}},
                        {"name": "Q1-25%", "id": "Q1-25%", "type": "numeric", "format": {"specifier": ".3f"}},
                        {"name": "Q3-75%", "id": "Q3-75%", "type": "numeric", "format": {"specifier": ".3f"}},
                        {"name": "Q1-Q3", "id": "Q1-Q3", "type": "numeric", "format": {"specifier": ".3f"}},
                        {"name": "Notes", "id": "Notes", "presentation": "input"}
                    ],
                    data=[],
                    editable=True,
                    style_table={'height': '300px', 'overflowY': 'auto'},
                    style_cell={
                        'minWidth': 95,
                        'maxWidth': 200,
                        'width': 95,
                        'textAlign': 'left'
                    },
                    style_cell_conditional=[
                        {
                            'if': {'column_id': 'Notes'},
                            'width': '200px',
                            'textAlign': 'left'
                        }
                    ],
                    style_data_conditional=[
                        {
                            'if': {'row_index': 'odd'},
                            'backgroundColor': 'rgb(248, 248, 248)'
                        }
                    ]
                )
            ], style={'marginBottom': '30px', 'padding': '20px', 'backgroundColor': '#ffffff', 'borderRadius': '5px', 'boxShadow': '0 2px 4px rgba(0,0,0,0.1)'}),

            # P4 Section
            html.Div([
                html.H2('P4 Analysis'),
                html.Div([
                    html.Div([
                        html.Label('Select P4 Worksheets:', style={'marginBottom': '5px'}),
                        dcc.Dropdown(
                            id='p4-worksheet-selector',
                            options=[
                                {'label': name, 'value': name} 
                                for name in sheet_data.keys() 
                                if 'P4' in name
                            ],
                            value=[],
                            multi=True,
                            placeholder="Select P4 worksheets (leave empty for all)",
                            style={'width': '500px'}
                        ),
                    ], style={'flex': '1'}),
                    html.Div([
                        html.Label('Filename: '),
                        dcc.Input(
                            id='p4-filename-input',
                            type='text',
                            placeholder='Enter filename (without .xlsx)',
                            value='mnl_py_results',
                            style={'width': '200px', 'marginRight': '10px'}
                        ),
                        html.Button(
                            'Save P4 Results',
                            id='save-p4-button',
                            n_clicks=0,
                            style={'marginLeft': '10px', 'height': '40px'}
                        ),
                    ], style={'display': 'flex', 'alignItems': 'center'}),
                    html.Div(id='p4-save-status')
                ], style={'display': 'flex', 'alignItems': 'flex-end', 'marginBottom': '10px'}),
                
                html.Div([
                    html.Button(
                        'Select All', 
                        id='p4-select-all', 
                        n_clicks=0,
                        style={
                            'marginRight': '10px',
                            'padding': '5px 10px',
                            'backgroundColor': '#f8f9fa',
                            'border': '1px solid #ddd',
                            'borderRadius': '4px',
                            'cursor': 'pointer'
                        }
                    ),
                    html.Button(
                        'Deselect All', 
                        id='p4-deselect-all', 
                        n_clicks=0,
                        style={
                            'padding': '5px 10px',
                            'backgroundColor': '#f8f9fa',
                            'border': '1px solid #ddd',
                            'borderRadius': '4px',
                            'cursor': 'pointer'
                        }
                    ),
                    dcc.ConfirmDialog(
                        id='p4-confirm-deselect',
                        message='Are you sure you want to deselect all rows?',
                    ),
                ], style={'marginBottom': '10px'}),
                
                dcc.Graph(id='p4-graph'),
                html.H3('P4 Inflection Points', style={'marginTop': '20px', 'marginBottom': '10px'}),
                
                # P4 Table with VAS column
                dash_table.DataTable(
                    id='p4-table',
                    columns=[
                        {"name": "Include", "id": "selected", "type": "any"},
                        {"name": "Offset", "id": "Offset", "type": "numeric", "format": {"specifier": ".3f"}},
                        {"name": "Value", "id": "Value", "type": "numeric", "format": {"specifier": ".2f"}},
                        {"name": "Type", "id": "Type"},
                        {"name": "Annotation", "id": "Annotation"},
                        {"name": "Event", "id": "Event"},
                        {"name": "sheet_name", "id": "sheet_name"},
                        {"name": "mbase", "id": "mbase", "type": "numeric", "format": {"specifier": ".2f"}},
                        {"name": "mbase time", "id": "mbase_time", "type": "numeric", "format": {"specifier": ".0f"}},
                        {"name": "mpeak", "id": "mpeak", "type": "numeric", "format": {"specifier": ".2f"}},
                        {"name": "mpeak time", "id": "mpeak_time", "type": "numeric", "format": {"specifier": ".0f"}},
                        {"name": "pos_delta", "id": "pos_delta", "type": "numeric", "format": {"specifier": ".2f"}},
                        {"name": "neg_delta", "id": "neg_delta", "type": "numeric", "format": {"specifier": ".2f"}},
                        {"name": "VAS", "id": "VAS", "type": "numeric", "format": {"specifier": ".1f"}},
                        {"name": "Notes", "id": "Notes", "presentation": "input"}
                    ],
                    data=[],
                    persistence=True,
                    persistence_type='session',
                    editable=True,
                    row_selectable='multi',
                    selected_rows=[],
                    style_table={'height': '300px', 'overflowY': 'auto'},
                    style_cell={
                        'minWidth': 95,
                        'maxWidth': 200,
                        'width': 95,
                        'textAlign': 'left'
                    },
                    style_cell_conditional=[
                        {
                            'if': {'column_id': 'Notes'},
                            'width': '200px',
                            'textAlign': 'left'
                        },
                        {
                            'if': {'column_id': 'selected'},
                            'textAlign': 'center'
                        }
                    ],
                    style_data_conditional=[
                        {
                            'if': {'row_index': 'odd'},
                            'backgroundColor': 'rgb(248, 248, 248)'
                        },
                        {
                            'if': {'column_id': 'pos_delta', 'filter_query': '{pos_delta} > 0'},
                            'backgroundColor': '#e6ffe6'
                        },
                        {
                            'if': {'column_id': 'neg_delta', 'filter_query': '{neg_delta} < 0'},
                            'backgroundColor': '#ffe6e6'
                        }
                    ]
                ),
                
                html.H3('P4 Event Statistics', style={'marginTop': '20px', 'marginBottom': '10px'}),
                dash_table.DataTable(
                    id='p4-stats-table',
                    columns=[
                        {"name": "sheet_name", "id": "sheet_name"},
                        {"name": "Event", "id": "Event"},
                        {"name": "Start", "id": "Start", "type": "numeric", "format": {"specifier": ".3f"}},
                        {"name": "End", "id": "End", "type": "numeric", "format": {"specifier": ".3f"}},
                        {"name": "Avg", "id": "Avg", "type": "numeric", "format": {"specifier": ".3f"}},
                        {"name": "Std_Dev", "id": "Std_Dev", "type": "numeric", "format": {"specifier": ".3f"}},
                        {"name": "Max", "id": "Max", "type": "numeric", "format": {"specifier": ".3f"}},
                        {"name": "Min", "id": "Min", "type": "numeric", "format": {"specifier": ".3f"}},
                        {"name": "Delta", "id": "Delta", "type": "numeric", "format": {"specifier": ".3f"}},
                        {"name": "Median", "id": "Median", "type": "numeric", "format": {"specifier": ".3f"}},
                        {"name": "Q1-25%", "id": "Q1-25%", "type": "numeric", "format": {"specifier": ".3f"}},
                        {"name": "Q3-75%", "id": "Q3-75%", "type": "numeric", "format": {"specifier": ".3f"}},
                        {"name": "Q1-Q3", "id": "Q1-Q3", "type": "numeric", "format": {"specifier": ".3f"}},
                        {"name": "Diff", "id": "Diff", "type": "numeric", "format": {"specifier": ".3f"}},
                        {"name": "Notes", "id": "Notes", "presentation": "input"}
                    ],
                    data=[],
                    editable=True,
                    style_table={'height': '300px', 'overflowY': 'auto'},
                    style_cell={
                        'minWidth': 95,
                        'maxWidth': 200,
                        'width': 95,
                        'textAlign': 'left'
                    },
                    style_cell_conditional=[
                        {
                            'if': {'column_id': 'Notes'},
                            'width': '200px',
                            'textAlign': 'left'
                        }
                    ],
                    style_data_conditional=[
                        {
                            'if': {'row_index': 'odd'},
                            'backgroundColor': 'rgb(248, 248, 248)'
                        }
                    ]
                )
            ], style={'marginBottom': '30px', 'padding': '20px', 'backgroundColor': '#ffffff', 'borderRadius': '5px', 'boxShadow': '0 2px 4px rgba(0,0,0,0.1)'}),
                    
            # Phase Overlay Analysis Section
            html.Div([
                html.H2('Phase Overlay Analysis', style={'marginBottom': '20px'}),
            
                html.Div([
                    html.Button(
                        'Toggle Action Labels',
                        id='overlay-action-toggle',
                        n_clicks=0,
                        style={
                            'marginRight': '10px',
                            'padding': '5px 10px',
                            'backgroundColor': '#f8f9fa',
                            'border': '1px solid #ddd',
                            'borderRadius': '4px',
                            'cursor': 'pointer'
                        }
                    ),
                    html.Div(
                        "Action Labels OFF",
                        id='overlay-action-toggle-status',
                        style={'marginLeft': '10px', 'color': 'red'}
                    )
                ], style={'marginBottom': '20px', 'display': 'flex', 'alignItems': 'center'}),
                # Add Word and Event Filters section         
                html.Div([
                # Word Filters
                html.Div([
                    html.Div([
                        html.Label('Include Words (comma-separated):'),
                        dcc.Input(
                            id='overlay-include-words', 
                            type='text', 
                            value=','.join(DEFAULT_INCLUDE_WORDS),
                            style={'width': '200px'}
                        ),
                    ], style={'marginRight': '20px'}),
                    
                    html.Div([
                        html.Label('Exclude Words (comma-separated):'),
                        dcc.Input(
                            id='overlay-exclude-words', 
                            type='text', 
                            value=','.join(DEFAULT_EXCLUDE_WORDS),
                            style={'width': '200px'}
                        ),
                    ])
                ], style={'display': 'flex', 'marginBottom': '10px'}),
                
                # Event Filters
                html.Div([
                    html.Div([
                        html.Label('Include Events (comma-separated):'),
                        dcc.Input(
                            id='overlay-include-events', 
                            type='text', 
                            value=','.join(DEFAULT_INCLUDE_EVENTS),
                            style={'width': '200px'}
                        ),
                    ], style={'marginRight': '20px'}),
                    
                    html.Div([
                        html.Label('Exclude Events (comma-separated):'),
                        dcc.Input(
                            id='overlay-exclude-events', 
                            type='text', 
                            value=','.join(DEFAULT_EXCLUDE_EVENTS),
                            style={'width': '200px'}
                        ),
                    ])
                ], style={'display': 'flex', 'marginBottom': '10px'})
            ], style={'padding': '10px', 'backgroundColor': '#f8f9fa', 'borderRadius': '5px', 'marginBottom': '20px'}),
                
                # P1 Overlay Selection
                html.Div([
                    html.H3('P1 Overlay Selection'),
                    dcc.Dropdown(
                        id='p1-overlay-selector',
                        options=[
                            {'label': name, 'value': name} 
                            for name in sheet_data.keys() 
                            if 'P1' in name
                        ],
                        value=[],
                        multi=True,
                        placeholder="Select P1 files for overlay",
                        style={'width': '100%', 'marginBottom': '20px'}
                    ),
                ]),
            
                # P3 Overlay Selection
                html.Div([
                    html.H3('P3 Overlay Selection'),
                    dcc.Dropdown(
                        id='p3-overlay-selector',
                        options=[
                            {'label': name, 'value': name} 
                            for name in sheet_data.keys() 
                            if 'P3' in name
                        ],
                        value=[],
                        multi=True,
                        placeholder="Select P3 files for overlay",
                        style={'width': '100%', 'marginBottom': '20px'}
                    ),
                ]),
            
                # P4 Overlay Selection
                html.Div([
                    html.H3('P4 Overlay Selection'),
                    dcc.Dropdown(
                        id='p4-overlay-selector',
                        options=[
                            {'label': name, 'value': name} 
                            for name in sheet_data.keys() 
                            if 'P4' in name
                        ],
                        value=[],
                        multi=True,
                        placeholder="Select P4 files for overlay",
                        style={'width': '100%', 'marginBottom': '20px'}
                    ),
                ]),
            
                # Overlay Graph
                dcc.Graph(id='overlay-graph'),
            
                # Save Controls
                html.Div([
                    html.Button(
                        'Save Overlay Results',
                        id='save-overlay-button',
                        n_clicks=0,
                        style={'marginTop': '20px'}
                    ),
                    html.Div(id='overlay-save-status')
                ]), 
            ], style={'marginBottom': '30px', 'padding': '20px', 'backgroundColor': '#ffffff', 'borderRadius': '5px', 'boxShadow': '0 2px 4px rgba(0,0,0,0.1)'})
        ]),
        # Results Analysis Tab
        dcc.Tab(label='Results Analysis', children=[
            html.Div([
                html.H1('Results Analysis'),
                
                # Worksheet selector with improved styling
                html.Div([
                    html.Label('Select Results Worksheets:', style={'marginBottom': '5px'}),
                    html.Div([
                        dcc.Dropdown(
                            id='worksheet-results-selector',
                            options=[
#                                 {'label': name, 'value': name} 
#                                 for name in sheet_data.keys() 
#                                 if '' in name
                            ],
                            value=[],
                            multi=True,
                            placeholder="Select worksheets or use Select All button",
                            style={'width': '500px'}
                        ),
                        html.Button(
                            'Select All',
                            id='select-all-worksheets',
                            n_clicks=0,
                            style={
                                'marginLeft': '10px',
                                'height': '36px',
                                'verticalAlign': 'middle',
                                'backgroundColor': '#f8f9fa',
                                'border': '1px solid #ddd',
                                'borderRadius': '4px',
                                'cursor': 'pointer'
                            }
                        ),
                        html.Button(
                            'Clear All',
                            id='clear-all-worksheets',
                            n_clicks=0,
                            style={
                                'marginLeft': '10px',
                                'height': '36px',
                                'verticalAlign': 'middle',
                                'backgroundColor': '#f8f9fa',
                                'border': '1px solid #ddd',
                                'borderRadius': '4px',
                                'cursor': 'pointer'
                            }
                        )
                    ], style={'display': 'flex', 'alignItems': 'center'})
                ], style={'marginBottom': '20px'}),
                
                # Phase-specific analysis sections
                html.Div([
                    dcc.Tabs([
                        dcc.Tab(label='P1 Analysis', children=[
                            html.Div([
                                html.H3('P1 Overall Statistics'),
                                dcc.Graph(id='p1-overall-stats-graph'),
                                html.H3('P1 Delta Analysis'),
                                dcc.Graph(id='p1-delta-stats-graph'),
                                html.H3('P1 Peak Counts'),
                                dcc.Graph(id='p1-peak-counts-graph'),
                                html.H3('P1 Statistical Analysis'),
                                dash_table.DataTable(
                                    id='p1-stats-analysis',
                                    columns=[
                                        {"name": "Comparison", "id": "comparison"},
                                        {"name": "Mean", "id": "mean", "type": "numeric", "format": {"specifier": ".3f"}},
                                        {"name": "Std Dev", "id": "std", "type": "numeric", "format": {"specifier": ".3f"}},
                                        {"name": "Min", "id": "min", "type": "numeric", "format": {"specifier": ".3f"}},
                                        {"name": "Max", "id": "max", "type": "numeric", "format": {"specifier": ".3f"}},
                                        {"name": "T-Statistic", "id": "t_statistic", "type": "numeric", "format": {"specifier": ".3f"}},
                                        {"name": "P-Value", "id": "p_value", "type": "numeric", "format": {"specifier": ".3f"}}
                                    ],
                                    data=[],
                                    style_table={'overflowX': 'auto'}
                                ),
                                # Event Analysis Section
                                html.H3('P1 Event Analysis'),
                                dcc.Graph(id='p1-event-avg-graph'),
                                dcc.Graph(id='p1-event-delta-graph')
                            ])
                        ]),
                        
                        dcc.Tab(label='P3 Analysis', children=[
                            html.Div([
                                html.H3('P3 Overall Statistics'),
                                dcc.Graph(id='p3-overall-stats-graph'),
                                html.H3('P3 Delta Analysis'),
                                dcc.Graph(id='p3-delta-stats-graph'),
                                html.H3('P3 Peak Counts'),
                                dcc.Graph(id='p3-peak-counts-graph'),
                                html.H3('P3 Statistical Analysis'),
                                dash_table.DataTable(
                                    id='p3-stats-analysis',
                                    columns=[
                                        {"name": "Comparison", "id": "comparison"},
                                        {"name": "Mean", "id": "mean", "type": "numeric", "format": {"specifier": ".3f"}},
                                        {"name": "Std Dev", "id": "std", "type": "numeric", "format": {"specifier": ".3f"}},
                                        {"name": "Min", "id": "min", "type": "numeric", "format": {"specifier": ".3f"}},
                                        {"name": "Max", "id": "max", "type": "numeric", "format": {"specifier": ".3f"}},
                                        {"name": "T-Statistic", "id": "t_statistic", "type": "numeric", "format": {"specifier": ".3f"}},
                                        {"name": "P-Value", "id": "p_value", "type": "numeric", "format": {"specifier": ".3f"}}
                                    ],
                                    data=[],
                                    style_table={'overflowX': 'auto'}
                                ),
                                # Event Analysis Section
                                html.H3('P3 Event Analysis'),
                                dcc.Graph(id='p3-event-avg-graph'),
                                dcc.Graph(id='p3-event-delta-graph')
                            ])
                        ]),
                        
                        dcc.Tab(label='P4 Analysis', children=[
                            html.Div([
                                html.H3('P4 Overall Statistics'),
                                dcc.Graph(id='p4-overall-stats-graph'),
                                html.H3('P4 Delta Analysis'),
                                dcc.Graph(id='p4-delta-stats-graph'),
                                html.H3('P4 Peak Counts'),
                                dcc.Graph(id='p4-peak-counts-graph'),
                                html.H3('P4 Statistical Analysis'),
                                dash_table.DataTable(
                                    id='p4-stats-analysis',
                                    columns=[
                                        {"name": "Comparison", "id": "comparison"},
                                        {"name": "Mean", "id": "mean", "type": "numeric", "format": {"specifier": ".3f"}},
                                        {"name": "Std Dev", "id": "std", "type": "numeric", "format": {"specifier": ".3f"}},
                                        {"name": "Min", "id": "min", "type": "numeric", "format": {"specifier": ".3f"}},
                                        {"name": "Max", "id": "max", "type": "numeric", "format": {"specifier": ".3f"}},
                                        {"name": "T-Statistic", "id": "t_statistic", "type": "numeric", "format": {"specifier": ".3f"}},
                                        {"name": "P-Value", "id": "p_value", "type": "numeric", "format": {"specifier": ".3f"}}
                                    ],
                                    data=[],
                                    style_table={'overflowX': 'auto'}
                                ),
                                # Event Analysis Section
                                html.H3('P4 Event Analysis'),
                                dcc.Graph(id='p4-event-avg-graph'),
                                dcc.Graph(id='p4-event-delta-graph')
                            ])
                        ]),
                        
                        dcc.Tab(label='Cross-Phase Analysis', children=[
                                                        
                            html.Div([
                                html.H3('Phase Comparison'),
                                dcc.Dropdown(
                                    id='cross-phase-metric',
                                    options=[
                                        {'label': 'Overall Average', 'value': 'overall_avg'},
                                        {'label': 'Positive Deltas', 'value': 'avg_pos_delta'},
                                        {'label': 'Negative Deltas', 'value': 'avg_neg_delta'},
                                        {'label': 'Peak Counts', 'value': 'peak_count'}
                                    ],
                                    value='overall_avg'
                                ),
                                dcc.Graph(id='cross-phase-comparison'),
                                html.H3('Statistical Analysis'),
                                dash_table.DataTable(
                                    id='cross-phase-stats',
                                    columns=[
                                        {"name": i, "id": i} for i in ['comparison', 't_statistic', 'p_value']
                                    ],
                                    data=[]
                                )
                            ]) 
                        ])
                    ])
                ])
            ])
        ])        
    ])
])

# data validation and error checking
def filter_selected_points(events_data, table_data):
    """Filter event statistics to only include events with selected points"""
    if not table_data:
        return events_data
    selected_points = pd.DataFrame(table_data)
    if 'selected' not in selected_points.columns:
        return events_data
    # Only include events that have selected points
    selected_events = selected_points[selected_points['selected']]['Event'].unique()
    return [event for event in events_data if event['Event'] in selected_events]

def create_default_outputs():
    """Create default outputs for error cases"""
    empty_fig = go.Figure()
    empty_data = []
    return (
        empty_fig,  # p1_graph
        empty_data,  # p1_table
        [],  # p1_selected_rows
        empty_data,  # p1_stats
        empty_fig,  # p3_graph
        empty_data,  # p3_table
        [],  # p3_selected_rows
        empty_data,  # p3_stats
        empty_fig,  # p4_graph
        empty_data,  # p4_table
        [],  # p4_selected_rows
        empty_data,  # p4_stats
        "Action Labels OFF",  # action_toggle_status
        False,  # p1_confirm
        False,  # p3_confirm
        False   # p4_confirm
    )

@app.callback(
    [Output('p1-graph', 'figure'),
     Output('p1-table', 'data'),
     Output('p1-table', 'selected_rows'),
     Output('p1-stats-table', 'data'),
     Output('p3-graph', 'figure'),
     Output('p3-table', 'data'),
     Output('p3-table', 'selected_rows'),
     Output('p3-stats-table', 'data'),
     Output('p4-graph', 'figure'),
     Output('p4-table', 'data'),
     Output('p4-table', 'selected_rows'),
     Output('p4-stats-table', 'data'),
     Output('action-toggle-status', 'children'),
     Output('p1-confirm-deselect', 'displayed'),
     Output('p3-confirm-deselect', 'displayed'),     
     Output('p4-confirm-deselect', 'displayed')],
    [Input('include-words', 'value'),
     Input('exclude-words', 'value'),
     Input('include-events', 'value'),
     Input('exclude-events', 'value'),
     Input('p1-worksheet-selector', 'value'),
     Input('p3-worksheet-selector', 'value'),
     Input('p4-worksheet-selector', 'value'),
     Input('action-toggle', 'n_clicks'),
     Input('window-size-input', 'value'),
     Input('p1-graph', 'relayoutData'),
     Input('p3-graph', 'relayoutData'),
     Input('p4-graph', 'relayoutData'),
     Input('p1-table', 'selected_rows'),
     Input('p3-table', 'selected_rows'),
     Input('p4-table', 'selected_rows'),
     Input('p1-table', 'data'),
     Input('p3-table', 'data'),
     Input('p4-table', 'data'),
     Input('p1-select-all', 'n_clicks'),
     Input('p1-deselect-all', 'n_clicks'),
     Input('p3-select-all', 'n_clicks'),
     Input('p3-deselect-all', 'n_clicks'),
     Input('p4-select-all', 'n_clicks'),
     Input('p4-deselect-all', 'n_clicks'),
     Input('p1-confirm-deselect', 'submit_n_clicks'),
     Input('p3-confirm-deselect', 'submit_n_clicks'),
     Input('p4-confirm-deselect', 'submit_n_clicks')],
    [State('p1-table', 'data'),
     State('p3-table', 'data'),
     State('p4-table', 'data')],
    prevent_initial_call=True
)

def update_graphs_and_tables(include_words, exclude_words, include_events, exclude_events, p1_sheets, p3_sheets, p4_sheets, n_clicks, window_size,
                           p1_relayout, p3_relayout, p4_relayout, 
                           p1_selected_rows, p3_selected_rows, p4_selected_rows,
                           p1_table_data, p3_table_data, p4_table_data,
                           p1_select_all, p1_deselect_all,
                           p3_select_all, p3_deselect_all,
                           p4_select_all, p4_deselect_all,
                           p1_confirm, p3_confirm, p4_confirm,
                           p1_table_state, p3_table_state, p4_table_state):
    """Update all graphs and tables based on user inputs and parameters"""
    ctx = dash.callback_context
    if not ctx.triggered:
        raise PreventUpdate
        
    trigger_id = ctx.triggered[0]['prop_id'].split('.')[0]

    try:
        # Initialize action toggle state based on n_clicks
        show_actions = bool(n_clicks % 2)  # Changed this line
        
        # Create status displays
        toggle_status = html.Div(
            "Action Labels ON" if show_actions else "Action Labels OFF",
            style={'color': 'green' if show_actions else 'red'}
        )
        
        # Initialize dictionaries
        results = {}
        stats_data = {}
        
        # Get all selected sheets
        all_selected_sheets = []
        if p1_sheets:
            all_selected_sheets.extend(p1_sheets)
        if p3_sheets:
            all_selected_sheets.extend(p3_sheets)
        if p4_sheets:
            all_selected_sheets.extend(p4_sheets)

        if not all_selected_sheets:
            raise PreventUpdate
              
        # Process filters
        include_list = [word.strip() for word in include_words.split(',')] if include_words else DEFAULT_INCLUDE_WORDS
        exclude_list = [word.strip() for word in exclude_words.split(',')] if exclude_words else DEFAULT_EXCLUDE_WORDS
        include_event = [word.strip() for word in include_events.split(',')] if include_events else DEFAULT_INCLUDE_EVENTS
        exclude_event = [word.strip() for word in exclude_events.split(',')] if exclude_events else DEFAULT_EXCLUDE_EVENTS

        # Initialize confirmation flags
        show_p1_confirm = False
        show_p3_confirm = False
        show_p4_confirm = False
        
        # Handle select/deselect button clicks
        ctx = dash.callback_context
        if ctx.triggered:
            button_id = ctx.triggered[0]['prop_id'].split('.')[0]
            
            # Handle confirmation dialog submissions
            if button_id == 'p1-confirm-deselect' and p1_confirm:
                p1_selected_rows = []
                show_p1_confirm = False
            elif button_id == 'p3-confirm-deselect' and p3_confirm:
                p3_selected_rows = []
                show_p3_confirm = False
            elif button_id == 'p4-confirm-deselect' and p4_confirm:
                p4_selected_rows = []
                show_p4_confirm = False
            # Handle initial deselect button clicks
            elif button_id == 'p1-deselect-all':
                show_p1_confirm = True
            elif button_id == 'p3-deselect-all':
                show_p3_confirm = True
            elif button_id == 'p4-deselect-all':
                show_p4_confirm = True
            # Handle select all button clicks
            elif button_id == 'p1-select-all' and p1_table_state:
                p1_selected_rows = list(range(len(p1_table_state)))
            elif button_id == 'p3-select-all' and p3_table_state:
                p3_selected_rows = list(range(len(p3_table_state)))
            elif button_id == 'p4-select-all' and p4_table_state:
                p4_selected_rows = list(range(len(p4_table_state)))

        # Process visible ranges
        visible_ranges = {}
        for phase, relayout_data in [('P1', p1_relayout), ('P3', p3_relayout), ('P4', p4_relayout)]:
            if relayout_data and 'xaxis.range[0]' in relayout_data:
                visible_ranges[phase] = (
                    relayout_data['xaxis.range[0]'],
                    relayout_data['xaxis.range[1]']
                )
            
        def get_combined_y_range(results_dict):
            all_y_values = []
            for result in results_dict.values():
                if isinstance(result.get('table'), pd.DataFrame) and not result['table'].empty:
                    all_y_values.extend(result['table']['Value'].dropna().tolist())

            if all_y_values:
                min_y = min(all_y_values)
                max_y = max(all_y_values)
                padding = (max_y - min_y) * 0.05
                return [min_y - padding, max_y + padding]
            return None
        
        # Process each phase
        for phase, phase_sheets, selected_rows, table_data, table_state in [
            ('P1', p1_sheets, p1_selected_rows, p1_table_data, p1_table_state),
            ('P3', p3_sheets, p3_selected_rows, p3_table_data, p3_table_state),
            ('P4', p4_sheets, p4_selected_rows, p4_table_data, p4_table_state)
        ]:
            if not phase_sheets:
                continue

            # Create graph and table
            fig, table = create_phase_graph_and_table(
                sheet_data, 
                measurement_data, 
                phase, 
                phase_sheets,
                include_list, 
                exclude_list,
                include_event, 
                exclude_event, 
                show_actions,
                int(window_size)
            )

            # Handle selection state and delta calculations
            if isinstance(table, pd.DataFrame) and not table.empty:
                if selected_rows is None or len(selected_rows) == 0:
                    selected_rows = list(range(len(table)))
                
                selected_rows = [i for i in selected_rows if i < len(table)]
                table['selected'] = False
                if selected_rows:
                    table.iloc[selected_rows, table.columns.get_loc('selected')] = True
                
                table = calculate_deltas_with_selection(table)
                table = calculate_manual_deltas_with_selection(table)
                
                # Update graph with selected points
                selected_points = table[table['selected']]
                
                # Keep all non-marker traces (lines, VAS bars, text)
                base_traces = []
                for trace in fig.data:
                    # Keep line traces for main data
                    if isinstance(trace, go.Scatter) and trace.mode == 'lines':
                        base_traces.append(trace)
                    # Keep VAS bar traces
                    elif isinstance(trace, go.Bar):
                        base_traces.append(trace)
                    # Keep text traces for annotations
                    elif isinstance(trace, go.Scatter) and trace.mode == 'text':
                        base_traces.append(trace)
                fig.data = tuple(base_traces)  # Update figure data
                
                # Add back selected points with appropriate styling
                for point_type in ['Rolling Base', 'Rolling Peak', 'Detected Base', 'Detected Peak', 'Manual Base', 'Manual Peak']:
                    points = selected_points[selected_points['Type'] == point_type]
                    if not points.empty:
                        marker_color = {
                            'Rolling Base': 'rgb(102, 255, 51)',
                            'Rolling Peak': 'rgb(0, 176, 240)',
                            'Detected Base': COLORS['detected_base'],
                            'Detected Peak': COLORS['detected_peak'],
                            'Manual Base': COLORS['mbase'],
                            'Manual Peak': COLORS['mpeak']
                        }.get(point_type)
                        
                        marker_symbol = 'diamond' if 'Manual' in point_type else 'circle'
                        marker_size = 8 if 'Manual' in point_type else 12
                        
                        fig.add_trace(
                            go.Scatter(
                                name=f'{phase} ({point_type})',
                                x=points['Offset'],
                                y=points['Value'],
                                mode='markers',
                                marker=dict(
                                    color=marker_color,
                                    size=marker_size,
                                    symbol=marker_symbol,
                                    line=dict(color='white', width=1)
                                ),
                                showlegend=True
                            ),
                            secondary_y=False  # Add points on primary y-axis
                        )

                results[phase] = {
                    'fig': fig,
                    'table': table,
                    'selected_rows': selected_rows if selected_rows is not None else list(range(len(table) if isinstance(table, pd.DataFrame) else 0))
                }

        # Calculate y-range for all graphs
        y_range = get_combined_y_range(results)
        if y_range:
            for phase, relayout_data in [('P1', p1_relayout), ('P3', p3_relayout), ('P4', p4_relayout)]:
                if phase in results:
                    if relayout_data and 'yaxis.range[0]' in relayout_data:
                        manual_y_range = [relayout_data['yaxis.range[0]'], relayout_data['yaxis.range[1]']]
                        results[phase]['fig'].update_layout(yaxis=dict(range=manual_y_range))
                    else:
                        results[phase]['fig'].update_layout(yaxis=dict(range=y_range))

        # Calculate event statistics
        p1_stats = calculate_event_statistics(sheet_data, p1_sheets) if p1_sheets else []
        p1_stats = filter_selected_points(p1_stats, p1_table_data)

        p3_stats = calculate_event_statistics(sheet_data, p3_sheets) if p3_sheets else []
        p3_stats = filter_selected_points(p3_stats, p3_table_data)

        p4_stats = calculate_event_statistics(sheet_data, p4_sheets) if p4_sheets else []
        p4_stats = filter_selected_points(p4_stats, p4_table_data)

        # Get results for each phase
        empty_figure = go.Figure()
        empty_data = []

        p1_result = results.get('P1', {'fig': empty_figure, 'table': pd.DataFrame(), 'selected_rows': []})
        p3_result = results.get('P3', {'fig': empty_figure, 'table': pd.DataFrame(), 'selected_rows': []})
        p4_result = results.get('P4', {'fig': empty_figure, 'table': pd.DataFrame(), 'selected_rows': []})
        
        # Return all components
        return (
            p1_result['fig'],
            p1_result['table'].to_dict('records') if not p1_result['table'].empty else empty_data,
            list(range(len(p1_result['table']))) if p1_result['table'].empty else p1_result['selected_rows'],
            p1_stats,
            p3_result['fig'],
            p3_result['table'].to_dict('records') if not p3_result['table'].empty else empty_data,
            list(range(len(p3_result['table']))) if p3_result['table'].empty else p3_result['selected_rows'],
            p3_stats,
            p4_result['fig'],
            p4_result['table'].to_dict('records') if not p4_result['table'].empty else empty_data,
            list(range(len(p4_result['table']))) if p4_result['table'].empty else p4_result['selected_rows'],
            p4_stats,
            toggle_status,
            show_p1_confirm,
            show_p3_confirm,
            show_p4_confirm
        )

    except Exception as e:
        print(f"Error in update_graphs_and_tables: {str(e)}")
        traceback.print_exc()
        return create_default_outputs()
    
@app.callback(
    [Output('overlay-graph', 'figure'),
     Output('overlay-action-toggle-status', 'children'),
     Output('overlay-action-toggle-status', 'style')],
    [Input('p1-overlay-selector', 'value'),
     Input('p3-overlay-selector', 'value'),
     Input('p4-overlay-selector', 'value'),
     Input('overlay-action-toggle', 'n_clicks'),
     Input('overlay-include-words', 'value'),
     Input('overlay-exclude-words', 'value'),
     Input('overlay-include-events', 'value'),
     Input('overlay-exclude-events', 'value')]
)
def update_overlay_graph(p1_overlay_sheets, p3_overlay_sheets, p4_overlay_sheets, n_clicks, 
                        include_words, exclude_words, include_events, exclude_events):
    """Update the overlay graph based on selected sheets"""
    try:
        show_actions = bool(n_clicks and n_clicks % 2)
        toggle_text = "Action Labels ON" if show_actions else "Action Labels OFF"
        toggle_style = {'color': 'green'} if show_actions else {'color': 'red'}
        
        print("\nUpdating overlay graph")
        print(f"Selected P1 sheets: {p1_overlay_sheets}")
        print(f"Selected P3 sheets: {p3_overlay_sheets}")
        print(f"Selected P4 sheets: {p4_overlay_sheets}")
        print(f"Include words: {include_words}")
        print(f"Exclude words: {exclude_words}")
        print(f"Include events: {include_events}")
        print(f"Exclude events: {exclude_events}")
        print(f"Action labels are {'ON' if show_actions else 'OFF'}")

        # Process filters - remove empty strings and whitespace
        include_words_list = [word.strip() for word in include_words.split(',') if word.strip()] if include_words else []
        exclude_words_list = [word.strip() for word in exclude_words.split(',') if word.strip()] if exclude_words else []
        include_events_list = [event.strip() for event in include_events.split(',') if event.strip()] if include_events else []
        exclude_events_list = [event.strip() for event in exclude_events.split(',') if event.strip()] if exclude_events else []

        # Create empty figure if no sheets selected
        if not any([p1_overlay_sheets, p3_overlay_sheets, p4_overlay_sheets]):
            print("No sheets selected, returning empty figure")
            empty_fig = go.Figure()
            empty_fig.update_layout(
                title='Phase Overlay Comparison (No Data Selected)',
                xaxis_title='Time (seconds from start)',
                yaxis_title='Value',
                showlegend=False,
                plot_bgcolor='white',
                paper_bgcolor='white'
            )
            return empty_fig, toggle_text, toggle_style

        # Create figure
        fig = go.Figure()
        
        # Color scheme for phases
        colors = {
            'P1': 'rgb(31, 119, 180)',  # Blue
            'P3': 'rgb(255, 127, 14)',  # Orange
            'P4': 'rgb(44, 160, 44)'    # Green
        }
        
        # Process each phase
        for phase_name, selected_sheets, color in [
            ('P1', p1_overlay_sheets, colors['P1']),
            ('P3', p3_overlay_sheets, colors['P3']),
            ('P4', p4_overlay_sheets, colors['P4'])
        ]:
            if not selected_sheets:
                continue
                
            # Process each sheet in the phase
            for sheet_name in selected_sheets:
                if sheet_name in sheet_data:
                    data = sheet_data[sheet_name]['data']
                    if not data.empty:
                        # Apply filters
                        filtered_data = data.copy()
                        rows_before = len(filtered_data)
                        
                        # Apply word filters if any are specified
                        if include_words_list or exclude_words_list:
                            filtered_data = filter_by_words(filtered_data, include_words_list, exclude_words_list)
                            print(f"\nAfter word filters for {sheet_name}:")
                            print(f"Rows before: {rows_before}")
                            print(f"Rows after: {len(filtered_data)}")
                            
                        # Apply event filters if any are specified
                        if include_events_list or exclude_events_list:
                            filtered_data = filter_by_events(filtered_data, include_events_list, exclude_events_list)
                            print(f"\nAfter event filters for {sheet_name}:")
                            print(f"Rows before: {rows_before}")
                            print(f"Rows after: {len(filtered_data)}")
                            
                        if filtered_data.empty:
                            print(f"No data remaining after filtering for {sheet_name}")
                            continue
                        
                        # Normalize time to start at 0
                        min_offset = filtered_data['Offset'].min()
                        normalized_time = (filtered_data['Offset'] - min_offset) / 1000  # Convert to seconds
                        
                        # Add main trace
                        fig.add_trace(
                            go.Scatter(
                                name=f'{phase_name} - {sheet_name}',
                                x=normalized_time,
                                y=filtered_data['Value'] * 10,  # Scale values
                                mode='lines',
                                line=dict(
                                    color=color,
                                    width=2
                                ),
                                opacity=0.7
                            )
                        )
                        
                        # Add action labels if enabled
                        if show_actions:
                            normalized_data = filtered_data.copy()
                            normalized_data['Offset'] = (normalized_data['Offset'] - min_offset) / 1000
                            
                            # Track used positions for annotation placement
                            used_positions = []

                            # Calculate event averages
                            event_stats = {}
                            event_starts = {}
                            for event_name, event_group in normalized_data.groupby('Event'):
                                if pd.notna(event_name) and event_name != '':
                                    avg_value = event_group['Value'].mean() * 10
                                    event_stats[event_name] = avg_value
                                    event_starts[event_name] = event_group['Offset'].min()

                            # Add annotations
                            non_empty_annotations = normalized_data[normalized_data['Annotations'].notna()]
                            if not non_empty_annotations.empty:
                                labeled_events = set()

                                # Sort annotations by x position to process them in order
                                sorted_annotations = non_empty_annotations.sort_values('Offset')
                                
                                for _, row in sorted_annotations.iterrows():
                                    event_name = row['Event']
                                    normalized_x = (row['Offset'] - min_offset) / 1000
                                    base_y = (row['Value']*10) + 0.75

                                    # Find an available y position that doesn't overlap
                                    adjusted_y = find_available_y_position(
                                        normalized_x, 
                                        base_y, 
                                        used_positions,
                                        min_gap=1.0  # Adjust this value to control spacing
                                    )
                                    
                                    # Add average if this is first occurrence of event
                                    if event_name in event_stats and event_name not in labeled_events and \
                                        row['Offset'] == event_starts[event_name]:
                                        avg_text = f"\nAvg: {event_stats[event_name]:.2f}"
                                        labeled_events.add(event_name)
                                    else:
                                        avg_text = ""
                                        
                                    fig.add_trace(
                                        go.Scatter(
                                            name=f'{phase_name} - {sheet_name} (Actions)',
                                            x=[row['Offset']],
                                            y=[adjusted_y],
                                            mode='text',
                                            text=[f"{row['Annotations']}{avg_text}"],
                                            textposition='top center',                                            
                                            textfont=dict(size=12, color=color),
                                            showlegend=False,
                                            hoverinfo='skip'
                                        )
                                    )

                                    # Record the used position
                                    used_positions.append((row['Offset'], adjusted_y))
        
        # Update layout
        fig.update_layout(
            title={
                'text': 'Phase Overlay Comparison (Time Normalized)',
                'x': 0.5,
                'xanchor': 'center'
            },
            xaxis_title='Time (seconds from start)',
            yaxis_title='Value',
            plot_bgcolor='white',
            paper_bgcolor='white',
            hovermode='closest',
            showlegend=True,
            legend=dict(
                yanchor="top",
                y=0.99,
                xanchor="left",
                x=1.05,
                bgcolor='rgba(255, 255, 255, 0.8)',
                bordercolor='black',
                borderwidth=1
            ),
            margin=dict(l=50, r=150, t=50, b=50),
            height=600
        )

        # Update axes
        fig.update_xaxes(
            showgrid=True,
            gridwidth=1,
            gridcolor='LightGray',
            zeroline=True,
            zerolinewidth=1,
            zerolinecolor='LightGray'
        )
        
        fig.update_yaxes(
            showgrid=True,
            gridwidth=1,
            gridcolor='LightGray',
            zeroline=True,
            zerolinewidth=1,
            zerolinecolor='LightGray'
        )
        
        return fig, toggle_text, toggle_style
        
    except Exception as e:
        print(f"Error in update_overlay_graph: {str(e)}")
        traceback.print_exc()
        empty_fig = go.Figure()
        empty_fig.update_layout(
            title='Error Creating Phase Overlay',
            xaxis_title='Time (seconds from start)',
            yaxis_title='Value',
            plot_bgcolor='white',
            paper_bgcolor='white'
        )
        return empty_fig, "Action Labels OFF", {'color': 'red'}

# Save button callbacks
@app.callback(
    Output('p1-save-status', 'children'),
    [Input('save-p1-button', 'n_clicks')],
    [State('p1-graph', 'figure'),
     State('p1-table', 'data'),
     State('p1-stats-table', 'data'),
     State('p1-filename-input', 'value')]
)
def save_p1_results(n_clicks, fig_data, table_data, event_data, filename):
    try:
        if n_clicks > 0:
            if not filename:
                filename = 'mnl_py_results'
            success, message = save_single_phase_to_excel('P1', fig_data, table_data, event_data, filename)
            return html.Div(message, style={'color': 'green' if success else 'red'})
        return ""
    except Exception as e:
        print(f"Error saving P1 results: {str(e)}")
        return html.Div(f"Error saving results: {str(e)}", style={'color': 'red'})

@app.callback(
    Output('p3-save-status', 'children'),
    [Input('save-p3-button', 'n_clicks')],
    [State('p3-graph', 'figure'),
     State('p3-table', 'data'),
     State('p3-stats-table', 'data'),
     State('p3-filename-input', 'value')]
)
def save_p3_results(n_clicks, fig_data, table_data, event_data, filename):
    try:
        if n_clicks > 0:
            if not filename:
                filename = 'mnl_py_results'
            success, message = save_single_phase_to_excel('P3', fig_data, table_data, event_data, filename)
            return html.Div(message, style={'color': 'green' if success else 'red'})
        return ""
    except Exception as e:
        print(f"Error saving P3 results: {str(e)}")
        return html.Div(f"Error saving results: {str(e)}", style={'color': 'red'})

@app.callback(
    Output('p4-save-status', 'children'),
    [Input('save-p4-button', 'n_clicks')],
    [State('p4-graph', 'figure'),
     State('p4-table', 'data'),
     State('p4-stats-table', 'data'),
     State('p4-filename-input', 'value')]
)
def save_p4_results(n_clicks, fig_data, table_data, event_data, filename):
    try:
        if n_clicks > 0:
            if not filename:
                filename = 'mnl_py_results'
            success, message = save_single_phase_to_excel('P4', fig_data, table_data, event_data, filename)
            return html.Div(message, style={'color': 'green' if success else 'red'})
        return ""
    except Exception as e:
        print(f"Error saving P4 results: {str(e)}")
        return html.Div(f"Error saving results: {str(e)}", style={'color': 'red'})

# Selection management callbacks
# @app.callback(
#     Output('p1-table', 'selected_rows'),
#     [Input('p1-select-all', 'n_clicks'),
#      Input('p1-deselect-all', 'n_clicks')],
#     [State('p1-table', 'data')]
#)
def manage_p1_selection(select_clicks, deselect_clicks, table_data):
    ctx = dash.callback_context
    if not ctx.triggered:
        return []
    button_id = ctx.triggered[0]['prop_id'].split('.')[0]
    
    if button_id == 'p1-select-all' and table_data:
        return list(range(len(table_data)))
    elif button_id == 'p1-deselect-all':
        return []
    return []

# @app.callback(
#     Output('p3-table', 'selected_rows'),
#     [Input('p3-select-all', 'n_clicks'),
#      Input('p3-deselect-all', 'n_clicks')],
#     [State('p3-table', 'data')]
#)
def manage_p3_selection(select_clicks, deselect_clicks, table_data):
    ctx = dash.callback_context
    if not ctx.triggered:
        return []
    button_id = ctx.triggered[0]['prop_id'].split('.')[0]
    
    if button_id == 'p3-select-all' and table_data:
        return list(range(len(table_data)))
    elif button_id == 'p3-deselect-all':
        return []
    return []

# @app.callback(
#     Output('p4-table', 'selected_rows'),
#     [Input('p4-select-all', 'n_clicks'),
#      Input('p4-deselect-all', 'n_clicks')],
#     [State('p4-table', 'data')]
#)
def manage_p4_selection(select_clicks, deselect_clicks, table_data):
    ctx = dash.callback_context
    if not ctx.triggered:
        return []
    button_id = ctx.triggered[0]['prop_id'].split('.')[0]
    
    if button_id == 'p4-select-all' and table_data:
        return list(range(len(table_data)))
    elif button_id == 'p4-deselect-all':
        return []
    return []

# Initialize data
print("\nInitializing Peak Detection Analysis...")
try:
    print("Loading measurement data...")
    valid_ids, measurement_data = read_measurement_data('Mnl_Analysis.xlsx')
    print(f"Successfully loaded measurement data with {len(valid_ids)} valid IDs")
    
    print("Loading phase data...")
    sheet_data = read_excel_sheets('All_Phase_Sheets.xlsx')
    print(f"Successfully loaded {len(sheet_data)} phase sheets")
    
except FileNotFoundError as e:
    print(f"Error loading files: {str(e)}")
    sheet_data = {}
    measurement_data = pd.DataFrame()
    print("No data loaded - please check file paths")
except Exception as e:
    print(f"Error during initialization: {str(e)}")
    traceback.print_exc()
    sheet_data = {}
    measurement_data = pd.DataFrame()
    print("Error loading data - check error messages above")

@app.callback(
    Output('overlay-save-status', 'children'),
    [Input('save-overlay-button', 'n_clicks')],
    [State('overlay-graph', 'figure'),
     State('p1-overlay-selector', 'value'),
     State('p3-overlay-selector', 'value'),
     State('p4-overlay-selector', 'value')]
)
def save_overlay_results(n_clicks, fig_data, p1_overlay_sheets, p3_overlay_sheets, p4_overlay_sheets):
    """Save overlay graph results to Excel"""
    try:
        if not n_clicks:
            return ""
            
        if not any([p1_overlay_sheets, p3_overlay_sheets, p4_overlay_sheets]):
            return html.Div("No data selected to save", style={'color': 'red'})
            
        filename = 'phase_overlay_results.xlsx'
        temp_image_dir = 'temp_images'
        
        # Create temp directory if it doesn't exist
        os.makedirs(temp_image_dir, exist_ok=True)
        
        # Use os.path.join for proper path handling
        img_path = os.path.join(temp_image_dir, "overlay_chart.png")
        
        try:
            # Convert the figure data to a plotly figure object
            fig = go.Figure(fig_data)
            
            # Save current graph state as image
            pio.write_image(
                fig,
                img_path,
                width=1200, 
                height=800,
                scale=2,
                engine='kaleido'
            )
            
            # Create Excel workbook
            wb = Workbook()
            if 'Sheet' in wb.sheetnames:
                wb.remove(wb.active)
                
            # Create sheets
            ws_chart = wb.create_sheet("Overlay_Chart")
            ws_data = wb.create_sheet("Overlay_Data")
            
            # Add image to chart sheet
            if os.path.exists(img_path) and os.path.getsize(img_path) > 0:
                try:
                    img = Image(img_path)
                    ws_chart.add_image(img, 'A1')
                except Exception as e:
                    print(f"Warning: Could not add image to worksheet: {str(e)}")
                    ws_chart.cell(row=1, column=1, value="Graph image could not be loaded")
            
            # Gather data for data sheet
            data_rows = []
            headers = ['Phase', 'Sheet Name', 'Offset', 'Value']
            
            # Process each phase's data
            for phase_name, selected_sheets in [
                ('P1', p1_overlay_sheets),
                ('P3', p3_overlay_sheets),
                ('P4', p4_overlay_sheets)
            ]:
                if not selected_sheets:
                    continue
                
                for sheet_name in selected_sheets:
                    if sheet_name in sheet_data:
                        df = sheet_data[sheet_name]['data']
                        if not df.empty:
                            # Normalize offset
                            min_offset = df['Offset'].min()
                            normalized_offset = (df['Offset'] - min_offset) / 1000
                            
                            for offset, value in zip(normalized_offset, df['Value'] * 10):
                                data_rows.append([
                                    phase_name,
                                    sheet_name,
                                    float(offset),
                                    float(value)
                                ])
            
            # Write headers
            for col, header in enumerate(headers, 1):
                ws_data.cell(row=1, column=col, value=header)
            
            # Write data
            for row_idx, row_data in enumerate(data_rows, 2):
                for col_idx, value in enumerate(row_data, 1):
                    ws_data.cell(row=row_idx, column=col_idx, value=value)
            
            # Save workbook
            wb.save(filename)
            return html.Div(f"Successfully saved overlay results to {filename}", style={'color': 'green'})
            
        except Exception as e:
            print(f"Error saving overlay: {str(e)}")
            traceback.print_exc()
            return html.Div(f"Error saving overlay: {str(e)}", style={'color': 'red'})
            
        finally:
            # Clean up temporary files
            try:
                if os.path.exists(img_path):
                    os.remove(img_path)
                if os.path.exists(temp_image_dir) and not os.listdir(temp_image_dir):
                    os.rmdir(temp_image_dir)
            except Exception as e:
                print(f"Error cleaning up temporary files: {str(e)}")
        
    except Exception as e:
        print(f"Error in save_overlay_results: {str(e)}")
        traceback.print_exc()
        return html.Div(f"Error: {str(e)}", style={'color': 'red'})

# Results Analysis Tab Callbacks
@app.callback(
    Output('worksheet-results-selector', 'options'),
    Input('worksheet-results-selector', 'value')
)
def update_worksheet_options(_):
    """Update the dropdown options with available results files"""
    try:
        files = get_results_files()
        files.sort()
        return [{'label': f, 'value': f} for f in files]
    except Exception as e:
        print(f"Error updating worksheet options: {str(e)}")
        return []

@app.callback(
    Output('worksheet-results-selector', 'value'),
    [Input('select-all-worksheets', 'n_clicks'),
     Input('clear-all-worksheets', 'n_clicks')],
    [State('worksheet-results-selector', 'options')]
)
def update_worksheet_selection(select_clicks, clear_clicks, options):
    """Handle Select All/Clear All button clicks"""
    try:
        ctx = dash.callback_context
        if not ctx.triggered:
            return []
        
        button_id = ctx.triggered[0]['prop_id'].split('.')[0]
        
        if button_id == 'select-all-worksheets':
            return [option['value'] for option in options] if options else []
        elif button_id == 'clear-all-worksheets':
            return []
        
        return []
    except Exception as e:
        print(f"Error in worksheet selection: {str(e)}")
        return []

# Phase analysis callbacks
@app.callback(
    [Output('p1-overall-stats-graph', 'figure'),
     Output('p1-delta-stats-graph', 'figure'),
     Output('p1-peak-counts-graph', 'figure'),
     Output('p1-stats-analysis', 'data'),
     Output('p3-overall-stats-graph', 'figure'),
     Output('p3-delta-stats-graph', 'figure'),
     Output('p3-peak-counts-graph', 'figure'),
     Output('p3-stats-analysis', 'data'),
     Output('p4-overall-stats-graph', 'figure'),
     Output('p4-delta-stats-graph', 'figure'),
     Output('p4-peak-counts-graph', 'figure'),
     Output('p4-stats-analysis', 'data'),
     Output('p1-event-avg-graph', 'figure'),
     Output('p3-event-avg-graph', 'figure'),
     Output('p4-event-avg-graph', 'figure'),
     Output('p1-event-delta-graph', 'figure'),
     Output('p3-event-delta-graph', 'figure'),
     Output('p4-event-delta-graph', 'figure')],
    [Input('worksheet-results-selector', 'value')]
)
def update_phase_analyses(selected_files):
    """Update all phase analyses based on selected files"""
    try:
        # Initialize empty defaults
        empty_fig = go.Figure()
        empty_fig.update_layout(
            title="No data available",
            xaxis_title="",
            yaxis_title="",
            showlegend=False
        )
        
        # Return empty state if no files selected
        if not selected_files:
            return [empty_fig] * 18  # Total of 18 outputs

        # Initialize containers for each phase
        phase_results = {
            '1': {'overall_fig': empty_fig, 'delta_fig': empty_fig, 'counts_fig': empty_fig, 'stats_data': []},
            '3': {'overall_fig': empty_fig, 'delta_fig': empty_fig, 'counts_fig': empty_fig, 'stats_data': []},
            '4': {'overall_fig': empty_fig, 'delta_fig': empty_fig, 'counts_fig': empty_fig, 'stats_data': []}
        }
        event_graphs = {
            '1': {'avg_fig': empty_fig, 'delta_fig': empty_fig},
            '3': {'avg_fig': empty_fig, 'delta_fig': empty_fig},
            '4': {'avg_fig': empty_fig, 'delta_fig': empty_fig}
        }
        
        # Load and process data for each phase
        phase_data = {}
        event_data = {}
        
        # First, load and process all data for each phase
        for phase in ['1', '3', '4']:
            try:
                # Process peak/base data
                peak_dfs = []
                event_dfs = []
                
                for file in selected_files:
                    try:
                        # Load peak/base data
                        df = load_peaks_data(file, phase)
                        if not df.empty:
                            stats = calculate_phase_statistics(df)
                            stats['number'], stats['phase'] = parse_filename(file)
                            # Round numerical values
                            for col in stats.index:
                                if isinstance(stats[col], (float, np.float64)):
                                    stats[col] = round(stats[col], 3)
                            peak_dfs.append(stats)
                        
                        # Load event data
                        event_df = load_and_process_events(file, phase)
                        if not event_df.empty:
                            event_dfs.append(event_df)
                            
                    except Exception as e:
                        print(f"Error processing file {file} for phase {phase}: {str(e)}")
                        continue
                
                # Create phase DataFrame
                phase_data[phase] = pd.DataFrame(peak_dfs) if peak_dfs else pd.DataFrame()
                event_data[phase] = pd.concat(event_dfs) if event_dfs else pd.DataFrame()
                
            except Exception as e:
                print(f"Error processing phase {phase}: {str(e)}")
                phase_data[phase] = pd.DataFrame()
                event_data[phase] = pd.DataFrame()
        
        # Process each phase
        for phase in ['1', '3', '4']:
            phase_df = phase_data.get(phase, pd.DataFrame())
            
            if not phase_df.empty:
                try:
                    # Create comparison graphs
                    overall_fig, delta_fig, counts_fig = create_comparison_graphs(phase_df, phase)
                    
                    # Update layouts
                    for fig in [overall_fig, delta_fig, counts_fig]:
                        fig.update_layout(
                            title_x=0.5,
                            margin=dict(t=50, b=50, l=50, r=50),
                            plot_bgcolor='white',
                            paper_bgcolor='white'
                        )
                    
                    # Store results for this phase
                    phase_results[phase].update({
                        'overall_fig': overall_fig,
                        'delta_fig': delta_fig,
                        'counts_fig': counts_fig
                    })

                    # Statistical analysis with proper error handling
                    try:
                        stats_analysis = perform_statistical_analysis(phase_df, 'overall_avg')
                        stats_data = []
                        if isinstance(stats_analysis, pd.DataFrame) and not stats_analysis.empty:
                            for _, row in stats_analysis.iterrows():
                                stats_entry = {
                                    "comparison": str(row["comparison"]),
                                    "mean": float(row["mean"]) if pd.notna(row["mean"]) else None,
                                    "std": float(row["std"]) if pd.notna(row["std"]) else None,
                                    "min": float(row["min"]) if pd.notna(row["min"]) else None,
                                    "max": float(row["max"]) if pd.notna(row["max"]) else None,
                                    "t_statistic": float(row["t_statistic"]) if pd.notna(row["t_statistic"]) else None,
                                    "p_value": float(row["p_value"]) if pd.notna(row["p_value"]) else None
                                }
                                stats_data.append(stats_entry)

                            # Ensure numeric values are properly converted
                            for row in stats_data:
                                for key, value in row.items():
                                    if pd.isna(value):
                                        row[key] = None
                                    elif isinstance(value, (float, np.float64)):
                                        row[key] = float(value)
                                    else:
                                        row[key] = str(value)
                            
                        phase_results[phase]['stats_data'] = stats_data
                    except Exception as e:
                        print(f"Error in statistical analysis for phase {phase}: {str(e)}")
                        phase_results[phase]['stats_data'] = []
                except Exception as e:
                    print(f"Error generating analysis for phase {phase}: {str(e)}")
                    
            # Process event data
            try:
                event_df = event_data.get(phase, pd.DataFrame())
                if not event_df.empty:
                    # Create event average graph
                    avg_fig = px.box(
                        event_df,
                        x='Event',
                        y='Avg',
                        color='number',
                        title=f'P{phase} Event Averages by Sample',
                        points='all'
                    )
                    avg_fig.update_layout(
                        xaxis_title='Event',
                        yaxis_title='Average Value',
                        showlegend=True,
                        boxmode='group',
                        title_x=0.5,
                        margin=dict(t=50, b=50, l=50, r=50),
                        plot_bgcolor='white',
                        paper_bgcolor='white'
                    )

                    # Create event delta graph
                    delta_fig = px.bar(
                        event_df,
                        x='Event',
                        y=['Delta', 'Diff'],
                        color='number',
                        barmode='group',
                        title=f'P{phase} Event Deltas and Differences'
                    )
                    delta_fig.update_layout(
                        xaxis_title='Event',
                        yaxis_title='Value',
                        showlegend=True,
                        title_x=0.5,
                        margin=dict(t=50, b=50, l=50, r=50),
                        plot_bgcolor='white',
                        paper_bgcolor='white'
                    )

                    event_graphs[phase].update({
                        'avg_fig': avg_fig,
                        'delta_fig': delta_fig
                    })
            except Exception as e:
                print(f"Error generating event analysis for phase {phase}: {str(e)}")

        # Return all results in the correct order
        return [
            # P1 analysis results
            phase_results['1']['overall_fig'],
            phase_results['1']['delta_fig'],
            phase_results['1']['counts_fig'],
            phase_results['1']['stats_data'],
            # P3 analysis results
            phase_results['3']['overall_fig'],
            phase_results['3']['delta_fig'],
            phase_results['3']['counts_fig'],
            phase_results['3']['stats_data'],
            # P4 analysis results
            phase_results['4']['overall_fig'],
            phase_results['4']['delta_fig'],
            phase_results['4']['counts_fig'],
            phase_results['4']['stats_data'],
            # Event analysis graphs
            event_graphs['1']['avg_fig'],
            event_graphs['3']['avg_fig'],
            event_graphs['4']['avg_fig'],
            event_graphs['1']['delta_fig'],
            event_graphs['3']['delta_fig'],
            event_graphs['4']['delta_fig']
        ]

    except Exception as e:
        print(f"Error in update_phase_analyses: {str(e)}")
        traceback.print_exc()
        return ([empty_fig] * 12) + ([[]] * 3) + ([empty_fig] * 3) 
    
@app.callback(
    [Output('cross-phase-comparison', 'figure'),
     Output('cross-phase-stats', 'data')],
    [Input('worksheet-results-selector', 'value'),
     Input('cross-phase-metric', 'value')]
)
def update_cross_phase_analysis(selected_files, metric):
    """Update cross-phase analysis based on selected files and metric"""
    if not selected_files or not metric:
        return go.Figure(), []
    
    try:
        # Combine data from all phases
        all_data = []
        for phase in ['1', '3', '4']:
            for file in selected_files:
                df = load_peaks_data(file, phase)
                if not df.empty:
                    stats = calculate_phase_statistics(df)
                    stats['number'], stats['phase'] = parse_filename(file)
                    all_data.append(stats)
        
        if not all_data:
            return go.Figure(), []
        
        combined_data = pd.DataFrame(all_data)
        
        # Create comparison figure
        fig = px.bar(
            combined_data,
            x='number',
            y=metric,
            color='phase',
            barmode='group',
            title=f'Phase Comparison - {metric}',
            labels={
                'number': 'Sample',
                metric: metric.replace('_', ' ').title(),
                'phase': 'Phase'
            }
        )
        
        fig.update_layout(
            xaxis_title='Sample',
            yaxis_title=metric.replace('_', ' ').title(),
            showlegend=True,
            legend_title='Phase'
        )
        
        # Add error bars
        for phase in combined_data['phase'].unique():
            phase_data = combined_data[combined_data['phase'] == phase]
            std = phase_data[metric].std()
            fig.add_traces([
                go.Bar(
                    name=f'{phase} Error',
                    x=phase_data['number'],
                    y=[std] * len(phase_data),
                    error_y=dict(type='data', array=[std] * len(phase_data)),
                    showlegend=False
                )
            ])
        
        # Statistical analysis
        stats_results = perform_statistical_analysis(combined_data, metric)
        stats_results = stats_results.round(3)
        
        # Convert to list of dictionaries for DataTable
        stats_data = []
        if not stats_results.empty:
            for _, row in stats_results.iterrows():
                stats_data.append({
                    'comparison': str(row['comparison']),
                    't_statistic': float(row['t_statistic']) if pd.notna(row['t_statistic']) else None,
                    'p_value': float(row['p_value']) if pd.notna(row['p_value']) else None
                })
        
        return fig, stats_data
        
    except Exception as e:
        print(f"Error in update_cross_phase_analysis: {str(e)}")
        traceback.print_exc()
        return go.Figure(), []

# Run the server
if __name__ == '__main__':
    app.run_server(debug=True, port= 3001)


# In[ ]:





# In[ ]:




