# %%
### MANUAL AND PYTHON BASES RWIN PEAK DETECTION WITH RESULTS
from dash import Dash, html, dcc, dash_table, ctx
from dash.dependencies import Input, Output, State, ALL
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

def detect_peaks_rolling_window(data, sheet_name):
    """
    Detect peaks and valleys using a rolling window comparison method with edge case handling.
    """
    peaks_data = []
    n = len(data)
    window_size = 5
    
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
    
DEFAULT_BASE_PARAMS = {
    'height': None,
    'prominence': 0.001,
    'distance': 1,
    'threshold': None,
    'width': None,
    'wlen': 5,
    'plateau_size': None
}

DEFAULT_PEAK_PARAMS = {
    'height': None,
    'prominence': 0.001,
    'distance': 1,
    'threshold': None,
    'width': None,
    'wlen': 5,
    'plateau_size': None
}

def get_dynamic_peak_params_multi(combined_stats, individual_stats, manual_params=None, sheet_name=None):
    """Generate peak detection parameters based on statistics with manual parameter override"""
    
    # Initialize with default parameters
    base_params = DEFAULT_BASE_PARAMS.copy()
    peak_params = DEFAULT_PEAK_PARAMS.copy()
    
    # Get appropriate stats
    stats = None
    if sheet_name and sheet_name in individual_stats:
        stats = individual_stats.get(sheet_name, None)
    else:
        stats = combined_stats

    if not stats:
        print("No valid stats found, returning default parameters.")
        return base_params, peak_params

    # Get statistics values with safe defaults
    mean = stats.get('mean', 0.0)
    std = stats.get('std', 0.0)
    range_val = stats.get('range', 0.0)

    print("\nCalculating dynamic parameters from statistics:")
    print(f"Mean: {mean:.3f}")
    print(f"Std: {std:.3f}")
    print(f"Range: {range_val:.3f}")

    # Calculate dynamic base parameters if no manual override
    if not manual_params or not manual_params.get('base', {}).get('height'):
        base_params['height'] = float(mean - 0.1 * std)
    if not manual_params or not manual_params.get('base', {}).get('prominence'):
        base_params['prominence'] = float(0.1 * range_val)
    if not manual_params or not manual_params.get('base', {}).get('distance'):
        base_params['distance'] = 20
    if not manual_params or not manual_params.get('base', {}).get('wlen'):
        base_params['wlen'] = 10
    if not manual_params or not manual_params.get('base', {}).get('threshold'):
        base_params['threshold'] = float(0.05 * std)
    if not manual_params or not manual_params.get('base', {}).get('width'):
        base_params['width'] = 5
    if not manual_params or not manual_params.get('base', {}).get('plateau_size'):
        base_params['plateau_size'] = None

    # Calculate dynamic peak parameters if no manual override
    if not manual_params or not manual_params.get('peak', {}).get('height'):
        peak_params['height'] = float(mean - 0.5 * std)
    if not manual_params or not manual_params.get('peak', {}).get('prominence'):
        peak_params['prominence'] = float(0.1 * range_val)
    if not manual_params or not manual_params.get('peak', {}).get('distance'):
        peak_params['distance'] = 20
    if not manual_params or not manual_params.get('peak', {}).get('wlen'):
        peak_params['wlen'] = 10
    if not manual_params or not manual_params.get('peak', {}).get('threshold'):
        peak_params['threshold'] = float(0.05 * std)
    if not manual_params or not manual_params.get('peak', {}).get('width'):
        peak_params['width'] = 5
    if not manual_params or not manual_params.get('peak', {}).get('plateau_size'):
        peak_params['plateau_size'] = None

    # Apply manual parameters if provided
    if manual_params:
        for param, value in manual_params.get('base', {}).items():
            if value is not None and value != 0:  # Only override if not None and not 0
                base_params[param] = value
                print(f"Using manual base {param}: {value}")

        for param, value in manual_params.get('peak', {}).items():
            if value is not None and value != 0:  # Only override if not None and not 0
                peak_params[param] = value
                print(f"Using manual peak {param}: {value}")

    print("\nFinal Parameters:")
    print("Base Parameters:", base_params)
    print("Peak Parameters:", peak_params)

    return base_params, peak_params

def create_base_params_table(base_params, individual_base_params):
    """Create a formatted DataTable to display base parameters with formulas"""
    if base_params is None:
        return dash_table.DataTable(
            id='base_params-table',
            columns=[],
            data=[]
        )
    
    # Define parameter display names and formulas
    param_info = {
        'height': {
            'name': 'Height (Relative to mean)',
            'base_formula': 'mean - 0.1 * std (if manual = 0)\nMore negative = more maxima peaks',
            'description': 'Baseline level for detecting maxima peaks'
        },
        'prominence': {
            'name': 'Prominence (% of range)',
            'base_formula': '0.1 * range (if manual = 0)\nSmaller = more maxima peaks',
            'description': 'Minimum height difference required between peaks'
        },
        'distance': {
            'name': 'Min Distance Between Peaks',
            'base_formula': '20 points (if manual = 0)\nSmaller = more maxima peaks',
            'description': 'Minimum number of points between detected peaks'
        },
        'wlen': {
            'name': 'Window Length',
            'base_formula': '10 points (if manual = 0)\nSmaller = more precise peak detection',
            'description': 'Size of window for peak detection'
        },
        'threshold': {
            'name': 'Detection Threshold',
            'base_formula': '0.05 * std (if manual = 0)\nSmaller = more maxima peaks',
            'description': 'Minimum threshold for peak detection'
        },
        'width': {
            'name': 'Min Peak Width',
            'base_formula': '5 points (if manual = 0)\nSmaller = more peaks detected',
            'description': 'Minimum width required for peak detection'
        },
        'plateau_size': {
            'name': 'Min Plateau Size',
            'base_formula': 'None (if manual = 0)',
            'description': 'Minimum size required for flat peaks'
        }
    }
    
    columns = [
        {'name': 'Parameter', 'id': 'Parameter'},
        {'name': 'Description', 'id': 'Description'},
        {'name': 'Dynamic Formula', 'id': 'Formula'}
    ]
    
    for sheet_name in individual_base_params.keys():
        columns.append({'name': sheet_name, 'id': sheet_name})
        
    base_params_keys = ['height', 'prominence', 'distance', 'wlen', 'threshold', 'width', 'plateau_size']
    table_data = []
    
    for base_param in base_params_keys:
        info = param_info.get(base_param, {})
        row = {
            'Parameter': base_param.upper(),
            'Description': f"{info.get('name', base_param)}\n{info.get('description', '')}",
            'Formula': info.get('base_formula', 'N/A')
        }
    
        for sheet_name, base_param_values in individual_base_params.items():
            value = base_param_values.get(base_param, 'N/A')
            if value is None:
                row[sheet_name] = 'None'
            elif value == 'N/A':
                row[sheet_name] = 'N/A'
            elif isinstance(value, (int, float)):
                row[sheet_name] = f"{value:.3f}"
            else:
                row[sheet_name] = str(value)
      
        table_data.append(row)
    
    return dash_table.DataTable(
        id='base_params-table',
        columns=columns,
        data=table_data,
        style_table={
            'overflowX': 'auto',
            'maxHeight': '400px',
            'overflowY': 'auto'
        },
        style_cell={
            'textAlign': 'left',
            'minWidth': '150px',
            'maxWidth': '300px',
            'whiteSpace': 'pre-line',
            'padding': '10px'
        },
        style_header={
            'backgroundColor': 'rgb(230, 230, 230)',
            'fontWeight': 'bold'
        },
        style_data_conditional=[
            {
                'if': {'column_id': 'Parameter'},
                'fontWeight': 'bold',
                'backgroundColor': 'rgb(248, 248, 248)',
                'width': '120px'
            },
            {
                'if': {'column_id': 'Description'},
                'backgroundColor': 'rgb(248, 248, 248)',
                'width': '250px'
            },
            {
                'if': {'column_id': 'Formula'},
                'backgroundColor': 'rgb(248, 248, 248)',
                'width': '300px'
            }
        ],
        tooltip_delay=0,
        tooltip_duration=None,
        style_cell_conditional=[
            {
                'if': {'column_id': 'Formula'},
                'textAlign': 'left',
                'whiteSpace': 'pre-line'
            }
        ]
    )

def create_peak_params_table(peak_params, individual_peak_params):
    """Create a formatted DataTable to display peak parameters with formulas"""
    if peak_params is None:
        return dash_table.DataTable(
            id='peak_params-table',
            columns=[],
            data=[]
        )
    
    # Define parameter display names and formulas
    param_info = {
        'height': {
            'name': 'Height (Relative to mean)',
            'base_formula': 'mean - 0.1 * std (if manual = 0)\nMore negative = more maxima peaks',
            'peak_formula': 'mean - 0.5 * std (if manual = 0)\nMore negative = more minima peaks'
        },
        'prominence': {
            'name': 'Prominence (% of range)',
            'base_formula': '0.1 * range (if manual = 0)\nSmaller = more maxima peaks',
            'peak_formula': '0.1 * range (if manual = 0)\nSmaller = more minima peaks'
        },
        'distance': {
            'name': 'Min Distance Between Peaks',
            'base_formula': '20 points (if manual = 0)\nSmaller = more maxima peaks',
            'peak_formula': '20 points (if manual = 0)\nSmaller = more minima peaks'
        },
        'wlen': {
            'name': 'Window Length',
            'base_formula': '10 points (if manual = 0)\nSmaller = more precise peak detection',
            'peak_formula': '10 points (if manual = 0)\nSmaller = more precise peak detection'
        },
        'threshold': {
            'name': 'Detection Threshold',
            'base_formula': '0.05 * std (if manual = 0)\nSmaller = more maxima peaks',
            'peak_formula': '0.05 * std (if manual = 0)\nSmaller = more minima peaks'
        },
        'width': {
            'name': 'Min Peak Width',
            'base_formula': '5 points (if manual = 0)\nSmaller = more peaks detected',
            'peak_formula': '5 points (if manual = 0)\nSmaller = more peaks detected'
        },
        'plateau_size': {
            'name': 'Min Plateau Size',
            'base_formula': 'None (if manual = 0)',
            'peak_formula': 'None (if manual = 0)'
        }
    }
    
    columns = [
        {'name': 'Parameter', 'id': 'Parameter'},
        {'name': 'Description', 'id': 'Description'},
        {'name': 'Dynamic Formula', 'id': 'Formula'}
    ]
    
    for sheet_name in individual_peak_params.keys():
        columns.append({'name': sheet_name, 'id': sheet_name})
        
    peak_params_keys = ['height', 'prominence', 'distance', 'wlen', 'threshold', 'width', 'plateau_size']
    table_data = []
    
    for peak_param in peak_params_keys:
        info = param_info.get(peak_param, {})
        row = {
            'Parameter': peak_param.upper(),
            'Description': info.get('name', peak_param),
            'Formula': f"Base: {info.get('base_formula', 'N/A')}\n\nPeak: {info.get('peak_formula', 'N/A')}"
        }
    
        for sheet_name, peak_param_values in individual_peak_params.items():
            value = peak_param_values.get(peak_param, 'N/A')
            if value is None:
                row[sheet_name] = 'None'
            elif value == 'N/A':
                row[sheet_name] = 'N/A'
            elif isinstance(value, (int, float)):
                row[sheet_name] = f"{value:.3f}"
            else:
                row[sheet_name] = str(value)
      
        table_data.append(row)
    
    return dash_table.DataTable(
        id='peak_params-table',
        columns=columns,
        data=table_data,
        style_table={
            'overflowX': 'auto',
            'maxHeight': '400px',
            'overflowY': 'auto'
        },
        style_cell={
            'textAlign': 'left',
            'minWidth': '150px',
            'maxWidth': '300px',
            'whiteSpace': 'pre-line',
            'padding': '10px'
        },
        style_header={
            'backgroundColor': 'rgb(230, 230, 230)',
            'fontWeight': 'bold'
        },
        style_data_conditional=[
            {
                'if': {'column_id': 'Parameter'},
                'fontWeight': 'bold',
                'backgroundColor': 'rgb(248, 248, 248)',
                'width': '120px'
            },
            {
                'if': {'column_id': 'Description'},
                'backgroundColor': 'rgb(248, 248, 248)',
                'width': '200px'
            },
            {
                'if': {'column_id': 'Formula'},
                'backgroundColor': 'rgb(248, 248, 248)',
                'width': '300px'
            }
        ]
    )

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

def read_excel_sheets(file_path):
    """Read worksheets from Excel file that contain offset, value, annotation, and event data"""
    sheet_data = {}
    with pd.ExcelFile(file_path) as xls:
        for sheet_name in xls.sheet_names:
            if any(phase in sheet_name for phase in ['P1', 'P3', 'P4']):
                try:
                    df = pd.read_excel(xls, sheet_name=sheet_name, usecols=[0, 2, 3, 4, 5], 
                                     names=['Offset', 'Value', 'Annotations', 'Annotation', 'Event'])
                    
                    actions = df['Annotations'].dropna().unique().tolist()
                    
                    df = df.dropna(subset=['Offset', 'Value', 'Annotation', 'Event'])
                    
                    sheet_data[sheet_name] = {'data': df, 'actions': actions}
                    #print(f"Loaded sheet: {sheet_name} with {len(df)} rows")
                except Exception as e:
                    print(f"Error reading sheet '{sheet_name}': {str(e)}")
    return sheet_data

def read_measurement_data(file_path):
    """Read manual analysis data from all worksheets"""
    #print(f"\nReading measurement data from {file_path}")
    data = []
    valid_ids = set()

    xlsx = pd.ExcelFile(file_path)
    for sheet_name in xlsx.sheet_names:
        #print(f"Processing sheet: {sheet_name}")
        
        df = pd.read_excel(xlsx, sheet_name=sheet_name, header=None)
        headers = df.iloc[0]
        #print(f"Headers found: {headers.tolist()}")
        
        for idx, row in df.iloc[1:].iterrows():
            id_value = str(row[0])
            if pd.isna(id_value):
                continue
                
            print(f"Processing ID: {id_value}")
            
            for col in range(1, len(row), 4):
                if pd.notna(row[col]):
                    try:
                        base_value = float(row[col])
                        peak_value = float(row[col + 1])
                        base_time = float(row[col + 2])
                        peak_time = float(row[col + 3])
                        
                        pos_delta = peak_value - base_value if peak_value > base_value else None
                        neg_delta = peak_value - base_value if peak_value < base_value else None
                        
                        base_name = str(headers[col]).strip()
                        peak_name = str(headers[col + 1]).strip()
                        
                        # Add base measurement
                        data.append({
                            'sheet_name': id_value,
                            'Type': 'Manual Base',
                            'Annotation': base_name,
                            'Offset': base_time,
                            'Value': base_value,
                            'mbase': base_value,
                            'mbase_time': base_time,
                            'mpeak': None,
                            'mpeak_time': None,
                            'pos_delta': pos_delta,
                            'neg_delta': neg_delta,
                            'Notes': ''
                        })
                        
                        # Add peak measurement
                        data.append({
                            'sheet_name': id_value,
                            'Type': 'Manual Peak',
                            'Annotation': peak_name,
                            'Offset': peak_time,
                            'Value': peak_value,
                            'mbase': None,
                            'mbase_time': None,
                            'mpeak': peak_value,
                            'mpeak_time': peak_time,
                            'pos_delta': pos_delta,
                            'neg_delta': neg_delta,
                            'Notes': ''
                        })
                        
                    except Exception as e:
                        print(f"Error processing measurements at column {col} for ID {id_value}: {str(e)}")
                        continue
            
            valid_ids.add(id_value)
    
    result_df = pd.DataFrame(data)
    return valid_ids, result_df

def find_peaks_and_bases(x, y, annotation, event, measurement_data, sheet_name, base_params, peak_params):
    """Find peaks (mins) and bases (maxes) in the data with specified parameters"""
    print(f"\nFinding peaks and bases for {sheet_name}")
    print(f"Data lengths - x: {len(x)}, y: {len(y)}, annotation: {len(annotation)}")
    
    try:
        y_array = np.array(y)
        x_array = np.array(x)
        annotation_array = np.array(annotation)
        event_array = np.array(event)
        
        print("\nData statistics:")
        print(f"Y value range - min: {np.min(y_array)}, max: {np.max(y_array)}")
        print(f"Y mean: {np.mean(y_array)}, std: {np.std(y_array)}")
        print(f"X value range - min: {np.min(x_array)}, max: {np.max(x_array)}")
        
        print("\nPeak detection parameters:")
        print(f"Base parameters: {base_params}")
        print(f"Peak parameters: {peak_params}")
        
        # Find peaks and bases
        bases, base_properties = find_peaks(y_array, **base_params)
        peaks, peak_properties = find_peaks(-y_array, **peak_params)
        
        print(f"\nDetection results:")
        print(f"Found {len(bases)} bases at indices: {bases}")
        print(f"Found {len(peaks)} peaks at indices: {peaks}")
        
        data = []
        
        # Process bases
        for i in bases:
            if i < len(x_array) and i < len(y_array):
                data.append({
                    'sheet_name': sheet_name,
                    'Type': 'Detected Base',
                    'Offset': float(x_array[i]/1000),
                    'Value': float(y_array[i]*10),
                    'Annotation': str(annotation_array[i]) if i < len(annotation_array) else '',
                    'Event': str(event_array[i]) if i < len(event_array) else '',
                    'mbase': None,
                    'mbase_time': None,
                    'mpeak': None,
                    'mpeak_time': None,
                    'pos_delta': None,
                    'neg_delta': None,
                    'Notes': ''
                })
        
        # Process peaks
        for i in peaks:
            if i < len(x_array) and i < len(y_array):
                data.append({
                    'sheet_name': sheet_name,
                    'Type': 'Detected Peak',
                    'Offset': float(x_array[i]/1000),
                    'Value': float(y_array[i]*10),
                    'Annotation': str(annotation_array[i]) if i < len(annotation_array) else '',
                    'Event': str(event_array[i]) if i < len(event_array) else '',
                    'mbase': None,
                    'mbase_time': None,
                    'mpeak': None,
                    'mpeak_time': None,
                    'pos_delta': None,
                    'neg_delta': None,
                    'Notes': ''
                })
        
        df = pd.DataFrame(data)
        if df.empty:
            print("No peaks or bases detected!")
            return df
            
        print(f"\nDetected points summary:")
        print(f"Total points detected: {len(df)}")
        print(f"Bases detected: {len(df[df['Type'] == 'Detected Base'])}")
        print(f"Peaks detected: {len(df[df['Type'] == 'Detected Peak'])}")
        
        return df.sort_values('Offset') if 'Offset' in df.columns else df
        
    except Exception as e:
        print(f"Error processing peaks and bases for {sheet_name}: {str(e)}")
        print("Traceback:")
        traceback.print_exc()
        return pd.DataFrame(columns=['sheet_name', 'Type', 'Annotation', 'Event', 'Offset', 'Value', 
                                   'mbase', 'mbase_time', 'mpeak', 'mpeak_time', 'Notes'])
    


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

def create_phase_overlay(data, metric, title):
    """Create line graph overlaying phases"""
    fig = px.line(
        data,
        x='number',
        y=metric,
        color='phase',
        title=title
    )
    return fig

def create_phase_graph_and_table(sheet_data, measurement_data, phase, selected_sheets, 
                               base_params, peak_params, include_words=None, 
                               exclude_words=None, include_events=None, 
                               exclude_events=None, show_actions=False):
    """Create graph and table for a phase with selection support"""
    print(f"\nCreating {phase} graph and table")
    print(f"Selected sheets: {selected_sheets}")
    print(f"Word filters - include: {include_words}, exclude: {exclude_words}")
    print(f"Event filters - include: {include_events}, exclude: {exclude_events}")
    
    fig = go.Figure()
    all_data = []
    
    # Filter sheets for the specified phase
    phase_sheets = {name: data for name, data in sheet_data.items() if phase in name}
    if selected_sheets:
        phase_sheets = {name: data for name, data in phase_sheets.items() if name in selected_sheets}
    
    print(f"Processing sheets: {list(phase_sheets.keys())}")
    
    for sheet_name, sheet_info in phase_sheets.items():
        print(f"\nProcessing sheet: {sheet_name}")
        data = sheet_info['data']
        
        # Filter data based on Annotation column
        filtered_data = data.copy()
        if include_words or exclude_words:
            if include_words:
                mask = filtered_data['Annotation'].str.contains('|'.join(include_words), case=False, na=False)
                filtered_data = filtered_data[mask]
            if exclude_words:
                mask = ~filtered_data['Annotation'].str.contains('|'.join(exclude_words), case=False, na=False)
                filtered_data = filtered_data[mask]
                
        # Filter data based on Event column
        if include_events or exclude_events:
            if include_events:
                mask = filtered_data['Event'].str.contains('|'.join(include_events), case=False, na=False)
                filtered_data = filtered_data[mask]
            if exclude_events:
                mask = ~filtered_data['Event'].str.contains('|'.join(exclude_events), case=False, na=False)
                filtered_data = filtered_data[mask]

        # Add line trace with filtered data
        if not filtered_data.empty:
            fig.add_trace(
                go.Scatter(
                    name=f'{sheet_name}',
                    x=filtered_data['Offset']/1000,
                    y=filtered_data['Value']*10,
                    mode='lines',
                    line=dict(color=COLORS['line'], width=2),
                    showlegend=True,                        
                )
            )

        # Add action labels if toggle is on     
        if show_actions:
            non_empty_annotations = filtered_data[filtered_data['Annotations'].notna()]
            if not non_empty_annotations.empty:
                fig.add_trace(
                    go.Scatter(
                        name=f'{sheet_name} (Actions)',
                        x=non_empty_annotations['Offset'] / 1000,
                        y=(non_empty_annotations['Value'] * 10) + 0.5,  # Adjusted Y values
                        mode='text',
                        text=non_empty_annotations['Annotations'],
                        textposition='top center',
                        textfont=dict(
                            size=14,
                            color='black'
                        ),
                        showlegend=False,
                        hoverinfo='skip'
                    )
                )


        
        # Find peaks and bases using filtered data
        if not filtered_data.empty:
            # Get detected points
            detected_points = find_peaks_and_bases(
                x=filtered_data['Offset'].values,
                y=filtered_data['Value'].values,
                annotation=filtered_data['Annotation'].values,
                event=filtered_data['Event'].values,
                measurement_data=measurement_data,
                sheet_name=sheet_name,
                base_params=base_params,
                peak_params=peak_params
            )
            
            # New rolling window detection
            rolling_points = detect_peaks_rolling_window(filtered_data, sheet_name)
            
            # Add detected points to graph and combine results
            if not detected_points.empty:
                detected_points['selected'] = True  # Default to selected
                all_data.append(detected_points)
                
            if not rolling_points.empty:
                rolling_points['selected'] = True  # Default to selected
                # Add rolling window bases to graph
                rolling_bases = rolling_points[rolling_points['Type'] == 'Rolling Base']
                if not rolling_bases.empty:
                    fig.add_trace(
                        go.Scatter(
                            name=f'{sheet_name} (Rolling Base)',
                            x=rolling_bases['Offset'],
                            y=rolling_bases['Value'],
                            mode='markers',
                            marker=dict(
                                color='rgb(102, 255, 51)',  # Green color
                                size=12,
                                symbol='circle',
                                line=dict(color='white', width=1)
                            ),
                            showlegend=True
                        )
                    )
                
                # Add rolling window peaks to graph
                rolling_peaks = rolling_points[rolling_points['Type'] == 'Rolling Peak']
                if not rolling_peaks.empty:
                    fig.add_trace(
                        go.Scatter(
                            name=f'{sheet_name} (Rolling Peak)',
                            x=rolling_peaks['Offset'],
                            y=rolling_peaks['Value'],
                            mode='markers',
                            marker=dict(
                                color='rgb(0, 176, 240)',  # Light Blue color
                                size=12,
                                symbol='circle',
                                line=dict(color='white', width=1)
                            ),
                            showlegend=True
                        )
                    )
                
                all_data.append(rolling_points)
            
            if not detected_points.empty:
                # Add detected bases to graph
                detected_bases = detected_points[detected_points['Type'] == 'Detected Base']
                if not detected_bases.empty:
                    fig.add_trace(
                        go.Scatter(
                            name=f'{sheet_name} (Detected Base)',
                            x=detected_bases['Offset'],
                            y=detected_bases['Value'],
                            mode='markers',
                            marker=dict(
                                color=COLORS['detected_base'],
                                size=12,
                                symbol='circle',
                                line=dict(color='white', width=1)
                            ),
                            showlegend=True
                        )
                    )
                
                # Add detected peaks to graph
                detected_peaks = detected_points[detected_points['Type'] == 'Detected Peak']
                if not detected_peaks.empty:
                    fig.add_trace(
                        go.Scatter(
                            name=f'{sheet_name} (Detected Peak)',
                            x=detected_peaks['Offset'],
                            y=detected_peaks['Value'],
                            mode='markers',
                            marker=dict(
                                color=COLORS['detected_peak'],
                                size=12,
                                symbol='circle',
                                line=dict(color='white', width=1)
                            ),
                            showlegend=True
                        )
                    )

        # Filter and add manual measurements
        manual_data = measurement_data[measurement_data['sheet_name'] == sheet_name].copy()
        if include_words or exclude_words:
            manual_data = filter_by_words(manual_data, include_words, exclude_words)

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
    
    # Update layout
    fig.update_layout(
        title=dict(
            text=f'{phase} Series Analysis',
            x=0.5,
            font=dict(size=24)
        ),
        xaxis_title='Time (seconds)',
        yaxis_title='Value',
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
        paper_bgcolor='white',
        xaxis=dict(
            showgrid=True,
            gridwidth=1,
            gridcolor='LightGray',
            zeroline=True,
            zerolinewidth=1,
            zerolinecolor='LightGray'
        ),
        yaxis=dict(
            showgrid=True,
            gridwidth=1,
            gridcolor='LightGray',
            zeroline=True,
            zerolinewidth=1,
            zerolinecolor='LightGray'
        )
    )
    
    # Combine all data for the table
    if all_data:
        combined_data = pd.concat(all_data, ignore_index=True)
        combined_data = combined_data.sort_values(['sheet_name', 'Offset'])
        combined_data['selected'] = True
        return fig, combined_data
    else:
        return fig, pd.DataFrame(columns=['sheet_name', 'Type', 'Annotation', 'Event', 'Offset', 'Value', 
                                        'mbase', 'mbase_time', 'mpeak', 'mpeak_time', 
                                        'pos_delta', 'neg_delta', 'Notes', 'selected'])
        empty_df['selected'] = True
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
    
    # Verify matching data
    print("\nData Validation:")
    
except FileNotFoundError as e:
    print(f"Error loading files: {str(e)}")
    sheet_data = {}
    measurement_data = pd.DataFrame()
    print("No data loaded")

# App layout
app.layout = html.Div([
    dcc.Tabs([
        dcc.Tab(label='Peak Detection', children=[
            # All original content starts here
            html.H1('Peak Detection Analysis Dashboard'),
            
            # Base Detection Parameter Controls
            html.Div([
                html.H3('Base Detection Parameters'),
                html.Div([
                    html.Label('Base Height:'),
                    dcc.Input(id='base_height-input', type='number', value=DEFAULT_BASE_PARAMS['height']),
                    html.Label('Base Prominence:'),
                    dcc.Input(id='base_prominence-input', type='number', value=DEFAULT_BASE_PARAMS['prominence']),
                    html.Label('Base Distance:'),
                    dcc.Input(id='base_distance-input', type='number', value=DEFAULT_BASE_PARAMS['distance']),
                    html.Label('Base Window Length:'),
                    dcc.Input(id='base_wlen-input', type='number', value=DEFAULT_BASE_PARAMS['wlen']),
                    html.Label('Base Threshold:'),
                    dcc.Input(id='base_threshold-input', type='number', value=DEFAULT_BASE_PARAMS['threshold']),
                    html.Label('Base Width:'),
                    dcc.Input(id='base_width-input', type='number', value=DEFAULT_BASE_PARAMS['width']),
                    html.Label('Base Plateau Size:'),
                    dcc.Input(id='base_plateau_size-input', type='number', value=DEFAULT_BASE_PARAMS['plateau_size']),
                ], style={'marginBottom': '20px', 'display': 'flex', 'gap': '10px', 'alignItems': 'center'}),
            ], style={'padding': '10px', 'backgroundColor': '#f8f9fa', 'borderRadius': '5px', 'marginBottom': '20px'}),
            
            # Peak Detection Parameter Controls
            html.Div([
                html.H3('Peak Detection Parameters'),
                html.Div([
                    html.Label('Peak Height:'),
                    dcc.Input(id='peak_height-input', type='number', value=DEFAULT_PEAK_PARAMS['height']),
                    html.Label('Peak Prominence:'),
                    dcc.Input(id='peak_prominence-input', type='number', value=DEFAULT_PEAK_PARAMS['prominence']),
                    html.Label('Peak Distance:'),
                    dcc.Input(id='peak_distance-input', type='number', value=DEFAULT_PEAK_PARAMS['distance']),
                    html.Label('Peak Window Length:'),
                    dcc.Input(id='peak_wlen-input', type='number', value=DEFAULT_PEAK_PARAMS['wlen']),
                    html.Label('Peak Threshold:'),
                    dcc.Input(id='peak_threshold-input', type='number', value=DEFAULT_PEAK_PARAMS['threshold']),
                    html.Label('Peak Width:'),
                    dcc.Input(id='peak_width-input', type='number', value=DEFAULT_PEAK_PARAMS['width']),
                    html.Label('Peak Plateau Size:'),
                    dcc.Input(id='peak_plateau_size-input', type='number', value=DEFAULT_PEAK_PARAMS['plateau_size']),
                ], style={'marginBottom': '20px', 'display': 'flex', 'gap': '10px', 'alignItems': 'center'}),
            ], style={'padding': '10px', 'backgroundColor': '#f8f9fa', 'borderRadius': '5px', 'marginBottom': '20px'}),

            # Base Parameters Section
            html.Div([
                html.H3('Base Parameters'),
                html.Div(id='base_params-table-container')
            ], style={'marginBottom': '20px', 'padding': '10px', 'backgroundColor': '#f8f9fa', 'borderRadius': '5px'}),
            
            # Peak Parameters Section
            html.Div([
                html.H3('Peak Parameters'),
                html.Div(id='peak_params-table-container')
            ], style={'marginBottom': '20px', 'padding': '10px', 'backgroundColor': '#f8f9fa', 'borderRadius': '5px'}),
            
            # Statistics Section
            html.Div([
                html.H3('Graph Statistics'),
                html.Div(id='stats-table-container')
            ], style={'marginBottom': '20px', 'padding': '10px', 'backgroundColor': '#f8f9fa', 'borderRadius': '5px'}),
            
            # Display Controls
            html.Div([
                html.H3('Display Controls', style={'marginBottom': '10px'}),
                html.Div([
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
                    )
                ], style={'marginBottom': '20px', 'display': 'flex', 'alignItems': 'center'})
            ], style={'padding': '10px', 'backgroundColor': '#f8f9fa', 'borderRadius': '5px', 'marginBottom': '20px'}),
            
            # Word Filters
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
                ),
            ], style={'marginBottom': '20px', 'display': 'flex', 'gap': '10px', 'alignItems': 'center'}),
            
            # Event Filters
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
                ),
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
                        {"name": "Diff", "id": "Diff", "type": "numeric", "format": {"specifier": ".3f"}},
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
                )
            ], style={'marginBottom': '30px', 'padding': '20px', 'backgroundColor': '#ffffff', 'borderRadius': '5px', 'boxShadow': '0 2px 4px rgba(0,0,0,0.1)'})
        ]),
    ]),    
        # Results Analysis Tab
        dcc.Tab(label='Results Analysis', children=[
        html.Div([
            html.H1('Results Analysis'),
            dcc.Download(id="download-dataframe-xlsx"),
                    html.Div([
                        # Phase Selection
                        html.Div([
                            # Phase Header
                            html.Div([
                                html.Label("Phase", style={'font-weight': 'bold', 'font-size': '16px'}),
                                html.Button("Clear All", id='phase-clear', className="clear-button"),
                            ], className="group-header"),

                            # Phase Radio Buttons
                            dcc.RadioItems(
                                id='phase-selection',
                                options=[
                                    {'label': 'All Phases', 'value': 'p1,p3,p4'},
                                    {'label': 'P1', 'value': 'p1'},
                                    {'label': 'P3', 'value': 'p3'},
                                    {'label': 'P4', 'value': 'p4'},
                                    {'label': 'P1 and P3', 'value': 'p1,p3'},
                                    {'label': 'P1 and P4', 'value': 'p1,p4'},
                                    {'label': 'P3 and P4', 'value': 'p3,p4'}
                                ],
                                className="radio-group",
                                style={'display': 'flex', 'flexDirection': 'column', 'gap': '10px'}
                            ),
                        ], className="filter-group"),

                        # Event Selection
                        html.Div([
                            # Event Header
                            html.Div([
                                html.Label("Event", style={'font-weight': 'bold', 'font-size': '16px'}),
                                html.Button("Clear All", id='event-clear', className="clear-button"),
                            ], className="group-header"),

                            # Event Radio Buttons
                            dcc.RadioItems(
                                id='event-selection',
                                options=[
                                    {'label': 'All Events', 'value': 'all'},
                                    {'label': 'A', 'value': 'walk'},
                                    {'label': 'B', 'value': 'up_and_go'},
                                    {'label': 'C', 'value': 'mp'},
                                    {'label': 'D', 'value': 'lp'},
                                    {'label': 'A and B', 'value': 'AB'},
                                    {'label': 'C and D', 'value': 'CD'}
                                ],
                                className="radio-group",
                                style={'display': 'flex', 'flexDirection': 'column', 'gap': '10px'}
                            )
                        ], className="filter-group"),

                        # Display for selected values
                        html.Div(id='selection-display', className="mt-4 p-4")
                    ], className="filter-container"),

                    # Worksheet selector with improved styling
                    html.Div([
                        html.Label('Select Results Worksheets:', style={'marginBottom': '10px', 'display': 'block'}),
                        html.Div([
                            dcc.Dropdown(
                                id='worksheet-results-selector',
                                options=[],
                                value=[],
                                multi=True,
                                placeholder="Select worksheets or use Select All button",
                                style={
                                    'width': '80%', 
                                    'display': 'inline-block', 
                                    'verticalAlign': 'middle',
                                    'maxHeight': '400px',  # Increased height
                                    'minHeight': '100px'   # Minimum height
                                },
                                optionHeight=50,
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
        html.Div([
            dcc.Tabs([
                # P1 Analysis Tab
                dcc.Tab(label='P1 Analysis', children=[
                    html.Div([
                        html.H3('P1 Overall Statistics'),
                        dcc.Graph(id='p1-overall-stats-graph'),
                        html.H3('P1 Delta Analysis'),
                        dcc.Graph(id='p1-delta-stats-graph'),
                        html.H3('P1 Peak Counts'),
                        dcc.Graph(id='p1-peak-counts-graph'),
                        html.H3('P1 Statistical Analysis'),
                        dash_table.DataTable(id='p1-stats-analysis'),
                        html.H3('P1 Event Analysis'),
                        dcc.Graph(id='p1-event-avg-graph'),
                        dcc.Graph(id='p1-event-delta-graph'),
                        html.Button("Export P1 Analysis", id="export-p1", n_clicks=0),
                    ])
                ]),
            html.Div([
                html.Button("Export P1 Analysis", id="export-button-p1", n_clicks=0,
                    style={
                        'marginLeft': '10px',
                        'height': '36px',
                        'backgroundColor': '#f8f9fa',
                        'border': '1px solid #ddd',
                        'borderRadius': '4px',
                        'cursor': 'pointer'
                    }
                )
            ]),
                # P3 Analysis Tab
                dcc.Tab(label='P3 Analysis', children=[
                    html.Div([
                        html.H3('P3 Overall Statistics'),
                        dcc.Graph(id='p3-overall-stats-graph'),
                        html.H3('P3 Delta Analysis'),
                        dcc.Graph(id='p3-delta-stats-graph'),
                        html.H3('P3 Peak Counts'),
                        dcc.Graph(id='p3-peak-counts-graph'),
                        html.H3('P3 Statistical Analysis'),
                        dash_table.DataTable(id='p3-stats-analysis'),
                        html.H3('P3 Event Analysis'),
                        dcc.Graph(id='p3-event-avg-graph'),
                        dcc.Graph(id='p3-event-delta-graph'),
                        html.Button("Export P3 Analysis", id="export-p3", n_clicks=0),
                    ])
                ]),
             html.Div([
                html.Button("Export P3 Analysis", id="export-p3", n_clicks=0,
                    style={
                        'marginLeft': '10px',
                        'height': '36px',
                        'backgroundColor': '#f8f9fa',
                        'border': '1px solid #ddd',
                        'borderRadius': '4px',
                        'cursor': 'pointer'
                    }
                )
            ]),   
                # P4 Analysis Tab
                dcc.Tab(label='P4 Analysis', children=[
                    html.Div([
                        html.H3('P4 Overall Statistics'),
                        dcc.Graph(id='p4-overall-stats-graph'),
                        html.H3('P4 Delta Analysis'),
                        dcc.Graph(id='p4-delta-stats-graph'),
                        html.H3('P4 Peak Counts'),
                        dcc.Graph(id='p4-peak-counts-graph'),
                        html.H3('P4 Statistical Analysis'),
                        dash_table.DataTable(id='p4-stats-analysis'),
                        html.H3('P4 Event Analysis'),
                        dcc.Graph(id='p4-event-avg-graph'),
                        dcc.Graph(id='p4-event-delta-graph'),
                        html.Button("Export P4 Analysis", id="export-p4", n_clicks=0),
                    ])
                ]),
             html.Div([
                html.Button("Export P4 Analysis", id="export-p4", n_clicks=0,
                    style={
                        'marginLeft': '10px',
                        'height': '36px',
                        'backgroundColor': '#f8f9fa',
                        'border': '1px solid #ddd',
                        'borderRadius': '4px',
                        'cursor': 'pointer'
                    }
                )
            ]),   
                # Cross-Phase Analysis Tab
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
                        ),
                        html.Div([
                            html.Button("Export Cross-Phase Analysis", id="export-cross-phase", n_clicks=0,
                                style={
                                    'marginLeft': '10px',
                                    'height': '36px',
                                    'backgroundColor': '#f8f9fa',
                                    'border': '1px solid #ddd',
                                    'borderRadius': '4px',
                                    'cursor': 'pointer'
                                }
                            )
                        ])
                    ])
                ])
            ])
        ])
    ])
])
])

# Add this function definition before the callback
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

@app.callback(
     Output('p1-graph', 'figure'),
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
     Output('base_params-table-container', 'children'),
     Output('peak_params-table-container', 'children'),
     Output('stats-table-container', 'children'),
     Output('p1-confirm-deselect', 'displayed'),
     Output('p3-confirm-deselect', 'displayed'),     
     Output('p4-confirm-deselect', 'displayed'),    
     [Input('base_height-input', 'value'),
     Input('base_prominence-input', 'value'),
     Input('base_distance-input', 'value'),
     Input('base_wlen-input', 'value'),
     Input('base_threshold-input', 'value'),
     Input('base_width-input', 'value'),
     Input('base_plateau_size-input', 'value'),
     Input('peak_height-input', 'value'),
     Input('peak_prominence-input', 'value'),
     Input('peak_distance-input', 'value'),
     Input('peak_wlen-input', 'value'),
     Input('peak_threshold-input', 'value'),
     Input('peak_width-input', 'value'),
     Input('peak_plateau_size-input', 'value'),
     Input('include-words', 'value'),
     Input('exclude-words', 'value'),
     Input('include-events', 'value'),
     Input('exclude-events', 'value'),
     Input('p1-worksheet-selector', 'value'),
     Input('p3-worksheet-selector', 'value'),
     Input('p4-worksheet-selector', 'value'),
     Input('action-toggle', 'n_clicks'),
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
     State('p4-table', 'data')]
)

def update_graphs_and_tables(base_height, base_prominence, base_distance, base_wlen, base_threshold, base_width, base_plateau_size,
                           peak_height, peak_prominence, peak_distance, peak_wlen, peak_threshold, peak_width, peak_plateau_size,
                           include_words, exclude_words, include_events, exclude_events, p1_sheets, p3_sheets, p4_sheets, n_clicks,
                           p1_relayout, p3_relayout, p4_relayout, 
                           p1_selected_rows, p3_selected_rows, p4_selected_rows,
                           p1_table_data, p3_table_data, p4_table_data,
                           p1_select_all, p1_deselect_all,
                           p3_select_all, p3_deselect_all,
                           p4_select_all, p4_deselect_all,
                           p1_confirm, p3_confirm, p4_confirm,  
                           p1_table_state, p3_table_state, p4_table_state):
    """Update all graphs and tables based on user inputs and parameters"""
    
    # Initialize dictionaries
    results = {}
    stats_data = {}
    base_params_data = {}
    peak_params_data = {}
    
    # Get all selected sheets
    all_selected_sheets = []
    if p1_sheets:
        all_selected_sheets.extend(p1_sheets)
    if p3_sheets:
        all_selected_sheets.extend(p3_sheets)
    if p4_sheets:
        all_selected_sheets.extend(p4_sheets)
    
    # Process filters
    include_list = [word.strip() for word in include_words.split(',')] if include_words else DEFAULT_INCLUDE_WORDS
    exclude_list = [word.strip() for word in exclude_words.split(',')] if exclude_words else DEFAULT_EXCLUDE_WORDS
    include_event = [word.strip() for word in include_events.split(',')] if include_events else DEFAULT_INCLUDE_EVENTS
    exclude_event = [word.strip() for word in exclude_events.split(',')] if exclude_events else DEFAULT_EXCLUDE_EVENTS

    show_actions = bool(n_clicks % 2)

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
        # Handle initial deselect button clicks to show confirmation
        elif button_id == 'p1-deselect-all':
            show_p1_confirm = True
        elif button_id == 'p3-deselect-all':
            show_p3_confirm = True
        elif button_id == 'p4-deselect-all':
            show_p4_confirm = True
        # Handle select all button clicks
        elif button_id == 'p1-select-all' and p1_table_state:  # Note: using table_state here
            p1_selected_rows = list(range(len(p1_table_state)))
        elif button_id == 'p3-select-all' and p3_table_state:
            p3_selected_rows = list(range(len(p3_table_state)))
        elif button_id == 'p4-select-all' and p4_table_state:
            p4_selected_rows = list(range(len(p4_table_state)))
        # Initialize selections if not yet set
        elif not any(btn in button_id for btn in ['select-all', 'deselect-all', 'confirm-deselect']):
            if p1_selected_rows is None or len(p1_selected_rows) == 0:
                p1_selected_rows = list(range(len(p1_table_state))) if p1_table_state else []
            if p3_selected_rows is None or len(p3_selected_rows) == 0:
                p3_selected_rows = list(range(len(p3_table_state))) if p3_table_state else []
            if p4_selected_rows is None or len(p4_selected_rows) == 0:
                p4_selected_rows = list(range(len(p4_table_state))) if p4_table_state else []

    # Collect manual parameters
    manual_params = {
        'base': {
            'height': base_height if base_height != 0 else None,
            'prominence': base_prominence if base_prominence != 0 else None,
            'distance': base_distance if base_distance != 0 else None,
            'wlen': base_wlen if base_wlen != 0 else None,
            'threshold': base_threshold if base_threshold != 0 else None,
            'width': base_width if base_width != 0 else None,
            'plateau_size': base_plateau_size if base_plateau_size != 0 else None
        },
        'peak': {
            'height': peak_height if peak_height != 0 else None,
            'prominence': peak_prominence if peak_prominence != 0 else None,
            'distance': peak_distance if peak_distance != 0 else None,
            'wlen': peak_wlen if peak_wlen != 0 else None,
            'threshold': peak_threshold if peak_threshold != 0 else None,
            'width': peak_width if peak_width != 0 else None,
            'plateau_size': peak_plateau_size if peak_plateau_size != 0 else None
        }
    }

    # Process visible ranges
    visible_ranges = {}
    for phase, relayout_data in [('P1', p1_relayout), ('P3', p3_relayout), ('P4', p4_relayout)]:
        if relayout_data and 'xaxis.range[0]' in relayout_data:
            visible_ranges[phase] = (
                relayout_data['xaxis.range[0]'],
                relayout_data['xaxis.range[1]']
            )

    def get_combined_y_range(results_dict):
        """Calculate the combined y-range across all graphs"""
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
    for phase, phase_sheets, selected_rows, table_data in [
        ('P1', p1_sheets, p1_selected_rows, p1_table_data),
        ('P3', p3_sheets, p3_selected_rows, p3_table_data),
        ('P4', p4_sheets, p4_selected_rows, p4_table_data)
    ]:
        if not phase_sheets:
            continue

        # Get statistics for this phase using all selected sheets
        combined_stats, individual_stats = analyze_multiple_graphs_statistics(
            sheet_data, 
            all_selected_sheets,
            visible_ranges.get(phase)
        )
        stats_data[phase] = (combined_stats, individual_stats)

        # Get parameters with manual overrides
        base_params_result, peak_params_result = get_dynamic_peak_params_multi(
            combined_stats, 
            individual_stats,
            manual_params=manual_params
        )

        base_params_data[phase] = base_params_result
        peak_params_data[phase] = peak_params_result

        # Create initial graphs and tables
        fig, table = create_phase_graph_and_table(
            sheet_data, 
            measurement_data, 
            phase, 
            all_selected_sheets,
            base_params_result, 
            peak_params_result,
            include_list, 
            exclude_list,
            include_event, 
            exclude_event, 
            show_actions
        )

        # Handle selection state and delta calculations
        if isinstance(table, pd.DataFrame) and not table.empty:
            # If we have no selection state yet, default to all rows selected
            if selected_rows is None or len(selected_rows) == 0 or max(selected_rows) >= len(table):
                selected_rows = list(range(len(table)))
                
            # Ensure selected_rows are within bounds
            selected_rows = [i for i in selected_rows if i < len(table)]
            
            # Mark selected state based on selected_rows
            table['selected'] = False  # First mark all as deselected
            if selected_rows:
                table.iloc[selected_rows, table.columns.get_loc('selected')] = True
            
            # Recalculate deltas based on current selection state
            table = calculate_deltas_with_selection(table)
            
            # Calculate manual deltas
            table = calculate_manual_deltas_with_selection(table)
            
            # Update graph to show only selected points
            selected_points = table[table['selected']]
            
            # Keep the line traces and action labels
            base_traces = [trace for trace in fig.data if trace.mode == 'lines' or trace.mode == 'text']
            fig.data = base_traces
            
            # Add back only the selected points
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
                        )
                    )

        results[phase] = {
            'fig': fig, 
            'table': table,
            'selected_rows': selected_rows if selected_rows is not None else list(range(len(table) if isinstance(table, pd.DataFrame) else 0))
        }
        
    y_range = get_combined_y_range(results)
    
    # Update all figures with the combined y-axis range
    if y_range:
        for phase, relayout_data in [('P1', p1_relayout), ('P3', p3_relayout), ('P4', p4_relayout)]:
            if phase in results:
                # Check if there's a manual y-axis range from user interaction
                if relayout_data and 'yaxis.range[0]' in relayout_data:
                    manual_y_range = [relayout_data['yaxis.range[0]'], relayout_data['yaxis.range[1]']]
                    results[phase]['fig'].update_layout(yaxis=dict(range=manual_y_range))
                else:
                    # Use the calculated combined range
                    results[phase]['fig'].update_layout(yaxis=dict(range=y_range))
        
    # Calculate event statistics using only selected points
    p1_stats = calculate_event_statistics(sheet_data, p1_sheets) if p1_sheets else []
    p1_stats = filter_selected_points(p1_stats, p1_table_data)
    
    p3_stats = calculate_event_statistics(sheet_data, p3_sheets) if p3_sheets else []
    p3_stats = filter_selected_points(p3_stats, p3_table_data)
    
    p4_stats = calculate_event_statistics(sheet_data, p4_sheets) if p4_sheets else []
    p4_stats = filter_selected_points(p4_stats, p4_table_data)

    # Create base params tables
    base_params_tables = {}
    for phase, individual_base_params in base_params_data.items():
        if individual_base_params:
            base_params_tables[phase] = create_base_params_table(
                base_params_data[phase],
                {sheet: params for sheet, params in base_params_data.items()}
            )
    
    # Create peak params tables
    peak_params_tables = {}
    for phase, individual_peak_params in peak_params_data.items():
        if individual_peak_params:
            peak_params_tables[phase] = create_peak_params_table(
                peak_params_data[phase],
                {sheet: params for sheet, params in peak_params_data.items()}
            )

    # Create combined statistics table
    all_stats = {}
    for phase_stats in stats_data.values():
        combined, individual = phase_stats
        if combined:
            all_stats.update(individual)

    stats_table = create_multi_stats_table(
        stats_data.get('P1', (None, {}))[0],
        all_stats
    )

    # Create toggle status display
    toggle_status = html.Div(
        "Action Labels ON" if show_actions else "Action Labels OFF",
        style={'color': 'green' if show_actions else 'red'}
    )

    # Prepare empty defaults
    empty_figure = go.Figure()
    empty_data = []

    # Get results for each phase
    p1_result = results.get('P1', {'fig': empty_figure, 'table': pd.DataFrame(), 'selected_rows': []})
    p3_result = results.get('P3', {'fig': empty_figure, 'table': pd.DataFrame(), 'selected_rows': []})
    p4_result = results.get('P4', {'fig': empty_figure, 'table': pd.DataFrame(), 'selected_rows': []})

    # Return all updated components
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
        base_params_tables.get('P1', html.Div()),
        peak_params_tables.get('P1', html.Div()),
        stats_table,
        show_p1_confirm,
        show_p3_confirm,
        show_p4_confirm
    )

# Callback for P1 save button
@app.callback(
    Output('p1-save-status', 'children'),
    [Input('save-p1-button', 'n_clicks')],
    [State('p1-graph', 'figure'),
     State('p1-table', 'data'),
     State('p1-stats-table', 'data'),
     State('p1-filename-input', 'value')]
)
def save_p1_results(n_clicks, fig_data, table_data, event_data, filename):
    if n_clicks > 0:
        if not filename:
            filename = 'mnl_py_results'
        success, message = save_single_phase_to_excel('P1', fig_data, table_data, event_data, filename)
        return html.Div(message, style={'color': 'green' if success else 'red'})
    return ""

# P1 Select All callback

def update_p1_selection(select_clicks, deselect_clicks, table_data):
    ctx = dash.callback_context
    if not ctx.triggered:
        return []
    button_id = ctx.triggered[0]['prop_id'].split('.')[0]
    
    if button_id == 'p1-select-all':
        return list(range(len(table_data)))
    elif button_id == 'p1-deselect-all':
        return []
    return []

# Callback for P3 save button
@app.callback(
    Output('p3-save-status', 'children'),
    [Input('save-p3-button', 'n_clicks')],
    [State('p3-graph', 'figure'),
     State('p3-table', 'data'),
     State('p3-stats-table', 'data'),
     State('p3-filename-input', 'value')]
)
def save_p3_results(n_clicks, fig_data, table_data, event_data, filename):
    if n_clicks > 0:
        if not filename:
            filename = 'mnl_py_results'
        success, message = save_single_phase_to_excel('P3', fig_data, table_data, event_data, filename)
        return html.Div(message, style={'color': 'green' if success else 'red'})
    return ""

# P3 Select All callback

def update_p3_selection(select_clicks, deselect_clicks, table_data):
    ctx = dash.callback_context
    if not ctx.triggered:
        return []
    button_id = ctx.triggered[0]['prop_id'].split('.')[0]
    
    if button_id == 'p3-select-all':
        return list(range(len(table_data)))
    elif button_id == 'p3-deselect-all':
        return []
    return []

# Callback for P4 save button
@app.callback(
    Output('p4-save-status', 'children'),
    [Input('save-p4-button', 'n_clicks')],
    [State('p4-graph', 'figure'),
     State('p4-table', 'data'),
     State('p4-stats-table', 'data'),
     State('p4-filename-input', 'value')]
)
def save_p4_results(n_clicks, fig_data, table_data, event_data, filename):
    if n_clicks > 0:
        if not filename:
            filename = 'mnl_py_results'
        success, message = save_single_phase_to_excel('P4', fig_data, table_data, event_data, filename)
        return html.Div(message, style={'color': 'green' if success else 'red'})
    return ""

# P4 Select All callback

def update_p4_selection(select_clicks, deselect_clicks, table_data):
    ctx = dash.callback_context
    if not ctx.triggered:
        return []
    button_id = ctx.triggered[0]['prop_id'].split('.')[0]
    
    if button_id == 'p4-select-all':
        return list(range(len(table_data)))
    elif button_id == 'p4-deselect-all':
        return []
    return []

# Initialize data
try:
    print("\nInitializing data...")
    # First read measurement data to get valid IDs
    valid_ids, measurement_data = read_measurement_data('Mnl_Analysis.xlsx')
    print(f"Loaded measurement data with {len(valid_ids)} valid IDs")
    
    # Then read the phase data
    sheet_data = read_excel_sheets('All_Phase_Sheets.xlsx')
    print(f"Loaded {len(sheet_data)} sheets from phase data")
    
except FileNotFoundError as e:
    print(f"Error loading files: {str(e)}")
    sheet_data = {}
    measurement_data = pd.DataFrame()
    print("No data loaded")
except Exception as e:
    print(f"Error during initialization: {str(e)}")
    traceback.print_exc()
    sheet_data = {}
    measurement_data = pd.DataFrame()
    print("Error loading data")
    
@app.callback(
    Output('worksheet-results-selector', 'options'),
    [Input('phase-selection', 'value'),
     Input('event-selection', 'value')]
)
def update_worksheet_options(selected_phase, selected_event):
    """Update the dropdown options based on phase and event selections"""
    # Get all available files first
    files = get_results_files()
    filtered_files = files.copy()

    try:
        print(f"Initial files: {files}")  # Debug print
        
        # Filter by phase if selected
        if selected_phase and selected_phase != 'p1,p3,p4':
            phases = selected_phase.split(',')
            # Updated pattern to match "_p1", "_p3", "_p4" regardless of ID number
            filtered_files = [f for f in filtered_files 
                            if any(f'_{p}.' in f.lower() or f'_{p}_' in f.lower() 
                                  for p in phases)]
            print(f"After phase filter: {filtered_files}")  # Debug print

        # Filter by event if selected
        if selected_event and selected_event != 'all':
            if selected_event == 'AB':
                filtered_files = [f for f in filtered_files 
                                if any(term in f.lower() for term in ['walk', 'up_and_go'])]
            elif selected_event == 'CD':
                filtered_files = [f for f in filtered_files 
                                if any(term in f.lower() for term in ['mp', 'lp'])]
            else:
                filtered_files = [f for f in filtered_files if selected_event in f.lower()]
            print(f"After event filter: {filtered_files}")  # Debug print

        # Sort files for easier selection
        filtered_files.sort()
        return [{'label': f, 'value': f} for f in filtered_files]

    except Exception as e:
        print(f"Error in update_worksheet_options: {str(e)}")
        traceback.print_exc()
        return [{'label': f, 'value': f} for f in files]

# Update the Select All callback
@app.callback(
    Output('worksheet-results-selector', 'value'),
    [Input('select-all-worksheets', 'n_clicks'),
     Input('clear-all-worksheets', 'n_clicks'),
     Input('phase-selection', 'value'),
     Input('event-selection', 'value')],
    [State('worksheet-results-selector', 'options')]
)
def update_worksheet_selection(select_clicks, clear_clicks, selected_phase, selected_event, options):
    """Handle Select All/Clear All button clicks and filter changes"""
    if not ctx.triggered:
        return []

    button_id = ctx.triggered[0]['prop_id'].split('.')[0]

    # Clear selection when filters change
    if button_id in ['phase-selection', 'event-selection']:
        return []

    # Handle button clicks
    if button_id == 'select-all-worksheets' and options:
        return [option['value'] for option in options]
    elif button_id == 'clear-all-worksheets':
        return []

    return []

@app.callback(
    Output('selection-display', 'children'),
    [Input('phase-selection', 'value'),
     Input('event-selection', 'value'),
     Input('worksheet-results-selector', 'value')]
)
def update_selection_display(selected_phase, selected_event, selected_worksheets):
    phase_display = selected_phase if selected_phase else 'None'
    event_display = selected_event if selected_event else 'None'
    worksheet_count = len(selected_worksheets) if selected_worksheets else 0
    
    return html.Div([
        html.P(f"Selected Phase: {phase_display}"),
        html.P(f"Selected Event: {event_display}"),
        html.P(f"Selected Worksheets: {worksheet_count}"),
        html.P(f"Available Worksheets: {worksheet_count}")  # Added to show available worksheets
    ])

@app.callback(
    [Output('phase-selection', 'value'),
     Output('event-selection', 'value')],
    [Input('phase-clear', 'n_clicks'),
     Input('event-clear', 'n_clicks')]
)
def clear_selections(phase_clear, event_clear):
    if not ctx.triggered:
        raise dash.exceptions.PreventUpdate
    
    button_id = ctx.triggered[0]['prop_id'].split('.')[0]
    
    if button_id == 'phase-clear':
        return None, dash.no_update
    elif button_id == 'event-clear':
        return dash.no_update, None
    
    raise dash.exceptions.PreventUpdate

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
    [Input('worksheet-results-selector', 'value'),
     Input('phase-selection', 'value'),
     Input('event-selection', 'value')]
)

def update_phase_analyses(selected_files, selected_phase, selected_event):  # Added missing parameters
    # Initialize empty defaults with proper layout
    empty_fig = go.Figure()
    empty_fig.update_layout(
        title="No data available",
        xaxis_title="",
        yaxis_title="",
        showlegend=False,
        height=400,  # Set a default height
        margin=dict(t=50, b=50, l=50, r=50),
        plot_bgcolor='white',
        paper_bgcolor='white'
    )

    # Return empty state if no files selected
    if not selected_files:
        return ([empty_fig] * 12) + [[]] * 3 + [empty_fig] * 3
    
    try:
        # Initialize empty defaults
        empty_fig = go.Figure()
        empty_fig.update_layout(
            title="No data available",
            xaxis_title="",
            yaxis_title="",
            showlegend=False
        )

        # Filter files based on selections
        filtered_files = selected_files if selected_files else []
        
        # Remove the phase_checked and event_checked checks since we're using direct values now
        if selected_phase and selected_phase != 'P1,P3,P4':
            phases = selected_phase.split(',')
            filtered_files = [f for f in filtered_files if any(f'_p{p[-1].lower()}_' in f.lower() for p in phases)]
        
        # Event filtering
        if selected_event and selected_event != 'all':
            if selected_event == 'AB':
                filtered_files = [f for f in filtered_files if any(event in f.lower() for event in ['walk', 'up_and_go'])]
            elif selected_event == 'CD':
                filtered_files = [f for f in filtered_files if any(event in f.lower() for event in ['mp', 'lp'])]
            else:
                filtered_files = [f for f in filtered_files if selected_event in f.lower()]

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

# Add the export callbacks
from openpyxl import Workbook
from openpyxl.drawing.image import Image
import io

@app.callback(
    Output("download-dataframe-xlsx", "data"),
    [Input("export-p1", "n_clicks"),
     Input("export-p3", "n_clicks"),
     Input("export-p4", "n_clicks"),
     Input("export-cross-phase", "n_clicks")],
    [State('p1-export-filename', 'value'),
     State('p3-export-filename', 'value'),
     State('p4-export-filename', 'value'),
     State('cross-phase-export-filename', 'value'),
     State('p1-stats-analysis', 'data'),
     State('p3-stats-analysis', 'data'),
     State('p4-stats-analysis', 'data'),
     State('cross-phase-stats', 'data'),
     State('p1-overall-stats-graph', 'figure'),
     State('p3-overall-stats-graph', 'figure'),
     State('p4-overall-stats-graph', 'figure'),
     State('cross-phase-comparison', 'figure')]
)
def export_analysis(p1_clicks, p3_clicks, p4_clicks, cross_phase_clicks,
                   p1_filename, p3_filename, p4_filename, cross_phase_filename,
                   p1_stats, p3_stats, p4_stats, cross_phase_stats,
                   p1_graph, p3_graph, p4_graph, cross_phase_graph):
    """Export analysis data and graphs to Excel"""
    
    ctx_triggered = ctx.triggered[0]['prop_id'].split('.')[0]
    if not ctx_triggered or not any([p1_clicks, p3_clicks, p4_clicks, cross_phase_clicks]):
        raise dash.exceptions.PreventUpdate
        
    wb = Workbook()
    
    if ctx_triggered == "export-button-p1":
        export_phase_analysis(wb, "P1", p1_stats, p1_graph)
        filename = f"{p1_filename if p1_filename else 'p1_analysis'}.xlsx"
    elif ctx_triggered == "export-p3":
        export_phase_analysis(wb, "P3", p3_stats, p3_graph)
        filename = f"{p3_filename if p3_filename else 'p3_analysis'}.xlsx"
    elif ctx_triggered == "export-p4":
        export_phase_analysis(wb, "P4", p4_stats, p4_graph)
        filename = f"{p4_filename if p4_filename else 'p4_analysis'}.xlsx"
    elif ctx_triggered == "export-cross-phase":
        export_cross_phase_analysis(wb, cross_phase_stats, cross_phase_graph)
        filename = f"{cross_phase_filename if cross_phase_filename else 'cross_phase_analysis'}.xlsx"
        
    # Save to BytesIO
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return dcc.send_bytes(output.getvalue(), filename)

def export_phase_analysis(wb, phase, stats_data, graph_fig):
    """Helper function to export phase-specific analysis"""
    # Create sheets
    graph_sheet = wb.create_sheet(f"{phase} Graphs")
    data_sheet = wb.create_sheet(f"{phase} Data")
    
    # Export graph as image
    img_bytes = io.BytesIO()
    graph_fig.write_image(img_bytes, format='png')
    img = Image(img_bytes)
    graph_sheet.add_image(img, 'A1')
    
    # Export data
    headers = list(stats_data[0].keys()) if stats_data else []
    data_sheet.append(headers)
    for row in stats_data:
        data_sheet.append([row[h] for h in headers])

def export_cross_phase_analysis(wb, stats_data, graph_fig):
    """Helper function to export cross-phase analysis"""
    # Create sheets
    graph_sheet = wb.create_sheet("Cross-Phase Graphs")
    data_sheet = wb.create_sheet("Cross-Phase Data")
    
    # Export graph as image
    img_bytes = io.BytesIO()
    graph_fig.write_image(img_bytes, format='png')
    img = Image(img_bytes)
    graph_sheet.add_image(img, 'A1')
    
    # Export data
    headers = list(stats_data[0].keys()) if stats_data else []
    data_sheet.append(headers)
    for row in stats_data:
        data_sheet.append([row[h] for h in headers])

# Run the server
if __name__ == '__main__':
    app.run_server(debug=True, port=8050)

# %%


# %%


# %%


# %%



