#!/usr/bin/env python
# coding: utf-8

# In[1]:


### MANUAL AND PYTHON BASES RWIN PEAK DETECTION WITH RESULTS
import dash
from dash import Dash, html, dcc, dash_table, callback_context
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
import glob
import re
import json
import base64
import io
from scipy import stats
import plotly.express as px
import plotly.subplots as sp
import plotly.graph_objs as go


# In[2]:


# Initialize the app
app = dash.Dash(__name__, suppress_callback_exceptions=True, title="Peak Detection Analysis")

# Set up directories for projects and CSV files
PROJECTS_DIRECTORY = os.path.join(os.getcwd(), "projects")
CSV_DIRECTORY = os.path.join(os.getcwd(), "data")
print(f"Current working directory: {os.getcwd()}")
print(f"CSV directory: {CSV_DIRECTORY}")

PROJECT_CONFIG_FILE = "project_config.json"
print(f"Project config file name: {PROJECT_CONFIG_FILE}")

# Make sure the directory exists
if not os.path.exists(PROJECTS_DIRECTORY):
    os.makedirs(PROJECTS_DIRECTORY)
if not os.path.exists(CSV_DIRECTORY):
    os.makedirs(CSV_DIRECTORY)  # Create it if it doesn't exist

# Register callbacks
def initialize_app():
    """Initialize application state and load existing projects"""
    # Initialize application variables
    if not hasattr(app, 'event_selections'):
        app.event_selections = []
    if not hasattr(app, 'phase_selections'):
        app.phase_selections = []
    if not hasattr(app, 'base_selections'):
        app.base_selections = []
    if not hasattr(app, 'peak_selections'):
        app.peak_selections = []
    if not hasattr(app, 'selected_point'):
        app.selected_point = None
    if not hasattr(app, 'event_data'):
        app.event_data = pd.DataFrame()
    if not hasattr(app, 'original_filename'):
        app.original_filename = None
    
    # Project management variables - always reset these
    app.current_project_id = None
    app.current_project_name = None
    app.current_scaling_factor = 10  # Default to Scaling Factor A
    
    print("App initialized with variables:")
    print(f"app.current_project_id = {app.current_project_id}")
    print(f"app.current_project_name = {app.current_project_name}")
    print(f"app.current_scaling_factor = {app.current_scaling_factor}")
    
    # Scan for existing projects
    projects = get_projects_list()
    print(f"Initialization: Found {len(projects)} projects")
    return projects

def create_project_structure(project_name):
    """Create a new project with the required folder structure"""
    try:
        print(f"Attempting to create project: {project_name}")
        
        # Validate project name
        if not project_name or not project_name.strip():
            print("Project name cannot be empty")
            return False, "Project name cannot be empty"
        
        # Clean project name (remove special characters, replace spaces with underscores)
        clean_name = ''.join(c if c.isalnum() else '_' for c in project_name)
        print(f"Cleaned project name: {clean_name}")
        
        # Create project directory
        project_dir = os.path.join(PROJECTS_DIRECTORY, clean_name)
        print(f"Project directory path: {project_dir}")
        
        # Check if project already exists
        if os.path.exists(project_dir):
            print(f"Project '{project_name}' already exists at {project_dir}")
            return False, f"Project '{project_name}' already exists"
        
        # Create project directory and subfolders
        os.makedirs(project_dir)
        print(f"Created project directory: {project_dir}")
        
        # Add raw_data folder to the structure
        os.makedirs(os.path.join(project_dir, "raw_data"))
        os.makedirs(os.path.join(project_dir, "trace_data"))
        os.makedirs(os.path.join(project_dir, "peak_data"))
        os.makedirs(os.path.join(project_dir, "results"))
        os.makedirs(os.path.join(project_dir, "statistics"))
        
        # Create project configuration file
        config = {
            "project_name": project_name,
            "created_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "scaling_factor": 10,  # Default to Scaling Factor A
            "active_csv_file": None
        }
        
        config_path = os.path.join(project_dir, PROJECT_CONFIG_FILE)
        print(f"Writing config to: {config_path}")
        with open(config_path, 'w') as f:
            json.dump(config, f, indent=4)
        
        print(f"Project '{project_name}' created successfully")
        return True, f"Project '{project_name}' created successfully"
    
    except Exception as e:
        print(f"Error creating project: {str(e)}")
        traceback.print_exc()
        return False, f"Error creating project: {str(e)}"

def get_projects_list():
    """Get list of available projects"""
    try:
        print(f"Getting projects list from directory: {PROJECTS_DIRECTORY}")
        projects = []
        
        # Check if directory exists
        if not os.path.exists(PROJECTS_DIRECTORY):
            print(f"Project directory doesn't exist: {PROJECTS_DIRECTORY}")
            os.makedirs(PROJECTS_DIRECTORY)  # Create it if it doesn't exist
            print(f"Created projects directory: {PROJECTS_DIRECTORY}")
            return []
            
        # List all directories in the projects folder
        items = os.listdir(PROJECTS_DIRECTORY)
        print(f"Found {len(items)} items in projects directory: {items}")
        
        for item in items:
            project_dir = os.path.join(PROJECTS_DIRECTORY, item)
            config_file = os.path.join(project_dir, PROJECT_CONFIG_FILE)
            
            # Check if it's a directory and has a config file
            if os.path.isdir(project_dir) and os.path.exists(config_file):
                try:
                    # Read the project name from the config file
                    with open(config_file, 'r') as f:
                        config = json.load(f)
                        project_name = config.get("project_name", item)
                        projects.append({"label": project_name, "value": item})
                        print(f"Found project: {project_name} ({item})")
                except Exception as e:
                    print(f"Error reading config for {item}: {str(e)}")
                    # Fall back to directory name if config can't be read
                    projects.append({"label": item, "value": item})
                    print(f"Added project using directory name: {item}")
        
        print(f"Returning {len(projects)} projects: {projects}")
        return sorted(projects, key=lambda x: x["label"])
    
    except Exception as e:
        print(f"Error getting projects list: {str(e)}")
        traceback.print_exc()
        return []

def load_project(project_id):
    """Load project configuration and return details"""
    try:
        print(f"Loading project: {project_id}")
        project_dir = os.path.join(PROJECTS_DIRECTORY, project_id)
        config_file = os.path.join(project_dir, PROJECT_CONFIG_FILE)
        
        if not os.path.exists(project_dir):
            print(f"Project directory not found: {project_dir}")
            return None, f"Project directory not found: {project_dir}"
        
        if not os.path.exists(config_file):
            print(f"Project configuration not found: {config_file}")
            return None, f"Project configuration not found"
        
        with open(config_file, 'r') as f:
            config = json.load(f)
        
        print(f"Successfully loaded project: {project_id}")
        return config, None
    
    except json.JSONDecodeError as e:
        print(f"Error parsing project configuration: {str(e)}")
        traceback.print_exc()
        return None, f"Error parsing project configuration: {str(e)}"
    except Exception as e:
        print(f"Error loading project: {str(e)}")
        traceback.print_exc()
        return None, f"Error loading project: {str(e)}"

def save_project_config(project_id, config):
    """Save updated project configuration"""
    try:
        project_dir = os.path.join(PROJECTS_DIRECTORY, project_id)
        config_file = os.path.join(project_dir, PROJECT_CONFIG_FILE)
        
        with open(config_file, 'w') as f:
            json.dump(config, f, indent=4)
        
        return True, "Project configuration saved"
    
    except Exception as e:
        print(f"Error saving project configuration: {str(e)}")
        traceback.print_exc()
        return False, f"Error saving configuration: {str(e)}"

def get_project_csv_files(project_id):
    """Get list of CSV files in the project's trace_data directory"""
    try:
        if not project_id:
            return []
        
        trace_data_dir = os.path.join(PROJECTS_DIRECTORY, project_id, "trace_data")
        if not os.path.exists(trace_data_dir):
            return []
        
        # Get all CSV files in the trace_data directory
        csv_files = glob.glob(os.path.join(trace_data_dir, "*.csv"))
        return [os.path.basename(f) for f in csv_files]
    
    except Exception as e:
        print(f"Error getting project CSV files: {str(e)}")
        traceback.print_exc()
        return []

def parse_csv_contents(contents, filename):
    """Parse contents of uploaded CSV file"""
    content_type, content_string = contents.split(',')
    decoded = base64.b64decode(content_string)
    
    try:
        df = pd.read_csv(io.StringIO(decoded.decode('utf-8')))
        return df, None
    except Exception as e:
        print(f"Error parsing CSV contents: {str(e)}")
        traceback.print_exc()
        return None, str(e)
    
def find_column_by_pattern(df, pattern):
    """Find a column that matches a pattern (case-insensitive)"""
    pattern = pattern.lower()
    for col in df.columns:
        if pattern in col.lower():
            return col
    return None

def normalize_dataframe_columns(df, scaling_factor=None):
    """Normalize column names and add missing required columns"""
    # Identify key columns using pattern matching
    offset_col = find_column_by_pattern(df, 'offset')
    trace_col = find_column_by_pattern(df, 'trace')
    annotation_col = find_column_by_pattern(df, 'annotation')
    
    # Create a copy to avoid modifying the original
    normalized_df = df.copy()
    
    # Rename columns if found
    rename_dict = {}
    if offset_col and offset_col != 'Offset':
        rename_dict[offset_col] = 'Offset'
    if trace_col and trace_col != 'Trace':
        rename_dict[trace_col] = 'Trace'
    if annotation_col and annotation_col != 'Annotations':
        rename_dict[annotation_col] = 'Annotations'
    
    if rename_dict:
        normalized_df = normalized_df.rename(columns=rename_dict)
    
    # Add missing required columns
    for col in ['Offset', 'Trace', 'Annotations']:
        if col not in normalized_df.columns:
            if col == 'Annotations':
                normalized_df[col] = ''
            else:
                normalized_df[col] = np.nan
    
    # Add Time column if it doesn't exist
    if 'Time' not in normalized_df.columns:
        normalized_df['Time'] = normalized_df['Offset'] / 1000  # assuming offset is in milliseconds
    
    # Create the Value column right after adding it
    if 'Trace' in normalized_df.columns and 'Value' not in normalized_df.columns and scaling_factor is not None:
        normalized_df['Value'] = normalized_df['Trace'] * scaling_factor
        print(f"Added Value column with scaling factor {scaling_factor}")
    
    # Add additional required columns
    for col in ['Annotation', 'Event', 'Phase', 'VAS', 'Mnl_Base', 'Mnl_Peak', 'Pos_Delta', 'Neg_Delta']:
        if col not in normalized_df.columns:
            if col in ['Mnl_Base', 'Mnl_Peak', 'Pos_Delta', 'Neg_Delta']:
                normalized_df[col] = None
            else:
                normalized_df[col] = ''
    
    # Now reorder the columns explicitly to ensure proper order
    # Get the list of all columns
    all_columns = normalized_df.columns.tolist()
    
    # Remove the core columns that we want to be first
    for col in ['Offset', 'Time', 'Trace', 'Value']:
        if col in all_columns:
            all_columns.remove(col)
    
    # Create a new order with our core columns first, then the rest
    new_order = []
    
    # Add Offset if it exists
    if 'Offset' in normalized_df.columns:
        new_order.append('Offset')
    
    # Add Time if it exists
    if 'Time' in normalized_df.columns:
        new_order.append('Time')
    
    # Add Trace if it exists
    if 'Trace' in normalized_df.columns:
        new_order.append('Trace')
    
    # Add Value if it exists
    if 'Value' in normalized_df.columns:
        new_order.append('Value')
    
    # Add all remaining columns
    new_order.extend(all_columns)
    
    # Apply the new order
    normalized_df = normalized_df[new_order]
    
    print(f"Final column order: {normalized_df.columns.tolist()}")
    return normalized_df

def process_annotations(df):
    """Process Annotations column to populate Annotation column"""
    if 'Annotations' in df.columns and 'Annotation' in df.columns:
        # Find rows with annotations
        annotated_rows = df[df['Annotations'].notna() & (df['Annotations'] != '')]
        
        if not annotated_rows.empty:
            # Sort by offset
            annotated_rows = annotated_rows.sort_values('Offset')
            
            # Get the sorted annotation offsets
            annotation_offsets = annotated_rows['Offset'].tolist()
            annotation_values = annotated_rows['Annotations'].tolist()
            
            # Add an "end" boundary if needed
            if len(annotation_offsets) > 0:
                annotation_offsets.append(df['Offset'].max() + 1)
                annotation_values.append('')
            
            # Populate the Annotation column based on ranges
            for i in range(len(annotation_offsets) - 1):
                start_offset = annotation_offsets[i]
                end_offset = annotation_offsets[i + 1]
                annotation = annotation_values[i]
                
                # Set annotation for all rows in range
                mask = (df['Offset'] >= start_offset) & (df['Offset'] < end_offset)
                df.loc[mask, 'Annotation'] = annotation
    
    return df

def save_uploaded_csv_with_scaling(contents, filename, project_id, scaling_factor, output_filename=None):
    """
    Parse, apply scaling factor, and save uploaded CSV to project's trace_data directory
    with scaling factor indicated in the filename.
    """
    try:
        if not project_id:
            return None, "No project selected"
        
        trace_data_dir = os.path.join(PROJECTS_DIRECTORY, project_id, "trace_data")
        if not os.path.exists(trace_data_dir):
            os.makedirs(trace_data_dir)
        
        # Parse the CSV contents
        content_type, content_string = contents.split(',')
        decoded = base64.b64decode(content_string)
        
        try:
            # Load the CSV into a DataFrame
            df = pd.read_csv(io.StringIO(decoded.decode('utf-8')))
            
            # Normalize column names and add missing required columns
            df = normalize_dataframe_columns(df, scaling_factor)  # Pass scaling factor here
            
            # Determine output filename
            if output_filename and output_filename.strip():
                # Use user-provided filename
                final_filename = f"{output_filename.strip()}_trace.csv"
            else:
                # Use original filename with _trace suffix
                name_parts = os.path.splitext(filename)
                final_filename = f"{name_parts[0]}_trace{name_parts[1]}"
            
            # Save to CSV file
            file_path = os.path.join(trace_data_dir, final_filename)
            df.to_csv(file_path, index=False)
            
            # Return the DataFrame along with success message
            return df, f"Saved {final_filename} to project with scaling factor {scaling_factor}x applied"
            
        except Exception as e:
            print(f"Error processing CSV data: {str(e)}")
            traceback.print_exc()
            
            # If we fail to process, still save the original file
            file_path = os.path.join(trace_data_dir, filename)
            with open(file_path, 'wb') as f:
                f.write(decoded)
            
            return None, f"Error processing CSV: {str(e)}. Saved original file without scaling."
    
    except Exception as e:
        print(f"Error saving uploaded CSV: {str(e)}")
        traceback.print_exc()
        return None, f"Error saving file: {str(e)}"

def extract_scaling_info_from_filename(filename):
    """Extract scaling factor from filename if present"""
    try:
        # Look for pattern like "_ScaleA_x10" or "_Custom_x5" in the filename
        scaling_patterns = [
            r'_ScaleA_x(\d+)',
            r'_ScaleB_x(\d+)',
            r'_ScaleC_x(\d+)',
            r'_ScaleD_x(\d+)',
            r'_Custom_x(\d+)'
        ]
        
        for pattern in scaling_patterns:
            match = re.search(pattern, filename)
            if match:
                scaling_factor = int(match.group(1))
                return scaling_factor, True
        
        return None, False
    except Exception as e:
        print(f"Error extracting scaling info: {str(e)}")
        return None, False

def load_csv_data(filename, project_id, scaling_factor=10):
    """
    Load CSV data with smart scaling detection - if filename contains scaling info,
    it will not apply additional scaling. Otherwise, it applies the provided scaling factor.
    """
    print(f"Loading CSV file: {filename}, project: {project_id}")
    
    try:
        # Get the full path to the CSV file in the project
        if not project_id:
            return pd.DataFrame(), "No project selected"
        
        file_path = os.path.join(PROJECTS_DIRECTORY, project_id, "trace_data", filename)
        if not os.path.exists(file_path):
            return pd.DataFrame(), f"File not found: {filename}"
        
        # Check if filename already contains scaling information
        file_scaling_factor, has_scaling = extract_scaling_info_from_filename(filename)
        
        # If file already has scaling info, use that instead of applying additional scaling
        if has_scaling:
            print(f"File {filename} already has scaling factor {file_scaling_factor}x applied")
            apply_scaling = False
            display_scaling = file_scaling_factor
        else:
            print(f"File {filename} has no scaling info, will apply scaling factor {scaling_factor}x")
            apply_scaling = True
            display_scaling = scaling_factor
        
        # Load the CSV file
        df = pd.read_csv(file_path)
        
        # Normalize column names and add missing required columns
        df = normalize_dataframe_columns(df)
        
        # Apply scaling factor to the Value column if needed
        if 'Value' not in df.columns and 'Trace' in df.columns:
            # Create Value column if it doesn't exist
            df['Value'] = df['Trace'] * (scaling_factor if apply_scaling else 1)
        elif 'Value' in df.columns and apply_scaling and scaling_factor != 1:
            df['Value'] = df['Value'] * scaling_factor
        
        # Process annotations
        df = process_annotations(df)
        
        return df, None
    
    except Exception as e:
        print(f"Error loading CSV file {filename}: {str(e)}")
        traceback.print_exc()
        # Return empty DataFrame with required columns
        return pd.DataFrame(columns=['Offset', 'Time', 'Value', 'Annotations', 'Annotation', 'Event', 'Phase', 'VAS']), str(e)

def get_project_file_details(project_id):
    """Get detailed information about files in the project"""
    try:
        if not project_id:
            return []
        
        file_details = []
        project_dir = os.path.join(PROJECTS_DIRECTORY, project_id)
        
        # Process files in raw_data directory
        raw_data_dir = os.path.join(project_dir, "raw_data")
        if os.path.exists(raw_data_dir):
            for file in os.listdir(raw_data_dir):
                if file.endswith('.csv'):
                    file_path = os.path.join(raw_data_dir, file)
                    file_details.append({
                        'filename': file,
                        'type': 'Raw Data',
                        'date_added': datetime.fromtimestamp(os.path.getmtime(file_path)).strftime('%Y-%m-%d %H:%M:%S'),
                        'folder': 'raw_data'
                    })
        
        # Process files in trace_data directory
        trace_data_dir = os.path.join(project_dir, "trace_data")
        if os.path.exists(trace_data_dir):
            for file in os.listdir(trace_data_dir):
                if file.endswith('.csv'):
                    file_path = os.path.join(trace_data_dir, file)
                    # Check if it's a work-in-progress file
                    if '_wip.csv' in file:
                        file_type = 'Work in Progress'
                    else:
                        file_type = 'Trace Data'
                    
                    file_details.append({
                        'filename': file,
                        'type': file_type,
                        'date_added': datetime.fromtimestamp(os.path.getmtime(file_path)).strftime('%Y-%m-%d %H:%M:%S'),
                        'folder': 'trace_data'
                    })
        
        # Process files in peak_data directory
        peak_data_dir = os.path.join(project_dir, "peak_data")
        if os.path.exists(peak_data_dir):
            for file in os.listdir(peak_data_dir):
                if file.endswith('.csv') or file.endswith('.xlsx'):
                    file_path = os.path.join(peak_data_dir, file)
                    file_details.append({
                        'filename': file,
                        'type': 'Peak Data',
                        'date_added': datetime.fromtimestamp(os.path.getmtime(file_path)).strftime('%Y-%m-%d %H:%M:%S'),
                        'folder': 'peak_data'
                    })
        
        # Process files in results directory
        results_dir = os.path.join(project_dir, "results")
        if os.path.exists(results_dir):
            for file in os.listdir(results_dir):
                if file.endswith('.csv'):
                    file_path = os.path.join(results_dir, file)
                    file_details.append({
                        'filename': file,
                        'type': 'Results Data',
                        'date_added': datetime.fromtimestamp(os.path.getmtime(file_path)).strftime('%Y-%m-%d %H:%M:%S'),
                        'folder': 'results'
                    })
        
        # Process files in statistics directory
        statistics_dir = os.path.join(project_dir, "statistics")
        if os.path.exists(statistics_dir):
            for file in os.listdir(statistics_dir):
                if file.endswith('.csv'):
                    file_path = os.path.join(statistics_dir, file)
                    file_details.append({
                        'filename': file,
                        'type': 'Statistics',
                        'date_added': datetime.fromtimestamp(os.path.getmtime(file_path)).strftime('%Y-%m-%d %H:%M:%S'),
                        'folder': 'statistics'
                    })
        
        return sorted(file_details, key=lambda x: x['date_added'], reverse=True)
    
    except Exception as e:
        print(f"Error getting project file details: {str(e)}")
        traceback.print_exc()
        return []

def detect_peaks_rolling_window(data, sheet_name, window_size):
    """
    Detect peaks and valleys using a rolling window comparison method with edge case handling.
    """
    species_scale = 10
    peaks_data = []
    n = len(data)

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
                            'Value': float(current_value*species_scale),
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
            df = calculate_auto_peak_deltas(df)
        
        return df
    
    except Exception as e:
        print(f"Error processing sheet {sheet_name}: {str(e)}")
        traceback.print_exc()
        return pd.DataFrame()

# Callback to save data with a new name
@app.callback(
    Output('save-csv-file-status', 'children'),
    Input('save-data-button', 'n_clicks'),
    [State('save-csv-filename-input', 'value'),
     State('upload-preview-table', 'data'),
     State('scaling-factor-dropdown', 'value')],
    prevent_initial_call=True
)
def save_data_with_new_name(n_clicks, filename, data, scaling_factor):
    print(f"\n=== SAVE DATA CALLBACK ===")
    print(f"Button clicked: {n_clicks}, Filename: {filename}")
    print(f"Number of rows in data: {len(data) if data else 0}")
    
    if n_clicks is None:
        return ""
        
    if not filename or not filename.strip():
        return html.Div("Please enter a filename", style={'color': 'red'})
    
    if not data:
        return html.Div("No data to save", style={'color': 'red'})
    
    if not hasattr(app, 'current_project_id') or not app.current_project_id:
        return html.Div("No project loaded", style={'color': 'red'})
    
    try:
        # Convert data back to DataFrame
        df = pd.DataFrame(data)
        print(f"Original columns: {df.columns.tolist()}")

        # SIMPLIFIED APPROACH: Force the column order directly
        # 1. First make sure Value exists
        if 'Value' not in df.columns and 'Trace' in df.columns:
            df['Value'] = df['Trace'] * scaling_factor
            print(f"Added Value column with scaling factor {scaling_factor}")
        
        # 2. Identify all columns
        columns = df.columns.tolist()
        
        # 3. Force an explicit reordering
        desired_order = ['Offset', 'Time', 'Trace', 'Value']
        remaining_columns = [col for col in columns if col not in desired_order]
        
        # 4. Check each column exists before including it in final order
        final_order = []
        for col in desired_order:
            if col in columns:
                final_order.append(col)
        
        # 5. Add remaining columns
        final_order.extend(remaining_columns)
        
        print(f"Attempting to reorder columns to: {final_order}")
        
        # 6. Reorder the DataFrame
        df = df[final_order]
        
        # 7. Verify the new order
        print(f"Actual column order after reordering: {df.columns.tolist()}")
        
        # Create the output file path
        trace_data_dir = os.path.join(PROJECTS_DIRECTORY, app.current_project_id, "trace_data")
        if not os.path.exists(trace_data_dir):
            print(f"Creating trace_data directory: {trace_data_dir}")
            os.makedirs(trace_data_dir)
        
        # Ensure filename has _trace suffix
        clean_filename = filename.strip()
        if not clean_filename.endswith('_trace'):
            final_filename = f"{clean_filename}_trace.csv"
        else:
            final_filename = f"{clean_filename}.csv"
        
        file_path = os.path.join(trace_data_dir, final_filename)
        print(f"Saving to: {file_path}")
        
        # Save to CSV
        df.to_csv(file_path, index=False)
        print(f"Successfully saved file with {len(df)} rows")
        
        # Refresh the project files list
        file_details = get_project_file_details(app.current_project_id)
        
        # Return success message
        return html.Div(f"Saved processed data with scaling factor {scaling_factor}x as {final_filename}", 
                       style={'color': 'green'})
    
    except Exception as e:
        print(f"Error saving data: {str(e)}")
        traceback.print_exc()  # Print full traceback for debugging
        return html.Div(f"Error saving data: {str(e)}", style={'color': 'red'})

# Define the Load Data tab layout
load_data_tab = dcc.Tab(
    label='Load Data',
    children=[
        html.Div([
            html.H1('Load Data', style={'textAlign': 'center'}),
            
            # Project Management section
            html.Div([
                html.H3('Project Management', style={'marginBottom': '10px'}),
                html.Div([
                    html.Div([
                        html.Label('Project Name:', style={'marginRight': '10px', 'fontWeight': 'bold'}),
                        dcc.Input(
                            id='project-name-input',
                            type='text',
                            placeholder='Enter project name',
                            style={'width': '250px', 'marginRight': '10px'}
                        ),
                        html.Button(
                            'Create Project',
                            id='create-project-button',
                            style={
                                'padding': '5px 10px',
                                'backgroundColor': '#4CAF50',
                                'color': 'white',
                                'border': 'none',
                                'borderRadius': '4px',
                                'cursor': 'pointer'
                            }
                        ),
                    ], style={'display': 'flex', 'alignItems': 'center', 'marginBottom': '10px'}),
                    html.Div(id='project-create-status'),
                ], style={'marginBottom': '10px'}),
                
                html.Div([
                    html.Label('Current Projects:', style={'marginRight': '10px', 'fontWeight': 'bold'}),
                    dcc.Dropdown(
                        id='load-project-selector',
                        options=[],
                        placeholder="Select an existing project",
                        style={'width': '300px', 'marginRight': '10px'}
                    ),
                    html.Button(
                        'Load Project',
                        id='load-project-button',
                        style={
                            'padding': '5px 10px',
                            'backgroundColor': '#2196F3',
                            'color': 'white',
                            'border': 'none',
                            'borderRadius': '4px',
                            'cursor': 'pointer',
                            'marginRight': '10px'
                        }
                    ),
                ], style={'display': 'flex', 'alignItems': 'center', 'marginBottom': '10px'}),
                html.Div(id='project-load-status'),
            ], style={'marginBottom': '20px', 'padding': '10px', 'backgroundColor': '#f8f9fa', 'borderRadius': '5px'}),
            
            # Scaling Factor section
            html.Div([
                html.Label('Species:', style={'marginRight': '10px', 'fontWeight': 'bold'}),
                dcc.Dropdown(
                    id='scaling-factor-dropdown',
                    options=[
                        {'label': 'Human (x10)', 'value': 10},
                        {'label': 'Equine (x7)', 'value': 7},
                        {'label': 'Canine (x3)', 'value': 3},
                        {'label': 'Bull (x1)', 'value': 1}
                    ],
                    value=10,  # Default to Scale A
                    style={'width': '200px'}
                ),
                html.Div(id='scaling-factor-status', style={'marginLeft': '10px'})
            ], style={'display': 'flex', 'alignItems': 'center', 'marginBottom': '20px'}),
            
            # File Upload section
            html.Div([
                html.H3('Upload CSV Files to Project', style={'marginBottom': '10px'}),
                dcc.Upload(
                    id='upload-csv',
                    children=html.Div([
                        'Drag and Drop or ',
                        html.A('Select CSV Files')
                    ]),
                    style={
                        'width': '100%',
                        'height': '60px',
                        'lineHeight': '60px',
                        'borderWidth': '1px',
                        'borderStyle': 'dashed',
                        'borderRadius': '5px',
                        'textAlign': 'center',
                        'margin': '10px 0'
                    },
                    multiple=True
                ),
                html.Div(id='upload-status'),
            ], style={'marginBottom': '20px'}),
            
            # Preview uploaded data
            html.Div([
                html.H3('Uploaded Data Preview', style={'marginBottom': '10px'}),
                html.Div([  # Add this wrapper div for the button and filename input
                    html.Label('Save As: ', style={'marginRight': '10px', 'fontWeight': 'bold'}),
                    dcc.Input(
                        id='save-csv-filename-input',
                        type='text',
                        placeholder='Enter filename to save (no extension needed)',
                        style={'width': '300px', 'marginRight': '10px'}
                    ),
                    html.Button(
                        'Save Data',
                        id='save-data-button',
                        style={
                            'padding': '5px 10px',
                            'backgroundColor': '#4CAF50',
                            'color': 'white',
                            'border': 'none',
                            'borderRadius': '4px',
                            'cursor': 'pointer'
                        }
                    ),
                    html.Div(id='save-csv-file-status', style={'marginLeft': '10px'})
                ], style={'display': 'flex', 'alignItems': 'center', 'marginBottom': '15px'}),
                
                dash_table.DataTable(
                    id='upload-preview-table',
                    columns=[
                        {'name': 'Offset', 'id': 'Offset', 'type': 'numeric'},
                        {'name': 'Time', 'id': 'Time', 'type': 'numeric'},
                        {'name': 'Trace', 'id': 'Trace', 'type': 'numeric'},
                        {'name': 'Value', 'id': 'Value', 'type': 'numeric'},
                        {'name': 'Annotations', 'id': 'Annotations'},
                    ],
                    data=[],
                    page_size=25,  # Increase the number of rows per page
                    style_table={
                        'overflowX': 'auto',
                        'maxHeight': '500px',  # Increase maximum height
                        'overflowY': 'scroll'  # Force scroll to always be enabled
                    },
                    style_cell={'textAlign': 'left', 'minWidth': '100px', 'maxWidth': '180px'},
                    # Add pagination controls
                    page_action='native',
                    page_current=0,
                ), 
            ], style={'marginBottom': '20px'}),
            
            # File selection section
            html.Div([
                html.H3('Project Files', style={'marginBottom': '10px'}),
                html.Button(
                    'Upload Files',
                    id='upload-files-button',
                    style={
                        'padding': '5px 10px',
                        'backgroundColor': '#f8f9fa',
                        'border': '1px solid #ddd',
                        'borderRadius': '4px',
                        'cursor': 'pointer',
                        'marginBottom': '10px'
                    }
                ),
                dash_table.DataTable(
                    id='project-files-table',
                    columns=[
                        {'name': 'Filename', 'id': 'filename'},
                        {'name': 'Type', 'id': 'type'},
                        {'name': 'Date Added', 'id': 'date_added'}
                    ],
                    data=[],
                    page_size=10,
                    style_table={'overflowX': 'auto'},
                    style_cell={'textAlign': 'left'},
                    row_selectable='single'
                ),
            ], style={'marginBottom': '20px'}),
            
            # Hidden div for storing temporary data
            html.Div(id='hidden-div', style={'display': 'none'})
        ], style={'padding': '20px'})
    ]
)

# @app.callback(
#     Output('initialization-trigger', 'children'),
#     Input('tabs', 'value'),
#     prevent_initial_call=False
# )
# def initialize_on_load(tab_value):
#     """Initialization callback that runs when the app loads"""
#     print("App initialization triggered")
#     projects = initialize_app()
#     print(f"Initial projects found: {len(projects)}")
#     return f"Initialized with {len(projects)} projects"

@app.callback(
    Output('load-project-selector', 'options'),
    [Input('tabs', 'value'),
     Input('create-project-button', 'n_clicks')],
    prevent_initial_call=False  # Allow initial call to populate options at startup
)
def update_projects_list(tab_value, n_clicks):
    print(f"Updating project list. Tab: {tab_value}, n_clicks: {n_clicks}")
    projects = get_projects_list()
    print(f"Found {len(projects)} projects: {projects}")
    return projects

# Callback to create a new project
@app.callback(
    [Output('project-create-status', 'children', allow_duplicate=True),
     Output('project-name-input', 'value', allow_duplicate=True),
     Output('load-project-selector', 'options', allow_duplicate=True)],  
    Input('create-project-button', 'n_clicks'),
    State('project-name-input', 'value'),
    prevent_initial_call=True
)
def create_project(n_clicks, project_name):
    if n_clicks is None or not project_name:
        # Return empty values if no click or no project name
        return "", project_name, dash.no_update
    
    success, message = create_project_structure(project_name)
    
    style = {'color': 'green', 'margin': '10px 0'} if success else {'color': 'red', 'margin': '10px 0'}
    
    # Get updated project list
    updated_projects = get_projects_list()
    
    print(f"Project created. Success: {success}, Message: {message}")
    print(f"Updated project list: {updated_projects}")
    
    # Return status message, reset input field (if successful), and updated dropdown options
    return html.Div(message, style=style), "" if success else project_name, updated_projects

# Callback to load a project
@app.callback(
    [Output('project-load-status', 'children'),
     Output('scaling-factor-dropdown', 'value'),
     Output('project-files-table', 'data')],
    Input('load-project-button', 'n_clicks'),
    State('load-project-selector', 'value'),  
    prevent_initial_call=True
)
def load_selected_project(n_clicks, project_id):
    print(f"\n=== LOAD PROJECT CALLBACK ===")
    print(f"Button clicks: {n_clicks}, Project ID: {project_id}")
    
    if n_clicks is None:
        print("No button clicks")
        raise PreventUpdate
        
    if not project_id:
        print("No project selected")
        return html.Div("No project selected", style={'color': 'red', 'margin': '10px 0'}), 10, []
    
    print(f"Loading project: {project_id}")
    config, error = load_project(project_id)
    
    if error or not config:
        print(f"Error loading project: {error}")
        return html.Div(error or "Error loading project", style={'color': 'red', 'margin': '10px 0'}), 10, []
    
    # Set app variables
    app.current_project_id = project_id
    app.current_project_name = config.get("project_name", project_id)
    app.current_scaling_factor = config.get("scaling_factor", 10)
    
    print(f"Set app.current_project_id = {app.current_project_id}")
    print(f"Set app.current_project_name = {app.current_project_name}")
    print(f"Set app.current_scaling_factor = {app.current_scaling_factor}")
    
    # Get project files
    file_details = get_project_file_details(project_id)
    print(f"Found {len(file_details)} files in project")
    
    return (
        html.Div(f"Project '{app.current_project_name}' loaded successfully", style={'color': 'green', 'margin': '10px 0'}),
        app.current_scaling_factor,
        file_details
    )

# Callback to refresh project files
@app.callback(
    Output('project-files-table', 'data', allow_duplicate=True),
    Input('upload-files-button', 'n_clicks'),
    prevent_initial_call=True
)
def refresh_project_files(n_clicks):
    if n_clicks is None:
        raise PreventUpdate
    
    if not hasattr(app, 'current_project_id') or not app.current_project_id:
        print("No current project ID to refresh files")
        return []
    
    print(f"Refreshing files for project: {app.current_project_id}")
    return get_project_file_details(app.current_project_id)

# Callback to update scaling factor when dropdown changes
@app.callback(
    Output('scaling-factor-status', 'children'),
    Input('scaling-factor-dropdown', 'value')
)
def update_scaling_factor(scaling_factor):
    if scaling_factor is None:
        return ""
    
    app.current_scaling_factor = scaling_factor
    
    # Update project config if a project is loaded
    if hasattr(app, 'current_project_id') and app.current_project_id:
        config, _ = load_project(app.current_project_id)
        if config:
            config['scaling_factor'] = scaling_factor
            save_project_config(app.current_project_id, config)
    
    factor_labels = {10: "Human (x10)", 7: "Equine (x7)", 3: "Canine (x3)", 1: "Bull (x1)"}
    species = factor_labels.get(scaling_factor, f"Custom")
    
    return html.Span(f"Using {species} scaling factor (x{scaling_factor})", style={'color': 'blue'})

@app.callback(
    Output('project-storage', 'data'),
    [Input('load-project-button', 'n_clicks')],
    [State('load-project-selector', 'value')],
    prevent_initial_call=True
)
def update_project_storage(n_clicks, project_id):
    if n_clicks is None:
        raise PreventUpdate
    
    if not project_id:
        print("No project ID to store")
        return {}
    
    print(f"Storing project ID in storage: {project_id}")
    return {'project_id': project_id}

# Then use this storage in your upload callback
@app.callback(
    [Output('upload-status', 'children'),
     Output('upload-preview-table', 'data'),
     Output('upload-preview-table', 'columns')],
    Input('upload-csv', 'contents'),
    [State('upload-csv', 'filename'),
     State('scaling-factor-dropdown', 'value'),
     State('project-storage', 'data')],
    prevent_initial_call=True
)
def upload_and_preview_csv(contents, filenames, scaling_factor, stored_project):
    print("\n=== UPLOAD CSV CALLBACK ===")
    
    if contents is None:
        print("No contents provided")
        return "", [], []
    
    # First, check if we have a project ID stored in the app
    project_id = None
    if hasattr(app, 'current_project_id') and app.current_project_id:
        project_id = app.current_project_id
        print(f"Using app.current_project_id: {project_id}")
    # If not, try to get it from storage
    elif stored_project and 'project_id' in stored_project:
        project_id = stored_project['project_id']
        # Also set it on the app for future use
        app.current_project_id = project_id
        print(f"Retrieved project ID from storage: {project_id}")
    
    # If we still don't have a project ID, return an error
    if not project_id:
        print("No project loaded (checked both app state and storage)")
        return html.Div("No project loaded. Please load a project first.", style={'color': 'red'}), [], []
    
    # Ensure we're dealing with a list
    if not isinstance(contents, list):
        contents = [contents]
        filenames = [filenames]
    
    # Process only the first file for now
    content = contents[0]
    filename = filenames[0]
    
    # Debug information
    print(f"Uploading file: {filename}")
    print(f"Current project: {app.current_project_id}")
    print(f"Current scaling factor: {scaling_factor}")
    
    # Check if a project is loaded
    if not hasattr(app, 'current_project_id') or not app.current_project_id:
        print(f"No project loaded. app.current_project_id = {getattr(app, 'current_project_id', None)}")
        return html.Div("No project loaded. Please load a project first.", style={'color': 'red'}), [], []
    
    # ===== CHANGE 1: Save the raw file first to raw_data folder =====
    try:
        # Ensure raw_data directory exists
        raw_data_dir = os.path.join(PROJECTS_DIRECTORY, project_id, "raw_data")
        if not os.path.exists(raw_data_dir):
            os.makedirs(raw_data_dir)
            print(f"Created raw_data directory: {raw_data_dir}")
            
        # Parse the raw content
        content_type, content_string = content.split(',')
        decoded = base64.b64decode(content_string)
        
        # Save the raw file directly without processing
        raw_file_path = os.path.join(raw_data_dir, filename)
        with open(raw_file_path, 'wb') as f:
            f.write(decoded)
        
        print(f"Saved raw file to: {raw_file_path}")
    except Exception as e:
        print(f"Error saving raw file: {str(e)}")
        return html.Div(f"Error saving raw file: {str(e)}", style={'color': 'red'}), [], []
    
    # ===== CHANGE 2: Now parse the CSV for preview only without saving to trace_data =====
    # Parse the CSV contents for preview
    df, error = parse_csv_contents(content, filename)
    
    if error or df is None:
        print(f"Error parsing CSV: {error}")
        return html.Div(f"Raw file saved, but error parsing CSV for preview: {error}", 
                      style={'color': 'orange'}), [], []
    
    # Normalize the DataFrame and apply scaling for preview only
    try:
        df = normalize_dataframe_columns(df)
        
        # Apply scaling factor to create Value column
        if 'Trace' in df.columns:
            df['Value'] = df['Trace'] * scaling_factor
            
        # Force column order for display
        if all(col in df.columns for col in ['Offset', 'Time', 'Trace', 'Value']):
            # Get all columns that aren't in the core set
            other_cols = [col for col in df.columns if col not in ['Offset', 'Time', 'Trace', 'Value']]
            # Set final order
            df = df[['Offset', 'Time', 'Trace', 'Value'] + other_cols]
            print(f"Preview table column order: {df.columns.tolist()}")
    except Exception as e:
        print(f"Error processing CSV: {str(e)}")
        return html.Div(f"Raw file saved, but error processing for preview: {str(e)}", 
                      style={'color': 'orange'}), [], []
    
    # ===== CHANGE 3: No longer automatically save the processed file =====
    # We'll only preview the data here, saving with scaling is done with the Save Data button
    
    # Create columns for the preview table based on actual DataFrame columns
    preview_columns = []
    # Add our core columns in the desired order
    for col in ['Offset', 'Time', 'Trace', 'Value']:
        if col in df.columns:
            preview_columns.append({'name': col, 'id': col})
    
    # Add remaining columns
    for col in df.columns:
        if col not in ['Offset', 'Time', 'Trace', 'Value']:
            preview_columns.append({'name': col, 'id': col})
    
    # Update project files table
    _ = get_project_file_details(project_id)
    
    # Return success message and preview data
    return (
        html.Div(f"Raw file '{filename}' saved to project's raw_data folder. Use 'Save Data' to save with scaling applied.", 
               style={'color': 'green', 'margin': '10px 0'}), 
        df.to_dict('records'), 
        preview_columns
    )

# @app.callback(
#     Output('hidden-div', 'children'),
#     [Input('load-project-button', 'n_clicks'),
#      Input('upload-csv', 'contents')],
#     [State('project-storage', 'data')],
#     prevent_initial_call=True
# )
# def debug_project_state(load_clicks, upload_contents, stored_project):
#     """Debug callback to track project state"""
#     ctx = dash.callback_context
#     trigger_id = ctx.triggered[0]['prop_id'].split('.')[0] if ctx.triggered else 'No triggers'
    
#     print(f"\n=== DEBUG PROJECT STATE ===")
#     print(f"Triggered by: {trigger_id}")
#     print(f"app.current_project_id: {getattr(app, 'current_project_id', 'Not set')}")
#     print(f"app.current_project_name: {getattr(app, 'current_project_name', 'Not set')}")
#     print(f"stored_project: {stored_project}")
    
#     return f"Project ID: {getattr(app, 'current_project_id', 'None')}"

# Callback to select and preview a file from the project files table
# Modified callback to select and preview a file from the project files table
@app.callback(
    [Output('upload-preview-table', 'data', allow_duplicate=True),
     Output('upload-preview-table', 'columns', allow_duplicate=True),
     Output('save-csv-filename-input', 'value')],
    Input('project-files-table', 'selected_rows'),
    [State('project-files-table', 'data')],
    prevent_initial_call=True
)
def preview_selected_project_file(selected_rows, file_data):
    if not selected_rows or not file_data:
        return [], [], ""
    
    # Get selected file info
    selected_file = file_data[selected_rows[0]]
    filename = selected_file['filename']
    file_type = selected_file['type']
    folder = selected_file.get('folder', '')  # Get the folder name if available
    
    # Only process certain file types
    valid_types = ['Raw Data', 'Trace Data', 'Work in Progress']
    if file_type not in valid_types or not hasattr(app, 'current_project_id') or not app.current_project_id:
        return [], [], ""
    
    # Determine the correct folder path based on file type
    if folder:
        # If folder is provided in file_data, use it
        folder_path = folder
    else:
        # Otherwise, determine folder based on file type
        if file_type == 'Raw Data':
            folder_path = "raw_data"
        elif file_type in ['Trace Data', 'Work in Progress']:
            folder_path = "trace_data"
        else:
            return [], [], ""
    
    # Load the file
    file_path = os.path.join(PROJECTS_DIRECTORY, app.current_project_id, folder_path, filename)
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return [], [], ""
    
    try:
        # Load the CSV file
        df = pd.read_csv(file_path)
        
        # Create a suggested filename for saving based on the current file
        suggested_filename = ""
        if file_type == 'Raw Data':
            # For raw data, suggest a filename without extension for processed version
            base_name = os.path.splitext(filename)[0]
            suggested_filename = base_name
        elif file_type == 'Work in Progress':
            # For WIP files, suggest the original name (without _wip)
            suggested_filename = filename.replace('_wip.csv', '')
        
        # Create columns for the preview table
        preview_columns = [{'name': col, 'id': col} for col in df.columns]
        
        # Return the data, columns, and suggested filename
        return df.to_dict('records'), preview_columns, suggested_filename
    except Exception as e:
        print(f"Error loading file for preview: {str(e)}")
        traceback.print_exc()
        return [], [], ""

# Sync Project Selection between tabs - fixed to match your ID changes
@app.callback(
    Output('event-project-selector', 'options'),
    Input('load-project-selector', 'options')
)
def sync_project_options(options):
    """Sync project options from Load Data tab to the Event Preparation tab"""
    return options

@app.callback(
    Output('csv-file-selector', 'options', allow_duplicate=True),
    [Input('event-project-selector', 'value'),
     Input('event-load-project-button', 'n_clicks')],
    prevent_initial_call=True
)
def update_event_csv_files(project_id, n_clicks):
    """Update the CSV file selector in Event Preparation when a project is loaded"""
    if not project_id:
        return []
    
    try:
        app.current_project_id = project_id
        trace_files = get_project_csv_files(project_id)
        file_options = [{'label': f, 'value': f} for f in trace_files if f.endswith('_trace.csv')]
        return file_options
    except Exception as e:
        print(f"Error getting trace files: {str(e)}")
        return []


# In[3]:


### EVENT PREPARATION TAB ###
# Set up directory for CSV files
CSV_DIRECTORY = os.path.join(os.getcwd(), "data")
print(f"Current working directory: {os.getcwd()}")
print(f"CSV directory: {CSV_DIRECTORY}")

# Make sure the directory exists
if not os.path.exists(CSV_DIRECTORY):
    os.makedirs(CSV_DIRECTORY)  # Create it if it doesn't exist

# Function to get CSV files from a specified directory
def get_trace_files(directory="."):
    """Find all trace files (CSV and Excel) in the specified directory with detailed debugging"""
    all_files = []
    
    print(f"Looking for trace files in directory: {directory}")
    
    # Search in specified directory
    if os.path.exists(directory):
        # Get CSV files
        csv_search_path = os.path.join(directory, "*.csv")
        csv_files = glob.glob(csv_search_path)
        print(f"  CSV files found: {len(csv_files)} - {[os.path.basename(f) for f in csv_files]}")
        
        # Get Excel files
        xlsx_search_path = os.path.join(directory, "*.xlsx")
        xlsx_files = glob.glob(xlsx_search_path)
        print(f"  XLSX files found: {len(xlsx_files)} - {[os.path.basename(f) for f in xlsx_files]}")
        
        # Add XLS files as well
        xls_search_path = os.path.join(directory, "*.xls")
        xls_files = glob.glob(xls_search_path)
        print(f"  XLS files found: {len(xls_files)} - {[os.path.basename(f) for f in xls_files]}")
        
        # Combine all file types
        dir_files = csv_files + xlsx_files + xls_files
        all_files.extend([os.path.basename(f) for f in dir_files])
    else:
        print(f"  Directory doesn't exist: {directory}")
    
    # Also search in current directory if different from specified directory
    current_dir = os.getcwd()
    if current_dir != os.path.abspath(directory):
        print(f"Also checking current directory: {current_dir}")
        current_csv_files = glob.glob("*.csv")
        current_xlsx_files = glob.glob("*.xlsx")
        current_xls_files = glob.glob("*.xls")
        
        print(f"  Current dir - CSV files: {current_csv_files}")
        print(f"  Current dir - XLSX files: {current_xlsx_files}")
        print(f"  Current dir - XLS files: {current_xls_files}")
        
        current_dir_files = current_csv_files + current_xlsx_files + current_xls_files
        all_files.extend([os.path.basename(f) for f in current_dir_files])
    
    # Remove duplicates
    unique_files = list(set(all_files))
    unique_files.sort()  # Sort alphabetically
    print(f"Final list of unique trace files: {unique_files}")
    
    return unique_files

def load_event_prep_data(filename):
    """Load CSV or Excel data for Event Preparation tab with minimal processing"""
    try:
        # Try multiple locations to find the file
        potential_paths = []
        
        # If we have a current project, check there first
        if hasattr(app, 'current_project_id') and app.current_project_id:
            project_trace_dir = os.path.join(PROJECTS_DIRECTORY, app.current_project_id, "trace_data")
            potential_paths.append(os.path.join(project_trace_dir, filename))
            
        # Then try other standard locations
        potential_paths.extend([
            filename,  # Try direct path
            os.path.join(CSV_DIRECTORY, filename),  # Then try in data directory
            os.path.join(os.getcwd(), filename)  # Then try in current directory
        ])
        
        df = None
        used_path = None
        is_excel = filename.lower().endswith(('.xlsx', '.xls'))
        
        for path in potential_paths:
            try:
                if os.path.exists(path):
                    print(f"Event Prep: Loading file from: {path}")
                    
                    if is_excel:
                        # Handle Excel file
                        excel_file = pd.ExcelFile(path)
                        df = pd.read_excel(excel_file, sheet_name='Data', engine='openpyxl')
                        
                        # Load selections if available
                        try:
                            # Initialize selection lists
                            app.event_selections = []
                            app.phase_selections = []
                            app.base_selections = []
                            app.peak_selections = []
                            
                            if 'Events' in excel_file.sheet_names:
                                event_df = pd.read_excel(excel_file, sheet_name='Events', engine='openpyxl')
                                app.event_selections = [tuple(x) for x in event_df.values]
                                print(f"Event Prep: Loaded {len(app.event_selections)} events from Excel")
                            
                            if 'Phases' in excel_file.sheet_names:
                                phase_df = pd.read_excel(excel_file, sheet_name='Phases', engine='openpyxl')
                                app.phase_selections = [tuple(x) for x in phase_df.values]
                                print(f"Event Prep: Loaded {len(app.phase_selections)} phases from Excel")
                            
                            if 'Bases' in excel_file.sheet_names:
                                base_df = pd.read_excel(excel_file, sheet_name='Bases', engine='openpyxl')
                                app.base_selections = [tuple(x) for x in base_df.values]
                                print(f"Event Prep: Loaded {len(app.base_selections)} base points from Excel")
                            
                            if 'Peaks' in excel_file.sheet_names:
                                peak_df = pd.read_excel(excel_file, sheet_name='Peaks', engine='openpyxl')
                                app.peak_selections = [tuple(x) for x in peak_df.values]
                                print(f"Event Prep: Loaded {len(app.peak_selections)} peak points from Excel")
                                
                        except Exception as e:
                            print(f"Event Prep: Could not load selection data from Excel: {str(e)}")
                            # We'll fall back to reconstruction method in the load_and_display_csv callback
                    else:
                        # Regular CSV file
                        df = pd.read_csv(path, low_memory=False)
                        
                    used_path = path
                    print(f"Event Prep: Successfully loaded data from: {path}")
                    print(f"Event Prep: Data shape: {df.shape}")
                    print(f"Event Prep: Columns: {df.columns.tolist()}")
                    break
            except Exception as e:
                print(f"Event Prep: Failed to load from {path}: {str(e)}")
                continue
        
        if df is None:
            print(f"Event Prep: Could not find file {filename} in any location")
            # Return empty DataFrame with required columns
            return pd.DataFrame(columns=['Offset', 'Time', 'Value', 'Annotations', 'Annotation', 
                                       'Event', 'Phase', 'VAS', 'Mnl_Base', 'Mnl_Peak', 'Pos_Delta', 'Neg_Delta'])
        
        if df.empty:
            print(f"Event Prep: File {filename} exists but contains no data")
            return pd.DataFrame(columns=['Offset', 'Time', 'Value', 'Annotations', 'Annotation', 
                                       'Event', 'Phase', 'VAS', 'Mnl_Base', 'Mnl_Peak', 'Pos_Delta', 'Neg_Delta'])
        
        # Check for required columns and add them if missing
        required_columns = ['Offset', 'Time', 'Value', 'Annotations', 'Annotation', 
                           'Event', 'Phase', 'VAS', 'Mnl_Base', 'Mnl_Peak', 'Pos_Delta', 'Neg_Delta']
        
        for col in required_columns:
            if col not in df.columns:
                print(f"Event Prep: Adding missing column: {col}")
                if col in ['Mnl_Base', 'Mnl_Peak', 'Pos_Delta', 'Neg_Delta']:
                    df[col] = None
                else:
                    df[col] = ''
        
        print("Event Prep: Sample data (first 3 rows):")
        try:
            print(df.head(3).to_string())
        except:
            print("Unable to print sample data")
        
        return df
    
    except Exception as e:
        print(f"Event Prep: Error loading CSV file {filename}: {str(e)}")
        traceback.print_exc()
        # Return empty DataFrame with required columns
        return pd.DataFrame(columns=['Offset', 'Time', 'Value', 'Annotations', 'Annotation', 
                                   'Event', 'Phase', 'VAS', 'Mnl_Base', 'Mnl_Peak', 'Pos_Delta', 'Neg_Delta'])

# Function to process event selection
def process_event_selection(df, selected_point, event_label):
    """Process event selection and update dataframe"""
    if selected_point is None or not event_label or df.empty:
        return df
    
    try:
        # Get the offset of the selected point
        offset = selected_point.get('points', [{}])[0].get('x', None)
        if offset is None:
            return df
        
        # Convert to milliseconds if the graph displays Time (seconds)
        if 'Time' in df.columns and df['Time'].equals(df['Offset']/1000):
            offset = offset * 1000
        
        # Find the closest row to the selected offset
        closest_idx = (df['Offset'] - offset).abs().idxmin()
        selected_offset = df.loc[closest_idx, 'Offset']
        
        # Store the selected event with its offset
        return selected_offset, event_label
    
    except Exception as e:
        print(f"Error processing event selection: {str(e)}")
        return None, None

# Function to update event column based on selected events
def update_event_column(df, event_selections):
    """Update the Event column based on event selections"""
    if df.empty or not event_selections:
        return df
    
    try:
        # Sort event selections by offset
        sorted_events = sorted(event_selections, key=lambda x: x[0])
        
        # Add maximum offset as the end boundary
        sorted_events.append((df['Offset'].max() + 1, ''))
        
        # Update Event column
        for i in range(len(sorted_events) - 1):
            start_offset = sorted_events[i][0]
            end_offset = sorted_events[i + 1][0]
            event_label = sorted_events[i][1]
            
            # Set event for all rows in range
            mask = (df['Offset'] >= start_offset) & (df['Offset'] < end_offset)
            df.loc[mask, 'Event'] = event_label
            
        return df
    
    except Exception as e:
        print(f"Error updating event column: {str(e)}")
        return df
    
# Function to process base selection
def process_base_selection(df, selected_point, base_label):
    """Process base selection and update dataframe"""
    if selected_point is None or not base_label or df.empty:
        return df
    
    try:
        # Get the offset of the selected point
        offset = selected_point.get('points', [{}])[0].get('x', None)
        if offset is None:
            return df
        
        # Convert to milliseconds if the graph displays Time (seconds)
        if 'Time' in df.columns and df['Time'].equals(df['Offset']/1000):
            offset = offset * 1000
        
        # Find the closest row to the selected offset
        closest_idx = (df['Offset'] - offset).abs().idxmin()
        selected_offset = df.loc[closest_idx, 'Offset']
        
        # Store the selected base with its offset
        return selected_offset, base_label
    
    except Exception as e:
        print(f"Error processing event selection: {str(e)}")
        return None, None

# Function to update base column based on selected events
def update_base_column(df, base_selections):
    """Update the base column based on peak selections"""
    if df.empty or not base_selections:
        return df
    
    try:
        # Sort base selections by offset
        sorted_bases = sorted(base_selections, key=lambda x: x[0])
        
        # Add maximum offset as the end boundary
        sorted_bases.append((df['Offset'].max() + 1, ''))
        
        # Update peak column
        for i in range(len(sorted_bases) - 1):
            base_offset = sorted_bases[i][0]
            base_time = sorted_bases[i + 1][0]
            base_label = sorted_bases[i][1]
            
            # Set base for row at base point
            mask = (df['Offset'] == base_offset) 
            df.loc[mask, 'Mnl_Base'] = base_label
            
        return df
    
    except Exception as e:
        print(f"Error updating base column: {str(e)}")
        traceback.print_exc()
        return df

def calculate_neg_delta(df, peak_point, app_base_selections):
            """
            Calculate neg_delta by finding the nearest base prior to the peak point
            and taking the difference between the peak value and base value.
            
            Args:
                df: DataFrame containing the data
                peak_point: Dictionary with offset, time, value of the peak point
                app_base_selections: List of base selections (offset, label, time, value)
                
            Returns:
                neg_delta value or None if no prior base exists
            """
            peak_offset = peak_point['offset']
            peak_value = peak_point['value']
            
            # Find all bases that occur prior to this peak
            prior_bases = [(offset, value) for offset, _, _, value in app_base_selections if offset < peak_offset]
            
            if not prior_bases:
                print(f"No prior bases found for peak at offset {peak_offset}")
                return None
            
            # Find the nearest prior base
            nearest_prior_base = max(prior_bases, key=lambda x: x[0])  # Get the one with highest offset (most recent)
            base_offset, base_value = nearest_prior_base
            
            # Calculate neg_delta as the absolute difference
            neg_delta = abs(peak_value - base_value)
            print(f"Calculated neg_delta={neg_delta} between peak({peak_offset}, {peak_value}) and base({base_offset}, {base_value})")
            
            return neg_delta

# Calculate Pos_Delta from Selcted Base - Prior Selcted Peak    
def calculate_pos_delta(df, base_point, app_peak_selections):
            """
            Calculate pos_delta by finding the nearest peak prior to the base point
            and taking the difference between the base value and peak value.
            
            Args:
                df: DataFrame containing the data
                base_point: Dictionary with offset, time, value of the base point
                app_peak_selections: List of peak selections (offset, label, time, value)
                
            Returns:
                pos_delta value or None if no prior peak exists
            """
            base_offset = base_point['offset']
            base_value = base_point['value']
            
            # Find all peaks that occur prior to this base
            prior_peaks = [(offset, value) for offset, _, _, value in app_peak_selections if offset < base_offset]
            
            if not prior_peaks:
                print(f"No prior peaks found for base at offset {base_offset}")
                return None
            
            # Find the nearest prior peak
            nearest_prior_peak = max(prior_peaks, key=lambda x: x[0])  # Get the one with highest offset (most recent)
            peak_offset, peak_value = nearest_prior_peak
            
            # Calculate pos_delta as the absolute difference
            pos_delta = abs(base_value - peak_value)
            print(f"Calculated pos_delta={pos_delta} between base({base_offset}, {base_value}) and peak({peak_offset}, {peak_value})")
            
            return pos_delta

# Function to process peak selection
def process_peak_selection(df, selected_point, peak_label):
    """Process peak selection and update dataframe"""
    if selected_point is None or not peak_label or df.empty:
        return df
    
    try:
        # Get the offset of the selected point
        offset = selected_point.get('points', [{}])[0].get('x', None)
        if offset is None:
            return df
        
        # Convert to milliseconds if the graph displays Time (seconds)
        if 'Time' in df.columns and df['Time'].equals(df['Offset']/1000):
            offset = offset * 1000
        
        # Find the closest row to the selected offset
        closest_idx = (df['Offset'] - offset).abs().idxmin()
        selected_offset = df.loc[closest_idx, 'Offset']
        
        # Store the selected peak with its offset
        return selected_offset, peak_label
    
    except Exception as e:
        print(f"Error processing event selection: {str(e)}")
        return None, None

# Function to update peak column based on selected events
def update_peak_column(df, peak_selections):
    """Update the peak column based on peak selections"""
    if df.empty or not peak_selections:
        return df
    
    try:
        # Sort peak selections by offset
        sorted_peaks = sorted(peak_selections, key=lambda x: x[0])
        
        # Add maximum offset as the end boundary
        sorted_peaks.append((df['Offset'].max() + 1, ''))
        
        # Update peak column
        for i in range(len(sorted_peaks) - 1):
            peak_offset = sorted_peaks[i][0]
            peak_time = sorted_peaks[i + 1][0]
            peak_label = sorted_peaks[i][1]
            
            # Set peak for row at peak point
            mask = (df['Offset'] == peak_offset) 
            df.loc[mask, 'Mnl_Peak'] = peak_label
            
        return df
    
    except Exception as e:
        print(f"Error updating peak column: {str(e)}")
        traceback.print_exc()
        return df

# Function to process phase selection
def process_phase_selection(df, selected_point, phase_label):
    """Process phase selection and update dataframe"""
    if selected_point is None or not phase_label or df.empty:
        return df
    
    try:
        # Get the offset of the selected point
        offset = selected_point.get('points', [{}])[0].get('x', None)
        if offset is None:
            return df
        
        # Convert to milliseconds if the graph displays Time (seconds)
        if 'Time' in df.columns and df['Time'].equals(df['Offset']/1000):
            offset = offset * 1000
        
        # Find the closest row to the selected offset
        closest_idx = (df['Offset'] - offset).abs().idxmin()
        selected_offset = df.loc[closest_idx, 'Offset']
        
        # Store the selected phase with its offset
        return selected_offset, phase_label
    
    except Exception as e:
        print(f"Error processing phase selection: {str(e)}")
        return None, None

# Function to update phase column based on selected phases
def update_phase_column(df, phase_selections):
    """Update the Phase column based on phase selections"""
    if df.empty or not phase_selections:
        return df
    
    try:
        # Sort phase selections by offset
        sorted_phases = sorted(phase_selections, key=lambda x: x[0])
        
        # Add maximum offset as the end boundary
        sorted_phases.append((df['Offset'].max() + 1, ''))
        
        # Update Phase column
        for i in range(len(sorted_phases) - 1):
            start_offset = sorted_phases[i][0]
            end_offset = sorted_phases[i + 1][0]
            phase_label = sorted_phases[i][1]
            
            # Set phase for all rows in range
            mask = (df['Offset'] >= start_offset) & (df['Offset'] < end_offset)
            df.loc[mask, 'Phase'] = phase_label
            
        return df
    
    except Exception as e:
        print(f"Error updating phase column: {str(e)}")
        return df
    
def save_graph_as_image(fig, filename='temp_graph.png'):
    """Save the plotly figure as a temporary image file"""
    try:
        # Save the figure as a PNG image
        fig.write_image(filename, scale=2)  # Scale=2 for higher resolution
        return filename
    except Exception as e:
        print(f"Error saving graph as image: {str(e)}")
        traceback.print_exc()
        return None

def save_processed_data_with_graph(df, fig, filename, output_path):
    """Save the processed data and graph to an Excel file"""
    if df.empty:
        return False, "No data to save"
    
    try:
        # Create a temporary graph image
        temp_image_path = save_graph_as_image(fig)
        if not temp_image_path:
            return False, "Error saving graph image"
        
        # Create Excel writer
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Save data to first worksheet
            df.to_excel(writer, sheet_name='Data', index=False)
            
            # Create a second worksheet for the graph
            workbook = writer.book
            worksheet = workbook.create_sheet('Graph')
            
            # Load the saved image
            img = Image(temp_image_path)
            
            # Add the image to the worksheet
            worksheet.add_image(img, 'A1')
            
            # Set column widths for better display
            for col in worksheet.columns:
                worksheet.column_dimensions[col[0].column_letter].width = 20
            
            # Set row heights for better display
            for i in range(1, 30):  # Adjust row heights for first 30 rows
                worksheet.row_dimensions[i].height = 20
            
        # Remove the temporary image file
        try:
            os.remove(temp_image_path)
        except:
            pass
        
        return True, f"Data and graph successfully saved to {output_path}"
    
    except Exception as e:
        print(f"Error saving processed data with graph: {str(e)}")
        traceback.print_exc()
        return False, f"Error saving data: {str(e)}"

# Function to create the data visualization
def create_event_preparation_graph(df):
    """Create the event preparation graph with annotations"""
    print(f"Creating graph with dataframe: empty={df.empty}, shape={df.shape if not df.empty else 'N/A'}")
    
    if df.empty:
        # Return empty figure with a message if no data
        empty_fig = go.Figure()
        empty_fig.update_layout(
            title='No Data Available',
            xaxis_title='Time (seconds)',
            yaxis_title='Value',
            annotations=[dict(
                text="Please load a CSV file",
                xref="paper",
                yref="paper",
                x=0.5,
                y=0.5,
                showarrow=False,
                font=dict(size=20)
            )]
        )
        return empty_fig

    try:
        # Create the main figure
        fig = go.Figure()
        
        # Get min and max y values for the full dataset for better coverage
        y_min = df['Value'].min()
        y_max = df['Value'].max()
        y_range = y_max - y_min
        
        # Add the main data line
        fig.add_trace(
            go.Scatter(
                x=df['Time'] if 'Time' in df.columns else df['Offset'],
                y=df['Value'],
                mode='lines',
                name='Data',
                line=dict(color='rgb(0, 140, 255)', width=2)
            )
        )    
        
        # Add event regions with transparent shading
        event_labels = df['Event'].unique()
        for event in event_labels:
            if event == '':
                continue
                
            event_data = df[df['Event'] == event]
            if len(event_data) > 0:
                # Get start and end points
                start_time = event_data['Time'].iloc[0] if 'Time' in event_data.columns else event_data['Offset'].iloc[0]
                end_time = event_data['Time'].iloc[-1] if 'Time' in event_data.columns else event_data['Offset'].iloc[-1]
                
                # Add transparent shape for the event region
                fig.add_shape(
                    type="rect",
                    x0=start_time,
                    x1=end_time,
                    y0=y_min - (y_range * 0.05),  # Extend a bit below
                    y1=y_max + (y_range * 0.05),  # Extend a bit above
                    fillcolor="rgba(0, 200, 0, 0.15)",  # Light green with transparency
                    line=dict(width=0),
                    layer="below"
                )
                
                # Add event label at the bottom
                fig.add_annotation(
                    x=(start_time + end_time) / 2,  # Center of the event
                    y=y_min - (y_range * 0.1),      # Below the bottom
                    text=f"{event}",
                    showarrow=False,
                    font=dict(size=12, color="rgb(0, 100, 0)"),
                    xanchor="center",
                    yanchor="top"
                )
        
        # Add phase regions with dotted vertical lines
        phase_labels = df['Phase'].unique()
        for phase in phase_labels:
            if phase == '':
                continue
                
            phase_data = df[df['Phase'] == phase]
            if len(phase_data) > 0:
                # Get start and end points
                start_time = phase_data['Time'].iloc[0] if 'Time' in phase_data.columns else phase_data['Offset'].iloc[0]
                end_time = phase_data['Time'].iloc[-1] if 'Time' in phase_data.columns else phase_data['Offset'].iloc[-1]
                
                # Print debugging info about phase lines
                print(f"Drawing phase lines: {phase} from {start_time} to {end_time}")
                
                # Add vertical dotted line at phase start - EXPLICITLY SET LINE PROPERTIES
                fig.add_shape(
                    type="line",
                    x0=start_time,
                    x1=start_time,
                    y0=y_min - (y_range * 0.05),
                    y1=y_max + (y_range * 0.05),
                    line=dict(
                        color="rgba(255, 165, 0, 0.8)",
                        width=2,
                        dash="dash"
                    ),
                    layer="above"  # Ensure it's drawn above other elements
                )
                
                # Add vertical dotted line at phase end - EXPLICITLY SET LINE PROPERTIES
                fig.add_shape(
                    type="line",
                    x0=end_time,
                    x1=end_time,
                    y0=y_min - (y_range * 0.05),
                    y1=y_max + (y_range * 0.05),
                    line=dict(
                        color="rgba(255, 165, 0, 0.8)",
                        width=2,
                        dash="dash"
                    ),
                    layer="above"  # Ensure it's drawn above other elements
                )
                
                # Add phase label at the top
                fig.add_annotation(
                    x=(start_time + end_time) / 2,  # Center of the phase
                    y=y_max + (y_range * 0.1),      # Above the top
                    text=f"Phase: {phase}",
                    showarrow=False,
                    font=dict(size=12, color="rgb(255, 140, 0)"),
                    xanchor="center",
                    yanchor="bottom"
                )
        
        # Add annotations with varied positions as requested
        annotated_points = df[df['Annotations'].notna() & (df['Annotations'] != '')]
        if not annotated_points.empty:
            # Create a pattern that repeats every 4 annotations
            for idx, row in annotated_points.iterrows():
                x_val = row['Time'] if 'Time' in row else row['Offset']
                y_val = row['Value']
                
                # Calculate longer line offsets for more visible annotation lines
                x_offset = 0.05 * (df['Time'].max() - df['Time'].min() if 'Time' in df.columns else df['Offset'].max() - df['Offset'].min())
                y_offset = 0.05 * y_range
                
                # Pattern position based on idx % 4
                pattern_idx = idx % 4
                
                if pattern_idx == 0:
                    # First in pattern: below line, 45 degrees
                    ax = x_val 
                    ay = y_val - (y_offset * 4.0)
                elif pattern_idx == 1:
                    # Second in pattern: above line, 225 degrees
                    ax = x_val 
                    ay = y_val + (y_offset * 5.0)
                elif pattern_idx == 2:
                    # Third in pattern: more above line, 45 degrees
                    ax = x_val 
                    ay = y_val - (y_offset * 9.0)  # Higher than the second
                elif pattern_idx == 3:
                    # Fourth in pattern: more below line, 225 degrees
                    ax = x_val 
                    ay = y_val + (y_offset * 7.5)  # Lower than the first
                
                # Add annotation with arrow pointing at the pattern determined above
                fig.add_annotation(
                    x=x_val,
                    y=y_val,  # Point to the actual data point
                    text=row['Annotations'],
                    showarrow=True,
                    arrowhead=2,
                    arrowsize=0.5,
                    arrowwidth=1,
                    arrowcolor="rgba(255, 0, 0, 0.7)",
                    standoff=0,
                    axref="x",
                    ayref="y",
                    ax=ax,
                    ay=ay,
                    font=dict(size=10, color="rgb(255, 0, 0)"),
                )
        
        # Add base points with light purple color
        base_points = df.dropna(subset=['Mnl_Base'])
        if not base_points.empty:
            # Create display text with pos_delta values
            text_labels = []
            for _, row in base_points.iterrows():
                text_labels.append("B")
            
            # Add base points
            fig.add_trace(
                go.Scatter(
                    x=base_points['Time'] if 'Time' in base_points.columns else base_points['Offset'],
                    y=base_points['Value'],  # Use Value column instead of Mnl_Base for y-coordinate
                    mode='markers+text',
                    text=text_labels,
                    textposition='bottom center',
                    marker=dict(
                        color='rgb(180, 160, 220)',  # Light purple
                        size=10,
                        symbol='circle'
                    ),
                    name='Base Points'
                )
            )
            
            # Add pos_delta values as separate text above the line
            for _, row in base_points.iterrows():
                pos_delta = row.get('Pos_Delta')
                if pos_delta is not None and not pd.isna(pos_delta):
                    x_val = row['Time'] if 'Time' in base_points.columns else row['Offset']
                    y_val = row['Value']  # Use Value instead of Mnl_Base for positioning
                    
                    fig.add_annotation(
                        x=x_val,
                        y=y_val + (y_range * 0.05),  # Position above the point
                        text=f"Δ:{pos_delta:.2f}",
                        showarrow=False,
                        font=dict(size=10, color="rgb(180, 160, 220)"),
                        xanchor="center",
                        yanchor="bottom"
                    )

        # Add peak points with color coding based on neg_delta
        peak_points = df.dropna(subset=['Mnl_Peak'])
        if not peak_points.empty:
            # Create a color list based on neg_delta values
            colors = []
            for _, row in peak_points.iterrows():
                # Get neg_delta value
                neg_delta = row.get('Neg_Delta')
                
                if neg_delta is None or pd.isna(neg_delta):
                    colors.append('rgb(178, 34, 34)')  # Default Firebrick Red
                else:
                    # Color based on absolute delta ranges
                    abs_delta = abs(neg_delta)
                    if abs_delta >= 10.000:
                        colors.append('rgb(0, 0, 0)')  # Black
                    elif 6.000 <= abs_delta <= 9.999:
                        colors.append('rgb(255, 0, 0)')  # Red
                    elif 3.000 <= abs_delta <= 5.999:
                        colors.append('rgb(255, 165, 0)')  # Orange
                    elif 0.050 <= abs_delta <= 2.999:
                        colors.append('rgb(255, 255, 0)')  # Yellow
                    else:  # < 0.050
                        colors.append('rgb(0, 0, 255)')  # Blue
            
            # Add peak points with simple 'P' label
            fig.add_trace(
                go.Scatter(
                    x=peak_points['Time'] if 'Time' in peak_points.columns else peak_points['Offset'],
                    y=peak_points['Value'],  # Use Value column instead of Mnl_Peak for y-coordinate
                    mode='markers+text',
                    text=["P" for _ in range(len(peak_points))],
                    textposition='top center',
                    marker=dict(
                        color=colors,
                        size=10,
                        symbol='circle'
                    ),
                    name='Peak Points'
                )
            )
            
            # Add neg_delta values as separate text below the line
            for _, row in peak_points.iterrows():
                neg_delta = row.get('Neg_Delta')
                if neg_delta is not None and not pd.isna(neg_delta):
                    x_val = row['Time'] if 'Time' in peak_points.columns else row['Offset']
                    y_val = row['Value']  # Use Value instead of Mnl_Peak for positioning
                    
                    # Use black text for all neg_delta values for better readability
                    fig.add_annotation(
                        x=x_val,
                        y=y_val - (y_range * 0.05),  # Position below the point
                        text=f"Δ:{neg_delta:.2f}",
                        showarrow=False,
                        font=dict(size=10, color="rgb(0, 0, 0)"),  # Always use black text
                        xanchor="center",
                        yanchor="top"
                    )

        # Customize layout
        x_title = 'Time (seconds)' if 'Time' in df.columns else 'Offset'
        fig.update_layout(
            title='Event Preparation',
            xaxis_title=x_title,
            yaxis_title='Value',
            hovermode='closest',
            showlegend=True,
            plot_bgcolor='white',
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1
            )
        )
        
        # Customize axes
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

        print(f"Successfully created graph with {len(fig.data)} traces")
        return fig
        
    except Exception as e:
        print(f"Error creating event preparation graph: {str(e)}")
        traceback.print_exc()

        # Return a minimal working figure in case of errors
        error_fig = go.Figure()
        error_fig.update_layout(
            title='Error Creating Graph',
            xaxis_title='Time',
            yaxis_title='Value',
            annotations=[dict(
                text=f"Error: {str(e)}",
                xref="paper",
                yref="paper",
                x=0.5,
                y=0.5,
                showarrow=False,
                font=dict(size=14, color="red")
            )]
        )
        return error_fig

def update_base_point(df, base_point, base_label):
    """Update the base column for the row of the base point"""
    if df.empty or not base_point or not base_label:
        return df
    
    try:
        # Get base point information
        base_offset = base_point['offset']
        base_value = base_point['value']
        
        # Set base value for the row at base offset
        mask = (df['Offset'] == base_offset)
        df.loc[mask, 'Mnl_Base'] = base_value
        
        print(f"Updated Base column: offset {base_offset} with value '{base_value}' and label '{base_label}'")
        print(f"Number of rows updated: {mask.sum()}")
        
        return df
    
    except Exception as e:
        print(f"Error updating base point: {str(e)}")
        traceback.print_exc()
        return df
    
def update_peak_point(df, peak_point, peak_label):
    """Update the peak column for the row of the peak point"""
    if df.empty or not peak_point or not peak_label:
        return df
    
    try:
        # Get peak point information
        peak_offset = peak_point['offset']
        peak_value = peak_point['value']
        
        # Set peak value for the row at peak offset
        mask = (df['Offset'] == peak_offset)
        df.loc[mask, 'Mnl_Peak'] = peak_value
        
        print(f"Updated Peak column: offset {peak_offset} with value '{peak_value}' and label '{peak_label}'")
        print(f"Number of rows updated: {mask.sum()}")
        
        return df
    
    except Exception as e:
        print(f"Error updating peak point: {str(e)}")
        traceback.print_exc()
        return df
    
def update_base(n_clicks, base_index, new_label):
    """Update the selected base with new values and recalculate pos_delta"""
    if not n_clicks or base_index is None:
        return dash.no_update, dash.no_update, dash.no_update, dash.no_update
            
    try:
        # Convert to integer index if needed
        if isinstance(base_index, str):
            base_index = int(base_index)
                
        if base_index < 0 or base_index >= len(app.base_selections):
            return html.Div("Invalid base selection", style={'color': 'red'}), dash.no_update, dash.no_update, dash.no_update
                
        if not new_label:
            return html.Div("Base label cannot be empty", style={'color': 'red'}), dash.no_update, dash.no_update, dash.no_update
                
        # Get original base data
        old_base_offset, old_base_label, old_base_time, old_base_value = app.base_selections[base_index]
            
        # Prepare new base data
        if app.base_edit_point:
            new_base_offset = app.base_edit_point['offset']
            new_base_time = app.base_edit_point['time']
            new_base_value = app.base_edit_point['value']
        else:
            new_base_offset = old_base_offset
            new_base_time = old_base_time
            new_base_value = old_base_value
                
        # Update the base value in the dataframe
        if not app.event_data.empty:
            # First, ensure the Mnl_Base column exists
            if 'Mnl_Base' not in app.event_data.columns:
                app.event_data['Mnl_Base'] = None

            # Clear the old base point
            mask = (app.event_data['Offset'] == old_base_offset)
            rows_affected = mask.sum()
            print(f"Clearing old base point at offset {old_base_offset}, rows affected: {rows_affected}")
            app.event_data.loc[mask, 'Mnl_Base'] = None
            app.event_data.loc[mask, 'Pos_Delta'] = None
                
            # Set the new base value
            mask = (app.event_data['Offset'] == new_base_offset)
            rows_affected = mask.sum()
            print(f"Setting new base point at offset {new_base_offset}, rows affected: {rows_affected}")
            app.event_data.loc[mask, 'Mnl_Base'] = new_base_value
                
            # Calculate pos_delta if there's a prior peak
            if app.base_edit_point:
                point_for_delta = app.base_edit_point
            else:
                point_for_delta = {'offset': new_base_offset, 'time': new_base_time, 'value': new_base_value}
                
            pos_delta = calculate_pos_delta(app.event_data, point_for_delta, app.peak_selections)
            if pos_delta is not None:
                app.event_data.loc[mask, 'Pos_Delta'] = pos_delta
                
            # Update the base selection
            app.base_selections[base_index] = (new_base_offset, new_label, new_base_time, new_base_value)
            print(f"Updated base selection: {app.base_selections[base_index]}")
                
            # Create success message with pos_delta if available
            if pos_delta is not None:
                status_text = f"Updated base '{new_label}' at {new_base_time:.2f}s, value: {new_base_value:.2f}, Pos_Delta: {pos_delta:.2f}"
            else:
                status_text = f"Updated base '{new_label}' at {new_base_time:.2f}s, value: {new_base_value:.2f}, No prior peak found"
                
            status = html.Div(status_text, style={'color': 'green'})
                
            # Update base list display with pos_delta values
            base_list_items = []
            for offset, label, time, value in sorted(app.base_selections, key=lambda x: x[0]):
                # Get pos_delta from the dataframe for this offset
                pos_delta_value = None
                if 'Pos_Delta' in app.event_data.columns:
                    mask = (app.event_data['Offset'] == offset)
                    if mask.any() and not app.event_data.loc[mask, 'Pos_Delta'].isnull().all():
                        pos_delta_value = app.event_data.loc[mask, 'Pos_Delta'].iloc[0]
                
                # Create list item text with pos_delta if available
                if pos_delta_value is not None:
                    list_text = f"At {time:.2f}s, Value: {value:.2f}, Pos_Delta: {pos_delta_value:.2f}"
                else:
                    list_text = f"At {time:.2f}s, Value: {value:.2f}"
                
                base_list_items.append(html.Li([
                    html.Strong(f"{label}"), 
                    list_text
                ]))
            
            base_list = html.Div([html.Ul(base_list_items)])
                
            # Update the graph
            updated_fig = create_event_preparation_graph(app.event_data)
                
            # Reset base edit point
            app.base_edit_point = None
                
            return status, base_list, updated_fig, app.event_data.to_dict('records')
        
        return html.Div("Error: No data loaded", style={'color': 'red'}), dash.no_update, dash.no_update, dash.no_update
        
    except Exception as e:
        print(f"Error updating base: {str(e)}")
        traceback.print_exc()
        return html.Div(f"Error: {str(e)}", style={'color': 'red'}), dash.no_update, dash.no_update, dash.no_update
    
def update_peak(n_clicks, peak_index, new_label):
    """Update the selected peak with new values and recalculate neg_delta"""
    if not n_clicks or peak_index is None:
        return dash.no_update, dash.no_update, dash.no_update, dash.no_update
            
    try:
        print(f"Updating peak: index={peak_index}, new_label={new_label}")
        
        # Convert to integer index if needed
        if isinstance(peak_index, str):
            peak_index = int(peak_index)
                
        if peak_index < 0 or peak_index >= len(app.peak_selections):
            print(f"Invalid peak index: {peak_index}, selections: {app.peak_selections}")
            return html.Div("Invalid peak selection", style={'color': 'red'}), dash.no_update, dash.no_update, dash.no_update
                
        if not new_label:
            return html.Div("Peak label cannot be empty", style={'color': 'red'}), dash.no_update, dash.no_update, dash.no_update
        
        # Debug print
        print(f"app.peak_selections: {app.peak_selections}")
        print(f"peak_index: {peak_index}")
        
        # Get original peak data
        old_peak_offset, old_peak_label, old_peak_time, old_peak_value = app.peak_selections[peak_index]
        print(f"Old peak data: offset={old_peak_offset}, label={old_peak_label}, time={old_peak_time}, value={old_peak_value}")
            
        # Prepare new peak data
        if app.peak_edit_point:
            new_peak_offset = app.peak_edit_point['offset']
            new_peak_time = app.peak_edit_point['time']
            new_peak_value = app.peak_edit_point['value']
            print(f"New peak point selected: offset={new_peak_offset}, time={new_peak_time}, value={new_peak_value}")
        else:
            new_peak_offset = old_peak_offset
            new_peak_time = old_peak_time
            new_peak_value = old_peak_value
            print(f"Using old peak data (no new point selected)")
                
        # Update the peak value in the dataframe
        if not app.event_data.empty:
            # Check if 'Mnl_Peak' column exists
            if 'Mnl_Peak' not in app.event_data.columns:
                print("'Mnl_Peak' column does not exist in the dataframe")
                app.event_data['Mnl_Peak'] = None
                
            # Clear the old peak point
            mask = (app.event_data['Offset'] == old_peak_offset)
            old_rows_affected = mask.sum()
            print(f"Clearing old peak point: offset={old_peak_offset}, rows affected={old_rows_affected}")
            app.event_data.loc[mask, 'Mnl_Peak'] = None
            app.event_data.loc[mask, 'Neg_Delta'] = None
                
            # Set the new peak value
            mask = (app.event_data['Offset'] == new_peak_offset)
            new_rows_affected = mask.sum()
            print(f"Setting new peak point: offset={new_peak_offset}, rows affected={new_rows_affected}")
            app.event_data.loc[mask, 'Mnl_Peak'] = new_peak_value
                
            # Calculate neg_delta if there's a prior base
            if app.peak_edit_point:
                point_for_delta = app.peak_edit_point
            else:
                point_for_delta = {'offset': new_peak_offset, 'time': new_peak_time, 'value': new_peak_value}
                
            neg_delta = calculate_neg_delta(app.event_data, point_for_delta, app.base_selections)
            if neg_delta is not None:
                app.event_data.loc[mask, 'Neg_Delta'] = neg_delta
                
            # Update the peak selection
            app.peak_selections[peak_index] = (new_peak_offset, new_label, new_peak_time, new_peak_value)
            print(f"Updated peak selection: {app.peak_selections[peak_index]}")
                
            # Create success message with neg_delta if available
            if neg_delta is not None:
                status_text = f"Updated peak '{new_label}' at {new_peak_time:.2f}s, value: {new_peak_value:.2f}, Neg_Delta: {neg_delta:.2f}"
            else:
                status_text = f"Updated peak '{new_label}' at {new_peak_time:.2f}s, value: {new_peak_value:.2f}, No prior base found"
                
            status = html.Div(status_text, style={'color': 'green'})
                
            # Update peak list display with neg_delta values
            peak_list_items = []
            for offset, label, time, value in sorted(app.peak_selections, key=lambda x: x[0]):
                # Get neg_delta from the dataframe for this offset
                neg_delta_value = None
                if 'Neg_Delta' in app.event_data.columns:
                    mask = (app.event_data['Offset'] == offset)
                    if mask.any() and not app.event_data.loc[mask, 'Neg_Delta'].isnull().all():
                        neg_delta_value = app.event_data.loc[mask, 'Neg_Delta'].iloc[0]
                
                # Create list item text with neg_delta if available
                if neg_delta_value is not None:
                    list_text = f"At {time:.2f}s, Value: {value:.2f}, Neg_Delta: {neg_delta_value:.2f}"
                else:
                    list_text = f"At {time:.2f}s, Value: {value:.2f}"
                
                peak_list_items.append(html.Li([
                    html.Strong(f"{label}"), 
                    list_text
                ]))
            
            peak_list = html.Div([html.Ul(peak_list_items)])
                
            # Update the graph
            updated_fig = create_event_preparation_graph(app.event_data)
                
            # Reset peak edit point
            app.peak_edit_point = None
                
            return status, peak_list, updated_fig, app.event_data.to_dict('records')
        
        print("Error: app.event_data is empty")
        return html.Div("Error: No data loaded", style={'color': 'red'}), dash.no_update, dash.no_update, dash.no_update
        
    except Exception as e:
        print(f"Error updating peak: {str(e)}")
        traceback.print_exc()
        return html.Div(f"Error: {str(e)}", style={'color': 'red'}), dash.no_update, dash.no_update, dash.no_update

def update_event_range(df, start_point, end_point, event_label):
    """Update the Event column for all rows between start and end offsets"""
    if df.empty or not start_point or not end_point or not event_label:
        return df
    
    try:
        # Sort start/end to ensure start < end
        start_offset = min(start_point['offset'], end_point['offset'])
        end_offset = max(start_point['offset'], end_point['offset'])
        
        # Set event label for all rows in range - use inclusive range
        mask = (df['Offset'] >= start_offset) & (df['Offset'] <= end_offset)
        df.loc[mask, 'Event'] = event_label
        
        print(f"Updated Event column: {start_offset} to {end_offset} with label '{event_label}'")
        print(f"Number of rows updated: {mask.sum()}")
        
        return df
    
    except Exception as e:
        print(f"Error updating event range: {str(e)}")
        traceback.print_exc()
        return df

def update_phase_range(df, start_point, end_point, phase_label):
    """Update the Phase column for all rows between start and end offsets"""
    if df.empty or not start_point or not end_point or not phase_label:
        return df
    
    try:
        # Sort start/end to ensure start < end
        start_offset = min(start_point['offset'], end_point['offset'])
        end_offset = max(start_point['offset'], end_point['offset'])
        
        # Set phase label for all rows in range - use inclusive range
        mask = (df['Offset'] >= start_offset) & (df['Offset'] <= end_offset)
        df.loc[mask, 'Phase'] = phase_label
        
        print(f"Updated Phase column: {start_offset} to {end_offset} with label '{phase_label}'")
        print(f"Number of rows updated: {mask.sum()}")
        
        return df
    
    except Exception as e:
        print(f"Error updating phase range: {str(e)}")
        traceback.print_exc()
        return df

def clear_selection_in_range(df, selected_rows, column_name):
    """Clear the specified column for the selected rows"""
    if df.empty or not selected_rows:
        return df
    
    try:
        # Get indices of selected rows
        selected_indices = [row['Offset'] for row in selected_rows]
        
        # Create a mask for the selected rows
        mask = df['Offset'].isin(selected_indices)
        
        # Clear the specified column for selected rows
        df.loc[mask, column_name] = ''
        
        return df
    
    except Exception as e:
        print(f"Error clearing {column_name}: {str(e)}")
        traceback.print_exc()
        return df

# Function to update VAS column based on slider selections
def update_vas_column(df, selected_point, vas_value):
    """Update the VAS column based on slider selection"""
    if selected_point is None or vas_value is None or df.empty:
        return df
    
    try:
        # Get the offset of the selected point
        offset = selected_point.get('points', [{}])[0].get('x', None)
        if offset is None:
            return df
        
        # Convert to milliseconds if the graph displays Time (seconds)
        if 'Time' in df.columns and df['Time'].equals(df['Offset']/1000):
            offset = offset * 1000
        
        # Find the closest row to the selected offset
        closest_idx = (df['Offset'] - offset).abs().idxmin()
        selected_offset = df.loc[closest_idx, 'Offset']
        
        # Find the next event or phase change or end of data
        event_changes = df[df['Event'].shift() != df['Event']]['Offset'].tolist()
        phase_changes = df[df['Phase'].shift() != df['Phase']]['Offset'].tolist()
        all_changes = sorted([o for o in event_changes + phase_changes if o > selected_offset])
        
        end_offset = df['Offset'].max() + 1
        if all_changes:
            end_offset = all_changes[0]
        
        # Set VAS value for all rows in range
        mask = (df['Offset'] >= selected_offset) & (df['Offset'] < end_offset)
        df.loc[mask, 'VAS'] = vas_value
        
        return df
    
    except Exception as e:
        print(f"Error updating VAS column: {str(e)}")
        return df

def save_to_project_peak_data(df, fig, filename, project_id):
    """
    Save processed data and graph to an Excel file in the project's peak_data folder,
    with "_manual" appended to the filename.
    
    Parameters:
    -----------
    df : pandas.DataFrame
        The dataframe containing the processed data
    fig : plotly.graph_objects.Figure
        The figure to save as a graph
    filename : str
        The base filename (without extension)
    project_id : str
        The ID of the current project
        
    Returns:
    --------
    tuple
        (success, message) where success is a boolean and message is a string
    """
    if df.empty:
        return False, "No data to save"
    
    if not project_id:
        return False, "No project selected"
    
    try:
        # Create the peak_data directory path
        peak_data_dir = os.path.join(PROJECTS_DIRECTORY, project_id, "peak_data")
        
        # Ensure the directory exists
        if not os.path.exists(peak_data_dir):
            os.makedirs(peak_data_dir)
            print(f"Created peak_data directory: {peak_data_dir}")
        
        # Create output filename with "_manual" suffix
        base_name = os.path.splitext(filename)[0]
        output_filename = f"{base_name}_manual.xlsx"
        output_path = os.path.join(peak_data_dir, output_filename)
        
        print(f"Saving data to: {output_path}")
        
        # Create a temporary graph image
        temp_image_path = save_graph_as_image(fig)
        if not temp_image_path:
            return False, "Error saving graph image"
        
        # Create Excel writer
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Save data to first worksheet
            df.to_excel(writer, sheet_name='Data', index=False)
            
            # Create a second worksheet for the graph
            workbook = writer.book
            worksheet = workbook.create_sheet('Graph')
            
            # Load the saved image
            img = Image(temp_image_path)
            
            # Add the image to the worksheet
            worksheet.add_image(img, 'A1')
            
            # Set column widths for better display
            for col in worksheet.columns:
                worksheet.column_dimensions[col[0].column_letter].width = 20
            
            # Set row heights for better display
            for i in range(1, 30):  # Adjust row heights for first 30 rows
                worksheet.row_dimensions[i].height = 20
        
        # Remove the temporary image file
        try:
            os.remove(temp_image_path)
        except Exception as e:
            print(f"Warning: Could not remove temporary image file: {str(e)}")
        
        print(f"Successfully saved data to {output_path}")
        return True, f"Data and graph successfully saved to project's peak_data folder as {output_filename}"
    
    except Exception as e:
        print(f"Error saving processed data with graph: {str(e)}")
        traceback.print_exc()
        return False, f"Error saving data: {str(e)}"

# Function to save the modified dataframe
def save_processed_data(df, original_filename):
    """Save the processed data to an Excel file"""
    if df.empty:
        return False, "No data to save"
    
    try:
        # Create output filename
        base_name = os.path.splitext(original_filename)[0]
        output_filename = f"{base_name}_processed.xlsx"
        output_path = os.path.join(CSV_DIRECTORY, output_filename)
        
        # Save to Excel
        df.to_excel(output_path, index=False)
        
        return True, f"Data successfully saved to {output_filename}"
    
    except Exception as e:
        print(f"Error saving processed data: {str(e)}")
        return False, f"Error saving data: {str(e)}"

# Define the Event Preparation tab layout
event_preparation_tab = dcc.Tab(
    label='Event Preparation',
    children=[
        html.Div([
            html.H1('Event Preparation', style={'textAlign': 'center'}),
            
            # Project selection section (added to integrate with Load Data tab)
            html.Div([
                html.H3('Project Selection', style={'marginBottom': '10px'}),
                html.Div([
                    html.Label('Select Project:', style={'marginRight': '10px', 'fontWeight': 'bold'}),
                    dcc.Dropdown(
                        id='event-project-selector',
                        options=[],
                        placeholder="Select a project",
                        style={'width': '300px', 'marginRight': '10px'}
                    ),
                    html.Button(
                        'Load Project',
                        id='event-load-project-button',
                        style={
                            'padding': '5px 10px',
                            'backgroundColor': '#2196F3',
                            'color': 'white',
                            'border': 'none',
                            'borderRadius': '4px',
                            'cursor': 'pointer',
                            'marginRight': '10px'
                        }
                    ),
                ], style={'display': 'flex', 'alignItems': 'center', 'marginBottom': '10px'}),
                html.Div(id='event-load-project-status'),
            ], style={'marginBottom': '20px', 'padding': '10px', 'backgroundColor': '#f8f9fa', 'borderRadius': '5px'}),
            
            # File selection section
            html.Div([
                html.H3('Select Trace File', style={'marginBottom': '10px'}),  # CHANGED FROM 'Select CSV File'
                dcc.Dropdown(
                    id='csv-file-selector',
                    options=[],
                    placeholder="Select a trace file",  # UPDATED placeholder text
                    style={'width': '500px'}
                ),
                html.Button(
                    'Refresh Files',
                    id='refresh-files-button',
                    style={
                        'marginLeft': '10px',
                        'padding': '5px 10px',
                        'backgroundColor': '#f8f9fa',
                        'border': '1px solid #ddd',
                        'borderRadius': '4px',
                        'cursor': 'pointer'
                    }
                ),
                html.Div(id='csv-load-status')
            ], style={'marginBottom': '20px', 'display': 'flex', 'alignItems': 'center'}),
            
            # Work in Progress save button section
            html.Div([
                html.Button(
                    'Save Work in Progress',
                    id='save-wip-button',
                    style={
                        'padding': '10px 20px',
                        'backgroundColor': '#2196F3',  # Blue color to differentiate from final save
                        'color': 'white',
                        'border': 'none',
                        'borderRadius': '4px',
                        'cursor': 'pointer',
                        'fontSize': '16px',
                        'marginRight': '10px'
                    }
                ),
                html.Div(id='save-wip-status', style={'display': 'inline-block'})
            ], style={'marginBottom': '20px'}),
            
            # Graph for visualization
            dcc.Graph(
                id='event-prep-graph',
                figure=go.Figure(),
                style={'height': '600px'},
                config={'scrollZoom': True}
            ),
            
            # Selection status display
            html.Div([
                html.Div(id='selection-status', style={'padding': '10px', 'backgroundColor': '#f0f0f0', 'borderRadius': '5px'}),
            ], style={'marginTop': '10px', 'marginBottom': '10px'}),
            
            # FIRST ROW: Base (Left) and Event (Right)
            html.Div([
                # L1: Base Section (Top Left)
                html.Div([
                    html.H3('Manual Base Point', style={'marginBottom': '10px'}),
                    
                    html.Div([
                        dcc.Input(
                            id='base-point-label-input',
                            type='text',
                            placeholder='Enter base label',
                            style={'width': '200px', 'marginRight': '10px'}
                        ),
                    ], style={'display': 'flex', 'alignItems': 'center', 'marginBottom': '10px'}),
                    
                    html.Div([
                        html.Button(
                            'Select Base Point',
                            id='select-base-point-button',
                            style={
                                'padding': '5px 10px',
                                'backgroundColor': '#f8f9fa',
                                'border': '1px solid #ddd',
                                'borderRadius': '4px',
                                'cursor': 'pointer',
                                'marginRight': '10px'
                            }
                        ),
                        html.Div(id='base-point-status', style={'marginLeft': '10px'})
                    ], style={'display': 'flex', 'alignItems': 'center', 'marginBottom': '10px'}),        
                        
                    html.Button(
                        'Add Base Point',
                        id='add-base-point-button',
                        style={
                            'padding': '5px 10px',
                            'backgroundColor': '#4CAF50',
                            'color': 'white',
                            'border': 'none',
                            'borderRadius': '4px',
                            'cursor': 'pointer',
                            'marginRight': '10px'
                        }
                    ),
                        
                    html.Button(
                        'Clear Base Point',
                        id='clear-base-point-button',
                        style={
                            'padding': '5px 10px',
                            'backgroundColor': '#f44336',
                            'color': 'white',
                            'border': 'none',
                            'borderRadius': '4px',
                            'cursor': 'pointer'
                        }
                    ),        
                        
                    html.Div(id='base-point-add-status', style={'marginTop': '10px'}),
                    
                    # Display selected base point details
                    html.Div([
                        html.H4("Selected Base Point:", style={'marginBottom': '5px'}),
                        html.Div(id='base-point-display', style={'padding': '5px', 'backgroundColor': '#f5f5f5', 'borderRadius': '4px'})
                    ], style={'marginTop': '10px'}),
                    
                    # Base list with edit functionality
                    html.Div([
                        html.H4("Current Bases:"),
                        html.Div(id='base-point-list-display'),
                        
                        # Base edit controls
                        html.Div([
                            html.H4("Edit Base:", style={'marginTop': '15px'}),
                            dcc.Dropdown(
                                id='base-point-edit-selector',
                                options=[],
                                placeholder="Select base to edit",
                                style={'width': '100%', 'marginBottom': '10px'}
                            ),
                            dcc.Input(
                                id='base-point-edit-label',
                                type='text',
                                placeholder='New base label',
                                style={'width': '100%', 'marginBottom': '10px'}
                            ),
                            html.Div([
                                html.Button(
                                    'Select New Point',
                                    id='base-point-edit-button',
                                    style={
                                        'padding': '5px 10px',
                                        'backgroundColor': '#f8f9fa',
                                        'border': '1px solid #ddd',
                                        'borderRadius': '4px',
                                        'cursor': 'pointer',
                                        'marginRight': '10px'
                                    }
                                ),
                                html.Div(id='base-point-edit-status', style={'marginLeft': '10px'})
                            ], style={'display': 'flex', 'alignItems': 'center', 'marginBottom': '10px'}),
                            
                            html.Button(
                                'Update Base',
                                id='update-base-point-button',
                                style={
                                    'padding': '5px 10px',
                                    'backgroundColor': '#2196F3',
                                    'color': 'white',
                                    'border': 'none',
                                    'borderRadius': '4px',
                                    'cursor': 'pointer',
                                    'width': '100%'
                                }
                            ),
                            html.Div(id='base-point-update-status', style={'marginTop': '10px'})
                        ], style={'marginTop': '10px', 'padding': '10px', 'backgroundColor': '#f8f9fa', 'borderRadius': '5px'})
                    ], style={'marginTop': '10px'})
                ], style={'flex': '1', 'marginRight': '20px'}),
                
                # L2: Event Section (Top Right)
                html.Div([
                    html.H3('Add Events', style={'marginBottom': '10px'}),
                    html.Div([
                        dcc.Input(
                            id='event-label-input',
                            type='text',
                            placeholder='Enter event label',
                            style={'width': '200px', 'marginRight': '10px'}
                        ),
                    ], style={'display': 'flex', 'alignItems': 'center', 'marginBottom': '10px'}),
                    
                    html.Div([
                        html.Div([
                            html.Button(
                                'Select Start Point',
                                id='select-event-start-button',
                                style={
                                    'padding': '5px 10px',
                                    'backgroundColor': '#f8f9fa',
                                    'border': '1px solid #ddd',
                                    'borderRadius': '4px',
                                    'cursor': 'pointer',
                                    'marginRight': '10px'
                                }
                            ),
                            html.Div(id='event-start-status', style={'marginLeft': '10px'})
                        ], style={'display': 'flex', 'alignItems': 'center', 'marginBottom': '10px'}),
                        
                        html.Div([
                            html.Button(
                                'Select End Point',
                                id='select-event-end-button',
                                style={
                                    'padding': '5px 10px',
                                    'backgroundColor': '#f8f9fa',
                                    'border': '1px solid #ddd',
                                    'borderRadius': '4px',
                                    'cursor': 'pointer',
                                    'marginRight': '10px'
                                }
                            ),
                            html.Div(id='event-end-status', style={'marginLeft': '10px'})
                        ], style={'display': 'flex', 'alignItems': 'center', 'marginBottom': '10px'}),
                        
                        html.Button(
                            'Add Event',
                            id='add-event-button',
                            style={
                                'padding': '5px 10px',
                                'backgroundColor': '#4CAF50',
                                'color': 'white',
                                'border': 'none',
                                'borderRadius': '4px',
                                'cursor': 'pointer',
                                'marginRight': '10px'
                            }
                        ),
                        
                        html.Button(
                            'Clear Event Selection',
                            id='clear-event-button',
                            style={
                                'padding': '5px 10px',
                                'backgroundColor': '#f44336',
                                'color': 'white',
                                'border': 'none',
                                'borderRadius': '4px',
                                'cursor': 'pointer'
                            }
                        )
                    ], style={'display': 'flex', 'flexDirection': 'column'}),
                    
                    html.Div(id='event-add-status', style={'marginTop': '10px'}),
                    
                    # Event list with edit functionality
                    html.Div([
                        html.H4("Current Events:"),
                        html.Div(id='event-list-display'),
                        
                        # Event edit controls
                        html.Div([
                            html.H4("Edit Event:", style={'marginTop': '15px'}),
                            dcc.Dropdown(
                                id='event-edit-selector',
                                options=[],
                                placeholder="Select event to edit",
                                style={'width': '100%', 'marginBottom': '10px'}
                            ),
                            dcc.Input(
                                id='event-edit-label',
                                type='text',
                                placeholder='New event label',
                                style={'width': '100%', 'marginBottom': '10px'}
                            ),
                            html.Div([
                                html.Button(
                                    'Select New Start',
                                    id='event-edit-start-button',
                                    style={
                                        'padding': '5px 10px',
                                        'backgroundColor': '#f8f9fa',
                                        'border': '1px solid #ddd',
                                        'borderRadius': '4px',
                                        'cursor': 'pointer',
                                        'marginRight': '10px'
                                    }
                                ),
                                html.Div(id='event-edit-start-status', style={'marginLeft': '10px'})
                            ], style={'display': 'flex', 'alignItems': 'center', 'marginBottom': '10px'}),
                            html.Div([
                                html.Button(
                                    'Select New End',
                                    id='event-edit-end-button',
                                    style={
                                        'padding': '5px 10px',
                                        'backgroundColor': '#f8f9fa',
                                        'border': '1px solid #ddd',
                                        'borderRadius': '4px',
                                        'cursor': 'pointer',
                                        'marginRight': '10px'
                                    }
                                ),
                                html.Div(id='event-edit-end-status', style={'marginLeft': '10px'})
                            ], style={'display': 'flex', 'alignItems': 'center', 'marginBottom': '10px'}),
                            html.Button(
                                'Update Event',
                                id='update-event-button',
                                style={
                                    'padding': '5px 10px',
                                    'backgroundColor': '#2196F3',
                                    'color': 'white',
                                    'border': 'none',
                                    'borderRadius': '4px',
                                    'cursor': 'pointer',
                                    'width': '100%'
                                }
                            ),
                            html.Div(id='event-update-status', style={'marginTop': '10px'})
                        ], style={'marginTop': '10px', 'padding': '10px', 'backgroundColor': '#f8f9fa', 'borderRadius': '5px'})
                    ], style={'marginTop': '10px'})
                ], style={'flex': '1'})
            ], style={'display': 'flex', 'marginBottom': '20px'}),

            # SECOND ROW: Peak (Left) and Phase (Right)
            html.Div([
                # L3: Peak Section (Bottom Left)
                html.Div([
                    html.H3('Manual Peak Point', style={'marginBottom': '10px'}),

                    html.Div([
                        dcc.Input(
                            id='peak-point-label-input',
                            type='text',
                            placeholder='Enter peak label',
                            style={'width': '200px', 'marginRight': '10px'}
                        ),
                    ], style={'display': 'flex', 'alignItems': 'center', 'marginBottom': '10px'}),

                    html.Div([
                        html.Button(
                            'Select Peak Point',
                            id='select-peak-point-button',
                            style={
                                'padding': '5px 10px',
                                'backgroundColor': '#f8f9fa',
                                'border': '1px solid #ddd',
                                'borderRadius': '4px',
                                'cursor': 'pointer',
                                'marginRight': '10px'
                            }
                        ),        
                        html.Div(id='peak-point-status', style={'marginLeft': '10px'})
                    ], style={'display': 'flex', 'alignItems': 'center', 'marginBottom': '10px'}),        
                        
                    html.Button(
                        'Add Peak Point',
                        id='add-peak-point-button',
                        style={
                            'padding': '5px 10px',
                            'backgroundColor': '#4CAF50',
                            'color': 'white',
                            'border': 'none',
                            'borderRadius': '4px',
                            'cursor': 'pointer',
                            'marginRight': '10px'
                        }
                    ),
                        
                    html.Button(
                        'Clear Peak Point',
                        id='clear-peak-point-button',
                        style={
                            'padding': '5px 10px',
                            'backgroundColor': '#f44336',
                            'color': 'white',
                            'border': 'none',
                            'borderRadius': '4px',
                            'cursor': 'pointer'
                        }
                    ),    
                        
                    html.Div(id='peak-point-add-status', style={'marginTop': '10px'}),
                    
                    # Display selected peak point details
                    html.Div([
                        html.H4("Selected Peak Point:", style={'marginBottom': '5px'}),
                        html.Div(id='peak-point-display', style={'padding': '5px', 'backgroundColor': '#f5f5f5', 'borderRadius': '4px'})
                    ], style={'marginTop': '10px'}),

                    # Peak Point list with edit functionality
                    html.Div([
                        html.H4("Current Peaks:"),
                        html.Div(id='peak-point-list-display'),
                        
                        # Peak edit controls
                        html.Div([
                            html.H4("Edit Peak:", style={'marginTop': '15px'}),
                            dcc.Dropdown(
                                id='peak-point-edit-selector',
                                options=[],
                                placeholder="Select peak to edit",
                                style={'width': '100%', 'marginBottom': '10px'}
                            ),
                            dcc.Input(
                                id='peak-point-edit-label',
                                type='text',
                                placeholder='New peak label',
                                style={'width': '100%', 'marginBottom': '10px'}
                            ),
                            html.Div([
                                html.Button(
                                    'Select New Point',
                                    id='peak-point-edit-button',
                                    style={
                                        'padding': '5px 10px',
                                        'backgroundColor': '#f8f9fa',
                                        'border': '1px solid #ddd',
                                        'borderRadius': '4px',
                                        'cursor': 'pointer',
                                        'marginRight': '10px'
                                    }
                                ),
                                html.Div(id='peak-point-edit-status', style={'marginLeft': '10px'})
                            ], style={'display': 'flex', 'alignItems': 'center', 'marginBottom': '10px'}),
                            
                            html.Button(
                                'Update Peak',
                                id='update-peak-point-button',
                                style={
                                    'padding': '5px 10px',
                                    'backgroundColor': '#2196F3',
                                    'color': 'white',
                                    'border': 'none',
                                    'borderRadius': '4px',
                                    'cursor': 'pointer',
                                    'width': '100%'
                                }
                            ),
                            html.Div(id='peak-point-update-status', style={'marginTop': '10px'})
                        ], style={'marginTop': '10px', 'padding': '10px', 'backgroundColor': '#f8f9fa', 'borderRadius': '5px'})
                    ], style={'marginTop': '10px'})
                ], style={'flex': '1', 'marginRight': '20px'}),
                
                # L4: Phase Section (Bottom Right)
                html.Div([
                    html.H3('Add Phases', style={'marginBottom': '10px'}),
                    html.Div([
                        dcc.Input(
                            id='phase-label-input',
                            type='text',
                            placeholder='Enter phase label',
                            style={'width': '200px', 'marginRight': '10px'}
                        ),
                    ], style={'display': 'flex', 'alignItems': 'center', 'marginBottom': '10px'}),
                    
                    html.Div([
                        html.Div([
                            html.Button(
                                'Select Start Point',
                                id='select-phase-start-button',
                                style={
                                    'padding': '5px 10px',
                                    'backgroundColor': '#f8f9fa',
                                    'border': '1px solid #ddd',
                                    'borderRadius': '4px',
                                    'cursor': 'pointer',
                                    'marginRight': '10px'
                                }
                            ),
                            html.Div(id='phase-start-status', style={'marginLeft': '10px'})
                        ], style={'display': 'flex', 'alignItems': 'center', 'marginBottom': '10px'}),
                        
                        html.Div([
                            html.Button(
                                'Select End Point',
                                id='select-phase-end-button',
                                style={
                                    'padding': '5px 10px',
                                    'backgroundColor': '#f8f9fa',
                                    'border': '1px solid #ddd',
                                    'borderRadius': '4px',
                                    'cursor': 'pointer',
                                    'marginRight': '10px'
                                }
                            ),
                            html.Div(id='phase-end-status', style={'marginLeft': '10px'})
                        ], style={'display': 'flex', 'alignItems': 'center', 'marginBottom': '10px'}),
                        
                        html.Button(
                            'Add Phase',
                            id='add-phase-button',
                            style={
                                'padding': '5px 10px',
                                'backgroundColor': '#4CAF50',
                                'color': 'white',
                                'border': 'none',
                                'borderRadius': '4px',
                                'cursor': 'pointer',
                                'marginRight': '10px'
                            }
                        ),
                        
                        html.Button(
                            'Clear Phase Selection',
                            id='clear-phase-button',
                            style={
                                'padding': '5px 10px',
                                'backgroundColor': '#f44336',
                                'color': 'white',
                                'border': 'none',
                                'borderRadius': '4px',
                                'cursor': 'pointer'
                            }
                        )
                    ], style={'display': 'flex', 'flexDirection': 'column'}),
                    
                    html.Div(id='phase-add-status', style={'marginTop': '10px'}),
                    
                    # Phase list with edit functionality
                    html.Div([
                        html.H4("Current Phases:"),
                        html.Div(id='phase-list-display'),
                        
                        # Phase edit controls
                        html.Div([
                            html.H4("Edit Phase:", style={'marginTop': '15px'}),
                            dcc.Dropdown(
                                id='phase-edit-selector',
                                options=[],
                                placeholder="Select phase to edit",
                                style={'width': '100%', 'marginBottom': '10px'}
                            ),
                            dcc.Input(
                                id='phase-edit-label',
                                type='text',
                                placeholder='New phase label',
                                style={'width': '100%', 'marginBottom': '10px'}
                            ),
                            html.Div([
                                html.Button(
                                    'Select New Start',
                                    id='phase-edit-start-button',
                                    style={
                                        'padding': '5px 10px',
                                        'backgroundColor': '#f8f9fa',
                                        'border': '1px solid #ddd',
                                        'borderRadius': '4px',
                                        'cursor': 'pointer',
                                        'marginRight': '10px'
                                    }
                                ),
                                html.Div(id='phase-edit-start-status', style={'marginLeft': '10px'})
                            ], style={'display': 'flex', 'alignItems': 'center', 'marginBottom': '10px'}),
                            html.Div([
                                html.Button(
                                    'Select New End',
                                    id='phase-edit-end-button',
                                    style={
                                        'padding': '5px 10px',
                                        'backgroundColor': '#f8f9fa',
                                        'border': '1px solid #ddd',
                                        'borderRadius': '4px',
                                        'cursor': 'pointer',
                                        'marginRight': '10px'
                                    }
                                ),
                                html.Div(id='phase-edit-end-status', style={'marginLeft': '10px'})
                            ], style={'display': 'flex', 'alignItems': 'center', 'marginBottom': '10px'}),
                            html.Button(
                                'Update Phase',
                                id='update-phase-button',
                                style={
                                    'padding': '5px 10px',
                                    'backgroundColor': '#2196F3',
                                    'color': 'white',
                                    'border': 'none',
                                    'borderRadius': '4px',
                                    'cursor': 'pointer',
                                    'width': '100%'
                                }
                            ),
                            html.Div(id='phase-update-status', style={'marginTop': '10px'})
                        ], style={'marginTop': '10px', 'padding': '10px', 'backgroundColor': '#f8f9fa', 'borderRadius': '5px'})
                    ], style={'marginTop': '10px'})
                ], style={'flex': '1'})
            ], style={'display': 'flex', 'marginBottom': '20px'}),
            
            # VAS input section
            html.Div([
                html.H3('Add VAS Rating', style={'marginBottom': '10px'}),
                html.Div([
                    dcc.Slider(
                        id='vas-slider',
                        min=0,
                        max=10,
                        step=0.5,
                        marks={i: str(i) for i in range(11)},
                        value=0
                    ),
                    html.Button(
                        'Add VAS at Selected Point',
                        id='add-vas-button',
                        style={
                            'padding': '5px 10px',
                            'backgroundColor': '#f8f9fa',
                            'border': '1px solid #ddd',
                            'borderRadius': '4px',
                            'cursor': 'pointer',
                            'marginTop': '10px'
                        }
                    )
                ], style={'width': '100%', 'marginBottom': '10px'}),
                html.Div(id='vas-add-status')
            ], style={'marginBottom': '20px'}),

            # Data table section with clear buttons
            html.H3('Data Preview and Edit', style={'marginBottom': '10px'}),
            
            html.Div([
                html.Button(
                    'Clear Selected Event',
                    id='clear-selected-event-button',
                    style={
                        'padding': '5px 10px',
                        'backgroundColor': '#f8f9fa',
                        'border': '1px solid #ddd',
                        'borderRadius': '4px',
                        'cursor': 'pointer',
                        'marginRight': '10px'
                    }
                ),
                html.Button(
                    'Clear Selected Phase',
                    id='clear-selected-phase-button',
                    style={
                        'padding': '5px 10px',
                        'backgroundColor': '#f8f9fa',
                        'border': '1px solid #ddd',
                        'borderRadius': '4px',
                        'cursor': 'pointer',
                        'marginRight': '10px'
                    }
                ),

                html.Div(id='clear-status', style={'display': 'inline-block'})
            ], style={'marginBottom': '10px'}),
            
            dash_table.DataTable(
                id='event-prep-table',
                columns=[
                    {'name': 'Offset', 'id': 'Offset', 'type': 'numeric'},
                    {'name': 'Time', 'id': 'Time', 'type': 'numeric'},
                    {'name': 'Value', 'id': 'Value', 'type': 'numeric'},
                    {'name': 'Annotations', 'id': 'Annotations'},
                    {'name': 'Annotation', 'id': 'Annotation'},
                    {'name': 'Event', 'id': 'Event'},
                    {'name': 'Phase', 'id': 'Phase'},
                    {'name': 'VAS', 'id': 'VAS', 'type': 'numeric'},
                    {'name': 'Mnl_Base', 'id': 'Mnl_Base', 'type': 'numeric'},
                    {'name': 'Mnl_Peak', 'id': 'Mnl_Peak', 'type': 'numeric'},
                    {'name': 'Pos_Delta', 'id': 'Pos_Delta', 'type': 'numeric'},
                    {'name': 'Neg_Delta', 'id': 'Neg_Delta', 'type': 'numeric'}
                ],
                data=[],
                editable=True,
                filter_action='native',
                sort_action='native',
                page_size=15,
                style_table={'overflowX': 'auto', 'height': '400px', 'overflowY': 'auto'},
                style_cell={'textAlign': 'left', 'minWidth': '100px', 'maxWidth': '180px'},
                row_selectable='multi'
            ),
            
            # Save section with filename input
            html.Div([
                html.Div([
                    html.Label('Filename: ', style={'marginRight': '10px', 'fontWeight': 'bold'}),
                    dcc.Input(
                        id='save-filename-input',
                        type='text',
                        placeholder='Enter filename (without extension)',
                        style={'width': '300px', 'marginRight': '10px'}
                    ),
                ], style={'display': 'flex', 'alignItems': 'center', 'marginBottom': '10px'}),
                
                html.Button(
                    'Save Processed Data',
                    id='save-processed-button',
                    style={
                        'padding': '10px 20px',
                        'backgroundColor': '#4CAF50',
                        'color': 'white',
                        'border': 'none',
                        'borderRadius': '4px',
                        'cursor': 'pointer',
                        'fontSize': '16px'
                    }
                ),
                html.Div(id='save-status')
            ], style={'marginTop': '20px', 'marginBottom': '40px'})
        ], style={'padding': '20px'})
    ],
    style={'padding': '0 20px'}
)

# Store variables as attributes of the app
def register_event_preparation_callbacks(app):
    """Register callbacks for the Event Preparation tab"""
    
    # Initialize app variables
    app.event_selections = []  # Now store (start_offset, end_offset, label, start_time, end_time) tuples
    app.phase_selections = []  # Now store (start_offset, end_offset, label, start_time, end_time) tuples
    app.base_selections = []   # Now store (base_label, base_offset, base_point) tuples
    app.peak_selections = []   # Now store (peak_label, peak_offset, peak_point) tuples
    app.selected_point = None
    app.event_data = pd.DataFrame()
    app.original_filename = None
    
    # Variables for tracking selection points
    app.event_start_point = None
    app.event_end_point = None
    app.phase_start_point = None
    app.phase_end_point = None
    app.base_point = None
    app.peak_point = None
    app.selection_mode = None

    # Add status message variables
    app.event_start_status = ""
    app.event_end_status = ""
    app.phase_start_status = ""
    app.phase_end_status = ""
    app.base_point_status = ""
    app.peak_point_status = ""
    
    # New variables for edit mode
    app.event_edit_start_point = None
    app.event_edit_end_point = None
    app.phase_edit_start_point = None
    app.phase_edit_end_point = None
    app.event_edit_index = None
    app.phase_edit_index = None
    app.base_edit_point = None
    app.peak_edit_point = None

    # 1. Callback to update file dropdown options
    # Enhanced file refresh callback to highlight work-in-progress files
    @app.callback(
        Output('csv-file-selector', 'options'),
        [Input('refresh-files-button', 'n_clicks'),
        Input('event-load-project-button', 'n_clicks'),
        Input('save-wip-button', 'n_clicks')],  # Add save-wip-button as a trigger
        [State('event-project-selector', 'value')],
        prevent_initial_call=False
    )
    def refresh_files(refresh_clicks, load_project_clicks, save_wip_clicks, project_id):
        """Update the trace file options when Refresh Files button is clicked or files are saved"""
        ctx = dash.callback_context
        trigger_id = ctx.triggered[0]['prop_id'].split('.')[0] if ctx.triggered else 'initial'
        
        print(f"refresh_files called. Trigger: {trigger_id}")
        
        try:
            trace_files = []
            
            # If we have a current project ID set, use that project's trace_data folder
            if hasattr(app, 'current_project_id') and app.current_project_id:
                project_dir = os.path.join(PROJECTS_DIRECTORY, app.current_project_id, "trace_data")
                print(f"Looking for trace files in project directory: {project_dir}")
                
                if os.path.exists(project_dir):
                    # Get all files from the trace_data directory using our enhanced function
                    trace_files = get_trace_files(project_dir)
                else:
                    print(f"Project trace_data directory doesn't exist: {project_dir}")
            else:
                # Fallback to the default CSV directory if no project is selected
                print(f"No project selected, looking in default CSV directory: {CSV_DIRECTORY}")
                trace_files = get_trace_files(CSV_DIRECTORY)
            
            # Prepare options with special styling for work-in-progress files
            options = []
            for filename in trace_files:
                if '_wip' in filename.lower():
                    # Highlight work-in-progress files with a prefix and different style
                    options.append({
                        'label': f"🔄 {filename} (Work in Progress)",
                        'value': filename,
                        'title': "This is a previously saved work-in-progress file"
                    })
                else:
                    options.append({
                        'label': filename,
                        'value': filename
                    })
            
            # Sort options to show WIP files first, then alphabetically
            options.sort(key=lambda x: (0 if '_wip' in x['value'].lower() else 1, x['value'].lower()))
            
            print(f"Final dropdown options: {[o['value'] for o in options]}")
            return options
        except Exception as e:
            print(f"Error refreshing files: {str(e)}")
            traceback.print_exc()
            return []

    # 2. Callback to load CSV data and update graph
    @app.callback(
        [Output('event-prep-graph', 'figure'),
         Output('event-prep-table', 'data'),
         Output('csv-load-status', 'children')],
        Input('csv-file-selector', 'value'),
        prevent_initial_call=True
    )
    def load_and_display_csv(filename):
        """Load CSV data and update display components"""
        if not filename:
            empty_fig = go.Figure()
            empty_fig.update_layout(
                title='No CSV File Selected',
                annotations=[dict(
                    text="Please select a CSV file",
                    xref="paper",
                    yref="paper",
                    x=0.5,
                    y=0.5,
                    showarrow=False,
                    font=dict(size=20)
                )]
            )
            return empty_fig, [], html.Div("No file selected", style={'color': 'orange', 'marginLeft': '10px'})
        
        try:
            # Load the event prep data
            print(f"Loading CSV file: {filename}")
            df = load_event_prep_data(filename)
            
            if df.empty:
                print("Warning: Loaded DataFrame is empty")
                empty_fig = go.Figure()
                empty_fig.update_layout(
                    title='Empty Data',
                    annotations=[dict(
                        text=f"The file {filename} contains no valid data",
                        xref="paper",
                        yref="paper",
                        x=0.5,
                        y=0.5,
                        showarrow=False
                    )]
                )
                return empty_fig, [], html.Div(f"Warning: {filename} contains no data", style={'color': 'orange', 'marginLeft': '10px'})
            
            # Reset stored selections
            app.event_selections = []
            app.phase_selections = []
            app.base_selections = []   # Now store (base_label, base_offset, base_point) tuples
            app.peak_selections = []   # Now store (peak_label, peak_offset, peak_point) tuples
            app.selected_point = None
            app.event_start_point = None
            app.event_end_point = None
            app.phase_start_point = None
            app.phase_end_point = None
            app.base_point = None
            app.peak_point = None
            app.selection_mode = None
            app.event_data = df.copy()  # Ensure we're working with a copy
            app.original_filename = filename
            
            # Reset status messages
            app.event_start_status = ""
            app.event_end_status = ""
            app.phase_start_status = ""
            app.phase_end_status = ""
            app.base_point_status = ""
            app.peak_point_status = ""
            
            # Create the graph
            print(f"Creating graph with shape: {df.shape}")
            fig = create_event_preparation_graph(df)
            
            # Create success message
            status = html.Div(f"Successfully loaded {filename}", style={'color': 'green', 'marginLeft': '10px'})
            
            return fig, df.to_dict('records'), status
        
        except Exception as e:
            print(f"Error loading CSV file: {str(e)}")
            traceback.print_exc()
            
            # Create an error figure
            error_fig = go.Figure()
            error_fig.update_layout(
                title='Error Loading Data',
                annotations=[dict(
                    text=f"Error: {str(e)}",
                    xref="paper",
                    yref="paper",
                    x=0.5,
                    y=0.5,
                    showarrow=False,
                    font=dict(size=14, color="red")
                )]
            )
            
            return error_fig, [], html.Div(f"Error: {str(e)}", style={'color': 'red', 'marginLeft': '10px'})
    
    # 3. Update event edit dropdown options when events change
    @app.callback(
        Output('event-edit-selector', 'options'),
        [Input('event-list-display', 'children')],
        prevent_initial_call=True
    )
    def update_event_edit_options(_):
        """Update event dropdown options based on available events"""
        try:
            if not app.event_selections:
                return []
                    
            options = []
            for i, (_, _, label, _, _) in enumerate(app.event_selections):
                options.append({'label': f"{label}", 'value': i})
                    
            return options
        except Exception as e:
            print(f"Error updating event edit options: {str(e)}")
            traceback.print_exc()
            return []
    
    # 4. Populate event edit inputs when selection changes
    @app.callback(
        [Output('event-edit-label', 'value'),
        Output('event-edit-start-status', 'children'),
        Output('event-edit-end-status', 'children')],
        [Input('event-edit-selector', 'value')],
        prevent_initial_call=True
    )
    def populate_event_edit_fields(event_index):
        """Populate event edit fields with selected event data"""
        if event_index is None or not app.event_selections:
            return "", "", ""
                
        try:
            # Convert to integer index if needed
            if isinstance(event_index, str):
                event_index = int(event_index)
                    
            # Get selected event
            if 0 <= event_index < len(app.event_selections):
                app.event_edit_index = event_index
                _, _, label, start_time, end_time = app.event_selections[event_index]
                    
                # Reset edit points
                app.event_edit_start_point = None
                app.event_edit_end_point = None
                    
                return label, f"Current: {start_time:.2f}s", f"Current: {end_time:.2f}s"
                    
            return "", "", ""
        except Exception as e:
            print(f"Error populating event edit fields: {str(e)}")
            traceback.print_exc()
            return "", "", ""
    
    # 5. Handle event edit point selection buttons
    @app.callback(
        [Output('selection-status', 'children', allow_duplicate=True),
         Output('event-edit-start-status', 'children', allow_duplicate=True),
         Output('event-edit-end-status', 'children', allow_duplicate=True)],
        [Input('event-edit-start-button', 'n_clicks'),
         Input('event-edit-end-button', 'n_clicks'),
         Input('event-prep-graph', 'clickData')],
        prevent_initial_call=True
    )
    def handle_event_edit_selection(start_clicks, end_clicks, click_data):
        """Handle event edit point selection"""
        ctx = dash.callback_context
        if not ctx.triggered:
            return dash.no_update, dash.no_update, dash.no_update
                
        trigger_id = ctx.triggered[0]['prop_id'].split('.')[0]
            
        # Set selection mode based on button clicks
        if trigger_id == 'event-edit-start-button':
            app.selection_mode = 'event_edit_start'
            return "Click on graph to select new EVENT START point", "Waiting for selection...", dash.no_update
                
        elif trigger_id == 'event-edit-end-button':
            app.selection_mode = 'event_edit_end'
            return "Click on graph to select new EVENT END point", dash.no_update, "Waiting for selection..."
                
        # Process click data if in edit selection mode
        elif trigger_id == 'event-prep-graph' and app.selection_mode in ['event_edit_start', 'event_edit_end'] and click_data:
            try:
                print(f"Event edit click data: {click_data}")  # Debug: Print the click data
                x_value = click_data['points'][0]['x']
                y_value = click_data['points'][0]['y']
                    
                # Convert to milliseconds if needed
                time_in_seconds = True
                if app.event_data is not None and not app.event_data.empty and 'Time' in app.event_data.columns:
                    # Check if Time column is Offset/1000
                    if app.event_data['Time'].iloc[0] == app.event_data['Offset'].iloc[0]/1000:
                        offset_value = x_value * 1000
                        time_in_seconds = True
                    else:
                        offset_value = x_value
                        time_in_seconds = False
                else:
                    offset_value = x_value
                    time_in_seconds = False
                    
                # Find closest point in data
                if not app.event_data.empty:
                    closest_idx = (app.event_data['Offset'] - offset_value).abs().idxmin()
                    closest_offset = app.event_data.loc[closest_idx, 'Offset']
                    
                    # Determine if we need to convert to seconds for display
                    if 'Time' in app.event_data.columns:
                        closest_time = app.event_data.loc[closest_idx, 'Time']
                    else:
                        closest_time = closest_offset / 1000 if time_in_seconds else closest_offset
                    
                    closest_value = app.event_data.loc[closest_idx, 'Value']
                    
                    print(f"Closest point for event edit: offset={closest_offset}, time={closest_time}, value={closest_value}")
                else:
                    closest_offset = offset_value
                    closest_time = x_value
                    closest_value = y_value
                    
                point_info = {
                    'offset': closest_offset,
                    'time': closest_time,
                    'value': closest_value
                }
                    
                status_msg = f"Selected: Time={closest_time:.2f}s, Value={closest_value:.2f}"
                    
                if app.selection_mode == 'event_edit_start':
                    app.event_edit_start_point = point_info
                    app.selection_mode = None  # Reset selection mode
                    return f"Selected new event start point at {closest_time:.2f}s", status_msg, dash.no_update
                        
                elif app.selection_mode == 'event_edit_end':
                    app.event_edit_end_point = point_info
                    app.selection_mode = None  # Reset selection mode
                    return f"Selected new event end point at {closest_time:.2f}s", dash.no_update, status_msg
                    
            except Exception as e:
                print(f"Error handling event edit selection: {str(e)}")
                traceback.print_exc()
                return f"Error selecting point: {str(e)}", dash.no_update, dash.no_update
                    
        return dash.no_update, dash.no_update, dash.no_update
    
    # 6. Update event callback
    @app.callback(
        [Output('event-update-status', 'children'),
         Output('event-list-display', 'children', allow_duplicate=True),
         Output('event-prep-graph', 'figure', allow_duplicate=True),
         Output('event-prep-table', 'data', allow_duplicate=True)],
        [Input('update-event-button', 'n_clicks')],
        [State('event-edit-selector', 'value'),
         State('event-edit-label', 'value')],
        prevent_initial_call=True
    )
    def update_event(n_clicks, event_index, new_label):
        """Update the selected event with new values"""
        if not n_clicks or event_index is None:
            return dash.no_update, dash.no_update, dash.no_update, dash.no_update
                
        try:
            # Convert to integer index if needed
            if isinstance(event_index, str):
                event_index = int(event_index)
                    
            if event_index < 0 or event_index >= len(app.event_selections):
                return html.Div("Invalid event selection", style={'color': 'red'}), dash.no_update, dash.no_update, dash.no_update
                    
            if not new_label:
                return html.Div("Event label cannot be empty", style={'color': 'red'}), dash.no_update, dash.no_update, dash.no_update
                    
            # Get original event data
            old_start_offset, old_end_offset, old_label, old_start_time, old_end_time = app.event_selections[event_index]
                
            # Prepare new event data
            new_start_offset = app.event_edit_start_point['offset'] if app.event_edit_start_point else old_start_offset
            new_end_offset = app.event_edit_end_point['offset'] if app.event_edit_end_point else old_end_offset
            new_start_time = app.event_edit_start_point['time'] if app.event_edit_start_point else old_start_time
            new_end_time = app.event_edit_end_point['time'] if app.event_edit_end_point else old_end_time
                
            # Ensure start is before end
            if new_start_offset > new_end_offset:
                new_start_offset, new_end_offset = new_end_offset, new_start_offset
                new_start_time, new_end_time = new_end_time, new_start_time
                    
            # Update the event in the dataframe
            if not app.event_data.empty:
                # First, clear the old event
                mask = (app.event_data['Event'] == old_label)
                app.event_data.loc[mask, 'Event'] = ''
                    
                # Then set the new event
                mask = (app.event_data['Offset'] >= new_start_offset) & (app.event_data['Offset'] <= new_end_offset)
                app.event_data.loc[mask, 'Event'] = new_label
                    
                # Update the event selection
                app.event_selections[event_index] = (new_start_offset, new_end_offset, new_label, new_start_time, new_end_time)
                    
                # Create success message
                status = html.Div(f"Updated event '{new_label}' from {new_start_time:.2f}s to {new_end_time:.2f}s", 
                                style={'color': 'green'})
                    
                # Update event list display
                event_list = html.Div([
                    html.Ul([
                        html.Li([
                            html.Strong(f"{label}"), 
                            f"From {start_time:.2f}s to {end_time:.2f}s"
                        ]) 
                        for start_offset, end_offset, label, start_time, end_time in sorted(
                            app.event_selections,
                            key=lambda x: x[0]  # Sort by start_offset
                        )
                    ])
                ])
                    
                # Update the graph
                updated_fig = create_event_preparation_graph(app.event_data)
                    
                # Reset edit points
                app.event_edit_start_point = None
                app.event_edit_end_point = None
                    
                return status, event_list, updated_fig, app.event_data.to_dict('records')
            
            return html.Div("Error: No data loaded", style={'color': 'red'}), dash.no_update, dash.no_update, dash.no_update
            
        except Exception as e:
            print(f"Error updating event: {str(e)}")
            traceback.print_exc()
            return html.Div(f"Error: {str(e)}", style={'color': 'red'}), dash.no_update, dash.no_update, dash.no_update

    # 7. Update base edit dropdown options with base point change
    @app.callback(
        Output('base-point-edit-selector', 'options'),
        [Input('base-point-list-display', 'children')],
        prevent_initial_call=True
    )
    def update_base_edit_options(_):
        """Update base dropdown options based on available bases"""
        try:
            if not app.base_selections:
                return []
                    
            options = []
            # Fix here: The base_selections has the format (offset, label, time, value)
            for i, (_, label, _, _) in enumerate(app.base_selections):
                options.append({'label': f"{label}", 'value': i})
                    
            return options
        except Exception as e:
            print(f"Error updating base edit options: {str(e)}")
            traceback.print_exc()
            return []
    
    # 8. Populate base edit input when selection changes
    @app.callback(
        [Output('base-point-edit-label', 'value'),
        Output('base-point-edit-status', 'children', allow_duplicate=True)],
        [Input('base-point-edit-selector', 'value')],
        prevent_initial_call=True
    )
    def populate_base_edit_fields(base_index):
        """Populate base edit fields with selected base data"""
        if base_index is None or not app.base_selections:
            return "", ""
                
        try:
            # Convert to integer index if needed
            if isinstance(base_index, str):
                base_index = int(base_index)
                    
            # Get selected base
            if 0 <= base_index < len(app.base_selections):
                app.base_edit_index = base_index
                base_offset, base_label, base_time, base_value = app.base_selections[base_index]
                    
                # Reset edit points
                app.base_edit_point = None
                    
                return base_label, f"Current: {base_time:.2f}s"
                    
            return "", ""
        except Exception as e:
            print(f"Error populating base edit fields: {str(e)}")
            traceback.print_exc()
            return "", ""
    
    # 9. Handle base edit point selection button
    @app.callback(
        [Output('selection-status', 'children', allow_duplicate=True),
        Output('base-point-edit-status', 'children')],
        [Input('base-point-edit-button', 'n_clicks'),
        Input('event-prep-graph', 'clickData')],
        prevent_initial_call=True
    )
    def handle_base_edit_selection(point_clicks, click_data):
        """Handle base edit point selection"""
        ctx = dash.callback_context
        if not ctx.triggered:
            return dash.no_update, dash.no_update
                
        trigger_id = ctx.triggered[0]['prop_id'].split('.')[0]
            
        # Set selection mode based on button clicks
        if trigger_id == 'base-point-edit-button':
            app.selection_mode = 'base_edit_point'
            return "Click on graph to select new base point", "Waiting for selection..."
                
        # Process click data if in edit selection mode
        elif trigger_id == 'event-prep-graph' and app.selection_mode == 'base_edit_point' and click_data:
            try:
                print(f"Base edit click data: {click_data}")  # Debug: Print the click data
                x_value = click_data['points'][0]['x']
                y_value = click_data['points'][0]['y']
                    
                # Convert to milliseconds if needed
                time_in_seconds = True
                if app.event_data is not None and not app.event_data.empty and 'Time' in app.event_data.columns:
                    # Check if Time column is Offset/1000
                    if app.event_data['Time'].iloc[0] == app.event_data['Offset'].iloc[0]/1000:
                        offset_value = x_value * 1000
                        time_in_seconds = True
                    else:
                        offset_value = x_value
                        time_in_seconds = False
                else:
                    offset_value = x_value
                    time_in_seconds = False
                    
                # Find closest point in data
                if not app.event_data.empty:
                    closest_idx = (app.event_data['Offset'] - offset_value).abs().idxmin()
                    closest_offset = app.event_data.loc[closest_idx, 'Offset']
                    
                    # Determine if we need to convert to seconds for display
                    if 'Time' in app.event_data.columns:
                        closest_time = app.event_data.loc[closest_idx, 'Time']
                    else:
                        closest_time = closest_offset / 1000 if time_in_seconds else closest_offset
                    
                    closest_value = app.event_data.loc[closest_idx, 'Value']
                    
                    print(f"Closest point for base edit: offset={closest_offset}, time={closest_time}, value={closest_value}")
                else:
                    closest_offset = offset_value
                    closest_time = x_value
                    closest_value = y_value
                    
                point_info = {
                    'offset': closest_offset,
                    'time': closest_time,
                    'value': closest_value
                }
                    
                status_msg = f"Selected: Time={closest_time:.2f}s, Value={closest_value:.2f}"
                
                app.base_edit_point = point_info
                app.selection_mode = None  # Reset selection mode
                return f"Selected new base point at {closest_time:.2f}s, value={closest_value:.2f}", status_msg
                    
            except Exception as e:
                print(f"Error handling base edit selection: {str(e)}")
                traceback.print_exc()
                return f"Error selecting point: {str(e)}", dash.no_update
                    
        return dash.no_update, dash.no_update
    
    # 10. Update base callback
    @app.callback(
        [Output('base-point-add-status', 'children'),
        Output('base-point-list-display', 'children', allow_duplicate=True),
        Output('event-prep-graph', 'figure', allow_duplicate=True),
        Output('event-prep-table', 'data', allow_duplicate=True)],
        Input('add-base-point-button', 'n_clicks'),
        [State('base-point-label-input', 'value'),
        State('event-prep-graph', 'figure'),
        State('event-prep-table', 'data')],
        prevent_initial_call=True
    )
    def add_base_point(n_clicks, base_label, current_figure, current_data):
        """Add base at base point with pos_delta calculation"""
        if not n_clicks or not base_label:
            return "", dash.no_update, dash.no_update, dash.no_update
        
        try:
            # Validate base point
            if not app.base_point:
                return html.Div("Please select base point", style={'color': 'red'}), dash.no_update, dash.no_update, dash.no_update
            
            # Update base data
            if not app.event_data.empty:
                # Store the base selection - directly access the values from app.base_point
                base_offset = app.base_point['offset']
                base_time = app.base_point['time']
                base_value = round(app.base_point['value'], 2)
                
                # Calculate pos_delta if there are prior peaks
                pos_delta = calculate_pos_delta(app.event_data, app.base_point, app.peak_selections)
                
                # Store as a tuple with all needed information
                app.base_selections.append((base_offset, base_label, base_time, base_value))
                
                # Update the dataframe - pass just the base_point and label
                base_value = round(app.base_point['value'], 2)  # Round to 2 decimal places
                app.event_data = update_base_point(app.event_data, app.base_point, base_label)
                
                # Add pos_delta to the dataframe at the base point offset
                if pos_delta is not None:
                    pos_delta = round(pos_delta, 2)  # Round to 2 decimal places
                    mask = (app.event_data['Offset'] == base_offset)
                    app.event_data.loc[mask, 'Pos_Delta'] = pos_delta
                
                # Create success message with pos_delta if available
                if pos_delta is not None:
                    status_text = f"Added base '{base_label}' at time {base_time:.2f}s with value {base_value:.2f}, Pos_Delta: {pos_delta:.2f}"
                else:
                    status_text = f"Added base '{base_label}' at time {base_time:.2f}s with value {base_value:.2f}, No prior peak found"
                    
                status = html.Div(status_text, style={'color': 'green'})
                
                # Create base list with improved display including pos_delta
                base_list_items = []
                for offset, label, time, value in sorted(app.base_selections, key=lambda x: x[0]):
                    # Get pos_delta from the dataframe for this offset
                    pos_delta_value = None
                    if 'Pos_Delta' in app.event_data.columns:
                        mask = (app.event_data['Offset'] == offset)
                        if mask.any() and not app.event_data.loc[mask, 'Pos_Delta'].isnull().all():
                            pos_delta_value = app.event_data.loc[mask, 'Pos_Delta'].iloc[0]
                    
                    # Create list item text with pos_delta if available
                    if pos_delta_value is not None:
                        list_text = f"At {time:.2f}s, Value: {value:.2f}, Pos_Delta: {pos_delta_value:.2f}"
                    else:
                        list_text = f"At {time:.2f}s, Value: {value:.2f}"
                    
                    base_list_items.append(html.Li([
                        html.Strong(f"{label}"), 
                        list_text
                    ]))
                
                base_list = html.Div([html.Ul(base_list_items)])
                
                # Update the graph
                updated_fig = create_event_preparation_graph(app.event_data)
                
                # Reset selection points
                app.base_point = None
                app.base_point_status = ""
                
                return status, base_list, updated_fig, app.event_data.to_dict('records')
            
            return html.Div("Error: No data loaded", style={'color': 'red'}), dash.no_update, dash.no_update, dash.no_update
    
        except Exception as e:
            print(f"Error adding base point: {str(e)}")
            traceback.print_exc()
            return html.Div(f"Error: {str(e)}", style={'color': 'red'}), dash.no_update, dash.no_update, dash.no_update

    # 11. Update peak edit dropdown options with peak point change
    @app.callback(
        Output('peak-point-edit-selector', 'options'),
        [Input('peak-point-list-display', 'children')],
    )
    def update_peak_edit_options(_):
        """Update peak dropdown options based on available peaks"""
        try:
            if not app.peak_selections:
                return []
                    
            options = []
            for i, (_, label, _, _) in enumerate(app.peak_selections):
                options.append({'label': f"{label}", 'value': i})
                    
            return options
        except Exception as e:
            print(f"Error updating peak edit options: {str(e)}")
            traceback.print_exc()
            return []
    
    # 12. Populate peak edit input when selection changes
    @app.callback(
        [Output('peak-point-edit-label', 'value'),
        Output('peak-point-edit-status', 'children', allow_duplicate=True)],  # Add allow_duplicate=True here
        [Input('peak-point-edit-selector', 'value')],
        prevent_initial_call=True  # You might want to add this as well
    )
    def populate_peak_edit_fields(peak_index):
        """Populate peak edit fields with selected peak data"""
        if peak_index is None or not app.peak_selections:
            return "", ""
                
        try:
            # Convert to integer index if needed
            if isinstance(peak_index, str):
                peak_index = int(peak_index)
                    
            # Get selected base
            if 0 <= peak_index < len(app.peak_selections):
                app.peak_edit_index = peak_index
                peak_offset, peak_label, peak_time, peak_value = app.peak_selections[peak_index]
                    
                # Reset edit points
                app.peak_edit_point = None
                    
                return peak_label, f"Current: {peak_time:.2f}s"
                    
            return "", ""
        except Exception as e:
            print(f"Error populating peak edit fields: {str(e)}")
            traceback.print_exc()
            return "", ""
    
    # 13. Handle peak edit point selection button
    @app.callback(
        [Output('selection-status', 'children', allow_duplicate=True),
        Output('peak-point-edit-status', 'children')],
        [Input('peak-point-edit-button', 'n_clicks'),
        Input('event-prep-graph', 'clickData')],
        prevent_initial_call=True
    )
    def handle_peak_edit_selection(point_clicks, click_data):
        """Handle peak edit point selection"""
        ctx = dash.callback_context
        if not ctx.triggered:
            return dash.no_update, dash.no_update
                
        trigger_id = ctx.triggered[0]['prop_id'].split('.')[0]
            
        # Set selection mode based on button clicks
        if trigger_id == 'peak-point-edit-button':
            app.selection_mode = 'peak_edit_point'
            return "Click on graph to select new peak point", "Waiting for selection..."
                
        # Process click data if in edit selection mode
        elif trigger_id == 'event-prep-graph' and app.selection_mode == 'peak_edit_point' and click_data:
            try:
                print(f"Peak edit click data: {click_data}")  # Debug: Print the click data
                x_value = click_data['points'][0]['x']
                y_value = click_data['points'][0]['y']
                    
                # Convert to milliseconds if needed
                time_in_seconds = True
                if app.event_data is not None and not app.event_data.empty and 'Time' in app.event_data.columns:
                    # Check if Time column is Offset/1000
                    if app.event_data['Time'].iloc[0] == app.event_data['Offset'].iloc[0]/1000:
                        offset_value = x_value * 1000
                        time_in_seconds = True
                    else:
                        offset_value = x_value
                        time_in_seconds = False
                else:
                    offset_value = x_value
                    time_in_seconds = False
                    
                # Find closest point in data
                if not app.event_data.empty:
                    closest_idx = (app.event_data['Offset'] - offset_value).abs().idxmin()
                    closest_offset = app.event_data.loc[closest_idx, 'Offset']
                    
                    # Determine if we need to convert to seconds for display
                    if 'Time' in app.event_data.columns:
                        closest_time = app.event_data.loc[closest_idx, 'Time']
                    else:
                        closest_time = closest_offset / 1000 if time_in_seconds else closest_offset
                    
                    closest_value = app.event_data.loc[closest_idx, 'Value']
                    
                    print(f"Closest point for peak edit: offset={closest_offset}, time={closest_time}, value={closest_value}")
                else:
                    closest_offset = offset_value
                    closest_time = x_value
                    closest_value = y_value
                    
                point_info = {
                    'offset': closest_offset,
                    'time': closest_time,
                    'value': closest_value
                }
                    
                status_msg = f"Selected: Time={closest_time:.2f}s, Value={closest_value:.2f}"
                    
                app.peak_edit_point = point_info
                app.selection_mode = None  # Reset selection mode
                return f"Selected new peak point at {closest_time:.2f}s, value={closest_value:.2f}", status_msg
                    
            except Exception as e:
                print(f"Error handling peak edit selection: {str(e)}")
                traceback.print_exc()
                return f"Error selecting point: {str(e)}", dash.no_update
                    
        return dash.no_update, dash.no_update
    
    # 14. Update peak callback
    @app.callback(
        [Output('peak-point-add-status', 'children'),
        Output('peak-point-list-display', 'children', allow_duplicate=True),
        Output('event-prep-graph', 'figure', allow_duplicate=True),
        Output('event-prep-table', 'data', allow_duplicate=True)],
        Input('add-peak-point-button', 'n_clicks'),
        [State('peak-point-label-input', 'value'),
        State('event-prep-graph', 'figure'),
        State('event-prep-table', 'data')],
        prevent_initial_call=True
    )
    def add_peak_point(n_clicks, peak_label, current_figure, current_data):
        """Add peak at peak point with neg_delta calculation"""
        if not n_clicks or not peak_label:
            return "", dash.no_update, dash.no_update, dash.no_update
        
        try:
            # Validate peak point
            if not app.peak_point:
                return html.Div("Please select peak point", style={'color': 'red'}), dash.no_update, dash.no_update, dash.no_update
            
            # Update peak data
            if not app.event_data.empty:
                # Store the peak selection - directly access the values from app.peak_point
                peak_offset = app.peak_point['offset']
                peak_time = app.peak_point['time']
                peak_value = app.peak_point['value']
                
                # Calculate neg_delta if there are prior bases
                neg_delta = calculate_neg_delta(app.event_data, app.peak_point, app.base_selections)
                
                # Store as a tuple with all needed information
                app.peak_selections.append((peak_offset, peak_label, peak_time, peak_value))
                
                # Update the dataframe - pass just the peak_point and label
                peak_value = round(app.peak_point['value'], 2)  # Round to 2 decimal places
                app.event_data = update_peak_point(app.event_data, app.peak_point, peak_label)
                
                # Add neg_delta to the dataframe at the peak point offset
                if neg_delta is not None:
                    neg_delta = round(neg_delta, 2)  # Round to 2 decimal places
                    mask = (app.event_data['Offset'] == peak_offset)
                    app.event_data.loc[mask, 'Neg_Delta'] = neg_delta
                
                # Create success message with neg_delta if available
                if neg_delta is not None:
                    status_text = f"Added peak '{peak_label}' at time {peak_time:.2f}s with value {peak_value:.2f}, Neg_Delta: {neg_delta:.2f}"
                else:
                    status_text = f"Added peak '{peak_label}' at time {peak_time:.2f}s with value {peak_value:.2f}, No prior base found"
                    
                status = html.Div(status_text, style={'color': 'green'})
                
                # Create peak list with improved display including neg_delta
                peak_list_items = []
                for offset, label, time, value in sorted(app.peak_selections, key=lambda x: x[0]):
                    # Get neg_delta from the dataframe for this offset
                    neg_delta_value = None
                    if 'Neg_Delta' in app.event_data.columns:
                        mask = (app.event_data['Offset'] == offset)
                        if mask.any() and not app.event_data.loc[mask, 'Neg_Delta'].isnull().all():
                            neg_delta_value = app.event_data.loc[mask, 'Neg_Delta'].iloc[0]
                    
                    # Create list item text with neg_delta if available
                    if neg_delta_value is not None:
                        list_text = f"At {time:.2f}s, Value: {value:.2f}, Neg_Delta: {neg_delta_value:.2f}"
                    else:
                        list_text = f"At {time:.2f}s, Value: {value:.2f}"
                    
                    peak_list_items.append(html.Li([
                        html.Strong(f"{label}"), 
                        list_text
                    ]))
                
                peak_list = html.Div([html.Ul(peak_list_items)])
                
                # Update the graph
                updated_fig = create_event_preparation_graph(app.event_data)
                
                # Reset selection points
                app.peak_point = None
                app.peak_point_status = ""
                
                return status, peak_list, updated_fig, app.event_data.to_dict('records')
            
            return html.Div("Error: No data loaded", style={'color': 'red'}), dash.no_update, dash.no_update, dash.no_update
        
        except Exception as e:
            print(f"Error adding peak point: {str(e)}")
            traceback.print_exc()
            return html.Div(f"Error: {str(e)}", style={'color': 'red'}), dash.no_update, dash.no_update, dash.no_update
    
    # 15. Update phase edit dropdown options when phases change
    @app.callback(
        Output('phase-edit-selector', 'options'),
        [Input('phase-list-display', 'children')]
    )
    def update_phase_edit_options(_):
        """Update phase dropdown options based on available phases"""
        try:
            if not app.phase_selections:
                return []
                
            options = []
            for i, (_, _, label, _, _) in enumerate(app.phase_selections):
                options.append({'label': f"{label}", 'value': i})
                
            return options
        except Exception as e:
            print(f"Error updating phase edit options: {str(e)}")
            traceback.print_exc()
            return []

    # 16. Populate phase edit inputs when selection changes
    @app.callback(
        [Output('phase-edit-label', 'value'),
         Output('phase-edit-start-status', 'children'),
         Output('phase-edit-end-status', 'children')],
        [Input('phase-edit-selector', 'value')]
    )
    def populate_phase_edit_fields(phase_index):
        """Populate phase edit fields with selected phase data"""
        if phase_index is None or not app.phase_selections:
            return "", "", ""
            
        try:
            # Convert to integer index if needed
            if isinstance(phase_index, str):
                phase_index = int(phase_index)
                
            # Get selected phase
            if 0 <= phase_index < len(app.phase_selections):
                app.phase_edit_index = phase_index
                _, _, label, start_time, end_time = app.phase_selections[phase_index]
                
                # Reset edit points
                app.phase_edit_start_point = None
                app.phase_edit_end_point = None
                
                return label, f"Current: {start_time:.2f}s", f"Current: {end_time:.2f}s"
                
            return "", "", ""
        except Exception as e:
            print(f"Error populating phase edit fields: {str(e)}")
            traceback.print_exc()
            return "", "", ""

    # 17. Handle phase edit point selection buttons
    @app.callback(
        [Output('selection-status', 'children', allow_duplicate=True),
        Output('phase-edit-start-status', 'children', allow_duplicate=True),
        Output('phase-edit-end-status', 'children', allow_duplicate=True)],
        [Input('phase-edit-start-button', 'n_clicks'),
        Input('phase-edit-end-button', 'n_clicks'),
        Input('event-prep-graph', 'clickData')],
        prevent_initial_call=True
    )
    def handle_phase_edit_selection(start_clicks, end_clicks, click_data):
        """Handle phase edit point selection"""
        ctx = dash.callback_context
        if not ctx.triggered:
            return dash.no_update, dash.no_update, dash.no_update
            
        trigger_id = ctx.triggered[0]['prop_id'].split('.')[0]
        
        # Set selection mode based on button clicks
        if trigger_id == 'phase-edit-start-button':
            app.selection_mode = 'phase_edit_start'
            return "Click on graph to select new PHASE START point", "Waiting for selection...", dash.no_update
            
        elif trigger_id == 'phase-edit-end-button':
            app.selection_mode = 'phase_edit_end'
            return "Click on graph to select new PHASE END point", dash.no_update, "Waiting for selection..."
            
        # Process click data if in edit selection mode
        elif trigger_id == 'event-prep-graph' and app.selection_mode in ['phase_edit_start', 'phase_edit_end'] and click_data:
            try:
                print(f"Phase edit click data: {click_data}")  # Debug: Print the click data
                x_value = click_data['points'][0]['x']
                y_value = click_data['points'][0]['y']
                
                # Convert to milliseconds if needed
                time_in_seconds = True
                if app.event_data is not None and not app.event_data.empty and 'Time' in app.event_data.columns:
                    # Check if Time column is Offset/1000
                    if app.event_data['Time'].iloc[0] == app.event_data['Offset'].iloc[0]/1000:
                        offset_value = x_value * 1000
                        time_in_seconds = True
                    else:
                        offset_value = x_value
                        time_in_seconds = False
                else:
                    offset_value = x_value
                    time_in_seconds = False
                
                # Find closest point in data
                if not app.event_data.empty:
                    closest_idx = (app.event_data['Offset'] - offset_value).abs().idxmin()
                    closest_offset = app.event_data.loc[closest_idx, 'Offset']
                    
                    # Determine if we need to convert to seconds for display
                    if 'Time' in app.event_data.columns:
                        closest_time = app.event_data.loc[closest_idx, 'Time']
                    else:
                        closest_time = closest_offset / 1000 if time_in_seconds else closest_offset
                    
                    closest_value = app.event_data.loc[closest_idx, 'Value']
                    print(f"Closest point for phase edit: offset={closest_offset}, time={closest_time}, value={closest_value}")
                else:
                    closest_offset = offset_value
                    closest_time = x_value
                    closest_value = y_value
                
                point_info = {
                    'offset': closest_offset,
                    'time': closest_time,
                    'value': closest_value
                }
                
                status_msg = f"Selected: Time={closest_time:.2f}s, Value={closest_value:.2f}"
                
                if app.selection_mode == 'phase_edit_start':
                    app.phase_edit_start_point = point_info
                    app.selection_mode = None  # Reset selection mode
                    return f"Selected new phase start point at {closest_time:.2f}s", status_msg, dash.no_update
                    
                elif app.selection_mode == 'phase_edit_end':
                    app.phase_edit_end_point = point_info
                    app.selection_mode = None  # Reset selection mode
                    return f"Selected new phase end point at {closest_time:.2f}s", dash.no_update, status_msg
                
            except Exception as e:
                print(f"Error handling phase edit selection: {str(e)}")
                traceback.print_exc()
                return f"Error selecting point: {str(e)}", dash.no_update, dash.no_update
                
        return dash.no_update, dash.no_update, dash.no_update

    # 18. Update phase callback
    @app.callback(
        [Output('phase-update-status', 'children'),
         Output('phase-list-display', 'children', allow_duplicate=True),
         Output('event-prep-graph', 'figure', allow_duplicate=True),
         Output('event-prep-table', 'data', allow_duplicate=True)],
        [Input('update-phase-button', 'n_clicks')],
        [State('phase-edit-selector', 'value'),
         State('phase-edit-label', 'value')],
        prevent_initial_call=True
    )
    def update_phase(n_clicks, phase_index, new_label):
        """Update the selected phase with new values"""
        if not n_clicks or phase_index is None:
            return dash.no_update, dash.no_update, dash.no_update, dash.no_update
            
        try:
            # Convert to integer index if needed
            if isinstance(phase_index, str):
                phase_index = int(phase_index)
                
            if phase_index < 0 or phase_index >= len(app.phase_selections):
                return html.Div("Invalid phase selection", style={'color': 'red'}), dash.no_update, dash.no_update, dash.no_update
                
            if not new_label:
                return html.Div("Phase label cannot be empty", style={'color': 'red'}), dash.no_update, dash.no_update, dash.no_update
                
            # Get original phase data
            old_start_offset, old_end_offset, old_label, old_start_time, old_end_time = app.phase_selections[phase_index]
            
            # Prepare new phase data
            new_start_offset = app.phase_edit_start_point['offset'] if app.phase_edit_start_point else old_start_offset
            new_end_offset = app.phase_edit_end_point['offset'] if app.phase_edit_end_point else old_end_offset
            new_start_time = app.phase_edit_start_point['time'] if app.phase_edit_start_point else old_start_time
            new_end_time = app.phase_edit_end_point['time'] if app.phase_edit_end_point else old_end_time
            
            # Ensure start is before end
            if new_start_offset > new_end_offset:
                new_start_offset, new_end_offset = new_end_offset, new_start_offset
                new_start_time, new_end_time = new_end_time, new_start_time
                
            # Update the phase in the dataframe
            if not app.event_data.empty:
                # First, clear the old phase
                mask = (app.event_data['Phase'] == old_label)
                app.event_data.loc[mask, 'Phase'] = ''
                
                # Then set the new phase
                mask = (app.event_data['Offset'] >= new_start_offset) & (app.event_data['Offset'] <= new_end_offset)
                app.event_data.loc[mask, 'Phase'] = new_label
                
                # Update the phase selection
                app.phase_selections[phase_index] = (new_start_offset, new_end_offset, new_label, new_start_time, new_end_time)
                
                # Create success message
                status = html.Div(f"Updated phase '{new_label}' from {new_start_time:.2f}s to {new_end_time:.2f}s", 
                              style={'color': 'green'})
                
                # Update phase list display
                phase_list = html.Div([
                    html.Ul([
                        html.Li([
                            html.Strong(f"{label}"), 
                            f"From {start_time:.2f}s to {end_time:.2f}s"
                        ]) 
                        for start_offset, end_offset, label, start_time, end_time in sorted(
                            app.phase_selections,
                            key=lambda x: x[0]  # Sort by start_offset
                        )
                    ])
                ])
                
                # Update the graph
                updated_fig = create_event_preparation_graph(app.event_data)
                
                # Reset edit points
                app.phase_edit_start_point = None
                app.phase_edit_end_point = None
                
                return status, phase_list, updated_fig, app.event_data.to_dict('records')
            
            return html.Div("Error: No data loaded", style={'color': 'red'}), dash.no_update, dash.no_update, dash.no_update
            
        except Exception as e:
            print(f"Error updating phase: {str(e)}")
            traceback.print_exc()
            return html.Div(f"Error: {str(e)}", style={'color': 'red'}), dash.no_update, dash.no_update, dash.no_update
    
    # 19. Callback to handle point selections
    @app.callback(
        [Output('selection-status', 'children'),
        Output('event-start-status', 'children'),
        Output('event-end-status', 'children'),
        Output('phase-start-status', 'children'),
        Output('phase-end-status', 'children'),
        Output('base-point-status', 'children'),
        Output('peak-point-status', 'children')],
        [Input('event-prep-graph', 'clickData'),
        Input('select-event-start-button', 'n_clicks'),
        Input('select-event-end-button', 'n_clicks'),
        Input('select-phase-start-button', 'n_clicks'),
        Input('select-phase-end-button', 'n_clicks'),
        Input('select-base-point-button', 'n_clicks'),
        Input('select-peak-point-button', 'n_clicks'),
        Input('clear-event-button', 'n_clicks'),
        Input('clear-phase-button', 'n_clicks'),
        Input('clear-base-point-button', 'n_clicks'),
        Input('clear-peak-point-button', 'n_clicks')],
        prevent_initial_call=True
    )
    def handle_point_selection(click_data, event_start_clicks, event_end_clicks,
                            phase_start_clicks, phase_end_clicks, base_point_clicks, peak_point_clicks,
                            clear_event_clicks, clear_phase_clicks, clear_base_clicks, clear_peak_clicks):
        """Handle point selection based on active selection mode"""
        ctx = dash.callback_context
        if not ctx.triggered:
            return "No selection active", "", "", "", "", "", ""
        
        # Determine which input triggered the callback
        trigger_id = ctx.triggered[0]['prop_id'].split('.')[0]
        
        # Clear selections if clear buttons were clicked
        if trigger_id == 'clear-event-button':
            app.event_start_point = None
            app.event_end_point = None
            app.selection_mode = None
            app.event_start_status = ""
            app.event_end_status = ""
            return "Event selection cleared", "", "", app.phase_start_status, app.phase_end_status, app.base_point_status, app.peak_point_status
        
        # Clear selections if clear buttons were clicked
        if trigger_id == 'clear-base-point-button':
            app.base_point = None
            app.selection_mode = None
            app.base_point_status = ""
            return "Base selection cleared", app.event_start_status, app.event_end_status, app.phase_start_status, app.phase_end_status, "", app.peak_point_status
        
        # Clear selections if clear buttons were clicked
        if trigger_id == 'clear-peak-point-button':
            app.peak_point = None
            app.selection_mode = None
            app.peak_point_status = ""
            return "Peak selection cleared", app.event_start_status, app.event_end_status, app.phase_start_status, app.phase_end_status, app.base_point_status, ""
        
        if trigger_id == 'clear-phase-button':
            app.phase_start_point = None
            app.phase_end_point = None
            app.selection_mode = None
            app.phase_start_status = ""
            app.phase_end_status = ""
            return "Phase selection cleared", app.event_start_status, app.event_end_status, "", "", app.base_point_status, app.peak_point_status
        
        # Set selection mode based on button clicks
        if trigger_id == 'select-event-start-button':
            app.selection_mode = 'event_start'
            app.event_start_status = "Waiting for selection..."
            return "Click on graph to select EVENT START point", app.event_start_status, app.event_end_status, app.phase_start_status, app.phase_end_status, app.base_point_status, app.peak_point_status
        
        if trigger_id == 'select-event-end-button':
            app.selection_mode = 'event_end'
            app.event_end_status = "Waiting for selection..."
            return "Click on graph to select EVENT END point", app.event_start_status, app.event_end_status, app.phase_start_status, app.phase_end_status, app.base_point_status, app.peak_point_status
        
        # Set selection mode based on button clicks
        if trigger_id == 'select-base-point-button':
            app.selection_mode = 'base_point'
            app.base_point_status = "Waiting for selection..."
            return "Click on graph to select base point", app.event_start_status, app.event_end_status, app.phase_start_status, app.phase_end_status, app.base_point_status, app.peak_point_status
        
        # Set selection mode based on button clicks
        if trigger_id == 'select-peak-point-button':
            app.selection_mode = 'peak_point'
            app.peak_point_status = "Waiting for selection..."
            return "Click on graph to select peak point", app.event_start_status, app.event_end_status, app.phase_start_status, app.phase_end_status, app.base_point_status, app.peak_point_status
        
        if trigger_id == 'select-phase-start-button':
            app.selection_mode = 'phase_start'
            app.phase_start_status = "Waiting for selection..."
            return "Click on graph to select PHASE START point", app.event_start_status, app.event_end_status, app.phase_start_status, app.phase_end_status, app.base_point_status, app.peak_point_status
        
        if trigger_id == 'select-phase-end-button':
            app.selection_mode = 'phase_end'
            app.phase_end_status = "Waiting for selection..."
            return "Click on graph to select PHASE END point", app.event_start_status, app.event_end_status, app.phase_start_status, app.phase_end_status, app.base_point_status, app.peak_point_status
        
        # Process click data if in selection mode
        if trigger_id == 'event-prep-graph' and app.selection_mode and click_data:
            try:
                print(f"Click data: {click_data}")  # Debug: Print the click data
                x_value = click_data['points'][0]['x']
                y_value = click_data['points'][0]['y']
                
                # Convert to milliseconds if needed
                time_in_seconds = True
                if app.event_data is not None and not app.event_data.empty and 'Time' in app.event_data.columns:
                    # Check if Time column is Offset/1000
                    if app.event_data['Time'].iloc[0] == app.event_data['Offset'].iloc[0]/1000:
                        offset_value = x_value * 1000
                        time_in_seconds = True
                    else:
                        offset_value = x_value
                        time_in_seconds = False
                else:
                    offset_value = x_value
                    time_in_seconds = False
                
                # Find closest point in data
                if not app.event_data.empty:
                    closest_idx = (app.event_data['Offset'] - offset_value).abs().idxmin()
                    closest_offset = app.event_data.loc[closest_idx, 'Offset']
                    
                    # Determine if we need to convert to seconds for display
                    if 'Time' in app.event_data.columns:
                        closest_time = app.event_data.loc[closest_idx, 'Time']
                    else:
                        closest_time = closest_offset / 1000 if time_in_seconds else closest_offset
                    
                    closest_value = app.event_data.loc[closest_idx, 'Value']
                    
                    print(f"Closest point: offset={closest_offset}, time={closest_time}, value={closest_value}")  # Debug
                else:
                    closest_offset = offset_value
                    closest_time = x_value
                    closest_value = y_value
                    print(f"No data, using clicked point: offset={closest_offset}, time={closest_time}, value={closest_value}")  # Debug
                
                point_info = {
                    'offset': closest_offset,
                    'time': closest_time,
                    'value': closest_value
                }
                
                # Store the point based on selection mode
                status_msg = f"Selected: Time={closest_time:.2f}s, Value={closest_value:.2f}"
                
                if app.selection_mode == 'event_start':
                    app.event_start_point = point_info
                    app.event_start_status = status_msg
                    app.selection_mode = None  # Reset selection mode
                    return f"Point selected: Time={closest_time:.2f}s, Value={closest_value:.2f}", status_msg, app.event_end_status, app.phase_start_status, app.phase_end_status, app.base_point_status, app.peak_point_status
                
                elif app.selection_mode == 'event_end':
                    app.event_end_point = point_info
                    app.event_end_status = status_msg
                    app.selection_mode = None  # Reset selection mode
                    return f"Point selected: Time={closest_time:.2f}s, Value={closest_value:.2f}", app.event_start_status, status_msg, app.phase_start_status, app.phase_end_status, app.base_point_status, app.peak_point_status
                
                elif app.selection_mode == 'phase_start':
                    app.phase_start_point = point_info
                    app.phase_start_status = status_msg
                    app.selection_mode = None  # Reset selection mode
                    return f"Point selected: Time={closest_time:.2f}s, Value={closest_value:.2f}", app.event_start_status, app.event_end_status, status_msg, app.phase_end_status, app.base_point_status, app.peak_point_status
                
                elif app.selection_mode == 'phase_end':
                    app.phase_end_point = point_info
                    app.phase_end_status = status_msg
                    app.selection_mode = None  # Reset selection mode
                    return f"Point selected: Time={closest_time:.2f}s, Value={closest_value:.2f}", app.event_start_status, app.event_end_status, app.phase_start_status, status_msg, app.base_point_status, app.peak_point_status
                
                elif app.selection_mode == 'base_point':
                    app.base_point = point_info
                    app.base_point_status = status_msg
                    app.selection_mode = None  # Reset selection mode
                    return f"Point selected: Time={closest_time:.2f}s, Value={closest_value:.2f}", app.event_start_status, app.event_end_status, app.phase_start_status, app.phase_end_status, status_msg, app.peak_point_status
                
                elif app.selection_mode == 'peak_point':
                    app.peak_point = point_info
                    app.peak_point_status = status_msg
                    app.selection_mode = None  # Reset selection mode
                    return f"Point selected: Time={closest_time:.2f}s, Value={closest_value:.2f}", app.event_start_status, app.event_end_status, app.phase_start_status, app.phase_end_status, app.base_point_status, status_msg
                
                # Store click data for other uses
                app.selected_point = click_data
                
                # Default return if selection mode doesn't match any specifics
                selection_status = f"Point selected: Time={closest_time:.2f}s, Value={closest_value:.2f}"
                return selection_status, app.event_start_status, app.event_end_status, app.phase_start_status, app.phase_end_status, app.base_point_status, app.peak_point_status
                
            except Exception as e:
                print(f"Error handling point selection: {str(e)}")
                traceback.print_exc()
                return f"Error selecting point: {str(e)}", app.event_start_status, app.event_end_status, app.phase_start_status, app.phase_end_status, app.base_point_status, app.peak_point_status
                                    
        # Default return if none of the above conditions are met
        return "No action", app.event_start_status, app.event_end_status, app.phase_start_status, app.phase_end_status, app.base_point_status, app.peak_point_status                                     
    
    # 20. Callback to add event
    @app.callback(
        [Output('event-add-status', 'children'),
         Output('event-list-display', 'children'),
         Output('event-prep-graph', 'figure', allow_duplicate=True),
         Output('event-prep-table', 'data', allow_duplicate=True)],
        Input('add-event-button', 'n_clicks'),
        [State('event-label-input', 'value'),
         State('event-prep-graph', 'figure'),
         State('event-prep-table', 'data')],
        prevent_initial_call=True
    )
    def add_event_range(n_clicks, event_label, current_figure, current_data):
        """Add event between start and end points"""
        if not n_clicks or not event_label:
            return "", dash.no_update, dash.no_update, dash.no_update
        
        try:
            # Validate start and end points
            if not app.event_start_point or not app.event_end_point:
                return html.Div("Please select both start and end points", style={'color': 'red'}), dash.no_update, dash.no_update, dash.no_update
            
            # Update event data
            if not app.event_data.empty:
                # Store the event selection
                start_offset = app.event_start_point['offset']
                end_offset = app.event_end_point['offset']
                start_time = app.event_start_point['time']
                end_time = app.event_end_point['time']
                
                # Store as a tuple with all needed information
                app.event_selections.append((start_offset, end_offset, event_label, start_time, end_time))
                
                # Update the dataframe - ensure we're passing the full point dictionaries
                app.event_data = update_event_range(app.event_data, app.event_start_point, app.event_end_point, event_label)
                
                # Create success message
                status = html.Div(f"Added event '{event_label}' from offset {start_time:.2f}s to {end_time:.2f}s", 
                                style={'color': 'green'})
                
                # Create event list with improved display
                event_list = html.Div([
                    html.Ul([
                        html.Li([
                            html.Strong(f"{label}"), 
                            f"From {start_time:.2f}s to {end_time:.2f}s"
                        ]) 
                        for start_offset, end_offset, label, start_time, end_time in sorted(
                            app.event_selections,
                            key=lambda x: x[0]  # Sort by start_offset
                        )
                    ])
                ])
                
                # Update the graph
                updated_fig = create_event_preparation_graph(app.event_data)
                
                # Reset selection points
                app.event_start_point = None
                app.event_end_point = None
                app.event_start_status = ""
                app.event_end_status = ""
                
                return status, event_list, updated_fig, app.event_data.to_dict('records')
            
            return html.Div("Error: No data loaded", style={'color': 'red'}), dash.no_update, dash.no_update, dash.no_update
        
        except Exception as e:
            print(f"Error adding event range: {str(e)}")
            traceback.print_exc()
            return html.Div(f"Error: {str(e)}", style={'color': 'red'}), dash.no_update, dash.no_update, dash.no_update
    
    # 21. Callback to add base
    @app.callback(
        [Output('base-point-add-status', 'children', allow_duplicate=True),
        Output('base-point-list-display', 'children', allow_duplicate=True),
        Output('event-prep-graph', 'figure', allow_duplicate=True),
        Output('event-prep-table', 'data', allow_duplicate=True)],
        Input('edit-base-point-button', 'n_clicks'),
        [State('base-point-label-input', 'value'),
        State('event-prep-graph', 'figure'),
        State('event-prep-table', 'data')],
        prevent_initial_call=True
    )
    def edit_base_point(n_clicks, base_label, current_figure, current_data):
        """Edit base at base point with pos_delta calculation"""
        if not n_clicks or not base_label:
            return "", dash.no_update, dash.no_update, dash.no_update
        
        try:
            # Validate base point
            if not app.base_point:
                return html.Div("Please select base point", style={'color': 'red'}), dash.no_update, dash.no_update, dash.no_update
            
            # Update base data
            if not app.event_data.empty:
                # Store the base selection - directly access the values from app.base_point
                base_offset = app.base_point['offset']
                base_time = app.base_point['time']
                base_value = round(app.base_point['value'], 2)  # Round to 2 decimal places
                
                # Calculate pos_delta if there are prior peaks
                pos_delta = calculate_pos_delta(app.event_data, app.base_point, app.peak_selections)
                
                # Store as a tuple with all needed information
                app.base_selections.append((base_offset, base_label, base_time, base_value))
                
                # Update the dataframe - pass just the base_point and label
                app.event_data = update_base_point(app.event_data, app.base_point, base_label)
                
                # Make sure the base value is stored with 2 decimal places in Mnl_Base
                mask = (app.event_data['Offset'] == base_offset)
                app.event_data.loc[mask, 'Mnl_Base'] = base_value
                
                # Add pos_delta to the dataframe at the base point offset with 2 decimal places
                if pos_delta is not None:
                    pos_delta = round(pos_delta, 2)  # Round to 2 decimal places
                    mask = (app.event_data['Offset'] == base_offset)
                    app.event_data.loc[mask, 'Pos_Delta'] = pos_delta
                
                # Create success message with pos_delta if available
                if pos_delta is not None:
                    status_text = f"Added base '{base_label}' at time {base_time:.2f}s with value {base_value:.2f}, Pos_Delta: {pos_delta:.2f}"
                else:
                    status_text = f"Added base '{base_label}' at time {base_time:.2f}s with value {base_value:.2f}, No prior peak found"
                    
                status = html.Div(status_text, style={'color': 'green'})
                
                # Create base list with improved display including pos_delta
                base_list_items = []
                for offset, label, time, value in sorted(app.base_selections, key=lambda x: x[0]):
                    # Get pos_delta from the dataframe for this offset
                    pos_delta_value = None
                    if 'Pos_Delta' in app.event_data.columns:
                        mask = (app.event_data['Offset'] == offset)
                        if mask.any() and not app.event_data.loc[mask, 'Pos_Delta'].isnull().all():
                            pos_delta_value = app.event_data.loc[mask, 'Pos_Delta'].iloc[0]
                    
                    # Create list item text with pos_delta if available
                    if pos_delta_value is not None:
                        list_text = f"At {time:.2f}s, Value: {value:.2f}, Pos_Delta: {pos_delta_value:.2f}"
                    else:
                        list_text = f"At {time:.2f}s, Value: {value:.2f}"
                    
                    base_list_items.append(html.Li([
                        html.Strong(f"{label}"), 
                        list_text
                    ]))
                
                base_list = html.Div([html.Ul(base_list_items)])
                
                # Update the graph
                updated_fig = create_event_preparation_graph(app.event_data)
                
                # Reset selection points
                app.base_point = None
                app.base_point_status = ""
                
                return status, base_list, updated_fig, app.event_data.to_dict('records')
            
            return html.Div("Error: No data loaded", style={'color': 'red'}), dash.no_update, dash.no_update, dash.no_update
        
        except Exception as e:
            print(f"Error editing base point: {str(e)}")
            traceback.print_exc()
            return html.Div(f"Error: {str(e)}", style={'color': 'red'}), dash.no_update, dash.no_update, dash.no_update

    # 22. Callback to add peak
    @app.callback(
        [Output('peak-point-add-status', 'children', allow_duplicate=True),
        Output('peak-point-list-display', 'children', allow_duplicate=True),
        Output('event-prep-graph', 'figure', allow_duplicate=True),
        Output('event-prep-table', 'data', allow_duplicate=True)],
        Input('edit-peak-point-button', 'n_clicks'),
        [State('peak-point-label-input', 'value'),
        State('event-prep-graph', 'figure'),
        State('event-prep-table', 'data')],
        prevent_initial_call=True
    )
    def edit_peak_point(n_clicks, peak_label, current_figure, current_data):
        """Edit peak at peak point with neg_delta calculation"""
        if not n_clicks or not peak_label:
            return "", dash.no_update, dash.no_update, dash.no_update
        
        try:
            # Validate peak point
            if not app.peak_point:
                return html.Div("Please select peak point", style={'color': 'red'}), dash.no_update, dash.no_update, dash.no_update
            
            # Update peak data
            if not app.event_data.empty:
                # Store the peak selection - directly access the values from app.peak_point
                peak_offset = app.peak_point['offset']
                peak_time = app.peak_point['time']
                peak_value = round(app.peak_point['value'], 2)  # Fixed: properly round to 2 decimal places
                
                # Calculate neg_delta if there are prior bases
                neg_delta = calculate_neg_delta(app.event_data, app.peak_point, app.base_selections)
                
                # Store as a tuple with all needed information
                app.peak_selections.append((peak_offset, peak_label, peak_time, peak_value))
                
                # Update the dataframe - pass just the peak_point and label
                app.event_data = update_peak_point(app.event_data, app.peak_point, peak_label)
                
                # Make sure the peak value is stored with 2 decimal places in Mnl_Peak
                mask = (app.event_data['Offset'] == peak_offset)
                app.event_data.loc[mask, 'Mnl_Peak'] = peak_value
                            
                # Add neg_delta to the dataframe at the peak point offset
                if neg_delta is not None:
                    neg_delta = round(neg_delta, 2)  # Round to 2 decimal places
                    mask = (app.event_data['Offset'] == peak_offset)
                    app.event_data.loc[mask, 'Neg_Delta'] = neg_delta
                
                # Create success message with neg_delta if available
                if neg_delta is not None:
                    status_text = f"Added peak '{peak_label}' at time {peak_time:.2f}s with value {peak_value:.2f}, Neg_Delta: {neg_delta:.2f}"
                else:
                    status_text = f"Added peak '{peak_label}' at time {peak_time:.2f}s with value {peak_value:.2f}, No prior base found"
                    
                status = html.Div(status_text, style={'color': 'green'})
                
                # Create peak list with improved display including neg_delta
                peak_list_items = []
                for offset, label, time, value in sorted(app.peak_selections, key=lambda x: x[0]):
                    # Get neg_delta from the dataframe for this offset
                    neg_delta_value = None
                    if 'Neg_Delta' in app.event_data.columns:
                        mask = (app.event_data['Offset'] == offset)
                        if mask.any() and not app.event_data.loc[mask, 'Neg_Delta'].isnull().all():
                            neg_delta_value = app.event_data.loc[mask, 'Neg_Delta'].iloc[0]
                    
                    # Create list item text with neg_delta if available
                    if neg_delta_value is not None:
                        list_text = f"At {time:.2f}s, Value: {value:.2f}, Neg_Delta: {neg_delta_value:.2f}"
                    else:
                        list_text = f"At {time:.2f}s, Value: {value:.2f}"
                    
                    peak_list_items.append(html.Li([
                        html.Strong(f"{label}"), 
                        list_text
                    ]))
                
                peak_list = html.Div([html.Ul(peak_list_items)])
                
                # Update the graph
                updated_fig = create_event_preparation_graph(app.event_data)
                
                # Reset selection points
                app.peak_point = None
                app.peak_point_status = ""
                
                return status, peak_list, updated_fig, app.event_data.to_dict('records')
            
            return html.Div("Error: No data loaded", style={'color': 'red'}), dash.no_update, dash.no_update, dash.no_update
        
        except Exception as e:
            print(f"Error editing peak point: {str(e)}")
            traceback.print_exc()
            return html.Div(f"Error: {str(e)}", style={'color': 'red'}), dash.no_update, dash.no_update, dash.no_update
    
    # 23 Update_base function to include pos_delta updating
    @app.callback(
        [Output('base-point-update-status', 'children'),
        Output('base-point-list-display', 'children', allow_duplicate=True),
        Output('event-prep-graph', 'figure', allow_duplicate=True),
        Output('event-prep-table', 'data', allow_duplicate=True)],
        [Input('update-base-point-button', 'n_clicks')],
        [State('base-point-edit-selector', 'value'),
        State('base-point-edit-label', 'value')],
        prevent_initial_call=True
    )
    def update_base(n_clicks, base_index, new_label):
        """Update the selected base with new values and recalculate pos_delta"""
        if not n_clicks or base_index is None:
            return dash.no_update, dash.no_update, dash.no_update, dash.no_update
                
        try:
            # Convert to integer index if needed
            if isinstance(base_index, str):
                base_index = int(base_index)
                    
            if base_index < 0 or base_index >= len(app.base_selections):
                return html.Div("Invalid base selection", style={'color': 'red'}), dash.no_update, dash.no_update, dash.no_update
                    
            if not new_label:
                return html.Div("Base label cannot be empty", style={'color': 'red'}), dash.no_update, dash.no_update, dash.no_update
                    
            # Get original base data
            old_base_offset, old_base_label, old_base_time, old_base_value = app.base_selections[base_index]
                
            # Prepare new base data
            if app.base_edit_point:
                new_base_offset = app.base_edit_point['offset']
                new_base_time = app.base_edit_point['time']
                new_base_value = app.base_edit_point['value']
            else:
                new_base_offset = old_base_offset
                new_base_time = old_base_time
                new_base_value = old_base_value
                    
            # Update the base value in the dataframe
            if not app.event_data.empty:
                # First, ensure the Mnl_Base column exists
                if 'Mnl_Base' not in app.event_data.columns:
                    app.event_data['Mnl_Base'] = None

                # Clear the old base point
                mask = (app.event_data['Offset'] == old_base_offset)
                rows_affected = mask.sum()
                print(f"Clearing old base point at offset {old_base_offset}, rows affected: {rows_affected}")
                app.event_data.loc[mask, 'Mnl_Base'] = None
                app.event_data.loc[mask, 'Pos_Delta'] = None
                    
                # Set the new base value
                mask = (app.event_data['Offset'] == new_base_offset)
                rows_affected = mask.sum()
                print(f"Setting new base point at offset {new_base_offset}, rows affected: {rows_affected}")
                new_base_value = round(new_base_value, 2)  # Add this line to round the value
                app.event_data.loc[mask, 'Mnl_Base'] = new_base_value
                    
                # Calculate pos_delta if there's a prior peak
                if app.base_edit_point:
                    point_for_delta = app.base_edit_point
                else:
                    point_for_delta = {'offset': new_base_offset, 'time': new_base_time, 'value': new_base_value}
                    
                pos_delta = calculate_pos_delta(app.event_data, point_for_delta, app.peak_selections)
                if pos_delta is not None:
                    app.event_data.loc[mask, 'Pos_Delta'] = pos_delta
                    
                # Update the base selection
                app.base_selections[base_index] = (new_base_offset, new_label, new_base_time, new_base_value)
                print(f"Updated base selection: {app.base_selections[base_index]}")
                    
                # Create success message with pos_delta if available
                if pos_delta is not None:
                    status_text = f"Updated base '{new_label}' at {new_base_time:.2f}s, value: {new_base_value:.2f}, Pos_Delta: {pos_delta:.2f}"
                else:
                    status_text = f"Updated base '{new_label}' at {new_base_time:.2f}s, value: {new_base_value:.2f}, No prior peak found"
                    
                status = html.Div(status_text, style={'color': 'green'})
                    
                # Update base list display with pos_delta values
                base_list_items = []
                for offset, label, time, value in sorted(app.base_selections, key=lambda x: x[0]):
                    # Get pos_delta from the dataframe for this offset
                    pos_delta_value = None
                    if 'Pos_Delta' in app.event_data.columns:
                        mask = (app.event_data['Offset'] == offset)
                        if mask.any() and not app.event_data.loc[mask, 'Pos_Delta'].isnull().all():
                            pos_delta_value = app.event_data.loc[mask, 'Pos_Delta'].iloc[0]
                    
                    # Create list item text with pos_delta if available
                    if pos_delta_value is not None:
                        list_text = f"At {time:.2f}s, Value: {value:.2f}, Pos_Delta: {pos_delta_value:.2f}"
                    else:
                        list_text = f"At {time:.2f}s, Value: {value:.2f}"
                    
                    base_list_items.append(html.Li([
                        html.Strong(f"{label}"), 
                        list_text
                    ]))
                
                base_list = html.Div([html.Ul(base_list_items)])
                    
                # Update the graph
                updated_fig = create_event_preparation_graph(app.event_data)
                    
                # Reset base edit point
                app.base_edit_point = None
                    
                return status, base_list, updated_fig, app.event_data.to_dict('records')
            
            return html.Div("Error: No data loaded", style={'color': 'red'}), dash.no_update, dash.no_update, dash.no_update
            
        except Exception as e:
            print(f"Error updating base: {str(e)}")
            traceback.print_exc()
            return html.Div(f"Error: {str(e)}", style={'color': 'red'}), dash.no_update, dash.no_update, dash.no_update

    # 24 Update_peak function to include neg_delta updating
    @app.callback(
        [Output('peak-point-update-status', 'children'),
        Output('peak-point-list-display', 'children', allow_duplicate=True),
        Output('event-prep-graph', 'figure', allow_duplicate=True),
        Output('event-prep-table', 'data', allow_duplicate=True)],
        [Input('update-peak-point-button', 'n_clicks')],
        [State('peak-point-edit-selector', 'value'),
        State('peak-point-edit-label', 'value')],
        prevent_initial_call=True
    )
    def update_peak(n_clicks, peak_index, new_label):
        """Update the selected peak with new values and recalculate neg_delta"""
        if not n_clicks or peak_index is None:
            return dash.no_update, dash.no_update, dash.no_update, dash.no_update
                
        try:
            print(f"Updating peak: index={peak_index}, new_label={new_label}")
            
            # Convert to integer index if needed
            if isinstance(peak_index, str):
                peak_index = int(peak_index)
                    
            if peak_index < 0 or peak_index >= len(app.peak_selections):
                print(f"Invalid peak index: {peak_index}, selections: {app.peak_selections}")
                return html.Div("Invalid peak selection", style={'color': 'red'}), dash.no_update, dash.no_update, dash.no_update
                    
            if not new_label:
                return html.Div("Peak label cannot be empty", style={'color': 'red'}), dash.no_update, dash.no_update, dash.no_update
            
            # Debug print
            print(f"app.peak_selections: {app.peak_selections}")
            print(f"peak_index: {peak_index}")
            
            # Get original peak data
            old_peak_offset, old_peak_label, old_peak_time, old_peak_value = app.peak_selections[peak_index]
            print(f"Old peak data: offset={old_peak_offset}, label={old_peak_label}, time={old_peak_time}, value={old_peak_value}")
                
            # Prepare new peak data
            if app.peak_edit_point:
                new_peak_offset = app.peak_edit_point['offset']
                new_peak_time = app.peak_edit_point['time']
                new_peak_value = app.peak_edit_point['value']
                print(f"New peak point selected: offset={new_peak_offset}, time={new_peak_time}, value={new_peak_value}")
            else:
                new_peak_offset = old_peak_offset
                new_peak_time = old_peak_time
                new_peak_value = old_peak_value
                print(f"Using old peak data (no new point selected)")
                    
            # Update the peak value in the dataframe
            if not app.event_data.empty:
                # Check if 'Mnl_Peak' column exists
                if 'Mnl_Peak' not in app.event_data.columns:
                    print("'Mnl_Peak' column does not exist in the dataframe")
                    app.event_data['Mnl_Peak'] = None
                    
                # Clear the old peak point
                mask = (app.event_data['Offset'] == old_peak_offset)
                old_rows_affected = mask.sum()
                print(f"Clearing old peak point: offset={old_peak_offset}, rows affected={old_rows_affected}")
                app.event_data.loc[mask, 'Mnl_Peak'] = None
                app.event_data.loc[mask, 'Neg_Delta'] = None
                    
                # Set the new peak value
                mask = (app.event_data['Offset'] == new_peak_offset)
                new_rows_affected = mask.sum()
                print(f"Setting new peak point: offset={new_peak_offset}, rows affected={new_rows_affected}")
                new_peak_value = round(new_peak_value, 2)  # Add this line to round the value
                app.event_data.loc[mask, 'Mnl_Peak'] = new_peak_value
                    
                # Calculate neg_delta if there's a prior base
                if app.peak_edit_point:
                    point_for_delta = app.peak_edit_point
                else:
                    point_for_delta = {'offset': new_peak_offset, 'time': new_peak_time, 'value': new_peak_value}
                    
                neg_delta = calculate_neg_delta(app.event_data, point_for_delta, app.base_selections)
                if neg_delta is not None:
                    app.event_data.loc[mask, 'Neg_Delta'] = neg_delta
                    
                # Update the peak selection
                app.peak_selections[peak_index] = (new_peak_offset, new_label, new_peak_time, new_peak_value)
                print(f"Updated peak selection: {app.peak_selections[peak_index]}")
                    
                # Create success message with neg_delta if available
                if neg_delta is not None:
                    status_text = f"Updated peak '{new_label}' at {new_peak_time:.2f}s, value: {new_peak_value:.2f}, Neg_Delta: {neg_delta:.2f}"
                else:
                    status_text = f"Updated peak '{new_label}' at {new_peak_time:.2f}s, value: {new_peak_value:.2f}, No prior base found"
                    
                status = html.Div(status_text, style={'color': 'green'})
                    
                # Update peak list display with neg_delta values
                peak_list_items = []
                for offset, label, time, value in sorted(app.peak_selections, key=lambda x: x[0]):
                    # Get neg_delta from the dataframe for this offset
                    neg_delta_value = None
                    if 'Neg_Delta' in app.event_data.columns:
                        mask = (app.event_data['Offset'] == offset)
                        if mask.any() and not app.event_data.loc[mask, 'Neg_Delta'].isnull().all():
                            neg_delta_value = app.event_data.loc[mask, 'Neg_Delta'].iloc[0]
                    
                    # Create list item text with neg_delta if available
                    if neg_delta_value is not None:
                        list_text = f"At {time:.2f}s, Value: {value:.2f}, Neg_Delta: {neg_delta_value:.2f}"
                    else:
                        list_text = f"At {time:.2f}s, Value: {value:.2f}"
                    
                    peak_list_items.append(html.Li([
                        html.Strong(f"{label}"), 
                        list_text
                    ]))
                
                peak_list = html.Div([html.Ul(peak_list_items)])
                    
                # Update the graph
                updated_fig = create_event_preparation_graph(app.event_data)
                    
                # Reset peak edit point
                app.peak_edit_point = None
                    
                return status, peak_list, updated_fig, app.event_data.to_dict('records')
            
            print("Error: app.event_data is empty")
            return html.Div("Error: No data loaded", style={'color': 'red'}), dash.no_update, dash.no_update, dash.no_update
            
        except Exception as e:
            print(f"Error updating peak: {str(e)}")
            traceback.print_exc()
            return html.Div(f"Error: {str(e)}", style={'color': 'red'}), dash.no_update, dash.no_update, dash.no_update
    
    # 25. Callback to add phase
    @app.callback(
        [Output('phase-add-status', 'children'),
         Output('phase-list-display', 'children'),
         Output('event-prep-graph', 'figure', allow_duplicate=True),
         Output('event-prep-table', 'data', allow_duplicate=True)],
        Input('add-phase-button', 'n_clicks'),
        [State('phase-label-input', 'value'),
         State('event-prep-graph', 'figure'),
         State('event-prep-table', 'data')],
        prevent_initial_call=True
    )
    def add_phase_range(n_clicks, phase_label, current_figure, current_data):
        """Add phase between start and end points"""
        if not n_clicks or not phase_label:
            return "", dash.no_update, dash.no_update, dash.no_update
        
        try:
            # Validate start and end points
            if not app.phase_start_point or not app.phase_end_point:
                return html.Div("Please select both start and end points", style={'color': 'red'}), dash.no_update, dash.no_update, dash.no_update
            
            # Update phase data
            if not app.event_data.empty:
                # Store the phase selection
                start_offset = app.phase_start_point['offset']
                end_offset = app.phase_end_point['offset']
                start_time = app.phase_start_point['time']
                end_time = app.phase_end_point['time']
                
                # Store as a tuple with all needed information
                app.phase_selections.append((start_offset, end_offset, phase_label, start_time, end_time))
                
                # Update the dataframe
                app.event_data = update_phase_range(app.event_data, app.phase_start_point, app.phase_end_point, phase_label)
                
                # Create success message
                status = html.Div(f"Added phase '{phase_label}' from offset {start_time:.2f}s to {end_time:.2f}s", 
                                style={'color': 'green'})
                
                # Create phase list with improved display
                phase_list = html.Div([
                    html.Ul([
                        html.Li([
                            html.Strong(f"{label}"), 
                            f"From {start_time:.2f}s to {end_time:.2f}s"
                        ]) 
                        for start_offset, end_offset, label, start_time, end_time in sorted(
                            app.phase_selections,
                            key=lambda x: x[0]  # Sort by start_offset
                        )
                    ])
                ])
                
                # Update the graph
                updated_fig = create_event_preparation_graph(app.event_data)
                
                # Reset selection points
                app.phase_start_point = None
                app.phase_end_point = None
                app.phase_start_status = ""
                app.phase_end_status = ""
                
                return status, phase_list, updated_fig, app.event_data.to_dict('records')
            
            return html.Div("Error: No data loaded", style={'color': 'red'}), dash.no_update, dash.no_update, dash.no_update
        
        except Exception as e:
            print(f"Error adding phase range: {str(e)}")
            traceback.print_exc()
            return html.Div(f"Error: {str(e)}", style={'color': 'red'}), dash.no_update, dash.no_update, dash.no_update

    # 26. Callback to clear selected events in the table
    @app.callback(
        [Output('clear-status', 'children'),
         Output('event-prep-graph', 'figure', allow_duplicate=True),
         Output('event-prep-table', 'data', allow_duplicate=True)],
        [Input('clear-selected-event-button', 'n_clicks'),
         Input('clear-selected-phase-button', 'n_clicks')],
        [State('event-prep-table', 'selected_rows'),
         State('event-prep-table', 'data')],
        prevent_initial_call=True
    )
    def clear_selected_items(clear_event_clicks, clear_phase_clicks, selected_rows, table_data):
        """Clear the Event or Phase column for selected rows"""
        if not selected_rows or not table_data:
            return "No rows selected", dash.no_update, dash.no_update
        
        try:
            ctx = dash.callback_context
            if not ctx.triggered:
                return "No action selected", dash.no_update, dash.no_update
            
            # Determine which button was clicked
            trigger_id = ctx.triggered[0]['prop_id'].split('.')[0]
            
            # Get selected rows data
            selected_data = [table_data[i] for i in selected_rows]
            
            # Check if we have app.event_data
            if app.event_data.empty:
                return "No data loaded", dash.no_update, dash.no_update
            
            # Make a copy to avoid modifying the original
            df_copy = app.event_data.copy()
            
            # Clear the appropriate column
            if trigger_id == 'clear-selected-event-button':
                # Check if we have any selected rows
                if not selected_data:
                    return "No rows selected", dash.no_update, dash.no_update
                
                # Get unique events from selected rows
                events_to_clear = set()
                for row in selected_data:
                    if row['Event'] and row['Event'] != '':
                        events_to_clear.add(row['Event'])
                
                if not events_to_clear:
                    return "No events found in selected rows", dash.no_update, dash.no_update
                
                # Clear all rows with these events
                for event in events_to_clear:
                    mask = df_copy['Event'] == event
                    df_copy.loc[mask, 'Event'] = ''
                
                # Also remove from app.event_selections
                app.event_selections = [e for e in app.event_selections if e[2] not in events_to_clear]
                
                message = f"Cleared events: {', '.join(events_to_clear)}"
            
            elif trigger_id == 'clear-base-point-button':
                # Check if we have any selected rows
                if not selected_data:
                    return "No rows selected", dash.no_update, dash.no_update
                
                # Get unique bases from selected rows
                bases_to_clear = set()
                for row in selected_data:
                    if row['Mnl_Base'] and row['Mnl_Base'] != '':
                        phases_to_clear.add(row['Mnl_Base'])
                
                if not bases_to_clear:
                    return "No phases found in selected rows", dash.no_update, dash.no_update
                
                # Clear all rows with these bases
                for base in bases_to_clear:
                    mask = df_copy['Mnl_Base'] == base
                    df_copy.loc[mask, 'Mnl_Base'] = ''
                
                # Also remove from app.base_selections
                app.base_selections = [b for b in app.base_selections if p[2] not in bases_to_clear]
                
                message = f"Cleared bases: {', '.join(bases_to_clear)}"
            
            elif trigger_id == 'clear-peak-point-button':
                # Check if we have any selected rows
                if not selected_data:
                    return "No rows selected", dash.no_update, dash.no_update
                
                # Get unique peaks from selected rows
                peaks_to_clear = set()
                for row in selected_data:
                    if row['Mnl_Peak'] and row['Mnl_Peak'] != '':
                        phases_to_clear.add(row['Mnl_Peak'])
                
                if not peaks_to_clear:
                    return "No phases found in selected rows", dash.no_update, dash.no_update
                
                # Clear all rows with these peaks
                for peak in peaks_to_clear:
                    mask = df_copy['Mnl_Peak'] == peak
                    df_copy.loc[mask, 'Mnl_Peak'] = ''
                
                # Also remove from app.peak_selections
                app.peak_selections = [k for k in app.peak_selections if p[2] not in peaks_to_clear]
                
                message = f"Cleared peaks: {', '.join(peaks_to_clear)}"
            
            elif trigger_id == 'clear-selected-phase-button':
                # Check if we have any selected rows
                if not selected_data:
                    return "No rows selected", dash.no_update, dash.no_update
                
                # Get unique phases from selected rows
                phases_to_clear = set()
                for row in selected_data:
                    if row['Phase'] and row['Phase'] != '':
                        phases_to_clear.add(row['Phase'])
                
                if not phases_to_clear:
                    return "No phases found in selected rows", dash.no_update, dash.no_update
                
                # Clear all rows with these phases
                for phase in phases_to_clear:
                    mask = df_copy['Phase'] == phase
                    df_copy.loc[mask, 'Phase'] = ''
                
                # Also remove from app.phase_selections
                app.phase_selections = [p for p in app.phase_selections if p[2] not in phases_to_clear]
                
                message = f"Cleared phases: {', '.join(phases_to_clear)}"
            
            else:
                return "No action taken", dash.no_update, dash.no_update
            
            # Update app.event_data with modified DataFrame
            app.event_data = df_copy
            
            # Update the graph
            updated_fig = create_event_preparation_graph(df_copy)
            
            return html.Div(message, style={'color': 'green'}), updated_fig, df_copy.to_dict('records')
        
        except Exception as e:
            print(f"Error clearing selection: {str(e)}")
            traceback.print_exc()
            return html.Div(f"Error: {str(e)}", style={'color': 'red'}), dash.no_update, dash.no_update

    # 27. VAS slider callback
    @app.callback(
        [Output('vas-add-status', 'children'),
         Output('event-prep-graph', 'figure', allow_duplicate=True),
         Output('event-prep-table', 'data', allow_duplicate=True)],
        Input('add-vas-button', 'n_clicks'),
        [State('vas-slider', 'value'),
         State('event-prep-graph', 'clickData'),
         State('event-prep-graph', 'figure'),
         State('event-prep-table', 'data')],
        prevent_initial_call=True
    )
    def add_vas_value(n_clicks, vas_value, click_data, current_figure, current_data):
        """Add VAS value at the selected point"""
        if not n_clicks or app.selected_point is None:
            return "", dash.no_update, dash.no_update
        
        try:
            if app.event_data.empty:
                return html.Div("No data loaded", style={'color': 'red'}), dash.no_update, dash.no_update
            
            # Update the VAS column
            updated_df = update_vas_column(app.event_data, app.selected_point, vas_value)
            app.event_data = updated_df
            
            # Create success message
            status = html.Div(f"Added VAS value: {vas_value}", style={'color': 'green'})
            
            # Update the graph
            updated_fig = create_event_preparation_graph(updated_df)
            
            return status, updated_fig, updated_df.to_dict('records')
        
        except Exception as e:
            print(f"Error adding VAS value: {str(e)}")
            traceback.print_exc()
            return html.Div(f"Error: {str(e)}", style={'color': 'red'}), dash.no_update, dash.no_update
    
    # 28 Save work in progress
    # Function to save work in progress
    def save_work_in_progress(df, filename, project_id):
        """
        Save the current work in progress to an Excel file in the project's trace_data folder.
        
        Parameters:
        -----------
        df : pandas.DataFrame
            The dataframe containing the current state of the data
        filename : str
            The original filename (without extension)
        project_id : str
            The ID of the current project
            
        Returns:
        --------
        tuple
            (success, message) where success is a boolean and message is a string
        """
        if df.empty:
            return False, "No data to save"
        
        if not project_id:
            return False, "No project selected"
        
        try:
            # Create the trace_data directory path
            trace_data_dir = os.path.join(PROJECTS_DIRECTORY, project_id, "trace_data")
            
            # Ensure the directory exists
            if not os.path.exists(trace_data_dir):
                os.makedirs(trace_data_dir)
                print(f"Created trace_data directory: {trace_data_dir}")
            
            # Create output filename with "_wip" suffix - ensuring we don't add _wip twice
            base_name = os.path.splitext(filename)[0]
            if base_name.lower().endswith('_wip'):
                output_filename = f"{base_name}.xlsx"
            else:
                output_filename = f"{base_name}_wip.xlsx"
                
            output_path = os.path.join(trace_data_dir, output_filename)
            
            print(f"Saving work in progress to: {output_path}")
            
            # Save data to Excel file with multiple sheets
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                # Main data sheet
                df.to_excel(writer, sheet_name='Data', index=False)
                
                # Save selections as separate tables for easier reconstruction
                # Events
                if hasattr(app, 'event_selections') and app.event_selections:
                    event_df = pd.DataFrame(app.event_selections, 
                                        columns=['start_offset', 'end_offset', 'label', 'start_time', 'end_time'])
                    event_df.to_excel(writer, sheet_name='Events', index=False)
                
                # Phases
                if hasattr(app, 'phase_selections') and app.phase_selections:
                    phase_df = pd.DataFrame(app.phase_selections, 
                                        columns=['start_offset', 'end_offset', 'label', 'start_time', 'end_time'])
                    phase_df.to_excel(writer, sheet_name='Phases', index=False)
                
                # Base Points
                if hasattr(app, 'base_selections') and app.base_selections:
                    base_df = pd.DataFrame(app.base_selections, 
                                        columns=['offset', 'label', 'time', 'value'])
                    base_df.to_excel(writer, sheet_name='Bases', index=False)
                
                # Peak Points
                if hasattr(app, 'peak_selections') and app.peak_selections:
                    peak_df = pd.DataFrame(app.peak_selections, 
                                        columns=['offset', 'label', 'time', 'value'])
                    peak_df.to_excel(writer, sheet_name='Peaks', index=False)
            
            print(f"Successfully saved work in progress to {output_path}")
            return True, f"Work in progress saved as {output_filename} in project's trace_data folder"
        
        except Exception as e:
            print(f"Error saving work in progress: {str(e)}")
            traceback.print_exc()
            return False, f"Error saving work in progress: {str(e)}"

    @app.callback(
        [Output('save-wip-status', 'children'),
        Output('last-saved-wip-file', 'data')],
        Input('save-wip-button', 'n_clicks'),
        prevent_initial_call=True
    )
    def save_current_progress(n_clicks):
        """Save the current state of the data as work in progress"""
        if not n_clicks:
            return "", None
        
        try:
            if app.event_data.empty:
                return html.Div("No data to save", style={'color': 'red'}), None
            
            if app.original_filename is None:
                return html.Div("No file selected", style={'color': 'red'}), None
            
            # Check if a project is selected
            if not hasattr(app, 'current_project_id') or not app.current_project_id:
                return html.Div("No project loaded. Please select a project first.", style={'color': 'red'}), None
            
            # Save data to work in progress file
            success, message = save_work_in_progress(
                app.event_data, 
                app.original_filename, 
                app.current_project_id
            )
            
            # Create a styled message based on success
            style = {'color': 'green'} if success else {'color': 'red'}
            
            # Store the saved filename to trigger dropdown refresh
            if success:
                base_name = os.path.splitext(app.original_filename)[0]
                if base_name.lower().endswith('_wip'):
                    saved_file = f"{base_name}.xlsx"
                else:
                    saved_file = f"{base_name}_wip.xlsx"
                return html.Div(message, style=style), saved_file
            
            return html.Div(message, style=style), None
        
        except Exception as e:
            print(f"Error saving work in progress: {str(e)}")
            traceback.print_exc()
            return html.Div(f"Error saving work in progress: {str(e)}", style={'color': 'red'}), None
    
    # 29. Update callback to use save_to_project_peak_data
    @app.callback(
        Output('save-status', 'children'),
        Input('save-processed-button', 'n_clicks'),
        [State('save-filename-input', 'value')],
        prevent_initial_call=True
    )
    def save_processed_data_with_filename(n_clicks, custom_filename):
        """Save the processed data to the project's peak_data folder with _manual suffix"""
        if not n_clicks:
            return ""
        
        try:
            if app.event_data.empty:
                return html.Div("No data to save", style={'color': 'red'})
            
            if app.original_filename is None:
                return html.Div("No file selected", style={'color': 'red'})
            
            # Check if a project is selected
            if not hasattr(app, 'current_project_id') or not app.current_project_id:
                return html.Div("No project loaded. Please select a project first.", style={'color': 'red'})
            
            # Use custom filename if provided, otherwise use original filename
            filename_to_use = custom_filename if custom_filename else app.original_filename
            
            # Create the current graph
            current_fig = create_event_preparation_graph(app.event_data)
            
            # Save data and graph to project's peak_data folder
            success, message = save_to_project_peak_data(
                app.event_data, 
                current_fig, 
                filename_to_use, 
                app.current_project_id
            )
            
            # Create a styled message based on success
            style = {'color': 'green'} if success else {'color': 'red'}
            return html.Div(message, style=style)
        
        except Exception as e:
            print(f"Error saving data: {str(e)}")
            traceback.print_exc()
            return html.Div(f"Error saving data: {str(e)}", style={'color': 'red'})

# Register all callbacks
register_event_preparation_callbacks(app)
# END EVENT PREPARATION TAB #


# In[4]:


##### PEAK DETECTION TAB 1 #####
# import dash
# from dash import dcc, html, dash_table
# import plotly.graph_objects as go
# Define the project selection layout
project_selection_layout = html.Div([
    html.H3('Project Selection', className='section-header'),
    html.Div([
        html.Label('Select Project:', style={'marginRight': '10px', 'fontWeight': 'bold'}),
        dcc.Dropdown(
            id='peak-detection-project-dropdown',
            options=[],  # Will be populated by callback
            placeholder="Select a project",
            style={'width': '300px', 'marginRight': '10px'}
        ),
        html.Button(
            'Refresh Projects',
            id='peak-detection-refresh-projects-button',
            className='action-button'
        ),
    ], style={'display': 'flex', 'alignItems': 'center', 'marginBottom': '10px'}),
    
    html.Button(
        'Load Project',
        id='load-peak-detection-project-button',
        className='primary-button',
        n_clicks=0
    ),
    html.Div(id='load-peak-detection-project-status', className='status-message'),
    
    html.Div([
        html.Label('Select Peak Data File:', style={'marginRight': '10px', 'fontWeight': 'bold'}),
        dcc.Dropdown(
            id='peak-detection-peak-data-files-dropdown',
            options=[],  # Will be populated by callback
            placeholder="Select a peak data file",
            multi=True,  # Allow selecting multiple files
            style={'width': '500px', 'marginRight': '10px'}
        ),
        html.Button(
            'Refresh Files',
            id='peak-detection-refresh-files-button',
            className='action-button',
            style={'marginLeft': '10px'}
        ),
    ], style={'display': 'flex', 'alignItems': 'center', 'marginBottom': '10px', 'marginTop': '20px'}),
    
    html.Button(
        'Load Selected Files',
        id='load-peak-data-files-button',
        className='primary-button',
        n_clicks=0
    ),
    html.Div(id='load-peak-data-files-status', className='status-message'),
], style={'marginBottom': '20px', 'padding': '20px', 'backgroundColor': '#f8f9fa', 'borderRadius': '5px'})

# Define the P1 Analysis section layout with worksheet selector
p1_analysis_layout = html.Div([
    html.H2('P1 Analysis', className='section-header'),
    html.Div([
        html.Div([
            html.Label('Select P1 Worksheets:', style={'marginBottom': '5px', 'fontWeight': 'bold'}),
            dcc.Dropdown(
                id='p1-worksheet-selector',
                options=[],  # Will be populated by callback
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
                className='primary-button',
                n_clicks=0
            ),
        ], style={'display': 'flex', 'alignItems': 'center'}),
        # html.Div(id='p1-save-status', className='status-message')
    ], style={'display': 'flex', 'alignItems': 'flex-end', 'marginBottom': '10px'}),
    
    html.Div([
        html.Button(
            'Select All', 
            id='p1-select-all',
            className='action-button',
            n_clicks=0,
            style={'marginRight': '10px'}
        ),
        html.Button(
            'Deselect All', 
            id='p1-deselect-all',
            className='action-button',
            n_clicks=0
        ),
        dcc.ConfirmDialog(
            id='p1-confirm-deselect',
            message='Are you sure you want to deselect all rows?',
        ),
    ], style={'marginBottom': '10px'}),
    
    dcc.Graph(id='p1-graph'),
    html.H3('P1 Inflection Points', style={'marginTop': '20px', 'marginBottom': '10px'}),
    
    # P1 Table
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
            {"name": "pos_delta", "id": "pos_delta", "type": "numeric", "format": {"specifier": ".2f"}},
            {"name": "neg_delta", "id": "neg_delta", "type": "numeric", "format": {"specifier": ".2f"}},
            {"name": "VAS", "id": "VAS", "type": "numeric", "format": {"specifier": ".1f"}},
            {"name": "Notes", "id": "Notes", "presentation": "input"}
        ],
        data=[],
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
], style={'marginBottom': '30px', 'padding': '20px', 'backgroundColor': '#ffffff', 'borderRadius': '5px', 'boxShadow': '0 2px 4px rgba(0,0,0,0.1)'})

# Define some CSS styles to use in the app
app_css = {
    'action-button': {
        'padding': '5px 10px',
        'backgroundColor': '#f8f9fa',
        'border': '1px solid #ddd',
        'borderRadius': '4px',
        'cursor': 'pointer',
        'height': '36px'
    },
    'primary-button': {
        'padding': '8px 15px',
        'backgroundColor': '#007bff',
        'color': 'white',
        'border': 'none',
        'borderRadius': '4px',
        'cursor': 'pointer',
        'height': '36px',
        'marginRight': '10px'
    },
    'status-message': {
        'marginLeft': '10px',
        'marginTop': '5px'
    },
    'section-header': {
        'marginBottom': '15px',
        'color': '#333'
    }
}

# Define the P3 Analysis section layout with worksheet selector
p3_analysis_layout = html.Div([
    html.H2('P3 Analysis', className='section-header'),
    html.Div([
        html.Div([
            html.Label('Select P3 Worksheets:', style={'marginBottom': '5px', 'fontWeight': 'bold'}),
            dcc.Dropdown(
                id='p3-worksheet-selector',
                options=[],  # Will be populated by callback
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
                className='primary-button',
                n_clicks=0
            ),
        ], style={'display': 'flex', 'alignItems': 'center'}),
        # html.Div(id='p3-save-status', className='status-message')
    ], style={'display': 'flex', 'alignItems': 'flex-end', 'marginBottom': '10px'}),
    
    html.Div([
        html.Button(
            'Select All', 
            id='p3-select-all',
            className='action-button',
            n_clicks=0,
            style={'marginRight': '10px'}
        ),
        html.Button(
            'Deselect All', 
            id='p3-deselect-all',
            className='action-button',
            n_clicks=0
        ),
        dcc.ConfirmDialog(
            id='p3-confirm-deselect',
            message='Are you sure you want to deselect all rows?',
        ),
    ], style={'marginBottom': '10px'}),
    
    dcc.Graph(id='p3-graph'),
    html.H3('P3 Inflection Points', style={'marginTop': '20px', 'marginBottom': '10px'}),
    
    # P3 Table
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
            {"name": "pos_delta", "id": "pos_delta", "type": "numeric", "format": {"specifier": ".2f"}},
            {"name": "neg_delta", "id": "neg_delta", "type": "numeric", "format": {"specifier": ".2f"}},
            {"name": "VAS", "id": "VAS", "type": "numeric", "format": {"specifier": ".1f"}},
            {"name": "Notes", "id": "Notes", "presentation": "input"}
        ],
        data=[],
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
], style={'marginBottom': '30px', 'padding': '20px', 'backgroundColor': '#ffffff', 'borderRadius': '5px', 'boxShadow': '0 2px 4px rgba(0,0,0,0.1)'})

# Define the P4 Analysis section layout with worksheet selector
p4_analysis_layout = html.Div([
    html.H2('P4 Analysis', className='section-header'),
    html.Div([
        html.Div([
            html.Label('Select P4 Worksheets:', style={'marginBottom': '5px', 'fontWeight': 'bold'}),
            dcc.Dropdown(
                id='p4-worksheet-selector',
                options=[],  # Will be populated by callback
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
                className='primary-button',
                n_clicks=0
            ),
        ], style={'display': 'flex', 'alignItems': 'center'}),
        # html.Div(id='p4-save-status', className='status-message')
    ], style={'display': 'flex', 'alignItems': 'flex-end', 'marginBottom': '10px'}),
    
    html.Div([
        html.Button(
            'Select All', 
            id='p4-select-all',
            className='action-button',
            n_clicks=0,
            style={'marginRight': '10px'}
        ),
        html.Button(
            'Deselect All', 
            id='p4-deselect-all',
            className='action-button',
            n_clicks=0
        ),
        dcc.ConfirmDialog(
            id='p4-confirm-deselect',
            message='Are you sure you want to deselect all rows?',
        ),
    ], style={'marginBottom': '10px'}),
    
    dcc.Graph(id='p4-graph'),
    html.H3('P4 Inflection Points', style={'marginTop': '20px', 'marginBottom': '10px'}),
    
    # P4 Table
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
            {"name": "pos_delta", "id": "pos_delta", "type": "numeric", "format": {"specifier": ".2f"}},
            {"name": "neg_delta", "id": "neg_delta", "type": "numeric", "format": {"specifier": ".2f"}},
            {"name": "VAS", "id": "VAS", "type": "numeric", "format": {"specifier": ".1f"}},
            {"name": "Notes", "id": "Notes", "presentation": "input"}
        ],
        data=[],
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
], style={'marginBottom': '30px', 'padding': '20px', 'backgroundColor': '#ffffff', 'borderRadius': '5px', 'boxShadow': '0 2px 4px rgba(0,0,0,0.1)'})

# Main application layout
def create_app_layout():
    """Create the main application layout"""
    return html.Div([
        dcc.Store(id='project-storage', storage_type='session'),
        dcc.Store(id='peak-data-file-store', storage_type='memory'),
        dcc.Store(id='sheet-data-store', storage_type='memory'),
        dcc.Store(id='p1-sheet-data-store', storage_type='memory'),
        dcc.Store(id='p3-sheet-data-store', storage_type='memory'),
        dcc.Store(id='p4-sheet-data-store', storage_type='memory'),
        
        html.H1('Peak Detection Analysis Dashboard', 
                style={'textAlign': 'center', 'marginBottom': '20px', 'marginTop': '20px'}),
        
        # Project and File Selection Section
        project_selection_layout,
        
        # Phase Analysis Sections
        p1_analysis_layout,
        p3_analysis_layout,
        p4_analysis_layout,
        
        # This is a hidden div that will trigger initialization
        html.Div(id='initialization-trigger', style={'display': 'none'})
    ])
# END PEAK DETECTION TAB 1


# In[5]:


##### PEAK DETECTION TAB 2 #####
# import os
# import dash
# from dash import dcc, html, dash_table, Input, Output, State, callback_context
# import plotly.graph_objects as go
# import pandas as pd
# import numpy as np
# import traceback
# from dash.exceptions import PreventUpdate
# import re
# Define important constants
PROJECTS_DIRECTORY = os.path.join(os.getcwd(), 'projects')

# Directory structure functions
def get_projects():
    """Get list of project folders with improved logging"""
    try:
        if not os.path.exists(PROJECTS_DIRECTORY):
            print(f"Projects directory not found: {PROJECTS_DIRECTORY}")
            return []
        
        projects = [d for d in os.listdir(PROJECTS_DIRECTORY) 
                 if os.path.isdir(os.path.join(PROJECTS_DIRECTORY, d))]
        
        print(f"Found {len(projects)} projects in {PROJECTS_DIRECTORY}")
        return projects
        
    except Exception as e:
        print(f"Error getting projects: {str(e)}")
        traceback.print_exc()
        return []

def get_peak_data_files(project_name):
    """Get list of peak data files for a specific project with improved logging"""
    try:
        if not project_name:
            print("No project name provided")
            return []
            
        peak_data_path = os.path.join(PROJECTS_DIRECTORY, project_name, 'peak_data')
        
        if not os.path.exists(peak_data_path):
            print(f"Peak data directory not found: {peak_data_path}")
            return []
            
        # Get all files in the peak_data directory
        files = [f for f in os.listdir(peak_data_path) 
                if os.path.isfile(os.path.join(peak_data_path, f))]
        
        print(f"Found {len(files)} files in {peak_data_path}")
        return sorted(files)
    
    except Exception as e:
        print(f"Error reading peak data directory: {str(e)}")
        traceback.print_exc()
        return []

def load_peak_data_file_single(project_name, file_name):
    """
    Load and process peak data from a file with improved logging and correct offset handling
    """
    try:
        if not project_name or not file_name:
            print("Missing project_name or file_name")
            return pd.DataFrame()
            
        file_path = os.path.join(PROJECTS_DIRECTORY, project_name, 'peak_data', file_name)
        
        if not os.path.exists(file_path):
            print(f"Peak data file not found: {file_path}")
            return pd.DataFrame()
            
        print(f"Loading peak data file: {file_path}")
        
        # If this is a trace data file (CSV)
        if file_name.endswith('.csv'):
            try:
                # Read the CSV file
                df = pd.read_csv(file_path)
                
                # Process the data - store original offset
                if 'Offset' in df.columns:
                    # Store original offset for calculations
                    df['original_offset'] = df['Offset'].copy()
                    
                    # Don't convert to seconds here - we'll do that per phase later
                    # This avoids inconsistencies between the table and graph
                
                # Add required columns if they don't exist
                required_columns = ['Offset', 'Value', 'Type', 'Annotation', 'Event', 'Phase', 'selected']
                for col in required_columns:
                    if col not in df.columns:
                        if col == 'selected':
                            df[col] = True
                        else:
                            df[col] = None
                
                # Set sheet_name to file_name
                df['sheet_name'] = file_name
                
                # Handle base and peak data types
                if 'Type' in df.columns:
                    # Ensure Type uses standard naming (Base, Peak) with Manual prefix
                    df['Type'] = df['Type'].apply(lambda x: 
                        'Manual Base' if pd.notna(x) and 'base' in str(x).lower() else 
                        'Manual Peak' if pd.notna(x) and 'peak' in str(x).lower() else x)
                
                # Set mbase and mpeak columns
                df['mbase'] = None
                df['mbase_time'] = None
                df['mpeak'] = None
                df['mpeak_time'] = None
                
                # Fill in base and peak data
                base_mask = df['Type'].astype(str).str.lower().str.contains('base', na=False)
                peak_mask = df['Type'].astype(str).str.lower().str.contains('peak', na=False)
                
                df.loc[base_mask, 'mbase'] = df.loc[base_mask, 'Value']
                df.loc[base_mask, 'mbase_time'] = df.loc[base_mask, 'original_offset']
                
                df.loc[peak_mask, 'mpeak'] = df.loc[peak_mask, 'Value']
                df.loc[peak_mask, 'mpeak_time'] = df.loc[peak_mask, 'original_offset']
                
                # Try to infer Phase if not present
                if 'Phase' not in df.columns or df['Phase'].isna().all():
                    # Try to extract phase info from filename
                    if 'p1' in file_name.lower():
                        df['Phase'] = 'P1'
                    elif 'p3' in file_name.lower():
                        df['Phase'] = 'P3'
                    elif 'p4' in file_name.lower():
                        df['Phase'] = 'P4'
                    else:
                        # Default to no phase - will be filtered out
                        df['Phase'] = None
                
                # Calculate pos_delta and neg_delta
                df = calculate_auto_peak_deltas(df)
                
                # Add Notes column if not exists
                if 'Notes' not in df.columns:
                    df['Notes'] = ''
                
                # Add VAS column if not exists  
                if 'VAS' not in df.columns:
                    df['VAS'] = None
                
                print(f"Loaded CSV file with {len(df)} rows")
                return df
            except Exception as e:
                print(f"Error reading CSV file {file_path}: {str(e)}")
                traceback.print_exc()
                return pd.DataFrame()
                
        # If this is an Excel file
        elif file_name.endswith('.xlsx'):
            try:
                df = pd.read_excel(file_path)
                
                # Process the data - store original offset
                if 'Offset' in df.columns:
                    # Store original offset for calculations
                    df['original_offset'] = df['Offset'].copy()
                    
                    # Don't convert to seconds here - we'll do that per phase later
                
                # Add selected column if it doesn't exist
                if 'selected' not in df.columns:
                    df['selected'] = True
                
                # Set sheet_name to file_name
                df['sheet_name'] = file_name
                
                # Handle base and peak data types
                if 'Type' in df.columns:
                    # Ensure Type uses standard naming (Base, Peak) with Manual prefix
                    df['Type'] = df['Type'].apply(lambda x: 
                        'Manual Base' if pd.notna(x) and 'base' in str(x).lower() else 
                        'Manual Peak' if pd.notna(x) and 'peak' in str(x).lower() else x)
                
                # Set mbase and mpeak columns if they don't exist
                required_columns = ['mbase', 'mbase_time', 'mpeak', 'mpeak_time', 'Annotation', 'Event', 'Phase']
                for col in required_columns:
                    if col not in df.columns:
                        df[col] = None
                
                # Fill in base and peak data
                base_mask = df['Type'].astype(str).str.lower().str.contains('base', na=False)
                peak_mask = df['Type'].astype(str).str.lower().str.contains('peak', na=False)
                
                df.loc[base_mask, 'mbase'] = df.loc[base_mask, 'Value']
                df.loc[base_mask, 'mbase_time'] = df.loc[base_mask, 'original_offset']
                
                df.loc[peak_mask, 'mpeak'] = df.loc[peak_mask, 'Value']
                df.loc[peak_mask, 'mpeak_time'] = df.loc[peak_mask, 'original_offset']
                
                # Try to infer Phase if not present
                if 'Phase' not in df.columns or df['Phase'].isna().all():
                    # Try to extract phase info from filename
                    if 'p1' in file_name.lower():
                        df['Phase'] = 'P1'
                    elif 'p3' in file_name.lower():
                        df['Phase'] = 'P3'
                    elif 'p4' in file_name.lower():
                        df['Phase'] = 'P4'
                    else:
                        # Default to no phase - will be filtered out
                        df['Phase'] = None
                
                # Calculate pos_delta and neg_delta
                df = calculate_auto_peak_deltas(df)
                
                # Add Notes column if not exists
                if 'Notes' not in df.columns:
                    df['Notes'] = ''
                
                # Add VAS column if not exists  
                if 'VAS' not in df.columns:
                    df['VAS'] = None
                
                print(f"Loaded Excel file with {len(df)} rows")
                return df
            except Exception as e:
                print(f"Error reading Excel file {file_path}: {str(e)}")
                traceback.print_exc()
                return pd.DataFrame()
        else:
            print(f"Unsupported file format: {file_path}")
            return pd.DataFrame()
            
    except Exception as e:
        print(f"Error loading peak data from {file_name}: {str(e)}")
        traceback.print_exc()
        return pd.DataFrame()

def process_peak_data_columns(df):
    """
    Process the specific column structure you mentioned
    Offset, Time, Trace, Value, Annotations, Annotation, Event, Phase, VAS, Mnl_Base, Mnl_Peak, Pod_Delta, Neg_Delta
    """
    try:
        # Ensure original dataframe is preserved
        processed_df = df.copy()
        
        # Handle column name inconsistencies
        if 'Annotations' in processed_df.columns and 'Annotation' not in processed_df.columns:
            processed_df['Annotation'] = processed_df['Annotations']
        
        if 'Pod_Delta' in processed_df.columns:
            processed_df['pos_delta'] = processed_df['Pod_Delta']
        
        if 'Neg_Delta' in processed_df.columns:
            processed_df['neg_delta'] = processed_df['Neg_Delta']
        
        # Convert Offset to seconds if needed
        if 'Offset' in processed_df.columns:
            # Store original value
            processed_df['original_offset'] = processed_df['Offset'].copy()
            
            # Check if values seem large (milliseconds)
            if processed_df['Offset'].median() > 1000:  # Likely in milliseconds
                processed_df['Offset'] = processed_df['Offset'] / 1000
        
        # Set Type column based on Mnl_Base and Mnl_Peak columns
        processed_df['Type'] = None
        
        if 'Mnl_Base' in processed_df.columns:
            base_mask = processed_df['Mnl_Base'].notna() & (processed_df['Mnl_Base'] != 0)
            processed_df.loc[base_mask, 'Type'] = 'Manual Base'
            processed_df.loc[base_mask, 'mbase'] = processed_df.loc[base_mask, 'Mnl_Base']
            processed_df.loc[base_mask, 'mbase_time'] = processed_df.loc[base_mask, 'original_offset']
        
        if 'Mnl_Peak' in processed_df.columns:
            peak_mask = processed_df['Mnl_Peak'].notna() & (processed_df['Mnl_Peak'] != 0)
            processed_df.loc[peak_mask, 'Type'] = 'Manual Peak'
            processed_df.loc[peak_mask, 'mpeak'] = processed_df.loc[peak_mask, 'Mnl_Peak']
            processed_df.loc[peak_mask, 'mpeak_time'] = processed_df.loc[peak_mask, 'original_offset']
        
        # Handle missing required columns
        required_columns = ['Offset', 'Value', 'Type', 'Annotation', 'Event', 'Phase', 'selected',
                          'mbase', 'mbase_time', 'mpeak', 'mpeak_time', 'pos_delta', 'neg_delta',
                          'VAS', 'Notes', 'sheet_name']
        
        for col in required_columns:
            if col not in processed_df.columns:
                if col == 'selected':
                    processed_df[col] = True
                elif col == 'Notes':
                    processed_df[col] = ''
                elif col == 'sheet_name':
                    # Use filename or determine from phase
                    if 'Phase' in processed_df.columns:
                        # Apply to each row based on Phase
                        def get_sheet_name(phase):
                            if pd.isna(phase):
                                return 'Unknown'
                            phase_str = str(phase).upper()
                            if 'P1' in phase_str:
                                return 'P1_Data'
                            elif 'P3' in phase_str:
                                return 'P3_Data'
                            elif 'P4' in phase_str:
                                return 'P4_Data'
                            else:
                                return 'Unknown'
                        
                        processed_df[col] = processed_df['Phase'].apply(get_sheet_name)
                    else:
                        processed_df[col] = 'Imported_Data'
                else:
                    processed_df[col] = None
        
        # Calculate deltas if not provided
        if 'pos_delta' not in df.columns or 'neg_delta' not in df.columns:
            processed_df = calculate_auto_peak_deltas(processed_df)
        
        # Remove temporary columns
        if 'original_offset' in processed_df.columns:
            processed_df = processed_df.drop(columns=['original_offset'])
        
        print(f"Processed dataframe with {len(processed_df)} rows")
        return processed_df
    
    except Exception as e:
        print(f"Error processing peak data columns: {str(e)}")
        traceback.print_exc()
        return df  # Return original dataframe if processing fails

# Make a function to verify the column structure
def verify_peak_data_columns(df):
    """Verify if the dataframe has the expected columns for peak data"""
    expected_columns = [
        'Offset', 'Time', 'Trace', 'Value', 'Annotations', 'Annotation', 
        'Event', 'Phase', 'VAS', 'Mnl_Base', 'Mnl_Peak', 'Pod_Delta', 'Neg_Delta'
    ]
    
    found_columns = [col for col in expected_columns if col in df.columns]
    missing_columns = [col for col in expected_columns if col not in df.columns]
    
    print(f"Found columns: {found_columns}")
    if missing_columns:
        print(f"Missing columns: {missing_columns}")
    
    # These columns are required for basic functionality
    required_columns = ['Offset', 'Value']
    missing_required = [col for col in required_columns if col not in df.columns]
    
    if missing_required:
        print(f"ERROR: Missing required columns: {missing_required}")
        return False
    
    return True

def read_excel_sheets(file_path):
    """Read worksheets from Excel file that contain offset, value, annotation, and event data"""
    sheet_data = {}
    try:
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
                        print(f"Loaded sheet: {sheet_name} with {len(df)} rows")
                    except Exception as e:
                        print(f"Error reading sheet '{sheet_name}': {str(e)}")
    except Exception as e:
        print(f"Error reading Excel file: {str(e)}")
        traceback.print_exc()
    return sheet_data

def calculate_auto_peak_deltas(df):
    """Calculate deltas for the data with selection flag"""
    if df.empty:
        return df
        
    try:
        # Make a copy to avoid modifying the original
        result_df = df.copy()
        
        # Only process rows that are selected
        selected_df = result_df[result_df['selected']]
        
        # Sort by Offset
        selected_df = selected_df.sort_values('Offset')
        
        # Process Base points
        base_points = selected_df[selected_df['Type'] == 'Base']
        
        # Process Peak points
        peak_points = selected_df[selected_df['Type'] == 'Peak']
        
        # Calculate pos_delta for each base (distance to prior peak)
        for idx, base_row in base_points.iterrows():
            # Find all peaks that occur prior to this base
            prior_peaks = peak_points[peak_points['Offset'] < base_row['Offset']]
            
            if not prior_peaks.empty:
                # Find the nearest prior peak
                nearest_prior_peak = prior_peaks.iloc[prior_peaks['Offset'].values.argmax()]
                
                # Calculate pos_delta as the absolute difference
                pos_delta = abs(base_row['Value'] - nearest_prior_peak['Value'])
                
                # Update the pos_delta in the result dataframe
                result_df.at[idx, 'pos_delta'] = pos_delta
        
        # Calculate neg_delta for each peak (distance to prior base)
        for idx, peak_row in peak_points.iterrows():
            # Find all bases that occur prior to this peak
            prior_bases = base_points[base_points['Offset'] < peak_row['Offset']]
            
            if not prior_bases.empty:
                # Find the nearest prior base
                nearest_prior_base = prior_bases.iloc[prior_bases['Offset'].values.argmax()]
                
                # Calculate neg_delta as the absolute difference
                neg_delta = abs(peak_row['Value'] - nearest_prior_base['Value'])
                
                # Update the neg_delta in the result dataframe
                result_df.at[idx, 'neg_delta'] = neg_delta
        
        return result_df
        
    except Exception as e:
        print(f"Error calculating deltas: {str(e)}")
        traceback.print_exc()
        return df

def read_manual_trace_data(file_path):
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
            print(f"Event row values: {event_row.values}")
            
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
                            
                            # Add peak data
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

def calculate_manual_deltas(data):
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

def process_manual_data_files(combined_data):
    """Process a combined dataframe into separate phase dataframes"""
    p1_manual_data = pd.DataFrame()
    p3_manual_data = pd.DataFrame()
    p4_manual_data = pd.DataFrame()
    
    if combined_data.empty:
        return p1_manual_data, p3_manual_data, p4_manual_data
    
    try:
        # Check if Phase column exists
        if 'Phase Name' not in combined_data.columns:
            print("Phase Name not found in data")
            return p1_manual_data, p3_manual_data, p4_manual_data
        
        # Filter data for each phase
        p1_mask = combined_data['Phase'].str.contains('P1|p1', regex=True, na=False)
        p3_mask = combined_data['Phase'].str.contains('P3|p3', regex=True, na=False)
        p4_mask = combined_data['Phase'].str.contains('P4|p4', regex=True, na=False)
        
        # Split data into phase dataframes
        if p1_mask.any():
            p1_data = combined_data[p1_mask].copy()
        
        if p3_mask.any():
            p3_data = combined_data[p3_mask].copy()
        
        if p4_mask.any():
            p4_data = combined_data[p4_mask].copy()
        
        # Calculate deltas for each phase if needed
        for phase_data in [p1_manual_data, p3_manual_data, p4_manual_data]:
            if not phase_data.empty and 'selected' in phase_data.columns:
                # If you have a calculate_deltas_with_selection function, call it here
                # phase_data = calculate_deltas_with_selection(phase_data)
                pass
        
        return p1_manual_data, p3_manual_data, p4_manual_data
    except Exception as e:
        print(f"Error processing peak data files: {str(e)}")
        traceback.print_exc()
        return p1_manual_data, p3_manual_data, p4_manual_data

def load_peak_deltas(n_clicks, table_data, selected_rows):
    """Update selection state for current table data"""
    if not n_clicks or not table_data:
        return dash.no_update
        
    try:
        # Convert table data to dataframe
        df = pd.DataFrame(table_data)
        if df.empty:
            return dash.no_update
            
        # Set selected state based on selected rows
        df['selected'] = True  # Default all to selected
        
        # If specific rows are selected, update selection state
        if selected_rows:
            # First set all to False
            df['selected'] = False
            # Then set selected rows to True
            for idx in selected_rows:
                if idx < len(df):
                    df.at[idx, 'selected'] = True
        
        # No need to recalculate deltas - they're already in the data
        # Just return the updated selection state
        return df.to_dict('records')
        
    except Exception as e:
        print(f"Error updating selection state: {str(e)}")
        traceback.print_exc()
        return dash.no_update

def process_peak_data_files(combined_data):
    """Process a combined dataframe into separate phase dataframes with improved phase detection"""
    p1_data = pd.DataFrame()
    p3_data = pd.DataFrame()
    p4_data = pd.DataFrame()
    
    if combined_data.empty:
        print("No data to process in process_peak_data_files")
        return p1_data, p3_data, p4_data
    
    try:
        # First check if 'Phase' column exists
        if 'Phase' not in combined_data.columns:
            print("Phase column not found in data. Columns available:", combined_data.columns.tolist())
            print("Trying to look for phase information in other columns...")
            
            # Try to infer phase from sheet_name or other columns
            if 'sheet_name' in combined_data.columns:
                print("Attempting to infer phase from sheet_name column")
                combined_data['Phase'] = combined_data['sheet_name'].apply(
                    lambda x: 'P1' if 'P1' in str(x) or 'p1' in str(x) else
                             ('P3' if 'P3' in str(x) or 'p3' in str(x) else
                             ('P4' if 'P4' in str(x) or 'p4' in str(x) else None))
                )
            else:
                print("No suitable column to infer phase information")
                # Create a placeholder Phase column for manual assignment
                combined_data['Phase'] = None
        
        # Convert Phase column to string for consistent processing
        combined_data['Phase'] = combined_data['Phase'].astype(str)
        
        # Print unique phase values for debugging
        print("Phase values detected:", combined_data['Phase'].unique().tolist())
        
        # Create flexible case-insensitive patterns for phase matching
        p1_pattern = r'(?i).*p1.*'
        p3_pattern = r'(?i).*p3.*'
        p4_pattern = r'(?i).*p4.*'
        
        # Filter data using regex patterns for more flexible matching
        p1_mask = combined_data['Phase'].str.contains(p1_pattern, regex=True, na=False)
        p3_mask = combined_data['Phase'].str.contains(p3_pattern, regex=True, na=False)
        p4_mask = combined_data['Phase'].str.contains(p4_pattern, regex=True, na=False)
        
        # Print how many rows matched each pattern
        print(f"Phase matching results - P1: {p1_mask.sum()}, P3: {p3_mask.sum()}, P4: {p4_mask.sum()}")
        
        # Split data into phase dataframes
        if p1_mask.any():
            p1_data = combined_data[p1_mask].copy()
            print(f"Found {len(p1_data)} rows of P1 data")
            
            # Ensure Offset is in seconds
            if 'Offset' in p1_data.columns and p1_data['Offset'].median() > 1000:
                print("Converting P1 Offset from milliseconds to seconds")
                p1_data['Offset'] = p1_data['Offset'] / 1000
        else:
            print("No P1 data found after filtering")
        
        if p3_mask.any():
            p3_data = combined_data[p3_mask].copy()
            print(f"Found {len(p3_data)} rows of P3 data")
            
            # Ensure Offset is in seconds
            if 'Offset' in p3_data.columns and p3_data['Offset'].median() > 1000:
                print("Converting P3 Offset from milliseconds to seconds")
                p3_data['Offset'] = p3_data['Offset'] / 1000
        else:
            print("No P3 data found after filtering")
        
        if p4_mask.any():
            p4_data = combined_data[p4_mask].copy()
            print(f"Found {len(p4_data)} rows of P4 data")
            
            # Ensure Offset is in seconds
            if 'Offset' in p4_data.columns and p4_data['Offset'].median() > 1000:
                print("Converting P4 Offset from milliseconds to seconds")
                p4_data['Offset'] = p4_data['Offset'] / 1000
        else:
            print("No P4 data found after filtering")
        
        # Ensure all dataframes have the required columns
        for phase_data, phase_label in [(p1_data, "P1"), (p3_data, "P3"), (p4_data, "P4")]:
            if not phase_data.empty:
                # Add a selected column if it doesn't exist
                if 'selected' not in phase_data.columns:
                    phase_data['selected'] = True
                    print(f"Added 'selected' column to {phase_label} data")
                
                # Make sure Type column exists
                if 'Type' not in phase_data.columns:
                    print(f"No 'Type' column in {phase_label} data, trying to infer from other columns")
                    phase_data['Type'] = None  # Initialize Type column
                    
                    # Try to infer type from other columns
                    if 'Mnl_Base' in phase_data.columns:
                        # Rows with non-null Mnl_Base are Base points
                        base_mask = phase_data['Mnl_Base'].notna() & (phase_data['Mnl_Base'] != 0)
                        phase_data.loc[base_mask, 'Type'] = 'Manual Base'
                        print(f"Added {base_mask.sum()} 'Manual Base' types to {phase_label} data")
                    
                    if 'Mnl_Peak' in phase_data.columns:
                        # Rows with non-null Mnl_Peak are Peak points
                        peak_mask = phase_data['Mnl_Peak'].notna() & (phase_data['Mnl_Peak'] != 0)
                        phase_data.loc[peak_mask, 'Type'] = 'Manual Peak'
                        print(f"Added {peak_mask.sum()} 'Manual Peak' types to {phase_label} data")
                        
                    # For rows without Type, set a default
                    null_type_mask = phase_data['Type'].isna()
                    if null_type_mask.any():
                        phase_data.loc[null_type_mask, 'Type'] = 'Data Point'
                        print(f"Added default 'Data Point' type to {null_type_mask.sum()} rows in {phase_label} data")
                
                # Calculate deltas if needed
                if ('pos_delta' not in phase_data.columns or phase_data['pos_delta'].isna().all()) and \
                   ('neg_delta' not in phase_data.columns or phase_data['neg_delta'].isna().all()):
                    print(f"Calculating deltas for {phase_label} data")
                    phase_data = calculate_auto_peak_deltas(phase_data)
        
        return p1_data, p3_data, p4_data
    except Exception as e:
        print(f"Error processing peak data files: {str(e)}")
        traceback.print_exc()
        return p1_data, p3_data, p4_data

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

# Initialize app
def initialize_app():
    """Initialize app by setting up required directories and files"""
    print("Initializing application...")
    
    # Make sure projects directory exists
    if not os.path.exists(PROJECTS_DIRECTORY):
        print(f"Creating projects directory: {PROJECTS_DIRECTORY}")
        os.makedirs(PROJECTS_DIRECTORY)
    
    # Additional initialization if needed
    print("Application initialized successfully")
    
    return

# Load data from Excel sheets (All_Phase_Sheets.xlsx) # 1ST UPLOAD
def load_all_phase_sheets(file_path="All_Phase_Sheets.xlsx"):
    """Load phase data from the All_Phase_Sheets Excel file with enhanced error handling"""
    try:
        if not os.path.exists(file_path):
            print(f"All_Phase_Sheets file not found: {file_path}")
            return {}
            
        print(f"Loading phase data from {file_path}")
        
        # Try to load the file directly first to catch file-level issues
        try:
            xls = pd.ExcelFile(file_path)
            print(f"Excel file opened successfully with sheets: {xls.sheet_names}")
        except Exception as e:
            print(f"Error opening Excel file: {str(e)}")
            return {}
        
        return read_excel_sheets(file_path)
    except Exception as e:
        print(f"Error loading All_Phase_Sheets: {str(e)}")
        traceback.print_exc()
        return {}
# END PEAK DETECTION TAB 2 #


# In[6]:


##### NEW PEAK DETECTION TAB 3 #####
# import dash
# from dash import Input, Output, State, callback_context
# from dash.exceptions import PreventUpdate
# import pandas as pd
# import plotly.graph_objects as go
# import traceback
# import os
# import json
# Callback to initialize the project dropdown on app startup

VAS_CONFIG = {
    'show_vas': True,  # Always show VAS
    'value': True,     # Default VAS value
    'color': 'red',    # Default VAS color
    'opacity': 0.3     # Default VAS opacity
}

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

@app.callback(
    [Output('sheet-data-store', 'data'),
     Output('p1-worksheet-selector', 'options'),
     Output('p3-worksheet-selector', 'options'),
     Output('p4-worksheet-selector', 'options')],
    Input('initialization-trigger', 'children')
)
def load_all_worksheet_selectors(_):
    """Load phase sheets from All_Phase_Sheets.xlsx and populate all worksheet selectors"""
    try:
        # Load All_Phase_Sheets.xlsx
        sheet_data = load_all_phase_sheets()
        
        if not sheet_data:
            print("No sheet data loaded from All_Phase_Sheets.xlsx")
            # Return empty values instead of raising exception
            return {}, [], [], []
            
        # Extract sheet names for each phase
        p1_sheets = [name for name in sheet_data.keys() if 'P1' in name]
        p3_sheets = [name for name in sheet_data.keys() if 'P3' in name]
        p4_sheets = [name for name in sheet_data.keys() if 'P4' in name]
        
        # Debug output
        print(f"Found sheets - P1: {len(p1_sheets)}, P3: {len(p3_sheets)}, P4: {len(p4_sheets)}")
        
        p1_options = [{'label': name, 'value': name} for name in sorted(p1_sheets)]
        p3_options = [{'label': name, 'value': name} for name in sorted(p3_sheets)]
        p4_options = [{'label': name, 'value': name} for name in sorted(p4_sheets)]
        
        # Prepare sheet data for storage - use simple structure to avoid serialization issues
        sheet_data_store = {
            'sheet_names': list(sheet_data.keys())
        }
        
        # Add each sheet's data separately to avoid nested dictionary issues
        for sheet_name, sheet_info in sheet_data.items():
            if 'data' in sheet_info and not sheet_info['data'].empty:
                # Convert to records for serialization
                sheet_data_store[f"{sheet_name}_data"] = sheet_info['data'].to_dict('records')
                if 'actions' in sheet_info:
                    sheet_data_store[f"{sheet_name}_actions"] = sheet_info['actions']
        
        return sheet_data_store, p1_options, p3_options, p4_options
        
    except Exception as e:
        print(f"Error loading phase sheets: {str(e)}")
        traceback.print_exc()
        return {}, [], [], []

@app.callback(
    Output('peak-detection-project-dropdown', 'options'),
    Input('initialization-trigger', 'children')
)
def initialize_project_dropdown(_):
    """Initialize the project dropdown with available projects"""
    try:
        projects = get_projects()
        return [{'label': p, 'value': p} for p in sorted(projects)]
    except Exception as e:
        print(f"Error initializing project dropdown: {str(e)}")
        traceback.print_exc()
        return []

# Callback to refresh the project dropdown
@app.callback(
    Output('peak-detection-project-dropdown', 'options', allow_duplicate=True),
    Input('peak-detection-refresh-projects-button', 'n_clicks'),
    prevent_initial_call=True
)
def refresh_projects(n_clicks):
    """Refresh the projects dropdown options"""
    if not n_clicks:
        raise PreventUpdate
        
    try:
        projects = get_projects()
        return [{'label': p, 'value': p} for p in sorted(projects)]
    except Exception as e:
        print(f"Error refreshing projects: {str(e)}")
        return []

# Callback to load the selected project
@app.callback(
    [Output('load-peak-detection-project-status', 'children'),
     Output('load-peak-detection-project-status', 'style'),
     Output('peak-detection-peak-data-files-dropdown', 'options')],
    Input('load-peak-detection-project-button', 'n_clicks'),
    State('peak-detection-project-dropdown', 'value'),
    prevent_initial_call=True
)
def load_selected_project(n_clicks, project_name):
    """Load the selected project and update the peak data files dropdown"""
    if not n_clicks or not project_name:
        raise PreventUpdate
        
    try:
        # Verify project directory exists
        project_path = os.path.join(PROJECTS_DIRECTORY, project_name)
        if not os.path.exists(project_path):
            return (
                f"Project directory not found: {project_path}", 
                {'color': 'red'}, 
                []
            )
        
        # Verify peak_data directory exists
        peak_data_path = os.path.join(project_path, 'peak_data')
        if not os.path.exists(peak_data_path):
            os.makedirs(peak_data_path)
            print(f"Created peak_data directory: {peak_data_path}")
        
        # Get peak data files
        peak_data_files = get_peak_data_files(project_name)
        
        if not peak_data_files:
            return (
                f"No peak data files found in {peak_data_path}", 
                {'color': 'orange'}, 
                []
            )
        
        # Create dropdown options
        file_options = [{'label': f, 'value': f} for f in sorted(peak_data_files)]
        
        return (
            f"Successfully loaded project with {len(peak_data_files)} files", 
            {'color': 'green'}, 
            file_options
        )
            
    except Exception as e:
        print(f"Error loading project: {str(e)}")
        traceback.print_exc()
        return (
            f"Error loading project: {str(e)}", 
            {'color': 'red'}, 
            []
        )

# Callback to refresh the peak data files dropdown
@app.callback(
    Output('peak-detection-peak-data-files-dropdown', 'options', allow_duplicate=True),
    [Input('peak-detection-refresh-files-button', 'n_clicks'),
     Input('peak-detection-project-dropdown', 'value')],
    prevent_initial_call=True
)
def refresh_peak_data_files(n_clicks, project_name):
    """Refresh the peak data files dropdown options"""
    ctx = callback_context
    if not ctx.triggered:
        raise PreventUpdate
        
    # Only respond to button click and ignore dropdown change
    trigger_id = ctx.triggered[0]['prop_id'].split('.')[0]
    if trigger_id != 'peak-detection-refresh-files-button' and n_clicks is None:
        raise PreventUpdate
        
    if not project_name:
        return []
        
    try:
        # Get peak data files
        files = get_peak_data_files(project_name)
        
        # Return options
        return [{'label': f, 'value': f} for f in sorted(files)]
    except Exception as e:
        print(f"Error refreshing peak data files: {str(e)}")
        return []

# Callback to load selected peak data files
@app.callback(
    [Output('load-peak-data-files-status', 'children'),
     Output('load-peak-data-files-status', 'style'),
     Output('peak-data-file-store', 'data')],
    Input('load-peak-data-files-button', 'n_clicks'),
    [State('peak-detection-project-dropdown', 'value'),
     State('peak-detection-peak-data-files-dropdown', 'value')],
    prevent_initial_call=True
)

def load_peak_data_files(n_clicks, project_name, selected_files):
    """Load selected peak data files and store the data with improved phase detection"""
    if not n_clicks:
        raise PreventUpdate
        
    if not project_name:
        return (
            "Please select a project first", 
            {'color': 'red'}, 
            {}
        )
        
    if not selected_files:
        return (
            "Please select one or more files to load", 
            {'color': 'red'}, 
            {}
        )
    
    try:
        # Load and process each selected file
        combined_data = pd.DataFrame()
        loaded_files = []
        
        # Ensure selected_files is a list
        if not isinstance(selected_files, list):
            selected_files = [selected_files]
            
        for file_name in selected_files:
            # Verify file exists
            file_path = os.path.join(PROJECTS_DIRECTORY, project_name, 'peak_data', file_name)
            if not os.path.exists(file_path):
                print(f"File not found: {file_path}")
                continue
                
            print(f"Loading file: {file_path}")
            df = load_peak_data_file_single(project_name, file_name)
            
            # Print detailed info about the loaded data
            if not df.empty:
                print(f"Loaded {len(df)} rows from {file_name}")
                print(f"Columns in loaded data: {df.columns.tolist()}")
                
                # Check if Phase column exists
                if 'Phase' in df.columns:
                    phase_values = df['Phase'].dropna().unique()
                    print(f"Phase values found: {phase_values}")
                else:
                    print("WARNING: No 'Phase' column found in the data")
                
                if combined_data.empty:
                    combined_data = df
                else:
                    combined_data = pd.concat([combined_data, df], ignore_index=True)
                loaded_files.append(file_name)
        
        if combined_data.empty:
            return (
                "No valid data found in selected files", 
                {'color': 'red'}, 
                {}
            )
        
        # Process data by phase with improved handling
        p1_data, p3_data, p4_data = process_peak_data_files(combined_data)
        
        # Debug info
        print(f"After processing: P1 data: {len(p1_data)} rows, P3 data: {len(p3_data)} rows, P4 data: {len(p4_data)} rows")
        
        # Prepare data for storage
        file_store = {
            'loaded': True,
            'p1_data': p1_data.to_dict('records') if not p1_data.empty else [],
            'p3_data': p3_data.to_dict('records') if not p3_data.empty else [],
            'p4_data': p4_data.to_dict('records') if not p4_data.empty else []
        }
        
        # Create status message with details
        status_msg = f"Loaded {len(loaded_files)}/{len(selected_files)} file(s)"
        phase_counts = []
        
        if not p1_data.empty:
            phase_counts.append(f"{len(p1_data)} P1 rows")
        if not p3_data.empty:
            phase_counts.append(f"{len(p3_data)} P3 rows")
        if not p4_data.empty:
            phase_counts.append(f"{len(p4_data)} P4 rows")
            
        if phase_counts:
            status_msg += ": " + ", ".join(phase_counts)
        
        return (
            status_msg, 
            {'color': 'green'}, 
            file_store
        )
            
    except Exception as e:
        print(f"Error loading peak data files: {str(e)}")
        traceback.print_exc()
        return (
            f"Error: {str(e)}", 
            {'color': 'red'}, 
            {}
        )

# Create P1 figure function # P1 1ST LAYOUT 
def create_p1_figure(df):
    """Create a figure for P1 analysis based on data"""
    if df.empty:
        fig = go.Figure()
        fig.update_layout(
            title='No P1 Data Available',
            annotations=[dict(
                text="No P1 data to display",
                xref="paper", yref="paper",
                x=0.5, y=0.5,
                showarrow=False
            )]
        )
        return fig
    
    fig = go.Figure()
    
    # Add line trace for all data
    fig.add_trace(
        go.Scatter(
            name='Raw Data',
            x=df['Offset'],
            y=df['Value'],
            mode='lines',
            line=dict(color='rgb(0, 140, 255)', width=2),
            showlegend=True
        )
    )
    
    # Add points grouped by Type if the column exists
    if 'Type' in df.columns:
        for point_type in df['Type'].unique():
            if pd.isna(point_type):
                continue
                
            type_data = df[df['Type'] == point_type]
            
            # Determine marker properties based on type
            if 'Rolling Base' in str(point_type):
                marker_color = 'rgb(102, 255, 51)'  # Light green for rolling bases
                marker_symbol = 'circle'
            elif 'Rolling Peak' in str(point_type):
                marker_color = 'rgb(0, 176, 240)'  # Light blue for rolling peaks
                marker_symbol = 'circle'
            elif 'Base' in str(point_type):
                marker_color = 'rgb(255, 165, 0)'  # Orange for regular bases
                marker_symbol = 'circle'
            elif 'Peak' in str(point_type):
                marker_color = 'rgb(255, 0, 0)'  # Red for regular peaks
                marker_symbol = 'circle'
            else:
                marker_color = 'rgb(128, 128, 128)'
                marker_symbol = 'circle'
            
            fig.add_trace(
                go.Scatter(
                    name=str(point_type),
                    x=type_data['Offset'],
                    y=type_data['Value'],
                    mode='markers',
                    marker=dict(
                        color=marker_color,
                        size=10,
                        symbol=marker_symbol,
                        line=dict(color='white', width=1)
                    ),
                    showlegend=True
                )
            )
    
    # Update layout
    fig.update_layout(
        title='P1 Analysis',
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
        plot_bgcolor='white',
        paper_bgcolor='white'
    )
    
    return fig

def create_phase_graph_and_table(sheet_data, manual_data, phase, selected_sheets, 
                                 include_words=None, exclude_words=None, include_events=None, 
                                 exclude_events=None, show_actions=False, window_size=1,
                                 peak_data_files=None):
    """
    Create graph and table for a phase with proper integration of manual and automated data
    
    Parameters:
    -----------
    sheet_data : dict
        Dictionary containing sheet data from All_Phase_Sheets
    manual_data : pd.DataFrame
        Manual data from Mnl_Analysis
    phase : str
        Phase identifier (P1, P3, P4)
    selected_sheets : list
        List of selected sheet names
    include_words : list, optional
        Words to include in filtering
    exclude_words : list, optional
        Words to exclude in filtering
    include_events : list, optional
        Events to include in filtering
    exclude_events : list, optional
        Events to exclude in filtering
    show_actions : bool, optional
        Whether to show action labels
    window_size : int, optional
        Window size for rolling detection
    peak_data_files : pd.DataFrame, optional
        Additional peak data from loaded peak_data files
        
    Returns:
    --------
    tuple
        (fig, combined_data) Plotly figure and DataFrame with combined data
    """
    import re

    species_scale = 10
    print(f"\nCreating {phase} graph and table")
    print(f"Selected sheets: {selected_sheets}")
    print(f"Action labels are {'ON' if show_actions else 'OFF'}")
    print(f"Manual data provided: {manual_data is not None and not manual_data.empty}")
    
    # Print information about peak_data_files
    if peak_data_files is not None and not isinstance(peak_data_files, pd.DataFrame):
        print(f"WARNING: peak_data_files is not a DataFrame, type: {type(peak_data_files)}")
        if isinstance(peak_data_files, list):
            # Convert list to DataFrame
            peak_data_files = pd.DataFrame(peak_data_files)
            print(f"Converted peak_data_files list to DataFrame with {len(peak_data_files)} rows")
    
    if peak_data_files is not None and not peak_data_files.empty:
        print(f"Peak data provided with {len(peak_data_files)} rows")
        print(f"Peak data columns: {peak_data_files.columns.tolist()}")
        if 'Type' in peak_data_files.columns:
            type_counts = peak_data_files['Type'].value_counts().to_dict()
            print(f"Type distribution in peak data: {type_counts}")
    else:
        print("No peak data files provided or empty")
    
    # Create subplot figure with secondary y-axis for VAS bars
    fig = sp.make_subplots(specs=[[{"secondary_y": True}]])
    all_data = []
    
    # Return empty figure and DataFrame if no sheets are selected and no manual data and no peak data
    if not selected_sheets and (manual_data is None or manual_data.empty) and (peak_data_files is None or peak_data_files.empty):
        empty_df = pd.DataFrame(columns=['sheet_name', 'Type', 'Annotation', 'Event', 'Offset', 'Value', 
                                       'mbase', 'mbase_time', 'mpeak', 'mpeak_time', 
                                       'pos_delta', 'neg_delta', 'VAS', 'Notes', 'selected'])
        empty_df['selected'] = True
        print(f"\nNo data selected for {phase}, returning empty DataFrame")
        return fig, empty_df
    
    # Filter sheets for the specified phase
    phase_sheets = {name: data for name, data in sheet_data.items() 
                  if phase in name and (not selected_sheets or name in selected_sheets)}
    
    print(f"Processing sheets: {list(phase_sheets.keys())}")
    
    # FIRST PASS: Process sheet data for rolling peak detection
    for sheet_name, sheet_info in phase_sheets.items():
        print(f"\nProcessing sheet: {sheet_name}")
        data = sheet_info['data']
        
        # Skip if data is empty
        if data.empty:
            print(f"Sheet {sheet_name} has no data, skipping")
            continue
            
        # Filter data based on words and events
        filtered_data = filter_by_words(data.copy(), include_words, exclude_words)
        filtered_data = filter_by_events(filtered_data, include_events, exclude_events)

        if not filtered_data.empty:
            # Add main line trace
            fig.add_trace(
                go.Scatter(
                    name=f'{sheet_name}',
                    x=filtered_data['Offset']/1000,
                    y=filtered_data['Value']*species_scale,
                    mode='lines',
                    line=dict(color=COLORS['line'], width=2),
                    showlegend=True
                ),
                secondary_y=False
            )

            # Add event span lines
            events = filtered_data.groupby('Event')
            used_y_positions = []
            
            for event_name, event_group in events:
                if pd.notna(event_name) and event_name != '':
                    # Get min and max offset for this event
                    min_offset = event_group['Offset'].min() / 1000
                    max_offset = event_group['Offset'].max() / 1000
                    
                    # Add vertical lines
                    fig.add_trace(
                        go.Scatter(
                            x=[min_offset, min_offset],
                            y=[filtered_data['Value'].min() * species_scale, filtered_data['Value'].max() * species_scale],
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
                            y=[filtered_data['Value'].min() * species_scale, filtered_data['Value'].max() * species_scale],
                            mode='lines',
                            line=dict(color='rgba(128, 128, 128, 0.5)', width=1, dash='dash'),
                            name=f'Event {event_name} End',
                            showlegend=False
                        ),
                        secondary_y=False
                    )
                    
                    # Find available y position for label
                    base_y = filtered_data['Value'].max() * species_scale
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
                        y=filtered_data['Value'].max() * species_scale,
                        text=f'Event {event_name}',
                        showarrow=False,
                        yshift=10,
                        textangle=45,
                        font=dict(size=10)
                    )
            
            # Perform rolling window peak detection
            rolling_points = detect_peaks_rolling_window(filtered_data, sheet_name, window_size)

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
                                hoverinfo='text'
                            ),
                            secondary_y=False
                        )
                
                # Add rolling points to the combined data
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

    # SECOND PASS: Process manual data (from Mnl_Analysis.xlsx)
    manual_data_from_file = pd.DataFrame()
    if manual_data is not None and not manual_data.empty:
        print("\nProcessing manual data")
        
        # First filter by phase - make this more robust with case-insensitive matching
        phase_pattern = f"(?i).*{phase}.*"  # Case-insensitive pattern that matches any string containing the phase
        phase_data_filter = manual_data['sheet_name'].astype(str).str.contains(phase_pattern, regex=True, na=False)
        phase_filtered_data = manual_data[phase_data_filter].copy()
        
        if phase_filtered_data.empty:
            print(f"No manual data found for phase {phase}")
        else:
            print(f"Found {len(phase_filtered_data)} rows for phase {phase}")
            print(f"Manual data sheet_names: {phase_filtered_data['sheet_name'].unique().tolist()}")
            
            # If no specific sheets are selected, include all manual data for this phase
            if not selected_sheets:
                print(f"No specific sheets selected, including all manual data for phase {phase}")
                manual_data_from_file = phase_filtered_data
            else:
                # Extract identifiers from selected sheet names
                selected_worksheet_identifiers = []
                for sheet_name in selected_sheets:
                    # Extract identifier from sheet name (e.g. "P1_1" -> "1")
                    parts = sheet_name.split('_')
                    if len(parts) > 1:
                        identifier = parts[1]
                        selected_worksheet_identifiers.append(identifier)
                    else:
                        # Try to extract numeric part from sheet name for other formats
                        import re
                        match = re.search(r'(\d+)', sheet_name)
                        if match:
                            identifier = match.group(1)
                            selected_worksheet_identifiers.append(identifier)
                        else:
                            # Just use the whole sheet name as an identifier
                            selected_worksheet_identifiers.append(sheet_name)
                
                print(f"Extracted worksheet identifiers: {selected_worksheet_identifiers}")
                
                if selected_worksheet_identifiers:
                    # Initialize the mask as all False
                    identifier_matches = pd.Series(False, index=phase_filtered_data.index)
                    
                    # Check and print data IDs for debugging
                    meas_ids = phase_filtered_data['sheet_name'].unique()
                    print(f"Available measurement IDs: {meas_ids}")
                    
                    # For each identifier, update the mask with OR operation
                    for identifier in selected_worksheet_identifiers:
                        # Print what we're searching for
                        print(f"Looking for ID matches with: '{identifier}'")
                        
                        # Store original match count for comparison
                        original_match_count = identifier_matches.sum()
                        
                        # Try exact matches first
                        exact_match = phase_filtered_data['sheet_name'] == identifier
                        if exact_match.any():
                            print(f"Found {exact_match.sum()} exact matches for '{identifier}'")
                            identifier_matches = identifier_matches | exact_match
                        else:
                            # Try exact match with phase prefix (e.g. "1" -> "P1_1")
                            prefixed_match = phase_filtered_data['sheet_name'] == f"{phase}_{identifier}"
                            if prefixed_match.any():
                                print(f"Found {prefixed_match.sum()} prefixed matches for '{phase}_{identifier}'")
                                identifier_matches = identifier_matches | prefixed_match
                            else:
                                # Try flexible regex matching
                                try:
                                    # First try more specific boundary matches
                                    flexible_match = phase_filtered_data['sheet_name'].str.contains(
                                        fr'\b{re.escape(identifier)}\b|\b{phase}[_\s]*{re.escape(identifier)}\b', 
                                        regex=True, 
                                        na=False
                                    )
                                    
                                    if flexible_match.any():
                                        matched_ids = phase_filtered_data.loc[flexible_match, 'sheet_name'].unique()
                                        print(f"Found {flexible_match.sum()} regex matches for '{identifier}': {matched_ids}")
                                        identifier_matches = identifier_matches | flexible_match
                                    else:
                                        # Try even more relaxed matching as a last resort
                                        relaxed_match = phase_filtered_data['sheet_name'].str.contains(
                                            fr'{re.escape(identifier)}', 
                                            regex=True, 
                                            na=False
                                        )
                                        
                                        if relaxed_match.any():
                                            matched_ids = phase_filtered_data.loc[relaxed_match, 'sheet_name'].unique()
                                            print(f"Found {relaxed_match.sum()} relaxed matches for '{identifier}': {matched_ids}")
                                            identifier_matches = identifier_matches | relaxed_match
                                        else:
                                            print(f"No matches found for '{identifier}'")
                                except Exception as e:
                                    print(f"Error in regex matching: {str(e)}")
                                    # Try simple contains as fallback
                                    contains_match = phase_filtered_data['sheet_name'].astype(str).str.contains(identifier, na=False)
                                    if contains_match.any():
                                        print(f"Found {contains_match.sum()} contains matches as fallback")
                                        identifier_matches = identifier_matches | contains_match
                        
                        # Check if we added any new matches
                        new_matches = identifier_matches.sum() - original_match_count
                        if new_matches > 0:
                            print(f"Added {new_matches} new matches with identifier '{identifier}'")
                    
                    # Apply the filter
                    manual_data_from_file = phase_filtered_data[identifier_matches].copy()
                    print(f"After filtering, {len(manual_data_from_file)} rows match selected identifiers")
                    
                    # If no matches found, make it explicit
                    if manual_data_from_file.empty:
                        print(f"WARNING: No manual data matched the selected identifiers: {selected_worksheet_identifiers}")
                        # Fallback - use all phase data if we couldn't match any identifiers
                        manual_data_from_file = phase_filtered_data
                        print(f"Falling back to using all {len(manual_data_from_file)} rows for phase {phase}")
                else:
                    # No valid identifiers were extracted, use all phase data
                    manual_data_from_file = phase_filtered_data
                    print(f"No valid identifiers extracted, using all {len(manual_data_from_file)} rows for phase {phase}")

    # Apply word and event filters to the manual data
    if not manual_data_from_file.empty:
        print(f"Applying word/event filters to {len(manual_data_from_file)} manual data rows")
        if include_words or exclude_words:
            manual_data_from_file = filter_by_words(manual_data_from_file, include_words, exclude_words)
        if include_events or exclude_events:
            manual_data_from_file = filter_by_events(manual_data_from_file, include_events, exclude_events)
        
        # After filtering
        print(f"After word/event filters, {len(manual_data_from_file)} manual data rows remain")
    else:
        print("No manual data to filter")

    # Add manual data to the combined dataset only if any rows passed the filters
    if not manual_data_from_file.empty:
        print(f"Adding {len(manual_data_from_file)} filtered manual data rows to combined dataset")
        manual_data_from_file['selected'] = True  # Default to selected
        all_data.append(manual_data_from_file)
        
        # Add visual representation of manual bases
        manual_bases = manual_data_from_file[manual_data_from_file['Type'] == 'Manual Base']
        if not manual_bases.empty:
            fig.add_trace(
                go.Scatter(
                    name=f'Manual Base ({phase})',
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
        
        # Add visual representation of manual peaks
        manual_peaks = manual_data_from_file[manual_data_from_file['Type'] == 'Manual Peak']
        if not manual_peaks.empty:
            fig.add_trace(
                go.Scatter(
                    name=f'Manual Peak ({phase})',
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
        
        # Add delta labels for manual peaks with negative deltas
        manual_delta_labels = manual_peaks[manual_peaks['neg_delta'].notna()]
        if not manual_delta_labels.empty:
            fig.add_trace(
                go.Scatter(
                    name=f'Manual Peak Deltas ({phase})',
                    x=manual_delta_labels['Offset'],
                    y=[val - 0.3 for val in manual_delta_labels['Value']],
                    mode='text',
                    text=[f'Δ: {delta:.1f}' for delta in manual_delta_labels['neg_delta']],
                    textposition='bottom right',
                    textfont=dict(size=10, color='darkred'),
                    showlegend=False,
                    hoverinfo='text'
                )
            )
    else:
        print(f"No manual data to add for phase {phase}")
    
    # THIRD PASS: Add peak data files (with improved handling)
    if peak_data_files is not None and (isinstance(peak_data_files, pd.DataFrame) or isinstance(peak_data_files, list)):
        # Convert list to DataFrame if needed
        if isinstance(peak_data_files, list):
            peak_data_files = pd.DataFrame(peak_data_files)
        
        if not peak_data_files.empty:
            print(f"Processing peak data files for {phase} with {len(peak_data_files)} rows")
            
            # Make a copy to avoid modifying the original
            peak_data = peak_data_files.copy()
            
            # Filter peak data by phase (case-insensitive)
            if 'Phase' in peak_data.columns:
                # Create a more flexible pattern for phase matching
                phase_pattern = f"{phase.lower()}"
                
                # First try direct string contains
                phase_mask = peak_data['Phase'].astype(str).str.lower().str.contains(phase_pattern, na=False)
                
                if phase_mask.sum() == 0:
                    # If no matches, try more flexible matching with regex
                    phase_pattern = f"^{phase.lower()}$|^{phase.lower()}_|_{phase.lower()}$|_{phase.lower()}_"
                    phase_mask = peak_data['Phase'].astype(str).str.lower().str.contains(phase_pattern, regex=True, na=False)
                
                # Apply the phase filter
                phase_filtered_peak_data = peak_data[phase_mask].copy()
                print(f"After phase filtering, {len(phase_filtered_peak_data)} peak data rows remain for phase {phase}")
            else:
                # No Phase column, try to infer from sheet_name or just use everything
                if 'sheet_name' in peak_data.columns:
                    phase_mask = peak_data['sheet_name'].astype(str).str.lower().str.contains(phase.lower(), na=False)
                    phase_filtered_peak_data = peak_data[phase_mask].copy()
                    print(f"Inferred phase from sheet_name, found {len(phase_filtered_peak_data)} rows for {phase}")
                else:
                    # No way to filter by phase, use all data (might mix phases)
                    phase_filtered_peak_data = peak_data.copy()
                    print(f"WARNING: No Phase or sheet_name column found, using all {len(phase_filtered_peak_data)} peak data rows")
            
            # Only proceed if we have data after phase filtering
            if not phase_filtered_peak_data.empty:
                # Make sure all offsets are in seconds for consistency
                if 'Offset' in phase_filtered_peak_data.columns:
                    # Store original values for reference if needed
                    if 'original_offset' not in phase_filtered_peak_data.columns:
                        phase_filtered_peak_data['original_offset'] = phase_filtered_peak_data['Offset'].copy()
                    
                    # Check if conversion is needed (milliseconds to seconds)
                    if phase_filtered_peak_data['Offset'].median() > 1000:
                        print(f"Converting Offset from milliseconds to seconds for {phase} peak data")
                        phase_filtered_peak_data['Offset'] = phase_filtered_peak_data['Offset'] / 1000
                        
                    print(f"Offset range for {phase} peak data: {phase_filtered_peak_data['Offset'].min():.2f}s - {phase_filtered_peak_data['Offset'].max():.2f}s")

                # Apply word and event filters
                if include_words or exclude_words:
                    phase_filtered_peak_data = filter_by_words(phase_filtered_peak_data, include_words, exclude_words)
                if include_events or exclude_events:
                    phase_filtered_peak_data = filter_by_events(phase_filtered_peak_data, include_events, exclude_events)
                
                print(f"After word/event filters, {len(phase_filtered_peak_data)} peak data rows remain")
                
                # Only proceed if we still have data after all filtering
                if not phase_filtered_peak_data.empty:
                    # Ensure all required columns exist
                    required_columns = ['Offset', 'Value', 'Type', 'Annotation', 'Event', 'sheet_name', 
                                    'mbase', 'mbase_time', 'mpeak', 'mpeak_time', 'pos_delta', 'neg_delta', 
                                    'VAS', 'Notes', 'selected']
                    
                    for col in required_columns:
                        if col not in phase_filtered_peak_data.columns:
                            if col == 'selected':
                                phase_filtered_peak_data[col] = True
                            elif col == 'Notes':
                                phase_filtered_peak_data[col] = ''
                            elif col == 'sheet_name' and 'Phase' in phase_filtered_peak_data.columns:
                                phase_filtered_peak_data[col] = phase_filtered_peak_data['Phase']
                            elif col == 'Type' and 'Mnl_Base' in phase_filtered_peak_data.columns and 'Mnl_Peak' in phase_filtered_peak_data.columns:
                                # Infer Type from Mnl_Base and Mnl_Peak columns
                                phase_filtered_peak_data[col] = None
                                base_mask = phase_filtered_peak_data['Mnl_Base'].notna() & (phase_filtered_peak_data['Mnl_Base'] != 0)
                                peak_mask = phase_filtered_peak_data['Mnl_Peak'].notna() & (phase_filtered_peak_data['Mnl_Peak'] != 0)
                                
                                phase_filtered_peak_data.loc[base_mask, 'Type'] = 'Manual Base'
                                phase_filtered_peak_data.loc[peak_mask, 'Type'] = 'Manual Peak'
                                
                                # For rows with neither base nor peak
                                other_mask = ~(base_mask | peak_mask)
                                phase_filtered_peak_data.loc[other_mask, 'Type'] = 'Data Point'
                                
                                print(f"Inferred 'Type' column: {base_mask.sum()} bases, {peak_mask.sum()} peaks, {other_mask.sum()} other")
                            else:
                                phase_filtered_peak_data[col] = None
                    
                    # Make sure neg_delta is calculated if possible
                    if 'neg_delta' not in phase_filtered_peak_data.columns or phase_filtered_peak_data['neg_delta'].isna().all():
                        print("Calculating neg_delta values for peak data")
                        phase_filtered_peak_data = calculate_auto_peak_deltas(phase_filtered_peak_data)
                        
                    # Add to combined data for table display
                    all_data.append(phase_filtered_peak_data)
                    
                    # Identify base and peak points for display
                    # Identify base and peak points for display
                    if 'Type' in phase_filtered_peak_data.columns:
                        # Use flexible pattern matching for type detection
                        base_mask = phase_filtered_peak_data['Type'].astype(str).str.lower().str.contains('base', na=False)
                        peak_mask = phase_filtered_peak_data['Type'].astype(str).str.lower().str.contains('peak', na=False)
                        
                        # Extract all bases and peaks first
                        peak_data_bases = phase_filtered_peak_data[base_mask].copy()
                        peak_data_all_peaks = phase_filtered_peak_data[peak_mask].copy()
                        
                        # Ensure we have delta values for all peaks
                        if not peak_data_all_peaks.empty:
                            # Sort by offset to ensure chronological processing
                            all_points = phase_filtered_peak_data.sort_values('Offset').copy()
                            
                            # Initialize or reset neg_delta column if needed
                            if 'neg_delta' not in all_points.columns:
                                all_points['neg_delta'] = None
                            
                            # Calculate deltas for all peaks
                            for idx in all_points.index:
                                if all_points.loc[idx, 'Type'].lower().find('peak') >= 0:
                                    peak_offset = all_points.loc[idx, 'Offset']
                                    peak_value = all_points.loc[idx, 'Value']
                                    
                                    # Find the closest prior base
                                    prior_bases = all_points[
                                        (all_points['Type'].astype(str).str.lower().str.contains('base', na=False)) & 
                                        (all_points['Offset'] < peak_offset)
                                    ]
                                    
                                    if not prior_bases.empty:
                                        # Get the most recent base before this peak
                                        prior_base = prior_bases.loc[prior_bases['Offset'].idxmax()]
                                        base_value = prior_base['Value']
                                        
                                        # Calculate delta (peak - base)
                                        delta = peak_value - base_value
                                        
                                        # Store the delta value
                                        all_points.loc[idx, 'neg_delta'] = delta
                                        
                                        # Debug output
                                        print(f"Calculated delta for peak at offset {peak_offset:.2f}s: {delta:.2f} (peak: {peak_value:.2f}, base: {base_value:.2f})")
                            
                            # Update the peak_data_all_peaks with calculated deltas
                            for idx in peak_data_all_peaks.index:
                                if idx in all_points.index:
                                    peak_data_all_peaks.loc[idx, 'neg_delta'] = all_points.loc[idx, 'neg_delta']
                            
                            # Update the main dataframe too
                            phase_filtered_peak_data = all_points.copy()
                        
                        # Now filter to include only peaks with valid delta values
                        peak_with_delta_mask = peak_mask & phase_filtered_peak_data['neg_delta'].notna()
                        peak_data_peaks = phase_filtered_peak_data[peak_with_delta_mask].copy()
                        
                        print(f"Found {len(peak_data_bases)} bases and {len(peak_data_peaks)} peaks with calculated deltas in peak data")
                    else:
                        # No Type column - try to use mbase and mpeak to identify
                        peak_data_bases = phase_filtered_peak_data[phase_filtered_peak_data['mbase'].notna()].copy()
                        
                        # Calculate deltas for peaks using mpeak and closest prior mbase
                        if 'mpeak' in phase_filtered_peak_data.columns and 'mbase' in phase_filtered_peak_data.columns:
                            # Get all points with mpeak values
                            all_peaks = phase_filtered_peak_data[phase_filtered_peak_data['mpeak'].notna()].copy()
                            all_points = phase_filtered_peak_data.sort_values('Offset').copy()
                            
                            # Calculate deltas
                            for idx in all_peaks.index:
                                peak_offset = all_peaks.loc[idx, 'Offset']
                                peak_value = all_peaks.loc[idx, 'mpeak']
                                
                                # Find prior bases
                                prior_bases = all_points[
                                    (all_points['mbase'].notna()) & 
                                    (all_points['Offset'] < peak_offset)
                                ]
                                
                                if not prior_bases.empty:
                                    prior_base = prior_bases.loc[prior_bases['Offset'].idxmax()]
                                    base_value = prior_base['mbase']
                                    
                                    # Calculate and store delta
                                    delta = peak_value - base_value
                                    all_peaks.loc[idx, 'neg_delta'] = delta
                                    
                                    # Update the main dataframe
                                    phase_filtered_peak_data.loc[idx, 'neg_delta'] = delta
                            
                            # Filter to include only peaks with calculated deltas
                            peak_data_peaks = all_peaks[all_peaks['neg_delta'].notna()].copy()
                        else:
                            peak_data_peaks = pd.DataFrame()  # Empty if no way to calculate deltas

                    # Add peak data base markers (as triangles)
                    if not peak_data_bases.empty:
                        fig.add_trace(
                            go.Scatter(
                                name=f'Peak Data Base ({phase})',
                                x=peak_data_bases['Offset'],  # Offset should now be in seconds
                                y=peak_data_bases['Value'],
                                mode='markers',
                                marker=dict(
                                    color='rgb(0, 100, 0)',  # Darker green for file bases
                                    size=10,
                                    symbol='triangle-up',  # Triangle pointing up for bases
                                    line=dict(color='white', width=1)
                                ),
                                showlegend=True,
                                hovertemplate='Offset: %{x:.2f}s<br>Value: %{y:.2f}<br>Type: Base<br>File: %{customdata}',
                                customdata=peak_data_bases['sheet_name']
                            ),
                            secondary_y=False
                        )

                    # Add peak data peak markers (as triangles) - only those with neg_delta values
                    if not peak_data_peaks.empty:
                        fig.add_trace(
                            go.Scatter(
                                name=f'Peak Data Peaks ({phase})',
                                x=peak_data_peaks['Offset'],  # Offset should now be in seconds
                                y=peak_data_peaks['Value'],
                                mode='markers',
                                marker=dict(
                                    color='rgb(100, 0, 100)',  # Darker purple for file peaks
                                    size=10,
                                    symbol='triangle-down',  # Triangle pointing down for peaks
                                    line=dict(color='white', width=1)
                                ),
                                showlegend=True,
                                hovertemplate='Offset: %{x:.2f}s<br>Value: %{y:.2f}<br>Type: Peak<br>Delta: %{text}<br>File: %{customdata}',
                                text=[f"{delta:.2f}" if pd.notna(delta) else "N/A" for delta in peak_data_peaks['neg_delta']],
                                customdata=peak_data_peaks['sheet_name']
                            ),
                            secondary_y=False
                        )

                    # Add delta labels for peaks with negative deltas
                    if not peak_data_peaks.empty:
                        fig.add_trace(
                            go.Scatter(
                                name=f'Peak Data Deltas ({phase})',
                                x=peak_data_peaks['Offset'],
                                y=[val - 0.3 for val in peak_data_peaks['Value']],
                                mode='text',
                                text=[f'Δ: {delta:.1f}' for delta in peak_data_peaks['neg_delta']],
                                textposition='bottom right',
                                textfont=dict(size=10, color='darkred'),
                                showlegend=False,
                                hoverinfo='text'
                            ),
                            secondary_y=False
                        )
                    
                    # Add delta labels for peaks with negative deltas
                    peak_data_delta_labels = peak_data_peaks[peak_data_peaks['neg_delta'].notna()]
                    if not peak_data_delta_labels.empty:
                        fig.add_trace(
                            go.Scatter(
                                name=f'Peak Data Deltas ({phase})',
                                x=peak_data_delta_labels['Offset'],
                                y=[val - 0.3 for val in peak_data_delta_labels['Value']],
                                mode='text',
                                text=[f'Δ: {delta:.1f}' for delta in peak_data_delta_labels['neg_delta']],
                                textposition='bottom right',
                                textfont=dict(size=10, color='darkred'),
                                showlegend=False,
                                hoverinfo='text'
                            ),
                            secondary_y=False
                        )
                        
                    # Handle any additional custom point types
                    if 'Type' in phase_filtered_peak_data.columns:
                        unique_types = phase_filtered_peak_data['Type'].dropna().unique()
                        for point_type in unique_types:
                            # Skip base and peak types we've already handled
                            point_type_str = str(point_type).lower()
                            if pd.notna(point_type) and not ('base' in point_type_str or 'peak' in point_type_str):
                                # Only include points with neg_delta values if they exist
                                custom_points = phase_filtered_peak_data[
                                    (phase_filtered_peak_data['Type'] == point_type) & 
                                    (phase_filtered_peak_data['neg_delta'].notna() | phase_filtered_peak_data['pos_delta'].notna())
                                ]
                                
                                if not custom_points.empty:
                                    # Use a distinct color based on the type name
                                    color_hash = hash(str(point_type)) % 256
                                    marker_color = f'rgb({color_hash}, {(color_hash + 85) % 256}, {(color_hash + 170) % 256})'
                                    
                                    fig.add_trace(
                                        go.Scatter(
                                            name=f'{point_type} ({phase})',
                                            x=custom_points['Offset'],
                                            y=custom_points['Value'],
                                            mode='markers',
                                            marker=dict(
                                                color=marker_color,
                                                size=10,
                                                symbol='circle',
                                                line=dict(color='white', width=1)
                                            ),
                                            showlegend=True,
                                            hovertemplate='Offset: %{x:.2f}s<br>Value: %{y:.2f}<br>Type: %{text}<br>File: %{customdata}',
                                            text=[str(point_type) for _ in range(len(custom_points))],
                                            customdata=custom_points['sheet_name']
                                        ),
                                        secondary_y=False
                                    )
                    
                    # Remove original_offset column if it was added
                    if 'original_offset' in phase_filtered_peak_data.columns:
                        phase_filtered_peak_data = phase_filtered_peak_data.drop(columns=['original_offset'])
                else:
                    print(f"No peak data files data to add for phase {phase} after filtering")

    # Add action labels for all data sources if enabled
    if show_actions and any(all_data):
        # Combine all data to get event information
        combined_for_events = pd.concat(all_data, ignore_index=True) if all_data else pd.DataFrame()
        
        if not combined_for_events.empty and 'Event' in combined_for_events.columns:
            used_positions = []
            
            # Calculate event averages
            event_stats = {}
            event_starts = {}
            
            # Process events - handle potential null/empty values
            for event_name, event_group in combined_for_events.groupby('Event', dropna=False):
                if pd.notna(event_name) and str(event_name).strip() != '':
                    # Filter out null/NaN values for calculations
                    valid_values = event_group['Value'].dropna()
                    if not valid_values.empty:
                        avg_value = valid_values.mean()
                        event_stats[event_name] = avg_value
                        event_starts[event_name] = event_group['Offset'].min()
            
            # Add annotations
            # Handle potential missing Annotation column
            if 'Annotation' in combined_for_events.columns:
                non_empty_annotations = combined_for_events[combined_for_events['Annotation'].notna() & 
                                                           (combined_for_events['Annotation'].astype(str) != '')]
                if not non_empty_annotations.empty:
                    labeled_events = set()
                    
                    for _, row in non_empty_annotations.sort_values('Offset').iterrows():
                        event_name = row.get('Event', '') if pd.notna(row.get('Event', '')) else ''
                        x_pos = row['Offset']
                        base_y = row['Value'] + 0.75
                        
                        # Find available y position
                        adjusted_y = find_available_y_position(
                            x_pos,
                            base_y,
                            used_positions,
                            min_gap=1.0
                        )
                        
                        # Add average info for first occurrence
                        if event_name in event_stats and event_name not in labeled_events and \
                           row['Offset'] == event_starts.get(event_name, None):
                            avg_text = f"\nAvg: {event_stats[event_name]:.2f}"
                            labeled_events.add(event_name)
                        else:
                            avg_text = ""
                        
                        # Add annotation text
                        annotation_text = f"{row['Annotation']}{avg_text}"
                        
                        fig.add_trace(
                            go.Scatter(
                                name=f'{phase} (Actions)',
                                x=[x_pos],
                                y=[adjusted_y],
                                mode='text',
                                text=[annotation_text],
                                textposition='top center',
                                textfont=dict(size=12, color='black'),
                                showlegend=False,
                                hoverinfo='text'
                            ),
                            secondary_y=False
                        )
                        
                        used_positions.append((x_pos, adjusted_y))
                else:
                    print("No non-empty annotations found for action labels")
            else:
                print("No 'Annotation' column available for action labels")

    # Update layout
    fig.update_layout(
        uirevision=phase,
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

    # Combine all data for the table - make this more robust with error handling
    if all_data:
        try:
            # Safe concatenation with error handling
            dataframes_to_concat = []
            for df_idx, df in enumerate(all_data):
                if isinstance(df, pd.DataFrame) and not df.empty:
                    # Verify and fix each dataframe before concatenation
                    if 'Offset' not in df.columns or 'Value' not in df.columns:
                        print(f"Warning: Dataframe at index {df_idx} is missing required columns")
                        continue
                    
                    # Make a copy to avoid modifying the original
                    df_copy = df.copy()
                    
                    # Ensure consistent types for key columns
                    for col in ['Offset', 'Value']:
                        if col in df_copy.columns:
                            df_copy[col] = pd.to_numeric(df_copy[col], errors='coerce')
                    
                    # Ensure offset is in seconds
                    if 'Offset' in df_copy.columns and df_copy['Offset'].median() > 1000:
                        print(f"Converting Offset from milliseconds to seconds for dataframe {df_idx} in table")
                        df_copy['Offset'] = df_copy['Offset'] / 1000
                    
                    dataframes_to_concat.append(df_copy)
                else:
                    print(f"Warning: Item at index {df_idx} is not a valid DataFrame, skipping")
            
            if dataframes_to_concat:
                combined_data = pd.concat(dataframes_to_concat, ignore_index=True)
                
                # Ensure consistent sorting by Offset
                # First, make sure Offset is numeric
                combined_data['Offset'] = pd.to_numeric(combined_data['Offset'], errors='coerce')
                
                # Sort by Offset first, then by Type if it exists
                sort_cols = ['Offset']
                if 'Type' in combined_data.columns:
                    sort_cols.append('Type')
                
                combined_data = combined_data.sort_values(sort_cols)
                
                # Ensure 'selected' column exists and is set to True
                combined_data['selected'] = True
                
                # Print summary
                print(f"\nTable summary for {phase}:")
                print(f"Number of rows: {len(combined_data)}")
                print(f"Columns: {combined_data.columns.tolist()}")
                print(f"Offset range in table: {combined_data['Offset'].min():.2f}s - {combined_data['Offset'].max():.2f}s")
                
                # Print sample rows if available
                if not combined_data.empty:
                    print(f"Sample of first few rows:")
                    print(combined_data.head())
                
                return fig, combined_data
            else:
                print(f"No valid dataframes to combine for {phase}")
                # Create empty DataFrame with required columns
                empty_df = pd.DataFrame(columns=['sheet_name', 'Type', 'Annotation', 'Event', 'Offset', 'Value', 
                                            'mbase', 'mbase_time', 'mpeak', 'mpeak_time', 
                                            'pos_delta', 'neg_delta', 'VAS', 'Notes', 'selected'])
                empty_df['selected'] = True
                return fig, empty_df
                
        except Exception as e:
            print(f"Error combining data for {phase}: {str(e)}")
            traceback.print_exc()
            empty_df = pd.DataFrame(columns=['sheet_name', 'Type', 'Annotation', 'Event', 'Offset', 'Value', 
                                        'mbase', 'mbase_time', 'mpeak', 'mpeak_time', 
                                        'pos_delta', 'neg_delta', 'VAS', 'Notes', 'selected'])
            empty_df['selected'] = True
            return fig, empty_df
    else:
        empty_df = pd.DataFrame(columns=['sheet_name', 'Type', 'Annotation', 'Event', 'Offset', 'Value', 
                                    'mbase', 'mbase_time', 'mpeak', 'mpeak_time', 
                                    'pos_delta', 'neg_delta', 'VAS', 'Notes', 'selected'])
        empty_df['selected'] = True
        print(f"\nNo data found for {phase}, returning empty DataFrame")
        return fig, empty_df
    
@app.callback(
    [Output('p1-graph', 'figure'),
     Output('p1-table', 'data'),
     Output('p1-table', 'selected_rows')],
    [Input('p1-worksheet-selector', 'value'),
     Input('peak-data-file-store', 'data'),
     Input('window-size-input', 'value'),
     Input('action-toggle', 'n_clicks'),
     Input('include-words', 'value'),
     Input('exclude-words', 'value'),
     Input('include-events', 'value'),
     Input('exclude-events', 'value')],
    [State('sheet-data-store', 'data')],
    prevent_initial_call=True
)
def update_p1_analysis(selected_worksheets, peak_data_store, window_size, action_toggle, 
                       include_words, exclude_words, include_events, exclude_events,
                       sheet_data_store):
    """Update P1 analysis with enhanced peak data handling"""
    try:
        # Process filter inputs
        include_words_list = [w.strip() for w in include_words.split(',') if w.strip()] if include_words else []
        exclude_words_list = [w.strip() for w in exclude_words.split(',') if w.strip()] if exclude_words else []
        include_events_list = [e.strip() for e in include_events.split(',') if e.strip()] if include_events else []
        exclude_events_list = [e.strip() for e in exclude_events.split(',') if e.strip()] if exclude_events else []
        
        # Determine if action labels should be shown
        show_actions = (action_toggle % 2 == 1) if action_toggle else False
        
        # Debug information about peak data store
        if peak_data_store and peak_data_store.get('loaded', False):
            p1_data = peak_data_store.get('p1_data', [])
            print(f"Peak data store contains {len(p1_data)} P1 data rows")
        else:
            print("No peak data store loaded or empty")
        
        # Get sheet data from store, or load it if not available
        if sheet_data_store and 'sheet_names' in sheet_data_store:
            print(f"Sheet data store contains {len(sheet_data_store['sheet_names'])} sheets")
            sheet_data = {}
            
            # Reconstruct sheet_data from store
            for sheet_name in sheet_data_store['sheet_names']:
                data_key = f"{sheet_name}_data"
                actions_key = f"{sheet_name}_actions"
                
                if data_key in sheet_data_store:
                    sheet_df = pd.DataFrame(sheet_data_store[data_key])
                    actions = sheet_data_store.get(actions_key, [])
                    
                    sheet_data[sheet_name] = {
                        'data': sheet_df,
                        'actions': actions
                    }
        else:
            # Fallback to loading directly if not in store
            print("Loading sheet data directly (not from store)")
            sheet_data = load_all_phase_sheets()
        
        # Get manual data if available
        manual_data = pd.DataFrame()
        project_name = dash.callback_context.states.get('peak-detection-project-dropdown.value')
        if project_name:
            mnl_file_path = os.path.join(PROJECTS_DIRECTORY, project_name, 'Mnl_Analysis.xlsx')
            if os.path.exists(mnl_file_path):
                _, manual_data = read_manual_trace_data(mnl_file_path)
        
        # Get peak data from files if available
        peak_data_files = pd.DataFrame()
        if peak_data_store and peak_data_store.get('loaded', False):
            p1_data = peak_data_store.get('p1_data', [])
            if p1_data:
                peak_data_files = pd.DataFrame(p1_data)
                print(f"Loaded {len(peak_data_files)} rows of P1 data from peak data store")
                print(f"P1 data columns: {peak_data_files.columns.tolist()}")
            else:
                print("No P1 data found in peak data store")
        
        # Use the create_phase_graph_and_table function
        fig, combined_data = create_phase_graph_and_table(
            sheet_data=sheet_data,
            manual_data=manual_data,
            phase='P1',
            selected_sheets=selected_worksheets,
            include_words=include_words_list,
            exclude_words=exclude_words_list,
            include_events=include_events_list,
            exclude_events=exclude_events_list,
            show_actions=show_actions,
            window_size=window_size,
            peak_data_files=peak_data_files
        )
        
        # Convert to records for table
        table_data = combined_data.to_dict('records')
        print(f"P1 analysis returning {len(table_data)} rows for the table")
        
        # Return results
        return fig, table_data, list(range(min(len(table_data), 10)))
        
    except Exception as e:
        print(f"Error updating P1 analysis: {str(e)}")
        traceback.print_exc()
        
        # Create error figure
        error_fig = go.Figure()
        error_fig.update_layout(
            title='Error Loading P1 Data',
            annotations=[dict(
                text=f"Error: {str(e)}",
                xref="paper", yref="paper",
                x=0.5, y=0.5,
                showarrow=False
            )]
        )
        return error_fig, [], []

# Callback to handle P1 table select/deselect all buttons
@app.callback(
    [Output('p1-table', 'selected_rows', allow_duplicate=True),
     Output('p1-confirm-deselect', 'displayed')],
    [Input('p1-select-all', 'n_clicks'),
     Input('p1-deselect-all', 'n_clicks'),
     Input('p1-confirm-deselect', 'submit_n_clicks')],
    [State('p1-table', 'data')],
    prevent_initial_call=True
)
def manage_p1_selection(select_clicks, deselect_clicks, confirm_clicks, table_data):
    """Manage selection/deselection of P1 table rows"""
    ctx = callback_context
    if not ctx.triggered:
        return [], False
        
    trigger_id = ctx.triggered[0]['prop_id'].split('.')[0]
    
    if trigger_id == 'p1-select-all' and table_data:
        return list(range(len(table_data))), False
    elif trigger_id == 'p1-deselect-all':
        # Show confirmation dialog
        return dash.no_update, True
    elif trigger_id == 'p1-confirm-deselect' and confirm_clicks:
        # User confirmed deselection
        return [], False
        
    return dash.no_update, False

# # Callback to save P1 results
# @app.callback(
#     [Output('p1-save-status', 'children'),
#      # Output('p1-save-status', 'style')],
#     [Input('save-p1-button', 'n_clicks'),]
#     [State('p1-table', 'data'),
#      State('p1-filename-input', 'value'),
#      State('peak-detection-project-dropdown', 'value')],
#     prevent_initial_call=True
# )
# def save_p1_results(n_clicks, table_data, filename, project_name):
#     """Save P1 results to a file"""
#     if not n_clicks:
#         return "", {'color': 'black'}
        
#     if not table_data:
#         return "No data to save", {'color': 'red'}
        
#     if not project_name:
#         return "No project selected", {'color': 'red'}
        
#     if not filename:
#         filename = 'mnl_py_results'
        
#     try:
#         # Ensure the project and peak_data directories exist
#         peak_data_path = os.path.join(PROJECTS_DIRECTORY, project_name, 'peak_data')
#         if not os.path.exists(peak_data_path):
#             os.makedirs(peak_data_path)
            
#         # Convert table data to DataFrame
#         df = pd.DataFrame(table_data)
        
#         # Create output filename
#         output_filename = f"{filename}_p1.csv"
#         output_path = os.path.join(peak_data_path, output_filename)
        
#         # Save data
#         df.to_csv(output_path, index=False)
        
#         return f"Saved results to {output_filename}", {'color': 'green'}
        
#     except Exception as e:
#         print(f"Error saving P1 results: {str(e)}")
#         traceback.print_exc()
#         return f"Error saving results: {str(e)}", {'color': 'red'}
    
# Create P3 figure function # P3 1ST LAYOUT 
def create_p3_figure(df):
    """Create a figure for P3 analysis based on data"""
    if df.empty:
        fig = go.Figure()
        fig.update_layout(
            title='No P3 Data Available',
            annotations=[dict(
                text="No P3 data to display",
                xref="paper", yref="paper",
                x=0.5, y=0.5,
                showarrow=False
            )]
        )
        return fig
    
    fig = go.Figure()
    
    # Add line trace for all data
    fig.add_trace(
        go.Scatter(
            name='Raw Data',
            x=df['Offset'],
            y=df['Value'],
            mode='lines',
            line=dict(color='rgb(0, 140, 255)', width=2),
            showlegend=True
        )
    )
    
    # Add points grouped by Type if the column exists
    if 'Type' in df.columns:
        for point_type in df['Type'].unique():
            if pd.isna(point_type):
                continue
                
            type_data = df[df['Type'] == point_type]
            
            # Determine marker properties based on type
            if 'Rolling Base' in str(point_type):
                marker_color = 'rgb(102, 255, 51)'  # Light green for rolling bases
                marker_symbol = 'circle'
            elif 'Rolling Peak' in str(point_type):
                marker_color = 'rgb(0, 176, 240)'  # Light blue for rolling peaks
                marker_symbol = 'circle'
            elif 'Base' in str(point_type):
                marker_color = 'rgb(255, 165, 0)'  # Orange for regular bases
                marker_symbol = 'circle'
            elif 'Peak' in str(point_type):
                marker_color = 'rgb(255, 0, 0)'  # Red for regular peaks
                marker_symbol = 'circle'
            else:
                marker_color = 'rgb(128, 128, 128)'
                marker_symbol = 'circle'
            
            fig.add_trace(
                go.Scatter(
                    name=str(point_type),
                    x=type_data['Offset'],
                    y=type_data['Value'],
                    mode='markers',
                    marker=dict(
                        color=marker_color,
                        size=10,
                        symbol=marker_symbol,
                        line=dict(color='white', width=1)
                    ),
                    showlegend=True
                )
            )
    
    # Update layout
    fig.update_layout(
        title='P3 Analysis',
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
        plot_bgcolor='white',
        paper_bgcolor='white'
    )
    
    return fig

@app.callback(
    [Output('p3-graph', 'figure'),
     Output('p3-table', 'data'),
     Output('p3-table', 'selected_rows')],
    [Input('p3-worksheet-selector', 'value'),
     Input('peak-data-file-store', 'data'),
     Input('window-size-input', 'value'),
     Input('action-toggle', 'n_clicks'),
     Input('include-words', 'value'),
     Input('exclude-words', 'value'),
     Input('include-events', 'value'),
     Input('exclude-events', 'value')],
    [State('sheet-data-store', 'data')],
    prevent_initial_call=True
)
def update_p3_analysis(selected_worksheets, peak_data_store, window_size, action_toggle, 
                       include_words, exclude_words, include_events, exclude_events,
                       sheet_data_store):
    """Update P3 analysis with enhanced peak data handling"""
    try:
        # Process filter inputs
        include_words_list = [w.strip() for w in include_words.split(',') if w.strip()] if include_words else []
        exclude_words_list = [w.strip() for w in exclude_words.split(',') if w.strip()] if exclude_words else []
        include_events_list = [e.strip() for e in include_events.split(',') if e.strip()] if include_events else []
        exclude_events_list = [e.strip() for e in exclude_events.split(',') if e.strip()] if exclude_events else []
        
        # Determine if action labels should be shown
        show_actions = (action_toggle % 2 == 1) if action_toggle else False
        
        # Debug information about peak data store
        if peak_data_store and peak_data_store.get('loaded', False):
            p3_data = peak_data_store.get('p3_data', [])
            print(f"Peak data store contains {len(p3_data)} P3 data rows")
        else:
            print("No peak data store loaded or empty")
        
        # Get sheet data from store, or load it if not available
        if sheet_data_store and 'sheet_names' in sheet_data_store:
            print(f"Sheet data store contains {len(sheet_data_store['sheet_names'])} sheets")
            sheet_data = {}
            
            # Reconstruct sheet_data from store
            for sheet_name in sheet_data_store['sheet_names']:
                data_key = f"{sheet_name}_data"
                actions_key = f"{sheet_name}_actions"
                
                if data_key in sheet_data_store:
                    sheet_df = pd.DataFrame(sheet_data_store[data_key])
                    actions = sheet_data_store.get(actions_key, [])
                    
                    sheet_data[sheet_name] = {
                        'data': sheet_df,
                        'actions': actions
                    }
        else:
            # Fallback to loading directly if not in store
            print("Loading sheet data directly (not from store)")
            sheet_data = load_all_phase_sheets()
        
        # Get manual data if available
        manual_data = pd.DataFrame()
        project_name = dash.callback_context.states.get('peak-detection-project-dropdown.value')
        if project_name:
            mnl_file_path = os.path.join(PROJECTS_DIRECTORY, project_name, 'Mnl_Analysis.xlsx')
            if os.path.exists(mnl_file_path):
                _, manual_data = read_manual_trace_data(mnl_file_path)
        
        # Get peak data from files if available
        peak_data_files = pd.DataFrame()
        if peak_data_store and peak_data_store.get('loaded', False):
            p3_data = peak_data_store.get('p3_data', [])
            if p3_data:
                peak_data_files = pd.DataFrame(p3_data)
                print(f"Loaded {len(peak_data_files)} rows of P3 data from peak data store")
                print(f"P3 data columns: {peak_data_files.columns.tolist()}")
            else:
                print("No P3 data found in peak data store")
        
        # Use the create_phase_graph_and_table function
        fig, combined_data = create_phase_graph_and_table(
            sheet_data=sheet_data,
            manual_data=manual_data,
            phase='P3',
            selected_sheets=selected_worksheets,
            include_words=include_words_list,
            exclude_words=exclude_words_list,
            include_events=include_events_list,
            exclude_events=exclude_events_list,
            show_actions=show_actions,
            window_size=window_size,
            peak_data_files=peak_data_files
        )
        
        # Convert to records for table
        table_data = combined_data.to_dict('records')
        print(f"P3 analysis returning {len(table_data)} rows for the table")
        
        # Return results
        return fig, table_data, list(range(min(len(table_data), 10)))
        
    except Exception as e:
        print(f"Error updating P3 analysis: {str(e)}")
        traceback.print_exc()
        
        # Create error figure
        error_fig = go.Figure()
        error_fig.update_layout(
            title='Error Loading P3 Data',
            annotations=[dict(
                text=f"Error: {str(e)}",
                xref="paper", yref="paper",
                x=0.5, y=0.5,
                showarrow=False
            )]
        )
        return error_fig, [], []

# Callback to handle P3 table select/deselect all buttons
@app.callback(
    [Output('p3-table', 'selected_rows', allow_duplicate=True),
     Output('p3-confirm-deselect', 'displayed')],
    [Input('p3-select-all', 'n_clicks'),
     Input('p3-deselect-all', 'n_clicks'),
     Input('p3-confirm-deselect', 'submit_n_clicks')],
    [State('p3-table', 'data')],
    prevent_initial_call=True
)

def manage_p3_selection(select_clicks, deselect_clicks, confirm_clicks, table_data):
    """Manage selection/deselection of P3 table rows"""
    ctx = callback_context
    if not ctx.triggered:
        return [], False
        
    trigger_id = ctx.triggered[0]['prop_id'].split('.')[0]
    
    if trigger_id == 'p3-select-all' and table_data:
        return list(range(len(table_data))), False
    elif trigger_id == 'p3-deselect-all':
        # Show confirmation dialog
        return dash.no_update, True
    elif trigger_id == 'p3-confirm-deselect' and confirm_clicks:
        # User confirmed deselection
        return [], False
        
    return dash.no_update, False

# Callback to save P3 results
# @app.callback(
#     [Output('p3-save-status', 'children'),
#      Output('p3-save-status', 'style')],
#     Input('save-p3-button', 'n_clicks'),
#     [State('p3-table', 'data'),
#      State('p3-filename-input', 'value'),
#      State('peak-detection-project-dropdown', 'value')],
#     prevent_initial_call=True
# )
# def save_p3_results(n_clicks, table_data, filename, project_name):
#     """Save P3 results to a file"""
#     if not n_clicks:
#         return "", {'color': 'black'}
        
#     if not table_data:
#         return "No data to save", {'color': 'red'}
        
#     if not project_name:
#         return "No project selected", {'color': 'red'}
        
#     if not filename:
#         filename = 'mnl_py_results'
        
#     try:
#         # Ensure the project and peak_data directories exist
#         peak_data_path = os.path.join(PROJECTS_DIRECTORY, project_name, 'peak_data')
#         if not os.path.exists(peak_data_path):
#             os.makedirs(peak_data_path)
            
#         # Convert table data to DataFrame
#         df = pd.DataFrame(table_data)
        
#         # Create output filename
#         output_filename = f"{filename}_p3.csv"
#         output_path = os.path.join(peak_data_path, output_filename)
        
#         # Save data
#         df.to_csv(output_path, index=False)
        
#         return f"Saved results to {output_filename}", {'color': 'green'}
        
#     except Exception as e:
#         print(f"Error saving P3 results: {str(e)}")
#         traceback.print_exc()
#         return f"Error saving results: {str(e)}", {'color': 'red'}
    
# Create P4 figure function # P4 1ST LAYOUT 
def create_p4_figure(df):
    """Create a figure for P3 analysis based on data"""
    if df.empty:
        fig = go.Figure()
        fig.update_layout(
            title='No P4 Data Available',
            annotations=[dict(
                text="No P4 data to display",
                xref="paper", yref="paper",
                x=0.5, y=0.5,
                showarrow=False
            )]
        )
        return fig
    
    fig = go.Figure()
    
    # Add line trace for all data
    fig.add_trace(
        go.Scatter(
            name='Raw Data',
            x=df['Offset'],
            y=df['Value'],
            mode='lines',
            line=dict(color='rgb(0, 140, 255)', width=2),
            showlegend=True
        )
    )
    
    # Add points grouped by Type if the column exists
    if 'Type' in df.columns:
        for point_type in df['Type'].unique():
            if pd.isna(point_type):
                continue
                
            type_data = df[df['Type'] == point_type]
            
            # Determine marker properties based on type
            if 'Rolling Base' in str(point_type):
                marker_color = 'rgb(102, 255, 51)'  # Light green for rolling bases
                marker_symbol = 'circle'
            elif 'Rolling Peak' in str(point_type):
                marker_color = 'rgb(0, 176, 240)'  # Light blue for rolling peaks
                marker_symbol = 'circle'
            elif 'Base' in str(point_type):
                marker_color = 'rgb(255, 165, 0)'  # Orange for regular bases
                marker_symbol = 'circle'
            elif 'Peak' in str(point_type):
                marker_color = 'rgb(255, 0, 0)'  # Red for regular peaks
                marker_symbol = 'circle'
            else:
                marker_color = 'rgb(128, 128, 128)'
                marker_symbol = 'circle'
            
            fig.add_trace(
                go.Scatter(
                    name=str(point_type),
                    x=type_data['Offset'],
                    y=type_data['Value'],
                    mode='markers',
                    marker=dict(
                        color=marker_color,
                        size=10,
                        symbol=marker_symbol,
                        line=dict(color='white', width=1)
                    ),
                    showlegend=True
                )
            )
    
    # Update layout
    fig.update_layout(
        title='P4 Analysis',
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
        plot_bgcolor='white',
        paper_bgcolor='white'
    )
    
    return fig

# @app.callback(
#     [Output('p4-graph', 'figure'),
#      Output('p4-table', 'data'),
#      Output('p4-table', 'selected_rows')],
#     [Input('p4-worksheet-selector', 'value'),
#      Input('peak-data-file-store', 'data')],
#     prevent_initial_call=True
# )
# def update_p4_analysis(selected_worksheets, peak_data_store):
#     """Update P4 analysis based on selected worksheets, manual analysis, and peak data files"""
#     try:
#         # Initialize combined dataframe
#         combined_df = pd.DataFrame()
        
#         # 1. First, load selected worksheet data from All_Phase_Sheets.xlsx
#         if selected_worksheets:
#             worksheet_df = pd.DataFrame()
            
#             # Convert to list if it's a single value
#             if not isinstance(selected_worksheets, list):
#                 selected_worksheets = [selected_worksheets]
                
#             # Load All_Phase_Sheets
#             sheet_data = load_all_phase_sheets()
            
#             for sheet_name in selected_worksheets:
#                 if sheet_name in sheet_data:
#                     sheet_df = sheet_data[sheet_name]['data']
                    
#                     # Add sheet_name column
#                     sheet_df['sheet_name'] = sheet_name
                    
#                     # Add to worksheet dataframe
#                     if worksheet_df.empty:
#                         worksheet_df = sheet_df
#                     else:
#                         worksheet_df = pd.concat([worksheet_df, sheet_df], ignore_index=True)
            
#             if not worksheet_df.empty:
#                 # Set basic columns
#                 worksheet_df['Type'] = 'Worksheet Data'
#                 worksheet_df['selected'] = True
                
#                 # Merge with combined dataframe
#                 combined_df = worksheet_df
#                 print(f"Added {len(worksheet_df)} rows from selected worksheets")
        
#         # 2. Second, load and add manual analysis data (Mnl_Analysis)
#         try:
#             # Assuming Mnl_Analysis.xlsx is in a known location
#             project_name = dash.callback_context.states.get('peak-detection-project-dropdown.value')
#             if project_name:
#                 mnl_file_path = os.path.join(PROJECTS_DIRECTORY, project_name, 'Mnl_Analysis.xlsx')
                
#                 if os.path.exists(mnl_file_path):
#                     _, manual_df = read_manual_trace_data(mnl_file_path)
                    
#                     # Filter for P4 data only
#                     if 'sheet_name' in manual_df.columns:
#                         manual_p4_df = manual_df[manual_df['sheet_name'].str.contains('P4', case=False, na=False)]
                        
#                         if not manual_p4_df.empty:
#                             # Merge with combined data
#                             if combined_df.empty:
#                                 combined_df = manual_p4_df
#                             else:
#                                 combined_df = pd.concat([combined_df, manual_p4_df], ignore_index=True)
                            
#                             print(f"Added {len(manual_p4_df)} rows from Mnl_Analysis")
#         except Exception as e:
#             print(f"Error loading manual analysis data: {str(e)}")
                    
#         # 3. Finally, add peak data files if available
#         if peak_data_store and peak_data_store.get('loaded', False):
#             p4_data = peak_data_store.get('p4_data', [])
            
#             if p4_data:
#                 peak_df = pd.DataFrame(p4_data)
                
#                 if not peak_df.empty:
#                     # Merge with combined data
#                     if combined_df.empty:
#                         combined_df = peak_df
#                     else:
#                         combined_df = pd.concat([combined_df, peak_df], ignore_index=True)
                    
#                     print(f"Added {len(peak_df)} rows from peak data files")
        
#         # If we still have no data, return empty view
#         if combined_df.empty:
#             empty_fig = go.Figure()
#             empty_fig.update_layout(
#                 title='No P4 Data Available',
#                 annotations=[dict(
#                     text="No data available. Please select worksheets or load peak data files.",
#                     xref="paper", yref="paper",
#                     x=0.5, y=0.5,
#                     showarrow=False
#                 )]
#             )
#             return empty_fig, [], []
        
#         # Create figure based on data
#         fig = create_p4_figure(combined_df)
        
#         # Convert to records for table
#         table_data = combined_df.to_dict('records')
        
#         # Return results
#         return fig, table_data, list(range(min(len(table_data), 10)))  # Select first 10 rows by default
        
#     except Exception as e:
#         print(f"Error updating P4 analysis: {str(e)}")
#         traceback.print_exc()
        
#         # Create error figure
#         error_fig = go.Figure()
#         error_fig.update_layout(
#             title='Error Loading P4 Data',
#             annotations=[dict(
#                 text=f"Error: {str(e)}",
#                 xref="paper", yref="paper",
#                 x=0.5, y=0.5,
#                 showarrow=False
#             )]
#         )
#         return error_fig, [], []

@app.callback(
    [Output('p4-graph', 'figure'),
     Output('p4-table', 'data'),
     Output('p4-table', 'selected_rows')],
    [Input('p4-worksheet-selector', 'value'),
     Input('peak-data-file-store', 'data'),
     Input('window-size-input', 'value'),
     Input('action-toggle', 'n_clicks'),
     Input('include-words', 'value'),
     Input('exclude-words', 'value'),
     Input('include-events', 'value'),
     Input('exclude-events', 'value')],
    [State('sheet-data-store', 'data')],
    prevent_initial_call=True
)
def update_p4_analysis(selected_worksheets, peak_data_store, window_size, action_toggle, 
                       include_words, exclude_words, include_events, exclude_events,
                       sheet_data_store):
    """Update P4 analysis with enhanced peak data handling"""
    try:
        # Process filter inputs
        include_words_list = [w.strip() for w in include_words.split(',') if w.strip()] if include_words else []
        exclude_words_list = [w.strip() for w in exclude_words.split(',') if w.strip()] if exclude_words else []
        include_events_list = [e.strip() for e in include_events.split(',') if e.strip()] if include_events else []
        exclude_events_list = [e.strip() for e in exclude_events.split(',') if e.strip()] if exclude_events else []
        
        # Determine if action labels should be shown
        show_actions = (action_toggle % 2 == 1) if action_toggle else False
        
        # Debug information about peak data store
        if peak_data_store and peak_data_store.get('loaded', False):
            p4_data = peak_data_store.get('p4_data', [])
            print(f"Peak data store contains {len(p4_data)} P4 data rows")
        else:
            print("No peak data store loaded or empty")
        
        # Get sheet data from store, or load it if not available
        if sheet_data_store and 'sheet_names' in sheet_data_store:
            print(f"Sheet data store contains {len(sheet_data_store['sheet_names'])} sheets")
            sheet_data = {}
            
            # Reconstruct sheet_data from store
            for sheet_name in sheet_data_store['sheet_names']:
                data_key = f"{sheet_name}_data"
                actions_key = f"{sheet_name}_actions"
                
                if data_key in sheet_data_store:
                    sheet_df = pd.DataFrame(sheet_data_store[data_key])
                    actions = sheet_data_store.get(actions_key, [])
                    
                    sheet_data[sheet_name] = {
                        'data': sheet_df,
                        'actions': actions
                    }
        else:
            # Fallback to loading directly if not in store
            print("Loading sheet data directly (not from store)")
            sheet_data = load_all_phase_sheets()
        
        # Get manual data if available
        manual_data = pd.DataFrame()
        project_name = dash.callback_context.states.get('peak-detection-project-dropdown.value')
        if project_name:
            mnl_file_path = os.path.join(PROJECTS_DIRECTORY, project_name, 'Mnl_Analysis.xlsx')
            if os.path.exists(mnl_file_path):
                _, manual_data = read_manual_trace_data(mnl_file_path)
        
        # Get peak data from files if available
        peak_data_files = pd.DataFrame()
        if peak_data_store and peak_data_store.get('loaded', False):
            p4_data = peak_data_store.get('p4_data', [])
            if p4_data:
                peak_data_files = pd.DataFrame(p4_data)
                print(f"Loaded {len(peak_data_files)} rows of P1 data from peak data store")
                print(f"P1 data columns: {peak_data_files.columns.tolist()}")
            else:
                print("No P4 data found in peak data store")
        
        # Use the create_phase_graph_and_table function
        fig, combined_data = create_phase_graph_and_table(
            sheet_data=sheet_data,
            manual_data=manual_data,
            phase='P4',
            selected_sheets=selected_worksheets,
            include_words=include_words_list,
            exclude_words=exclude_words_list,
            include_events=include_events_list,
            exclude_events=exclude_events_list,
            show_actions=show_actions,
            window_size=window_size,
            peak_data_files=peak_data_files
        )
        
        # Convert to records for table
        table_data = combined_data.to_dict('records')
        print(f"P4 analysis returning {len(table_data)} rows for the table")
        
        # Return results
        return fig, table_data, list(range(min(len(table_data), 10)))
        
    except Exception as e:
        print(f"Error updating P4 analysis: {str(e)}")
        traceback.print_exc()
        
        # Create error figure
        error_fig = go.Figure()
        error_fig.update_layout(
            title='Error Loading P4 Data',
            annotations=[dict(
                text=f"Error: {str(e)}",
                xref="paper", yref="paper",
                x=0.5, y=0.5,
                showarrow=False
            )]
        )
        return error_fig, [], []

# Callback to handle P4 table select/deselect all buttons
@app.callback(
    [Output('p4-table', 'selected_rows', allow_duplicate=True),
     Output('p4-confirm-deselect', 'displayed')],
    [Input('p4-select-all', 'n_clicks'),
     Input('p4-deselect-all', 'n_clicks'),
     Input('p4-confirm-deselect', 'submit_n_clicks')],
    [State('p4-table', 'data')],
    prevent_initial_call=True
)
def manage_p4_selection(select_clicks, deselect_clicks, confirm_clicks, table_data):
    """Manage selection/deselection of P4 table rows"""
    ctx = callback_context
    if not ctx.triggered:
        return [], False
        
    trigger_id = ctx.triggered[0]['prop_id'].split('.')[0]
    
    if trigger_id == 'p4-select-all' and table_data:
        return list(range(len(table_data))), False
    elif trigger_id == 'p4-deselect-all':
        # Show confirmation dialog
        return dash.no_update, True
    elif trigger_id == 'p4-confirm-deselect' and confirm_clicks:
        # User confirmed deselection
        return [], False
        
    return dash.no_update, False

# # Callback to save P4 results
# @app.callback(
#     [Output('p4-save-status', 'children'),
#      Output('p4-save-status', 'style')],
#     Input('save-p4-button', 'n_clicks'),
#     [State('p4-table', 'data'),
#      State('p4-filename-input', 'value'),
#      State('peak-detection-project-dropdown', 'value')],
#     prevent_initial_call=True
# )
# def save_p4_results(n_clicks, table_data, filename, project_name):
#     """Save P4 results to a file"""
#     if not n_clicks:
#         return "", {'color': 'black'}
        
#     if not table_data:
#         return "No data to save", {'color': 'red'}
        
#     if not project_name:
#         return "No project selected", {'color': 'red'}
        
#     if not filename:
#         filename = 'mnl_py_results'
        
#     try:
#         # Ensure the project and peak_data directories exist
#         peak_data_path = os.path.join(PROJECTS_DIRECTORY, project_name, 'peak_data')
#         if not os.path.exists(peak_data_path):
#             os.makedirs(peak_data_path)
            
#         # Convert table data to DataFrame
#         df = pd.DataFrame(table_data)
        
#         # Create output filename
#         output_filename = f"{filename}_p4.csv"
#         output_path = os.path.join(peak_data_path, output_filename)
        
#         # Save data
#         df.to_csv(output_path, index=False)
        
#         return f"Saved results to {output_filename}", {'color': 'green'}
        
#     except Exception as e:
#         print(f"Error saving P4 results: {str(e)}")
#         traceback.print_exc()
#         return f"Error saving results: {str(e)}", {'color': 'red'}    

# END PEAK DETECTION TAB 3 #


# In[7]:


#### PEAK DETECTION TAB 4 ####
##### PEAK DETECTION TAB AND RESULTS TAB ##### 

species_scale = 10  # Using a default value of 10

def detect_peaks_rolling_window(data, sheet_name, window_size):
    """
    Detect peaks and valleys using a rolling window comparison method with edge case handling.
    """
    species_scale = 10
    peaks_data = []
    n = len(data)

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
                            'Value': float(current_value*species_scale),
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
            df = calculate_auto_peak_deltas(df)
        
        return df
    
    except Exception as e:
        print(f"Error processing sheet {sheet_name}: {str(e)}")
        traceback.print_exc()
        return pd.DataFrame()

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
    new_table = calculate_auto_peak_deltas(new_table)
            
    return new_table
    
# def analyze_multiple_graphs_statistics(sheet_data, selected_sheets, visible_range=None):
#     """
#     Calculate statistics for multiple graphs, both combined and individual.
    
#     Parameters:
#     -----------
#     sheet_data : dict
#         Dictionary containing data for each sheet
#     selected_sheets : list
#         List of selected sheet names to analyze
#     visible_range : tuple, optional
#         (min_x, max_x) range to calculate statistics for
        
#     Returns:
#     --------
#     tuple
#         (combined_stats, individual_stats)
#     """
#     try:
#         all_values = []
#         individual_stats = {}
        
#         # Process each selected sheet
#         for sheet_name in selected_sheets:
#             if sheet_name not in sheet_data:
#                 continue
                
#             data = sheet_data[sheet_name]['data']
#             if data.empty:
#                 continue
            
#             # Apply visible range filter if provided
#             if visible_range:
#                 min_x, max_x = visible_range
#                 mask = (data['Offset']/1000 >= min_x) & (data['Offset']/1000 <= max_x)
#                 data = data[mask]
            
#             values = data['Value'].values
            
#             if len(values) > 0:
#                 # Calculate individual sheet statistics
#                 stats = {
#                     'mean': np.mean(values),
#                     'std': np.std(values),
#                     'max': np.max(values),
#                     'min': np.min(values),
#                     'range': np.max(values) - np.min(values),
#                     'median': np.median(values),
#                     'q1': np.percentile(values, 25),
#                     'q3': np.percentile(values, 75),
#                     'iqr': np.percentile(values, 75) - np.percentile(values, 25)
#                 }
#                 individual_stats[sheet_name] = stats
#                 all_values.extend(values)
        
#         # Calculate combined statistics if we have any values
#         if all_values:
#             combined_stats = {
#                 'mean': np.mean(all_values),
#                 'std': np.std(all_values),
#                 'max': np.max(all_values),
#                 'min': np.min(all_values),
#                 'range': np.max(all_values) - np.min(all_values),
#                 'median': np.median(all_values),
#                 'q1': np.percentile(all_values, 25),
#                 'q3': np.percentile(all_values, 75),
#                 'iqr': np.percentile(all_values, 75) - np.percentile(all_values, 25)
#             }
#         else:
#             combined_stats = None
            
#         return combined_stats, individual_stats
        
#     except Exception as e:
#         print(f"Error in analyze_multiple_graphs_statistics: {str(e)}")
#         return None, {}
    

def create_multi_p1_stats_table(combined_stats, individual_stats):
    """Create a formatted DataTable to display statistics for multiple graphs"""
    if not combined_stats and not individual_stats:
        return dash_table.DataTable(
            id='p1-stats-table',
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
        id='p1-stats-table',
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

def create_multi_p3_stats_table(combined_stats, individual_stats):
    """Create a formatted DataTable to display statistics for multiple graphs"""
    if not combined_stats and not individual_stats:
        return dash_table.DataTable(
            id='p3-stats-table',
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
        id='p3-stats-table',
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

def create_multi_p4_stats_table(combined_stats, individual_stats):
    """Create a formatted DataTable to display statistics for multiple graphs"""
    if not combined_stats and not individual_stats:
        return dash_table.DataTable(
            id='p4-stats-table',
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
        id='p4-stats-table',
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



# if __name__ == '__main__':
#     print("Starting application...")
#     print(f"Projects directory: {PROJECTS_DIRECTORY}")
    
#     # Make sure directories exist
#     if not os.path.exists(PROJECTS_DIRECTORY):
#         print(f"Creating projects directory: {PROJECTS_DIRECTORY}")
#         os.makedirs(PROJECTS_DIRECTORY)
    
#     # Initialize app
#     initialize_app()
    
#     # Run server
#     app.run_server(debug=True, port=8050)
    ### End Peak detection Tab 4 ###


# In[8]:


### Peak Detection Tab 5 ###

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
            
# Load data 5TH UPLOAD
try:
    # First read measurement data to get valid IDs
    valid_ids, manual_trace_data = read_manual_trace_data('Mnl_Analysis.xlsx')

    # Then read the phase data
    sheet_data = read_excel_sheets('All_Phase_Sheets.xlsx')

    # Add validation message
    print("\nData Validation:")
    if not manual_trace_data.empty and sheet_data:
        print("Successfully loaded measurement and phase data")
    else:
        print("Warning: Some data may be missing or empty")

except FileNotFoundError as e:
    print(f"Error loading files: {str(e)}")
    sheet_data = {}
    manual_trace_data = pd.DataFrame()
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

# Define app layout
app.layout = html.Div([
    dcc.Store(id='project-storage', storage_type='session'),
    dcc.Store(id='last-saved-wip-file', storage_type='memory'),
    dcc.Store(id='peak-data-file-store', storage_type='memory'),
    dcc.Store(id='sheet-data-store', storage_type='memory'),
    dcc.Store(id='p1-sheet-data-store', storage_type='memory'),
    dcc.Store(id='p3-sheet-data-store', storage_type='memory'),
    dcc.Store(id='p4-sheet-data-store', storage_type='memory'),
    dcc.Tabs(id='tabs', value='load-data', children=[
        load_data_tab,
        event_preparation_tab,
        
        dcc.Tab(label='Peak Detection', children=[
            html.H1('Peak Detection Analysis Dashboard'),
                    
            # Project and File Selection Controls
            html.Div([
                html.Div([
                    html.Label('Select Project:'),
                    dcc.Dropdown(
                        id='peak-detection-project-dropdown',
                        options=[],  # Will be populated by callback
                        placeholder="Select a project",
                        value=None,
                        style={'width': '300px'}
                    ),
                ], style={'marginRight': '20px'}),
                
                # Load Project Button
                html.Button(
                    'Load Project',
                    id='load-peak-detection-project-button',
                    n_clicks=0,
                    style={
                        'marginRight': '10px',
                        'padding': '5px 10px',
                        'backgroundColor': '#f8f9fa',
                        'border': '1px solid #ddd',
                        'borderRadius': '4px',
                        'cursor': 'pointer',
                        'height': '36px'
                    }
                ),
                html.Div(id='load-peak-detection-project-status', style={'marginLeft': '10px'}),
            ], style={'display': 'flex', 'alignItems': 'flex-end', 'marginBottom': '15px'}),
            
            # Peak Data Files Selection
            html.Div([
                html.Div([
                    html.Label('Select Peak Data Files:'),
                    dcc.Dropdown(
                        id='peak-detection-peak-data-files-dropdown',
                        options=[],  # Will be populated by callback
                        placeholder="Select peak analysis files",
                        multi=True,
                        value=[],
                        style={'width': '400px'}
                    ),
                ]),    

                html.Div([
                    html.Button(
                        'Load Selected Files',
                        id='load-peak-data-files-button',
                        n_clicks=0,
                        style={
                            'marginLeft': '10px',
                            'padding': '5px 10px',
                            'backgroundColor': '#f8f9fa',
                            'border': '1px solid #ddd',
                            'borderRadius': '4px',
                            'cursor': 'pointer',
                            'height': '36px'
                        }
                    ),
                    html.Div(id='load-peak-data-files-status', style={'marginLeft': '10px'})
                ], style={'display': 'flex', 'alignItems': 'center', 'marginLeft': '20px'})
            ], style={'display': 'flex', 'alignItems': 'flex-end', 'marginBottom': '20px'}),
            
                    html.Button(
                        'Refresh Files',
                        id='peak-detection-refresh-files-button',
                        className='action-button',
                        style={'marginLeft': '10px'}
                        ),

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
                            value=1,
                            min=0,
                            max=20,
                            step=0.5,
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
                    # html.Div(id='p1-save-status')
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
                    # html.Div(id='p3-save-status')
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
                    # html.Div(id='p4-save-status')
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

                html.Div([
                    html.H3('Peak Deltas Analysis', style={'marginTop': '20px'}),
                    dcc.Graph(id='overlay-indiv-neg-delta-graph'),
                    dcc.Graph(id='overlay-indiv-pos-delta-graph'),
                    dcc.Graph(id='overlay-neg-delta-graph'),
                    dcc.Graph(id='overlay-pos-delta-graph')
                ]),

                # New cluster graph
                html.H3("Negative Delta Cluster Analysis", style={'textAlign': 'center', 'marginTop': '20px'}),
                html.Div([
                    dcc.Graph(
                        id='overlay-neg-delta-cluster-graph',
                        figure=go.Figure(),
                        style={'height': '500px'}
                    )
                ]),

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
                                #html.H3('P1 Overall Statistics'),
                                #dcc.Graph(id='p1-overall-stats-graph'),
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
    ]),

    html.Div(id='initialization-trigger', style={'display': 'none'})
])

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

### End Peak Detection Tab 5 ###


# In[9]:


# ### Peak Detection Tab 6 ###

# # data validation and error checking
# def filter_selected_points(events_data, table_data):
#     """Filter event statistics to only include events with selected points"""
#     if not table_data:
#         return events_data
#     selected_points = pd.DataFrame(table_data)
#     if 'selected' not in selected_points.columns:
#         return events_data
#     # Only include events that have selected points
#     selected_events = selected_points[selected_points['selected']]['Event'].unique()
#     return [event for event in events_data if event['Event'] in selected_events]

# def create_default_outputs():
#     """Create default outputs for error cases"""
#     empty_fig = go.Figure()
#     empty_data = []
#     return (
#         empty_fig,  # p1_graph
#         empty_data,  # p1_table
#         [],  # p1_selected_rows
#         empty_data,  # p1_stats
#         empty_fig,  # p3_graph
#         empty_data,  # p3_table
#         [],  # p3_selected_rows
#         empty_data,  # p3_stats
#         empty_fig,  # p4_graph
#         empty_data,  # p4_table
#         [],  # p4_selected_rows
#         empty_data,  # p4_stats
#         "Action Labels OFF",  # action_toggle_status
#         False,  # p1_confirm
#         False,  # p3_confirm
#         False   # p4_confirm
#     )

# @app.callback(
#     [Output('p1-graph', 'figure'),
#      Output('p1-table', 'data'),
#      Output('p1-table', 'selected_rows'),
#      Output('p1-stats-table', 'data'),
#      Output('p3-graph', 'figure'),
#      Output('p3-table', 'data'),
#      Output('p3-table', 'selected_rows'),
#      Output('p3-stats-table', 'data'),
#      Output('p4-graph', 'figure'),
#      Output('p4-table', 'data'),
#      Output('p4-table', 'selected_rows'),
#      Output('p4-stats-table', 'data'),
#      Output('action-toggle-status', 'children'),
#      Output('p1-confirm-deselect', 'displayed'),
#      Output('p3-confirm-deselect', 'displayed'),     
#      Output('p4-confirm-deselect', 'displayed')],
#     [Input('include-words', 'value'),
#      Input('exclude-words', 'value'),
#      Input('include-events', 'value'),
#      Input('exclude-events', 'value'),
#      Input('p1-worksheet-selector', 'value'),
#      Input('p3-worksheet-selector', 'value'),
#      Input('p4-worksheet-selector', 'value'),
#      Input('action-toggle', 'n_clicks'),
#      Input('window-size-input', 'value'),
#      Input('p1-graph', 'relayoutData'),
#      Input('p3-graph', 'relayoutData'),
#      Input('p4-graph', 'relayoutData'),
#      Input('p1-table', 'selected_rows'),
#      Input('p3-table', 'selected_rows'),
#      Input('p4-table', 'selected_rows'),
#      Input('p1-table', 'data'),
#      Input('p3-table', 'data'),
#      Input('p4-table', 'data'),
#      Input('p1-select-all', 'n_clicks'),
#      Input('p1-deselect-all', 'n_clicks'),
#      Input('p3-select-all', 'n_clicks'),
#      Input('p3-deselect-all', 'n_clicks'),
#      Input('p4-select-all', 'n_clicks'),
#      Input('p4-deselect-all', 'n_clicks'),
#      Input('p1-confirm-deselect', 'submit_n_clicks'),
#      Input('p3-confirm-deselect', 'submit_n_clicks'),
#      Input('p4-confirm-deselect', 'submit_n_clicks')],
#     [State('p1-table', 'data'),
#      State('p3-table', 'data'),
#      State('p4-table', 'data')],
#     prevent_initial_call=True
# )

# def update_graphs_and_tables(include_words, exclude_words, include_events, exclude_events, p1_sheets, p3_sheets, p4_sheets, n_clicks, window_size,
#                            p1_relayout, p3_relayout, p4_relayout, 
#                            p1_selected_rows, p3_selected_rows, p4_selected_rows,
#                            p1_table_data, p3_table_data, p4_table_data,
#                            p1_select_all, p1_deselect_all,
#                            p3_select_all, p3_deselect_all,
#                            p4_select_all, p4_deselect_all,
#                            p1_confirm, p3_confirm, p4_confirm,
#                            p1_table_state, p3_table_state, p4_table_state):
#     """Update all graphs and tables based on user inputs and parameters"""
#     ctx = dash.callback_context
#     if not ctx.triggered:
#         raise PreventUpdate
        
#     trigger_id = ctx.triggered[0]['prop_id'].split('.')[0]

#     try:
#         # Initialize action toggle state based on n_clicks
#         show_actions = bool(n_clicks % 2)  # Changed this line
        
#         # Create status displays
#         toggle_status = html.Div(
#             "Action Labels ON" if show_actions else "Action Labels OFF",
#             style={'color': 'green' if show_actions else 'red'}
#         )
        
#         # Initialize dictionaries
#         results = {}
#         stats_data = {}
        
#         # Get all selected sheets
#         all_selected_sheets = []
#         if p1_sheets:
#             all_selected_sheets.extend(p1_sheets)
#         if p3_sheets:
#             all_selected_sheets.extend(p3_sheets)
#         if p4_sheets:
#             all_selected_sheets.extend(p4_sheets)

#         if not all_selected_sheets:
#             raise PreventUpdate
              
#         # Process filters
#         include_list = [word.strip() for word in include_words.split(',')] if include_words else DEFAULT_INCLUDE_WORDS
#         exclude_list = [word.strip() for word in exclude_words.split(',')] if exclude_words else DEFAULT_EXCLUDE_WORDS
#         include_event = [word.strip() for word in include_events.split(',')] if include_events else DEFAULT_INCLUDE_EVENTS
#         exclude_event = [word.strip() for word in exclude_events.split(',')] if exclude_events else DEFAULT_EXCLUDE_EVENTS

#         # Initialize confirmation flags
#         show_p1_confirm = False
#         show_p3_confirm = False
#         show_p4_confirm = False
        
#         # Handle select/deselect button clicks
#         ctx = dash.callback_context
#         if ctx.triggered:
#             button_id = ctx.triggered[0]['prop_id'].split('.')[0]
            
#             # Handle confirmation dialog submissions
#             if button_id == 'p1-confirm-deselect' and p1_confirm:
#                 p1_selected_rows = []
#                 show_p1_confirm = False
#             elif button_id == 'p3-confirm-deselect' and p3_confirm:
#                 p3_selected_rows = []
#                 show_p3_confirm = False
#             elif button_id == 'p4-confirm-deselect' and p4_confirm:
#                 p4_selected_rows = []
#                 show_p4_confirm = False
#             # Handle initial deselect button clicks
#             elif button_id == 'p1-deselect-all':
#                 show_p1_confirm = True
#             elif button_id == 'p3-deselect-all':
#                 show_p3_confirm = True
#             elif button_id == 'p4-deselect-all':
#                 show_p4_confirm = True
#             # Handle select all button clicks
#             elif button_id == 'p1-select-all' and p1_table_state:
#                 p1_selected_rows = list(range(len(p1_table_state)))
#             elif button_id == 'p3-select-all' and p3_table_state:
#                 p3_selected_rows = list(range(len(p3_table_state)))
#             elif button_id == 'p4-select-all' and p4_table_state:
#                 p4_selected_rows = list(range(len(p4_table_state)))

#         # Process visible ranges
#         visible_ranges = {}
#         for phase, relayout_data in [('P1', p1_relayout), ('P3', p3_relayout), ('P4', p4_relayout)]:
#             if relayout_data and 'xaxis.range[0]' in relayout_data:
#                 visible_ranges[phase] = (
#                     relayout_data['xaxis.range[0]'],
#                     relayout_data['xaxis.range[1]']
#                 )
            
#         def get_combined_y_range(results_dict):
#             all_y_values = []
#             for result in results_dict.values():
#                 if isinstance(result.get('table'), pd.DataFrame) and not result['table'].empty:
#                     all_y_values.extend(result['table']['Value'].dropna().tolist())

#             if all_y_values:
#                 min_y = min(all_y_values)
#                 max_y = max(all_y_values)
#                 padding = (max_y - min_y) * 0.05
#                 return [min_y - padding, max_y + padding]
#             return None
        
#         # Process each phase
#         for phase, phase_sheets, selected_rows, table_data, table_state in [
#             ('P1', p1_sheets, p1_selected_rows, p1_table_data, p1_table_state),
#             ('P3', p3_sheets, p3_selected_rows, p3_table_data, p3_table_state),
#             ('P4', p4_sheets, p4_selected_rows, p4_table_data, p4_table_state)
#         ]:
#             if not phase_sheets:
#                 continue

#             # Create graph and table
#             fig, table = create_phase_graph_and_table(
#                 sheet_data, 
#                 measurement_data, 
#                 phase, 
#                 phase_sheets,
#                 include_list, 
#                 exclude_list,
#                 include_event, 
#                 exclude_event, 
#                 show_actions,
#                 int(window_size)
#             )

#             # Handle selection state and delta calculations
#             if isinstance(table, pd.DataFrame) and not table.empty:
#                 if selected_rows is None or len(selected_rows) == 0:
#                     selected_rows = list(range(len(table)))
                
#                 selected_rows = [i for i in selected_rows if i < len(table)]
#                 table['selected'] = False
#                 if selected_rows:
#                     table.iloc[selected_rows, table.columns.get_loc('selected')] = True
                
#                 table = calculate_auto_peak_deltas(table)
#                 table = calculate_manual_deltas(table)
                
#                 # Update graph with selected points
#                 selected_points = table[table['selected']]
                
#                 # Keep all non-marker traces (lines, VAS bars, text)
#                 base_traces = []
#                 for trace in fig.data:
#                     # Keep line traces for main data
#                     if isinstance(trace, go.Scatter) and trace.mode == 'lines':
#                         base_traces.append(trace)
#                     # Keep VAS bar traces
#                     elif isinstance(trace, go.Bar):
#                         base_traces.append(trace)
#                     # Keep text traces for annotations
#                     elif isinstance(trace, go.Scatter) and trace.mode == 'text':
#                         base_traces.append(trace)
#                 fig.data = tuple(base_traces)  # Update figure data
                
#                 # Add back selected points with appropriate styling
#                 for point_type in ['Rolling Base', 'Rolling Peak', 'Detected Base', 'Detected Peak', 'Manual Base', 'Manual Peak']:
#                     points = selected_points[selected_points['Type'] == point_type]
#                     if not points.empty:
#                         marker_color = {
#                             'Rolling Base': 'rgb(102, 255, 51)',
#                             'Rolling Peak': 'rgb(0, 176, 240)',
#                             'Detected Base': COLORS['detected_base'],
#                             'Detected Peak': COLORS['detected_peak'],
#                             'Manual Base': COLORS['mbase'],
#                             'Manual Peak': COLORS['mpeak']
#                         }.get(point_type)
                        
#                         marker_symbol = 'diamond' if 'Manual' in point_type else 'circle'
#                         marker_size = 8 if 'Manual' in point_type else 12
                        
#                         fig.add_trace(
#                             go.Scatter(
#                                 name=f'{phase} ({point_type})',
#                                 x=points['Offset'],
#                                 y=points['Value'],
#                                 mode='markers',
#                                 marker=dict(
#                                     color=marker_color,
#                                     size=marker_size,
#                                     symbol=marker_symbol,
#                                     line=dict(color='white', width=1)
#                                 ),
#                                 showlegend=True
#                             ),
#                             secondary_y=False  # Add points on primary y-axis
#                         )

#                 results[phase] = {
#                     'fig': fig,
#                     'table': table,
#                     'selected_rows': selected_rows if selected_rows is not None else list(range(len(table) if isinstance(table, pd.DataFrame) else 0))
#                 }

#         # Calculate y-range for all graphs
#         y_range = get_combined_y_range(results)
#         if y_range:
#             for phase, relayout_data in [('P1', p1_relayout), ('P3', p3_relayout), ('P4', p4_relayout)]:
#                 if phase in results:
#                     if relayout_data and 'yaxis.range[0]' in relayout_data:
#                         manual_y_range = [relayout_data['yaxis.range[0]'], relayout_data['yaxis.range[1]']]
#                         results[phase]['fig'].update_layout(yaxis=dict(range=manual_y_range))
#                     else:
#                         results[phase]['fig'].update_layout(yaxis=dict(range=y_range))

#         # Calculate event statistics
#         p1_stats = calculate_event_statistics(sheet_data, p1_sheets) if p1_sheets else []
#         p1_stats = filter_selected_points(p1_stats, p1_table_data)

#         p3_stats = calculate_event_statistics(sheet_data, p3_sheets) if p3_sheets else []
#         p3_stats = filter_selected_points(p3_stats, p3_table_data)

#         p4_stats = calculate_event_statistics(sheet_data, p4_sheets) if p4_sheets else []
#         p4_stats = filter_selected_points(p4_stats, p4_table_data)

#         # Get results for each phase
#         empty_figure = go.Figure()
#         empty_data = []

#         p1_result = results.get('P1', {'fig': empty_figure, 'table': pd.DataFrame(), 'selected_rows': []})
#         p3_result = results.get('P3', {'fig': empty_figure, 'table': pd.DataFrame(), 'selected_rows': []})
#         p4_result = results.get('P4', {'fig': empty_figure, 'table': pd.DataFrame(), 'selected_rows': []})
        
#         # Return all components
#         return (
#             p1_result['fig'],
#             p1_result['table'].to_dict('records') if not p1_result['table'].empty else empty_data,
#             list(range(len(p1_result['table']))) if p1_result['table'].empty else p1_result['selected_rows'],
#             p1_stats,
#             p3_result['fig'],
#             p3_result['table'].to_dict('records') if not p3_result['table'].empty else empty_data,
#             list(range(len(p3_result['table']))) if p3_result['table'].empty else p3_result['selected_rows'],
#             p3_stats,
#             p4_result['fig'],
#             p4_result['table'].to_dict('records') if not p4_result['table'].empty else empty_data,
#             list(range(len(p4_result['table']))) if p4_result['table'].empty else p4_result['selected_rows'],
#             p4_stats,
#             toggle_status,
#             show_p1_confirm,
#             show_p3_confirm,
#             show_p4_confirm
#         )

#     except Exception as e:
#         print(f"Error in update_graphs_and_tables: {str(e)}")
#         traceback.print_exc()
#         return create_default_outputs()

### End Peak Detection Tab 6 ###


# In[10]:


### Overlay Graph ###

def determine_peak_color(delta, is_positive=False):
    """
    Determine segment color based on delta magnitude and direction
    Returns shades of purple for positive peaks, otherwise colors based on negative delta magnitude
    """
    if is_positive:
        abs_delta = abs(delta)
        if abs_delta >= 10.000:
            return 'rgb(110, 30, 110)'    # Dark Purple for strongest positive peaks
        elif 6.000 <= abs_delta <= 9.999:
            return 'rgb(160, 40, 160)'    # Purple for strong peaks
        elif 3.000 <= abs_delta <= 5.999:
            return 'rgb(210, 70, 210)'    # Light Purple for moderate peaks
        elif 0.050 <= abs_delta <= 2.999:
            return 'rgb(240, 170, 240)'   # Lavender for mild peaks
        return 'rgb(0, 0, 255)'           # Blue for baseline
    
    # For negative peaks
    abs_delta = abs(delta)
    if abs_delta >= 10.000:
        return 'rgb(0, 0, 0)'           # Black for strongest peaks
    elif 6.000 <= abs_delta <= 9.999:
        return 'rgb(255, 0, 0)'         # Red for strong peaks
    elif 3.000 <= abs_delta <= 5.999:
        return 'rgb(255, 165, 0)'       # Orange for moderate peaks
    elif 0.050 <= abs_delta <= 2.999:
        return 'rgb(255, 255, 0)'       # Yellow for mild peaks
    return 'rgb(0, 0, 255)'             # Blue for weak/baseline

def detect_peaks_with_deltas(data, sheet_name, window_size=1):
    """
    Detect peaks and calculate deltas between peaks and bases,
    including duration and slope information.
    
    Args:
        data: DataFrame containing Offset and Value columns
        sheet_name: Name of the sheet for logging
        window_size: Size of the rolling window for peak detection (percentage of data length)
    
    Returns:
        DataFrame with peak information including deltas, segments, durations and slopes
    """
    try:
        if data.empty:
            return pd.DataFrame()
            
        # Create a copy to avoid modifying the original data
        normalized_data = data.copy()
        
        # Calculate rolling window statistics
        window = max(3, int(len(normalized_data) * window_size / 100))
        normalized_data['rolling_mean'] = normalized_data['Value'].rolling(window=window, center=True, min_periods=1).mean()
        normalized_data['rolling_std'] = normalized_data['Value'].rolling(window=window, center=True, min_periods=1).std()
        
        # Filter for significant peaks and bases
        threshold = normalized_data['rolling_std'].mean() * 1.0
        
        # Find local maxima and minima
        peaks_data = []
        
        for i in range(1, len(normalized_data) - 1):
            current = normalized_data.iloc[i]['Value']
            prev_val = normalized_data.iloc[i-1]['Value']
            next_val = normalized_data.iloc[i+1]['Value']
            
            # Check for peaks (local maxima)
            if current > prev_val and current > next_val and abs(current - normalized_data.iloc[i]['rolling_mean']) > threshold:
                peaks_data.append({
                    'index': i,
                    'Offset': normalized_data.iloc[i]['Offset'],
                    'Value': current,
                    'Type': 'Peak'
                })
            
            # Check for bases (local minima)
            elif current < prev_val and current < next_val and abs(current - normalized_data.iloc[i]['rolling_mean']) > threshold:
                peaks_data.append({
                    'index': i,
                    'Offset': normalized_data.iloc[i]['Offset'],
                    'Value': current,
                    'Type': 'Base'
                })
        
        # Convert to DataFrame and sort by offset
        peaks_df = pd.DataFrame(peaks_data)
        if peaks_df.empty:
            return pd.DataFrame()
            
        peaks_df = peaks_df.sort_values('Offset').reset_index(drop=True)
        
        # Initialize result storage
        result_data = []
        
        # Process each peak/base to find pairs
        for i in range(len(peaks_df) - 1):
            current = peaks_df.iloc[i]
            next_point = peaks_df.iloc[i + 1]
            
            # Check for peak-base or base-peak pairs
            if (current['Type'] == 'Peak' and next_point['Type'] == 'Base') or \
               (current['Type'] == 'Base' and next_point['Type'] == 'Peak'):
                
                # Calculate delta
                delta = next_point['Value'] - current['Value']
                
                # Get segment indices between the two points
                start_idx = current['index']
                end_idx = next_point['index']
                segment_indices = list(range(start_idx, end_idx + 1))
                
                # Calculate duration (new)
                duration = normalized_data.iloc[end_idx]['Offset'] - normalized_data.iloc[start_idx]['Offset']
                
                # Calculate slope (new)
                slope = delta / duration if duration > 0 else 0
                
                # Store result - include all original fields plus new ones
                result_data.append({
                    'start_idx': start_idx,
                    'end_idx': end_idx,
                    'start_offset': current['Offset'],
                    'end_offset': next_point['Offset'],
                    'start_value': current['Value'],
                    'end_value': next_point['Value'],
                    'delta': delta,
                    'abs_delta': abs(delta),
                    'is_positive': delta > 0,
                    'segment_indices': segment_indices,
                    'duration': duration,  # New field
                    'slope': slope         # New field
                })
        
        # Convert to DataFrame
        result_df = pd.DataFrame(result_data)
        
        # Ensure we maintain existing format and fields
        if not result_df.empty:
            # Log information about detected peaks (keep existing logging)
            print(f"Detected peaks for {sheet_name}: {len(result_df)} peak-base pairs")
            print(f"  Positive deltas: {sum(result_df['is_positive'])} | Negative deltas: {sum(~result_df['is_positive'])}")
            
        return result_df
        
    except Exception as e:
        print(f"Error detecting peaks for {sheet_name}: {str(e)}")
        traceback.print_exc()
        return pd.DataFrame()

def process_peak_deltas(phase_data, filtered_data):
    """
    Process peak data to find maximum negative delta and its corresponding offset
    Returns the maximum delta and its offset
    """
    max_delta = None
    max_offset = None
    
    if not phase_data.empty:
        # Filter for negative deltas only
        neg_deltas = phase_data[phase_data['neg_delta'].notna()]
        if not neg_deltas.empty:
            # Find the maximum absolute negative delta
            max_idx = neg_deltas['neg_delta'].abs().idxmax()
            max_delta = neg_deltas.loc[max_idx, 'neg_delta']
            max_offset = neg_deltas.loc[max_idx, 'Offset']
    
    return max_delta, max_offset

def process_remaining_segments(normalized_data, peak_segments, phase_color):
    """
    Process segments that aren't part of peak-base pairs or baseline
    Returns segments that should be colored with phase color
    """
    all_points = set(range(len(normalized_data)))
    used_points = set()
    for segment in peak_segments:
        used_points.update(segment)
        
    remaining_points = sorted(list(all_points - used_points))
    
    # Split into continuous segments
    segments = []
    current_segment = []
    
    for i, idx in enumerate(remaining_points):
        if not current_segment or idx - remaining_points[i-1] == 1:
            current_segment.append(idx)
        else:
            if current_segment:
                segments.append(current_segment)
            current_segment = [idx]
            
    if current_segment:
        segments.append(current_segment)
        
    return segments

def find_baseline_segments(normalized_data, peak_segments):
    """
    Find segments where rolling average is below 0.05
    Returns list of baseline segment indices
    """
    # Calculate rolling average
    rolling_avg = normalized_data['Value'].rolling(
        window=5,
        center=True,
        min_periods=1
    ).mean()
    
    # Find points that meet baseline criteria
    baseline_mask = abs(rolling_avg) < 0.05
    baseline_indices = normalized_data[baseline_mask].index
    
    # Remove points that are part of peak segments
    used_points = set()
    for segment in peak_segments:
        used_points.update(segment)
        
    baseline_indices = sorted(list(set(baseline_indices) - used_points))
    
    # Split into continuous segments
    segments = []
    current_segment = []
    
    for i, idx in enumerate(baseline_indices):
        if not current_segment or idx - baseline_indices[i-1] == 1:
            current_segment.append(idx)
        else:
            if current_segment:
                segments.append(current_segment)
            current_segment = [idx]
            
    if current_segment:
        segments.append(current_segment)
        
    return segments

def create_neg_delta_cluster_graph(p1_overlay_sheets, p3_overlay_sheets, p4_overlay_sheets, sheet_data,
                                 include_words_list=None, exclude_words_list=None,
                                 include_events_list=None, exclude_events_list=None):
    """
    Create a cluster graph for negative deltas showing the relationship between magnitude, slope, and duration
    across different phases (P1, P3, P4).
    
    Returns:
        go.Figure: A plotly figure containing a scatter plot for clustering visualization
    """
    try:
        print("\nCreating negative delta cluster graph")
        cluster_fig = go.Figure()
        
        # Define phase-specific marker symbols
        markers = {
            'P1': 'triangle-up',
            'P3': 'square',
            'P4': 'diamond'
        }
        
        # Store all negative delta data for clustering
        neg_delta_data = {'P1': [], 'P3': [], 'P4': []}
        
        # Process each phase
        for phase_name, selected_sheets in [
            ('P1', p1_overlay_sheets),
            ('P3', p3_overlay_sheets),
            ('P4', p4_overlay_sheets)
        ]:
            if not selected_sheets:
                continue
                
            for sheet_name in selected_sheets:
                if sheet_name not in sheet_data:
                    continue
                    
                # Get and filter data
                data = sheet_data[sheet_name]['data'].copy()
                if data.empty:
                    continue
                
                # Apply filters
                filtered_data = data.copy()
                if include_words_list:
                    filtered_data = filter_by_words(filtered_data, include_words_list, [])
                if exclude_words_list:
                    filtered_data = filter_by_words(filtered_data, [], exclude_words_list)
                if include_events_list:
                    filtered_data = filter_by_events(filtered_data, include_events_list, [])
                if exclude_events_list:
                    filtered_data = filter_by_events(filtered_data, [], exclude_events_list)
                
                if filtered_data.empty:
                    continue

                # Normalize data
                normalized_data = filtered_data.copy()
                min_offset = filtered_data['Offset'].min()
                normalized_data['Offset'] = (normalized_data['Offset'] - min_offset) / 1000
                normalized_data['Value'] = normalized_data['Value'] * 10

                # Detect peaks and calculate deltas
                peaks_df = detect_peaks_with_deltas(normalized_data, sheet_name)
                
                if not peaks_df.empty:
                    # Filter for negative deltas only
                    neg_peaks = peaks_df[~peaks_df['is_positive']].copy()
                    
                    # For each negative delta, calculate its slope and duration
                    for _, row in neg_peaks.iterrows():
                        segment_indices = row['segment_indices']
                        if segment_indices and len(segment_indices) >= 2:
                            segment_data = normalized_data.iloc[segment_indices].copy()
                            
                            # Calculate duration (in seconds)
                            duration = segment_data['Offset'].max() - segment_data['Offset'].min()
                            
                            # Calculate slope (change in value over time)
                            if duration > 0:
                                delta_abs = abs(row['delta'])
                                slope = delta_abs / duration
                                
                                # Store data for clustering
                                neg_delta_data[phase_name].append({
                                    'delta': delta_abs,
                                    'duration': duration,
                                    'slope': slope,
                                    'sheet': sheet_name
                                })
        
        # Create color mapping function based on delta magnitude
        def get_color(delta):
            if 0.500 <= delta < 2.999:
                return 'rgb(255, 255, 0)'  # Yellow
            elif 3.000 <= delta < 5.999:
                return 'rgb(255, 165, 0)'  # Orange
            elif 6.000 <= delta < 9.999:
                return 'rgb(255, 0, 0)'    # Red
            elif  delta >= 10.000:
                return 'rgb(0, 0, 0)'      # Black
            return 'rgb(0, 0, 255)'        # Blue for weak/baseline
        
        # Add traces for each phase
        for phase_name, marker_symbol in markers.items():
            if not neg_delta_data[phase_name]:
                continue
                
            # Extract data for plotting
            deltas = [item['delta'] for item in neg_delta_data[phase_name]]
            slopes = [item['slope'] for item in neg_delta_data[phase_name]]
            durations = [item['duration'] for item in neg_delta_data[phase_name]]
            sheets = [item['sheet'] for item in neg_delta_data[phase_name]]
            
            # Calculate marker sizes based on duration (scaled for visibility)
            marker_sizes = [max(10, min(50, d * 20)) for d in durations]
            
            # Get colors based on delta magnitudes
            marker_colors = [get_color(delta) for delta in deltas]
            
            # Create hover text with detailed information
            hover_texts = [
                f"Phase: {phase_name}<br>" +
                f"Sheet: {sheet}<br>" +
                f"Delta: {delta:.2f}<br>" +
                f"Duration: {duration:.2f}s<br>" +
                f"Slope: {slope:.2f}/s"
                for delta, duration, slope, sheet in zip(deltas, durations, slopes, sheets)
            ]
            
            # Add scatter trace for this phase
            cluster_fig.add_trace(
                go.Scatter(
                    x=deltas,
                    y=slopes,
                    mode='markers',
                    marker=dict(
                        symbol=marker_symbol,
                        size=marker_sizes,
                        color=marker_colors,
                        line=dict(width=1, color='black')
                    ),
                    name=phase_name,
                    text=hover_texts,
                    hoverinfo='text',

                )
            )
        
        # Add vertical lines for delta magnitude ranges
        range_lines = [1.0, 3.0, 6.0, 10.0]
        for line_val in range_lines:
            cluster_fig.add_shape(
                type="line",
                x0=line_val,
                y0=0,
                x1=line_val,
                y1=1,
                yref="paper",
                line=dict(
                    color="gray",
                    width=1,
                    dash="dash",
                )
            )
            
            # Add annotations for the ranges
            if line_val < 10.0:  # Don't annotate the last line
                next_val = range_lines[range_lines.index(line_val) + 1]
                mid_point = (line_val + next_val) / 2
                
                # Determine color for this range
                range_color = get_color((line_val + next_val) / 2)
                
                cluster_fig.add_annotation(
                    x=mid_point,
                    y=1.02,
                    yref="paper",
                    text=f"{line_val}-{next_val}",
                    showarrow=False,
                    font=dict(
                        color=range_color,
                        size=10
                    )
                )
        
        # Add annotation for >10.0 range
        cluster_fig.add_annotation(
            x=11.5,
            y=1.02,
            yref="paper",
            text=">10.0",
            showarrow=False,
            font=dict(
                color="black",
                size=10
            )
        )
        
        # Update layout
        cluster_fig.update_layout(
            title={
                'text': 'Negative Delta Cluster Analysis',
                'x': 0.5,
                'xanchor': 'center'
            },
            xaxis_title='Negative Delta Magnitude',
            yaxis_title='Slope (Delta/Duration)',
            plot_bgcolor='white',
            paper_bgcolor='white',
            hovermode='closest',
            legend=dict(
                title="Phase",
                yanchor="top",
                y=0.99,
                xanchor="right",
                x=0.99,
                bgcolor='rgba(255, 255, 255, 0.8)',
                bordercolor='black',
                borderwidth=1
            ),
            margin=dict(l=50, r=50, t=80, b=50),
            height=500
        )
        
        # Update axes
        cluster_fig.update_xaxes(
            showgrid=True,
            gridwidth=1,
            gridcolor='LightGray',
            zeroline=True,
            zerolinewidth=1,
            zerolinecolor='LightGray',
            range=[0.5, 15]  # Adjust range as needed
        )
        
        cluster_fig.update_yaxes(
            showgrid=True,
            gridwidth=1,
            gridcolor='LightGray',
            zeroline=True,
            zerolinewidth=1,
            zerolinecolor='LightGray'
        )
        
        return cluster_fig
        
    except Exception as e:
        print(f"Error in create_neg_delta_cluster_graph: {str(e)}")
        traceback.print_exc()
        return go.Figure()

def create_enhanced_phase_overlay_graph(p1_overlay_sheets, p3_overlay_sheets, p4_overlay_sheets, sheet_data,
                                      include_words_list=None, exclude_words_list=None,
                                      include_events_list=None, exclude_events_list=None,
                                      show_actions=False, window_size=1):
    try:
        print("\nCreating enhanced phase overlay graph")
        fig = go.Figure()
        
        colors = {
            'P1': 'rgb(30, 139, 188)',  # Aqua
            'P3': 'rgb(105, 220, 247)',  # Light Blue
            'P4': 'rgb(160, 169, 192)'   # Grey
        }
        
        # Store all peak-base pairs across all phases/sheets
        all_peak_base_pairs = {}
        
        def process_peaks_and_bases(peaks_df, normalized_data, sheet_name, phase_name):
            """Process peaks and bases to find pairs and calculate deltas"""
            if peaks_df.empty:
                return [], set()
                
            peaks = peaks_df[peaks_df['Type'] == 'Rolling Peak'].copy()
            bases = peaks_df[peaks_df['Type'] == 'Rolling Base'].copy()
            
            # Sort by offset
            peaks = peaks.sort_values('Offset').reset_index(drop=True)
            bases = bases.sort_values('Offset').reset_index(drop=True)
            
            peak_base_pairs = []
            points_in_peak_base = set()
            
            print(f"\nProcessing peaks for {phase_name} - {sheet_name}")
            print(f"Found {len(peaks)} peaks and {len(bases)} bases")
            
            # Process each peak to find matching base
            for peak_idx, peak_row in peaks.iterrows():
                try:
                    # Find closest preceding base
                    valid_bases = bases[bases['Offset'] < peak_row['Offset']]
                    
                    if valid_bases.empty:
                        print(f"No valid base found for peak at offset {peak_row['Offset']:.3f}")
                        continue
                        
                    matching_base = valid_bases.iloc[-1]
                    
                    # Calculate delta
                    delta = peak_row['Value'] - matching_base['Value']
                    abs_delta = abs(delta)
                    
                    if delta < 0:  # Only store negative deltas
                        # Get segment data
                        segment_mask = (normalized_data['Offset'] >= matching_base['Offset']) & \
                                     (normalized_data['Offset'] <= peak_row['Offset'])
                        segment_data = normalized_data[segment_mask]
                        
                        if not segment_data.empty:
                            peak_base_pairs.append({
                                'peak': peak_row,
                                'base': matching_base,
                                'delta': delta,
                                'abs_delta': abs_delta,
                                'segment_indices': segment_data.index.tolist()
                            })
                            points_in_peak_base.update(segment_data.index)
                            
                            print(f"Found peak-base pair {len(peak_base_pairs)}:")
                            print(f"Peak - Offset: {peak_row['Offset']:.3f}, Value: {peak_row['Value']:.3f}")
                            print(f"Base - Offset: {matching_base['Offset']:.3f}, Value: {matching_base['Value']:.3f}")
                            print(f"Delta: {delta:.3f} (abs: {abs_delta:.3f})")
                    
                except Exception as e:
                    print(f"Error processing peak {peak_idx}: {str(e)}")
                    continue
            
            return peak_base_pairs, points_in_peak_base

        # First pass - collect all peak-base pairs
        for phase_name, selected_sheets, phase_color in [
            ('P1', p1_overlay_sheets, colors['P1']),
            ('P3', p3_overlay_sheets, colors['P3']),
            ('P4', p4_overlay_sheets, colors['P4'])
        ]:
            if not selected_sheets:
                continue
                
            for sheet_name in selected_sheets:
                try:
                    if sheet_name not in sheet_data:
                        continue
                        
                    # Get and filter data
                    filtered_data = sheet_data[sheet_name]['data'].copy()
                    if filtered_data.empty:
                        continue
                    
                    # Apply filters
                    if include_words_list:
                        filtered_data = filter_by_words(filtered_data, include_words_list, [])
                    if exclude_words_list:
                        filtered_data = filter_by_words(filtered_data, [], exclude_words_list)
                    if include_events_list:
                        filtered_data = filter_by_events(filtered_data, include_events_list, [])
                    if exclude_events_list:
                        filtered_data = filter_by_events(filtered_data, [], exclude_events_list)
                    
                    if filtered_data.empty:
                        continue

                    # Normalize data
                    normalized_data = filtered_data.copy()
                    min_offset = filtered_data['Offset'].min()
                    normalized_data['Offset'] = (normalized_data['Offset'] - min_offset) / 1000
                    normalized_data['Value'] = normalized_data['Value'] * 10
                    normalized_data = normalized_data.reset_index(drop=True)
                    
                    # Detect peaks and find pairs
                    peaks_df = detect_peaks_rolling_window(normalized_data, sheet_name, window_size)
                    pairs, points = process_peaks_and_bases(peaks_df, normalized_data, sheet_name, phase_name)
                    
                    if pairs:
                        all_peak_base_pairs[sheet_name] = {
                            'pairs': pairs,
                            'points': points,
                            'data': normalized_data,
                            'phase': phase_name
                        }
                        
                except Exception as e:
                    print(f"Error processing sheet {sheet_name}: {str(e)}")
                    continue

        # Second pass - create visualization with delta-based coloring
        for sheet_name, sheet_info in all_peak_base_pairs.items():
            try:
                normalized_data = sheet_info['data']
                pairs = sheet_info['pairs']
                points_in_peak_base = sheet_info['points']
                phase_name = sheet_info['phase']
                
                # Calculate running average for baseline detection
                normalized_data['running_avg'] = normalized_data['Value'].rolling(
                    window=5, 
                    center=True, 
                    min_periods=1
                ).mean()
                
                # Initialize point tracking
                all_points = set(range(len(normalized_data)))
                
                # Add traces for peak-base pairs
                for pair_idx, pair in enumerate(pairs):
                    segment_indices = pair['segment_indices']
                    segment_data = normalized_data.loc[segment_indices]
                    
                    # Determine color based on delta
                    color = determine_peak_color(pair['delta'], False)
                    
                    fig.add_trace(
                        go.Scatter(
                            x=segment_data['Offset'],
                            y=segment_data['Value'],
                            mode='lines',
                            line=dict(color=color, width=2),
                            name=f'{phase_name} Peak {pair_idx+1} (Δ: {pair["delta"]:.1f})',
                            showlegend=True
                        )
                    )
                
                # Process baseline segments
                try:
                    baseline_mask = normalized_data['running_avg'].abs() < 0.05
                    baseline_indices = normalized_data[baseline_mask].index
                    points_in_baseline = set(baseline_indices) - points_in_peak_base
                    
                    if points_in_baseline:
                        baseline_data = normalized_data.loc[list(points_in_baseline)].sort_values('Offset')
                        current_segment = []
                        baseline_segments = []
                        
                        # Split into continuous segments
                        for idx in baseline_data.index:
                            if not current_segment or idx - current_segment[-1] == 1:
                                current_segment.append(idx)
                            else:
                                if current_segment:
                                    baseline_segments.append(current_segment)
                                current_segment = [idx]
                                
                        if current_segment:
                            baseline_segments.append(current_segment)
                        
                        # Add baseline segments
                        for segment in baseline_segments:
                            segment_data = normalized_data.loc[segment]
                            fig.add_trace(
                                go.Scatter(
                                    x=segment_data['Offset'],
                                    y=segment_data['Value'],
                                    mode='lines',
                                    line=dict(color='rgb(0, 0, 255)', width=2),
                                    name=f'{phase_name} Baseline',
                                    showlegend=False
                                )
                            )
                            
                except Exception as e:
                    print(f"Error processing baseline segments: {str(e)}")
                
                # Process remaining segments
                try:
                    remaining_points = all_points - points_in_peak_base - points_in_baseline
                    if remaining_points:
                        remaining_indices = sorted(list(remaining_points))
                        current_segment = []
                        segments = []
                        
                        for i, idx in enumerate(remaining_indices):
                            if not current_segment or idx - remaining_indices[i-1] == 1:
                                current_segment.append(idx)
                            else:
                                if current_segment:
                                    segments.append(current_segment)
                                current_segment = [idx]
                                
                        if current_segment:
                            segments.append(current_segment)
                            
                        # Add remaining segments with phase color
                        for segment in segments:
                            segment_data = normalized_data.loc[segment]
                            fig.add_trace(
                                go.Scatter(
                                    x=segment_data['Offset'],
                                    y=segment_data['Value'],
                                    mode='lines',
                                    line=dict(color=colors[phase_name], width=2),
                                    name=phase_name,
                                    showlegend=False,
                                    hoverinfo='text'
                                )
                            )
                            
                except Exception as e:
                    print(f"Error processing remaining segments: {str(e)}")
                    
                # Add action labels if enabled
                if show_actions:
                    try:
                        used_positions = []
                        
                        # Calculate event statistics
                        event_stats = {}
                        event_starts = {}
                        for event_name, event_group in normalized_data.groupby('Event'):
                            if pd.notna(event_name) and event_name != '':
                                avg_value = event_group['Value'].mean()
                                event_stats[event_name] = avg_value
                                event_starts[event_name] = event_group['Offset'].min()
                        
                        # Add annotations
                        non_empty_annotations = normalized_data[normalized_data['Annotations'].notna()]
                        if not non_empty_annotations.empty:
                            labeled_events = set()
                            for _, row in non_empty_annotations.sort_values('Offset').iterrows():
                                event_name = row['Event']
                                x_pos = row['Offset']
                                base_y = row['Value'] + 0.75
                                
                                adjusted_y = find_available_y_position(
                                    x_pos, 
                                    base_y, 
                                    used_positions,
                                    min_gap=1.0
                                )
                                
                                if event_name in event_stats and event_name not in labeled_events and \
                                   row['Offset'] == event_starts[event_name]:
                                    avg_text = f"\nAvg: {event_stats[event_name]:.2f}"
                                    labeled_events.add(event_name)
                                else:
                                    avg_text = ""
                                    
                                fig.add_trace(
                                    go.Scatter(
                                        name=f'{phase_name} (Actions)',
                                        x=[x_pos],
                                        y=[adjusted_y],
                                        mode='text',
                                        text=[f"{row['Annotations']}{avg_text}"],
                                        textposition='top center',
                                        textfont=dict(size=12, color='black'),
                                        showlegend=False,
                                        hoverinfo='text',
                                        hovertext=f'{phase_name} {sheet_name}: {abs(row["delta"]):.2f}'
                                    )
                                )
                                
                                used_positions.append((x_pos, adjusted_y))
                                
                    except Exception as e:
                        print(f"Error adding action labels: {str(e)}")
                        
            except Exception as e:
                print(f"Error processing visualization for sheet {sheet_name}: {str(e)}")
                continue

        # Create negative delta bar graph with all pairs
        indiv_neg_traces = []
        for phase_name in ['P1', 'P3', 'P4']:
            phase_pairs = []
            for sheet_name, info in all_peak_base_pairs.items():
                if info['phase'] == phase_name:
                    phase_pairs.extend([(pair, sheet_name) for pair in info['pairs']])
            
            if phase_pairs:
                # Sort pairs by offset
                phase_pairs.sort(key=lambda x: x[0]['peak']['Offset'])
                
                indiv_neg_traces.append(go.Bar(
                    name=f'{phase_name} Negative Deltas',
                    x=[f'{phase_name}_{i+1}' for i in range(len(phase_pairs))],
                    y=[abs(pair[0]['delta']) for pair in phase_pairs],
                    text=[f'Δ: {abs(pair[0]["delta"]):.1f}<br>Offset: {pair[0]["peak"]["Offset"]:.1f}<br>{sheet}'
                          for pair, sheet in phase_pairs],
                    textposition='outside',
                    textangle=0
                ))

        # Create figures
        neg_delta_fig = go.Figure(data=indiv_neg_traces)
        neg_delta_fig.update_layout(
            title='Individual Negative Peak Deltas by Phase',
            xaxis_title='Peak Number',
            yaxis_title='Absolute Delta Value',
            showlegend=True,
            plot_bgcolor='white',
            paper_bgcolor='white'
        )
        
        # Update main figure layout
        fig.update_layout(
            title={
                'text': 'Enhanced Phase Overlay Comparison (Time Normalized)',
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

        return fig, neg_delta_fig

    except Exception as e:
        print(f"Error in create_enhanced_phase_overlay_graph: {str(e)}")
        traceback.print_exc()
        return go.Figure(), go.Figure()

@app.callback(
    [Output('overlay-graph', 'figure'),
     Output('overlay-action-toggle-status', 'children'),
     Output('overlay-action-toggle-status', 'style'),
     Output('overlay-indiv-neg-delta-graph', 'figure'),
     Output('overlay-indiv-pos-delta-graph', 'figure'),
     Output('overlay-neg-delta-graph', 'figure'),
     Output('overlay-pos-delta-graph', 'figure'),
     Output('overlay-neg-delta-cluster-graph', 'figure')],  
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
    try:
        show_actions = bool(n_clicks and n_clicks % 2)
        toggle_text = "Action Labels ON" if show_actions else "Action Labels OFF"
        toggle_style = {'color': 'green'} if show_actions else {'color': 'red'}
        
        print("\nUpdating overlay graph")
        print(f"Selected P1 sheets: {p1_overlay_sheets}")
        print(f"Selected P3 sheets: {p3_overlay_sheets}")
        print(f"Selected P4 sheets: {p4_overlay_sheets}")
        
        # Create empty figure if no sheets selected
        if not any([p1_overlay_sheets, p3_overlay_sheets, p4_overlay_sheets]):
            empty_fig = go.Figure()
            empty_fig.update_layout(
                title='Phase Overlay Comparison (No Data Selected)',
                xaxis_title='Time (seconds from start)',
                yaxis_title='Value',
                showlegend=False,
                plot_bgcolor='white',
                paper_bgcolor='white'
            )
            return [empty_fig, toggle_text, toggle_style, empty_fig, empty_fig, empty_fig, empty_fig, empty_fig]  # Added empty_fig for cluster graph

        # Process filters
        include_words_list = [word.strip() for word in include_words.split(',') if word.strip()] if include_words else []
        exclude_words_list = [word.strip() for word in exclude_words.split(',') if word.strip()] if exclude_words else []
        include_events_list = [event.strip() for event in include_events.split(',') if event.strip()] if include_events else []
        exclude_events_list = [event.strip() for event in exclude_events.split(',') if event.strip()] if exclude_events else []

        # Event set mappings (preserve existing event sets)
        event_sets = {
            'Event Set 1: Walk': {
                'P1': '1 Event',
                'P3': '3 Event', 
                'P4': '5 Event',
                'color': 'rgb(31, 119, 180)'  # Blue
            },
            'Activity Set 1: Ext Less Painful': {
                'P1': '1 Activity',
                'P3': '5 Activity',
                'P4': '9 Activity', 
                'color': 'rgb(255, 127, 14)'  # Orange
            },
            'Activity Set 2: Flex Less Painful': {
                'P1': '2 Activity',
                'P3': '6 Activity',
                'P4': '10 Activity',
                'color': 'rgb(44, 160, 44)'  # Green
            },
            'Activity Set 3: Ext More Painful': {
                'P1': '3 Activity',
                'P3': '7 Activity',
                'P4': '11 Activity',
                'color': 'rgb(214, 39, 40)'  # Red
            },
            'Activity Set 4: Flex More Painful': {
                'P1': '4 Activity',
                'P3': '8 Activity',
                'P4': '12 Activity',
                'color': 'rgb(148, 103, 189)'  # Purple
            },
            'Event Set 2: Up and Go': {
                'P1': '2 Event',
                'P3': '4 Event',
                'P4': '6 Event',
                'color': 'rgb(140, 86, 75)'  # Brown
            }
        }

        # Initialize storage for delta data
        indiv_pos_data = {}
        negative_deltas = {'P1': [], 'P3': [], 'P4': []}
        pos_deltas = {'P1': [], 'P3': [], 'P4': []}
        
        # Initialize storage for cluster graph data (new)
        cluster_data = {'P1': [], 'P3': [], 'P4': []}
        
        # Track which legend items have been added
        shown_legend_items = set()
        
        # Create main figure
        main_fig = go.Figure()
        
        # Process each phase
        for phase_name, selected_sheets, phase_color in [
            ('P1', p1_overlay_sheets, 'rgb(30, 139, 188)'),
            ('P3', p3_overlay_sheets, 'rgb(105, 220, 247)'),
            ('P4', p4_overlay_sheets, 'rgb(160, 169, 192)')
        ]:
            if not selected_sheets:
                continue
                
            for sheet_name in selected_sheets:
                if sheet_name not in sheet_data:
                    continue
                    
                # Get and filter data
                data = sheet_data[sheet_name]['data'].copy()
                if data.empty:
                    continue
                
                # Apply filters
                filtered_data = data.copy()
                if include_words_list:
                    filtered_data = filter_by_words(filtered_data, include_words_list, [])
                if exclude_words_list:
                    filtered_data = filter_by_words(filtered_data, [], exclude_words_list)
                if include_events_list:
                    filtered_data = filter_by_events(filtered_data, include_events_list, [])
                if exclude_events_list:
                    filtered_data = filter_by_events(filtered_data, [], exclude_events_list)
                
                if filtered_data.empty:
                    continue

                # Normalize data
                normalized_data = filtered_data.copy()
                min_offset = filtered_data['Offset'].min()
                normalized_data['Offset'] = (normalized_data['Offset'] - min_offset) / 1000
                normalized_data['Value'] = normalized_data['Value'] * 10

                # Store all peak-base segment indices
                peak_segments = []
                
                # Detect peaks and calculate deltas
                peaks_df = detect_peaks_with_deltas(normalized_data, sheet_name)
                
                if not peaks_df.empty:
                    # Store deltas for bar graphs
                    for _, row in peaks_df.iterrows():
                        if not row['is_positive']:
                            negative_deltas[phase_name].append(row['delta'])
                            
                            # Store additional data for cluster graph (new)
                            cluster_data[phase_name].append({
                                'delta': abs(row['delta']),
                                'duration': row['duration'],
                                'slope': abs(row['slope']),
                                'sheet': sheet_name
                            })
                        else:
                            # Store positive deltas separately if needed
                            pos_deltas[phase_name].append(abs(row['delta']))
                    
                    # Add colored segments for each peak-base pair
                    for _, row in peaks_df.iterrows():
                        segment_indices = row['segment_indices']
                        if segment_indices:  # Check if we have indices
                            peak_segments.append(segment_indices)
                            extended_indices = segment_indices
                            if min(segment_indices) > 0:
                                extended_indices = [min(segment_indices) - 1] + extended_indices
                            if max(segment_indices) < len(normalized_data) - 1:
                                extended_indices = extended_indices + [max(segment_indices) + 1]
                            segment_data = normalized_data.iloc[extended_indices].copy()
                            
                            # Ensure we have at least two points for the line
                            if len(segment_data) >= 2:
                                color = determine_peak_color(row['delta'], row['is_positive'])
                                label_text = 'Positive Δ' if row['is_positive'] else 'Negative Δ'
                                legend_key = f"{phase_name} {label_text}"
                                
                                # Only show in legend if it's the first time this type appears
                                show_in_legend = legend_key not in shown_legend_items
                                if show_in_legend:
                                    shown_legend_items.add(legend_key)
                                    
                                main_fig.add_trace(
                                    go.Scatter(
                                        x=segment_data['Offset'],
                                        y=segment_data['Value'],
                                        mode='lines',
                                        line=dict(color=color, width=2),
                                        name=legend_key,
                                        showlegend=show_in_legend,
                                        legendgroup=legend_key,  # Group by phase+delta type
                                        hoverinfo='text',
                                        hovertext=f'{phase_name} {sheet_name} {label_text}: {abs(row["delta"]):.2f}'
                                    )
                                )

                # Find baseline segments
                baseline_segments = find_baseline_segments(normalized_data, peak_segments)
                
                # Add baseline segments
                for segment_indices in baseline_segments:
                    if len(segment_indices) >= 2:  # Ensure we have at least two points for the line
                        segment_data = normalized_data.iloc[segment_indices].copy()
                        main_fig.add_trace(
                            go.Scatter(
                                x=segment_data['Offset'],
                                y=segment_data['Value'],
                                mode='lines',
                                line=dict(color='rgb(0, 0, 255)', width=2),
                                name='Baseline',
                                showlegend=False,
                                legendgroup='baselines',
                                hoverinfo='text',
                                hovertext=f'{phase_name} {sheet_name} {label_text}: {abs(row["delta"]):.2f}'
                            )
                        )
                
                # Process remaining segments
                remaining_segments = process_remaining_segments(normalized_data, peak_segments + baseline_segments, phase_color)

                # Add remaining segments with phase color
                if not any(tr.name == phase_name for tr in main_fig.data):
                    show_in_legend = True
                else: 
                    show_in_legend = False

                for segment_indices in remaining_segments:
                    if len(segment_indices) >= 2:
                        segment_data = normalized_data.iloc[segment_indices].copy()
                        main_fig.add_trace(
                            go.Scatter(
                                x=segment_data['Offset'],
                                y=segment_data['Value'],
                                mode='lines', 
                                line=dict(color=phase_color, width=2),
                                name=phase_name,
                                showlegend=False,
                                legendgroup=phase_name,
                                hoverinfo='text',
                            )
                        )

                # Add action labels if enabled
                if show_actions:
                    used_positions = []
                    non_empty_annotations = normalized_data[normalized_data['Annotations'].notna()]
                    if not non_empty_annotations.empty:
                        for _, row in non_empty_annotations.iterrows():
                            x_pos = row['Offset']
                            base_y = row['Value'] + 0.75
                            
                            adjusted_y = find_available_y_position(
                                x_pos, 
                                base_y, 
                                used_positions,
                                min_gap=1.0
                            )
                            
                            main_fig.add_trace(
                                go.Scatter(
                                    name=f'{phase_name} (Actions)',
                                    x=[x_pos],
                                    y=[adjusted_y],
                                    mode='text',
                                    text=[f"{row['Annotations']}"],
                                    textposition='top center',
                                    textfont=dict(size=12, color='black'),
                                    showlegend=False,
                                    hoverinfo='skip'
                                )
                            )
                            
                            used_positions.append((x_pos, adjusted_y))

        # Create negative deltas bar graph
        neg_data = []
        for phase_name in ['P1', 'P3', 'P4']:
            if negative_deltas[phase_name]:
                mean_delta = abs(np.mean(negative_deltas[phase_name]))
                color = determine_peak_color(-mean_delta, False)  # Negative to match peak-base direction
                neg_data.append(go.Bar(
                    name=phase_name,
                    x=[phase_name],
                    y=[mean_delta],
                    error_y=dict(
                        type='data',
                        array=[np.std(negative_deltas[phase_name])],
                        visible=True
                    ),
                    marker_color=color,
                    text=[f'Δ: {mean_delta:.1f}'],
                    textposition='outside'
                ))

        neg_fig = go.Figure(data=neg_data)
        neg_fig.update_layout(
            title='Average Negative Peak Deltas by Phase',
            yaxis_title='Absolute Delta Value',
            showlegend=True,
            barmode='group',
            plot_bgcolor='white',
            paper_bgcolor='white'
        )

        # Create individual negative deltas graph
        indiv_neg_traces = []
        for phase_name in ['P1', 'P3', 'P4']:
            if negative_deltas[phase_name]:
                x_values = [f'{phase_name}_{i+1}' for i in range(len(negative_deltas[phase_name]))]
                abs_deltas = [abs(delta) for delta in negative_deltas[phase_name]]
                colors = [determine_peak_color(delta, False) for delta in negative_deltas[phase_name]]
                indiv_neg_traces.append(go.Bar(
                    name=f'{phase_name} Negative Deltas',
                    x=x_values,
                    y=[abs(delta) for delta in negative_deltas[phase_name]],
                    text=[f'Δ: {abs(delta):.1f}' for delta in negative_deltas[phase_name]],
                    textposition='outside',
                    textangle=0,
                    marker_color=colors,
                    showlegend=True
                ))

        indiv_neg_fig = go.Figure(data=indiv_neg_traces)
        indiv_neg_fig.update_layout(
            title='Individual Negative Peak Deltas by Phase',
            xaxis_title='Peak Number',
            yaxis_title='Absolute Delta Value',
            showlegend=True,
            plot_bgcolor='white',
            paper_bgcolor='white'
        )

        # Create positive deltas graph
        pos_data = []
        for phase_name in ['P1', 'P3', 'P4']:
            if pos_deltas[phase_name]:
                mean_delta = np.mean(pos_deltas[phase_name])
                color = determine_peak_color(mean_delta, True)
                pos_data.append(go.Bar(
                    name=phase_name,
                    x=[phase_name],
                    y=[mean_delta],
                    error_y=dict(
                        type='data',
                        array=[np.std(pos_deltas[phase_name])],
                        visible=True
                    ),
                    marker_color='rgb(112, 50, 160)',  # Purple for positive peaks
                    text=[f'Δ: {mean_delta:.1f}'],
                    textposition='outside'
                ))

        pos_fig = go.Figure(data=pos_data)
        pos_fig.update_layout(
            title='Average Positive Peak Deltas by Phase',
            yaxis_title='Delta Value',
            showlegend=True,
            barmode='group',
            plot_bgcolor='white',
            paper_bgcolor='white'
        )

        # Create individual positive deltas graph
        indiv_pos_traces = []
        for phase_name in ['P1', 'P3', 'P4']:
            if pos_deltas[phase_name]:
                x_values = [f'{phase_name}_{i+1}' for i in range(len(pos_deltas[phase_name]))]
                abs_deltas = [abs(delta) for delta in pos_deltas[phase_name]]
                colors = [determine_peak_color(delta, True) for delta in pos_deltas[phase_name]]  # Set is_positive to True
                indiv_pos_traces.append(go.Bar(
                    name=f'{phase_name} Positive Deltas',
                    x=x_values,
                    y=abs_deltas,
                    text=[f'Δ: {abs_delta:.1f}' for abs_delta in abs_deltas],
                    textposition='outside',
                    textangle=0,
                    marker_color=colors,
                    showlegend=True
                ))

        indiv_pos_fig = go.Figure(data=indiv_pos_traces)
        indiv_pos_fig.update_layout(
            title='Individual Positive Peak Deltas by Phase',
            xaxis_title='Peak Number',
            yaxis_title='Delta Value',
            showlegend=True,
            plot_bgcolor='white',
            paper_bgcolor='white'
        )
        
        # Create or update the cluster graph by calling the dedicated function
        cluster_fig = create_neg_delta_cluster_graph(
            p1_overlay_sheets, 
            p3_overlay_sheets, 
            p4_overlay_sheets, 
            sheet_data,
            include_words_list, 
            exclude_words_list,
            include_events_list, 
            exclude_events_list
        )
        
        # Define phase-specific marker symbols
        markers = {
            'P1': 'triangle-up',
            'P3': 'square',
            'P4': 'diamond'
        }
        
        # Define color function for delta magnitude ranges
        def get_color(delta):
            if 0.050 <= delta <= 2.999:
                return 'rgb(255, 255, 0)'  # Yellow
            elif 3.000 <= delta <= 5.999:
                return 'rgb(255, 165, 0)'  # Orange
            elif 6.000 <= delta <= 9.999:
                return 'rgb(255, 0, 0)'    # Red
            elif  delta >= 10.000:
                return 'rgb(0, 0, 0)'      # Black
            return 'rgb(0, 0, 255)'        # Blue for weak/baseline
        
        # Add traces for each phase to the cluster graph
        for phase_name, marker_symbol in markers.items():
            if not cluster_data[phase_name]:
                continue
                
            # Filter for deltas >= 1.0 only
            filtered_cluster_data = [item for item in cluster_data[phase_name] if item['delta'] >= 1.0]
            
            if not filtered_cluster_data:  # Skip if no data points remain after filtering
                continue
            
            # Extract data for plotting
            deltas = [item['delta'] for item in filtered_cluster_data]
            slopes = [item['slope'] for item in filtered_cluster_data]
            durations = [item['duration'] for item in filtered_cluster_data]
            sheets = [item['sheet'] for item in filtered_cluster_data]
            
            # Calculate marker sizes based on duration (scaled for visibility)
            marker_sizes = [max(10, min(50, d * 20)) for d in durations]
            
            # Get colors based on delta magnitudes
            marker_colors = [get_color(delta) for delta in deltas]
            
            # Create hover text with detailed information
            hover_texts = [
                f"Phase: {phase_name}<br>" +
                f"Sheet: {sheet}<br>" +
                f"Delta: {delta:.2f}<br>" +
                f"Duration: {duration:.2f}s<br>" +
                f"Slope: {slope:.2f}/s"
                for delta, duration, slope, sheet in zip(deltas, durations, slopes, sheets)
            ]
            
            # Add scatter trace for this phase
            cluster_fig.add_trace(
                go.Scatter(
                    x=deltas,
                    y=slopes,
                    mode='markers',
                    marker=dict(
                        symbol=marker_symbol,
                        size=marker_sizes,
                        color=marker_colors,
                        line=dict(width=1, color='black')
                    ),
                    name=phase_name,
                    text=hover_texts,
                    hoverinfo='text'
                )
            )
        
        # Add vertical lines for delta magnitude ranges
        range_lines = [1.0, 3.0, 6.0, 10.0]
        for line_val in range_lines:
            cluster_fig.add_shape(
                type="line",
                x0=line_val,
                y0=0,
                x1=line_val,
                y1=1,
                yref="paper",
                line=dict(
                    color="gray",
                    width=1,
                    dash="dash",
                )
            )
            
            # Add annotations for the ranges
            if line_val < 10.0:  # Don't annotate the last line
                next_val = range_lines[range_lines.index(line_val) + 1]
                mid_point = (line_val + next_val) / 2
                
                # Determine color for this range
                range_color = get_color((line_val + next_val) / 2)
                
                cluster_fig.add_annotation(
                    x=mid_point,
                    y=1.02,
                    yref="paper",
                    text=f"{line_val}-{next_val}",
                    showarrow=False,
                    font=dict(
                        color=range_color,
                        size=10
                    )
                )
        
        # Add annotation for >10.0 range
        cluster_fig.add_annotation(
            x=11.5,
            y=1.02,
            yref="paper",
            text=">10.0",
            showarrow=False,
            font=dict(
                color="black",
                size=10
            )
        )
        
        # Update cluster graph layout
        cluster_fig.update_layout(
            title={
                'text': 'Negative Delta Cluster Analysis',
                'x': 0.5,
                'xanchor': 'center'
            },
            xaxis_title='Negative Delta Magnitude',
            yaxis_title='Slope (Delta/Duration)',
            plot_bgcolor='white',
            paper_bgcolor='white',
            hovermode='closest',
            legend=dict(
                title="Phase",
                yanchor="top",
                y=0.99,
                xanchor="right",
                x=0.99,
                bgcolor='rgba(255, 255, 255, 0.8)',
                bordercolor='black',
                borderwidth=1
            ),
            margin=dict(l=50, r=50, t=80, b=50),
            height=500
        )
        
        # Update cluster graph axes
        cluster_fig.update_xaxes(
            showgrid=True,
            gridwidth=1,
            gridcolor='LightGray',
            zeroline=True,
            zerolinewidth=1,
            zerolinecolor='LightGray',
            range=[0.5, 15]  # Adjust range as needed
        )
        
        cluster_fig.update_yaxes(
            showgrid=True,
            gridwidth=1,
            gridcolor='LightGray',
            zeroline=True,
            zerolinewidth=1,
            zerolinecolor='LightGray'
        )

        # Update main figure layout
        main_fig.update_layout(
            title={
                'text': 'Enhanced Phase Overlay Comparison (Time Normalized)',
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
                xanchor="right",
                x=0.99,
                bgcolor='rgba(255, 255, 255, 0.8)',
                bordercolor='black',
                borderwidth=1,
                orientation="h",
                font=dict(size=8)
            ),
            margin=dict(l=50, r=50, t=50, b=50)  # Increased top margin for legend
        )
        
        # Update axes
        main_fig.update_xaxes(
            showgrid=True,
            gridwidth=1,
            gridcolor='LightGray',
            zeroline=True,
            zerolinewidth=1,
            zerolinecolor='LightGray'
        )
        
        main_fig.update_yaxes(
            showgrid=True,
            gridwidth=1,
            gridcolor='LightGray',
            zeroline=True,
            zerolinewidth=1,
            zerolinecolor='LightGray'
        )

        return [
            main_fig,             # overlay-graph
            toggle_text,          # overlay-action-toggle-status text
            toggle_style,         # overlay-action-toggle-status style
            indiv_neg_fig,        # overlay-indiv-neg-delta-graph
            indiv_pos_fig,        # overlay-indiv-pos-delta-graph
            neg_fig,              # overlay-neg-delta-graph
            pos_fig,              # overlay-pos-delta-graph
            cluster_fig           # overlay-neg-delta-cluster-graph
        ]

    except Exception as e:
        print(f"Error in update_overlay_graph: {str(e)}")
        traceback.print_exc()
        empty_fig = go.Figure()
        return [empty_fig, "Action Labels OFF", {'color': 'red'}, empty_fig, empty_fig, empty_fig, empty_fig, empty_fig]  # Added empty_fig for cluster 

# Save button callbacks
# @app.callback(
#     Output('p1-save-status', 'children', allow_duplicate=True),
#     [Input('save-p1-button', 'n_clicks')],
#     [State('p1-graph', 'figure'),
#      State('p1-table', 'data'),
#      State('p1-stats-table', 'data'),
#      State('p1-filename-input', 'value')]
# )
# def save_p1_results(n_clicks, fig_data, table_data, event_data, filename):
#     try:
#         if n_clicks > 0:
#             if not filename:
#                 filename = 'mnl_py_results'
#             success, message = save_single_phase_to_excel('P1', fig_data, table_data, event_data, filename)
#             return html.Div(message, style={'color': 'green' if success else 'red'})
#         return ""
#     except Exception as e:
#         print(f"Error saving P1 results: {str(e)}")
#         return html.Div(f"Error saving results: {str(e)}", style={'color': 'red'})

# @app.callback(
#     Output('p3-save-status', 'children', allow_duplicate=True),
#     [Input('save-p3-button', 'n_clicks')],
#     [State('p3-graph', 'figure'),
#      State('p3-table', 'data'),
#      State('p3-stats-table', 'data'),
#      State('p3-filename-input', 'value')]
# )
# def save_p3_results(n_clicks, fig_data, table_data, event_data, filename):
#     try:
#         if n_clicks > 0:
#             if not filename:
#                 filename = 'mnl_py_results'
#             success, message = save_single_phase_to_excel('P3', fig_data, table_data, event_data, filename)
#             return html.Div(message, style={'color': 'green' if success else 'red'})
#         return ""
#     except Exception as e:
#         print(f"Error saving P3 results: {str(e)}")
#         return html.Div(f"Error saving results: {str(e)}", style={'color': 'red'})

# @app.callback(
#     Output('p4-save-status', 'children', allow_duplicate=True),
#     [Input('save-p4-button', 'n_clicks')],
#     [State('p4-graph', 'figure'),
#      State('p4-table', 'data'),
#      State('p4-stats-table', 'data'),
#      State('p4-filename-input', 'value')]
# )
# def save_p4_results(n_clicks, fig_data, table_data, event_data, filename):
#     try:
#         if n_clicks > 0:
#             if not filename:
#                 filename = 'mnl_py_results'
#             success, message = save_single_phase_to_excel('P4', fig_data, table_data, event_data, filename)
#             return html.Div(message, style={'color': 'green' if success else 'red'})
#         return ""
#     except Exception as e:
#         print(f"Error saving P4 results: {str(e)}")
#         return html.Div(f"Error saving results: {str(e)}", style={'color': 'red'})

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
    print("Loading manual trace data...")
    valid_ids, manual_trace_data = read_manual_trace_data('Mnl_Analysis.xlsx')
    print(f"Successfully loaded manual trace data with {len(valid_ids)} valid IDs")
    
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

### End Overlay Graph ### 


# In[ ]:


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
# @app.callback(
#     [#Output('p1-overall-stats-graph', 'figure'),
#      #Output('p1-delta-stats-graph', 'figure'),
#      Output('p1-peak-counts-graph', 'figure'),
#      #Output('p1-stats-analysis', 'data'),
#      #Output('p3-overall-stats-graph', 'figure'),
#      #Output('p3-delta-stats-graph', 'figure'),
#      Output('p3-peak-counts-graph', 'figure'),
#      Output('p3-stats-analysis', 'data'),
#      Output('p4-overall-stats-graph', 'figure'),
#      Output('p4-delta-stats-graph', 'figure'),
#      Output('p4-peak-counts-graph', 'figure'),
#      Output('p4-stats-analysis', 'data'),
#      Output('p1-event-avg-graph', 'figure'),
#      Output('p3-event-avg-graph', 'figure'),
#      Output('p4-event-avg-graph', 'figure'),
#      Output('p1-event-delta-graph', 'figure'),
#      Output('p3-event-delta-graph', 'figure'),
#      Output('p4-event-delta-graph', 'figure')],
#     [Input('worksheet-results-selector', 'value')]
# )
# def update_phase_analyses(selected_files):
#     """Update all phase analyses based on selected files"""
#     try:
#         # Initialize empty defaults
#         empty_fig = go.Figure()
#         empty_fig.update_layout(
#             title="No data available",
#             xaxis_title="",
#             yaxis_title="",
#             showlegend=False
#         )
        
#         # Return empty state if no files selected
#         if not selected_files:
#             return [empty_fig] * 18  # Total of 18 outputs

#         # Initialize containers for each phase
#         phase_results = {
#             '1': {'overall_fig': empty_fig, 'delta_fig': empty_fig, 'counts_fig': empty_fig, 'stats_data': []},
#             '3': {'overall_fig': empty_fig, 'delta_fig': empty_fig, 'counts_fig': empty_fig, 'stats_data': []},
#             '4': {'overall_fig': empty_fig, 'delta_fig': empty_fig, 'counts_fig': empty_fig, 'stats_data': []}
#         }
#         event_graphs = {
#             '1': {'avg_fig': empty_fig, 'delta_fig': empty_fig},
#             '3': {'avg_fig': empty_fig, 'delta_fig': empty_fig},
#             '4': {'avg_fig': empty_fig, 'delta_fig': empty_fig}
#         }
        
#         # Load and process data for each phase
#         phase_data = {}
#         event_data = {}
        
#         # First, load and process all data for each phase
#         for phase in ['1', '3', '4']:
#             print(f"Phase {phase} stats_data type: {type(phase_results[phase]['stats_data'])}")
#             print(f"Phase {phase} stats_data: {phase_results[phase]['stats_data']}")
#             if phase_results[phase]['stats_data'] is None:
#                 phase_results[phase]['stats_data'] = []
#             elif not isinstance(phase_results[phase]['stats_data'], list):
#                 try:
#                     # Attempt to convert to list if possible
#                     phase_results[phase]['stats_data'] = list(phase_results[phase]['stats_data'])
#                 except:
#                     # If conversion fails, use an empty list
#                     phase_results[phase]['stats_data'] = []
            
#             try:
#                 # Process peak/base data
#                 peak_dfs = []
#                 event_dfs = []
                
#                 for file in selected_files:
#                     try:
#                         # Load peak/base data
#                         df = load_peaks_data(file, phase)
#                         if not df.empty:
#                             stats = calculate_phase_statistics(df)
#                             stats['number'], stats['phase'] = parse_filename(file)
#                             # Round numerical values
#                             for col in stats.index:
#                                 if isinstance(stats[col], (float, np.float64)):
#                                     stats[col] = round(stats[col], 3)
#                             peak_dfs.append(stats)
                        
#                         # Load event data
#                         event_df = load_and_process_events(file, phase)
#                         if not event_df.empty:
#                             event_dfs.append(event_df)
                            
#                     except Exception as e:
#                         print(f"Error processing file {file} for phase {phase}: {str(e)}")
#                         continue
                
#                 # Create phase DataFrame
#                 phase_data[phase] = pd.DataFrame(peak_dfs) if peak_dfs else pd.DataFrame()
#                 event_data[phase] = pd.concat(event_dfs) if event_dfs else pd.DataFrame()
                
#             except Exception as e:
#                 print(f"Error processing phase {phase}: {str(e)}")
#                 phase_data[phase] = pd.DataFrame()
#                 event_data[phase] = pd.DataFrame()
        
#         # Process each phase
#         for phase in ['1', '3', '4']:
#             print(f"Phase {phase} stats_data type: {type(phase_results[phase]['stats_data'])}")
#             print(f"Phase {phase} stats_data: {phase_results[phase]['stats_data']}")
#             if phase_results[phase]['stats_data'] is None:
#                 phase_results[phase]['stats_data'] = []
#             elif not isinstance(phase_results[phase]['stats_data'], list):
#                 try:
#                     # Attempt to convert to list if possible
#                     phase_results[phase]['stats_data'] = list(phase_results[phase]['stats_data'])
#                 except:
#                     # If conversion fails, use an empty list
#                     phase_results[phase]['stats_data'] = []
            
#             phase_df = phase_data.get(phase, pd.DataFrame())
            
#             if not phase_df.empty:
#                 try:
#                     # Create comparison graphs
#                     overall_fig, delta_fig, counts_fig = create_comparison_graphs(phase_df, phase)
                    
#                     # Update layouts
#                     for fig in [overall_fig, delta_fig, counts_fig]:
#                         fig.update_layout(
#                             title_x=0.5,
#                             margin=dict(t=50, b=50, l=50, r=50),
#                             plot_bgcolor='white',
#                             paper_bgcolor='white'
#                         )
                    
#                     # Store results for this phase
#                     phase_results[phase].update({
#                         'overall_fig': overall_fig,
#                         'delta_fig': delta_fig,
#                         'counts_fig': counts_fig
#                     })

#                     # Statistical analysis with proper error handling
#                     try:
#                         stats_analysis = perform_statistical_analysis(phase_df, 'overall_avg')
#                         stats_data = []
                        
#                         if isinstance(stats_analysis, pd.DataFrame) and not stats_analysis.empty:
#                             # Preserve the original data processing logic
#                             for _, row in stats_analysis.iterrows():
#                                 stats_entry = {
#                                     "comparison": str(row["comparison"]),
#                                     "mean": float(row["mean"]) if pd.notna(row["mean"]) else None,
#                                     "std": float(row["std"]) if pd.notna(row["std"]) else None,
#                                     "min": float(row["min"]) if pd.notna(row["min"]) else None,
#                                     "max": float(row["max"]) if pd.notna(row["max"]) else None,
#                                     "t_statistic": float(row["t_statistic"]) if pd.notna(row["t_statistic"]) else None,
#                                     "p_value": float(row["p_value"]) if pd.notna(row["p_value"]) else None
#                                 }
#                                 stats_data.append(stats_entry)

#                             # Ensure numeric values are properly converted
#                             for row in stats_data:
#                                 for key, value in row.items():
#                                     if pd.isna(value):
#                                         row[key] = None
#                                     elif isinstance(value, (float, np.float64)):
#                                         row[key] = float(value)
#                                     else:
#                                         row[key] = str(value)
                        
#                         # Ensure we have at least one item in the list, even if empty
#                         if not stats_data:
#                             stats_data = [{}]
                            
#                         phase_results[phase]['stats_data'] = stats_data
#                     except Exception as e:
#                         print(f"Error in statistical analysis for phase {phase}: {str(e)}")
#                         phase_results[phase]['stats_data'] = [{}]  # Always use at least an empty dict in a list
#                 except Exception as e:
#                     print(f"Error generating analysis for phase {phase}: {str(e)}")
#                     # Ensure stats_data is set even in case of overall analysis failure
#                     phase_results[phase]['stats_data'] = [{}]
                    
#             # Process event data
#             try:
#                 event_df = event_data.get(phase, pd.DataFrame())
#                 if not event_df.empty:
#                     # Create event average graph
#                     avg_fig = px.box(
#                         event_df,
#                         x='Event',
#                         y='Avg',
#                         color='number',
#                         title=f'P{phase} Event Averages by Sample',
#                         points='all'
#                     )
#                     avg_fig.update_layout(
#                         xaxis_title='Event',
#                         yaxis_title='Average Value',
#                         showlegend=True,
#                         boxmode='group',
#                         title_x=0.5,
#                         margin=dict(t=50, b=50, l=50, r=50),
#                         plot_bgcolor='white',
#                         paper_bgcolor='white'
#                     )

#                     # Create event delta graph
#                     delta_fig = px.bar(
#                         event_df,
#                         x='Event',
#                         y=['Delta', 'Diff'],
#                         color='number',
#                         barmode='group',
#                         title=f'P{phase} Event Deltas and Differences'
#                     )
#                     delta_fig.update_layout(
#                         xaxis_title='Event',
#                         yaxis_title='Value',
#                         showlegend=True,
#                         title_x=0.5,
#                         margin=dict(t=50, b=50, l=50, r=50),
#                         plot_bgcolor='white',
#                         paper_bgcolor='white'
#                     )

#                     event_graphs[phase].update({
#                         'avg_fig': avg_fig,
#                         'delta_fig': delta_fig
#                     })
#             except Exception as e:
#                 print(f"Error generating event analysis for phase {phase}: {str(e)}")

#         # Return all results in the correct order
#         return [
#             # P1 analysis results
#             phase_results['1']['overall_fig'],
#             phase_results['1']['delta_fig'],
#             phase_results['1']['counts_fig'],
#             phase_results['1']['stats_data'],
#             # P3 analysis results
#             phase_results['3']['overall_fig'],
#             phase_results['3']['delta_fig'],
#             phase_results['3']['counts_fig'],
#             phase_results['3']['stats_data'],
#             # P4 analysis results
#             phase_results['4']['overall_fig'],
#             phase_results['4']['delta_fig'],
#             phase_results['4']['counts_fig'],
#             phase_results['4']['stats_data'],
#             # Event analysis graphs
#             event_graphs['1']['avg_fig'],
#             event_graphs['3']['avg_fig'],
#             event_graphs['4']['avg_fig'],
#             event_graphs['1']['delta_fig'],
#             event_graphs['3']['delta_fig'],
#             event_graphs['4']['delta_fig']
#         ]

#     except Exception as e:
#         print(f"Error in update_phase_analyses: {str(e)}")
#         traceback.print_exc()
#         #return ([empty_fig] * 12) + ([[]] * 3) + ([empty_fig] * 3)
#         return [
#             empty_fig, empty_fig, empty_fig, [],  # P1
#             empty_fig, empty_fig, empty_fig, [],  # P3
#             empty_fig, empty_fig, empty_fig, [],  # P4
#             empty_fig, empty_fig, empty_fig,      # Event averages
#             empty_fig, empty_fig, empty_fig       # Event deltas
#         ] 
    
# @app.callback(
#     [Output('cross-phase-comparison', 'figure'),
#      Output('cross-phase-stats', 'data')],
#     [Input('worksheet-results-selector', 'value'),
#      Input('cross-phase-metric', 'value')]
# )
# def update_cross_phase_analysis(selected_files, metric):
#     """Update cross-phase analysis based on selected files and metric"""
#     if not selected_files or not metric:
#         return go.Figure(), []
    
#     try:
#         # Combine data from all phases
#         all_data = []
#         for phase in ['1', '3', '4']:
#             for file in selected_files:
#                 df = load_peaks_data(file, phase)
#                 if not df.empty:
#                     stats = calculate_phase_statistics(df)
#                     stats['number'], stats['phase'] = parse_filename(file)
#                     all_data.append(stats)
        
#         if not all_data:
#             return go.Figure(), []
        
#         combined_data = pd.DataFrame(all_data)
        
#         # Create comparison figure
#         fig = px.bar(
#             combined_data,
#             x='number',
#             y=metric,
#             color='phase',
#             barmode='group',
#             title=f'Phase Comparison - {metric}',
#             labels={
#                 'number': 'Sample',
#                 metric: metric.replace('_', ' ').title(),
#                 'phase': 'Phase'
#             }
#         )
        
#         fig.update_layout(
#             xaxis_title='Sample',
#             yaxis_title=metric.replace('_', ' ').title(),
#             showlegend=True,
#             legend_title='Phase'
#         )
        
#         # Add error bars
#         for phase in combined_data['phase'].unique():
#             phase_data = combined_data[combined_data['phase'] == phase]
#             std = phase_data[metric].std()
#             fig.add_traces([
#                 go.Bar(
#                     name=f'{phase} Error',
#                     x=phase_data['number'],
#                     y=[std] * len(phase_data),
#                     error_y=dict(type='data', array=[std] * len(phase_data)),
#                     showlegend=False
#                 )
#             ])
        
#         # Statistical analysis
#         stats_results = perform_statistical_analysis(combined_data, metric)
#         stats_results = stats_results.round(3)
        
#         # Convert to list of dictionaries for DataTable
#         stats_data = []
#         if not stats_results.empty:
#             for _, row in stats_results.iterrows():
#                 stats_data.append({
#                     'comparison': str(row['comparison']),
#                     't_statistic': float(row['t_statistic']) if pd.notna(row['t_statistic']) else None,
#                     'p_value': float(row['p_value']) if pd.notna(row['p_value']) else None
#                 })
        
#         return fig, stats_data
        
#     except Exception as e:
#         print(f"Error in update_cross_phase_analysis: {str(e)}")
#         traceback.print_exc()
#         return go.Figure(), []

if __name__ == '__main__':
    print("Starting application...")
    print(f"Projects directory: {PROJECTS_DIRECTORY}")
    
    # Make sure directories exist
    if not os.path.exists(PROJECTS_DIRECTORY):
        print(f"Creating projects directory: {PROJECTS_DIRECTORY}")
        os.makedirs(PROJECTS_DIRECTORY)
    
    # Initialize app
    initialize_app()
    
    # Run server
    app.run(debug=True, port=8050)

    ### End Results Tab ###

