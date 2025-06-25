import pandas as pd
import os
import sys
from datetime import datetime
import argparse
import tkinter as tk
from tkinter import Tk, filedialog, messagebox, TclError
import shutil
import tempfile 

def print_progress(message, is_stderr=False):
    """
    Prints messages that the UI can capture.
    For errors, direct to sys.stderr.
    Includes a special format for progress percentage updates: "PROGRESS: <float_percentage>"
    """
    if message.startswith("PROGRESS:"):
        print(message, flush=True) # Ensure progress updates are sent immediately
    elif is_stderr:
        print(message, file=sys.stderr, flush=True) # Flush stderr immediately too
    else:
        print(message, flush=True) # Flush all regular messages immediately

def prompt_for_file(title_msg="Select File", file_types=None):
    """
    Prompt the user to select a file using a Tkinter dialog.
    Returns the selected path or None if cancelled.
    """
    try:
        root = Tk()
        root.withdraw() # Hide the main Tkinter window
        file_path = filedialog.askopenfilename(
            title=title_msg,
            filetypes=file_types if file_types else [("All files", "*.*")]
        )
        root.destroy() # Destroy the hidden root window
        return file_path
    except TclError:
        print_progress("Warning: Tkinter GUI not available. Cannot prompt for file. Script might require direct arguments.", is_stderr=True)
        return None
    except Exception as e:
        print_progress(f"Error opening file dialog: {e}", is_stderr=True)
        return None

def _get_skus_from_input_file(sku_file_path):
    """
    Reads SKUs from either an Excel spreadsheet file (.xlsx) or a plain text file (.txt).
    Returns a list of unique SKUs.
    """
    skus = []
    if not sku_file_path or not os.path.exists(sku_file_path):
        print_progress(f"Error: SKU list file not found or path is empty: {sku_file_path}", is_stderr=True)
        return None

    file_extension = os.path.splitext(sku_file_path)[1].lower()

    try:
        if file_extension in ('.xlsx', '.xls'):
            print_progress(f"Reading SKUs from Excel spreadsheet: {sku_file_path}")
            # Assuming SKUs are in the first column of the Excel file
            df = pd.read_excel(sku_file_path, usecols=[0], dtype=str)
            skus = df.iloc[:, 0].dropna().astype(str).str.strip().tolist()
        elif file_extension == '.txt':
            print_progress(f"Reading SKUs from text file: {sku_file_path}")
            with open(sku_file_path, 'r', encoding='utf-8') as f:
                skus = [line.strip() for line in f if line.strip()]
        else:
            print_progress(f"Error: Unsupported SKU file type: {file_extension}. Please provide an .xlsx or .txt file.", is_stderr=True)
            return None
    except Exception as e:
        print_progress(f"Error reading SKU list from '{sku_file_path}': {e}", is_stderr=True)
        return None
    
    if not skus:
        print_progress("No valid SKUs found in the provided input file.", is_stderr=True)
        return None
    
    return list(pd.Series(skus).unique()) # Return unique SKUs

def main():
    # Define default STEP export paths based on operating system
    if sys.platform == "darwin":  # mac
        default_step_export_one = "/Volumes/Work_In_Progress/2024_DAM_DI_Projects/_STEP Export/STEP_Export_one.xlsx"
        default_step_export_two = "/Volumes/Work_In_Progress/2024_DAM_DI_Projects/_STEP Export/STEP_Export_two.xlsx"
    elif os.name == "nt":  # Windows
        default_step_export_one = r"W:\2024_DAM_DI_Projects\_STEP Export\STEP_Export_one.xlsx"
        default_step_export_two = r"W:\2024_DAM_DI_Projects\_STEP Export\STEP_Export_two.xlsx"
    else:
        print_progress(f"Warning: Unsupported OS '{sys.platform}'. Defaulting to generic paths, please verify manually.", is_stderr=True)
        default_step_export_one = "STEP_Export_one.xlsx"
        default_step_export_two = "STEP_Export_two.xlsx"

    parser = argparse.ArgumentParser(description="Retrieve measurement data from STEP exports for given SKUs.")
    parser.add_argument("--sku_list_file", help="Path to a file (Excel .xlsx or Text .txt) containing SKUs in the first column.", default=None)
    parser.add_argument("--step_one_file", help="Path to the first STEP export Excel file.", default=None)
    parser.add_argument("--step_two_file", help="Path to the second STEP export Excel file.", default=None)
    parser.add_argument("--output_folder", help="Optional: Folder to save the output Excel file.", default=None)
    
    args = parser.parse_args()

    print_progress("PROGRESS: 0.0") # Initial progress

    # --- Determine SKU Input Source ---
    sku_list_file_path = args.sku_list_file
    if not sku_list_file_path:
        print_progress("No SKU list file provided. Prompting user to select a file for SKUs...")
        sku_list_file_path = prompt_for_file(
            "Select SKU List File (.xlsx or .txt)", 
            file_types=[("Excel files", "*.xlsx *.xls"), ("Text files", "*.txt"), ("All files", "*.*")]
        )
        if not sku_list_file_path:
            print_progress("SKU list selection cancelled. Exiting.", is_stderr=True)
            sys.exit(0) # Exit gracefully on cancel

    skus_to_process = _get_skus_from_input_file(sku_list_file_path)
    
    if skus_to_process is None or not skus_to_process:
        print_progress("No valid SKUs found in the provided input. Exiting.", is_stderr=True)
        sys.exit(1)

    print_progress("PROGRESS: 10.0") # Progress after reading SKUs
    print_progress(f"Using SKUs from input data (count: {len(skus_to_process)})")


    # --- Determine STEP export file paths ---
    # Prioritize command-line arg, then check if default exists, then prompt
    final_step_export_one = args.step_one_file
    if not final_step_export_one: # If not provided as command-line arg
        if os.path.exists(default_step_export_one):
            final_step_export_one = default_step_export_one
            print_progress(f"Using default STEP Export One from: {final_step_export_one}")
        else:
            print_progress(f"No STEP Export One file provided and default not found. Attempting interactive prompt...")
            selected_path = prompt_for_file(
                "Select First STEP Export Excel File",
                file_types=[("Excel files", "*.xlsx *.xls")]
            )
            if not selected_path:
                print_progress("STEP Export One selection cancelled. Script may proceed with missing data.", is_stderr=True)
                final_step_export_one = None # Explicitly set to None if cancelled and no default
            else:
                final_step_export_one = selected_path
                print_progress(f"Using user-selected STEP Export One from: {final_step_export_one}")


    final_step_export_two = args.step_two_file
    if not final_step_export_two: # If not provided as command-line arg
        if os.path.exists(default_step_export_two):
            final_step_export_two = default_step_export_two
            print_progress(f"Using default STEP Export Two from: {final_step_export_two}")
        else:
            print_progress(f"No STEP Export Two file provided and default not found. Attempting interactive prompt...")
            selected_path = prompt_for_file(
                "Select Second STEP Export Excel File",
                file_types=[("Excel files", "*.xlsx *.xls")]
            )
            if not selected_path:
                print_progress("STEP Export Two selection cancelled. Script may proceed with missing data.", is_stderr=True)
                final_step_export_two = None # Explicitly set to None if cancelled and no default
            else:
                final_step_export_two = selected_path
                print_progress(f"Using user-selected STEP Export Two from: {final_step_export_two}")

    # Log the final paths being used, even if they are defaults or None
    print_progress(f"Final STEP Export One path: {final_step_export_one}")
    print_progress(f"Final STEP Export Two path: {final_step_export_two}")


    print_progress("PROGRESS: 20.0") # Progress after determining STEP files

    # Read and combine the STEP export files
    combined_df = pd.DataFrame()
    found_any_step_data = False
    for step_file in [final_step_export_one, final_step_export_two]: # Use final_step_export_one/two
        if not step_file: # Skip if the path is None (e.g., if dialog was cancelled and no default was found)
            continue
        if not os.path.exists(step_file):
            print_progress(f"Warning: STEP export file not found at '{step_file}'. Skipping this file.", is_stderr=True)
            continue
        try:
            df = pd.read_excel(step_file)
            combined_df = pd.concat([combined_df, df], ignore_index=True)
            found_any_step_data = True
        except Exception as e:
            print_progress(f"Warning: Error reading STEP export file '{step_file}': {e}. Skipping this file.", is_stderr=True)

    if combined_df.empty or not found_any_step_data:
        print_progress("Error: No valid data found in any of the provided or selected STEP export files. Cannot retrieve measurements. Exiting.", is_stderr=True)
        print_progress("PROGRESS: 0.0", is_stderr=True) # Reset progress on error
        sys.exit(1)

    print_progress("PROGRESS: 40.0") # Progress after loading STEP data

    # Normalize SKU column in the combined data
    if 'SKU' in combined_df.columns:
        combined_df['SKU'] = combined_df['SKU'].astype(str).str.strip()
        combined_dict = combined_df.set_index('SKU').to_dict(orient='index')
    else:
        print_progress("Error: 'SKU' column not found in combined STEP export data. Please ensure your STEP exports have an 'SKU' column. Exiting.", is_stderr=True)
        print_progress("PROGRESS: 0.0", is_stderr=True) # Reset progress on error
        sys.exit(1)

    headers = [
        "SKU", "Name", "Vendor Code", "Path", "Collection", "Color",
        "Additional Dimensions", "Length", "Width", "Height",
        "Inside Arm to Arm", "Seat Depth", "Seat Height", "Extended Length",
        "Extended Width", "Drawer Length", "Drawer Height", "Arm Height"
    ]

    output_data = []
    print_progress(f"Processing {len(skus_to_process)} SKUs to retrieve measurements...")
    total_skus = len(skus_to_process)
    for i, sku in enumerate(skus_to_process):
        row = {h: '' for h in headers}
        row['SKU'] = sku
        if sku in combined_dict:
            for h in headers[1:]:
                row[h] = combined_dict[sku].get(h, '')
            row['Status'] = 'Success'
        else:
            row['Status'] = 'SKU not found in STEP exports'
        output_data.append(row)
        
        # --- Send Progress Update to UI ---
        progress_percentage = (i + 1) / total_skus * 100
        print_progress(f"PROGRESS:{progress_percentage:.2f}")

    output_df = pd.DataFrame(output_data, columns=headers + ['Status'])
    
    print_progress("PROGRESS: 90.0") # Progress after processing SKUs

    # Determine final output folder
    final_output_folder = args.output_folder
    if not final_output_folder:
        # If output folder is not specified by argument, use Downloads folder.
        final_output_folder = os.path.join(os.path.expanduser("~"), "Downloads")
        print_progress(f"No output folder specified. Defaulting to Downloads folder: {final_output_folder}")

    os.makedirs(final_output_folder, exist_ok=True) 

    output_filename = f"measurementsFromSTEP_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    output_path = os.path.join(final_output_folder, output_filename)
    output_df.to_excel(output_path, index=False)

    print_progress(f"Output file saved to: {output_path}")
    print_progress("PROGRESS: 100.0")

    # The cleanup of any temporary SKU file created by the GUI is handled by the GUI.
    # This script should not delete `args.sku_list_file` as it might be a user's original file.

if __name__ == "__main__":
    main()
