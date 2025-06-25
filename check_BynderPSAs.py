import pandas as pd
import os
import platform
from datetime import datetime
import sys
import argparse
import tkinter as tk
from tkinter import filedialog, messagebox, TclError # Added TclError

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
        root = tk.Tk()
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
    # Define hardcoded Bynder Asset Report Path
    if platform.system() == 'Windows':
        hardcoded_sku_report_path = r"W:\2024_DAM_DI_Projects\_STEP Export\BynderProductSiteAssetsExport\sku_asset_report.csv"
    else: # macOS/Linux
        hardcoded_sku_report_path = r"/Volumes/Work_In_Progress/2024_DAM_DI_Projects/_STEP Export/BynderProductSiteAssetsExport/sku_asset_report.csv"
        
    # Set up argument parser
    parser = argparse.ArgumentParser(description="Check Bynder PSAs based on SKU list and asset report.")
    parser.add_argument("--sku_file", help="Path to a file (Excel .xlsx or Text .txt) containing SKUs.", default=None)

    args = parser.parse_args()

    print_progress("PROGRESS: 0.0") # Initial progress

    # --- Determine SKU List Input ---
    sku_input_data = None
    if args.sku_file:
        # Running from UI or via command line with --sku_file
        sku_input_data = _get_skus_from_input_file(args.sku_file)
    else:
        # Running standalone, prompt user for SKU list file
        print_progress("No SKU list file provided. Prompting user to select a file for SKUs...")
        sku_list_path = prompt_for_file("Select SKU List File (.xlsx or .txt)", 
                                        file_types=[("Excel files", "*.xlsx *.xls"), ("Text files", "*.txt"), ("All files", "*.*")])
        if not sku_list_path:
            print_progress("SKU list selection cancelled. Exiting.", is_stderr=True)
            sys.exit(0) # Exit gracefully on cancel
        sku_input_data = _get_skus_from_input_file(sku_list_path)
    
    if sku_input_data is None: # Covers cases where file not found, empty, or read error
        print_progress("PROGRESS: 0.0", is_stderr=True) # Reset progress on error
        sys.exit(1) # Exit with error code if SKU input parsing failed

    print_progress("PROGRESS: 10.0") # Progress after reading SKUs

    # --- SKU Report Path (Always Hardcoded) ---
    sku_report_path = hardcoded_sku_report_path
    print_progress(f"Using hardcoded Bynder Asset Report from: {sku_report_path}")

    print_progress(f"Using SKUs from input data (count: {len(sku_input_data)})")

    # Generate date and timestamp for output filename
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    # Output path (always to Downloads for simplicity)
    output_path = os.path.expanduser(f"~/Downloads/check_BynderPSAs_checked_{timestamp}.xlsx")

    # Convert the list of SKUs to a DataFrame for merging
    sku_list_df = pd.DataFrame({'SKU': sku_input_data})

    # Load the SKU report CSV file and ensure SKU column is treated as string
    try:
        print_progress("Loading Bynder Asset Report...")
        sku_report_df = pd.read_csv(sku_report_path, dtype=str)
    except FileNotFoundError:
        print_progress(f"Error: Bynder Asset Report file not found at {sku_report_path}. Please check the path and file name.", is_stderr=True)
        print_progress("PROGRESS: 0.0", is_stderr=True) # Reset progress on error
        sys.exit(1)
    except Exception as e:
        print_progress(f"Error loading Bynder Asset Report from '{sku_report_path}': {str(e)}", is_stderr=True)
        print_progress("PROGRESS: 0.0", is_stderr=True) # Reset progress on error
        sys.exit(1)
    
    print_progress("PROGRESS: 40.0") # Progress after loading report

    # Identify the first column (assumed to contain SKUs) in sku_report
    if sku_report_df.empty:
        print_progress("Warning: Bynder Asset Report file is empty. Cannot check for assets.", is_stderr=True)
        output_df = sku_list_df.copy()
        output_df['Status'] = 'No report data'
        output_df.to_excel(output_path, index=False)
        print_progress(f"Report with 'No report data' status saved to {output_path}")
        print_progress("PROGRESS: 100.00") # Final progress update even on warning path
        sys.exit(0)

    # The column from the Bynder report that contains the SKUs (the merge key)
    sku_column_name_report = sku_report_df.columns[0]
    
    # We want columns B, C, D, E, F from the *original* report.
    # In pandas (0-indexed), these are columns at index 1, 2, 3, 4, 5.
    desired_columns_indices = [1, 2, 3, 4, 5] 
    
    # Check if the report has enough columns for B-F
    if len(sku_report_df.columns) <= max(desired_columns_indices):
        print_progress("Warning: Bynder Asset Report does not contain enough columns (expected at least B-F). Including all available columns from index 1 onwards.", is_stderr=True)
        # If not enough columns, take all columns from index 1 to the end
        columns_to_select_for_merge = sku_report_df.columns[1:].tolist()
    else:
        # Select the desired columns (B-F) by their names using the indices
        columns_to_select_for_merge = sku_report_df.columns[desired_columns_indices].tolist()

    # Filter the SKU report to include rows for SKUs found in sku_list_df, and only select the desired columns (B-F)
    print_progress("Filtering Bynder asset report based on provided SKUs and selecting relevant columns (B-F)...")
    # We need to ensure sku_column_name_report is included temporarily for the .isin() filter
    # and then in the actual selection, we'll exclude it if it's not B-F.
    temp_columns_for_filter = [sku_column_name_report] + columns_to_select_for_merge
    temp_filtered_report_df = sku_report_df[sku_report_df[sku_column_name_report].isin(sku_list_df['SKU'])][temp_columns_for_filter]

    print_progress("PROGRESS: 70.0") # Progress after filtering

    # Merge the SKU list with the filtered report.
    # We merge on 'SKU' from sku_list_df and sku_column_name_report from temp_filtered_report_df.
    # After the merge, the 'SKU' column from sku_list_df will be the primary SKU column.
    print_progress("Merging SKU list with filtered report data...")
    output_df = sku_list_df.merge(temp_filtered_report_df, how='left', 
                                  left_on='SKU', right_on=sku_column_name_report)

    # Drop the duplicate SKU column from the report (it's the one we merged *from*)
    # and ensure the SKU column from the original list is preserved as 'SKU'.
    if sku_column_name_report in output_df.columns:
        if sku_column_name_report != 'SKU': 
            output_df = output_df.drop(columns=[sku_column_name_report])
        if 'SKU_x' in output_df.columns and 'SKU_y' in output_df.columns:
            output_df = output_df.rename(columns={'SKU_x': 'SKU'}).drop(columns=['SKU_y'])
        elif 'SKU_x' in output_df.columns:
             output_df = output_df.rename(columns={'SKU_x': 'SKU'})

    # Reorder columns to ensure 'SKU' is first, then the selected report columns
    final_columns = ['SKU'] + [col for col in output_df.columns if col != 'SKU']
    output_df = output_df[final_columns]
    
    print_progress("PROGRESS: 90.0") # Progress after merging and reordering

    # Save the output to an Excel file in the Downloads folder with timestamped filename
    print_progress(f"Saving output to Excel file: {output_path}")
    output_df.to_excel(output_path, index=False)

    print_progress(f"File saved to {output_path}")
    print_progress("PROGRESS: 100.0") # Final progress update

    # The cleanup of the temporary SKU file is handled by the GUI.
    # No need for this script to delete `args.sku_file` as it might be a user's original file.

if __name__ == "__main__":
    main()
