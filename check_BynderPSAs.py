import pandas as pd
import os
import platform
from datetime import datetime
import sys
import argparse
import tkinter as tk
from tkinter import filedialog, messagebox

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
    

def prompt_for_excel_file(title_msg="Select Excel Spreadsheet"):
    """Prompt the user to select an Excel file and return its path."""
    # Hide the main Tkinter window for the file dialog
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title=title_msg,
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    root.destroy() # Destroy the hidden root window
    return file_path

def _get_skus_from_input_data(sku_file_path, sku_list_str):
    """
    Reads SKUs from either a spreadsheet file path or a comma-separated string.
    Returns a list of unique SKUs.
    """
    skus = []
    if sku_file_path:
        print_progress(f"Reading SKUs from spreadsheet: {sku_file_path}")
        try:
            # Assuming SKUs are in the first column of the Excel file
            df = pd.read_excel(sku_file_path, usecols=[0], dtype=str)
            skus = df.iloc[:, 0].dropna().astype(str).str.strip().tolist()
        except FileNotFoundError:
            print_progress(f"Error: SKU list file not found at {sku_file_path}.", is_stderr=True)
            return None
        except Exception as e:
            print_progress(f"Error reading SKU list from '{sku_file_path}': {e}", is_stderr=True)
            return None
    elif sku_list_str:
        print_progress(f"Processing SKUs from text input...")
        # Split by comma or whitespace, then filter out empty strings
        skus = [s.strip() for s in sku_list_str.replace(',', ' ').split() if s.strip()]
    
    if not skus:
        print_progress("No valid SKUs found in the provided input.", is_stderr=True)
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
    parser.add_argument("--sku_file", help="Path to the Excel file containing SKUs.", default=None)
    parser.add_argument("--sku_list", help="Comma-separated string of SKUs (from text input).", default=None)

    args = parser.parse_args()

    print_progress("PROGRESS: 0.0") # Initial progress

    # --- Determine SKU List Input ---
    sku_input_data = None
    if args.sku_file:
        # Running from UI with spreadsheet input
        sku_input_data = _get_skus_from_input_data(args.sku_file, None)
    elif args.sku_list:
        # Running from UI with text box input
        sku_input_data = _get_skus_from_input_data(None, args.sku_list)
    else:
        # Running standalone, prompt user for SKU list file
        print_progress("No SKU list provided. Prompting user to select an Excel file for SKUs...")
        sku_list_path = prompt_for_excel_file("Select SKU List Excel Spreadsheet")
        if not sku_list_path:
            print_progress("SKU list selection cancelled. Exiting.", is_stderr=True)
            sys.exit(1)
        sku_input_data = _get_skus_from_input_data(sku_list_path, None)
    
    if sku_input_data is None:
        print_progress("PROGRESS: 0.0", is_stderr=True) # Reset progress on error
        sys.exit(1) # Exit if SKU input parsing failed (either from arguments or standalone prompt)

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
        # Only drop if the column name exists and it's not the primary 'SKU' column we want to keep
        if sku_column_name_report != 'SKU': # This check prevents dropping the primary SKU if names are identical
            output_df = output_df.drop(columns=[sku_column_name_report])
        # If sku_column_name_report was 'SKU' and it was duplicated during merge, pandas might append '_x', '_y'.
        # We ensure the original 'SKU' (from sku_list_df) remains.
        if 'SKU_x' in output_df.columns and 'SKU_y' in output_df.columns:
            output_df = output_df.rename(columns={'SKU_x': 'SKU'}).drop(columns=['SKU_y'])
        elif 'SKU_x' in output_df.columns: # If only _x, implies original was 'SKU', report had same name
             output_df = output_df.rename(columns={'SKU_x': 'SKU'})
        # In this specific case, since we explicitly merged on sku_column_name_report and kept 'SKU' as left_on,
        # pandas automatically handles the duplicate column if `sku_column_name_report` and 'SKU' were the same.
        # The main goal here is to ensure the final output starts with 'SKU' column (from the input list)
        # followed by the columns B-F from the report.

    # Reorder columns to ensure 'SKU' is first, then the selected report columns
    final_columns = ['SKU'] + [col for col in output_df.columns if col != 'SKU']
    output_df = output_df[final_columns]
    
    print_progress("PROGRESS: 90.0") # Progress after merging and reordering

    # Save the output to an Excel file in the Downloads folder with timestamped filename
    print_progress(f"Saving output to Excel file: {output_path}")
    output_df.to_excel(output_path, index=False)

    print_progress(f"File saved to {output_path}")
    print_progress("PROGRESS: 100.0") # Final progress update

if __name__ == "__main__":
    main()
