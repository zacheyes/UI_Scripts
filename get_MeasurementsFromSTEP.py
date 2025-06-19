import pandas as pd
import os
import sys
from datetime import datetime
import argparse
from tkinter import Tk, filedialog
import tempfile
import shutil

def main():
    # Define default STEP export paths based on operating system
    if sys.platform == "darwin":  # mac
        default_step_export_one = "/Volumes/Work_In_Progress/2024_DAM_DI_Projects/_STEP Export/STEP_Export_one.xlsx"
        default_step_export_two = "/Volumes/Work_In_Progress/2024_DAM_DI_Projects/_STEP Export/STEP_Export_two.xlsx"
    elif os.name == "nt":  # Windows
        default_step_export_one = r"W:\2024_DAM_DI_Projects\_STEP Export\STEP_Export_one.xlsx"
        default_step_export_two = r"W:\2024_DAM_DI_Projects\_STEP Export\STEP_Export_two.xlsx"
    else:
        print(f"Warning: Unsupported OS '{sys.platform}'. Defaulting to generic paths, please verify manually.", file=sys.stderr)
        default_step_export_one = "STEP_Export_one.xlsx"
        default_step_export_two = "STEP_Export_two.xlsx"

    parser = argparse.ArgumentParser(description="Retrieve measurement data from STEP exports for given SKUs.")
    parser.add_argument("--sku_list_file", help="Path to the Excel file containing SKUs in the first column.")
    parser.add_argument("--sku_list", help="Comma-separated string of SKUs directly from UI textbox.")
    parser.add_argument("--step_one_file", help="Path to the first STEP export Excel file.", default=None)
    parser.add_argument("--step_two_file", help="Path to the second STEP export Excel file.", default=None)
    parser.add_argument("--output_folder", help="Optional: Folder to save the output Excel file.")
    
    args = parser.parse_args()

    skus_to_process = []
    input_is_temp_file = False
    temp_sku_xlsx_path = None

    # --- Determine SKU Input Source ---
    input_sku_source = None
    
    if args.sku_list_file:
        print(f"Script: Reading SKUs from file: {args.sku_list_file}")
        input_sku_source = args.sku_list_file
        try:
            source_df = pd.read_excel(input_sku_source, dtype=str)
            skus_to_process = source_df.iloc[:, 0].dropna().astype(str).str.strip().tolist()
        except FileNotFoundError:
            print(f"Error: SKU list file not found at '{input_sku_source}'. Exiting.", file=sys.stderr)
            sys.exit(1)
        except Exception as e:
            print(f"Error reading SKU list from '{input_sku_source}': {e}. Exiting.", file=sys.stderr)
            sys.exit(1)
    elif args.sku_list:
        print("Script: Processing SKUs from command line string (UI textbox input).")
        skus_to_process = [sku.strip() for sku in args.sku_list.split(',') if sku.strip()]
        
        try:
            temp_dir = tempfile.mkdtemp(prefix="measurements_skus_temp_")
            temp_sku_xlsx_path = os.path.join(temp_dir, f"temp_skus_{os.getpid()}.xlsx")
            pd.DataFrame(skus_to_process, columns=['SKU']).to_excel(temp_sku_xlsx_path, index=False, header=False)
            input_sku_source = temp_sku_xlsx_path
            input_is_temp_file = True
        except Exception as e:
            print(f"Error creating temporary SKU file: {e}. Exiting.", file=sys.stderr)
            sys.exit(1)
    else:
        root = Tk()
        root.withdraw()
        print("Script: No SKU input provided via arguments. Prompting interactively for file...")
        selected_sku_file = filedialog.askopenfilename(
            title="Select spreadsheet with your SKUs in column A",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not selected_sku_file:
            print("No SKU file selected. Exiting.", file=sys.stderr)
            if root: root.destroy()
            sys.exit(0)
        
        input_sku_source = selected_sku_file
        print(f"Script: SKU file selected: {input_sku_source}")
        try:
            source_df = pd.read_excel(input_sku_source, dtype=str)
            skus_to_process = source_df.iloc[:, 0].dropna().astype(str).str.strip().tolist()
        except Exception as e:
            print(f"Error reading SKU list from '{input_sku_source}': {e}. Exiting.", file=sys.stderr)
            if root: root.destroy()
            sys.exit(1)
    
    if not skus_to_process:
        print("No valid SKUs found in the input. Exiting.", file=sys.stderr)
        if 'root' in locals() and root: root.destroy()
        sys.exit(0)

    # Determine STEP export file paths (from arguments or defaults)
    step_export_one = args.step_one_file if args.step_one_file else default_step_export_one
    step_export_two = args.step_two_file if args.step_two_file else default_step_export_two

    print(f"Script: Using STEP Export One from: {step_export_one}")
    print(f"Script: Using STEP Export Two from: {step_export_two}")

    # Read and combine the STEP export files
    combined_df = pd.DataFrame()
    for step_file in [step_export_one, step_export_two]:
        if not os.path.exists(step_file):
            print(f"Warning: STEP export file not found at '{step_file}'. Skipping this file.", file=sys.stderr)
            continue
        try:
            df = pd.read_excel(step_file)
            combined_df = pd.concat([combined_df, df], ignore_index=True)
        except Exception as e:
            print(f"Warning: Error reading STEP export file '{step_file}': {e}. Skipping this file.", file=sys.stderr)

    if combined_df.empty:
        print("Error: No valid data found in any of the provided STEP export files. Cannot retrieve measurements. Exiting.", file=sys.stderr)
        if 'root' in locals() and root: root.destroy()
        if temp_sku_xlsx_path and os.path.exists(os.path.dirname(temp_sku_xlsx_path)):
            try:
                shutil.rmtree(os.path.dirname(temp_sku_xlsx_path))
            except Exception as e:
                print(f"Script: Warning: Could not clean up temporary SKU directory: {e}", file=sys.stderr)
        sys.exit(1)

    # Normalize SKU column in the combined data
    if 'SKU' in combined_df.columns:
        combined_df['SKU'] = combined_df['SKU'].astype(str).str.strip()
        combined_dict = combined_df.set_index('SKU').to_dict(orient='index')
    else:
        print("Error: 'SKU' column not found in combined STEP export data. Please ensure your STEP exports have an 'SKU' column. Exiting.", file=sys.stderr)
        if 'root' in locals() and root: root.destroy()
        if temp_sku_xlsx_path and os.path.exists(os.path.dirname(temp_sku_xlsx_path)):
            try:
                shutil.rmtree(os.path.dirname(temp_sku_xlsx_path))
            except Exception as e:
                print(f"Script: Warning: Could not clean up temporary SKU directory: {e}", file=sys.stderr)
        sys.exit(1)

    headers = [
        "SKU", "Name", "Vendor Code", "Path", "Collection", "Color",
        "Additional Dimensions", "Length", "Width", "Height",
        "Inside Arm to Arm", "Seat Depth", "Seat Height", "Extended Length",
        "Extended Width", "Drawer Length", "Drawer Height", "Arm Height"
    ]

    output_data = []
    print(f"Script: Processing {len(skus_to_process)} SKUs to retrieve measurements...")
    total_skus = len(skus_to_process)
    for i, sku in enumerate(skus_to_process): # Use enumerate for progress tracking
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
        print(f"PROGRESS:{progress_percentage:.2f}", flush=True) # CRITICAL: print and flush

    output_df = pd.DataFrame(output_data, columns=headers + ['Status'])
    
    final_output_folder = args.output_folder
    if not final_output_folder:
        if input_is_temp_file:
            final_output_folder = os.path.join(os.path.expanduser("~"), "Downloads")
            print(f"Script: No output folder specified. Defaulting to Downloads folder for textbox input: {final_output_folder}")
        else:
            if input_sku_source and os.path.exists(input_sku_source):
                final_output_folder = os.path.dirname(input_sku_source)
                print(f"Script: No output folder specified. Defaulting to input SKU file's directory: {final_output_folder}")
            else:
                 final_output_folder = os.path.join(os.path.expanduser("~"), "Downloads")
                 print(f"Script: Could not determine input file's directory. Defaulting to Downloads: {final_output_folder}")

    os.makedirs(final_output_folder, exist_ok=True) 

    output_filename = f"measurementsFromSTEP_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    output_path = os.path.join(final_output_folder, output_filename)
    output_df.to_excel(output_path, index=False)

    print(f"Script: Output file saved to: {output_path}")

    if 'root' in locals() and root:
        root.destroy()
    
    if temp_sku_xlsx_path and os.path.exists(os.path.dirname(temp_sku_xlsx_path)):
        try:
            shutil.rmtree(os.path.dirname(temp_sku_xlsx_path))
            print(f"Script: Removed temporary directory: {os.path.dirname(temp_sku_xlsx_path)}")
        except Exception as e:
            print(f"Script: Warning: Could not remove temporary SKU directory {os.path.dirname(temp_sku_xlsx_path)}: {e}", file=sys.stderr)


if __name__ == "__main__":
    main()
