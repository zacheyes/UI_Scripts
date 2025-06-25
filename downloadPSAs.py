import os
import pandas as pd
import requests
import sys
import argparse
from datetime import datetime
from tkinter import Tk, filedialog, messagebox, TclError # <--- Added TclError here
import shutil 

def print_progress(message, is_stderr=False):
    """
    Prints messages that the UI can capture.
    For errors, direct to sys.stderr.
    Includes a special format for progress percentage updates: "PROGRESS: <float_percentage>"
    """
    if message.startswith("PROGRESS:"):
        print(message, flush=True) 
    elif is_stderr:
        print(message, file=sys.stderr, flush=True) 
    else:
        print(message, flush=True) 

def download_image(url, output_path):
    """Downloads an image from the given URL to the specified output path."""
    try:
        response = requests.get(url, stream=True, timeout=10) 
        response.raise_for_status() 
        with open(output_path, 'wb') as file:
            for chunk in response.iter_content(chunk_size=8192):
                file.write(chunk)
        return "Downloaded"
    except requests.exceptions.HTTPError as e:
        if e.response.status_code == 404:
            return "Not Found (404)"
        else:
            return f"HTTP Error {e.response.status_code}"
    except requests.exceptions.ConnectionError:
        return "Connection Error"
    except requests.exceptions.Timeout:
        return "Timeout Error"
    except requests.exceptions.RequestException as e:
        return f"Network Error: {str(e)}"
    except Exception as e:
        return f"Unexpected Error: {str(e)}"

def prompt_for_file_or_folder(is_file=True, title_msg="Select File", file_types=None):
    """
    Prompt the user to select a file or folder using a Tkinter dialog.
    Returns the selected path or None if cancelled.
    """
    try:
        root = Tk()
        root.withdraw() 
        if is_file:
            path = filedialog.askopenfilename(title=title_msg, filetypes=file_types)
        else:
            path = filedialog.askdirectory(title=title_msg)
        root.destroy() 
        return path
    except TclError: # <--- Changed from tk.TclError to TclError
        print_progress("Warning: Tkinter GUI not available. Cannot prompt for file/folder. Script might require direct arguments.", is_stderr=True)
        return None
    except Exception as e:
        print_progress(f"Error opening file/folder dialog: {e}", is_stderr=True)
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
    
    return list(pd.Series(skus).unique()) 


def main():
    parser = argparse.ArgumentParser(description="Download PSA images from Bynder.")
    parser.add_argument("--sku_file", help="Path to a file (Excel .xlsx or Text .txt) containing SKUs.", default=None)
    parser.add_argument("--output_folder", help="Path to the folder where images will be downloaded.", default=None)
    parser.add_argument("--image_types", help="Comma-separated list of image types to download (e.g., grid,100,dimension).", default=None)
    
    args = parser.parse_args()

    print_progress("PROGRESS: 0.0")

    # --- Determine SKU List Input ---
    skus_to_process = []
    if args.sku_file:
        print_progress(f"Script: Reading SKUs from file: {args.sku_file}")
        skus_to_process = _get_skus_from_input_file(args.sku_file)
    else:
        print_progress("Script: No SKU file provided. Attempting interactive prompt for file...")
        sku_list_path = prompt_for_file_or_folder(is_file=True, title_msg="Select SKU List File (.xlsx or .txt)", 
                                                 file_types=[("Excel files", "*.xlsx *.xls"), ("Text files", "*.txt"), ("All files", "*.*")])
        if not sku_list_path:
            print_progress("No SKU file selected. Exiting.", is_stderr=True)
            sys.exit(0) 
        skus_to_process = _get_skus_from_input_file(sku_list_path)
    
    if skus_to_process is None or not skus_to_process:
        print_progress("No valid SKUs found in the provided input. Exiting.", is_stderr=True)
        sys.exit(1)

    print_progress("PROGRESS: 10.0")

    # --- Output folder determination logic ---
    output_folder = args.output_folder
    if not output_folder:
        print_progress("Script: No output folder provided. Attempting interactive prompt for folder...")
        output_folder = prompt_for_file_or_folder(is_file=False, title_msg="Select Folder for Downloads")
        if not output_folder:
            print_progress("No output folder selected. Exiting.", file=sys.stderr)
            sys.exit(0) 
        print_progress(f"Script: Output folder selected: {output_folder}")
    
    os.makedirs(output_folder, exist_ok=True)
    print_progress(f"Images will be downloaded to: {output_folder}")

    print_progress("PROGRESS: 20.0")

    # --- Image types determination logic ---
    image_types_options = ["grid", "100", "200", "300", "400", "500", "dimension", "swatch"]
    image_types_flags = {k: False for k in image_types_options}

    if args.image_types:
        types_from_arg = [t.strip().lower() for t in args.image_types.split(',') if t.strip()]
        print_progress(f"Script: Image types requested via argument: {types_from_arg}")
        if not types_from_arg:
             print_progress("Script: No specific image types specified via argument. Defaulting to 'grid'.")
             image_types_flags["grid"] = True
        else:
            for img_type in types_from_arg:
                if img_type in image_types_flags:
                    image_types_flags[img_type] = True
                else:
                    print_progress(f"Warning: Unknown image type '{img_type}' requested, skipping.", is_stderr=True)
    else:
        print_progress("Script: No image types specified via argument. Attempting interactive prompt...")
        
        if sys.stdin.isatty(): 
            for img_type_key in image_types_options:
                response = input(f"Do you want the SKU_{img_type_key} image? (Y/N): ").strip().lower()
                if response == 'y':
                    image_types_flags[img_type_key] = True
            if not any(image_types_flags.values()):
                print_progress("No image types selected interactively. Defaulting to 'grid'.")
                image_types_flags["grid"] = True
        else:
            print_progress("Script: Running in non-interactive mode and no image types specified. Defaulting to 'grid' image.")
            image_types_flags["grid"] = True


    base_url = "https://www.bynder.raymourflanigan.com/match/Product_SKU_Position/"
    report_rows = []
    headers = ["SKU"] + [f"SKU_{t}_Status" for t in image_types_options]

    print_progress(f"Script: Starting download for {len(skus_to_process)} SKUs...")
    total_skus = len(skus_to_process)
    
    downloaded_count = 0
    not_found_count = 0
    error_count = 0

    for i, sku in enumerate(skus_to_process):
        row = {"SKU": sku}
        for img_type in image_types_options:
            column_name = f"SKU_{img_type}_Status"
            if image_types_flags[img_type]:
                image_url = f"{base_url}{sku}_{img_type}/"
                output_path = os.path.join(output_folder, f"{sku}_{img_type}.jpg")
                status = download_image(image_url, output_path)
                row[column_name] = status
                if status == "Downloaded":
                    downloaded_count += 1
                elif "Not Found" in status:
                    not_found_count += 1
                else:
                    error_count += 1
            else:
                row[column_name] = "Not Requested"
        report_rows.append(row)
        
        progress_percentage = (i + 1) / total_skus * 100
        print_progress(f"PROGRESS:{progress_percentage:.2f}")

    report_df = pd.DataFrame(report_rows, columns=headers)
    report_file = os.path.join(output_folder, f"download_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
    report_df.to_csv(report_file, index=False)

    print_progress(f"Script: Download process complete. Report saved to {report_file}")
    print_progress(f"Summary: Downloaded {downloaded_count} images, {not_found_count} not found, {error_count} errors.")
    print_progress("PROGRESS: 100.0")

if __name__ == "__main__":
    main()
