import os
import pandas as pd
import requests
import sys
import argparse
from tkinter import Tk, filedialog
import tempfile
import shutil

def download_image(url, output_path):
    # ... (same as before) ...
    try:
        response = requests.get(url, stream=True)
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
    except requests.exceptions.RequestException as e:
        return f"Network Error: {str(e)}"
    except Exception as e:
        return f"Unexpected Error: {str(e)}"

def main():
    parser = argparse.ArgumentParser(description="Download PSA images from Bynder.")
    parser.add_argument("--sku_file", help="Path to the Excel file containing SKUs in the first column.")
    parser.add_argument("--sku_list", help="Comma-separated string of SKUs directly from UI textbox.")
    parser.add_argument("--output_folder", help="Path to the folder where images will be downloaded.")
    parser.add_argument("--image_types", help="Comma-separated list of image types to download (e.g., grid,100,dimension).")
    
    args = parser.parse_args()

    skus_to_process = []
    temp_sku_xlsx_path = None

    # ... (SKU input determination logic - same as before) ...
    if args.sku_file:
        print(f"Script: Reading SKUs from file: {args.sku_file}")
        try:
            data = pd.read_excel(args.sku_file, dtype=str)
            skus_to_process = data.iloc[:, 0].dropna().astype(str).str.strip().tolist()
        except FileNotFoundError:
            print(f"Error: SKU file not found at '{args.sku_file}'. Exiting.", file=sys.stderr)
            sys.exit(1)
        except Exception as e:
            print(f"Error reading SKU input file '{args.sku_file}': {e}. Exiting.", file=sys.stderr)
            sys.exit(1)
    elif args.sku_list:
        print("Script: Processing SKUs from command line string (UI textbox input).")
        skus_to_process = [sku.strip() for sku in args.sku_list.split(',') if sku.strip()]
    else:
        root = Tk()
        root.withdraw()
        print("Script: No SKU input provided via arguments. Prompting interactively for file...")
        selected_sku_file = filedialog.askopenfilename(
            title="Select the Excel file with SKUs",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not selected_sku_file:
            print("No SKU file selected. Exiting.", file=sys.stderr)
            if root: root.destroy()
            sys.exit(0)
        print(f"Script: SKU file selected: {selected_sku_file}")
        try:
            data = pd.read_excel(selected_sku_file, dtype=str)
            skus_to_process = data.iloc[:, 0].dropna().astype(str).str.strip().tolist()
        except Exception as e:
            print(f"Error reading SKU input file '{selected_sku_file}': {e}. Exiting.", file=sys.stderr)
            if root: root.destroy()
            sys.exit(1)
    
    if not skus_to_process:
        print("No valid SKUs found in the provided input. Exiting.", file=sys.stderr)
        if 'root' in locals() and root: root.destroy()
        sys.exit(0)

    # ... (Output folder determination logic - ) ...
    output_folder = args.output_folder
    if not output_folder:
        if 'root' not in locals() or not root:
            root = Tk()
            root.withdraw()
        print("Script: No output folder provided. Prompting interactively...", file=sys.stderr)
        output_folder = filedialog.askdirectory(title="Select Folder for Downloads")
        if not output_folder:
            print("No output folder selected. Exiting.", file=sys.stderr)
            if root: root.destroy()
            sys.exit(0)
        print(f"Script: Output folder selected: {output_folder}")
    
    if 'root' in locals() and root:
        root.destroy()

    os.makedirs(output_folder, exist_ok=True)

    # ... (Image types determination logic - ) ...
    image_types_options = ["grid", "100", "200", "300", "400", "500", "dimension", "swatch"]
    image_types_flags = {k: False for k in image_types_options}

    if args.image_types:
        types_from_arg = [t.strip().lower() for t in args.image_types.split(',') if t.strip()]
        print(f"Script: Image types requested via argument: {types_from_arg}")
        if not types_from_arg:
             print("Script: No specific image types specified via argument. Defaulting to 'grid'.")
             image_types_flags["grid"] = True
        else:
            for img_type in types_from_arg:
                if img_type in image_types_flags:
                    image_types_flags[img_type] = True
                else:
                    print(f"Warning: Unknown image type '{img_type}' requested, skipping.", file=sys.stderr)
    elif sys.stdin.isatty():
        print("Script: No image types specified via argument. Prompting interactively...")
        for img_type_key in image_types_options:
            response = input(f"Do you want the SKU_{img_type_key} image? (Y/N): ").strip().lower()
            if response == 'y':
                image_types_flags[img_type_key] = True
        if not any(image_types_flags.values()):
            print("No image types selected interactively. Defaulting to 'grid'.")
            image_types_flags["grid"] = True
    else:
        print("Script: No image types specified. Defaulting to 'grid' image.")
        image_types_flags["grid"] = True


    base_url = "https://www.bynder.raymourflanigan.com/match/Product_SKU_Position/"
    report_rows = []
    headers = ["SKU"] + [f"SKU_{t}_Status" for t in image_types_options]

    print(f"Script: Starting download for {len(skus_to_process)} SKUs into '{output_folder}'...")
    total_skus = len(skus_to_process)
    for i, sku in enumerate(skus_to_process): # Use enumerate for progress tracking
        row = {"SKU": sku}
        for img_type in image_types_options:
            column_name = f"SKU_{img_type}_Status"
            if image_types_flags[img_type]:
                image_url = f"{base_url}{sku}_{img_type}/"
                output_path = os.path.join(output_folder, f"{sku}_{img_type}.jpg")
                status = download_image(image_url, output_path)
                row[column_name] = status
            else:
                row[column_name] = "Not Requested"
        report_rows.append(row)
        
        # --- Send Progress Update to UI ---
        progress_percentage = (i + 1) / total_skus * 100
        print(f"PROGRESS:{progress_percentage:.2f}", flush=True) # CRITICAL: print and flush

    report_df = pd.DataFrame(report_rows, columns=headers)
    report_file = os.path.join(output_folder, "download_report.csv")
    report_df.to_csv(report_file, index=False)

    print(f"Script: Download process complete. Report saved to {report_file}")

    if temp_sku_xlsx_path and os.path.exists(temp_sku_xlsx_path):
        try:
            os.remove(temp_sku_xlsx_path)
            print(f"Script: Removed temporary SKU file: {temp_sku_xlsx_path}")
            temp_dir = os.path.dirname(temp_sku_xlsx_path)
            if os.path.exists(temp_dir) and not os.listdir(temp_dir):
                shutil.rmtree(temp_dir)
                print(f"Script: Removed empty temporary directory: {temp_dir}")
        except Exception as e:
            print(f"Script: Warning: Could not remove temporary file or directory {temp_sku_xlsx_path}: {e}", file=sys.stderr)

if __name__ == "__main__":
    main()