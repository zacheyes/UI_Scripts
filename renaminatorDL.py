import os
import sys
import pandas as pd
import requests
import mimetypes
from urllib.parse import urlparse, unquote
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
import argparse # New import for command-line arguments

def prompt_for_file_tk(prompt_message, filetypes):
    """Prompt the user to select a file using a Tkinter dialog."""
    root = tk.Tk()
    root.withdraw() # Hide the main window
    file_path = filedialog.askopenfilename(title=prompt_message, filetypes=filetypes)
    root.destroy() # Close the Tkinter root window
    return file_path

def prompt_for_folder_tk(prompt_message):
    """Prompt the user to select a folder using a Tkinter dialog."""
    root = tk.Tk()
    root.withdraw() # Hide the main window
    folder_path = filedialog.askdirectory(title=prompt_message)
    root.destroy() # Close the Tkinter root window
    return folder_path

def get_downloads_folder():
    """Returns the user's Downloads folder, or home directory as a fallback."""
    home = Path.home()
    downloads = home / "Downloads"
    return downloads if downloads.exists() else home

def infer_extension(response, default=".jpg"):
    """Infers file extension from HTTP Content-Type header, defaults to .jpg."""
    content_type = response.headers.get("Content-Type", "").lower()
    ext = mimetypes.guess_extension(content_type.split(";")[0].strip()) if content_type else None
    if ext:
        return ext
    return default

def extract_filename_from_url(url):
    """Extracts the filename from a URL, handling URL decoding."""
    parsed = urlparse(url)
    name = os.path.basename(parsed.path)
    return unquote(name)

def download_and_save(url, dest_folder):
    """Downloads a file from a URL and saves it to the destination folder."""
    try:
        response = requests.get(url, stream=True, timeout=15)
        response.raise_for_status() # Raise an exception for HTTP errors (4xx or 5xx)
    except requests.exceptions.RequestException as e:
        raise RuntimeError(f"Download failed for URL '{url}': {e}")

    ext = infer_extension(response)
    raw_name = extract_filename_from_url(url)
    base, _ = os.path.splitext(raw_name)
    final_name = f"{base}{ext}"
    save_path = os.path.join(dest_folder, final_name)

    try:
        with open(save_path, "wb") as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)
    except IOError as e:
        raise RuntimeError(f"Saving file '{final_name}' failed: {e}")
    except Exception as e:
        raise RuntimeError(f"An unexpected error occurred while saving '{final_name}': {e}")

    return final_name

def main():
    parser = argparse.ArgumentParser(description="Renamer Downloader Script")
    parser.add_argument('--matrix', help='Path to the Excel spreadsheet containing URLs.')
    parser.add_argument('--output', help='Path to the folder to save downloaded files.')
    
    args = parser.parse_args()

    spreadsheet_path = None
    download_folder = None

    # Determine if arguments were provided (from GUI) or if we need to prompt the user (standalone)
    if not args.matrix or not args.output:
        print("\n--- Running Renamer Downloader in Standalone Mode ---")
        print("Please provide the required paths using graphical dialogs:")
        
        spreadsheet_path = prompt_for_file_tk(
            "Select the Excel spreadsheet containing URLs",
            [("Excel files", "*.xlsx *.xls")]
        )
        if not spreadsheet_path:
            sys.exit("No spreadsheet selected. Exiting.")

        download_folder = prompt_for_folder_tk("Select folder to save downloaded files")
        if not download_folder:
            sys.exit("No download folder selected. Exiting.")
    else:
        # Arguments were provided, likely from the GUI
        spreadsheet_path = args.matrix
        download_folder = args.output

    print(f"\nRenamer Downloader script starting with:")
    print(f"  Matrix: {spreadsheet_path}")
    print(f"  Output Folder: {download_folder}")

    # Ensure the download folder exists
    os.makedirs(download_folder, exist_ok=True)

    # Read the spreadsheet into a DataFrame
    try:
        df = pd.read_excel(spreadsheet_path)  # Assumes header row
    except FileNotFoundError:
        sys.exit(f"ERROR: Spreadsheet not found at {spreadsheet_path}. Exiting.")
    except Exception as e:
        sys.exit(f"ERROR: Failed to read spreadsheet: {e}. Exiting.")

    errors = []
    downloaded_count = 0
    total_urls = 0

    # Iterate over rows and columns C (index 2) through K (index 10)
    # Excel columns C-K correspond to pandas indices 2-10
    # Note: If your spreadsheet structure might vary, you might need to find columns by name.
    # For now, assuming fixed column indices.
    for row_idx in range(df.shape[0]):
        for col_idx in range(2, 11): # Columns C through K
            cell_value = df.iat[row_idx, col_idx]
            if pd.notna(cell_value) and isinstance(cell_value, str) and cell_value.strip():
                url = cell_value.strip()
                total_urls += 1
                # Convert column index to Excel letter (0=A, 1=B, etc.)
                cell_coord = f"{chr(ord('A') + col_idx)}{row_idx + 2}" # +2 for 1-based row and header

                try:
                    print(f"Attempting to download from Cell {cell_coord}: {url}")
                    final_name = download_and_save(url, download_folder)
                    print(f"  Downloaded: {final_name}")
                    downloaded_count += 1
                except RuntimeError as e:
                    print(f"  ERROR for Cell {cell_coord}: {e}")
                    errors.append({
                        "Cell": cell_coord,
                        "URL": url,
                        "Error": str(e)
                    })
                except Exception as e:
                    print(f"  UNEXPECTED ERROR for Cell {cell_coord}: {e}")
                    errors.append({
                        "Cell": cell_coord,
                        "URL": url,
                        "Error": f"Unexpected error: {str(e)}"
                    })

    # Prepare and save report
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_name = f"renamerDL_report_{timestamp}.xlsx"
    downloads_folder = get_downloads_folder()
    report_path = os.path.join(downloads_folder, report_name)

    print(f"\n--- Download Summary ---")
    print(f"Total URLs processed: {total_urls}")
    print(f"Successfully downloaded: {downloaded_count}")
    print(f"Downloads with errors: {len(errors)}")

    if errors:
        report_df = pd.DataFrame(errors)
        try:
            report_df.to_excel(report_path, index=False)
            print(f"Completed with errors. A detailed report can be found at: {report_path}")
        except Exception as e:
            sys.exit(f"ERROR: Failed to write error report to '{report_path}': {e}. Exiting.")
        sys.exit(1) # Exit with an error code if there were failures
    else:
        print("All specified URLs downloaded successfully!")
        sys.exit(0) # Exit with success code

if __name__ == "__main__":
    main()