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

    print_progress("PROGRESS: 0.0") # Initial progress

    # Determine if arguments were provided (from GUI) or if we need to prompt the user (standalone)
    if not args.matrix or not args.output:
        print_progress("\n--- Running Renamer Downloader in Standalone Mode ---")
        print_progress("Please provide the required paths using graphical dialogs:")
        
        spreadsheet_path = prompt_for_file_tk(
            "Select the Excel spreadsheet containing URLs",
            [("Excel files", "*.xlsx *.xls")]
        )
        if not spreadsheet_path:
            print_progress("No spreadsheet selected. Exiting.", is_stderr=True)
            print_progress("PROGRESS: 0.0", is_stderr=True) # Indicate failure
            sys.exit(1)

        download_folder = prompt_for_folder_tk("Select folder to save downloaded files")
        if not download_folder:
            print_progress("No download folder selected. Exiting.", is_stderr=True)
            print_progress("PROGRESS: 0.0", is_stderr=True) # Indicate failure
            sys.exit(1)
    else:
        # Arguments were provided, likely from the GUI
        spreadsheet_path = args.matrix
        download_folder = args.output

    print_progress(f"\nRenamer Downloader script starting with:")
    print_progress(f"  Matrix: {spreadsheet_path}")
    print_progress(f"  Output Folder: {download_folder}")

    # Ensure the download folder exists
    os.makedirs(download_folder, exist_ok=True)

    # Read the spreadsheet into a DataFrame
    try:
        print_progress("Loading spreadsheet...")
        df = pd.read_excel(spreadsheet_path)  # Assumes header row
    except FileNotFoundError:
        print_progress(f"ERROR: Spreadsheet not found at {spreadsheet_path}. Exiting.", is_stderr=True)
        print_progress("PROGRESS: 0.0", is_stderr=True)
        sys.exit(1)
    except Exception as e:
        print_progress(f"ERROR: Failed to read spreadsheet: {e}. Exiting.", is_stderr=True)
        print_progress("PROGRESS: 0.0", is_stderr=True)
        sys.exit(1)

    print_progress("PROGRESS: 10.0") # After loading spreadsheet

    errors = []
    downloaded_count = 0
    total_urls_to_process = 0
    urls_to_download = []

    # First pass: Collect all valid URLs and count them for accurate progress calculation
    print_progress("Scanning spreadsheet for URLs to download...")
    for row_idx in range(df.shape[0]):
        for col_idx in range(2, 11): # Columns C through K (indices 2-10)
            cell_value = df.iat[row_idx, col_idx]
            if pd.notna(cell_value) and isinstance(cell_value, str) and cell_value.strip():
                urls_to_download.append({
                    'url': cell_value.strip(),
                    'cell_coord': f"{chr(ord('A') + col_idx)}{row_idx + 2}" # +2 for 1-based row and header
                })
    total_urls_to_process = len(urls_to_download)

    if total_urls_to_process == 0:
        print_progress("No URLs found in the specified columns (C-K) of the spreadsheet. Exiting.")
        print_progress("PROGRESS: 100.0") # Complete immediately if nothing to do
        sys.exit(0)

    print_progress(f"Found {total_urls_to_process} URLs to process.")
    
    # Second pass: Perform downloads with progress updates
    for i, url_info in enumerate(urls_to_download):
        url = url_info['url']
        cell_coord = url_info['cell_coord']
        
        try:
            print_progress(f"Attempting to download from Cell {cell_coord}: {url}")
            final_name = download_and_save(url, download_folder)
            print_progress(f"  Downloaded: {final_name}")
            downloaded_count += 1
        except RuntimeError as e:
            print_progress(f"  ERROR for Cell {cell_coord}: {e}", is_stderr=True)
            errors.append({
                "Cell": cell_coord,
                "URL": url,
                "Error": str(e)
            })
        except Exception as e:
            print_progress(f"  UNEXPECTED ERROR for Cell {cell_coord}: {e}", is_stderr=True)
            errors.append({
                "Cell": cell_coord,
                "URL": url,
                "Error": f"Unexpected error: {str(e)}"
            })
        
        # Update progress bar
        current_progress = ((i + 1) / total_urls_to_process) * 100
        print_progress(f"PROGRESS: {current_progress:.1f}")

    # Prepare and save report
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_name = f"renamerDL_report_{timestamp}.xlsx"
    downloads_folder = get_downloads_folder()
    report_path = os.path.join(downloads_folder, report_name)

    print_progress(f"\n--- Download Summary ---")
    print_progress(f"Total URLs processed: {total_urls_to_process}")
    print_progress(f"Successfully downloaded: {downloaded_count}")
    print_progress(f"Downloads with errors: {len(errors)}")

    if errors:
        report_df = pd.DataFrame(errors)
        try:
            report_df.to_excel(report_path, index=False)
            print_progress(f"Completed with errors. A detailed report can be found at: {report_path}")
        except Exception as e:
            print_progress(f"ERROR: Failed to write error report to '{report_path}': {e}. Exiting.", is_stderr=True)
            print_progress("PROGRESS: 100.0", is_stderr=True) # Final progress update on error
            sys.exit(1) # Exit with an error code if there were failures
        print_progress("PROGRESS: 100.0") # Final progress update after report generation
        sys.exit(1) # Exit with an error code if there were failures
    else:
        print_progress("All specified URLs downloaded successfully!")
        print_progress("PROGRESS: 100.0") # Final progress update on success
        sys.exit(0) # Exit with success code

if __name__ == "__main__":
    main()
