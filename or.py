import pandas as pd
from datetime import datetime
import os
import subprocess
import sys

# We keep tkinter imports here because they are needed for the standalone GUI mode
# and for message boxes in that mode.
try:
    import tkinter as tk
    from tkinter import filedialog, messagebox
except ImportError:
    # Handle cases where tkinter might not be installed (e.g., in a server environment)
    tk = None
    filedialog = None
    messagebox = None

def process_input_and_get_result(input_path):
    """
    Processes the input file (Excel or text) and returns the OR boolean string.
    This function no longer handles file output or GUI messages directly.
    """
    try:
        if input_path.lower().endswith(('.xlsx', '.xls')):
            df = pd.read_excel(input_path)
            if df.empty or df.shape[1] == 0:
                raise ValueError("The selected Excel file is empty or has no columns.")
            values = df.iloc[:, 0]
        elif input_path.lower().endswith('.txt'):
            with open(input_path, 'r', encoding='utf-8') as f:
                values = [line.strip() for line in f if line.strip()]
            if not values:
                raise ValueError("The selected text file is empty.")
        else:
            raise ValueError("Unsupported file type. Please provide an Excel (.xlsx, .xls) or a text (.txt) file.")

        processed_values = pd.Series(values).dropna().astype(str).tolist()
        if not processed_values:
            return "No valid values found to create a boolean string." # Return a user-friendly message
        
        result_string = " OR ".join(processed_values)
        return result_string

    except Exception as e:
        # For command-line execution, printing to stderr is standard for errors
        print(f"Error during processing: {e}", file=sys.stderr)
        sys.exit(1) # Exit with a non-zero code to indicate an error
        return None # Should not be reached due to sys.exit

def run_gui_mode_standalone():
    """
    Runs the script in standalone GUI mode (when executed directly without arguments).
    This mode uses Tkinter for file selection and creates a file in Downloads.
    """
    if tk is None:
        print("Tkinter is not installed. Cannot run in standalone GUI mode.", file=sys.stderr)
        return

    root = tk.Tk()
    root.withdraw() # Hide the main window

    input_file = filedialog.askopenfilename(
        title="Select an Excel or Text File",
        filetypes=[("Excel or Text files", "*.xlsx *.xls *.txt")]
    )

    if not input_file:
        messagebox.showinfo("Cancelled", "File selection cancelled. No output generated.")
        root.destroy()
        return

    try:
        result_string = process_input_and_get_result(input_file)

        downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
        os.makedirs(downloads_folder, exist_ok=True)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = os.path.join(downloads_folder, f"or_{timestamp}.txt")

        with open(output_file, "w", encoding="utf-8") as f:
            f.write(result_string)

        if sys.platform.startswith('win'):
            os.startfile(output_file)
        elif sys.platform.startswith('darwin'):
            subprocess.run(['open', output_file], check=True)
        else:
            subprocess.run(['xdg-open', output_file], check=True)
            
        messagebox.showinfo("Success", f"Output saved to: {output_file}")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")
    finally:
        root.destroy()

def run_cli_mode_for_gui(input_path):
    """
    Runs the script in CLI mode, specifically when called by the main GUI.
    It prints the result to stdout for the GUI to capture.
    """
    result_string = process_input_and_get_result(input_path)
    if result_string is not None: # Only print if processing was successful
        print(result_string) # This is what the GUI will capture

if __name__ == "__main__":
    if len(sys.argv) > 1:
        # This branch is executed when the script is called with command-line arguments.
        # Your main GUI.py script will call it this way.
        input_file_path_from_args = sys.argv[1]
        run_cli_mode_for_gui(input_file_path_from_args)
    else:
        # This branch is executed when the script is run directly (e.g., double-clicked).
        run_gui_mode_standalone()