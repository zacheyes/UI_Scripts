import tkinter as tk
from tkinter import filedialog, scrolledtext, ttk, messagebox
import os
import subprocess
import shutil
import datetime
import sys
import json
import tkinter.font as tkFont
import requests
import filecmp
import threading
import pandas as pd
import tempfile

# --- Configuration ---
# Define your GitHub repository details
GITHUB_USERNAME = "zacheyes"
GITHUB_REPO_NAME = "UI_Scripts"
# This base URL points to the root of the 'main' branch for raw content.
# Ensure your scripts are directly in the root of the 'main' branch in your GitHub repo.
GITHUB_RAW_BASE_URL = f"https://raw.githubusercontent.com/{GITHUB_USERNAME}/{GITHUB_REPO_NAME}/main/"

# --- GUI Script specific constants ---
GUI_SCRIPT_FILENAME = "GUI.py"
UPDATE_IN_PROGRESS_MARKER = "gui_update_in_progress.tmp"


SCRIPT_FILENAMES = {
    "Main Renaminator Script": "renaminator.py",
    "Downloader Script": "renaminatorDL.py",
    "File Copier Script": "renaminatorCF.py",
    "Renamer Spreadsheet Template": "renaminator.xlsx",
    "Cropping Script Silo (1688)": "reformat1688_silo.py",
    "Cropping Script Room (1688)": "reformat1688_room.py",
    "Cropping Script Room CutLR (1688)": "reformat1688_room_cutLR.py",
    "Cropping Script Room CutTopBot (1688)": "reformat1688_room_cutTopBot.py",
    "Cropping Script Silo (2200)": "reformat2200_silo.py",
    "Cropping Script Room (2200)": "reformat2200_room.py",
    "Bynder Metadata Prep": "bynder_metadataPrep.py",
    "Check Bynder PSAs script": "check_BynderPSAs.py",
    "Download PSAs script": "downloadPSAs.py",
    "Get Measurements script": "get_MeasurementsFromSTEP.py",
    "Convert Bynder Metadata to XLS": "convertBynderMetadataToXls.py",
    "Move Files from Spreadsheet": "move_filename.py",
    "OR Boolean Search Creator": "or.py",
}

# NEW: GitHub URLs for Python scripts
GITHUB_SCRIPT_URLS = {
    "renaminator.py": GITHUB_RAW_BASE_URL + "renaminator.py",
    "renaminatorDL.py": GITHUB_RAW_BASE_URL + "renaminatorDL.py",
    "renaminatorCF.py": GITHUB_RAW_BASE_URL + "renaminatorCF.py",
    "reformat1688_silo.py": GITHUB_RAW_BASE_URL + "reformat1688_silo.py",
    "reformat1688_room.py": GITHUB_RAW_BASE_URL + "reformat1688_room.py",
    "reformat1688_room_cutLR.py": GITHUB_RAW_BASE_URL + "reformat1688_room_cutLR.py",
    "reformat1688_room_cutTopBot.py": GITHUB_RAW_BASE_URL + "reformat1688_room_cutTopBot.py",
    "reformat2200_silo.py": GITHUB_RAW_BASE_URL + "reformat2200_silo.py",
    "reformat2200_room.py": GITHUB_RAW_BASE_URL + "reformat2200_room.py",
    "bynder_metadataPrep.py": GITHUB_RAW_BASE_URL + "bynder_metadataPrep.py",
    "check_BynderPSAs.py": GITHUB_RAW_BASE_URL + "check_BynderPSAs.py",
    "downloadPSAs.py": GITHUB_RAW_BASE_URL + "downloadPSAs.py",
    "get_MeasurementsFromSTEP.py": GITHUB_RAW_BASE_URL + "get_MeasurementsFromSTEP.py",
    GUI_SCRIPT_FILENAME: GITHUB_RAW_BASE_URL + GUI_SCRIPT_FILENAME,
    "convertBynderMetadataToXls.py": GITHUB_RAW_BASE_URL + "convertBynderMetadataToXls.py",
    "move_filename.py": GITHUB_RAW_BASE_URL + "move_filename.py",
    "or.py": GITHUB_RAW_BASE_URL + "or.py",
}

RENAMER_EXCEL_URL = "https://www.bynder.raymourflanigan.com/m/333617bb041ff764/original/renaminator.xlsx"

CONFIG_FILE = "rf_renamer_config.json"

# --- General Helper Functions ---

def _append_to_log(log_widget, text, is_stderr=False):
    log_widget.configure(state='normal')
    if is_stderr:
        log_widget.insert(tk.END, text, 'error')
    else:
        log_widget.insert(tk.END, text)
    log_widget.see(tk.END)
    log_widget.configure(state='disabled')

# --- Progress Bar Specific Helper Functions ---

def _prepare_progress_ui(progress_bar, progress_label, run_button_wrapper, progress_wrapper, initial_text):
    run_button_wrapper.grid_remove()
    progress_wrapper.grid(row=0, column=1)

    progress_bar.config(value=0, maximum=100)
    progress_bar.start()
    progress_label.config(text=initial_text)


def _update_progress_ui(progress_bar, progress_label, value):
    progress_bar.stop()
    progress_bar['value'] = value
    progress_label.config(text=f"{value:.1f}%")

def _on_process_complete_with_progress_ui(success, full_output, progress_bar, progress_label, run_button_wrapper, progress_wrapper, success_callback, error_callback, log_output_widget):
    if progress_bar:
        progress_bar.stop()
        progress_bar['value'] = 0
    if progress_label:
        progress_label.config(text="")
    
    progress_wrapper.grid_remove()
    run_button_wrapper.grid(row=0, column=1)

    if progress_bar and progress_bar.winfo_toplevel():
        progress_bar.winfo_toplevel().config(cursor="")
    elif log_output_widget and log_output_widget.winfo_toplevel():
        log_output_widget.winfo_toplevel().config(cursor="")

    if success:
        log_output_widget.insert(tk.END, "\nScript completed successfully.\n", 'success')
        if success_callback:
            success_callback(full_output)
    else:
        log_output_widget.insert(tk.END, "\nScript failed. Please check the log above for errors.\n", 'error')
        if error_callback:
            error_callback(full_output)
    log_output_widget.see(tk.END)

# --- Run Script functions based on progress display needs ---

def _run_script_with_progress(script_full_path, args, log_output_widget, progress_bar, progress_label, run_button_wrapper, progress_wrapper, success_callback, error_callback, initial_progress_text):
    print("DEBUG (UI): Running script with progress bar.", file=sys.stderr)
    
    python_executable = sys.executable
    command = [python_executable, script_full_path]
    if args:
        command.extend(args)
    full_command_str = ' '.join(command)
    _append_to_log(log_output_widget, f"Executing subprocess command: {full_command_str}\n")

    log_output_widget.winfo_toplevel().after(0, lambda: _prepare_progress_ui(progress_bar, progress_label, run_button_wrapper, progress_wrapper, initial_progress_text))

    def _read_output_thread():
        process = None
        stdout_buffer = []
        stderr_buffer = []
        try:
            env = os.environ.copy()
            env['PYTHONIOENCODING'] = 'utf-8'

            process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                         bufsize=1, universal_newlines=True, env=env)

            def read_stream(stream, buffer, is_stderr=False):
                for line in stream:
                    buffer.append(line)
                    log_output_widget.after(0, lambda log=log_output_widget, l=line: _append_to_log(log, l, is_stderr))
                    if line.startswith("PROGRESS:"):
                        try:
                            percent_str = line.split("PROGRESS:")[1].strip()
                            percent_val = float(percent_str)
                            progress_bar.after(0, lambda pb=progress_bar, pl=progress_label, val=percent_val: _update_progress_ui(pb, pl, val))
                        except ValueError:
                            print(f"DEBUG (UI): Could not parse progress: {line.strip()}", file=sys.stderr)
                stream.close()

            stdout_thread = threading.Thread(target=read_stream, args=(process.stdout, stdout_buffer, False))
            stderr_thread = threading.Thread(target=read_stream, args=(process.stderr, stderr_buffer, True))
            
            stdout_thread.start()
            stderr_thread.start()

            stdout_thread.join()
            stderr_thread.join()

            process.wait()
            success = (process.returncode == 0)
            full_output = "".join(stdout_buffer) + "".join(stderr_buffer)
            log_output_widget.after(0, lambda: _on_process_complete_with_progress_ui(success, full_output, progress_bar, progress_label, run_button_wrapper, progress_wrapper, success_callback, error_callback, log_output_widget))

        except FileNotFoundError:
            error_msg = f"  Error: Python interpreter (or script) not found. Check paths and ensure Python is correctly installed and accessible.\n"
            log_output_widget.after(0, lambda: _append_to_log(log_output_widget, error_msg, is_stderr=True))
            log_output_widget.after(0, lambda: _on_process_complete_with_progress_ui(False, error_msg, progress_bar, progress_label, run_button_wrapper, progress_wrapper, success_callback, error_callback, log_output_widget))
        except Exception as e:
            error_msg = f"  An unexpected error occurred during subprocess execution: {e}\n"
            log_output_widget.after(0, lambda: _append_to_log(log_output_widget, error_msg, is_stderr=True))
            log_output_widget.after(0, lambda: _on_process_complete_with_progress_ui(False, error_msg, progress_bar, progress_label, run_button_wrapper, progress_wrapper, success_callback, error_callback, log_output_widget))

    subprocess_thread = threading.Thread(target=_read_output_thread)
    subprocess_thread.daemon = True
    subprocess_thread.start()
    return True, "Process started in background."


def _run_script_no_progress(script_full_path, args, log_output_widget, success_callback=None, error_callback=None):
    print("DEBUG (UI): Running script without progress bar.", file=sys.stderr)

    python_executable = sys.executable
    command = [python_executable, script_full_path]
    if args:
        command.extend(args)
    full_command_str = ' '.join(command)
    _append_to_log(log_output_widget, f"Executing subprocess command: {full_command_str}\n")

    try:
        env = os.environ.copy()
        env['PYTHONIOENCODING'] = 'utf-8'
        result = subprocess.run(command, capture_output=True, check=False, universal_newlines=True, env=env)
        
        stdout_str = result.stdout
        stderr_str = result.stderr
        
        full_output = stdout_str + stderr_str
        
        if stdout_str:
            _append_to_log(log_output_widget, "\n--- Script Output ---\n" + stdout_str)
        if stderr_str:
            _append_to_log(log_output_widget, "\n--- Script Errors (stderr) ---\n" + stderr_str, is_stderr=True)
        _append_to_log(log_output_widget, f"\nScript exited with return code: {result.returncode}\n")

        success = (result.returncode == 0)
        
        log_output_widget.winfo_toplevel().config(cursor="")
        if success:
            if success_callback: success_callback(full_output)
        else:
            if error_callback: error_callback(full_output)
        
        return success, full_output

    except FileNotFoundError:
        error_msg = f"  Error: Python interpreter (or script) not found. Check paths and ensure Python is correctly installed and accessible.\n"
        _append_to_log(log_output_widget, error_msg, is_stderr=True)
        log_output_widget.winfo_toplevel().config(cursor="")
        if error_callback: error_callback("")
        return False, error_msg
    except Exception as e:
        error_msg = f"  An unexpected error occurred during subprocess execution: {e}\n"
        _append_to_log(log_output_widget, error_msg, is_stderr=True)
        log_output_widget.winfo_toplevel().config(cursor="")
        if error_callback: error_callback("")
        return False, error_msg

# --- Main Dispatcher Function for running scripts (MODIFIED) ---
def run_script_wrapper(script_full_path, is_python_script, args=None, log_output_widget=None,
                       progress_bar=None, progress_label=None, run_button_wrapper=None,
                       progress_wrapper=None, success_callback=None, error_callback=None,
                       initial_progress_text="Starting..."):
    
    print("DEBUG (UI): Entered run_script_wrapper function.", file=sys.stderr)

    if not os.path.exists(script_full_path):
        error_msg = f"Error: File not found at {script_full_path}\n"
        _append_to_log(log_output_widget, error_msg, is_stderr=True)
        log_output_widget.winfo_toplevel().config(cursor="")
        if error_callback: error_callback("")
        return False, error_msg

    if is_python_script:
        if progress_bar is not None and progress_label is not None and \
           run_button_wrapper is not None and progress_wrapper is not None:
            return _run_script_with_progress(script_full_path, args, log_output_widget,
                                             progress_bar, progress_label, run_button_wrapper,
                                             progress_wrapper, success_callback, error_callback,
                                             initial_progress_text)
        else:
            return _run_script_no_progress(script_full_path, args, log_output_widget,
                                             success_callback, error_callback)
    else:
        _append_to_log(log_output_widget, f"Opening file: {script_full_path}\n")
        try:
            os.startfile(script_full_path)
            _append_to_log(log_output_widget, f"  File opened.\n")
            return True, f"Opened file: {script_full_path}"
        except Exception as e:
            _append_to_log(log_output_widget, f"  Error opening file: {e}\n", is_stderr=True)
            return False, f"Error opening file: {e}"


class Tooltip:
    def __init__(self, widget, text, bg_color, text_color):
        self.widget = widget
        self.text = text
        self.tooltip_window = None
        self.id = None
        self.x = 0
        self.y = 0
        self.bg_color = bg_color  
        self.text_color = text_color
        self.widget.bind("<Enter>", self.show_tooltip)
        self.widget.bind("<Leave>", self.hide_tooltip)

    def show_tooltip(self, event=None):
        self.x = self.widget.winfo_rootx() + 20  
        self.y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5  
        self.id = self.widget.after(500, self._display_tooltip)  

    def _display_tooltip(self):
        if self.tooltip_window:
            return
        self.tooltip_window = tk.Toplevel(self.widget)
        self.tooltip_window.wm_overrideredirect(True)  
        self.tooltip_window.wm_geometry(f"+{self.x}+{self.y}")

        label = ttk.Label(self.tooltip_window, text=self.text, background=self.bg_color, relief=tk.SOLID, borderwidth=1,
                                 font=("Arial", 11), foreground=self.text_color, wraplength=400)
        label.pack(padx=5, pady=5)

    def hide_tooltip(self, event=None):
        if self.id:
            self.widget.after_cancel(self.id)
        if self.tooltip_window:
            self.tooltip_window.destroy()
        self.tooltip_window = None

class RenamerApp:
    def __init__(self, master):
        self.master = master
        master.title("Raymour & Flanigan Renamer Tool")
        master.geometry("700x800")  
        master.resizable(True, True)  

        self.current_theme = tk.StringVar(value="Light")
        self.style = ttk.Style()  
        
        self.base_font = tkFont.Font(family="Arial", size=10)
        self.header_font = tkFont.Font(family="Arial", size=12, weight="bold")
        self.log_font = tkFont.Font(family="Consolas", size=9)

        self._restarting_for_update = False

        self._initialize_logger_widget()

        self._apply_theme(self.current_theme.get())  

        self.scripts_root_folder = tk.StringVar(value=os.path.dirname(os.path.abspath(__file__)))
        self.last_update_timestamp = tk.StringVar(value="Last update: Never")
        self.gui_last_update_timestamp = tk.StringVar(value="Last GUI update: Never")
        
        self.check_psa_sku_spreadsheet_path = tk.StringVar(value="")
        self.download_psa_sku_spreadsheet_path = tk.StringVar(value="")
        self.get_measurements_sku_spreadsheet_path = tk.StringVar(value="")
        self.bynder_metadata_csv_path = tk.StringVar(value="")
        self.move_files_source_folder = tk.StringVar(value="")
        self.move_files_destination_folder = tk.StringVar(value="")
        self.move_files_excel_path = tk.StringVar(value="")
        self.or_boolean_input_type = tk.StringVar(value="spreadsheet")
        self.or_boolean_spreadsheet_path = tk.StringVar(value="")
        self.or_boolean_output_text = tk.StringVar(value="")


        self.check_psa_input_type = tk.StringVar(value="spreadsheet")
        self.download_psa_input_type = tk.StringVar(value="spreadsheet")
        self.get_measurements_input_type = tk.StringVar(value="spreadsheet")
        self.move_files_input_type = tk.StringVar(value="spreadsheet")

        self.source_type = tk.StringVar(value="inline")  
        self.master_matrix_path = tk.StringVar(value="")
        self.rename_input_folder = tk.StringVar(value="")
        self.vendor_code = tk.StringVar(value="")
        self.inline_source_folder = tk.StringVar(value="")
        self.inline_matrix_path = tk.StringVar(value="")
        self.inline_output_folder = tk.StringVar(value="")
        self.pso1_matrix_path = tk.StringVar(value="")
        self.pso1_output_folder = tk.StringVar(value="")
        self.pso2_network_folder = tk.StringVar(value="")
        self.pso2_matrix_path = tk.StringVar(value="")
        self.pso2_output_folder = tk.StringVar(value="")
        self.prep_input_path = tk.StringVar(value="")
        self.bynder_assets_folder = tk.StringVar(value="")
        self.download_psa_output_folder = tk.StringVar(value="")  
        
        self.download_psa_grid = tk.BooleanVar(value=False)
        self.download_psa_100 = tk.BooleanVar(value=False)
        self.download_psa_200 = tk.BooleanVar(value=False)
        self.download_psa_300 = tk.BooleanVar(value=False)
        self.download_psa_400 = tk.BooleanVar(value=False)
        self.download_psa_500 = tk.BooleanVar(value=False)
        self.download_psa_dimension = tk.BooleanVar(value=False)
        self.download_psa_swatch = tk.BooleanVar(value=False)
        self.download_psa_5000 = tk.BooleanVar(value=False)
        self.download_psa_5100 = tk.BooleanVar(value=False)
        self.download_psa_5200 = tk.BooleanVar(value=False)
        self.download_psa_5300 = tk.BooleanVar(value=False)
        self.download_psa_squareThumbnail = tk.BooleanVar(value=False)

        # List of all BooleanVar objects for "Download PSAs"
        self.download_psa_checkboxes = [
            self.download_psa_grid, self.download_psa_100, self.download_psa_200,
            self.download_psa_300, self.download_psa_400, self.download_psa_500,
            self.download_psa_dimension, self.download_psa_swatch, self.download_psa_5000,
            self.download_psa_5100, self.download_psa_5200, self.download_psa_5300,
            self.download_psa_squareThumbnail
        ]


        self.log_expanded = False

        self._create_widgets()
        self._load_configuration()

        self.log_print(f"UI launched with Python {sys.version.split(' ')[0]} from: {sys.executable}\n")
        self.log_print("UI initialized. Please select paths and run operations.\n")

        self.master.after(100, self._handle_startup_update_check)

        master.protocol("WM_DELETE_WINDOW", self._on_closing)

    def _initialize_logger_widget(self):
        self.log_text_early_placeholder = scrolledtext.ScrolledText(self.master, width=1, height=1, state='disabled')
        
        def custom_print(*args, **kwargs):
            text = " ".join(map(str, args)) + kwargs.get('end', '\n')
            if hasattr(self, 'log_text') and self.log_text.winfo_exists():
                self.log_text.configure(state='normal')
                self.log_text.insert(tk.END, text)
                self.log_text.see(tk.END)
                self.log_text.configure(state='disabled')
            else:
                print(text, end='')  

        self.log_print = custom_print
        self.log_text = self.log_text_early_placeholder


    def _on_closing(self):
        if not self._restarting_for_update:
            self._save_configuration()
        self.master.destroy()

    def _save_configuration(self):
        config_data = {
            "theme": self.current_theme.get(),
            "scripts_root_folder": self.scripts_root_folder.get(),
            "last_update": self.last_update_timestamp.get(),
            "gui_last_update": self.gui_last_update_timestamp.get(),
        }
        try:
            with open(CONFIG_FILE, 'w') as f:
                json.dump(config_data, f, indent=4)
            self.log_print("Configuration saved successfully.\n")
        except Exception as e:
            self.log_print(f"Error saving configuration: {e}\n")

    def _load_configuration(self):
        """Loads specified configuration items from the JSON file."""
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r') as f:
                    config_data = json.load(f)
                
                self.scripts_root_folder.set(config_data.get("scripts_root_folder", os.path.dirname(os.path.abspath(__file__))))
                loaded_theme = config_data.get("theme", "Light")
                self.current_theme.set(loaded_theme)
                self._apply_theme(loaded_theme)

                last_update_from_config = config_data.get("last_update", "Last update: Never")
                if not last_update_from_config.startswith("Last update:"):
                    self.last_update_timestamp.set(f"Last update: {last_update_from_config}")
                else:
                    self.last_update_timestamp.set(last_update_from_config)

                gui_last_update_from_config = config_data.get("gui_last_update", "Last GUI update: Never")
                if not gui_last_update_from_config.startswith("Last GUI update:"):
                    self.gui_last_update_timestamp.set(f"Last GUI update: {gui_last_update_from_config}")
                else:
                    self.gui_last_update_timestamp.set(gui_last_update_from_config)


                self.log_print("Configuration loaded successfully.\n")
            except json.JSONDecodeError as e:
                self.log_print(f"Error reading configuration file (JSON format issue): {e}\n")
            except Exception as e:
                self.log_print(f"Error loading configuration: {e}\n")
        else:
            self.log_print("No existing configuration file found. Using default paths.\n")

        self._setup_initial_state()


    def _apply_theme(self, theme_name):
        self.current_theme.set(theme_name)

        self.RF_PURPLE_BASE = "#4f245e"  
        self.RF_WHITE_BASE = "#FFFFFF"  

        if theme_name == "Dark":
            self.primary_bg = "#2B2B2B"  
            self.secondary_bg = "#3C3C3C"  
            self.text_color = "#E0E0E0"  
            self.header_text_color = "#FFFFFF"  
            self.accent_color = self.RF_PURPLE_BASE
            self.border_color = "#555555"  
            self.log_bg = "#1E1E1E"  
            self.log_text_color = "#CCCCCC"  
            self.trough_color = "#555555"  
            self.slider_color = "#888888"  
            self.checkbox_indicator_off = "#3C3C3C"  
            self.checkbox_indicator_on = self.accent_color
            self.checkbox_hover_bg = "#505050"  
            self.radiobutton_hover_bg = "#505050"

        else:
            self.primary_bg = "#F0F0F0"  
            self.secondary_bg = "#FFFFFF"  
            self.text_color = "#333333"  
            self.header_text_color = self.RF_PURPLE_BASE  
            self.accent_color = self.RF_PURPLE_BASE
            self.border_color = "#CCCCCC"  
            self.log_bg = "#E8E8E8"  
            self.log_text_color = "#444444"  
            self.trough_color = "#E0E0E0"  
            self.slider_color = "#BBBBBB"  
            self.checkbox_indicator_off = "#E0E0E0"
            self.checkbox_indicator_on = self.accent_color
            self.checkbox_hover_bg = "#E0E0E0"
            self.radiobutton_hover_bg = "#E0E0E0"
            
        self.master.config(bg=self.primary_bg)
        if hasattr(self, 'canvas'):  
            self.canvas.config(bg=self.primary_bg)

        self.style.theme_use("clam")  

        self.style.configure('.',
                             font=self.base_font,
                             background=self.primary_bg,
                             foreground=self.text_color)
        
        self.style.configure('TFrame',
                             background=self.primary_bg)
        
        self.style.configure('SectionFrame.TFrame',
                             background=self.secondary_bg,
                             borderwidth=1,
                             relief="solid",  
                             padding=0)  

        self.style.configure('TLabel',
                             background=self.primary_bg,
                             foreground=self.text_color)
        
        self.style.configure('Header.TLabel',
                             font=self.header_font,
                             foreground=self.header_text_color,
                             background=self.secondary_bg)  

        self.style.configure('TButton',
                             background=self.accent_color,
                             foreground=self.RF_WHITE_BASE,  
                             font=self.base_font,
                             relief='flat',
                             padding=5)
        self.style.map('TButton',
                         background=[('active', self._shade_color(self.accent_color, -0.1))],  
                         foreground=[('active', self.RF_WHITE_BASE)])  

        self.style.configure('TEntry',
                             fieldbackground=self.secondary_bg,
                             foreground=self.text_color,
                             borderwidth=1,
                             relief="solid")
        
        self.style.configure('TScrollbar',
                             troughcolor=self.trough_color,
                             background=self.slider_color,
                             bordercolor=self.trough_color,
                             arrowcolor=self.text_color)
        self.style.map('TScrollbar',
                         background=[('active', self._shade_color(self.slider_color, -0.1))])

        self.style.configure('TNotebook',
                             background=self.primary_bg,
                             borderwidth=0)
        self.style.configure('TNotebook.Tab',
                             background=self._shade_color(self.primary_bg, -0.05),  
                             foreground=self.text_color,
                             font=self.base_font,
                             padding=[5, 2])
        self.style.map('TNotebook.Tab',
                         background=[('selected', self.accent_color)],
                         foreground=[('selected', self.RF_WHITE_BASE)],  
                         expand=[('selected', [1, 1, 1, 0])])  

        self.style.configure('TRadiobutton',
                             background=self.primary_bg,
                             foreground=self.text_color,
                             font=self.base_font,
                             indicatorcolor=self.accent_color)
        self.style.map('TRadiobutton',
                         background=[('active', self.radiobutton_hover_bg)],
                         foreground=[('active', self.text_color)],
                         indicatorcolor=[('selected', self.accent_color), ('!selected', self.checkbox_indicator_off)])
        
        self.style.configure('TCheckbutton',
                             background=self.primary_bg,
                             foreground=self.text_color,
                             font=self.base_font,
                             indicatorcolor=self.checkbox_indicator_off)
        self.style.map('TCheckbutton',
                         background=[('active', self.checkbox_hover_bg)],
                         foreground=[('active', self.text_color)],
                         indicatorcolor=[('selected', self.checkbox_indicator_on), ('!selected', self.checkbox_indicator_off)])

        self.style.configure('TSeparator', background=self.border_color, relief='solid', sashrelief='solid', sashwidth=3)
        self.style.layout('TSeparator',
                          [('TSeparator.separator', {'sticky': 'nswe'})])

        self.style.configure('TCombobox',
                             fieldbackground=self.secondary_bg,  
                             background=self.primary_bg,  
                             foreground=self.text_color,
                             arrowcolor=self.text_color)
        self.style.map('TCombobox',
                         fieldbackground=[('readonly', self.secondary_bg)],
                         background=[('readonly', self.primary_bg)],
                         foreground=[('readonly', self.text_color)],
                         selectbackground=[('readonly', self._shade_color(self.secondary_bg, -0.05))],  
                         selectforeground=[('readonly', self.text_color)])  

        if hasattr(self, 'log_text'):
            self.log_text.config(bg=self.log_bg, fg=self.log_text_color,
                                 insertbackground=self.log_text_color,
                                 selectbackground=self.accent_color,
                                 selectforeground=self.RF_WHITE_BASE)
            self.log_text.tag_config('error', foreground='#FF6B6B')
            self.log_text.tag_config('success', foreground='#6BFF6B')
        
        self._update_all_widget_colors()  

    def _shade_color(self, hex_color, percent):
        """Shades a hex color by a given percentage. Positive percent for lighter, negative for darker."""
        hex_color = hex_color.lstrip('#')
        rgb = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
        
        new_rgb = []
        for color_val in rgb:
            new_val = color_val * (1 + percent)
            new_val = max(0, min(255, int(new_val)))
            new_rgb.append(new_val)
            
        return '#%02x%02x%02x' % tuple(new_rgb)

    def _update_all_widget_colors(self):
        for widget in self.master.winfo_children():
            self._update_widget_color_recursive(widget)

    def _update_widget_color_recursive(self, widget):
        try:
            if hasattr(widget, 'config'):
                if 'background' in widget.config():
                    if isinstance(widget, ttk.Label) and widget.cget('style') == 'Header.TLabel':
                        widget.config(background=self.secondary_bg)  
                    else:
                        widget.config(background=self.primary_bg)
                if 'foreground' in widget.config():
                    if isinstance(widget, ttk.Label) and widget.cget('style') == 'Header.TLabel':
                        widget.config(foreground=self.header_text_color)
                    else:
                        widget.config(foreground=self.text_color)
            
            if isinstance(widget, tk.Canvas):
                widget.config(bg=self.primary_bg)
            elif isinstance(widget, scrolledtext.ScrolledText):
                widget.config(bg=self.log_bg, fg=self.log_text_color,
                                 insertbackground=self.log_text_color,
                                 selectbackground=self.accent_color,
                                 selectforeground=self.RF_WHITE_BASE)

        except tk.TclError:
            pass  

        for child_widget in widget.winfo_children():
            self._update_widget_color_recursive(child_widget)
            
    def _on_theme_change(self, event=None):
        selected_theme = self.current_theme.get()
        self._apply_theme(selected_theme)
        self._save_configuration()  

    def _toggle_log_size(self):
        if self.log_expanded:  
            self.log_text.pack_forget()  
            self.toggle_log_button.config(text="▲")  
            self.master.grid_rowconfigure(2, weight=0)  
            self.log_wrapper_frame.config(height=50)  
            self.log_expanded = False  
        else:  
            self.log_text.pack(padx=10, pady=(0, 10), fill="both", expand=True)  
            self.toggle_log_button.config(text="▼")  
            self.master.grid_rowconfigure(2, weight=1)  
            self.log_expanded = True  
        self._save_configuration()  

    def _browse_scripts_root_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.scripts_root_folder.set(folder_path)
            self._save_configuration()

    def _browse_folder(self, string_var):
        folder_path = filedialog.askdirectory()
        if folder_path:
            string_var.set(folder_path)

    def _browse_file(self, string_var, file_type):
        if file_type == "xlsx":
            file_types = [("Excel files", "*.xlsx"), ("All files", "*.*")]
        elif file_type == "csv":
            file_types = [("CSV files", "*.csv"), ("All files", "*.*")]
        elif file_type == "txt":
            file_types = [("Text files", "*.txt"), ("All files", "*.*")]
        else:
            file_types = [("All files", "*.*")]

        file_path = filedialog.askopenfilename(filetypes=file_types)
        if file_path:
            string_var.set(file_path)

    def _show_source_section(self):
        selected_source = self.source_type.get()
        print(f"DEBUG: Selected source type: {selected_source}", file=sys.stderr)

        for source, frame in self.source_sections.items():
            if source == selected_source:
                frame.pack(fill="both", expand=True, padx=0, pady=0)
            else:
                frame.pack_forget()
            
        self.master.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))


    def _show_input_method(self, tool_name, method):
        if tool_name == "check_psa":
            if method == "spreadsheet":
                self.check_psa_spreadsheet_frame.grid(row=1, column=0, columnspan=3, sticky="ew")
                self.check_psa_textbox_frame.grid_remove()
            else:
                self.check_psa_textbox_frame.grid(row=1, column=0, columnspan=3, sticky="nsew")
                self.check_psa_spreadsheet_frame.grid_remove()
        elif tool_name == "get_measurements":
            if method == "spreadsheet":
                self.get_measurements_spreadsheet_frame.grid(row=1, column=0, columnspan=3, sticky="ew")
                self.get_measurements_textbox_frame.grid_remove()
            else:
                self.get_measurements_textbox_frame.grid(row=1, column=0, columnspan=3, sticky="nsew")
                self.get_measurements_spreadsheet_frame.grid_remove()
        self.master.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))


    def _ensure_dir(self, path):
        """Ensures the directory for a given path exists. If path is a file, it ensures its parent directory exists."""
        directory = os.path.dirname(path) if os.path.isfile(path) or (os.path.basename(path) and '.' in os.path.basename(path)) else path
            
        if directory and not os.path.exists(directory):
            os.makedirs(directory)
            self.log_print(f"  Created directory: {directory}")

    def _download_and_compare_file(self, display_name, filename, download_url, local_target_folder):
        local_full_path = os.path.join(local_target_folder, filename)
        temp_file_path = local_full_path + ".tmp"
        
        self.log_print(f"Checking {display_name} ({filename})...")
        self.log_print(f"  Local path: {local_full_path}")
        self.log_print(f"  Download URL: {download_url}")

        try:
            response = requests.get(download_url, stream=True)
            response.raise_for_status()

            self._ensure_dir(local_full_path)  

            with open(temp_file_path, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)

            if os.path.exists(local_full_path):
                if filecmp.cmp(local_full_path, temp_file_path, shallow=False):
                    self.log_print(f"  '{filename}' is already up to date. No action needed.\n")
                    os.remove(temp_file_path)
                    return "skipped"
                else:
                    self.log_print(f"  New version of '{filename}' found. Updating...")
                    shutil.move(temp_file_path, local_full_path)
                    self.log_print(f"  '{filename}' updated successfully!\n")
                    return "updated"
            else:
                self.log_print(f"  '{filename}' not found locally. Downloading new script...")
                shutil.move(temp_file_path, local_full_path)
                self.log_print(f"  '{filename}' downloaded successfully!\n")
                return "downloaded"

        except requests.exceptions.RequestException as e:
            self.log_print(f"  ERROR downloading '{filename}' from '{download_url}': {e}\n", is_stderr=True)
            if os.path.exists(temp_file_path):
                os.remove(temp_file_path)
            return "error"
        except Exception as e:
            self.log_print(f"  An unexpected ERROR occurred while updating '{filename}': {e}\n", is_stderr=True)
            if os.path.exists(temp_file_path):
                os.remove(temp_file_path)
            return "error"

    def _update_all_scripts(self):
        scripts_folder = self.scripts_root_folder.get()
        if not scripts_folder or not os.path.isdir(scripts_folder):
            messagebox.showerror("Error", "Please set a valid 'Local Scripts Folder' first.")
            self.log_print("Script update aborted: 'Local Scripts Folder' is not set or invalid.\n", is_stderr=True)
            return

        self.log_print("\n--- Starting All Scripts Update Process ---")
        self.log_print(f"Using scripts root folder: {scripts_folder}\n")

        updated_count = 0
        downloaded_count = 0
        skipped_count = 0
        error_count = 0
        total_checked = 0

        for display_name, filename in SCRIPT_FILENAMES.items():
            if filename.endswith(".py") and filename in GITHUB_SCRIPT_URLS:
                total_checked += 1
                github_url = GITHUB_SCRIPT_URLS[filename]
                
                status = self._download_and_compare_file(display_name, filename, github_url, scripts_folder)
                
                if status == "updated":
                    updated_count += 1
                elif status == "downloaded":
                    downloaded_count += 1
                elif status == "skipped":
                    skipped_count += 1
                elif status == "error":
                    error_count += 1

        self.log_print("\n--- Script Update Process Complete ---")
        
        summary_message_parts = []
        if updated_count > 0:
            summary_message_parts.append(f"Updated {updated_count} script(s).")
        if downloaded_count > 0:
            summary_message_parts.append(f"Newly downloaded {downloaded_count} script(s).")
        if skipped_count > 0:
            summary_message_parts.append(f"{skipped_count} script(s) were already up to date.")
        if error_count > 0:
            summary_message_parts.append(f"{error_count} script(s) encountered errors.")

        if summary_message_parts:
            summary_message = "\n".join(summary_message_parts)
            current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")  
            self.last_update_timestamp.set(f"Last update: {current_time}")  
            messagebox.showinfo("Update Complete", f"Script update summary:\n{summary_message}\n\nCheck the Activity Log for full details.")
        elif total_checked == 0:
            messagebox.showinfo("Update Complete", "No Python scripts with valid GitHub URLs were found to check for updates.")
        else:
            messagebox.showinfo("Update Complete", "No scripts were updated, downloaded, or encountered errors. All checked scripts are already up to date.")
            
        self._save_configuration()

    def _check_for_gui_update(self):
        """Checks for a new version of the GUI script and updates/restarts if available."""
        self.log_print("\n--- Checking for GUI script update ---")
        local_gui_path = os.path.abspath(sys.argv[0])
        github_url = GITHUB_SCRIPT_URLS.get(GUI_SCRIPT_FILENAME)
        
        if not github_url:
            self.log_print("Error: GUI script URL not found in configuration.", is_stderr=True)
            messagebox.showerror("Update Error", "GUI script URL not configured.")
            return

        temp_download_path = local_gui_path + ".new_version_tmp"

        try:
            self.log_print(f"Downloading latest GUI from: {github_url}")
            response = requests.get(github_url, stream=True)
            response.raise_for_status()

            with open(temp_download_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
            
            if os.path.exists(local_gui_path) and filecmp.cmp(local_gui_path, temp_download_path, shallow=False):
                self.log_print("GUI script is already up to date.\n")
                os.remove(temp_download_path)
                messagebox.showinfo("Update Check", "The GUI is already up to date!")
                return
            else:
                self.log_print("New version of GUI script found. Applying update...")
                
                with open(UPDATE_IN_PROGRESS_MARKER, 'w') as f:
                    f.write(str(os.getpid()))

                shutil.copy(temp_download_path, local_gui_path)  
                os.remove(temp_download_path)

                current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                self.gui_last_update_timestamp.set(f"Last GUI update: {current_time}")
                self._save_configuration()

                self.log_print("GUI script updated successfully. Restarting application...\n")

                messagebox.showinfo("Update Complete", "The GUI has been updated. The application will now restart to apply changes.")
                
                self._restarting_for_update = True
                
                python = sys.executable
                os.execl(python, python, *sys.argv)  
        
        except requests.exceptions.RequestException as e:
            self.log_print(f"Error checking/downloading GUI update: {e}\n", is_stderr=True)
            messagebox.showerror("Update Error", f"Failed to check for GUI update: {e}")
        except Exception as e:
            self.log_print(f"An unexpected error occurred during GUI update: {e}\n", is_stderr=True)
            messagebox.showerror("Update Error", f"An unexpected error occurred: {e}")
        finally:
            if os.path.exists(temp_download_path):
                try:
                    os.remove(temp_download_path)
                except Exception as e:
                    self.log_print(f"Warning: Could not remove temporary download file: {e}", is_stderr=True)

    def _handle_startup_update_check(self):
        """Checks for and cleans up the update marker file on startup."""
        if os.path.exists(UPDATE_IN_PROGRESS_MARKER):
            try:
                os.remove(UPDATE_IN_PROGRESS_MARKER)
                self.log_print("GUI update completed successfully. Welcome back!", 'success')
            except Exception as e:
                self.log_print(f"Warning: Could not remove update marker file: {e}", is_stderr=True)

    def _download_renamer_excel(self):
        self.log_print("\n--- Downloading Renamer Excel Template ---")
        default_filename = "renaminator.xlsx"
        
        output_path = filedialog.asksaveasfilename(
             defaultextension=".xlsx",
             filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
             initialfile=default_filename
        )

        if not output_path:
            self.log_print("Download cancelled by user.\n")
            return

        self.log_print(f"Attempting to download from: {RENAMER_EXCEL_URL}")
        self.log_print(f"Saving to: {output_path}")

        try:
            response = requests.get(RENAMER_EXCEL_URL, stream=True)
            response.raise_for_status()

            output_dir = os.path.dirname(output_path)
            if output_dir:
                os.makedirs(output_dir, exist_ok=True)

            with open(output_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
            
            self.log_print(f"Renamer Excel file downloaded successfully to: {output_path}\n", is_stderr=False)
            messagebox.showinfo("Download Complete", f"Renamer Excel template downloaded successfully to:\n{output_path}")

        except requests.exceptions.RequestException as e:
            error_msg = f"Error downloading Renamer Excel file: {e}\n"
            self.log_print(error_msg, is_stderr=True)
            messagebox.showerror("Download Error", f"Failed to download Renamer Excel file.\nDetails: {e}")
        except Exception as e:
            error_msg = f"An unexpected error occurred during Excel download: {e}\n"
            self.log_print(error_msg, is_stderr=True)
            messagebox.showerror("Download Error", f"An unexpected error occurred.\nDetails: {e}")

    def _start_master_renamer_threaded(self, force_continue=False):
        self.master.config(cursor="wait")
        self.master.update_idletasks()

        self.log_text.configure(state='normal')
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state='disabled')
        self.log_print("Starting Renamer script...\n")

        thread = threading.Thread(target=self._run_master_renamer_in_thread, args=(force_continue,))
        thread.daemon = True
        thread.start()

    def _run_master_renamer_in_thread(self, force_continue):
        matrix_path = self.master_matrix_path.get()
        input_folder = self.rename_input_folder.get()
        vendor_code = self.vendor_code.get().strip()
        renaminator_script_path = os.path.join(self.scripts_root_folder.get(), SCRIPT_FILENAMES["Main Renaminator Script"])

        if not os.path.exists(renaminator_script_path):
            self.master.after(0, lambda: messagebox.showerror("Error", f"Main Renaminator Script not found: {renaminator_script_path}"))
            self.master.after(0, self._enable_renamer_button)
            return
        if not matrix_path or not os.path.exists(matrix_path):
            self.master.after(0, lambda: messagebox.showerror("Error", "Please select a valid Renamer Matrix (.xlsx)."))
            self.master.after(0, self._enable_renamer_button)
            return
        if not input_folder or not os.path.exists(input_folder):
            self.master.after(0, lambda: messagebox.showerror("Error", "Please select a valid Input Images Folder."))
            self.master.after(0, self._enable_renamer_button)
            return
        if not vendor_code:
            self.master.after(0, lambda: messagebox.showerror("Error", "Please enter a Vendor Code."))
            self.master.after(0, self._enable_renamer_button)
            return

        self.log_print(f"--- Executing Renamer Script (Force Continue: {force_continue}) ---\n")
        
        args = ['--matrix', matrix_path, '--input', input_folder, '--vendor_code', vendor_code]
        if force_continue:
            args.append('--force_continue')
            
        success, _ = run_script_wrapper(renaminator_script_path, True, args, self.log_text,
                                         progress_bar=None, progress_label=None,
                                         run_button_wrapper=None, progress_wrapper=None,
                                         success_callback=lambda output: self._process_renamer_result(True, force_continue, output),
                                         error_callback=lambda output: self._process_renamer_result(False, force_continue, output))

    def _process_renamer_result(self, success, was_forced_attempt, output):
        if success:
            messagebox.showinfo("Success", "Master Renamer script completed successfully!")
        else:
            if not was_forced_attempt:
                response = messagebox.askyesno(
                    "Renamer Failed",
                    "The Master Renamer script failed. This might be due to non-JPG files, missing files, or other issues. Would you like to try running it again with the '--force_continue' option enabled?\n\n(Check the Activity Log for details on the previous run.)"
                )
                if response:
                    self._start_master_renamer_threaded(force_continue=True)
                    return
                else:
                    messagebox.showerror("Renamer Failed", "Master Renamer script failed. Please check the log for details.")
            else:
                    messagebox.showerror("Renamer Failed", "Master Renamer script failed, even with '--force_continue' enabled. Please check the log for details.")
        
        self._enable_renamer_button()


    def _enable_renamer_button(self):
        self.master.config(cursor="")


    def _start_inline_copy(self):
        network_folder = self.inline_source_folder.get()
        matrix_path = self.inline_matrix_path.get()
        output_folder = self.inline_output_folder.get()
        copier_script_path = os.path.join(self.scripts_root_folder.get(), SCRIPT_FILENAMES["File Copier Script"])

        if not os.path.exists(copier_script_path):
            messagebox.showerror("Error", f"File Copier Script not found: {copier_script_path}")
            return
        if not network_folder or not os.path.exists(network_folder):
            messagebox.showerror("Error", "Please select a valid Source Folder (Network Assets).")
            return
        if not matrix_path or not os.path.exists(matrix_path):
            messagebox.showerror("Error", "Please select a valid Renamer Matrix (with Filenames).")
            return
        if not output_folder:
            messagebox.showerror("Error", "Please select an Output Folder for Copied Images.")
            return
        os.makedirs(output_folder, exist_ok=True)

        self.log_print(f"\n--- Starting Inline Project Copy (using File Copier Script) ---")
        args = ['--matrix', matrix_path, '--input', network_folder, '--output', output_folder]
        
        def inline_copy_success_callback(output):
            self.run_inline_copy_button.config(state='normal')
            messagebox.showinfo("Success", "Inline Project Copy completed successfully!")
        def inline_copy_error_callback(output):
            self.run_inline_copy_button.config(state='normal')
            messagebox.showerror("Error", "Inline Project Copy failed. Please check the log for details.")

        self.run_inline_copy_button.config(state='disabled')

        success, _ = run_script_wrapper(copier_script_path, True, args, self.log_text,
                                         self.inline_copy_progress_bar, self.inline_copy_progress_label,
                                         self.inline_copy_run_button_wrapper, self.inline_copy_progress_wrapper,
                                         inline_copy_success_callback, inline_copy_error_callback,
                                         initial_progress_text="Copying Images...")


    def _start_pso1_download(self):
        matrix_path = self.pso1_matrix_path.get()
        output_folder = self.pso1_output_folder.get()
        downloader_script_path = os.path.join(self.scripts_root_folder.get(), SCRIPT_FILENAMES["Downloader Script"])

        if not os.path.exists(downloader_script_path):
            messagebox.showerror("Error", f"Downloader Script not found: {downloader_script_path}")
            return
        if not matrix_path or not os.path.exists(matrix_path):
            messagebox.showerror("Error", "Please select a valid Renamer Matrix (with URLs).")
            return
        if not output_folder:
            messagebox.showerror("Error", "Please select an Output Folder for Downloaded Images.")
            return
        os.makedirs(output_folder, exist_ok=True)
        
        self.log_print(f"\n--- Starting PSO Option 1 Download ---")
        args = ['--matrix', matrix_path, '--output', output_folder]
        
        def pso1_download_success_callback(output):
            self.run_pso1_download_button.config(state='normal')
            messagebox.showinfo("Success", "Download (PSO Option 1) completed successfully!")
        def pso1_download_error_callback(output):
            self.run_pso1_download_button.config(state='normal')
            messagebox.showerror("Error", "Download (PSO Option 1) failed. Please check the log for details.")

        self.run_pso1_download_button.config(state='disabled')

        success, _ = run_script_wrapper(downloader_script_path, True, args, self.log_text,
                                         self.pso1_download_progress_bar, self.pso1_download_progress_label,
                                         self.pso1_download_run_button_wrapper, self.pso1_download_progress_wrapper,
                                         pso1_download_success_callback, pso1_download_error_callback,
                                         initial_progress_text="Downloading Images...")

    def _start_pso2_copy(self):
        network_folder = self.pso2_network_folder.get()
        matrix_path = self.pso2_matrix_path.get()
        output_folder = self.pso2_output_folder.get()
        copier_script_path = os.path.join(self.scripts_root_folder.get(), SCRIPT_FILENAMES["File Copier Script"])

        if not os.path.exists(copier_script_path):
            messagebox.showerror("Error", f"File Copier Script not found: {copier_script_path}")
            return
        if not network_folder or not os.path.exists(network_folder):
            messagebox.showerror("Error", "Please select a valid Network Assets Source Folder.")
            return
        if not matrix_path or not os.path.exists(matrix_path):
            messagebox.showerror("Error", "Please select a valid Renamer Matrix (with Filenames).")
            return
        if not output_folder:
            messagebox.showerror("Error", "Please select an Output Folder for Copied Images.")
            return
        os.makedirs(output_folder, exist_ok=True)

        self.log_print(f"\n--- Starting PSO Option 2 Copy ---")
        args = ['--matrix', matrix_path, '--input', network_folder, '--output', output_folder]
        
        def pso2_copy_success_callback(output):
            self.run_pso2_copy_button.config(state='normal')
            messagebox.showinfo("Success", "Copy (PSO Option 2) completed successfully!")
        def pso2_copy_error_callback(output):
            self.run_pso2_copy_button.config(state='normal')
            messagebox.showerror("Error", "Copy (PSO Option 2) failed. Please check the log for details.")

        self.run_pso2_copy_button.config(state='disabled')

        success, _ = run_script_wrapper(copier_script_path, True, args, self.log_text,
                                         self.pso2_copy_progress_bar, self.pso2_copy_progress_label,
                                         self.pso2_copy_run_button_wrapper, self.pso2_copy_progress_wrapper,
                                         pso2_copy_success_callback, pso2_copy_error_callback,
                                         initial_progress_text="Copying Images...")

    def _run_cropping_script(self, script_filename):
        input_folder = self.prep_input_path.get()  

        if not input_folder or not os.path.isdir(input_folder):
            messagebox.showerror("Error", "Cropping scripts require a valid *folder* for preparation. Please select a folder.")
            return
        if not os.path.exists(input_folder):
            messagebox.showerror("Error", f"Images to Crop with Scripts folder not found: {input_folder}")
            return

        cropping_script_path = os.path.join(self.scripts_root_folder.get(), script_filename)

        if not os.path.exists(cropping_script_path):
            messagebox.showerror("Error", f"Cropping script '{script_filename}' not found: {cropping_script_path}")
            return
        
        def cropping_success_callback(output):
            self.cropping_run_button_wrapper.grid(row=0, column=1)
            self.cropping_progress_wrapper.grid_remove()
            messagebox.showinfo("Success", f"Cropping with {script_filename} completed successfully!")

        def cropping_error_callback(output):
            self.cropping_run_button_wrapper.grid(row=0, column=1)
            self.cropping_progress_wrapper.grid_remove()
            messagebox.showerror("Error", f"Cropping with {script_filename} failed. Please check the log for details.")

        self.cropping_run_button_wrapper.grid_remove()
        self.cropping_progress_wrapper.grid(row=0, column=1)

        self.log_print(f"\n--- Running Cropping Script: {script_filename} ---")
        args = ['--input', input_folder]  

        success, _ = run_script_wrapper(cropping_script_path, True, args, self.log_text,
                                         self.cropping_progress_bar, self.cropping_progress_label,
                                         self.cropping_run_button_wrapper, self.cropping_progress_wrapper,
                                         cropping_success_callback, cropping_error_callback,
                                         initial_progress_text=f"Cropping {script_filename.split('_')[0].replace('reformat', '').upper()}...")


    def _run_bynder_metadata_prep(self):
        assets_folder = self.bynder_assets_folder.get()

        if not assets_folder or not os.path.isdir(assets_folder):
            messagebox.showerror("Input Error", "Please select a valid folder containing assets for Bynder metadata preparation.")
            return
        if not os.path.exists(assets_folder):
            messagebox.showerror("Input Error", f"Assets folder not found: {assets_folder}")
            return

        bynder_script_name = SCRIPT_FILENAMES["Bynder Metadata Prep"]
        bynder_script_path = os.path.join(self.scripts_root_folder.get(), bynder_script_name)

        if not os.path.exists(bynder_script_path):
            messagebox.showerror("Error", f"Bynder Metadata Prep script not found: {bynder_script_path}\n"
                                          f"Please ensure '{bynder_script_name}' is in your scripts folder.")
            return

        self.log_print(f"\n--- Running Bynder Metadata Prep Script ({bynder_script_name}) ---")
        self.log_print(f"Preparing metadata for assets in: {assets_folder}")

        args = [
            '--input', assets_folder
        ]
        
        def bynder_prep_success_callback(output):
            self.run_bynder_prep_button.config(state='normal')
            messagebox.showinfo("Success", "Bynder Metadata Prep script completed successfully!\n"
                                           "The metadata importer CSV should be in your downloads folder.")
        def bynder_prep_error_callback(output):
            self.run_bynder_prep_button.config(state='normal')
            messagebox.showerror("Error", "Bynder Metadata Prep script failed. Please check the log for details.")

        self.run_bynder_prep_button.config(state='disabled')

        success, _ = run_script_wrapper(bynder_script_path, True, args, self.log_text,
                                         self.bynder_prep_progress_bar, self.bynder_prep_progress_label,  
                                         self.bynder_prep_run_button_wrapper, self.bynder_prep_progress_wrapper,  
                                         bynder_prep_success_callback, bynder_prep_error_callback,
                                         initial_progress_text="Preparing Metadata...")


    def _get_skus_from_input(self, input_type_var, spreadsheet_path_var, text_widget, file_prefix="skus_"):
        """
        Helper to get SKUs/filenames either from a spreadsheet (returns path) or textbox.
        If from textbox, it writes the content to a temporary .txt file and returns its path.
        Returns (data, is_file_path) tuple.
        """
        
        if input_type_var.get() == "spreadsheet":
            input_path = spreadsheet_path_var.get()
            if not input_path or not os.path.exists(input_path) or not input_path.lower().endswith('.xlsx'):
                messagebox.showerror("Input Error", "Please select a valid SKU Spreadsheet (.xlsx).")
                return None, False
            self.log_print(f"Reading SKUs/filenames from spreadsheet: {input_path}")
            return input_path, True

        elif input_type_var.get() == "textbox":
            raw_text = text_widget.get("1.0", tk.END).strip()
            if not raw_text:
                messagebox.showerror("Input Error", "Please paste SKUs/filenames into the text box.")
                return None, False
            
            temp_fd, temp_file_path = tempfile.mkstemp(suffix=".txt", prefix=file_prefix, dir=tempfile.gettempdir())
            os.close(temp_fd)

            try:
                content_to_write = "\n".join(line.strip() for line in raw_text.splitlines() if line.strip())
                with open(temp_file_path, "w", encoding="utf-8") as f:
                    f.write(content_to_write)
                self.log_print(f"Content from text box written to temporary file: {temp_file_path}")
                return temp_file_path, True
            except Exception as e:
                messagebox.showerror("File Error", f"Failed to write temporary file: {e}")
                self.log_print(f"Error writing temporary file: {e}\n", is_stderr=True)
                if os.path.exists(temp_file_path):
                    os.remove(temp_file_path)
                return None, False
                
        return None, False


    def _run_check_psas_script(self):
        scripts_folder = self.scripts_root_folder.get()
        check_psas_script_name = SCRIPT_FILENAMES["Check Bynder PSAs script"]
        check_psas_script_path = os.path.join(scripts_folder, check_psas_script_name)

        if not os.path.exists(check_psas_script_path):
            messagebox.showerror("Error", f"Check Bynder PSAs script not found: {check_psas_script_path}\n"
                                          f"Please ensure '{check_psas_script_name}' is in your scripts folder.")
            return
        
        sku_input_data, is_file_path = self._get_skus_from_input(  
            self.check_psa_input_type,  
            self.check_psa_sku_spreadsheet_path,  
            self.check_psa_text_widget,
            file_prefix="check_psas_skus_"
        )
        if sku_input_data is None:
            return

        self.log_print(f"\n--- Running Check Bynder PSAs Script ({check_psas_script_name}) ---")
        
        args = []
        self.log_print(f"Passing SKU input file: {sku_input_data}")
        args.extend(["--sku_file", sku_input_data])
            
        def check_psas_success_callback(output):
            self.run_check_psas_button.config(state='normal')
            messagebox.showinfo("Success", "Check Bynder PSAs script completed successfully!\n"
                                           "Results should be in your downloads folder.")
            if self.check_psa_input_type.get() == "textbox" and is_file_path and os.path.exists(sku_input_data):
                try:
                    os.remove(sku_input_data)
                except Exception as e:
                    self.log_print(f"Warning: Could not remove temporary file {sku_input_data}: {e}\n", is_stderr=True)

        def check_psas_error_callback(output):
            self.run_check_psas_button.config(state='normal')
            messagebox.showerror("Error", "Check Bynder PSAs script failed. Please check the log for details.")
            if self.check_psa_input_type.get() == "textbox" and is_file_path and os.path.exists(sku_input_data):
                try:
                    os.remove(sku_input_data)
                except Exception as e:
                    self.log_print(f"Warning: Could not remove temporary file {sku_input_data}: {e}\n", is_stderr=True)

        self.run_check_psas_button.config(state='disabled')

        success, _ = run_script_wrapper(check_psas_script_path, True, args, self.log_text,
                                         self.check_psas_progress_bar, self.check_psas_progress_label,
                                         self.check_psas_run_button_wrapper, self.check_psas_progress_wrapper,
                                         check_psas_success_callback, check_psas_error_callback,
                                         initial_progress_text="Checking Bynder PSAs...")


    def _run_download_psas_script(self):
        scripts_folder = self.scripts_root_folder.get()
        download_psas_script_name = SCRIPT_FILENAMES["Download PSAs script"]
        download_psas_script_path = os.path.join(scripts_folder, download_psas_script_name)

        if not os.path.exists(download_psas_script_path):
            messagebox.showerror("Error", f"Download PSAs script not found: {download_psas_script_path}\n"
                                          f"Please ensure '{download_psas_script_name}' is in your scripts folder.")
            return
        
        sku_input_data, is_file_path = self._get_skus_from_input(
            self.download_psa_input_type,
            self.download_psa_sku_spreadsheet_path,
            self.download_psa_text_widget,
            file_prefix="download_psas_skus_"
        )
        if sku_input_data is None:
            return

        output_folder_path = self.download_psa_output_folder.get()
        if not output_folder_path:
            messagebox.showerror("Error", "Please select an Output Folder for Download PSAs.")
            return
        
        os.makedirs(output_folder_path, exist_ok=True)
        
        selected_image_types = []
        if self.download_psa_grid.get(): selected_image_types.append("grid")
        if self.download_psa_100.get(): selected_image_types.append("100")
        if self.download_psa_200.get(): selected_image_types.append("200")
        if self.download_psa_300.get(): selected_image_types.append("300")
        if self.download_psa_400.get(): selected_image_types.append("400")
        if self.download_psa_500.get(): selected_image_types.append("500")
        if self.download_psa_dimension.get(): selected_image_types.append("dimension")
        if self.download_psa_swatch.get(): selected_image_types.append("swatch")
        if self.download_psa_5000.get(): selected_image_types.append("5000")
        if self.download_psa_5100.get(): selected_image_types.append("5100")
        if self.download_psa_5200.get(): selected_image_types.append("5200")
        if self.download_psa_5300.get(): selected_image_types.append("5300")
        if self.download_psa_squareThumbnail.get(): selected_image_types.append("squareThumbnail")


        image_types_arg = ",".join(selected_image_types)

        self.log_print(f"\n--- Running Download PSAs Script ({download_psas_script_name}) ---")
        self.log_print(f"Output folder provided: {output_folder_path}")
        if image_types_arg:
            self.log_print(f"Image types requested via UI: {image_types_arg}")
        else:
            self.log_print("No specific image types selected in UI. Script might default or prompt.")

        args = []
        self.log_print(f"Passing SKU input file: {sku_input_data}")
        args.extend(["--sku_file", sku_input_data])

        args.extend(["--output_folder", output_folder_path])
        if image_types_arg:
            args.extend(["--image_types", image_types_arg])
            
        def download_success_callback(output):
            self.run_download_psas_button.config(state='normal')
            messagebox.showinfo("Success", f"Download PSAs script completed successfully!\n"
                                           f"Results are in the selected output folder: {output_folder_path}")
            if self.download_psa_input_type.get() == "textbox" and is_file_path and os.path.exists(sku_input_data):
                try:
                    os.remove(sku_input_data)
                except Exception as e:
                    self.log_print(f"Warning: Could not remove temporary file {sku_input_data}: {e}\n", is_stderr=True)

        def download_error_callback(output):
            self.run_download_psas_button.config(state='normal')
            messagebox.showerror("Error", "Download PSAs script failed. Please check the log for details.")
            if self.download_psa_input_type.get() == "textbox" and is_file_path and os.path.exists(sku_input_data):
                try:
                    os.remove(sku_input_data)
                except Exception as e:
                    self.log_print(f"Warning: Could not remove temporary file {sku_input_data}: {e}\n", is_stderr=True)

        self.run_download_psas_button.config(state='disabled')

        run_script_wrapper(download_psas_script_path, True, args, self.log_text,  
                                       self.download_psas_progress_bar, self.download_psas_progress_label,
                                       self.download_psas_run_button_wrapper,  
                                       self.download_psas_progress_wrapper,  
                                       download_success_callback, download_error_callback,
                                       initial_progress_text="Downloading...")

    def _select_all_psas(self):
        """Sets all Download PSA checkboxes to True."""
        for var in self.download_psa_checkboxes:
            var.set(True)

    def _clear_all_psas(self):
        """Sets all Download PSA checkboxes to False."""
        for var in self.download_psa_checkboxes:
            var.set(False)

    def _run_get_measurements_script(self):
        scripts_folder = self.scripts_root_folder.get()
        get_measurements_script_name = SCRIPT_FILENAMES["Get Measurements script"]
        get_measurements_script_path = os.path.join(scripts_folder, get_measurements_script_name)

        if not os.path.exists(get_measurements_script_path):
            messagebox.showerror("Error", f"Get Measurements script not found: {get_measurements_script_path}\n"
                                          f"Please ensure '{get_measurements_script_name}' is in your scripts folder.")
            return

        sku_input_data, is_file_path = self._get_skus_from_input(
            self.get_measurements_input_type,  
            self.get_measurements_sku_spreadsheet_path,  
            self.get_measurements_text_widget,
            file_prefix="get_measurements_skus_"
        )
        if sku_input_data is None:
            return

        output_location_message = ""
        output_folder_for_script = ""

        if self.get_measurements_input_type.get() == "spreadsheet":
            output_folder_for_script = os.path.dirname(sku_input_data)
            output_location_message = f"Results should be in the same folder as your spreadsheet: {output_folder_for_script}"
            self.log_print(f"SKU input from spreadsheet: {sku_input_data}")
        else:
            output_folder_for_script = os.path.join(os.path.expanduser("~"), "Downloads")
            output_location_message = "Results should be in your Downloads folder."
            self.log_print(f"SKU input from text box (now temp file): {sku_input_data}")

        self._ensure_dir(output_folder_for_script)

        self.log_print(f"\n--- Running Get Measurements Script ({get_measurements_script_name}) ---")
        self.log_print(f"Output will be saved to: {output_folder_for_script}")
        self.log_print("NOTE: The script will use its default or hardcoded paths for STEP exports.")

        args = []
        args.extend(["--sku_list_file", sku_input_data])

        args.extend(["--output_folder", output_folder_for_script])
            
        def get_measurements_success_callback(output):
            self.run_get_measurements_button.config(state='normal')
            messagebox.showinfo("Success", f"Get Measurements script completed successfully!\n"
                                           f"{output_location_message}")
            if self.get_measurements_input_type.get() == "textbox" and is_file_path and os.path.exists(sku_input_data):
                try:
                    os.remove(sku_input_data)
                except Exception as e:
                    self.log_print(f"Warning: Could not remove temporary file {sku_input_data}: {e}\n", is_stderr=True)

        def get_measurements_error_callback(output):
            self.run_get_measurements_button.config(state='normal')
            messagebox.showerror("Error", "Get Measurements script failed. Please check the log for details.")
            if self.get_measurements_input_type.get() == "textbox" and is_file_path and os.path.exists(sku_input_data):
                try:
                    os.remove(sku_input_data)
                except Exception as e:
                    self.log_print(f"Warning: Could not remove temporary file {sku_input_data}: {e}\n", is_stderr=True)

        self.run_get_measurements_button.config(state='disabled')

        run_script_wrapper(get_measurements_script_path, True, args, self.log_text,
                                       self.get_measurements_progress_bar, self.get_measurements_progress_label,
                                       self.get_measurements_run_button_wrapper,  
                                       self.get_measurements_progress_wrapper,  
                                       get_measurements_success_callback, get_measurements_error_callback,
                                       initial_progress_text="Getting Measurements...")

    def _run_bynder_metadata_convert_script(self):
        input_csv_path = self.bynder_metadata_csv_path.get()
        scripts_folder = self.scripts_root_folder.get()
        convert_script_name = SCRIPT_FILENAMES["Convert Bynder Metadata to XLS"]
        convert_script_path = os.path.join(scripts_folder, convert_script_name)
        
        output_folder = os.path.join(os.path.expanduser("~"), "Downloads")
        os.makedirs(output_folder, exist_ok=True)


        if not os.path.exists(convert_script_path):
            messagebox.showerror("Error", f"Bynder Metadata Conversion script not found: {convert_script_path}\n"
                                          f"Please ensure '{convert_script_name}' is in your scripts folder.")
            return
        if not input_csv_path or not os.path.exists(input_csv_path) or not input_csv_path.lower().endswith('.csv'):
            messagebox.showerror("Input Error", "Please select a valid Bynder Metadata CSV file (.csv).")
            return

        self.log_print(f"\n--- Running Bynder Metadata Conversion Script ({convert_script_name}) ---")
        self.log_print(f"Input CSV: {input_csv_path}")
        self.log_print(f"Output to: {output_folder}")

        args = [input_csv_path, output_folder]

        def convert_success_callback(output):
            self.run_bynder_metadata_convert_button.config(state='normal')
            messagebox.showinfo("Success", f"Bynder Metadata CSV converted successfully!\n"
                                           f"The converted Excel file is in your Downloads folder.")
        def convert_error_callback(output):
            self.run_bynder_metadata_convert_button.config(state='normal')
            messagebox.showerror("Error", "Bynder Metadata conversion failed. Please check the log for details.")

        self.run_bynder_metadata_convert_button.config(state='disabled')

        run_script_wrapper(convert_script_path, True, args, self.log_text,
                                       self.bynder_metadata_convert_progress_bar, self.bynder_metadata_convert_progress_label,
                                       self.bynder_metadata_convert_run_button_wrapper,
                                       self.bynder_metadata_convert_progress_wrapper,
                                       convert_success_callback, convert_error_callback,
                                       initial_progress_text="Converting CSV to XLS...")

    def _run_move_files_script(self):
        scripts_folder = self.scripts_root_folder.get()
        move_script_name = SCRIPT_FILENAMES["Move Files from Spreadsheet"]
        move_script_path = os.path.join(scripts_folder, move_script_name)

        if not os.path.exists(move_script_path):
            messagebox.showerror("Error", f"Move Files script not found: {move_script_path}\n"
                                          f"Please ensure '{move_script_name}' is in your scripts folder.")
            return

        source_folder = self.move_files_source_folder.get()
        destination_folder = self.move_files_destination_folder.get()

        if not source_folder or not os.path.isdir(source_folder):
            messagebox.showerror("Input Error", "Please select a valid Source Folder.")
            return
        if not destination_folder:
            messagebox.showerror("Input Error", "Please select a Destination Folder.")
            return
        os.makedirs(destination_folder, exist_ok=True)

        file_input_data, is_file_path = self._get_skus_from_input(
            self.move_files_input_type,
            self.move_files_excel_path,
            self.move_files_text_widget,
            file_prefix="move_files_filenames_"
        )
        if file_input_data is None:
            return

        self.log_print(f"\n--- Running Move Files Script ({move_script_name}) ---")
        self.log_print(f"Source Folder: {source_folder}")
        self.log_print(f"Destination Folder: {destination_folder}")

        args = ["--source_folder", source_folder, "--destination_folder", destination_folder]
        self.log_print(f"Passing filenames input file: {file_input_data}")
        args.extend(["--filenames_file", file_input_data])
        
        def move_files_success_callback(output):
            self.run_move_files_button.config(state='normal')
            total_attempted = 0
            moved_count = 0
            for line in output.splitlines():
                if "Total files attempted:" in line:
                    try:
                        total_attempted = int(line.split(":")[1].strip())
                    except ValueError:
                        pass
                if "Files successfully moved:" in line:
                    try:
                        moved_count = int(line.split(":")[1].strip())
                    except ValueError:
                        pass

            if total_attempted > 0 and moved_count == 0:
                messagebox.showwarning("No Files Moved", "The script completed, but 0 files were successfully moved. This might mean the files listed were not found in the source folder. Please check the Activity Log for details.")
            elif total_attempted == 0:
                messagebox.showinfo("No Files Specified", "The script completed, but no files were specified in the input Excel or textbox.")
            else:
                messagebox.showinfo("Success", f"Move Files script completed successfully! {moved_count} of {total_attempted} files moved.")
            
            if self.move_files_input_type.get() == "textbox" and is_file_path and os.path.exists(file_input_data):
                try:
                    os.remove(file_input_data)
                except Exception as e:
                    self.log_print(f"Warning: Could not remove temporary file {file_input_data}: {e}\n", is_stderr=True)
        
        def move_files_error_callback(output):
            self.run_move_files_button.config(state='normal')
            messagebox.showerror("Error", "Move Files script failed. Please check the log for details.")
            if self.move_files_input_type.get() == "textbox" and is_file_path and os.path.exists(file_input_data):
                try:
                    os.remove(file_input_data)
                except Exception as e:
                    self.log_print(f"Warning: Could not remove temporary file {file_input_data}: {e}\n", is_stderr=True)

        self.run_move_files_button.config(state='disabled')

        run_script_wrapper(move_script_path, True, args, self.log_text,
                                       self.move_files_progress_bar, self.move_files_progress_label,
                                       self.move_files_run_button_wrapper,
                                       self.move_files_progress_wrapper,
                                       move_files_success_callback, move_files_error_callback,
                                       initial_progress_text="Moving Files...")

    def _run_or_boolean_script(self):
        scripts_folder = self.scripts_root_folder.get()
        or_script_name = SCRIPT_FILENAMES["OR Boolean Search Creator"]
        or_script_path = os.path.join(scripts_folder, or_script_name)

        if not os.path.exists(or_script_path):
            messagebox.showerror("Error", f"OR Boolean Search Creator script not found: {or_script_path}\n"
                                          f"Please ensure '{or_script_name}' is in your scripts folder.")
            return

        input_data, is_file_path = self._get_skus_from_input(
            self.or_boolean_input_type,
            self.or_boolean_spreadsheet_path,
            self.or_boolean_text_widget,
            file_prefix="or_boolean_skus_"
        )
        if input_data is None:
            return

        self.log_print(f"\n--- Running OR Boolean Search Creator Script ({or_script_name}) ---")
        self.log_print(f"Input source: {'Spreadsheet' if self.or_boolean_input_type.get() == 'spreadsheet' else 'Text Box'}")
        self.log_print(f"Input file: {input_data}")

        args = [input_data]

        def or_boolean_success_callback(full_output):
            self.run_or_boolean_button.config(state='normal')
            
            # Update the dedicated results textbox
            self.or_boolean_results_textbox.configure(state='normal')
            self.or_boolean_results_textbox.delete("1.0", tk.END)
            self.or_boolean_results_textbox.insert(tk.END, full_output)
            self.or_boolean_results_textbox.configure(state='disabled')
            self.or_boolean_results_textbox.see(tk.END)

            messagebox.showinfo("Success", "OR Boolean Search Creator script completed successfully! The result is displayed in the textbox.")
            
            if self.or_boolean_input_type.get() == "textbox" and is_file_path and os.path.exists(input_data):
                try:
                    os.remove(input_data)
                except Exception as e:
                    self.log_print(f"Warning: Could not remove temporary file {input_data}: {e}\n", is_stderr=True)

        def or_boolean_error_callback(full_output):
            self.run_or_boolean_button.config(state='normal')
            # Update the dedicated results textbox with error info
            self.or_boolean_results_textbox.configure(state='normal')
            self.or_boolean_results_textbox.delete("1.0", tk.END)
            self.or_boolean_results_textbox.insert(tk.END, full_output, 'error')
            self.or_boolean_results_textbox.configure(state='disabled')
            self.or_boolean_results_textbox.see(tk.END)
            
            messagebox.showerror("Error", "OR Boolean Search Creator script failed. Please check the log for details.")
            
            if self.or_boolean_input_type.get() == "textbox" and is_file_path and os.path.exists(input_data):
                try:
                    os.remove(input_data)
                except Exception as e:
                    self.log_print(f"Warning: Could not remove temporary file {input_data}: {e}\n", is_stderr=True)

        self.run_or_boolean_button.config(state='disabled')

        run_script_wrapper(or_script_path, True, args, self.log_text,
                                       self.or_boolean_progress_bar, self.or_boolean_progress_label,
                                       self.or_boolean_run_button_wrapper,
                                       self.or_boolean_progress_wrapper,
                                       or_boolean_success_callback, or_boolean_error_callback,
                                       initial_progress_text="Creating OR Boolean Search...")

    def _setup_initial_state(self):
        self._show_source_section()  
        self._show_input_method("check_psa", self.check_psa_input_type.get())
        self._show_input_method_download_psa(self.download_psa_input_type.get())
        self._show_input_method("get_measurements", self.get_measurements_input_type.get())
        self._show_input_method_move_files(self.move_files_input_type.get())
        self._show_input_method_or_boolean(self.or_boolean_input_type.get())
        if not self.log_expanded:  
            self.log_text.pack_forget()  
            self.toggle_log_button.config(text="▲")  
            self.master.grid_rowconfigure(2, weight=0)  
            self.log_wrapper_frame.config(height=50)  
        
    def _create_widgets(self):
        self.master.grid_rowconfigure(0, weight=0)
        self.master.grid_rowconfigure(1, weight=2)
        self.master.grid_rowconfigure(2, weight=1)
        self.master.grid_rowconfigure(3, weight=0)
        self.master.grid_columnconfigure(0, weight=1)  

        top_bar_frame = ttk.Frame(self.master, style='TFrame')
        top_bar_frame.grid(row=0, column=0, padx=(10, 10), pady=(2, 2), sticky="new")  
        
        top_bar_frame.grid_columnconfigure(0, weight=1)
        top_bar_frame.grid_columnconfigure(1, weight=1)
        top_bar_frame.grid_columnconfigure(2, weight=0)

        update_all_scripts_section = ttk.Frame(top_bar_frame, style='TFrame')
        update_all_scripts_section.grid(row=0, column=0, padx=(0, 10), sticky="w")
        update_all_scripts_section.grid_columnconfigure(0, weight=1)
        
        self.update_all_scripts_button = ttk.Button(update_all_scripts_section, text="Update All Scripts", command=self._update_all_scripts, style='TButton')
        self.update_all_scripts_button.pack(fill="x", expand=True)
        Tooltip(self.update_all_scripts_button, "Checks GitHub for updated versions of Python scripts and downloads them to your local scripts folder if newer versions are available. If a script is missing, it will download it.", self.secondary_bg, self.text_color)
        
        self.last_update_label = ttk.Label(update_all_scripts_section, textvariable=self.last_update_timestamp, style='TLabel')
        self.last_update_label.pack(pady=(2,0))


        update_gui_section = ttk.Frame(top_bar_frame, style='TFrame')
        update_gui_section.grid(row=0, column=1, padx=(0, 10), sticky="w")
        update_gui_section.grid_columnconfigure(0, weight=1)

        self.check_gui_update_button = ttk.Button(update_gui_section, text="Update GUI", command=self._check_for_gui_update, style='TButton')
        self.check_gui_update_button.pack(fill="x", expand=True)
        Tooltip(self.check_gui_update_button, "Checks for and applies updates to this GUI application itself, then restarts.", self.secondary_bg, self.text_color)
        
        self.gui_last_update_label = ttk.Label(update_gui_section, textvariable=self.gui_last_update_timestamp, style='TLabel')
        self.gui_last_update_label.pack(pady=(2,0))


        theme_frame = ttk.Frame(top_bar_frame, style='TFrame')
        theme_frame.grid(row=0, column=2, sticky="e")  
        
        self.theme_label = ttk.Label(theme_frame, text="Theme:", style='TLabel')
        self.theme_label.pack(side="left", padx=(0, 5))
        
        self.theme_selector = ttk.Combobox(theme_frame, textvariable=self.current_theme,  
                                         values=["Light", "Dark"], state="readonly", width=6)
        self.theme_selector.pack(side="left")
        self.theme_selector.bind("<<ComboboxSelected>>", self._on_theme_change)
        

        container = ttk.Frame(self.master)
        container.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")  

        self.canvas = tk.Canvas(container, highlightthickness=0, bg=self.primary_bg)  
        self.canvas.pack(side="left", fill="both", expand=True)  

        scrollbar = ttk.Scrollbar(container, orient="vertical", command=self.canvas.yview)
        scrollbar.pack(side="right", fill="y")
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.scrollable_frame = ttk.Frame(self.canvas, style='TFrame')  
        canvas_frame_id = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        
        self.scrollable_frame.grid_columnconfigure(0, weight=1)

        def on_frame_configure(event):
            canvas_width = event.width
            self.canvas.itemconfig(canvas_frame_id, width=canvas_width)  
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))

        self.scrollable_frame.bind("<Configure>", on_frame_configure)
        self.canvas.bind("<Configure>", lambda event: self.canvas.itemconfig(canvas_frame_id, width=event.width))

        def _on_mouse_wheel(event):
            if sys.platform == "darwin":  
                self.canvas.yview_scroll(int(-1*(event.delta)), "units")
            else:  
                self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

        self.canvas.bind_all("<MouseWheel>", _on_mouse_wheel)


        row_counter = 0  

        # SECTION: Local Scripts Folder Path
        scripts_folder_wrapper_frame = ttk.Frame(self.scrollable_frame, style='SectionFrame.TFrame')
        scripts_folder_wrapper_frame.grid(row=row_counter, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
        row_counter += 1

        header_sub_frame_scripts = ttk.Frame(scripts_folder_wrapper_frame, style='TFrame')
        header_sub_frame_scripts.pack(side="top", fill="x", pady=(0, 5), padx=0)
        header_label_scripts = ttk.Label(header_sub_frame_scripts, text="Local Scripts Folder", style='Header.TLabel')
        header_label_scripts.pack(side="left", padx=(0, 5))
        info_label_scripts = ttk.Label(header_sub_frame_scripts, text=" ⓘ", font=self.base_font)  
        Tooltip(info_label_scripts, "This is the local folder where all your Python scripts are located. The application will look for and save scripts in this directory.", self.secondary_bg, self.text_color)  
        info_label_scripts.pack(side="left", anchor="center")

        scripts_folder_frame = ttk.Frame(scripts_folder_wrapper_frame, style='TFrame')
        scripts_folder_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        scripts_folder_frame.grid_columnconfigure(1, weight=1)

        ttk.Label(scripts_folder_frame, text="Path to Scripts Folder:", style='TLabel').grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(scripts_folder_frame, textvariable=self.scripts_root_folder, width=40, style='TEntry').grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(scripts_folder_frame, text="Browse", command=self._browse_scripts_root_folder, style='TButton').grid(row=0, column=2, padx=5, pady=5)

        # NEW SECTION: Download Renamer Excel (Moved here)
        renamer_excel_wrapper_frame = ttk.Frame(self.scrollable_frame, style='SectionFrame.TFrame')
        renamer_excel_wrapper_frame.grid(row=row_counter, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
        row_counter += 1

        header_sub_frame_excel = ttk.Frame(renamer_excel_wrapper_frame, style='TFrame')
        header_sub_frame_excel.pack(side="top", fill="x", pady=(0, 5), padx=0)
        header_label_excel = ttk.Label(header_sub_frame_excel, text="Download Renamer Excel Template", style='Header.TLabel')
        header_label_excel.pack(side="left", padx=(0, 5))
        info_label_excel = ttk.Label(header_sub_frame_excel, text=" ⓘ", font=self.base_font)  
        Tooltip(info_label_excel, "Downloads a fresh copy of the Renamer Excel template directly from Bynder.", self.secondary_bg, self.text_color)
        info_label_excel.pack(side="left", anchor="center")

        renamer_excel_frame = ttk.Frame(renamer_excel_wrapper_frame, style='TFrame')
        renamer_excel_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        renamer_excel_frame.grid_columnconfigure(0, weight=1)

        self.download_renamer_excel_button = ttk.Button(renamer_excel_frame, text="Download New Renamer Excel", command=self._download_renamer_excel, style='TButton')
        self.download_renamer_excel_button.grid(row=0, column=0, pady=5, sticky="ew")

        initial_acquisition_wrapper_frame = ttk.Frame(self.scrollable_frame, style='SectionFrame.TFrame')
        initial_acquisition_wrapper_frame.grid(row=row_counter, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
        row_counter += 1

        header_sub_frame_3 = ttk.Frame(initial_acquisition_wrapper_frame, style='TFrame')
        header_sub_frame_3.pack(side="top", fill="x", pady=(0, 5), padx=0)
        header_label_3 = ttk.Label(header_sub_frame_3, text="Initial Image Acquisition (Download/Copy)", style='Header.TLabel')
        header_label_3.pack(side="left", padx=(0, 5))
        info_label_3 = ttk.Label(header_sub_frame_3, text=" ⓘ", font=self.base_font)  
        Tooltip(info_label_3, "This section allows you to acquire initial image assets for renaming, either by copying local DI'd images, downloading from URLs (PSO Option 1), or copying from network locations (PSO Option 2).", self.secondary_bg, self.text_color)  
        info_label_3.pack(side="left", anchor="center")

        initial_acquisition_frame = ttk.Frame(initial_acquisition_wrapper_frame, style='TFrame')
        initial_acquisition_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        ttk.Radiobutton(initial_acquisition_frame, text="Inline Project (copy from folder via Spreadsheet)", variable=self.source_type, value="inline", command=self._show_source_section, style='TRadiobutton').pack(anchor="w", padx=5)  
        ttk.Radiobutton(initial_acquisition_frame, text="PSO Option 1 (Download from URLs in Spreadsheet)", variable=self.source_type, value="pso1", command=self._show_source_section, style='TRadiobutton').pack(anchor="w", padx=5)
        ttk.Radiobutton(initial_acquisition_frame, text="PSO Option 2 (Copy from Network via Spreadsheet)", variable=self.source_type, value="pso2", command=self._show_source_section, style='TRadiobutton').pack(anchor="w", padx=5)
        
        self.source_sections = {}

        self.inline_section = ttk.Frame(initial_acquisition_frame, style='TFrame')
        self.inline_section.grid_columnconfigure(1, weight=1)
        ttk.Label(self.inline_section, text="Source Folder (Network Assets):", style='TLabel').grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(self.inline_section, textvariable=self.inline_source_folder, width=40, style='TEntry').grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(self.inline_section, text="Browse", command=lambda: self._browse_folder(self.inline_source_folder), style='TButton').grid(row=0, column=2, padx=5, pady=5)
        
        ttk.Label(self.inline_section, text="Renamer Matrix (with Filenames):", style='TLabel').grid(row=1, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(self.inline_section, textvariable=self.inline_matrix_path, width=40, style='TEntry').grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(self.inline_section, text="Browse", command=lambda: self._browse_file(self.inline_matrix_path, "xlsx"), style='TButton').grid(row=1, column=2, padx=5, pady=5)

        ttk.Label(self.inline_section, text="Output Folder for Copied Images:", style='TLabel').grid(row=2, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(self.inline_section, textvariable=self.inline_output_folder, width=40, style='TEntry').grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(self.inline_section, text="Browse", command=lambda: self._browse_folder(self.inline_output_folder), style='TButton').grid(row=2, column=2, padx=5, pady=5)

        self.inline_copy_run_control_frame = ttk.Frame(self.inline_section, style='TFrame')
        self.inline_copy_run_control_frame.grid(row=3, column=0, columnspan=3, pady=10, sticky="ew")
        self.inline_copy_run_control_frame.grid_columnconfigure(0, weight=1)
        self.inline_copy_run_control_frame.grid_columnconfigure(1, weight=0)
        self.inline_copy_run_control_frame.grid_columnconfigure(2, weight=1)

        self.inline_copy_run_button_wrapper = ttk.Frame(self.inline_copy_run_control_frame, style='TFrame')
        self.inline_copy_run_button_wrapper.grid(row=0, column=1, sticky="")
        self.run_inline_copy_button = ttk.Button(self.inline_copy_run_button_wrapper, text="Start Copy (Inline Project)", command=self._start_inline_copy, style='TButton')
        self.run_inline_copy_button.pack(padx=5, pady=0)

        self.inline_copy_progress_wrapper = ttk.Frame(self.inline_copy_run_control_frame, style='TFrame')
        self.inline_copy_progress_wrapper.grid(row=0, column=1, sticky="ew")
        self.inline_copy_progress_bar = ttk.Progressbar(self.inline_copy_progress_wrapper, orient="horizontal", length=200, mode="determinate")
        self.inline_copy_progress_bar.pack(side="left", fill="x", expand=True, padx=5)
        self.inline_copy_progress_label = ttk.Label(self.inline_copy_progress_wrapper, text="", style='TLabel')
        self.inline_copy_progress_label.pack(side="right", padx=5)
        self.inline_copy_progress_wrapper.grid_remove()

        self.source_sections["inline"] = self.inline_section

        self.pso1_section = ttk.Frame(initial_acquisition_frame, style='TFrame')
        self.pso1_section.grid_columnconfigure(1, weight=1)
        ttk.Label(self.pso1_section, text="Renamer Matrix (with URLs):", style='TLabel').grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.pso1_matrix_path = tk.StringVar()
        ttk.Entry(self.pso1_section, textvariable=self.pso1_matrix_path, width=40, style='TEntry').grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(self.pso1_section, text="Browse", command=lambda: self._browse_file(self.pso1_matrix_path, "xlsx"), style='TButton').grid(row=0, column=2, padx=5, pady=5)
        ttk.Label(self.pso1_section, text="Output Folder for Downloaded Images:", style='TLabel').grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.pso1_output_folder = tk.StringVar()
        ttk.Entry(self.pso1_section, textvariable=self.pso1_output_folder, width=40, style='TEntry').grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(self.pso1_section, text="Browse", command=lambda: self._browse_folder(self.pso1_output_folder), style='TButton').grid(row=1, column=2, padx=5, pady=5)
        
        self.pso1_download_run_control_frame = ttk.Frame(self.pso1_section, style='TFrame')
        self.pso1_download_run_control_frame.grid(row=2, column=0, columnspan=3, pady=10, sticky="ew")
        self.pso1_download_run_control_frame.grid_columnconfigure(0, weight=1)
        self.pso1_download_run_control_frame.grid_columnconfigure(1, weight=0)
        self.pso1_download_run_control_frame.grid_columnconfigure(2, weight=1)

        self.pso1_download_run_button_wrapper = ttk.Frame(self.pso1_download_run_control_frame, style='TFrame')
        self.pso1_download_run_button_wrapper.grid(row=0, column=1, sticky="")
        self.run_pso1_download_button = ttk.Button(self.pso1_download_run_button_wrapper, text="Start Download (PSO Option 1)", command=self._start_pso1_download, style='TButton')
        self.run_pso1_download_button.pack(padx=5, pady=0)

        self.pso1_download_progress_wrapper = ttk.Frame(self.pso1_download_run_control_frame, style='TFrame')
        self.pso1_download_progress_wrapper.grid(row=0, column=1, sticky="ew")
        self.pso1_download_progress_bar = ttk.Progressbar(self.pso1_download_progress_wrapper, orient="horizontal", length=200, mode="determinate")
        self.pso1_download_progress_bar.pack(side="left", fill="x", expand=True, padx=5)
        self.pso1_download_progress_label = ttk.Label(self.pso1_download_progress_wrapper, text="", style='TLabel')
        self.pso1_download_progress_label.pack(side="right", padx=5)
        self.pso1_download_progress_wrapper.grid_remove()

        self.source_sections["pso1"] = self.pso1_section

        self.pso2_section = ttk.Frame(initial_acquisition_frame, style='TFrame')
        self.pso2_section.grid_columnconfigure(1, weight=1)
        ttk.Label(self.pso2_section, text="Network Assets Source Folder:", style='TLabel').grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.pso2_network_folder = tk.StringVar()
        ttk.Entry(self.pso2_section, textvariable=self.pso2_network_folder, width=40, style='TEntry').grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(self.pso2_section, text="Browse", command=lambda: self._browse_folder(self.pso2_network_folder), style='TButton').grid(row=0, column=2, padx=5, pady=5)
        ttk.Label(self.pso2_section, text="Renamer Matrix (with Filenames):", style='TLabel').grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.pso2_matrix_path = tk.StringVar()
        ttk.Entry(self.pso2_section, textvariable=self.pso2_matrix_path, width=40, style='TEntry').grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(self.pso2_section, text="Browse", command=lambda: self._browse_file(self.pso2_matrix_path, "xlsx"), style='TButton').grid(row=1, column=2, padx=5, pady=5)
        ttk.Label(self.pso2_section, text="Output Folder for Copied Images:", style='TLabel').grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.pso2_output_folder = tk.StringVar()
        ttk.Entry(self.pso2_section, textvariable=self.pso2_output_folder, width=40, style='TEntry').grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(self.pso2_section, text="Browse", command=lambda: self._browse_folder(self.pso2_output_folder), style='TButton').grid(row=2, column=2, padx=5, pady=5)
        
        self.pso2_copy_run_control_frame = ttk.Frame(self.pso2_section, style='TFrame')
        self.pso2_copy_run_control_frame.grid(row=3, column=0, columnspan=3, pady=10, sticky="ew")
        self.pso2_copy_run_control_frame.grid_columnconfigure(0, weight=1)
        self.pso2_copy_run_control_frame.grid_columnconfigure(1, weight=0)
        self.pso2_copy_run_control_frame.grid_columnconfigure(2, weight=1)

        self.pso2_copy_run_button_wrapper = ttk.Frame(self.pso2_copy_run_control_frame, style='TFrame')
        self.pso2_copy_run_button_wrapper.grid(row=0, column=1, sticky="")
        self.run_pso2_copy_button = ttk.Button(self.pso2_copy_run_button_wrapper, text="Start Copy (PSO Option 2)", command=self._start_pso2_copy, style='TButton')
        self.run_pso2_copy_button.pack(padx=5, pady=0)

        self.pso2_copy_progress_wrapper = ttk.Frame(self.pso2_copy_run_control_frame, style='TFrame')
        self.pso2_copy_progress_wrapper.grid(row=0, column=1, sticky="ew")
        self.pso2_copy_progress_bar = ttk.Progressbar(self.pso2_copy_progress_wrapper, orient="horizontal", length=200, mode="determinate")
        self.pso2_copy_progress_bar.pack(side="left", fill="x", expand=True, padx=5)
        self.pso2_copy_progress_label = ttk.Label(self.pso2_copy_progress_wrapper, text="", style='TLabel')
        self.pso2_copy_progress_label.pack(side="right", padx=5)
        self.pso2_copy_progress_wrapper.grid_remove()

        self.source_sections["pso2"] = self.pso2_section
        
        row_counter += 1  

        master_renamer_wrapper_frame = ttk.Frame(self.scrollable_frame, style='SectionFrame.TFrame')
        master_renamer_wrapper_frame.grid(row=row_counter, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
        row_counter += 1

        header_sub_frame_2 = ttk.Frame(master_renamer_wrapper_frame, style='TFrame')
        header_sub_frame_2.pack(side="top", fill="x", pady=(0, 5), padx=0)
        header_label_2 = ttk.Label(header_sub_frame_2, text="Main Renamer Script", style='Header.TLabel')
        header_label_2.pack(side="left", padx=(0, 5))
        info_label_2 = ttk.Label(header_sub_frame_2, text=" ⓘ", font=self.base_font)  
        Tooltip(info_label_2, "Use this section to run the primary renaming process. It renames images based on a matrix, handles JPGs, vendor codes, and organizes outputs.", self.secondary_bg, self.text_color)  
        info_label_2.pack(side="left", anchor="center")
        master_renamer_frame = ttk.Frame(master_renamer_wrapper_frame, style='TFrame')
        master_renamer_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))  

        ttk.Label(master_renamer_frame, text="Renamer Matrix (.xlsx):", style='TLabel').grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.master_matrix_path = tk.StringVar()
        ttk.Entry(master_renamer_frame, textvariable=self.master_matrix_path, width=45, style='TEntry').grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(master_renamer_frame, text="Browse", command=lambda: self._browse_file(self.master_matrix_path, "xlsx"), style='TButton').grid(row=0, column=2, padx=5, pady=5)
        master_renamer_frame.grid_columnconfigure(1, weight=1)

        ttk.Label(master_renamer_frame, text="Input Images Folder:", style='TLabel').grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.rename_input_folder = tk.StringVar()
        ttk.Entry(master_renamer_frame, textvariable=self.rename_input_folder, width=45, style='TEntry').grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(master_renamer_frame, text="Browse", command=lambda: self._browse_folder(self.rename_input_folder), style='TButton').grid(row=1, column=2, padx=5, pady=5)

        ttk.Label(master_renamer_frame, text="Vendor Code:", style='TLabel').grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.vendor_code = tk.StringVar()
        ttk.Entry(master_renamer_frame, textvariable=self.vendor_code, width=15, style='TEntry').grid(row=2, column=1, padx=5, pady=5, sticky="w")
        
        ttk.Button(master_renamer_frame, text="Run Renamer", command=self._start_master_renamer_threaded, style='TButton').grid(row=3, column=0, columnspan=3, pady=10)


        bynder_prep_wrapper_frame = ttk.Frame(self.scrollable_frame, style='SectionFrame.TFrame')
        bynder_prep_wrapper_frame.grid(row=row_counter, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
        row_counter += 1

        header_sub_frame_5 = ttk.Frame(bynder_prep_wrapper_frame, style='TFrame')
        header_sub_frame_5.pack(side="top", fill="x", pady=(0, 5), padx=0)
        header_label_5 = ttk.Label(header_sub_frame_5, text="Bynder Metadata Preparation", style='Header.TLabel')
        header_label_5.pack(side="left", padx=(0, 5))
        info_label_5 = ttk.Label(header_sub_frame_5, text=" ⓘ", font=self.base_font)  
        Tooltip(info_label_5, "Prepare metadata for assets to be uploaded to Bynder using information from STEP exports.", self.secondary_bg, self.text_color)  
        info_label_5.pack(side="left", anchor="center")

        bynder_prep_frame = ttk.Frame(bynder_prep_wrapper_frame, style='TFrame')
        bynder_prep_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
            
        ttk.Label(bynder_prep_frame, text="Folder of Assets for Bynder Metadata Prep:", style='TLabel').grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.bynder_assets_folder = tk.StringVar()
        ttk.Entry(bynder_prep_frame, textvariable=self.bynder_assets_folder, width=45, style='TEntry').grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(bynder_prep_frame, text="Browse Folder", command=lambda: self._browse_folder(self.bynder_assets_folder), style='TButton').grid(row=0, column=2, padx=5, pady=5)
        
        self.bynder_prep_run_control_frame = ttk.Frame(bynder_prep_frame, style='TFrame')
        self.bynder_prep_run_control_frame.grid(row=1, column=0, columnspan=3, pady=10, sticky="ew")
        
        self.bynder_prep_run_control_frame.grid_columnconfigure(0, weight=1)
        self.bynder_prep_run_control_frame.grid_columnconfigure(1, weight=0)
        self.bynder_prep_run_control_frame.grid_columnconfigure(2, weight=1)

        self.bynder_prep_run_button_wrapper = ttk.Frame(self.bynder_prep_run_control_frame, style='TFrame')
        self.bynder_prep_run_button_wrapper.grid(row=0, column=1, sticky="")

        self.run_bynder_prep_button = ttk.Button(self.bynder_prep_run_button_wrapper, text="Prepare metadata for Bynder upload", command=self._run_bynder_metadata_prep, style='TButton')
        self.run_bynder_prep_button.pack(padx=5, pady=0)

        self.bynder_prep_progress_wrapper = ttk.Frame(self.bynder_prep_run_control_frame, style='TFrame')
        self.bynder_prep_progress_wrapper.grid(row=0, column=1, sticky="ew")

        self.bynder_prep_progress_bar = ttk.Progressbar(self.bynder_prep_progress_wrapper, orient="horizontal", length=200, mode="determinate")
        self.bynder_prep_progress_bar.pack(side="left", fill="x", expand=True, padx=5)
        self.bynder_prep_progress_label = ttk.Label(self.bynder_prep_progress_wrapper, text="", style='TLabel')
        self.bynder_prep_progress_label.pack(side="right", padx=5)

        self.bynder_prep_progress_wrapper.grid_remove()

        bynder_prep_frame.grid_columnconfigure(1, weight=1)


        ttk.Separator(self.scrollable_frame, orient="horizontal", style='TSeparator').grid(row=row_counter, column=0, columnspan=2, padx=10, pady=(20, 15), sticky="ew")
        row_counter += 1

        addendum_header_frame = ttk.Frame(self.scrollable_frame, style='TFrame')
        addendum_header_frame.grid(row=row_counter, column=0, columnspan=2, padx=10, pady=(0, 5), sticky="w")
        addendum_label = ttk.Label(addendum_header_frame, text="Extra Tools", style='Header.TLabel')
        addendum_label.pack(side="left", padx=(0, 5))
        info_label_addendum = ttk.Label(addendum_header_frame, text=" ⓘ", font=self.base_font)  
        Tooltip(info_label_addendum, "This section contains additional tools for image preparation and Bynder-related operations.", self.secondary_bg, self.text_color)  
        info_label_addendum.pack(side="left", anchor="center")
        row_counter += 1

        image_prep_wrapper_frame = ttk.Frame(self.scrollable_frame, style='SectionFrame.TFrame')
        image_prep_wrapper_frame.grid(row=row_counter, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
        row_counter += 1

        header_sub_frame_4 = ttk.Frame(image_prep_wrapper_frame, style='TFrame')
        header_sub_frame_4.pack(side="top", fill="x", pady=(0, 5), padx=0)
        header_label_4 = ttk.Label(header_sub_frame_4, text="Image Preparation & Cropping", style='Header.TLabel')
        header_label_4.pack(side="left", padx=(0, 5))
        info_label_4 = ttk.Label(header_sub_frame_4, text=" ⓘ", font=self.base_font)  
        Tooltip(info_label_4, "Tools to prepare and crop images before renaming or uploading to Bynder.", self.secondary_bg, self.text_color)  
        info_label_4.pack(side="left", anchor="center")

        image_prep_frame = ttk.Frame(image_prep_wrapper_frame, style='TFrame')
        image_prep_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        ttk.Label(image_prep_frame, text="Images to Crop with Scripts:", style='TLabel').grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.prep_input_path = tk.StringVar()  
        ttk.Entry(image_prep_frame, textvariable=self.prep_input_path, width=45, style='TEntry').grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(image_prep_frame, text="Browse Folder", command=lambda: self._browse_folder(self.prep_input_path), style='TButton').grid(row=0, column=2, padx=5, pady=5)
        
        image_prep_frame.grid_columnconfigure(1, weight=1)

        ttk.Separator(image_prep_frame, orient="horizontal", style='TSeparator').grid(row=1, column=0, columnspan=3, sticky="ew", pady=5)  
        ttk.Label(image_prep_frame, text="Run Cropping Scripts (Require Folder Input):", font=self.base_font, foreground=self.text_color, background=self.primary_bg).grid(row=2, column=0, columnspan=3, sticky="w", padx=5, pady=5)  
        
        self.cropping_run_control_frame = ttk.Frame(image_prep_frame, style='TFrame')
        self.cropping_run_control_frame.grid(row=3, column=0, columnspan=3, pady=10, sticky="ew")
        self.cropping_run_control_frame.grid_columnconfigure(0, weight=1)
        self.cropping_run_control_frame.grid_columnconfigure(1, weight=0)
        self.cropping_run_control_frame.grid_columnconfigure(2, weight=1)

        self.cropping_run_button_wrapper = ttk.Frame(self.cropping_run_control_frame, style='TFrame')
        self.cropping_run_button_wrapper.grid(row=0, column=1, sticky="")

        self.cropping_buttons = {}
        self.cropping_buttons["1688_silo"] = ttk.Button(self.cropping_run_button_wrapper, text="Crop Silo (3000x1688)", command=lambda: self._run_cropping_script("reformat1688_silo.py"), style='TButton')
        self.cropping_buttons["1688_room"] = ttk.Button(self.cropping_run_button_wrapper, text="Crop Room (3000x1688)", command=lambda: self._run_cropping_script("reformat1688_room.py"), style='TButton')
        self.cropping_buttons["1688_room_cutLR"] = ttk.Button(self.cropping_run_button_wrapper, text="Crop Room CutLR (3000x1688)", command=lambda: self._run_cropping_script("reformat1688_room_cutLR.py"), style='TButton')
        self.cropping_buttons["1688_room_cutTopBot"] = ttk.Button(self.cropping_run_button_wrapper, text="Crop Room CutTopBot (3000x1688)", command=lambda: self._run_cropping_script("reformat1688_room_cutTopBot.py"), style='TButton')
        self.cropping_buttons["2200_silo"] = ttk.Button(self.cropping_run_button_wrapper, text="Crop Silo (3000x2200)", command=lambda: self._run_cropping_script("reformat2200_silo.py"), style='TButton')
        self.cropping_buttons["2200_room"] = ttk.Button(self.cropping_run_button_wrapper, text="Crop Room (3000x2200)", command=lambda: self._run_cropping_script("reformat2200_room.py"), style='TButton')

        self.cropping_buttons["1688_silo"].grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        self.cropping_buttons["2200_silo"].grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.cropping_buttons["1688_room"].grid(row=1, column=0, padx=5, pady=5, sticky="ew")
        self.cropping_buttons["2200_room"].grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        self.cropping_buttons["1688_room_cutLR"].grid(row=2, column=0, padx=5, pady=5, sticky="ew")
        self.cropping_buttons["1688_room_cutTopBot"].grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        self.cropping_run_button_wrapper.grid_columnconfigure(0, weight=1)
        self.cropping_run_button_wrapper.grid_columnconfigure(1, weight=1)


        self.cropping_progress_wrapper = ttk.Frame(self.cropping_run_control_frame, style='TFrame')
        self.cropping_progress_wrapper.grid(row=0, column=1, sticky="ew")
        
        self.cropping_progress_bar = ttk.Progressbar(self.cropping_progress_wrapper, orient="horizontal", length=200, mode="determinate")
        self.cropping_progress_bar.pack(side="left", fill="x", expand=True, padx=5)
        self.cropping_progress_label = ttk.Label(self.cropping_progress_wrapper, text="", style='TLabel')
        self.cropping_progress_label.pack(side="right", padx=5)
        self.cropping_progress_wrapper.grid_remove()

        row_counter += 1

        # NEW SECTION: Bynder Metadata Convert to XLS
        bynder_metadata_convert_wrapper_frame = ttk.Frame(self.scrollable_frame, style='SectionFrame.TFrame')
        bynder_metadata_convert_wrapper_frame.grid(row=row_counter, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
        row_counter += 1

        header_sub_frame_bynder_convert = ttk.Frame(bynder_metadata_convert_wrapper_frame, style='TFrame')
        header_sub_frame_bynder_convert.pack(side="top", fill="x", pady=(0, 5), padx=0)
        header_label_bynder_convert = ttk.Label(header_sub_frame_bynder_convert, text="Convert Bynder Metadata CSV to XLS", style='Header.TLabel')
        header_label_bynder_convert.pack(side="left", padx=(0, 5))
        info_label_bynder_convert = ttk.Label(header_sub_frame_bynder_convert, text=" ⓘ", font=self.base_font)  
        Tooltip(info_label_bynder_convert, "Did you download a metadata csv from assets in Bynder? Use this tool to easily convert that csv to XLSX! It will be exported to your Downloads folder.", self.secondary_bg, self.text_color)  
        info_label_bynder_convert.pack(side="left", anchor="center")

        bynder_metadata_convert_frame = ttk.Frame(bynder_metadata_convert_wrapper_frame, style='TFrame')
        bynder_metadata_convert_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        bynder_metadata_convert_frame.grid_columnconfigure(1, weight=1)

        ttk.Label(bynder_metadata_convert_frame, text="Bynder Metadata CSV File:", style='TLabel').grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(bynder_metadata_convert_frame, textvariable=self.bynder_metadata_csv_path, width=45, style='TEntry').grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(bynder_metadata_convert_frame, text="Browse", command=lambda: self._browse_file(self.bynder_metadata_csv_path, "csv"), style='TButton').grid(row=0, column=2, padx=5, pady=5)
        
        self.bynder_metadata_convert_run_control_frame = ttk.Frame(bynder_metadata_convert_frame, style='TFrame')
        self.bynder_metadata_convert_run_control_frame.grid(row=1, column=0, columnspan=3, pady=10, sticky="ew")
        
        self.bynder_metadata_convert_run_control_frame.grid_columnconfigure(0, weight=1)
        self.bynder_metadata_convert_run_control_frame.grid_columnconfigure(1, weight=0)
        self.bynder_metadata_convert_run_control_frame.grid_columnconfigure(2, weight=1)

        self.bynder_metadata_convert_run_button_wrapper = ttk.Frame(self.bynder_metadata_convert_run_control_frame, style='TFrame')
        self.bynder_metadata_convert_run_button_wrapper.grid(row=0, column=1, sticky="")

        self.run_bynder_metadata_convert_button = ttk.Button(self.bynder_metadata_convert_run_button_wrapper, text="Convert CSV to XLS", command=self._run_bynder_metadata_convert_script, style='TButton')
        self.run_bynder_metadata_convert_button.pack(padx=5, pady=0)

        self.bynder_metadata_convert_progress_wrapper = ttk.Frame(self.bynder_metadata_convert_run_control_frame, style='TFrame')
        self.bynder_metadata_convert_progress_wrapper.grid(row=0, column=1, sticky="ew")
        self.bynder_metadata_convert_progress_bar = ttk.Progressbar(self.bynder_metadata_convert_progress_wrapper, orient="horizontal", length=200, mode="determinate")
        self.bynder_metadata_convert_progress_bar.pack(side="left", fill="x", expand=True, padx=5)
        self.bynder_metadata_convert_progress_label = ttk.Label(self.bynder_metadata_convert_progress_wrapper, text="", style='TLabel')
        self.bynder_metadata_convert_progress_label.pack(side="right", padx=5)
        self.bynder_metadata_convert_progress_wrapper.grid_remove()

        row_counter += 1

        check_psas_wrapper_frame = ttk.Frame(self.scrollable_frame, style='SectionFrame.TFrame')
        check_psas_wrapper_frame.grid(row=row_counter, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
        row_counter += 1

        header_sub_frame_check_psa = ttk.Frame(check_psas_wrapper_frame, style='TFrame')
        header_sub_frame_check_psa.pack(side="top", fill="x", pady=(0, 5), padx=0)
        header_label_check_psa = ttk.Label(header_sub_frame_check_psa, text="Check Bynder PSAs", style='Header.TLabel')
        header_label_check_psa.pack(side="left", padx=(0, 5))
        info_label_check_psa = ttk.Label(header_sub_frame_check_psa, text=" ⓘ", font=self.base_font)  
        Tooltip(info_label_check_psa, "Checks Bynder to see if specific PSAs (Product Shot Assets) exist for the given SKUs.", self.secondary_bg, self.text_color)  
        info_label_check_psa.pack(side="left", anchor="center")

        check_psas_frame = ttk.Frame(check_psas_wrapper_frame, style='TFrame')
        check_psas_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        check_psas_frame.grid_columnconfigure(0, weight=1)

        input_method_frame_check_psa = ttk.Frame(check_psas_frame, style='TFrame')
        input_method_frame_check_psa.grid(row=0, column=0, columnspan=3, sticky="w", padx=5, pady=(0,5))
        ttk.Label(input_method_frame_check_psa, text="Input Method:", style='TLabel').pack(side="left", padx=(0,5))
        ttk.Radiobutton(input_method_frame_check_psa, text="From Spreadsheet", variable=self.check_psa_input_type, value="spreadsheet", command=lambda: self._show_input_method("check_psa", "spreadsheet"), style='TRadiobutton').pack(side="left", padx=5)
        ttk.Radiobutton(input_method_frame_check_psa, text="From Text Box", variable=self.check_psa_input_type, value="textbox", command=lambda: self._show_input_method("check_psa", "textbox"), style='TRadiobutton').pack(side="left", padx=5)

        self.check_psa_spreadsheet_frame = ttk.Frame(check_psas_frame, style='TFrame')
        self.check_psa_spreadsheet_frame.grid(row=1, column=0, columnspan=3, sticky="ew")
        ttk.Label(self.check_psa_spreadsheet_frame, text="SKU Spreadsheet (.xlsx):", style='TLabel').grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.check_psa_sku_spreadsheet_path = tk.StringVar()
        ttk.Entry(self.check_psa_spreadsheet_frame, textvariable=self.check_psa_sku_spreadsheet_path, width=45, style='TEntry').grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(self.check_psa_spreadsheet_frame, text="Browse", command=lambda: self._browse_file(self.check_psa_sku_spreadsheet_path, "xlsx"), style='TButton').grid(row=0, column=2, padx=5, pady=5)
        self.check_psa_spreadsheet_frame.grid_columnconfigure(1, weight=1)

        self.check_psa_textbox_frame = ttk.Frame(check_psas_frame, style='TFrame')
        ttk.Label(self.check_psa_textbox_frame, text="Paste SKUs (one per line):", style='TLabel').pack(padx=5, pady=5, anchor="w")
        self.check_psa_text_widget = scrolledtext.ScrolledText(self.check_psa_textbox_frame, width=60, height=8, font=self.base_font,
                                                               bg=self.secondary_bg, fg=self.text_color, wrap=tk.WORD,
                                                               insertbackground=self.text_color, relief="solid", borderwidth=1)
        self.check_psa_text_widget.pack(padx=5, pady=(0, 5), fill="both", expand=True)

        self.check_psas_run_control_frame = ttk.Frame(check_psas_frame, style='TFrame')
        self.check_psas_run_control_frame.grid(row=2, column=0, columnspan=3, pady=10, sticky="ew")
        self.check_psas_run_control_frame.grid_columnconfigure(0, weight=1)
        self.check_psas_run_control_frame.grid_columnconfigure(1, weight=0)
        self.check_psas_run_control_frame.grid_columnconfigure(2, weight=1)

        self.check_psas_run_button_wrapper = ttk.Frame(self.check_psas_run_control_frame, style='TFrame')
        self.check_psas_run_button_wrapper.grid(row=0, column=1, sticky="")
        self.run_check_psas_button = ttk.Button(self.check_psas_run_button_wrapper, text="Run Check Bynder PSAs", command=self._run_check_psas_script, style='TButton')
        self.run_check_psas_button.pack(padx=5, pady=0)

        self.check_psas_progress_wrapper = ttk.Frame(self.check_psas_run_control_frame, style='TFrame')
        self.check_psas_progress_wrapper.grid(row=0, column=1, sticky="ew")
        self.check_psas_progress_bar = ttk.Progressbar(self.check_psas_progress_wrapper, orient="horizontal", length=200, mode="determinate")
        self.check_psas_progress_bar.pack(side="left", fill="x", expand=True, padx=5)
        self.check_psas_progress_label = ttk.Label(self.check_psas_progress_wrapper, text="", style='TLabel')
        self.check_psas_progress_label.pack(side="right", padx=5)
        self.check_psas_progress_wrapper.grid_remove()


        row_counter += 1

        download_psas_wrapper_frame = ttk.Frame(self.scrollable_frame, style='SectionFrame.TFrame')
        download_psas_wrapper_frame.grid(row=row_counter, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
        row_counter += 1

        header_sub_frame_download_psa = ttk.Frame(download_psas_wrapper_frame, style='TFrame')
        header_sub_frame_download_psa.pack(side="top", fill="x", pady=(0, 5), padx=0)
        header_label_download_psa = ttk.Label(header_sub_frame_download_psa, text="Download PSAs", style='Header.TLabel')
        header_label_download_psa.pack(side="left", padx=(0, 5))
        info_label_download_psa = ttk.Label(header_sub_frame_download_psa, text=" ⓘ", font=self.base_font)  
        Tooltip(info_label_download_psa, "Downloads specified PSA (Product Shot Asset) image types from Bynder for a list of SKUs.", self.secondary_bg, self.text_color)  
        info_label_download_psa.pack(side="left", anchor="center")

        download_psas_frame = ttk.Frame(download_psas_wrapper_frame, style='TFrame')
        download_psas_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        download_psas_frame.grid_columnconfigure(1, weight=1)

        input_method_frame_download_psa = ttk.Frame(download_psas_frame, style='TFrame')
        input_method_frame_download_psa.grid(row=0, column=0, columnspan=3, sticky="w", padx=5, pady=(0,5))
        ttk.Label(input_method_frame_download_psa, text="Input Method:", style='TLabel').pack(side="left", padx=(0,5))
        ttk.Radiobutton(input_method_frame_download_psa, text="From Spreadsheet", variable=self.download_psa_input_type, value="spreadsheet", command=lambda: self._show_input_method_download_psa("spreadsheet"), style='TRadiobutton').pack(side="left", padx=5)
        ttk.Radiobutton(input_method_frame_download_psa, text="From Text Box", variable=self.download_psa_input_type, value="textbox", command=lambda: self._show_input_method_download_psa("textbox"), style='TRadiobutton').pack(side="left", padx=5)

        self.download_psa_spreadsheet_frame = ttk.Frame(download_psas_frame, style='TFrame')
        ttk.Label(self.download_psa_spreadsheet_frame, text="SKU Spreadsheet (.xlsx):", style='TLabel').grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.download_psa_sku_spreadsheet_path = tk.StringVar()
        ttk.Entry(self.download_psa_spreadsheet_frame, textvariable=self.download_psa_sku_spreadsheet_path, width=45, style='TEntry').grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(self.download_psa_spreadsheet_frame, text="Browse", command=lambda: self._browse_file(self.download_psa_sku_spreadsheet_path, "xlsx"), style='TButton').grid(row=0, column=2, padx=5, pady=5)
        self.download_psa_spreadsheet_frame.grid_columnconfigure(1, weight=1)

        self.download_psa_textbox_frame = ttk.Frame(download_psas_frame, style='TFrame')
        ttk.Label(self.download_psa_textbox_frame, text="Paste SKUs (one per line):", style='TLabel').pack(padx=5, pady=5, anchor="w")
        self.download_psa_text_widget = scrolledtext.ScrolledText(self.download_psa_textbox_frame, width=60, height=8, font=self.base_font,
                                                                bg=self.secondary_bg, fg=self.text_color, wrap=tk.WORD,
                                                                insertbackground=self.text_color, relief="solid", borderwidth=1)
        self.download_psa_text_widget.pack(padx=5, pady=(0, 5), fill="both", expand=True)

        self.download_psa_spreadsheet_frame.grid(row=1, column=0, columnspan=3, sticky="ew")
        self.download_psa_textbox_frame.grid(row=1, column=0, columnspan=3, sticky="nsew")

        ttk.Label(download_psas_frame, text="Output Folder:", style='TLabel').grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.download_psa_output_folder = tk.StringVar()
        ttk.Entry(download_psas_frame, textvariable=self.download_psa_output_folder, width=45, style='TEntry').grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(download_psas_frame, text="Browse", command=lambda: self._browse_folder(self.download_psa_output_folder), style='TButton').grid(row=2, column=2, padx=5, pady=5)

        ttk.Label(download_psas_frame, text="Select Image Types:", style='TLabel').grid(row=3, column=0, padx=5, pady=5, sticky="w")
        
        # Frame for checkboxes and new buttons
        image_types_controls_frame = ttk.Frame(download_psas_frame, style='TFrame')
        image_types_controls_frame.grid(row=3, column=1, columnspan=2, sticky="w", padx=5, pady=5)

        image_types_frame = ttk.Frame(image_types_controls_frame, style='TFrame')
        image_types_frame.pack(side="top", fill="x", expand=True)

        ttk.Checkbutton(image_types_frame, text="grid", variable=self.download_psa_grid).grid(row=0, column=0, sticky="w", padx=2)
        ttk.Checkbutton(image_types_frame, text="100", variable=self.download_psa_100).grid(row=0, column=1, sticky="w", padx=2)
        ttk.Checkbutton(image_types_frame, text="200", variable=self.download_psa_200).grid(row=0, column=2, sticky="w", padx=2)
        ttk.Checkbutton(image_types_frame, text="300", variable=self.download_psa_300).grid(row=0, column=3, sticky="w", padx=2)
        ttk.Checkbutton(image_types_frame, text="400", variable=self.download_psa_400).grid(row=1, column=0, sticky="w", padx=2)
        ttk.Checkbutton(image_types_frame, text="500", variable=self.download_psa_500).grid(row=1, column=1, sticky="w", padx=2)
        ttk.Checkbutton(image_types_frame, text="dimension", variable=self.download_psa_dimension).grid(row=1, column=2, sticky="w", padx=2)
        ttk.Checkbutton(image_types_frame, text="swatch", variable=self.download_psa_swatch).grid(row=1, column=3, sticky="w", padx=2)
        ttk.Checkbutton(image_types_frame, text="5000", variable=self.download_psa_5000).grid(row=2, column=0, sticky="w", padx=2)
        ttk.Checkbutton(image_types_frame, text="5100", variable=self.download_psa_5100).grid(row=2, column=1, sticky="w", padx=2)
        ttk.Checkbutton(image_types_frame, text="5200", variable=self.download_psa_5200).grid(row=2, column=2, sticky="w", padx=2)
        ttk.Checkbutton(image_types_frame, text="5300", variable=self.download_psa_5300).grid(row=2, column=3, sticky="w", padx=2)
        ttk.Checkbutton(image_types_frame, text="squareThumbnail", variable=self.download_psa_squareThumbnail).grid(row=3, column=0, sticky="w", padx=2)
        
        # New Select All and Clear All buttons
        selection_buttons_frame = ttk.Frame(image_types_controls_frame, style='TFrame')
        selection_buttons_frame.pack(side="bottom", fill="x", pady=(5,0))
        ttk.Button(selection_buttons_frame, text="Select All", command=self._select_all_psas, style='TButton', width=10).pack(side="left", padx=2)
        ttk.Button(selection_buttons_frame, text="Clear All", command=self._clear_all_psas, style='TButton', width=10).pack(side="left", padx=2)


        self.download_psas_run_control_frame = ttk.Frame(download_psas_frame, style='TFrame')
        self.download_psas_run_control_frame.grid(row=4, column=0, columnspan=3, pady=10, sticky="ew")
        
        self.download_psas_run_control_frame.grid_columnconfigure(0, weight=1)
        self.download_psas_run_control_frame.grid_columnconfigure(1, weight=0)
        self.download_psas_run_control_frame.grid_columnconfigure(2, weight=1)

        self.download_psas_run_button_wrapper = ttk.Frame(self.download_psas_run_control_frame, style='TFrame')
        self.download_psas_run_button_wrapper.grid(row=0, column=1, sticky="")
        
        self.run_download_psas_button = ttk.Button(self.download_psas_run_button_wrapper, text="Run Download PSAs", command=self._run_download_psas_script, style='TButton')
        self.run_download_psas_button.pack(padx=5, pady=0)

        self.download_psas_progress_wrapper = ttk.Frame(self.download_psas_run_control_frame, style='TFrame')
        self.download_psas_progress_wrapper.grid(row=0, column=1, sticky="ew")

        self.download_psas_progress_bar = ttk.Progressbar(self.download_psas_progress_wrapper, orient="horizontal", length=200, mode="determinate")
        self.download_psas_progress_bar.pack(side="left", fill="x", expand=True, padx=5)
        self.download_psas_progress_label = ttk.Label(self.download_psas_progress_wrapper, text="", style='TLabel')
        self.download_psas_progress_label.pack(side="right", padx=5)

        self.download_psas_progress_wrapper.grid_remove()


        row_counter += 1

        get_measurements_wrapper_frame = ttk.Frame(self.scrollable_frame, style='SectionFrame.TFrame')
        get_measurements_wrapper_frame.grid(row=row_counter, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
        row_counter += 1

        header_sub_frame_get_measurements = ttk.Frame(get_measurements_wrapper_frame, style='TFrame')
        header_sub_frame_get_measurements.pack(side="top", fill="x", pady=(0, 5), padx=0)
        header_label_get_measurements = ttk.Label(header_sub_frame_get_measurements, text="Get Measurements Script", style='Header.TLabel')
        header_label_get_measurements.pack(side="left", padx=(0, 5))
        info_label_get_measurements = ttk.Label(header_sub_frame_get_measurements, text=" ⓘ", font=self.base_font)  
        Tooltip(info_label_get_measurements, "Retrieves product measurements for specified SKUs from a STEP export (or similar source) and outputs them to an Excel file.", self.secondary_bg, self.text_color)  
        info_label_get_measurements.pack(side="left", anchor="center")

        get_measurements_frame = ttk.Frame(get_measurements_wrapper_frame, style='TFrame')
        get_measurements_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        get_measurements_frame.grid_columnconfigure(0, weight=1)

        input_method_frame_get_measurements = ttk.Frame(get_measurements_frame, style='TFrame')
        input_method_frame_get_measurements.grid(row=0, column=0, columnspan=3, sticky="w", padx=5, pady=(0,5))
        ttk.Label(input_method_frame_get_measurements, text="Input Method:", style='TLabel').pack(side="left", padx=(0,5))
        ttk.Radiobutton(input_method_frame_get_measurements, text="From Spreadsheet", variable=self.get_measurements_input_type, value="spreadsheet", command=lambda: self._show_input_method("get_measurements", "spreadsheet"), style='TRadiobutton').pack(side="left", padx=5)
        ttk.Radiobutton(input_method_frame_get_measurements, text="From Text Box", variable=self.get_measurements_input_type, value="textbox", command=lambda: self._show_input_method("get_measurements", "textbox"), style='TRadiobutton').pack(side="left", padx=5)

        self.get_measurements_spreadsheet_frame = ttk.Frame(get_measurements_frame, style='TFrame')
        self.get_measurements_spreadsheet_frame.grid(row=1, column=0, columnspan=3, sticky="ew")
        ttk.Label(self.get_measurements_spreadsheet_frame, text="SKU Spreadsheet (.xlsx):", style='TLabel').grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.get_measurements_sku_spreadsheet_path = tk.StringVar()
        ttk.Entry(self.get_measurements_spreadsheet_frame, textvariable=self.get_measurements_sku_spreadsheet_path, width=45, style='TEntry').grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(self.get_measurements_spreadsheet_frame, text="Browse", command=lambda: self._browse_file(self.get_measurements_sku_spreadsheet_path, "xlsx"), style='TButton').grid(row=0, column=2, padx=5, pady=5)
        self.get_measurements_spreadsheet_frame.grid_columnconfigure(1, weight=1)

        self.get_measurements_textbox_frame = ttk.Frame(get_measurements_frame, style='TFrame')
        self.get_measurements_textbox_frame.grid(row=1, column=0, columnspan=3, sticky="nsew")
        ttk.Label(self.get_measurements_textbox_frame, text="Paste SKUs (one per line):", style='TLabel').pack(padx=5, pady=5, anchor="w")
        self.get_measurements_text_widget = scrolledtext.ScrolledText(self.get_measurements_textbox_frame, width=60, height=8, font=self.base_font,
                                                                bg=self.secondary_bg, fg=self.text_color, wrap=tk.WORD,
                                                                insertbackground=self.text_color, relief="solid", borderwidth=1)
        self.get_measurements_text_widget.pack(padx=5, pady=(0, 5), fill="both", expand=True)

        self.get_measurements_run_control_frame = ttk.Frame(get_measurements_frame, style='TFrame')
        self.get_measurements_run_control_frame.grid(row=2, column=0, columnspan=3, pady=10, sticky="ew")
        
        self.get_measurements_run_control_frame.grid_columnconfigure(0, weight=1)
        self.get_measurements_run_control_frame.grid_columnconfigure(1, weight=0)
        self.get_measurements_run_control_frame.grid_columnconfigure(2, weight=1)

        self.get_measurements_run_button_wrapper = ttk.Frame(self.get_measurements_run_control_frame, style='TFrame')
        self.get_measurements_run_button_wrapper.grid(row=0, column=1, sticky="")
        
        self.run_get_measurements_button = ttk.Button(self.get_measurements_run_button_wrapper, text="Run Get Measurements", command=self._run_get_measurements_script, style='TButton')
        self.run_get_measurements_button.pack(padx=5, pady=0)
        
        self.get_measurements_progress_wrapper = ttk.Frame(self.get_measurements_run_control_frame, style='TFrame')
        self.get_measurements_progress_wrapper.grid(row=0, column=1, sticky="ew")

        self.get_measurements_progress_bar = ttk.Progressbar(self.get_measurements_progress_wrapper, orient="horizontal", length=200, mode="determinate")
        self.get_measurements_progress_bar.pack(side="left", fill="x", expand=True, padx=5)
        self.get_measurements_progress_label = ttk.Label(self.get_measurements_progress_wrapper, text="", style='TLabel')
        self.get_measurements_progress_label.pack(side="right", padx=5)

        self.get_measurements_progress_wrapper.grid_remove()


        row_counter += 1

        # NEW SECTION: Move Files from Spreadsheet/List
        move_files_wrapper_frame = ttk.Frame(self.scrollable_frame, style='SectionFrame.TFrame')
        move_files_wrapper_frame.grid(row=row_counter, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
        row_counter += 1

        header_sub_frame_move_files = ttk.Frame(move_files_wrapper_frame, style='TFrame')
        header_sub_frame_move_files.pack(side="top", fill="x", pady=(0, 5), padx=0)
        header_label_move_files = ttk.Label(header_sub_frame_move_files, text="Move Files from Spreadsheet/List", style='Header.TLabel')
        header_label_move_files.pack(side="left", padx=(0, 5))
        info_label_move_files = ttk.Label(header_sub_frame_move_files, text=" ⓘ", font=self.base_font)
        Tooltip(info_label_move_files, "Moves files from a source folder to a destination folder based on a list of filenames provided in an Excel spreadsheet (first column) or a text box.", self.secondary_bg, self.text_color)
        info_label_move_files.pack(side="left", anchor="center")

        move_files_frame = ttk.Frame(move_files_wrapper_frame, style='TFrame')
        move_files_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        move_files_frame.grid_columnconfigure(1, weight=1)

        ttk.Label(move_files_frame, text="Source Folder:", style='TLabel').grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(move_files_frame, textvariable=self.move_files_source_folder, width=45, style='TEntry').grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(move_files_frame, text="Browse", command=lambda: self._browse_folder(self.move_files_source_folder), style='TButton').grid(row=0, column=2, padx=5, pady=5)

        ttk.Label(move_files_frame, text="Destination Folder:", style='TLabel').grid(row=1, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(move_files_frame, textvariable=self.move_files_destination_folder, width=45, style='TEntry').grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(move_files_frame, text="Browse", command=lambda: self._browse_folder(self.move_files_destination_folder), style='TButton').grid(row=1, column=2, padx=5, pady=5)

        input_method_frame_move_files = ttk.Frame(move_files_frame, style='TFrame')
        input_method_frame_move_files.grid(row=2, column=0, columnspan=3, sticky="w", padx=5, pady=(0,5))
        ttk.Label(input_method_frame_move_files, text="Input Method:", style='TLabel').pack(side="left", padx=(0,5))
        ttk.Radiobutton(input_method_frame_move_files, text="From Spreadsheet", variable=self.move_files_input_type, value="spreadsheet", command=lambda: self._show_input_method_move_files("spreadsheet"), style='TRadiobutton').pack(side="left", padx=5)
        ttk.Radiobutton(input_method_frame_move_files, text="From Text Box", variable=self.move_files_input_type, value="textbox", command=lambda: self._show_input_method_move_files("textbox"), style='TRadiobutton').pack(side="left", padx=5)

        self.move_files_spreadsheet_frame = ttk.Frame(move_files_frame, style='TFrame')
        self.move_files_spreadsheet_frame.grid(row=3, column=0, columnspan=3, sticky="ew")
        ttk.Label(self.move_files_spreadsheet_frame, text="Filenames Spreadsheet (.xlsx):", style='TLabel').grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.move_files_excel_path = tk.StringVar()
        ttk.Entry(self.move_files_spreadsheet_frame, textvariable=self.move_files_excel_path, width=45, style='TEntry').grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(self.move_files_spreadsheet_frame, text="Browse", command=lambda: self._browse_file(self.move_files_excel_path, "xlsx"), style='TButton').grid(row=0, column=2, padx=5, pady=5)
        self.move_files_spreadsheet_frame.grid_columnconfigure(1, weight=1)

        self.move_files_textbox_frame = ttk.Frame(move_files_frame, style='TFrame')
        self.move_files_textbox_frame.grid(row=3, column=0, columnspan=3, sticky="nsew")
        ttk.Label(self.move_files_textbox_frame, text="Paste Filenames (one per line):", style='TLabel').pack(padx=5, pady=5, anchor="w")
        self.move_files_text_widget = scrolledtext.ScrolledText(self.move_files_textbox_frame, width=60, height=8, font=self.base_font,
                                                                bg=self.secondary_bg, fg=self.text_color, wrap=tk.WORD,
                                                                insertbackground=self.text_color, relief="solid", borderwidth=1)
        self.move_files_text_widget.pack(padx=5, pady=(0, 5), fill="both", expand=True)

        self.move_files_spreadsheet_frame.grid(row=3, column=0, columnspan=3, sticky="ew")
        self.move_files_textbox_frame.grid_remove()

        self.move_files_run_control_frame = ttk.Frame(move_files_frame, style='TFrame')
        self.move_files_run_control_frame.grid(row=4, column=0, columnspan=3, pady=10, sticky="ew")

        self.move_files_run_control_frame.grid_columnconfigure(0, weight=1)
        self.move_files_run_control_frame.grid_columnconfigure(1, weight=0)
        self.move_files_run_control_frame.grid_columnconfigure(2, weight=1)

        self.move_files_run_button_wrapper = ttk.Frame(self.move_files_run_control_frame, style='TFrame')
        self.move_files_run_button_wrapper.grid(row=0, column=1, sticky="")
        self.run_move_files_button = ttk.Button(self.move_files_run_button_wrapper, text="Run Move Files", command=self._run_move_files_script, style='TButton')
        self.run_move_files_button.pack(padx=5, pady=0)

        self.move_files_progress_wrapper = ttk.Frame(self.move_files_run_control_frame, style='TFrame')
        self.move_files_progress_wrapper.grid(row=0, column=1, sticky="ew")
        self.move_files_progress_bar = ttk.Progressbar(self.move_files_progress_wrapper, orient="horizontal", length=200, mode="determinate")
        self.move_files_progress_bar.pack(side="left", fill="x", expand=True, padx=5)
        self.move_files_progress_label = ttk.Label(self.move_files_progress_wrapper, text="", style='TLabel')
        self.move_files_progress_label.pack(side="right", padx=5)
        self.move_files_progress_wrapper.grid_remove()

        row_counter += 1

        # NEW SECTION: OR Boolean Search Creator
        or_boolean_wrapper_frame = ttk.Frame(self.scrollable_frame, style='SectionFrame.TFrame')
        or_boolean_wrapper_frame.grid(row=row_counter, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
        row_counter += 1

        header_sub_frame_or_boolean = ttk.Frame(or_boolean_wrapper_frame, style='TFrame')
        header_sub_frame_or_boolean.pack(side="top", fill="x", pady=(0, 5), padx=0)
        header_label_or_boolean = ttk.Label(header_sub_frame_or_boolean, text="OR Boolean Search Creator", style='Header.TLabel')
        header_label_or_boolean.pack(side="left", padx=(0, 5))
        info_label_or_boolean = ttk.Label(header_sub_frame_or_boolean, text=" ⓘ", font=self.base_font)
        Tooltip(info_label_or_boolean, "Use this tool to create a boolean search for Bynder with your SKUs separated by “ OR “. This is especially helpful for when you need to search up all assets for a particular list of SKUs, such as when you’re collecting imagery for 3D model projects.", self.secondary_bg, self.text_color)
        info_label_or_boolean.pack(side="left", anchor="center")

        or_boolean_frame = ttk.Frame(or_boolean_wrapper_frame, style='TFrame')
        or_boolean_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        or_boolean_frame.grid_columnconfigure(1, weight=1)

        input_method_frame_or_boolean = ttk.Frame(or_boolean_frame, style='TFrame')
        input_method_frame_or_boolean.grid(row=0, column=0, columnspan=3, sticky="w", padx=5, pady=(0,5))
        ttk.Label(input_method_frame_or_boolean, text="Input Method:", style='TLabel').pack(side="left", padx=(0,5))
        ttk.Radiobutton(input_method_frame_or_boolean, text="From Spreadsheet", variable=self.or_boolean_input_type, value="spreadsheet", command=lambda: self._show_input_method_or_boolean("spreadsheet"), style='TRadiobutton').pack(side="left", padx=5)
        ttk.Radiobutton(input_method_frame_or_boolean, text="From Text Box", variable=self.or_boolean_input_type, value="textbox", command=lambda: self._show_input_method_or_boolean("textbox"), style='TRadiobutton').pack(side="left", padx=5)

        self.or_boolean_spreadsheet_frame = ttk.Frame(or_boolean_frame, style='TFrame')
        self.or_boolean_spreadsheet_frame.grid(row=1, column=0, columnspan=3, sticky="ew")
        ttk.Label(self.or_boolean_spreadsheet_frame, text="SKU Spreadsheet (.xlsx):", style='TLabel').grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(self.or_boolean_spreadsheet_frame, textvariable=self.or_boolean_spreadsheet_path, width=45, style='TEntry').grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ttk.Button(self.or_boolean_spreadsheet_frame, text="Browse", command=lambda: self._browse_file(self.or_boolean_spreadsheet_path, "xlsx"), style='TButton').grid(row=0, column=2, padx=5, pady=5)
        self.or_boolean_spreadsheet_frame.grid_columnconfigure(1, weight=1)

        self.or_boolean_textbox_frame = ttk.Frame(or_boolean_frame, style='TFrame')
        self.or_boolean_textbox_frame.grid(row=1, column=0, columnspan=3, sticky="nsew")
        ttk.Label(self.or_boolean_textbox_frame, text="Paste SKUs (one per line):", style='TLabel').pack(padx=5, pady=5, anchor="w")
        self.or_boolean_text_widget = scrolledtext.ScrolledText(self.or_boolean_textbox_frame, width=60, height=8, font=self.base_font,
                                                                bg=self.secondary_bg, fg=self.text_color, wrap=tk.WORD,
                                                                insertbackground=self.text_color, relief="solid", borderwidth=1)
        self.or_boolean_text_widget.pack(padx=5, pady=(0, 5), fill="both", expand=True)

        self.or_boolean_spreadsheet_frame.grid(row=1, column=0, columnspan=3, sticky="ew")
        self.or_boolean_textbox_frame.grid_remove()

        ttk.Label(or_boolean_frame, text="Results:", style='TLabel').grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.or_boolean_results_textbox = scrolledtext.ScrolledText(or_boolean_frame, width=60, height=5, font=self.base_font,
                                                                     bg=self.secondary_bg, fg=self.text_color, wrap=tk.WORD,
                                                                     insertbackground=self.text_color, relief="solid", borderwidth=1, state='disabled')
        self.or_boolean_results_textbox.grid(row=3, column=0, columnspan=3, padx=5, pady=(0, 5), sticky="nsew")
        or_boolean_frame.grid_rowconfigure(3, weight=1)

        self.or_boolean_run_control_frame = ttk.Frame(or_boolean_frame, style='TFrame')
        self.or_boolean_run_control_frame.grid(row=4, column=0, columnspan=3, pady=10, sticky="ew")

        self.or_boolean_run_control_frame.grid_columnconfigure(0, weight=1)
        self.or_boolean_run_control_frame.grid_columnconfigure(1, weight=0)
        self.or_boolean_run_control_frame.grid_columnconfigure(2, weight=1)

        self.or_boolean_run_button_wrapper = ttk.Frame(self.or_boolean_run_control_frame, style='TFrame')
        self.or_boolean_run_button_wrapper.grid(row=0, column=1, sticky="")
        self.run_or_boolean_button = ttk.Button(self.or_boolean_run_button_wrapper, text="Create OR Boolean Search", command=self._run_or_boolean_script, style='TButton')
        self.run_or_boolean_button.pack(padx=5, pady=0)

        self.or_boolean_progress_wrapper = ttk.Frame(self.or_boolean_run_control_frame, style='TFrame')
        self.or_boolean_progress_wrapper.grid(row=0, column=1, sticky="ew")
        self.or_boolean_progress_bar = ttk.Progressbar(self.or_boolean_progress_wrapper, orient="horizontal", length=200, mode="determinate")
        self.or_boolean_progress_bar.pack(side="left", fill="x", expand=True, padx=5)
        self.or_boolean_progress_label = ttk.Label(self.or_boolean_progress_wrapper, text="", style='TLabel')
        self.or_boolean_progress_label.pack(side="right", padx=5)
        self.or_boolean_progress_wrapper.grid_remove()

        row_counter += 1


        self.log_wrapper_frame = ttk.Frame(self.master, style='SectionFrame.TFrame')
        self.log_wrapper_frame.grid(row=2, column=0, padx=10, pady=5, sticky="nsew")  

        self.log_header_frame = ttk.Frame(self.log_wrapper_frame, style='TFrame')
        self.log_header_frame.pack(fill="x", padx=5, pady=2, side="top")  
        
        log_title_label = ttk.Label(self.log_header_frame, text="Activity Log", font=self.header_font, foreground=self.header_text_color, background=self.secondary_bg)
        log_title_label.pack(side="left", padx=(0, 5))  
        
        self.toggle_log_button = ttk.Button(self.log_header_frame, text="▼", command=self._toggle_log_size, width=2, style='TButton')
        self.toggle_log_button.pack(side="right")


        self.log_text = scrolledtext.ScrolledText(self.log_wrapper_frame, width=90, height=15,  
                                                 font=self.log_font, state='disabled',
                                                 bg=self.log_bg, fg=self.log_text_color,
                                                 insertbackground=self.log_text_color,  
                                                 selectbackground=self.accent_color,  
                                                 selectforeground=self.RF_WHITE_BASE,  
                                                 relief="solid", borderwidth=1)
        self.log_text.pack(padx=10, pady=(0, 10), fill="both", expand=True)  

        if not self.log_expanded:  
            self.log_text.pack_forget()  
            self.toggle_log_button.config(text="▲")  
            self.master.grid_rowconfigure(2, weight=0)  
            self.log_wrapper_frame.config(height=50)  
        
    def _show_input_method_download_psa(self, method):
        """Shows either the spreadsheet input or textbox input for the Download PSAs tool."""
        if method == "spreadsheet":
            self.download_psa_spreadsheet_frame.grid(row=1, column=0, columnspan=3, sticky="ew")
            self.download_psa_textbox_frame.grid_remove()
        else:
            self.download_psa_textbox_frame.grid(row=1, column=0, columnspan=3, sticky="nsew")
            self.download_psa_spreadsheet_frame.grid_remove()
        self.master.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _show_input_method_move_files(self, method):
        """Shows either the spreadsheet input or textbox input for the Move Files tool."""
        if method == "spreadsheet":
            self.move_files_spreadsheet_frame.grid(row=3, column=0, columnspan=3, sticky="ew")
            self.move_files_textbox_frame.grid_remove()
        else:
            self.move_files_textbox_frame.grid(row=3, column=0, columnspan=3, sticky="nsew")
            self.move_files_spreadsheet_frame.grid_remove()
        self.master.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _show_input_method_or_boolean(self, method):
        """Shows either the spreadsheet input or textbox input for the OR Boolean Search Creator tool."""
        if method == "spreadsheet":
            self.or_boolean_spreadsheet_frame.grid(row=1, column=0, columnspan=3, sticky="ew")
            self.or_boolean_textbox_frame.grid_remove()
        else:
            self.or_boolean_textbox_frame.grid(row=1, column=0, columnspan=3, sticky="nsew")
            self.or_boolean_spreadsheet_frame.grid_remove()
        self.master.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))


if __name__ == "__main__":
    root = tk.Tk()
    app = RenamerApp(root)

    creator_frame = ttk.Frame(root, style='TFrame')
    creator_frame.grid(row=3, column=0, sticky="se", padx=10, pady=5)
    creator_label = ttk.Label(creator_frame, text="Created By: Zachary Eisele", font=("Arial", 8), foreground="#888888", background=root.cget('bg'))
    creator_label.pack(side="right", anchor="se")

    root.mainloop()
