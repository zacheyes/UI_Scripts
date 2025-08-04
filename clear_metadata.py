# Save this as clear_metadata.py
import os
import subprocess
import sys
import argparse
import shutil
import tkinter as tk
from tkinter import filedialog

# --- Configuration for Metadata Fields ---

# Mapping of user-friendly names to a list of ExifTool tags.
# This dictionary has been updated to include common EXIF, IPTC, and XMP tags
# for each metadata property, ensuring comprehensive clearing.
METADATA_PROPERTIES = {
    # Description: Targets common EXIF, IPTC, and XMP fields used for general descriptions and captions.
    "Description": [
        "-Description=",            # Generic ExifTool tag for description
        "-ImageDescription=",       # EXIF Image Description 
        "-Caption-Abstract=",       # IPTC Caption-Abstract 
        "-XMP-dc:Description="      # XMP Dublin Core Description 
    ],

    # Keywords: Targets common IPTC and XMP fields for keywords and subjects.
    "Keywords": [
        "-Keywords=",               # IPTC Keywords 
        "-XMP-dc:Subject=",         # XMP Dublin Core Subject 
        "-XMP:photoshop:SupplementalCategories=", # Photoshop's Supplemental Categories
        "-XMP:lr:HierarchicalSubject=", # Lightroom's Hierarchical Subject
        "-TagList="                 # Bynder's virtual TagList (if applicable, not found in provided dump)
    ],

    # Title: Targets common IPTC and XMP fields for the image title or object name.
    "Title": [
        "-Title=",                  # Generic ExifTool tag for title
        "-DocumentTitle=",          # Document Title (sometimes used as a general title)
        "-ObjectName=",             # IPTC Object Name (traditional IPTC title) 
        "-XMP-dc:Title="            # XMP Dublin Core Title 
    ],

    # 'Generated Image' in Title: Specifically targets AI-related metadata or history entries.
    "'Generated Image' in Title": [
        "-JUMBF:all=",              # Targets all JUMBF metadata, often used for content provenance
        "-XMP-xmpMM:History="       # XMP Media Management History (can contain AI generation info) 
    ],

    # Headline: Targets common IPTC and XMP fields for a news headline.
    "Headline": [
        "-Headline=",               # IPTC Headline
        "-XMP-photoshop:Headline="  # XMP Photoshop Headline
    ],

    # Event: Targets common IPTC and XMP fields for the event associated with the image.
    "Event": [
        "-Event=",                  # IPTC Event
        "-XMP-photoshop:Event="     # XMP Photoshop Event
    ],
}

# --- Special command for aggressive stripping ---
AGGRESSIVE_STRIP_FLAG = "--STRIP_ALL_METADATA_EXCEPT_ICC--"

def get_exiftool_executable_path():
    """
    Finds the correct ExifTool executable based on the operating system.
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Check for the OS-specific bundled executable first
    if sys.platform == "win32":
        # On Windows, look for exiftool.exe
        portable_path = os.path.join(script_dir, 'tools', 'exiftool_PC', 'exiftool.exe')
        if os.path.exists(portable_path):
            return portable_path
    elif sys.platform == "darwin":
        # On macOS, look for the Unix executable
        portable_path = os.path.join(script_dir, 'tools', 'exiftool_MAC', 'exiftool')
        if os.path.exists(portable_path):
            return portable_path

    # If a bundled version isn't found, check the system PATH
    exiftool_in_path = shutil.which("exiftool")
    if exiftool_in_path:
        return exiftool_in_path

    return None # ExifTool not found


def clear_metadata_from_image_exiftool(file_path, exiftool_args_to_run):
    """
    Clears metadata from a single image file using ExifTool.
    """
    print(f"\n  Processing: {os.path.basename(file_path)}")
    
    exiftool_path = get_exiftool_executable_path()
    if not exiftool_path:
        sys.stderr.write(f"  Error: ExifTool not found. Please ensure the correct version is in the 'tools' subdirectory or system PATH.\n")
        sys.stderr.write(f"  Download from: https://exiftool.org/\n")
        return

    print(f"  Using ExifTool from: {exiftool_path}")

    command = [exiftool_path]

    # Determine command based on arguments
    if exiftool_args_to_run and exiftool_args_to_run[0] == AGGRESSIVE_STRIP_FLAG:
        print("  Applying AGGRESSIVE STRIP: Removing all metadata except the ICC Profile...")
        command.extend(['-all=', '--ICC_Profile'])
    elif exiftool_args_to_run:
        print(f"  Clearing standard tags: {', '.join(exiftool_args_to_run)}")
        command.extend(exiftool_args_to_run)
    else:
        print(f"  No metadata fields selected for clearing. Skipping ExifTool operation.")
        return

    command.append("-overwrite_original")
    command.append(file_path)

    try:
        # Run ExifTool command
        process = subprocess.run(command, capture_output=True, text=True, check=True, encoding='utf-8')
        print(f"  ExifTool stdout: {process.stdout.strip()}")
        if process.stderr:
            sys.stderr.write(f"  ExifTool stderr: {process.stderr.strip()}\n")
        print(f"  Successfully processed metadata for: {os.path.basename(file_path)}")
    except subprocess.CalledProcessError as e:
        sys.stderr.write(f"  Error running ExifTool for {os.path.basename(file_path)}: {e}\n")
        sys.stderr.write(f"  ExifTool stderr: {e.stderr.strip()}\n")
    except Exception as e:
        sys.stderr.write(f"  An unexpected error occurred for {os.path.basename(file_path)}: {e}\n")


def get_user_choices_interactive():
    """
    Interactively prompts the user for which metadata properties to clear.
    """
    print("\n--- Metadata Configuration (Interactive Mode) ---")
    print("For each property, do you want to CLEAR it? (y/n)")
    print("-------------------------------------------------")

    final_tags_to_clear_exiftool = []
    
    # Sort keys for consistent prompting order
    sorted_prop_keys = sorted(METADATA_PROPERTIES.keys())
    
    for prop_display_name in sorted_prop_keys:
        exiftool_tags = METADATA_PROPERTIES[prop_display_name]
        user_input = ""
        while user_input.lower() not in ['y', 'n']:
            user_input = input(f"Clear '{prop_display_name}'? (y/n): ").strip().lower()
            if user_input == 'y':
                final_tags_to_clear_exiftool.extend(exiftool_tags)
                print(f"  -> '{prop_display_name}' will be CLEARED.")
            elif user_input == 'n':
                print(f"  -> '{prop_display_name}' will be KEPT.")
            else:
                print("Invalid input. Please enter 'y' or 'n'.")
    
    return list(set(final_tags_to_clear_exiftool))


def main():
    sorted_prop_keys = sorted(METADATA_PROPERTIES.keys())
    available_properties_str = ", ".join(sorted_prop_keys)

    parser = argparse.ArgumentParser(
        description="Clear specific embedded metadata from image files using ExifTool.",
        formatter_class=argparse.RawTextHelpFormatter
    )
    parser.add_argument("--input_folder", help="Path to the folder containing image files.")
    parser.add_argument("--clear_properties", nargs='*',
                        help=f"Space-separated list of metadata properties to clear.\n"
                             f"Options: {available_properties_str}.\n"
                             f"If used, script runs non-interactively.")
    parser.add_argument("--strip_ai_metadata", action='store_true',
                        help="DANGER: Aggressively strips ALL metadata, keeping only the ICC color profile.\n"
                             "Use this for removing stubborn 'Generated Image' data from AI-edited files.\n"
                             "This option overrides --clear_properties.")
    args = parser.parse_args()

    tags_to_clear_for_all_files = []
    input_folder = args.input_folder

    # --- Determine Operation Mode ---

    # 1. Aggressive Strip Mode (highest priority)
    if args.strip_ai_metadata:
        print("\n" + "="*50)
        print("     D A N G E R O U S   O P E R A T I O N")
        print("="*50)
        print("Selected: Aggressively strip ALL metadata except ICC Profile.")
        print("This will override any other selections.")
        print("="*50 + "\n")
        tags_to_clear_for_all_files = [AGGRESSIVE_STRIP_FLAG]
        if not input_folder:
            sys.stderr.write("Error: --strip_ai_metadata requires the --input_folder argument.\n")
            sys.exit(1)

    # 2. Non-Interactive (CLI) Mode
    elif input_folder and args.clear_properties is not None:
        print("\n--- Metadata Configuration (Command Line) ---")
        for prop_name in args.clear_properties:
            if prop_name in METADATA_PROPERTIES:
                tags_to_clear_for_all_files.extend(METADATA_PROPERTIES[prop_name])
                print(f"  -> '{prop_name}' will be CLEARED.")
            else:
                sys.stderr.write(f"Warning: Unknown property '{prop_name}' specified. Skipping.\n")
        tags_to_clear_for_all_files = list(set(tags_to_clear_for_all_files))

    # 3. Fully Interactive Mode
    else:
        print("\n--- Interactive Mode ---")
        if not input_folder:
            root = tk.Tk()
            root.withdraw()
            input_folder = filedialog.askdirectory(title="Select Folder to Process")
            root.destroy()
            if not input_folder:
                print("No folder selected. Exiting.")
                sys.exit(0)
        
        print(f"Selected folder: {input_folder}")
        tags_to_clear_for_all_files = get_user_choices_interactive()

    # --- Execute Processing ---

    if not os.path.isdir(input_folder):
        sys.stderr.write(f"Error: Input folder not found at '{input_folder}'\n")
        sys.exit(1)

    image_files = [f for f in os.listdir(input_folder) if f.lower().endswith(('.jpg', '.jpeg', '.png', '.tif', '.tiff'))]
    
    if not image_files:
        print(f"No supported image files found in '{input_folder}'.")
        sys.exit(0)

    if not tags_to_clear_for_all_files:
        print("\nNo metadata fields selected for clearing. Exiting.")
        sys.exit(0)
    
    print("\n--- Starting File Processing ---")
    total_files = len(image_files)
    for i, filename in enumerate(image_files):
        file_path = os.path.join(input_folder, filename)
        clear_metadata_from_image_exiftool(
            file_path,
            exiftool_args_to_run=tags_to_clear_for_all_files
        )
        # Report progress to stdout for the calling UI
        progress = (i + 1) / total_files * 100
        sys.stdout.write(f"PROGRESS:{progress:.2f}\n")
        sys.stdout.flush()

    print(f"\nFinished processing {len(image_files)} files.")
    sys.exit(0)


if __name__ == "__main__":
    main()
