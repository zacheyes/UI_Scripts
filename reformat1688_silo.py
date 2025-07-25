import os
import sys
import argparse
import tkinter as tk
from tkinter import filedialog
from PIL import Image

def prompt_for_folder_tk(prompt_message):
    """Prompt the user to select a folder using a Tkinter dialog."""
    root = tk.Tk()
    root.withdraw() # Hide the main window
    folder_path = filedialog.askdirectory(title=prompt_message)
    root.destroy() # Close the Tkinter root window
    return folder_path

def process_image(file_path, output_dir, canvas_size=(3000, 1688), margin=10):
    """
    Processes a single image:
    1. Opens it with alpha to detect transparency.
    2. Finds the bounding box of non-white and non-transparent pixels.
    3. Crops the image to its content.
    4. Scales the cropped content to fit within the specified canvas size, respecting margins.
    5. Pastes the resized content onto a white background canvas, centered.
    6. Saves the result to the output directory.
    Returns True for success, False for skip/error.
    """
    try:
        # Open image with alpha so we can detect transparency too
        im = Image.open(file_path).convert("RGBA")
        width, height = im.size
        pixels = im.load()

        # Find bounding box of all non-white (and non-transparent) pixels
        min_x, min_y = width, height
        max_x, max_y = 0, 0
        found_content = False
        for y in range(height):
            for x in range(width):
                r, g, b, a = pixels[x, y]
                # Check if pixel is not fully transparent AND not pure white
                if a != 0 and not (r == 255 and g == 255 and b == 255):
                    found_content = True
                    min_x = min(min_x, x)
                    min_y = min(min_y, y)
                    max_x = max(max_x, x)
                    max_y = max(max_y, y)

        # If no non-white/non-transparent pixels found, skip
        if not found_content or min_x > max_x or min_y > max_y:
            print(f"  ➔ Skipped (no discernible content): {os.path.basename(file_path)}")
            return False

        # Crop to content
        cropped = im.crop((min_x, min_y, max_x + 1, max_y + 1))

        # Compute scale to fit within canvas minus margins
        canvas_w, canvas_h = canvas_size
        content_area_w = canvas_w - 2 * margin
        content_area_h = canvas_h - 2 * margin

        crop_w, crop_h = cropped.size

        # Prevent division by zero if crop_w or crop_h is 0 (though found_content should prevent this)
        if crop_w == 0 or crop_h == 0:
            print(f"  ➔ Skipped (zero dimension after crop): {os.path.basename(file_path)}")
            return False

        scale = min(content_area_w / crop_w, content_area_h / crop_h)

        # Resize with high-quality LANCZOS
        new_w, new_h = int(crop_w * scale), int(crop_h * scale)
        resized = cropped.resize((new_w, new_h), Image.LANCZOS)

        # Create white background canvas and center the resized crop
        canvas = Image.new("RGB", canvas_size, (255, 255, 255))
        paste_x = (canvas_w - new_w) // 2
        paste_y = (canvas_h - new_h) // 2
        canvas.paste(resized, (paste_x, paste_y), mask=resized) # Use mask for proper alpha blending

        # Save to output folder
        out_path = os.path.join(output_dir, os.path.basename(file_path))
        canvas.save(out_path)
        print(f"  ✔ Processed and saved: {os.path.basename(file_path)}")
        return True

    except Exception as e:
        print(f"  ✖ Error processing {os.path.basename(file_path)}: {e}")
        return False

def main():
    parser = argparse.ArgumentParser(description="Process and crop silo-style images to 3000x1688 with content auto-cropping.")
    parser.add_argument('--input', help='Path to the input folder containing images to process.')
    
    args = parser.parse_args()

    input_folder = None

    # Determine if input path was provided (from GUI) or if we need to prompt (standalone)
    if not args.input:
        print("\n--- Running Silo Cropping Script in Standalone Mode ---")
        input_folder = prompt_for_folder_tk("Select folder with images to process")
        if not input_folder:
            sys.exit("No input folder selected. Exiting.")
    else:
        input_folder = args.input
        if not os.path.isdir(input_folder):
            sys.exit(f"Error: Input path '{input_folder}' is not a valid directory. Exiting.")

    # Prepare output directory
    output_dir = os.path.join(input_folder, "Processed_1688_Silo")
    os.makedirs(output_dir, exist_ok=True)

    # Supported extensions
    exts = (".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".gif")

    print(f"\nProcessing images in: {input_folder}")
    print(f"Output will be saved to: {output_dir}\n")

    processed_count = 0
    skipped_count = 0
    error_count = 0

    image_files = [f for f in sorted(os.listdir(input_folder)) if f.lower().endswith(exts)]

    if not image_files:
        print("No supported image files found in the specified folder. Exiting.")
        # Report 0% progress if no files to process
        print("PROGRESS: 0")
        sys.exit(0)

    total_files = len(image_files)

    for i, fname in enumerate(image_files):
        file_path = os.path.join(input_folder, fname)
        
        # Call the image processing function
        success = process_image(file_path, output_dir)
        
        if success:
            processed_count += 1
        else:
            # Currently, process_image returns False for both skips and errors.
            # For a more accurate summary, you'd need process_image to return an enum or specific codes.
            # For this example, we'll increment error_count for simplicity if it wasn't successful.
            # If you want to differentiate skips, you'd modify process_image to return a distinct value for "skipped".
            error_count += 1 # Or add a specific skipped_count if process_image returns a different signal for skips

        # Calculate progress and print for the UI
        progress_percent = (i + 1) / total_files * 100
        print(f"PROGRESS: {progress_percent:.1f}")
        sys.stdout.flush() # Ensure the progress is sent immediately

    print(f"\n--- Processing Summary ---")
    print(f"Total image files found: {total_files}")
    print(f"Successfully processed: {processed_count}")
    # Note: If process_image returns False for both skips and errors, error_count will include skips.
    print(f"Failed or Skipped: {error_count}") # Combined for simplicity without deeper return logic in process_image
    print(f"All done! Processed images are in: {output_dir}")
    
    if error_count > 0:
        sys.exit(1) # Indicate an error occurred
    sys.exit(0) # Indicate success

if __name__ == "__main__":
    main()
