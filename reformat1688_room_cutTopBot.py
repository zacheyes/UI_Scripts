import os
import sys
import argparse
import tkinter as tk
from tkinter import filedialog
from PIL import Image

def prompt_for_folder_tk(prompt_message):
    """
    Prompts the user to select a folder using a Tkinter dialog.
    """
    root = tk.Tk()
    root.withdraw() # Hide the main window
    folder_path = filedialog.askdirectory(title=prompt_message)
    root.destroy() # Close the Tkinter root window
    return folder_path

def process_image(file_path, output_dir, canvas_size=(3000, 1688), margin=0):
    """
    Processes a single image:
    1. Opens it with alpha to detect transparency.
    2. Finds the bounding box of non-white and non-transparent pixels.
    3. Crops the image to its content.
    4. Scales the cropped content to match the canvas width (height may overflow).
    5. Pastes the resized content onto a white background canvas, aligning left and centering vertically.
    6. Saves the result to the output directory.
    """
    try:
        # Open image with alpha so we can detect transparency
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

        # Compute scale to match canvas width (ignoring height, it may overflow)
        canvas_w, canvas_h = canvas_size
        
        crop_w, crop_h = cropped.size

        # Prevent division by zero if crop_w is 0
        if crop_w == 0:
            print(f"  ➔ Skipped (zero width after crop): {os.path.basename(file_path)}")
            return False

        scale = canvas_w / crop_w  # Scale ONLY by width

        # Resize with high-quality LANCZOS
        new_w, new_h = canvas_w, int(crop_h * scale)  # Force width to match canvas
        resized = cropped.resize((new_w, new_h), Image.LANCZOS)

        # Create white background canvas and paste, aligning to left edge and centering vertically
        canvas = Image.new("RGB", canvas_size, (255, 255, 255))
        paste_x = 0  # Align to left edge
        paste_y = (canvas_h - new_h) // 2  # Center vertically
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
    parser = argparse.ArgumentParser(description="Process and crop room images to 3000x1688, scaling by width (Cut Top/Bottom version).")
    parser.add_argument('--input', help='Path to the input folder containing images to process.')
    
    args = parser.parse_args()

    input_folder = None

    # Determine if input path was provided (from GUI) or if we need to prompt (standalone)
    if not args.input:
        print("\n--- Running Room Cropping Script (3000x1688 Cut Top/Bottom) in Standalone Mode ---")
        input_folder = prompt_for_folder_tk("Select folder with images to process")
        if not input_folder:
            sys.exit("No folder selected. Exiting.")
    else:
        input_folder = args.input
        if not os.path.isdir(input_folder):
            sys.exit(f"Error: Input path '{input_folder}' is not a valid directory. Exiting.")

    # Prepare output directory
    output_dir = os.path.join(input_folder, "Processed_1688_Room_CutTopBot")
    os.makedirs(output_dir, exist_ok=True)

    # Supported extensions (Pillow handles many, but explicitly list common ones)
    exts = (".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".gif")

    print(f"\nProcessing images in: {input_folder}")
    print(f"Output will be saved to: {output_dir}\n")

    processed_count = 0
    skipped_count = 0
    error_count = 0

    image_files = [f for f in sorted(os.listdir(input_folder)) if f.lower().endswith(exts)]

    if not image_files:
        print("No supported image files found in the specified folder. Exiting.")
        sys.exit(0)

    for fname in image_files:
        file_path = os.path.join(input_folder, fname)
        result = process_image(file_path, output_dir)
        if result is True: # Successfully processed
            processed_count += 1
        else: # Skipped due to no content or error during processing
            error_count += 1 


    print(f"\n--- Processing Summary ---")
    print(f"Total image files found: {len(image_files)}")
    print(f"Successfully processed: {processed_count}")
    print(f"Files with errors or skipped (e.g., no content): {error_count}")
    print(f"All done! Processed images are in: {output_dir}")
    
    if error_count > 0:
        sys.exit(1) # Indicate an error occurred
    sys.exit(0) # Indicate success

if __name__ == "__main__":
    main()