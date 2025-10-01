# **HEIC to JPG Converter & Enhancer**

## **Overview**

This Python script is your all-in-one tool for **converting HEIC images to high-quality JPGs** and performing **automated image enhancement**. Designed with programmers in mind, it lets you preserve image quality while adding a touch of magic to make your photos pop! No GUI, no bloated tools—just a clean, efficient script to get the job done.

## **Features**

- **HEIC to JPG Conversion**: Converts `.heic` images to `.jpg` format while preserving maximum quality (`quality=100` and `subsampling=0`).
- **Image Enhancement**: Automatically adjusts:
  - **Brightness** (+20%)
  - **Contrast** (+30%)
  - **Sharpness** (+50%)
- **Optional Resizing**: Scale images to your preferred width while maintaining the aspect ratio.
- **File Organization**:
  - Original JPGs saved in `Originals` folder.
  - Enhanced JPGs saved in `Enhanced` folder.

## **Requirements**

- **Python 3.7+**
- Libraries:
  - `Pillow`
  - `pillow-heif`

Install dependencies:

```bash
pip install pillow pillow-heif
```

## **Usage**

1. Clone or copy this script to your machine.
2. Run the script in a terminal:
   ```bash
   python convert_and_enhance.py
   ```
3. Enter the following when prompted:
   - Path to the folder containing your `.heic` files.
   - Path to the output folder where the converted and enhanced files will be saved.
   - (Optional) Resize width (leave blank to skip resizing).

## **Example Workflow**

Suppose you have a folder `D:/Photos/HEIC_Images` containing:

- `vacation1.heic`
- `vacation2.heic`

### Input

When prompted:

- Input folder: `D:/Photos/HEIC_Images`
- Output folder: `D:/Photos/Processed_Images`
- Resize width: `1920`

### Output

- **Originals**:
  - `D:/Photos/Processed_Images/Originals/vacation1.jpg`
  - `D:/Photos/Processed_Images/Originals/vacation2.jpg`
- **Enhanced**:
  - `D:/Photos/Processed_Images/Enhanced/vacation1.jpg`
  - `D:/Photos/Processed_Images/Enhanced/vacation2.jpg`

---

## **Why You'll Love This Script**

- **No quality compromise**: Your photos retain their original brilliance.
- **No bloat**: Pure Python magic, no unnecessary tools.
- **Fully customizable**: Adjust enhancement factors, resizing, or folder structure to fit your needs.

---

## **Customization Tips**

- **Tweak enhancements**: Modify `enhance_image()` for different levels of brightness, contrast, or sharpness.
- **Change output settings**: Adjust output paths, naming conventions, or JPG compression levels as needed.

---

## **License**

This script is free and open-source. Use it, modify it, and make it your own! 🚀

---

### Written by a Programmer, for Programmers ✨
