import os
from PIL import Image, ImageEnhance
import pillow_heif


def enhance_image(image):
    """
    Enhance the image by adjusting brightness, contrast, and sharpness.
    """
    enhancer = ImageEnhance.Brightness(image)
    image = enhancer.enhance(1.05)  # Increase brightness by 20%

    # enhancer = ImageEnhance.Contrast(image)
    # image = enhancer.enhance(1.03)  # Increase contrast by 30%

    enhancer = ImageEnhance.Sharpness(image)
    image = enhancer.enhance(1.05)  # Increase sharpness by 50%

    return image


def convert_and_enhance(input_folder, output_folder, resize_width=None):
    """
    Convert HEIC images to JPG, enhance them, and save with maximum quality.
    """
    original_folder = os.path.join(output_folder, "Originals")
    enhanced_folder = os.path.join(output_folder, "Enhanced")

    # Ensure output subfolders exist
    os.makedirs(original_folder, exist_ok=True)
    os.makedirs(enhanced_folder, exist_ok=True)

    for file_name in os.listdir(input_folder):
        if file_name.lower().endswith('.heic'):
            input_path = os.path.join(input_folder, file_name)
            output_file_name = os.path.splitext(file_name)[0] + '.jpg'
            original_path = os.path.join(original_folder, output_file_name)
            enhanced_path = os.path.join(enhanced_folder, output_file_name)

            try:
                # Convert HEIC to Image object
                heif_file = pillow_heif.read_heif(input_path)
                image = Image.frombytes(
                    heif_file.mode,
                    heif_file.size,
                    heif_file.data
                )

                # Save the original as JPG
                image.save(original_path, "JPEG", quality=100, subsampling=0)
                print(f"Converted: {file_name} -> {output_file_name}")

                # Enhance the image
                enhanced_image = enhance_image(image)

                # Resize if specified
                if resize_width:
                    aspect_ratio = image.height / image.width
                    new_height = int(resize_width * aspect_ratio)
                    enhanced_image = enhanced_image.resize((resize_width, new_height), Image.ANTIALIAS)

                # Save the enhanced image
                enhanced_image.save(enhanced_path, "JPEG", quality=100, subsampling=0)
                print(f"Enhanced and saved: {file_name} -> {output_file_name}")

            except Exception as e:
                print(f"Failed to process {file_name}: {e}")


if __name__ == "__main__":
    input_folder = input("Enter the path to the folder containing HEIC files: ").strip()
    output_folder = input("Enter the path to the output folder: ").strip()
    resize_width = input("Enter the width to resize images (or press Enter to skip resizing): ").strip()

    if resize_width.isdigit():
        resize_width = int(resize_width)
    else:
        resize_width = None

    if not os.path.isdir(input_folder):
        print(f"Error: The folder {input_folder} does not exist or is not a directory.")
    else:
        convert_and_enhance(input_folder, output_folder, resize_width)
        print("Conversion and enhancement process completed.")
