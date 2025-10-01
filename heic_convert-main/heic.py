from PIL import Image
import pillow_heif
import os

def convert_heic_to_jpg(input_folder, output_folder):
    # Ensure output folder exists
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for file_name in os.listdir(input_folder):
        # Process only .heic files
        if file_name.lower().endswith('.heic'):
            input_path = os.path.join(input_folder, file_name)
            output_file_name = os.path.splitext(file_name)[0] + '.jpg'
            output_path = os.path.join(output_folder, output_file_name)

            try:
                # Open HEIC file with pillow-heif
                heif_file = pillow_heif.open_heif(input_path)
                image = Image.frombytes(heif_file.mode, heif_file.size, heif_file.data)

                # Save as JPEG with maximum quality
                image.save(output_path, "JPEG", quality=100, subsampling=0)
                print(f"Converted: {file_name} -> {output_file_name}")
            except Exception as e:
                print(f"Failed to convert {file_name}: {e}")

if __name__ == "__main__":
    input_folder = input("Enter the path to the folder containing HEIC files: ").strip()
    output_folder = input("Enter the path to the output folder for JPG files: ").strip()

    if not os.path.isdir(input_folder):
        print(f"Error: The folder {input_folder} does not exist or is not a directory.")
    else:
        convert_heic_to_jpg(input_folder, output_folder)
        print("Conversion process completed.")
