import os
import shutil

def organize_and_rename_files(input_folder):
    heic_folder = os.path.join(input_folder, "HEIC_Files")
    jpg_folder = os.path.join(input_folder, "JPG_Files")

    # Create subfolders if they don't exist
    os.makedirs(heic_folder, exist_ok=True)
    os.makedirs(jpg_folder, exist_ok=True)

    heic_count = 1
    jpg_count = 1

    for file_name in os.listdir(input_folder):
        file_path = os.path.join(input_folder, file_name)

        # Skip directories
        if os.path.isdir(file_path):
            continue

        # Handle HEIC files
        if file_name.lower().endswith('.heic'):
            new_name = f"image_{heic_count:03d}.heic"
            heic_count += 1
            new_path = os.path.join(heic_folder, new_name)
            shutil.move(file_path, new_path)
            print(f"Moved and renamed: {file_name} -> {new_name}")

        # Handle JPG files
        elif file_name.lower().endswith(('.jpg', '.jpeg')):
            new_name = f"image_{jpg_count:03d}.jpg"
            jpg_count += 1
            new_path = os.path.join(jpg_folder, new_name)
            shutil.move(file_path, new_path)
            print(f"Moved and renamed: {file_name} -> {new_name}")

if __name__ == "__main__":
    input_folder = input("Enter the path to the folder containing HEIC and JPG files: ").strip()

    if not os.path.isdir(input_folder):
        print(f"Error: The folder {input_folder} does not exist or is not a directory.")
    else:
        organize_and_rename_files(input_folder)
        print("Organizing and renaming completed.")
