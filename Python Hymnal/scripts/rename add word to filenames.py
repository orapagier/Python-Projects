import os

# Specify the directory where your files are located
folder_path = r"D:\Jelmar Orapa\NEMA HYMNS\Data\2 Hymns"

# Function to rename files in a folder and its subfolders
def rename_files_in_folder(folder):
    for root, _, files in os.walk(folder):
        for filename in files:
            # Split the filename and its extension
            name, extension = os.path.splitext(filename)
            
            # Add any word to the name
            new_name = f"{name} PHIL EDITION{extension}"
            
            # Construct the old and new paths
            old_path = os.path.join(root, filename)
            new_path = os.path.join(root, new_name)
            
            # Rename the file
            os.rename(old_path, new_path)
            
            print(f"Renamed: {old_path} to {new_path}")

# Check if the folder path exists
if os.path.exists(folder_path):
    rename_files_in_folder(folder_path)
    print("File names in subfolders have been updated successfully.")
else:
    print(f"The directory '{folder_path}' does not exist.")
