import os

# Specify the folder path where the files are located
folder_path = r"C:\Users\Admin\Desktop\SDAHymnalPPT"

# List all files in the folder
files = os.listdir(folder_path)

# Iterate through each file
for file_name in files:
    # Construct the new file name by removing the first letter
    new_file_name = file_name[1:]

    # Get the full file paths
    old_file_path = os.path.join(folder_path, file_name)
    new_file_path = os.path.join(folder_path, new_file_name)

    # Rename the file
    os.rename(old_file_path, new_file_path)

print("File names have been updated.")
