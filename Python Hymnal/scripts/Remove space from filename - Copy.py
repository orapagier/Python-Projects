import os

# Define the directory path
folder_path = r"C:\Users\Admin\Desktop\White BG"

# List all files in the directory
file_list = os.listdir(folder_path)

# Iterate through the files
for filename in file_list:
    if filename.endswith(".pptx"):
        # Check if the filename starts with a space
        if filename[0] == ' ':
            # Remove the space at the beginning of the filename
            new_filename = filename[1:]
            
            # Create the new file path
            old_path = os.path.join(folder_path, filename)
            new_path = os.path.join(folder_path, new_filename)
            
            # Rename the file
            os.rename(old_path, new_path)
            print(f"Renamed '{filename}' to '{new_filename}'")
            
print("File renaming completed.")
