import os

# Specify the folder path where the files are located
folder_path = r"C:\Users\Admin\Desktop\White BG 1"

# Get a list of all files in the folder
file_list = os.listdir(folder_path)

# Iterate through the files and rename .ppt to .pps
for file_name in file_list:
    if file_name.endswith('.pps'):
        # Split the filename to remove the .ppt extension
        base_name = os.path.splitext(file_name)[0]
        
        # Create the new filename with .pps extension
        new_file_name = base_name + '.ppt'
        
        # Create the full path for the old and new filenames
        old_file_path = os.path.join(folder_path, file_name)
        new_file_path = os.path.join(folder_path, new_file_name)
        
        # Rename the file
        os.rename(old_file_path, new_file_path)
        print(f'Renamed: {file_name} to {new_file_name}')

