import os
import comtypes.client

# Specify the folder path where the .ppt files are located
folder_path = r"C:\Users\Admin\Desktop\SDAHymnalPPT"

# Specify the folder path where you want to save the converted .pptx files
output_folder = r"C:\Users\Admin\Desktop\Output pptx"

# Create a PowerPoint application object
powerpoint = comtypes.client.CreateObject("PowerPoint.Application")

# Iterate through the .ppt files in the folder
for filename in os.listdir(folder_path):
    if filename.endswith('.ppt'):
        # Create the full path to the .ppt file
        ppt_file_path = os.path.join(folder_path, filename)

        # Open the .ppt file
        presentation = powerpoint.Presentations.Open(ppt_file_path)

        # Create the new filename with .pptx extension in the output folder
        new_file_name = os.path.splitext(filename)[0] + ".pptx"
        new_file_path = os.path.join(output_folder, new_file_name)

        # Save the presentation in .pptx format to the output folder
        presentation.SaveAs(new_file_path, 24)  # 32 is the value for .pptx format
        presentation.Close()

# Quit PowerPoint application
powerpoint.Quit()
