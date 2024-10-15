from pptx import Presentation
from pptx.util import Inches
import os

# Specify the folder where your PowerPoint files are located
folder_path = "C:\ConvertedPPTX - Copy"

#List all the .pptx files in the folder
pptx_files = [f for f in os.listdir(folder_path) if f.endswith('.pptx')]

# Function to change background color to white
def change_background_to_white(slide):
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = (255, 255, 255)  # White color

# Loop through each PowerPoint file
for pptx_file in pptx_files:
    pptx_path = os.path.join(folder_path, pptx_file)

    # Open the PowerPoint file
    presentation = Presentation(pptx_path)

    # Loop through all slides in the presentation
    for slide in presentation.slides:
        change_background_to_white(slide)

    # Save the modified PowerPoint file
    presentation.save(pptx_path)

print("Backgrounds changed to white in all PowerPoint files.")

