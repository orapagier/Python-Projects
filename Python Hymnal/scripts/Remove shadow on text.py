import os
import win32com.client as win32

def run_vba_script_on_pptx(pptx_file_path, vba_script_path):
    try:
        # Create a PowerPoint application object
        ppt_app = win32.gencache.EnsureDispatch("PowerPoint.Application")

        # Open the PowerPoint presentation
        presentation = ppt_app.Presentations.Open(pptx_file_path)

        # Load the VBA module from the .bas file
        presentation.VBProject.VBComponents.Import(vba_script_path)

        # Run the VBA macro from the module
        ppt_app.Run("Module1.RemoveShadows")  # Replace YourMacroName with the actual macro name

        # Save and close the presentation
        presentation.Save()
        presentation.Close()

        print(f"VBA script executed successfully on {pptx_file_path}")
    except Exception as e:
        print(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    folder_path = r"C:\Users\Admin\Desktop\White BG"  # Specify the folder containing your .pptx files
    vba_script_path = r"C:\Users\Admin\Desktop\Module1.bas"

    # Iterate through all .pptx files in the folder
    for filename in os.listdir(folder_path):
        if filename.endswith(".pptx"):
            pptx_file_path = os.path.join(folder_path, filename)
            run_vba_script_on_pptx(pptx_file_path, vba_script_path)
