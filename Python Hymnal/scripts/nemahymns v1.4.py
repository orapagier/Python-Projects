import os
import psutil
import shutil
import tkinter as tk
from tkinter import Scrollbar, Listbox, Entry, Button, Menu, messagebox
from PIL import Image, ImageTk
from tkinter import filedialog


def search_files(event=None):
    search_term = search_entry.get().lower()
    allowed_extensions = [".pps", ".ppsx", ".ppt", ".pptx", ".mp4"]
    search_results = []

    for root, dirs, files in os.walk(dir_path, topdown=True):
        for file in files:
            if any(file.lower().endswith(ext) for ext in allowed_extensions) and search_term in file.lower():
                search_results.append(file)

    result_listbox.delete(0, tk.END)

    if not search_results:
        result_listbox.insert(tk.END, "No hymn found with that word in the title. Try another!")
    else:
        for result in search_results:
            # Remove the file extension from the result before displaying
            result_without_extension = os.path.splitext(result)[0]
            result_listbox.insert(tk.END, result_without_extension)

def open_selected(event):
    selected_item_index = result_listbox.curselection()
    if selected_item_index:
        selected_item = result_listbox.get(selected_item_index)
        selected_file_with_extension = None

        # Recursively search for the selected file in dir_path and its subfolders
        for root, _, files in os.walk(dir_path):
            for file in files:
                if selected_item.lower() in file.lower():
                    selected_file_with_extension = os.path.join(root, file)
                    break
            if selected_file_with_extension:
                break

        if selected_file_with_extension:
            os.startfile(selected_file_with_extension)
            
def update_background():
    global resized_bg_image
    bg_image = Image.open(r"Data\bg.png")  # Replace with your image file path
    resized_bg_image = bg_image.resize((root.winfo_width(), root.winfo_height()), Image.LANCZOS)
    bg_image_tk = ImageTk.PhotoImage(resized_bg_image)
    background_label.config(image=bg_image_tk)
    background_label.image = bg_image_tk

def toggle_focus(event=None):
    if search_entry.focus_get() == search_entry:
        result_listbox.select_set(0)
        result_listbox.focus_set()
    else:
        result_listbox.select_clear(0, tk.END)
        search_entry.focus_set()
        search_files()

def clear_search_entry():
    search_entry.delete(0, tk.END)

def select_next_result(event):
    current_selection = result_listbox.curselection()
    if current_selection:
        next_index = (current_selection[0] + 1) % result_listbox.size()
        if next_index == 0:  # Check if the next index is the first item
            next_index = current_selection[0]  # Keep the selection on the current item
        result_listbox.select_clear(current_selection)
        result_listbox.select_set(next_index)
        result_listbox.event_generate("<<ListboxSelect>>")

def select_previous_result(event):
    current_selection = result_listbox.curselection()
    if current_selection:
        previous_index = current_selection[0] - 1
        if previous_index < 0:
            previous_index = 0
        result_listbox.select_clear(current_selection)
        result_listbox.select_set(previous_index)
        result_listbox.event_generate("<<ListboxSelect>>")

def add_hymns():
    file_paths = filedialog.askopenfilenames(
        title="Select Hymn Files",
        filetypes=[("PowerPoint Files", "*.pps *.ppsx")]
    )
    
    if file_paths:
        hymns_directory = os.path.join(dir_path, "Data", "More Hymns")
        os.makedirs(hymns_directory, exist_ok=True)

        for file_path in file_paths:
            file_name = os.path.basename(file_path)
            destination_path = os.path.join(hymns_directory, file_name)
            try:
                shutil.copy(file_path, destination_path)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to copy {file_name}: {str(e)}")

        messagebox.showinfo("Success", f"{len(file_paths)} hymn(s) added successfully!")


def helps():
    tk.messagebox.showinfo("Help", "Keyboard Shortcuts: \n\nShift (Right): - Switch between search entry and results' list. \nArrow Up/Down: - Select from the results' list up or down. \nEnter: - To open the selected hymn. \nEsc: - To close or exit from the current hymn played. \n\nAdd Hymns: \n\nTo add a hymns that are not on the app's database, \nclick on `File` from the menu bar and select `Add Hymns`, \nthen from the file dialog, select the hymns you want to add. \n\nNote that only .pps or .ppsx file formats are accepted.")
       
def about():
    tk.messagebox.showinfo("About", "Northeastern Mindanao Academy Church. \n\nDeveloper: Jelmar A. Orapa \nEmail: orapajelmar@gmail.com")


def close_all_powerpoint_shows():
    # Create a list to store PowerPoint Show process PIDs
    powerpoint_show_pids = []

    # Iterate through all running processes
    for process in psutil.process_iter(attrs=['pid', 'name', 'cmdline']):
        try:
            process_info = process.info
            process_name = process_info['name'].lower()
            cmdline = process_info['cmdline']

            # Check if the process is PowerPoint and has a .pps or .ppsx file open
            if "powerpnt.exe" in process_name and any(arg.lower().endswith((".pps", ".ppsx")) for arg in cmdline):
                process_pid = process_info['pid']
                powerpoint_show_pids.append(process_pid)

        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass

    # Terminate all identified PowerPoint Show processes
    for pid in powerpoint_show_pids:
        psutil.Process(pid).terminate()
   
dir_path = os.path.dirname(os.path.realpath(__file__))

root = tk.Tk()
root.title("Northeastern Mindanao Academy Church")

window_width = 510
window_height = 322
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x_position = (screen_width - window_width) // 2
y_position = (screen_height - window_height) // 2
root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")
root.resizable(False, False)

background_label = tk.Label(root)
background_label.place(relwidth=1, relheight=1)
update_background()





# Create a menu bar
menu_bar = Menu(root)
root.config(menu=menu_bar)

# Create a File menu
file_menu = Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="File", menu=file_menu)
file_menu.add_command(label="Add Hymns", command=lambda: add_hymns())
file_menu.add_separator()
file_menu.add_command(label="Exit", command=root.destroy)

# Create a Help menu
help_menu = Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="More", menu=help_menu)
help_menu.add_command(label="Help", command=helps)
help_menu.add_command(label="About", command=about)

# Create a button in the menu bar for closing PowerPoint shows
menu_bar.add_command(label="Close Hymn", command=close_all_powerpoint_shows)



search_entry = Entry(root, highlightbackground="white", highlightthickness=1)
search_entry.grid(row=0, column=1, padx=0, pady=0)
search_entry.bind("<Return>", search_files)
search_entry.focus_set()

search_button = Button(root, text="Search", command=search_files)
search_button.grid(row=0, column=2, padx=5, pady=0)

result_listbox = Listbox(root, selectmode=tk.SINGLE, borderwidth=0, highlightthickness=0)
scrollbar = Scrollbar(root, orient=tk.VERTICAL)
scrollbar.config(command=result_listbox.yview)
result_listbox.config(yscrollcommand=scrollbar.set, font=("Times New Roman", 12))
scrollbar.grid(row=1, column=1, padx=0, pady=(0, 24), sticky="ns", rowspan=3)

search_entry.place(in_=result_listbox, x=0, y=0, relx=0.7, relwidth=0.2, relheight=0.1)
search_button.place(in_=result_listbox, x=1, y=0, relx=0.885, relwidth=0.1, relheight=0.1)
search_entry.lift()
search_button.lift()
result_listbox.grid(row=1, column=0, padx=10, pady=(0, 24), sticky="nsew", rowspan=3, columnspan=3)

result_listbox.bind("<Double-Button-1>", open_selected)
result_listbox.bind("<Return>", open_selected)

#add some space
#root.grid_rowconfigure(4, weight=1)

root.grid_rowconfigure(1, weight=1)
root.grid_columnconfigure(0, weight=1)

root.bind("<Configure>", lambda event: update_background())
root.bind("<Shift_R>", lambda event: [toggle_focus(), clear_search_entry()])
root.bind("<Up>", select_previous_result)
root.bind("<Down>", select_next_result)

search_files()

root.mainloop()
