import os
import pythoncom
import threading
import win32com.client
import tkinter as tk
from tkinter import Scrollbar, Listbox, Entry, Button, Menu, messagebox
from PIL import Image, ImageTk
from tkinter import filedialog
from tkinter import StringVar
from screeninfo import get_monitors
import shutil


dir_path = os.path.dirname(os.path.realpath(__file__))

def search_files(event=None):
    search_term = search_var.get().lower()
    allowed_extensions = [".pps", ".ppsx", ".ppt", ".pptx"]
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
            result_without_extension = os.path.splitext(result)[0]
            result_listbox.insert(tk.END, result_without_extension)

def quit_powerpoint():
    clear_search_entry()
    search_entry.focus_set()

    def quit_powerpoint_thread():
        try:
            pythoncom.CoInitialize()
            ppt = win32com.client.Dispatch('PowerPoint.Application')
            ppt.Quit()
        except Exception as e:
            print("Error quitting PowerPoint:", e)
        finally:
            pythoncom.CoUninitialize()
        root.after(0, search_files)

    threading.Thread(target=quit_powerpoint_thread, daemon=True).start()

def clear_search_entry():
    search_entry.delete(0, tk.END)

def open_selected(event):
    selected_item_index = result_listbox.curselection()
    if selected_item_index:
        selected_item = result_listbox.get(selected_item_index)
        selected_file_with_extension = None

        for root, _, files in os.walk(dir_path):
            for file in files:
                if selected_item.lower() in file.lower():
                    selected_file_with_extension = os.path.join(root, file)
                    break
            if selected_file_with_extension:
                break

        if selected_file_with_extension:
            try:
                ppt = win32com.client.Dispatch('PowerPoint.Application')
                ppt.Visible = True
                presentation = ppt.Presentations.Open(selected_file_with_extension, WithWindow=True)
                presentation.SlideShowSettings.AdvanceMode = 1
                presentation.SlideShowSettings.ShowType = 1
                presentation.SlideShowSettings.Run()
                ppt.WindowState = 2
            except Exception as e:
                print("Error opening presentation in Presenter View:", e)
                return

def update_background():
    global resized_bg_image
    bg_image = Image.open(r"_internal/Data\bg.png")
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

def select_next_result(event):
    current_selection = result_listbox.curselection()
    if current_selection:
        next_index = (current_selection[0] + 1) % result_listbox.size()
        if next_index == 0:
            next_index = current_selection[0]
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
        filetypes=[("PowerPoint Files", "*.pps *.ppsx")])

    if file_paths:
        hymns_directory = os.path.join(dir_path, "Data", "4 More Hymns")
        os.makedirs(hymns_directory, exist_ok=True)

        for file_path in file_paths:
            file_name = os.path.basename(file_path)
            destination_path = os.path.join(hymns_directory, file_name)
            try:
                shutil.copy(file_path, destination_path)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to copy {file_name}: {str(e)}")

        messagebox.showinfo("Success", f"{len(file_paths)} hymn(s) added successfully!")

def apphelp():
    tk.messagebox.showinfo("Help", "Keyboard Shortcuts: \n\nShift (Right): - Switch between search entry and results' list. \nArrow Up/Down: - Select from the results' list up or down. \nEnter: - To open the selected hymn. \nEsc: - To close or exit from the current hymn played. \n\nAdd Hymns: \n\nTo add hymns that are not on the app's database, \nclick on `File` from the menu bar and select `Add Hymns`, \nthen from the file dialog, select the hymns you want to add. \n\nNote that only .pps or .ppsx file formats are accepted.")

def about():
    tk.messagebox.showinfo("About", "Seventh Day Adventist Church Hymnal. \n\nDeveloper: Jelmar A. Orapa \nEmail: orapajelmar@gmail.com")

def delete_temp_folder():
    try:
        temp_folder = os.path.join(os.getcwd(), "Temp")
        if os.path.exists(temp_folder):
            shutil.rmtree(temp_folder)
            tk.messagebox.showinfo("Success", "Temp folder deleted successfully!")
        else:
            tk.messagebox.showinfo("Info", "Temp folder does not exist.")
    except Exception as e:
        tk.messagebox.showerror("Error", f"Failed to delete Temp folder: {str(e)}")


root = tk.Tk()
root.title("Seventh Day Adventist Church Hymnal")

icon_image = Image.open("_internal/Data/favicon.ico")
icon_photo = ImageTk.PhotoImage(icon_image)
root.winfo_toplevel().iconphoto(True, icon_photo)

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

menu_bar = Menu(root)
root.config(menu=menu_bar)

file_menu = Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="File", menu=file_menu)
file_menu.add_command(label="Add Hymns", command=lambda: add_hymns())
file_menu.add_command(label="Delete Temporary Files", command=lambda: delete_temp_folder())
file_menu.add_command(label="Exit", command=root.destroy)

help_menu = Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="More", menu=help_menu)
help_menu.add_command(label="Help", command=apphelp)
help_menu.add_command(label="About", command=about)

separator = Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="I", menu=separator)
menu_bar.add_command(label="Clear", command=quit_powerpoint)

search_var = StringVar()
search_entry = Entry(root, highlightbackground="white", highlightthickness=1, textvariable=search_var)
search_entry.grid(row=0, column=1, padx=0, pady=0)
search_var.trace_add("write", lambda *args: search_files())
search_entry.focus_set()

search_button = Button(root, text="Search", command=search_files)
search_button.grid(row=0, column=2, padx=5, pady=0)

result_listbox = Listbox(root, selectmode=tk.SINGLE, borderwidth=0, highlightthickness=0)
scrollbar = Scrollbar(root, orient=tk.VERTICAL)
scrollbar.config(command=result_listbox.yview)
result_listbox.config(yscrollcommand=scrollbar.set, font=("Times New Roman", 12))
scrollbar.grid(row=1, column=1, padx=0, pady=(0, 24), sticky="ns", rowspan=3)

search_entry.place(in_=result_listbox, x=0, y=0, relx=0.7, relwidth=0.2, relheight=0.07)
search_button.place(in_=result_listbox, x=1, y=0, relx=0.885, relwidth=0.1, relheight=0.07)
search_entry.lift()
search_button.lift()
result_listbox.grid(row=1, column=0, padx=10, pady=(0, 24), sticky="nsew", rowspan=3, columnspan=3)

result_listbox.bind("<Double-Button-1>", open_selected)
result_listbox.bind("<Return>", open_selected)

root.grid_rowconfigure(1, weight=1)
root.grid_columnconfigure(0, weight=1)

root.bind("<Configure>", lambda event: update_background())
root.bind("<Shift_R>", lambda event: toggle_focus())
root.bind("<Up>", select_previous_result)
root.bind("<Down>", select_next_result)

search_files()
root.mainloop()
