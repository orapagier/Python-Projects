import os
import pythoncom
import threading
import win32com.client
import tkinter as tk
from tkinter import Scrollbar, Listbox, Entry, Menu, messagebox, filedialog, StringVar
from PIL import Image, ImageTk
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
            result_listbox.insert(tk.END, os.path.splitext(result)[0])

def quit_powerpoint():
    clear_search_entry()
    search_entry.focus_set()
    search_files()
    def quit_thread():
        try:
            pythoncom.CoInitialize()
            ppt = win32com.client.Dispatch('PowerPoint.Application')
            ppt.Quit()
        except Exception as e:
            print("Error quitting PowerPoint:", e)
        finally:
            pythoncom.CoUninitialize()
    threading.Thread(target=quit_thread, daemon=True).start()

def clear_search_entry():
    search_entry.delete(0, tk.END)

def open_selected(event):
    selected = result_listbox.curselection()
    if selected:
        item = result_listbox.get(selected[0])
        for root_dir, _, files in os.walk(dir_path):
            for file in files:
                if item.lower() in file.lower():
                    full_path = os.path.join(root_dir, file)
                    try:
                        pythoncom.CoInitialize()
                        ppt = win32com.client.Dispatch('PowerPoint.Application')
                        ppt.Visible = True
                        
                        # Add these lines to bypass security warnings
                        ppt.DisplayAlerts = False
                        
                        pres = ppt.Presentations.Open(full_path, WithWindow=True)
                        pres.SlideShowSettings.AdvanceMode = 1
                        pres.SlideShowSettings.ShowType = 1
                        pres.SlideShowSettings.Run()
                        ppt.WindowState = 2
                        pres.SlideShowSettings.ShowPresenterView = True
                        
                        # Reset alerts after opening
                        ppt.DisplayAlerts = False
                        
                        pythoncom.CoUninitialize()
                        return
                    except Exception as e:
                        print("Error:", e)
                        pythoncom.CoUninitialize()

def toggle_focus(event=None):
    if search_entry.focus_get() == search_entry:
        result_listbox.select_set(0)
        result_listbox.focus_set()
    else:
        result_listbox.select_clear(0, tk.END)
        search_entry.focus_set()
        search_files()

def select_next_result(event):
    current = result_listbox.curselection()
    if current:
        idx = (current[0] + 1) % result_listbox.size()
        result_listbox.select_clear(0, tk.END)
        result_listbox.select_set(idx)
        result_listbox.event_generate("<<ListboxSelect>>")

def select_previous_result(event):
    current = result_listbox.curselection()
    if current:
        idx = max(current[0] - 1, 0)
        result_listbox.select_clear(0, tk.END)
        result_listbox.select_set(idx)
        result_listbox.event_generate("<<ListboxSelect>>")

def add_hymns():
    paths = filedialog.askopenfilenames(title="Select Hymn Files", filetypes=[("PowerPoint Files", "*.pps *.ppsx")])
    if paths:
        target = os.path.join(dir_path, "_internal", "Data", "Added More Hymns") 
        os.makedirs(target, exist_ok=True)
        for path in paths:
            try:
                shutil.copy(path, os.path.join(target, os.path.basename(path)))
            except Exception as e:
                messagebox.showerror("Error", f"Could not copy {path}: {e}")
        messagebox.showinfo("Done", f"{len(paths)} hymn(s) added!")

def apphelp():
    messagebox.showinfo("Help", "Keyboard Shortcuts:\n\nShift → Switch focus\nArrow ↑↓ → Navigate\nEnter → Open hymn\nEsc → Close presentation\n\nAdd Hymns:\nGo to File > Add Hymns and select .pps or .ppsx files to add.")

def about():
    messagebox.showinfo("About", "Seventh Day Adventist Church Hymnal\n\nDeveloper: Jelmar A. Orapa\nEmail: orapajelmar@gmail.com")

def delete_temp_folder():
    temp = os.path.join(os.getcwd(), "Temp")
    if os.path.exists(temp):
        shutil.rmtree(temp)
        messagebox.showinfo("Deleted", "Temporary folder deleted.")
    else:
        messagebox.showinfo("Info", "Temporary folder does not exist.")

def update_background():
    global resized_bg_image
    try:
        bg_image = Image.open(r"_internal/Data/bg.png")
        resized_bg_image = bg_image.resize((root.winfo_width(), root.winfo_height()), Image.LANCZOS)
        bg_image_tk = ImageTk.PhotoImage(resized_bg_image)
        background_label.config(image=bg_image_tk)
        background_label.image = bg_image_tk
    except:
        pass

# Dictionary to track active menus
active_menus = {}

# Function to handle menu popup on hover
def show_menu(event, menu_button, menu):
    # Hide all other menus first
    for m in active_menus.values():
        m.unpost()
    
    # Get button coordinates
    x = menu_button.winfo_rootx()
    y = menu_button.winfo_rooty() + menu_button.winfo_height()
    menu.post(x, y)
    
    # Store the active menu
    active_menus[menu_button] = menu

# Function to show menu on click as well
def show_menu_click(event, menu_button, menu):
    show_menu(event, menu_button, menu)

# Function to determine if mouse is over a widget
def is_mouse_over_widget(widget):
    x, y = root.winfo_pointerxy()
    widget_x = widget.winfo_rootx()
    widget_y = widget.winfo_rooty()
    widget_width = widget.winfo_width()
    widget_height = widget.winfo_height()
    
    return (widget_x <= x <= widget_x + widget_width and 
            widget_y <= y <= widget_y + widget_height)

# Global tracking for menu handling
current_menu = None
menu_hide_scheduled = False

def schedule_menu_hide(menu_button, menu):
    global menu_hide_scheduled
    
    # If there's already a hide scheduled, don't schedule another
    if menu_hide_scheduled:
        return
    
    menu_hide_scheduled = True
    
    # Schedule the hide check
    root.after(100, check_menu_hide, menu_button, menu)

def check_menu_hide(menu_button, menu):
    global menu_hide_scheduled
    menu_hide_scheduled = False
    
    # If mouse is not over menu button or menu, hide the menu
    menu_widget = menu
    
    if not is_mouse_over_widget(menu_button):
        # Check if mouse is over any menu
        try:
            menu_coords = menu_widget.winfo_geometry().split('+')
            menu_width = int(menu_coords[0].split('x')[0])
            menu_height = int(menu_coords[0].split('x')[1])
            menu_x = int(menu_coords[1])
            menu_y = int(menu_coords[2])
            
            x, y = root.winfo_pointerxy()
            
            if not (menu_x <= x <= menu_x + menu_width and 
                    menu_y <= y <= menu_y + menu_height):
                menu.unpost()
        except:
            # Menu might not be visible, just unpost
            menu.unpost()

# === GUI Setup ===
root = tk.Tk()
root.title("Seventh Day Adventist Church Hymnal")

icon_image = Image.open("_internal/Data/favicon.ico")
icon_photo = ImageTk.PhotoImage(icon_image)
root.iconphoto(True, icon_photo)

window_width, window_height = 620, 422
screen_width, screen_height = root.winfo_screenwidth(), root.winfo_screenheight()
x, y = (screen_width - window_width) // 2, (screen_height - window_height) // 2
root.geometry(f"{window_width}x{window_height}+{x}+{y}")
root.resizable(False, False)

# === Background ===
background_label = tk.Label(root)
background_label.place(relwidth=1, relheight=1)
update_background()

# === Menu Bar ===
menu_bar = tk.Frame(root, bg="white", height=30, relief="flat", bd=0)
menu_bar.pack(fill=tk.X, side=tk.TOP)

# Button styling configuration
button_font = ("Times New Roman", 10)
button_style = {"bg": "white", "relief": "flat", "cursor": "hand2", "padx": 5, "pady": 3}

# Add individual buttons for all menu items - flattened structure
add_hymns_button = tk.Button(menu_bar, text="Add hymns", font=button_font, command=add_hymns, **button_style)
add_hymns_button.pack(side=tk.LEFT)

help_button = tk.Button(menu_bar, text="Help", font=button_font, command=apphelp, **button_style)
help_button.pack(side=tk.LEFT)

about_button = tk.Button(menu_bar, text="About", font=button_font, command=about, **button_style)
about_button.pack(side=tk.LEFT)

# Clear Button
clear_button = tk.Button(menu_bar, text="Clear app", font=button_font, command=quit_powerpoint, **button_style)
clear_button.pack(side=tk.LEFT)

# Search Frame inside menu bar
search_frame = tk.Frame(menu_bar, bg="white")
search_frame.pack(side=tk.RIGHT, padx=10)

# Create a subtle background frame to make entry recognizable without border
entry_background = tk.Frame(search_frame, bg="#f0f0f0", bd=0, highlightthickness=0)
entry_background.pack(side=tk.LEFT, pady=2)

search_var = StringVar()
# Complete removal of borders on search entry
search_entry = Entry(entry_background, textvariable=search_var, width=20, 
                     relief="raised", bd=0,  # Flat with no border
                     highlightbackground="white",
                     highlightthickness=0,  # No highlight border
                     font=("Times New Roman", 11))
search_entry.pack(side=tk.LEFT, pady=2, padx=2)

# Load and attach search icon inside Entry field
search_icon_path = os.path.join("_internal", "Data", "search_icon.png")
search_icon_img = Image.open(search_icon_path).resize((16, 16), Image.LANCZOS)
search_icon = ImageTk.PhotoImage(search_icon_img)

search_button = tk.Label(search_frame, image=search_icon, bg="white", cursor="hand2")
search_button.image = search_icon
search_button.pack(side=tk.LEFT, padx=(2, 0))
search_button.bind("<Button-1>", lambda e: search_files())

search_var.trace_add("write", lambda *args: search_files())

# === Main Content ===
main_frame = tk.Frame(root, bg="white")
main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 35))

result_listbox = Listbox(main_frame, selectmode=tk.SINGLE, borderwidth=0, highlightthickness=0, font=("Times New Roman", 12))
scrollbar = Scrollbar(main_frame, orient=tk.VERTICAL, command=result_listbox.yview)
result_listbox.config(yscrollcommand=scrollbar.set)

result_listbox.grid(row=0, column=0, sticky="nsew")
scrollbar.grid(row=0, column=1, sticky="ns")

main_frame.grid_rowconfigure(0, weight=1)
main_frame.grid_columnconfigure(0, weight=1)

# === Event Bindings ===
result_listbox.bind("<Double-Button-1>", open_selected)
result_listbox.bind("<Return>", open_selected)

root.bind("<Configure>", lambda e: update_background())
root.bind("<Shift_R>", toggle_focus)
root.bind("<Up>", select_previous_result)
root.bind("<Down>", select_next_result)

# Function to check mouse position and handle menus
def check_mouse_position():
    # This ensures menus open/close correctly during mouse movement
    root.after(200, check_mouse_position)

# Start the mouse position checking
check_mouse_position()

search_entry.focus_set()
search_files()
root.mainloop()
