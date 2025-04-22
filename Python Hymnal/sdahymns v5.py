import os
import sys
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QListWidget, QLineEdit,
    QPushButton, QLabel, QFileDialog, QMessageBox, QScrollArea, QFrame
)
from PySide6.QtGui import QIcon, QPixmap
from PySide6.QtCore import Qt, QSize
import pythoncom
import win32com.client
import shutil
import threading

class HymnalApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Seventh Day Adventist Hymnal")
        self.setFixedSize(640, 480)
        self.setWindowIcon(QIcon("_internal/Data/favicon.ico"))

        self.dir_path = os.path.dirname(os.path.realpath(__file__))
        self.init_ui()
        self.search_bar.setFocus()

    def init_ui(self):
        main_layout = QVBoxLayout(self)
        
        # Top menu bar layout
        menu_layout = QHBoxLayout()

        add_btn = QPushButton("Add Hymns")
        add_btn.clicked.connect(self.add_hymns)
        help_btn = QPushButton("Help")
        help_btn.clicked.connect(self.show_help)
        about_btn = QPushButton("About")
        about_btn.clicked.connect(self.show_about)
        clear_btn = QPushButton("Clear App")
        clear_btn.clicked.connect(self.quit_powerpoint)

        menu_layout.addWidget(add_btn)
        menu_layout.addWidget(help_btn)
        menu_layout.addWidget(about_btn)
        menu_layout.addWidget(clear_btn)
        menu_layout.addStretch()

        # Search bar
        self.search_bar = QLineEdit()
        self.search_bar.setPlaceholderText("Search hymns...")
        self.search_bar.textChanged.connect(self.search_files)
        menu_layout.addWidget(self.search_bar)

        main_layout.addLayout(menu_layout)

        # Result list
        self.result_list = QListWidget()
        self.result_list.itemDoubleClicked.connect(self.open_selected)
        main_layout.addWidget(self.result_list)

        self.search_files()

    def search_files(self):
        term = self.search_bar.text().lower()
        allowed_extensions = [".pps", ".ppsx", ".ppt", ".pptx"]
        self.result_list.clear()
        found = False

        for root, dirs, files in os.walk(self.dir_path):
            for file in files:
                if any(file.lower().endswith(ext) for ext in allowed_extensions):
                    if term in file.lower():
                        self.result_list.addItem(os.path.splitext(file)[0])
                        found = True

        if not found:
            self.result_list.addItem("No hymn found. Try another!")

        if term.strip() == "":
            self.result_list.scrollToTop()

    def open_selected(self):
        selected = self.result_list.currentItem()
        if selected:
            name = selected.text()
            for root, _, files in os.walk(self.dir_path):
                for file in files:
                    if name.lower() in file.lower():
                        full_path = os.path.join(root, file)
                        threading.Thread(target=self.launch_ppt, args=(full_path,), daemon=True).start()
                        return

    def launch_ppt(self, file_path):
        try:
            pythoncom.CoInitialize()
            ppt = win32com.client.Dispatch("PowerPoint.Application")
            ppt.Visible = True
            ppt.DisplayAlerts = False
            pres = ppt.Presentations.Open(file_path, WithWindow=True)
            pres.SlideShowSettings.Run()
            ppt.WindowState = 2
            ppt.DisplayAlerts = False
        except Exception as e:
            print("Error launching presentation:", e)
        finally:
            pythoncom.CoUninitialize()

    def quit_powerpoint(self):
        self.search_bar.clear()
        self.search_files()
        self.result_list.scrollToTop()
        self.toggle_focus()
        threading.Thread(target=self._quit_ppt, daemon=True).start()

    def _quit_ppt(self):
        try:
            pythoncom.CoInitialize()
            ppt = win32com.client.Dispatch("PowerPoint.Application")
            if ppt.Presentations.Count > 0:
                ppt.Quit()
            else:
                # Even if no presentation, attempt to quit safely
                ppt.Quit()
        except Exception as e:
            print("PowerPoint not running or failed to quit:", e)
        finally:
            pythoncom.CoUninitialize()
            
    def toggle_focus(self):
        if self.search_bar.hasFocus():
            self.result_list.setFocus()
            if self.result_list.count() > 0:
                self.result_list.setCurrentRow(0)
        else:
            self.result_list.clearSelection()
            self.search_bar.setFocus()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Shift:
            self.toggle_focus()
        elif event.key() == Qt.Key_Up:
            self.select_previous_result()
        elif event.key() == Qt.Key_Down:
            self.select_next_result()
        elif event.key() == Qt.Key_Return or event.key() == Qt.Key_Enter:
            self.open_selected()
        elif event.key() == Qt.Key_Escape:
            self.quit_powerpoint()

    def select_next_result(self):
        current = self.result_list.currentRow()
        if current < self.result_list.count() - 1:
            self.result_list.setCurrentRow(current + 1)

    def select_previous_result(self):
        current = self.result_list.currentRow()
        if current > 0:
            self.result_list.setCurrentRow(current - 1)

    def add_hymns(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "Select Hymns", filter="PowerPoint Files (*.pps *.ppsx)")
        if paths:
            target = os.path.join(self.dir_path, "_internal", "Data", "Added More Hymns")
            os.makedirs(target, exist_ok=True)
            for path in paths:
                try:
                    shutil.copy(path, os.path.join(target, os.path.basename(path)))
                except Exception as e:
                    QMessageBox.critical(self, "Error", f"Could not copy {path}: {e}")
            QMessageBox.information(self, "Done", f"{len(paths)} hymn(s) added!")
        self.search_bar.setFocus()
        
    def show_help(self):
        QMessageBox.information(self, "Help", "Keyboard Shortcuts:\n\nDouble-click a result to open.\nEnter search keywords in the top bar.\nAdd hymns via the 'Add Hymns' button.")
        self.search_bar.setFocus()
        
    def show_about(self):
        QMessageBox.information(self, "About", "Seventh Day Adventist Church Hymnal\n\nDeveloper: Jelmar A. Orapa\nEmail: orapajelmar@gmail.com")
        self.search_bar.setFocus()
        
if __name__ == "__main__":
    app = QApplication(sys.argv)
    hymnal_app = HymnalApp()
    hymnal_app.show()
    sys.exit(app.exec())
