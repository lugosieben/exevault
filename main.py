import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
import configparser
import os
import subprocess
import win32com.client
import shutil
import ctypes
import sys
import pythoncom

class ExeVaultGui:
    def __init__(self, root):
        self.root = root
        self.root.title("ExeVault")
        
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "exevault.ico")
        if os.path.exists(icon_path):
            myappid = 'exevault'
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
            self.root.iconbitmap(icon_path)

            
        self.storage_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "vault")
        self.ensure_storage_dir()

        self.main_frame = ttk.Frame(root)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.canvas = tk.Canvas(self.main_frame, highlightthickness=1, highlightbackground="black")
        self.scrollbar = ttk.Scrollbar(self.main_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw", tags="self.scrollable_frame")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.bind_all("<MouseWheel>", self.on_mousewheel)

        self.canvas.bind("<Configure>", self.on_canvas_configure)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="left", fill="y")

        self.entries = {}

        self.control_frame = ttk.Frame(self.main_frame)
        self.control_frame.pack(side="right", fill=tk.Y, padx=10, pady=10)

        self.entry_text = tk.StringVar()
        self.entry_field = ttk.Entry(self.control_frame, textvariable=self.entry_text, width=40)
        self.entry_field.insert(0, "Enter shortcut name")
        self.entry_field.bind('<FocusIn>', lambda e: self.entry_field.delete(0, 'end') if self.entry_field.get() == "Enter shortcut name" else None)
        self.entry_field.bind('<FocusOut>', lambda e: self.entry_field.insert(0, "Enter shortcut name") if not self.entry_field.get() else None)
        self.entry_field.pack(pady=5)

        self.exepath_text = tk.StringVar()
        self.exepath_field = ttk.Entry(self.control_frame, textvariable=self.exepath_text, width=40)
        self.exepath_field.insert(0, "Select executable path")
        self.exepath_field.bind('<FocusIn>', lambda e: self.exepath_field.delete(0, 'end') if self.exepath_field.get() == "Select executable path" else None)
        self.exepath_field.bind('<FocusOut>', lambda e: self.exepath_field.insert(0, "Select executable path") if not self.exepath_field.get() else None)
        self.exepath_field.pack(pady=5)

        self.browse_button = ttk.Button(self.control_frame, text="Browse", command=self.browse_exe)
        self.browse_button.pack(pady=5)

        self.start_menu_var = tk.BooleanVar()
        self.start_menu_checkbox = ttk.Checkbutton(
            self.control_frame, 
            text="Add to Start Menu", 
            variable=self.start_menu_var
        )
        self.start_menu_checkbox.pack(pady=5)

        self.add_button = ttk.Button(self.control_frame, text="Add Entry", command=self.add_entry_from_ui)
        self.add_button.pack(pady=5)

        self.status_label = ttk.Label(root, text="Ready", relief="sunken", anchor="w")
        self.status_label.pack(side="bottom", fill="x", padx=5, pady=2)

        self.config = configparser.ConfigParser()
        self.config_file = "entries.ini"
        self.clean_orphaned_files()
        self.load_entries()

        if not self.check_storage_writable() and not self.is_elevated():
            self.elevate_and_restart()
            root.destroy()
            return

    def ensure_storage_dir(self):
        if not os.path.exists(self.storage_dir):
            os.makedirs(self.storage_dir)

    def clone_executable(self, source_path, entry_name):
        self.update_status(f"Cloning executable: {entry_name}...")
        if not os.path.exists(source_path):
            self.update_status("Error: Source file not found")
            return None
        
        safe_name = "".join(c for c in entry_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
        dest_path = os.path.join(self.storage_dir, f"{safe_name}.exe")
        
        try:
            shutil.copy2(source_path, dest_path)
            self.update_status("Executable cloned successfully")
            return dest_path
        except Exception as e:
            self.update_status(f"Error: Failed to clone executable - {str(e)}")
            messagebox.showerror("Error", f"Failed to clone executable: {e}")
            return None

    def on_canvas_configure(self, event):
        canvas_width = event.width
        self.canvas.itemconfig("self.scrollable_frame", width=canvas_width - 5)

    def on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta/120)), "units")

    def get_start_menu_folder(self):
        start_menu_path = os.path.join(
            os.getenv('APPDATA'),
            r'Microsoft\Windows\Start Menu\Programs\exevault'
        )
        if not os.path.exists(start_menu_path):
            try:
                os.makedirs(start_menu_path)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to create Start Menu folder: {e}")
                return None
        return start_menu_path

    def create_start_menu_shortcut(self, text, exepath):
        start_menu_folder = self.get_start_menu_folder()
        if not start_menu_folder:
            return None
            
        shortcut_path = os.path.join(start_menu_folder, f"{text}.lnk")
        
        if os.path.exists(shortcut_path):
            try:
                os.remove(shortcut_path)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to remove existing shortcut: {e}")
                return None
            
        try:
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut(shortcut_path)
            shortcut.Targetpath = exepath
            shortcut.WorkingDirectory = os.path.dirname(exepath)
            shortcut.save()
            return shortcut_path
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create Start Menu shortcut: {e}")
            return None

    def remove_start_menu_shortcut(self, text):
        start_menu_folder = self.get_start_menu_folder()
        shortcut_path = os.path.join(start_menu_folder, f"{text}.lnk")
        if os.path.exists(shortcut_path):
            os.remove(shortcut_path)

    def rename_entry(self, old_text, label):
        entry = ttk.Entry(label.master)
        entry.insert(0, old_text)
        entry.place(relx=0, rely=0, relwidth=0.7, relheight=1)
        
        def finish_rename(event=None):
            new_text = entry.get().strip()
            if new_text and new_text != old_text:
                if new_text not in self.entries:
                    frame, exepath, start_menu_var = self.entries[old_text]
                    if start_menu_var.get():
                        self.remove_start_menu_shortcut(old_text)
                        self.create_start_menu_shortcut(new_text, exepath)
                    self.entries[new_text] = self.entries.pop(old_text)
                    label.configure(text=new_text)
                    self.save_entries()
                else:
                    messagebox.showerror("Error", f"Entry '{new_text}' already exists.")
                    label.configure(text=old_text)
            else:
                label.configure(text=old_text)
            entry.destroy()
            label.place_forget()
            label.pack(side="left", fill="both", expand=True, padx=5)

        entry.bind('<Return>', finish_rename)
        entry.bind('<FocusOut>', finish_rename)
        entry.focus_set()

    def add_entry(self, text, exepath, start_menu=False, from_config=False):
        self.update_status(f"Adding entry: {text}...")
        if text and text not in self.entries:
            if from_config and exepath.startswith(self.storage_dir):
                cloned_path = exepath
            else:
                cloned_path = self.clone_executable(exepath, text)
                if not cloned_path:
                    self.update_status("Error: Failed to add entry")
                    messagebox.showerror("Error", "Failed to clone executable.")
                    return

            frame = ttk.Frame(self.scrollable_frame, borderwidth=2, relief="solid")
            frame.pack(pady=5, padx=5, fill="x", expand=True)

            label = ttk.Label(frame, text=text, anchor="w", justify="left", cursor="hand2")
            label.pack(side="left", fill="both", expand=True, padx=5)
            label.bind('<Button-1>', lambda e, t=text, l=label: self.rename_entry(t, l))

            run_button = ttk.Button(frame, text="Run", command=lambda: self.run_exe(cloned_path))
            run_button.pack(side="right", padx=5)

            start_menu_var = tk.BooleanVar(value=start_menu)
            start_menu_check = ttk.Checkbutton(
                frame, 
                text="Start Menu", 
                variable=start_menu_var,
                command=lambda: self.toggle_start_menu(label.cget("text"), cloned_path, start_menu_var.get())
            )
            start_menu_check.pack(side="right", padx=5)

            remove_button = ttk.Button(frame, text="Remove", command=lambda: self.remove_entry(label.cget("text")))
            remove_button.pack(side="right", padx=5)

            self.entries[text] = (frame, cloned_path, start_menu_var)
            
            if start_menu:
                self.create_start_menu_shortcut(text, cloned_path)
            
            if not from_config:
                self.update_status("Entry added successfully")
                self.config[text] = {
                    'exepath': cloned_path,
                    'start_menu': str(start_menu)
                }
                self.save_entries()

    def toggle_start_menu(self, text, exepath, enable):
        if enable:
            self.create_start_menu_shortcut(text, exepath)
        else:
            self.remove_start_menu_shortcut(text)
        self.save_entries()

    def remove_entry(self, text):
        self.update_status(f"Removing entry: {text}...")
        if text in self.entries:
            frame, exepath, start_menu_var = self.entries[text]
            if start_menu_var.get():
                self.remove_start_menu_shortcut(text)
            frame.destroy()
            if os.path.isfile(exepath) and exepath.startswith(self.storage_dir):
                try:
                    os.remove(exepath)
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to remove cloned executable: {e}")
            self.entries.pop(text)
            if self.config.has_section(text):
                self.config.remove_section(text)
            self.update_status("Entry removed successfully")
            self.save_entries()
        else:
            self.update_status(f"Error: Entry '{text}' not found")
            messagebox.showerror("Error", f"Entry '{text}' does not exist.")

    def add_entry_from_ui(self):
        if not self.check_storage_writable() and not self.is_elevated():
            response = messagebox.askyesno(
                "Administrator Rights Required",
                "Adding entries requires administrator privileges. Do you want to restart with elevated privileges?"
            )
            if response:
                self.elevate_and_restart()
            return

        text = self.entry_text.get()
        exepath = self.exepath_text.get()
        start_menu = self.start_menu_var.get()
        
        if text not in self.entries:
            if os.path.isfile(exepath):
                self.add_entry(text, exepath, start_menu)
                self.entry_text.set("")
                self.exepath_text.set("")
                self.entry_field.insert(0, "Enter shortcut name")
                self.exepath_field.insert(0, "Select executable path")
                self.start_menu_var.set(False)
            else:
                messagebox.showerror("Error", "Invalid executable path.")
        else:
            messagebox.showerror("Error", f"Entry '{text}' already exists.")

    def save_entries(self):
        for entry, (_, exepath, start_menu_var) in self.entries.items():
            if not self.config.has_section(entry):
                self.config.add_section(entry)
            self.config[entry]['exepath'] = exepath
            self.config[entry]['start_menu'] = str(start_menu_var.get())
        
        for section in self.config.sections():
            if section not in self.entries:
                self.config.remove_section(section)
        
        with open(self.config_file, 'w') as configfile:
            self.config.write(configfile)

    def load_entries(self):
        self.update_status("Loading entries...")
        if not os.path.exists(self.config_file):
            self.update_status("No entries file found")
            return
            
        self.config.read(self.config_file)
        
        invalid_entries = []
        
        for entry in self.config.sections():
            try:
                if self.validate_entry(entry):
                    exepath = self.config[entry]['exepath']
                    start_menu = self.config[entry].getboolean('start_menu', False)
                    self.add_entry(entry, exepath, start_menu, from_config=True)
                else:
                    invalid_entries.append(entry)
            except Exception as e:
                print(f"Error loading entry {entry}: {e}")
                invalid_entries.append(entry)
        
        for entry in invalid_entries:
            self.config.remove_section(entry)
        
        if invalid_entries:
            self.update_status(f"Loaded with {len(invalid_entries)} invalid entries removed")
        else:
            self.update_status("Entries loaded successfully")

    def run_exe(self, exepath):
        if os.path.isfile(exepath):
            try:
                self.update_status("Launching executable...")
                try:
                    process = subprocess.Popen(exepath)
                    self.update_status("Executable launched successfully")
                except WindowsError as e:
                    # 740 = ERROR_ELEVATION_REQUIRED
                    if hasattr(e, 'winerror') and e.winerror == 740:
                        response = messagebox.askyesno(
                            "Administrator Rights Required",
                            "This executable requires administrator privileges. Do you want to run as administrator?"
                        )
                        if response:
                            result = ctypes.windll.shell32.ShellExecuteW(
                                None,
                                "runas",
                                str(exepath),
                                None,
                                os.path.dirname(str(exepath)),
                                1  # SW_SHOWNORMAL
                            )
                            # ShellExecute returns value > 32 if successful
                            if result > 32:
                                self.update_status("Executable launched with admin privileges")
                            else:
                                self.update_status("Failed to launch with admin privileges")
                                messagebox.showerror("Error", "Failed to launch with administrator privileges")
                    else:
                        raise
            except Exception as e:
                self.update_status(f"Error: Failed to run executable - {str(e)}")
                messagebox.showerror("Error", f"Failed to run executable: {e}")
        else:
            self.update_status("Error: Executable file not found")
            messagebox.showerror("Error", "Executable file not found.")

    def browse_exe(self):
        exepath = filedialog.askopenfilename(filetypes=[("Executable files", "*.exe")])
        if exepath:
            self.exepath_text.set(exepath)

    def clean_orphaned_files(self):
        self.update_status("Cleaning orphaned files...")
        self.config.read(self.config_file)
        
        if os.path.exists(self.storage_dir):
            for file in os.listdir(self.storage_dir):
                file_path = os.path.join(self.storage_dir, file)
                is_referenced = any(
                    section in self.config 
                    and self.config[section].get('exepath') == file_path 
                    for section in self.config.sections()
                )
                if not is_referenced:
                    try:
                        os.remove(file_path)
                    except Exception:
                        pass

        start_menu_folder = self.get_start_menu_folder()
        if start_menu_folder and os.path.exists(start_menu_folder):
            for file in os.listdir(start_menu_folder):
                if file.endswith('.lnk'):
                    entry_name = os.path.splitext(file)[0]
                    if entry_name not in self.config.sections():
                        try:
                            os.remove(os.path.join(start_menu_folder, file))
                        except Exception:
                            pass
        self.update_status("Cleanup completed")

    def validate_entry(self, entry):
        if not entry in self.config:
            return False
        
        try:
            exepath = self.config[entry].get('exepath', '')
            if not os.path.isfile(exepath):
                return False
            if not exepath.startswith(self.storage_dir):
                return False
            return True
        except Exception:
            return False

    def update_status(self, message):
        self.status_label.config(text=message)
        self.root.update_idletasks()

    def check_storage_writable(self):
        if not os.path.exists(self.storage_dir):
            try:
                os.makedirs(self.storage_dir)
                return True
            except PermissionError:
                return False
        
        test_file = os.path.join(self.storage_dir, "write_test")
        try:
            with open(test_file, 'w') as f:
                f.write('test')
            os.remove(test_file)
            return True
        except (PermissionError, OSError):
            return False

    def is_elevated(self):
        try:
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False

    def elevate_and_restart(self):
        if not self.is_elevated():
            self.update_status("Elevating privileges...")
            pythoncom.CoInitialize()
            ctypes.windll.shell32.ShellExecuteW(
                None,
                "runas",
                os.path.join(sys.exec_prefix, 'pythonw.exe'),
                " ".join(sys.argv),
                None,
                1  # SW_SHOWNORMAL
            )
            self.update_status("Restarting with elevated privileges...")

if __name__ == "__main__":
    root = tk.Tk()
    gui = ExeVaultGui(root)
    root.mainloop()