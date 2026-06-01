import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog
import pandas as pd
import shutil
import os
import subprocess
import sys

import re # Import the regular expression module
import colorsys
import numpy as np
import openpyxl
import xlsxwriter

from FinancialsSummary import FinancialsSummary

def upload_file():
    file_path = filedialog.askopenfilename(
        title="Select a file",
        filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
    )
    if file_path:
        try:
            # Get the base name of the file
            file_name = os.path.basename(file_path)
            # Define the destination path in the current working directory
            destination_path = os.getcwd() + "\\"

            # Copy the file to the working directory
            shutil.copy(file_path, destination_path+file_name)
            messagebox.showinfo(f"File Uploaded",f"Creating Financials Summary file for {file_name}\r\n{destination_path}{file_name}")

            # Process the file and create the destination summary file
            dest_path = FinancialsSummary(destination_path, file_name)

            # Create a small window to present the destination as a downloadable/openable link
            link_win = tk.Toplevel(root)
            link_win.title("Summary Created")
            link_win.geometry("600x120")

            tk.Label(link_win, text=f"Summary created at:", font=("Arial", 10, "bold")).pack(pady=(10, 0))
            path_label = tk.Label(link_win, text=dest_path, fg="blue", cursor="hand2")
            path_label.pack(pady=(5, 10))
            path_label.bind("<Button-1>", lambda e: open_file())

            def open_file():
                try:
                    if os.name == 'nt':
                        os.startfile(dest_path)
                    elif sys.platform == 'darwin':
                        subprocess.call(['open', dest_path])
                    else:
                        subprocess.call(['xdg-open', dest_path])
                except Exception as e:
                    messagebox.showerror("Error", f"Could not open file: {e}")

            def open_folder():
                folder = os.path.dirname(dest_path)
                try:
                    if os.name == 'nt':
                        os.startfile(folder)
                    elif sys.platform == 'darwin':
                        subprocess.call(['open', folder])
                    else:
                        subprocess.call(['xdg-open', folder])
                except Exception as e:
                    messagebox.showerror("Error", f"Could not open folder: {e}")



            open_btn = tk.Button(link_win, text="Open File", command=open_file, bg="#0078d4", fg="white")
            open_btn.pack(side="left", padx=(150, 10))
            open_folder_btn = tk.Button(link_win, text="Open Folder", command=open_folder)
            open_folder_btn.pack(side="left")
            # Add a Close button to dismiss this link window
            close_message_btn = tk.Button(link_win, text="Close", command=link_win.destroy)
            close_message_btn.pack(side="left", padx=(10, 0))
            # Allow Esc key to close the window as well
            link_win.bind("<Escape>", lambda e: link_win.destroy())
            close_message_btn = tk.Button(link_win, text="Close")
            close_message_btn.pack(side="left")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")
    else:
        messagebox.showinfo("No File", "No file was selected.")

# Create the main application window
root = tk.Tk()
root.title("Excel Summary Creation")
root.geometry("600x200")
root.configure(bg="#f0f0f0")

# Add a text label
label = tk.Label(
    root,
    text="Welcome to Sundeep's prototype!\r\nFinancial Summary creator for Bloom, Aster, ACIL",
    font=("Arial", 14, "bold"),
    bg="#f0f0f0",
    fg="#333333"
)
label.pack(pady=30)

# Add an interactive button
button = tk.Button(
    root,
    text="Upload File",
    command=upload_file,
    font=("Arial", 11),
    bg="#0078d4",
    fg="white",
    padx=10,
    pady=5
)
button.pack()

# Start the application loop
root.mainloop()