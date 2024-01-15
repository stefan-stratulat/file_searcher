import os
import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl

# Function to search files based on the path and keywords
def search_files():
    base_path = path_entry.get()
    keywords = [keyword.strip().lower() for keyword in keywords_entry.get().split(',')]

    # Check if path is valid
    if not os.path.exists(base_path):
        messagebox.showerror("Error", "Invalid path")
        return

    # Search for files
    matched_files = []
    for root, dirs, files in os.walk(base_path):
        for file in files:
            if all(keyword in file.lower() for keyword in keywords):
                file_path = os.path.join(root, file)
                mod_time = datetime.fromtimestamp(os.path.getmtime(file_path)).strftime('%Y-%m-%d %H:%M:%S')
                matched_files.append([file, mod_time, root])

    # Create DataFrame and generate Excel report
    if matched_files:
        df = pd.DataFrame(matched_files, columns=['File Name', 'Date Modified', 'Folder'])
        try:
            df.to_excel('report.xlsx', index=False)
            messagebox.showinfo("Completed", "Your search is finished, report was generated successfully")
        except PermissionError:
            messagebox.showerror("Permission Error", "The report file is open in another program. Please close it and try again.")
    else:
        messagebox.showinfo("No Files Found", "No files match your search criteria")


# Setting up the GUI
root = tk.Tk()
root.title("File Searcher")
root.geometry("600x150")  # Width x Height


# Configure grid for center alignment
root.columnconfigure(0, weight=1)
root.columnconfigure(1, weight=2)
root.columnconfigure(2, weight=1)

# Windows Path Entry
tk.Label(root, text="Windows Path:").grid(row=0, column=0, sticky=tk.W)
path_entry = tk.Entry(root)
path_entry.grid(row=0, column=1, sticky=tk.EW)

# Keywords Entry
tk.Label(root, text="Keywords:").grid(row=1, column=0, sticky=tk.W)
keywords_entry = tk.Entry(root)
keywords_entry.grid(row=1, column=1, sticky=tk.EW)

# Search Button
search_button = tk.Button(root, text="Search", command=search_files)
search_button.grid(row=2, column=0, columnspan=3)

root.mainloop()
