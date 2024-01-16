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

    if not os.path.exists(base_path):
        messagebox.showerror("Error", "Invalid path")
        return

    matched_files = []
    for root, dirs, files in os.walk(base_path):
        for file in files:
            if all(keyword in file.lower() for keyword in keywords):
                file_path = "\\\\?\\" + os.path.join(root, file)
                try:
                    mod_time = datetime.fromtimestamp(os.path.getmtime(file_path)).strftime('%Y-%m-%d %H:%M:%S')
                    matched_files.append([file, mod_time, root])
                except FileNotFoundError:
                    print(f"File not found: {file_path}")

    if matched_files:
        df = pd.DataFrame(matched_files, columns=['File Name', 'Date Modified', 'Folder'])
        # Get the directory of the script
        script_dir = os.path.dirname(os.path.realpath(__file__))
        report_path = os.path.join(script_dir, 'report.xlsx')
        try:
            df.to_excel(report_path, index=False)
            print(f"Report generated at {report_path}")  # Debug print
            messagebox.showinfo("Completed", f"Your search is finished, report was generated successfully at {report_path}")
        except Exception as e:
            print(f"Error generating report: {e}")  # Print any error encountered
            messagebox.showerror("Error", "An error occurred while generating the report.")


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
