import re
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os
from bs4 import BeautifulSoup

def extract_data_from_txt(input_file, output_folder):
    """
    Reads a text file, extracts date, entrance time, and exit time (both English and Turkish),
    then saves the data to an Excel file.
    """
    try:
        # Read the file content
        with open(input_file, "r", encoding="utf-8") as file:
            text = file.read()

        # Parse text with BeautifulSoup
        soup = BeautifulSoup(text, "html.parser")

        # Extract relevant entries
        entries = []

        # Find all date elements (assuming they have a specific tag or class)
        dates = soup.find_all("div", class_="sapExtentUilibHistoryItemDate")

        # Iterate through each date block and find corresponding entry and exit times
        for date_element in dates:
            date = date_element.get_text(strip=True)
            
            # Find the parent element containing entrance and exit times
            parent = date_element.find_parent("div", class_="sapExtentUilibHistoryItem")
            if parent:
                entrance_time = parent.find("span", title=re.compile("Giriş Saati|Entrance Time"))
                exit_time = parent.find("span", title=re.compile("Çıkış Saati|Exit Time"))
                
                if entrance_time and exit_time:
                    entrance_value = entrance_time.find_next_sibling("span").get_text(strip=True) if entrance_time.find_next_sibling("span") else ""
                    exit_value = exit_time.find_next_sibling("span").get_text(strip=True) if exit_time.find_next_sibling("span") else ""
                    
                    entries.append({"Date": date, "Entrance Time": entrance_value, "Exit Time": exit_value})

        # If no data found, display a warning
        if not entries:
            messagebox.showwarning("No Data Found", "No matching data found in the file.")
            return

        # Convert matches to a DataFrame
        df = pd.DataFrame(entries)

        # Set output file path
        output_excel = os.path.join(output_folder, "attendance_records.xlsx")

        # Save to Excel
        df.to_excel(output_excel, index=False)

        # Show success message
        messagebox.showinfo("Success", f"Data extracted and saved to:\n{output_excel}")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n{e}")

def browse_txt_file():
    """ Opens file dialog to select a .txt file """
    file_path = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt")])
    txt_entry.delete(0, tk.END)
    txt_entry.insert(0, file_path)

def browse_output_folder():
    """ Opens folder dialog to select an output directory """
    folder_path = filedialog.askdirectory()
    output_entry.delete(0, tk.END)
    output_entry.insert(0, folder_path)

def process_file():
    """ Gets file path and output folder, then processes the file """
    txt_file = txt_entry.get()
    output_folder = output_entry.get()

    if not txt_file or not output_folder:
        messagebox.showwarning("Missing Input", "Please select both a text file and an output folder.")
        return

    extract_data_from_txt(txt_file, output_folder)

# Create GUI Window
root = tk.Tk()
root.title("TXT to Excel Converter")
root.geometry("500x300")

# File Selection
tk.Label(root, text="Select TXT File:").pack(pady=5)
txt_entry = tk.Entry(root, width=50)
txt_entry.pack(pady=2)
tk.Button(root, text="Browse", command=browse_txt_file).pack(pady=2)

# Output Folder Selection
tk.Label(root, text="Select Output Folder:").pack(pady=10)
output_entry = tk.Entry(root, width=50)
output_entry.pack(pady=2)
tk.Button(root, text="Browse", command=browse_output_folder).pack(pady=2)

# Process Button
tk.Button(root, text="Convert to Excel", command=process_file, bg="green", fg="white").pack(pady=20)

# Run the GUI
root.mainloop()
