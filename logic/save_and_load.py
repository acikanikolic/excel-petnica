import tkinter as tk
from tkinter import messagebox, filedialog
import csv

def save_file(cells):
    file_path = tk.filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", ".csv")])
    if file_path:
        try:
            with open(file_path, 'w', newline='') as file:
                writer = csv.writer(file)
                for row in range(1, 11):
                    row_data = []
                    for col in range(1, 11):
                        value = cells[(row, col)].get()
                        row_data.append(value)
                    writer.writerow(row_data)
            messagebox.showinfo("Success", "File saved successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file: {str(e)}")


def load_file(cells):
    file_path = tk.filedialog.askopenfilename(filetypes=[("CSV files", ".csv")])
    if file_path:
        try:
            with open(file_path, 'r') as file:
                reader = csv.reader(file)
                for row_idx, row_data in enumerate(reader):
                    for col_idx, value in enumerate(row_data):
                        cells[(row_idx + 1, col_idx + 1)].delete(0, tk.END)
                        cells[(row_idx + 1, col_idx + 1)].insert(0, value)
            messagebox.showinfo("Success", "File loaded successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {str(e)}")