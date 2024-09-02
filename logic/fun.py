import tkinter as tk
from tkinter import filedialog, messagebox
import csv
from logic.undo_redo import UndoRedoManager


class ExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel")
        self.root.geometry("800x600")
        self.manager = UndoRedoManager()
        self.create_buttons()
        self.create_grid()

        self.selected_row = None
        self.selected_col = None
        self.selected_cell = None


    # selekcija
    def start_selection(self, event):
        cell = self.get_cell_coordinates(event.widget)
        if cell:
            self.selection_start = cell
            self.update_selection()

    def extend_selection(self, event):
        """Extend the selection while dragging the mouse."""
        cell = self.get_cell_coordinates(event.widget)
        if cell:
            self.selection_end = cell
            self.update_selection()

    def end_selection(self, event):
        """Finalize the selection when mouse button is released."""
        cell = self.get_cell_coordinates(event.widget)
        if cell:
            self.selection_end = cell
            self.update_selection()

    def update_selection(self):
        """Highlight the selected cells."""
        self.clear_selection()  # Clear previous selection

        if self.selection_start and self.selection_end:
            start_row, start_col = self.selection_start
            end_row, end_col = self.selection_end

            # Determine the selection range
            for row in range(min(start_row, end_row), max(start_row, end_row) + 1):
                for col in range(min(start_col, end_col), max(start_col, end_col) + 1):
                    if (row, col) in self.cells:
                        self.cells[(row, col)].config(bg="#ADD8E6")  # Highlight color

    def clear_selection(self):
        for (row, col), entry in self.cells.items():
            entry.config(bg="white")

    def select_row(self, event):
        row = int(event.widget.cget("text"))
        self.clear_selection()  # Clear previous selection
        self.highlight_row(row)  # Highlight the clicked row
        self.selected_row = row

    def select_column(self, event):
        col = ord(event.widget.cget("text")) - 64  # Convert 'A'-'J' to 1-10
        self.clear_selection()  # Clear previous selection
        self.highlight_column(col)  # Highlight the clicked column
        self.selected_col = col

    def highlight_row(self, row):
        for col in range(1, 11):
            cell = (row, col)
            if cell in self.cells:
                entry = self.cells[cell]
                entry.config(borderwidth=1, relief="solid")

    def highlight_column(self, col):
        for row in range(1, 11):
            cell = (row, col)
            if cell in self.cells:
                entry = self.cells[cell]
                entry.config(borderwidth=1, relief="solid")

    def select_cell(self, event):
        clicked_cell = self.get_cell_coordinates(event.widget)
        if clicked_cell:
            self.clear_selection()  # Clear previous selection
            self.highlight_border(clicked_cell)  # Highlight the clicked cell
            self.selected_cell = clicked_cell

    def highlight_border(self, cell):
        if cell in self.cells:
            entry = self.cells[cell]
            entry.config(borderwidth=1, relief="solid")  # Add border

    def clear_selection(self):
        for (row, col), entry in self.cells.items():
            entry.config(borderwidth=1, relief="flat")

    def create_buttons(self):
        button_frame = tk.Frame(self.root, bg="#f0f0f0", bd=1, relief="solid")
        button_frame.grid(row=0, column=0, columnspan=11, pady=10, padx=10, sticky="ew")

        save_button = tk.Button(button_frame, text="Save", command=self.save_file, bg="#4CAF50", fg="white", padx=20)
        save_button.pack(side=tk.LEFT, padx=10, pady=5)

        load_button = tk.Button(button_frame, text="Load", command=self.load_file, bg="#008CBA", fg="white", padx=20)
        load_button.pack(side=tk.LEFT, padx=10, pady=5)

        undo_button = tk.Button(button_frame, text="Undo", command=self.undo_action, bg="#FFC107", fg="black", padx=20)
        undo_button.pack(side=tk.LEFT, padx=10, pady=5)

        redo_button = tk.Button(button_frame, text="Redo", command=self.redo_action, bg="#FFC107", fg="black", padx=20)
        redo_button.pack(side=tk.LEFT, padx=10, pady=5)

        self.cells = {}

        for col in range(10):
            label = tk.Entry(self.root, text=chr(65 + col), borderwidth=1, relief="solid", bg="#D3D3D3")
            label.grid(row=1, column=col + 1, sticky="nsew", padx=1, pady=1)

        for row in range(10):
            label = tk.Entry(self.root, text=str(row + 1), borderwidth=1, relief="solid", bg="#D3D3D3")
            label.grid(row=row + 2, column=0, sticky="nsew", padx=1, pady=1)
            label.bind('<Button-1>', self.select_row)

        for row in range(10):
            for col in range(10):
                entry = tk.Entry(self.root, width=10, justify="center", font=("Arial", 12))
                entry.grid(row=row + 2, column=col + 1, sticky="nsew", padx=1, pady=1)
                self.cells[(row + 1, col + 1)] = entry
                entry.bind('<Return>', self.process_formula)
                entry.bind('<Button-1>', self.select_cell)
                entry.bind('<FocusOut>', self.save_state)

        for i in range(11):
            self.root.grid_columnconfigure(i, weight=1)
        for i in range(12):
            self.root.grid_rowconfigure(i, weight=1)

    def create_grid(self):
        self.cells = {}

        for row in range(10):
            label = tk.Label(self.root, text=str(row + 1), borderwidth=1, relief="solid", bg="#D3D3D3")
            label.grid(row=row + 2, column=0, sticky="nsew", padx=1, pady=1)
            label.configure(width=4)
            label.bind('<Button-1>', self.select_row)

        for col in range(10):
            label = tk.Label(self.root, text=chr(65 + col), borderwidth=1, relief="solid", bg="#D3D3D3")
            label.grid(row=1, column=col + 1, sticky="nsew", padx=1, pady=1)
            label.bind('<Button-1>', self.select_column)

        for row in range(10):
            for col in range(10):
                entry = tk.Entry(self.root, width=10, justify="center", font=("Arial", 12))
                entry.grid(row=row + 2, column=col + 1, sticky="nsew", padx=1, pady=1)
                self.cells[(row + 1, col + 1)] = entry
                entry.bind('<Return>', self.process_formula)
                entry.bind('<Button-1>', self.select_cell)
                entry.bind('<FocusOut>', self.save_state)

        for i in range(11):
            self.root.grid_columnconfigure(i, weight=1)
        for i in range(12):
            self.root.grid_rowconfigure(i, weight=1)

    def process_formula(self, event):
        widget = event.widget
        cell = self.get_cell_coordinates(widget)
        if cell is None:
            return

        formula = self.cells[cell].get()
        if not formula.startswith('='):
            return

        formula = formula[1:]
        if '+' in formula:
            self.calculate_sum(cell, formula)
        elif '*' in formula:
            self.calculate_product(cell, formula)
        elif formula.lower().startswith('avr(') and formula.endswith(')'):
            self.calculate_average(cell, formula[4:-1])
        elif formula.lower().startswith('max(') and formula.endswith(')'):
            self.calculate_max(cell, formula[4:-1])
        elif formula.lower().startswith('min(') and formula.endswith(')'):
            self.calculate_min(cell, formula[4:-1])
        else:
            messagebox.showerror("Error", "Invalid formula.")

    def calculate_sum(self, cell, formula):
        cell_refs = formula.split('+')
        total_sum = 0.0
        for cell_ref in cell_refs:
            cell_ref = cell_ref.strip()
            try:
                value = self.get_cell_value(cell_ref)
                total_sum += float(value)
            except ValueError:
                messagebox.showerror("Error", f"Invalid number in cell {cell_ref}.")
                return
            except KeyError:
                messagebox.showerror("Error", f"Invalid cell reference: {cell_ref}.")
                return

        self.cells[cell].delete(0, tk.END)
        self.cells[cell].insert(0, str(total_sum))

    def calculate_product(self, cell, formula):
        cell_refs = formula.split('*')
        total_product = 1.0
        for cell_ref in cell_refs:
            cell_ref = cell_ref.strip()
            try:
                value = self.get_cell_value(cell_ref)
                total_product *= float(value)
            except ValueError:
                messagebox.showerror("Error", f"Invalid number in cell {cell_ref}.")
                return
            except KeyError:
                messagebox.showerror("Error", f"Invalid cell reference: {cell_ref}.")
                return

        self.cells[cell].delete(0, tk.END)
        self.cells[cell].insert(0, str(total_product))

    def calculate_average(self, cell, formula):
        cell_refs = formula.split(',')
        total_sum = 0.0
        count = 0
        for cell_ref in cell_refs:
            cell_ref = cell_ref.strip()
            try:
                value = self.get_cell_value(cell_ref)
                total_sum += float(value)
                count += 1
            except ValueError:
                messagebox.showerror("Error", f"Invalid number in cell {cell_ref}.")
                return
            except KeyError:
                messagebox.showerror("Error", f"Invalid cell reference: {cell_ref}.")
                return

        if count > 0:
            average = total_sum / count
            self.cells[cell].delete(0, tk.END)
            self.cells[cell].insert(0, str(average))
        else:
            messagebox.showerror("Error", "No valid cells to average.")

    def calculate_max(self, cell, formula):
        cell_refs = formula.split(',')
        values = []
        for cell_ref in cell_refs:
            cell_ref = cell_ref.strip()
            try:
                value = self.get_cell_value(cell_ref)
                values.append(float(value))
            except ValueError:
                messagebox.showerror("Error", f"Invalid number in cell {cell_ref}.")
                return
            except KeyError:
                messagebox.showerror("Error", f"Invalid cell reference: {cell_ref}.")
                return

        if values:
            max_value = max(values)
            self.cells[cell].delete(0, tk.END)
            self.cells[cell].insert(0, str(max_value))
        else:
            messagebox.showerror("Error", "No valid cells to find the maximum.")

    def calculate_min(self, cell, formula):
        cell_refs = formula.split(',')
        values = []
        for cell_ref in cell_refs:
            cell_ref = cell_ref.strip()
            try:
                value = self.get_cell_value(cell_ref)
                values.append(float(value))
            except ValueError:
                messagebox.showerror("Error", f"Invalid number in cell {cell_ref}.")
                return
            except KeyError:
                messagebox.showerror("Error", f"Invalid cell reference: {cell_ref}.")
                return

        if values:
            min_value = min(values)
            self.cells[cell].delete(0, tk.END)
            self.cells[cell].insert(0, str(min_value))
        else:
            messagebox.showerror("Error", "No valid cells to find the minimum.")

    def get_cell_coordinates(self, widget):
        for (row, col), entry in self.cells.items():
            if entry == widget:
                return row, col
        return None

    def get_cell_value(self, cell_ref):
        row_col = self.convert_cell_reference(cell_ref)
        if row_col and row_col in self.cells:
            return self.cells[row_col].get()
        return None

    def convert_cell_reference(self, cell_ref):
        if len(cell_ref) < 2:
            return None

        col = ord(cell_ref[0].upper()) - ord('A') + 1
        try:
            row = int(cell_ref[1:])
            return row, col
        except ValueError:
            return None

    def save_state(self, event=None):
        state = {key: entry.get() for key, entry in self.cells.items()}
        self.manager.push(state)

    def undo_action(self):
        state = self.manager.undo()
        if state is not None:
            for key, value in state.items():
                self.cells[key].delete(0, tk.END)
                self.cells[key].insert(0, value)

    def redo_action(self):
        state = self.manager.redo()
        if state is not None:
            for key, value in state.items():
                self.cells[key].delete(0, tk.END)
                self.cells[key].insert(0, value)

    def save_file(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if file_path:
            with open(file_path, 'w', newline='') as file:
                writer = csv.writer(file)
                for row in range(1, 11):
                    writer.writerow([self.cells.get((row, col), '').get() for col in range(1, 11)])

    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if file_path:
            with open(file_path, 'r') as file:
                reader = csv.reader(file)
                for row_index, row in enumerate(reader, start=1):
                    for col_index, cell_value in enumerate(row, start=1):
                        if (row_index, col_index) in self.cells:
                            self.cells[(row_index, col_index)].delete(0, tk.END)
                            self.cells[(row_index, col_index)].insert(0, cell_value)

