import tkinter as tk
from tkinter import filedialog, messagebox, font, colorchooser
import csv
from logic.undo_redo import UndoRedoManager
from logic.save_and_load import save_file, load_file
import re


class ExcelApp:
    def __init__(self, root):
        self.cells = {}
        self.root = root
        self.root.title("Excel")
        self.root.geometry("800x600")
        self.initial_state = {}

        self.fonts = ["Arial", "Courier", "Times", "Helvetica", "Verdana"]
        self.font_sizes = ["8", "10", "12", "14", "16", "18", "20", "24", "28", "32", "36"]

        self.selected_font = tk.StringVar(value=self.fonts[0])
        self.selected_font_size = tk.StringVar(value=self.font_sizes[2])

        self.manager = UndoRedoManager()
        self.create_buttons()
        self.create_grid()

        self.selected_row = None
        self.selected_col = None
        self.selected_cell = None

    def start_selection(self, event):
        cell = self.get_cell_coordinates(event.widget)
        if cell:
            self.selection_start = cell
            self.update_selection()

    def extend_selection(self, event):
        cell = self.get_cell_coordinates(event.widget)
        if cell:
            self.selection_end = cell
            self.update_selection()

    def end_selection(self, event):
        cell = self.get_cell_coordinates(event.widget)
        if cell:
            self.selection_end = cell
            self.update_selection()

    def update_selection(self):
        self.clear_selection()

        # Check if selection_start and selection_end are provided
        if self.selection_start and self.selection_end:
            start_ref, end_ref = self.selection_start, self.selection_end
            # Convert cell references to row and column indices
            start_row_col = self.convert_cell_reference(start_ref)
            end_row_col = self.convert_cell_reference(end_ref)

            if start_row_col and end_row_col:
                start_row, start_col = start_row_col
                end_row, end_col = end_row_col

                # Ensure start_row <= end_row and start_col <= end_col
                for row in range(min(start_row, end_row), max(start_row, end_row) + 1):
                    for col in range(min(start_col, end_col), max(start_col, end_col) + 1):
                        if (row, col) in self.cells:
                            self.cells[(row, col)].config(bg="#ADD8E6")

    def clear_selection(self):
        for (row, col), entry in self.cells.items():
            entry.config(bg="white")

    def select_row(self, event):
        row = int(event.widget.cget("text"))
        self.clear_selection()
        self.highlight_row(row)
        self.selected_row = row

    def select_column(self, event):
        col = ord(event.widget.cget("text")) - 64
        self.clear_selection()
        self.highlight_column(col)
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
            self.clear_selection()
            self.highlight_border(clicked_cell)
            self.selected_cell = clicked_cell

    def highlight_border(self, cell):
        if cell in self.cells:
            entry = self.cells[cell]
            entry.config(borderwidth=1, relief="solid")

    def clear_selection(self):
        for (row, col), entry in self.cells.items():
            entry.config(borderwidth=1, relief="flat")

    def format_bold(self):
        if self.selected_cell:
            row, col = self.selected_cell
            entry = self.cells[(row, col)]
            current_font = entry.cget("font")
            new_font = ("Arial", 12, "bold" if "bold" not in current_font else "normal")
            entry.config(font=new_font)

    def format_italic(self):
        if self.selected_cell:
            row, col = self.selected_cell
            entry = self.cells[(row, col)]
            current_font = entry.cget("font")
            new_font = ("Arial", 12, "italic" if "italic" not in current_font else "normal")
            entry.config(font=new_font)

    def underline_text(self):
        if self.selected_cell:
            row, col = self.selected_cell
            entry = self.cells[(row, col)]

            current_font = font.Font(font=entry.cget("font"))
            underline = current_font.cget("underline")

            current_font.configure(underline=0 if underline else 1)
            entry.config(font=current_font)

    def change_font(self, selected_font):
        if self.selected_cell:
            row, col = self.selected_cell
            entry = self.cells[(row, col)]

            current_font = font.Font(font=entry.cget("font"))
            current_font.configure(family=selected_font)

            entry.config(font=current_font)

    def change_font_size(self, selected_size):
        if self.selected_cell:
            row, col = self.selected_cell
            entry = self.cells[(row, col)]

            current_font = font.Font(font=entry.cget("font"))
            current_font.configure(size=int(selected_size))

            entry.config(font=current_font)

    def change_text_color(self):
        if self.selected_cell:
            row, col = self.selected_cell
            entry = self.cells[(row, col)]

            color = colorchooser.askcolor()[1]
            if color:
                entry.config(fg=color)

    def change_cell_color(self):
        if self.selected_cell:
            row, col = self.selected_cell
            entry = self.cells[(row, col)]
            color = colorchooser.askcolor()[1]
            if color:
                entry.config(bg=color)

    def align_left(self):
        if self.selected_cell:
            row, col = self.selected_cell
            entry = self.cells[(row, col)]
            entry.config(justify="left")

    def align_center(self):
        if self.selected_cell:
            row, col = self.selected_cell
            entry = self.cells[(row, col)]
            entry.config(justify="center")

    def align_right(self):
        if self.selected_cell:
            row, col = self.selected_cell
            entry = self.cells[(row, col)]
            entry.config(justify="right")

    def create_buttons(self):
        self.cells = {}
        button_frame = tk.Frame(self.root, bg="#f0f0f0", bd=1, relief="solid")
        button_frame.grid(row=0, column=0, columnspan=11, pady=10, padx=10, sticky="ew")

        save_button = tk.Button(
            button_frame,
            text="Save",
            command=lambda: save_file(self.cells),
            bg="#4CAF50",
            fg="white",
            padx=20
        )
        save_button.pack(side=tk.LEFT, padx=10, pady=5)

        load_button = tk.Button(
            button_frame,
            text="Load",
            command=lambda: load_file(self.cells),
            bg="#008CBA",
            fg="white",
            padx=20
        )
        load_button.pack(side=tk.LEFT, padx=10, pady=5)

        undo_button = tk.Button(button_frame, text="Undo", command=self.undo_action, bg="#FFC107", fg="black", padx=20)
        undo_button.pack(side=tk.LEFT, padx=10, pady=5)

        redo_button = tk.Button(button_frame, text="Redo", command=self.redo_action, bg="#FFC107", fg="black", padx=20)
        redo_button.pack(side=tk.LEFT, padx=10, pady=5)

        separator = tk.Frame(button_frame, width=2, bg="#d3d3d3", height=40)
        separator.pack(side=tk.LEFT, padx=10, pady=5, fill=tk.Y)

        format_frame = tk.Frame(button_frame, bg="#f0f0f0")
        format_frame.pack(side=tk.LEFT, padx=10, pady=5)

        bold_button = tk.Button(format_frame, text="Bold", command=self.format_bold, bg="#FFC107", fg="black", padx=15)
        bold_button.pack(side=tk.TOP, padx=5, pady=5)

        italic_button = tk.Button(format_frame, text="Italic", command=self.format_italic, bg="#FFC107", fg="black",
                                  padx=15)
        italic_button.pack(side=tk.TOP, padx=5, pady=5)

        underline_button = tk.Button(button_frame, text="Underline", command=self.underline_text, bg="#f0f0f0",
                                     fg="black", padx=20)
        underline_button.pack(side=tk.LEFT, padx=10, pady=5)

        font_menu = tk.OptionMenu(button_frame, self.selected_font, *self.fonts, command=self.change_font)
        font_menu.pack(side=tk.LEFT, padx=10, pady=5)

        font_size_menu = tk.OptionMenu(button_frame, self.selected_font_size, *self.font_sizes,
                                       command=self.change_font_size)
        font_size_menu.pack(side=tk.LEFT, padx=10, pady=5)

        color_button = tk.Button(button_frame, text="Text Color", command=self.change_text_color, bg="#FFC107",
                                 fg="black", padx=20)
        color_button.pack(side=tk.LEFT, padx=10, pady=5)

        cell_color_button = tk.Button(button_frame, text="Cell Color", command=self.change_cell_color, bg="#FFC107",
                                      fg="black", padx=20)
        cell_color_button.pack(side=tk.LEFT, padx=10, pady=5)

        align_frame = tk.Frame(button_frame, bg="#f0f0f0")
        align_frame.pack(side=tk.LEFT, padx=10, pady=5)

        left_align_button = tk.Button(align_frame, text="Left", command=self.align_left, bg="#FFC107", fg="black",
                                      padx=10)
        left_align_button.pack(side=tk.LEFT, padx=5, pady=5)

        center_align_button = tk.Button(align_frame, text="Center", command=self.align_center, bg="#FFC107", fg="black",
                                        padx=10)
        center_align_button.pack(side=tk.LEFT, padx=5, pady=5)

        right_align_button = tk.Button(align_frame, text="Right", command=self.align_right, bg="#FFC107", fg="black",
                                       padx=10)
        right_align_button.pack(side=tk.LEFT, padx=5, pady=5)


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

                # Bind events to the entry widgets
                entry.bind('<Return>', self.process_formula)
                entry.bind('<Button-1>', self.select_cell)
                entry.bind('<FocusOut>', self.save_state)
                entry.bind("<FocusIn>", self.save_initial_state)  # Add this line

        for i in range(11):
            self.root.grid_columnconfigure(i, weight=1)
        for i in range(12):
            self.root.grid_rowconfigure(i, weight=1)

    def save_initial_state(self, event=None):
        # Save the initial state of all cells
        self.initial_state = {key: entry.get() for key, entry in self.cells.items()}

    def get_cell_coordinates(self, widget):
        for (row, col), entry in self.cells.items():
            if entry == widget:
                return row, col
        return None

    def process_cell_input(self, input_value):
        if input_value.startswith('='):
            cell_ref = input_value[1:]
            if self.is_cell_reference(cell_ref):
                value = self.get_cell_value(cell_ref)
                return value
        return input_value

    def is_cell_reference(self, reference):
        if len(reference) < 2:
            return False
        col = reference[0].upper()
        row: int = reference[1:]
        return col.isalpha() and row.isdigit()

    def convert_to_indices(self, reference):
        col = ord(reference[0].upper()) - ord('A')
        row = int(reference[1:]) - 1
        return (row, col)

    def get_cell_value_from_ref(self, cell_ref):
        # Ova funkcija treba da pretvori cell_ref u red i kolonu i da dobije vrednost iz tih koordinata
        row, col = self.parse_cell_reference(cell_ref)
        return self.get_cell_value(row, col)

    def parse_cell_reference(self, cell_ref):
        # Ovo pretvara referencu ćelije (npr. "A1") u redni i kolonski broj
        # Ovo treba implementirati prema vašoj logici, npr. koristeći ASCII vrednosti za kolone
        pass



    def reference_cell(self, formula):

        if formula.lower().startswith('ref(') and formula.endswith(')'):

            reference_cell = formula[4:-1]

            value = self.get_cell_value(reference_cell)

            if value is not None:
                return value
            else:
                messagebox.showerror("Error", f"Vrednost ćelije {reference_cell} nije dostupna.")
                return None
        else:
            messagebox.showerror("Error", "Formula nije u ispravnom formatu. Očekuje se REF(A1).")
            return None

    def update_cell_value(self, row, col, value):
        entry = self.cells[(row, col)]
        if value.startswith('='):
            formula = value[1:]
            result = self.process_formula(formula)
            if result is not None:
                entry.delete(0, tk.END)
                entry.insert(0, result)
                return
        entry.delete(0, tk.END)
        entry.insert(0, value)

    def set_cell_value(self, cell, value):
        if cell in self.cells:
            self.cells[cell].delete(0, tk.END)
            self.cells[cell].insert(0, value)

    def calculate_obj(self, cell_ref):
        cell_ref = cell_ref.upper()
        cell_coords = self.parse_cell_reference(cell_ref)
        if cell_coords:
            row, col = cell_coords
            cell_value = self.get_cell_value(row, col)
            return cell_value
        else:
            return "ERROR"

    def connecting_cell(self, cell_ref):
        if self.is_cell_reference(cell_ref):
            value = self.get_cell_value(cell_ref)
            if value is not None:
                self.cells[self.current_cell].set(value)
            else:
                print(f"Invalid cell reference: {cell_ref}")
        else:
            print(f"Invalid cell reference format: {cell_ref}")

    def process_formula(self, event):
        widget = event.widget
        cell = self.get_cell_coordinates(widget)
        if cell is None:
            return

        formula = self.cells[cell].get()
        if not formula.startswith('='):
            return

        formula = formula[1:]



        if formula.lower().startswith('sum(') and formula.endswith(')'):
            self.calculate_sum(cell, formula[4:-1])
        elif formula.lower().startswith('prd(') and formula.endswith(')'):
            self.calculate_product(cell, formula[4:-1])
        elif formula.lower().startswith('avr(') and formula.endswith(')'):
            self.calculate_average(cell, formula[4:-1])
        elif formula.lower().startswith('max(') and formula.endswith(')'):
            self.calculate_max(cell, formula[4:-1])
        elif formula.lower().startswith('min(') and formula.endswith(')'):
            self.calculate_min(cell, formula[4:-1])
        elif formula.lower().startswith('det(') and formula.endswith(')'):
            self.calculate_detraction(cell, formula[4:-1])
        elif formula.lower().startswith('mod(') and formula.endswith(')'):
            self.calculate_modul(cell, formula[4:-1])
        elif formula.lower().startswith('pow(') and formula.endswith(')'):
            self.calculate_power(cell, formula[4:-1])
        elif formula.lower().startswith('sumif(') and formula.endswith(')'):
            self.calculate_sumif(cell, formula[6:-1])
        elif formula.lower().startswith('prdif(') and formula.endswith(')'):
            self.calculate_productif(cell, formula[6:-1])
        elif formula.lower().startswith('avrif(') and formula.endswith(')'):
            self.calculate_avrif(cell, formula[6:-1])
        elif formula.lower().startswith('obj(') and formula.endswith(')'):
            cell_ref = formula[4:-1]  # Uklanja 'OBJ(' i ')'
            result = self.calculate_obj(cell_ref)
        else :
            result = self.get_cell_value(formula)
            self.cells[cell].delete(0, tk.END)

            self.cells[cell].insert(0, str(result))

    def calculate_sumif(self, cell, formula):
        first_division = formula.split(';')


        if len(first_division) != 2:
            messagebox.showerror("Error", "sumif formula requires exactly two arguments separated by ';'.")
            return

        condition = first_division[0].strip()
        cell_refs = first_division[1].split(',')

        condition_match = re.match(r'^(-?\d+(\.\d+)?)(>|<|=)$', condition)
        if not condition_match:
            messagebox.showerror("Error", "Invalid condition format. Correct format: number followed by >, <, or =.")
            return

        condition_value = float(condition_match.group(1))
        operator = condition_match.group(3)
        total_sum = 0.0

        for cell_ref in cell_refs:
            cell_ref = cell_ref.strip()
            try:
                value_str = self.get_cell_value(cell_ref)
                value = float(value_str)

                if ((operator == '<' and value > condition_value) or
                        (operator == '>' and value < condition_value) or
                        (operator == '=' and value == condition_value)):
                    total_sum += value

            except ValueError:
                messagebox.showerror("Error", f"Invalid number in cell {cell_ref}.")
                return

        self.cells[cell].delete(0, tk.END)
        self.cells[cell].insert(0, str(total_sum))

    def calculate_productif(self, cell, formula):
        first_division = formula.split(';')

        if len(first_division) != 2:
            messagebox.showerror("Error", "prdif formula requires exactly two arguments separated by ';'.")
            return

        condition = first_division[0].strip()
        cell_refs = first_division[1].split(',')


        condition_match = re.match(r'^(-?\d+(\.\d+)?)(>|<|=)$', condition)
        if not condition_match:
            messagebox.showerror("Error", "Invalid condition format. Correct format: number followed by >, <, or =.")
            return

        condition_value = float(condition_match.group(1))
        operator = condition_match.group(3)
        total_prd = 1.0

        for cell_ref in cell_refs:
            cell_ref = cell_ref.strip()
            try:
                value_str = self.get_cell_value(cell_ref)
                value = float(value_str)


                if ((operator == '<' and value > condition_value) or
                        (operator == '>' and value < condition_value) or
                        (operator == '=' and value == condition_value)):
                    total_prd *= value

            except ValueError:
                messagebox.showerror("Error", f"Invalid number in cell {cell_ref}.")
                return

        self.cells[cell].delete(0, tk.END)
        self.cells[cell].insert(0, str(total_prd))

    def calculate_avrif(self, cell, formula):
        first_division = formula.split(';')


        if len(first_division) != 2:
            messagebox.showerror("Error", "avrif formula requires exactly two arguments separated by ';'.")
            return

        condition = first_division[0].strip()
        cell_refs = first_division[1].split(',')

        condition_match = re.match(r'^(-?\d+(\.\d+)?)(>|<|=)$', condition)
        if not condition_match:
            messagebox.showerror("Error", "Invalid condition format. Correct format: number followed by >, <, or =.")
            return

        condition_value = float(condition_match.group(1))
        operator = condition_match.group(3)
        total_avr = 0.0
        count = 0

        for cell_ref in cell_refs:
            cell_ref = cell_ref.strip()
            try:
                value_str = self.get_cell_value(cell_ref)
                value = float(value_str)

                if ((operator == '<' and value > condition_value) or
                        (operator == '>' and value < condition_value) or
                        (operator == '=' and value == condition_value)):
                    total_avr += value
                    count+=1

            except ValueError:
                messagebox.showerror("Error", f"Invalid number in cell {cell_ref}.")
                return
        if count > 0:
            average = total_avr / count
            self.cells[cell].delete(0, tk.END)
            self.cells[cell].insert(0, str(average))
        else:
            messagebox.showerror("Error", "No valid cells to average.")

    #obične operqcije
    def calculate_sum(self, cell, formula):
        cell_refs = formula.split(',')
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
        cell_refs = formula.split(',')
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

    def calculate_detraction(self, cell, formula):
        cell_refs = formula.split(",")
        total_det = 0.0
        if len(cell_refs) != 2:
            messagebox.showerror("Error", f"The formula must contain exactly two cell references.")

        cell_ref1 = cell_refs[0].strip()
        cell_ref2 = cell_refs[1].strip()

        try:
            value1 = self.get_cell_value(cell_ref1)
            value2 = self.get_cell_value(cell_ref2)
            total_det = float(value1) - float(value2)
            self.cells[cell].delete(0, tk.END)
            self.cells[cell].insert(0, str(total_det))

            total_det = float(value1) - float(value2)
            self.cells[cell].delete(0, tk.END)
            self.cells[cell].insert(0, str(total_det))
        except ValueError:
            messagebox.showerror("Error", f"Invalid number in cell {cell_ref1}.")
            return

        except KeyError:
            messagebox.showerror("Error", f"Invalid cell reference: {cell_ref1}.")
            return

    def calculate_modul(self, cell, formula):
        cell_refs = formula.split(",")
        total_det = 0.0
        if len(cell_refs) != 2:
            messagebox.showerror("Error", f"The formula must contain exactly two cell references.")

        cell_ref1 = cell_refs[0].strip()
        cell_ref2 = cell_refs[1].strip()

        try:
            value1 = self.get_cell_value(cell_ref1)
            value2 = self.get_cell_value(cell_ref2)
            total_det = float(value1) - float(value2)
            self.cells[cell].delete(0, tk.END)
            self.cells[cell].insert(0, str(total_det))

            total_det = float(value1) / float(value2)
            self.cells[cell].delete(0, tk.END)
            self.cells[cell].insert(0, str(total_det))
        except ValueError:
            messagebox.showerror("Error", f"Invalid number in cell {cell_ref1}.")
            return

        except KeyError:
            messagebox.showerror("Error", f"Invalid cell reference: {cell_ref1}.")
            return

    def calculate_power(self, cell, formula):
        cell_refs = formula.split(",")
        total_det = 0.0
        if len(cell_refs) != 2:
            messagebox.showerror("Error", f"The formula must contain exactly two cell references.")
        else:
            cell_ref1 = cell_refs[0].strip()
            cell_ref2 = cell_refs[1].strip()

            try:
                value1 = self.get_cell_value(cell_ref1)
                value2 = self.get_cell_value(cell_ref2)
                total_det = pow(float(value1), float(value2))
                self.cells[cell].delete(0, tk.END)
                self.cells[cell].insert(0, str(total_det))

            except ValueError:
                messagebox.showerror("Error", f"Invalid number in cell {cell_ref1}.")
                return

            except KeyError:
                messagebox.showerror("Error", f"Invalid cell reference: {cell_ref1}.")
                return

    def get_cell_coordinates(self, widget):
        for (row, col), entry in self.cells.items():
            if entry == widget:
                return row, col
        return None


    def get_cell_value(self, cell_ref):
        row_col = self.convert_cell_reference(cell_ref)
        print(row_col)
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
        new_state = {key: entry.get() for key, entry in self.cells.items()}
        if new_state != self.initial_state:
            self.manager.push(new_state)

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

