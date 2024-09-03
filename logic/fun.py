import tkinter as tk
from tkinter import filedialog, messagebox, font, colorchooser
import csv
from logic.undo_redo import UndoRedoManager

class ExcelApp:
    def __init__(self, root):
        self.col_count = 20
        self.row_count = 20

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
        

        #Mila
        self.create_canvas_and_scrollbars()

        self.selected_row = None
        self.selected_col = None
        self.selected_cell = None

        self.grid_frame = tk.Frame(self.canvas, bg="#f0f0f0")
        self.canvas.create_window((0, 0), window=self.grid_frame, anchor="nw")

        self.grid_frame.update_idletasks()
        self.canvas.config(scrollregion=self.canvas.bbox("all"))

        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        self.cell_width = 10
        self.cell_height = 2

        self.create_buttons()
        self.create_grid()

    #selekcija

    def clear_selection(self):
        for (row, col), entry in self.cells.items():
            entry.config(borderwidth=1, relief="flat")

    def select_row(self, event):
        row = int(event.widget.cget("text"))
        self.clear_selection()
        self.highlight_row(row)
        self.selected_row = row

    def select_column(self, event):
        col = self.find_col_index(event)
        self.clear_selection()
        self.highlight_column(col)
        self.selected_col = col

    def find_col_index(self, event):
        if len(event.widget.cget("text")) == 1:
            return ord(event.widget.cget("text")) - 64
        else:
            return (ord(event.widget.cget("text")[0]) - 64) * 26 + ord(event.widget.cget("text")[1]) - 64

    def highlight_row(self, row):
        for col in range(1, self.col_count):
            cell = (row, col)
            if cell in self.cells:
                entry = self.cells[cell]
                entry.config(borderwidth=1, relief="solid")

    def highlight_column(self, col):
        for row in range(1, self.row_count):
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

    #scrollbar

    def update_scrollregion(self):
        self.canvas.update_idletasks()
        self.canvas.config(scrollregion=self.canvas.bbox("all"))

    def on_vertical_scroll(self, *args):
        self.canvas.yview_moveto(float(args[1]))
        self.check_vertical_scroll_end()

    def check_vertical_scroll_end(self):
        pos = self.canvas.yview()[1]
        
        if pos >= 1.0:
            self.add_more_rows()

    def on_horizontal_scroll(self, *args):
        self.canvas.xview_moveto(float(args[1]))
        self.check_horizontal_scroll_end()

    def check_horizontal_scroll_end(self):
        pos = self.canvas.xview()[1]

        if pos >= 1.0:
            self.add_more_columns()

    def add_more_columns(self):
            dodate_kolone = 2
            self.col_count += dodate_kolone

            if 65 + self.col_count > 91:
                first_letter = chr((self.col_count - 1) // 26 + 64)
            else:
                first_letter = ""


            for col in range(self.col_count - dodate_kolone, self.col_count):
                label = tk.Label(self.grid_frame, text=first_letter + chr((col % 26) + 65), borderwidth=1, relief="solid", bg="#D3D3D3",
                                width=self.cell_width, height=self.cell_height)
                label.grid(row=1, column=col+1, sticky="nsew", padx=1, pady=1)
                label.bind('<Button-1>', self.select_column)

            for row in range(0, self.row_count):
                for col in range(self.col_count - dodate_kolone, self.col_count):
                    entry = tk.Entry(self.grid_frame, width=self.cell_width, font=("Arial", 12), justify="center", relief="flat" if self.selected_cell or self.selected_col or self.selected_row else "sunken")
                    entry.grid(row = row + 2, column=col+1, sticky="nsew", padx=1, pady=1)
                    self.cells[(row + 1, col + 1)] = entry
                    entry.bind('<Return>', self.process_formula)
                    entry.bind('<Button-1>', self.select_cell)

            if self.selected_row != None:
                row = self.selected_row
                for col in range(self.col_count - dodate_kolone, self.col_count):
                    entry = self.cells[(row, col)]
                    entry.config(borderwidth=1, relief="solid")


            self.update_scrollregion()

    def add_more_rows(self):
        dodati_redovi = 2
        self.row_count += dodati_redovi

        for row in range(self.row_count - dodati_redovi, self.row_count):
            label = tk.Label(self.grid_frame, text=str(row + 1), borderwidth=1, relief="solid", bg="#D3D3D3",
                             width=self.cell_width, height=self.cell_height)
            label.grid(row=row+2, column=0, sticky="nsew", padx=1, pady=1)
            label.configure(width=4)
            label.bind('<Button-1>', self.select_row)

        for col in range(0, self.col_count):
                for row in range(self.row_count - dodati_redovi, self.row_count):
                    entry = tk.Entry(self.grid_frame, width=self.cell_width, font=("Arial", 12), justify="center", relief="flat" if self.selected_cell or self.selected_col or self.selected_row else "sunken")
                    entry.grid(row = row + 2, column=col+1, sticky="nsew", padx=1, pady=1)
                    self.cells[(row + 1, col + 1)] = entry
                    entry.bind('<Return>', self.process_formula)
                    entry.bind('<Button-1>', self.select_cell)

        if self.selected_col != None:
            col = self.selected_col
            for row in range(self.row_count - dodati_redovi, self.row_count):
                entry = self.cells[(row, col)]
                entry.config(borderwidth=1, relief="solid")

        self.update_scrollregion()

    def on_canvas_resize(self, event):
        h = self.root.winfo_height() - 120
        w = self.root.winfo_width() - 20
        self.canvas.config(width=w, height=h)
        self.canvas.update_idletasks() 
        self.canvas.configure(scrollregion=self.grid_frame.bbox("all"))

    def create_canvas_and_scrollbars(self):
        self.canvas = tk.Canvas(self.root, height=self.root.winfo_height() - 120, width=self.root.winfo_width() - 20)
        self.canvas.grid(row=1, column=0)

        self.v_scroll = tk.Scrollbar(self.root, orient="vertical", command=self.on_vertical_scroll)
        self.v_scroll.grid(row=1, column=1, sticky="ns")
        
        self.h_scroll = tk.Scrollbar(self.root, orient="horizontal", command=self.on_horizontal_scroll)
        self.h_scroll.grid(row=2, column=0, sticky="ew") 

        self.canvas.configure(yscrollcommand=self.v_scroll.set, xscrollcommand=self.h_scroll.set)

        self.root.grid_rowconfigure(1, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        self.canvas.bind("<Configure>", self.on_canvas_resize)

        self.update_scrollregion()


    #Milica
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
            command=lambda: self.save_file(self.cells),
            bg="#4CAF50",
            fg="white",
            padx=20
        )
        save_button.pack(side=tk.LEFT, padx=10, pady=5)

        load_button = tk.Button(
            button_frame,
            text="Load",
            command=lambda: self.load_file(self.cells),
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

    def create_grid(self):
        self.cells = {}

        for row in range(self.row_count):
            label = tk.Label(self.grid_frame, text=str(row + 1), borderwidth=1, relief="solid", bg="#D3D3D3",
                             width=self.cell_width, height=self.cell_height)
            
            label.grid(row=row + 2, column=0, sticky="nsew", padx=1, pady=1)
            label.configure(width=4)
            label.bind('<Button-1>', self.select_row)

        for col in range(self.col_count):
            label = tk.Label(self.grid_frame, text=chr(65 + col), borderwidth=1, relief="solid", bg="#D3D3D3",
                             width=self.cell_width, height=self.cell_height)
            label.grid(row=1, column=col + 1, sticky="nsew", padx=1, pady=1)
            label.bind('<Button-1>', self.select_column)

        for row in range(self.row_count):
            for col in range(self.col_count):
                entry = tk.Entry(self.grid_frame, width=self.cell_width, justify="center", font=("Arial", 12))
                entry.grid(row=row + 2, column=col + 1, sticky="nsew", padx=1, pady=1)
                self.cells[(row + 1, col + 1)] = entry

                entry.bind('<Return>', self.process_formula)
                entry.bind('<Button-1>', self.select_cell)
                entry.bind('<FocusOut>', self.save_state)
                entry.bind("<FocusIn>", self.save_initial_state)

        for i in range(self.col_count):
            self.root.grid_columnconfigure(i, weight=1)
        for i in range(self.row_count):
            self.root.grid_rowconfigure(i, weight=1)


    def save_initial_state(self, event=None):
        self.initial_state = {key: entry.get() for key, entry in self.cells.items()}

    def get_cell_coordinates(self, widget):
        for (row, col), entry in self.cells.items():
            if entry == widget:
                return row, col
        return None


    #formule

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
        else:
            messagebox.showerror("Error", "Invalid formula.")


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

    def save_initial_state(self, event=None):
        self.initial_state = {key: entry.get() for key, entry in self.cells.items()}

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




    #save and load
    def save_file(self, cells):
        file_path = tk.filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", ".csv")])
        if file_path:
            try:
                with open(file_path, 'w', newline='') as file:
                    writer = csv.writer(file)
                    for row in range(1, self.row_count):
                        row_data = []
                        for col in range(1, self.col_count):
                            value = cells[(row, col)].get()
                            row_data.append(value)
                        writer.writerow(row_data)
                messagebox.showinfo("Success", "File saved successfully.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save file: {str(e)}")


    def load_file(self, cells):
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
