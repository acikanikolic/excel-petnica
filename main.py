import tkinter as tk
from logic.fun import ExcelApp

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelApp(root)
    app.save_state()
    root.mainloop()
