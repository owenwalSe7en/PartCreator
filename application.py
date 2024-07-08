import tkinter as tk
from tkinter import ttk, messagebox
from erp_manager import OperationType
from forms import CreateForm, OverwriteForm, DeleteForm


class Application:
    def __init__(self, erp_manager):
        self.erp_manager = erp_manager
        self.root = tk.Tk()
        self.root.title("ERP Part Manager")
        self.root.geometry("650x170")
        self.root.minsize(650, 170)

    def create_ui(self):
        style = ttk.Style()
        style.configure('TButton', font=('Arial', 14))
        style.configure('TLabel', font=('Arial', 18))

        # Create a new style for taller buttons
        style.configure('Tall.TButton', font=('Arial', 14), padding=(10, 20))

        ttk.Label(self.root, text="Select an operation:", style='TLabel').pack(pady=(20, 10))

        button_frame = ttk.Frame(self.root)
        button_frame.pack(pady=10)

        ttk.Button(button_frame, text="Create", command=lambda: self.open_form(CreateForm, OperationType.CREATE),
                   style='Tall.TButton', width=15).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="Overwrite", command=lambda: self.open_form(OverwriteForm,
                                                                                  OperationType.OVERWRITE),
                   style='Tall.TButton', width=15).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="Delete", command=lambda: self.open_form(DeleteForm, OperationType.DELETE),
                   style='Tall.TButton', width=15).pack(side=tk.LEFT, padx=10)

    def open_form(self, FormClass, operation_type):
        form_window = tk.Toplevel(self.root)
        form = FormClass(form_window)
        form.create_form()
        self.root.withdraw() # Hide the main window
        self.root.wait_window(form_window)
        if form.data:
            try:
                self.erp_manager.perform_operation(operation_type, form.data)
                messagebox.showinfo("Success", f"{operation_type.name} operation completed successfully.")
            except Exception as e:
                messagebox.showerror("Error", str(e))
        self.root.quit()

    def run(self):
        self.create_ui()
        self.root.mainloop()
