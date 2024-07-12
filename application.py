import tkinter as tk
from tkinter import ttk, messagebox
from erp_manager import OperationType
from forms import CreateForm, OverwriteForm, DeleteForm
import shutil


def print_fancy_separator(text="", char='-'):
    """
    This function generates a visually appealing separator in the terminal.

    :param text: The text to be centered within the separator. Default is an empty string.
    :param char: The character used to build the separator. Default is '-'.
    :return: None
    """

    terminal_width, _ = shutil.get_terminal_size()
    if text:
        text = f" {text} "
    separator_width = (terminal_width - len(text)) // 2
    print(f"{char * separator_width}{text}{char * separator_width}")


class Application:
    def __init__(self, erp_manager):
        """
        Initializes the Application class instance, an instance of the ERPManager, and the original tkinter root

        :param erp_manager: An instance of the ERPManager class
        """

        self.erp_manager = erp_manager
        self.root = tk.Tk()
        self.root.title("Operation Selection")
        self.root.geometry("650x170")
        self.root.minsize(650, 170)

    def create_ui(self):
        """
        Configures the operation selection form to have 3 buttons each of which correspond with a specific OperationType

        :return: None
        """
        style = ttk.Style()
        style.configure('TButton', font=('Arial', 14))
        style.configure('TLabel', font=('Arial', 18))

        # Create a new style for taller buttons
        style.configure('Tall.TButton', font=('Arial', 14), padding=(10, 20))

        ttk.Label(self.root, text="Select an operation:", style='TLabel').pack(pady=(20, 10))

        # Create button frame
        button_frame = ttk.Frame(self.root)
        button_frame.pack(pady=10)

        # Create buttons
        ttk.Button(button_frame, text="Create", command=lambda: self.open_form(CreateForm, OperationType.CREATE),
                   style='Tall.TButton', width=15).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="Overwrite", command=lambda: self.open_form(OverwriteForm,
                                                                                  OperationType.OVERWRITE),
                   style='Tall.TButton', width=15).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="Delete", command=lambda: self.open_form(DeleteForm, OperationType.DELETE),
                   style='Tall.TButton', width=15).pack(side=tk.LEFT, padx=10)

    def open_form(self, form_class, operation_type):
        """
        Creates the File Information form which leads to the collection of all user data in both File Information
        and Label Information. Verifies that the user data exist and then executes the specific looping method that
        cooresponds to a specific Operation subclass in erp_manager.

        :param form_class: The specific form of use (CreateForm, OverwriteForm, or DeleteForm)
        :param operation_type: The specific operation type (OperationType.Create, OperationType.OVERWRITE, or
        OperationType.DELETE)
        :raises: Any error that gets caught during the execute method
        """

        form_window = tk.Toplevel(self.root)
        form = form_class(form_window)
        form.create_file_form(operation_type)  # Create the File Information form
        self.root.withdraw()  # Hide the main window
        self.root.wait_window(form_window)

        # Check for existing file data and termination global variable
        if form.file_data and not form.is_terminated:
            try:
                self.erp_manager.perform_operation(operation_type, form.file_data, form.label_data)
                print_fancy_separator("Program Terminated")
                messagebox.showinfo("Success", f"{operation_type.name} operation completed successfully.")
            except Exception as e:
                messagebox.showerror("Error", str(e))
                raise e
        self.root.quit()

    def run(self):
        """
        Run the UI

        :return: None
        """
        self.create_ui()
        self.root.mainloop()
