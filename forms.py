import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, filedialog, messagebox
from combobox_options import (TYPE_OPTIONS, CLASS_OPTIONS, REPORTING_GROUP_OPTIONS,
                              ON_HOLD_REASON_OPTIONS, GROUP_OPTIONS, LABEL_GROUP_OPTIONS)
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter, exceptions
import openpyxl
import sys
import os
import re


# region Validation Methods

def sheet_exists(excel_file_path, sheet_name):
    try:
        workbook = openpyxl.load_workbook(excel_file_path)
        return sheet_name in workbook.sheetnames
    except FileNotFoundError:
        print(f"File not found: {excel_file_path}")
        return False
    except openpyxl.utils.exceptions.InvalidFileException:
        print(f"Invalid Excel file: {excel_file_path}")
        return False
    except Exception(BaseException):
        print("An error occured. Please try again.")


def get_sheet_index(excel_file_path, sheet_name):
    """
    Get the index of a sheet in an Excel file.

    Args:
    excel_file_path (str): The file path of the Excel file.
    sheet_name (str): The name of the sheet to find.

    Returns:
    int: The index of the sheet in the Excel file.
    """
    try:
        workbook = openpyxl.load_workbook(excel_file_path)
        sheet_index = workbook.sheetnames.index(sheet_name)
        return sheet_index
    except FileNotFoundError:
        print(f"File not found: {excel_file_path}")
        return False
    except openpyxl.utils.exceptions.InvalidFileException:
        print(f"Invalid Excel file: {excel_file_path}")
        return False
    except ValueError:
        print(f"Sheet '{sheet_name}' not found in the Excel file.")
        return False


def validate_file_location(file_path):
    # Check if the file path is not empty
    if not file_path:
        return False

    # Check if the file path exists
    if not os.path.exists(file_path):
        return False

    # Check if the file path points to a file (not a directory)
    if os.path.isdir(file_path):
        return False

    return True


def is_file_open(filepath):
    try:
        os.rename(filepath, filepath)  # Attempt to rename the file to itself
        return False  # File is not open
    except OSError as e:
        return True  # File is open


def is_valid_row_combo(first_row, last_row):
    if is_valid_integer(first_row) and is_valid_integer(last_row):
        if int(first_row) < int(last_row):
            return True
        else:
            return False
    else:
        return False


def is_valid_column(column):
    return re.match(r'^[A-Za-z]+$', column)


def is_valid_integer(var):
    return isinstance(var, str) and var.isdigit()


def browse_file(var):
    file_path = filedialog.askopenfilename()
    var.set(file_path)


# endregion


class BaseForm:
    def __init__(self, master):
        # tkinter variables and configuration
        self.master = master
        self.master.title("Create Part")
        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)

        # Misc variables
        self.file_data = {}
        self.label_data = {}
        self.file_widgets = []
        self.label_widgets = []
        self.small_font = tkfont.Font(size=12)
        self.is_terminated = False

    # region Widget Creation
    def create_entry_widget(self, frame, label, row, col, arr, var_type=tk.StringVar):

        label_widget = ttk.Label(frame, text=label, width=22, anchor='w', font=self.small_font)
        label_widget.grid(row=row, column=col, padx=(0, 10), pady=7, sticky='w')

        var = var_type()
        entry = ttk.Entry(frame, textvariable=var)
        entry.grid(row=row, column=col + 1, padx=(0, 10), pady=7, sticky='ew')

        arr.append((label, var))

    def create_file_widget(self, frame, label, row, col, arr, var_type=tk.StringVar):

        label_widget = ttk.Label(frame, text=label, width=20, anchor='w', font=self.small_font)
        label_widget.grid(row=row, column=col, padx=(0, 10), pady=7, sticky='w')

        var = var_type()
        entry = ttk.Entry(frame, textvariable=var)
        entry.grid(row=row, column=col + 1, padx=(0, 10), pady=7, sticky='ew')

        browse_button = tk.Button(frame, text="Browse", command=lambda: browse_file(var))
        browse_button.grid(row=row, column=col + 2, padx=(0, 10), pady=7, sticky='ew')

        arr.append((label, var))

    def create_dropdown_widget(self, frame, label, width, options, row, col, arr, var_type=tk.StringVar):
        label_widget = ttk.Label(frame, text=label, width=20, anchor='w', font=self.small_font)
        label_widget.grid(row=row, column=col, padx=(0, 10), pady=7, sticky='w')

        var = var_type()
        combobox = ttk.Combobox(frame, textvariable=var, values=options, state="readonly")
        combobox.grid(row=row, column=col + 1, padx=(0, 10), pady=7, sticky='w')
        combobox.config(width=width)

        arr.append((label, var))

    def create_checkbox_widget(self, frame, label, row, col, arr, var_type=tk.BooleanVar):
        var = var_type()  # Create a BooleanVar instance that updates when the checkbox state is changed

        checkbox = ttk.Checkbutton(frame, text=label, variable=var)
        checkbox.grid(row=row, column=col, padx=(0, 10), pady=7, sticky='w')

        arr.append((label, var))

    # endregion

    def on_closing(self):
        if messagebox.askokcancel("Quit", "Do you want to quit the program?"):
            self.is_terminated = True
            self.master.destroy()
            sys.exit()

    def create_file_form(self, operation_type):
        # To be implemented by subclasses
        pass

    def create_label_form(self, operation_type):
        # To be implemented by subclasses
        pass

    def submit_file_data(self, target_dict, operation_type):
        for label, var in self.file_widgets:
            if var.get().strip() == "":
                messagebox.showerror("Error", "There are missing fields in the current form")
                return
            target_dict[label] = var.get()

        # Validate the user-inputted Excel file
        if not validate_file_location(target_dict["Input File"]):
            messagebox.showerror("Error", "Invalid file input")
            return

        # Validate that the file is not open
        if is_file_open(target_dict["Input File"]):
            messagebox.showerror("Error", "Excel file is currently open. Please close it and try again")
            return

        # Validate that the sheet is within the user-inputted Excel file
        if sheet_exists(target_dict["Input File"], target_dict["Sheet Name"]):
            sheet_index = get_sheet_index(target_dict["Input File"], target_dict["Sheet Name"])

            # Replace 'Sheet Name' with 'Sheet Index' for simplicity
            target_dict["Sheet Index"] = target_dict.pop("Sheet Name")
            target_dict["Sheet Index"] = sheet_index
        else:
            messagebox.showerror("Error", "Invalid sheet name")
            return

        # Validate Column Letters
        if not is_valid_column(target_dict["Part Column Letter"]):
            messagebox.showerror("Error", "Invalid part column letter")
            return
        if "Description Column Letter" in target_dict:
            if target_dict["Description Column Letter"] != target_dict["Part Column Letter"]:
                if not is_valid_column(target_dict["Description Column Letter"]):
                    messagebox.showerror("Error", "Invalid description column letter")
                    return
            else:
                messagebox.showerror("Error", "Invalid description column letter")
                return

        # Validate row order
        if not is_valid_row_combo(target_dict["First Row"], target_dict["Last Row"]):
            messagebox.showerror("Error", "Invalid row or row combination")
            return

        # Verify the current subclass isn't DeleteForm
        class_name = type(self).__name__
        if class_name != "DeleteForm":
            self.create_label_form(operation_type)
        else:
            # Close the current form
            self.master.destroy()

    def submit_label_data(self, target_dict, operation_type):
        # Variable that checks for the CREATE operation type
        # This type should not have any empty spaces
        is_create_operation = False
        if operation_type.name == "CREATE":
            is_create_operation = True

        # Variable that counts the amount of total empty fields
        empty_fields = 0
        # Variable that counts the amount of empty dropdown fields
        empty_dropdown_fields = 0

        for label, var in self.label_widgets:
            if var.get() == "" or var.get() == False:
                empty_fields += 1

            if var.get() == "":
                empty_dropdown_fields += 1

            target_dict[label] = var.get()

        # Check for empty dropdown fields in case of user creating
        if is_create_operation and empty_dropdown_fields > 0:
            messagebox.showerror("Input Error", "When you are creating parts you must fill in all dropdown"
                                                " fields")
            return

        # In the case of user overwriting with no inputs, ask for confirmation
        if empty_fields == len(target_dict):
            if not messagebox.askyesno("Warning", "You haven't made any changes. "
                                                  "Are you sure you want to proceed?"):
                return

        self.master.destroy()


class CreateForm(BaseForm):

    # Build custom File Form
    def create_file_form(self, operation_type):
        self.master.minsize(420, 240)
        self.first_frame = ttk.Frame(self.master, padding="10")
        self.first_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.create_file_widget(self.first_frame, "Input File", 0, 0, self.file_widgets)
        self.create_entry_widget(self.first_frame, "Sheet Name", 1, 0, self.file_widgets)
        self.create_entry_widget(self.first_frame, "Part Column Letter", 2, 0, self.file_widgets)
        self.create_entry_widget(self.first_frame, "Description Column Letter", 3, 0, self.file_widgets)
        self.create_entry_widget(self.first_frame, "First Row", 4, 0, self.file_widgets)
        self.create_entry_widget(self.first_frame, "Last Row", 5, 0, self.file_widgets)
        tk.Button(self.first_frame, text="Submit",
                  command=lambda: self.submit_file_data(self.file_data, operation_type)).grid(row=5,
                                                                                              column=2,
                                                                                              padx=(
                                                                                                  0,
                                                                                                  10),
                                                                                              pady=7)

    # Build custom Label Form
    def create_label_form(self, operation_type):
        self.second_frame = ttk.Frame(self.master, padding="10")
        self.second_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.create_dropdown_widget(self.second_frame, "Type", 13, TYPE_OPTIONS,
                                    0, 0, self.label_widgets)
        self.create_dropdown_widget(self.second_frame, "On Hold Reason", 28, ON_HOLD_REASON_OPTIONS,
                                    5, 0, self.label_widgets)
        self.create_dropdown_widget(self.second_frame, "Group", 28, GROUP_OPTIONS,
                                    1, 0, self.label_widgets)
        self.create_dropdown_widget(self.second_frame, "Class", 28, CLASS_OPTIONS,
                                    2, 0, self.label_widgets)
        self.create_dropdown_widget(self.second_frame, "Label Group", 28, LABEL_GROUP_OPTIONS,
                                    3, 0, self.label_widgets)
        self.create_dropdown_widget(self.second_frame, "Reporting Group", 28, REPORTING_GROUP_OPTIONS,
                                    4, 0, self.label_widgets)

        self.create_checkbox_widget(self.second_frame, "Priced Part", 0, 2, self.label_widgets)
        self.create_checkbox_widget(self.second_frame, "Salesforce Sync", 1, 2, self.label_widgets)
        self.create_checkbox_widget(self.second_frame, "Catalog Part", 2, 2, self.label_widgets)
        self.create_checkbox_widget(self.second_frame, "Kit Catalog", 3, 2, self.label_widgets)

        tk.Button(self.second_frame, text="Submit",
                  command=lambda: self.submit_label_data(self.label_data, operation_type)).grid(row=5, column=2, padx=(0
                                                                                        , 10), pady=7, )


class OverwriteForm(BaseForm):

    # Build custom File Form
    def create_file_form(self, operation_type):
        self.master.minsize(410, 200)
        self.first_frame = ttk.Frame(self.master, padding="10")
        self.first_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.create_file_widget(self.first_frame, "Input File", 0, 0, self.file_widgets)
        self.create_entry_widget(self.first_frame, "Sheet Name", 1, 0, self.file_widgets)
        self.create_entry_widget(self.first_frame, "Part Column Letter", 2, 0, self.file_widgets)
        self.create_entry_widget(self.first_frame, "First Row", 3, 0, self.file_widgets)
        self.create_entry_widget(self.first_frame, "Last Row", 4, 0, self.file_widgets)
        tk.Button(self.first_frame, text="Submit",
                  command=lambda: self.submit_file_data(self.file_data, operation_type)).grid(row=4,
                                                                                              column=2,
                                                                                              padx=(
                                                                                                  0,
                                                                                                  10)
                                                                                              , pady=7)

    # Build custom Label Form
    def create_label_form(self, operation_type):
        self.second_frame = ttk.Frame(self.master, padding="10")
        self.second_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.create_dropdown_widget(self.second_frame, "Type", 13, TYPE_OPTIONS,
                                    0, 0, self.label_widgets)
        self.create_dropdown_widget(self.second_frame, "On Hold Reason", 28, ON_HOLD_REASON_OPTIONS,
                                    5, 0, self.label_widgets)
        self.create_dropdown_widget(self.second_frame, "Group", 28, GROUP_OPTIONS,
                                    1, 0, self.label_widgets)
        self.create_dropdown_widget(self.second_frame, "Class", 28, CLASS_OPTIONS,
                                    2, 0, self.label_widgets)
        self.create_dropdown_widget(self.second_frame, "Label Group", 28, LABEL_GROUP_OPTIONS,
                                    3, 0, self.label_widgets)
        self.create_dropdown_widget(self.second_frame, "Reporting Group", 28, REPORTING_GROUP_OPTIONS,
                                    4, 0, self.label_widgets)

        self.create_checkbox_widget(self.second_frame, "Priced Part", 0, 2, self.label_widgets)
        self.create_checkbox_widget(self.second_frame, "Salesforce Sync", 1, 2, self.label_widgets)
        self.create_checkbox_widget(self.second_frame, "Catalog Part", 2, 2, self.label_widgets)
        self.create_checkbox_widget(self.second_frame, "Kit Catalog", 3, 2, self.label_widgets)

        tk.Button(self.second_frame, text="Submit", command=lambda: self.submit_label_data(self.label_data,
                                                                                           operation_type)).grid(row=5,
                                                                                                                 column=2,
                                                                                                                 padx=(
                                                                                                                     0,
                                                                                                                     10),
                                                                                                                 pady=7)


class DeleteForm(BaseForm):

    # Build custom File Form
    def create_file_form(self, operation_type):
        self.master.minsize(410, 200)
        self.first_frame = ttk.Frame(self.master, padding="10")
        self.first_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.create_file_widget(self.first_frame, "Input File", 0, 0, self.file_widgets)
        self.create_entry_widget(self.first_frame, "Sheet Name", 1, 0, self.file_widgets)
        self.create_entry_widget(self.first_frame, "Part Column Letter", 2, 0, self.file_widgets)
        self.create_entry_widget(self.first_frame, "First Row", 3, 0, self.file_widgets)
        self.create_entry_widget(self.first_frame, "Last Row", 4, 0, self.file_widgets)
        tk.Button(self.first_frame, text="Submit", command=lambda: self.submit_file_data(self.file_data, operation_type)
                  ).grid(row=4, column=2, padx=(0, 10), pady=7)
