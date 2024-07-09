from enum import Enum
from abc import ABC, abstractmethod
from pywinauto import Application
from pywinauto.keyboard import send_keys
from tkinter import messagebox
import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter, exceptions
from openpyxl.styles import Font


class OperationType(Enum):
    # Shared dictionaries
    file_data = {}
    label_data = {}

    CREATE = 1
    OVERWRITE = 2
    DELETE = 3

    @classmethod
    def initialize_workbook(cls):
        cls.workbook = Workbook()
        cls.sheet = cls.workbook.active
        cls.sheet.title = "Operations Log"
        cls.sheet.append(["Operation", "Part Number", "Description", "Status"])


class Operation(ABC):
    @abstractmethod
    def execute(self, file_data, label_data):
        pass

    def log_operation(self, operation, part_number, description, status):
        OperationType.sheet.append([operation, part_number, description, status])
        OperationType.workbook.save("operations_log.xlsx")


class CreateOperation(Operation):
    def execute(self, file_data, label_data):
        # Send message to user
        print(f"File Data: {file_data}\nLabel Data: {label_data}")
        print("Initializing Part Creation...")

        try:
            # Connect the application to Label Maintenance and send confirmation message
            app = Application(backend="uia").connect(title="Part Maintenance")
            print('Connection to Part Maintenance achieved')

            # Clear current information
            app.window(title='Part Maintenance').child_window(title="Clear").click_input()

            # Load in the user-provided workbook and access the specific sheet
            book = load_workbook(file_data["Input File"])
            sheet_index = file_data["Sheet Index"]
            sheets = book.sheetnames
            active_sheet = book[sheets[sheet_index]]

            # Loop through all the part numbers
            for i in range(file_data["First Row"], file_data["Last Row"] + 1):
                # Reconnect to the form toe ensure it doesn't fall asleep
                app = Application(backend="uia").connect(title="Part Maintenance")
                main_window = app.window(title='Part Maintenance')

                # Find the cell that hosts the part numbers and descriptions
                pn_cell = file_data["Part Column Letter"] + str(i)
                part_number = active_sheet[pn_cell].value
                part_description = active_sheet[file_data["Description Column Letter"]].value

                # Type cell value into text box
                main_window.child_window(auto_id='tbPart').type_keys(part_number)
                send_keys("{TAB}")

                # Confirm that the part does not already exist
                if main_window.child_window(title="Add New Confirmation").exists():
                    main_window.child_window(auto_id='btnYes2').click_input()
                    self.log_operation("Create", str(part_number), part_description, "Completed")
                else:
                    # Write PN into Excel file
                    self.log_operation("Create", str(part_number), part_description,
                                       "Incomplete - Part already exists")
                    continue

                # Begin writing data into Epicor
                main_window.child_window(auto_id="tbPartDescription").type_keys(part_description, with_spaces=True)
                main_window.child_window(auto_id="cboTypeCode").type_keys(label_data["Type"], with_spaces=True)
                main_window.child_window(auto_id="cbProdCode").type_keys(label_data["Group"], with_spaces=True)
                main_window.child_window(auto_id="cbClass").type_keys(label_data["Class"], with_spaces=True)
                main_window.child_window(auto_id="ucbLabelGroup").type_keys(label_data["Label Group"])
                main_window.child_window(auto_id="cboReportGroup").type_keys(label_data["Reporting Group"])
                main_window.child_window(auto_id="cbOnHoldReasonCode").type_keys(label_data["On Hold Reason"])

                # Here, we use simple logic to determine whether a checkbox should be clicked
                # Either the box is checked in our form and unchecked in Epicor or it's unchecked in our form and
                # checked in Epicor
                if (label_data["Priced Part"] and main_window.child_window(auto_id="epiCheckBox1").
                        get_toggle_state() == 0):
                    main_window.child_window(auto_id="epiCheckBox1").click_input()
                elif (not label_data["Priced Part"] and main_window.child_window(auto_id="epiCheckBox1").
                        get_toggle_state() == 1):
                    main_window.child_window(auto_id="epiCheckBox1").click_input()

                if (label_data["Salesforce Sync"] and main_window.child_window(auto_id="epiCheckBox2").
                        get_toggle_state() == 0):
                    main_window.child_window(auto_id="epiCheckBox2").click_input()
                elif (not label_data["Salesforce Sync"] and main_window.child_window(auto_id="epiCheckBox2").
                        get_toggle_state() == 1):
                    main_window.child_window(auto_id="epiCheckBox2").click_input()

                if (label_data["Catalog Part"] and main_window.child_window(auto_id="chkCatalogPart").
                        get_toggle_state() == 0):
                    main_window.child_window(auto_id="chkCatalogPart").click_input()
                elif (not label_data["Catalog Part"] and main_window.child_window(auto_id="chkCatalogPart").
                        get_toggle_state() == 1):
                    main_window.child_window(auto_id="chkCatalogPart").click_input()

                if (label_data["Kit Catalog"] and main_window.child_window(auto_id="chkKitCatalog").
                        get_toggle_state() == 0):
                    main_window.child_window(auto_id="chkKitCatalog").click_input()
                elif (not label_data["Kit Catalog"] and main_window.child_window(auto_id="chkKitCatalog").
                        get_toggle_state() == 1):
                    main_window.child_window(auto_id="chkKitCatalog").click_input()

                # Save the form and check for any unexpected errors
                main_window.child_window(title="Save").click_input()
                if main_window.child_window(title="Error").exists():
                    messagebox.showerror(
                        "Error",
                        "If you are creating parts and not overwriting existing ones, you must add a "
                        "description in the first form of the program. "
                    )
                main_window.child_window(title="Clear").click_input()

        except Exception as e:
            print(e)
            raise e


class OverwriteOperation(Operation):
    def execute(self, file_data, label_data):
        # Send message to user
        print(f"File Data: {file_data}\nLabel Data: {label_data}")
        print("Initializing Part Overwriting...")


class DeleteOperation(Operation):
    def execute(self, file_data, label_data):
        # Send message to user
        print(f"File Data: {file_data}\nLabel Data: {label_data}")
        print("Initializing Part Deletion...")

        try:
            # Connect the application to Label Maintenance and send confirmation message
            app = Application(backend="uia").connect(title="Part Maintenance")
            print('Connection to Part Maintenance achieved')

            # Clear current information
            app.window(title='Part Maintenance').child_window(title="Clear").click_input()

            # Load in the user-provided workbook and access the specific sheet
            book = load_workbook(file_data["Input File"])
            sheet_index = file_data["Sheet Index"]
            sheets = book.sheetnames
            active_sheet = book[sheets[sheet_index]]

            # Loop through all the part numbers
            for i in range(file_data["First Row"], file_data["Last Row"] + 1):
                # Reconnect to the form toe ensure it doesn't fall asleep
                app = Application(backend="uia").connect(title="Part Maintenance")
                main_window = app.window(title='Part Maintenance')

                # Find the cell that hosts the part numbers and descriptions
                pn_cell = file_data["Part Column Letter"] + str(i)
                part_number = active_sheet[pn_cell].value
                part_description = active_sheet[file_data["Description Column Letter"]].value

                # Type cell value into text box
                main_window.child_window(auto_id='tbPart').type_keys(part_number)
                send_keys("{TAB}")

                # Confirm that the part already exist
                if main_window.child_window(title="Add New Confirmation").exists():
                    main_window.child_window(auto_id='btnNo2').click_input()
                    self.log_operation("Delete", part_number, part_description, "Incomplete - "
                                                                                "Part doesn't exist and therefore can't"
                                                                                "be deleted")
                else:
                    main_window.child_window(title="Delete").click_input()
                    if main_window.child_window(title="Delete Confirmation").exists():
                        main_window.child_window(auto_id='btnYes2').click_input()
                        self.log_operation("Delete", part_number, part_description, "Completed")

        except Exception as e:
            print(e)
            raise e


class ERPManager:
    def __init__(self, create_op: Operation, overwrite_op: Operation, delete_op: Operation):
        self.operations = {
            OperationType.CREATE: create_op,
            OperationType.OVERWRITE: overwrite_op,
            OperationType.DELETE: delete_op
        }

    def perform_operation(self, op_type: OperationType, form_data, label_data):
        operation = self.operations.get(op_type)
        if operation:
            operation.execute(form_data, label_data)
        else:
            raise ValueError("Invalid operation type")
