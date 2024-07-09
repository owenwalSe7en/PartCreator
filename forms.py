import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, filedialog, messagebox
from combobox_options import (TYPE_OPTIONS, CLASS_OPTIONS, REPORTING_GROUP_OPTIONS,
                              ON_HOLD_REASON_OPTIONS, GROUP_OPTIONS, LABEL_GROUP_OPTIONS)


# TODO: Validate each input escpecially the sheet name parameter. Be sure to convert it to the sheet index in the dict

class BaseForm:
    def __init__(self, master):
        self.master = master
        self.master.title("Create Part")
        self.file_data = {}
        self.label_data = {}
        self.widgets = []
        self.small_font = tkfont.Font(size=12)

    def create_entry_widget(self, frame, label, row, col, var_type=tk.StringVar):

        label_widget = ttk.Label(frame, text=label, width=22, anchor='w', font=self.small_font)
        label_widget.grid(row=row, column=col, padx=(0, 10), pady=7, sticky='w')

        var = var_type()
        entry = ttk.Entry(frame, textvariable=var)
        entry.grid(row=row, column=col + 1, padx=(0, 10), pady=7, sticky='ew')

        self.widgets.append((label, var))

    def create_file_widget(self, frame, label, row, col, var_type=tk.StringVar):

        label_widget = ttk.Label(frame, text=label, width=20, anchor='w', font=self.small_font)
        label_widget.grid(row=row, column=col, padx=(0, 10), pady=7, sticky='w')

        var = var_type()
        entry = ttk.Entry(frame, textvariable=var)
        entry.grid(row=row, column=col + 1, padx=(0, 10), pady=7, sticky='ew')

        browse_button = tk.Button(frame, text="Browse", command=lambda: self.browse_file(var))
        browse_button.grid(row=row, column=col + 2, padx=(0, 10), pady=7, sticky='ew')

        self.widgets.append((label, var))

    def create_dropdown_widget(self, frame, label, width, options, row, col, var_type=tk.StringVar):
        label_widget = ttk.Label(frame, text=label, width=20, anchor='w', font=self.small_font)
        label_widget.grid(row=row, column=col, padx=(0, 10), pady=7, sticky='w')

        var = var_type()
        combobox = ttk.Combobox(frame, textvariable=var, values=options, state="readonly")
        combobox.grid(row=row, column=col + 1, padx=(0, 10), pady=7, sticky='w')
        combobox.config(width=width)

        self.widgets.append((label, var))

    def create_checkbox_widget(self, frame, label, row, col, var_type=tk.BooleanVar):
        var = var_type()  # Create a BooleanVar instance that updates when the checkbox state is changed

        checkbox = ttk.Checkbutton(frame, text=label, variable=var)
        checkbox.grid(row=row, column=col, padx=(0, 10), pady=7, sticky='w')

        self.widgets.append((label, var))

    def browse_file(self, var):
        file_path = filedialog.askopenfilename()
        var.set(file_path)

    def create_file_form(self):
        # To be implemented by subclasses
        pass

    def create_label_form(self):
        # To be implemented by subclasses
        pass

    def submit_file_data(self, target_dict, frame):
        for label, var in self.widgets:
            if var.get().strip() == "":
                messagebox.showerror("Error", "There are missing fields in the current form")
                return

            target_dict[label] = var.get()
        print(self.file_data)

        # Verify the current subclass isn't DeleteForm
        class_name = type(self).__name__
        if class_name != "DeleteForm":
            self.create_label_form()
        else:
            # Close the current form
            self.master.destroy()


    def submit_label_data(self, target_dict, frame):
        for label, var in self.widgets:
            if var.get().strip() == "":
                messagebox.showerror("Error", "There are missing fields in the current form")
                return
            target_dict[label] = var.get()
        print(self.label_data)
        frame.destroy()


class CreateForm(BaseForm):

    def create_file_form(self):
        self.master.minsize(420, 240)
        self.first_frame = ttk.Frame(self.master, padding="10")
        self.first_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.create_file_widget(self.first_frame, "Input File", 0, 0)
        self.create_entry_widget(self.first_frame, "Sheet Index", 1, 0)
        self.create_entry_widget(self.first_frame, "Part Column Letter", 2, 0)
        self.create_entry_widget(self.first_frame, "Description Column Letter", 3, 0)
        self.create_entry_widget(self.first_frame, "First Row", 4, 0)
        self.create_entry_widget(self.first_frame, "Last Row", 5, 0)
        tk.Button(self.first_frame, text="Submit", command=lambda: self.submit_file_data(self.file_data,
                                                    self.first_frame)).grid(row=5, column=2, padx=(0, 10), pady=7)

    def create_label_form(self):
        self.second_frame = ttk.Frame(self.master, padding="10")
        self.second_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.create_dropdown_widget(self.second_frame, "Type", 13, TYPE_OPTIONS, 0, 0)
        self.create_dropdown_widget(self.second_frame, "On Hold Reason", 28, ON_HOLD_REASON_OPTIONS,
                                    5, 0)
        self.create_dropdown_widget(self.second_frame, "Group", 28, GROUP_OPTIONS, 1, 0)
        self.create_dropdown_widget(self.second_frame, "Class", 28, CLASS_OPTIONS, 2, 0)
        self.create_dropdown_widget(self.second_frame, "Label Group", 28, LABEL_GROUP_OPTIONS,
                                    3, 0)
        self.create_dropdown_widget(self.second_frame, "Reporting Group", 28, REPORTING_GROUP_OPTIONS,
                                    4, 0)

        self.create_checkbox_widget(self.second_frame, "Priced Part", 0, 2)
        self.create_checkbox_widget(self.second_frame, "Salesforce Sync", 1, 2)
        self.create_checkbox_widget(self.second_frame, "Catalog Part", 2, 2)
        self.create_checkbox_widget(self.second_frame, "Kit Catalog", 3, 2)

        tk.Button(self.second_frame, text="Submit", command=lambda: self.submit_label_data(self.file_data,
                                                    self.second_frame)).grid(row=5, column=2, padx=(0, 10), pady=7,)


class OverwriteForm(BaseForm):
    def create_file_form(self):
        self.master.minsize(410, 200)
        self.first_frame = ttk.Frame(self.master, padding="10")
        self.first_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.create_file_widget(self.first_frame, "Input File", 0, 0)
        self.create_entry_widget(self.first_frame, "Sheet Index", 1, 0)
        self.create_entry_widget(self.first_frame, "Part Column Letter", 2, 0)
        self.create_entry_widget(self.first_frame, "First Row", 3, 0)
        self.create_entry_widget(self.first_frame, "Last Row", 4, 0)
        tk.Button(self.first_frame, text="Submit", command=lambda: self.submit_file_data(self.file_data,
                                                    self.first_frame)).grid(row=4, column=2, padx=(0, 10), pady=7)

    def create_label_form(self):
        self.second_frame = ttk.Frame(self.master, padding="10")
        self.second_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.create_dropdown_widget(self.second_frame, "Type", 13, TYPE_OPTIONS, 0, 0)
        self.create_dropdown_widget(self.second_frame, "On Hold Reason", 28, ON_HOLD_REASON_OPTIONS,
                                    5, 0)
        self.create_dropdown_widget(self.second_frame, "Group", 28, GROUP_OPTIONS, 1, 0)
        self.create_dropdown_widget(self.second_frame, "Class", 28, CLASS_OPTIONS, 2, 0)
        self.create_dropdown_widget(self.second_frame, "Label Group", 28, LABEL_GROUP_OPTIONS,
                                    3, 0)
        self.create_dropdown_widget(self.second_frame, "Reporting Group", 28, REPORTING_GROUP_OPTIONS,
                                    4, 0)

        self.create_checkbox_widget(self.second_frame, "Priced Part", 0, 2)
        self.create_checkbox_widget(self.second_frame, "Salesforce Sync", 1, 2)
        self.create_checkbox_widget(self.second_frame, "Catalog Part", 2, 2)
        self.create_checkbox_widget(self.second_frame, "Kit Catalog", 3, 2)

        tk.Button(self.second_frame, text="Submit", command=lambda: self.submit_label_data(self.file_data,
                                                    self.second_frame)).grid(row=5, column=2, padx=(0, 10), pady=7)



class DeleteForm(BaseForm):
    def create_file_form(self):
        self.master.minsize(410, 200)
        self.first_frame = ttk.Frame(self.master, padding="10")
        self.first_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.create_file_widget(self.first_frame, "Input File", 0, 0)
        self.create_entry_widget(self.first_frame, "Sheet Index", 1, 0)
        self.create_entry_widget(self.first_frame, "Part Column Letter", 2, 0)
        self.create_entry_widget(self.first_frame, "First Row", 3, 0)
        self.create_entry_widget(self.first_frame, "Last Row", 4, 0)
        tk.Button(self.first_frame, text="Submit", command=lambda: self.submit_file_data(self.file_data,
                                                    self.first_frame)).grid(row=4, column=2, padx=(0, 10), pady=7)


