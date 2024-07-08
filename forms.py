import tkinter as tk
from tkinter import ttk, filedialog


class BaseForm:
    def __init__(self, master):
        self.master = master
        self.data = {}
        self.widgets = []

    def create_widget(self, label, var_type=tk.StringVar):
        frame = tk.Frame(self.master)
        frame.pack(pady=5)

        label = tk.Label(frame, text=label, width=20, anchor='w')
        label.pack(side=tk.LEFT, padx=5)

        var = var_type()
        entry = tk.Entry(frame, textvariable=var)
        entry.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        self.widgets.append((label['text'], var))

    def create_file_widget(self, label, var_type=tk.StringVar):
        frame = tk.Frame(self.master)
        frame.pack(pady=5)

        label = tk.Label(frame, text=label, width=20, anchor='w')
        label.pack(side=tk.LEFT, padx=5)

        var = var_type()
        entry = tk.Entry(frame, textvariable=var)
        entry.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        browse_button = tk.Button(frame, text="Browse", command=lambda: self.browse_file(var))
        browse_button.pack(side=tk.LEFT, padx=5)

        self.widgets.append((label['text'], var))

    def browse_file(self, var):
        file_path = filedialog.askopenfilename()
        var.set(file_path)

    def create_form(self):
        # To be implemented by subclasses
        pass

    def submit(self):
        for label, var in self.widgets:
            self.data[label] = var.get()
        self.master.destroy()


class CreateForm(BaseForm):

    def create_form(self):
        self.master.title("Create Part")
        self.master.geometry = "400x400"
        self.create_file_widget("Input File")
        self.create_widget("Sheet Index")
        self.create_widget("Part Column Letter")
        self.create_widget("Description Column Letter")
        self.create_widget("First Row")
        self.create_widget("Second Row")
        tk.Button(self.master, text="Submit", command=self.submit).pack(pady=10)


class OverwriteForm(BaseForm):
    def create_form(self):
        self.master.title("Overwrite Part")
        self.create_widget("Part Number")
        self.create_widget("New Description")
        self.create_widget("New Type")
        self.create_widget("New Group")
        tk.Button(self.master, text="Submit", command=self.submit).pack(pady=10)


class DeleteForm(BaseForm):
    def create_form(self):
        self.master.title("Delete Part")
        self.create_widget("Part Number")
        self.create_widget("Reason for Deletion")
        tk.Button(self.master, text="Submit", command=self.submit).pack(pady=10)

