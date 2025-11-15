# ---------------------------- import ------------------------------- #
from reworker import DataRework, A
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading


# ---------------------------- __init__ ------------------------------- #
class App:
    def __init__(self, app_root):
        self.root = app_root
        self.selected_file = None
        self.work = None
        self.progress_value = 0
        self.total_sheets = 0
        self.completed_sheets = 0
        self.root.title("Tesla Data Converter")
        self.root.config(width=400, height=400)

# ---------------------------- GUI ------------------------------- #

        self.entry_path = ttk.Entry(self.root,width=50)
        self.entry_path.grid(row=0, column=0,columnspan=2, padx=10, pady=20)

        self.upload = ttk.Button(self.root, text="Datei Upload",command= self.data)
        self.upload.grid(row=1, column=0)

        self.start_button = ttk.Button(text="Converter Start", command=self.start_thread)
        self.start_button.grid(row=1, column=1)

        self.progressbar = ttk.Progressbar(self.root, orient="horizontal", length=200, mode="determinate")
        self.progressbar.grid(row=2,column=0, columnspan=2, padx=10, pady=20)

        self.progress_label = ttk.Label()
        self.progress_label.grid(row=3, column=0, columnspan=2, padx=10, pady=20)

        self.analysis = ttk.Button(text="Start Analysis", command= lambda : A(self.selected_file))
        self.analysis.grid(row=4, column=0, columnspan=2, padx=10, pady=20)

# ---------------------------- Function ------------------------------- #

    def data(self):
        messagebox.showinfo(title="Upload-Info", message=
        "This tool adds the following categories based on the “Postleitzahl” (postal code) and “Stadt” (city) columns using data from the OpenPLZ API.\n\nThe column headers must be exactly “Postleitzahl” and “Stadt”, with the corresponding values listed in the rows below.\n\nIt supports multiple sheets, and the file must be uploaded in Excel format.")
        self.selected_file = filedialog.askopenfilename()
        if self.selected_file:
            self.entry_path.delete(0, "end")
            self.entry_path.insert(0, self.selected_file)
        self.entry_path.master.focus()

    def start_thread(self):
        self.start_button.config(state="disabled")
        self.progressbar["value"] = 0
        self.completed_sheets = 0
        self.progress_value = 0

        self.work = DataRework(self.selected_file)
        self.total_sheets = len(self.work.excel)

        background_thread = threading.Thread(target=self.run_converter)
        background_thread.start()
        self.update_progressbar()

    def run_converter(self):
        def report_progress(done, total, sheet_name):
            self.root.after(0, self.update_progress_values, done, total, sheet_name)
        self.work.check_plz_parallel(callback=report_progress)

    def update_progress_values(self, done, total, sheet_name):
        if not hasattr(self, "sheet_progress"):
            self.sheet_progress = {}
            self.total_sheets = len(self.work.excel)
            self.completed_sheets = 0

        self.sheet_progress[sheet_name] = done / total if total else 0

        overall_progress = sum(self.sheet_progress.values()) / self.total_sheets
        self.progressbar["value"] = overall_progress * 100

        if done == total and sheet_name not in getattr(self, "finished_sheets", set()):
            if not hasattr(self, "finished_sheets"):
                self.finished_sheets = set()
            self.finished_sheets.add(sheet_name)
            self.completed_sheets += 1

        if self.completed_sheets == self.total_sheets:
            self.start_button.config(state="normal")
            self.work.save_excel(self.selected_file)
            messagebox.showinfo(message="All sheets converted")

        self.progress_label.config(text=f"Sheets: {self.completed_sheets}/{self.total_sheets}\nData row: {done}/{total}")

    def update_progressbar(self):
        if self.progress_value < 100:
            self.root.after(200, self.update_progressbar)

# ---------------------------- Mainloop ------------------------------- #

root = tk.Tk()
App(root)
root.mainloop()

