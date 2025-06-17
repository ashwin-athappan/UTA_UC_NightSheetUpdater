import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import Calendar
from datetime import datetime
from office365_api import Sharepoint
from night_sheet_updater import run_on_sharepoint_file

DEFAULT_FILE_PATH = "Apps/Mazevo/Night Sheet - Multi Day Test.xlsx"

class DateRangeDialog(tk.Toplevel):
    def __init__(self, parent, start_date=None, end_date=None):
        super().__init__(parent)
        self.title("Select Date Range")
        self.grab_set()  # modal

        self.start_date = None
        self.end_date = None

        ttk.Label(self, text="Start Date:").grid(row=0, column=0, padx=10, pady=5)
        self.start_cal = Calendar(self, selectmode='day')
        self.start_cal.grid(row=1, column=0, padx=10)

        ttk.Label(self, text="End Date:").grid(row=0, column=1, padx=10, pady=5)
        self.end_cal = Calendar(self, selectmode='day')
        self.end_cal.grid(row=1, column=1, padx=10)

        btn_frame = ttk.Frame(self)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=10)

        ttk.Button(btn_frame, text="OK", command=self.confirm).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Cancel", command=self.destroy).pack(side="right", padx=5)

        if start_date:
            self.start_cal.set_date(start_date)
        if end_date:
            self.end_cal.set_date(end_date)

    def confirm(self):
        start = self.start_cal.get_date()
        end = self.end_cal.get_date()
        if datetime.strptime(start, "%m/%d/%y") > datetime.strptime(end, "%m/%d/%y"):
            messagebox.showerror("Invalid Range", "Start date cannot be after end date.")
            return
        self.start_date = start
        self.end_date = end
        self.destroy()


class SharepointBrowser(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Remote SharePoint File Browser + Script Runner")
        self.geometry("800x600")

        self.sharepoint = Sharepoint()
        self.current_path = ""
        self.history = []
        self.history_index = -1
        self.selected_file = None
        self.start_date_str = None
        self.end_date_str = None

        # Layout
        self.top_frame = ttk.Frame(self)
        self.top_frame.pack(fill="both", expand=True, padx=5, pady=5)

        self.bottom_frame = ttk.LabelFrame(self, text="Run Script on Selected File")
        self.bottom_frame.pack(fill="x", padx=5, pady=5)

        self.init_file_browser()
        self.init_script_controls()

        folder, filename = DEFAULT_FILE_PATH.rsplit("/", 1)
        self.default_folder = folder
        self.default_filename = filename
        self.navigate_to(folder)

    # store for later selection

    def init_file_browser(self):
        nav_frame = ttk.Frame(self.top_frame)
        nav_frame.pack(fill="x")

        self.back_btn = ttk.Button(nav_frame, text="⬅ Back", command=self.go_back)
        self.back_btn.pack(side="left")
        self.forward_btn = ttk.Button(nav_frame, text="➡ Forward", command=self.go_forward)
        self.forward_btn.pack(side="left")

        self.path_frame = ttk.Frame(self.top_frame)
        self.path_frame.pack(fill="x", pady=(5, 0))

        tree_frame = ttk.Frame(self.top_frame)
        tree_frame.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(tree_frame, columns=("Type",), show="headings")
        self.tree.heading("Type", text="Item")
        self.tree.column("Type", anchor="w")
        self.tree.pack(side="left", fill="both", expand=True)

        self.tree.bind("<Double-1>", self.on_item_double_click)
        self.tree.bind("<<TreeviewSelect>>", self.on_select_file)

        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

    def init_script_controls(self):
        self.bottom_frame.columnconfigure(1, weight=1)
        self.bottom_frame.columnconfigure(3, weight=1)

        ttk.Label(self.bottom_frame, text="Selected File:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.selected_file_label = ttk.Label(self.bottom_frame, text="None", foreground="gray")
        self.selected_file_label.grid(row=0, column=1, columnspan=3, sticky="w", padx=5, pady=5)

        ttk.Label(self.bottom_frame, text="Date Range:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.date_range_label = ttk.Label(self.bottom_frame, text="Not selected", foreground="gray")
        self.date_range_label.grid(row=1, column=1, columnspan=2, sticky="w", padx=5, pady=5)
        self.select_range_btn = ttk.Button(self.bottom_frame, text="Select Range", command=self.open_date_range)
        self.select_range_btn.grid(row=1, column=3, padx=5, pady=5)

        self.run_btn = ttk.Button(self.bottom_frame, text="Run Script", command=self.run_script)
        self.run_btn.grid(row=2, column=0, columnspan=4, pady=15)

    def open_date_range(self):
        dialog = DateRangeDialog(self)
        self.wait_window(dialog)
        if dialog.start_date and dialog.end_date:
            self.start_date_str = dialog.start_date
            self.end_date_str = dialog.end_date
            self.date_range_label.config(
                text=f"{self.start_date_str} to {self.end_date_str}",
                foreground="black"
            )

    def update_breadcrumb(self):
        for widget in self.path_frame.winfo_children():
            widget.destroy()

        path_parts = self.current_path.split("/") if self.current_path else []

        def navigate_handler(folder_path):
            def handler():
                self.navigate_to(folder_path, add_to_history=True)

            return handler

        ttk.Button(self.path_frame, text="Root", command=navigate_handler("")).pack(side="left")
        if path_parts:
            ttk.Label(self.path_frame, text=" / ").pack(side="left")

        for i, part in enumerate(path_parts):
            accumulated_path = "/".join(path_parts[:i + 1])
            ttk.Button(self.path_frame, text=part, command=navigate_handler(accumulated_path)).pack(side="left")
            if i < len(path_parts) - 1:
                ttk.Label(self.path_frame, text=" / ").pack(side="left")

    def navigate_to(self, path, add_to_history=True):
        try:
            res = self.sharepoint.get_files_folders_list(path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load folder: {e}")
            return

        self.current_path = path
        self.update_breadcrumb()

        for item in self.tree.get_children():
            self.tree.delete(item)

        for folder in res['folders']:
            name = folder.properties['Name']
            self.tree.insert("", "end", iid=name, values=(f"📁 {name}",))

        for file in res['files']:
            name = file.properties['Name']
            iid = f"file::{name}"
            self.tree.insert("", "end", iid=iid, values=(f"📄 {name}",))

        if add_to_history:
            self.history = self.history[:self.history_index + 1]
            self.history.append(path)
            self.history_index += 1

        self.update_nav_buttons()

        # Auto-select the default file if we're in the default folder
        if hasattr(self, "default_filename") and path == self.default_folder:
            for item in self.tree.get_children():
                label = self.tree.item(item)["values"][0]
                if label.endswith(self.default_filename):
                    self.tree.selection_set(item)
                    self.tree.focus(item)
                    self.tree.see(item)
                    self.selected_file = self.default_filename
                    self.selected_file_label.config(text=self.selected_file, foreground="black")
                    del self.default_filename  # prevent re-selection on future navigations
                    break

    def on_item_double_click(self, event):
        selected = self.tree.focus()
        values = self.tree.item(selected, "values")
        if not values:
            return
        label = values[0]
        if label.startswith("📁"):
            folder_name = label[2:].strip()
            new_path = f"{self.current_path}/{folder_name}" if self.current_path else folder_name
            self.navigate_to(new_path)

    def on_select_file(self, event):
        selected = self.tree.focus()
        values = self.tree.item(selected, "values")
        if values and values[0].startswith("📄"):
            self.selected_file = values[0][2:].strip()
            self.selected_file_label.config(text=self.selected_file, foreground="black")
        else:
            self.selected_file = None
            self.selected_file_label.config(text="None", foreground="gray")

    def run_script(self):
        if not self.selected_file:
            messagebox.showwarning("No file", "Please select a file to run the script on.")
            return

        if not self.selected_file.lower().endswith(".xlsx"):
            messagebox.showerror("Invalid File", "Only .xlsx files are supported.")
            return

        if not self.start_date_str or not self.end_date_str:
            messagebox.showwarning("Date Range Required", "Please select a date range.")
            return

        try:
            folder_path = self.current_path
            file_name = self.selected_file
            start = datetime.strptime(self.start_date_str, "%m/%d/%y")
            end = datetime.strptime(self.end_date_str, "%m/%d/%y")

            result_msg = run_on_sharepoint_file(folder_path, file_name, start, end)
            messagebox.showinfo("Success", result_msg)

        except Exception as e:
            messagebox.showerror("Script Failed", str(e))

    def go_back(self):
        if self.history_index > 0:
            self.history_index -= 1
            self.navigate_to(self.history[self.history_index], add_to_history=False)

    def go_forward(self):
        if self.history_index < len(self.history) - 1:
            self.history_index += 1
            self.navigate_to(self.history[self.history_index], add_to_history=False)

    def update_nav_buttons(self):
        self.back_btn.config(state="normal" if self.history_index > 0 else "disabled")
        self.forward_btn.config(state="normal" if self.history_index < len(self.history) - 1 else "disabled")


if __name__ == "__main__":
    app = SharepointBrowser()
    app.mainloop()
