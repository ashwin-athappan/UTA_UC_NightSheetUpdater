from kivymd.uix.dialog import MDDialog
from kivymd.uix.list import OneLineIconListItem, IconLeftWidget, MDList
from kivymd.uix.button import MDRaisedButton
from kivymd.uix.snackbar import MDSnackbar
from kivymd.uix.pickers import MDDatePicker
from kivymd.uix.scrollview import MDScrollView
from kivymd.uix.boxlayout import MDBoxLayout
from kivymd.uix.label import MDLabel
from kivy.uix.screenmanager import Screen
from kivy.clock import Clock
from kivy.properties import StringProperty
from kivy.app import App
from datetime import datetime
from functools import partial

from api.night_sheet_updater import run_on_sharepoint_file

DOUBLE_CLICK_DELAY = 0.4  # seconds


class DashboardScreen(Screen):
    night_sheet_path = StringProperty("")
    turnover_sheet_path = StringProperty("")
    start_date_value = StringProperty("")
    end_date_value = StringProperty("")

    current_path = ""
    path_history = []
    selected_file = None
    dialog = None
    file_type = None
    folder_cache = {}  # Cache: { "path": { "files": [...], "folders": [...] } }

    def select_file(self, file_type):
        self.file_type = file_type
        self.sharepoint = App.get_running_app().sharepoint
        self._open_browser(path="")

    def _open_snackbar(self, message):
        snackbar_text = MDLabel(text=message)
        MDSnackbar(snackbar_text, snackbar_x="10dp", snackbar_y="10dp").open()

    def _open_browser(self, path, add_to_history=True):
        self.current_path = path
        self.selected_file = None
        self.last_click_time = {}

        # Container: NEW every time
        container = MDBoxLayout(orientation="vertical", size_hint_y=None)
        container.bind(minimum_height=container.setter("height"))
        list_container = MDList()
        list_container.size_hint_y = None
        list_container.bind(minimum_height=list_container.setter("height"))

        # Breadcrumb
        breadcrumb = MDLabel(
            text=f"📁 /{self.current_path}" if self.current_path else "📁 /",
            theme_text_color="Secondary",
            halign="left",
            size_hint_y=None,
            height=40,
            padding=(10, 10),
        )
        container.add_widget(breadcrumb)

        try:
            response = {}
            if path in self.folder_cache:
                response = self.folder_cache[path]
            else:
                # Fetch from SharePoint
                response = self.sharepoint.get_files_folders_list(path)
                self.folder_cache[path] = response

            for folder in response['folders']:
                name = folder.properties['Name']

                def on_folder_press(item, folder_name):
                    now = Clock.get_boottime()
                    last = self.last_click_time.get(folder_name, 0)

                    if now - last < DOUBLE_CLICK_DELAY:
                        self._navigate_to(folder_name)
                    else:
                        self.last_click_time[folder_name] = now

                item = OneLineIconListItem(
                    text=name,
                    on_release=lambda x, folder_name=name: on_folder_press(x, folder_name),
                )
                item.add_widget(IconLeftWidget(icon="folder"))
                list_container.add_widget(item)

            def on_file_select(file_name, list_item):
                now = Clock.get_boottime()
                last = self.last_click_time.get(file_name, 0)

                self.selected_file = file_name  # ✅ Must be set before confirm_selection

                if now - last < DOUBLE_CLICK_DELAY:
                    confirm_selection(None)
                else:
                    self.last_click_time[file_name] = now

            for file in response['files']:
                name = file.properties['Name']
                item = OneLineIconListItem(
                    text=name,
                    on_release=lambda x, file_name=name, list_item=item: on_file_select(file_name, list_item),
                )
                item.add_widget(IconLeftWidget(icon="file"))
                list_container.add_widget(item)

        except Exception as e:
            self._open_snackbar(message=f"Error loading: {e}")
            return

        scroll = MDScrollView(size_hint=(1, None), height=400)
        scroll.add_widget(list_container)
        container.add_widget(scroll)

        def go_root(_):
            self.path_history.clear()
            self._open_browser("")

        def go_back(_):
            if self.path_history:
                prev = self.path_history.pop()
                self._open_browser(prev, add_to_history=False)

        def confirm_selection(_):
            if not self.selected_file:
                self._open_snackbar("❗ Please select a file first.")
                return

            full_path = f"{self.current_path}/{self.selected_file}".strip("/")
            if self.file_type == "night":
                self.night_sheet_path = full_path
            else:
                self.turnover_sheet_path = full_path

            self._open_snackbar(message=f"Selected: {full_path}")
            self._close_dialog()

        self._close_dialog()  # just in case a dialog was open

        self.dialog = MDDialog(
            title=f"Browsing: /{path or 'Root'}",
            type="custom",
            content_cls=container,
            buttons=[
                MDRaisedButton(text="Root", on_release=go_root),
                MDRaisedButton(text="Back", on_release=go_back),
                MDRaisedButton(text="Confirm", on_release=confirm_selection),
                MDRaisedButton(text="Close", on_release=self._close_dialog),
            ],
        )
        self.dialog.open()

    def _navigate_to(self, folder_name, *args):
        self.path_history.append(self.current_path)
        new_path = f"{self.current_path}/{folder_name}".strip("/")
        self._close_dialog()
        self._open_browser(new_path)

    def _close_dialog(self, *args):
        if self.dialog:
            self.dialog.dismiss()
            self.dialog = None

    def show_date_picker(self, date_type):
        date_picker = MDDatePicker()
        date_picker.bind(on_save=lambda instance, value, date_range: self.set_date(date_type, value))
        date_picker.open()

    def set_date(self, date_type, value):
        formatted = value.strftime("%Y-%m-%d")
        if date_type == "start":
            self.start_date_value = formatted
        else:
            self.end_date_value = formatted

        if self.start_date_value and self.end_date_value:
            self.validate_date_range()

    def validate_date_range(self):
        try:
            start_dt = datetime.strptime(self.start_date_value, "%Y-%m-%d")
            end_dt = datetime.strptime(self.end_date_value, "%Y-%m-%d")
            if start_dt >= end_dt:
                self._open_snackbar(message="❌ Start date must be before end date.")
            else:
                self._open_snackbar(message="✅ Date range is valid.")
        except Exception:
            self._open_snackbar(message="⚠️ Invalid date format.")

    def run_script(self):
        try:
            if not self.start_date_value or not self.end_date_value:
                self._open_snackbar(message="Please select both dates.")
                return

            start_dt = datetime.strptime(self.start_date_value, "%Y-%m-%d")
            end_dt = datetime.strptime(self.end_date_value, "%Y-%m-%d")

            night_sheet_file_name = self.night_sheet_path.split("/")[-1]
            turnover_sheet_file_name = self.turnover_sheet_path.split("/")[-1]
            print("Night Sheet:", night_sheet_file_name)
            print("Turnovers Sheet:", turnover_sheet_file_name)
            result = run_on_sharepoint_file(self.sharepoint, start_dt, end_dt, self.current_path, night_sheet_file_name, turnover_sheet_file_name)
            print(result)
        except Exception as e:
            print("Error running script:", e)
            self._open_snackbar(message="⚠️ Error running script")
