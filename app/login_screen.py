import os
from kivymd.uix.screen import MDScreen
from kivymd.uix.dialog import MDDialog
from kivymd.uix.button import MDRaisedButton
from kivy.app import App
from kivy.properties import ObjectProperty

from api.office365_api import Sharepoint

REMEMBER_FILE = "remember_email.txt"

class LoginScreen(MDScreen):
    dialog = None
    checkbox = ObjectProperty(None)

    def on_pre_enter(self, *args):
        """Called automatically when the screen is about to be displayed."""
        if os.path.exists(REMEMBER_FILE):
            try:
                with open(REMEMBER_FILE, "r") as f:
                    saved_email = f.read().strip()
                    self.ids.email.text = saved_email
                    self.ids.remember_checkbox.active = True
            except Exception as e:
                print(f"Error reading saved email: {e}")

    def validate_credentials(self):
        email = self.ids.email.text
        password = self.ids.password.text

        try:
            sharepoint = Sharepoint(email, password)
            sharepoint.get_files_folders_list('')  # Validate credentials
            print('✅ Login Successful')
            App.get_running_app().sharepoint = sharepoint

            # Handle "Remember Me"
            if self.ids.remember_checkbox.active:
                with open(REMEMBER_FILE, "w") as f:
                    f.write(email)
            else:
                if os.path.exists(REMEMBER_FILE):
                    os.remove(REMEMBER_FILE)

            self.manager.current = "dashboard"

        except Exception as e:
            print(f"❌ Login error: {e}")
            self.show_invalid_credentials_dialog()

    def show_invalid_credentials_dialog(self):
        if not self.dialog:
            self.dialog = MDDialog(
                title="Login Failed",
                text="Invalid credentials.\nPlease try again.",
                buttons=[
                    MDRaisedButton(
                        text="OK",
                        on_release=lambda x: self.dialog.dismiss()
                    )
                ],
            )
        self.dialog.open()
