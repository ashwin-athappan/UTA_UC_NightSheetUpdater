from kivymd.app import MDApp
from kivy.uix.screenmanager import ScreenManager
from kivy.lang import Builder
from login_screen import LoginScreen
from dashboard_screen import DashboardScreen

class MainApp(MDApp):
    def build(self):
        # Load KV files
        Builder.load_file("login.kv")
        Builder.load_file("dashboard.kv")

        # Optional: Set theme
        self.theme_cls.theme_style = "Dark"  # or "Light"
        self.theme_cls.primary_palette = "Blue"

        # Setup screens
        sm = ScreenManager()
        sm.add_widget(LoginScreen(name="login"))
        sm.add_widget(DashboardScreen(name="dashboard"))
        return sm

if __name__ == "__main__":
    MainApp().run()
