import os
from pathlib import PurePath
import environ
import threading
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.files.file import File

env = environ.Env()
env.read_env()

SHAREPOINT_SITE_URL = env("SHAREPOINT_SITE_URL")
SHAREPOINT_SITE_NAME = env("SHAREPOINT_SITE_NAME")
SHAREPOINT_DOC_LIBRARY = env("SHAREPOINT_DOC_LIBRARY")


class Sharepoint:
    def __init__(self, email, password):
        self.lock = threading.Lock()  # For thread-safe folder creation
        self.email = email
        self.password = password

    def _auth(self):
        # Always return a new ClientContext to ensure thread safety
        return ClientContext(SHAREPOINT_SITE_URL).with_credentials(
            UserCredential(self.email, self.password)
        )

    def _get_files_list(self, folder_name):
        conn = self._auth()
        target_folder_url = f'{SHAREPOINT_DOC_LIBRARY}/{folder_name}'
        root_folder = conn.web.get_file_by_server_relative_url(target_folder_url)
        root_folder.expand(["Files", "Folders"]).get().execute_query()
        return root_folder.files

    def get_files_folders_list(self, folder_name):
        conn = self._auth()
        target_folder_url = f'{SHAREPOINT_DOC_LIBRARY}/{folder_name}'
        root_folder = conn.web.get_folder_by_server_relative_url(target_folder_url)
        files = root_folder.files
        conn.load(files).execute_query()
        root_folder.expand(["Folders"]).get().execute_query()
        return {
            "files": files,
            "folders": root_folder.folders
        }

    def download_file(self, file_name, folder_path):
        if not file_name:
            return {"error": "File name cannot be empty.", "downloaded_file_path": None}

        conn = self._auth()
        file_url = f'/sites/{SHAREPOINT_SITE_NAME}/{SHAREPOINT_DOC_LIBRARY}/{folder_path}/{file_name}'

        try:
            files_list = self.get_files_folders_list(folder_path)
        except Exception as e:
            return {"error": str(e), "downloaded_file_path": None}

        if not files_list['files']:
            return {"error": "No files found in the specified folder.", "downloaded_file_path": None}

        file_exists = any(f.properties['Name'] == file_name for f in files_list['files'])
        if not file_exists:
            return {"error": f"File '{file_name}' not found in folder '{folder_path}'.", "downloaded_file_path": None}

        try:
            file = File.open_binary(conn, file_url)
        except Exception as e:
            return {"error": f"Download error: {e}", "downloaded_file_path": None}

        # Build path safely with threading lock
        folder_hierarchy = 'api/local_directory'
        path_parts = folder_path.split('/') if folder_path else []

        with self.lock:
            for folder in path_parts:
                folder_hierarchy = os.path.join(folder_hierarchy, folder)
                if not os.path.exists(folder_hierarchy):
                    os.mkdir(folder_hierarchy)

        file_dir_path = PurePath(folder_hierarchy, file_name)
        try:
            with open(file_dir_path, 'wb') as f:
                f.write(file.content)
        except Exception as e:
            return {"error": f"File write error: {e}", "downloaded_file_path": None}

        return {"error": None, "downloaded_file_path": str(file_dir_path)}

    def upload_file(self, file_name, folder_path, content):
        conn = self._auth()
        target_folder_url = f'/sites/{SHAREPOINT_SITE_NAME}/{SHAREPOINT_DOC_LIBRARY}/{folder_path}'
        target_folder = conn.web.get_folder_by_server_relative_path(target_folder_url)
        response = target_folder.upload_file(file_name, content).execute_query()
        return {"error": None, "response": response}

    def create_folder(self, folder_name):
        conn = self._auth()
        target_folder_url = f'{SHAREPOINT_DOC_LIBRARY}/{folder_name}'
        try:
            folder = conn.web.folders.add(target_folder_url).execute_query()
            return {"error": None, "folder": folder}
        except Exception as e:
            return {"error": str(e), "folder": None}

    def check_if_folder_exists(self, folder_name):
        conn = self._auth()
        target_folder_url = f'{SHAREPOINT_DOC_LIBRARY}/{folder_name}'
        try:
            folder = conn.web.get_folder_by_server_relative_url(target_folder_url)
            conn.load(folder).execute_query()
            return {"exists": True, "error": None}
        except Exception as e:
            return {"exists": False, "error": str(e)}

