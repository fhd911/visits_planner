from pathlib import Path

from google.oauth2 import service_account
from googleapiclient.discovery import build


BASE_DIR = Path(__file__).resolve().parents[2]
SERVICE_ACCOUNT_FILE = BASE_DIR / "secrets" / "service-account.json"
SCOPES = ["https://www.googleapis.com/auth/drive"]


def get_drive_service():
    credentials = service_account.Credentials.from_service_account_file(
        str(SERVICE_ACCOUNT_FILE),
        scopes=SCOPES,
    )
    return build("drive", "v3", credentials=credentials)


def list_files_in_folder(folder_id: str, page_size: int = 100):
    service = get_drive_service()
    results = service.files().list(
        q=f"'{folder_id}' in parents and trashed = false",
        pageSize=page_size,
        fields="files(id, name, mimeType, webViewLink)",
        orderBy="name",
    ).execute()
    return results.get("files", [])


def find_file_by_name(folder_id: str, filename: str):
    service = get_drive_service()
    results = service.files().list(
        q=f"'{folder_id}' in parents and name = '{filename}' and trashed = false",
        pageSize=10,
        fields="files(id, name, mimeType, webViewLink)",
    ).execute()
    files = results.get("files", [])
    return files[0] if files else None


def find_folder_by_name(parent_folder_id: str, folder_name: str):
    service = get_drive_service()
    results = service.files().list(
        q=(
            f"'{parent_folder_id}' in parents and "
            f"name = '{folder_name}' and "
            f"mimeType = 'application/vnd.google-apps.folder' and "
            f"trashed = false"
        ),
        pageSize=10,
        fields="files(id, name, mimeType, webViewLink)",
    ).execute()
    files = results.get("files", [])
    return files[0] if files else None