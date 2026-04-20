# ============================================================
# drive_manager.py
# Handles Google Drive operations:
#   - Copy template to target folder
#   - Upload image to Drive (for inserting into Slides)
# ============================================================

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseUpload
import io
import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import CREDENTIALS_FILE, SCOPES, TARGET_DRIVE_FOLDER_ID


def get_drive_service():
    creds = service_account.Credentials.from_service_account_file(
        CREDENTIALS_FILE, scopes=SCOPES
    )
    return build("drive", "v3", credentials=creds)


def copy_template(template_id: str, new_title: str, folder_id: str = TARGET_DRIVE_FOLDER_ID) -> str:
    """
    Copy a Google Slides template to the target folder.
    Returns the new presentation ID.
    """
    service = get_drive_service()
    body = {
        "name": new_title,
        "parents": [folder_id]
    }
    result = service.files().copy(fileId=template_id, body=body).execute()
    new_id = result["id"]

    # Make it accessible to anyone with the link (view)
    service.permissions().create(
        fileId=new_id,
        body={"type": "anyone", "role": "reader"}
    ).execute()

    return new_id


def upload_image_to_drive(image_bytes: bytes, filename: str, folder_id: str = TARGET_DRIVE_FOLDER_ID) -> str:
    """
    Upload an image (as bytes) to Google Drive.
    Returns a publicly accessible URL for use in Slides API.
    """
    service = get_drive_service()

    file_metadata = {
        "name": filename,
        "parents": [folder_id],
        "mimeType": "image/png"
    }
    media = MediaIoBaseUpload(
        io.BytesIO(image_bytes),
        mimetype="image/png",
        resumable=False
    )
    file = service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id,webContentLink,webViewLink"
    ).execute()

    file_id = file["id"]

    # Make public so Slides API can access it
    service.permissions().create(
        fileId=file_id,
        body={"type": "anyone", "role": "reader"}
    ).execute()

    # Return the direct download URL
    url = f"https://drive.google.com/uc?export=download&id={file_id}"
    return url, file_id


def delete_drive_file(file_id: str):
    """Clean up temporary image files from Drive."""
    service = get_drive_service()
    service.files().delete(fileId=file_id).execute()
