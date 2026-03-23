#!/usr/bin/env -S uv run
# /// script
# requires-python = ">=3.11"
# dependencies = [
#   "google-api-python-client",
#   "google-auth",
# ]
# ///
"""Upload a .pptx file to Google Drive (supports Shared Drives).

Usage:
  uv run upload_to_drive.py --file report.pptx --title "Weekly Report 2026-W12" --folder-id FOLDER_ID
"""

import argparse
import os
import sys

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

SCOPES = ["https://www.googleapis.com/auth/drive"]

MIME_PPTX = "application/vnd.openxmlformats-officedocument.presentationml.presentation"


def _get_drive():
    creds_path = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS")
    if not creds_path:
        print("Error: GOOGLE_APPLICATION_CREDENTIALS not set.", file=sys.stderr)
        sys.exit(1)
    creds = service_account.Credentials.from_service_account_file(creds_path, scopes=SCOPES)
    return build("drive", "v3", credentials=creds)


def cmd_upload(file_path: str, title: str, folder_id: str) -> str:
    drive = _get_drive()

    media = MediaFileUpload(file_path, mimetype=MIME_PPTX, resumable=False)
    file = drive.files().create(
        body={
            "name": title,
            "mimeType": MIME_PPTX,
            "parents": [folder_id],
        },
        media_body=media,
        supportsAllDrives=True,
        fields="id",
    ).execute()

    file_id = file["id"]
    return f"https://drive.google.com/file/d/{file_id}/view"


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--file", required=True, help="Path to .pptx file")
    parser.add_argument("--title", required=True)
    parser.add_argument("--folder-id", required=True)
    args = parser.parse_args()

    try:
        url = cmd_upload(args.file, args.title, args.folder_id)
        print(url)
    except Exception as e:
        print(f"Upload failed: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
