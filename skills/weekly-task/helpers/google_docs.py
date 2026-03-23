#!/usr/bin/env -S uv run
# /// script
# requires-python = ">=3.11"
# dependencies = [
#   "google-api-python-client",
#   "google-auth",
# ]
# ///
"""Shared Google Docs helper for weekly-task and weekly-report skills.

Usage:
  uv run google_docs.py create --title "Title" --content draft.md --folder-id FOLDER_ID
  uv run google_docs.py read --url https://docs.google.com/document/d/...
"""

import argparse
import os
import sys

from google.oauth2 import service_account
from googleapiclient.discovery import build

SCOPES = [
    "https://www.googleapis.com/auth/documents",
    "https://www.googleapis.com/auth/drive",
]


def extract_doc_id(url: str) -> str:
    """Extract document ID from a Google Docs URL."""
    try:
        return url.split("/d/")[1].split("/")[0]
    except IndexError:
        raise ValueError(f"Could not extract document ID from URL: {url}")


def text_to_insert_requests(text: str) -> list:
    """Build a batchUpdate insertText request for a new (empty) document."""
    return [{"insertText": {"location": {"index": 1}, "text": text}}]


def _get_credentials():
    creds_path = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS")
    if not creds_path:
        print("Error: GOOGLE_APPLICATION_CREDENTIALS environment variable not set.", file=sys.stderr)
        sys.exit(1)
    return service_account.Credentials.from_service_account_file(creds_path, scopes=SCOPES)


def _get_services():
    creds = _get_credentials()
    docs_service = build("docs", "v1", credentials=creds)
    drive_service = build("drive", "v3", credentials=creds)
    return docs_service, drive_service


def cmd_create(title: str, content_path: str, folder_id: str) -> str:
    """Create a Google Doc with the given content directly in folder_id."""
    with open(content_path, "r", encoding="utf-8") as f:
        content = f.read()

    docs_service, drive_service = _get_services()

    # Create empty doc directly in the target folder via Drive API (supports Shared Drives)
    file = drive_service.files().create(
        body={
            "name": title,
            "mimeType": "application/vnd.google-apps.document",
            "parents": [folder_id],
        },
        supportsAllDrives=True,
        fields="id",
    ).execute()
    doc_id = file["id"]

    # Insert content via Docs API
    docs_service.documents().batchUpdate(
        documentId=doc_id,
        body={"requests": text_to_insert_requests(content)},
    ).execute()

    return f"https://docs.google.com/document/d/{doc_id}"


def cmd_read(url: str) -> str:
    """Read a Google Doc and return its plain text content."""
    doc_id = extract_doc_id(url)
    docs_service, _ = _get_services()
    doc = docs_service.documents().get(documentId=doc_id).execute()

    text_parts = []
    for elem in doc.get("body", {}).get("content", []):
        for para_elem in elem.get("paragraph", {}).get("elements", []):
            text_parts.append(para_elem.get("textRun", {}).get("content", ""))
    return "".join(text_parts)


def main():
    parser = argparse.ArgumentParser()
    subparsers = parser.add_subparsers(dest="command", required=True)

    create_parser = subparsers.add_parser("create")
    create_parser.add_argument("--title", required=True)
    create_parser.add_argument("--content", required=True, help="Path to markdown file")
    create_parser.add_argument("--folder-id", required=True)

    read_parser = subparsers.add_parser("read")
    read_parser.add_argument("--url", required=True)

    args = parser.parse_args()

    if args.command == "create":
        try:
            url = cmd_create(args.title, args.content, args.folder_id)
            print(url)
        except Exception as e:
            print("Google Docs API call failed — check GOOGLE_APPLICATION_CREDENTIALS and WEEKLY_REPORT_FOLDER_ID.", file=sys.stderr)
            sys.exit(1)
    elif args.command == "read":
        try:
            content = cmd_read(args.url)
            print(content)
        except Exception as e:
            print("Failed to read Google Doc — check permissions and GOOGLE_APPLICATION_CREDENTIALS.", file=sys.stderr)
            sys.exit(1)


if __name__ == "__main__":
    main()
