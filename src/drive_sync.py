"""Synchronizacja pliku .xlsx z Google Drive przez konto serwisowe.

Wymaga:
- zmiennej środowiskowej GDRIVE_SERVICE_ACCOUNT_JSON (pełna zawartość klucza JSON),
- zmiennej środowiskowej GDRIVE_FILE_ID (ID pliku .xlsx na Drive),
- udostępnienia pliku adresowi e-mail konta serwisowego z uprawnieniem "Edytor".
"""

from __future__ import annotations

import io
import json
import logging
import os

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload

log = logging.getLogger(__name__)

SCOPES = ["https://www.googleapis.com/auth/drive"]
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _service():
    raw = os.environ.get("GDRIVE_SERVICE_ACCOUNT_JSON")
    if not raw:
        raise RuntimeError("Brak zmiennej GDRIVE_SERVICE_ACCOUNT_JSON")
    info = json.loads(raw)
    creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def file_id() -> str:
    fid = os.environ.get("GDRIVE_FILE_ID", "").strip()
    if not fid:
        raise RuntimeError("Brak zmiennej GDRIVE_FILE_ID")
    return fid


def download_xlsx(local_path: str) -> None:
    """Pobiera plik .xlsx z Drive do local_path."""
    svc = _service()
    request = svc.files().get_media(fileId=file_id())
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _status, done = downloader.next_chunk()
    with open(local_path, "wb") as fh:
        fh.write(buf.getvalue())
    log.info("Pobrano plik z Drive (%d bajtów)", len(buf.getvalue()))


def upload_xlsx(local_path: str) -> None:
    """Nadpisuje plik na Drive nową wersją (historia wersji Drive zostaje zachowana)."""
    svc = _service()
    media = MediaFileUpload(local_path, mimetype=XLSX_MIME, resumable=True)
    svc.files().update(fileId=file_id(), media_body=media).execute()
    log.info("Wysłano zaktualizowany plik na Drive")
