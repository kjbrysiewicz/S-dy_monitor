"""Synchronizacja pliku bazy z Google Drive przez konto serwisowe.

Obsługuje OBA formaty automatycznie:
- zwykły plik .xlsx na Dysku -> pobranie/wysyłka bez zmian formatu,
- natywny Arkusz Google (application/vnd.google-apps.spreadsheet)
  -> pobranie przez eksport do .xlsx, wysyłka przez aktualizację treści
     (plik na Dysku pozostaje Arkuszem Google).

Wymaga:
- zmiennej środowiskowej GDRIVE_SERVICE_ACCOUNT_JSON (pełna zawartość klucza JSON),
- zmiennej środowiskowej GDRIVE_FILE_ID (ID pliku z adresu URL),
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
GSHEET_MIME = "application/vnd.google-apps.spreadsheet"


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


def _file_mime(svc) -> str:
    meta = svc.files().get(fileId=file_id(), fields="mimeType,name").execute()
    log.info("Plik na Drive: '%s' (%s)", meta.get("name"), meta.get("mimeType"))
    return meta.get("mimeType", "")


def download_xlsx(local_path: str) -> None:
    """Pobiera bazę jako .xlsx (eksportując Arkusz Google, jeśli trzeba)."""
    svc = _service()
    mime = _file_mime(svc)
    if mime == GSHEET_MIME:
        request = svc.files().export_media(fileId=file_id(), mimeType=XLSX_MIME)
    else:
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
    """Nadpisuje bazę na Drive nową wersją (historia wersji zostaje zachowana).

    Jeśli plik na Dysku jest Arkuszem Google, treść .xlsx jest importowana
    do tego samego Arkusza (plik pozostaje Arkuszem Google).
    """
    svc = _service()
    media = MediaFileUpload(local_path, mimetype=XLSX_MIME, resumable=True)
    svc.files().update(fileId=file_id(), media_body=media).execute()
    log.info("Wysłano zaktualizowany plik na Drive")
