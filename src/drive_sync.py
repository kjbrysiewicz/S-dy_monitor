"""Synchronizacja pliku bazy z Google Drive przez konto serwisowe.

Obsługuje automatycznie:
- zwykły plik .xlsx ORAZ natywny Arkusz Google (eksport/import),
- pliki na "Moim dysku" ORAZ na Dyskach współdzielonych (supportsAllDrives).

Wymaga:
- GDRIVE_SERVICE_ACCOUNT_JSON (pełna zawartość klucza JSON),
- GDRIVE_FILE_ID (ID pliku z adresu URL),
- udostępnienia pliku adresowi konta serwisowego z uprawnieniem "Edytor".
"""

from __future__ import annotations

import io
import json
import logging
import os

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload

log = logging.getLogger(__name__)

SCOPES = ["https://www.googleapis.com/auth/drive"]
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
GSHEET_MIME = "application/vnd.google-apps.spreadsheet"

_client_email: str = ""


def _service():
    global _client_email
    raw = os.environ.get("GDRIVE_SERVICE_ACCOUNT_JSON")
    if not raw:
        raise RuntimeError("Brak zmiennej GDRIVE_SERVICE_ACCOUNT_JSON")
    info = json.loads(raw)
    _client_email = info.get("client_email", "?")
    creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def file_id() -> str:
    fid = os.environ.get("GDRIVE_FILE_ID", "").strip()
    if not fid:
        raise RuntimeError("Brak zmiennej GDRIVE_FILE_ID")
    return fid


def _explain_404() -> RuntimeError:
    fid = file_id()
    return RuntimeError(
        "Google Drive nie widzi pliku (404). To oznacza jedną z dwóch rzeczy:\n"
        f"  1) Plik NIE jest udostępniony kontu serwisowemu.\n"
        f"     -> Otwórz plik na Dysku, przycisk 'Udostępnij' i dodaj DOKŁADNIE ten adres\n"
        f"        z uprawnieniem 'Edytujący':  {_client_email}\n"
        f"  2) ID pliku w sekrecie GDRIVE_FILE_ID jest błędne.\n"
        f"     -> Użyte ID (pierwsze/ostatnie znaki): {fid[:6]}...{fid[-6:]} (długość {len(fid)}).\n"
        f"        Porównaj z adresem URL pliku: docs.google.com/spreadsheets/d/ID/edit\n"
        "Po poprawce uruchom workflow ponownie."
    )


def _file_mime(svc) -> str:
    try:
        meta = (
            svc.files()
            .get(fileId=file_id(), fields="mimeType,name,driveId", supportsAllDrives=True)
            .execute()
        )
    except HttpError as exc:
        if exc.resp.status == 404:
            raise _explain_404() from exc
        raise
    log.info(
        "Plik na Drive: '%s' (%s)%s",
        meta.get("name"),
        meta.get("mimeType"),
        " [Dysk współdzielony]" if meta.get("driveId") else "",
    )
    return meta.get("mimeType", "")


def download_xlsx(local_path: str) -> None:
    """Pobiera bazę jako .xlsx (eksportując Arkusz Google, jeśli trzeba)."""
    svc = _service()
    mime = _file_mime(svc)
    if mime == GSHEET_MIME:
        request = svc.files().export_media(fileId=file_id(), mimeType=XLSX_MIME)
    else:
        request = svc.files().get_media(fileId=file_id(), supportsAllDrives=True)
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
    try:
        svc.files().update(
            fileId=file_id(), media_body=media, supportsAllDrives=True
        ).execute()
    except HttpError as exc:
        if exc.resp.status == 404:
            raise _explain_404() from exc
        raise
    log.info("Wysłano zaktualizowany plik na Drive")
