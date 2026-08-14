#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Raport GSC -> Google Sheet dla kbrysiewicz.pl

Codziennie (dni robocze, GitHub Actions) pobiera z Google Search Console:
  1. metryki fraz docelowych z planu treści (dopasowanie dokładne + "wszystkie słowa"),
  2. metryki stron filarowych,
  3. TOP 10 zapytań dnia,
  4. dzienne podsumowanie całej witryny,
i dopisuje wiersze do wskazanego arkusza (zakładki tworzone automatycznie).

Idempotentny: jeśli dana data jest już w zakładce, pomija ją (bezpieczne re-runy).

Wymagane zmienne środowiskowe:
  GOOGLE_SERVICE_ACCOUNT_JSON  - treść klucza JSON konta serwisowego (lub ścieżka do pliku)
  SHEET_ID                     - ID arkusza Google
Opcjonalne:
  GSC_SITE_URL   - domyślnie "sc-domain:kbrysiewicz.pl"
                   (dla usługi z prefiksem URL: "https://kbrysiewicz.pl/")
  GSC_LAG_DAYS   - opóźnienie danych GSC w dniach, domyślnie 3
"""

import json
import os
import sys
from datetime import date, timedelta

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# ------------------------------------------------------------------ konfiguracja

SITE_URL = os.environ.get("GSC_SITE_URL", "sc-domain:kbrysiewicz.pl")
LAG_DAYS = int(os.environ.get("GSC_LAG_DAYS", "3"))

# Frazy z planu treści (małymi literami — GSC zwraca zapytania lowercase)
TRACKED_QUERIES = [
    "wezwanie do zwrotu dotacji",
    "zwrot dotacji unijnej",
    "zwrot dofinansowania",
    "kpo horeca zwrot",
    "kontrola horeca",
    "zwrot środków kpo",
    "odwołanie od decyzji o zwrocie",
    "zwrot subwencji pfr",
    "korekta finansowa",
    "odpowiedzialność członka zarządu dotacja",
    "informacja pokontrolna zastrzeżenia",
]

# Strony filarowe (ścieżki); wpis o zarządzie już istnieje, reszta wg planu
TRACKED_PAGES = [
    "/zwrot-srodkow-kpo-horeca/",
    "/wezwanie-do-zwrotu-dotacji/",
    "/decyzja-o-zwrocie-dofinansowania-odwolanie/",
    "/odpowiedzialnosc-menedzera-za-niezwrocona-dotacje/",
    "/zwrot-subwencji-pfr/",
    "/korekta-finansowa-jak-kwestionowac/",
    "/kontrola-projektu-unijnego/",
]

TOP_N = 10

TAB_QUERIES = "GSC_Frazy"
TAB_PAGES = "GSC_Strony"
TAB_TOP = "GSC_Top"
TAB_SUMMARY = "GSC_Dziennie"

HEADERS = {
    TAB_QUERIES: ["data", "fraza", "dopasowanie", "klikniecia", "wyswietlenia", "ctr_proc", "pozycja"],
    TAB_PAGES:   ["data", "strona", "klikniecia", "wyswietlenia", "ctr_proc", "pozycja"],
    TAB_TOP:     ["data", "miejsce", "zapytanie", "klikniecia", "wyswietlenia", "ctr_proc", "pozycja"],
    TAB_SUMMARY: ["data", "klikniecia", "wyswietlenia", "ctr_proc", "pozycja_srednia"],
}

SCOPES = [
    "https://www.googleapis.com/auth/webmasters.readonly",
    "https://www.googleapis.com/auth/spreadsheets",
]

# ------------------------------------------------------------------ pomocnicze


def get_credentials():
    raw = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not raw:
        sys.exit("BŁĄD: brak zmiennej GOOGLE_SERVICE_ACCOUNT_JSON")
    if os.path.isfile(raw):
        with open(raw, encoding="utf-8") as f:
            raw = f.read()
    info = json.loads(raw)
    return service_account.Credentials.from_service_account_info(info, scopes=SCOPES)


def gsc_query(gsc, day_iso, dimensions):
    """Jedno zapytanie do Search Analytics dla jednego dnia."""
    body = {
        "startDate": day_iso,
        "endDate": day_iso,
        "rowLimit": 25000,
        "dataState": "all",
    }
    if dimensions:
        body["dimensions"] = dimensions
    try:
        resp = gsc.searchanalytics().query(siteUrl=SITE_URL, body=body).execute()
    except HttpError as e:
        if e.resp.status == 403:
            print(
                "BŁĄD 403: konto serwisowe nie ma dostępu do usługi GSC "
                f"'{SITE_URL}'. Dodaj e-mail konta w GSC: Ustawienia -> "
                "Użytkownicy i uprawnienia (wystarczy 'Ograniczony')."
            )
        raise
    return resp.get("rows", [])


def pct(x):
    return round(x * 100, 2)


def pos(x):
    return round(x, 1)


def norm_path(url_or_path):
    """Zwraca ścieżkę bez końcowego '/' — do porównywania stron."""
    p = url_or_path
    if p.startswith("http"):
        p = "/" + p.split("/", 3)[3] if p.count("/") >= 3 else "/"
    return p.rstrip("/") or "/"


# ------------------------------------------------------------------ budowa wierszy


def rows_for_queries(day_iso, query_rows):
    """Dla każdej frazy: wiersz 'dokladne' i 'slowa' (wszystkie słowa frazy w zapytaniu)."""
    by_query = {}
    for r in query_rows:
        q = (r.get("keys") or [""])[0]
        by_query[q] = r

    out = []
    for phrase in TRACKED_QUERIES:
        ph = phrase.lower().strip()

        # dopasowanie dokładne
        r = by_query.get(ph)
        if r:
            out.append([day_iso, phrase, "dokladne",
                        r.get("clicks", 0), r.get("impressions", 0),
                        pct(r.get("ctr", 0.0)), pos(r.get("position", 0.0))])
        else:
            out.append([day_iso, phrase, "dokladne", 0, 0, 0, 0])

        # dopasowanie: wszystkie słowa frazy występują w zapytaniu
        tokens = ph.split()
        clicks = impressions = 0
        pos_weighted = 0.0
        for q, r in by_query.items():
            if all(t in q for t in tokens):
                clicks += r.get("clicks", 0)
                impressions += r.get("impressions", 0)
                pos_weighted += r.get("position", 0.0) * r.get("impressions", 0)
        avg_pos = pos(pos_weighted / impressions) if impressions else 0
        ctr = pct(clicks / impressions) if impressions else 0
        out.append([day_iso, phrase, "slowa", clicks, impressions, ctr, avg_pos])
    return out


def rows_for_pages(day_iso, page_rows):
    by_path = {}
    for r in page_rows:
        url = (r.get("keys") or [""])[0]
        by_path[norm_path(url)] = r

    out = []
    for slug in TRACKED_PAGES:
        r = by_path.get(norm_path(slug))
        if r:
            out.append([day_iso, slug,
                        r.get("clicks", 0), r.get("impressions", 0),
                        pct(r.get("ctr", 0.0)), pos(r.get("position", 0.0))])
        else:
            out.append([day_iso, slug, 0, 0, 0, 0])
    return out


def rows_for_top(day_iso, query_rows):
    ranked = sorted(
        query_rows,
        key=lambda r: (r.get("clicks", 0), r.get("impressions", 0)),
        reverse=True,
    )[:TOP_N]
    out = []
    for i, r in enumerate(ranked, start=1):
        q = (r.get("keys") or [""])[0]
        out.append([day_iso, i, q,
                    r.get("clicks", 0), r.get("impressions", 0),
                    pct(r.get("ctr", 0.0)), pos(r.get("position", 0.0))])
    return out


def rows_for_summary(day_iso, total_rows):
    if total_rows:
        r = total_rows[0]
        return [[day_iso, r.get("clicks", 0), r.get("impressions", 0),
                 pct(r.get("ctr", 0.0)), pos(r.get("position", 0.0))]]
    return [[day_iso, 0, 0, 0, 0]]


# ------------------------------------------------------------------ arkusz


def ensure_tabs(sheets, sheet_id):
    meta = sheets.spreadsheets().get(
        spreadsheetId=sheet_id, fields="sheets.properties.title"
    ).execute()
    existing = {s["properties"]["title"] for s in meta.get("sheets", [])}

    requests = [
        {"addSheet": {"properties": {"title": tab}}}
        for tab in HEADERS if tab not in existing
    ]
    if requests:
        sheets.spreadsheets().batchUpdate(
            spreadsheetId=sheet_id, body={"requests": requests}
        ).execute()

    # nagłówki, jeśli zakładka pusta
    for tab, header in HEADERS.items():
        vals = sheets.spreadsheets().values().get(
            spreadsheetId=sheet_id, range=f"'{tab}'!A1:A1"
        ).execute().get("values")
        if not vals:
            sheets.spreadsheets().values().update(
                spreadsheetId=sheet_id, range=f"'{tab}'!A1",
                valueInputOption="RAW", body={"values": [header]},
            ).execute()


def date_already_written(sheets, sheet_id, tab, day_iso):
    vals = sheets.spreadsheets().values().get(
        spreadsheetId=sheet_id, range=f"'{tab}'!A:A"
    ).execute().get("values", [])
    return any(row and row[0] == day_iso for row in vals)


def append_rows(sheets, sheet_id, tab, rows):
    if not rows:
        return
    sheets.spreadsheets().values().append(
        spreadsheetId=sheet_id, range=f"'{tab}'!A1",
        valueInputOption="RAW", insertDataOption="INSERT_ROWS",
        body={"values": rows},
    ).execute()


# ------------------------------------------------------------------ main


def main():
    sheet_id = os.environ.get("SHEET_ID")
    if not sheet_id:
        sys.exit("BŁĄD: brak zmiennej SHEET_ID")

    report_day = date.today() - timedelta(days=LAG_DAYS)
    day_iso = report_day.isoformat()
    print(f"Raport GSC dla {SITE_URL}, dzień: {day_iso} (lag {LAG_DAYS} dni)")

    creds = get_credentials()
    gsc = build("searchconsole", "v1", credentials=creds, cache_discovery=False)
    sheets = build("sheets", "v4", credentials=creds, cache_discovery=False)

    query_rows = gsc_query(gsc, day_iso, ["query"])
    page_rows = gsc_query(gsc, day_iso, ["page"])
    total_rows = gsc_query(gsc, day_iso, [])
    print(f"GSC: {len(query_rows)} zapytań, {len(page_rows)} stron")

    ensure_tabs(sheets, sheet_id)

    plan = [
        (TAB_QUERIES, rows_for_queries(day_iso, query_rows)),
        (TAB_PAGES,   rows_for_pages(day_iso, page_rows)),
        (TAB_TOP,     rows_for_top(day_iso, query_rows)),
        (TAB_SUMMARY, rows_for_summary(day_iso, total_rows)),
    ]

    for tab, rows in plan:
        if date_already_written(sheets, sheet_id, tab, day_iso):
            print(f"[{tab}] {day_iso} już zapisany — pomijam")
            continue
        append_rows(sheets, sheet_id, tab, rows)
        print(f"[{tab}] dopisano {len(rows)} wierszy")

    print("Gotowe.")


if __name__ == "__main__":
    main()
