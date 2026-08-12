"""Klient API SAOS (System Analizy Orzeczeń Sądowych, saos.org.pl).

SAOS udostępnia otwarte REST API indeksujące m.in. orzecznictwo
sądów administracyjnych. Dokumentacja: https://www.saos.org.pl/help/index.php/dokumentacja-api
"""

from __future__ import annotations

import logging
import time
from datetime import date

import requests

log = logging.getLogger(__name__)

SEARCH_URL = "https://www.saos.org.pl/api/search/judgments"
DETAILS_URL = "https://www.saos.org.pl/api/judgments/{id}"
JUDGMENT_URL = "https://www.saos.org.pl/judgments/{id}"

HEADERS = {
    "User-Agent": "nsa-monitor/2.0 (monitoring orzecznictwa)",
    "Accept": "application/json",
}

PAGE_SIZE = 100          # maksimum SAOS
REQUEST_DELAY = 0.4      # grzecznościowa przerwa między zapytaniami


def _court_name(item: dict) -> str:
    """Wyciąga nazwę sądu — struktura różni się między typami sądów."""
    division = item.get("division") or {}
    court = division.get("court") or {}
    if court.get("name"):
        return court["name"]
    if item.get("courtType") == "ADMINISTRATIVE":
        return item.get("courtName") or "Sąd administracyjny (NSA/WSA)"
    return item.get("courtType", "")


def _case_number(item: dict) -> str:
    cases = item.get("courtCases") or []
    return cases[0].get("caseNumber", "").strip() if cases else ""


def search_keyword(
    keyword: str,
    date_from: date,
    court_types: list[str],
    max_results: int = 100,
    excerpt_length: int = 400,
) -> list[dict]:
    """Zwraca orzeczenia pasujące do frazy od zadanej daty.

    Pełna paginacja: pobiera kolejne strony aż do wyczerpania wyników
    lub osiągnięcia max_results (na typ sądu).
    """
    results: list[dict] = []
    for court_type in court_types:
        page = 0
        fetched = 0
        while fetched < max_results:
            params = {
                "all": keyword,
                "courtType": court_type,
                "judgmentDateFrom": date_from.isoformat(),
                "pageSize": PAGE_SIZE,
                "pageNumber": page,
                "sortingField": "JUDGMENT_DATE",
                "sortingDirection": "DESC",
            }
            try:
                resp = requests.get(SEARCH_URL, params=params, headers=HEADERS, timeout=30)
                resp.raise_for_status()
                payload = resp.json()
            except Exception as exc:  # noqa: BLE001
                log.error("SAOS: błąd zapytania '%s' (%s, str. %d): %s",
                          keyword, court_type, page, exc)
                break

            items = payload.get("items") or []
            if not items:
                break

            for item in items:
                sygnatura = _case_number(item)
                if not sygnatura:
                    continue
                text = (item.get("textContent") or "").strip().replace("\n", " ")
                results.append(
                    {
                        "saos_id": item.get("id"),
                        "data": item.get("judgmentDate", ""),
                        "sad": _court_name(item),
                        "sygnatura": sygnatura,
                        "rodzaj": item.get("judgmentType", ""),
                        "fraza": keyword,
                        "fragment": text[:excerpt_length],
                        "link": JUDGMENT_URL.format(id=item.get("id")),
                        "zrodlo": "SAOS",
                        "court_type": court_type,
                    }
                )
            fetched += len(items)
            log.info("SAOS: '%s' (%s) str. %d -> %d wyników (razem %d)",
                     keyword, court_type, page, len(items), fetched)
            if len(items) < PAGE_SIZE:
                break
            page += 1
            time.sleep(REQUEST_DELAY)

    return results


def get_full_text(saos_id: int) -> str:
    """Pobiera pełną treść orzeczenia (endpoint szczegółów)."""
    if not saos_id:
        return ""
    try:
        resp = requests.get(DETAILS_URL.format(id=saos_id), headers=HEADERS, timeout=30)
        resp.raise_for_status()
        data = resp.json().get("data") or {}
        return (data.get("textContent") or "").strip()
    except Exception as exc:  # noqa: BLE001
        log.warning("SAOS: nie udało się pobrać pełnego tekstu id=%s: %s", saos_id, exc)
        return ""
