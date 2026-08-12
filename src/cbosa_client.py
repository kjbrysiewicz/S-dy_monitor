"""Klient CBOSA (orzeczenia.nsa.gov.pl) — wyszukiwanie i weryfikacja symbolu sprawy.

CBOSA nie udostępnia API — moduł odpytuje formularz wyszukiwania i parsuje HTML.
Struktura strony może się zmienić bez zapowiedzi; wtedy funkcje logują błąd
i zwracają puste wyniki (główny przebieg nie jest przerywany).

Kluczowa funkcja filtra kategorii: sprawy dotyczące środków unijnych mają
w CBOSA symbol 6559 (konfigurowalne w config.yaml). Moduł pozwala:
- wyszukiwać bezpośrednio po symbolu (search_by_symbol),
- weryfikować symbol konkretnego orzeczenia po sygnaturze (get_symbols),
- wyszukiwać po frazie (search_keyword).

Zasady: niskie tempo zapytań (delay), identyfikujący User-Agent.
"""

from __future__ import annotations

import logging
import re
import time
from datetime import date

import requests
from bs4 import BeautifulSoup

log = logging.getLogger(__name__)

BASE = "https://orzeczenia.nsa.gov.pl"
SEARCH_PATH = "/cbo/search"

HEADERS = {
    "User-Agent": "nsa-monitor/2.0 (monitoring orzecznictwa dot. srodkow unijnych)",
}

SYMBOL_RE = re.compile(r"Symbol z opisem[^0-9]*(\d{3,4})", re.IGNORECASE)


def _parse_result_rows(soup: BeautifulSoup, fraza: str, excerpt_length: int) -> list[dict]:
    out: list[dict] = []
    for row in soup.select("table tr"):
        link_el = row.select_one("a[href*='/doc/']")
        if not link_el:
            continue
        cells = [c.get_text(" ", strip=True) for c in row.select("td")]
        row_text = " | ".join(cells)

        link_text = link_el.get_text(" ", strip=True)
        sygnatura = link_text.split(" - ")[0].strip()
        rodzaj, sad = "", ""
        if " - " in link_text:
            parts = link_text.split(" - ", 1)[1].split()
            if parts:
                rodzaj, sad = parts[0], " ".join(parts[1:])

        data_orz = ""
        for token in row_text.replace("|", " ").split():
            if len(token) == 10 and token[4] == "-" and token[7] == "-":
                data_orz = token
                break

        if not sygnatura:
            continue
        out.append(
            {
                "data": data_orz,
                "sad": sad or "NSA/WSA",
                "sygnatura": sygnatura,
                "rodzaj": rodzaj,
                "fraza": fraza,
                "fragment": row_text[:excerpt_length],
                "link": BASE + link_el["href"],
                "zrodlo": "CBOSA",
                "court_type": "ADMINISTRATIVE",
            }
        )
    return out


def _paged_search(
    form: dict, fraza: str, delay_seconds: float, max_pages: int, excerpt_length: int
) -> list[dict]:
    results: list[dict] = []
    session = requests.Session()
    session.headers.update(HEADERS)
    try:
        resp = session.post(BASE + SEARCH_PATH, data=form, timeout=30)
        resp.raise_for_status()
    except Exception as exc:  # noqa: BLE001
        log.error("CBOSA: błąd wyszukiwania (%s): %s", fraza, exc)
        return results

    for _ in range(max_pages):
        soup = BeautifulSoup(resp.text, "html.parser")
        results.extend(_parse_result_rows(soup, fraza, excerpt_length))
        next_el = soup.find("a", string=lambda s: s and "następ" in s.lower())
        if not next_el or not next_el.get("href"):
            break
        time.sleep(delay_seconds)
        try:
            resp = session.get(BASE + next_el["href"], timeout=30)
            resp.raise_for_status()
        except Exception as exc:  # noqa: BLE001
            log.error("CBOSA: błąd paginacji: %s", exc)
            break
    return results


def search_by_symbol(
    symbol: str,
    date_from: date,
    delay_seconds: float = 3.0,
    max_pages: int = 50,
    excerpt_length: int = 400,
) -> list[dict]:
    """Wyszukuje WSZYSTKIE orzeczenia z danym symbolem sprawy (np. 6559)."""
    form = {
        "symbole": symbol,
        "odDaty": date_from.strftime("%Y-%m-%d"),
        "doDaty": date.today().strftime("%Y-%m-%d"),
    }
    return _paged_search(form, f"symbol {symbol}", delay_seconds, max_pages, excerpt_length)


def search_keyword(
    keyword: str,
    date_from: date,
    delay_seconds: float = 3.0,
    max_pages: int = 3,
    excerpt_length: int = 400,
    symbol: str | None = None,
) -> list[dict]:
    """Wyszukuje frazę; opcjonalnie zawężone do symbolu sprawy."""
    form = {
        "wszystkieSlowa": keyword,
        "odDaty": date_from.strftime("%Y-%m-%d"),
        "doDaty": date.today().strftime("%Y-%m-%d"),
    }
    if symbol:
        form["symbole"] = symbol
    return _paged_search(form, keyword, delay_seconds, max_pages, excerpt_length)


def get_symbols(sygnatura: str, delay_seconds: float = 1.5) -> set[str] | None:
    """Zwraca symbole sprawy dla sygnatury (wg strony orzeczenia w CBOSA).

    None = nie udało się ustalić (błąd sieci/parsowania lub brak w CBOSA)
    — wywołujący decyduje, czy w razie wątpliwości zachować rekord.
    """
    session = requests.Session()
    session.headers.update(HEADERS)
    try:
        resp = session.post(
            BASE + SEARCH_PATH, data={"sygnatura": sygnatura}, timeout=30
        )
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")
        link_el = soup.select_one("a[href*='/doc/']")
        if not link_el:
            return None
        time.sleep(delay_seconds)
        doc = session.get(BASE + link_el["href"], timeout=30)
        doc.raise_for_status()
        symbols = set(SYMBOL_RE.findall(doc.text))
        # symbol bywa też w wierszu tabeli "Symbol z opisem" jako osobna komórka
        dsoup = BeautifulSoup(doc.text, "html.parser")
        for tr in dsoup.select("tr"):
            cells = [c.get_text(" ", strip=True) for c in tr.select("td")]
            if len(cells) >= 2 and "symbol" in cells[0].lower():
                symbols.update(re.findall(r"\b(\d{3,4})\b", cells[1]))
        return symbols or None
    except Exception as exc:  # noqa: BLE001
        log.warning("CBOSA: weryfikacja symbolu %s nie powiodła się: %s", sygnatura, exc)
        return None
