"""Baza orzeczeń Sądu Najwyższego (www.sn.pl) — moduł eksperymentalny.

SN nie udostępnia API — moduł odpytuje wyszukiwarkę bazy orzeczeń i parsuje
HTML. Struktura strony może się zmienić bez zapowiedzi; wtedy funkcje logują
błąd i zwracają puste wyniki. Uwaga: orzecznictwo SN jest też indeksowane
w SAOS (courtType=SUPREME) — ten moduł jest uzupełnieniem, nie jedynym kanałem.
"""

from __future__ import annotations

import logging
import re
import time
from datetime import date

import requests
from bs4 import BeautifulSoup

log = logging.getLogger(__name__)

BASE = "https://www.sn.pl"
SEARCH_PATH = "/orzecznictwo/SitePages/Baza_orzeczen.aspx"

HEADERS = {
    "User-Agent": "nsa-monitor/2.0 (monitoring orzecznictwa dot. srodkow unijnych)",
}

# np. "II CSK 625/24", "I CSK 2106/25", "III CZP 6/24"
SIG_RE = re.compile(r"\b([IVX]+\s+[A-Z]{2,5}[a-z]*\s+\d+/\d{2})\b")
DATE_RE = re.compile(r"\b(\d{4}-\d{2}-\d{2})\b")
DATE_PL_RE = re.compile(r"\b(\d{2})[.-](\d{2})[.-](\d{4})\b")


def _norm_date(ctx: str) -> str:
    m = DATE_RE.search(ctx)
    if m:
        return m.group(1)
    m = DATE_PL_RE.search(ctx)
    if m:
        return f"{m.group(3)}-{m.group(2)}-{m.group(1)}"
    return ""


def search_keyword(
    keyword: str,
    date_from: date,
    delay_seconds: float = 3.0,
    max_pages: int = 5,
    excerpt_length: int = 400,
) -> list[dict]:
    """Wyszukuje frazę w bazie orzeczeń SN."""
    results: list[dict] = []
    session = requests.Session()
    session.headers.update(HEADERS)

    params = {"Szukaj": keyword}
    try:
        resp = session.get(BASE + SEARCH_PATH, params=params, timeout=30)
        resp.raise_for_status()
    except Exception as exc:  # noqa: BLE001
        log.error("SN: błąd wyszukiwania '%s': %s", keyword, exc)
        return results

    seen: set[str] = set()
    for _ in range(max_pages):
        soup = BeautifulSoup(resp.text, "html.parser")
        found_any = False
        for a in soup.find_all("a", href=True):
            text = a.get_text(" ", strip=True)
            m = SIG_RE.search(text)
            if not m:
                continue
            node = a
            for _i in range(4):
                if node.parent is None:
                    break
                node = node.parent
                if len(node.get_text(" ", strip=True)) > 40:
                    break
            ctx = node.get_text(" ", strip=True)
            sygnatura = m.group(1)
            if sygnatura in seen:
                continue
            seen.add(sygnatura)
            found_any = True
            href = a["href"]
            results.append(
                {
                    "data": _norm_date(ctx),
                    "sad": "Sąd Najwyższy",
                    "sygnatura": sygnatura,
                    "rodzaj": "",
                    "fraza": keyword,
                    "fragment": ctx[:excerpt_length],
                    "link": href if href.startswith("http") else BASE + href,
                    "zrodlo": "SN",
                    "court_type": "SUPREME",
                }
            )
        next_el = soup.find("a", string=lambda s: s and ("następ" in s.lower() or s.strip() == ">"))
        if not found_any or not next_el or not next_el.get("href"):
            break
        time.sleep(delay_seconds)
        href = next_el["href"]
        try:
            resp = session.get(href if href.startswith("http") else BASE + href, timeout=30)
            resp.raise_for_status()
        except Exception as exc:  # noqa: BLE001
            log.error("SN: błąd paginacji: %s", exc)
            break

    cutoff = date_from.isoformat()
    return [r for r in results if not r["data"] or r["data"] >= cutoff]
