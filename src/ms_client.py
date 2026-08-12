"""Portal Orzeczeń Sądów Powszechnych (orzeczenia.ms.gov.pl) — moduł eksperymentalny.

Portal nie udostępnia API — moduł odpytuje wyszukiwarkę i parsuje HTML.
Struktura strony może się zmienić bez zapowiedzi; wtedy funkcje logują błąd
i zwracają puste wyniki (główny przebieg nie jest przerywany).

Parser jest celowo elastyczny: wyłuskuje z wyników linki do orzeczeń
i rozpoznaje sygnatury/sądy/daty wzorcami, zamiast polegać na konkretnych
klasach CSS.
"""

from __future__ import annotations

import logging
import re
import time
from datetime import date

import requests
from bs4 import BeautifulSoup

log = logging.getLogger(__name__)

BASE = "https://orzeczenia.ms.gov.pl"
SEARCH_PATH = "/search/advanced"

HEADERS = {
    "User-Agent": "nsa-monitor/2.0 (monitoring orzecznictwa dot. srodkow unijnych)",
}

# np. "I C 123/24", "V ACa 55/23", "XVI GC 100/26"
SIG_RE = re.compile(r"\b([IVXL]+\s+[A-Za-z]{1,5}[a-z]*\s+\d+/\d{2})\b")
DATE_RE = re.compile(r"\b(\d{4}-\d{2}-\d{2})\b")
COURT_RE = re.compile(r"(Sąd (?:Rejonowy|Okręgowy|Apelacyjny)[^,|<\n]{0,60})")


def _parse_results(html: str, fraza: str, excerpt_length: int) -> list[dict]:
    soup = BeautifulSoup(html, "html.parser")
    out: list[dict] = []
    seen: set[str] = set()
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if not any(p in href for p in ("/content/", "/details/", "/judgement", "/orzeczenie")):
            continue
        # kontekst wiersza wyniku: najbliższy przodek z sensowną ilością tekstu
        node = a
        for _ in range(4):
            if node.parent is None:
                break
            node = node.parent
            if len(node.get_text(" ", strip=True)) > 40:
                break
        ctx = node.get_text(" ", strip=True)

        m = SIG_RE.search(a.get_text(" ", strip=True)) or SIG_RE.search(ctx)
        if not m:
            continue
        sygnatura = m.group(1)
        dedup = sygnatura + "|" + href
        if dedup in seen:
            continue
        seen.add(dedup)

        dm = DATE_RE.search(ctx)
        cm = COURT_RE.search(ctx)
        link = href if href.startswith("http") else BASE + href
        out.append(
            {
                "data": dm.group(1) if dm else "",
                "sad": cm.group(1).strip() if cm else "Sąd powszechny",
                "sygnatura": sygnatura,
                "rodzaj": "",
                "fraza": fraza,
                "fragment": ctx[:excerpt_length],
                "link": link,
                "zrodlo": "PO SP",  # Portal Orzeczeń Sądów Powszechnych
                "court_type": "COMMON",
            }
        )
    return out


def search_keyword(
    keyword: str,
    date_from: date,
    delay_seconds: float = 3.0,
    max_pages: int = 5,
    excerpt_length: int = 400,
) -> list[dict]:
    """Wyszukuje frazę w Portalu Orzeczeń Sądów Powszechnych."""
    results: list[dict] = []
    session = requests.Session()
    session.headers.update(HEADERS)

    params = {
        "t:lb_text": keyword,
        "t:lb_dateFrom": date_from.strftime("%Y-%m-%d"),
        "t:lb_dateTo": date.today().strftime("%Y-%m-%d"),
    }
    try:
        resp = session.get(BASE + SEARCH_PATH, params=params, timeout=30)
        resp.raise_for_status()
    except Exception as exc:  # noqa: BLE001
        log.error("PO SP: błąd wyszukiwania '%s': %s", keyword, exc)
        return results

    for _ in range(max_pages):
        batch = _parse_results(resp.text, keyword, excerpt_length)
        results.extend(batch)
        soup = BeautifulSoup(resp.text, "html.parser")
        next_el = soup.find("a", string=lambda s: s and ("następ" in s.lower() or s.strip() == ">"))
        if not batch or not next_el or not next_el.get("href"):
            break
        time.sleep(delay_seconds)
        href = next_el["href"]
        try:
            resp = session.get(href if href.startswith("http") else BASE + href, timeout=30)
            resp.raise_for_status()
        except Exception as exc:  # noqa: BLE001
            log.error("PO SP: błąd paginacji: %s", exc)
            break

    # filtr dat po stronie klienta (na wypadek zignorowania parametrów przez portal)
    cutoff = date_from.isoformat()
    return [r for r in results if not r["data"] or r["data"] >= cutoff]
