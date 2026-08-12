"""Heurystyczna analiza treści orzeczenia — klasyfikacja słownikowa.

Przypisuje wartości WYŁĄCZNIE ze słowników kontrolowanych mastera:
- "Problemy prawne" (wiele wartości, rozdzielane "; " jak w istniejących rekordach),
- "Etap sporu" (jedna wartość, wg pierwszeństwa dopasowań).

Celowo NIE ocenia wyniku dla beneficjenta ani nie formułuje tezy —
to wymaga lektury i pozostaje ręczne (Instrukcja, reguła 3). Analiza jest
podpowiedzią przy wstępnej klasyfikacji, stąd status "Metadane wstępne".
"""

from __future__ import annotations

import re

# --- Problemy prawne: wzorce -> wartość ze słownika mastera ---
PROBLEM_PATTERNS: list[tuple[str, list[str]]] = [
    ("nieprawidłowość i szkoda", [
        r"nieprawidłowoś", r"szkod[aęy] w budżecie", r"art\.?\s*2\s*pkt\s*36",
    ]),
    ("proporcjonalność i taryfikator", [
        r"taryfikator", r"proporcjonalnoś", r"korekt[aęy] finansow", r"stawk[ai] procentow",
    ]),
    ("zamówienia publiczne", [
        r"zamówie[nń] publicznych", r"prawo zamówień", r"\bpzp\b", r"zasad[aęy] konkurencyjnoś",
    ]),
    ("kwalifikowalność wydatków", [
        r"kwalifikowalnoś", r"wydat(ek|ki|ków) niekwalifikowaln", r"koszt[óy]w kwalifikowaln",
    ]),
    ("cele i wskaźniki projektu", [
        r"wskaźnik(i|ów)? (rezultatu|produktu)", r"osiągnięci[ae] cel", r"nieosiągnięci",
    ]),
    ("trwałość projektu", [
        r"trwałoś(ć|ci) (projektu|operacji)", r"zasadnicz[aą] modyfikacj",
    ]),
    ("change of control / MŚP", [
        r"status MŚP", r"małe(go)? i średnie(go)?", r"zmian[aęy] kontroli", r"przedsiębiorstw powiązanych",
    ]),
    ("pomoc publiczna", [
        r"pomoc(y|ą)? publiczn", r"pomoc(y|ą)? państwa", r"art\.?\s*107", r"\bGBER\b",
    ]),
    ("odsetki i zaliczki", [
        r"odset(ki|ek|kach)", r"zalicz(ka|ki|ek)",
    ]),
    ("przedawnienie", [
        r"przedawni", r"rozporządzeni[ae].{0,20}2988/95",
    ]),
    ("procedura i uzasadnienie", [
        r"uzasadnieni[ae] decyzji", r"art\.?\s*107\s*§", r"wad(a|liwość) uzasadnienia",
        r"naruszeni[ae] przepisów postępowania",
    ]),
    ("bezstronność i zaufanie", [
        r"bezstronnoś", r"zasad[aęy] zaufania", r"art\.?\s*8\s*k\.?p\.?a",
    ]),
    ("alokacja i dostęp do dofinansowania", [
        r"alokacj", r"wyczerpani[ae] (środków|alokacji)", r"list[aęy] rankingow",
    ]),
    ("droga cywilna i weksel", [
        r"weksl", r"subwencj[aięy]", r"powództw", r"pozw(u|em)\b",
    ]),
    ("art. 5 k.c.", [
        r"art\.?\s*5\s*k\.?c", r"nadużyci[ae] prawa podmiotowego",
    ]),
    ("wykluczenie", [
        r"wykluczeni[ae] (z możliwości|beneficjenta)", r"art\.?\s*207\s*ust\.?\s*4",
    ]),
    ("informacja publiczna", [
        r"informacj[iaę] publiczn", r"dostęp(u|ie) do informacji",
    ]),
    ("odszkodowanie / dochodzenie dofinansowania", [
        r"odszkodowa", r"dochodzeni[ae] dofinansowania",
    ]),
]

# --- Etap sporu: pierwszeństwo dopasowań (pierwszy trafiony wygrywa) ---
ETAP_PATTERNS: list[tuple[str, list[str]]] = [
    ("Pomoc publiczna / odzyskanie", [
        r"odzyskani[ae] pomocy", r"windykacj[aię] pomocy",
    ]),
    ("Postępowanie cywilne / wekslowe", [
        r"weksl", r"powództw", r"subwencj[aięy] finansow",
    ]),
    ("Nabór i ocena projektu", [
        r"protest(u|em)?\b", r"ocen[aęy] (projektu|wniosku)", r"nab[óo]r(u|ze)?\b",
        r"negatywn[aąej]+ ocen",
    ]),
    ("Korekta i postępowanie zwrotowe", [
        r"zwrot(u|em)? (środków|dofinansowania)", r"korekt[aęy] finansow",
        r"art\.?\s*207", r"określeni[ae] .{0,30}do zwrotu",
    ]),
    ("Realizacja, zmiana i trwałość", [
        r"trwałoś(ć|ci)", r"zmian[aęy] umowy o dofinansowanie", r"aneks",
    ]),
    ("Kontrola i zamówienia", [
        r"kontrol[aięy] (projektu|na miejscu)", r"zamówie[nń] publicznych",
    ]),
    ("Kontrola sądowoadministracyjna", [
        r"skarg[aęi] kasacyjn", r"wojewódzki sąd administracyjny",
    ]),
]


def _hits(text: str, patterns: list[str]) -> int:
    return sum(1 for p in patterns if re.search(p, text, re.IGNORECASE))


def classify(text: str, base_problem: str = "") -> dict:
    """Zwraca {'problemy': str, 'etap': str|None, 'trafienia': dict}.

    problemy — wartości ze słownika połączone '; ' (max 4, wg liczby trafień),
    etap — pierwsza pasująca wartość ze słownika lub None (zostaje default),
    trafienia — diagnostyka do notatki wewnętrznej.
    """
    if not text:
        return {"problemy": base_problem, "etap": None, "trafienia": {}}

    scored = []
    for label, patterns in PROBLEM_PATTERNS:
        h = _hits(text, patterns)
        if h:
            scored.append((h, label))
    scored.sort(reverse=True)
    problems = [label for _h, label in scored[:4]]
    if base_problem and base_problem not in problems:
        problems.insert(0, base_problem)

    etap = None
    for label, patterns in ETAP_PATTERNS:
        if _hits(text, patterns):
            etap = label
            break

    return {
        "problemy": "; ".join(problems[:4]),
        "etap": etap,
        "trafienia": {label: h for h, label in scored},
    }
