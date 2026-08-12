# Monitor orzeczeń — środki unijne (aktualizacja mastera)

Automatyczny monitoring orzecznictwa dotyczącego środków unijnych (fundusze UE, KPO, Czyste Powietrze, PFR/Tarcza Finansowa, pomoc publiczna), dopisujący nowe rekordy bezpośrednio do pliku **„Baza_orzecznictwa_po_stronie_beneficjenta___MASTER"** na Google Drive — zgodnie z jego strukturą i konwencjami. Uruchamiany przez GitHub Actions w dni robocze, bez własnego serwera.

## Jak to działa

1. Skrypt pobiera master (`.xlsx`) z Google Drive.
2. Zbiera kandydatów z **czterech źródeł**:
   - **CBOSA po symbolu sprawy** — wszystkie orzeczenia sądów administracyjnych z kategorii **6559** (środki unijne), niezależnie od sformułowań w treści,
   - **API SAOS** (saos.org.pl) wg grup fraz z `config.yaml` — sądy administracyjne ORAZ sądy powszechne i SN (frazy cywilne: umowa o dofinansowanie, powództwa o zwrot, weksle, subwencje PFR, kary umowne, bezpodstawne wzbogacenie),
   - **Portal Orzeczeń Sądów Powszechnych** (orzeczenia.ms.gov.pl) — frazy cywilne,
   - **Baza orzeczeń Sądu Najwyższego** (www.sn.pl) — frazy cywilne, jako uzupełnienie SAOS.
3. **Filtr wstępny po kategorii sprawy:** kandydatom z SAOS (sądowoadministracyjnym) weryfikuje symbol po sygnaturze w CBOSA — sprawy spoza symboli z listy są odrzucane przed dopisaniem (wynik weryfikacji trafia do Notatki wewnętrznej). Sprawy cywilne PFR nie mają symboli CBOSA i przechodzą bez tego filtra.
4. Deduplikuje po znormalizowanej **Sygnaturze** i **Kluczu sprawy** względem arkusza „Baza orzeczeń" (zasada jednego rekordu z Instrukcji).
5. Dopisuje nowe rekordy w pełnej strukturze mastera:
   - **ID** — kontynuacja numeracji ORZ-NNNN,
   - **Rok** — z końcówki sygnatury (jak w istniejących rekordach),
   - **Jurysdykcja / Rodzaj finansowania / Problemy prawne** — z mapowania grup fraz,
   - **Status publikacji** = „Do weryfikacji", **Status weryfikacji** = „Metadane wstępne", **Wynik** = „Do oceny", **Etap sporu** = „Inne / do klasyfikacji",
   - **Klucz sprawy** = `PL-{sygnatura}-{rok}`,
   - fragment treści i trafiona fraza lądują w **Notatce wewnętrznej** z prefiksem `[AUTO SAOS data]`,
   - pola merytoryczne (teza publiczna, znaczenie, rozstrzygnięcie, podstawa prawna) pozostają **puste** — zgodnie z Instrukcją uzupełniane po lekturze pełnej treści.
6. Rozszerza autofiltr i walidacje danych na nowe wiersze, przelicza formuły Dashboardu (LibreOffice) i odsyła plik na Drive. Historia wersji Drive zostaje zachowana — każdą zmianę można cofnąć.

Skrypt **niczego nie modyfikuje ani nie usuwa** w istniejących rekordach i pozostałych arkuszach (Dashboard, Kolejka publikacji, Duplikaty, Słowniki, Instrukcja, Źródła wtórne, Kandydaci z publikacji, Korpusy, Audyt) — wyłącznie dopisuje wiersze na końcu „Bazy orzeczeń". Przed dopisaniem weryfikuje zgodność nagłówków wszystkich 36 kolumn i przerywa pracę, jeśli struktura pliku się zmieniła.

## Konfiguracja krok po kroku

### 1. Konto serwisowe Google (jednorazowo, ok. 10 minut)

1. Wejdź na [console.cloud.google.com](https://console.cloud.google.com) i utwórz nowy projekt (np. `nsa-monitor`).
2. W menu **APIs & Services → Library** wyszukaj **Google Drive API** i kliknij **Enable**.
3. W **APIs & Services → Credentials** kliknij **Create Credentials → Service account**. Nadaj nazwę (np. `nsa-monitor-bot`), pomiń role, zakończ.
4. Otwórz utworzone konto serwisowe → zakładka **Keys → Add key → Create new key → JSON**. Pobierze się plik klucza — zachowaj go bezpiecznie.
5. Skopiuj **adres e-mail konta serwisowego** (kończy się na `...iam.gserviceaccount.com`).

### 2. Udostępnienie pliku Excel

1. Master już jest na Google Drive — upewnij się tylko, że pozostaje plikiem `.xlsx` (bez konwersji do Arkuszy Google; skrypt pracuje na formacie Excel).
2. Kliknij plik prawym przyciskiem → **Udostępnij** → wklej e-mail konta serwisowego → uprawnienie **Edytujący**.
3. Skopiuj **ID pliku** z adresu URL: `https://drive.google.com/file/d/`**`TO_JEST_ID`**`/view`.

### 3. Repozytorium GitHub

1. Utwórz **prywatne** repozytorium i wgraj zawartość tego katalogu.
2. W **Settings → Secrets and variables → Actions → New repository secret** dodaj dwa sekrety:
   - `GDRIVE_SERVICE_ACCOUNT_JSON` — cała zawartość pobranego pliku klucza JSON (wklej tekst),
   - `GDRIVE_FILE_ID` — ID pliku z kroku 2.
3. W zakładce **Actions** włącz workflowy (GitHub może o to poprosić przy pierwszym uruchomieniu) i uruchom ręcznie workflow **Monitor orzeczeń NSA/WSA** przyciskiem **Run workflow**, żeby przetestować.

Harmonogram: dni robocze o 05:30 UTC (07:30 czasu letniego w Warszawie). Zmienisz go w `.github/workflows/monitor.yml` (linia `cron`).

### Deduplikacja spraw cywilnych

Sygnatury sądów powszechnych powtarzają się między sądami (wiele „I C 100/23" w kraju), dlatego dla jurysdykcji cywilnej (poza SN) deduplikacja i „Klucz sprawy" uwzględniają także sąd: `PL-I C 100/23-2023-Sąd Okręgowy w Poznaniu`. Sprawy sądowoadministracyjne i SN zachowują dotychczasowy format `PL-{sygnatura}-{rok}`.

### Moduły PO SP i SN — zastrzeżenie

Portal Orzeczeń Sądów Powszechnych i baza SN nie mają API — moduły `ms_client.py` i `sn_client.py` parsują HTML elastycznymi wzorcami (sygnatury, daty, nazwy sądów) i mogą wymagać korekty selektorów po pierwszym realnym uruchomieniu lub po zmianie struktury stron. Awaria któregokolwiek modułu nie przerywa przebiegu — skrypt loguje błąd i kontynuuje z pozostałymi źródłami. Wyłączenie: `ms_portal.enabled: false` / `sn_portal.enabled: false`. Orzecznictwo SN jest równolegle indeksowane w SAOS, więc nawet przy awarii modułu SN sprawy kasacyjne będą wpływać.

## Backfill: zaciągnięcie historii od 1 stycznia 2023

Jednorazowe uzupełnienie bazy o całe orzecznictwo tematyczne od `backfill.since` (domyślnie **2023-01-01** — Instrukcja mastera wskazuje lukę 2022–2025):

- **z GitHuba:** zakładka Actions → „Monitor orzeczeń" → **Run workflow** → tryb **backfill**,
- **lokalnie:** `python src/monitor.py --backfill`.

W tym trybie skrypt przechodzi pełną paginację SAOS (do 1000 wyników na frazę), pobiera **pełny tekst każdego nowego orzeczenia** i wykonuje klasyfikację słownikową:

- **Problemy prawne** — do 4 wartości ze słownika mastera wg liczby trafień wzorców w treści (np. „proporcjonalność i taryfikator; nieprawidłowość i szkoda; zamówienia publiczne"),
- **Etap sporu** — jedna wartość ze słownika wg reguł pierwszeństwa (np. korekta/art. 207 → „Korekta i postępowanie zwrotowe", protest/ocena → „Nabór i ocena projektu", weksel/subwencja → „Postępowanie cywilne / wekslowe"),
- diagnostyka trafień trafia do **Notatki wewnętrznej** (`[AUTO SAOS data] … klasyfikacja słownikowa wg trafień: …`).

Klasyfikacja jest heurystyczna — to podpowiedź przy triażu, nie ocena merytoryczna. Dlatego **Wynik dla beneficjenta pozostaje „Do oceny"**, a teza publiczna, znaczenie i rozstrzygnięcie są puste (Instrukcja, reguła 3). Status weryfikacji: „Metadane wstępne".

Backfill można bezpiecznie powtarzać i przerywać: deduplikacja po sygnaturze pomija wszystko, co już jest w masterze, więc kolejne uruchomienie dopisze tylko brakujące rekordy. Orientacyjny czas: przy kilkuset nowych orzeczeniach ok. 10–30 minut (limit workflowu: 120 min). Ta sama analiza treści działa też w trybie dziennym (`analyze_in_standard_mode: true`), gdzie nowych orzeczeń jest niewiele.

## Dostosowanie

Wszystko istotne jest w `config.yaml`:

- **`keyword_groups`** — grupy fraz z przypisanym „Rodzajem finansowania", opcjonalnie „Problemami prawnymi" (wartości muszą pochodzić ze słowników mastera) i typami sądów SAOS; listę warto dostroić po pierwszych uruchomieniach,
- **`lookback_days`** — jak daleko wstecz sięga każde uruchomienie (deduplikacja i tak chroni przed duplikatami),
- **`master.defaults`** — wartości słownikowe wpisywane do nowych rekordów,
- **`backfill`** — data początkowa, limit wyników na frazę, włączenie analizy treści,
- **wzorce klasyfikacji** — słowniki wyrażeń w `src/analyze.py` (`PROBLEM_PATTERNS`, `ETAP_PATTERNS`); łatwo dodać własne wzorce pod nowe linie orzecznicze.

Ręczna praca z bazą pozostaje bez zmian: nowe rekordy wchodzą ze statusem „Do weryfikacji"/„Metadane wstępne", więc odfiltrujesz je jednym kliknięciem i klasyfikujesz (etap sporu, problemy, wynik, teza) we własnym tempie.

### Filtr po symbolu sprawy (6559)

Sprawy dotyczące środków unijnych mają w CBOSA symbol **6559** — to najprecyzyjniejszy filtr wstępny i działa dwutorowo:

- `cbosa.search_by_symbol: true` — pobieranie **wszystkich** spraw z symbolami z listy `cbosa.symbols` (łapie także orzeczenia, które nie używają żadnej z fraz),
- `symbol_filter.enabled: true` — weryfikacja symbolu kandydatów z SAOS po sygnaturze w CBOSA; sprawy spoza kategorii są odrzucane. `keep_unverified: true` zachowuje (z adnotacją „symbol nieustalony") sprawy, których symbolu nie udało się ustalić — bezpieczniej przejrzeć ręcznie niż stracić orzeczenie.

Listę symboli można rozszerzyć (np. o pokrewne podsymbole grupy 655) w obu sekcjach konfiguracji.

### Moduł CBOSA — zastrzeżenie

CBOSA nie ma API — moduł `cbosa_client.py` parsuje HTML i może przestać działać po zmianie struktury strony (skrypt wtedy loguje błąd i kontynuuje z samym SAOS). Moduł stosuje kilkusekundowe przerwy między zapytaniami, a weryfikacja symboli — `symbol_filter.delay_seconds` między sprawdzeniami. Przy backfillu z setkami kandydatów weryfikacja symboli wydłuża przebieg (ok. 3 s na sprawę z SAOS); sprawy trafione po symbolu w CBOSA nie są weryfikowane ponownie.

## Uruchomienie lokalne (test)

```bash
pip install -r requirements.txt
export GDRIVE_SERVICE_ACCOUNT_JSON="$(cat klucz.json)"
export GDRIVE_FILE_ID="..."
python src/monitor.py
```

## Struktura

```
├── config.yaml                  # grupy fraz, okno czasowe, wartości domyślne
├── requirements.txt
├── src/
│   ├── monitor.py               # główny przebieg (świadomy struktury mastera)
│   ├── saos_client.py           # klient API SAOS (paginacja + pełne teksty)
│   ├── analyze.py               # klasyfikacja słownikowa treści orzeczeń
│   ├── cbosa_client.py          # CBOSA: wyszukiwanie po symbolu + weryfikacja
│   ├── ms_client.py             # Portal Orzeczeń Sądów Powszechnych
│   ├── sn_client.py             # Baza orzeczeń SN
│   └── drive_sync.py            # pobieranie/wysyłka pliku na Drive
└── .github/workflows/monitor.yml
```

## Znane ograniczenia

- SAOS pokrywa orzecznictwo niekompletnie i z opóźnieniem względem CBOSA — monitoring wyłapie większość istotnych nowości, ale nie zastępuje ręcznej kwerendy przy badaniu konkretnej sprawy.
- Dopasowanie po frazach jest tekstowe: orzeczenie, które nie używa żadnej z fraz, nie zostanie wychwycone — listę warto iteracyjnie rozbudowywać (największa luka wg Instrukcji: lata 2022–2025 oraz KPO/Czyste Powietrze/PFR, stąd dedykowane grupy fraz).
- Jeśli na maszynie GitHub Actions zabraknie LibreOffice, formuły Dashboardu nie zostaną przeliczone w pliku — Excel/Arkusze przeliczą je same przy otwarciu, ale podgląd pliku na Drive może chwilowo pokazywać puste liczniki.
- Klasyfikacja merytoryczna (etap sporu, wynik, teza) pozostaje ręczna — skrypt celowo nie zgaduje treści orzeczeń.
