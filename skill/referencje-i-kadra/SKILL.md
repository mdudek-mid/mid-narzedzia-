---
name: referencje-i-kadra
description: Baza referencji i kadry Pracowni MiD – dobór usług i osób do warunków udziału, wykaz usług, wykaz osób i załącznik o doświadczeniu projektanta (kryterium), wpisane w formularze DOCX zamawiającego.
---

# Referencje i kadra (MiD)

Używaj, gdy trzeba wykazać spełnienie warunków udziału albo zdobyć punkty za doświadczenie. Typowe prośby: „czy spełniamy warunek doświadczenia”, „dobierz referencje do nr 34”, „wypełnij wykaz usług / wykaz osób”, „ile dokumentacji ma projektant mostowy do zał. 1A”, „kogo wskazać jako sprawdzającego”.

Narzędzie `kadra_tool.py` przeszukuje bazę, ocenia każdą referencję i osobę względem kryteriów (✓ spełnia, ✗ nie spełnia, ? brak danych w bazie, ! uwaga formalna) i wpisuje wybrane pozycje w formularz zamawiającego. Decyzję, co wpisać, podejmujesz ty na podstawie dokładnego brzmienia SWZ. Ostatnie słowo ma Marcin.

## Instalacja (≈10 s)

```bash
rm -rf /tmp/mid-narzedzia && git clone -q --depth 1 https://github.com/mdudek-mid/mid-narzedzia- /tmp/mid-narzedzia
mkdir -p ~/.local/kadra && cp /tmp/mid-narzedzia/kadra/kadra_tool.py ~/.local/kadra/
pip install -q --break-system-packages python-docx openpyxl
T=~/.local/kadra/kadra_tool.py
```

**Baza jest prywatna**: katalog `baza_mid/` w repo `mdudek-mid/mid-przetargi` (`referencje.json`, `kadra.json`, `zrodla/`). Dołącz repo narzędziem do repozytoriów (odczyt) i wskaż katalog:
```bash
export MID_BAZA=/sciezka/mid-przetargi/baza_mid     # albo: python3 $T --baza KATALOG <komenda>
python3 $T stan
```
Bez bazy narzędzie nie ma czego dobierać. Nie wymyślaj referencji ani osób; napisz, że baza nie jest dostępna.

## Zasady

- **Dane osobowe kadry** (nazwiska, numery uprawnień, historia zatrudnienia) trafiają tylko do dokumentów oferty MiD. Nie publikuj ich, nie wklejaj do publicznych repo ani artefaktów udostępnianych poza firmę.
- **Niczego nie dopisuj z głowy.** Brak danych w bazie = `[DO UZUPEŁNIENIA]` w formularzu i pozycja na liście braków.
- **`do_weryfikacji`**: rekord ma rozbieżności między dokumentami (wartość, daty), pochodzi tylko z nazwy skanu albo z doświadczenia JDG. Przed złożeniem oferty sprawdź go w dowodzie (pole `dowody`, ścieżka w OneDrive) i napisz Marcinowi, którą wersję przyjąłeś.
- **JDG a spółka.** Ofertę składa zwykle Pracownia Projektowa MiD Sp. z o.o. Usługi wykonane przez JDG „Pracownia Projektowa MiD Marcin Dudek” (pole `podmiot` zawiera „JDG”) nie są doświadczeniem spółki. Można je wykazać tylko przez udostępnienie zasobów (art. 118 Pzp): zobowiązanie podmiotu udostępniającego, a przy doświadczeniu ten podmiot musi faktycznie wykonać część usług, do których zdolność jest wymagana (art. 118 ust. 2). Narzędzie oznacza takie pozycje „!”; `--tylko-spolka` je pomija. Zawsze najpierw szukaj kompletu w doświadczeniu spółki.
- **Uprawnienia**: numer i datę potwierdź w zaświadczeniu izby lub kopii decyzji. Okres „X lat od uzyskania uprawnień” licz na dzień składania ofert (`--od`).
- **Formularz zawsze obejrzyj** po wypełnieniu (PDF → PNG). Wzory zamawiających mają scalone komórki, nagłówki wielopoziomowe i dziwne formatowanie.
- Nie nadpisuj plików w OneDrive. Wynik zapisuj pod nową nazwą (np. `7_wykaz_uslug_MiD.docx`) w katalogu roboczym i wyślij użytkownikowi.
- Załącznik do kryterium pozacenowego (np. 1A „doświadczenie projektanta”) zwykle **nie podlega uzupełnieniu**. Błąd to 0 pkt. Każdą pozycję sprawdź z definicją z SWZ (co jest „obiektem mostowym”, czy liczy się PB, czy dopiero PB+PW, czy osobne obiekty w jednej umowie liczą się osobno, okres).

## Przebieg

**1. Warunki z SWZ.** Wypisz dosłownie (z kartą `Analiza_SWZ.md`, skill analiza-swz): rodzaj usługi i obiektu, zakres (PB, PW, ZRID…), okres („w ciągu ostatnich N lat przed upływem terminu składania ofert”), liczba usług, wartość (każda czy łączna; brutto czy netto), parametry (długość, rozpiętość, klasa drogi, nad czym), osoby (specjalność, „bez ograniczeń”, lata od uprawnień lub doświadczenie zawodowe), a przy częściach — który zestaw wymagań obowiązuje przy ofercie na kilka.

**2. Usługi.**
```bash
python3 $T uslugi --typ most,wiadukt,kladka --dowolny-zakres PB,PW --lat 3 --od 2026-10-09 --min-wartosc 120000
#   inne filtry: --zakres PB,PW (wszystkie), --netto, --dlugosc-min 50, --rozpietosc-min 30, --klasa-drogi G,
#                --nad kolej, --szukaj "Pruszcz", --tylko-spolka, --max 15
python3 $T pokaz R012          # pelny rekord z cytatem z referencji i rozbieznosciami
```
Wybierz minimalny zestaw spełniający warunek plus jedną pozycję zapasową. Preferuj: doświadczenie spółki, rekordy z cytatem i wartością z referencji, bez rozbieżności. Wartość „łączną” policz sam i podaj sumę. Gdy nic nie pasuje, poluzuj filtr, żeby zobaczyć najbliższe pozycje, i napisz, czego brakuje (konsorcjum, podmiot udostępniający, podwykonawca).

**3. Osoby.**
```bash
python3 $T osoby --specjalnosc mostowa --bez-ograniczen --lat-od-uprawnien 5 --od 2026-10-09
#   --rownowazne      uprawnienia konstrukcyjno-budowlane wg starych przepisow traktuj jak mostowe/drogowe
#   --dokumentacje-min 3 --typ most --lat 10   gdy warunek dotyczy liczby opracowan
#   --wszyscy         takze osoby znane tylko z list projektantow w referencjach
```
Projektant i sprawdzający to różne osoby. Sprawdź podstawę dysponowania (personel własny / umowa o współpracy / udostępnienie). Kategoria „tylko w referencjach” oznacza, że osoba nie była dotąd wskazywana w wykazach MiD — zapytaj Marcina, czy jest dostępna.

**4. Kryterium doświadczenia (np. zał. 1A).**
```bash
python3 $T doswiadczenie K01 --typ most --tylko-projekty --lat 10 --od 2026-10-09
```
Lista łączy doświadczenie osoby z wykazów i referencje, w których jest na liście projektantów (bez duplikatów). `--tylko-projekty` pomija ekspertyzy, przeglądy i nadzory. Policz punkty wg tabeli z SWZ (skill kryteria-oferty). Wybierz pozycje z pełnymi danymi (inwestor, parametry, data) i z dowodem.

**5. Formularze.**
```bash
python3 $T wykaz-uslug R012,R008,R011 --form "7 wykaz usług.docx" --out 7_wykaz_uslug_MiD.docx
python3 $T wykaz-osob "K01:projektant branży mostowej,K17:sprawdzający branży mostowej" --od 2026-10-09 \
        --form "8 wykaz osob.docx" --out 8_wykaz_osob_MiD.docx
python3 $T doswiadczenie K01 --typ most --tylko-projekty --form "Zał_1A.docx" --wybierz 1,2,3,4,5,6 --out Zal_1A_MiD.docx
soffice --headless --convert-to pdf 7_wykaz_uslug_MiD.docx && pdftoppm -r 60 -png 7_wykaz_uslug_MiD.pdf podglad
```
Narzędzie rozpoznaje kolumny po nagłówkach i wypisuje mapowanie (`kolumny: {...}`). Sprawdź je. Bez `--form` drukuje treść wierszy w Markdown do ręcznego wpisania (gdy wzór jest PDF albo tabela nietypowa). Formularza nie wypełnisz poprawnie bez obejrzenia podglądu: sprawdź Lp., czy rekordy nie rozjechały się między wierszami, komórki scalone (data / miejsce wykonania) i `[DO UZUPEŁNIENIA]`. „Miejsce wykonania” i dane wykonawcy w nagłówku formularza uzupełnij ręcznie.

**6. Odpowiedź dla Marcina** (krótko, tabelą):
- warunek → czy spełniony, którymi pozycjami (ID, wartość, data zakończenia), suma, jeśli liczy się łączna;
- osoby → kto, uprawnienia, lata na dzień składania ofert, podstawa dysponowania;
- kryterium → liczba pozycji i punkty;
- do weryfikacji: rozbieżności, JDG (art. 118 + zobowiązanie), brakujące dowody (skany do załączenia);
- pliki: wypełnione formularze (SendUserFile).

## Utrzymanie bazy

Źródła w OneDrive (konektor Microsoft 365, tylko odczyt): `_PRZETARGI_/skany_referencji/MiD` (referencje), `_PRZETARGI_/Projektanci - doświadczenie/` (dane branżystów), złożone wykazy usług i osób w teczkach `_PRZETARGI_/!_SWZ_RRRR/...`. Schemat:
- `referencje.json → referencje[]`: `id` (R001…), `etykieta`, `nazwa`, `zamawiajacy_nazwa/adres`, `inwestor_koncowy`, `rola_mid`, `droga{numer, kategoria, klasa}`, `obiekty[{typ, nad, dlugosc_calkowita_m, rozpietosc_max_m, klasa_obciazenia, konstrukcja}]`, `zakres[]` (PB, PW, STWiORB, kosztorys, ZRID, projekt_rozbiorki, ekspertyza…), `wartosc_brutto_pln`, `wartosc_netto_pln`, `data_rozpoczecia`, `data_zakonczenia`, `projektanci[{osoba, rola}]`, `dowody[]`, `wykazy[]`, `podmiot` (JDG/spółka), `cytat`, `rozbieznosci[]`, `pewnosc`, `do_weryfikacji`, opcjonalnie `miejsce`.
- `kadra.json → osoby[]`: `id` (K01…), `imie_nazwisko`, `tytul`, `kategoria`, `uprawnienia[{specjalnosc, zakres, rodzaj, numer, data_wydania}]`, `podstawa_dysponowania`, `firma_zewnetrzna`, `funkcje_w_wykazach[]`, `doswiadczenie[{zadanie, rola, zamawiajacy, data, zakres}]`, `lata_doswiadczenia_deklarowane`, `wyksztalcenie`, `zrodla[]`, `uwagi`, `do_weryfikacji`.

Zmiany w bazie tylko na prośbę Marcina albo po jego potwierdzeniu. Nowy rekord dostaje kolejne ID. Rekordów nie usuwaj; nieaktualne oznacz `"archiwalna": true` z uwagą. Wpisz źródło (`dowody`, `cytat`). Commit do `mid-przetargi` dotyczy tylko katalogu `baza_mid/`.

Ok. 55 referencji jest znanych tylko z nazwy skanu (`pewnosc: niska`). Gdy rozmowa jest połączona z komputerem Marcina, pobierz skany z `skany_referencji/MiD` (stage), odczytaj OCR (`swz_tool.py teksty`, skill analiza-swz) i uzupełnij pola. `python3 $T stan` pokazuje luki. `python3 $T xlsx --out Baza_referencji_i_kadry.xlsx` eksportuje bazę do przeglądu w Excelu.
