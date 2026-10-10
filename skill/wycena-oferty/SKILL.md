---
name: wycena-oferty
description: Wycena prac projektowych Pracowni MiD do oferty przetargowej – kalkulacja metodą MiD (rbg × stawka, podwykonawcy z narzutem, rezerwa), scenariusze A/B/C, cel cenowy Marcina, rozbicie na formularz cenowy zamawiającego, arkusz XLSX i opis MD w stylu MiD.
---

# Wycena oferty (MiD)

Używaj, gdy trzeba policzyć cenę oferty na prace projektowe (dokumentacja, koncepcja, nadzór autorski, P&B dla wykonawcy robót). Typowe prośby: „wyceń nr 34”, „ile to kosztuje”, „rozpisz na formularz cenowy”, „Marcin chce 380 tys. netto, rozłóż”, „próg bólu”, „przygotuj wyjaśnienia RNC”.

Narzędzie `wycena_tool.py` liczy kalkulację ze specyfikacji JSON, robi scenariusze i rozbicie na formularz zamawiającego, sprawdza formalności i tworzy dokumenty. Zakres prac i liczby rbg ustalasz ty, z OPZ i wzorców, a cenę zatwierdza Marcin.

## Instalacja (≈5 s)

```bash
rm -rf /tmp/mid-narzedzia && git clone -q --depth 1 https://github.com/mdudek-mid/mid-narzedzia- /tmp/mid-narzedzia
mkdir -p ~/.local/wycena && cp /tmp/mid-narzedzia/wycena/wycena_tool.py ~/.local/wycena/
pip install -q --break-system-packages openpyxl
W=~/.local/wycena/wycena_tool.py; python3 -I $W --help
```
Dane prywatne w repo `mdudek-mid/mid-przetargi` (dołącz do sesji, odczyt): `baza_mid/wzorce_wycen.json` (pozycje wcześniejszych wycen MiD), `WIEDZA_Rynek.md`, `wyniki_postepowan.json`, `karta_wyceny.py`. Parametry metody (`baza_mid/parametry_wyceny.json`) i wzorce wycen to know-how firmy. Nie umieszczaj ich w publicznych repo ani w udostępnianych artefaktach.

## Metoda MiD (zasady Marcina)

Liczby metody, czyli stawki, narzut, rezerwę, parametry jnp i P&B, oraz spisane zasady Marcina trzymamy prywatnie w `baza_mid/parametry_wyceny.json` (pola `parametry` i `zasady`). Narzędzie wczytuje je samo (`--baza` albo `MID_BAZA`). Przed wyceną przeczytaj `zasady`.

- **Praca własna**: rbg × stawka MiD. Nadzór autorski i drobne zmiany można liczyć metodą jnp (pole `jnp`).
- **Branże zlecane podwykonawcom**: wg ich ofert + narzut (`"podwykonawca": true`). Badania, mapy, opłaty, delegacje i wydruki liczone są po kosztach, chyba że pozycja ma `narzut`.
- **A — pełna kalkulacja**: koszt kalkulacyjny + rezerwa ryzyka.
- **B — rekomendowany**: wstępnie koszt kalkulacyjny. Po kalibracji rynkowej i decyzji Marcina używasz celu (`--cel` netto albo `--cel-brutto`). Pozycje **sztywne** (PB/PT/PW mostowe, geodezja, decyzje, geologia, uzgodnienia z koleją) zostają bez zmian. Tniesz **miękkie** (kompletacja, organizacja ruchu, nadzór i pobyty, koordynacja, narady, egzemplarze), maksymalnie o 40%. Głębsze cięcie narzędzie rozkłada też na sztywne i ostrzega.
- **C — próg bólu**: niższa stawka, koszty zewnętrzne po kosztach własnych, bez rezerwy. Poniżej wychodzimy na minus.
- **Nie wykazujemy pozycji warunkowych.** Wliczamy je w cenę (`"warunkowa": true` tylko dla opisu).
- **P&B** (MiD jako projektant u wykonawcy robót): `"formula": "pib"` (współczynnik `wsp_pib`) i kontrola udziału dokumentacji w wartości robót (`wartosc_robot_netto`).
- **Prawo opcji** podawaj osobno (`"opcja": true` w formularzu). Zawsze podawaj netto i brutto (VAT 23%). Baza wyników przetargów jest brutto, a wyceny MiD netto.
- **Wadium**: poniżej 10 tys. zł przelew, od 10 tys. zł gwarancja wadialna. Status MŚP do formularzy: małe przedsiębiorstwo.
- **Najpierw warunki udziału, potem wycena.** Gdy warunek doświadczenia wykracza poza profil mostowy (drogi klasy G na km, sieci), zatrzymaj się i zapytaj Marcina o referencje albo konsorcjanta (skill referencje-i-kadra). Narzędzie przypomina o tym, dopóki spec nie ma `"warunki_sprawdzone": true`.
- Przed wyceną i przed złożeniem oferty sprawdź aktualność materiałów (zmiany SWZ, odpowiedzi). Skill analiza-swz.
- Nie zapisuj w OneDrive i nie wysyłaj maili. Pliki wyślij użytkownikowi, a Marcin odkłada je do `_wycena_MiD`.

## Przebieg

**1. Zakres.** Z OPZ/PFU, wzoru umowy i odpowiedzi na pytania (karta Analiza_SWZ.md) wypisz elementy dokumentacji z parametrami obiektów, decyzje, badania, liczbę pobytów nadzoru, egzemplarze i terminy. Każda pozycja kalkulacji ma uzasadnienie w `uwagi` z odnośnikiem [OPZ s. N].

**2. Formularz cenowy zamawiającego** (ZZER, TOP, formularz ofertowy) przepisz do `formularz` dokładnie z numeracją pozycji. Obsługiwane są:
- ryczałty i ilości (`ilosc`, `jm`; cena jednostkowa liczona jest z wartości);
- pozycje procentowe (np. „kwota tymczasowa 5%” → `"procent": 0.05`, liczona od sumy pozostałych);
- stałe udziały w grupie (np. ZDW Kraków: koncepcja 40% / DUŚ 60% → `"grupa"`, `"udzial"`);
- opcje.
Pozycja formularza z wartością 0 jest zgłaszana, bo często oznacza odrzucenie oferty.

**3. Kalkulacja.**
```bash
python3 -I $W wzor --out wycena.json                     # szablon specyfikacji
python3 -I $W wzorce --baza $MID_BAZA --szukaj "geotech|odwiert"   # jak liczyliśmy podobne pozycje
python3 -I $W wzorce --baza $MID_BAZA --lista             # wcześniejsze wyceny i ich wyniki
```
Pozycja `praca`: `id`, `zakres`, `rbg` (albo `jnp`), `poz` (numer w formularzu albo `{"11": 0.6, "12": 0.4}`), `sztywna`, `uwagi`. Pozycja `zewnetrzne`: `zakres`, `kwota`, `podwykonawca`, `poz`, `sztywna`. Kwoty podwykonawców bierz z ich ofert. Jeśli ich nie ma, przyjmij szacunek ze wzorców i oznacz w `uwagi` „szacunek — potwierdzić ofertą”. Klasyfikację sztywna/miękka ustaw jawnie. Bez niej narzędzie zgaduje po słowach i wypisuje, co zgadło.

**4. Liczenie i rynek.**
```bash
python3 -I $W licz wycena.json                  # A/B/C, formularz, kontrole (warunki, wadium, zera, RNC, budżet, P&B)
```
Kalibracja rynkowa:
- `karta_wyceny.py` (mid-przetargi) i `swz_tool.py rynek` (skill analiza-swz): ceny podobnych postępowań, profil zamawiającego, konkurenci;
- skill kryteria-oferty (`scenariusze --budzet`, po otwarciu `prog`): cena, przy której MiD wygrywa przy danych punktach.
Wpisz w `rynek`: `budzet_brutto`, `oferty_brutto` (lista, po otwarciu) i `opis`. Następnie przedstaw Marcinowi A/B/C na tle rynku. Po jego decyzji:
```bash
python3 -I $W licz wycena.json --cel 380000            # albo --cel-brutto 467400
```

**5. Dokumenty.**
```bash
python3 -I $W xlsx wycena.json --cel 380000 --out Wycena_<zamawiajacy>_<temat>_RRRR-MM.xlsx
python3 -I $W md   wycena.json --cel 380000 --out Wycena_<zamawiajacy>_<temat>_RRRR-MM.md
```
XLSX ma arkusze „Formularz cenowy” i „Kalkulacja” w stylu MiD: Arial, nagłówki 1F4E79, niebieskie pola do edycji, formuły sum, VAT, scenariuszy i % rynku, wiersz kontrolny „formularz − kalkulacja”, stopka „Dokument roboczy – wewnętrzny, nie do publikacji”. Arkusz Kalkulacja jest też podstawą wyjaśnień rażąco niskiej ceny (art. 224 Pzp). Sprawdź przeliczenie formuł:
```bash
soffice --headless --convert-to csv:"Text - txt - csv (StarCalc)":44,34,76,1,,0,false,true,false,false,false,-1 Wycena.xlsx
```
Porównaj wynik z `licz` i obejrzyj stronę 1 (PDF → PNG). MD to opis dla Marcina w układzie wycen MiD: status, dane postępowania, kalkulacja, warianty, rozbicie na formularz, benchmark, ryzyka, kontrole, do zrobienia.

**6. Odpowiedź** (krótko): kalkulacja bazowa (rbg, koszty zewnętrzne), A/B/C netto i brutto, pozycja na tle rynku (miejsce cenowo, % min/mediany/budżetu), ryzyka kosztowe (badania, uzgodnienia, decyzje), formalności (wadium, zera w formularzu, opcje), pliki XLSX i MD. Gdy cena wygrywająca (skill kryteria-oferty, `prog`) jest poniżej C, napisz to wprost i pokaż, ile trzeba by ciąć (`licz --cel-brutto`).
