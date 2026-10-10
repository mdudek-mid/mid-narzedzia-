---
name: kryteria-oferty
description: Kryteria oceny ofert w przetargach MiD – symulacja punktacji (cena i kryteria pozacenowe), próg ceny wygrywającej, scenariusze przed i po otwarciu ofert, rażąco niska cena, dokumenty do kryteriów, które nie podlegają uzupełnieniu.
---

# Kryteria oceny ofert (MiD)

Używaj, gdy pytanie dotyczy punktów, a nie samego spełnienia warunków. Typowe prośby: „ile punktów dostaniemy”, „jaką cenę trzeba dać, żeby wygrać”, „czy opłaca się wskazać projektanta z 6 dokumentacjami”, „kto wygra po otwarciu”, „co w kryteriach nie podlega uzupełnieniu”, „zrób arkusz do symulacji”.

Narzędzie `kryteria_tool.py` liczy punkty dokładnie wg wzorów z SWZ (z zaokrągleniem do 0,01 pkt i rozstrzyganiem remisu), szuka ceny progowej bisekcją i tworzy arkusz XLSX z formułami. Wzory i progi przepisujesz ty, z SWZ, do pliku specyfikacji JSON.

## Instalacja (≈5 s)

```bash
rm -rf /tmp/mid-narzedzia && git clone -q --depth 1 https://github.com/mdudek-mid/mid-narzedzia- /tmp/mid-narzedzia
mkdir -p ~/.local/kryteria && cp /tmp/mid-narzedzia/kryteria/kryteria_tool.py ~/.local/kryteria/
pip install -q --break-system-packages openpyxl
K=~/.local/kryteria/kryteria_tool.py; python3 -I $K --help
```

## Zasady

- **Wzory tylko z SWZ.** Każde kryterium w specyfikacji ma źródło (`zrodlo`: „SWZ pkt 17 s. 9”). Gdy SWZ i ogłoszenie się różnią albo wzór jest niejednoznaczny, opisz to w odpowiedzi i zaproponuj pytanie do zamawiającego (termin na pytania: skill analiza-swz, `terminy`).
- **Deklaracje konkurencji są zwykle nieznane** (informacja z otwarcia podaje tylko ceny). Domyślnie narzędzie zakłada, że konkurenci mają maksimum punktów pozacenowych (`--konkurencja max`, wariant ostrożny). Zawsze napisz, jakie założenie przyjąłeś. Dla porównania możesz pokazać `--konkurencja min`.
- **Deklaracje MiD** bierz z bazy (skill referencje-i-kadra, np. `doswiadczenie K01 --typ most --tylko-projekty`) i z decyzji Marcina (termin, gwarancja). Liczbę „dokumentacji” licz wg definicji z SWZ.
- **Dokumenty do kryteriów** (np. zał. 1A, wykaz doświadczenia osoby, deklaracja terminu w formularzu) są treścią oferty. Zwykle nie podlegają uzupełnieniu ani zmianie (art. 223 ust. 1 Pzp; często SWZ mówi to wprost). Błąd = 0 pkt. Wypisz je jako listę kontrolną.
- Podawaj fakty i progi. Decyzję o starcie i cenie podejmuje Marcin.

## Przebieg

**1. Kryteria z SWZ.** Na tekstach z `swz_tool.py teksty` (skill analiza-swz):
```bash
python3 -I $K swz $KAT/txt          # strony z kryteriami, "nie podlega uzupełnieniu", 0 pkt, remis, dokumenty
python3 -I $K wzor --out kryteria.json
```
Przeczytaj wskazane strony i zapisz w specyfikacji:
```json
{"postepowanie": "nr 34 ZDW Katowice, zad. 1", "zrodlo": "SWZ pkt 17 s. 9-10", "remis": "cena",
 "kryteria": [
  {"id": "C", "nazwa": "Cena brutto", "typ": "cena", "waga": 60},
  {"id": "B", "nazwa": "Doświadczenie projektanta mostowego (zał. 1A)", "typ": "progi", "waga": 40,
   "progi": [[2,0],[3,10],[4,20],[5,30],[6,40]], "ponizej": 0, "dokument": "zał. 1A", "uzupelnienie": false}],
 "oferty": [{"wykonawca": "MiD", "my": true, "C": 350000, "B": 6},
            {"wykonawca": "RAFINS", "C": 307377, "B": null}]}
```
Typy: `cena` (Cmin/C × W), `cena_liniowa`, `min` (np. termin: Xmin/X × W, `dolna` = wartość minimalna liczona przez zamawiającego), `max` (np. gwarancja: X/Xmax × W, `limit`), `liniowy` (`od`, `do`, `pkt_od`, `pkt_do`), `progi`, `tak_nie`, `punkty` (np. ocena koncepcji przez komisję). `zaokraglenie` (domyślnie 2). `remis: "cena"`: przy równej sumie wygrywa niższa cena. Zgodnie z art. 248 Pzp decyduje najpierw kryterium o najwyższej wadze, a dopiero potem oferty dodatkowe. Pełny remis (te same punkty i cena) narzędzie liczy na niekorzyść MiD.

**2. Przed otwarciem** — ile można zaoferować przy danych punktach:
```bash
python3 -I $K scenariusze kryteria.json --budzet 400000 --udzialy 0.6,0.7,0.8,0.9
```
Wiersze to poziomy punktów pozacenowych MiD, kolumny to cena najtańszego konkurenta jako udział kwoty na sfinansowanie (z maksimum punktów, gdy `--konkurencja max`). Udziały dobierz z rynku: `swz_tool.py rynek` i `WIEDZA_Rynek.md` w mid-przetargi (np. w ZDW Katowice najniższe oferty wynosiły 67–81% mediany). Reguła dla Cmin/C × 60: 10 pkt przewagi pozwala być droższym o 20%, a 10 pkt straty wymaga ceny niższej o 16,7%.

**3. Po otwarciu** (informacja z otwarcia ofert):
```bash
python3 -I $K otwarcie "$KAT/txt/info z otwarcia ofert.txt" --out oferty.json   # rozbiór heurystyczny: sprawdź nazwy i liczbę ofert
# wklej oferty do specyfikacji (deklaracje konkurencji: null, chyba że są w informacji)
python3 -I $K ocena kryteria.json --budzet 400000
python3 -I $K prog kryteria.json --budzet 400000
```
`ocena` pokazuje ranking i rażąco niską cenę: cena niższa o ≥30% od średniej wszystkich ofert albo od wartości zamówienia z VAT oznacza obowiązkowe wezwanie do wyjaśnień (art. 224 ust. 2 pkt 1 Pzp). Kwota na sfinansowanie zastępuje tu wartość zamówienia, z zastrzeżeniem. Pokazuje też oferty powyżej budżetu. `prog` podaje najwyższą cenę wygrywającą MiD i cenę, przy której MiD wyprzedza każdego konkurenta. Ostrzega, jeśli przy cenie progowej grozi wezwanie RNC. Wtedy potrzebna jest kalkulacja (skill wycena-oferty).

**4. Arkusz dla Marcina:**
```bash
python3 -I $K xlsx kryteria.json --out Symulacja_punktacji_nr34.xlsx
```
Arkusz „Symulacja” ma żółte komórki wejściowe (ceny, deklaracje) i formuły punktów, sumy i miejsca. Arkusz „Progi” zawiera progi z chwili eksportu. Sprawdź przeliczenie: `soffice --headless --convert-to csv ...` i porównaj z `ocena`.

**5. Odpowiedź** (krótko):
- tabela kryteriów: waga, wzór, dokument, czy podlega uzupełnieniu;
- punkty MiD przy planowanej deklaracji i cena progowa (zł i % najniższej ceny / mediany / budżetu), z przyjętym założeniem o konkurencji;
- ile „kosztuje” każdy poziom punktów pozacenowych (tabela scenariuszy);
- ryzyka: RNC, cena powyżej budżetu, dokumenty bez możliwości uzupełnienia (lista kontrolna z wymaganymi danymi dla każdej pozycji);
- pliki: arkusz XLSX (SendUserFile).
