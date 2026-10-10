---
name: ekspertyza-mostu
description: Ekspertyzy techniczne obiektów mostowych Pracowni MiD – ocena elementów w skali GDDKiA 0–5 (ocena średnia i ogólna, tryby robót A/1/2/3), interpretacja badań (sklerometr, karbonatyzacja, chlorki, potencjały, rezystywność, ubytki zbrojenia), nośność użytkowa i MLC z MES, raport DOCX wg Instrukcji 02.
---

# Ekspertyza mostu (MiD)

Używaj przy ekspertyzach, opiniach technicznych i przeglądach szczegółowych mostów, wiaduktów i przepustów. Typowe prośby: „przygotuj ekspertyzę mostu w …”, „oceń elementy i podaj ocenę ogólną”, „zinterpretuj wyniki sklerometru i karbonatyzacji”, „określ nośność i oznakowanie”, „zalecenia z trybami”, „sprawdź starą ekspertyzę”.

Skill łączy trzy narzędzia:
- `ekspertyza_tool.py`: oceny, badania, rozdziały ekspertyzy;
- `mes_tool.py` (skill mes-eurokody): nośność użytkowa, MLC, RF;
- `opis_tool.py` (skill opis-techniczny): lokalizacja z mapami, punkty 1–4 Instrukcji 02, aktualne akty prawne, DOCX.

Ustalenia z oględzin, oceny i wnioski należą do autora ekspertyzy. Narzędzie pilnuje skali, spójności i formy.

## Instalacja (≈15 s)

```bash
rm -rf /tmp/mid-narzedzia && git clone -q --depth 1 https://github.com/mdudek-mid/mid-narzedzia- /tmp/mid-narzedzia
mkdir -p ~/.local/ekspertyza ~/.local/opis ~/.local/mes
cp /tmp/mid-narzedzia/ekspertyza/ekspertyza_tool.py ~/.local/ekspertyza/
cp /tmp/mid-narzedzia/opis/opis_tool.py ~/.local/opis/; cp /tmp/mid-narzedzia/mes/mes_tool.py ~/.local/mes/
pip install -q --break-system-packages python-docx pyproj pillow numpy scipy matplotlib
E=~/.local/ekspertyza/ekspertyza_tool.py; O=~/.local/opis/opis_tool.py; M=~/.local/mes/mes_tool.py
python3 -I $E skala
```

## Przebieg

**1. Materiały.** Zbierz: OPZ zamawiającego (wymagany zakres i układ, wersja instrukcji przeglądów), poprzednie ekspertyzy i przeglądy, dokumentację archiwalną (normatyw i klasa obciążenia), zdjęcia i notatki z oględzin, wyniki badań. Starą ekspertyzę sprawdź `opis_tool.py sprawdz`: wychwyci nieaktualne Dz.U., uchylone rozporządzenia z 1999 i 2000 r. oraz brak lokalizacji.

**2. Specyfikacja.**
```bash
python3 -I $E wzor --out ekspertyza.json
python3 -I $O lokalizacja --xy X,Y --out lokalizacja        # mapy do pkt 2 (skill opis-techniczny)
```
Sekcja `opis` to część wg Instrukcji 02 (pola jak w opis_tool). `obiekt.parametry` trafia do tabeli. `charakterystyka.elementy` to podrozdziały.

**3. Oceny elementów** (skala 0–5: 5 odpowiedni, 4 zadowalający, 3 niepokojący, 2 niedostateczny, 1 przedawaryjny, 0 awaryjny; izolacja tylko 5/2/0):
```bash
python3 -I $E ocena ekspertyza.json
```
Liczy ocenę średnią (średnia ocenionych elementów) i ogólną: minimum z oceny średniej, pomostu, dźwigarów głównych oraz (min przyczółek + min filar)/2. Sprawdza też spójność: element z oceną ≤ 2 musi mieć opis i zalecenie, a z oceną ≤ 1 zalecenie w trybie A albo 1. Skala i reguły pochodzą z Instrukcji GDDKiA z 2005 r. (Zarz. nr 14). Jeżeli OPZ wskazuje nowszą instrukcję (np. GDDKiA 2020), sprawdź jej listę elementów i reguły i w razie różnic podaj je w `zrodlo_skali` i w opisie.

**4. Badania:**
```bash
python3 -I $E badania ekspertyza.json
```
- sklerometr: z odczytów R krzywa podstawowa PN-EN 13791:2007 (orientacyjnie, bez kalibracji odwiertami; odrzuca odczyty odbiegające o > 5 od średniej; f_ck,is = min(f_m − 1,48 s, f_min + 4)). Wynik przyrządu jako wytrzymałość kostkowa daje klasę PN-EN 206 (B45 → C35/45). Do obliczeń nośności zalecaj odwierty rdzeniowe.
- karbonatyzacja: głębokość wobec otuliny i czas dojścia frontu do zbrojenia z d = K√t;
- chlorki: % masy cementu (próg 0,4% żelbet, 0,2% sprężony; przeliczenie z % masy betonu przy założonej zawartości cementu);
- potencjały (ASTM C876: > −200 mV, −200…−350, < −350 mV), rezystywność (orientacyjnie);
- ubytki zbrojenia: współczynnik As do MRd.

**5. Nośność** (skill mes-eurokody): model rusztu z inwentaryzacji, `mes_tool.py nosnosc model.json --json nosnosc.json --png nosnosc.png`, warianty modelu poprzecznego. W `nosnosc.wyniki_mes` podaj plik JSON. Do raportu trafia tabela kategorii wg M i V, MLC i RF. Wniosek o oznakowaniu (B-18) formułuje autor. Gdy V daje niższą kategorię niż M, zaleć sprawdzenie V_Rd przy podporach.

**6. Zalecenia i warianty.** `zalecenia`: tryb (A – natychmiast; 1 – w następnym roku; 2, 3 – w kolejnych latach), numery elementów, zakres robót. `warianty`: I/II/III (remont, wzmocnienie, przebudowa), z rekomendacją i przesłankami do koncepcji. Podaj okres ważności ekspertyzy (`waznosc`).

**7. Dokument:**
```bash
python3 -I $E raport ekspertyza.json --out Ekspertyza_<obiekt>.docx --md Ekspertyza_<obiekt>.md
python3 -I $O sprawdz Ekspertyza_<obiekt>.docx              # Zał. 4 ZEW poz. 14
```
Układ: strona tytułowa z metryką i zespołem; 1–4 wg Instrukcji 02 (lokalizacja z mapą Polski i topograficzną); charakterystyka; stan techniczny (tabela ocen, ocena średnia i ogólna); badania; obliczenia nośności (tabele, rysunki); wnioski i zalecenia (tabela trybów, warianty, ważność); wykorzystane materiały [DA][N][U][R][W][I]. Obejrzyj render (PDF → PNG). Dokumentację fotograficzną i protokoły badań dołącz jako załączniki albo rozdział dodatkowy (`opis.rozdzialy_dodatkowe`).

## Zasady

- Oceny wynikają z oględzin. Gdy przenosisz je z wcześniejszego opracowania, napisz to w tekście („wg ekspertyzy 2021 – do aktualizacji”).
- Dane kadry (uprawnienia) bierz z prywatnej bazy (skill referencje-i-kadra).
- Nie zapisuj w OneDrive i nie wysyłaj maili. Pliki przekaż użytkownikowi.

## Test na obiekcie MiD

Most DW 226 w Wiślinie (dane z ekspertyzy MiD 2021): ocena średnia 2,50, ogólna 2,50; beton dźwigarów C35/45, przyczółka C16/20; strzemiona w strefie skarbonatyzowanej; nośność wg M 2/S32 (jak w 2021), wg V 3/S24 (2021 nie sprawdzano). Raport testowy: 8 stron, sprawdzenie Instrukcji 02 bez niezgodności poza numerem umowy do uzupełnienia.

## Odpowiedź (krótko)

Ocena ogólna i średnia, elementy w stanie ≤ 2, kategoria nośności i MLC (z miejscem decydującym), najważniejsze zalecenia z trybami, warianty, plik DOCX i lista rzeczy do potwierdzenia przez autora (oceny przeniesione, założenia modelu, brakujące badania).
