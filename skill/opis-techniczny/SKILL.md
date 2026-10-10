---
name: opis-techniczny
description: Opisy techniczne i części opisowe opracowań Pracowni MiD wg Instrukcji 02 (przedmiot, lokalizacja z mapami, cel i zakres, podstawa, wykorzystane materiały [DA][N][U][R][W][L][I][P]) – generowanie DOCX, aktualne cytaty aktów prawnych z API ELI Sejmu, działki z ULDK, sprawdzenie zgodności przed wydaniem (Zał. 4 ZEW poz. 14).
---

# Opis techniczny wg Instrukcji 02 (MiD)

Używaj przy pisaniu lub sprawdzaniu części opisowej każdego opracowania MiD: PB, PT, PW, PZT, koncepcja, ekspertyza, STWiORB, PFU, opinia. Typowe prośby: „napisz opis techniczny do PB”, „zrób lokalizację z mapami”, „wykaz działek”, „podaj aktualny Dz.U. Prawa budowlanego”, „sprawdź opis przed wydaniem”, „czy ta ekspertyza jest zgodna z Instrukcją 02”.

Narzędzie `opis_tool.py` składa dokument ze specyfikacji JSON, pobiera dane z usług publicznych (GUGiK, ELI Sejmu) i sprawdza gotowe opisy. Treść merytoryczną piszesz ty, z materiałów wyjściowych.

## Instalacja (≈5 s)

```bash
rm -rf /tmp/mid-narzedzia && git clone -q --depth 1 https://github.com/mdudek-mid/mid-narzedzia- /tmp/mid-narzedzia
mkdir -p ~/.local/opis && cp /tmp/mid-narzedzia/opis/opis_tool.py ~/.local/opis/
pip install -q --break-system-packages python-docx pyproj pillow
O=~/.local/opis/opis_tool.py; python3 -I $O --help
```

## Instrukcja 02 – wymagania (REW01, 11.12.2024)

Na początku opracowania, w tej kolejności:
1. **Przedmiot opracowania** – konkretnie: co jest opracowywane (obiekt, roboty, inwestycja).
2. **Lokalizacja** – opis tekstowy (województwo, powiat, gmina, obręb, km drogi/linii, przeszkoda), mapa Polski ze wskazaniem miejsca, mapa topograficzna; dla PZT spis działek.
3. **Cel i zakres** – cel z jednej z kategorii: wynikający z przepisów prawa (decyzja, pozwolenie, zgłoszenie), ekspercki (ocena, ekspertyza) albo inny; zakres jako lista.
4. **Podstawa opracowania** – formalna (umowa / zlecenie: nr, data, strony) i merytoryczna (materiały wyjściowe z odwołaniem do wykazu).

Na końcu: **Wykorzystane materiały** w grupach [DA] dokumentacja archiwalna, [N] normy, [U] ustawy, [R] rozporządzenia, [W] wytyczne, [L] literatura, [I] źródła internetowe, [P] pozostałe. W tekście odwołuj się znacznikami [N1], [U2]. Każdą informację podawaj tylko raz. Zgodność opisu z Instrukcją 02 to poz. 14 listy kontrolnej wydania (Zał. 4 do ZEW).

## Przebieg

**1. Specyfikacja.**
```bash
python3 -I $O wzor --typ PB --out opis.json      # PB|PT|PW|PZT|koncepcja|ekspertyza|STWiORB|PFU|inne
```
Wypełnij pola z umowy, OPZ i materiałów wyjściowych. Odwołania w tekście pisz jako `[@klucz]`: klucz aktu z katalogu (`pb`, `udp`, `ptb_drogi_2022`…), ELI (`DU/2005/582`), numer normy z katalogu (`PN-EN 1991-2`) albo własny klucz pozycji (`DA1`). Narzędzie zamieni je na [U1], [N2] wg kolejności w wykazie. Nagłówek firmy można zmienić polem `firma` (`nazwa`, `adres`).

**2. Lokalizacja** (ULDK, UUG, WMS PRG/TOPO/ORTO/KIEG; GUGiK):
```bash
python3 -I $O lokalizacja --xy 481644,714438 --out lokalizacja          # PL-1992; --uklad 4326 dla lon,lat
python3 -I $O lokalizacja --miejscowosc Wiślina --gmina Pruszcz --out lokalizacja
python3 -I $O dzialki --xy "X1,Y1;X2,Y2;X3,Y3"                           # działki wzdłuż obiektu
python3 -I $O dzialki --id 220404_2.0006.1,220404_2.0006.76/3
```
Daje `lokalizacja.json` (z gotowym zdaniem `opis`), `rys_polska.png`, `rys_topo.png`, `rys_orto.png` (ortofotomapa z działkami). **Obejrzyj mapy** i sprawdź, czy znacznik wskazuje obiekt. Środek miejscowości z UUG to nie obiekt: przesuń punkt na obiekt (ortofotomapa) i uruchom ponownie z `--xy`. Działka w punkcie to zwykle działka drogi albo cieku. Dla PZT zbierz działki z kilku punktów wzdłuż obiektu i porównaj z mapą do celów projektowych.

**3. Przepisy – zawsze aktualny adres publikacyjny** (API ELI Sejmu):
```bash
python3 -I $O przepisy                       # katalog: aktualny t.j., „z późn. zm.”, status
python3 -I $O przepisy --szukaj "skrzyżowania linii kolejowych z drogami"
python3 -I $O cytuj pb                       # albo cytuj DU/2000/735 → informacja o uchyleniu i akcie zastępującym
```
Nie przepisuj adresów Dz.U. ze starych opracowań. Typowe błędy: stary t.j. Prawa budowlanego, rozporządzenia z 1999 i 2000 r. o warunkach technicznych dróg i obiektów inżynierskich (uchylone 21.09.2022, zastąpione Dz.U. 2022 poz. 1518), numer Dz.U. innego aktu. Norm katalog PKN nie udostępnia przez API, więc aktualność norm sprawdź ręcznie. Normy wycofane (PN-85/S-10030, PN-91/S-10042, PN-82/S-10052, PN-66/B-02015) przywołuj tylko jako podstawę oceny obiektu istniejącego, z takim opisem.

**4. Dokument:**
```bash
python3 -I $O generuj opis.json --out Opis_techniczny.docx --md Opis_techniczny.md
```
Strona tytułowa z metryką i zespołem, punkty 1–4, rozdziały merytoryczne z `rozdzialy`, wykaz materiałów. Narzędzie wypisze nierozwiązane odwołania `[@…]`. Przejrzyj render: `soffice --headless --convert-to pdf` → `pdftoppm -r 60 -png`.

**5. Sprawdzenie przed wydaniem** (także dokumentów cudzych i archiwalnych):
```bash
python3 -I $O sprawdz Opis_techniczny.docx --out raport_I02.md     # .docx / .pdf / .md / .txt
```
Sprawdza: obecność, kolejność i położenie punktów 1–4, wykaz materiałów na końcu, elementy lokalizacji, kategorię celu, podstawę formalną i merytoryczną, odwołania [X#] bez pozycji w wykazie i pozycje bez odwołań, aktualność aktów (uchylone, stary t.j., numer innego aktu, stary zapis „Dz. U. Nr 63 poz. 735”), normy wycofane, powtórzone zdania, niewypełnione pola (`[DO UZUPEŁNIENIA…]`, `[NR…]`, `[ ]`). Wynik ✗/⚠ oznacza poz. 14 listy kontrolnej „do poprawy”.

## Zasady

- Dane osobowe i numery uprawnień projektantów bierz z prywatnej bazy kadry (skill referencje-i-kadra), nie wpisuj ich do publicznych repozytoriów.
- Usługi GUGiK i ELI czasem nie odpowiadają (502, zerwane połączenie). Narzędzie ponawia zapytanie. Blokady proxy (403) zgłoś użytkownikowi, nie obchodź ich.
- Nie zapisuj w OneDrive i nie wysyłaj maili. Pliki przekaż użytkownikowi (SendUserFile).

## Odpowiedź (krótko)

Plik DOCX (i MD), raport `sprawdz`, lista pól do uzupełnienia, cytaty aktów zmienionych względem materiałów wyjściowych oraz uwagi do lokalizacji (czy znacznik i działki są potwierdzone na ortofotomapie).
