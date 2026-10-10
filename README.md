# mid-narzedzia

Narzędzia pomocnicze dla Claude w Pracowni Projektowej MiD: czytniki rysunków DWG/DXF i modeli IFC, narzędzia przetargowe (analiza SWZ, referencje i kadra, kryteria oceny, wycena, przegląd umowy) oraz projektowe (opis techniczny wg Instrukcji 02, MES i nośność mostów, ekspertyzy). Repozytorium jest publiczne i zawiera tylko kod. Dane firmy (baza referencji i kadry, parametry i wzorce wycen) leżą w prywatnym repo `mid-przetargi`, w katalogu `baza_mid/`.

## Czytnik DWG (LibreDWG 0.13.3)

Gotowe programy do czytania rysunków AutoCAD, żeby Claude nie budował ich od zera w każdej sesji (ok. 5 min). Korzysta z nich skill „dwg”, którego kopia leży w `skill/dwg/SKILL.md`.

Zawartość:
- `libredwg/libredwg-0.13.3-linux-x86_64.tar.xz` – programy `dwg2dxf`, `dwgread`, `dwg2SVG` (build statyczny, Linux x86_64, zależą tylko od libc) oraz skrypt `dwg_tool.py`.
- `libredwg/libredwg-0.13.3-linux-x86_64.tar.xz.sha256` – suma kontrolna.

Pobranie w sesji Claude (repo jest publiczne, nie trzeba go dołączać do sesji):

```bash
rm -rf /tmp/mid-narzedzia
git clone -q --depth 1 https://github.com/mdudek-mid/mid-narzedzia- /tmp/mid-narzedzia
cd /tmp/mid-narzedzia/libredwg && sha256sum -c libredwg-0.13.3-linux-x86_64.tar.xz.sha256
mkdir -p ~/.local/libredwg && tar -xJf libredwg-0.13.3-linux-x86_64.tar.xz -C ~/.local/libredwg --strip-components=1
~/.local/libredwg/bin/dwg2dxf --version
```

Źródła: https://github.com/LibreDWG/libredwg (tag 0.13.3). Licencja LibreDWG: GPLv3.
Zbudowano 2026-10-09, gcc 13.3, Ubuntu 24.04:
`cmake -G Ninja -DCMAKE_BUILD_TYPE=Release -DBUILD_SHARED_LIBS=OFF -DENABLE_LTO=OFF -DDISABLE_WERROR=ON`

## Czytnik IFC (IfcOpenShell)

Skrypt `ifc/ifc_tool.py` do modeli BIM (IFC2x3, IFC4, IFC4.3, także .ifcZIP): struktura obiektu, elementy z właściwościami, ilości (z modelu i z geometrii), materiały, osie tras z pikietażem i kontrolą ciągłości, georeferencja, podgląd PNG, walidacja. Korzysta z niego skill „ifc” (`skill/ifc/SKILL.md`). Biblioteka IfcOpenShell (LGPL-3.0) instaluje się z PyPI:

```bash
pip install -q --break-system-packages ifcopenshell matplotlib pytest
rm -rf /tmp/mid-narzedzia && git clone -q --depth 1 https://github.com/mdudek-mid/mid-narzedzia- /tmp/mid-narzedzia
mkdir -p ~/.local/ifc && cp /tmp/mid-narzedzia/ifc/ifc_tool.py ~/.local/ifc/
python3 ~/.local/ifc/ifc_tool.py info model.ifc
```

## Analiza dokumentacji przetargowej (swz_tool.py)

Skrypt `swz/swz_tool.py` do SWZ/IDW, OPZ, PFU, wzorów umów i wyjaśnień. Robi następujące rzeczy:
- pobiera dokumenty z e-Zamówień i platformazakupowa.pl (inne platformy przez prywatne repo systemu przetargów, jeśli jest w sesji);
- rozpakowuje archiwa;
- wyciąga tekst z numerami stron (DOCX/DOC przez LibreOffice, skany przez OCR tesseract z modelem polskim);
- oznacza dokumenty podmienione w zmianach SWZ;
- robi spis rozdziałów i klauzul, wyszukuje tematy (kary, limit kar, opóźnienie/zwłoka, waloryzacja, prawa autorskie, klauzula AI, materiały na FTP…);
- rozbiera tabele pytań i odpowiedzi;
- porównuje wersje umowy;
- liczy terminy na pytania wg Pzp.

Korzysta z niego skill „analiza-swz” (`skill/analiza-swz/SKILL.md`), który opisuje też wzór karty `Analiza_SWZ.md`.

```bash
rm -rf /tmp/mid-narzedzia && git clone -q --depth 1 https://github.com/mdudek-mid/mid-narzedzia- /tmp/mid-narzedzia
mkdir -p ~/.local/swz && cp /tmp/mid-narzedzia/swz/swz_tool.py ~/.local/swz/
pip install -q --break-system-packages requests beautifulsoup4 lxml openpyxl python-docx pdfplumber py7zr
python3 -I ~/.local/swz/swz_tool.py --help
```
Wymaga `pdftotext`/`pdftoppm` (poppler), `tesseract` i `soffice` (LibreOffice). Polski model OCR pobiera się sam z repozytorium tesseract-ocr/tessdata_fast (licencja Apache-2.0).

## Referencje i kadra (kadra_tool.py)

Skrypt `kadra/kadra_tool.py` korzysta z bazy referencji i kadry (`referencje.json`, `kadra.json`, prywatne). Robi następujące rzeczy:
- dobiera usługi do warunku udziału: typ obiektu, zakres, okres, wartość, parametry, i oznacza doświadczenie JDG wymagające udostępnienia zasobów;
- dobiera osoby wg specjalności i lat od uprawnień;
- zestawia dokumentacje projektanta do kryterium doświadczenia;
- wpisuje wybrane pozycje w formularze DOCX zamawiającego: wykaz usług, wykaz osób, załącznik o doświadczeniu. Obsługuje nagłówki wielopoziomowe i komórki scalone.

Korzysta z niego skill „referencje-i-kadra” (`skill/referencje-i-kadra/SKILL.md`).

```bash
rm -rf /tmp/mid-narzedzia && git clone -q --depth 1 https://github.com/mdudek-mid/mid-narzedzia- /tmp/mid-narzedzia
mkdir -p ~/.local/kadra && cp /tmp/mid-narzedzia/kadra/kadra_tool.py ~/.local/kadra/
pip install -q --break-system-packages python-docx openpyxl
python3 ~/.local/kadra/kadra_tool.py --baza /sciezka/mid-przetargi/baza_mid stan
```

## Kryteria oceny ofert (kryteria_tool.py)

Skrypt `kryteria/kryteria_tool.py` symuluje punktację ofert. Obsługuje typy kryteriów: cena Cmin/C, cena liniowa, min/max (termin, gwarancja), liniowe, progi, tak/nie, punkty. Liczy:
- najwyższą cenę, przy której oferta wygrywa, i cenę wyprzedzenia każdego konkurenta;
- scenariusze przed otwarciem ofert (udział najtańszego konkurenta w budżecie);
- próg rażąco niskiej ceny (art. 224 Pzp).

Rozbiera informację z otwarcia ofert, wyszukuje w SWZ opis kryteriów i dokumenty, które nie podlegają uzupełnieniu, i tworzy arkusz XLSX z formułami. Korzysta z niego skill „kryteria-oferty” (`skill/kryteria-oferty/SKILL.md`). Wymaga `openpyxl`.

## Wycena oferty (wycena_tool.py)

Skrypt `wycena/wycena_tool.py` liczy kalkulację prac projektowych ze specyfikacji JSON: pracę własną w rbg lub jnp, koszty zewnętrzne z narzutem na podwykonawców, rezerwę. Robi scenariusze A/B/C i rozkłada cel cenowy z zachowaniem pozycji sztywnych. Rozbija cenę na formularz zamawiającego: ryczałty, ilości, pozycje procentowe (np. kwota tymczasowa), stałe udziały i opcje. Sprawdza formalności i tworzy XLSX (formularz + kalkulacja z formułami) oraz opis MD w stylu MiD. Parametry metody wczytuje z prywatnego `baza_mid/parametry_wyceny.json`. Korzysta z niego skill „wycena-oferty” (`skill/wycena-oferty/SKILL.md`). Wymaga `openpyxl`.

## Opis techniczny wg Instrukcji 02 (opis_tool.py)

Skrypt `opis/opis_tool.py` składa część opisową opracowań MiD ze specyfikacji JSON w układzie Instrukcji 02: przedmiot, lokalizacja, cel i zakres, podstawa, wykorzystane materiały [DA][N][U][R][W][L][I][P]. Robi następujące rzeczy:
- lokalizacja z usług GUGiK: jednostki administracyjne i działka z ULDK, mapa Polski (PRG), mapa topograficzna, ortofotomapa z działkami (KIEG);
- aktualne cytaty aktów prawnych z API ELI Sejmu: tekst jednolity, „z późn. zm.”, uchylenia i akty zastępujące;
- generowanie DOCX i MD z tabelami, rysunkami i odwołaniami `[@klucz]` → [U1], [N2];
- sprawdzenie gotowego opisu (DOCX/PDF/MD/TXT) przed wydaniem (Zał. 4 ZEW poz. 14).

Korzysta z niego skill „opis-techniczny” (`skill/opis-techniczny/SKILL.md`). Wymaga `python-docx`, `pyproj`, `pillow`.

## MES i nośność mostów (mes_tool.py)

Skrypt `mes/mes_tool.py` to solver MES dla przęseł mostowych: belka ciągła albo ruszt (dźwigary + elementy poprzeczne sztywne lub przegubowe), powierzchnie wpływu metodą sprzężoną. Liczy obciążenia ruchome:
- LM1/LM2 z współczynnikami PTB 2022;
- PN-85/S-10030 klasy A–E, PN-66/B-02015 klasa I;
- samochody modelowe GDDKiA 1/S42…5/S10;
- pojazdy MLC kołowe i gąsienicowe (PTB 2022 zał. 2).

Wyznacza nośność użytkową wg Zarz. 17 GDDKiA (kategoria, m_u, znak B-18), klasę MLC, RF oraz M_Rd i V_Rd,c przekroju żelbetowego (PN-EN 1992). Testy (`test`) porównują wyniki ze wzorami zamkniętymi. Belkę ciągłą sprawdzono też z PyNite. Korzysta z niego skill „mes-eurokody” (`skill/mes-eurokody/SKILL.md`). Wymaga `numpy`, `scipy`, `matplotlib`.

## Ekspertyza mostu (ekspertyza_tool.py)

Skrypt `ekspertyza/ekspertyza_tool.py` obsługuje ekspertyzę od ocen do dokumentu:
- oceny elementów w skali GDDKiA 0–5, ocena średnia i ogólna obiektu, tryby robót A/1/2/3, kontrola spójności ocen i zaleceń;
- interpretacja badań: sklerometr (PN-EN 13791), klasa betonu, karbonatyzacja z prognozą, chlorki, potencjały, rezystywność, ubytki zbrojenia;
- raport DOCX wg Instrukcji 02 (przez `opis_tool.py`) z wynikami nośności z `mes_tool.py`.

Korzysta z niego skill „ekspertyza-mostu” (`skill/ekspertyza-mostu/SKILL.md`).

## Przegląd umowy (umowa_tool.py)

Skrypt `umowa/umowa_tool.py` dzieli umowę (DOCX/PDF, także kilka plików kontraktu FIDIC) na klauzule z lokalizacją i stroną i przechodzi listę kontrolną z perspektywy biura projektowego:
- kary (stawki, podstawa, zwłoka/opóźnienie), limit kar, odszkodowanie i ograniczenie odpowiedzialności;
- płatności i odbiór, płatności częściowe, waloryzacja, zabezpieczenie;
- prawa autorskie, AI, nadzór autorski, terminy zamawiającego, przewlekłość organów, odstąpienie, ograniczenie zakresu.

Liczy ekspozycję na kary, generuje pytania do zamawiającego albo tabelę propozycji do negocjacji, wstawia komentarze Worda przy klauzulach (python-docx ≥ 1.2) i dla P&B wypisuje kary do przeniesienia do umowy z GW. Korzysta z niego skill „przeglad-umowy” (`skill/przeglad-umowy/SKILL.md`).


## Protokoły (protokoly_tool.py)

Skrypt `protokoly/protokoly_tool.py` przygotowuje protokoły z realizacji umowy:
- rzeczowo-finansowe: wczytuje harmonogram z umowy (XLSX, PDF, DOCX, CSV, tekst), prowadzi historię protokołów w JSON, kontroluje zaawansowanie narastająco i tworzy XLSX w układzie wzoru GW z formułami oraz kwotą do faktury;
- przekazania dokumentacji: DOCX z wykazem pozycji i plików z sumami SHA-256 (manifest CSV);
- narady, rady techniczne i pobyty nadzoru autorskiego: DOCX z ustaleniami, decyzjami i zadaniami, kontrolą odstąpień (art. 36a i 36b PB), porządkowaniem transkryptu Teams (VTT/DOCX) i zbiorczą listą zadań.

Korzysta z niego skill „protokoly” (`skill/protokoly/SKILL.md`).
