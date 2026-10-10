# mid-narzedzia

Narzędzia pomocnicze dla Claude w Pracowni Projektowej MiD: czytniki rysunków DWG/DXF i modeli IFC oraz narzędzia przetargowe (analiza SWZ, referencje i kadra, kryteria oceny, wycena). Repozytorium jest publiczne i zawiera tylko kod. Dane firmy (baza referencji i kadry, parametry i wzorce wycen) leżą w prywatnym repo `mid-przetargi`, w katalogu `baza_mid/`.

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
