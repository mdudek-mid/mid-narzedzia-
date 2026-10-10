# mid-narzedzia

Narzędzia pomocnicze dla Claude w Pracowni Projektowej MiD: czytniki rysunków DWG/DXF i modeli IFC oraz narzędzie do analizy dokumentacji przetargowej. Repozytorium jest publiczne i nie zawiera danych firmy ani danych systemu przetargów.

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
