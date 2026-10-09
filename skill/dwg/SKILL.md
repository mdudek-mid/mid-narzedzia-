---
name: dwg
description: Odczyt rysunków AutoCAD .dwg i .dxf (warstwy, teksty, opisy, wymiary, tabelki, podgląd PNG) przez LibreDWG + ezdxf. Używaj zawsze, gdy zadanie dotyczy pliku .dwg lub .dxf (załącznik, plik z komputera, folder projektu, ekspertyza). Nigdy nie odpowiadaj, że nie da się czytać DWG.
---

# Czytnik DWG/DXF (LibreDWG + ezdxf)

Środowisko robocze nie ma gotowego czytnika DWG i każda nowa sesja zaczyna od zera. Gotowe programy LibreDWG 0.13.3 leżą w publicznym repozytorium Pracowni MiD `mdudek-mid/mid-narzedzia-` (nazwa kończy się myślnikiem), w katalogu `libredwg/`. Instalacja z niego trwa kilka sekund i działa w każdej sesji, także u współpracowników bez podłączonego GitHuba. Gdy ta droga zawiedzie, budujemy ze źródeł (ok. 5 minut). Obie drogi sprawdzone 2026-10-09 na rysunkach R14, 2000, 2004, 2007, 2010, 2013 i 2018.

## Krok 1 – instalacja (raz na sesję)

Najpierw sprawdź, czy już jest: `test -x ~/.local/libredwg/bin/dwg2dxf && echo jest`. Jeśli jest, przejdź do kroku 3.

### 1a. Szybka droga: gotowe programy z GitHuba (kilka sekund)

Repozytorium jest publiczne, więc nie trzeba go dołączać do sesji. Pobierz je i zainstaluj programy:

```bash
rm -rf /tmp/mid-narzedzia
git clone -q --depth 1 https://github.com/mdudek-mid/mid-narzedzia- /tmp/mid-narzedzia
cd /tmp/mid-narzedzia/libredwg && sha256sum -c libredwg-0.13.3-linux-x86_64.tar.xz.sha256
mkdir -p ~/.local/libredwg && tar -xJf libredwg-0.13.3-linux-x86_64.tar.xz -C ~/.local/libredwg --strip-components=1
pip install -q --break-system-packages ezdxf matplotlib cairosvg
export PATH=~/.local/libredwg/bin:$PATH && dwg2dxf --version   # oczekiwane: dwg2dxf 0.13.3
```

Paczka zawiera `dwg2dxf`, `dwgread`, `dwg2SVG` oraz skrypt `~/.local/libredwg/dwg_tool.py`, identyczny z tym z kroku 2. Po udanej instalacji krok 2 pomiń.

Jeśli `git clone` zwróci błąd dostępu, a w sesji jest narzędzie `add_repo`, dołącz nim repo (owner `mdudek-mid`, repo `mid-narzedzia-`, access `read`) i spróbuj raz jeszcze. Gdy repo nie istnieje pod tą nazwą, mogło zostać przemianowane: poszukaj go narzędziem `list_repos` z zapytaniem `narzedzia`.

Jeśli mimo to którykolwiek krok zawiedzie (repo niedostępne, suma kontrolna się nie zgadza, program nie startuje), przejdź do drogi 1b.

### 1b. Zapas: budowa ze źródeł (ok. 5 minut)

Budowę uruchom w tle, a w tym czasie wykonuj kroki 2 i 3:

```bash
mkdir -p ~/.local/src && cd ~/.local/src
[ -d libredwg ] || git clone -q --depth 1 --branch 0.13.3 https://github.com/LibreDWG/libredwg.git
cd libredwg
git submodule update --init --depth 1 jsmn   # BEZ TEGO kompilacja pada na ../jsmn/jsmn.h
mkdir -p b && cd b
cmake -G Ninja -DCMAKE_BUILD_TYPE=Release -DBUILD_SHARED_LIBS=OFF -DENABLE_LTO=OFF -DDISABLE_WERROR=ON .. > ../cmake.log 2>&1
(ninja -j"$(nproc)" dwg2dxf dwgread dwg2SVG > ../ninja.log 2>&1; echo "EXIT $?" >> ../ninja.log) &
```

Równolegle zainstaluj biblioteki Pythona:

```bash
pip install -q --break-system-packages ezdxf matplotlib cairosvg
```

Poczekaj na koniec budowy i skopiuj programy:

```bash
until grep -q EXIT ~/.local/src/libredwg/ninja.log; do sleep 15; done; tail -2 ~/.local/src/libredwg/ninja.log
mkdir -p ~/.local/libredwg/bin && cp ~/.local/src/libredwg/b/{dwg2dxf,dwgread,dwg2SVG} ~/.local/libredwg/bin/
export PATH=~/.local/libredwg/bin:$PATH && dwg2dxf --version   # oczekiwane: dwg2dxf 0.13.3
```

Nie trać czasu na inne źródła, bo nie działają w tym środowisku: apt (brak pakietu libredwg-tools), pobieranie paczek z github.com/…/releases, ODA File Converter. Działa tylko `git clone`. Build jest statyczny, więc programy zależą wyłącznie od libc.

## Krok 2 – skrypt pomocniczy (tylko przy drodze 1b)

Zapisz poniższy skrypt w całości jako `~/.local/libredwg/dwg_tool.py` (narzędziem Write):

```python
#!/usr/bin/env python3
"""dwg_tool.py - odczyt rysunkow DWG/DXF (LibreDWG + ezdxf).

Uzycie:
  python3 dwg_tool.py info    PLIK.dwg|dxf            # wersja, warstwy, liczba obiektow, zasieg
  python3 dwg_tool.py texts   PLIK [--layer WARSTWA]  # wszystkie teksty/wymiary/atrybuty -> CSV na stdout
  python3 dwg_tool.py render  PLIK [-o out.png] [--layout NAZWA] [--layers A,B] [--window=x1,y1,x2,y2] [--dpi 200]
      (--window z ZNAKIEM =, bo wspolrzedne ujemne inaczej wygladaja jak opcje)
  python3 dwg_tool.py layouts PLIK                    # lista arkuszy (Model + layouty)

Plik .dwg jest najpierw konwertowany do .dxf (dwg2dxf z LibreDWG) obok, w katalogu roboczym.
"""
import argparse, csv, os, shutil, subprocess, sys
from collections import Counter

BIN = os.path.expanduser("~/.local/libredwg/bin")
DWG2DXF = shutil.which("dwg2dxf") or os.path.join(BIN, "dwg2dxf")
DWGREAD = shutil.which("dwgread") or os.path.join(BIN, "dwgread")


def u(s):
    """Polskie litery: DXF R2000/2004 zapisuje znaki spoza strony kodowej jako \\U+0105 - dekodujemy."""
    from ezdxf.lldxf.encoding import decode_dxf_unicode, has_dxf_unicode
    s = s or ""
    return decode_dxf_unicode(s) if has_dxf_unicode(s) else s


def to_dxf(path, workdir=None):
    if path.lower().endswith(".dxf"):
        return path
    if not os.path.exists(DWG2DXF):
        sys.exit("Brak dwg2dxf - najpierw zbuduj LibreDWG (krok 1 skilla dwg).")
    workdir = workdir or os.path.join(os.getcwd(), "dwg_work")
    os.makedirs(workdir, exist_ok=True)
    out = os.path.join(workdir, os.path.splitext(os.path.basename(path))[0] + ".dxf")
    if os.path.exists(out) and os.path.getmtime(out) >= os.path.getmtime(path):
        return out
    r = subprocess.run([DWG2DXF, "-y", "-o", out, path], capture_output=True, text=True)
    if not os.path.exists(out) or os.path.getsize(out) == 0:
        sys.exit(f"dwg2dxf nie wygenerowal DXF (kod {r.returncode}):\n{r.stderr[-2000:]}")
    if r.returncode != 0:
        print(f"[uwaga] dwg2dxf zakonczyl sie kodem {r.returncode}, DXF powstal - czytam z odzyskiwaniem.", file=sys.stderr)
    return out


def load(path):
    import ezdxf
    from ezdxf import recover
    dxf = to_dxf(path)
    try:
        doc = ezdxf.readfile(dxf)
        aud = None
    except Exception:
        doc, aud = recover.readfile(dxf)
        if aud and aud.has_errors:
            print(f"[uwaga] odzyskano z {len(aud.errors)} bledami", file=sys.stderr)
    fix_unicode(doc)
    return doc


def fix_unicode(doc):
    """Dekoduje \\U+XXXX w tekstach calego rysunku (w pamieci), zeby polskie litery byly poprawne
    takze na podgladzie PNG."""
    from ezdxf.lldxf.encoding import has_dxf_unicode, decode_dxf_unicode
    containers = list(doc.layouts) + [b for b in doc.blocks]
    for c in containers:
        for e in c:
            t = e.dxftype()
            if t in ("TEXT", "ATTRIB", "ATTDEF", "MTEXT", "DIMENSION"):
                txt = e.dxf.get("text", None) if t != "MTEXT" else e.text
                if txt and has_dxf_unicode(txt):
                    if t == "MTEXT":
                        e.text = decode_dxf_unicode(txt)
                    else:
                        e.dxf.text = decode_dxf_unicode(txt)
            if t == "INSERT":
                for att in e.attribs:
                    txt = att.dxf.get("text", "")
                    if txt and has_dxf_unicode(txt):
                        att.dxf.text = decode_dxf_unicode(txt)


def cmd_info(a):
    from ezdxf import bbox
    doc = load(a.file)
    msp = doc.modelspace()
    print(f"Plik: {a.file}")
    print(f"Wersja DXF: {doc.dxfversion}  ($ACADVER={doc.header.get('$ACADVER')})  "
          f"jednostki $INSUNITS={doc.header.get('$INSUNITS')}  codepage={doc.header.get('$DWGCODEPAGE')}")
    ext = bbox.extents(msp, fast=True)
    if ext.has_data:
        print(f"Zasieg modelu: {tuple(round(v, 3) for v in ext.extmin)} .. {tuple(round(v, 3) for v in ext.extmax)}")
    print("Arkusze:", ", ".join(l.name for l in doc.layouts))
    cnt = Counter(e.dxftype() for e in msp)
    print("Obiekty w modelu:", ", ".join(f"{k}={v}" for k, v in cnt.most_common()))
    per_layer = Counter(e.dxf.get("layer", "0") for e in msp)
    print("\nWarstwy (obiekty w modelu | wl/wyl | zamr.):")
    for lay in sorted(doc.layers, key=lambda l: l.dxf.name.lower()):
        print(f"  {u(lay.dxf.name):40s} {per_layer.get(lay.dxf.name, 0):6d}  "
              f"{'wl ' if lay.is_on() else 'WYL'}  {'Z' if lay.is_frozen() else ''}")
    blocks = [u(b.name) for b in doc.blocks if not b.name.startswith("*")]
    print(f"\nBloki ({len(blocks)}):", ", ".join(blocks[:60]) + (" ..." if len(blocks) > 60 else ""))


def iter_texts(doc, layer=None):
    for lay in doc.layouts:
        for e in lay:
            try:
                yield from _texts_of(e, lay.name, layer)
            except Exception as ex:  # pojedynczy uszkodzony obiekt nie zatrzymuje calosci
                print(f"[uwaga] pominieto {e.dxftype()} {e.dxf.handle}: {ex}", file=sys.stderr)


def _texts_of(e, space, layer, depth=0):
    t = e.dxftype()
    if t not in ("TEXT", "ATTRIB", "ATTDEF", "MTEXT", "DIMENSION", "INSERT"):
        return
    lyr = u(e.dxf.get("layer", ""))
    ins = e.dxf.get("insert", None) if e.dxf.is_supported("insert") else None
    xy = (round(ins.x, 3), round(ins.y, 3)) if ins is not None else ("", "")
    if t in ("TEXT", "ATTRIB", "ATTDEF"):
        val = e.dxf.get("text", "")
        if t != "TEXT":  # atrybuty: TAG=wartosc
            val = f"{e.dxf.get('tag', '')}={val}"
        if not layer or lyr == layer:
            yield (space, lyr, t, xy[0], xy[1], u(val).strip())
    elif t == "MTEXT":
        if not layer or lyr == layer:
            yield (space, lyr, t, xy[0], xy[1], u(e.plain_text()).strip())
    elif t == "DIMENSION":
        if not layer or lyr == layer:
            meas = e.dxf.get("actual_measurement", None)
            if meas is None:
                try:
                    meas = e.get_measurement()
                except Exception:
                    meas = ""
            if isinstance(meas, (int, float)):
                m = f"{meas:.3f}"
            elif hasattr(meas, "x"):  # wymiar rzednej zwraca wektor
                m = f"X={meas.x:.3f} Y={meas.y:.3f}"
            else:
                m = str(meas)
            txt = e.dxf.get("text", "")
            val = u((txt.replace("<>", m) if "<>" in txt else txt) if txt and txt != "<>" else m)
            p = e.dxf.get("text_midpoint", None) or e.dxf.get("defpoint", None)
            yield (space, lyr, t, round(p.x, 3) if p else "", round(p.y, 3) if p else "", val)
    elif t == "INSERT" and depth < 4:
        for att in e.attribs:
            yield from _texts_of(att, space, layer, depth + 1)


def texts_from_json(path, layer=None):
    """Zapasowa sciezka: surowy odczyt DWG przez dwgread -O JSON (gdy DXF nie wczytuje sie w ezdxf)."""
    import json, tempfile
    out = os.path.join(tempfile.mkdtemp(), "d.json")
    subprocess.run([DWGREAD, "-O", "JSON", "-o", out, path], capture_output=True)
    d = json.load(open(out, encoding="utf-8", errors="replace"))
    objs = d.get("OBJECTS", [])
    layers = {o["handle"][-1]: o.get("name", "") for o in objs if o.get("object") == "LAYER" and o.get("handle")}
    for o in objs:
        t = o.get("entity")
        if t not in ("TEXT", "MTEXT", "ATTRIB", "ATTDEF", "DIMENSION_LINEAR", "DIMENSION_ALIGNED",
                     "DIMENSION_ANG2LN", "DIMENSION_ANG3PT", "DIMENSION_RADIUS", "DIMENSION_DIAMETER",
                     "DIMENSION_ORDINATE"):
            continue
        ref = o.get("layer") or [0]
        lyr = layers.get(ref[-1], str(ref))
        if layer and lyr != layer:
            continue
        p = o.get("ins_pt") or o.get("text_midpt") or ["", ""]
        val = o.get("text_value") or o.get("text") or o.get("user_text") or ""
        if t.startswith("DIMENSION") and not val:
            val = f"{o.get('act_measurement', '')}"
        if t in ("ATTRIB", "ATTDEF"):
            val = f"{o.get('tag', '')}={val}"
        if t == "MTEXT":
            from ezdxf.tools.text import plain_mtext
            val = plain_mtext(val)
        x = round(p[0], 3) if isinstance(p[0], (int, float)) else ""
        y = round(p[1], 3) if len(p) > 1 and isinstance(p[1], (int, float)) else ""
        yield ("?", u(lyr), t, x, y, u(val).strip())


def cmd_texts(a):
    w = csv.writer(sys.stdout)
    w.writerow(["arkusz", "warstwa", "typ", "x", "y", "tekst"])
    try:
        rows = list(iter_texts(load(a.file), a.layer))
    except (Exception, SystemExit) as ex:
        if a.file.lower().endswith(".dxf"):
            raise
        print(f"[uwaga] DXF nieczytelny ({ex}); czytam DWG bezposrednio przez dwgread JSON", file=sys.stderr)
        rows = list(texts_from_json(a.file, a.layer))
    for row in rows:
        if row[-1]:
            w.writerow(row)


def cmd_layouts(a):
    doc = load(a.file)
    for l in doc.layouts:
        print(f"{l.name}\t{len(l)} obiektow")


def cmd_render(a):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from ezdxf.addons.drawing import RenderContext, Frontend
    from ezdxf.addons.drawing.matplotlib import MatplotlibBackend
    from ezdxf.addons.drawing.config import Configuration, BackgroundPolicy, ColorPolicy

    doc = load(a.file)
    layout = doc.layouts.get(a.layout) if a.layout else doc.modelspace()
    for name in {e.dxf.get("layer", "0") for e in layout}:  # warstwy bez definicji rysowalyby sie na bialo
        if name not in doc.layers:
            doc.layers.add(name)
    if a.layers:  # wylacz w kopii w pamieci wszystkie warstwy poza wskazanymi
        keep = {s.strip().lower() for s in a.layers.split(",")}
        for l in doc.layers:
            (l.on if l.dxf.name.lower() in keep else l.off)()
    ctx = RenderContext(doc)
    # kadr: --window albo zasieg obiektow bez nieskonczonych XLINE/RAY (inaczej rozwalaja widok)
    if a.window:
        x1, y1, x2, y2 = (float(v) for v in a.window.split(","))
    else:
        from ezdxf import bbox
        hidden = {l.dxf.name.lower() for l in doc.layers if not l.is_on() or l.is_frozen()}
        ents = (e for e in layout if e.dxftype() not in ("XLINE", "RAY", "VIEWPORT")
                and e.dxf.get("layer", "0").lower() not in hidden)
        boxes = [b for b in (bbox.extents([e], fast=False) for e in ents) if b.has_data]
        if not boxes:
            x1, y1, x2, y2 = 0, 0, 1, 1
        elif len(boxes) < 50 or a.full:
            x1, y1 = min(b.extmin.x for b in boxes), min(b.extmin.y for b in boxes)
            x2, y2 = max(b.extmax.x for b in boxes), max(b.extmax.y for b in boxes)
        else:  # odporny kadr: obcina 1% skrajnych obiektow (zablakane elementy daleko od rysunku)
            q = lambda vals, p: sorted(vals)[min(len(vals) - 1, max(0, int(p * len(vals))))]
            x1, y1 = q([b.extmin.x for b in boxes], 0.01), q([b.extmin.y for b in boxes], 0.01)
            x2, y2 = q([b.extmax.x for b in boxes], 0.99), q([b.extmax.y for b in boxes], 0.99)
            print(f"[info] kadr odporny: {x1:.1f},{y1:.1f},{x2:.1f},{y2:.1f} (calosc: --full)", file=sys.stderr)
    x1, x2 = sorted((x1, x2)); y1, y2 = sorted((y1, y2))
    mx, my = (x2 - x1) * 0.02 or 1, (y2 - y1) * 0.02 or 1
    x1, x2, y1, y2 = x1 - mx, x2 + mx, y1 - my, y2 + my
    h = min(max(a.width * (y2 - y1) / (x2 - x1), 2), 40)  # proporcje obrazu = proporcje kadru
    fig = plt.figure(figsize=(a.width, h))
    ax = fig.add_axes([0, 0, 1, 1])
    cfg = Configuration(background_policy=BackgroundPolicy.WHITE,
                        color_policy=ColorPolicy.BLACK if a.mono else ColorPolicy.COLOR)
    Frontend(ctx, MatplotlibBackend(ax, adjust_figure=False), config=cfg).draw_layout(layout, finalize=True)
    ax.set_xlim(x1, x2)
    ax.set_ylim(y1, y2)
    ax.set_aspect("equal", adjustable="box")
    out = a.output or os.path.splitext(os.path.basename(a.file))[0] + (f"_{a.layout}" if a.layout else "") + ".png"
    fig.savefig(out, dpi=a.dpi, facecolor="white")
    print(out)


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = p.add_subparsers(dest="cmd", required=True)
    for n in ("info", "texts", "render", "layouts"):
        s = sub.add_parser(n)
        s.add_argument("file")
        if n == "texts":
            s.add_argument("--layer")
        if n == "render":
            s.add_argument("-o", "--output")
            s.add_argument("--layout")
            s.add_argument("--layers", help="pokaz tylko te warstwy, rozdzielone przecinkiem")
            s.add_argument("--window", help="x1,y1,x2,y2 - wycinek w jednostkach rysunku")
            s.add_argument("--dpi", type=int, default=200)
            s.add_argument("--width", type=float, default=16, help="szerokosc obrazu w calach")
            s.add_argument("--mono", action="store_true", help="wszystko na czarno")
            s.add_argument("--full", action="store_true", help="kadr na wszystkie obiekty, bez obcinania skrajnych")
    a = p.parse_args()
    {"info": cmd_info, "texts": cmd_texts, "render": cmd_render, "layouts": cmd_layouts}[a.cmd](a)


if __name__ == "__main__":
    main()
```

## Krok 3 – skąd wziąć rysunek

- Załącznik w rozmowie: czytaj po nazwie pliku.
- Plik na komputerze użytkownika (folder projektu, ekspertyzy, archiwum): skopiuj go do środowiska roboczego narzędziem do pobierania plików z komputera i pracuj na kopii. Na komputerze użytkownika (Windows) niczego nie buduj ani nie instaluj.
- Oryginałów nie zmieniaj. Pośrednie pliki DXF trafiają do `./dwg_work/`.

## Krok 4 – czytanie

```bash
export PATH=~/.local/libredwg/bin:$PATH; T=~/.local/libredwg/dwg_tool.py
python3 $T info rysunek.dwg                    # wersja, jednostki, arkusze, warstwy z liczbą obiektów, bloki
python3 $T texts rysunek.dwg > teksty.csv      # TEXT/MTEXT/atrybuty/wymiary z warstwą i współrzędnymi
python3 $T texts rysunek.dwg --layer OPISY     # tylko jedna warstwa
python3 $T layouts rysunek.dwg                 # lista arkuszy papieru
python3 $T render rysunek.dwg -o model.png --dpi 120               # przegląd modelu
python3 $T render rysunek.dwg --layout "A1-01" -o arkusz.png       # arkusz papieru
python3 $T render rysunek.dwg --window=1200,300,1800,700 -o detal.png   # powiększenie
python3 $T render rysunek.dwg --layers=KONSTRUKCJA,WYMIARY --mono -o konstr.png
```

Wygenerowany PNG obejrzyj narzędziem Read.

Dobra kolejność pracy:
- Zacznij od `info` i `texts`. Większość pytań (opisy, tabelka rysunkowa, numer i skala rysunku, rzędne, wymiary, oznaczenia elementów) rozstrzyga się na tekstach, bez renderowania. Tabelki rysunkowe to zwykle atrybuty bloków (typ ATTRIB, zapis `TAG=wartość`).
- Drobny tekst na całym arkuszu jest nieczytelny. Najpierw przegląd przy niskim dpi, potem `--window` wokół współrzędnych znalezionych w `texts`.
- Wymiary w CSV to wartość zmierzona albo tekst nadpisany przez projektanta (wtedy pokazany jest tekst). Jednostki podaje `$INSUNITS` w `info` (4 = mm, 6 = m).
- Wiele rysunków: pętla `for f in *.dwg; do ...; done`, wyniki zbierz w jednym CSV lub arkuszu.

## Znane pułapki (sprawdzone)

- **Polskie litery.** W rysunkach R2000/2004 dwg2dxf zapisuje znaki spoza strony kodowej jako `\U+0105`. Skrypt je dekoduje, także na podglądzie PNG. Jeśli piszesz własny kod na ezdxf, użyj `ezdxf.lldxf.encoding.decode_dxf_unicode`.
- **Opcja `--window` wymaga znaku `=`**, bo ujemne współrzędne bez niego wyglądają jak opcje.
- **Kadr.** Linie XLINE/RAY i zabłąkane obiekty daleko od rysunku psują automatyczny kadr. Skrypt pomija XLINE/RAY, a przy co najmniej 50 obiektach obcina 1% skrajnych. `--full` wyłącza obcinanie.
- **Warstwy bez definicji** ezdxf rysowałby na biało. Skrypt dodaje brakujące definicje przed renderowaniem.
- **Uszkodzony DXF.** Gdy ezdxf nie wczyta DXF, `texts` sam przełącza się na surowy odczyt `dwgread -O JSON` (bez nazw arkuszy). Podgląd awaryjny: `dwg2SVG plik.dwg > plik.svg`, potem `python3 -c "import cairosvg; cairosvg.svg2png(url='plik.svg', write_to='plik.png', output_width=2000, background_color='white')"`. ImageMagick w tym środowisku nie czyta SVG.
- **Niezerowy kod dwg2dxf.** Czasem dwg2dxf kończy się błędem, a DXF i tak powstaje. Skrypt wtedy czyta go w trybie odzyskiwania i wypisuje ostrzeżenie.
- **Czcionki SHX** nie są dostępne, więc teksty na PNG mają czcionkę zastępczą. Treść jest poprawna, wygląd nieco inny niż w AutoCADzie.
