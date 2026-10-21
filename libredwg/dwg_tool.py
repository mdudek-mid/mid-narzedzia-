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

# Linux: ~/.local/libredwg/bin/dwg2dxf; Windows: %USERPROFILE%\\.local\\libredwg\\bin\\dwg2dxf.exe
# (inny katalog: zmienna LIBREDWG_BIN)
EXE = ".exe" if os.name == "nt" else ""
BIN = os.environ.get("LIBREDWG_BIN") or os.path.expanduser(os.path.join("~", ".local", "libredwg", "bin"))
DWG2DXF = shutil.which("dwg2dxf") or os.path.join(BIN, "dwg2dxf" + EXE)
DWGREAD = shutil.which("dwgread") or os.path.join(BIN, "dwgread" + EXE)


def u(s):
    """Polskie litery: DXF R2000/2004 zapisuje znaki spoza strony kodowej jako \\U+0105 - dekodujemy."""
    from ezdxf.lldxf.encoding import decode_dxf_unicode, has_dxf_unicode
    s = s or ""
    return decode_dxf_unicode(s) if has_dxf_unicode(s) else s


def to_dxf(path, workdir=None):
    if path.lower().endswith(".dxf"):
        return path
    if not os.path.exists(DWG2DXF):
        sys.exit("Brak dwg2dxf - najpierw zainstaluj LibreDWG (krok 1 skilla dwg; na Windows krok 1w).")
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
