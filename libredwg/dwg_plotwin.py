#!/usr/bin/env python3
"""dwg_plotwin.py - podglad PNG wycinka rysunku DWG/DXF, gdy `dwg_tool.py render` daje czarny/pusty obraz albo pada.

Rysuje wprost matplotlibem: LINE, LWPOLYLINE, POLYLINE, ARC, CIRCLE, ELLIPSE, SPLINE, TEXT, MTEXT, DIMENSION (tekst),
INSERT (rozwiniete bloki, do 4 poziomow), opcjonalnie obrysy HATCH. Wczytuje przez ezdxf.recover (odporny na bledy).
Plik .dwg najpierw konwertuje dwg2dxf (LibreDWG) i naprawia dwg_fixdxf.py (ten sam katalog).

Uzycie:
  python3 dwg_plotwin.py PLIK.dwg|dxf -o out.png --window=x1,y1,x2,y2 [--layout NAZWA] [--dpi 90] [--hatch] [--notext]
  python3 dwg_plotwin.py PLIK.dxf -o out.png --center=cx,cy,R          (kwadrat 2R x 2R)
  python3 dwg_plotwin.py PLIK.dxf --extents                            (tylko wypisz zasieg modelu)
Okno podawaj w jednostkach rysunku (zasieg: `dwg_tool.py info` albo --extents). Znak = przy ujemnych liczbach.
"""
import argparse, math, os, shutil, subprocess, sys

HERE = os.path.dirname(os.path.abspath(__file__))
BIN = os.environ.get("LIBREDWG_BIN") or os.path.expanduser("~/.local/libredwg/bin")

def to_dxf(path):
    if path.lower().endswith(".dxf"):
        return path
    d2d = shutil.which("dwg2dxf") or os.path.join(BIN, "dwg2dxf")
    wd = os.path.join(os.getcwd(), "dwg_work"); os.makedirs(wd, exist_ok=True)
    out = os.path.join(wd, os.path.splitext(os.path.basename(path))[0] + ".dxf")
    subprocess.run([d2d, "-y", "-o", out, path], capture_output=True)
    if not os.path.exists(out) or os.path.getsize(out) == 0:
        sys.exit("dwg2dxf nie wygenerowal DXF")
    sys.path.insert(0, HERE)
    try:
        from dwg_fixdxf import fix
        fix(out)
    except ImportError:
        print("[uwaga] brak dwg_fixdxf.py obok - pomijam naprawe", file=sys.stderr)
    return out

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("file"); ap.add_argument("-o", "--output", default="plotwin.png")
    ap.add_argument("--window"); ap.add_argument("--center"); ap.add_argument("--layout")
    ap.add_argument("--dpi", type=int, default=90); ap.add_argument("--size", type=float, default=14)
    ap.add_argument("--hatch", action="store_true"); ap.add_argument("--notext", action="store_true")
    ap.add_argument("--extents", action="store_true")
    a = ap.parse_args()
    import ezdxf
    from ezdxf import recover, bbox
    doc, aud = recover.readfile(to_dxf(a.file))
    lay = doc.layouts.get(a.layout) if a.layout else doc.modelspace()
    if a.extents or not (a.window or a.center):
        # odporny zasieg: kwantyle 1%-99% obiektow (zablakane elementy daleko od rysunku psuja zwykly zasieg)
        boxes = []
        for e in lay:
            if e.dxftype() in ("XLINE", "RAY", "VIEWPORT"):
                continue
            try:
                b = bbox.extents([e], fast=True)
                if b.has_data: boxes.append(b)
            except Exception:
                pass
        if not boxes:
            sys.exit("brak obiektow do narysowania")
        q = lambda v, p: sorted(v)[min(len(v) - 1, max(0, int(p * len(v))))]
        x1, y1 = q([b.extmin.x for b in boxes], 0.01), q([b.extmin.y for b in boxes], 0.01)
        x2, y2 = q([b.extmax.x for b in boxes], 0.99), q([b.extmax.y for b in boxes], 0.99)
        print(f"zasieg (1-99%): {x1:.3f},{y1:.3f},{x2:.3f},{y2:.3f}  obiektow: {len(boxes)}")
        if a.extents:
            return
    elif a.center:
        cx, cy, r = (float(v) for v in a.center.split(",")); x1, y1, x2, y2 = cx - r, cy - r, cx + r, cy + r
    else:
        x1, y1, x2, y2 = (float(v) for v in a.window.split(","))
    x1, x2 = sorted((x1, x2)); y1, y2 = sorted((y1, y2))
    import matplotlib; matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.patches import Arc, Circle
    w = a.size; h = max(2.0, min(40.0, w * (y2 - y1) / max(x2 - x1, 1e-9)))
    fig, ax = plt.subplots(figsize=(w, h))
    inside = lambda x, y: x1 <= x <= x2 and y1 <= y <= y2
    def seg(pts, style, lw):
        if len(pts) > 1 and any(inside(x, y) for x, y in pts):
            ax.plot([p[0] for p in pts], [p[1] for p in pts], style, lw=lw)
    def draw(e, depth=0):
        t = e.dxftype()
        try:
            if t == "LINE":
                s, en = e.dxf.start, e.dxf.end
                seg([(s.x, s.y), (en.x, en.y)], "k-", 0.5)
            elif t == "LWPOLYLINE":
                pts = [(v.x, v.y) for v in e.flattening(0.01)] if e.has_arc else [p[:2] for p in e.get_points("xy")]
                if e.closed and pts: pts.append(pts[0])
                seg(pts, "b-", 0.6)
            elif t == "POLYLINE":
                pts = [(v.x, v.y) for v in e.points()]
                if e.is_closed and pts: pts.append(pts[0])
                seg(pts, "b-", 0.6)
            elif t == "CIRCLE":
                c = e.dxf.center
                if inside(c.x, c.y): ax.add_patch(Circle((c.x, c.y), e.dxf.radius, fill=False, lw=0.5))
            elif t == "ARC":
                c = e.dxf.center
                if inside(c.x, c.y):
                    ax.add_patch(Arc((c.x, c.y), 2 * e.dxf.radius, 2 * e.dxf.radius,
                                     theta1=e.dxf.start_angle, theta2=e.dxf.end_angle, lw=0.5))
            elif t in ("ELLIPSE", "SPLINE"):
                seg([(v.x, v.y) for v in e.flattening(0.01)], "g-", 0.5)
            elif t in ("TEXT", "MTEXT") and not a.notext:
                p = e.dxf.insert
                if inside(p.x, p.y):
                    tx = e.plain_text() if t == "MTEXT" else e.dxf.text
                    ax.text(p.x, p.y, tx[:40], fontsize=6, color="r")
            elif t == "DIMENSION" and not a.notext:
                p = e.dxf.get("text_midpoint", None) or e.dxf.get("defpoint", None)
                if p is not None and inside(p.x, p.y):
                    try: m = e.get_measurement()
                    except Exception: m = ""
                    txt = e.dxf.get("text", "") or ""
                    val = txt.replace("<>", f"{m:.3f}" if isinstance(m, float) else str(m)) if txt and txt != "<>" else (f"{m:.3f}" if isinstance(m, float) else str(m))
                    ax.text(p.x, p.y, val, fontsize=6, color="m")
            elif t == "HATCH" and a.hatch:
                for path in e.paths:
                    vs = getattr(path, "vertices", None)
                    if vs:
                        pts = [(v[0], v[1]) for v in vs]; pts.append(pts[0]); seg(pts, "c-", 0.4)
            elif t == "INSERT" and depth < 4:
                for v in e.virtual_entities():
                    draw(v, depth + 1)
        except Exception:
            pass
    for e in lay:
        draw(e)
    ax.set_xlim(x1, x2); ax.set_ylim(y1, y2); ax.set_aspect("equal")
    fig.savefig(a.output, dpi=a.dpi, bbox_inches="tight", facecolor="white")
    print(a.output)

if __name__ == "__main__":
    main()
