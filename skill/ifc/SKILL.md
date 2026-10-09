---
name: ifc
description: Odczyt modeli BIM w formacie IFC (.ifc, .ifcZIP; IFC2x3, IFC4, IFC4.3) – struktura obiektu, elementy, właściwości, ilości i objętości, materiały, osie tras z pikietażem, georeferencja, podgląd PNG, walidacja. Używaj zawsze, gdy zadanie dotyczy pliku IFC. Nigdy nie odpowiadaj, że nie da się czytać IFC.
---

# Czytnik IFC (IfcOpenShell)

Do IFC służy biblioteka IfcOpenShell, instalowana z PyPI: ok. 20 sekund w zwykłej sesji, do ok. 3 minut, gdy w środowisku nie ma jeszcze numpy i matplotlib. Instalację uruchom w tle i w tym czasie pobierz model (krok 3). Skrypt pomocniczy `ifc_tool.py` leży w publicznym repozytorium Pracowni MiD `mdudek-mid/mid-narzedzia-` (nazwa kończy się myślnikiem), w katalogu `ifc/`. Procedura sprawdzona 2026-10-09 na modelach IFC2x3, IFC4 i IFC4.3: mosty, drogi, kolej, osie tras z klotoidami, cosinusoidami i przechyłką, georeferencja, zbrojenie, pliki .ifcZIP, polskie nazwy.

## Krok 1 – instalacja (raz na sesję)

Najpierw sprawdź, czy już jest: `test -f ~/.local/ifc/ifc_tool.py && python3 -c "import ifcopenshell" && echo jest`. Jeśli jest, przejdź do kroku 3.

```bash
pip install -q --break-system-packages ifcopenshell matplotlib pytest
rm -rf /tmp/mid-narzedzia
git clone -q --depth 1 https://github.com/mdudek-mid/mid-narzedzia- /tmp/mid-narzedzia
mkdir -p ~/.local/ifc && cp /tmp/mid-narzedzia/ifc/ifc_tool.py ~/.local/ifc/
python3 -c "import ifcopenshell; print('ifcopenshell', ifcopenshell.version)"
```

Repozytorium jest publiczne, więc nie trzeba go dołączać do sesji. Jeśli `git clone` zawiedzie, zapisz skrypt z kroku 2. Pakiet `pytest` jest potrzebny tylko do pełnej walidacji (`validate --rules`).

## Krok 2 – skrypt pomocniczy (tylko gdy klonowanie zawiodło)

Zapisz poniższy skrypt w całości jako `~/.local/ifc/ifc_tool.py` (narzędziem Write):

```python
#!/usr/bin/env python3
"""ifc_tool.py - odczyt modeli IFC (IFC2x3, IFC4, IFC4.3) przez IfcOpenShell.

Uzycie:
  python3 ifc_tool.py info      PLIK.ifc                 # schemat, autor/program, jednostki, georeferencja, struktura, klasy, materialy
  python3 ifc_tool.py tree      PLIK.ifc [--elements]    # drzewo przestrzenne (teren/most/czesc/budynek/kondygnacja)
  python3 ifc_tool.py elements  PLIK.ifc [--type IfcBeam,IfcSlab] [--props] [--geom]  > elementy.csv
  python3 ifc_tool.py quantities PLIK.ifc [--by type|material|container] [--geom]   # zestawienie ilosci
  python3 ifc_tool.py props     PLIK.ifc TEKST           # wszystko o elementach o danym GlobalId / fragmencie nazwy
  python3 ifc_tool.py alignment PLIK.ifc [--plot os.png] # osie tras IFC4.3: elementy geometryczne, niweleta, przechylka, pikietaz
  python3 ifc_tool.py render    PLIK.ifc [-o out.png] [--view plan|front|side|iso|elevation] [--in OBIEKT] [--type ...] [--exclude ...] [--window=x1,y1,x2,y2]
  (--in OBIEKT dziala tez dla elements i quantities: tylko elementy obiektu, ktorego nazwa zawiera tekst)
  python3 ifc_tool.py validate  PLIK.ifc                 # zgodnosc ze schematem IFC
"""
import argparse, csv, math, multiprocessing, os, sys
from collections import Counter, defaultdict

SKIP_RENDER = {"IfcOpeningElement", "IfcSpace", "IfcVirtualElement", "IfcAnnotation", "IfcGrid",
               "IfcAlignment", "IfcAlignmentSegment", "IfcAlignmentHorizontal", "IfcAlignmentVertical",
               "IfcAlignmentCant", "IfcReferent", "IfcSite", "IfcBuilding", "IfcBuildingStorey",
               "IfcBridge", "IfcBridgePart", "IfcRoad", "IfcRoadPart", "IfcRailway", "IfcRailwayPart",
               "IfcFacility", "IfcFacilityPart", "IfcMarineFacility", "IfcLinearPositioningElement"}


def open_model(path):
    import ifcopenshell
    return ifcopenshell.open(path)


def by_type(m, t):
    try:
        return m.by_type(t)
    except RuntimeError:  # klasy brak w tym schemacie (np. IfcAlignment w IFC2x3)
        return []


def unit_info(m):
    import ifcopenshell.util.unit as uu
    out = {}
    for kind in ("LENGTHUNIT", "AREAUNIT", "VOLUMEUNIT", "PLANEANGLEUNIT", "MASSUNIT"):
        try:
            u = uu.get_project_unit(m, kind)
            out[kind] = uu.get_full_unit_name(u) if u else ""
        except Exception:
            out[kind] = ""
    try:
        out["skala_do_m"] = uu.calculate_unit_scale(m)
    except Exception:
        out["skala_do_m"] = 1.0
    return out


def si_scales(m):
    """Mnozniki z jednostek projektu do SI: dlugosc->m, pole->m2, objetosc->m3, masa->kg."""
    import ifcopenshell.util.unit as uu
    out = {}
    for name, kind in (("dlugosc", "LENGTHUNIT"), ("pole", "AREAUNIT"), ("objetosc", "VOLUMEUNIT"), ("masa", "MASSUNIT")):
        try:
            out[name] = uu.calculate_unit_scale(m, kind) if uu.get_project_unit(m, kind) else 1.0
        except Exception:
            out[name] = 1.0
    return out


def path_of(e):
    """Sciezka przestrzenna elementu: Projekt > Teren > Most > Czesc ... (po nazwach)."""
    import ifcopenshell.util.element as ue
    parts, cur, guard = [], e, 0
    while cur is not None and guard < 20:
        guard += 1
        nxt = ue.get_container(cur) or ue.get_aggregate(cur) or (ue.get_nest(cur) if hasattr(ue, "get_nest") else None)
        if nxt is None:
            break
        parts.append(nxt.Name or nxt.is_a())
        cur = nxt
    return " > ".join(reversed(parts))


def materials_of(e):
    import ifcopenshell.util.element as ue
    try:
        return "; ".join(sorted({(m.Name or "") for m in ue.get_materials(e) if m}))
    except Exception:
        return ""


def iter_shapes(m, elements, threads=None):
    """Lista (element, geometria) w ukladzie swiata [m] dla podanych elementow.
    Jadro geometrii (C++) potrafi wypisywac na stdout diagnostyke - wyciszamy ja na poziomie deskryptora."""
    import ifcopenshell.geom
    s = ifcopenshell.geom.settings()
    s.set("use-world-coords", True)
    if not elements:
        return []
    out = []
    sys.stdout.flush()
    saved, devnull = os.dup(1), os.open(os.devnull, os.O_WRONLY)
    os.dup2(devnull, 1)
    try:
        it = ifcopenshell.geom.iterator(s, m, threads or multiprocessing.cpu_count(), include=list(elements))
        if it.initialize():
            while True:
                sh = it.get()
                out.append((m.by_id(sh.id), sh.geometry))
                if not it.next():
                    break
    finally:
        os.dup2(saved, 1)
        os.close(saved)
        os.close(devnull)
    return out


def geom_volumes(m, elements):
    """Objetosci z siatki [m3] - IfcOpenShell zwraca geometrie zawsze w metrach (SI)."""
    import ifcopenshell.util.shape as us
    vols = {}
    for el, g in iter_shapes(m, elements):
        try:
            vols[el.id()] = us.get_volume(g)
        except Exception:
            vols[el.id()] = None
    return vols


def physical_elements(m, types=None, inside=None):
    """Elementy fizyczne; types - lista klas; inside - fragment nazwy obiektu/czesci w sciezce
    przestrzennej (np. nazwa jednego mostu w modelu calego kontraktu)."""
    if types:
        out = []
        for t in types:
            out += by_type(m, t)
    else:
        out = [e for e in by_type(m, "IfcElement") if e.is_a() not in SKIP_RENDER]
    if inside:
        q = inside.lower()
        out = [e for e in out if q in path_of(e).lower() or q in (e.Name or "").lower()]
    return out


# ---------------------------------------------------------------- info / tree
def cmd_info(a):
    m = open_model(a.file)
    h = m.header
    fn = h.file_name
    print(f"Plik: {a.file} ({os.path.getsize(a.file) / 1e6:.1f} MB)")
    print(f"Schemat: {m.schema_identifier if hasattr(m, 'schema_identifier') else m.schema}")
    print(f"Program: {fn.originating_system} | preprocesor: {fn.preprocessor_version} | data: {fn.time_stamp}")
    print(f"Autor: {', '.join(fn.author or [])} | organizacja: {', '.join(fn.organization or [])}")
    try:
        print(f"Widok (MVD): {', '.join(h.file_description.description)}")
    except Exception:
        pass
    for p in by_type(m, "IfcProject"):
        print(f"Projekt: {p.Name} | {p.LongName or ''} | faza: {getattr(p, 'Phase', '') or ''}")
    u = unit_info(m)
    print(f"Jednostki: dlugosc={u['LENGTHUNIT']} pole={u['AREAUNIT']} objetosc={u['VOLUMEUNIT']} kat={u['PLANEANGLEUNIT']} (1 jedn. = {u['skala_do_m']} m)")
    for mc in by_type(m, "IfcMapConversion"):
        crs = mc.TargetCRS
        print(f"Georeferencja: CRS={getattr(crs, 'Name', '')} ({getattr(crs, 'Description', '') or ''}) "
              f"E={mc.Eastings} N={mc.Northings} H={mc.OrthogonalHeight} skala={getattr(mc, 'Scale', None)} "
              f"os X=({mc.XAxisAbscissa},{mc.XAxisOrdinate})")
    sites = by_type(m, "IfcSite")
    for s_ in sites:
        if getattr(s_, "RefLatitude", None):
            print(f"Teren {s_.Name}: szer={s_.RefLatitude} dl={s_.RefLongitude} wys={s_.RefElevation}")
    print("\nStruktura przestrzenna:")
    print_tree(m, max_depth=4, show_elements=False)
    els = physical_elements(m)
    print(f"\nElementy fizyczne ({len(els)}):")
    for k, v in Counter(e.is_a() for e in els).most_common():
        print(f"  {k:32s} {v}")
    mats = Counter()
    for e in els:
        for mm in materials_of(e).split("; "):
            if mm:
                mats[mm] += 1
    if mats:
        print("\nMaterialy (liczba elementow):")
        for k, v in mats.most_common(30):
            print(f"  {k:40s} {v}")
    print(f"\nZestawy wlasciwosci: {len(by_type(m, 'IfcPropertySet'))} | zestawy ilosci: {len(by_type(m, 'IfcElementQuantity'))} "
          f"| typy elementow: {len(by_type(m, 'IfcTypeObject'))}")
    al = by_type(m, "IfcAlignment")
    if al:
        print(f"Osie tras (IfcAlignment): {len(al)} -> szczegoly: ifc_tool.py alignment")


def print_tree(m, max_depth=10, show_elements=False):
    import ifcopenshell.util.element as ue
    projects = by_type(m, "IfcProject")

    def children(x):
        out = []
        for rel in getattr(x, "IsDecomposedBy", []) or []:
            out += [c for c in rel.RelatedObjects]
        return out

    def contained(x):
        out = []
        for rel in getattr(x, "ContainsElements", []) or []:
            out += list(rel.RelatedElements)
        return out

    def walk(x, depth):
        cont = contained(x)
        extra = f"  [{len(cont)} el.: " + ", ".join(f"{k} {v}" for k, v in Counter(c.is_a() for c in cont).most_common(5)) + "]" if cont else ""
        pt = ue.get_predefined_type(x) if hasattr(ue, "get_predefined_type") else ""
        print("  " * depth + f"- {x.is_a()} \"{x.Name or ''}\"" + (f" ({pt})" if pt and pt != "NOTDEFINED" else "") + extra)
        if show_elements:
            for c in cont:
                print("  " * (depth + 1) + f"* {c.is_a()} \"{c.Name or ''}\" {c.GlobalId}")
        if depth < max_depth:
            for c in children(x):
                if c.is_a("IfcSpatialElement") or c.is_a("IfcSpatialStructureElement") or c.is_a("IfcProject"):
                    walk(c, depth + 1)

    for p in projects:
        walk(p, 0)


def cmd_tree(a):
    print_tree(open_model(a.file), show_elements=a.elements)


# ---------------------------------------------------------------- elements / quantities / props
def flat_props(e):
    import ifcopenshell.util.element as ue
    out = {}
    try:
        for pset, props in ue.get_psets(e).items():
            for k, v in props.items():
                if k == "id":
                    continue
                out[f"{pset}.{k}"] = v
    except Exception:
        pass
    return out


def cmd_elements(a):
    m = open_model(a.file)
    import ifcopenshell.util.element as ue
    types = [t.strip() for t in a.type.split(",")] if a.type else None
    els = physical_elements(m, types, a.inside)
    vols = geom_volumes(m, els) if a.geom else {}
    rows, keys = [], []
    for e in els:
        t = ue.get_type(e)
        r = {"GlobalId": e.GlobalId, "klasa": e.is_a(), "typ_predef": ue.get_predefined_type(e) or "",
             "nazwa": e.Name or "", "ObjectType": getattr(e, "ObjectType", "") or "", "Tag": getattr(e, "Tag", "") or "",
             "typ_elementu": (t.Name if t else ""), "material": materials_of(e), "polozenie": path_of(e)}
        if a.geom:
            v = vols.get(e.id())
            r["objetosc_geom_m3"] = round(v, 4) if v is not None else ""
        if a.props:
            fp = flat_props(e)
            for k in fp:
                if k not in keys:
                    keys.append(k)
            r.update(fp)
        rows.append(r)
    base = ["GlobalId", "klasa", "typ_predef", "nazwa", "ObjectType", "Tag", "typ_elementu", "material", "polozenie"]
    if a.geom:
        base.append("objetosc_geom_m3")
    w = csv.DictWriter(sys.stdout, fieldnames=base + sorted(keys), extrasaction="ignore")
    w.writeheader()
    for r in rows:
        w.writerow(r)


QTO_KEYS = {"objetosc": ("NetVolume", "GrossVolume", "Volume"),
            "pole": ("NetArea", "GrossArea", "NetSurfaceArea", "GrossSurfaceArea", "NetSideArea", "GrossSideArea", "Area"),
            "dlugosc": ("Length", "NetLength", "GrossLength"),
            "masa": ("NetWeight", "GrossWeight", "Weight", "Mass")}


def qto_values(e):
    """Pierwsza znaleziona wartosc kazdej wielkosci z zestawow ilosci (Qto_*)."""
    import ifcopenshell.util.element as ue
    out = {}
    try:
        q = ue.get_psets(e, qtos_only=True)
    except Exception:
        q = {}
    for name, keys in QTO_KEYS.items():
        for qset in q.values():
            for k in keys:
                if isinstance(qset.get(k), (int, float)):
                    out[name] = qset[k]
                    break
            if name in out:
                break
    return out


def cmd_quantities(a):
    m = open_model(a.file)
    import ifcopenshell.util.element as ue
    els = physical_elements(m, None, a.inside)
    vols = geom_volumes(m, els) if a.geom else {}
    si = si_scales(m)
    agg = defaultdict(lambda: defaultdict(float))
    for e in els:
        if a.by == "material":
            key = materials_of(e) or "(brak materialu)"
        elif a.by == "container":
            key = path_of(e) or "(brak)"
        else:
            key = e.is_a() + (f" {ue.get_predefined_type(e)}" if ue.get_predefined_type(e) not in (None, "NOTDEFINED", "USERDEFINED") else "")
        d = agg[key]
        d["szt"] += 1
        for k, v in qto_values(e).items():
            d[k] += v * si[k]
            d[k + "_n"] += 1
        if a.geom and vols.get(e.id()) is not None:
            d["objetosc_geom_m3"] += vols[e.id()]
    w = csv.writer(sys.stdout)
    w.writerow([a.by, "szt", "objetosc_qto_m3", "pole_qto_m2", "dlugosc_qto_m", "masa_qto_kg"] + (["objetosc_geom_m3"] if a.geom else []) + ["uwagi"])
    for key, d in sorted(agg.items(), key=lambda kv: -kv[1]["szt"]):
        notes = [f"{k}: Qto dla {int(d[k + '_n'])}/{int(d['szt'])} el." for k in QTO_KEYS if 0 < d.get(k + "_n", 0) < d["szt"]]
        row = [key, int(d["szt"])] + [round(d[k], 4) if d.get(k + "_n") else "" for k in QTO_KEYS]
        if a.geom:
            row.append(round(d["objetosc_geom_m3"], 4))
        w.writerow(row + ["; ".join(notes)])
    print(f"# Qto przeliczone z jednostek projektu na SI (mnozniki {si}); objetosc_geom z siatki, tylko dla bryl zamknietych", file=sys.stderr)


def cmd_props(a):
    m = open_model(a.file)
    import ifcopenshell.util.element as ue
    hits = []
    try:
        hits = [m.by_guid(a.query)]
    except Exception:
        q = a.query.lower()
        hits = [e for e in by_type(m, "IfcProduct") if q in (e.Name or "").lower() or q in (getattr(e, "Tag", "") or "").lower()]
    if not hits:
        sys.exit("Nie znaleziono elementu o takim GlobalId/nazwie/Tag.")
    for e in hits[:a.limit]:
        t = ue.get_type(e)
        print(f"=== {e.is_a()} \"{e.Name}\" GlobalId={e.GlobalId} #{e.id()}")
        info = e.get_info(recursive=False)
        for k, v in info.items():
            if k in ("id", "type") or v is None or k.startswith("Owner") or hasattr(v, "is_a"):
                continue
            print(f"  {k}: {v}")
        print(f"  polozenie: {path_of(e)}")
        print(f"  typ: {t.is_a() + ' ' + str(t.Name) if t else '-'}")
        print(f"  material: {materials_of(e) or '-'}")
        for pset, props in ue.get_psets(e).items():
            print(f"  [{pset}]")
            for k, v in props.items():
                if k != "id":
                    print(f"     {k} = {v}")
    if len(hits) > a.limit:
        print(f"... i {len(hits) - a.limit} kolejnych (zwieksz --limit)")


# ---------------------------------------------------------------- alignment (IFC4.3)
def nested(x):
    out = []
    for rel in getattr(x, "IsNestedBy", []) or []:
        out += list(rel.RelatedObjects)
    return out


def station_str(s):
    sign = "-" if s < 0 else ""
    s = abs(s)
    km = int(s // 1000)
    return f"{sign}{km}+{s - km * 1000:07.3f}"


# przebieg krzywizny na krzywej przejsciowej: k(t) = k0 + (k1-k0)*f(t), t = s/L
TRANSITION = {
    "CLOTHOID": lambda t: t,
    "CUBIC": lambda t: t,                     # parabola 3. stopnia ~ klotoida (przyblizenie)
    "COSINECURVE": lambda t: (1 - math.cos(math.pi * t)) / 2,
    "SINECURVE": lambda t: t - math.sin(2 * math.pi * t) / (2 * math.pi),
    "BLOSSCURVE": lambda t: 3 * t * t - 2 * t ** 3,
    "HELMERTCURVE": lambda t: 2 * t * t if t <= 0.5 else 1 - 2 * (1 - t) ** 2,
    "VIENNESEBEND": lambda t: 35 * t ** 4 - 84 * t ** 5 + 70 * t ** 6 - 20 * t ** 7,
}


def horiz_points(segs, step=1.0, sub=10):
    """Os w planie z parametrow projektowych (prosta, luk, krzywe przejsciowe wg typu),
    calkowanie kierunku metoda punktu srodkowego."""
    pts = []
    for p in segs:
        x, y = p.StartPoint.Coordinates[:2]
        th = p.StartDirection
        L = abs(p.SegmentLength or 0.0)
        k0 = 1 / p.StartRadiusOfCurvature if p.StartRadiusOfCurvature else 0.0
        k1 = 1 / p.EndRadiusOfCurvature if p.EndRadiusOfCurvature else 0.0
        t = p.PredefinedType
        if t == "LINE":
            f = lambda u: 0.0; k0 = k1 = 0.0
        elif t == "CIRCULARARC":
            f = lambda u: 0.0; k1 = k0
        else:
            f = TRANSITION.get(t, lambda u: u)
        n = max(2, int(L / step) + 1)
        h = L / ((n - 1) * sub) if L else 0
        seg_pts = [(x, y)]
        s = 0.0
        for i in range(1, n):
            for _ in range(sub):
                sm = s + h / 2
                k_mid = k0 + (k1 - k0) * f(sm / L) if L else 0
                # krok o dlugosci h w kierunku ze srodka kroku
                x += h * math.cos(th + k_mid * h / 2)
                y += h * math.sin(th + k_mid * h / 2)
                th += k_mid * h
                s += h
            seg_pts.append((x, y))
        pts.append(seg_pts)
    return pts


def start_station(al):
    """Pikietaz poczatku osi: IfcReferent typu STATION z Pset_Stationing (pierwszy znaleziony)."""
    import ifcopenshell.util.element as ue
    for r in nested(al):
        if r.is_a("IfcReferent") and (r.PredefinedType or "") == "STATION":
            st = ue.get_pset(r, "Pset_Stationing") or {}
            if st.get("Station") is not None:
                return st["Station"], r.Name
    return 0.0, None


def vert_profile(segs, step=1.0):
    out = []
    for p in segs:
        s0, L, h0 = p.StartDistAlong, p.HorizontalLength, p.StartHeight
        g0, g1 = p.StartGradient, p.EndGradient
        n = max(2, int(L / step) + 1)
        for i in range(n):
            x = L * i / (n - 1)
            h = h0 + g0 * x + ((g1 - g0) / (2 * L) * x * x if L else 0)
            out.append((s0 + x, h))
    return out


def cmd_alignment(a):
    m = open_model(a.file)
    sc = unit_info(m)["skala_do_m"]
    als = by_type(m, "IfcAlignment")
    if not als:
        print("Brak osi tras (IfcAlignment). Osie sa dostepne od IFC4.3; w IFC2x3/IFC4 trasa bywa zapisana jako zwykla geometria.")
        return
    plots = []
    for al in als:
        print(f"=== Os: \"{al.Name}\" {al.GlobalId} {al.PredefinedType if hasattr(al, 'PredefinedType') else ''}")
        st0, st0_name = start_station(al)
        print(f"  Pikietaz poczatku: km {station_str(st0 * sc)}" + (f" (referent \"{st0_name}\")" if st0_name else " (brak referenta STATION - liczony od 0)"))
        layouts = nested(al)
        hsegs, vsegs = [], []
        for lay in layouts:
            segs = [x for x in nested(lay) if x.is_a("IfcAlignmentSegment")]
            params = [x.DesignParameters for x in segs if x.DesignParameters]
            if lay.is_a("IfcAlignmentHorizontal"):
                hsegs = params
                tot = sum(p.SegmentLength or 0 for p in params)
                print(f"  Plan: {len(params)} elementow, dlugosc {tot * sc:.3f} m")
                s = 0.0
                for i, p in enumerate(params):
                    r0 = f"R={p.StartRadiusOfCurvature * sc:.2f}" if p.StartRadiusOfCurvature else "R=inf"
                    r1 = f"R={p.EndRadiusOfCurvature * sc:.2f}" if p.EndRadiusOfCurvature else "R=inf"
                    print(f"    {i + 1:3d}. km {station_str((st0 + s) * sc)}  {p.PredefinedType:16s} L={p.SegmentLength * sc:10.3f}  {r0} -> {r1}"
                          + (f"  A={math.sqrt(abs(p.SegmentLength * (p.EndRadiusOfCurvature or p.StartRadiusOfCurvature or 0))) * sc:.2f}" if p.PredefinedType == "CLOTHOID" and (p.EndRadiusOfCurvature or p.StartRadiusOfCurvature) else ""))
                    s += abs(p.SegmentLength or 0)
            elif lay.is_a("IfcAlignmentVertical"):
                vsegs = params
                print(f"  Niweleta: {len(params)} elementow")
                for i, p in enumerate(params):
                    rad = f"  R={p.RadiusOfCurvature * sc:.1f}" if getattr(p, "RadiusOfCurvature", None) else ""
                    print(f"    {i + 1:3d}. km {station_str((st0 + p.StartDistAlong) * sc)}  {p.PredefinedType:16s} L={p.HorizontalLength * sc:9.3f}  "
                          f"H0={p.StartHeight * sc:.3f}  i={p.StartGradient * 100:+.3f}% -> {p.EndGradient * 100:+.3f}%{rad}")
            elif lay.is_a("IfcAlignmentCant"):
                print(f"  Przechylka: {len(params)} elementow, rozstaw szyn {getattr(lay, 'RailHeadDistance', '')}")
                for i, p in enumerate(params):
                    print(f"    {i + 1:3d}. km {station_str((st0 + p.StartDistAlong) * sc)}  {p.PredefinedType:16s} L={p.HorizontalLength * sc:9.3f}  "
                          f"lewa {p.StartCantLeft}->{p.EndCantLeft}  prawa {p.StartCantRight}->{p.EndCantRight}")
        # kontrola ciaglosci planu: koniec obliczony vs poczatek nastepnego
        if hsegs:
            pts = horiz_points(hsegs)
            gaps = []
            for i in range(len(hsegs) - 1):
                ex, ey = pts[i][-1]
                nx, ny = hsegs[i + 1].StartPoint.Coordinates[:2]
                gaps.append(math.hypot(ex - nx, ey - ny) * sc)
            if gaps:
                print(f"  Kontrola ciaglosci planu: max odchylka konca elementu od poczatku nastepnego {max(gaps):.4f} m")
            prof = [(st0 + s, z) for s, z in vert_profile(vsegs)] if vsegs else []
            plots.append((al.Name or al.GlobalId, pts, prof))
        refs = [r for r in nested(al) if r.is_a("IfcReferent")]
        if refs:
            print(f"  Punkty odniesienia (IfcReferent): {len(refs)}")
            import ifcopenshell.util.element as ue
            for r in refs[:40]:
                st = ue.get_pset(r, "Pset_Stationing") or {}
                print(f"    {r.Name or ''} {r.PredefinedType or ''} {('km ' + station_str(st['Station'] * sc)) if st.get('Station') is not None else ''}")
    if a.plot and plots:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        has_v = any(p[2] for p in plots)
        fig, axes = plt.subplots(2 if has_v else 1, 1, figsize=(14, 10 if has_v else 7), squeeze=False)
        ax = axes[0][0]
        for name, pts, _ in plots:
            for seg in pts:
                xs, ys = zip(*seg)
                ax.plot([x * sc for x in xs], [y * sc for y in ys], lw=1.5)
            ax.plot([seg[0][0] * sc for seg in pts], [seg[0][1] * sc for seg in pts], "k|", ms=8)
        ax.set_aspect("equal", adjustable="datalim"); ax.set_title("Plan osi (kreski = poczatki elementow)"); ax.grid(alpha=.3)
        ax.ticklabel_format(useOffset=False, style="plain")
        if has_v:
            ax2 = axes[1][0]
            for name, _, prof in plots:
                if prof:
                    s, h = zip(*prof)
                    ax2.plot([x * sc for x in s], [y * sc for y in h], lw=1.5, label=name)
            ax2.set_title("Niweleta"); ax2.set_xlabel("pikietaz [m]"); ax2.set_ylabel("wysokosc [m]"); ax2.grid(alpha=.3); ax2.legend()
        fig.tight_layout()
        fig.savefig(a.plot, dpi=130)
        print(a.plot)


# ---------------------------------------------------------------- render
def cmd_render(a):
    import numpy as np
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.collections import PolyCollection
    import ifcopenshell.util.shape as us

    m = open_model(a.file)
    types = [t.strip() for t in a.type.split(",")] if a.type else None
    excl = {t.strip() for t in a.exclude.split(",")} if a.exclude else set()
    els = [e for e in physical_elements(m, types, a.inside) if e.is_a() not in excl]
    tris, cols = [], []
    for el, g in iter_shapes(m, els):
        v = np.array(g.verts, dtype=float).reshape(-1, 3)
        f = np.array(g.faces, dtype=int).reshape(-1, 3)
        if not len(f):
            continue
        try:
            mc = us.get_material_colors(g)
            mid = np.array(g.material_ids, dtype=int)
            c = np.array([mc[i][:3] if 0 <= i < len(mc) else (0.7, 0.7, 0.7) for i in mid])
        except Exception:
            c = np.tile((0.7, 0.7, 0.7), (len(f), 1))
        tris.append(v[f])
        cols.append(c)
    if not tris:
        sys.exit("Brak bryl do narysowania: sprawdz --type / --exclude / --in. Jesli model zawiera tylko os trasy, uzyj: alignment --plot os.png")
    T = np.concatenate(tris)            # (n,3,3)
    C = np.concatenate(cols)
    views = {"plan": (-90, 90), "front": (-90, 0), "side": (0, 0), "iso": (-50, 30)}
    if a.view == "elevation":           # widok z boku wzdluz najdluzszego kierunku obiektu (np. most)
        xy = T.reshape(-1, 3)[:, :2]
        xy = xy - xy.mean(0)
        w, vec = np.linalg.eigh(np.cov(xy.T))
        main = vec[:, -1]
        az = math.degrees(math.atan2(main[1], main[0])) - 90
        el_deg = 0
    else:
        az, el_deg = views[a.view]
    az, el_deg = (a.azimuth if a.azimuth is not None else az), (a.elev if a.elev is not None else el_deg)
    azr, elr = math.radians(az), math.radians(el_deg)
    d = np.array([math.cos(elr) * math.cos(azr), math.cos(elr) * math.sin(azr), math.sin(elr)])  # od obiektu do kamery
    up0 = np.array([0, 0, 1.0]) if abs(d[2]) < 0.99 else np.array([0, 1.0, 0])  # z gory: X w prawo, Y (polnoc) w gore
    right = np.cross(up0, d); right /= np.linalg.norm(right)
    up = np.cross(d, right)
    P = T @ np.stack([right, up, d], axis=1)  # (n,3,3) -> wspolrzedne ekranu + glebokosc
    depth = P[:, :, 2].mean(1)
    order = np.argsort(depth)          # najpierw najdalsze
    n = np.cross(T[:, 1] - T[:, 0], T[:, 2] - T[:, 0])
    nl = np.linalg.norm(n, axis=1); nl[nl == 0] = 1
    shade = 0.45 + 0.55 * np.abs((n / nl[:, None]) @ (d * 0.8 + up * 0.2) / np.linalg.norm(d * 0.8 + up * 0.2))
    F = np.clip(C * shade[:, None], 0, 1)
    poly = P[order][:, :, :2]
    if a.window:
        x1, y1, x2, y2 = (float(v) for v in a.window.split(","))
        x1, x2 = sorted((x1, x2)); y1, y2 = sorted((y1, y2))
    else:
        x1, y1 = poly[:, :, 0].min(), poly[:, :, 1].min()
        x2, y2 = poly[:, :, 0].max(), poly[:, :, 1].max()
    mx, my = (x2 - x1) * 0.03 or 1, (y2 - y1) * 0.03 or 1
    x1, x2, y1, y2 = x1 - mx, x2 + mx, y1 - my, y2 + my
    h = min(max(a.width * (y2 - y1) / (x2 - x1), 2), 30)
    fig = plt.figure(figsize=(a.width, h))
    ax = fig.add_axes([0, 0, 1, 1])
    ax.add_collection(PolyCollection(poly, facecolors=F[order], edgecolors=(0, 0, 0, 0.15) if len(poly) < 20000 else "none", linewidths=0.2))
    ax.set_xlim(x1, x2); ax.set_ylim(y1, y2); ax.set_aspect("equal", adjustable="box"); ax.axis("off")
    out = a.output or os.path.splitext(os.path.basename(a.file))[0] + f"_{a.view}.png"
    fig.savefig(out, dpi=a.dpi, facecolor="white")
    print(f"{out}  ({len(els)} elementow, {len(poly)} trojkatow, widok az={az:.1f} el={el_deg:.1f})")


# ---------------------------------------------------------------- validate
def cmd_validate(a):
    import ifcopenshell.validate
    m = open_model(a.file)
    log = ifcopenshell.validate.json_logger()
    ifcopenshell.validate.validate(m, log, express_rules=a.rules)
    stmts = log.statements
    print(f"Bledy zgodnosci ze schematem {m.schema}: {len(stmts)}")
    for s_ in stmts[:a.limit]:
        print(f"  - {s_.get('message', s_)}"[:400])
    if len(stmts) > a.limit:
        print(f"  ... i {len(stmts) - a.limit} kolejnych")


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = p.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("info"); s.add_argument("file")
    s = sub.add_parser("tree"); s.add_argument("file"); s.add_argument("--elements", action="store_true")
    s = sub.add_parser("elements"); s.add_argument("file"); s.add_argument("--type")
    s.add_argument("--props", action="store_true", help="dolacz wszystkie Pset/Qto jako kolumny")
    s.add_argument("--geom", action="store_true", help="dolacz objetosc liczona z geometrii [m3]")
    s.add_argument("--in", dest="inside", help="tylko elementy obiektu/czesci o nazwie zawierajacej TEKST")
    s = sub.add_parser("quantities"); s.add_argument("file"); s.add_argument("--by", choices=["type", "material", "container"], default="type")
    s.add_argument("--geom", action="store_true", help="dodaj objetosc liczona z geometrii [m3]")
    s.add_argument("--in", dest="inside", help="tylko elementy obiektu/czesci o nazwie zawierajacej TEKST")
    s = sub.add_parser("props"); s.add_argument("file"); s.add_argument("query"); s.add_argument("--limit", type=int, default=20)
    s = sub.add_parser("alignment"); s.add_argument("file"); s.add_argument("--plot", help="zapisz plan osi i niwelete do PNG")
    s = sub.add_parser("render"); s.add_argument("file"); s.add_argument("-o", "--output")
    s.add_argument("--view", choices=["plan", "front", "side", "iso", "elevation"], default="iso")
    s.add_argument("--azimuth", type=float, help="wlasny kierunek kamery w planie [deg od osi X]")
    s.add_argument("--elev", type=float, help="wlasne nachylenie kamery [deg]")
    s.add_argument("--type", help="tylko te klasy, np. IfcBeam,IfcSlab")
    s.add_argument("--exclude", help="pomin te klasy, np. IfcEarthworksFill")
    s.add_argument("--in", dest="inside", help="tylko elementy obiektu/czesci o nazwie zawierajacej TEKST")
    s.add_argument("--window", help="x1,y1,x2,y2 [m] we wspolrzednych ekranu; dla widoku plan = wspolrzedne X,Y modelu")
    s.add_argument("--dpi", type=int, default=150); s.add_argument("--width", type=float, default=16)
    s = sub.add_parser("validate"); s.add_argument("file"); s.add_argument("--rules", action="store_true", help="takze reguly EXPRESS (wolniej)")
    s.add_argument("--limit", type=int, default=50)
    a = p.parse_args()
    {"info": cmd_info, "tree": cmd_tree, "elements": cmd_elements, "quantities": cmd_quantities, "props": cmd_props,
     "alignment": cmd_alignment, "render": cmd_render, "validate": cmd_validate}[a.cmd](a)


if __name__ == "__main__":
    main()
```

## Krok 3 – skąd wziąć model

- Załącznik w rozmowie: czytaj po nazwie pliku.
- Plik na komputerze użytkownika (folder projektu, dokumentacja przetargowa, archiwum): skopiuj go do środowiska roboczego narzędziem do pobierania plików z komputera i pracuj na kopii. Na komputerze użytkownika (Windows) niczego nie instaluj.
- Oryginałów nie zmieniaj. `.ifcZIP` czyta się bezpośrednio, bez rozpakowywania.

## Krok 4 – czytanie

```bash
T=~/.local/ifc/ifc_tool.py
python3 $T info model.ifc                       # schemat, program i autor, jednostki, georeferencja, drzewo obiektu, klasy, materiały
python3 $T tree model.ifc --elements            # pełne drzewo z elementami i GlobalId
python3 $T elements model.ifc --props --geom > elementy.csv   # każdy element: klasa, nazwa, typ, materiał, położenie, wszystkie Pset/Qto, objętość z geometrii
python3 $T elements model.ifc --type IfcBeam,IfcSlab --in "WD-12"
python3 $T quantities model.ifc --geom          # zestawienie ilości wg klas (Qto z modelu w SI + objętość z geometrii)
python3 $T quantities model.ifc --by material --geom          # albo --by container (wg obiektów/części)
python3 $T props model.ifc "P1"                 # wszystko o elementach o danym GlobalId / fragmencie nazwy lub Tag
python3 $T alignment model.ifc --plot os.png    # osie tras IFC4.3: plan, niweleta, przechyłka, pikietaż, kontrola ciągłości
python3 $T render model.ifc --view iso -o iso.png          # widoki: plan, front, side, iso, elevation
python3 $T render model.ifc --view elevation --in "WD-12" --exclude IfcEarthworksFill -o widok.png
python3 $T validate model.ifc                   # zgodność ze schematem; --rules = także reguły EXPRESS (wolniej)
```

Wygenerowany PNG obejrzyj narzędziem Read. Wyniki CSV można od razu przenieść do arkusza.

Dobra kolejność pracy:
- Zacznij od `info`: schemat, program autorski, jednostki, georeferencja, drzewo obiektu, klasy i materiały. To odpowiada na większość pytań o to, co jest w modelu.
- Model kontraktu z kilkoma obiektami: zawęź pracę opcją `--in "fragment nazwy obiektu lub części"` (działa w `elements`, `quantities` i `render`).
- Widok `elevation` ustawia kamerę prostopadle do najdłuższego kierunku obiektu, czyli daje widok boczny mostu. Najlepiej działa z `--in` dla jednego obiektu.
- Przy ilościach porównuj kolumny `objetosc_qto_m3` (zapisane w modelu) i `objetosc_geom_m3` (policzone z brył). Kolumna `uwagi` mówi, ile elementów w grupie ma ilości w modelu. Nie sumuj na ślepo, gdy pokrycie jest niepełne.

## Znane pułapki (sprawdzone)

- **Jednostki.** Ilości w zestawach Qto są w jednostkach projektu (często mm dla długości). `quantities` przelicza je na m, m², m³ i kg, ale `elements --props` i `props` pokazują surowe wartości, więc sprawdź jednostki w `info`. Geometria z IfcOpenShell jest zawsze w metrach, także współrzędne w `--window`.
- **Objętość z geometrii** jest wiarygodna tylko dla zamkniętych brył. `IfcElementAssembly` ma zwykle 0, bo geometrię niosą jego części.
- **IFC2x3 i IFC4** nie mają klas `IfcBridge`, `IfcRoad`, `IfcAlignment`. Most bywa zapisany jako `IfcBuilding`, a elementy jako `IfcBuildingElementProxy` z typem w `ObjectType`. Oś trasy jest wtedy zwykłą geometrią.
- **Pikietaż osi** liczony jest od referenta typu STATION (`Pset_Stationing`). Bez niego skrypt liczy od 0 i wyraźnie o tym informuje.
- **Kontrola ciągłości planu** porównuje obliczony koniec każdego elementu z początkiem następnego. Na poprawnych osiach wynik to 0,0000 m. Odchyłka powyżej 1 cm wskazuje błąd w modelu albo nietypową krzywą. Obsługiwane krzywe: prosta, łuk, klotoida, cosinusoida, sinusoida, Bloss, Helmert, krzywa wiedeńska oraz parabola 3. stopnia (przybliżana klotoidą).
- **Polskie znaki** (w pliku zakodowane jako `\X2\...`) IfcOpenShell dekoduje automatycznie.
- **Opcja `--window` wymaga znaku `=`**, bo ujemne współrzędne bez niego wyglądają jak opcje.
- **Duże modele (100+ MB)** wczytują się dłużej i zajmują dużo pamięci. Podgląd zawężaj opcjami `--in`, `--type` lub `--exclude`.
- **Tworzenie geometrii** przez `ifcopenshell.api` (gdy trzeba coś zapisać do IFC) przyjmuje wymiary w metrach, niezależnie od jednostek projektu.
