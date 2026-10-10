#!/usr/bin/env python3
"""opis_tool.py - opisy techniczne Pracowni MiD wg Instrukcji 02
(przedmiot opracowania, lokalizacja, cel i zakres, podstawa opracowania, wykorzystane materialy).

Komendy:
  wzor [--typ PB|PT|PW|PZT|koncepcja|ekspertyza|STWiORB|PFU|inne] [--out opis.json]
  lokalizacja (--xy X,Y [--uklad 2180|4326] | --miejscowosc NAZWA [--gmina G] | --dzialka ID_ULDK)
              --out KATALOG [--promien 1500]
        -> lokalizacja.json (wojewodztwo, powiat, gmina, obreb, dzialka), rys_polska.png, rys_topo.png, rys_orto.png
           (uslugi GUGiK: ULDK, UUG, WMS PRG/TOPO/ORTO/KIEG)
  dzialki (--id ID[,ID..] | --xy X,Y[;X,Y..] [--uklad 2180|4326]) [--md]   wykaz dzialek z ULDK
  przepisy [--szukaj TEKST] [--klucz K]        katalog aktow prawnych z aktualnym t.j. (API ELI Sejmu)
  cytuj ELI|klucz                              poprawny cytat aktu na dzis (np. cytuj DU/1994/414, cytuj pb)
  generuj SPEC.json --out Opis.docx [--md Opis.md] [--offline]
  sprawdz PLIK(.docx/.pdf/.md/.txt) [--offline] [--out raport.md]
        zgodnosc z Instrukcja 02 + aktualnosc przywolanych aktow (ELI) + odwolania [X#] + powtorzenia tresci
"""
import argparse, json, os, re, sys, time, unicodedata
from datetime import date
from pathlib import Path

ELI_API = "https://api.sejm.gov.pl/eli/acts/"
CACHE = Path(os.environ.get("XDG_CACHE_HOME", Path.home() / ".cache")) / "opis_tool_eli.json"
KATEGORIE = [("DA", "Dokumentacja archiwalna"), ("N", "Normy"), ("U", "Ustawy"), ("R", "Rozporządzenia"),
             ("W", "Wytyczne"), ("L", "Literatura"), ("I", "Źródła internetowe"), ("P", "Pozostałe")]
MIESIACE = ["stycznia", "lutego", "marca", "kwietnia", "maja", "czerwca", "lipca", "sierpnia", "września",
            "października", "listopada", "grudnia"]

# Katalog aktow (ELI aktu pierwotnego). Aktualny tekst jednolity i nowelizacje pobierane sa z API ELI.
KATALOG_AKTOW = {
    "pb": ("U", "DU/1994/414"), "udp": ("U", "DU/1985/60"), "specustawa_drogowa": ("U", "DU/2003/721"),
    "prawo_wodne": ("U", "DU/2017/1566"), "oos": ("U", "DU/2008/1227"), "pos": ("U", "DU/2001/627"),
    "ochrona_przyrody": ("U", "DU/2004/880"), "pgik": ("U", "DU/1989/163"), "zabytki": ("U", "DU/2003/1568"),
    "planowanie": ("U", "DU/2003/717"), "transport_kolejowy": ("U", "DU/2003/789"), "wyroby_budowlane": ("U", "DU/2004/881"),
    "pzp": ("U", "DU/2019/2019"), "samorzady_zawodowe": ("U", "DU/2001/42"), "prawo_autorskie": ("U", "DU/1994/83"),
    "prd": ("U", "DU/1997/602"),
    "ptb_drogi_2022": ("R", "DU/2022/1518"), "pb_zakres_formy": ("R", "DU/2020/1609"),
    "dok_projektowa_stwiorb_pfu": ("R", "DU/2021/2454"), "bioz": ("R", "DU/2003/1126"),
    "numeracja_ewidencja_2005": ("R", "DU/2005/582"), "znaki_sygnaly": ("R", "DU/2002/1393"),
    "warunki_znaki": ("R", "DU/2003/2181"), "zarzadzanie_ruchem": ("R", "DU/2017/784"),
    "skrzyzowania_kolej": ("R", "DU/2025/1105"),
}
# Normy i wytyczne - katalog pomocniczy (aktualnosc norm sprawdzac w katalogu PKN; API brak)
KATALOG_NORM = {
    "PN-EN 1990": "PN-EN 1990:2004 Eurokod. Podstawy projektowania konstrukcji",
    "PN-EN 1990/A1": "PN-EN 1990:2004/A1:2008 Eurokod. Podstawy projektowania konstrukcji. Zmiana A1 (mosty)",
    "PN-EN 1991-1-1": "PN-EN 1991-1-1:2004 Eurokod 1. Oddziaływania na konstrukcje. Część 1-1: Oddziaływania ogólne",
    "PN-EN 1991-1-4": "PN-EN 1991-1-4:2008 Eurokod 1. Oddziaływania na konstrukcje. Część 1-4: Oddziaływania wiatru",
    "PN-EN 1991-1-5": "PN-EN 1991-1-5:2005 Eurokod 1. Oddziaływania na konstrukcje. Część 1-5: Oddziaływania termiczne",
    "PN-EN 1991-1-7": "PN-EN 1991-1-7:2008 Eurokod 1. Oddziaływania na konstrukcje. Część 1-7: Oddziaływania wyjątkowe",
    "PN-EN 1991-2": "PN-EN 1991-2:2007 Eurokod 1. Oddziaływania na konstrukcje. Część 2: Obciążenia ruchome mostów",
    "PN-EN 1992-1-1": "PN-EN 1992-1-1:2008 Eurokod 2. Projektowanie konstrukcji z betonu. Część 1-1",
    "PN-EN 1992-2": "PN-EN 1992-2:2010 Eurokod 2. Projektowanie konstrukcji z betonu. Część 2: Mosty z betonu",
    "PN-EN 1993-1-1": "PN-EN 1993-1-1:2006 Eurokod 3. Projektowanie konstrukcji stalowych. Część 1-1",
    "PN-EN 1993-2": "PN-EN 1993-2:2010 Eurokod 3. Projektowanie konstrukcji stalowych. Część 2: Mosty stalowe",
    "PN-EN 1994-2": "PN-EN 1994-2:2010 Eurokod 4. Projektowanie zespolonych konstrukcji stalowo-betonowych. Część 2: Mosty",
    "PN-EN 1997-1": "PN-EN 1997-1:2008 Eurokod 7. Projektowanie geotechniczne. Część 1: Zasady ogólne",
    "PN-EN 206": "PN-EN 206+A2:2021-08 Beton. Wymagania, właściwości użytkowe, produkcja i zgodność",
    "PN-EN 13791": "PN-EN 13791:2019-12 Ocena wytrzymałości betonu na ściskanie w konstrukcjach i prefabrykowanych wyrobach betonowych",
    "PN-EN 14630": "PN-EN 14630:2007 Wyroby i systemy do ochrony i napraw konstrukcji betonowych. Oznaczanie głębokości karbonatyzacji",
    "PN-85/S-10030": "PN-85/S-10030 Obiekty mostowe. Obciążenia (norma wycofana – stosowana do oceny obiektów istniejących)",
}
WYCOFANE_NORMY = {"PN-85/S-10030": "wycofana; dopuszczalna do oceny obiektów zaprojektowanych wg niej",
                  "PN-91/S-10042": "wycofana; obiekty betonowe projektuje się wg PN-EN 1992",
                  "PN-82/S-10052": "wycofana; obiekty stalowe projektuje się wg PN-EN 1993",
                  "PN-B-03264": "wycofana dla obiektów mostowych", "PN-66/B-02015": "historyczny normatyw obciążeń"}


# ----------------------------------------------------------------------------- ELI (API Sejmu)
def _cache():
    try:
        return json.loads(CACHE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def eli_get(eli, offline=False):
    eli = eli.strip().strip("/")
    c = _cache()
    rec = c.get(eli)
    if rec and (offline or time.time() - rec.get("_t", 0) < 7 * 86400):
        return rec
    if offline:
        return None
    import urllib.request
    for proba in range(3):
        try:
            d = json.load(urllib.request.urlopen(ELI_API + eli, timeout=25))
            break
        except Exception as e:
            if proba == 2 or "403" in str(e) or "404" in str(e):
                return rec or {"_blad": str(e)}
            time.sleep(2 * (proba + 1))
    d = {k: d.get(k) for k in ("ELI", "title", "type", "status", "inForce", "displayAddress", "announcementDate",
                               "repealDate", "expirationDate", "references", "year", "pos")}
    d["_t"] = time.time()
    c[eli] = d
    try:
        CACHE.parent.mkdir(parents=True, exist_ok=True)
        CACHE.write_text(json.dumps(c, ensure_ascii=False), encoding="utf-8")
    except Exception:
        pass
    return d


def _refs(d, klucz):
    return [x["id"] for x in ((d or {}).get("references") or {}).get(klucz, [])]


def _po(a, b):
    """czy akt a (DU/RRRR/POZ) ogloszono po b"""
    ya, pa = map(int, a.split("/")[1:3])
    yb, pb = map(int, b.split("/")[1:3])
    return (ya, pa) > (yb, pb)


def _adres(eli):
    _, y, p = eli.split("/")
    return f"Dz.U. {y} poz. {p}"


def eli_rozwiaz(eli, offline=False):
    """Zwraca opis aktu: akt bazowy, status, aktualny t.j., czy sa zmiany po t.j., czym uchylony, cytat."""
    d = eli_get(eli, offline)
    if not d or d.get("_blad"):
        return {"eli": eli, "blad": (d or {}).get("_blad", "brak danych (offline)")}
    wynik = {"eli": eli, "cytowany_typ": d.get("type")}
    baza = eli
    if d.get("type") == "Obwieszczenie":
        b = _refs(d, "Tekst jednolity dla aktu")
        if b:
            baza = b[0]
            d2 = eli_get(baza, offline)
            if not d2 or d2.get("_blad"):
                return {"eli": eli, "blad": (d2 or {}).get("_blad", "brak danych aktu bazowego (offline)")}
            d = d2
    wynik.update(baza=baza, adres=d.get("displayAddress") or _adres(baza),
                 tytul=re.sub(r"\.$", "", (d.get("title") or "")).replace(" - ", " – "),
                 typ=d.get("type"), status=d.get("status"), w_mocy=d.get("inForce") == "IN_FORCE",
                 uchylony=d.get("repealDate") or d.get("expirationDate"))
    tj = _refs(d, "Inf. o tekście jednolitym")
    tj_akt = tj[0] if tj else None
    zmiany = _refs(d, "Akty zmieniające")
    if tj_akt:
        dtj = eli_get(tj_akt, offline) or {}
        po = _refs(dtj, "Nowelizacje po tekście jednolitym")
        wynik.update(tj=tj_akt, zmiany_po_tj=bool(po) or any(_po(z, tj_akt) for z in zmiany))
    else:
        wynik.update(tj=None, zmiany_po_tj=bool(zmiany))
    wynik["uchylony_przez"] = _refs(d, "Uchylenia wynikające z") or _refs(d, "Akty uchylające")
    wynik["cytat"] = cytat(wynik)
    return wynik


def cytat(w):
    if w.get("blad"):
        return w["eli"]
    adres = (f"t.j. {_adres(w['tj'])}" if w.get("tj") else (w.get("adres") or _adres(w["baza"])))
    zm = " z późn. zm." if w.get("zmiany_po_tj") else ""
    return f"{w['tytul']} ({adres}{zm})"


def eli_szukaj(tekst):
    """Wyszukiwarka ELI dopasowuje tytul po fragmencie; pytamy o najdluzsze slowo i filtrujemy po pozostalych."""
    import urllib.parse, urllib.request
    slowa = [w for w in re.findall(r"\w+", tekst) if len(w) > 3] or [tekst]
    klucz = max(slowa, key=len)
    wyn = []
    for off in range(0, 2000, 500):
        q = urllib.parse.urlencode({"title": klucz, "publisher": "DU", "limit": 500, "offset": off})
        try:
            d = json.load(urllib.request.urlopen(ELI_API + "search?" + q, timeout=40))
        except Exception as e:
            sys.exit(f"Wyszukiwarka ELI niedostępna: {e}")
        it = d.get("items") or []
        wyn += it
        if len(it) < 500:
            break
    rdz = [_norm(w)[:-2] if len(w) > 5 else _norm(w) for w in slowa]
    wyn = [x for x in wyn if all(r in _norm(x.get("title") or "") for r in rdz) and x.get("type") != "Obwieszczenie"]
    wyn.sort(key=lambda x: (x.get("status") != "obowiązujący" and "jednolity" not in (x.get("status") or ""), -(x.get("year") or 0)))
    return wyn[:25]


# ----------------------------------------------------------------------------- GUGiK: lokalizacja i dzialki
def _get(url, timeout=40):
    import urllib.request
    return urllib.request.urlopen(url, timeout=timeout).read()


def _do2180(x, y, uklad):
    if str(uklad) == "2180":
        return float(x), float(y)
    from pyproj import Transformer
    t = Transformer.from_crs(int(uklad), 2180, always_xy=True)
    return t.transform(float(x), float(y))


def _do4326(x, y):
    from pyproj import Transformer
    return Transformer.from_crs(2180, 4326, always_xy=True).transform(x, y)


def uldk_xy(x, y):
    t = _get(f"https://uldk.gugik.gov.pl/?request=GetParcelByXY&xy={x:.2f},{y:.2f},2180"
             "&result=id,teryt,voivodeship,county,commune,region,parcel").decode("utf-8").strip().splitlines()
    if not t or t[0].strip() != "0" or len(t) < 2:
        return None
    p = t[1].split("|")
    return {"id": p[0], "wojewodztwo": p[2], "powiat": p[3], "gmina": p[4], "obreb": p[5], "dzialka": p[6],
            "jednostka_ewid": p[0].split(".")[0]}


def uldk_id(ident):
    import urllib.parse
    t = _get("https://uldk.gugik.gov.pl/?request=GetParcelById&id=" + urllib.parse.quote(ident) +
             "&result=id,voivodeship,county,commune,region,parcel,geom_extent").decode("utf-8").strip().splitlines()
    if not t or t[0].strip() != "0" or len(t) < 2:
        return None
    p = t[1].split("|")
    ext = [float(v) for v in p[6].split(",")] if len(p) > 6 and p[6] else None
    return {"id": p[0], "wojewodztwo": p[1], "powiat": p[2], "gmina": p[3], "obreb": p[4], "dzialka": p[5],
            "jednostka_ewid": p[0].split(".")[0], "extent": ext}


def uug_miejscowosc(nazwa, gmina=None):
    import urllib.parse
    d = json.loads(_get("https://services.gugik.gov.pl/uug/?request=GetAddress&address=" + urllib.parse.quote(nazwa)))
    wyn = list((d.get("results") or {}).values())
    if gmina:
        wyn = [w for w in wyn if gmina.lower() in (w.get("commune") or "").lower()] or wyn
    return wyn


def wms(url, warstwy, bbox, w, h, fmt="image/png", inne="", proby=3):
    import io
    from PIL import Image
    u = (f"{url}{'&' if '?' in url else '?'}SERVICE=WMS&VERSION=1.1.1&REQUEST=GetMap&LAYERS={warstwy}&STYLES="
         f"&SRS=EPSG:2180&BBOX={bbox[0]:.1f},{bbox[1]:.1f},{bbox[2]:.1f},{bbox[3]:.1f}&WIDTH={w}&HEIGHT={h}&FORMAT={fmt}{inne}")
    for i in range(proby):
        try:
            return Image.open(io.BytesIO(_get(u, 90))).convert("RGBA")
        except Exception:
            if i == proby - 1:
                raise
            time.sleep(3 * (i + 1))


TOPO = "https://mapy.geoportal.gov.pl/wss/service/img/guest/TOPO/MapServer/WMSServer"
ORTO = "https://mapy.geoportal.gov.pl/wss/service/PZGIK/ORTO/WMS/StandardResolution"
PRG = "https://mapy.geoportal.gov.pl/wss/service/PZGIK/PRG/WMS/AdministrativeBoundaries"
KIEG = "https://integracja.gugik.gov.pl/cgi-bin/KrajowaIntegracjaEwidencjiGruntow"


def _znacznik(img, px, py, r=14, kolor=(220, 20, 60, 255)):
    from PIL import ImageDraw
    dr = ImageDraw.Draw(img)
    dr.ellipse([px - r, py - r, px + r, py + r], outline=kolor, width=4)
    dr.line([px - 2 * r, py, px - r, py], fill=kolor, width=3)
    dr.line([px + r, py, px + 2 * r, py], fill=kolor, width=3)
    dr.line([px, py - 2 * r, px, py - r], fill=kolor, width=3)
    dr.line([px, py + r, px, py + 2 * r], fill=kolor, width=3)


def _font(rozm):
    from PIL import ImageFont
    for f in ("DejaVuSans-Bold.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"):
        try:
            return ImageFont.truetype(f, rozm)
        except Exception:
            pass
    return ImageFont.load_default()


def _podziałka_i_polnoc(img, m_na_px):
    from PIL import ImageDraw
    dr = ImageDraw.Draw(img)
    W, H = img.size
    cel = W * m_na_px / 5
    krok = [50, 100, 200, 250, 500, 1000, 2000, 5000, 10000, 20000, 50000, 100000]
    L = max(k for k in krok if k <= cel) if cel >= 50 else 50
    lpx = L / m_na_px
    x0, y0 = 30, H - 40
    dr.rectangle([x0 - 8, y0 - 28, x0 + lpx + 60, y0 + 14], fill=(255, 255, 255, 220))
    dr.rectangle([x0, y0, x0 + lpx / 2, y0 + 8], fill=(0, 0, 0, 255))
    dr.rectangle([x0 + lpx / 2, y0, x0 + lpx, y0 + 8], outline=(0, 0, 0, 255), fill=(255, 255, 255, 255))
    txt = f"{L / 1000:g} km" if L >= 1000 else f"{L:g} m"
    dr.text((x0, y0 - 24), f"0          {txt}", fill=(0, 0, 0, 255), font=_font(16))
    # strzalka polnocy
    xs, ys = W - 50, 30
    dr.polygon([(xs, ys), (xs - 14, ys + 40), (xs, ys + 30), (xs + 14, ys + 40)], fill=(0, 0, 0, 255))
    dr.text((xs - 7, ys + 42), "N", fill=(0, 0, 0, 255), font=_font(20))


def mapa_polski(x, y, out):
    bbox = (140000, 120000, 880000, 800000)
    W, H = 900, int(900 * (bbox[3] - bbox[1]) / (bbox[2] - bbox[0]))
    from PIL import Image
    tlo = Image.new("RGBA", (W, H), (255, 255, 255, 255))
    try:
        g = wms(PRG, "A01_Granice_wojewodztw,A00_Granice_panstwa", bbox, W, H, inne="&TRANSPARENT=TRUE")
        tlo.alpha_composite(g)
    except Exception as e:
        print(f"PRG WMS niedostępny ({e}) - mapa Polski bez granic", file=sys.stderr)
    px = (x - bbox[0]) / (bbox[2] - bbox[0]) * W
    py = (bbox[3] - y) / (bbox[3] - bbox[1]) * H
    _znacznik(tlo, px, py, r=12)
    tlo.convert("RGB").save(out, quality=92)


def mapa_topo(x, y, out, promien=1500, orto=False):
    W, H = 1500, 1000
    bbox = (x - promien * 1.5, y - promien, x + promien * 1.5, y + promien)
    img = wms(ORTO if orto else TOPO, "Raster" if not orto else "Raster", bbox, W, H,
              fmt="image/jpeg")
    if orto:
        try:
            dz = wms(KIEG, "dzialki,numery_dzialek", bbox, W, H, inne="&TRANSPARENT=TRUE")
            # czesc serwerow powiatowych zwraca biale, polprzezroczyste tlo - usuwamy je
            px = dz.load()
            for j in range(dz.height):
                for i in range(dz.width):
                    r, g, b, al = px[i, j]
                    if r > 225 and g > 225 and b > 225:
                        px[i, j] = (0, 0, 0, 0)
            img.alpha_composite(dz)
        except Exception as e:
            print(f"KIEG (działki) niedostępny: {e}", file=sys.stderr)
    _znacznik(img, W / 2, H / 2, r=18)
    _podziałka_i_polnoc(img, (bbox[2] - bbox[0]) / W)
    img.convert("RGB").save(out, quality=90)


def _miejscownik(przym):
    """pomorskie -> pomorskim, gdański -> gdańskim (przymiotniki nazw jednostek); inne bez zmian"""
    def f(w):
        if not w[:1].islower():
            return w
        if w.endswith("ie"):
            return w[:-1] + "m"
        if w.endswith(("i", "y")):
            return w + "m"
        return w
    return " ".join("-".join(f(c) for c in w.split("-")) for w in przym.split(" "))


def _powiat_miejscownik(p):
    nazwa = re.sub(r"^powiat\s+", "", p.strip())
    if nazwa[:1].islower():
        return f"powiecie {_miejscownik(nazwa)}"
    return f"mieście na prawach powiatu {nazwa}"


def cmd_lokalizacja(a):
    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    if a.miejscowosc:
        wyn = uug_miejscowosc(a.miejscowosc, a.gmina)
        if not wyn:
            sys.exit("Nie znaleziono miejscowości w PRNG/UUG.")
        if len(wyn) > 1 and not a.gmina:
            print("Kilka miejscowości o tej nazwie - użyto pierwszej; zawęź --gmina:", file=sys.stderr)
            for w in wyn:
                print(f"  {w['city']} - gm. {w['commune']}, pow. {w['county']}, woj. {w['voivodeship']}", file=sys.stderr)
        x, y = float(wyn[0]["x"]), float(wyn[0]["y"])
    elif a.dzialka:
        d = uldk_id(a.dzialka)
        if not d:
            sys.exit("ULDK nie zna takiej działki.")
        e = d["extent"]
        x, y = (e[0] + e[2]) / 2, (e[1] + e[3]) / 2
    else:
        xs, ys = a.xy.split(",")
        x, y = _do2180(xs, ys, a.uklad)
    adm = uldk_xy(x, y) or {}
    lon, lat = _do4326(x, y)
    info = {"x_2180": round(x, 2), "y_2180": round(y, 2), "lon": round(lon, 6), "lat": round(lat, 6), **adm,
            "geoportal": f"https://mapy.geoportal.gov.pl/imap/Imgp_2.html?locale=pl&gui=new&sessionID=&bbox={x - 800:.0f},{y - 500:.0f},{x + 800:.0f},{y + 500:.0f}",
            "zrodlo": "GUGiK: ULDK, UUG, WMS PRG/TOPO/ORTO/KIEG; dane na " + date.today().isoformat()}
    if adm:
        info["opis"] = (f"Obiekt położony jest w województwie {_miejscownik(adm['wojewodztwo'])}, w "
                        f"{_powiat_miejscownik(adm['powiat'])}, w gminie {adm['gmina']}, w obrębie ewidencyjnym "
                        f"{adm['obreb']} (jednostka ewidencyjna {adm['jednostka_ewid']}), na działce nr {adm['dzialka']}.")
    (out / "lokalizacja.json").write_text(json.dumps(info, ensure_ascii=False, indent=1), encoding="utf-8")
    mapa_polski(x, y, out / "rys_polska.png")
    mapa_topo(x, y, out / "rys_topo.png", a.promien)
    try:
        mapa_topo(x, y, out / "rys_orto.png", min(a.promien, 400), orto=True)
    except Exception as e:
        print(f"Ortofotomapa niedostępna: {e}", file=sys.stderr)
    print(json.dumps(info, ensure_ascii=False, indent=1))
    print(f"\nMapy: {out}/rys_polska.png, rys_topo.png, rys_orto.png - sprawdź, czy znacznik wskazuje obiekt (działka z ULDK to działka w punkcie).")


def cmd_dzialki(a):
    wiersze = []
    if a.id:
        for i in a.id.split(","):
            d = uldk_id(i.strip())
            wiersze.append(d or {"id": i, "blad": "brak w ULDK"})
    else:
        for p in a.xy.split(";"):
            xs, ys = p.split(",")
            x, y = _do2180(xs, ys, a.uklad)
            wiersze.append(uldk_xy(x, y) or {"id": f"{xs},{ys}", "blad": "brak działki w punkcie"})
    print("| Lp. | Województwo | Powiat | Gmina | Obręb | Nr działki | Identyfikator |\n|---|---|---|---|---|---|---|")
    for i, w in enumerate(wiersze, 1):
        if w.get("blad"):
            print(f"| {i} | | | | | | {w['id']} – {w['blad']} |")
        else:
            print(f"| {i} | {w['wojewodztwo']} | {w['powiat']} | {w['gmina']} | {w['obreb']} | {w['dzialka']} | {w['id']} |")


# ----------------------------------------------------------------------------- przepisy
def cmd_przepisy(a):
    if a.szukaj:
        for it in eli_szukaj(a.szukaj):
            print(f"- {it.get('ELI')} | {it.get('status')} | {it.get('title')}")
        print("\nCytat aktu: opis_tool.py cytuj <ELI>  (dla aktu z tekstem jednolitym podawany jest aktualny t.j.)")
        return
    klucze = [a.klucz] if a.klucz else list(KATALOG_AKTOW)
    print("| klucz | kat. | cytat (stan na dziś) | status |\n|---|---|---|---|")
    for k in klucze:
        kat, eli = KATALOG_AKTOW[k]
        w = eli_rozwiaz(eli, a.offline)
        print(f"| {k} | [{kat}] | {w.get('cytat')} | {w.get('status') or w.get('blad')} |")


def cmd_cytuj(a):
    eli = KATALOG_AKTOW[a.akt][1] if a.akt in KATALOG_AKTOW else a.akt
    w = eli_rozwiaz(eli, a.offline)
    print(w.get("cytat"))
    if w.get("uchylony") or (w.get("status") and "uchyl" in w["status"]):
        print(f"UWAGA: akt {w['status']} ({w.get('uchylony')}); uchylenie wynika z: {', '.join(w.get('uchylony_przez') or ['?'])}")


# ----------------------------------------------------------------------------- materialy i odwolania
def zbuduj_materialy(spec, offline=False):
    """Zwraca (lista sekcji [(kat, nazwa, [(znacznik, tekst)])], mapa klucz->znacznik)."""
    mat = spec.get("materialy") or {}
    sekcje, mapa = [], {}
    for kat, nazwa in KATEGORIE:
        pozycje = []
        for i, e in enumerate(mat.get(kat) or [], 1):
            znacz = f"[{kat}{i}]"
            if isinstance(e, dict):
                klucz, tekst = e.get("klucz"), e.get("tekst") or ""
            else:
                klucz, tekst = e, e
            if kat in ("U", "R") and (klucz in KATALOG_AKTOW or re.fullmatch(r"DU/\d{4}/\d+", str(klucz))):
                eli = KATALOG_AKTOW[klucz][1] if klucz in KATALOG_AKTOW else klucz
                tekst = eli_rozwiaz(eli, offline).get("cytat")
            elif kat == "N" and klucz in KATALOG_NORM:
                tekst = KATALOG_NORM[klucz]
            pozycje.append((znacz, tekst))
            if klucz:
                mapa[str(klucz)] = znacz
        if pozycje:
            sekcje.append((kat, nazwa, pozycje))
    return sekcje, mapa


def podstaw_odwolania(tekst, mapa, braki):
    def f(m):
        k = m.group(1)
        if k in mapa:
            return mapa[k]
        braki.add(k)
        return f"[?{k}]"
    return re.sub(r"\[@([^\]]+)\]", f, tekst or "")


def data_pl(s):
    try:
        d = date.fromisoformat(s)
        return f"{d.day} {MIESIACE[d.month - 1]} {d.year} r."
    except Exception:
        return s or "[DATA]"


def tekst_podstawy_formalnej(spec):
    u = spec.get("umowa") or {}
    if spec.get("podstawa_formalna"):
        return spec["podstawa_formalna"]
    if not u:
        return "[DO UZUPEŁNIENIA: umowa / zlecenie – numer, data, zamawiający]"
    t = (f"Umowa nr {u.get('nr', '[NR]')} z dnia {data_pl(u.get('data'))} zawarta pomiędzy "
         f"{u.get('zamawiajacy', '[ZAMAWIAJĄCY]')} a Pracownią Projektową MiD Sp. z o.o."
         + (f" ({u['uwagi']})" if u.get("uwagi") else ""))
    return t if t.endswith(".") else t + "."


# ----------------------------------------------------------------------------- generuj DOCX / MD
WZORY = {
    "PB": ("PROJEKT BUDOWLANY", "prawo", "Celem opracowania jest uzyskanie decyzji o pozwoleniu na budowę [lub decyzji o zezwoleniu na realizację inwestycji drogowej] dla zamierzenia budowlanego opisanego w pkt 1."),
    "PT": ("PROJEKT TECHNICZNY", "prawo", "Celem opracowania jest uszczegółowienie rozwiązań projektu budowlanego w zakresie niezbędnym do wykonania robót budowlanych (art. 34 ust. 3 pkt 4 [@pb])."),
    "PW": ("PROJEKT WYKONAWCZY", "inne", "Celem opracowania jest uzupełnienie i uszczegółowienie projektu budowlanego i technicznego w zakresie i stopniu dokładności niezbędnym do sporządzenia przedmiaru robót, kosztorysu inwestorskiego, przygotowania oferty i realizacji robót budowlanych."),
    "PZT": ("PROJEKT ZAGOSPODAROWANIA TERENU", "prawo", "Celem opracowania jest określenie usytuowania obiektów budowlanych na działkach objętych inwestycją, jako części projektu budowlanego."),
    "koncepcja": ("KONCEPCJA PROJEKTOWA", "inne", "Celem opracowania jest przedstawienie wariantów rozwiązania i wskazanie wariantu rekomendowanego do dalszych prac projektowych."),
    "ekspertyza": ("EKSPERTYZA TECHNICZNA", "eksperckie", "Celem opracowania jest ocena stanu technicznego obiektu, określenie przyczyn uszkodzeń, aktualnej nośności oraz sposobu naprawy."),
    "STWiORB": ("SPECYFIKACJE TECHNICZNE WYKONANIA I ODBIORU ROBÓT BUDOWLANYCH", "prawo", "Celem opracowania jest określenie wymagań dotyczących wykonania i odbioru robót budowlanych [@dok_projektowa_stwiorb_pfu]."),
    "PFU": ("PROGRAM FUNKCJONALNO-UŻYTKOWY", "prawo", "Celem opracowania jest określenie wymagań zamawiającego dla zamówienia w formule „zaprojektuj i wybuduj” [@dok_projektowa_stwiorb_pfu]."),
    "inne": ("OPRACOWANIE", "inne", "[DO UZUPEŁNIENIA: cel opracowania]"),
}
KAT_CELU = {"prawo": "opracowanie wynikające z przepisów prawa", "eksperckie": "opracowanie eksperckie", "inne": "opracowanie inne"}


def cmd_wzor(a):
    t = WZORY.get(a.typ, WZORY["inne"])
    spec = {
        "typ": a.typ, "tytul_opracowania": t[0], "branza": "mostowa", "tom": "", "faza": t[0].title(),
        "zamierzenie": "[nazwa zamierzenia budowlanego / zadania z umowy]",
        "obiekt": "[obiekt – rodzaj, droga, km, przeszkoda]",
        "inwestor": {"nazwa": "[Inwestor]", "adres": "[adres]"},
        "umowa": {"nr": "[nr]", "data": "RRRR-MM-DD", "zamawiajacy": "[Zamawiający]"},
        "lokalizacja_json": "lokalizacja/lokalizacja.json",
        "lokalizacja_opis_dodatkowy": "[km drogi, przeszkoda, dojazd]",
        "dzialki": [{"jednostka_ewid": "", "obreb": "", "dzialka": ""}],
        "zespol": [{"funkcja": "Projektant", "osoba": "dr inż. Marcin Dudek", "uprawnienia": "[nr, specjalność]"},
                   {"funkcja": "Sprawdzający", "osoba": "[ ]", "uprawnienia": "[ ]"}],
        "data": date.today().strftime("%m.%Y"), "rewizja": "00",
        "przedmiot": "Przedmiotem opracowania jest [konkretnie: obiekt / roboty / inwestycja].",
        "cel": {"kategoria": t[1], "tekst": t[2]},
        "zakres": ["[element zakresu 1]", "[element zakresu 2]"],
        "podstawa_formalna": "",
        "podstawa_merytoryczna": ["[@DA1] – [materiał wyjściowy]"],
        "rozdzialy": [{"tytul": "Opis stanu istniejącego", "tresc": "[treść; odwołania do materiałów: [@pb], [@PN-EN 1991-2]]"}],
        "materialy": {"DA": [{"klucz": "DA1", "tekst": "[dokumentacja archiwalna – autor, tytuł, rok]"}],
                      "N": ["PN-EN 1990", "PN-EN 1991-2"], "U": ["pb"], "R": ["ptb_drogi_2022"], "W": [], "L": [], "I": [], "P": []},
    }
    Path(a.out).write_text(json.dumps(spec, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"Zapisano {a.out}. Lokalizację wypełnij komendą 'lokalizacja', odwołania w tekście pisz jako [@klucz].")


def _md_akapity(tekst):
    return [p.strip() for p in re.split(r"\n\s*\n", tekst or "") if p.strip()]


def cmd_generuj(a):
    import docx
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt, RGBColor
    spec = json.loads(Path(a.spec).read_text(encoding="utf-8"))
    baza = Path(a.spec).parent
    sekcje, mapa = zbuduj_materialy(spec, a.offline)
    braki = set()
    lok = {}
    if spec.get("lokalizacja_json") and (baza / spec["lokalizacja_json"]).exists():
        lok = json.loads((baza / spec["lokalizacja_json"]).read_text(encoding="utf-8"))
    rys_dir = (baza / spec["lokalizacja_json"]).parent if spec.get("lokalizacja_json") else baza
    NIEB = RGBColor(0x1F, 0x4E, 0x79)
    d = docx.Document()
    st = d.styles["Normal"]
    st.font.name = "Arial"
    st.font.size = Pt(10.5)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    for s in d.sections:
        s.left_margin = s.right_margin = Cm(2.0)
        s.top_margin = s.bottom_margin = Cm(2.0)

    def akapit(t, b=False, rozm=None, wyr=None, kolor=None, odstep=6):
        p = d.add_paragraph()
        r = p.add_run(t)
        r.bold = b
        if rozm:
            r.font.size = Pt(rozm)
        if kolor:
            r.font.color.rgb = kolor
        if wyr == "c":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif wyr == "j":
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_after = Pt(odstep)
        return p

    def naglowek(t, poziom=1):
        h = d.add_heading(t, level=poziom)
        for r in h.runs:
            r.font.name = "Arial"
            r.font.color.rgb = NIEB
            r.font.size = Pt(13 if poziom == 1 else 11.5)
        return h

    def tabela(dane, szer=None, naglowkowa=True):
        t = d.add_table(rows=len(dane), cols=len(dane[0]))
        t.style = "Table Grid"
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        for i, w in enumerate(dane):
            for j, v in enumerate(w):
                c = t.cell(i, j)
                c.text = ""
                r = c.paragraphs[0].add_run(str(v))
                r.font.size = Pt(9.5)
                if naglowkowa and i == 0:
                    r.bold = True
                    tc = c._tc.get_or_add_tcPr()
                    sh = OxmlElement("w:shd")
                    sh.set(qn("w:val"), "clear")
                    sh.set(qn("w:fill"), "D9E1F2")
                    tc.append(sh)
        if szer:
            t.autofit = False
            for j, s in enumerate(szer):
                t.columns[j].width = Cm(s)
                for row in t.rows:
                    row.cells[j].width = Cm(s)
        d.add_paragraph()
        return t

    # strona tytulowa
    firma = spec.get("firma") or {}
    akapit((firma.get("nazwa") or "Pracownia Projektowa MiD Sp. z o.o.").upper().replace("MID SP. Z O.O.", "MiD Sp. z o.o."),
           b=True, rozm=12, wyr="c", kolor=NIEB, odstep=0 if firma.get("adres") else 24)
    if firma.get("adres"):
        akapit(firma["adres"], rozm=9, wyr="c", odstep=24)
    akapit(spec.get("tytul_opracowania", ""), b=True, rozm=18, wyr="c", kolor=NIEB, odstep=4)
    if spec.get("branza"):
        akapit(f"BRANŻA {spec['branza'].upper()}", b=True, rozm=12, wyr="c", odstep=4)
    if spec.get("tom"):
        akapit(spec["tom"], b=True, rozm=11, wyr="c", odstep=14)
    inw = spec.get("inwestor") or {}
    dz_txt = ", ".join(f"{x.get('dzialka')}" for x in spec.get("dzialki") or [] if x.get("dzialka"))
    metryka = [["Zamierzenie budowlane:", spec.get("zamierzenie", "")], ["Obiekt:", spec.get("obiekt", "")],
               ["Adres / lokalizacja:", (lok.get("opis") or spec.get("lokalizacja_opis_dodatkowy") or "")],
               ["Działki ewidencyjne:", (dz_txt + (f" (obręb {spec['dzialki'][0].get('obreb')}, jedn. ewid. {spec['dzialki'][0].get('jednostka_ewid')})" if dz_txt and spec['dzialki'][0].get('obreb') else "")) or "—"],
               ["Inwestor:", f"{inw.get('nazwa', '')}, {inw.get('adres', '')}".strip(", ")],
               ["Nr umowy:", (spec.get("umowa") or {}).get("nr", "—")], ["Data / rewizja:", f"{spec.get('data', '')} / REW{spec.get('rewizja', '00')}"]]
    tabela(metryka, szer=[4.5, 12.5], naglowkowa=False)
    zesp = [["Funkcja", "Imię i nazwisko", "Specjalność i nr uprawnień", "Data", "Podpis"]]
    for z in spec.get("zespol") or []:
        zesp.append([z.get("funkcja", ""), z.get("osoba", ""), z.get("uprawnienia", ""), spec.get("data", ""), ""])
    tabela(zesp, szer=[3, 4.5, 5.5, 2, 2])
    d.add_page_break()

    # 1-4 wg Instrukcji 02
    naglowek("1. Przedmiot opracowania")
    akapit(podstaw_odwolania(spec.get("przedmiot"), mapa, braki), wyr="j")
    naglowek("2. Lokalizacja")
    opis_lok = " ".join(x for x in [lok.get("opis", ""), spec.get("lokalizacja_opis_dodatkowy", "")] if x and not x.startswith("["))
    akapit(podstaw_odwolania(opis_lok or "[DO UZUPEŁNIENIA: województwo, powiat, gmina, obręb, km drogi]", mapa, braki), wyr="j")
    if lok.get("x_2180"):
        akapit(f"Współrzędne obiektu (PL-1992, EPSG:2180): X = {lok['x_2180']:.0f}, Y = {lok['y_2180']:.0f}; "
               f"WGS-84: {lok['lat']:.5f}° N, {lok['lon']:.5f}° E.", rozm=9.5)
    nr_rys = 0
    for plik, podpis in (("rys_polska.png", "Lokalizacja obiektu na mapie Polski (granice województw: PRG, GUGiK)"),
                         ("rys_topo.png", "Lokalizacja obiektu na mapie topograficznej (źródło: Geoportal, GUGiK)")):
        if (rys_dir / plik).exists():
            nr_rys += 1
            d.add_picture(str(rys_dir / plik), width=Cm(12 if "polska" in plik else 16))
            d.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
            akapit(f"Rys. {nr_rys}. {podpis}", rozm=9, wyr="c")
    if spec.get("dzialki") and any(x.get("dzialka") for x in spec["dzialki"]):
        akapit("Wykaz działek objętych opracowaniem:", b=True)
        wiersze = [["Lp.", "Jednostka ewidencyjna", "Obręb", "Nr działki"]]
        for i, x in enumerate(spec["dzialki"], 1):
            wiersze.append([i, x.get("jednostka_ewid", ""), x.get("obreb", ""), x.get("dzialka", "")])
        tabela(wiersze, szer=[1.2, 5, 5, 4])
    naglowek("3. Cel i zakres opracowania")
    cel = spec.get("cel") or {}
    akapit(f"Rodzaj opracowania: {KAT_CELU.get(cel.get('kategoria'), cel.get('kategoria', '—'))}.", rozm=9.5)
    akapit(podstaw_odwolania(cel.get("tekst"), mapa, braki), wyr="j")
    if spec.get("zakres"):
        akapit("Zakres opracowania obejmuje:")
        for z in spec["zakres"]:
            d.add_paragraph(podstaw_odwolania(z, mapa, braki), style="List Bullet")
    naglowek("4. Podstawa opracowania")
    naglowek("4.1. Podstawa formalna", 2)
    akapit(tekst_podstawy_formalnej(spec), wyr="j")
    naglowek("4.2. Podstawa merytoryczna", 2)
    for m in spec.get("podstawa_merytoryczna") or ["[DO UZUPEŁNIENIA: materiały wyjściowe z odwołaniem do listy wykorzystanych materiałów]"]:
        d.add_paragraph(podstaw_odwolania(m, mapa, braki), style="List Bullet")
    # rozdzialy merytoryczne (tresc, tabele, rysunki, podrozdzialy)
    licz_tab = [0]

    def blok(r, nr, poziom):
        naglowek(f"{nr}. {r.get('tytul', '')}", poziom)
        for p in _md_akapity(r.get("tresc")):
            if p.startswith("- "):
                for li in p.split("\n"):
                    d.add_paragraph(podstaw_odwolania(li.lstrip("- ").strip(), mapa, braki), style="List Bullet")
            else:
                akapit(podstaw_odwolania(p, mapa, braki), wyr="j")
        for t in r.get("tabele") or []:
            licz_tab[0] += 1
            if t.get("podpis"):
                akapit(f"Tabela {licz_tab[0]}. {podstaw_odwolania(t['podpis'], mapa, braki)}", rozm=9, b=True, odstep=2)
            tabela([t["naglowek"]] + [[podstaw_odwolania(str(v), mapa, braki) for v in w] for w in t["wiersze"]], szer=t.get("szerokosci"))
            if t.get("uwaga"):
                akapit(podstaw_odwolania(t["uwaga"], mapa, braki), rozm=8.5)
        for rys in r.get("rysunki") or []:
            pr = (baza / rys["plik"])
            if pr.exists():
                nonlocal_rys[0] += 1
                d.add_picture(str(pr), width=Cm(rys.get("szerokosc_cm", 16)))
                d.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
                akapit(f"Rys. {nonlocal_rys[0]}. {podstaw_odwolania(rys.get('podpis', ''), mapa, braki)}", rozm=9, wyr="c")
            else:
                print(f"UWAGA: brak pliku rysunku {pr}", file=sys.stderr)
        for k, pr in enumerate(r.get("podrozdzialy") or [], 1):
            blok(pr, f"{nr}.{k}", min(poziom + 1, 3))

    nonlocal_rys = [nr_rys]
    for i, r in enumerate(spec.get("rozdzialy") or [], 5):
        blok(r, str(i), 1)
    nr_rys = nonlocal_rys[0]
    # wykorzystane materialy (na koncu)
    n = 5 + len(spec.get("rozdzialy") or [])
    naglowek(f"{n}. Wykorzystane materiały")
    for kat, nazwa, poz in sekcje:
        akapit(f"{nazwa} [{kat}]", b=True, odstep=2)
        for znacz, tekst in poz:
            p = d.add_paragraph()
            r = p.add_run(f"{znacz} ")
            r.bold = True
            p.add_run(tekst)
            p.paragraph_format.left_indent = Cm(1.2)
            p.paragraph_format.first_line_indent = Cm(-1.2)
            p.paragraph_format.space_after = Pt(2)
    # stopka
    for s in d.sections:
        f = s.footer.paragraphs[0]
        f.text = f"{(spec.get('firma') or {}).get('nazwa') or 'Pracownia Projektowa MiD Sp. z o.o.'} | {spec.get('tytul_opracowania', '')} | REW{spec.get('rewizja', '00')}"
        f.runs[0].font.size = Pt(8)
    d.save(a.out)
    print(f"Zapisano {a.out}: {sum(len(p) for _, _, p in sekcje)} pozycji materiałów, rysunków: {nr_rys}.")
    if braki:
        print(f"UWAGA: nierozwiązane odwołania [@...]: {', '.join(sorted(braki))} - dodaj je do 'materialy' (pole klucz).")
    if a.md:
        L = [f"# {spec.get('tytul_opracowania', '')}", "", "## 1. Przedmiot opracowania", podstaw_odwolania(spec.get("przedmiot"), mapa, set()),
             "", "## 2. Lokalizacja", opis_lok, "", "## 3. Cel i zakres opracowania", podstaw_odwolania(cel.get("tekst"), mapa, set())]
        L += [f"- {podstaw_odwolania(z, mapa, set())}" for z in spec.get("zakres") or []]
        L += ["", "## 4. Podstawa opracowania", "### 4.1. Podstawa formalna", tekst_podstawy_formalnej(spec), "### 4.2. Podstawa merytoryczna"]
        L += [f"- {podstaw_odwolania(m, mapa, set())}" for m in spec.get("podstawa_merytoryczna") or []]
        def blok_md(r, nr, poziom):
            out = ["", f"{'#' * (poziom + 1)} {nr}. {r.get('tytul')}", podstaw_odwolania(r.get("tresc") or "", mapa, set())]
            for t in r.get("tabele") or []:
                out += ["", f"*{podstaw_odwolania(t.get('podpis', ''), mapa, set())}*", "", "| " + " | ".join(map(str, t["naglowek"])) + " |", "|" + "---|" * len(t["naglowek"])]
                out += ["| " + " | ".join(podstaw_odwolania(str(v), mapa, set()).replace("\n", " ") for v in w) + " |" for w in t["wiersze"]]
                if t.get("uwaga"):
                    out += ["", podstaw_odwolania(t["uwaga"], mapa, set())]
            for rys in r.get("rysunki") or []:
                out += ["", f"![{rys.get('podpis', '')}]({rys['plik']})"]
            for k, pr in enumerate(r.get("podrozdzialy") or [], 1):
                out += blok_md(pr, f"{nr}.{k}", min(poziom + 1, 3))
            return out
        for i, r in enumerate(spec.get("rozdzialy") or [], 5):
            L += blok_md(r, str(i), 1)
        L += ["", f"## {n}. Wykorzystane materiały"]
        for kat, nazwa, poz in sekcje:
            L += [f"**{nazwa} [{kat}]**", ""] + [f"{z} {t}  " for z, t in poz] + [""]
        Path(a.md).write_text("\n".join(L), encoding="utf-8")


# ----------------------------------------------------------------------------- sprawdz
def tekst_pliku(p):
    p = Path(p)
    if p.suffix.lower() == ".docx":
        import docx
        dd = docx.Document(str(p))
        linie = []
        for blk in dd.element.body.iterchildren():
            tag = blk.tag.split("}")[1]
            if tag == "p":
                t = "".join(x.text or "" for x in blk.iter() if x.tag.endswith("}t"))
                linie.append(t)
            elif tag == "tbl":
                for tr in blk.iter():
                    if tr.tag.endswith("}tr"):
                        linie.append(" | ".join("".join(x.text or "" for x in tc.iter() if x.tag.endswith("}t"))
                                                for tc in tr if tc.tag.endswith("}tc")))
        return "\n".join(linie)
    if p.suffix.lower() == ".pdf":
        import subprocess
        return subprocess.run(["pdftotext", "-layout", str(p), "-"], capture_output=True, text=True).stdout
    return p.read_text(encoding="utf-8", errors="replace")


def _norm(s):
    s = unicodedata.normalize("NFKD", s.lower())
    return "".join(c for c in s if not unicodedata.combining(c)).replace("ł", "l")


SEKCJE_WZ = [
    ("przedmiot", r"przedmiot\w*\s+(opracowania|zamowienia|ekspertyzy|projektu)|^\s*\d*\.?\s*przedmiot\b"),
    ("lokalizacja", r"lokalizacj|usytuowani\w* obiektu|polozenie obiektu"),
    ("cel", r"\bcel\w*\s*(i\s+zakres)?\s*(opracowania|ekspertyzy|projektu|zamowienia)|^\s*\d*\.?\s*cel\b"),
    ("zakres", r"zakres\w*\s+(opracowania|ekspertyzy|projektu|zamowienia)|cel i zakres"),
    ("podstawa", r"podstaw\w*\s+(opracowania|formaln|merytoryczn)"),
    ("materialy", r"wykorzystan\w*\s+material|material\w*\s+wykorzystan|spis\s+(norm|literatury)|bibliografi|\bliteratura\b|normy,? przepisy|przepisy zwiazane|materialy zrodlowe"),
]


def cmd_sprawdz(a):
    tekst = tekst_pliku(a.plik)
    linie = tekst.splitlines()
    spis = [bool(re.search(r"(\.{4,}|…{2,}|\. \. \. )\s*\d+\s*$", l)) for l in linie]
    N = max(len(linie), 1)
    norm_linie = [_norm(l) for l in linie]
    wyniki = []

    def wynik(stan, co, szcz=""):
        wyniki.append((stan, co, szcz))

    # 1. sekcje Instrukcji 02 (naglowki: krotkie linie)
    poz = {}
    for k, rx in SEKCJE_WZ:
        for i, l in enumerate(norm_linie):
            if not spis[i] and len(l.strip()) <= 90 and re.search(rx, l.strip()):
                poz.setdefault(k, i)
                if k != "materialy":
                    break
        if k == "materialy":
            ost = [i for i, l in enumerate(norm_linie) if not spis[i] and len(l.strip()) <= 90 and re.search(rx, l.strip())]
            if ost:
                poz[k] = ost[-1]
    nazwy = {"przedmiot": "Przedmiot opracowania", "lokalizacja": "Lokalizacja", "cel": "Cel opracowania",
             "zakres": "Zakres opracowania", "podstawa": "Podstawa opracowania", "materialy": "Wykorzystane materiały"}
    for k in ("przedmiot", "lokalizacja", "cel", "zakres", "podstawa", "materialy"):
        if k in poz:
            wynik("✓", f"Punkt „{nazwy[k]}”", f"wiersz {poz[k] + 1} ({100 * poz[k] / N:.0f}% dokumentu)")
        else:
            wynik("✗", f"Punkt „{nazwy[k]}”", "nie znaleziono nagłówka")
    pierwsze = [poz[k] for k in ("przedmiot", "lokalizacja", "cel", "podstawa") if k in poz]
    if pierwsze:
        # pierwszy rozdzial merytoryczny: numerowany naglowek, ktory nie jest zadnym z 4 punktow ani metryka
        rx4 = "|".join(rx for k, rx in SEKCJE_WZ if k != "materialy")
        rozdz = [i for i, l in enumerate(norm_linie)
                 if not spis[i] and i > min(pierwsze) and re.match(r"\s*\d+(\.\d+)*\.?\s+\w", l) and len(l.strip()) <= 90
                 and not re.search(rx4 + r"|zamawiaj|spis|zawartosc|kopie|uprawnien|metryka|oswiadcz|podstawa (formalna|merytoryczna)", l)]
        przed = [i for i in rozdz if i < max(pierwsze)]
        if przed:
            wynik("⚠", "Cztery pierwsze punkty na początku opracowania",
                  f"przed ostatnim z nich jest rozdział „{linie[przed[0]].strip()[:60]}” (wiersz {przed[0] + 1})")
        elif N > 300 and max(pierwsze) / N > 0.35:
            wynik("⚠", "Cztery pierwsze punkty na początku opracowania", f"ostatni z nich w {100 * max(pierwsze) / N:.0f}% dokumentu")
        else:
            wynik("✓", "Cztery pierwsze punkty na początku opracowania")
        kol = [k for k in ("przedmiot", "lokalizacja", "cel", "podstawa") if k in poz]
        if [poz[k] for k in kol] != sorted(poz[k] for k in kol):
            wynik("⚠", "Kolejność punktów", "oczekiwana: przedmiot → lokalizacja → cel i zakres → podstawa")
    if "materialy" in poz:
        wynik("✓" if poz["materialy"] / N > 0.6 else "⚠", "Wykorzystane materiały na końcu opracowania",
              f"{100 * poz['materialy'] / N:.0f}% dokumentu")
    # 2. lokalizacja
    t_n = _norm(tekst)
    brak = [w for w, rx in (("województwo", r"wojewodztw"), ("powiat", r"powiat|powiec"), ("gmina", r"gmin"),
                            ("obręb", r"obreb"), ("działka", r"dzialk|dz\. ?nr|nr dz"), ("mapa / rysunek lokalizacyjny", r"\bmap[aiy]|rys\.|orientacj"))
            if not re.search(rx, t_n)]
    wynik("✓" if not brak else "⚠", "Lokalizacja: opis administracyjny i mapy", ("brak: " + ", ".join(brak)) if brak else "")
    # 3. cel - kategoria
    if re.search(r"w celu uzyskania|uzyskani\w* (decyzji|pozwolenia|zrid|zgody)|zgloszeni", t_n):
        kat = "wynikające z przepisów prawa (decyzja/pozwolenie/zgłoszenie)"
    elif re.search(r"ocen\w* stanu|ekspertyz|nosnosc|diagnostyk", t_n):
        kat = "eksperckie (ocena, ekspertyza)"
    else:
        kat = None
    wynik("✓" if kat else "⚠", "Cel opracowania określony (kategoria wg Instrukcji 02)", kat or "nie rozpoznano celu: decyzja/uzgodnienie albo ocena ekspercka")
    # 4. podstawa formalna i merytoryczna
    formalna = re.search(r"umow\w*(\s+o\s+\w+)?(\s+\w+)?\s+(nr|numer)|zleceni\w*\s+(nr|z dnia)", t_n)
    wynik("✓" if formalna else "✗", "Podstawa formalna (umowa / zlecenie z numerem i datą)",
          "" if formalna else "nie znaleziono „umowa nr …” ani „zlecenie …”")
    meryt = re.search(r"podstaw\w* merytoryczn|material\w* wyjsciow|\[DA\d+\]", t_n, re.I)
    wynik("✓" if meryt else "⚠", "Podstawa merytoryczna (materiały wyjściowe)")
    # 5. odwolania [X#]
    znaczniki = re.findall(r"\[(DA|N|U|R|W|L|I|P)\s?(\d+)\]", tekst)
    if znaczniki:
        zdef = set()
        i0 = poz.get("materialy", int(N * 0.7))
        for l in linie[i0:]:
            m = re.match(r"\s*\[(DA|N|U|R|W|L|I|P)\s?(\d+)\]", l)
            if m:
                zdef.add(m.groups())
        uzyte = set(m for m in (re.findall(r"\[(DA|N|U|R|W|L|I|P)\s?(\d+)\]", "\n".join(linie[:i0]))))
        nieznane = sorted(f"[{k}{n}]" for k, n in uzyte - zdef)
        nieuzyte = sorted(f"[{k}{n}]" for k, n in zdef - uzyte)
        wynik("✓" if not nieznane else "✗", "Odwołania [X#] w tekście mają pozycję w wykazie",
              ("brak w wykazie: " + ", ".join(nieznane)) if nieznane else f"{len(uzyte)} odwołań")
        if nieuzyte:
            wynik("⚠", "Pozycje wykazu bez odwołania w tekście", ", ".join(nieuzyte[:20]))
        kats = sorted(set(k for k, _ in zdef))
        wynik("✓", "Oznaczenia kategorii materiałów", ", ".join(f"[{k}]" for k in kats))
    else:
        wynik("✗", "Oznaczenia materiałów wg Instrukcji 02 ([DA1], [N1], [U1], [R1], [W1], [L1], [I1], [P1])", "nie znaleziono")
    # 6. akty prawne - aktualnosc (ELI)
    cyt = []
    for m in re.finditer(r"Dz\.?\s?U\.?\s*(?:z\s*)?(\d{4})?\s*(?:r\.)?\s*(?:,?\s*(?:Nr|nr)\.?\s*\d+\s*,?)?\s*poz\.?\s*(\d+)", tekst):
        kont = tekst[max(0, m.start() - 260):m.start()]
        rok = m.group(1)
        if not rok:
            # stary zapis "Dz. U. Nr 63, poz. 735" - rok z daty aktu w kontekscie (albo rok nastepny)
            md = list(re.finditer(r"z dnia \d{1,2} \w+ (\d{4})", kont))
            if not md:
                cyt.append((m.group(0), None, kont))
                continue
            r0 = int(md[-1].group(1))
            kand = [f"DU/{r}/{m.group(2)}" for r in (r0, r0 + 1)]
            eli = kand[0]
            for k in kand:
                w0 = eli_get(k, a.offline) or {}
                mt = re.search(r"z dnia (\d{1,2}) (\w+) (\d{4})", w0.get("title") or "")
                mk0 = re.search(r"z dnia (\d{1,2}) (\w+) (\d{4})", kont[md[-1].start():])
                if mt and mk0 and (mt.group(1), mt.group(3)) == (mk0.group(1), mk0.group(3)):
                    eli = k
                    break
            cyt.append((m.group(0), eli, kont))
            continue
        cyt.append((m.group(0), f"DU/{rok}/{m.group(2)}", kont))
    vis = set()
    for oryg, eli, kont in cyt:
        if eli is None:
            wynik("?", f"Akt {oryg}", "brak roku publikacji i daty aktu – nie można zidentyfikować")
            continue
        if eli in vis:
            continue
        vis.add(eli)
        w = eli_rozwiaz(eli, a.offline)
        if w.get("blad"):
            wynik("?", f"Akt {oryg}", w["blad"])
            continue
        # zgodnosc daty aktu z kontekstem cytatu
        mk = list(re.finditer(r"z dnia (\d{1,2}) (\w+) (\d{4})", kont))
        rozjazd = ""
        if mk:
            dz = mk[-1]
            m_t = re.search(r"z dnia (\d{1,2}) (\w+) (\d{4})", w.get("tytul") or "")
            if m_t and (dz.group(1), dz.group(3)) != (m_t.group(1), m_t.group(3)):
                rozjazd = f"cytowany numer Dz.U. dotyczy innego aktu: „{w['tytul'][:110]}”"
        if rozjazd:
            wynik("✗", f"{oryg}", rozjazd)
        elif w.get("uchylony") or (w.get("status") and ("uchyl" in w["status"] or "wygaś" in w["status"])):
            zast = ", ".join(eli_rozwiaz(x, a.offline).get("cytat", x) for x in (w.get("uchylony_przez") or [])[:1])
            wynik("✗", f"{oryg}: {w['tytul'][:90]}", f"{w['status']} ({w.get('uchylony')}); zastąpiony: {zast or '?'}")
        elif w.get("tj") and eli != w["tj"]:
            wynik("⚠", f"{oryg}: {w['tytul'][:90]}", f"nieaktualny adres – cytuj: {w['cytat']}")
        else:
            wynik("✓", f"{oryg}: {w['tytul'][:90]}", w["cytat"])
    if not cyt:
        wynik("⚠", "Akty prawne z adresem publikacyjnym (Dz.U. RRRR poz. N)", "nie znaleziono cytowań Dz.U.")
    # 7. normy wycofane
    for nr, opis in WYCOFANE_NORMY.items():
        if nr.lower() in tekst.lower():
            wynik("⚠", f"Norma {nr}", opis)
    # 8. powtorzenia tresci (informacja tylko raz)
    zdania = [z.strip() for z in re.split(r"(?<=[.;:])\s+", re.sub(r"\s+", " ", tekst)) if len(z.split()) >= 12]
    licz = {}
    for z in zdania:
        k = _norm(z)[:160]
        licz[k] = licz.get(k, 0) + 1
    powt = [(k, n) for k, n in licz.items() if n > 1]
    wynik("✓" if not powt else "⚠", "Informacja podana tylko raz (brak powtórzeń zdań)",
          (f"{len(powt)} powtórzonych zdań, np.: „{powt[0][0][:100]}…” ×{powt[0][1]}") if powt else "")
    # 9. niewypelnione pola
    pola = re.findall(r"\[(?:DO UZUPEŁNIENIA[^\]]*|NR[^\]]*|DATA|ZAMAWIAJĄCY|sprawdzić[^\]]*|nr uprawnień[^\]]*|\s*|\?[^\]]+)\]", tekst)
    wynik("✓" if not pola else "✗", "Brak niewypełnionych pól i znaczników roboczych",
          (f"{len(pola)}: " + ", ".join(sorted(set(pola))[:8])) if pola else "")
    # raport
    L = [f"# Zgodność z Instrukcją 02 — {Path(a.plik).name}", "",
         f"Sprawdzono {date.today().strftime('%d.%m.%Y')}. Legenda: ✓ zgodne, ⚠ do poprawy, ✗ niezgodne, ? nie sprawdzono.", "",
         "| | Wymaganie | Szczegóły |", "|---|---|---|"]
    L += [f"| {s} | {c} | {x} |" for s, c, x in wyniki]
    n_bad = sum(1 for s, _, _ in wyniki if s == "✗")
    n_warn = sum(1 for s, _, _ in wyniki if s == "⚠")
    L += ["", f"Podsumowanie: {n_bad} niezgodności, {n_warn} uwag. Lista kontrolna wydania (Zał. 4 do ZEW), poz. 14: "
          + ("**do poprawy**" if n_bad or n_warn else "**zgodne**") + "."]
    raport = "\n".join(L)
    if a.out:
        Path(a.out).write_text(raport + "\n", encoding="utf-8")
    print(raport)


def main():
    import signal
    if hasattr(signal, "SIGPIPE"):
        signal.signal(signal.SIGPIPE, signal.SIG_DFL)
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("wzor"); s.add_argument("--typ", default="PB", choices=list(WZORY)); s.add_argument("--out", default="opis.json")
    s = sub.add_parser("lokalizacja")
    g = s.add_mutually_exclusive_group(required=True)
    g.add_argument("--xy"); g.add_argument("--miejscowosc"); g.add_argument("--dzialka")
    s.add_argument("--uklad", default="2180"); s.add_argument("--gmina"); s.add_argument("--out", required=True)
    s.add_argument("--promien", type=float, default=1500)
    s = sub.add_parser("dzialki")
    g = s.add_mutually_exclusive_group(required=True)
    g.add_argument("--id"); g.add_argument("--xy")
    s.add_argument("--uklad", default="2180"); s.add_argument("--md", action="store_true")
    s = sub.add_parser("przepisy"); s.add_argument("--szukaj"); s.add_argument("--klucz"); s.add_argument("--offline", action="store_true")
    s = sub.add_parser("cytuj"); s.add_argument("akt"); s.add_argument("--offline", action="store_true")
    s = sub.add_parser("generuj"); s.add_argument("spec"); s.add_argument("--out", required=True); s.add_argument("--md")
    s.add_argument("--offline", action="store_true")
    s = sub.add_parser("sprawdz"); s.add_argument("plik"); s.add_argument("--offline", action="store_true"); s.add_argument("--out")
    a = ap.parse_args()
    {"wzor": cmd_wzor, "lokalizacja": cmd_lokalizacja, "dzialki": cmd_dzialki, "przepisy": cmd_przepisy, "cytuj": cmd_cytuj,
     "generuj": cmd_generuj, "sprawdz": cmd_sprawdz}[a.cmd](a)


if __name__ == "__main__":
    main()
