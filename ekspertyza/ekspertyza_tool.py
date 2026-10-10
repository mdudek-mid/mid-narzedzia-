#!/usr/bin/env python3
"""ekspertyza_tool.py - ekspertyzy techniczne obiektow mostowych Pracowni MiD.

Komendy:
  wzor --out ekspertyza.json                 szablon specyfikacji (obiekt, elementy, badania, nosnosc, zalecenia)
  skala                                      skala ocen GDDKiA, lista elementow, tryby robot, ocena ogolna
  ocena SPEC.json                            ocena srednia i ogolna obiektu, kontrola spojnosci ocen i zalecen
  badania SPEC.json                          interpretacja badan: sklerometr, wytrzymalosc, karbonatyzacja/pH,
                                             chlorki, potencjaly, rezystywnosc, ubytki zbrojenia
  raport SPEC.json --out Ekspertyza.docx [--md Ekspertyza.md] [--opis-tool sciezka/opis_tool.py]
        dokument wg Instrukcji 02 (opis_tool.py) z rozdzialami: charakterystyka, stan techniczny, badania,
        nosnosc (wyniki mes_tool.py nosnosc --json), wnioski i zalecenia, warianty

Zrodla skal: Instrukcja przeprowadzania przegladow drogowych obiektow inzynierskich GDDKiA (Zarz. nr 14 GDDKiA z 7.07.2005,
cz. III); dla biezacych zlecen sprawdz wersje instrukcji wskazana w OPZ.
"""
import argparse, json, math, os, subprocess, sys
from datetime import date
from pathlib import Path

SKALA = {5: "odpowiedni", 4: "zadowalający", 3: "niepokojący", 2: "niedostateczny", 1: "przedawaryjny", 0: "awaryjny"}
SKALA_OPIS = {
    5: "brak uszkodzeń i zanieczyszczeń możliwych do wykrycia podczas przeglądu",
    4: "zanieczyszczenia lub pierwsze oznaki uszkodzeń pogarszające wygląd",
    3: "uszkodzenia, które pozostawione bez naprawy skrócą bezpieczny okres użytkowania",
    2: "uszkodzenia obniżające przydatność, możliwe do naprawy",
    1: "nieodwracalne uszkodzenia dyskwalifikujące element z użytkowania",
    0: "element zniszczony albo nie istnieje",
}
SKALA_IZOLACJI = {5: "odpowiednia – brak przecieków", 2: "niedostateczna – nieliczne, niewielkie przecieki",
                  0: "awaryjna – rozległe przecieki obniżające trwałość"}
SKALA_PRZYDATNOSCI = {5: "odpowiednia", 2: "ograniczona", 0: "niedostateczna"}
TRYBY = {"A": "roboty awaryjne – natychmiast, poza planem robót bieżącego roku",
         "1": "roboty do wykonania w następnym roku",
         "2": "roboty w drugiej kolejności w kolejnych latach",
         "3": "roboty w trzeciej kolejności w kolejnych latach"}
ELEMENTY = [
    (1, "Nasypy i skarpy"), (2, "Dojazdy w obrębie skrzydeł"), (3, "Nawierzchnia jezdni"),
    (4, "Nawierzchnia chodników i krawężniki"), (5, "Balustrady, bariery ochronne i ekrany"),
    (6, "Belki podporęczowe i gzymsy"), (7, "Urządzenia odwadniające"), (8, "Izolacja pomostu"),
    (9, "Pomost"), (10, "Dźwigary główne"), (11, "Łożyska"), (12, "Urządzenia dylatacyjne"),
    (13, "Przyczółki"), (14, "Filary"), (15, "Koryto cieku i przestrzeń pod obiektem"), (16, "Przeguby"),
    (17, "Konstrukcje oporowe i skrzydła"), (18, "Urządzenia ochrony środowiska"), (19, "Zakotwienia kabli"),
    (20, "Kable sprężające"), (21, "Urządzenia obce"),
]
ZRODLO_SKALI = ("Instrukcja przeprowadzania przeglądów drogowych obiektów inżynierskich, cz. III – przeglądy podstawowe "
                "i rozszerzone, wprowadzona Zarządzeniem nr 14 Generalnego Dyrektora Dróg Krajowych i Autostrad z dnia 7 lipca 2005 r.")


# ----------------------------------------------------------------------------- ocena stanu
def ocena_obiektu(spec):
    el = [e for e in spec.get("elementy", []) if isinstance(e.get("ocena"), (int, float))]
    if not el:
        return None
    srednia = sum(e["ocena"] for e in el) / len(el)

    def ocena_nr(nr):
        v = [e["ocena"] for e in el if e.get("nr") == nr]
        return min(v) if v else None
    pomost, dzw = ocena_nr(9), ocena_nr(10)
    przycz = [e["ocena"] for e in el if e.get("nr") == 13]
    filary = [e["ocena"] for e in el if e.get("nr") == 14]
    podpory = None
    if przycz and filary:
        podpory = (min(przycz) + min(filary)) / 2
    elif przycz:
        podpory = min(przycz)
    skl = {"średnia wszystkich ocenionych elementów": srednia, "pomost": pomost, "dźwigary główne": dzw,
           "podpory ((min przyczółek + min filar)/2)": podpory}
    skl = {k: v for k, v in skl.items() if v is not None}
    ogolna_k = min(skl, key=skl.get)
    return {"srednia": round(srednia, 2), "ogolna": round(skl[ogolna_k], 2), "decyduje": ogolna_k, "skladniki": skl,
            "liczba_elementow": len(el)}


def kontrola(spec):
    uw = []
    for e in spec.get("elementy", []):
        o = e.get("ocena")
        if o is None or o == "n/d":
            continue
        if e.get("nr") == 8 and o not in (5, 2, 0):
            uw.append(f"Izolacja (el. 8): ocena {o} spoza skali 5/2/0.")
        if o not in SKALA:
            uw.append(f"{e.get('nazwa')}: ocena {o} spoza skali 0–5.")
        zal = [z for z in spec.get("zalecenia", []) if e.get("nr") in (z.get("elementy") or [])]
        if o <= 2 and not zal:
            uw.append(f"{e.get('nazwa')} ({o} – {SKALA.get(o)}): brak zalecenia naprawy.")
        if o <= 1 and not any(z.get("tryb") in ("A", "1") for z in zal):
            uw.append(f"{e.get('nazwa')} ({o} – {SKALA.get(o)}): zalecenie powinno mieć tryb A albo 1.")
        if o <= 2 and not (e.get("opis") or e.get("uszkodzenia")):
            uw.append(f"{e.get('nazwa')}: brak opisu uszkodzeń uzasadniającego ocenę {o}.")
    for z in spec.get("zalecenia", []):
        if str(z.get("tryb")) not in TRYBY:
            uw.append(f"Zalecenie „{z.get('opis', '')[:50]}”: tryb {z.get('tryb')} spoza A/1/2/3.")
    return uw


# ----------------------------------------------------------------------------- badania
def sklerometr_fR(R):
    """PN-EN 13791:2007, krzywa podstawowa (orientacyjnie; wymaga kalibracji odwiertami)"""
    if R < 20:
        return None
    if R <= 24:
        return 1.25 * R - 23
    if R <= 50:
        return 1.73 * R - 34.5
    return 1.73 * 50 - 34.5


def klasa_kostkowa(fc_cube):
    klasy = [(12, 15), (16, 20), (20, 25), (25, 30), (30, 37), (35, 45), (40, 50), (45, 55), (50, 60)]
    wyb = None
    for fck, fcc in klasy:
        if fcc <= fc_cube + 1e-9:
            wyb = (fck, fcc)
    return f"C{wyb[0]}/{wyb[1]}" if wyb else "< C12/15"


def klasa_cylindryczna_insitu(fck_is):
    """PN-EN 13791: f_ck,is >= 0,85 f_ck (min. wymaganie dla klasy) - najwyzsza klasa spelniajaca"""
    klasy = [(12, 15), (16, 20), (20, 25), (25, 30), (30, 37), (35, 45), (40, 50), (45, 55), (50, 60)]
    wyb = None
    for fck, fcc in klasy:
        if 0.85 * fck <= fck_is + 1e-9:
            wyb = (fck, fcc)
    return f"C{wyb[0]}/{wyb[1]}" if wyb else "< C12/15"


def interpretuj_badania(spec):
    b = spec.get("badania", {})
    wiek = b.get("wiek_lat") or (date.today().year - int(spec.get("obiekt", {}).get("rok_budowy", date.today().year)))
    out = {"wiek_lat": wiek, "sklerometr": [], "wytrzymalosc": [], "karbonatyzacja": [], "chlorki": [], "potencjaly": [],
           "rezystywnosc": [], "korozja": []}
    for s in b.get("sklerometr", []):
        R = s.get("odczyty") or []
        if R:
            import statistics
            Rm = statistics.mean(R)
            odrzuc = [r for r in R if abs(r - Rm) > 5]   # odczyty odbiegajace od sredniej o > 5 jednostek
            Rk = [r for r in R if abs(r - Rm) <= 5] or R
            fr = [sklerometr_fR(r) for r in Rk]
            fr = [f for f in fr if f is not None]
            if fr:
                fm = statistics.mean(fr)
                sd = statistics.stdev(fr) if len(fr) > 1 else 0.0
                fck_is = min(fm - 1.48 * max(sd, 2.0), min(fr) + 4)
                out["sklerometr"].append({"miejsce": s.get("miejsce"), "n": len(Rk), "odrzucone": len(odrzuc), "R_sr": round(statistics.mean(Rk), 1),
                                          "f_R_sr": round(fm, 1), "s": round(sd, 1), "f_ck_is": round(fck_is, 1),
                                          "klasa": klasa_cylindryczna_insitu(fck_is)})
        elif s.get("wytrzymalosc_kostkowa"):
            out["wytrzymalosc"].append({"miejsce": s.get("miejsce"), "f_c_cube": s["wytrzymalosc_kostkowa"],
                                        "klasa": klasa_kostkowa(s["wytrzymalosc_kostkowa"]), "metoda": s.get("metoda", "sklerometr – wynik przyrządu")})
    for k in b.get("karbonatyzacja", []):
        dk, c = k.get("glebokosc_mm"), k.get("otulina_mm")
        K = dk / math.sqrt(wiek) if dk and wiek else None
        rec = {"miejsce": k.get("miejsce"), "d_k": dk, "c": c, "K": round(K, 2) if K else None}
        if c is not None and dk is not None:
            if dk >= c:
                rec["wniosek"] = "front karbonatyzacji osiągnął zbrojenie – stal zdepasywowana, korozja możliwa przy dostępie wilgoci"
                rec["lat_do_zbrojenia"] = 0
            elif K:
                t = (c / K) ** 2 - wiek
                rec["lat_do_zbrojenia"] = round(t)
                rec["wniosek"] = f"zbrojenie w strefie alkalicznej; przy d = K√t front dojdzie do zbrojenia za ok. {lat(t)}"
        out["karbonatyzacja"].append(rec)
    for k in b.get("chlorki", []):
        v, odn = k.get("zawartosc_proc"), k.get("odniesienie", "cement")
        if odn == "beton":
            cem = k.get("cement_kg_m3", 350)
            v_c = v * 2400 / cem
        else:
            v_c = v
        prog = 0.2 if k.get("sprezony") else 0.4
        out["chlorki"].append({"miejsce": k.get("miejsce"), "glebokosc_mm": k.get("glebokosc_mm"), "Cl_masa_cementu_proc": round(v_c, 3),
                               "prog": prog, "wniosek": "przekroczony próg inicjacji korozji" if v_c >= prog else "poniżej progu inicjacji korozji"})
    for k in b.get("potencjaly", []):
        E = k.get("E_mV")
        if E is None:
            continue
        ocena = "prawdopodobieństwo korozji < 10%" if E > -200 else ("strefa niepewna" if E >= -350 else "prawdopodobieństwo korozji > 90%")
        out["potencjaly"].append({"miejsce": k.get("miejsce"), "E_mV": E, "ocena": ocena})
    for k in b.get("rezystywnosc", []):
        r = k.get("kOhm_cm")
        if r is None:
            continue
        ocena = "pomijalne" if r > 100 else ("niskie" if r > 50 else ("umiarkowane do wysokiego" if r > 10 else "wysokie"))
        out["rezystywnosc"].append({"miejsce": k.get("miejsce"), "kOhm_cm": r, "ryzyko_korozji": ocena})
    for k in b.get("korozja_zbrojenia", []):
        ub = k.get("ubytek_sredni_proc", 0.0)
        out["korozja"].append({"miejsce": k.get("miejsce"), "ubytek_sredni_proc": ub, "wsp_As": round(1 - ub / 100, 3),
                               "opis": k.get("opis", "")})
    return out


# ----------------------------------------------------------------------------- raport (spec dla opis_tool)
def lat(n):
    n = int(round(n))
    if n == 1:
        return "1 rok"
    if n % 10 in (2, 3, 4) and n % 100 not in (12, 13, 14):
        return f"{n} lata"
    return f"{n} lat"


def fmt(v, n=1):
    if v is None:
        return "—"
    if isinstance(v, float):
        return f"{v:,.{n}f}".replace(",", " ").replace(".", ",")
    return str(v)


def znajdz_opis_tool(sciezka=None):
    kand = [sciezka, Path(__file__).resolve().parent.parent / "opis" / "opis_tool.py", Path(__file__).resolve().parent / "opis_tool.py",
            Path.home() / ".local/opis/opis_tool.py", Path("/tmp/mid-narzedzia/opis/opis_tool.py")]
    for k in kand:
        if k and Path(k).exists():
            return Path(k)
    sys.exit("Nie znaleziono opis_tool.py (skill opis-techniczny). Podaj --opis-tool.")


def rozdzialy_ekspertyzy(spec, baza):
    ob = spec.get("obiekt", {})
    R = []
    # Charakterystyka
    par = ob.get("parametry", {})
    ch = {"tytul": "Charakterystyka obiektu", "tresc": spec.get("charakterystyka", {}).get("ogolna", ""), "tabele": [], "podrozdzialy": []}
    if par:
        ch["tabele"].append({"podpis": "Podstawowe parametry obiektu", "naglowek": ["Parametr", "Wartość"],
                             "wiersze": [[k, v] for k, v in par.items()], "szerokosci": [7, 10]})
    for k, v in (spec.get("charakterystyka", {}).get("elementy") or {}).items():
        ch["podrozdzialy"].append({"tytul": k, "tresc": v})
    R.append(ch)
    # Stan techniczny
    oc = ocena_obiektu(spec)
    wiersze = []
    for e in sorted(spec.get("elementy", []), key=lambda e: e.get("nr", 99)):
        o = e.get("ocena")
        ocena_txt = "n/d" if o in (None, "n/d") else f"{o} – {SKALA.get(o, '?')}"
        wiersze.append([e.get("nr", ""), e.get("nazwa", ""), ocena_txt, (e.get("opis") or "") + (f" ({e['foto']})" if e.get("foto") else "")])
    st = {"tytul": "Opis i ocena stanu technicznego", "tresc": spec.get("stan", {}).get("ogolny", ""), "tabele": []}
    st["tabele"].append({"podpis": "Oceny stanu technicznego elementów (skala 0–5 wg [@W_przeglady])",
                         "naglowek": ["Nr", "Element", "Ocena", "Opis stanu i uszkodzeń"], "wiersze": wiersze, "szerokosci": [1, 4.2, 3, 8.8]})
    if oc:
        st["tresc"] += ("\n\n" if st["tresc"] else "") + (
            f"Ocena średnia obiektu (średnia arytmetyczna {oc['liczba_elementow']} ocenionych elementów) wynosi {fmt(oc['srednia'], 2)}. "
            f"Ocena ogólna obiektu, przyjęta jako najmniejsza z: oceny średniej, oceny pomostu, oceny dźwigarów głównych i oceny podpór, "
            f"wynosi {fmt(oc['ogolna'], 2)} (decyduje: {oc['decyduje']}).")
    R.append(st)
    # Badania
    b = interpretuj_badania(spec)
    bad = {"tytul": "Badania i pomiary", "tresc": spec.get("badania", {}).get("wstep", ""), "podrozdzialy": []}
    if b["wytrzymalosc"] or b["sklerometr"]:
        p = {"tytul": "Wytrzymałość betonu na ściskanie", "tresc": spec.get("badania", {}).get("sklerometr_opis", ""), "tabele": []}
        if b["sklerometr"]:
            p["tabele"].append({"podpis": "Badanie sklerometryczne – krzywa podstawowa PN-EN 13791:2007 (orientacyjnie, bez kalibracji odwiertami)",
                                "naglowek": ["Miejsce", "n", "R śr.", "f_R śr. [MPa]", "s [MPa]", "f_ck,is [MPa]", "Klasa (orient.)"],
                                "wiersze": [[x["miejsce"], x["n"], fmt(x["R_sr"]), fmt(x["f_R_sr"]), fmt(x["s"]), fmt(x["f_ck_is"]), x["klasa"]] for x in b["sklerometr"]]})
        if b["wytrzymalosc"]:
            p["tabele"].append({"podpis": "Wytrzymałość betonu (wytrzymałość kostkowa z badania) i odpowiadająca klasa wg PN-EN 206",
                                "naglowek": ["Miejsce", "Metoda", "f_c,cube [MPa]", "Klasa"],
                                "wiersze": [[x["miejsce"], x["metoda"], fmt(float(x["f_c_cube"])), x["klasa"]] for x in b["wytrzymalosc"]]})
        bad["podrozdzialy"].append(p)
    if b["karbonatyzacja"]:
        bad["podrozdzialy"].append({"tytul": "Karbonatyzacja i odczyn pH betonu", "tresc": spec.get("badania", {}).get("karbonatyzacja_opis", ""),
                                    "tabele": [{"podpis": f"Głębokość karbonatyzacji a otulina (wiek betonu {lat(b['wiek_lat'])}; d = K√t)",
                                                "naglowek": ["Miejsce", "d_k [mm]", "c [mm]", "K [mm/√rok]", "Wniosek"],
                                                "wiersze": [[x["miejsce"], fmt(x["d_k"]), fmt(x["c"]), fmt(x["K"], 2), x.get("wniosek", "")] for x in b["karbonatyzacja"]]}]})
    if b["chlorki"]:
        bad["podrozdzialy"].append({"tytul": "Zawartość chlorków", "tabele": [{"podpis": "Chlorki (% masy cementu); próg 0,4% dla żelbetu, 0,2% dla betonu sprężonego",
                                    "naglowek": ["Miejsce", "Głębokość [mm]", "Cl⁻ [% m.c.]", "Próg", "Wniosek"],
                                    "wiersze": [[x["miejsce"], fmt(x["glebokosc_mm"]), fmt(x["Cl_masa_cementu_proc"], 3), fmt(x["prog"], 1), x["wniosek"]] for x in b["chlorki"]]}]})
    if b["potencjaly"]:
        bad["podrozdzialy"].append({"tytul": "Potencjały korozyjne zbrojenia", "tabele": [{"podpis": "Potencjał stali względem Cu/CuSO₄ – kryteria ASTM C876",
                                    "naglowek": ["Miejsce", "E [mV]", "Ocena"], "wiersze": [[x["miejsce"], x["E_mV"], x["ocena"]] for x in b["potencjaly"]]}]})
    if b["rezystywnosc"]:
        bad["podrozdzialy"].append({"tytul": "Rezystywność betonu", "tabele": [{"podpis": "Rezystywność (metoda Wennera) – orientacyjne ryzyko korozji",
                                    "naglowek": ["Miejsce", "ρ [kΩ·cm]", "Ryzyko korozji"], "wiersze": [[x["miejsce"], x["kOhm_cm"], x["ryzyko_korozji"]] for x in b["rezystywnosc"]]}]})
    if b["korozja"]:
        bad["podrozdzialy"].append({"tytul": "Odkrywki i ubytki korozyjne zbrojenia", "tresc": spec.get("badania", {}).get("korozja_opis", ""),
                                    "tabele": [{"podpis": "Ubytki przekroju zbrojenia", "naglowek": ["Miejsce", "Ubytek średni [%]", "Współczynnik As", "Opis"],
                                                "wiersze": [[x["miejsce"], fmt(float(x["ubytek_sredni_proc"])), fmt(x["wsp_As"], 3), x["opis"]] for x in b["korozja"]]}]})
    R.append(bad)
    # Nosnosc
    n = spec.get("nosnosc", {})
    nos = {"tytul": "Obliczenia nośności", "tresc": n.get("opis_modelu", ""), "podrozdzialy": [], "tabele": [], "rysunki": []}
    wj = None
    if n.get("wyniki_mes") and (baza / n["wyniki_mes"]).exists():
        wj = json.loads((baza / n["wyniki_mes"]).read_text(encoding="utf-8"))
    if wj:
        g = wj["nosnosc"]
        wt = g.get("wg_typu", {})
        wiersze = []
        for typ, naz in (("M", "moment zginający"), ("V", "siła poprzeczna")):
            if typ in wt:
                t = wt[typ]
                wiersze.append([naz, t["kategoria"], fmt(t["mu"], 1) if t.get("mu") else "—", f"przęsło {t['przeslo']}, x = {fmt(float(t['x']), 2)} m, dźwigar {t['dzwigar']}"])
        nos["tabele"].append({"podpis": f"Nośność użytkowa wg [@W1] – porównanie z obciążeniem normowym {g['norma']}",
                              "naglowek": ["Wielkość", "Kategoria", "m_u [t]", "Miejsce decydujące"], "wiersze": wiersze})
        m = g.get("miarodajny") or {}
        nos["tresc"] += ("\n\n" if nos["tresc"] else "") + (
            f"Nośność użytkowa obiektu: kategoria {m.get('kategoria')}" + (f", m_u = {fmt(m.get('mu'), 1)} t" if m.get("mu") else "") + ".")
        mlc = wj.get("MLC")
        if mlc:
            nos["tabele"].append({"podpis": f"Klasa obciążenia wojskowego MLC wg [@ptb_drogi_2022] (zał. 2), sprawdzono: {mlc.get('_odpowiedzi', 'wszystkie odpowiedzi')}",
                                  "naglowek": ["Pojazdy", "Kolumny", "Dopuszczalna klasa MLC"],
                                  "wiersze": [["kołowe", 1, mlc.get("K1") or "< 40"], ["kołowe", 2, mlc.get("K2") or "< 40"],
                                              ["gąsienicowe", 1, mlc.get("G1") or "< 40"], ["gąsienicowe", 2, mlc.get("G2") or "< 40"]]})
        if wj.get("RF"):
            nos["tabele"].append({"podpis": "Współczynnik nośności RF (najmniejsze wartości)", "naglowek": ["Przęsło", "Odp.", "Dźwigar", "Model", "RF"],
                                  "wiersze": [[d["przeslo"], d["typ"], d["dzwigar"], d["model"], fmt(d["RF"], 2)] for d in wj["RF"][:6]]})
    for w in n.get("warianty_modelu", []):
        nos["podrozdzialy"].append({"tytul": w.get("tytul", "Wariant modelu"), "tresc": w.get("tresc", "")})
    for r in n.get("rysunki", []):
        nos["rysunki"].append(r)
    if n.get("wnioski"):
        nos["tresc"] += "\n\n" + n["wnioski"]
    R.append(nos)
    # Wnioski i zalecenia
    zal = spec.get("zalecenia", [])
    wz = {"tytul": "Wnioski i zalecenia", "tresc": spec.get("wnioski", ""), "tabele": [], "podrozdzialy": []}
    if zal:
        nazwy = {nr: nz for nr, nz in ELEMENTY}
        porz = {"A": 0, "1": 1, "2": 2, "3": 3}
        wz["tabele"].append({"podpis": "Zalecane roboty i tryb realizacji (A – natychmiast; 1 – w następnym roku; 2, 3 – w kolejnych latach)",
                             "naglowek": ["Lp.", "Tryb", "Element", "Zakres robót"],
                             "wiersze": [[i, z.get("tryb"), ", ".join(nazwy.get(e, str(e)) for e in z.get("elementy", [])), z.get("opis", "")]
                                         for i, z in enumerate(sorted(zal, key=lambda z: porz.get(str(z.get("tryb")), 9)), 1)],
                             "szerokosci": [1, 1.4, 4.6, 10]})
    for w in spec.get("warianty", []):
        wz["podrozdzialy"].append({"tytul": w.get("tytul", "Wariant"), "tresc": w.get("tresc", "")})
    if spec.get("waznosc"):
        wz["tresc"] += ("\n\n" if wz["tresc"] else "") + spec["waznosc"]
    R.append(wz)
    return R


def cmd_raport(a):
    spec = json.loads(Path(a.spec).read_text(encoding="utf-8"))
    baza = Path(a.spec).resolve().parent
    opis = dict(spec.get("opis", {}))
    opis.setdefault("typ", "ekspertyza")
    opis.setdefault("tytul_opracowania", "EKSPERTYZA TECHNICZNA")
    opis["rozdzialy"] = rozdzialy_ekspertyzy(spec, baza) + (opis.get("rozdzialy_dodatkowe") or [])
    mat = opis.setdefault("materialy", {})
    W = mat.setdefault("W", [])
    if not any(isinstance(x, dict) and x.get("klucz") == "W_przeglady" for x in W):
        W.append({"klucz": "W_przeglady", "tekst": spec.get("zrodlo_skali", ZRODLO_SKALI)})
    tmp = baza / f".{Path(a.out).stem}_opis.json"
    tmp.write_text(json.dumps(opis, ensure_ascii=False, indent=1), encoding="utf-8")
    ot = znajdz_opis_tool(a.opis_tool)
    cmd = [sys.executable, "-I", str(ot), "generuj", str(tmp), "--out", a.out] + (["--md", a.md] if a.md else [])
    r = subprocess.run(cmd, capture_output=True, text=True)
    sys.stdout.write(r.stdout)
    sys.stderr.write(r.stderr)
    if r.returncode:
        sys.exit(r.returncode)
    uw = kontrola(spec)
    if uw:
        print("\nKontrola spójności ocen i zaleceń:")
        for u in uw:
            print(" - " + u)
    print(f"\nSpecyfikacja dokumentu: {tmp} (można poprawić i wygenerować ponownie przez opis_tool.py generuj).")


def cmd_ocena(a):
    spec = json.loads(Path(a.spec).read_text(encoding="utf-8"))
    oc = ocena_obiektu(spec)
    print("| Nr | Element | Ocena |\n|---|---|---|")
    for e in sorted(spec.get("elementy", []), key=lambda e: e.get("nr", 99)):
        o = e.get("ocena")
        print(f"| {e.get('nr')} | {e.get('nazwa')} | {o if o is not None else 'n/d'} {('– ' + SKALA[o]) if o in SKALA else ''} |")
    if oc:
        print(f"\nOcena średnia: {oc['srednia']:.2f}; ocena ogólna: {oc['ogolna']:.2f} (decyduje: {oc['decyduje']})")
        for k, v in oc["skladniki"].items():
            print(f"  {k}: {v:.2f}")
    uw = kontrola(spec)
    print("\nKontrola spójności:" + ("" if uw else " bez uwag"))
    for u in uw:
        print(" - " + u)


def cmd_badania(a):
    spec = json.loads(Path(a.spec).read_text(encoding="utf-8"))
    print(json.dumps(interpretuj_badania(spec), ensure_ascii=False, indent=1))


def cmd_skala(a):
    print("Skala ocen stanu technicznego elementów:")
    for k in sorted(SKALA, reverse=True):
        print(f"  {k} – {SKALA[k]}: {SKALA_OPIS[k]}")
    print("Izolacja:", SKALA_IZOLACJI)
    print("Przydatność do użytkowania:", SKALA_PRZYDATNOSCI)
    print("Tryby robót:", TRYBY)
    print("Elementy:")
    for nr, n in ELEMENTY:
        print(f"  {nr:2d}. {n}")
    print("Ocena ogólna obiektu = min(średnia ocenionych elementów, pomost, dźwigary główne, (min przyczółek + min filar)/2).")
    print("Źródło:", ZRODLO_SKALI)


def cmd_wzor(a):
    spec = {
        "opis": {"tytul_opracowania": "EKSPERTYZA TECHNICZNA", "branza": "mostowa", "tom": "[obiekt, droga, km, JNI]",
                 "zamierzenie": "Ocena stanu technicznego i nośności obiektu", "obiekt": "[obiekt]",
                 "inwestor": {"nazwa": "[Zamawiający]", "adres": "[adres]"},
                 "umowa": {"nr": "[nr]", "data": "RRRR-MM-DD", "zamawiajacy": "[Zamawiającym – narzędnik]"},
                 "lokalizacja_json": "lokalizacja/lokalizacja.json", "lokalizacja_opis_dodatkowy": "[km drogi, przeszkoda]",
                 "dzialki": [], "zespol": [{"funkcja": "Autor ekspertyzy", "osoba": "[ ]", "uprawnienia": "[ ]"}],
                 "data": date.today().strftime("%m.%Y"), "rewizja": "00",
                 "przedmiot": "Przedmiotem opracowania jest ekspertyza techniczna [obiektu].",
                 "cel": {"kategoria": "eksperckie", "tekst": "Celem opracowania jest ocena stanu technicznego, nośności użytkowej wg [@W1] i klasy MLC oraz określenie zakresu i pilności robót."},
                 "zakres": ["oględziny i ocena elementów wg [@W_przeglady]", "badania materiałowe", "obliczenia nośności", "wnioski i zalecenia"],
                 "podstawa_merytoryczna": ["[@DA1] – dokumentacja archiwalna / poprzednia ekspertyza", "wyniki oględzin i badań autora"],
                 "materialy": {"DA": [{"klucz": "DA1", "tekst": "[ ]"}], "N": ["PN-EN 1990", "PN-EN 1991-2", "PN-EN 1992-1-1"],
                               "U": ["pb", "udp"], "R": ["ptb_drogi_2022"],
                               "W": [{"klucz": "W1", "tekst": "Instrukcja do określania nośności użytkowej drogowych obiektów mostowych, wprowadzona Zarządzeniem nr 17 Generalnego Dyrektora Dróg Krajowych i Autostrad z dnia 1 czerwca 2004 r."}]}},
        "obiekt": {"rok_budowy": 1978, "normatyw": "PN-66/B-02015 kl. I", "parametry": {"Długość obiektu": "[m]", "Rozpiętości teoretyczne": "[m]"}},
        "charakterystyka": {"ogolna": "", "elementy": {"Ustrój nośny": "", "Podpory": "", "Wyposażenie": ""}},
        "stan": {"ogolny": ""},
        "elementy": [{"nr": nr, "nazwa": n, "ocena": None, "opis": "", "foto": ""} for nr, n in ELEMENTY],
        "badania": {"wiek_lat": None, "sklerometr": [{"miejsce": "dźwigar – spód", "odczyty": []}, {"miejsce": "przyczółek", "wytrzymalosc_kostkowa": None}],
                    "karbonatyzacja": [{"miejsce": "", "glebokosc_mm": None, "otulina_mm": None}],
                    "chlorki": [], "potencjaly": [], "rezystywnosc": [], "korozja_zbrojenia": []},
        "nosnosc": {"wyniki_mes": "nosnosc.json", "opis_modelu": "", "warianty_modelu": [], "rysunki": [], "wnioski": ""},
        "zalecenia": [{"tryb": "1", "elementy": [8], "opis": ""}],
        "warianty": [{"tytul": "Wariant I – remont", "tresc": ""}, {"tytul": "Wariant II – przebudowa", "tresc": ""}],
        "wnioski": "", "waznosc": "",
    }
    Path(a.out).write_text(json.dumps(spec, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"Zapisano {a.out}")


def main():
    import signal
    if hasattr(signal, "SIGPIPE"):
        signal.signal(signal.SIGPIPE, signal.SIG_DFL)
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("wzor"); s.add_argument("--out", default="ekspertyza.json")
    sub.add_parser("skala")
    s = sub.add_parser("ocena"); s.add_argument("spec")
    s = sub.add_parser("badania"); s.add_argument("spec")
    s = sub.add_parser("raport"); s.add_argument("spec"); s.add_argument("--out", required=True); s.add_argument("--md"); s.add_argument("--opis-tool")
    a = ap.parse_args()
    {"wzor": cmd_wzor, "skala": cmd_skala, "ocena": cmd_ocena, "badania": cmd_badania, "raport": cmd_raport}[a.cmd](a)


if __name__ == "__main__":
    main()
