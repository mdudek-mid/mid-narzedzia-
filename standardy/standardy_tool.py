#!/usr/bin/env python3
"""standardy_tool.py - nazwy plikow i folderow wg standardow biura (konfiguracja w prywatnym JSON).

  nazwa CZLON [CZLON ...] [--data RRRR-MM-DD|brak] [--wersja N] [--ext docx]
        sekcje laczone "_", slowa w sekcji "-", bez polskich znakow, spacji, # i !; prefiks RRRRMMDD_
        np.: nazwa PAB M OT --ext docx  ->  20261012_PAB_M_OT.docx
  sprawdz KATALOG [--bez-daty] [--max-sciezka 200]
        nazwy niezgodne ze standardem (znaki, brak prefiksu daty w plikach), puste foldery, za dlugie sciezki
  folder RODZAJ --standardy standardy.json      gdzie zapisac dany rodzaj dokumentu (lista: folder --standardy X)
  na TYP NR BRANZA OPIS [--data] --standardy standardy.json
        nazwa folderu wystapienia w nadzorze autorskim, np. ZSP 12 M "lozyska podpory P1"
  skrot "pelna nazwa instytucji" --standardy standardy.json   skrot zamawiajacego/odbiorcy z tabeli skrotow
"""
import argparse, json, os, re, sys, unicodedata
from datetime import date
from pathlib import Path

PL = str.maketrans("ąćęłńóśźżĄĆĘŁŃÓŚŹŻ", "acelnoszzACELNOSZZ")
ZAKAZANE = re.compile(r"[ #!ąćęłńóśźżĄĆĘŁŃÓŚŹŻ]")


def ascii_pl(s):
    s = s.translate(PL)
    s = unicodedata.normalize("NFKD", s)
    return "".join(c for c in s if not unicodedata.combining(c))


def sekcja(s):
    """jedna sekcja nazwy: slowa rozdzielone '-', tylko [A-Za-z0-9.-]"""
    s = ascii_pl(str(s)).strip()
    s = re.sub(r"[#!]", "", s)
    s = re.sub(r"[\s/\\:;,]+", "-", s)
    s = re.sub(r"[^A-Za-z0-9.\-+]", "", s)
    return re.sub(r"-{2,}", "-", s).strip("-")


def zbuduj(czlony, data=None, wersja=None, ext=None):
    sek = [sekcja(c) for c in czlony if str(c).strip()]
    if data != "brak":
        d = date.fromisoformat(data) if data else date.today()
        sek.insert(0, d.strftime("%Y%m%d"))
    n = "_".join(s for s in sek if s)
    if wersja:
        n += f"_v{int(wersja):02d}"
    if ext:
        n += "." + ext.lstrip(".")
    return n


def wczytaj(p, wymagany=True):
    if not p:
        if not wymagany:
            return {}
        sys.exit("Brak pliku konfiguracji biura (--standardy albo MID_STANDARDY). Foldery tematu i skróty instytucji: "
                 "wewnętrzna Instrukcja 01 biura. Komendy `nazwa`, `na` i `sprawdz` działają bez konfiguracji.")
    return json.loads(Path(p).read_text(encoding="utf-8"))


def cmd_nazwa(a):
    print(zbuduj(a.czlony, a.data, a.wersja, a.ext))


def cmd_sprawdz(a):
    k = Path(a.katalog)
    zle, bez_daty, dlugie, puste = [], [], [], []
    for p in sorted(k.rglob("*")):
        rel = p.relative_to(k)
        if any(x.startswith(".") for x in rel.parts) or p.name.lower() in ("desktop.ini", "thumbs.db"):
            continue
        if ZAKAZANE.search(p.name):
            zle.append(str(rel))
        if p.is_file() and not a.bez_daty and not re.match(r"\d{8}_", p.name):
            bez_daty.append(str(rel))
        if len(str(p.resolve())) > a.max_sciezka:
            dlugie.append(str(rel))
        if p.is_dir() and not any(p.iterdir()):
            puste.append(str(rel))
    print(f"# Sprawdzenie nazw: {k}\n")
    for tyt, lst, rada in (("Niedozwolone znaki (polskie litery, spacja, #, !)", zle, "zmień nazwę: `nazwa` podpowie wersję zgodną"),
                           ("Pliki bez prefiksu daty RRRRMMDD_", bez_daty, "dotyczy plików projektowych; wzory i pliki z zewnątrz mogą mieć własne nazwy"),
                           (f"Ścieżki dłuższe niż {a.max_sciezka} znaków", dlugie, "skróć nazwy folderów"),
                           ("Puste foldery", puste, "nie tworzy się pustych folderów – usuń albo zostaw do decyzji koordynatora")):
        print(f"## {tyt}: {len(lst)}")
        for x in lst[:a.limit]:
            print(f"- {x}" + (f"  →  {zbuduj([Path(x).stem], 'brak')}{Path(x).suffix}" if lst is zle else ""))
        if len(lst) > a.limit:
            print(f"- … i {len(lst) - a.limit} więcej")
        if lst:
            print(f"  ({rada})")
        print()
    print("Narzędzie niczego nie zmienia. Zmiany nazw tylko po decyzji koordynatora tematu (linki i odwołania w innych plikach).")


def cmd_folder(a):
    st = wczytaj(a.standardy)
    f = st.get("foldery", {})
    if not a.rodzaj:
        for k, v in f.items():
            print(f"{k:34} {v}")
        return
    trafione = [k for k in f if a.rodzaj.lower() in k.lower()]
    if not trafione:
        sys.exit(f"Nie znam rodzaju „{a.rodzaj}”. Dostępne: {', '.join(f)}")
    tem = re.sub(r"\s*\(.*\)", "", st.get("folder_tematu", "KOD_ZAMAWIAJACY_NR-DROGI_NAZWA"))
    for k in trafione:
        print(f"{k}: " + (f[k] if f[k].startswith("MiD_dokumenty") else f"{tem}/{f[k]}"))
    wr = st.get("dane_wrazliwe", {})
    if wr and re.search(r"protok|rf|umow|faktur|ofert|zestaw|rozlicz", a.rodzaj, re.I):
        print(f"\nDane wrażliwe: {wr.get('zasada', '')}")
        for k, v in (wr.get("foldery_MiD_dokumenty") or {}).items():
            if re.search(r"protok|rozlicz|zestaw", v + k, re.I):
                print(f"  {k}: {v}")


def cmd_na(a):
    st = wczytaj(a.standardy, wymagany=False)      # bez konfiguracji: bez kontroli typu i branzy
    na = st.get("nadzor_autorski", {})
    typy, branze = na.get("typy", {}), st.get("branze", {})
    if typy and a.typ.upper() not in typy:
        print(f"UWAGA: typ {a.typ} spoza listy ({', '.join(typy)}) – dopuszczalne tymczasowo, uzupełnij po wyjaśnieniu.", file=sys.stderr)
    if branze and a.branza.upper() not in branze:
        sys.exit(f"Nieznana branża {a.branza}. Dostępne: {', '.join(f'{k}={v}' for k, v in branze.items())}")
    print(zbuduj([a.typ.upper(), f"{int(a.nr):03d}", a.branza.upper(), a.opis[:1].lower() + a.opis[1:]], a.data))


def cmd_skrot(a):
    st = wczytaj(a.standardy)
    sk = {k: v for k, v in st.get("skroty", {}).items() if not k.startswith("_")}
    slowa = lambda t: set(re.findall(r"[a-z0-9]{3,}", ascii_pl(t).lower()))
    n = slowa(a.nazwa)
    traf = []
    for k, v in sk.items():
        w = slowa(v)
        if w and (len(w & n) / len(w) >= 0.6 or ascii_pl(k).lower() == ascii_pl(a.nazwa).lower()):
            traf.append((len(w & n), k, v))
    traf = [(k, v) for _, k, v in sorted(traf, key=lambda x: -x[0])]
    if not traf:
        print(f"Brak w tabeli skrótów – nadaj skrót indywidualnie (decyzja koordynatora), np. {sekcja(a.nazwa).upper()[:12]}")
    for k, v in traf[:5]:
        print(f"{k:12} {v}")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("nazwa"); s.add_argument("czlony", nargs="+"); s.add_argument("--data"); s.add_argument("--wersja", type=int); s.add_argument("--ext")
    s = sub.add_parser("sprawdz"); s.add_argument("katalog"); s.add_argument("--bez-daty", action="store_true")
    s.add_argument("--max-sciezka", type=int, default=200); s.add_argument("--limit", type=int, default=40)
    s = sub.add_parser("folder"); s.add_argument("rodzaj", nargs="?"); s.add_argument("--standardy", default=os.environ.get("MID_STANDARDY"))
    s = sub.add_parser("na"); s.add_argument("typ"); s.add_argument("nr"); s.add_argument("branza"); s.add_argument("opis"); s.add_argument("--data")
    s.add_argument("--standardy", default=os.environ.get("MID_STANDARDY"))
    s = sub.add_parser("skrot"); s.add_argument("nazwa"); s.add_argument("--standardy", default=os.environ.get("MID_STANDARDY"))
    a = ap.parse_args()
    {"nazwa": cmd_nazwa, "sprawdz": cmd_sprawdz, "folder": cmd_folder, "na": cmd_na, "skrot": cmd_skrot}[a.cmd](a)


if __name__ == "__main__":
    main()
