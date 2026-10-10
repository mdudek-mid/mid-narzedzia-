#!/usr/bin/env python3
"""swz_tool.py - analiza dokumentacji przetargowej (SWZ/IDW, OPZ, PFU, umowa, wyjasnienia). Pracownia Projektowa MiD.

Komendy:
  pobierz URL|--nr N --out KATALOG [--repo mid-przetargi] [--max-mb 60] [--max-dok 40]
        dokumenty z platformy (wbudowane: e-Zamowienia, platformazakupowa; z repo mid-przetargi takze eB2B,
        PLK, logintrade, Baza Konkurencyjnosci) do KATALOG/pliki + _lista.json; archiwa rozpakowane
  teksty WEJSCIE... --out KATALOG_TXT [--ocr auto|tak|nie] [--maks-ocr 40]
        PDF/DOCX/DOC/RTF/ODT/XLSX/XLS/TXT/ZIP/7Z -> .txt ze znacznikami stron [[s. N]] + _indeks.md
        (rodzaj dokumentu, data publikacji, paczka zmian, z ktorej pochodzi, OCR skanow)
  spis PLIK_TXT|KATALOG_TXT          naglowki rozdzialow / (Sub)klauzul / paragrafow z numerami stron
  strony PLIK_TXT --strony 9-11,15   tekst wybranych stron
  szukaj KATALOG_TXT|PLIK [--temat a,b] [--wzor REGEX] [--plik FRAGMENT_NAZWY] [--kontekst 350] [--max 10]
        fragmenty z numerem strony dla tematow analizy (lista: komenda tematy)
  pytania KATALOG_TXT [--wzor REGEX] [--max 60]
        pytania i odpowiedzi z wyjasnien jako osobne pozycje [Lp/zestaw/nr] ze strona, np. --wzor "most|obiekt"
  roznice STARY.txt NOWY.txt         zmiany miedzy wersjami dokumentu (zdaniami: BYLO / JEST)
  terminy --skladanie RRRR-MM-DD [--tryb podstawowy|nieograniczony] [--godzina 10:00] [--pierwotny RRRR-MM-DD]
        ostatni dzien na pytania do zamawiajacego wg Pzp i liczba dni do terminu
  rynek --repo mid-przetargi (--zamawiajacy TEKST | --slowa a,b) [--max 15]
        podobne postepowania z bazy wynikow mid-przetargi: budzet, ceny ofert, zwyciezca
  tematy                             lista tematow wyszukiwania
"""
import argparse, json, os, re, shutil, subprocess, sys, tempfile, unicodedata, zipfile
from datetime import date, datetime, timedelta
from pathlib import Path

EXT_TEKST = {".pdf", ".docx", ".doc", ".rtf", ".odt", ".xlsx", ".xlsm", ".xls", ".ods", ".txt"}
EXT_ARCH = {".zip", ".7z"}

# ----------------------------------------------------------------- normalizacja (dlugosc 1:1)
_PL = str.maketrans("ąćęłńóśźżĄĆĘŁŃÓŚŹŻ", "acelnoszzACELNOSZZ")


def norm(s: str) -> str:
    """Bez polskich znakow i malymi literami, z zachowaniem dlugosci (pozycje zgodne z oryginalem)."""
    s = s.translate(_PL)
    return "".join(c.lower() if len(c.lower()) == 1 else c for c in s)


def bezpieczna(n: str, maks: int = 120) -> str:
    n = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", n).strip(" .")
    if len(n) > maks:
        stem, ext = os.path.splitext(n)
        n = stem[: maks - len(ext)] + ext
    return n or "plik"


# ----------------------------------------------------------------- rodzaj dokumentu po nazwie
RODZAJE = [
    ("wyniki", r"otwarci\w* ofert|z otwarcia|kwot\w* (przeznaczon|jaka zamierza)|wybor\w* (najkorzystniejsz|oferty)|rozstrzygn|uniewazni"),
    ("wyjasnienia/zmiany", r"wyjasn|odpowiedz|pytan|zapytan|zmian|modyfikac|sprostow|informacja dla wykonawc"),
    ("SWZ/IDW", r"(^|[^a-z])s?i?wz([^a-z]|$)|specyfikacj\w* warunk|instrukcj\w* dla wykonaw|(^|[^a-z])idw"),
    ("umowa", r"umow|(^|[^a-z])pou([^a-z]|$)|istotn\w* postanowien|warunk\w* kontrakt|projektowane postanowien|(^|[^a-z])(swk|owk)([^a-z]|$)|dane kontraktowe|akt umowy|gwarancj\w* (jakosci|nalezyt|zwrotu)|wzor\w* gwarancji"),
    ("OPZ", r"(^|[^a-z])opz([^a-z]|$)|opis\w*[ _-]*przedmiot"),
    ("PFU", r"(^|[^a-z])pfu([^a-z]|$)|program\w*[ _-]*funkcjonaln"),
    ("ogloszenie", r"ogloszeni|notice"),
    ("formularze", r"formularz|jedz|oswiadczen|wykaz|zalacznik|zal\.? ?nr"),
]
PRIORYTET = {"wyjasnienia/zmiany": 0, "SWZ/IDW": 1, "umowa": 2, "OPZ": 3, "PFU": 3, "wyniki": 4, "ogloszenie": 5,
             "formularze": 6, "inne": 7}


def rodzaj(nazwa: str) -> str:
    n = norm(nazwa)
    for r, wz in RODZAJE:
        if re.search(wz, n):
            return r
    return "inne"


# ----------------------------------------------------------------- tematy wyszukiwania
TEMATY = {
    "termin_skladania": r"termin\w* skladania ofert|ofert\w* nalezy (zlozyc|skladac)|otwarci\w* ofert nastapi|skladania ofert uplywa",
    "zwiazanie": r"zwiazan\w* ofert",
    "termin_realizacji": r"termin\w* (realizacji|wykonania)|w terminie (do )?\d+ (dni|miesiecy|mies\.|tygodni)|harmonogram\w*|\betap\w* (i|ii|1|2)\b",
    "wadium": r"\bwadi(um|a|ach|ow)\b",
    "kryteria": r"kryteri\w* (oceny|wyboru)|waga kryterium|liczba punktow|maksymaln\w* liczb\w* punkt|\d+ ?pkt|\d+ ?%.{0,25}(cena|termin|doswiadcz|gwarancj|rekojm)",
    "warunki_udzialu": r"warunk\w* udzialu|zdolnosc\w* (techniczn|zawodow)|w okresie ostatnich (3|5|trzech|pieciu) lat|wykonal\w* (co najmniej|min)|wykonani\w* (co najmniej|min)",
    "osoby": r"uprawnieni\w* (budowlan|do projektowania)|specjalnosc\w* (inzynieryjn|drogow|mostow|konstrukcyjn)|bez ograniczen|osob\w* (skierowan|zdoln)|projektant\w*|kierownik\w* (projektu|zespolu)|koordynator\w*",
    "finanse_oc": r"ubezpieczon\w*|odpowiedzialnosci cywilnej|polis\w*|sytuacj\w* (ekonomiczn|finansow)|przychod\w*|zdolnosc\w* kredytow|srodk\w* finansow",
    "zakres": r"przedmiot\w* zamowienia|zakres\w* (prac|zamowienia|opracowania|dokumentacji)|opis\w* przedmiotu|w sklad dokumentacji|projekt\w* (budowlan|wykonawcz|techniczn)|koncepcj\w*|stwiorb|specyfikacj\w* techniczn|przedmiar\w*|kosztorys\w*",
    "obiekty": r"\bmost\w*|wiadukt\w*|kladk\w*|przepust\w*|estakad\w*|tunel\w*|przejsci\w* podziemn|rozpietos\w*|dlugos\w* (obiektu|calkowit|przesl)|przesl\w*|dzwigar\w*|podpor\w*",
    "decyzje": r"\bzrid\b|decyzj\w* (o zezwoleniu|o srodowiskow|srodowiskow|lokalizacyjn|o pozwoleniu)|pozwoleni\w* (na budowe|wodnoprawn|na rozbiorke)|zgloszeni\w* (robot|rozbiorki)|operat\w* wodnoprawn|raport\w* (o|oddzialywania)|uzgodnieni\w*|opini\w* (rdos|konserwator)",
    "badania": r"badani\w* (geotechn|podloza|geologiczn|betonu|materialow)|dokumentacj\w* (geologiczn|geotechn)|opini\w* geotechn|odwiert\w*|sondowan\w*|map\w* do celow projektowych|inwentaryzacj\w*|przeglad\w* (szczegolow|stanu|okresow)|ekspertyz\w*",
    "bim": r"\bbim\b|\bifc\b|model\w* (3d|informacyjn|bim)|\beir\b|\bbep\b|\bcde\b",
    "nadzor_autorski": r"nadzor\w* autorsk\w*|pobyt\w* na budowie|wizyt\w* na (budowie|terenie)|udzial\w* w (radach|naradach)",
    "wynagrodzenie": r"wynagrodzeni\w* (ryczaltow|kosztorysow|umown)|ryczalt\w*|platnos\w* (czesciow|koncow|za etap)|faktur\w* (czesciow|koncow|przejsciow)|termin\w* zaplaty|w terminie \d+ dni od (dnia )?(otrzymania|doreczenia|zlozenia)",
    "kary": r"kar\w* umown\w*|kare umown\w*|naliczy\w* kar|kara (za|w wysokosci)",
    "limit_kar": r"laczn\w* (maksymaln\w* )?(wysokos|wartos|kwot)\w* kar|kar\w*.{0,80}nie (moze|moga) przekroczyc|limit\w* kar",
    "opoznienie_zwloka": r"\bopoznieni\w*|\bzwlok\w*",
    "odpowiedzialnosc": r"ogranicz\w* odpowiedzialnos|odszkodowani\w* (przenoszac|uzupelniajac|na zasadach ogolnych)|w pelnej wysokosci|szkod\w* (rzeczywist|powstal)",
    "zabezpieczenie": r"zabezpieczeni\w* nalezyt\w*",
    "waloryzacja": r"waloryzac\w*|zmian\w* (wysokosci )?wynagrodzenia|art\.? ?439|wskaznik\w* (cen|gus|zmiany cen)",
    "zmiany_umowy": r"art\.? ?455|zmian\w* (postanowien )?umowy|zmian\w* terminu|aneks\w*",
    "prawa_autorskie": r"praw\w* autorsk\w*|pol\w* eksploatacji|praw\w* zalezn\w*|wykonywani\w* praw zaleznych|licencj\w*",
    "gwarancja_rekojmia": r"rekojm\w*|gwarancj\w*",
    "odstapienie": r"odstapi\w*|wypowiedz\w*|rozwiaza\w* umow",
    "podwykonawcy": r"podwykonaw\w*",
    "umowa_o_prace": r"umow\w* o prace|art\.? ?95",
    "wizja_lokalna": r"wizj\w* lokaln|zapoznani\w* sie z terenem|ogledzin\w*",
    "pytania": r"wyjasnieni\w* tresci (swz|specyfikacji)|wniosek o wyjasnienie|zapytani\w* do (tresci|swz)|art\.? ?(135|284)",
    "czesci": r"czesc\w* (nr )?(i|ii|iii|\d)\b|podzial\w* (zamowienia )?na czesci|ofert\w* czesciow\w*",
    "formalne": r"\bjedz\b|\besPD\b|oswiadczeni\w* (wykonawcy|o niepodleganiu|o spelnianiu)|przedmiotow\w* srodk\w* dowodow|podmiotow\w* srodk\w* dowodow|kwalifikowan\w* podpis|podpis\w* (zaufan|osobist)|pelnomocnictw\w*|formularz\w* (ofert|cenow)",
    "rodo": r"\brodo\b|przetwarzani\w* danych osobowych|powierzeni\w* przetwarzania",
    "ubezpieczenie_umowa": r"ubezpieczeni\w* (odpowiedzialnosci|oc)|polis\w* oc|sum\w* ubezpieczenia",
    "ai": r"sztuczn\w* inteligencj|system\w* ai\b|\bai act\b|2024/1689",
    "materialy_zewnetrzne": r"\bftp\b|serwer\w* (ftp|zamawiajacego)|dysk\w* (sieciow|zewnetrzn)|wetransfer|do pobrania (z|ze) (serwera|strony)|udostepni\w* na (serwerze|dysku)|login\w*:|haslo:",
}
OPIS_TEMATOW = {
    "termin_skladania": "termin i miejsce skladania/otwarcia ofert", "zwiazanie": "termin zwiazania oferta",
    "termin_realizacji": "termin wykonania, etapy, harmonogram", "wadium": "wadium (kwota, forma, termin)",
    "kryteria": "kryteria oceny i wagi, sposob liczenia punktow", "warunki_udzialu": "doswiadczenie wykonawcy",
    "osoby": "osoby, uprawnienia, doswiadczenie kadry", "finanse_oc": "sytuacja finansowa, OC jako warunek",
    "zakres": "przedmiot i zakres, fazy dokumentacji", "obiekty": "obiekty inzynierskie i ich parametry",
    "decyzje": "decyzje, pozwolenia, uzgodnienia", "badania": "badania, geotechnika, mapy, inwentaryzacje, ekspertyzy",
    "bim": "wymagania BIM", "nadzor_autorski": "nadzor autorski (zakres, liczba pobytow)",
    "wynagrodzenie": "rodzaj wynagrodzenia i platnosci", "kary": "kary umowne", "limit_kar": "laczny limit kar",
    "opoznienie_zwloka": "opoznienie vs zwloka (art. 433 pkt 1)", "odpowiedzialnosc": "zakres odpowiedzialnosci, odszkodowania",
    "zabezpieczenie": "zabezpieczenie nalezytego wykonania", "waloryzacja": "waloryzacja wynagrodzenia (art. 439)",
    "zmiany_umowy": "zmiany umowy (art. 455)", "prawa_autorskie": "prawa autorskie, pola eksploatacji, prawa zalezne",
    "gwarancja_rekojmia": "rekojmia i gwarancja", "odstapienie": "odstapienie, wypowiedzenie",
    "podwykonawcy": "podwykonawstwo", "umowa_o_prace": "wymog zatrudnienia na umowe o prace",
    "wizja_lokalna": "wizja lokalna", "pytania": "wyjasnienia tresci SWZ, termin na pytania",
    "czesci": "podzial na czesci", "formalne": "dokumenty skladane z oferta, podpis",
    "rodo": "RODO, powierzenie danych", "ubezpieczenie_umowa": "OC wymagane umowa",
    "ai": "zakaz/ograniczenia uzycia AI w utworach (prawa autorskie)",
    "materialy_zewnetrzne": "dokumenty poza platforma (FTP, dyski) - do recznego pobrania",
}


# ----------------------------------------------------------------- pobieranie z platform
def _sesja():
    import requests
    s = requests.Session()
    s.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/126.0 Safari/537.36",
                      "Accept-Language": "pl-PL,pl;q=0.9"})
    return s


def _lista_wbudowana(url: str) -> dict:
    """Minimalna obsluga e-Zamowien i platformazakupowa.pl (gdy brak repo mid-przetargi)."""
    from urllib.parse import urljoin, urlparse, unquote
    s = _sesja()
    wyn = {"platforma": None, "dokumenty": [], "uwagi": [], "opis": ""}
    host = urlparse(url).netloc.lower()
    if host == "ezamowienia.gov.pl":
        wyn["platforma"] = "ezamowienia"
        m = re.search(r"(ocds-[0-9a-z\-]+)", url)
        if not m:
            wyn["uwagi"].append("brak identyfikatora ocds w adresie")
            return wyn
        tid, api = m.group(1), "https://ezamowienia.gov.pl/mp-readmodels/api"
        r = s.get(f"{api}/Search/GetTender", params={"id": tid}, headers={"Accept": "application/json"}, timeout=60)
        if r.ok:
            t = r.json()
            wyn["opis"] = "\n".join(f"{k}: {t.get(k)}" for k in ("title", "organizationName", "referenceNumber",
                                                                  "bzpNumber", "submissionDate", "openDate") if t.get(k))
        rd = s.get(f"{api}/Search/GetTenderDocuments", params={"tenderId": tid}, headers={"Accept": "application/json"}, timeout=60)
        for d in (rd.json() if rd.ok else []) or []:
            if d.get("deleteDate") or d.get("tenderDocumentState") not in (None, "Published"):
                continue
            u = f"{api}/Tender/DownloadDocument/{tid}/{d.get('objectId')}"
            wyn["dokumenty"].append({"id": u, "nazwa": d.get("fileName") or d.get("name") or d.get("objectId"),
                                     "data": (d.get("publishedDate") or d.get("createDate") or "")[:10] or None,
                                     "rozmiar": None, "pobierz": {"metoda": "get", "url": u}})
    elif host.endswith("platformazakupowa.pl") and "plk-sa" not in host:
        from bs4 import BeautifulSoup
        wyn["platforma"] = "platformazakupowa"
        r = s.get(url, headers={"X-Requested-With": "XMLHttpRequest"}, timeout=60)
        soup = BeautifulSoup(r.text, "lxml")
        seen = set()
        for a in soup.select('a[href*="file/get_new"]'):
            href = urljoin(url, a["href"])
            if href in seen:
                continue
            seen.add(href)
            tr = a.find_parent("tr")
            td = tr.find_all("td") if tr else []
            nazwa = re.sub(r"\s+", " ", td[0].get_text(" ")).strip() if td else a.get_text(" ").strip()
            nazwa = nazwa or unquote(href.rsplit("/", 1)[-1])
            data = None
            if tr:
                m = re.search(r"(\d{2})[.\-](\d{2})[.\-](\d{4})", tr.get_text(" "))
                data = f"{m.group(3)}-{m.group(2)}-{m.group(1)}" if m else None
            wyn["dokumenty"].append({"id": href, "nazwa": nazwa, "data": data, "rozmiar": None,
                                     "pobierz": {"metoda": "get", "url": href}})
        wyn["opis"] = soup.get_text(" ")[:20000]
    else:
        wyn["uwagi"].append("platforma bez wbudowanej obslugi - pobierz dokumenty recznie albo uzyj --repo mid-przetargi")
    return wyn


def _pobierz_wbudowane(d: dict, cel: Path, limit_b: int) -> dict:
    s = _sesja()
    r = s.get(d["pobierz"]["url"], stream=True, timeout=300)
    if r.status_code != 200:
        raise RuntimeError(f"HTTP {r.status_code}")
    n = 0
    with open(cel, "wb") as f:
        for k in r.iter_content(1 << 18):
            n += len(k)
            if n > limit_b:
                f.close(); cel.unlink(missing_ok=True)
                raise RuntimeError(f"plik > {limit_b >> 20} MB")
            f.write(k)
    cd = r.headers.get("content-disposition", "")
    m = re.search(r"filename\*=UTF-8''([^;]+)|filename=\"?([^\";]+)", cd)
    from urllib.parse import unquote
    return {"rozmiar": n, "nazwa_serwera": unquote((m.group(1) or m.group(2))) if m else None}


def _zip_nazwa(info: zipfile.ZipInfo) -> str:
    if info.flag_bits & 0x800:
        return info.filename
    try:
        return info.filename.encode("cp437").decode("cp852")
    except (UnicodeEncodeError, UnicodeDecodeError):
        return info.filename


def rozpakuj(arch: Path, cel: Path, maks_b: int) -> list:
    out = []
    cel.mkdir(parents=True, exist_ok=True)
    try:
        if arch.suffix.lower() == ".zip":
            with zipfile.ZipFile(arch) as z:
                for i in z.infolist():
                    if i.is_dir() or i.file_size > maks_b:
                        continue
                    rel = Path(*[bezpieczna(p) for p in Path(_zip_nazwa(i)).parts if p not in ("..", "/")])
                    dst = cel / rel
                    dst.parent.mkdir(parents=True, exist_ok=True)
                    with z.open(i) as src, open(dst, "wb") as f:
                        shutil.copyfileobj(src, f)
                    out.append(dst)
        elif arch.suffix.lower() == ".7z":
            try:
                import py7zr
            except ImportError:
                subprocess.run([sys.executable, "-m", "pip", "install", "-q", "--break-system-packages", "py7zr"], check=False)
                import py7zr
            with py7zr.SevenZipFile(arch, "r") as z:
                z.extractall(path=cel)
            out = [p for p in cel.rglob("*") if p.is_file()]
    except Exception as e:  # noqa: BLE001
        print(f"  [uwaga] archiwum {arch.name}: {type(e).__name__} {str(e)[:120]}", file=sys.stderr)
    # archiwa w archiwach
    for p in list(out):
        if p.suffix.lower() in EXT_ARCH:
            out += rozpakuj(p, p.with_suffix("") if p.with_suffix("") != p else p.parent / (p.stem + "_x"), maks_b)
    return out


def cmd_pobierz(a):
    out = Path(a.out); pl = out / "pliki"; pl.mkdir(parents=True, exist_ok=True)
    repo = Path(a.repo) if a.repo else None
    info = {}
    url = a.url
    if a.nr:
        if not repo:
            sys.exit("--nr wymaga --repo (sciezka do klonu mid-przetargi)")
        d = json.load(open(repo / "przetargi_active.json", encoding="utf-8"))
        t = d.get("tenders", d)
        hit = [v for v in t.values() if isinstance(v, dict) and str(v.get("nr")) == str(a.nr)]
        if not hit:
            sys.exit(f"Brak przetargu nr {a.nr} w przetargi_active.json")
        info = {k: hit[0].get(k) for k in ("nr", "tytul", "zamawiajacy", "termin", "termin_godzina", "url",
                                           "nr_ref", "wartosc", "status_mid", "uwagi")}
        url = hit[0].get("url")
        if not str(url).startswith("http"):
            sys.exit(f"Przetarg nr {a.nr} nie ma strony postepowania (zapytanie mailowe?): {url}")
    T = None
    if repo and (repo / "chmura" / "teczka_pliki.py").exists():
        sys.path.insert(0, str(repo / "chmura"))
        import teczka_pliki as T  # noqa: N812
    if T:
        se = T.Sesja()
        wyn = T.lista(se, url)
    else:
        wyn = _lista_wbudowana(url)
    dok = wyn.get("dokumenty", [])
    limit_b = int(a.max_mb * 1024 * 1024)
    dok.sort(key=lambda d: (PRIORYTET.get(rodzaj(d["nazwa"]), 7), d.get("rozmiar") or 0))
    lista = []
    for i, d in enumerate(dok):
        poz = {"nazwa": d["nazwa"], "data": d.get("data"), "rozmiar": d.get("rozmiar"), "rodzaj": rodzaj(d["nazwa"]),
               "folder_platformy": d.get("folder_platformy"), "status": None}
        if i >= a.max_dok:
            poz["status"] = f"pominiety (limit {a.max_dok} dokumentow)"
        elif d.get("rozmiar") and d["rozmiar"] > limit_b:
            poz["status"] = f"pominiety ({d['rozmiar'] >> 20} MB > {a.max_mb} MB)"
        else:
            cel = pl / bezpieczna(d["nazwa"])
            try:
                if T:
                    r = T.pobierz(se, d, cel, limit=limit_b)
                else:
                    r = _pobierz_wbudowane(d, cel, limit_b)
                if r.get("nazwa_serwera") and not Path(d["nazwa"]).suffix:
                    nowy = cel.with_name(bezpieczna(d["nazwa"] + Path(r["nazwa_serwera"]).suffix))
                    cel.rename(nowy); cel = nowy
                poz.update(status="pobrany", plik=str(cel.relative_to(out)), rozmiar=r.get("rozmiar"))
                if cel.suffix.lower() in EXT_ARCH:
                    rozp = rozpakuj(cel, cel.parent / (cel.stem + "_rozpakowane"), limit_b)
                    poz["rozpakowane"] = len(rozp)
            except Exception as e:  # noqa: BLE001
                poz["status"] = f"blad: {str(e)[:150]}"
        lista.append(poz)
        print(f"  {poz['status'][:40]:40s} {poz['rodzaj']:20s} {poz['nazwa'][:90]}")
    meta = {"url": url, "platforma": wyn.get("platforma"), "przetarg": info, "uwagi": wyn.get("uwagi", []),
            "pobrano": datetime.now().isoformat(timespec="minutes"), "dokumenty": lista}
    json.dump(meta, open(out / "_lista.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    (out / "_opis_postepowania.txt").write_text(wyn.get("opis") or "", encoding="utf-8")
    ok = sum(1 for p in lista if p["status"] == "pobrany")
    print(f"POBRANO={ok} POMINIETO={len(lista) - ok} PLATFORMA={wyn.get('platforma')} -> {out}/_lista.json")


# ----------------------------------------------------------------- ekstrakcja tekstu
_TESS = None


def _tessdata() -> tuple:
    """(katalog tessdata, jezyk) - raz na uruchomienie; dociaga pol.traineddata z GitHuba, gdy brak."""
    global _TESS
    if _TESS:
        return _TESS
    for d in ("/usr/share/tesseract-ocr/5/tessdata", "/usr/share/tesseract-ocr/4.00/tessdata"):
        if Path(d, "pol.traineddata").exists():
            _TESS = (d, "pol")
            return _TESS
    own = Path.home() / ".local/share/tessdata"
    if not (own / "pol.traineddata").exists():
        own.mkdir(parents=True, exist_ok=True)
        try:
            import requests
            r = requests.get("https://raw.githubusercontent.com/tesseract-ocr/tessdata_fast/main/pol.traineddata", timeout=60)
            if r.ok and len(r.content) > 1_000_000:
                (own / "pol.traineddata").write_bytes(r.content)
        except Exception:  # noqa: BLE001
            pass
    _TESS = (str(own), "pol") if (own / "pol.traineddata").exists() else (None, "eng")
    if _TESS[1] == "eng":
        print("  [uwaga] brak polskiego modelu OCR - rozpoznawanie modelem angielskim (gorsze polskie znaki)", file=sys.stderr)
    return _TESS


def _ocr_strona(pdf: Path, nr: int, klucz: str = "") -> str:
    cache = Path.home() / ".cache" / "swz_ocr" / f"{klucz}_{nr}.txt" if klucz else None
    if cache and cache.exists():
        return cache.read_text(encoding="utf-8")
    t = _ocr_strona_bez_cache(pdf, nr)
    if cache:
        cache.parent.mkdir(parents=True, exist_ok=True)
        cache.write_text(t, encoding="utf-8")
    return t


def _ocr_strona_bez_cache(pdf: Path, nr: int) -> str:
    td, lang = _tessdata()
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run(["pdftoppm", "-r", "200", "-gray", "-f", str(nr), "-l", str(nr), "-png", str(pdf), f"{tmp}/p"],
                       capture_output=True, timeout=120)
        png = sorted(Path(tmp).glob("p*.png"))
        if not png:
            return ""
        cmd = ["tesseract", str(png[0]), "-", "-l", lang, "--psm", "3"]
        if td:
            cmd += ["--tessdata-dir", td]
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
        return r.stdout


def tekst_pdf(p: Path, ocr: str, maks_ocr: int, tabele: bool = False) -> tuple:
    r = subprocess.run(["pdftotext"] + (["-layout"] if tabele else []) + [str(p), "-"], capture_output=True, timeout=600)
    strony = r.stdout.decode("utf-8", errors="replace").split("\f")
    if strony and not strony[-1].strip():
        strony = strony[:-1]
    if not strony:  # pdftotext zawiodl - pdfplumber
        try:
            import pdfplumber
            with pdfplumber.open(str(p)) as pdf:
                strony = [pg.extract_text() or "" for pg in pdf.pages]
        except Exception:  # noqa: BLE001
            strony = []
    puste = [i for i, s in enumerate(strony) if len(s.strip()) < 40]
    ocr_str = 0
    if strony and puste and (ocr == "tak" or (ocr == "auto" and len(puste) >= max(1, len(strony) // 3))):
        from concurrent.futures import ThreadPoolExecutor
        _tessdata()
        do_ocr = puste[:maks_ocr]
        print(f"    OCR {len(do_ocr)} stron skanu: {p.name[:60]} ...", flush=True)
        with ThreadPoolExecutor(max_workers=max(2, os.cpu_count() or 2)) as ex:
            import hashlib
            klucz = hashlib.md5(p.read_bytes()).hexdigest()
            for i, t in zip(do_ocr, ex.map(lambda i: _ocr_strona(p, i + 1, klucz), do_ocr)):
                strony[i] = "[OCR] " + t
                ocr_str += 1
    uwaga = ""
    if len(puste) > ocr_str:
        uwaga = f"{len(puste) - ocr_str} stron bez warstwy tekstowej (skan/rysunek) bez OCR"
    return strony, ocr_str, uwaga


def _do_pdf(p: Path, tmp: Path) -> Path | None:
    r = subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(tmp), str(p)],
                       capture_output=True, timeout=300)
    wyn = tmp / (p.stem + ".pdf")
    return wyn if wyn.exists() else None


def tekst_pliku(p: Path, ocr: str, maks_ocr: int, tabele: bool = False) -> tuple:
    """(lista stron, liczba stron OCR, uwaga). Dokumenty biurowe -> PDF przez LibreOffice, zeby miec numery stron."""
    ext = p.suffix.lower()
    if ext == ".pdf":
        return tekst_pdf(p, ocr, maks_ocr, tabele)
    if ext in (".docx", ".doc", ".rtf", ".odt"):
        with tempfile.TemporaryDirectory() as tmp:
            pdf = _do_pdf(p, Path(tmp)) if shutil.which("soffice") else None
            if pdf:  # dokument biurowy - bez OCR (skany sa w PDF-ach)
                return tekst_pdf(pdf, "nie", 0, tabele)
        if ext == ".docx":
            import docx
            dk = docx.Document(str(p))
            t = "\n".join(x.text for x in dk.paragraphs)
            for tb in dk.tables:
                for w in tb.rows:
                    t += "\n" + " | ".join(c.text.strip() for c in w.cells)
            return [t], 0, "bez numerow stron (brak LibreOffice)"
        return [""], 0, "nie udalo sie odczytac"
    if ext in (".xlsx", ".xlsm", ".xls", ".ods"):
        src = p
        tmpd = None
        if ext in (".xls", ".ods"):
            tmpd = tempfile.mkdtemp()
            subprocess.run(["soffice", "--headless", "--convert-to", "xlsx", "--outdir", tmpd, str(p)], capture_output=True, timeout=300)
            src = Path(tmpd) / (p.stem + ".xlsx")
        import openpyxl
        wb = openpyxl.load_workbook(src, data_only=True, read_only=True)
        strony = []
        for ws in wb.worksheets:
            wiersze = []
            for row in ws.iter_rows(values_only=True):
                v = [str(c).strip() for c in row if c not in (None, "")]
                if v:
                    wiersze.append(" | ".join(v))
            strony.append(f"(arkusz: {ws.title})\n" + "\n".join(wiersze))
        if tmpd:
            shutil.rmtree(tmpd, ignore_errors=True)
        return strony, 0, ""
    if ext == ".txt":
        return [p.read_text(encoding="utf-8", errors="replace")], 0, ""
    return [], 0, "format nieobslugiwany"


def kompaktuj(strony: list, tabele: bool = False) -> list:
    """Mniej tokenow do czytania: bez naglowkow/stopek powtarzanych na stronach, bez numeracji 'Strona X z Y',
    bez wciec i pustych linii; kolumny tabel (3+ spacje) zostaja jako ' | '."""
    from collections import Counter
    if len(strony) >= 3:
        def brzegi(s):  # naglowki i stopki: 3 pierwsze i 3 ostatnie niepuste linie strony
            w = [x.strip() for x in s.splitlines() if x.strip()]
            return set(w[:3] + w[-3:])
        ile = Counter(x for s in strony for x in brzegi(s))
        stopka = {x for x, n in ile.items() if n >= max(3, len(strony) // 2) and len(x) < 300}
    else:
        stopka = set()
    out = []
    for s in strony:
        w = []
        for x in s.splitlines():
            x = x.strip()
            if not x or x in stopka or re.fullmatch(r"(Strona|str\.?)\s*\d+\s*(z|/|of)\s*\d+|\d+\s*/\s*\d+|- ?\d+ ?-", x, re.I):
                continue
            w.append(re.sub(r" {3,}", " | ", x) if tabele else re.sub(r"\s{2,}", " ", x))
        out.append("\n".join(w))
    return out


def cmd_teksty(a):
    out = Path(a.out); out.mkdir(parents=True, exist_ok=True)
    lista, zrodla = {}, {}  # plik z platformy -> wpis _lista.json; katalog rozpakowanego archiwum -> wpis
    pliki = []
    for w in a.wejscie:
        w = Path(w)
        if w.is_dir():
            lj = w / "_lista.json"
            if lj.exists():
                for d in json.load(open(lj, encoding="utf-8")).get("dokumenty", []):
                    if d.get("plik"):
                        pp = (w / d["plik"]).resolve()
                        lista[pp] = d
                        zrodla[pp.parent / (pp.stem + "_rozpakowane")] = d
            # pliki zaczynajace sie od "_" to metadane narzedzia (_lista.json, _opis_postepowania.txt)
            # ...a katalogi z _indeks.md to wczesniejsze wyniki tej komendy - nie czytamy ich jak dokumentow
            pliki += [p for p in sorted(w.rglob("*")) if p.is_file() and not p.name.startswith("_")
                      and not (p.parent / "_indeks.md").exists() and out.resolve() not in p.resolve().parents]
        elif w.is_file():
            pliki.append(w)

    def pochodzenie(p):
        """(wpis z _lista.json, nazwa archiwum albo None) - data publikacji i paczka, z ktorej plik pochodzi."""
        rp = p.resolve()
        if rp in lista:
            return lista[rp], None
        for q in rp.parents:
            if q in zrodla:
                return zrodla[q], zrodla[q]["nazwa"]
        return {}, None
    rozpak = []
    for p in pliki:
        if p.suffix.lower() in EXT_ARCH and not any(q.name.startswith(p.stem + "_rozpakowane") for q in p.parent.iterdir()):
            rozpak += rozpakuj(p, p.parent / (p.stem + "_rozpakowane"), 300 << 20)
    pliki = [p for p in pliki + rozpak if p.suffix.lower() in EXT_TEKST]
    pliki = list(dict.fromkeys(pliki))

    def nasz_wynik(p):  # .txt wygenerowany wczesniej przez te komende (naglowek '### nazwa | rodzaj: ...')
        if p.suffix.lower() != ".txt":
            return False
        with open(p, encoding="utf-8", errors="replace") as f:
            return bool(re.match(r"### .+ \| rodzaj: ", f.readline()))
    pliki = [p for p in pliki if not nasz_wynik(p)]
    import hashlib
    unik, dup = {}, []
    for p in pliki:  # identyczne pliki (np. z platformy i z archiwum) czytamy raz
        h = hashlib.md5(p.read_bytes()).hexdigest()
        if h in unik:
            dup.append((p, unik[h]))
        else:
            unik[h] = p
    pliki = list(unik.values())
    for p, q in dup:
        print(f"  pominiety duplikat: {p.name} (= {q.name})", flush=True)
    wiersze, teksty_hash, uzyte = [], {}, set()
    for p in pliki:
        rodz = rodzaj(p.name)
        try:
            # tabele (pytanie | odpowiedz, przed | po zmianie, wyniki otwarcia): uklad kolumn zachowany jako " | "
            tabele = rodz in ("wyniki", "formularze", "wyjasnienia/zmiany")
            strony, n_ocr, uwaga = tekst_pliku(p, a.ocr, a.maks_ocr, tabele)
        except Exception as e:  # noqa: BLE001
            strony, n_ocr, uwaga = [], 0, f"blad: {type(e).__name__} {str(e)[:100]}"
        strony = kompaktuj(strony, tabele)
        txt = "".join(f"\n[[s. {i + 1}]]\n{s}" for i, s in enumerate(strony))
        ht = hashlib.md5(re.sub(r"\s+", "", txt).encode()).hexdigest()
        if ht in teksty_hash and len(txt) > 200:
            print(f"  pominiety (ta sama tresc co {teksty_hash[ht]}): {p.name}", flush=True)
            continue
        teksty_hash[ht] = p.name
        meta, arch = pochodzenie(p)
        data = meta.get("data") or ""
        if arch and rodzaj(arch) == "wyjasnienia/zmiany" and rodz not in ("wyjasnienia/zmiany", "wyniki"):
            # np. PFU podmienione razem z wyjasnieniami - ta wersja zastepuje dokument z pierwotnej SWZ
            uwaga = (uwaga + "; " if uwaga else "") + f"DODANY/ZMIENIONY w paczce: {arch}"
        nazwa_txt = bezpieczna(f"{PRIORYTET.get(rodz, 7)}_{p.stem}", 100)
        if nazwa_txt + ".txt" in uzyte:  # ta sama nazwa w roznych paczkach (np. kolejne 'Tresc zapytan...')
            nazwa_txt += f"_{data or len(uzyte)}"
        nazwa_txt += ".txt"
        uzyte.add(nazwa_txt)
        zr = f" | z paczki: {arch}" if arch else ""
        (out / nazwa_txt).write_text(f"### {p.name} | rodzaj: {rodz} | data: {data or '?'}{zr}\n{txt}", encoding="utf-8")
        wiersze.append((PRIORYTET.get(rodz, 7), data, p.name, rodz, len(strony), sum(len(s) for s in strony), n_ocr,
                        uwaga, nazwa_txt))
        print(f"  {rodz:20s} {len(strony):4d} s. {sum(len(s) for s in strony):8d} zn.  {p.name[:80]}" + (f"  [{uwaga}]" if uwaga else ""), flush=True)
    # ten sam dokument opublikowany ponownie (np. poprawiony formularz) - zaznacz, ktora wersja obowiazuje
    from collections import defaultdict
    wersje = defaultdict(list)
    for i, w in enumerate(wiersze):
        wersje[w[2]].append(i)
    for nazwa, idx in wersje.items():
        daty = sorted({wiersze[i][1] for i in idx if wiersze[i][1]})
        if len(idx) > 1 and len(daty) > 1:
            for i in idx:
                w = list(wiersze[i])
                dop = (f"NOWSZA WERSJA ({w[1]})" if w[1] == daty[-1] else f"zastapiona wersja z {daty[-1]}")
                w[7] = (w[7] + "; " if w[7] else "") + dop
                wiersze[i] = tuple(w)
    wiersze.sort()
    md = ["| rodzaj | dokument | data | stron | znakow | OCR | uwagi | plik txt |", "|---|---|---|---|---|---|---|---|"]
    md += [f"| {r} | {n} | {d} | {s} | {z} | {o or ''} | {u} | {t} |" for _, d, n, r, s, z, o, u, t in wiersze]
    (out / "_indeks.md").write_text("\n".join(md) + "\n", encoding="utf-8")
    print(f"DOKUMENTOW={len(wiersze)} -> {out}/_indeks.md")


# ----------------------------------------------------------------- wyszukiwanie
def _strona(znaczniki, poz):
    nr = 1
    for p, n in znaczniki:
        if p > poz:
            break
        nr = n
    return nr


def cmd_szukaj(a):
    katalog = Path(a.katalog)
    tematy = {}
    if a.wzor:
        tematy["wzor"] = a.wzor
    if a.temat:
        for t in a.temat.split(","):
            t = t.strip()
            if t not in TEMATY:
                sys.exit(f"Nieznany temat: {t}. Lista: {', '.join(TEMATY)}")
            tematy[t] = TEMATY[t]
    if not tematy:
        tematy = dict(TEMATY)
    pliki = sorted(katalog.glob("*.txt")) if katalog.is_dir() else [katalog]
    if a.plik:  # tylko wybrane dokumenty, np. --plik "PFU - DK60" (fragment nazwy, bez polskich znakow tez dziala)
        pliki = [p for p in pliki if norm(a.plik) in norm(p.name)]
    teksty = []
    for p in pliki:
        t = p.read_text(encoding="utf-8", errors="replace")
        zn = [(m.start(), int(m.group(1))) for m in re.finditer(r"\[\[s\. (\d+)\]\]", t)]
        teksty.append((p, t, norm(t), zn))
    for tem, wz in tematy.items():
        rx = re.compile(norm(wz) if tem != "wzor" else wz, re.I)
        print(f"\n## {tem}" + (f" - {OPIS_TEMATOW.get(tem, '')}" if tem in OPIS_TEMATOW else ""))
        trafienia = 0
        for p, t, tn, zn in teksty:
            nag = t.split("\n", 1)[0].lstrip("# ").split(" | ")
            nazwa = nag[0]
            dt = next((x[6:] for x in nag if x.startswith("data: ") and x[6:] != "?"), "")
            if dt and rodzaj(nazwa) == "wyjasnienia/zmiany":
                nazwa += f" z {dt}"  # kolejne odpowiedzi czesto maja te sama nazwe pliku
            koniec = -1  # nie pokazuj trafien, ktore mieszcza sie w poprzednim fragmencie
            for m in rx.finditer(tn if tem != "wzor" else t):
                if m.start() < koniec:
                    continue
                s = _strona(zn, m.start())
                a0, a1 = max(0, m.start() - a.kontekst), min(len(t), m.end() + a.kontekst)
                koniec = a1
                frag = re.sub(r"\[\[s\. \d+\]\]", " ", t[a0:a1])
                frag = re.sub(r"[ \t]+", " ", re.sub(r"\n\s*\n+", "\n", frag)).strip()
                print(f"\n**{nazwa}, s. {s}**\n> " + frag.replace("\n", "\n> "))
                trafienia += 1
                if trafienia >= a.max:
                    break
            if trafienia >= a.max:
                print(f"\n_(limit {a.max} trafien - zawez: --temat {tem} --max 30 albo --wzor)_")
                break
        if not trafienia:
            print("_brak trafien w dokumentach_")


RE_WIERSZ_PYT = re.compile(r"^(\d{1,4})((?: \| \d{1,3}){1,2}) \| (.*)")  # tabela: Lp | zestaw | nr | pytanie | odpowiedz
RE_PYTANIE = re.compile(r"^(?:Zapytanie|ZAPYTANIE|Pytanie|PYTANIE|Pyt\.)\s*(?:nr\.?\s*)?(\d{1,4})\b[:.)\-]*\s*(.*)")


def cmd_pytania(a):
    """Pytania i odpowiedzi z wyjasnien SWZ jako osobne pozycje (z numerem i strona), z filtrem --wzor."""
    katalog = Path(a.katalog)
    pliki = sorted(katalog.glob("*.txt")) if katalog.is_dir() else [katalog]
    rx = re.compile(norm(a.wzor), re.I) if a.wzor else None
    razem = 0
    for p in pliki:
        t = p.read_text(encoding="utf-8", errors="replace")
        nag = t.split("\n", 1)[0].lstrip("# ").split(" | ")
        if katalog.is_dir() and not any(x == "rodzaj: wyjasnienia/zmiany" for x in nag):
            continue
        dt = next((x[6:] for x in nag if x.startswith("data: ") and x[6:] != "?"), "")
        poz, cur, s = [], None, 1
        for linia in t.splitlines()[1:]:
            m = re.match(r"\[\[s\. (\d+)\]\]", linia)
            if m:
                s = int(m.group(1))
                continue
            m = RE_WIERSZ_PYT.match(linia) or RE_PYTANIE.match(linia)
            if m:
                nr = m.group(1) + (m.group(2).replace(" | ", "/") if m.re is RE_WIERSZ_PYT else "")
                cur = [nr, s, m.groups()[-1]]
                poz.append(cur)
            elif cur:
                cur[2] += " " + linia.strip()
        traf = [x for x in poz if not rx or rx.search(norm(x[2]))]
        if not poz or not traf:
            continue
        print(f"\n## {nag[0]}" + (f" z {dt}" if dt else "") + f" - pozycji {len(poz)}, pasuje {len(traf)}")
        for nr, s, tx in traf[: a.max]:
            tx = re.sub(r"\s+", " ", tx).strip()
            print(f"- [{nr}] s. {s}: {tx[: a.dlugosc]}" + ("..." if len(tx) > a.dlugosc else ""))
        razem += len(traf)
    print(f"\nPOZYCJI={razem}" + ("" if razem else " (nie rozpoznano tabeli pytan - uzyj: szukaj --plik ... --wzor ...)"))


RE_NAGLOWEK = re.compile(
    r"^(?:(?:rozdzia[lł]|cz[eę][sś][cć]|dzia[lł])\s+[IVXLC\d]+\b.*"
    r"|(?:Cz[eę][sś][cć]|CZĘŚĆ)\s+[A-Z]\s*[—–-].*"
    r"|§\s*\d+.*"
    r"|(?:Sub)?[Kk]lauzula\s+\d+(?:\.\d+)*\.?\s+\S.*"
    r"|\d{1,2}(?:\.\d{1,2}){0,1}\.?\s+[A-ZĄĆĘŁŃÓŚŹŻ][A-ZĄĆĘŁŃÓŚŹŻ ,\-–()/„”\"]{8,}.*"
    r"|\d{1,2}\.\s+[A-ZĄĆĘŁŃÓŚŹŻ][a-ząćęłńóśźż]+(?:\s+[^\s.]+){0,9}"
    r"|[A-H]\.\d{1,2}\.?\s+[A-ZĄĆĘŁŃÓŚŹŻ].{3,80}"
    r"|[IVX]{1,4}\.\s+[A-ZĄĆĘŁŃÓŚŹŻ].{5,})$")


def cmd_spis(a):
    """Spis rozdzialow (naglowki numerowane wielkimi literami, §, Rozdzial) z numerami stron."""
    for p in ([Path(a.plik)] if Path(a.plik).is_file() else sorted(Path(a.plik).glob("*.txt"))):
        t = p.read_text(encoding="utf-8", errors="replace")
        print(f"\n## {t.split(chr(10), 1)[0].lstrip('# ')}")
        s = 1
        linie = t.splitlines()
        for i, x in enumerate(linie):
            m = re.match(r"\[\[s\. (\d+)\]\]", x)
            if m:
                s = int(m.group(1))
                continue
            x2 = x.split(" | ")[0].strip()
            # numer w osobnej linii, tytul wielkimi literami w nastepnej (np. IDW GDDKiA: "7." / "TERMIN REALIZACJI")
            if re.fullmatch(r"\d{1,2}\.?", x2) and i + 1 < len(linie):
                nast = linie[i + 1].strip()
                if re.fullmatch(r"[A-ZĄĆĘŁŃÓŚŹŻ][A-ZĄĆĘŁŃÓŚŹŻ0-9 ,\-–()/„”\".]{5,}", nast):
                    print(f"s. {s:>3}  {x2.rstrip('.')}. {nast}")
                continue
            if RE_NAGLOWEK.match(x2) and len(x2) < 160:
                print(f"s. {s:>3}  {x2}")


def cmd_strony(a):
    """Tekst wybranych stron dokumentu, np. --strony 9-11,15."""
    t = Path(a.plik).read_text(encoding="utf-8", errors="replace")
    czesci = re.split(r"\n\[\[s\. (\d+)\]\]\n", t)
    strony = {int(czesci[i]): czesci[i + 1] for i in range(1, len(czesci) - 1, 2)}
    wyb = set()
    for z in a.strony.split(","):
        if "-" in z:
            x, y = z.split("-")
            wyb |= set(range(int(x), int(y) + 1))
        else:
            wyb.add(int(z))
    print(czesci[0].strip())
    for n in sorted(wyb):
        if n in strony:
            print(f"\n[[s. {n}]]\n{strony[n]}")


def cmd_roznice(a):
    """Zmiany miedzy dwiema wersjami dokumentu (np. wzor umowy przed i po wyjasnieniach) - akapitami."""
    import difflib

    def akapity(p):
        t = Path(p).read_text(encoding="utf-8", errors="replace").split("\n", 1)[-1]
        strony = re.split(r"\[\[s\. \d+\]\]", t)
        from collections import Counter
        ile = Counter(x.strip() for s in strony for x in set(s.splitlines()) if x.strip())
        stopka = {x for x, n in ile.items() if n >= max(3, len(strony) // 2)}  # naglowki/stopki stron
        t = "\n".join(x for x in t.splitlines() if x.strip() not in stopka)
        t = re.sub(r"\[\[s\. \d+\]\]|Strona \d+ z \d+", " ", t)
        t = re.sub(r"\s+", " ", t)  # niezalezne od lamania wierszy i stron
        zd = re.split(r"(?<=[.;:])\s+(?=[A-ZĄĆĘŁŃÓŚŹŻ§0-9(]|[a-z]\))", t)
        return [x.strip() for x in zd if len(x.strip()) > 3]

    A, B = akapity(a.stary), akapity(a.nowy)
    sm = difflib.SequenceMatcher(a=[norm(x) for x in A], b=[norm(x) for x in B], autojunk=False)
    zmian = 0
    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op == "equal":
            continue
        zmian += 1
        print(f"\n### zmiana {zmian} ({op})")
        for x in A[i1:i2]:
            print("- BYLO: " + x[:600].replace("\n", " "))
        for x in B[j1:j2]:
            print("+ JEST: " + x[:600].replace("\n", " "))
    print(f"\nZMIAN={zmian}")


def cmd_tematy(a):
    for k, v in OPIS_TEMATOW.items():
        print(f"{k:22s} {v}")


# ----------------------------------------------------------------- terminy
def cmd_terminy(a):
    d = date.fromisoformat(a.skladanie)
    dzis = date.fromisoformat(a.dzis) if a.dzis else date.today()
    print(f"Termin skladania: {d} {a.godzina or ''} - za {(d - dzis).days} dni (dzis {dzis})")
    if a.tryb == "podstawowy":
        wn, odp, pod = 4, 2, "art. 284 ust. 2 Pzp (tryb podstawowy)"
    else:
        wn, odp, pod = 14, 6, "art. 135 ust. 2 Pzp (przetarg nieograniczony)"
    wniosek = d - timedelta(days=wn)
    print(f"Pytania do SWZ z gwarancja odpowiedzi: wniosek najpozniej {wniosek} "
          f"(na {wn} dni przed terminem), odpowiedz zamawiajacego najpozniej {d - timedelta(days=odp)} - {pod}."
          + ("  << TEN TERMIN JUZ MINAL" if wniosek < dzis else f"  << zostalo {(wniosek - dzis).days} dni"))
    if a.pierwotny:
        p0 = date.fromisoformat(a.pierwotny)
        print(f"Pierwotny termin skladania {p0}: wniosek liczony od niego {p0 - timedelta(days=wn)}. Ma znaczenie tylko, "
              "gdy termin przedluzono, bo zamawiajacy spoznil sie z odpowiedziami (art. 135 ust. 3-4 / art. 284 ust. 3-4);"
              " przy przedluzeniu przez zmiane SWZ (art. 137) liczy sie aktualny termin.")
    print("Sprawdz w SWZ: zamawiajacy moze odpowiadac takze na pozniejsze pytania, ale nie musi (i nie musi wtedy przedluzac terminu).")


# ----------------------------------------------------------------- kontekst rynkowy
def cmd_rynek(a):
    d = json.load(open(Path(a.repo) / "wyniki_postepowan.json", encoding="utf-8"))
    post = d.get("postepowania", {})
    rek = list(post.values()) if isinstance(post, dict) else post
    zam = norm(a.zamawiajacy) if a.zamawiajacy else None
    slowa = [norm(s.strip()) for s in a.slowa.split(",")] if a.slowa else []

    def pasuje(r):
        if zam and zam not in norm(r.get("zamawiajacy") or ""):
            return False
        t = norm((r.get("tytul") or "") + " " + (r.get("typ") or ""))
        return all(s in t for s in slowa)

    traf = [r for r in rek if pasuje(r)]
    traf.sort(key=lambda r: r.get("termin_skladania") or "", reverse=True)
    print(f"Podobnych postepowan w bazie wynikow: {len(traf)} (waluta bazy: {d.get('_waluta')}, uwaga: czesc rekordow brutto - patrz uwagi)")
    print("| termin | zamawiajacy | tytul | budzet | oferty (min-max) | zwyciezca, cena | kryteria/uwagi |")
    print("|---|---|---|---|---|---|---|")
    for r in traf[: a.max]:
        bud = "; ".join(f"{b.get('kwota_pln'):,.0f}".replace(",", " ") for b in (r.get("budzet") or []) if b.get("kwota_pln"))
        ceny = [o.get("cena_pln") for o in (r.get("oferty") or []) if o.get("cena_pln")]
        of = f"{len(ceny)}: {min(ceny):,.0f}-{max(ceny):,.0f}".replace(",", " ") if ceny else "-"
        w = r.get("wynik") or {}
        zw = f"{(w.get('zwyciezca') or '')[:30]}, {w.get('cena_pln'):,.0f} {w.get('waluta') or ''}".replace(",", " ") if w.get("cena_pln") else ("uniewaznione" if w.get("uniewaznione") else "-")
        kr = re.search(r"[Kk]ryteria:?[^.]{0,80}", w.get("powod") or "")
        print(f"| {r.get('termin_skladania') or ''} | {(r.get('zamawiajacy') or '')[:35]} | {(r.get('tytul') or '')[:70]} | {bud or '-'} | {of} | {zw} | {kr.group(0) if kr else ''} |")


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = p.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("pobierz"); s.add_argument("url", nargs="?"); s.add_argument("--nr"); s.add_argument("--out", required=True)
    s.add_argument("--repo"); s.add_argument("--max-mb", type=float, default=60); s.add_argument("--max-dok", type=int, default=40)
    s = sub.add_parser("teksty"); s.add_argument("wejscie", nargs="+"); s.add_argument("--out", required=True)
    s.add_argument("--ocr", choices=["auto", "tak", "nie"], default="auto"); s.add_argument("--maks-ocr", type=int, default=40)
    s = sub.add_parser("szukaj"); s.add_argument("katalog"); s.add_argument("--temat"); s.add_argument("--wzor")
    s.add_argument("--plik", help="fragment nazwy pliku txt - szukaj tylko w nim")
    s.add_argument("--kontekst", type=int, default=350); s.add_argument("--max", type=int, default=10)
    sub.add_parser("tematy")
    s = sub.add_parser("roznice"); s.add_argument("stary"); s.add_argument("nowy")
    s = sub.add_parser("spis"); s.add_argument("plik")
    s = sub.add_parser("pytania"); s.add_argument("katalog"); s.add_argument("--wzor")
    s.add_argument("--max", type=int, default=60); s.add_argument("--dlugosc", type=int, default=700)
    s = sub.add_parser("strony"); s.add_argument("plik"); s.add_argument("--strony", required=True)
    s = sub.add_parser("terminy"); s.add_argument("--skladanie", required=True); s.add_argument("--godzina")
    s.add_argument("--pierwotny", help="pierwotny termin skladania RRRR-MM-DD, gdy byl przedluzany"); s.add_argument("--dzis")
    s.add_argument("--tryb", choices=["podstawowy", "nieograniczony"], default="podstawowy")
    s = sub.add_parser("rynek"); s.add_argument("--repo", required=True); s.add_argument("--zamawiajacy"); s.add_argument("--slowa")
    s.add_argument("--max", type=int, default=15)
    a = p.parse_args()
    if a.cmd == "pobierz" and not (a.url or a.nr):
        sys.exit("podaj URL postepowania albo --nr")
    {"pobierz": cmd_pobierz, "teksty": cmd_teksty, "szukaj": cmd_szukaj, "tematy": cmd_tematy,
     "terminy": cmd_terminy, "rynek": cmd_rynek, "roznice": cmd_roznice, "spis": cmd_spis,
     "strony": cmd_strony, "pytania": cmd_pytania}[a.cmd](a)


if __name__ == "__main__":
    main()
