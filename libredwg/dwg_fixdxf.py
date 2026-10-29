#!/usr/bin/env python3
"""dwg_fixdxf.py - naprawa DXF z dwg2dxf (LibreDWG), ktorego ezdxf nie wczytuje.

Objaw: ezdxf.readfile / recover przerywa ("Invalid group code", "DXFStructureError"), dwg_tool.py info/texts/render
pada, a DXF powstal. Przyczyna: dwg2dxf zapisuje wielowierszowe MTEXT z surowymi znakami konca wiersza, wiec
pary kod/wartosc sie rozjezdzaja. Skrypt skleja osierocone wiersze z poprzednia wartoscia jako \\P (akapit MTEXT).
Bajty zachowane (latin-1 round-trip) - dziala dla DXF ANSI (cp1250) i UTF-8.

Uzycie:  python3 dwg_fixdxf.py WE.dxf [WY.dxf]      (bez WY nadpisuje WE)
Z DWG:   dwg2dxf -y -o x.dxf x.dwg && python3 dwg_fixdxf.py x.dxf
"""
import sys

def fix(src, dst=None):
    dst = dst or src
    raw = open(src, "rb").read().decode("latin-1")
    nl = "\r\n" if "\r\n" in raw else "\n"
    lines = raw.split(nl)
    if lines and lines[-1] == "":
        lines.pop()
    out, i, n, fixes = [], 0, len(lines), 0
    while i < n:
        try:
            int(lines[i].strip())
        except ValueError:          # wiersz w miejscu kodu grupy = ciag dalszy poprzedniej wartosci
            if out:
                out[-1] += "\\P" + lines[i]
            fixes += 1
            i += 1
            continue
        out.append(lines[i])
        if i + 1 < n:
            out.append(lines[i + 1])
        i += 2
    out, seq = add_seqend(out)
    open(dst, "wb").write((nl.join(out) + nl).encode("latin-1"))
    return fixes + seq, dst

def add_seqend(lines):
    """Brakujace SEQEND po POLYLINE+VERTEX i po INSERT z atrybutami (66=1) - drugi znany blad dwg2dxf
    (ezdxf: 'Expected DXF entity POLYLINE or SEQEND')."""
    res, added, expect = [], 0, None      # expect: 'VERTEX' / 'ATTRIB' / None
    i = 0
    while i + 1 < len(lines):
        code, val = lines[i].strip(), lines[i + 1]
        if code == "0":
            v = val.strip()
            if expect and v not in (expect, "SEQEND"):
                res += ["  0", "SEQEND", "  8", "0"]; added += 1
            if v == "SEQEND":
                expect = None
            elif v == "POLYLINE":
                expect = "VERTEX"
            elif v == "INSERT":
                expect = None
            elif v not in ("VERTEX", "ATTRIB"):
                expect = None
        elif code == "66" and val.strip() == "1" and expect is None:
            # INSERT z atrybutami: kolejne encje ATTRIB i SEQEND
            j = len(res) - 1
            while j > 0 and res[j - 1].strip() != "0":
                j -= 2
            if j > 0 and res[j].strip() == "INSERT":
                expect = "ATTRIB"
        res += [lines[i], lines[i + 1]]
        i += 2
    res += lines[i:]
    return res, added

if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    f, d = fix(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
    print(f"naprawiono {f} wierszy -> {d}")
