#!/usr/bin/env python3
"""mes_tool.py - MES przesel mostowych dla Pracowni MiD: belka ciagla albo ruszt (grillage),
powierzchnie wplywu (metoda sprzezona), obciazenia ruchome, nosnosc uzytkowa i klasa MLC.

Modele obciazen (komenda 'modele'):
  LM1, LM2      PN-EN 1991-2 z wspolczynnikami alfa wg PTB 2022 par. 108 (klasa I / II)
  PN85_A..E     PN-85/S-10030: pojazd K (4 osie co 1,20 m, rozstaw kol 2,70 m) z phi=1,35-0,005L + q
  PN66_I        PN-66/B-02015 klasa I: pasma 0,60 m co 1,5 m, p=8 kPa + 2 x 80 kN/m co 1,6 m
  S42 S32 S24 S16 S10   samochody modelowe + q wg Instrukcji GDDKiA (Zarz. nr 17 GDDKiA z 1.06.2004)
  MLC40..MLC150 (K = kolowe, G = gasienicowe; 1 albo 2 kolumny)  PTB 2022 zal. 2 (STANAG 2021)

Komendy:
  wzor [--typ belka|ruszt] --out model.json       szablon modelu
  modele                                          parametry i zrodla modeli obciazen
  wplyw MODEL.json --odp M:5.85[:1] [--png plik]  linia / powierzchnia wplywu (typ:x[:dzwigar])
  obwiednia MODEL.json --obc LM1[,PN85_A,...] [--md raport.md] [--json wyniki.json]
  nosnosc MODEL.json [--md raport.md] [--json wyniki.json] [--png wykres.png]
        nosnosc uzytkowa (porownanie z obciazeniem normowym, interpolacja m_u), klasa MLC, wspolczynnik RF
  przekroj --b 500 --h 560 --d 510 --As 2400 --fck 35 --fyk 410 [--beff --hf --bw] [--gc 1.5 --gs 1.15]
        M_Rd i V_Rd,c przekroju zelbetowego (PN-EN 1992-1-1)
  test                                            testy weryfikacyjne (wzory zamkniete)

Jednostki: m, kN, kNm, kPa. Os x wzdluz obiektu od pierwszej podpory, y w poprzek od lewej krawedzi pomostu.
"""
import argparse, json, math, sys
from itertools import permutations
from pathlib import Path

import numpy as np

G = 9.81  # kN/t
_trapz = getattr(np, "trapezoid", None) or getattr(np, "trapz")

# ----------------------------------------------------------------------------- modele obciazen
POJAZDY_S = {  # Instrukcja GDDKiA (Zarz. 17/2004), rys. 2-6: osie od tylu do przodu [kN], rozstawy [m], q [kN/m]
    "S42": {"osie": [80, 80, 80, 100, 80], "rozstawy": [1.4, 1.4, 5.2, 3.2], "q": 5.0, "masa": 42, "kat": "1/S42"},
    "S32": {"osie": [100, 100, 70, 50], "rozstawy": [2.0, 5.2, 3.2], "q": 4.0, "masa": 32, "kat": "2/S32"},
    "S24": {"osie": [80, 80, 80], "rozstawy": [1.4, 4.2], "q": 4.0, "masa": 24, "kat": "3/S24"},
    "S16": {"osie": [100, 60], "rozstawy": [4.5], "q": 3.0, "masa": 16, "kat": "4/S16"},
    "S10": {"osie": [60, 40], "rozstawy": [4.0], "q": 2.0, "masa": 10, "kat": "5/S10"},
}
ROZSTAW_KOL_S = 1.75
ZNAK_B18 = {"2/S32": "32 t lub 36 t lub 40 t", "3/S24": "24 t lub 28 t", "4/S16": "16 t lub 20 t", "5/S10": "10 t lub 12 t"}

MLC_KOLOWE = {  # PTB 2022 zal. 2 ust. 6: naciski osi [t] od przodu, rozstawy [m], rozstaw kol [m]
    150: ([19.96, 38.10, 38.10, 29.03, 29.03], [3.66, 2.13, 6.71, 1.83], 2.48),
    120: ([16.33, 32.66, 32.66, 21.77, 21.77], [3.66, 1.83, 6.10, 1.52], 2.48),
    100: ([13.61, 27.22, 27.22, 18.14, 18.14], [3.66, 1.68, 6.25, 1.52], 2.58),
    80: ([10.89, 21.77, 21.77, 14.51, 14.51], [3.66, 1.52, 5.49, 1.52], 2.26),
    60: ([7.26, 16.33, 16.33, 11.79, 11.79], [3.66, 1.52, 4.57, 1.22], 2.38),
    40: ([6.35, 11.79, 11.79, 12.70], [3.66, 1.22, 4.88], 2.10),
}
MLC_GASIENICOWE = {  # masa [t], dlugosc gasienicy [m], szerokosc calkowita [m], szerokosc gasienicy [m]
    120: (108.86, 6.10, 4.27, 1.02), 100: (90.72, 5.49, 3.96, 0.94), 80: (72.58, 4.88, 3.66, 0.84),
    60: (54.43, 4.27, 3.35, 0.71), 40: (36.29, 3.66, 2.84, 0.56),
}
MLC_WYMAGANE = {"I": {"K1": 150, "K2": 100, "G1": 120, "G2": 80}, "II": {"K1": 120, "K2": 80, "G1": 100, "G2": 60}}

PN85 = {"A": (4.0, 800), "B": (3.0, 600), "C": (2.0, 400), "D": (1.6, 320), "E": (1.2, 240)}  # q [kPa], K [kN]

ZRODLA = {
    "LM1": "PN-EN 1991-2:2007 p. 4.3.2; alfa wg rozporządzenia MI z 24.06.2022 (Dz.U. 2022 poz. 1518) § 108 ust. 2: "
           "klasa I: αQ=1,00, αq1=1,33, αq2=2,40, αqi(i≥3)=αqr=1,20; klasa II: wszystkie 1,00",
    "LM2": "PN-EN 1991-2 p. 4.3.3; βQ=1,00 (PTB 2022 § 108 ust. 1 pkt 2)",
    "PN85": "PN-85/S-10030 (norma wycofana – do oceny obiektów zaprojektowanych wg niej): K – 4 osie co 1,20 m, rozstaw kół 2,70 m, "
            "oś K ≥ 2,00 m od krawężnika (2,50 m od bariery), jedno K na obiekcie, φ=1,35−0,005L ≤ 1,325 (L ≥ 70 m: 1,0); "
            "q na jezdni w strefach niekorzystnych (wg IBDiM 2013, Analiza możliwości poruszania się pojazdów o masie do 60 t)",
    "PN66": "PN-66/B-02015 klasa I: pasma 0,60 m co 1,5 m, p=8 kPa + dwa obciążenia liniowe 80 kN/m co 1,6 m, "
            "redukcja 0,8 przy > 3 pasmach, chodniki 4 kPa; φ wg ekspertyzy MiD 2021 DW226 (1+10/(20+3L)) – ZWERYFIKOWAĆ Z TEKSTEM NORMY",
    "S": "Instrukcja do określania nośności użytkowej drogowych obiektów mostowych, Zarządzenie nr 17 GDDKiA z 1.06.2004: "
         "samochody modelowe 1/S42…5/S10 + q, ≤ 2 pasma (oś 1,5 m od krawężnika i 3/4 szerokości jezdni), rozstaw kół 1,75 m; "
         "m_u = m_i + (m_i−1 − m_i)(W_N − W_i)/(W_i−1 − W_i)",
    "MLC": "PTB 2022 załącznik nr 2 (STANAG 2021): wartości charakterystyczne z nadwyżką dynamiczną, γQ=1,35; kolumna: 30,90 m między osiami "
           "sąsiednich pojazdów kołowych (30,50 m między gąsienicami); oś kół ≥ 0,65 m od krawężnika (gąsienica ≥ 0,35 m), "
           "bez krawężnika +0,50 m; dwie kolumny: osie kół ≥ 1,10 m (gąsienice ≥ 0,50 m)",
}


def phi_pn85(L):
    return 1.0 if L >= 70 else min(1.35 - 0.005 * L, 1.325)


def phi_wzor(wzor, L):
    if isinstance(wzor, (int, float)):
        return float(wzor)
    return float(eval(wzor, {"__builtins__": {}}, {"L": L, "min": min, "max": max}))


# ----------------------------------------------------------------------------- MES: ruszt / belka
class Ruszt:
    """Ruszt z dzwigarow podluznych (w, a=dw/dx, b=dw/dy w wezle) i elementow poprzecznych
    (sztywnych albo z przegubem w zamku miedzy dzwigarami). Jeden dzwigar = belka ciagla."""

    def __init__(self, m, przesla_idx=None):
        self.m = m
        L = m["przesla"]
        if przesla_idx is None:
            przesla_idx = list(range(len(L)))
        self.idx = przesla_idx
        x0 = sum(L[:przesla_idx[0]])
        self.x0 = x0
        dx = m.get("dx", 0.25)
        xs, podp = [0.0], [0.0]
        for i in przesla_idx:
            n = max(4, int(math.ceil(L[i] / dx - 1e-9)))
            st = np.linspace(podp[-1], podp[-1] + L[i], n + 1)[1:]
            xs += list(st)
            podp.append(podp[-1] + L[i])
        self.x = np.array(xs)
        self.podpory = podp
        dz = m["dzwigary"]
        self.y = np.array(dz["y"] if "y" in dz else [dz.get("y0", dz["rozstaw"] / 2) + i * dz["rozstaw"] for i in range(dz["liczba"])])
        self.ng, self.nx = len(self.y), len(self.x)
        EI = dz.get("EI", 1.0)
        self.EI = (lambda p: EI[p] if isinstance(EI, list) else EI)
        self.GJ = dz.get("GJ", 0.0)
        pop = m.get("poprzeczne", {"typ": "sztywne"})
        self.typ_pop = pop.get("typ", "sztywne") if self.ng > 1 else "brak"
        self.EIt, self.GJt = pop.get("EI_m", 1.0), pop.get("GJ_m", 0.0)
        self.skret_podp = m.get("skret_na_podporach", True)
        self._buduj()

    def _przeslo(self, x):
        for k in range(len(self.podpory) - 1):
            if x <= self.podpory[k + 1] + 1e-9:
                return self.idx[k]
        return self.idx[-1]

    def _buduj(self):
        from scipy.sparse import coo_matrix
        from scipy.sparse.linalg import splu
        ng, nx = self.ng, self.nx
        self.dof = {}
        n = 0
        for j in range(ng):
            for k in range(nx):
                self.dof[("g", j, k)] = (n, n + 1, n + 2)
                n += 3
        if self.typ_pop == "przegub":
            for j in range(ng - 1):
                for k in range(nx):
                    self.dof[("z", j, k)] = (n, n + 1, n + 2)  # w, bL, bR
                    n += 3
        self.ndof = n
        R, C, V = [], [], []

        def dodaj(dofs, k):
            for a, da in enumerate(dofs):
                for b, db in enumerate(dofs):
                    if da is not None and db is not None and k[a, b] != 0:
                        R.append(da); C.append(db); V.append(k[a, b])

        self.elem = {}
        for j in range(ng):
            for k in range(nx - 1):
                l = self.x[k + 1] - self.x[k]
                EI = self.EI(self._przeslo((self.x[k] + self.x[k + 1]) / 2))
                kb = belka_k(EI, l)
                d1, d2 = self.dof[("g", j, k)], self.dof[("g", j, k + 1)]
                dodaj((d1[0], d1[1], d2[0], d2[1]), kb)
                self.elem[(j, k)] = ((d1[0], d1[1], d2[0], d2[1]), kb)
                if self.GJ > 0:
                    dodaj((d1[2], d2[2]), self.GJ / l * np.array([[1, -1], [-1, 1.0]]))
        if ng > 1 and self.typ_pop != "brak":
            for k in range(nx):
                trib = ((self.x[min(k + 1, nx - 1)] - self.x[k]) + (self.x[k] - self.x[max(k - 1, 0)])) / 2
                EIt, GJt = self.EIt * trib, self.GJt * trib
                for j in range(ng - 1):
                    g1, g2 = self.dof[("g", j, k)], self.dof[("g", j + 1, k)]
                    if self.typ_pop == "przegub":
                        z = self.dof[("z", j, k)]
                        h1 = (self.y[j + 1] - self.y[j]) / 2
                        dodaj((g1[0], g1[2], z[0], z[1]), belka_k(EIt, h1))
                        dodaj((z[0], z[2], g2[0], g2[2]), belka_k(EIt, h1))
                        if GJt > 0:
                            dodaj((g1[1],), np.array([[GJt / h1]]))
                            dodaj((g2[1],), np.array([[GJt / h1]]))
                    else:
                        l = self.y[j + 1] - self.y[j]
                        dodaj((g1[0], g1[2], g2[0], g2[2]), belka_k(EIt, l))
                        if GJt > 0:
                            dodaj((g1[1], g2[1]), GJt / l * np.array([[1, -1], [-1, 1.0]]))
        # warunki brzegowe
        stale = set()
        kp = [int(np.argmin(abs(self.x - p))) for p in self.podpory]
        for j in range(ng):
            for k in kp:
                d = self.dof[("g", j, k)]
                stale.add(d[0])
                if self.skret_podp or ng == 1:
                    stale.add(d[2])
            if ng == 1 or self.typ_pop == "brak":
                for k in range(nx):
                    stale.add(self.dof[("g", j, k)][2])
        if self.typ_pop == "przegub":
            for j in range(ng - 1):
                for k in kp:
                    stale.add(self.dof[("z", j, k)][0])
        self.wolne = np.array([i for i in range(self.ndof) if i not in stale])
        K = coo_matrix((V, (R, C)), shape=(self.ndof, self.ndof)).tocsc()
        Kf = K[self.wolne][:, self.wolne]
        # mala sztywnosc na stopniach bez sztywnosci (np. b przy GJ=0) zapobiega osobliwosci
        diag = abs(Kf.diagonal())
        zero = diag < diag.max() * 1e-14
        if zero.any():
            from scipy.sparse import diags
            Kf = Kf + diags(np.where(zero, diag.max() * 1e-8, 0.0), format="csc")
        self.lu = splu(Kf.tocsc())
        self.mapa = -np.ones(self.ndof, dtype=int)
        self.mapa[self.wolne] = np.arange(len(self.wolne))

    def solve(self, f):
        u = np.zeros(self.ndof)
        u[self.wolne] = self.lu.solve(f[self.wolne])
        return u

    def wplyw(self, typ, x, j):
        """Powierzchnia wplywu odpowiedzi (M albo V) w dzwigarze j w przekroju x (lokalne x modelu).
        Zwraca macierz Z[ng, nx]: odpowiedz od sily jednostkowej (w dol) w wezle (dzwigar, stacja)."""
        k = int(np.argmin(abs(self.x - x)))
        lewa = typ == "V" and k > 0 and (x < self.x[k] - 1e-9 or k == self.nx - 1)
        if typ == "V" and not lewa and x > self.x[k] + 1e-9 and k < self.nx - 1:
            pass
        if typ == "M" or (typ == "V" and not lewa):
            ke = min(k, self.nx - 2)
            dofs, kb = self.elem[(j, ke)]
            if typ == "M" and k == self.nx - 1:
                ke = self.nx - 2
                dofs, kb = self.elem[(j, ke)]
                c_loc, znak = kb[3], -1.0
            elif typ == "M":
                c_loc, znak = kb[1], 1.0
            else:
                c_loc, znak = kb[0], -1.0
        else:  # V z lewej strony przekroju (koniec elementu k-1)
            dofs, kb = self.elem[(j, k - 1)]
            c_loc, znak = kb[2], 1.0
        c = np.zeros(self.ndof)
        for d, v in zip(dofs, c_loc):
            c[d] += znak * v
        z = self.solve(c)
        Z = np.array([[z[self.dof[("g", jj, kk)][0]] for kk in range(self.nx)] for jj in range(self.ng)])
        self.skok = k if typ == "V" else None   # nieciaglosc linii wplywu sily poprzecznej w wezle k
        return Z

    def wagi_y(self, y):
        """wagi rozdzialu sily w punkcie y na linie dzwigarow (dzwignia, ekstrapolacja liniowa na wspornikach)"""
        w = np.zeros(self.ng)
        if self.ng == 1:
            w[0] = 1.0
            return w
        j = int(np.clip(np.searchsorted(self.y, y) - 1, 0, self.ng - 2))
        t = (y - self.y[j]) / (self.y[j + 1] - self.y[j])
        w[j], w[j + 1] = 1 - t, t
        return w


def belka_k(EI, l):
    return EI / l ** 3 * np.array([[12, 6 * l, -12, 6 * l], [6 * l, 4 * l * l, -6 * l, 2 * l * l],
                                  [-12, -6 * l, 12, -6 * l], [6 * l, 2 * l * l, -6 * l, 4 * l * l]])


# ----------------------------------------------------------------------------- powierzchnia wplywu na siatce
class Powierzchnia:
    def __init__(self, r, Z, dy=0.05):
        self.r, self.Z = r, Z
        m = r.m
        self.B = m.get("szerokosc", float(r.y[-1] + (r.y[0] if r.ng > 1 else 0.5)))
        self.yg = np.arange(0, self.B + 1e-9, dy)
        W = np.array([r.wagi_y(y) for y in self.yg])        # ny x ng
        xg = r.x.copy()
        k = getattr(r, "skok", None)
        if k is not None:
            # skok linii wplywu V w przekroju: wartosci graniczne z lewej i z prawej (ekstrapolacja liniowa)
            n = len(xg)
            lewa = 2 * Z[:, k - 1] - Z[:, k - 2] if k >= 2 else (Z[:, k - 1] if k >= 1 else None)
            prawa = 2 * Z[:, k + 1] - Z[:, k + 2] if k <= n - 3 else (Z[:, k + 1] if k <= n - 2 else None)
            cols, xs = [], []
            for i in range(n):
                if i != k:
                    cols.append(Z[:, i]); xs.append(xg[i])
                    continue
                if lewa is not None:
                    cols.append(lewa); xs.append(xg[i] - (1e-6 if prawa is not None else 0))
                if prawa is not None:
                    cols.append(prawa); xs.append(xg[i] + (1e-6 if lewa is not None else 0))
            Z = np.array(cols).T
            xg = np.array(xs)
            self.Z = Z
        self.S = W @ Z                                       # ny x nx
        self.dy = dy
        self.xg = xg
        self.h = 0.05
        self.xf = np.arange(0.0, r.x[-1] + self.h / 2, self.h)
        self.Sf = np.array([np.interp(self.xf, self.xg, row) for row in self.S])
        self._ruch = {}

    def ruch(self, osie_x, obc, okres=None):
        """E[iy, i]: efekt osi (obciazenia obc w osiach osie_x wzgl. pierwszej osi) na linii yg[iy],
        pierwsza os w x = (i - nmax) * h. Kolumna pojazdow co 'okres'. Oba kierunki jazdy (max liczony dalej)."""
        key = (tuple(round(o, 3) for o in osie_x), tuple(round(q, 4) for q in obc), okres)
        if key in self._ruch:
            return self._ruch[key]
        h, nf = self.h, len(self.xf)
        offs, ws = list(osie_x), list(obc)
        if okres:
            nrep = int(self.xf[-1] // okres) + 2
            offs = [o + r * okres for r in range(nrep) for o in osie_x]
            ws = [q for r in range(nrep) for q in obc]
        nn = [int(round(o / h)) for o in offs]
        nmax = max(nn)
        Sp = np.zeros((self.Sf.shape[0], nf + 2 * nmax))
        Sp[:, nmax:nmax + nf] = self.Sf
        npos = nf + nmax
        E = np.zeros((self.Sf.shape[0], npos))
        if len(nn) > 20 and not okres and nn == list(range(nn[0], nn[0] + len(nn))) and len(set(ws)) == 1:
            C = np.concatenate([np.zeros((Sp.shape[0], 1)), np.cumsum(Sp, axis=1)], axis=1)
            a0, a1 = nn[0], nn[-1] + 1
            E = ws[0] * (C[:, a1:a1 + npos] - C[:, a0:a0 + npos])
            self._ruch[key] = (E,)
            return (E,)
        for n, w in zip(nn, ws):
            E += w * Sp[:, n:n + npos]
        # jazda w przeciwnym kierunku: lustrzane odbicie ukladu osi
        nr = [nmax - n for n in nn]
        if nr != sorted(nn) or ws != ws[::-1]:
            E2 = np.zeros_like(E)
            for n, w in zip(nr, ws):
                E2 += w * Sp[:, n:n + npos]
            E = (E, E2)
        else:
            E = (E,)
        self._ruch[key] = E
        return E

    def wiersz(self, E, y):
        i = np.clip((y - self.yg[0]) / self.dy, 0, len(self.yg) - 1.000001)
        i0 = int(i)
        t = i - i0
        return (1 - t) * E[i0] + t * E[min(i0 + 1, len(self.yg) - 1)]

    def linia(self, y):
        return np.interp(y, self.yg, np.arange(len(self.yg)))

    def wartosc_y(self, y):
        """wartosci wplywu wzdluz x dla sily na linii y"""
        i = np.clip((y - self.yg[0]) / self.dy, 0, len(self.yg) - 1.000001)
        i0 = int(i)
        t = i - i0
        return (1 - t) * self.S[i0] + t * self.S[min(i0 + 1, len(self.yg) - 1)]

    def calka(self, y0, y1, znak=1, tylko_niekorzystne=True):
        """calka wplywu po pasie y0..y1 (cala dlugosc), tylko czesci niekorzystne dla znaku"""
        msk = (self.yg >= y0 - 1e-9) & (self.yg <= y1 + 1e-9)
        if not msk.any() or y1 <= y0:
            return 0.0
        S = self.S[msk] * znak
        if tylko_niekorzystne:
            S = np.clip(S, 0, None)
        # trapezy w x, prostokaty w y przyciete do pasa
        Ix = _trapz(S, self.xg, axis=1)
        ys = self.yg[msk]
        if len(ys) == 1:
            return float(Ix[0] * (y1 - y0)) * znak
        return float(_trapz(Ix, ys) + Ix[0] * (ys[0] - y0) + Ix[-1] * (y1 - ys[-1])) * znak

    def calka_linii(self, y, znak=1, tylko_niekorzystne=True):
        v = self.wartosc_y(y) * znak
        if tylko_niekorzystne:
            v = np.clip(v, 0, None)
        return float(_trapz(v, self.xg)) * znak


def przejazd(P, y_kol, osie_x, obc_osi, znak=1, okres=None):
    """Maksymalny efekt (dla znaku) pojazdu przesuwanego wzdluz x, w obu kierunkach jazdy.
    y_kol: linie kol (obciazenie osi dzielone rowno na kola); osie_x: polozenia osi wzgl. pierwszej osi."""
    best, poz = -1e30, 0.0
    for E in P.ruch(osie_x, obc_osi, okres):
        row = sum(P.wiersz(E, y) for y in y_kol) / len(y_kol) * znak
        i = int(np.argmax(row))
        if row[i] > best:
            nmax = E.shape[1] - len(P.xf)
            best, poz = float(row[i]), (i - nmax) * P.h
    return best * znak, poz


def osie_z_rozstawow(rozst):
    xs = [0.0]
    for s in rozst:
        xs.append(xs[-1] + s)
    return xs


# ----------------------------------------------------------------------------- modele obciazen: efekt maksymalny
def jezdnia(m):
    j = m.get("jezdnia")
    if not j:
        sys.exit("Brak 'jezdnia': [y_lewy_krawęznik, y_prawy_krawęznik] w modelu.")
    return float(j[0]), float(j[1])


def efekt_LM1(P, m, znak, klasa="I", chodniki=0.0):
    a = {"I": dict(aQ=[1.0, 1.0, 1.0], aq=[1.33, 2.40, 1.20], aqr=1.20), "II": dict(aQ=[1.0, 1.0, 1.0], aq=[1.0, 1.0, 1.0], aqr=1.0)}[klasa]
    Q, q = [300.0, 200.0, 100.0], [9.0, 2.5, 2.5]
    y0, y1 = jezdnia(m)
    w = y1 - y0
    if w < 5.4:
        n, wl = 1, 3.0
    elif w < 6.0:
        n, wl = 2, w / 2
    else:
        n, wl = int(w // 3), 3.0
    najl = (-1e30, None)
    for off in np.arange(0, w - n * wl + 1e-9, 0.1) if w - n * wl > 1e-9 else [0.0]:
        pasy = [(y0 + off + i * wl, y0 + off + (i + 1) * wl) for i in range(n)]
        ts, udl = [], []
        for (a0, a1) in pasy:
            best = 0.0
            for yc in np.arange(a0 + 1.2, a1 - 1.2 + 1e-9, 0.1) if wl >= 2.4 else [(a0 + a1) / 2]:
                e, _ = przejazd(P, [yc - 1.0, yc + 1.0], [0.0, 1.2], [2.0, 2.0], znak)
                best = max(best, e * znak)
            ts.append(best)
            udl.append(P.calka(a0, a1, znak) * znak)
        reszta = P.calka(y0, pasy[0][0], znak) * znak + P.calka(pasy[-1][1], y1, znak) * znak
        for perm in permutations(range(n)):
            tot = 0.0
            for nr, p in enumerate(perm):
                i = min(nr, 2)
                tot += max(a["aQ"][i] * Q[i] / 2 * ts[p] if nr < 3 else 0.0, 0) + a["aq"][i] * q[i] * udl[p]
            tot += a["aqr"] * 2.5 * reszta
            if tot > najl[0]:
                najl = (tot, {"przesuniecie": round(float(off), 2), "kolejnosc": [int(p) + 1 for p in perm]})
    tot = najl[0]
    if chodniki:
        for c in m.get("chodniki", []):
            tot += chodniki * P.calka(c[0], c[1], znak) * znak
    return znak * tot, najl[1]


def efekt_LM2(P, m, znak, beta=1.0):
    y0, y1 = jezdnia(m)
    best = 0.0
    for yc in np.arange(y0 + 0.3 + 1.0, y1 - 0.3 - 1.0 + 1e-9, 0.05):
        e, _ = przejazd(P, [yc - 1.0, yc + 1.0], [0.0], [400.0 * beta], znak)
        best = max(best, e * znak)
    for yk in np.arange(y0 + 0.3, y1 - 0.3 + 1e-9, 0.05):   # jedno kolo 200 kN
        e, _ = przejazd(P, [yk], [0.0], [200.0 * beta], znak)
        best = max(best, e * znak)
    return znak * best, None


def efekt_PN85(P, m, znak, klasa, L, chodniki=None):
    q, K = PN85[klasa]
    y0, y1 = jezdnia(m)
    odst = 2.0 if m.get("kraweznik", True) else 2.5
    phi = phi_pn85(L)
    best = 0.0
    for yc in np.arange(y0 + odst, y1 - odst + 1e-9, 0.05) if (y1 - y0) >= 2 * odst else [(y0 + y1) / 2]:
        e, _ = przejazd(P, [yc - 1.35, yc + 1.35], [0.0, 1.2, 2.4, 3.6], [K / 4] * 4, znak)
        best = max(best, e * znak)
    tot = phi * best + q * P.calka(y0, y1, znak) * znak
    qc = 2.5 if chodniki is None else chodniki
    for c in m.get("chodniki", []):
        tot += qc * P.calka(c[0], c[1], znak) * znak
    return znak * tot, {"phi": round(phi, 3)}


def efekt_PN66(P, m, znak, L, phi="1+10/(20+3*L)", chodniki=4.0):
    y0, y1 = jezdnia(m)
    ph = phi_wzor(phi, L)
    najl = (0.0, None)
    for off in np.arange(0.0, 1.5, 0.05):
        sr = [c for c in np.arange(y0 + 0.3 + off, y1 - 0.3 + 1e-9, 1.5)]
        if not sr:
            continue
        ef = []
        for yc in sr:
            e_line, _ = przejazd(P, [yc - 0.2, yc, yc + 0.2], [0.0, 1.6], [48.0, 48.0], znak)
            e_p = 8.0 * P.calka(yc - 0.3, yc + 0.3, znak) * znak
            ef.append(e_line * znak + e_p)
        ef = sorted(ef, reverse=True)
        for k in range(1, len(ef) + 1):
            t = sum(ef[:k]) * (0.8 if k > 3 else 1.0)
            if t > najl[0]:
                najl = (t, {"pasm": k, "przesuniecie": round(float(off), 2)})
    tot = ph * najl[0]
    for c in m.get("chodniki", []):
        tot += chodniki * P.calka(c[0], c[1], znak) * znak
    return znak * tot, {**(najl[1] or {}), "phi": round(ph, 3)}


def efekt_S(P, m, znak, sym, L, phi="PN85"):
    v = POJAZDY_S[sym]
    y0, y1 = jezdnia(m)
    w = y1 - y0
    ph = phi_pn85(L) if phi == "PN85" else phi_wzor(phi, L)
    osie = osie_z_rozstawow(v["rozstawy"])
    best = (0.0, None)
    for strona in (0, 1):
        osi_pasm = [1.5, 0.75 * w]
        ys = [y0 + d if strona == 0 else y1 - d for d in osi_pasm]
        ef = []
        for yc in ys:
            e, _ = przejazd(P, [yc - ROZSTAW_KOL_S / 2, yc + ROZSTAW_KOL_S / 2], osie, v["osie"], znak)
            eq = v["q"] / 2 * (P.calka_linii(yc - ROZSTAW_KOL_S / 2, znak) + P.calka_linii(yc + ROZSTAW_KOL_S / 2, znak)) * znak
            ef.append(ph * e * znak + eq)
        for wyb, opis in (([0, 1], "2 pasma"), ([0], "pasmo 1"), ([1], "pasmo 2")):
            t = sum(ef[i] for i in wyb)
            if t > best[0]:
                best = (t, {"pasma": opis, "od": "lewego" if strona == 0 else "prawego", "phi": round(ph, 3)})
    return znak * best[0], best[1]


def efekt_MLC(P, m, znak, klasa, rodzaj="K", kolumny=1):
    y0, y1 = jezdnia(m)
    kr = m.get("kraweznik", True)
    if rodzaj == "K":
        osie_t, rozst, tor = MLC_KOLOWE[klasa]
        osie = osie_z_rozstawow(rozst)
        obc = [t * G for t in osie_t]
        okres = osie[-1] + 30.90
        od = 0.65 + (0 if kr else 0.5)
        lo, hi = y0 + od + tor / 2, y1 - od - tor / 2
        miedzy = 1.10 + tor  # rozstaw osi pojazdow w dwoch kolumnach: kolo-kolo >= 1,10 m
    else:
        masa, dl, B, b = MLC_GASIENICOWE[klasa]
        tor = B - b
        n = max(2, int(round(dl / 0.05)) + 1)
        osie = list(np.linspace(0, dl, n))
        obc = [masa * G / n] * n
        okres = dl + 30.50
        od = 0.35 + (0 if kr else 0.5)
        lo, hi = y0 + od + B / 2, y1 - od - B / 2
        miedzy = B + 0.50
    uwagi = []
    if hi < lo:
        lo = hi = (y0 + y1) / 2
        uwagi.append("jezdnia za wąska na odstępy ust. 5.2 – pojazd w osi jezdni (ust. 5.4 – sprawdzić)")
    def kolumna(yc):
        return przejazd(P, [yc - tor / 2, yc + tor / 2], osie, obc, znak, okres=okres)[0] * znak
    best = 0.0
    pozycje = np.arange(lo, hi + 1e-9, 0.05)
    e1s = [kolumna(yc) for yc in pozycje]
    best = max(e1s) if e1s else 0.0
    if kolumny == 2:
        if hi - lo < miedzy - 1e-9:
            uwagi.append("dwie kolumny nie mieszczą się z odstępami ust. 5.2–5.3 – wynik dla jednej kolumny; ust. 5.4 sprawdzić ręcznie")
        else:
            k2 = int(round(miedzy / 0.05))
            for i in range(len(pozycje) - k2):
                best = max(best, e1s[i] + e1s[i + k2])
    return znak * best, {"uwagi": uwagi} if uwagi else None


def efekt(P, m, model, znak, L):
    o = m.get("obciazenia", {})
    if model in ("LM1", "LM1_I", "LM1_II"):
        kl = "II" if model.endswith("II") else o.get("klasa_LM1", "I")
        return efekt_LM1(P, m, znak, kl, o.get("chodniki_gr1a", 3.0))
    if model == "LM2":
        return efekt_LM2(P, m, znak, o.get("betaQ", 1.0))
    if model.startswith("PN85_"):
        return efekt_PN85(P, m, znak, model[-1], L, o.get("chodniki_PN85"))
    if model == "PN66_I":
        return efekt_PN66(P, m, znak, L, o.get("phi_PN66", "1+10/(20+3*L)"), o.get("chodniki_PN66", 4.0))
    if model in POJAZDY_S:
        return efekt_S(P, m, znak, model, L, o.get("phi_S", "PN85"))
    if model.startswith("MLC"):
        # MLC80K1, MLC120G2 ...
        import re
        mm = re.fullmatch(r"MLC(\d+)([KG])([12])", model)
        if not mm:
            sys.exit(f"Model MLC: MLC<klasa><K|G><1|2>, np. MLC80K1 (podano {model})")
        return efekt_MLC(P, m, znak, int(mm.group(1)), mm.group(2), int(mm.group(3)))
    sys.exit(f"Nieznany model obciążenia: {model}")


# ----------------------------------------------------------------------------- stale
def efekt_stalych(P, m, r):
    wyn = []
    for s in m.get("stale", []):
        typ = s.get("typ", "pow")
        q = s["q"]
        if typ == "dzwigar":   # q [kN/m] na kazdy dzwigar (lista per przeslo albo liczba)
            qq = q[r.idx[0]] if isinstance(q, list) else q
            e = sum(qq * _trapz(P.Z[j], P.xg) for j in range(r.ng))
        elif typ == "liniowe":  # q [kN/m] na linii y
            e = q * _trapz(P.wartosc_y(s["y"]), P.xg)
        else:                   # q [kPa] na pasie y0..y1
            yy = s.get("y", [0, P.B])
            e = q * P.calka(yy[0], yy[1], 1, tylko_niekorzystne=False)
        g = s.get("gamma", [1.35, 1.0])
        wyn.append({"nazwa": s.get("nazwa", typ), "Ek": e, "gamma_sup": g[0], "gamma_inf": g[1]})
    return wyn


# ----------------------------------------------------------------------------- odpowiedzi i obwiednie
def modele_obliczeniowe(m):
    """lista (Ruszt, przeslo, L) - przesla swobodnie podparte osobno, belka ciagla jako calosc"""
    if m.get("ciagla", False):
        r = Ruszt(m)
        return [(r, None, max(m["przesla"]))]
    return [(Ruszt(m, [i]), i, m["przesla"][i]) for i in range(len(m["przesla"]))]


def odpowiedzi(m, r, nr):
    """domyslne przekroje: M w 0,3..0,7 L, V przy podporach; wszystkie dzwigary"""
    L = r.podpory[-1]
    wyn = []
    dz = m.get("dzwigary_obliczane") or list(range(r.ng))
    xsM = m.get("przekroje_M") or [0.5, 0.4, 0.6, 0.3, 0.7]
    pod = r.podpory
    for j in dz:
        for k in range(len(pod) - 1):
            Lk = pod[k + 1] - pod[k]
            for f in xsM:
                wyn.append(("M", pod[k] + f * Lk, j))
        wyn.append(("V", 0.0, j))
        wyn.append(("V", L, j))
        for p in pod[1:-1]:
            wyn.append(("M", p, j))
            wyn.append(("V", p - 1e-6, j))   # lewa strona podpory
            wyn.append(("V", p + 0.01, j))   # prawa strona podpory
    return wyn


_POW = []   # powierzchnie wplywu odpowiadajace kolejnym rekordom wynikow (do doliczania modeli)


def dolicz(m, wyniki, modele, indeksy):
    for i in indeksy:
        rec, P, Lp = wyniki[i], _POW[i][0], _POW[i][1]
        dom = 1 if P.S.max() >= -P.S.min() else -1
        znaki = (1, -1) if (m.get("ciagla") or m.get("obie_obwiednie")) else (dom,)
        for mod in modele:
            e = {z: efekt(P, m, mod, z, Lp) for z in znaki}
            emax, smax = e.get(1, (0.0, None))
            emin, smin = e.get(-1, (0.0, None))
            rec["ruchome"][mod] = {"max": emax, "min": emin, "ustawienie": smax if abs(emax) >= abs(emin) else smin}


def licz(m, modele, log=True):
    _POW.clear()
    wyniki = []
    for r, nr, Lmax in modele_obliczeniowe(m):
        for (typ, x, j) in odpowiedzi(m, r, nr):
            Z = r.wplyw(typ, x, j)
            P = Powierzchnia(r, Z)
            Lp = m["przesla"][r._przeslo(x)] if nr is None else m["przesla"][nr]
            rec = {"przeslo": (nr + 1) if nr is not None else "ciągła", "typ": typ, "x": round(r.x0 + x, 3), "dzwigar": j + 1,
                   "stale": efekt_stalych(P, m, r), "ruchome": {}}
            dom = 1 if P.S.max() >= -P.S.min() else -1
            znaki = (1, -1) if (m.get("ciagla") or m.get("obie_obwiednie")) else (dom,)
            for mod in modele:
                e = {z: efekt(P, m, mod, z, Lp) for z in znaki}
                emax, smax = e.get(1, (0.0, None))
                emin, smin = e.get(-1, (0.0, None))
                rec["ruchome"][mod] = {"max": emax, "min": emin, "ustawienie": smax if abs(emax) >= abs(emin) else smin}
            wyniki.append(rec)
            _POW.append((P, Lp))
        if log:
            print(f"  przęsło {nr + 1 if nr is not None else 'ciągłe'}: {len(odpowiedzi(m, r, nr))} odpowiedzi", file=sys.stderr)
    return wyniki


def miarodajny(rec, mod):
    v = rec["ruchome"][mod]
    if rec["typ"] == "V":
        return v["max"] if abs(v["max"]) >= abs(v["min"]) else v["min"]
    return v["max"] if abs(v["max"]) >= abs(v["min"]) else v["min"]


# ----------------------------------------------------------------------------- nosnosc uzytkowa i MLC
def ocena_nosnosci(m, wyniki):
    n = m.get("nosnosc", {})
    norma = n.get("obciazenie_normowe", "PN85_A")
    gN, gS = n.get("gamma_normowe", 1.0), n.get("gamma_uzytkowe", 1.0)
    kat = ["S42", "S32", "S24", "S16", "S10"]
    tab, najgorszy = [], None
    for rec in wyniki:
        WN = abs(miarodajny(rec, norma)) * gN
        if WN <= 1e-9:
            continue
        Wi = [abs(miarodajny(rec, s)) * gS for s in kat]
        ok = [w <= WN + 1e-9 for w in Wi]
        i = next((k for k, o in enumerate(ok) if o), None)
        if i is None:
            mu = None
        elif i == 0:
            mu = 42.0
        else:
            mi, mi1 = POJAZDY_S[kat[i]]["masa"], POJAZDY_S[kat[i - 1]]["masa"]
            mu = mi + (mi1 - mi) * (WN - Wi[i]) / (Wi[i - 1] - Wi[i])
        rek = {"przeslo": rec["przeslo"], "typ": rec["typ"], "x": rec["x"], "dzwigar": rec["dzwigar"], "WN": WN,
               **{s: w for s, w in zip(kat, Wi)}, "kategoria": POJAZDY_S[kat[i]]["kat"] if i is not None else "< 5/S10",
               "ik": i if i is not None else 99, "mu": mu, "W_S42/WN": Wi[0] / WN}
        tab.append(rek)
        if najgorszy is None or (rek["ik"], -(rek["mu"] or 0)) > (najgorszy["ik"], -(najgorszy["mu"] or 0)):
            najgorszy = rek
    wg_typu = {}
    for typ in ("M", "V"):
        tt = [d for d in tab if d["typ"] == typ]
        if tt:
            g = max(tt, key=lambda d: (d["ik"], -(d["mu"] or 0)))
            wg_typu[typ] = {"kategoria": g["kategoria"], "mu": g["mu"], "przeslo": g["przeslo"], "x": g["x"], "dzwigar": g["dzwigar"]}
    return {"norma": norma, "gamma_N": gN, "gamma_S": gS, "tabela": tab, "miarodajny": najgorszy, "wg_typu": wg_typu}


GAMMA_OBL_NORMY = {"PN85": 1.5, "PN66": 1.5, "LM1": 1.35, "LM2": 1.35}


def gamma_MLC_normy(n, norma):
    if "gamma_normowe_MLC" in n:
        return float(n["gamma_normowe_MLC"])
    for k, v in GAMMA_OBL_NORMY.items():
        if norma.startswith(k):
            return v
    return 1.35


def ocena_MLC(m, wyniki):
    n = m.get("nosnosc", {})
    norma = n.get("obciazenie_normowe", "PN85_A")
    gN = gamma_MLC_normy(n, norma)
    gM = 1.35
    wynik = {}
    for kod in ("K1", "K2", "G1", "G2"):
        klasy = sorted(MLC_KOLOWE if kod[0] == "K" else MLC_GASIENICOWE)
        dop = None
        for k in klasy:
            mod = f"MLC{k}{kod}"
            if mod not in wyniki[0]["ruchome"]:
                continue
            ok = all(abs(miarodajny(r, mod)) * gM <= abs(miarodajny(r, norma)) * gN + 1e-9
                     for r in wyniki if abs(miarodajny(r, norma)) > 1e-9)
            if ok:
                dop = k
            else:
                break
        wynik[kod] = dop
    wynik["_gamma_N"] = gN
    return wynik


def wspolczynnik_RF(m, wyniki):
    """RF = (R_d - sum gamma_G G) / (gamma_Q Q) dla odpowiedzi z nosnoscia (nosnosc.MRd / VRd na dzwigar)"""
    n = m.get("nosnosc", {})
    MRd, VRd = n.get("MRd"), n.get("VRd")
    if not (MRd or VRd):
        return None
    modele = n.get("RF_modele", ["LM1"])
    out = []
    for rec in wyniki:
        Rd = MRd if rec["typ"] == "M" else VRd
        if not Rd:
            continue
        if isinstance(Rd, dict):
            Rd = Rd.get(str(rec["przeslo"]), Rd.get("*"))
        Gd = sum(s["Ek"] * s["gamma_sup"] for s in rec["stale"])
        for mod in modele:
            if mod not in rec["ruchome"]:
                continue
            gQ = 1.35
            Q = abs(miarodajny(rec, mod))
            out.append({"przeslo": rec["przeslo"], "typ": rec["typ"], "x": rec["x"], "dzwigar": rec["dzwigar"], "model": mod,
                        "Rd": Rd, "Gd": Gd, "Qd": gQ * Q, "RF": (Rd - abs(Gd)) / (gQ * Q) if Q else None})
    return sorted(out, key=lambda d: d["RF"] if d["RF"] is not None else 1e9)


# ----------------------------------------------------------------------------- przekroj zelbetowy (EC2)
def przekroj_zelbetowy(b, h, d, As, fck, fyk, beff=None, hf=None, bw=None, gc=1.5, gs=1.15, acc=1.0):
    """M_Rd [kNm], V_Rd,c [kN]; wymiary mm, As mm2, MPa. Prostokatny blok naprezen lambda=0,8, eta=1 (fck<=50)."""
    fcd, fyd = acc * fck / gc, fyk / gs
    lam, eta = (0.8, 1.0) if fck <= 50 else (0.8 - (fck - 50) / 400, 1.0 - (fck - 50) / 200)
    bw = bw or b
    if beff and hf:
        x = As * fyd / (lam * eta * fcd * beff)
        if lam * x <= hf:
            MRd = As * fyd * (d - lam * x / 2)
            bx = beff
        else:
            Ff = eta * fcd * (beff - bw) * hf
            x = (As * fyd - Ff) / (lam * eta * fcd * bw)
            MRd = Ff * (d - hf / 2) + eta * fcd * bw * lam * x * (d - lam * x / 2)
            bx = bw
    else:
        x = As * fyd / (lam * eta * fcd * b)
        MRd = As * fyd * (d - lam * x / 2)
        bx = b
    ecu = 0.0035
    eyd = fyd / 200000
    xlim = ecu / (ecu + eyd) * d
    k = min(1 + math.sqrt(200 / d), 2.0)
    rho = min(As / (bw * d), 0.02)
    CRdc = 0.18 / gc
    vmin = 0.035 * k ** 1.5 * math.sqrt(fck)
    VRdc = max(CRdc * k * (100 * rho * fck) ** (1 / 3), vmin) * bw * d
    return {"MRd_kNm": MRd / 1e6, "x_mm": x, "x_lim_mm": xlim, "x_ok": x <= xlim, "VRdc_kN": VRdc / 1e3,
            "fcd": fcd, "fyd": fyd, "k": k, "rho_l": rho, "b_strefy": bx}


def klasa_z_wytrzymalosci_kostkowej(fc_cube):
    """klasa betonu PN-EN 206 o najwiekszej fck,cube <= wartosci (np. wynik sklerometru jako wytrz. kostkowa)"""
    klasy = [(12, 15), (16, 20), (20, 25), (25, 30), (30, 37), (35, 45), (40, 50), (45, 55), (50, 60)]
    wyb = None
    for fck, fcc in klasy:
        if fcc <= fc_cube + 1e-9:
            wyb = (fck, fcc)
    return f"C{wyb[0]}/{wyb[1]}" if wyb else "< C12/15"


# ----------------------------------------------------------------------------- raport
def fmt(v, n=1):
    if v is None:
        return "—"
    if isinstance(v, float):
        return f"{v:,.{n}f}".replace(",", " ").replace(".", ",")
    return str(v)


def raport_md(m, wyniki, nos=None, mlc=None, rf=None, modele=None):
    L = [f"# Obliczenia MES – {m.get('nazwa', '')}", ""]
    L += [f"Model: {'belka ciągła' if m.get('ciagla') else 'przęsła swobodnie podparte'} {' + '.join(fmt(float(x), 2) for x in m['przesla'])} m; "
          f"dźwigarów {len(m['dzwigary'].get('y', [])) or m['dzwigary'].get('liczba', 1)}; połączenie poprzeczne: "
          f"{m.get('poprzeczne', {}).get('typ', 'sztywne')}; siatka dx = {m.get('dx', 0.25)} m. Jednostki: kNm, kN.", ""]
    if nos and nos.get("miarodajny"):
        g = nos["miarodajny"]
        L += ["## Nośność użytkowa (Instrukcja GDDKiA, Zarz. nr 17/2004)", "",
              f"Obciążenie normowe: **{nos['norma']}** (γ = {fmt(nos['gamma_N'], 2)}); obciążenie użytkowe γ = {fmt(nos['gamma_S'], 2)}.", "",
              f"**Kategoria nośności użytkowej: {g['kategoria']}**" + (f", m_u = {fmt(g['mu'], 1)} t" if g.get("mu") else "") +
              f" – decyduje {g['typ']} w przęśle {g['przeslo']}, x = {fmt(g['x'], 2)} m, dźwigar {g['dzwigar']}.", ""]
        for typ, nazwa in (("M", "momentu zginającego"), ("V", "siły poprzecznej")):
            t = nos.get("wg_typu", {}).get(typ)
            if t:
                L.append(f"- wg {nazwa}: {t['kategoria']}" + (f", m_u = {fmt(t['mu'], 1)} t" if t.get("mu") else "") +
                         f" (przęsło {t['przeslo']}, x = {fmt(t['x'], 2)} m, dźwigar {t['dzwigar']})")
        L.append("")
        if g["kategoria"] in ZNAK_B18:
            L += [f"Oznakowanie (znak B-18): {ZNAK_B18[g['kategoria']]}.", ""]
        L += ["| Przęsło | Odp. | x [m] | Dźw. | W_N | S42 | S32 | S24 | S16 | S10 | Kategoria | m_u [t] |", "|---|---|---|---|---|---|---|---|---|---|---|---|"]
        top = sorted(nos["tabela"], key=lambda d: (-d["ik"], d.get("mu") or 0, -d["W_S42/WN"]))[:15]
        for d in top:
            L.append(f"| {d['przeslo']} | {d['typ']} | {fmt(d['x'], 2)} | {d['dzwigar']} | {fmt(d['WN'])} | {fmt(d['S42'])} | {fmt(d['S32'])} | "
                     f"{fmt(d['S24'])} | {fmt(d['S16'])} | {fmt(d['S10'])} | {d['kategoria']} | {fmt(d['mu'], 1)} |")
        L += ["", "Tabela: 15 najbardziej wytężonych odpowiedzi (obciążenia ruchome; obciążenie stałe jest jednakowe po obu stronach porównania).", ""]
    if mlc:
        L += ["## Klasa MLC (PTB 2022, zał. 2)", "",
              "| Pojazdy | Kolumny | Klasa MLC dopuszczalna | Wymagana dla nowych obiektów kl. I / II |", "|---|---|---|---|"]
        for kod, nazwa in (("K1", ("kołowe", 1)), ("K2", ("kołowe", 2)), ("G1", ("gąsienicowe", 1)), ("G2", ("gąsienicowe", 2))):
            if kod not in mlc:
                continue
            L.append(f"| {nazwa[0]} | {nazwa[1]} | {mlc.get(kod) or '< 40'} | {MLC_WYMAGANE['I'][kod]} / {MLC_WYMAGANE['II'][kod]} |")
        L += ["", f"Porównanie: efekt pojazdów MLC × γQ = 1,35 ≤ efekt obciążenia normowego × γ_N = {fmt(mlc.get('_gamma_N'), 2)} "
              f"(PTB 2022 § 108 ust. 7). Sprawdzono: {mlc.get('_odpowiedzi', 'wszystkie odpowiedzi')}.", ""]
    if rf:
        L += ["## Współczynnik nośności RF = (R_d − G_d) / (γ_Q Q_k)", "", "| Przęsło | Odp. | x | Dźw. | Model | R_d | G_d | Q_d | RF |", "|---|---|---|---|---|---|---|---|---|"]
        for d in rf[:10]:
            L.append(f"| {d['przeslo']} | {d['typ']} | {fmt(d['x'], 2)} | {d['dzwigar']} | {d['model']} | {fmt(d['Rd'])} | {fmt(d['Gd'])} | {fmt(d['Qd'])} | {fmt(d['RF'], 2)} |")
        L.append("")
    L += ["## Obwiednie (wartości charakterystyczne, z φ)", ""]
    mods = modele or list(wyniki[0]["ruchome"])
    L += ["| Przęsło | Odp. | x [m] | Dźw. | G_k | " + " | ".join(mods) + " |", "|---|---|---|---|---|" + "---|" * len(mods)]
    for typ in ("M", "V"):
        rr = [r for r in wyniki if r["typ"] == typ]
        rr.sort(key=lambda r: -max(abs(miarodajny(r, mods[0])), 0))
        for r in rr[:12]:
            L.append(f"| {r['przeslo']} | {typ} | {fmt(r['x'], 2)} | {r['dzwigar']} | {fmt(sum(s['Ek'] for s in r['stale']))} | "
                     + " | ".join(fmt(miarodajny(r, mo)) for mo in mods) + " |")
    L += ["", "## Źródła modeli obciążeń", ""]
    for k, v in ZRODLA.items():
        L.append(f"- **{k}**: {v}")
    L += ["", "Wyniki MES wymagają sprawdzenia przez projektanta (model, sztywności, podparcie). Narzędzie nie zastępuje programu certyfikowanego."]
    return "\n".join(L)


# ----------------------------------------------------------------------------- komendy
def wczytaj(p):
    return json.loads(Path(p).read_text(encoding="utf-8"))


def cmd_wzor(a):
    if a.typ == "belka":
        m = {"nazwa": "Belka ciągła – przykład", "przesla": [15.0, 20.0, 15.0], "ciagla": True, "dx": 0.25,
             "dzwigary": {"liczba": 1, "rozstaw": 1.0, "EI": 1.0}, "szerokosc": 10.0, "jezdnia": [1.0, 9.0], "chodniki": [],
             "stale": [{"nazwa": "ciężar własny", "typ": "dzwigar", "q": 150.0, "gamma": [1.35, 1.0]}],
             "obciazenia": {"klasa_LM1": "I"},
             "nosnosc": {"obciazenie_normowe": "PN85_A", "gamma_normowe": 1.5, "gamma_uzytkowe": 1.5}}
    else:
        m = {"nazwa": "Ruszt – przęsła swobodnie podparte z belek prefabrykowanych", "przesla": [8.6, 11.7, 8.6], "ciagla": False, "dx": 0.25,
             "dzwigary": {"liczba": 22, "rozstaw": 0.5, "y0": 0.25, "EI": [1.0, 1.0, 1.0], "GJ": 0.5,
                          "_uwaga": "sztywności względne; GJ/EI z przekroju belki (z redukcją na zarysowanie) – policz warianty 0,5 / 1,0 / sztywne"},
             "poprzeczne": {"typ": "przegub", "EI_m": 2.0, "GJ_m": 0.0, "_uwaga": "EI_m: sztywność poprzeczna belki na 1 m względem EI dźwigara (h^3/12 na 1 m / I belki)"},
             "szerokosc": 11.0, "jezdnia": [2.20, 9.25], "kraweznik": True, "chodniki": [[0.0, 2.20], [9.25, 11.0]],
             "stale": [{"nazwa": "dźwigary", "typ": "dzwigar", "q": [4.60, 5.27, 4.60], "gamma": [1.2, 0.9]},
                       {"nazwa": "nawierzchnia", "typ": "pow", "q": 2.30, "y": [2.20, 9.25], "gamma": [1.5, 0.9]}],
             "obciazenia": {"klasa_LM1": "I", "phi_PN66": "1+10/(20+3*L)", "chodniki_PN66": 4.0, "phi_S": "PN85"},
             "nosnosc": {"obciazenie_normowe": "PN66_I", "gamma_normowe": 1.0, "gamma_uzytkowe": 1.0, "MLC": True}}
    Path(a.out).write_text(json.dumps(m, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"Zapisano {a.out}")


def cmd_modele(a):
    print("Samochody modelowe GDDKiA (osie od tyłu [kN], rozstawy [m], q [kN/m], rozstaw kół 1,75 m):")
    for k, v in POJAZDY_S.items():
        print(f"  {v['kat']:6} {k}: osie {v['osie']} rozstawy {v['rozstawy']} q={v['q']} masa {v['masa']} t")
    print("\nPN-85/S-10030 klasy (q [kPa], K [kN]):", PN85)
    print("\nMLC kołowe (osie [t] od przodu, rozstawy [m], rozstaw kół [m]):")
    for k, v in sorted(MLC_KOLOWE.items()):
        print(f"  MLC {k}: {v[0]} {v[1]} tor {v[2]} m; masa {sum(v[0]):.2f} t")
    print("MLC gąsienicowe (masa [t], długość, szerokość całkowita, szerokość gąsienicy [m]):")
    for k, v in sorted(MLC_GASIENICOWE.items()):
        print(f"  MLC {k}: {v}")
    print("\nWymagane klasy MLC dla nowych obiektów (PTB 2022 zał. 2 ust. 2):", MLC_WYMAGANE)
    print("\nŹródła:")
    for k, v in ZRODLA.items():
        print(f"  {k}: {v}")


def cmd_wplyw(a):
    m = wczytaj(a.model)
    typ, x, *rest = a.odp.split(":")
    x = float(x)
    j = int(rest[0]) - 1 if rest else 0
    for r, nr, _ in modele_obliczeniowe(m):
        if r.x0 - 1e-9 <= x <= r.x0 + r.podpory[-1] + 1e-9:
            Z = r.wplyw(typ, x - r.x0, j)
            P = Powierzchnia(r, Z)
            if r.ng == 1:
                print("x [m]\twpływ")
                for xx, v in zip(r.x[:: max(1, len(r.x) // 40)] + r.x0, Z[0][:: max(1, len(r.x) // 40)]):
                    print(f"{xx:.2f}\t{v:.4f}")
            else:
                i = int(np.argmax(abs(P.S).max(axis=1)))
                print(f"Powierzchnia wpływu {typ} x={x} dźwigar {j + 1}: max {P.S.max():.4f}, min {P.S.min():.4f}")
                print("Rzędne poprzecznie (w przekroju x) dla dźwigarów:", " ".join(f"{v:.3f}" for v in Z[:, int(np.argmin(abs(r.x - (x - r.x0))))]))
            if a.png:
                import matplotlib
                matplotlib.use("Agg")
                import matplotlib.pyplot as plt
                fig, ax = plt.subplots(figsize=(9, 3.5 if r.ng == 1 else 5))
                if r.ng == 1:
                    ax.plot(r.x + r.x0, Z[0], color="#1F4E79")
                    ax.axhline(0, color="k", lw=0.5)
                    ax.set_xlabel("x [m]")
                    ax.invert_yaxis() if False else None
                else:
                    c = ax.contourf(r.x + r.x0, P.yg, P.S, 30, cmap="RdBu_r")
                    fig.colorbar(c, ax=ax)
                    ax.set_xlabel("x [m]"); ax.set_ylabel("y [m]")
                ax.set_title(f"Wpływ {typ} w x = {x} m, dźwigar {j + 1}")
                fig.tight_layout()
                fig.savefig(a.png, dpi=120)
            return
    sys.exit("x poza modelem")


def cmd_obwiednia(a):
    m = wczytaj(a.model)
    modele = a.obc.split(",")
    w = licz(m, modele)
    if a.json:
        Path(a.json).write_text(json.dumps(w, ensure_ascii=False, indent=1, default=float), encoding="utf-8")
    t = raport_md(m, w, modele=modele)
    if a.md:
        Path(a.md).write_text(t, encoding="utf-8")
    print(t)


def cmd_nosnosc(a):
    m = wczytaj(a.model)
    n = m.get("nosnosc", {})
    modele = [n.get("obciazenie_normowe", "PN85_A")] + list(POJAZDY_S)
    modele_mlc = [f"MLC{k}K{c}" for k in sorted(MLC_KOLOWE) for c in (1, 2)] + [f"MLC{k}G{c}" for k in sorted(MLC_GASIENICOWE) for c in (1, 2)]
    modele_rf = [mm for mm in n.get("RF_modele", ["LM1"]) if mm not in modele] if (n.get("MRd") or n.get("VRd")) else []
    if a.szybko:
        m["przekroje_M"] = [0.5]
    w = licz(m, modele)
    nos = ocena_nosnosci(m, w)
    # MLC i RF: dla wszystkich odpowiedzi (--pelne) albo dla najbardziej wytezonych wg S42/W_N
    norma = n.get("obciazenie_normowe", "PN85_A")
    def ratio(i):
        WN = abs(miarodajny(w[i], norma))
        return abs(miarodajny(w[i], "S42")) / WN if WN > 1e-9 else 0
    if a.pelne:
        wybrane = list(range(len(w)))
    else:
        wybrane = []
        for typ in ("M", "V"):
            ii = sorted((i for i in range(len(w)) if w[i]["typ"] == typ), key=ratio, reverse=True)
            wybrane += ii[:a.top]
    if n.get("MLC", True):
        print(f"  MLC: {len(wybrane)} odpowiedzi", file=sys.stderr)
        dolicz(m, w, modele_mlc, wybrane)
    if modele_rf:
        dolicz(m, w, modele_rf, range(len(w)))
    w_mlc = [w[i] for i in wybrane]
    mlc = ocena_MLC(m, w_mlc) if n.get("MLC", True) else None
    if mlc is not None:
        mlc["_odpowiedzi"] = "wszystkie" if a.pelne else f"{len(wybrane)} najbardziej wytężonych (wg S42/W_N)"
    rf = wspolczynnik_RF(m, w)
    t = raport_md(m, w, nos, mlc, rf, modele=[modele[0]] + list(POJAZDY_S))
    if a.json:
        Path(a.json).write_text(json.dumps({"nosnosc": nos, "MLC": mlc, "RF": rf, "wyniki": w}, ensure_ascii=False, indent=1, default=float), encoding="utf-8")
    if a.md:
        Path(a.md).write_text(t, encoding="utf-8")
    if a.png:
        wykres_nosnosci(nos, a.png)
    print(t)


def wykres_nosnosci(nos, out):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    tab = [d for d in nos["tabela"] if d["typ"] == "M"]
    if not tab:
        return
    prz = sorted(set(d["przeslo"] for d in tab), key=str)
    fig, axs = plt.subplots(len(prz), 1, figsize=(9, 2.6 * len(prz)), squeeze=False)
    for ax, p in zip(axs[:, 0], prz):
        dd = {}
        for d in tab:
            if d["przeslo"] == p:
                dd.setdefault(d["dzwigar"], d)
                if d["W_S42/WN"] > dd[d["dzwigar"]]["W_S42/WN"]:
                    dd[d["dzwigar"]] = d
        js = sorted(dd)
        for s, kol in (("S42", "#C00000"), ("S32", "#ED7D31"), ("S24", "#70AD47")):
            ax.plot(js, [dd[j][s] / dd[j]["WN"] for j in js], "o-", ms=3, color=kol, label=s)
        ax.axhline(1.0, color="k", lw=1)
        ax.set_ylabel("W_S / W_N")
        ax.set_title(f"Przęsło {p}: moment – stosunek obciążenia użytkowego do normowego")
        ax.legend(fontsize=8, ncol=3)
    axs[-1, 0].set_xlabel("dźwigar")
    fig.tight_layout()
    fig.savefig(out, dpi=120)


def cmd_przekroj(a):
    w = przekroj_zelbetowy(a.b, a.h, a.d, a.As, a.fck, a.fyk, a.beff, a.hf, a.bw, a.gc, a.gs, a.acc)
    print(json.dumps({k: (round(v, 3) if isinstance(v, float) else v) for k, v in w.items()}, ensure_ascii=False, indent=1))
    if not w["x_ok"]:
        print("UWAGA: x > x_lim – zbrojenie nie osiąga granicy plastyczności (przekrój przezbrojony).")


# ----------------------------------------------------------------------------- testy
def cmd_test(a):
    ok = True

    def spr(nazwa, v, ref, tol=0.01):
        nonlocal ok
        e = abs(v - ref) / max(abs(ref), 1e-9)
        s = "OK " if e <= tol else "BŁĄD"
        ok &= e <= tol
        print(f"[{s}] {nazwa}: {v:.4f} (oczekiwane {ref:.4f}, różnica {100 * e:.2f}%)")

    # 1. belka swobodnie podparta L=10: M w srodku od P=1 w srodku = L/4; od q=1: L^2/8
    m = {"przesla": [10.0], "dx": 0.1, "dzwigary": {"liczba": 1, "rozstaw": 1.0, "EI": 1.0}, "szerokosc": 1.0}
    r = Ruszt(m)
    Z = r.wplyw("M", 5.0, 0)
    spr("SS M(L/2) od P w L/2 = L/4", Z[0][50], 2.5)
    spr("SS M(L/2) od q=1 = L²/8", _trapz(Z[0], r.x), 12.5)
    Zv = r.wplyw("V", 0.0, 0)
    Pv = Powierzchnia(r, Zv)
    spr("SS V(0+) od P w x = 0,10 m = 0,99", float(np.interp(0.10, Pv.xg, Pv.wartosc_y(0.5))), 0.99, 0.001)
    spr("SS V(0+) od q=1 = qL/2", Pv.calka_linii(0.5, 1, False), 5.0, 0.001)
    # 2. belka ciagla 2 x 10 m: M nad podpora od q=1 na obu przeslach = -qL^2/8
    m2 = {"przesla": [10.0, 10.0], "ciagla": True, "dx": 0.1, "dzwigary": {"liczba": 1, "rozstaw": 1.0, "EI": 1.0}, "szerokosc": 1.0}
    r2 = Ruszt(m2)
    Z2 = r2.wplyw("M", 10.0, 0)
    spr("Ciągła 2×10: M_B od q=1 = −qL²/8", _trapz(Z2[0], r2.x), -12.5)
    # max moment przeslowy od q na jednym przesle: 0.0957 qL^2 (dla 2 przesel, obciazone jedno)
    Zp = r2.wplyw("M", 4.375, 0)
    spr("Ciągła 2×10: M(0,4375L) od q na przęśle 1 = 0,0957 qL²", _trapz(np.clip(Zp[0], 0, None), r2.x), 9.57, 0.01)
    # 3. ruszt sztywny poprzecznie (EIt duze): rozdzial rowny przy obciazeniu w osi
    m3 = {"przesla": [10.0], "dx": 0.25, "dzwigary": {"liczba": 5, "rozstaw": 1.0, "y0": 0.5, "EI": 1.0},
          "poprzeczne": {"typ": "sztywne", "EI_m": 1e6}, "szerokosc": 5.0}
    r3 = Ruszt(m3)
    Z3 = r3.wplyw("M", 5.0, 2)
    spr("Ruszt sztywny: środkowy dźwigar, P w osi → L/4/5", Z3[2][20], 0.5, 0.02)
    Z3e = r3.wplyw("M", 5.0, 0)
    # sztywna poprzecznica, P nad skrajnym: wsp. = 1/n + e^2/sum(y^2) = 0.2 + 4/10 = 0.6 (bez skrecania)
    spr("Ruszt sztywny: skrajny dźwigar, P nad nim → 0,6·L/4 (Courbon)", Z3e[0][20], 0.6 * 2.5, 0.03)
    # 4. pojazd: SS L=10, S10 (60+40 kN co 4,0 m): Mmax z wzoru (wypadkowa)
    P = Powierzchnia(r, Z)
    e, _ = przejazd(P, [0.5], [0.0, 4.0], [60.0, 40.0], 1)
    # wypadkowa R=100 w odl. 1,6 m od 60 kN; 60 kN w x = 5 - 0,8 = 4,2: M = RA*4,2 ; RA = 100*(10-5.8)/10... liczymy wprost
    xs = np.linspace(0, 10, 100001)
    best = 0
    for x1 in np.linspace(0, 10, 2001):
        x2 = x1 + 4.0
        Mx = 0
        for xp, Pp in ((x1, 60), (x2, 40)):
            if 0 <= xp <= 10:
                Mx += Pp * (xp * (10 - 5) / 10 if xp <= 5 else 5 * (10 - xp) / 10)
        best = max(best, Mx)
    spr("Przejazd 60+40 kN co 4 m, M(L/2)", e, best, 0.005)
    # 5. EC2 przekroj prostokatny
    w = przekroj_zelbetowy(1000, 500, 450, 2000, 30, 500)
    fyd = 500 / 1.15; fcd = 30 / 1.5
    x = 2000 * fyd / (0.8 * fcd * 1000)
    spr("EC2 M_Rd prostokąt", w["MRd_kNm"], 2000 * fyd * (450 - 0.4 * x) / 1e6, 1e-6)
    print("\nWszystkie testy zaliczone." if ok else "\nSĄ BŁĘDY.")
    if not ok:
        sys.exit(1)


def main():
    import signal
    if hasattr(signal, "SIGPIPE"):
        signal.signal(signal.SIGPIPE, signal.SIG_DFL)
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("wzor"); s.add_argument("--typ", default="ruszt", choices=["belka", "ruszt"]); s.add_argument("--out", default="model.json")
    sub.add_parser("modele")
    s = sub.add_parser("wplyw"); s.add_argument("model"); s.add_argument("--odp", required=True); s.add_argument("--png")
    s = sub.add_parser("obwiednia"); s.add_argument("model"); s.add_argument("--obc", default="LM1"); s.add_argument("--md"); s.add_argument("--json")
    s = sub.add_parser("nosnosc"); s.add_argument("model"); s.add_argument("--md"); s.add_argument("--json"); s.add_argument("--png")
    s.add_argument("--szybko", action="store_true", help="tylko M w L/2 i V przy podporach")
    s.add_argument("--pelne", action="store_true", help="MLC dla wszystkich odpowiedzi (wolniej)")
    s.add_argument("--top", type=int, default=8, help="liczba najbardziej wytężonych odpowiedzi M i V do sprawdzenia MLC")
    s = sub.add_parser("przekroj")
    for k in ("b", "h", "d", "As", "fck", "fyk"):
        s.add_argument("--" + k, type=float, required=True)
    for k in ("beff", "hf", "bw"):
        s.add_argument("--" + k, type=float)
    s.add_argument("--gc", type=float, default=1.5); s.add_argument("--gs", type=float, default=1.15); s.add_argument("--acc", type=float, default=1.0)
    sub.add_parser("test")
    a = ap.parse_args()
    {"wzor": cmd_wzor, "modele": cmd_modele, "wplyw": cmd_wplyw, "obwiednia": cmd_obwiednia, "nosnosc": cmd_nosnosc,
     "przekroj": cmd_przekroj, "test": cmd_test}[a.cmd](a)


if __name__ == "__main__":
    main()
