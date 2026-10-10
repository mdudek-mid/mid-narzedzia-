---
name: mes-eurokody
description: MES przęseł mostowych MiD: ruszt i belka ciągła, LM1/LM2, PN-85, PN-66, samochody S GDDKiA i MLC; nośność użytkowa (Zarz. 17), klasa MLC, RF, M_Rd/V_Rd wg EC2.
---

# MES i Eurokody dla mostów (MiD)

Używaj, gdy trzeba policzyć siły wewnętrzne w przęśle od obciążeń ruchomych albo ocenić nośność istniejącego obiektu. Typowe prośby: „policz nośność użytkową mostu”, „jaka kategoria S i znak B-18”, „czy most przeniesie MLC 60”, „obwiednia momentów od LM1”, „linia wpływu podpory”, „M_Rd belki”, „sprawdź obliczenia z ekspertyzy”.

`mes_tool.py` to własny solver MES. Elementy belkowe Eulera-Bernoulliego tworzą ruszt z dźwigarów podłużnych i elementów poprzecznych (sztywnych albo z przegubem w zamku). Powierzchnie wpływu liczone są metodą sprzężoną (jedno rozwiązanie układu na odpowiedź), a obciążenia ruchome przesuwane po siatce 5 cm z przeszukiwaniem ustawień poprzecznych. Model (geometria, sztywności, obciążenia stałe) przygotowujesz ty, z dokumentacji, inwentaryzacji i badań.

## Instalacja (≈10 s)

```bash
rm -rf /tmp/mid-narzedzia && git clone -q --depth 1 https://github.com/mdudek-mid/mid-narzedzia- /tmp/mid-narzedzia
mkdir -p ~/.local/mes && cp /tmp/mid-narzedzia/mes/mes_tool.py ~/.local/mes/
pip install -q --break-system-packages numpy scipy matplotlib
M=~/.local/mes/mes_tool.py; python3 -I $M test        # testy weryfikacyjne – muszą przejść
```

## Model (JSON)

```bash
python3 -I $M wzor --typ ruszt --out model.json      # albo --typ belka (cały przekrój jako jedna belka)
```
- `przesla` [m], `ciagla` (false = każde przęsło swobodnie podparte, liczone osobno), `dx` (0,25 m dla rusztu, 0,1 m dla belki).
- `dzwigary`: `liczba`, `rozstaw`, `y0` (oś pierwszego od lewej krawędzi pomostu) albo lista `y`; `EI` (liczba albo lista na przęsła), `GJ`. W porównaniach liczą się sztywności względne, więc wystarczy EI = 1 i stosunki.
- `poprzeczne`: `typ` `przegub` (prefabrykaty łączone zamkami, pasma płytowe), `sztywne` (płyta, poprzecznice) albo `brak`; `EI_m`, `GJ_m` na 1 m długości.
- `szerokosc`, `jezdnia` [y krawężnika L, y krawężnika P], `kraweznik`, `chodniki` [[y0, y1], …].
- `stale`: `dzwigar` (q kN/m na każdy dźwigar), `pow` (q kPa na pasie `y`), `liniowe` (q kN/m na linii `y`), z `gamma` [sup, inf].
- `obciazenia`: `klasa_LM1` I/II, `phi_PN66`, `chodniki_PN66`, `chodniki_PN85`, `phi_S` (domyślnie φ wg PN-85 dla samochodów modelowych), `betaQ`.
- `nosnosc`: `obciazenie_normowe` (np. `PN66_I`, `PN85_B`, `LM1`), `gamma_normowe`, `gamma_uzytkowe` (przy porównaniu wg Zarz. 17 jednakowe γ się skracają), `gamma_normowe_MLC`, `MLC`, opcjonalnie `MRd`, `VRd` [kNm, kN na dźwigar] i `RF_modele`.

**Sztywność poprzeczna i skręcanie decydują o rozdziale obciążenia.** Na moście w Wiślinie (22 belki Gromnik) przy przegubach i GJ/EI = 0,5–1,0 oraz przy połączeniu sztywnym wychodzi 2/S32 wg momentu, a przy GJ = 0 tylko 4/S16. Zawsze policz 2–3 warianty (`przegub` z GJ/EI 0,5 i 1,0 oraz `sztywne`), podaj rozrzut i uzasadnij przyjęty model. Jeżeli badania wykazały uszkodzone zamki albo przecieki między belkami, przyjmij wariant ostrożny.

## Komendy

```bash
python3 -I $M modele                                      # parametry modeli obciążeń i źródła
python3 -I $M wplyw model.json --odp M:14.45:13 --png wplyw.png      # typ:x[:dźwigar], x globalne
python3 -I $M obwiednia model.json --obc LM1,LM2,PN85_A --md obwiednia.md --json obwiednia.json
python3 -I $M nosnosc model.json --md nosnosc.md --json nosnosc.json --png nosnosc.png [--szybko] [--pelne]
python3 -I $M przekroj --b 490 --h 560 --d 510 --As 2400 --fck 35 --fyk 410 [--beff --hf --bw] [--gc 1.5 --gs 1.15]
```
`--szybko` liczy M tylko w L/2 (V zawsze przy podporach). Domyślnie liczone są M w 0,3–0,7 L każdego przęsła, V po obu stronach podpór i M nad podporami belki ciągłej. MLC liczone jest dla 8 najbardziej wytężonych odpowiedzi M i 8 V (wg S42/W_N). `--pelne` liczy wszystkie odpowiedzi, wolniej (most 3-przęsłowy z 22 dźwigarami: ≈30 s w trybie `--szybko`).

## Modele obciążeń (źródła w raporcie)

| Model | Zawartość |
|---|---|
| `LM1` | PN-EN 1991-2, pasy 3 m, TS 300/200/100 kN, UDL 9/2,5 kPa; α wg PTB 2022 § 108: kl. I αq1 = 1,33, αq2 = 2,40, αqi = αqr = 1,20, αQ = 1,00; kl. II 1,00; chodniki 3 kPa (gr1a) |
| `LM2` | oś 400 kN βQ (koła co 2,0 m) albo jedno koło 200 kN |
| `PN85_A…E` | K = 800/600/400/320/240 kN (4 osie co 1,20 m, koła 2,70 m, oś K ≥ 2,00 m od krawężnika), φ = 1,35 − 0,005L ≤ 1,325; q = 4,0/3,0/2,0/1,6/1,2 kPa |
| `PN66_I` | pasma 0,60 m co 1,5 m: 8 kPa + 2 × 80 kN/m co 1,6 m; redukcja 0,8 przy > 3 pasmach; chodniki 4 kPa |
| `S42 S32 S24 S16 S10` | samochody modelowe Zarz. 17/2004 + q 5/4/4/3/2 kN/m; ≤ 2 pasma (oś 1,5 m od krawężnika i 3/4 szerokości jezdni), koła co 1,75 m |
| `MLC<kl><K/G><1/2>` | pojazdy kołowe (K) i gąsienicowe (G) PTB 2022 zał. 2, 1 albo 2 kolumny, γQ = 1,35, z nadwyżką dynamiczną |

Do sprawdzenia z tekstem norm (oznaczone w raporcie): współczynnik dynamiczny PN-66 (domyślnie 1 + 10/(20 + 3L), jak w ekspertyzie MiD 2021 dla DW 226), obciążenie chodników PN-85 (domyślnie 2,5 kPa) i γ obliczeniowe obciążeń normowych do porównania z MLC (PN-85/PN-66: 1,5).

## Nośność użytkowa – procedura (Zarz. nr 17 GDDKiA z 1.06.2004)

1. Ustal normatyw i klasę, na które obiekt zaprojektowano (ostatnia przebudowa). Obiekty zaprojektowane na klasę A albo B wg PN-85 mają 1/S42 bez obliczeń.
2. Zbuduj model z inwentaryzacji: rozpiętości, przekrój, połączenie dźwigarów, obciążenia stałe. Uwzględnij ubytki z badań w MRd, jeśli liczysz RF.
3. `nosnosc`: dla każdego dźwigara porównywane są W_N (normowe) i W_i (kategorie S) dla M i V. Kategoria to najwyższa, przy której W_i ≤ W_N. Dla kategorii poniżej 1/S42 m_u = m_i + (m_i−1 − m_i)(W_N − W_i)/(W_i−1 − W_i). Gdy M i V dają różne wyniki, decyduje niższy. Wynik daje znak B-18 (np. 2/S32: 32, 36 albo 40 t).
4. MLC: najwyższa klasa, przy której efekt × 1,35 ≤ W_N × γ_N. Wymagania dla nowych obiektów (kl. I: 150/100 kołowe, 120/80 gąsienicowe; kl. II: 120/80, 100/60) podawaj jako tło. Ocena MLC istniejącego obiektu jest podstawą oznakowania (PTB 2022 § 108 ust. 7).
5. Jeśli podano `MRd`/`VRd`: RF = (R_d − G_d)/(γ_Q Q_k) dla wybranych modeli (np. LM1). RF < 1 oznacza niewystarczającą nośność na dany model.

**Weryfikacja na obiekcie MiD:** most DW 226 w Wiślinie, ekspertyza 2021: 2/S32 wg momentu, porównanie z kl. I PN-66. Narzędzie daje 2/S32 wg momentu (S42/W_N ≈ 1,03, S32/W_N ≈ 0,90). Wskazuje też, że siła poprzeczna przy podporach daje 3/S24, a ekspertyza 2021 tego nie sprawdzała. Testy `test` sprawdzają belkę swobodnie podpartą i ciągłą (wzory zamknięte, 0,0957 qL²), ruszt sztywny (Courbon), przejazd pojazdu i EC2. Belkę ciągłą sprawdzono też z PyNite (zgodność do 0,001 kNm).

## Zasady

- Wynik MES to narzędzie projektanta. W raporcie podaj model, warianty, założenia i źródła obciążeń, a rozstrzygnięcie zostaw autorowi ekspertyzy (uprawnienia).
- Beton: wynik sklerometru albo odwiertów jako wytrzymałość kostkową przelicz na klasę PN-EN 206 (`klasa_z_wytrzymalosci_kostkowej` w kodzie). Stare klasy B: B45 to f_c,cube 45 MPa (C35/45).
- γc: zalecane 1,5 (EN). Zastosowany załącznik krajowy i αcc podaj jawnie (`--gc`, `--acc`).
- Pliki przekaż użytkownikowi (SendUserFile). Do ekspertyzy wyniki trafiają przez skill ekspertyza-mostu.

## Odpowiedź (krótko)

Kategoria nośności użytkowej i m_u (M i V osobno, z miejscem decydującym), znak B-18, klasy MLC (1 i 2 kolumny, kołowe i gąsienicowe), wrażliwość na model poprzeczny, pliki MD/JSON/PNG oraz lista założeń do potwierdzenia.
