# mid-narzedzia

Narzędzia pomocnicze dla Claude w Pracowni Projektowej MiD. Repozytorium nie dotyczy systemu przetargów.

## Czytnik DWG (LibreDWG 0.13.3)

Gotowe programy do czytania rysunków AutoCAD, żeby Claude nie budował ich od zera w każdej sesji (ok. 5 min). Korzysta z nich skill „dwg”, którego kopia leży w `skill/dwg/SKILL.md`.

Zawartość:
- `libredwg/libredwg-0.13.3-linux-x86_64.tar.xz` – programy `dwg2dxf`, `dwgread`, `dwg2SVG` (build statyczny, Linux x86_64, zależą tylko od libc) oraz skrypt `dwg_tool.py`.
- `libredwg/libredwg-0.13.3-linux-x86_64.tar.xz.sha256` – suma kontrolna.

Pobranie w sesji Claude (po dołączeniu repo `mdudek-mid/mid-narzedzia-` do sesji):

```bash
rm -rf /tmp/mid-narzedzia
git clone -q --depth 1 https://github.com/mdudek-mid/mid-narzedzia- /tmp/mid-narzedzia
cd /tmp/mid-narzedzia/libredwg && sha256sum -c libredwg-0.13.3-linux-x86_64.tar.xz.sha256
mkdir -p ~/.local/libredwg && tar -xJf libredwg-0.13.3-linux-x86_64.tar.xz -C ~/.local/libredwg --strip-components=1
~/.local/libredwg/bin/dwg2dxf --version
```

Źródła: https://github.com/LibreDWG/libredwg (tag 0.13.3). Licencja LibreDWG: GPLv3.
Zbudowano 2026-10-09, gcc 13.3, Ubuntu 24.04:
`cmake -G Ninja -DCMAKE_BUILD_TYPE=Release -DBUILD_SHARED_LIBS=OFF -DENABLE_LTO=OFF -DDISABLE_WERROR=ON`
