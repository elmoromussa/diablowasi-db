"""Escriptura dels CSV per a la BD (UTF-8 amb BOM, que Access llig bé)."""

from __future__ import annotations

import csv
from collections.abc import Iterable, Sequence
from pathlib import Path


def fnum(v: float | None, dec: int = 3) -> str:
    return ("%." + str(dec) + "f") % v if v is not None else ""


def escriure_csv(path: str | Path, camps: Sequence[str], files: Iterable[dict]) -> Path:
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    with p.open("w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=list(camps), extrasaction="ignore")
        w.writeheader()
        for fila in files:
            w.writerow({k: fila.get(k, "") for k in camps})
    return p
