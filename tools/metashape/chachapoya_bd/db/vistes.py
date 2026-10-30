"""Vistes del region (frontal, posterior, superior, inferior, dreta, esquerra) tal com les anomena
Pontarch: els plans f/b/t/bo/r/l de ``CONVENCIO_SHAPES_METRIQUES.md`` hi corresponen; és el contracte
entre els ortomosaics de Pontarch i les mètriques de la base de dades.

La convenció és la d'*Exportar Ortomosaics* de Pontarch (Metashape; coordenades locals del region):
  · +Y del region = façana (la mateixa convenció que ``georef/ea_front.py``)
  · +Z del region = vertical (amunt)
No és la de les càmeres del region de l'extensió de Blender (``alcat_N/E/S/W``), que classifica pel món.
La vista *frontal* posa l'observador a +Y mirant cap a la façana, *superior* el posa a +Z (planta), etc.
Els eixos retornats formen una base dextrogira (X dreta de la imatge, Y amunt de la imatge, Z cap a
l'observador).
"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class Vista:
    id: str        # identificador estable (nom de fitxer, etiqueta de l'asset)
    etiqueta: str  # text de la GUI
    eix: str       # eix local del region cap a l'observador: X, -X, Y, -Y, Z, -Z


VISTES: tuple[Vista, ...] = (
    Vista("frontal",   "Frontal (façana, +Y)",     "Y"),
    Vista("posterior", "Posterior (−Y)",           "-Y"),
    Vista("superior",  "Superior (planta, +Z)",    "Z"),
    Vista("inferior",  "Inferior (−Z)",            "-Z"),
    Vista("dreta",     "Dreta (+X)",               "X"),
    Vista("esquerra",  "Esquerra (−X)",            "-X"),
)

VISTES_PER_ID: dict[str, Vista] = {v.id: v for v in VISTES}


def eixos_vista(rx, ry, rz, eix: str):
    """Retorna (X, Y, Z) de la projecció a partir dels eixos locals del region.

    ``rx, ry, rz`` són les columnes de ``region.rot`` (qualsevol tipus de vector
    que suporte ``v * -1``). Z apunta cap a l'observador, X cap a la dreta de la
    imatge i Y cap amunt. Les vistes oposades giren 180° X i Z per no produir
    una imatge en mirall.
    """
    if eix == "Z":
        return rx, ry, rz
    if eix == "-Z":
        return rx * -1, ry, rz * -1
    if eix == "Y":
        return rx * -1, rz, ry
    if eix == "-Y":
        return rx, rz, ry * -1
    if eix == "X":
        return ry, rz, rx
    if eix == "-X":
        return ry * -1, rz, rx * -1
    raise ValueError(f"eix de vista desconegut: {eix!r}")


def eixos_per_vista(rx, ry, rz, vista_id: str):
    return eixos_vista(rx, ry, rz, VISTES_PER_ID[vista_id].eix)
