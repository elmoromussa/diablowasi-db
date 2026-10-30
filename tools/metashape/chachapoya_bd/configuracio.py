"""Configuració del menú Base de dades: ``<jaciment>/bd.yaml`` i carpeta d'eixida ``<jaciment>/BD``.

A Oculat estes opcions eren el bloc ``database:`` de ``pipeline.yaml`` i la carpeta d'eixida era ``paths.db_csv``. Ací
viuen en un fitxer propi de la carpeta de jaciment, ``bd.yaml``, fora de ``jaciment.yaml`` (que és contracte entre
Oculat i Pontarch i no es toca), amb **les mateixes claus** que el bloc ``database:`` però al primer nivell
(``codi_regex``, ``sectors``, ``epsg_geo``, ``epsg_utm``, ``precisio_m``, ``eix_facana``, ``tilt_avis_deg``,
``min_horitz_deg``, ``divergencia_avis_m``). Es llig amb ``jaciment_yaml.escalars`` del contracte (sense PyYAML).

- **Carpeta de jaciment**: la del vincle del document (claus ``jaciment/*`` de ``doc.meta``, o la heretada que accepta
  ``jaciment_yaml.llegir_jaciment_meta``).
- **EPSG UTM**: ``bd.yaml`` (``epsg_utm``) > ``jaciment.yaml`` (``epsg``); ``font_epsg_utm`` diu d'on ve.
- **Carpeta d'eixida per defecte**: ``<jaciment>/BD``.
- Cap valor de reserva per a les claus obligatòries: sense vincle o sense ``bd.yaml`` alça ``ErrorConfiguracio``
  amb el text que es mostra a l'usuari; les claus absents les detecta ``codis.validar_opcions``.

Lògica pura (sense ``Metashape``): les eines només hi aporten ``Metashape.app.document.meta``.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from chachapoya_bd._vendor.contracte import jaciment_yaml
from chachapoya_bd.db import codis as db_codis

FITXER_BD = "bd.yaml"
CARPETA_EIXIDA = "BD"
CLAUS = tuple(db_codis.DEFAULTS)    # les del bloc database: de pipeline.yaml, sense reanomenar-ne cap

ERROR_SENSE_VINCLE = ("El projecte no està vinculat a cap carpeta de jaciment (no hi ha la clau jaciment/carpeta ni "
                      "la heretada pontarch/lliurament al document): vincula'l al jaciment abans d'usar este menú.")


class ErrorConfiguracio(ValueError):
    """Falta el vincle, la carpeta de jaciment o ``bd.yaml``; el missatge és el que es mostra a l'usuari."""


def error_sense_carpeta(carpeta: str | Path) -> str:
    return f"La carpeta de jaciment del vincle no existeix: {carpeta}"


def error_sense_bd(carpeta: str | Path) -> str:
    return (f"No hi ha {FITXER_BD} a la carpeta de jaciment: {Path(carpeta) / FITXER_BD} "
            f"(parteix de tools/metashape/chachapoya_bd/{FITXER_BD}.example).")


# ── Lògica pura sobre textos ─────────────────────────────────────────────────


def crs_de_jaciment(text_jaciment: str | None) -> str | None:
    """``"EPSG:<epsg>"`` de ``jaciment.yaml`` si declara un ``epsg`` enter; si no, None."""
    if not text_jaciment:
        return None
    epsg = jaciment_yaml.escalars(text_jaciment).get("epsg")
    if isinstance(epsg, bool) or not isinstance(epsg, int):
        return None
    return f"EPSG:{epsg}"


def opcions_de(text_bd: str, text_jaciment: str | None = None) -> dict[str, Any]:
    """Opcions del menú a partir del text de ``bd.yaml`` i, per a l'EPSG UTM, del de ``jaciment.yaml``."""
    return db_codis.opcions(jaciment_yaml.escalars(text_bd), crs_jaciment=crs_de_jaciment(text_jaciment))


def carpeta_eixida(carpeta: str | Path) -> Path:
    return Path(carpeta) / CARPETA_EIXIDA


# ── Carpeta de jaciment (vincle del document) i fitxers ──────────────────────


def carpeta_jaciment(meta: Any) -> Path:
    """Carpeta de jaciment del vincle de ``doc.meta`` (o d'un ``dict`` amb les mateixes claus)."""
    carpeta, _campanya, _font = jaciment_yaml.llegir_jaciment_meta(meta)
    if not carpeta:
        raise ErrorConfiguracio(ERROR_SENSE_VINCLE)
    return Path(carpeta)


def opcions(carpeta: str | Path) -> dict[str, Any]:
    """Opcions de ``<carpeta>/bd.yaml`` (+ ``epsg`` de ``<carpeta>/jaciment.yaml`` si n'hi ha)."""
    carpeta = Path(carpeta)
    if not carpeta.is_dir():
        raise ErrorConfiguracio(error_sense_carpeta(carpeta))
    f_bd = carpeta / FITXER_BD
    if not f_bd.is_file():
        raise ErrorConfiguracio(error_sense_bd(carpeta))
    f_jac = carpeta / jaciment_yaml.FITXER_JACIMENT
    text_jac = f_jac.read_text(encoding="utf-8") if f_jac.is_file() else None
    return opcions_de(f_bd.read_text(encoding="utf-8"), text_jac)


def opcions_de_meta(meta: Any) -> dict[str, Any]:
    return opcions(carpeta_jaciment(meta))


def carpeta_eixida_de_meta(meta: Any) -> Path:
    return carpeta_eixida(carpeta_jaciment(meta))
