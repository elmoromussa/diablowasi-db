"""Configuració del menú Base de dades: quin ``bd.yaml`` es llig i carpeta d'eixida per defecte.

A Oculat estes opcions eren el bloc ``database:`` de ``pipeline.yaml`` del clon i el menú funcionava amb qualsevol
projecte. Ací viuen en un fitxer propi, ``bd.yaml``, amb **les mateixes claus** que el bloc ``database:`` però al
primer nivell (``codi_regex``, ``sectors``, ``epsg_geo``, ``epsg_utm``, ``precisio_m``, ``eix_facana``,
``tilt_avis_deg``, ``min_horitz_deg``, ``divergencia_avis_m``). Es llig amb ``jaciment_yaml.escalars`` del contracte
(sense PyYAML). ``jaciment.yaml`` (contracte entre Oculat i Pontarch) no es toca.

- **Quin ``bd.yaml``**: ``<jaciment>/bd.yaml`` si el projecte està vinculat a una carpeta de jaciment (claus
  ``jaciment/*`` de ``doc.meta``, o la heretada que accepta ``jaciment_yaml.llegir_jaciment_meta``) i el fitxer hi
  és; si no, el ``bd.yaml`` d'este paquet (el de Diablo Wasi, versionat). El vincle no és mai obligatori. Un
  vincle que apunta a una carpeta inexistent es tracta com a «sense vincle», però deixa un ``avis`` de text.
- **EPSG UTM**: ``epsg_utm`` del ``bd.yaml`` triat > ``epsg`` de ``<jaciment>/jaciment.yaml`` (només amb vincle);
  ``font_epsg_utm`` és la ruta del fitxer d'on ve i ``origen_epsg_utm``, el mateix en paraules curtes.
- **Carpeta d'eixida per defecte**: ``<jaciment>/BD`` amb vincle; si no, la carpeta del ``.psx``; si el document no
  està desat, la carpeta de l'usuari.
- Cap valor de reserva per a les claus obligatòries: sense cap ``bd.yaml`` alça ``ErrorConfiguracio`` amb els llocs
  on l'ha buscat; les claus absents les detecta ``codis.validar_opcions``.

Lògica pura (sense ``Metashape``): les eines només hi aporten ``Metashape.app.document.meta`` i ``.path``.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

from chachapoya_bd._vendor.contracte import jaciment_yaml
from chachapoya_bd.db import codis as db_codis

FITXER_BD = "bd.yaml"
CARPETA_EIXIDA = "BD"
CARPETA_PAQUET = Path(__file__).resolve().parent     # on viu el bd.yaml del repositori
CLAUS = tuple(db_codis.DEFAULTS)    # les del bloc database: de pipeline.yaml, sense reanomenar-ne cap

FONT_JACIMENT = "jaciment"
FONT_REPO = "repo"
DESCRIPCIO_FONT = {FONT_JACIMENT: "de la carpeta de jaciment", FONT_REPO: "del repositori"}
# D'on ve l'EPSG UTM, per a la línia «Sistemes»: (font_epsg_utm de codis.opcions, font del bd.yaml) -> text curt.
ORIGEN_EPSG_UTM = {("bd.yaml", FONT_REPO): "de bd.yaml del repositori",
                   ("bd.yaml", FONT_JACIMENT): "de bd.yaml del jaciment",
                   ("jaciment.yaml", FONT_REPO): "de jaciment.yaml",
                   ("jaciment.yaml", FONT_JACIMENT): "de jaciment.yaml"}


class ErrorConfiguracio(ValueError):
    """No hi ha cap ``bd.yaml``; el missatge és el que es mostra a l'usuari."""


@dataclass(frozen=True)
class Configuracio:
    opcions: dict[str, Any]
    fitxer_bd: Path          # bd.yaml que s'ha llegit
    font: str                # FONT_JACIMENT | FONT_REPO
    carpeta_eixida: Path     # per defecte; el diàleg en deixa triar una altra
    origen_epsg_utm: str | None = None   # ORIGEN_EPSG_UTM; None si no hi ha EPSG UTM
    avis: str | None = None              # vincle trencat; None si no n'hi ha cap


def error_sense_bd(carpeta_jac: Path | None, carpeta_paquet: str | Path) -> str:
    """Text de l'error amb els dos llocs on s'ha buscat ``bd.yaml``."""
    if carpeta_jac is None:
        lloc_jac = f"<jaciment>/{FITXER_BD}: el projecte no està vinculat a cap carpeta de jaciment"
    else:
        lloc_jac = str(carpeta_jac / FITXER_BD)
    return (f"No hi ha {FITXER_BD} en cap dels dos llocs on es busca:\n"
            f"  1. {lloc_jac}\n"
            f"  2. {Path(carpeta_paquet) / FITXER_BD}\n"
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


# ── Vincle del document i fitxers ────────────────────────────────────────────


def carpeta_jaciment(meta: Any) -> Path | None:
    """Carpeta de jaciment del vincle de ``doc.meta`` (o d'un ``dict`` amb les mateixes claus), si existeix."""
    carpeta, _campanya, _font = jaciment_yaml.llegir_jaciment_meta(meta if meta is not None else {})
    if not carpeta or not Path(carpeta).is_dir():
        return None
    return Path(carpeta)


def avis_vincle(meta: Any) -> str | None:
    """Avís si el vincle de ``doc.meta`` apunta a una carpeta que no existeix; None sense vincle o amb vincle bo."""
    carpeta, _campanya, _font = jaciment_yaml.llegir_jaciment_meta(meta if meta is not None else {})
    if not carpeta or Path(carpeta).is_dir():
        return None
    return f"El vincle apunta a {carpeta}, que no existeix: s'usa el {FITXER_BD} del repositori"


def carpeta_eixida(meta: Any, ruta_psx: str | Path | None) -> Path:
    """``<jaciment>/BD`` amb vincle; si no, la carpeta del ``.psx``; si no està desat, la de l'usuari."""
    jac = carpeta_jaciment(meta)
    if jac is not None:
        return jac / CARPETA_EIXIDA
    if ruta_psx:
        return Path(ruta_psx).parent
    return Path.home()


def configuracio(meta: Any, ruta_psx: str | Path | None,
                 carpeta_paquet: str | Path = CARPETA_PAQUET) -> Configuracio:
    """Opcions, ``bd.yaml`` usat, font i carpeta d'eixida per al document (``meta``, ``ruta_psx``)."""
    jac = carpeta_jaciment(meta)
    candidats = ([(jac / FITXER_BD, FONT_JACIMENT)] if jac is not None else []) \
        + [(Path(carpeta_paquet) / FITXER_BD, FONT_REPO)]
    trobat = next(((f, font) for f, font in candidats if f.is_file()), None)
    if trobat is None:
        raise ErrorConfiguracio(error_sense_bd(jac, carpeta_paquet))
    f_bd, font = trobat

    f_jac = jac / jaciment_yaml.FITXER_JACIMENT if jac is not None else None
    if f_jac is not None and not f_jac.is_file():
        f_jac = None
    text_jac = f_jac.read_text(encoding="utf-8") if f_jac is not None else None
    o = opcions_de(f_bd.read_text(encoding="utf-8"), text_jac)
    origen = ORIGEN_EPSG_UTM.get((o["font_epsg_utm"], font))
    o["font_epsg_utm"] = {"bd.yaml": str(f_bd),
                          "jaciment.yaml": str(f_jac) if f_jac is not None else None}.get(o["font_epsg_utm"])
    return Configuracio(o, f_bd, font, carpeta_eixida(meta, ruta_psx), origen, avis_vincle(meta))
