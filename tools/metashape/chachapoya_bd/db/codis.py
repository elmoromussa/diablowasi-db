r"""Gramàtica del codi d'estructura i opcions de ``bd.yaml``.

El codi d'estructura és la clau primària de la BD (``Code`` a ``T_STRUCTURES``).
La regex que el reconeix, el catàleg de sectors i l'EPSG UTM són propis de
cada jaciment i **no tenen valor per defecte**: es declaren a ``bd.yaml``
(``codi_regex``, ``sectors``, ``epsg_utm``) o, per
a l'EPSG, a ``jaciment.yaml`` (``epsg``). Exemple d'un jaciment amb
sectors (necròpolis): codi ``<SECTOR>-EA<n>`` amb regex
``^(?P<codi>(?P<sector>[A-Z]+-[A-Z0-9]+)-EA\d+)(?:-(?P<variant>.+))?$`` i
chunks amb sufix de variant (``<SECTOR>-EA##-<any>``); vegeu el bloc comentat
de ``bd.yaml.example``.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any

DEFAULTS: dict[str, Any] = {
    # Regex amb grups ``codi`` (obligatori), ``variant`` (opcional, sufix del
    # chunk) i ``sector`` (opcional, per comprovar-lo contra el catàleg).
    # Obligatòria a bd.yaml: no hi ha valor de reserva.
    "codi_regex": None,
    # Catàleg de sectors (L_SECTORS de la BD). Buit = no comprovar.
    "sectors": [],
    "epsg_geo": 4326,          # coordenades geogràfiques del CSV
    # EPSG UTM de sortida. Si s'omet, es pren de jaciment.yaml (epsg).
    "epsg_utm": None,
    "precisio_m": 0.05,        # Coord_Precision_m per defecte
    "eix_facana": "frontal",   # vista d'ortomosaics que fa de façana (+Y del region)
    "tilt_avis_deg": 15.0,     # els plans són verticals: més inclinació és sospitosa
    "min_horitz_deg": 20.0,    # pla a menys d'aquest angle de l'horitzontal → azimut ND
    "divergencia_avis_m": 0.25,  # marcador en dos chunks més lluny que això → avís
}

ERROR_CODI_REGEX = ("codi_regex no definit: és obligatori a bd.yaml "
                    "(regex amb el grup (?P<codi>...)).")
ERROR_EPSG_UTM = ("epsg_utm no definit: posa'l a bd.yaml o declara "
                  "epsg (UTM) a jaciment.yaml.")


def _codi_epsg(crs: str | None) -> int | None:
    m = re.search(r"(\d{4,5})\s*$", str(crs or "").strip())
    return int(m.group(1)) if m else None


def opcions(config: dict | None = None, crs_jaciment: str | None = None) -> dict[str, Any]:
    """Fusiona les claus de bd.yaml amb els valors per defecte.

    ``epsg_utm`` es pren de la config (``font_epsg_utm = "bd.yaml"``) o,
    si no hi és, del CRS UTM del jaciment (``"jaciment.yaml"``) quan
    ``crs_jaciment`` és un UTM reconegut per ``utm_crs_info``. Si no hi ha
    res, queda ``None`` (``validar_opcions`` ho detecta).
    """
    out = dict(DEFAULTS)
    for k, v in (config or {}).items():
        if k in out and v is not None:
            out[k] = v
    out["sectors"] = [str(s) for s in (out["sectors"] or [])]
    out["font_epsg_utm"] = None
    if out["epsg_utm"] is not None:
        out["font_epsg_utm"] = "bd.yaml"
    elif crs_jaciment:
        from chachapoya_bd._vendor.contracte.utm import utm_crs_info

        codi = _codi_epsg(crs_jaciment)
        # utm_crs_info (contracte) retorna None si no és un UTM reconegut; només
        # s'accepta si el codi que ix és el mateix que s'ha declarat.
        info = utm_crs_info(crs_jaciment)
        if codi is not None and info is not None and info["epsg"] == f"EPSG:{codi}":
            out["epsg_utm"] = codi
            out["font_epsg_utm"] = "jaciment.yaml"
    return out


def validar_opcions(o: dict | None, requereix_utm: bool = True) -> list[str]:
    """Errors (en català) per cada clau obligatòria absent; buit = tot correcte.

    ``requereix_utm=False`` per als consumidors que no projecten a UTM (el
    diàleg de mètriques de shapes només usa ``epsg_geo``).
    """
    o = o or {}
    errors = []
    if not o.get("codi_regex"):
        errors.append(ERROR_CODI_REGEX)
    if requereix_utm and o.get("epsg_utm") is None:
        errors.append(ERROR_EPSG_UTM)
    return errors


@dataclass(frozen=True)
class Codi:
    codi: str
    variant: str | None
    sector: str | None


class Codis:
    """Reconeix codis d'estructura en etiquetes de chunks i marcadors."""

    def __init__(self, opts: dict | None = None):
        o = opcions(opts)
        if not o["codi_regex"]:
            raise ValueError(ERROR_CODI_REGEX)
        self.regex = re.compile(o["codi_regex"])
        if "codi" not in self.regex.groupindex:
            raise ValueError("codi_regex ha de tenir el grup (?P<codi>...)")
        self.sectors = tuple(o["sectors"])

    def analitzar(self, label: str | None) -> Codi | None:
        m = self.regex.match((label or "").strip())
        if not m:
            return None
        g = m.groupdict()
        return Codi(g["codi"], g.get("variant") or None, g.get("sector") or None)

    def es_exacte(self, label: str | None) -> str | None:
        """Codi si l'etiqueta és un codi complet sense variant; si no, None."""
        c = self.analitzar(label)
        return c.codi if c and not c.variant else None

    def sector_fora_cataleg(self, codi: str) -> str | None:
        """Sector del codi si no és al catàleg (i hi ha catàleg); si no, None."""
        if not self.sectors:
            return None
        c = self.analitzar(codi)
        if c is None or c.sector is None:
            return None
        return c.sector if c.sector not in self.sectors else None
