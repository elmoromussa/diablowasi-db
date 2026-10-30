"""
Zona UTM i codis EPSG (lògica pura; extret de ``georef/coords_convert.py`` d'Oculat).

Mòdul de contracte (``docs/CONTRACTE_JACIMENT.md``): còpia idèntica a ``pipeline/contracte/`` d'Oculat. Sense cap
valor per defecte de jaciment: si el CRS no és un UTM reconegut i no hi ha longitud/latitud, ``utm_crs_info``
retorna ``None``; el selector de CRS de la GUI (``qt_comu.SelectorCrsUtm``) ofereix ``candidats_zona`` de la zona
detectada en lloc de fer teclejar un codi.
"""

from __future__ import annotations

import math
import re


def utm_zone_from_lon(lon_deg: float) -> int:
    """Zona UTM estàndard a partir de la longitud (sense les excepcions de Noruega/Svalbard)."""
    z = int(math.floor((float(lon_deg) + 180.0) / 6.0)) + 1
    return max(1, min(60, z))


def utm_crs_info(crs: str | None = None, lon_deg=None, lat_deg=None) -> dict | None:
    """
    Resol zona/hemisferi/EPSG d'un CRS UTM.

    crs: "EPSG:32718" (WGS 84 sud), "EPSG:326xx" (WGS 84 nord), "EPSG:258xx" (ETRS89 nord);
    també "EPSG::32718" o el codi sol. Si no és un UTM reconeixible i hi ha lon/lat,
    s'autodetecta la zona (WGS 84). Retorna ``{"epsg", "name", "zone", "northern"}`` o ``None``.
    """
    code = None
    if crs:
        m = re.search(r"(\d{4,5})\s*$", str(crs).strip())
        if m:
            code = int(m.group(1))
    if code is not None:
        if 32601 <= code <= 32660:
            z = code - 32600
            return {"epsg": f"EPSG:{code}", "name": f"WGS 84 / UTM zone {z}N", "zone": z, "northern": True}
        if 32701 <= code <= 32760:
            z = code - 32700
            return {"epsg": f"EPSG:{code}", "name": f"WGS 84 / UTM zone {z}S", "zone": z, "northern": False}
        if 25828 <= code <= 25838:
            z = code - 25800
            return {"epsg": f"EPSG:{code}", "name": f"ETRS89 / UTM zone {z}N", "zone": z, "northern": True}
    if lon_deg is not None and lat_deg is not None:
        z = utm_zone_from_lon(lon_deg)
        northern = float(lat_deg) >= 0
        code = (32600 if northern else 32700) + z
        return {
            "epsg": f"EPSG:{code}",
            "name": f"WGS 84 / UTM zone {z}{'N' if northern else 'S'}",
            "zone": z, "northern": northern,
        }
    return None


# ── Zona dels chunks i candidats ─────────────────────────────────────────────


def zona_utm_des_de_lonlat(lon: float, lat: float) -> tuple[int, int, str]:
    """``(epsg WGS 84, zona, hemisferi 'N'|'S')`` del punt."""
    info = utm_crs_info(None, float(lon), float(lat))
    return int(info["epsg"].split(":")[-1]), int(info["zone"]), "N" if info["northern"] else "S"


def zona_utm_de_epsg(epsg: int) -> tuple[int, str] | None:
    """``(zona, hemisferi)`` si l'EPSG és un UTM reconegut (WGS 84 326xx/327xx, ETRS89 258xx); si no, None."""
    info = utm_crs_info(f"EPSG:{int(epsg)}")
    if info is None or info["epsg"] != f"EPSG:{int(epsg)}":
        return None
    return int(info["zone"]), "N" if info["northern"] else "S"


def epsg_de_zona(zona: int, hemisferi: str) -> list[int]:
    """Codis EPSG habituals d'una zona: ETRS89 (258xx, només zones 28–38 N) primer i després WGS 84."""
    return [e for e, _ in candidats_zona(zona, hemisferi)]


def candidats_zona(zona: int, hemisferi: str) -> list[tuple[int, str]]:
    """``[(epsg, nom llarg), …]`` dels CRS UTM d'una zona, en l'ordre que mostra el selector: ETRS89 primer si la zona
    és de l'àmbit europeu (28–38 N), després WGS 84."""
    zona = int(zona)
    hemi = str(hemisferi).upper()[:1]
    if not 1 <= zona <= 60 or hemi not in ("N", "S"):
        raise ValueError(f"Zona UTM no vàlida: {zona}{hemisferi}")
    out: list[tuple[int, str]] = []
    if hemi == "N" and 28 <= zona <= 38:
        out.append((25800 + zona, f"ETRS89 / UTM zone {zona}N"))
    e = (32600 if hemi == "N" else 32700) + zona
    out.append((e, f"WGS 84 / UTM zone {zona}{hemi}"))
    return out


def nom_crs_utm(epsg: int | None) -> str | None:
    """Nom llarg d'un UTM reconegut (``ETRS89 / UTM zone 30N``); None si l'EPSG no n'és cap (o és buit)."""
    if not epsg:
        return None
    info = utm_crs_info(f"EPSG:{int(epsg)}")
    if info is None or info["epsg"] != f"EPSG:{int(epsg)}":
        return None
    return str(info["name"])


def comprovar_epsg(epsg: int, lon: float, lat: float) -> str | None:
    """Avís si la zona UTM de l'EPSG no és la del punt (centre dels chunks); None si coincideix o si
    l'EPSG no és un UTM reconegut (no es pot comprovar)."""
    z_epsg = zona_utm_de_epsg(epsg)
    if z_epsg is None:
        return None
    _e, zona, hemi = zona_utm_des_de_lonlat(lon, lat)
    if z_epsg == (zona, hemi):
        return None
    codis = " o ".join(f"EPSG:{c}" for c in epsg_de_zona(zona, hemi))
    return (f"EPSG:{int(epsg)} és UTM {z_epsg[0]}{z_epsg[1]}; els chunks cauen a la zona {zona}{hemi}: "
            f"{codis}.")


def text_zona(lon: float, lat: float) -> str:
    """«30N (centre 38.9700°, -0.1800°)» per a les etiquetes de la GUI."""
    _e, zona, hemi = zona_utm_des_de_lonlat(lon, lat)
    return f"{zona}{hemi} (centre {float(lat):.4f}°, {float(lon):.4f}°)"
