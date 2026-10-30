"""Coordenades i orientació de façana per estructura → CSV per a la BD.

Una fila per estructura (codi), amb dues fonts independents:

- **orientació de façana**: normal exterior del pla del bounding box del chunk
  d'orientació (la mateixa vista ``eix_facana`` que *Generar Ortomosaics…*),
  expressada com a azimut (0 = Nord), sector de 8 i tilt (control de qualitat:
  els plans són verticals, |tilt| gran = bbox mal orientat);
- **coordenades del punt de circulació**: marcador amb el codi exacte de
  l'estructura, situat en un chunk de coordenades (típicament el sector sencer).

Aquest mòdul no depèn de Metashape: el diàleg li passa ``ChunkInfo`` i dos
callables (``calc_orientacio``, ``calc_coordenades``) que fan la geometria amb
l'API. Arredoniment (convenció del projecte): UTM i altitud al cm, lat/lon a
7 decimals, azimut en graus enters.
"""

from __future__ import annotations

import math
from collections.abc import Callable, Sequence
from dataclasses import dataclass, field

from chachapoya_bd.db.codis import Codis

SECTORS8 = ("N", "NE", "E", "SE", "S", "SW", "W", "NW")


# ── Orientació ───────────────────────────────────────────────────────────────


def orientacio_facana(n_enu: Sequence[float], min_horitz_deg: float = 20.0,
                      tilt_avis_deg: float = 15.0):
    """Azimut, sector de 8 i tilt d'una normal expressada en ENU (Est, Nord, Amunt).

    Retorna ``(azimut_o_None, sector, tilt_deg, notes)``.
    """
    e, n, u = float(n_enu[0]), float(n_enu[1]), float(n_enu[2])
    notes: list[str] = []
    horiz = math.hypot(e, n)
    tilt = math.degrees(math.atan2(u, horiz))
    if horiz < math.sin(math.radians(min_horitz_deg)):
        notes.append("pla quasi horitzontal: azimut indeterminat (ND)")
        return None, "ND", tilt, notes
    azim = math.degrees(math.atan2(e, n)) % 360.0
    sector = SECTORS8[int(round(azim / 45.0)) % 8]
    if abs(tilt) > tilt_avis_deg:
        notes.append("bbox inclinat %.1f graus: revisar 'rotate to view'" % tilt)
    return azim, sector, tilt, notes


# ── Inventari ────────────────────────────────────────────────────────────────


@dataclass
class Marcador:
    label: str
    posicio: tuple[float, float, float] | None   # geocèntrica; None = no situat


@dataclass
class ChunkInfo:
    label: str
    enabled: bool = True
    georeferenciat: bool = True
    te_region: bool = True
    marcadors: list[Marcador] = field(default_factory=list)


@dataclass
class Candidat:
    chunk: ChunkInfo
    codi: str
    variant: str | None


def candidats(chunks: Sequence[ChunkInfo], codis: Codis):
    """Cataloga els chunks per a les dues llistes del diàleg.

    Retorna ``(orientacio, coordenades)``: ``orientacio`` són ``Candidat`` (chunks
    amb etiqueta de codi, amb o sense variant); ``coordenades`` són parelles
    ``(chunk, n_marcadors_de_codi_situats)``.
    """
    orient: list[Candidat] = []
    coord: list[tuple[ChunkInfo, int]] = []
    for ch in chunks:
        c = codis.analitzar(ch.label)
        if c:
            orient.append(Candidat(ch, c.codi, c.variant))
        n = sum(1 for mk in ch.marcadors
                if codis.es_exacte(mk.label) and mk.posicio is not None)
        if n:
            coord.append((ch, n))
    return orient, coord


def preseleccio_orientacio(c: Candidat) -> bool:
    """Chunks actius de codi exacte; les variants mai es marquen soles."""
    return c.chunk.enabled and c.variant is None


@dataclass
class Entrada:
    codi: str
    auth: ChunkInfo | None = None
    variants: list[ChunkInfo] = field(default_factory=list)
    colocacions: list[tuple[ChunkInfo, Marcador]] = field(default_factory=list)
    o_errors: list[str] = field(default_factory=list)
    c_errors: list[str] = field(default_factory=list)
    avisos: list[str] = field(default_factory=list)

    @property
    def chunk_orientacio(self) -> ChunkInfo | None:
        if self.auth is not None:
            return self.auth
        if len(self.variants) == 1:
            return self.variants[0]
        return None

    @property
    def estat(self) -> str:
        """completa | orientacio | coordenades | bloquejada"""
        o, c = not self.o_errors, not self.c_errors
        if o and c:
            return "completa"
        if o:
            return "orientacio"
        if c:
            return "coordenades"
        return "bloquejada"


def _dist(a: Sequence[float], b: Sequence[float]) -> float:
    return math.dist(a, b)


def construir_inventari(orient: Sequence[Candidat], coord: Sequence[ChunkInfo],
                        codis: Codis, divergencia_avis_m: float = 0.25) -> list[Entrada]:
    """Construeix les entrades a partir dels chunks MARCATS per l'usuari."""
    entrades: dict[str, Entrada] = {}

    for cand in orient:
        e = entrades.setdefault(cand.codi, Entrada(cand.codi))
        if cand.variant:
            e.variants.append(cand.chunk)
        else:
            if e.auth is not None:
                e.o_errors.append("dos chunks marcats amb l'etiqueta exacta '%s'" % cand.codi)
            e.auth = cand.chunk

    for ch in coord:
        for mk in ch.marcadors:
            codi = codis.es_exacte(mk.label)
            if codi is None or mk.posicio is None:
                continue
            entrades.setdefault(codi, Entrada(codi)).colocacions.append((ch, mk))

    for e in entrades.values():
        sector = codis.sector_fora_cataleg(e.codi)
        if sector:
            e.avisos.append("sector '%s' fora del catàleg de L_SECTORS" % sector)

        if e.auth is None and len(e.variants) > 1:
            e.o_errors.append("%d variants marcades i cap chunk de codi exacte: "
                              "marca'n només una" % len(e.variants))
        elif e.auth is None and len(e.variants) == 1:
            e.avisos.append("orientació presa de la variant '%s'" % e.variants[0].label)
        ock = e.chunk_orientacio
        if ock is None and not e.o_errors:
            e.o_errors.append("cap chunk d'orientació marcat")
        if ock is not None:
            if not ock.georeferenciat:
                e.o_errors.append("chunk d'orientació sense georeferenciar")
            elif not ock.te_region:
                e.o_errors.append("chunk d'orientació sense bounding box")

        per_chunk: dict[str, list[Marcador]] = {}
        for ch, mk in e.colocacions:
            per_chunk.setdefault(ch.label, []).append(mk)
        for label, mks in per_chunk.items():
            if len(mks) > 1:
                e.c_errors.append("%d marcadors '%s' al chunk '%s'" % (len(mks), e.codi, label))
        if not e.colocacions and not e.c_errors:
            e.c_errors.append("cap marcador '%s' als chunks de coordenades marcats" % e.codi)
        if e.colocacions and not e.colocacions[0][0].georeferenciat:
            e.c_errors.append("el chunk del marcador ('%s') no està georeferenciat"
                              % e.colocacions[0][0].label)
        if len(per_chunk) > 1 and not e.c_errors:
            e.avisos.append("marcador a %d chunks: s'usa '%s'"
                            % (len(per_chunk), e.colocacions[0][0].label))
            v1 = e.colocacions[0][1].posicio
            v2 = e.colocacions[1][1].posicio
            if v1 is not None and v2 is not None:
                d = _dist(v1, v2)
                if d > divergencia_avis_m:
                    e.avisos.append("divergència de %.2f m entre '%s' i '%s'"
                                    % (d, e.colocacions[0][0].label,
                                       e.colocacions[1][0].label))

    return [entrades[k] for k in sorted(entrades)]


def resum_estats(entrades: Sequence[Entrada]) -> dict[str, int]:
    out = {"completa": 0, "orientacio": 0, "coordenades": 0, "bloquejada": 0}
    for e in entrades:
        out[e.estat] += 1
    out["exportables"] = out["completa"] + out["orientacio"] + out["coordenades"]
    return out


# ── CSV ──────────────────────────────────────────────────────────────────────


def camps_csv(epsg_utm: int) -> tuple[str, ...]:
    return ("Code", "Lat_WGS84", "Lon_WGS84", "E_UTM%d" % epsg_utm, "N_UTM%d" % epsg_utm,
            "Alt_masl", "Precision_m", "Azimuth_deg", "Sector8", "Tilt_deg",
            "Orientation_Chunk", "Marker_Chunk", "Notes")


def files_csv(entrades: Sequence[Entrada], epsg_utm: int, precisio: str,
              calc_orientacio: Callable[[ChunkInfo], tuple],
              calc_coordenades: Callable[[ChunkInfo, Marcador], tuple]):
    """Construeix les files del CSV.

    ``calc_orientacio(chunk)`` → ``(azimut|None, sector, tilt, notes)``;
    ``calc_coordenades(chunk, marcador)`` → ``(lat, lon, est, nord, alt)``.
    Retorna ``(files, n_coords, n_orient)``.
    """
    camps = camps_csv(epsg_utm)
    files: list[dict] = []
    n_coords = n_orient = 0
    for e in entrades:
        fila = {k: "" for k in camps}
        fila["Code"] = e.codi
        notes = list(e.avisos)

        if e.o_errors:
            notes.append("SENSE ORIENTACIÓ: " + "; ".join(e.o_errors))
        else:
            ock = e.chunk_orientacio
            assert ock is not None
            fila["Orientation_Chunk"] = ock.label
            azim, sector, tilt, o_notes = calc_orientacio(ock)
            notes.extend(o_notes)
            fila["Azimuth_deg"] = "" if azim is None else str(int(round(azim)) % 360)
            fila["Sector8"] = sector
            fila["Tilt_deg"] = "%.1f" % tilt
            n_orient += 1

        if e.c_errors:
            notes.append("SENSE COORDENADES: " + "; ".join(e.c_errors))
        else:
            ch, mk = e.colocacions[0]
            lat, lon, est, nord, alt = calc_coordenades(ch, mk)
            fila["Lat_WGS84"] = "%.7f" % lat
            fila["Lon_WGS84"] = "%.7f" % lon
            fila[camps[3]] = "%.2f" % est
            fila[camps[4]] = "%.2f" % nord
            fila["Alt_masl"] = "%.2f" % alt
            fila["Precision_m"] = precisio
            fila["Marker_Chunk"] = ch.label
            n_coords += 1

        fila["Notes"] = "; ".join(notes)
        files.append(fila)
    return files, n_coords, n_orient
