"""Mètriques des de shapes 3D dibuixades sobre el model → 4 CSV per a la BD.

GRAMÀTICA d'etiquetes:  ``PLA[-MÈTRICA]-[ELEMENT]-COS[+COS...][a-z][-COS][-res|-c]``

- Plans: ``f b t bo r l`` (frontal, posterior, superior, inferior, dreta,
  esquerra = les vistes de *Generar Ortomosaics…*).
- Mètriques (polilínies): ``ll`` longitud 3D multi-tram · ``hh`` component
  vertical · ``dd`` corda 3D absoluta, un sol tram.
- Elements: ``port int corn crown led bra wall und cav roof eave plat slot sill
  jamb lint tie pil bbeam tbeam ebeam ext``.
- Estatus: (cap) conservat · ``-c`` complet · ``-r`` reconstruït (``-res`` acceptat).
- Relacions: ``-COS`` interfície · ``+COS`` aplicabilitat (mateix mur, entre cossos).
- Lletres: instàncies (2n portal…); els fragments d'una cara repeteixen l'etiqueta.

Semàntica del prefix: polígons → pla de projecció (vigilat per la normal);
polilínies → pertinença al mur (vigilada per proximitat als polígons).
Jerarquia d'evidència: ``-c`` directa > ``-res`` directa > propagada (``+``).

SIS controls de qualitat: obliqüitat/prefix, pertinença, direcció, conflicte
``-res`` ∧ ``-c``, discrepància directa/propagada (>2 %), orfandat. A més:
etiquetes invàlides amb proposta de canònic i exclusió de shapes 2D llegades.

Aquest mòdul és pur (numpy): el diàleg converteix els vèrtexs de Metashape a
``Shape`` i els marcs (eixos del region, centre, vertical) a ``Marc``.
"""

from __future__ import annotations

import math
import re
from collections.abc import Sequence
from dataclasses import dataclass, field

import numpy as np

from chachapoya_bd.db.csv_out import fnum

# ── Constants ────────────────────────────────────────────────────────────────

PLA_VISTA = {"f": "frontal", "b": "posterior", "t": "superior",
             "bo": "inferior", "r": "dreta", "l": "esquerra"}
MURS = ("f", "b", "r", "l")
ELEMENTS = ("port", "int", "corn", "crown", "led", "bra", "wall", "und", "cav",
            "roof", "eave", "plat", "slot", "sill", "jamb", "lint", "tie", "pil",
            "bbeam", "tbeam", "ebeam", "ext")
_ELEMS = "|".join(ELEMENTS)

LABEL_RE = re.compile(
    r"^(f|b|t|bo|r|l)"
    r"(?:-(ll|hh|dd))?"
    r"(?:-(" + _ELEMS + r"))?"
    r"(?:-(\d+\.\d+)([a-z])?((?:\+(?:(?:f|b|r|l)-)?\d+\.\d+)*))?"
    r"(?:-(\d+\.\d+))?"
    r"(-res|-r|-c)?$")

F_EST = ("Code", "Chunk_Label", "Frame_Mode", "N_Shapes", "N_Valid",
         "N_Invalid", "N_Legacy2D", "N_QC_Flags", "Vert_Cons_m2",
         "F_Cons_m2", "T_Int_m2", "Notes")
F_MUR = ("Code", "Wall", "Body", "Instance", "Area_Cons_m2", "N_Fragments",
         "Width_Cons_m", "Height_Cons_m", "Area_Complete_m2",
         "Area_Restituted_m2", "LL_m", "LL_Source", "HH_m", "HH_Source",
         "DD_m", "DD_Source", "Area_Orig_m2", "Orig_Source",
         "Volume_Prism_m3", "Index_Conserv_pct", "QC_Notes")
F_PORT = ("Code", "Plane", "Port", "Body", "Status", "Area_Cons_m2",
          "Area_Orig_m2", "Ratio_vs_Body_Cons_pct", "Ratio_vs_Body_Orig_pct")
F_SHP = ("Code", "Label", "Canonical", "Valid", "Geometry", "Plane",
         "Metric", "Element", "Body", "Letter", "Also", "Body2", "Status",
         "Dims", "N_Vertices", "Area_Proj_m2", "Width_m", "Height_m",
         "Area_Real_m2", "Obliquity_deg", "Len3D_m", "Len_Vert_m",
         "Len_Horiz_m", "Chord_m", "Metric_Value_m", "QC_Flags", "Suggestion")

FITXERS_CSV = {
    "estructures": ("metriques_estructures.csv", F_EST),
    "murs": ("metriques_murs.csv", F_MUR),
    "portals": ("metriques_portals.csv", F_PORT),
    "shapes": ("metriques_shapes_audit.csv", F_SHP),
}

OBLIQUITAT_PREFIX_DEG = 35.0
DIRECCIO_MAX_DEG = 30.0
DIAMETRE_MIN_DEG = 60.0
DISCREPANCIA_PROP = 0.02

# ── Gramàtica ────────────────────────────────────────────────────────────────


def suggerir(label: str) -> str | None:
    """Proposa l'etiqueta canònica per a vocabulari llegat; None si no cal."""
    s = label.strip().lower()
    s = re.sub(r"^e-", "f-", s)
    s = re.sub(r"^p-", "t-", s)
    for vell, nou in (("rooftop", "roof"), ("roofin", "roof"), ("roofout", "roof"),
                      ("cov", "cav"), ("clo", "crown"), ("sup", "plat"), ("strat", "slot"),
                      ("llind", "sill"), ("jac", "bbeam"), ("dint", "lint"), ("bran", "jamb"),
                      ("pilar", "pil"), ("pilastre", "pil")):
        s = re.sub(r"(^|-)%s(?=-|$)" % vell, r"\1%s" % nou, s)
    # back/bottom: l'intradós mira avall
    s = re.sub(r"^b-(roof|eave)(?=-|$)", r"bo-\1", s)
    s = re.sub(r"^t-roofin\b", "bo-roof", s)
    # llegat h-/l-/d- del vocabulari antic
    if re.match(r"^(h|hh)-", s):
        s = re.sub(r"^(h|hh)-", "f-hh-", s)
    if re.match(r"^(d|dd)-", s):
        s = re.sub(r"^(d|dd)-", "f-dd-", s)
    if re.match(r"^l-(%s)(?=-|$)" % _ELEMS, s):
        s = re.sub(r"^l-", "f-ll-", s)
    # pla per defecte (f-) per a etiquetes que comencen per mètrica o element
    if re.match(r"^(ll|hh|dd)-", s):
        s = "f-" + s
    elif re.match(r"^(%s)(?=-|$)" % _ELEMS, s):
        s = "f-" + s
    # cossos i sufixos
    s = re.sub(r"n(\d+\.\d+)", r"\1", s)
    s = re.sub(r"(^|-)0?(\d)(\d)($|-)", r"\g<1>\g<2>.\g<3>\g<4>", s)
    if re.match(r"^dd-wall-(f|b|r|l)-", s):
        s = re.sub(r"^dd-wall-(f|b|r|l)-", r"\1-dd-wall-", s)
    s = re.sub(r"-res((?:\+[^+]*)*)$", r"-r\1", s)
    s = re.sub(r"-(r|c)((?:\+(?:(?:f|b|r|l)-)?\d+\.\d+)+)$", r"\2-\1", s)
    return s if s != label.strip().lower() and LABEL_RE.match(s) else None


def parse_etiqueta(label: str | None):
    """Retorna ``(dict, None, None)`` o ``(None, motiu, suggeriment)``."""
    s = (label or "").strip().lower()
    m = LABEL_RE.match(s)
    if not m:
        return None, "fora de gramatica", suggerir(s)
    plane, metric, elem, body1, letter, plus, body2, rest = m.groups()
    also = []
    for it in (plus or "").split("+"):
        if not it:
            continue
        m2 = re.match(r"^(?:(f|b|r|l)-)?(\d+\.\d+)$", it)
        if m2:
            also.append((m2.group(1) or plane, m2.group(2)))
    d = dict(plane=plane, metric=metric, element=elem, body=body1,
             letter=letter, also=also, body2=body2,
             restitution=(rest in ("-r", "-res")), complete=(rest == "-c"),
             canonical=s)
    if not (elem or body1):
        return None, "cal element o cos despres del pla", None
    # Cossos X.Y amb Y >= 1 (DELTA v24→v25, A4: els '1.0' eren typos d'etiqueta)
    for cos in [body1, body2] + [b for _p, b in also]:
        if cos and cos.split(".")[1].strip("0") == "":
            return None, "cos %s: Y ha de ser >= 1" % cos, None
    if body2 and not body1:
        return None, "parell de cossos incomplet", None
    if body2 and elem not in (None, "corn", "led"):
        return None, "parell de cossos nomes per a corn/led o pla nu", None
    if also and not metric:
        return None, "la llista + nomes val per a polilinies ll/hh/dd", None
    if also and body2:
        return None, "la llista + no es combinable amb interficie", None
    if also and not rest:
        return None, "la llista + requereix estatus declarat (-c o -r)", None
    return d, None, None


# ── Geometria ────────────────────────────────────────────────────────────────


def newell(pts: np.ndarray) -> np.ndarray:
    """Vector normal (mòdul = àrea) d'un polígon 3D pel mètode de Newell."""
    p = np.asarray(pts, dtype=float)
    q = np.roll(p, -1, axis=0)
    n = np.array([
        np.sum((p[:, 1] - q[:, 1]) * (p[:, 2] + q[:, 2])),
        np.sum((p[:, 2] - q[:, 2]) * (p[:, 0] + q[:, 0])),
        np.sum((p[:, 0] - q[:, 0]) * (p[:, 1] + q[:, 1])),
    ])
    return n * 0.5


def shoelace(uv: Sequence[Sequence[float]]) -> float:
    a = np.asarray(uv, dtype=float)
    if len(a) < 3:
        return 0.0
    x, y = a[:, 0], a[:, 1]
    return float(abs(np.dot(x, np.roll(y, -1)) - np.dot(y, np.roll(x, -1))) / 2.0)


def _norm(v) -> float:
    return float(np.linalg.norm(v))


def _unit(v) -> np.ndarray:
    n = _norm(v)
    return np.asarray(v, dtype=float) / n if n > 0 else np.asarray(v, dtype=float)


def _angle_deg(a, b) -> float:
    """Angle entre dues direccions (sense signe, 0–90)."""
    c = abs(float(np.dot(_unit(a), _unit(b))))
    return math.degrees(math.acos(max(min(c, 1.0), 0.0)))


def mediana(vals: Sequence[float]) -> float | None:
    v = sorted(vals)
    n = len(v)
    if n == 0:
        return None
    return v[n // 2] if n % 2 else 0.5 * (v[n // 2 - 1] + v[n // 2])


# ── Dades d'entrada ──────────────────────────────────────────────────────────


@dataclass
class Shape:
    label: str
    tipus: str                      # "polygon" | "line"
    punts: np.ndarray | None        # (n, 3) o None si és 2D llegada
    n_vertexs: int = 0

    def __post_init__(self):
        if self.punts is not None:
            self.punts = np.asarray(self.punts, dtype=float)
            self.n_vertexs = len(self.punts)


@dataclass
class Marc:
    """Eixos de projecció per pla (``{pla: (X, Y, Z)}``), centre i vertical."""
    eixos: dict[str, tuple[np.ndarray, np.ndarray, np.ndarray]]
    centre: np.ndarray
    amunt: np.ndarray


@dataclass
class Resultat:
    code: str
    chunk_label: str
    mode: str = ""
    shapes: list[dict] = field(default_factory=list)
    murs: list[dict] = field(default_factory=list)
    ports: list[dict] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)
    n_shapes: int = 0
    n_valid: int = 0
    n_invalid: int = 0
    n_2d: int = 0
    n_qc: int = 0
    vert_cons: float = 0.0
    f_cons: float = 0.0
    t_int: float = 0.0

    def problemes(self) -> list[str]:
        out = []
        for s in self.shapes:
            if s["QC_Flags"]:
                t = "%s: %s" % (s["Label"], s["QC_Flags"])
                if s["Suggestion"]:
                    t += "  →  %s" % s["Suggestion"]
                out.append(t)
        for m in self.murs:
            if m["QC_Notes"]:
                out.append("mur %s cos %s: %s" % (m["Wall"], m["Body"], m["QC_Notes"]))
        return out


def triar_marc(shapes: Sequence[Shape], marcs: dict[str, Marc]) -> str:
    """Mode amb la mediana d'obliqüitat més baixa dels polígons de pla."""
    if len(marcs) == 1:
        return next(iter(marcs))
    med = {}
    for mode, marc in marcs.items():
        obls = []
        for sh in shapes:
            parsed, _e, _s = parse_etiqueta(sh.label)
            if parsed is None or parsed["metric"] or sh.tipus != "polygon" or sh.punts is None:
                continue
            nv = newell(sh.punts)
            if _norm(nv) < 1e-9:
                continue
            obls.append(_angle_deg(nv, marc.eixos[parsed["plane"]][2]))
        med[mode] = mediana(obls)
    valid = {m: v for m, v in med.items() if v is not None}
    return min(valid, key=valid.get) if valid else next(iter(marcs))


# ── Anàlisi ──────────────────────────────────────────────────────────────────


def analitzar(code: str, chunk_label: str, shapes: Sequence[Shape],
              marcs: dict[str, Marc], notes: Sequence[str] = ()) -> Resultat:
    """Analitza les shapes d'una estructura i retorna files + comptadors."""
    res = Resultat(code=code, chunk_label=chunk_label, notes=list(notes))
    shapes = [s for s in shapes if s.punts is None or len(s.punts) >= 2]
    res.n_shapes = len(shapes)
    if not shapes:
        res.notes.append("sense shapes")
        return res

    res.mode = triar_marc(shapes, marcs)
    marc = marcs[res.mode]
    axes, center, UP = marc.eixos, marc.centre, marc.amunt

    # pre-pas: polígons de mur (centroide + normal) per al QC
    wall_refs: dict[str, list[tuple[np.ndarray, np.ndarray | None]]] = {}
    for sh in shapes:
        parsed, _e, _s = parse_etiqueta(sh.label)
        if (parsed is None or parsed["metric"] or parsed["element"] is not None
                or parsed["plane"] not in MURS or sh.tipus != "polygon" or sh.punts is None):
            continue
        nv = newell(sh.punts)
        nrm = _unit(nv) if _norm(nv) > 1e-9 else None
        wall_refs.setdefault(parsed["plane"], []).append((sh.punts.mean(axis=0), nrm))

    walls_cons: dict = {}; walls_c: dict = {}; walls_r: dict = {}; walls_frag: dict = {}  # noqa: E702
    walls_uv: dict = {}
    ports_cons: dict = {}; ports_o: dict = {}; ports_meta: dict = {}  # noqa: E702
    rest: dict = {}; comp: dict = {}; prop: dict = {}  # noqa: E702

    for sh in shapes:
        parsed, err, sug = parse_etiqueta(sh.label)
        row = {k: "" for k in F_SHP}
        row["Code"] = code
        row["Label"] = sh.label
        flags: list[str] = []

        if parsed is None:
            res.n_invalid += 1
            row["Valid"] = "no"
            row["QC_Flags"] = "INVALIDA:" + (err or "")
            row["Suggestion"] = sug or ""
            res.shapes.append(row)
            continue

        row["Canonical"] = parsed["canonical"]
        sugv = suggerir(sh.label)
        if sugv and sugv != parsed["canonical"]:
            row["Suggestion"] = sugv
        row["Plane"] = parsed["plane"]
        row["Metric"] = parsed["metric"] or ""
        row["Element"] = parsed["element"] or ""
        row["Body"] = parsed["body"] or ""
        row["Letter"] = parsed["letter"] or ""
        row["Also"] = "+".join("%s-%s" % pb for pb in parsed["also"])
        row["Body2"] = parsed["body2"] or ""
        row["Status"] = "-c" if parsed["complete"] else ("-r" if parsed["restitution"] else "cons")

        if parsed["element"] == "port" and not parsed["body"]:
            flags.append("PORT_SENSE_COS")

        pts = sh.punts
        if pts is None:
            res.n_2d += 1
            row["Dims"] = "2D"
            row["Valid"] = "no"
            flags.append("2D_LLEGAT")
            row["QC_Flags"] = ";".join(flags)
            res.shapes.append(row)
            continue

        row["Dims"] = "3D"
        row["N_Vertices"] = str(len(pts))
        res.n_valid += 1
        row["Valid"] = "si"

        is_line = sh.tipus == "line"
        row["Geometry"] = "line" if is_line else "polygon"
        if parsed["metric"] and not is_line:
            flags.append("METRICA_EN_POLIGON")
        if not parsed["metric"] and is_line:
            flags.append("LINIA_SENSE_TOKEN")

        if is_line:
            segs = np.diff(pts, axis=0)
            tot = float(np.sum(np.linalg.norm(segs, axis=1)))
            vert = float(np.sum(np.abs(segs @ UP)))
            horiz = math.sqrt(max(tot * tot - vert * vert, 0.0))
            chord = _norm(pts[-1] - pts[0])
            row["Len3D_m"] = fnum(tot)
            row["Len_Vert_m"] = fnum(vert)
            row["Len_Horiz_m"] = fnum(horiz)
            row["Chord_m"] = fnum(chord)
            mv = {"ll": tot, "hh": vert, "dd": chord}.get(parsed["metric"], tot)
            row["Metric_Value_m"] = fnum(mv)
            if parsed["metric"] == "dd" and len(pts) > 2:
                flags.append("DD_MULTITRAM")

            # QC orfandat: línia de mur amb prefix no-mur
            if parsed["element"] == "wall" and parsed["plane"] not in MURS:
                flags.append("ORFE_MUR")

            # QC pertinença
            if parsed["plane"] in MURS and wall_refs:
                mid = pts.mean(axis=0)
                dists = {pf: min(_norm(mid - c) for c, _n in refs)
                         for pf, refs in wall_refs.items()}
                if parsed["plane"] in dists:
                    best = min(dists, key=dists.get)
                    if best != parsed["plane"] and dists[parsed["plane"]] > 1.5 * dists[best] + 0.10:
                        flags.append("PERTINENCA:%s" % best)

            # QC direcció, conscient de la família de l'element:
            #  - paral·lels a la cara (wall, lint, ebeam, corn, led, sill, tie):
            #    ll corre per la cara, dd la travessa (gruix)
            #  - projectants (bra, tbeam, bbeam): ll ix per la normal de la
            #    cara, dd és el diàmetre (perpendicular a l'eix)
            if parsed["metric"] and chord > 1e-9:
                dirv = _unit(pts[-1] - pts[0])
                projecting = parsed["element"] in ("bra", "tbeam", "bbeam")
                expected = None
                perp_to = None
                if parsed["metric"] == "hh":
                    expected = UP
                elif parsed["plane"] in MURS:
                    nrm = axes[parsed["plane"]][2]
                    refs = wall_refs.get(parsed["plane"], [])
                    if refs:
                        mid2 = pts.mean(axis=0)
                        refs = sorted(refs, key=lambda cn: _norm(mid2 - cn[0]))
                        if refs[0][1] is not None:
                            nrm = refs[0][1]
                    if parsed["metric"] == "dd":
                        if projecting:
                            perp_to = nrm
                        else:
                            expected = nrm
                    elif parsed["metric"] == "ll":
                        expected = nrm if projecting else axes[parsed["plane"]][0]
                if expected is not None:
                    ang = _angle_deg(dirv, expected)
                    if ang > DIRECCIO_MAX_DEG:
                        flags.append("DIRECCIO:%.0f" % ang)
                elif perp_to is not None:
                    ang = _angle_deg(dirv, perp_to)
                    if ang < DIAMETRE_MIN_DEG:
                        flags.append("DIRECCIO:%.0f" % ang)

            if parsed["body"] and parsed["metric"] and parsed["plane"] in MURS:
                key = (parsed["plane"], parsed["body"], parsed["letter"] or "")
                if parsed["restitution"]:
                    rest.setdefault(key, {})[parsed["metric"]] = mv
                elif parsed["complete"]:
                    comp.setdefault(key, {})[parsed["metric"]] = mv
                suf = "-c" if parsed["complete"] else ("-r" if parsed["restitution"] else "")
                anchor = "%s-%s" % (parsed["plane"], parsed["body"])
                for p2, b2 in parsed["also"]:
                    prop.setdefault((p2, b2, ""), {})[parsed["metric"]] = (mv, anchor, suf)

            row["QC_Flags"] = ";".join(flags)
            res.n_qc += len(flags)
            res.shapes.append(row)
            continue

        # polígon
        AXp, AYp, AZp = axes[parsed["plane"]]
        rel = pts - center
        uv = np.column_stack([rel @ AXp, rel @ AYp])
        area_proj = shoelace(uv)
        w = float(uv[:, 0].max() - uv[:, 0].min())
        h = float(uv[:, 1].max() - uv[:, 1].min())
        row["Area_Proj_m2"] = fnum(area_proj)
        row["Width_m"] = fnum(w)
        row["Height_m"] = fnum(h)
        nv = newell(pts)
        if _norm(nv) > 1e-9:
            nvn = _unit(nv)
            obl = _angle_deg(nvn, AZp)
            best = max(PLA_VISTA, key=lambda k: abs(float(np.dot(nvn, axes[k][2]))))
            row["Area_Real_m2"] = fnum(_norm(nv))
            row["Obliquity_deg"] = fnum(obl, 1)
            if obl > OBLIQUITAT_PREFIX_DEG and best != parsed["plane"]:
                flags.append("PREFIX_DUBTOS:%s" % best)

        if parsed["element"] is None and parsed["body"] and parsed["plane"] in MURS:
            key = (parsed["plane"], parsed["body"], parsed["letter"] or "")
            if parsed["restitution"]:
                walls_r[key] = walls_r.get(key, 0.0) + area_proj
            else:
                walls_cons[key] = walls_cons.get(key, 0.0) + area_proj
                walls_frag[key] = walls_frag.get(key, 0) + 1
                walls_uv.setdefault(key, []).extend(uv.tolist())
                if parsed["complete"]:
                    walls_c[key] = walls_c.get(key, 0.0) + area_proj
        elif parsed["element"] == "port":
            pk = (parsed["plane"], (parsed["body"] or "?") + (parsed["letter"] or ""))
            ports_meta[pk] = dict(body=parsed["body"] or "", status=row["Status"])
            if not parsed["restitution"]:
                ports_cons[pk] = ports_cons.get(pk, 0.0) + area_proj
            if parsed["complete"] or parsed["restitution"]:
                ports_o[pk] = ports_o.get(pk, 0.0) + area_proj
        elif parsed["element"] == "int" and parsed["plane"] == "t" and not parsed["restitution"]:
            res.t_int += area_proj

        row["QC_Flags"] = ";".join(flags)
        res.n_qc += len(flags)
        res.shapes.append(row)

    # derivades per (mur, cos)
    keys = sorted(set(list(rest) + list(comp) + list(prop) +
                      list(walls_cons) + list(walls_c) + list(walls_r)))
    for key in keys:
        mur, body, inst = key
        mrow = {k: "" for k in F_MUR}
        mrow["Code"] = code
        mrow["Wall"] = mur
        mrow["Body"] = body
        mrow["Instance"] = inst
        qc: list[str] = []

        if key in walls_cons:
            mrow["Area_Cons_m2"] = fnum(walls_cons[key])
            mrow["N_Fragments"] = str(walls_frag.get(key, 0))
            uvs = np.asarray(walls_uv.get(key, []))
            if len(uvs):
                mrow["Width_Cons_m"] = fnum(float(uvs[:, 0].max() - uvs[:, 0].min()))
                mrow["Height_Cons_m"] = fnum(float(uvs[:, 1].max() - uvs[:, 1].min()))
        if key in walls_c:
            mrow["Area_Complete_m2"] = fnum(walls_c[key])
        if key in walls_r:
            mrow["Area_Restituted_m2"] = fnum(walls_r[key])

        dims: dict[str, float] = {}
        src_d: dict[str, str] = {}
        for m in ("ll", "hh", "dd"):
            in_r = m in rest.get(key, {})
            in_c = m in comp.get(key, {})
            in_p = m in prop.get(key, {})
            if in_r and in_c:
                qc.append("CONFLICTE_%s_RES_C" % m)
            if in_c:
                dims[m] = comp[key][m]; src_d[m] = "-c"  # noqa: E702
            elif in_r:
                dims[m] = rest[key][m]; src_d[m] = "-r"  # noqa: E702
            elif in_p:
                v, anc, suf = prop[key][m]
                dims[m] = v; src_d[m] = "prop:%s%s" % (anc, suf)  # noqa: E702
            if in_p and (in_c or in_r):
                v_dir = dims[m]
                v_pro = prop[key][m][0]
                if v_dir > 1e-9 and abs(v_dir - v_pro) / v_dir > DISCREPANCIA_PROP:
                    qc.append("DISCREPANCIA_%s_prop:%s" % (m, prop[key][m][1]))
            if m in dims:
                mrow[m.upper() + "_m"] = fnum(dims[m])
                mrow[m.upper() + "_Source"] = src_d[m]

        area_o = None
        if key in walls_c:
            area_o = walls_c[key]
            mrow["Orig_Source"] = "poligon-c"
        elif key in walls_r:
            area_o = walls_r[key]
            mrow["Orig_Source"] = "poligon-r"
        elif "ll" in dims and "hh" in dims:
            area_o = dims["ll"] * dims["hh"]
            mrow["Orig_Source"] = "rectangle(ll:%s,hh:%s)" % (src_d["ll"], src_d["hh"])
        if area_o is not None:
            mrow["Area_Orig_m2"] = fnum(area_o)
            if "dd" in dims:
                mrow["Volume_Prism_m3"] = fnum(area_o * dims["dd"])
            if key in walls_cons and area_o > 1e-9:
                mrow["Index_Conserv_pct"] = fnum(100.0 * walls_cons[key] / area_o, 1)
        mrow["QC_Notes"] = ";".join(qc)
        res.n_qc += len(qc)
        res.murs.append(mrow)

    # portals
    for pk in sorted(set(list(ports_cons) + list(ports_o))):
        plane, port = pk
        meta = ports_meta.get(pk, {})
        body = meta.get("body", "")
        prow = {k: "" for k in F_PORT}
        prow["Code"] = code
        prow["Plane"] = plane
        prow["Port"] = port
        prow["Body"] = body
        prow["Status"] = meta.get("status", "")
        a_c = ports_cons.get(pk)
        a_o = ports_o.get(pk)
        if a_c is not None:
            prow["Area_Cons_m2"] = fnum(a_c)
        if a_o is not None:
            prow["Area_Orig_m2"] = fnum(a_o)
        if plane == "f" and body:
            wkey = ("f", body, "")
            if a_c is not None and wkey in walls_cons and walls_cons[wkey] > 1e-9:
                prow["Ratio_vs_Body_Cons_pct"] = fnum(100.0 * a_c / walls_cons[wkey], 1)
            worig = walls_c.get(wkey) or walls_r.get(wkey)
            if worig is None:
                d = dict(comp.get(wkey, {}))
                r2 = rest.get(wkey, {})
                p2 = prop.get(wkey, {})
                ll = d.get("ll") or r2.get("ll") or (p2.get("ll", (None,))[0])
                hh = d.get("hh") or r2.get("hh") or (p2.get("hh", (None,))[0])
                if ll and hh:
                    worig = ll * hh
            if a_o is not None and worig:
                prow["Ratio_vs_Body_Orig_pct"] = fnum(100.0 * a_o / worig, 1)
        res.ports.append(prow)

    res.vert_cons = sum(walls_cons.values())
    res.f_cons = sum(a for (m, _b, _l), a in walls_cons.items() if m == "f")
    return res


# ── Files dels 4 CSV ─────────────────────────────────────────────────────────


def files_exportacio(resultats: Sequence[Resultat]) -> dict[str, list[dict]]:
    """``{"estructures": [...], "murs": [...], "portals": [...], "shapes": [...]}``"""
    est, murs, ports, shp = [], [], [], []
    for r in resultats:
        est.append({
            "Code": r.code, "Chunk_Label": r.chunk_label, "Frame_Mode": r.mode,
            "N_Shapes": r.n_shapes, "N_Valid": r.n_valid, "N_Invalid": r.n_invalid,
            "N_Legacy2D": r.n_2d, "N_QC_Flags": r.n_qc,
            "Vert_Cons_m2": fnum(r.vert_cons) if r.vert_cons else "",
            "F_Cons_m2": fnum(r.f_cons) if r.f_cons else "",
            "T_Int_m2": fnum(r.t_int) if r.t_int else "",
            "Notes": "; ".join(r.notes)})
        murs.extend(r.murs)
        ports.extend(r.ports)
        shp.extend(r.shapes)
    return {"estructures": est, "murs": murs, "portals": ports, "shapes": shp}


def exportar(resultats: Sequence[Resultat], carpeta) -> dict[str, tuple[str, int]]:
    """Escriu els 4 CSV a ``carpeta``. Retorna ``{clau: (ruta, n_files)}``."""
    from chachapoya_bd.db.csv_out import escriure_csv
    from pathlib import Path
    out = {}
    files = files_exportacio(resultats)
    for clau, (nom, camps) in FITXERS_CSV.items():
        p = escriure_csv(Path(carpeta) / nom, camps, files[clau])
        out[clau] = (str(p), len(files[clau]))
    return out
