# -*- coding: utf-8 -*-
"""
export_metriques_shapes.py — Agisoft Metashape Professional 2.3.0
────────────────────────────────────────────────────────────────────────────
EXTRACTOR DEFINITIU de mètriques des de shapes 3D dibuixades sobre el model.
Consolida la convenció validada en la fase de proves (tests v1–v7):

  GRAMÀTICA:  PLA[-MÈTRICA]-[ELEMENT]-COS[+COS...][a-z][-COS][-res|-c]
    Plans:     f b t bo r l   (= PLANE_AXES de batch_orto_bbox.py)
    Mètriques: ll (longitud 3D, multi-tram) | hh (component vertical)
               dd (corda 3D absoluta, UN sol tram)
    Elements:  port int corn crown led bra wall und cav roof eave plat
               slot sill jamb lint tie pil bbeam tbeam ebeam ext
    Estatus:   (cap) conservat | -c complet | -r reconstruït (-res acceptat)
    Relacions: -COS interfície | +COS aplicabilitat (mateix mur, entre cossos)
    Lletres:   INSTANCIES (2n portal, 3a mensula, 2n mur coplanari);
               els FRAGMENTS d'una mateixa cara repeteixen l'etiqueta

  SEMÀNTICA DEL PREFIX: polígons → pla de projecció (vigilat per la normal);
  polilínies → pertinença al mur (vigilada per proximitat als polígons).

  JERARQUIA D'EVIDÈNCIA: -c directa > -res directa > propagada (+), amb
  procedència declarada a cada valor derivat.

  SIS CONTROLS DE QUALITAT: obliqüitat/prefix (polígons), pertinença
  (línies vs polígons de mur), direcció (línies; dd contra la normal REAL
  del mur), conflicte -res∧-c, discrepància directa/propagada (>2%),
  orfandat (línia de mur amb prefix no-mur). A més: etiquetes invàlides
  amb proposta de canònic, i exclusió de shapes 2D llegades d'orthos.

SORTIDA (4 CSV al directori triat, una passada per tot el corpus marcat):
  metriques_estructures.csv   una fila per estructura (chunk)
  metriques_murs.csv          una fila per (estructura, mur, cos)
  metriques_portals.csv       una fila per portal, amb ràtios _cons i _orig
  metriques_shapes_audit.csv  una fila per shape: tot el calculat + QC

Els chunks candidats són els que tenen shapes; es preseleccionen els
activats (enabled). El codi de l'estructura és l'etiqueta del chunk.
L'script NO modifica el projecte: és de només lectura.

Execució: Tools → Run Script. Menú: vegeu el bloc final.
────────────────────────────────────────────────────────────────────────────
"""

import csv
import datetime
import math
import os
import re

import Metashape

try:
    from PySide2 import QtCore, QtWidgets
except ImportError:
    from PySide6 import QtCore, QtWidgets  # type: ignore

# ---------------------------------------------------------------- constants

PLANE_AXIS = {"f": "-Y", "b": "Y", "t": "Z", "bo": "-Z", "r": "X", "l": "-X"}
WALLS = ("f", "b", "r", "l")

LABEL_RE = re.compile(
    r"^(f|b|t|bo|r|l)"
    r"(?:-(ll|hh|dd))?"
    r"(?:-(port|int|corn|crown|led|bra|wall|und|cav|roof|eave|plat|slot|sill|jamb|lint|tie|pil|bbeam|tbeam|ebeam|ext))?"
    r"(?:-(\d+\.\d+)([a-z])?((?:\+(?:(?:f|b|r|l)-)?\d+\.\d+)*))?"
    r"(?:-(\d+\.\d+))?"
    r"(-res|-r|-c)?$")

CODE_RE = re.compile(r"^((LP|DW)-[A-Z0-9]+-EA\d+)(?:-(.+))?$")

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


def log(msg):
    print("[exportMetriques] " + str(msg))


def fnum(v, dec=3):
    return ("%." + str(dec) + "f") % v if v is not None else ""


# ---------------------------------------------------------------- gramàtica

def suggest(label):
    s = label.strip().lower()
    ELEMS = ("port|int|corn|crown|led|bra|wall|und|cav|roof|eave|plat|"
             "slot|sill|jamb|lint|tie|pil|bbeam|tbeam|ebeam|ext")
    # vocabulari llegat i variants (amb o sense pla davant)
    s = re.sub(r"^e-", "f-", s)
    s = re.sub(r"^p-", "t-", s)
    s = re.sub(r"(^|-)rooftop(?=-|$)", r"\1roof", s)
    s = re.sub(r"(^|-)roofin(?=-|$)", r"\1roof", s)
    s = re.sub(r"(^|-)roofout(?=-|$)", r"\1roof", s)
    s = re.sub(r"(^|-)cov(?=-|$)", r"\1cav", s)
    s = re.sub(r"(^|-)clo(?=-|$)", r"\1crown", s)
    s = re.sub(r"(^|-)sup(?=-|$)", r"\1plat", s)
    s = re.sub(r"(^|-)strat(?=-|$)", r"\1slot", s)
    s = re.sub(r"(^|-)llind(?=-|$)", r"\1sill", s)
    s = re.sub(r"(^|-)jac(?=-|$)", r"\1bbeam", s)
    s = re.sub(r"(^|-)dint(?=-|$)", r"\1lint", s)
    s = re.sub(r"(^|-)bran(?=-|$)", r"\1jamb", s)
    s = re.sub(r"(^|-)pilar(?=-|$)", r"\1pil", s)
    s = re.sub(r"(^|-)pilastre(?=-|$)", r"\1pil", s)
    # back/bottom: l'intrados mira avall
    s = re.sub(r"^b-(roof|eave)(?=-|$)", r"bo-\1", s)
    s = re.sub(r"^t-roofin\b", "bo-roof", s)
    # llegat h-/l-/d- del vocabulari antic
    if re.match(r"^(h|hh)-", s):
        s = re.sub(r"^(h|hh)-", "f-hh-", s)
    if re.match(r"^(d|dd)-", s):
        s = re.sub(r"^(d|dd)-", "f-dd-", s)
    if re.match(r"^l-(%s)(?=-|$)" % ELEMS, s):
        s = re.sub(r"^l-", "f-ll-", s)
    # pla per defecte (f-) per a etiquetes que comencen per metrica o element
    if re.match(r"^(ll|hh|dd)-", s):
        s = "f-" + s
    elif re.match(r"^(%s)(?=-|$)" % ELEMS, s):
        s = "f-" + s
    # cossos i sufixos
    s = re.sub(r"n(\d+\.\d+)", r"\1", s)
    s = re.sub(r"(^|-)0?(\d)(\d)($|-)", r"\g<1>\g<2>.\g<3>\g<4>", s)
    if re.match(r"^dd-wall-(f|b|r|l)-", s):
        s = re.sub(r"^dd-wall-(f|b|r|l)-", r"\1-dd-wall-", s)
    s = re.sub(r"-res((?:\+[^+]*)*)$", r"-r\1", s)
    s = re.sub(r"-(r|c)((?:\+(?:(?:f|b|r|l)-)?\d+\.\d+)+)$", r"\2-\1", s)
    return s if s != label.strip().lower() and LABEL_RE.match(s) else None


def parse_label(label):
    s = (label or "").strip().lower()
    m = LABEL_RE.match(s)
    if not m:
        return None, "fora de gramatica", suggest(s)
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


# ---------------------------------------------------------------- geometria

def ortho_axes(rot, axis):
    rx, ry, rz = rot.col(0), rot.col(1), rot.col(2)
    if axis == "Z":
        return rx, ry, rz
    if axis == "-Z":
        return rx * -1, ry, rz * -1
    if axis == "Y":
        return rx, rz, ry * -1
    if axis == "-Y":
        return rx * -1, rz, ry
    if axis == "X":
        return ry, rz, rx
    if axis == "-X":
        return ry * -1, rz, rx * -1
    raise ValueError(axis)


def newell(pts):
    n = Metashape.Vector([0.0, 0.0, 0.0])
    k = len(pts)
    for i in range(k):
        p, q = pts[i], pts[(i + 1) % k]
        n.x += (p.y - q.y) * (p.z + q.z)
        n.y += (p.z - q.z) * (p.x + q.x)
        n.z += (p.x - q.x) * (p.y + q.y)
    return n * 0.5


def shoelace(uv):
    a = 0.0
    k = len(uv)
    for i in range(k):
        x1, y1 = uv[i]
        x2, y2 = uv[(i + 1) % k]
        a += x1 * y2 - x2 * y1
    return abs(a) / 2.0


def vmean(pts):
    m = Metashape.Vector([0.0, 0.0, 0.0])
    for p in pts:
        m += p
    return m / len(pts)


def median(vals):
    v = sorted(vals)
    n = len(v)
    if n == 0:
        return None
    return v[n // 2] if n % 2 else 0.5 * (v[n // 2 - 1] + v[n // 2])


def build_frame(chunk, mode, geo):
    R = chunk.transform.rotation
    # LA VERTICAL DEPEN DEL CONTEXT DEL CHUNK, no del mode de les
    # shapes: en un chunk GEOREFERENCIAT el mon es geocentric i
    # (0,0,1) apunta a l'eix polar, no amunt (a Leymebamba, ~87
    # graus d'error - el bug d'EA40, dues hh a 88). La vertical
    # del mon es deriva sempre del localframe quan el centre viu
    # en coordenades geocentriques (norma > 1000 km).
    center_w = chunk.transform.matrix.mulp(chunk.region.center)
    if center_w.norm() > 1000000.0:
        UPw = ((geo.localframe(center_w).rotation().t()) * Metashape.Vector([0, 0, 1])).normalized()
    else:
        UPw = Metashape.Vector([0, 0, 1])
    if mode in ("geografic", "mon-local"):
        axes = {}
        for pfx, name in PLANE_AXIS.items():
            a1, a2, a3 = ortho_axes(chunk.region.rot, name)
            axes[pfx] = ((R * a1).normalized(), (R * a2).normalized(),
                         (R * a3).normalized())
        return axes, center_w, UPw
    if mode == "intern":
        s = chunk.transform.scale if chunk.transform.scale else 1.0
        center = chunk.region.center * s
        axes = {}
        for pfx, name in PLANE_AXIS.items():
            a1, a2, a3 = ortho_axes(chunk.region.rot, name)
            axes[pfx] = (a1.normalized(), a2.normalized(), a3.normalized())
        UP = (R.t() * UPw).normalized()
        return axes, center, UP
    raise ValueError(mode)


# ---------------------------------------------------------------- anàlisi

def analyze_chunk(chunk, geo):
    """Analitza un chunk i retorna el paquet de files + comptadors."""
    res = dict(code=chunk.label, chunk_label=chunk.label, mode="",
               shapes=[], murs=[], ports=[], notes=[],
               n_shapes=0, n_valid=0, n_invalid=0, n_2d=0, n_qc=0,
               vert_cons=0.0, f_cons=0.0, t_int=0.0)

    cm = CODE_RE.match((chunk.label or "").strip())
    if cm:
        res["code"] = cm.group(1)
        if cm.group(3):
            res["notes"].append("chunk variant '%s': codi %s" % (chunk.label, cm.group(1)))
    else:
        res["notes"].append("etiqueta de chunk fora de la gramatica de codis")

    if chunk.shapes is None or len(chunk.shapes) == 0:
        res["notes"].append("sense shapes")
        return res

    items = []
    allc = []
    for sh in chunk.shapes:
        g = sh.geometry
        label = sh.label or "(sense etiqueta)"
        try:
            cs = list(g.coordinates[0]) if g.type == Metashape.Geometry.Type.PolygonType else list(g.coordinates)
        except Exception as exc:
            res["notes"].append("[%s] error de geometria: %s" % (label, exc))
            continue
        if len(cs) >= 2:
            items.append((label, g.type, cs))
            allc.extend(cs)
    res["n_shapes"] = len(items)
    if not items:
        return res

    xs = [c.x for c in allc]; ys = [c.y for c in allc]
    spread = max(max(xs) - min(xs), max(ys) - min(ys))
    geographic = (max(map(abs, xs)) <= 180 and max(map(abs, ys)) <= 90 and spread < 0.01)
    scrs = chunk.shapes.crs

    def pts3d(cs):
        out = []
        for c in cs:
            if c.size < 3:
                return None
            if geographic:
                out.append(scrs.unproject(Metashape.Vector([c.x, c.y, c.z])))
            else:
                out.append(Metashape.Vector([c.x, c.y, c.z]))
        return out

    # -- marc --
    if geographic:
        mode = "geografic"
        frames = {mode: build_frame(chunk, mode, geo)}
    else:
        frames = {m: build_frame(chunk, m, geo) for m in ("mon-local", "intern")}
        med = {}
        for m, (axes_m, _c, _u) in frames.items():
            obls = []
            for label, gtype, cs in items:
                parsed, err, _ = parse_label(label)
                if parsed is None or parsed["metric"] or gtype != Metashape.Geometry.Type.PolygonType:
                    continue
                pts = pts3d(cs)
                if pts is None:
                    continue
                nv = newell(pts)
                if nv.norm() < 1e-9:
                    continue
                cosang = abs(nv.normalized() * axes_m[parsed["plane"]][2])
                obls.append(math.degrees(math.acos(max(min(cosang, 1.0), 0.0))))
            med[m] = median(obls)
        valid = {m: v for m, v in med.items() if v is not None}
        mode = min(valid, key=valid.get) if valid else "mon-local"
    res["mode"] = mode
    axes, center, UP = frames[mode]

    # -- pre-pas: poligons de mur (centroide + normal) per al QC --
    wall_refs = {}
    for label, gtype, cs in items:
        parsed, err, _ = parse_label(label)
        if (parsed is None or parsed["metric"] or parsed["element"] is not None
                or parsed["plane"] not in WALLS
                or gtype != Metashape.Geometry.Type.PolygonType):
            continue
        pts = pts3d(cs)
        if pts is not None:
            nv = newell(pts)
            nrm = nv.normalized() if nv.norm() > 1e-9 else None
            wall_refs.setdefault(parsed["plane"], []).append((vmean(pts), nrm))

    walls_cons = {}; walls_c = {}; walls_r = {}; walls_frag = {}
    walls_uv = {}
    ports_cons = {}; ports_o = {}; ports_meta = {}
    rest = {}; comp = {}; prop = {}

    # -- passada principal --
    for label, gtype, cs in items:
        parsed, err, sug = parse_label(label)
        row = {k: "" for k in F_SHP}
        row["Code"] = res["code"]
        row["Label"] = label
        flags = []

        if parsed is None:
            res["n_invalid"] += 1
            row["Valid"] = "no"
            row["QC_Flags"] = "INVALIDA:" + err
            row["Suggestion"] = sug or ""
            res["shapes"].append(row)
            continue

        row["Canonical"] = parsed["canonical"]
        sugv = suggest(label)
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

        pts = pts3d(cs)
        if pts is None:
            res["n_2d"] += 1
            row["Dims"] = "2D"
            row["Valid"] = "no"
            flags.append("2D_LLEGAT")
            row["QC_Flags"] = ";".join(flags)
            res["shapes"].append(row)
            continue

        row["Dims"] = "3D"
        row["N_Vertices"] = str(len(pts))
        res["n_valid"] += 1
        row["Valid"] = "si"

        is_line = (gtype == Metashape.Geometry.Type.LineStringType)
        row["Geometry"] = "line" if is_line else "polygon"
        if parsed["metric"] and not is_line:
            flags.append("METRICA_EN_POLIGON")
        if not parsed["metric"] and is_line:
            flags.append("LINIA_SENSE_TOKEN")

        if is_line:
            tot = sum((pts[i + 1] - pts[i]).norm() for i in range(len(pts) - 1))
            vert = sum(abs((pts[i + 1] - pts[i]) * UP) for i in range(len(pts) - 1))
            horiz = math.sqrt(max(tot * tot - vert * vert, 0.0))
            chord = (pts[-1] - pts[0]).norm()
            row["Len3D_m"] = fnum(tot)
            row["Len_Vert_m"] = fnum(vert)
            row["Len_Horiz_m"] = fnum(horiz)
            row["Chord_m"] = fnum(chord)
            mv = {"ll": tot, "hh": vert, "dd": chord}.get(parsed["metric"], tot)
            row["Metric_Value_m"] = fnum(mv)
            if parsed["metric"] == "dd" and len(pts) > 2:
                flags.append("DD_MULTITRAM")

            # QC orfandat: linia de mur amb prefix no-mur
            if parsed["element"] == "wall" and parsed["plane"] not in WALLS:
                flags.append("ORFE_MUR")

            # QC pertinenca
            if parsed["plane"] in WALLS and wall_refs:
                mid = vmean(pts)
                dists = {pf: min((mid - c).norm() for c, _n in refs)
                         for pf, refs in wall_refs.items()}
                if parsed["plane"] in dists:
                    best = min(dists, key=dists.get)
                    if best != parsed["plane"] and dists[parsed["plane"]] > 1.5 * dists[best] + 0.10:
                        flags.append("PERTINENCA:%s" % best)

            # QC direccio, conscient de la familia de l'element:
            #  - paral.lels a la cara (wall, lint, ebeam, corn, led, sill,
            #    tie): ll corre per la cara, dd la travessa (gruix)
            #  - PROJECTANTS (bra, tbeam, bbeam): ll ix per la NORMAL de
            #    la cara, dd es el DIAMETRE (perpendicular a l'eix)
            if parsed["metric"] and (pts[-1] - pts[0]).norm() > 1e-9:
                dirv = (pts[-1] - pts[0]).normalized()
                projecting = parsed["element"] in ("bra", "tbeam", "bbeam")
                expected = None
                perp_to = None
                if parsed["metric"] == "hh":
                    expected = UP
                elif parsed["plane"] in WALLS:
                    nrm = axes[parsed["plane"]][2]
                    refs = wall_refs.get(parsed["plane"], [])
                    if refs:
                        mid2 = vmean(pts)
                        refs = sorted(refs, key=lambda cn: (mid2 - cn[0]).norm())
                        if refs[0][1] is not None:
                            nrm = refs[0][1]
                    if parsed["metric"] == "dd":
                        if projecting:
                            perp_to = nrm
                        else:
                            expected = nrm
                    elif parsed["metric"] == "ll":
                        if projecting:
                            expected = nrm
                        else:
                            expected = axes[parsed["plane"]][0]
                if expected is not None:
                    ang = math.degrees(math.acos(max(min(abs(dirv * expected), 1.0), 0.0)))
                    if ang > 30.0:
                        flags.append("DIRECCIO:%.0f" % ang)
                elif perp_to is not None:
                    ang = math.degrees(math.acos(max(min(abs(dirv * perp_to), 1.0), 0.0)))
                    if ang < 60.0:
                        flags.append("DIRECCIO:%.0f" % ang)

            if parsed["body"] and parsed["metric"] and parsed["plane"] in WALLS:
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
            res["n_qc"] += len(flags)
            res["shapes"].append(row)
            continue

        # -- poligon --
        AXp, AYp, AZp = axes[parsed["plane"]]
        uv = [((p - center) * AXp, (p - center) * AYp) for p in pts]
        area_proj = shoelace(uv)
        w = max(u for u, v in uv) - min(u for u, v in uv)
        h = max(v for u, v in uv) - min(v for u, v in uv)
        row["Area_Proj_m2"] = fnum(area_proj)
        row["Width_m"] = fnum(w)
        row["Height_m"] = fnum(h)
        nv = newell(pts)
        if nv.norm() > 1e-9:
            nvn = nv.normalized()
            obl = math.degrees(math.acos(max(min(abs(nvn * AZp), 1.0), 0.0)))
            best = max(PLANE_AXIS, key=lambda k: abs(nvn * axes[k][2]))
            row["Area_Real_m2"] = fnum(nv.norm())
            row["Obliquity_deg"] = fnum(obl, 1)
            if obl > 35 and best != parsed["plane"]:
                flags.append("PREFIX_DUBTOS:%s" % best)

        if parsed["element"] is None and parsed["body"] and parsed["plane"] in WALLS:
            key = (parsed["plane"], parsed["body"], parsed["letter"] or "")
            if parsed["restitution"]:
                walls_r[key] = walls_r.get(key, 0.0) + area_proj
            else:
                walls_cons[key] = walls_cons.get(key, 0.0) + area_proj
                walls_frag[key] = walls_frag.get(key, 0) + 1
                pu = walls_uv.setdefault(key, [])
                pu.extend(uv)
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
            res["t_int"] += area_proj

        row["QC_Flags"] = ";".join(flags)
        res["n_qc"] += len(flags)
        res["shapes"].append(row)

    # -- derivades per (mur, cos) --
    keys = sorted(set(list(rest) + list(comp) + list(prop) +
                      list(walls_cons) + list(walls_c) + list(walls_r)))
    for key in keys:
        mur, body, inst = key
        mrow = {k: "" for k in F_MUR}
        mrow["Code"] = res["code"]
        mrow["Wall"] = mur
        mrow["Body"] = body
        mrow["Instance"] = inst
        qc = []

        if key in walls_cons:
            mrow["Area_Cons_m2"] = fnum(walls_cons[key])
            mrow["N_Fragments"] = str(walls_frag.get(key, 0))
            uvs = walls_uv.get(key, [])
            if uvs:
                mrow["Width_Cons_m"] = fnum(max(u for u, v in uvs) - min(u for u, v in uvs))
                mrow["Height_Cons_m"] = fnum(max(v for u, v in uvs) - min(v for u, v in uvs))
        if key in walls_c:
            mrow["Area_Complete_m2"] = fnum(walls_c[key])
        if key in walls_r:
            mrow["Area_Restituted_m2"] = fnum(walls_r[key])

        dims = {}
        src_d = {}
        for m in ("ll", "hh", "dd"):
            in_r = m in rest.get(key, {})
            in_c = m in comp.get(key, {})
            in_p = m in prop.get(key, {})
            if in_r and in_c:
                qc.append("CONFLICTE_%s_RES_C" % m)
            if in_c:
                dims[m] = comp[key][m]; src_d[m] = "-c"
            elif in_r:
                dims[m] = rest[key][m]; src_d[m] = "-r"
            elif in_p:
                v, anc, suf = prop[key][m]
                dims[m] = v; src_d[m] = "prop:%s%s" % (anc, suf)
            if in_p and (in_c or in_r):
                v_dir = dims[m]
                v_pro = prop[key][m][0]
                if v_dir > 1e-9 and abs(v_dir - v_pro) / v_dir > 0.02:
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
        res["n_qc"] += len(qc)
        res["murs"].append(mrow)

    # -- portals --
    for pk in sorted(set(list(ports_cons) + list(ports_o))):
        plane, port = pk
        meta = ports_meta.get(pk, {})
        body = meta.get("body", "")
        prow = {k: "" for k in F_PORT}
        prow["Code"] = res["code"]
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
                d = {m: v for m, v in (comp.get(wkey, {}) or {}).items()}
                r2 = rest.get(wkey, {})
                p2 = prop.get(wkey, {})
                ll = d.get("ll") or r2.get("ll") or (p2.get("ll", (None,))[0])
                hh = d.get("hh") or r2.get("hh") or (p2.get("hh", (None,))[0])
                if ll and hh:
                    worig = ll * hh
            if a_o is not None and worig:
                prow["Ratio_vs_Body_Orig_pct"] = fnum(100.0 * a_o / worig, 1)
        res["ports"].append(prow)

    res["vert_cons"] = sum(walls_cons.values())
    res["f_cons"] = sum(a for (m, b, l), a in walls_cons.items() if m == "f")
    return res


# ---------------------------------------------------------------- exportació

def export_all(results, folder):
    paths = {}
    def w(name, fields, rows):
        p = os.path.join(folder, name)
        with open(p, "w", newline="", encoding="utf-8-sig") as fh:
            wr = csv.DictWriter(fh, fieldnames=fields)
            wr.writeheader()
            wr.writerows(rows)
        paths[name] = p
        return p

    est_rows = []
    mur_rows = []
    port_rows = []
    shp_rows = []
    for r in results:
        est_rows.append({
            "Code": r["code"], "Chunk_Label": r["chunk_label"],
            "Frame_Mode": r["mode"], "N_Shapes": r["n_shapes"],
            "N_Valid": r["n_valid"], "N_Invalid": r["n_invalid"],
            "N_Legacy2D": r["n_2d"], "N_QC_Flags": r["n_qc"],
            "Vert_Cons_m2": fnum(r["vert_cons"]) if r["vert_cons"] else "",
            "F_Cons_m2": fnum(r["f_cons"]) if r["f_cons"] else "",
            "T_Int_m2": fnum(r["t_int"]) if r["t_int"] else "",
            "Notes": "; ".join(r["notes"])})
        mur_rows.extend(r["murs"])
        port_rows.extend(r["ports"])
        shp_rows.extend(r["shapes"])

    w("metriques_estructures.csv", F_EST, est_rows)
    w("metriques_murs.csv", F_MUR, mur_rows)
    w("metriques_portals.csv", F_PORT, port_rows)
    w("metriques_shapes_audit.csv", F_SHP, shp_rows)
    return paths, len(est_rows), len(mur_rows), len(port_rows), len(shp_rows)


# ---------------------------------------------------------------- diàleg

def _mk_item(text, checked, data):
    item = QtWidgets.QListWidgetItem(text)
    item.setFlags(item.flags() | QtCore.Qt.ItemIsUserCheckable)
    item.setCheckState(QtCore.Qt.Checked if checked else QtCore.Qt.Unchecked)
    item.setData(QtCore.Qt.UserRole, data)
    return item


class ExportMetriquesDialog(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Exportar mètriques de shapes (BD / R)")
        self.setMinimumWidth(680)
        self._geo = Metashape.CoordinateSystem("EPSG::4326")
        self._cache = {}
        self._loading = False

        layout = QtWidgets.QVBoxLayout(self)
        layout.setSpacing(8)

        hdr = QtWidgets.QHBoxLayout()
        self.lbl_resum = QtWidgets.QLabel("")
        btn_refresh = QtWidgets.QPushButton("\u21ba")
        btn_refresh.setFixedWidth(28)
        btn_refresh.setToolTip("Torna a analitzar el projecte")
        btn_refresh.clicked.connect(self._rescan)
        hdr.addWidget(self.lbl_resum)
        hdr.addStretch()
        hdr.addWidget(btn_refresh)
        layout.addLayout(hdr)

        self.lst = QtWidgets.QListWidget()
        self.lst.setMinimumHeight(160)
        layout.addWidget(self.lst)
        row_b = QtWidgets.QHBoxLayout()
        b_all = QtWidgets.QPushButton("Tots")
        b_none = QtWidgets.QPushButton("Cap")
        for b in (b_all, b_none):
            b.setFixedWidth(60)
        b_all.clicked.connect(lambda: self._set_all(True))
        b_none.clicked.connect(lambda: self._set_all(False))
        row_b.addWidget(b_all)
        row_b.addWidget(b_none)
        row_b.addStretch()
        layout.addLayout(row_b)

        self.detail = QtWidgets.QTextEdit()
        self.detail.setReadOnly(True)
        self.detail.setMinimumHeight(120)
        self.detail.setMaximumHeight(220)
        layout.addWidget(self.detail)

        grp = QtWidgets.QGroupBox("Sortida")
        form = QtWidgets.QFormLayout(grp)
        form.setContentsMargins(12, 8, 12, 8)
        default_dir = os.path.join(
            os.path.dirname(Metashape.app.document.path or os.path.expanduser("~")),
            "metriques_%s" % datetime.date.today().strftime("%Y%m%d"))
        self.edt_dir = QtWidgets.QLineEdit(default_dir)
        btn_browse = QtWidgets.QPushButton("\u2026")
        btn_browse.setFixedWidth(32)
        btn_browse.clicked.connect(self._browse)
        row_d = QtWidgets.QHBoxLayout()
        row_d.addWidget(self.edt_dir)
        row_d.addWidget(btn_browse)
        form.addRow("Carpeta CSV:", row_d)
        layout.addWidget(grp)

        nota = QtWidgets.QLabel(
            "<small>4 CSV: estructures, murs (mur x cos, amb originals, "
            "prismes i index de conservacio), portals (ratios _cons i _orig) "
            "i audit per shape amb els QC. Nomes lectura del projecte.</small>")
        nota.setWordWrap(True)
        layout.addWidget(nota)

        blay = QtWidgets.QHBoxLayout()
        self.btn_export = QtWidgets.QPushButton("Exportar")
        self.btn_export.setDefault(True)
        btn_cancel = QtWidgets.QPushButton("Cancel\u00b7lar")
        blay.addWidget(self.btn_export)
        blay.addStretch()
        blay.addWidget(btn_cancel)
        layout.addLayout(blay)
        self.btn_export.clicked.connect(self._run)
        btn_cancel.clicked.connect(self.reject)

        self.lst.itemChanged.connect(self._recompute)
        self._rescan()

    def _rescan(self):
        prev = None
        if self.lst.count():
            prev = {self.lst.item(i).data(QtCore.Qt.UserRole): self.lst.item(i).checkState() == QtCore.Qt.Checked
                    for i in range(self.lst.count())}
        self._loading = True
        self.lst.clear()
        self._cache = {}
        doc = Metashape.app.document
        for chunk in doc.chunks:
            if chunk.shapes is None or len(chunk.shapes) == 0:
                continue
            log("analitzant %s ..." % chunk.label)
            r = analyze_chunk(chunk, self._geo)
            self._cache[chunk.label] = r
            tags = ["%d shapes" % r["n_shapes"]]
            if r["n_invalid"] or r["n_2d"]:
                tags.append("%d inv/2D" % (r["n_invalid"] + r["n_2d"]))
            if r["n_qc"]:
                tags.append("%d QC" % r["n_qc"])
            if not chunk.enabled:
                tags.append("inactiu")
            text = "%s   [%s]" % (chunk.label, ", ".join(tags))
            checked = prev.get(chunk.label, chunk.enabled) if prev else chunk.enabled
            self.lst.addItem(_mk_item(text, checked, chunk.label))
        self._loading = False
        self._recompute()

    def _set_all(self, checked):
        self._loading = True
        state = QtCore.Qt.Checked if checked else QtCore.Qt.Unchecked
        for i in range(self.lst.count()):
            self.lst.item(i).setCheckState(state)
        self._loading = False
        self._recompute()

    def _selected(self):
        out = []
        for i in range(self.lst.count()):
            it = self.lst.item(i)
            if it.checkState() == QtCore.Qt.Checked:
                out.append(self._cache[it.data(QtCore.Qt.UserRole)])
        return out

    def _recompute(self, *_a):
        if self._loading:
            return
        sel = self._selected()
        n = len(sel)
        tot_v = sum(r["n_valid"] for r in sel)
        tot_i = sum(r["n_invalid"] + r["n_2d"] for r in sel)
        tot_q = sum(r["n_qc"] for r in sel)
        resum = "<b>%d</b> estructur%s" % (n, "es" if n != 1 else "a")
        resum += "&nbsp;&nbsp;<span style='color:#2ecc71'>(%d shapes valides)</span>" % tot_v
        if tot_i:
            resum += "&nbsp;&nbsp;<span style='color:#e67e22'>(%d invalides/2D)</span>" % tot_i
        if tot_q:
            resum += "&nbsp;&nbsp;<span style='color:#3498db'>(%d avisos QC)</span>" % tot_q
        self.lbl_resum.setText(resum)

        lines = []
        for r in sel:
            probs = []
            for s in r["shapes"]:
                if s["QC_Flags"]:
                    t = "%s: %s" % (s["Label"], s["QC_Flags"])
                    if s["Suggestion"]:
                        t += "  \u2192  %s" % s["Suggestion"]
                    probs.append(t)
            for m in r["murs"]:
                if m["QC_Notes"]:
                    probs.append("mur %s cos %s: %s" % (m["Wall"], m["Body"], m["QC_Notes"]))
            icon = "\u2713" if not probs else "\u26a0"
            color = "#2ecc71" if not probs else "#e67e22"
            lines.append("<span style='color:%s'>%s</span>&nbsp;&nbsp;<b>%s</b>"
                         % (color, icon, r["code"]))
            for p in probs[:6]:
                lines.append("&nbsp;&nbsp;&nbsp;<span style='color:#888; font-size:small'>%s</span>" % p)
            if len(probs) > 6:
                lines.append("&nbsp;&nbsp;&nbsp;<span style='color:#aaa; font-size:small'>... i %d mes</span>"
                             % (len(probs) - 6))
        self.detail.setHtml("<br>".join(lines) if lines else
                            "<i style='color:gray'>Cap estructura seleccionada.</i>")
        self.btn_export.setText("Exportar (%d)" % n)
        self.btn_export.setEnabled(n > 0)

    def _browse(self):
        path = QtWidgets.QFileDialog.getExistingDirectory(
            self, "Carpeta de sortida dels CSV", self.edt_dir.text())
        if path:
            self.edt_dir.setText(path)

    def _run(self):
        folder = self.edt_dir.text().strip()
        if not folder:
            QtWidgets.QMessageBox.warning(self, "Exportar", "Indica la carpeta de sortida.")
            return
        try:
            os.makedirs(folder, exist_ok=True)
        except Exception as exc:
            QtWidgets.QMessageBox.critical(self, "Exportar", "No es pot crear la carpeta: %s" % exc)
            return
        sel = self._selected()
        self.btn_export.setEnabled(False)
        self.btn_export.setText("Exportant\u2026")
        try:
            paths, n_e, n_m, n_p, n_s = export_all(sel, folder)
        except Exception as exc:
            QtWidgets.QMessageBox.critical(self, "Exportar", "ERROR: %s" % exc)
            self._recompute()
            return
        QtWidgets.QMessageBox.information(
            self, "Exportar",
            "Exportacio completada a:\n%s\n\nestructures: %d files\nmurs: %d\n"
            "portals: %d\nshapes (audit): %d" % (folder, n_e, n_m, n_p, n_s))
        self.accept()


def show_dialog():
    dlg = ExportMetriquesDialog()
    dlg.exec_()


# Menu (estil install_menu.py):
#   MENU = "Extensions/BD Chachapoya/Exportar metriques de shapes..."
#   try: Metashape.app.removeMenuItem(MENU)
#   except Exception: pass
#   Metashape.app.addMenuItem(MENU, show_dialog)

if __name__ == "__main__":
    show_dialog()
