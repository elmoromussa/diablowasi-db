"""
auditar_shapes.py
────────────────────────────────────────────────────────────────────────────
Instrument MÍNIM de depuració: imprimeix a la consola exactament el que hi ha
guardat a cada shape del CHUNK ACTIU, sense cap interpretació de marcs ni CRS.

Per shape: tipus, nombre de vèrtexs, dims (2D/3D), dos primers vèrtexs crus,
extensió crua en x/y(/z); polígons: àrea (shoelace sobre x,y crues) i
ompliment (àrea / envolvent; < 5 % delata geometria col·lapsada);
polilínies: longitud crua. També l'inventari d'orthomosaics del chunk.

Només lectura. Ús: Extensions → Base de dades → Auditar shapes del chunk actiu
────────────────────────────────────────────────────────────────────────────
"""

from __future__ import annotations

import math
import sys
from pathlib import Path

import Metashape

_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from chachapoya_bd.db.shapes import shoelace  # noqa: E402

OMPLIMENT_MIN_PCT = 5.0


def auditar(chunk) -> int:
    """Escriu l'informe a la consola. Retorna el nombre de shapes col·lapsades."""
    print("Chunk actiu: %s" % chunk.label)
    try:
        orthos = chunk.orthomosaics
    except AttributeError:
        orthos = [chunk.orthomosaic] if chunk.orthomosaic else []
    active = getattr(chunk, "orthomosaic", None)
    print("Orthomosaics al chunk: %d" % len(orthos))
    for o in orthos:
        try:
            mark = "  <- ACTIU" if (active is not None and o.key == active.key) else ""
            print("  [key=%s] %s  %dx%d px%s" % (o.key, getattr(o, "label", "") or "(sense etiqueta)",
                                                 o.width, o.height, mark))
        except Exception:
            print("  (ortho no llegible)")

    if chunk.shapes is None or len(chunk.shapes) == 0:
        print("El chunk no té shapes.")
        return 0
    scrs = chunk.shapes.crs
    print("CRS de shapes declarat: %s" % (scrs.name if scrs else "cap"))

    colapsades = 0
    for shape in chunk.shapes:
        g = shape.geometry
        label = shape.label or "(sense etiqueta)"
        try:
            coords = list(g.coordinates[0]) if g.type == Metashape.Geometry.Type.PolygonType \
                else list(g.coordinates)
        except Exception as exc:
            print("\n[%s] ERROR: %s" % (label, exc))
            continue
        if len(coords) < 2:
            print("\n[%s] sense prou vèrtexs" % label)
            continue

        dims = coords[0].size
        tname = str(g.type).split(".")[-1]
        print("\n[%s]  %s  vèrtexs=%d  dims=%dD" % (label, tname, len(coords), dims))
        for i, c in enumerate(coords[:2]):
            if dims >= 3:
                print("   v%d: x=%.4f  y=%.4f  z=%.4f" % (i, c.x, c.y, c.z))
            else:
                print("   v%d: x=%.4f  y=%.4f" % (i, c.x, c.y))

        xs = [c.x for c in coords]
        ys = [c.y for c in coords]
        ext_x = max(xs) - min(xs)
        ext_y = max(ys) - min(ys)
        line = "   ext_x: %.3f | ext_y: %.3f" % (ext_x, ext_y)
        if dims >= 3:
            zs = [c.z for c in coords]
            line += " | ext_z: %.3f" % (max(zs) - min(zs))
        print(line)

        if g.type == Metashape.Geometry.Type.PolygonType:
            area = shoelace(list(zip(xs, ys)))
            env = ext_x * ext_y
            fill = (100.0 * area / env) if env > 1e-12 else 0.0
            flag = ""
            if fill < OMPLIMENT_MIN_PCT:
                flag = "   *** COL·LAPSADA ***"
                colapsades += 1
            print("   àrea (x,y): %.3f | ompliment: %.0f%%%s" % (area, fill, flag))
        elif g.type == Metashape.Geometry.Type.LineStringType:
            if dims >= 3:
                L = sum(math.dist((coords[i].x, coords[i].y, coords[i].z),
                                  (coords[i + 1].x, coords[i + 1].y, coords[i + 1].z))
                        for i in range(len(coords) - 1))
            else:
                L = sum(math.hypot(coords[i + 1].x - coords[i].x, coords[i + 1].y - coords[i].y)
                        for i in range(len(coords) - 1))
            print("   longitud: %.3f" % L)
    return colapsades


def run():
    doc = Metashape.app.document
    chunk = doc.chunk if doc else None
    if chunk is None:
        Metashape.app.messageBox("Cap chunk actiu.")
        return
    n = auditar(chunk)
    n_shapes = len(chunk.shapes) if chunk.shapes else 0
    Metashape.app.messageBox("Auditoria de «%s» escrita a la consola.\n\nShapes: %d\n"
                             "Col·lapsades (ompliment < %.0f%%): %d"
                             % (chunk.label, n_shapes, OMPLIMENT_MIN_PCT, n))


if __name__ == "__main__":
    run()
