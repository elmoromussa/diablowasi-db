# -*- coding: utf-8 -*-
"""
batch_orto_bbox_v3.py -- Agisoft Metashape Professional 2.3.0

v3 (2026-08-21): replace_asset=False (les vistes COEXISTEIXEN com a
assets independents; amb True cada vista nova destruia l'anterior i
corrompia les shapes que s'hi recolzaven). DISCIPLINA DEL BBOX, que
este script comparteix amb azimuts i metriques:
  - la vista es tria SEMPRE amb el dialeg (PLANE_AXES), MAI girant
    el bounding box: girar-lo descalibra alhora azimuts, orthos i
    projeccions de shapes (cas EA01b, documentat);
  - el bbox es manté alineat amb el pla exposat (facana); la
    passada neta de l'extractor de metriques (zero PREFIX_DUBTOS)
    es el certificat de calibratge de tot el corpus.

En executar-se mostra un dialeg per triar el pla de projeccio (front, back,
top, bottom, right, left, relatius al bounding box) i si es vol backface
culling ("single-sided rendering"). Despres, per a cada chunk activat:
  1. Construeix una projeccio ortogonal de tipus PLANAR orientada segons el
     bounding box (region) del chunk i el pla triat (veure PLANE_AXES).
  2. Genera l'orthomosaic a 1 mm/px sobre el model (surface_data = ModelData).
  3. Exporta TIFF + JPG a la carpeta que trie l'usuari en una finestra emergent.

Execucio: Tools > Run Script (o consola Python de Metashape).
"""

import os
import re

import Metashape

# ---------------------------------------------------------------- parametres

PIXEL_SIZE = 0.001          # mida de pixel en unitats del projecte (1 mm si esta escalat en metres)

# Eix LOCAL del bbox que fa de direccio de vista (normal del pla de projeccio).
#   "Z"  -> es projecta el pla local XY  (el que deixa "rotate to view")
#   "Y"  -> es projecta el pla local XZ
#   "X"  -> es projecta el pla local YZ
# Qualsevol eix es pot prefixar amb "-" per mirar des de la cara oposada
# (es manté l'amunt de la imatge i es giren 180° X i Z de vista, per no generar
# una imatge en mirall).
#
# Correspondencia entre les vistes estandard de Metashape i l'eix de vista
# (l'eix apunta cap a l'observador). Es tria a la GUI en executar l'script;
# VIEW_AXIS nomes es el valor per defecte del dialeg.
PLANE_AXES = [
    ("front",  "-Y"),
    ("back",   "Y"),
    ("top",    "Z"),
    ("bottom", "-Z"),
    ("right",  "X"),
    ("left",   "-X"),
]
VIEW_AXIS = "-Y"

MARGIN = 0.0                # marge extra al voltant del bbox, en unitats del projecte
JPEG_QUALITY = 95
MAX_JPEG_SIDE = 65500       # limit dur del format JPEG (px)

BLENDING = Metashape.BlendingMode.MosaicBlending
FILL_HOLES = True
REFINE_SEAMLINES = False
CULL_FACES = False          # valor per defecte de la casella del dialeg

# ---------------------------------------------------------------- utilitats


def log(msg):
    print("[batchOrtho] " + msg)


def safe_name(label, fallback):
    name = re.sub(r"[^\w\-. ]+", "_", (label or "").strip())
    name = re.sub(r"\s+", "_", name).strip("._")
    return name or fallback


def ask_options():
    """Dialeg Qt amb el pla de projeccio i el backface culling.

    Retorna (view_axis, plane_name, cull_faces) o None si l'usuari cancel·la.
    Metashape Pro porta PySide2 integrat (PySide6 en versions futures).
    """
    try:
        from PySide2 import QtWidgets
    except ImportError:
        from PySide6 import QtWidgets

    dlg = QtWidgets.QDialog()
    dlg.setWindowTitle("batchOrtho - opcions")
    layout = QtWidgets.QVBoxLayout(dlg)

    form = QtWidgets.QFormLayout()
    combo = QtWidgets.QComboBox()
    for name, axis in PLANE_AXES:
        combo.addItem("%s  (eix %s)" % (name.capitalize(), axis))
        if axis == VIEW_AXIS:
            combo.setCurrentIndex(combo.count() - 1)
    form.addRow("Pla de projeccio:", combo)
    layout.addLayout(form)

    cull_check = QtWidgets.QCheckBox("Backface culling (single-sided rendering)")
    cull_check.setChecked(CULL_FACES)
    layout.addWidget(cull_check)

    buttons = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok |
                                         QtWidgets.QDialogButtonBox.Cancel)
    buttons.accepted.connect(dlg.accept)
    buttons.rejected.connect(dlg.reject)
    layout.addWidget(buttons)

    exec_dialog = getattr(dlg, "exec_", dlg.exec)
    if exec_dialog() != QtWidgets.QDialog.Accepted:
        return None

    name, axis = PLANE_AXES[combo.currentIndex()]
    return axis, name, cull_check.isChecked()


def ortho_axes(rot, axis):
    """Retorna (X, Y, Z) de la projeccio a partir de la rotacio del bbox.

    Les columnes de region.rot son els eixos locals del bbox expressats en
    coordenades internes del chunk. Z es la direccio de vista, X va cap a la
    dreta de la imatge i Y cap amunt (base dextrogira).
    """
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
    raise ValueError("VIEW_AXIS ha de ser 'X', '-X', 'Y', '-Y', 'Z' o '-Z'")


def build_projection(chunk, view_axis):
    """Projeccio planar + BBox de sortida (en coordenades de la projeccio)."""
    region = chunk.region
    T = chunk.transform.matrix
    R_world = chunk.transform.rotation          # rotacio chunk -> mon (sense escala)

    ax, ay, az = ortho_axes(region.rot, view_axis)
    ax = (R_world * ax).normalized()
    ay = (R_world * ay).normalized()
    az = (R_world * az).normalized()

    center = T.mulp(region.center)               # centre del bbox en coordenades mon

    # Files = eixos de la projeccio -> transforma mon a coordenades ortho.
    R_ortho = Metashape.Matrix([[ax.x, ax.y, ax.z],
                                [ay.x, ay.y, ay.z],
                                [az.x, az.y, az.z]])

    crs = chunk.world_crs                        # imprescindible per a projeccions planars
    if crs is None:
        crs = chunk.crs

    candidates = [
        Metashape.Matrix.Rotation(R_ortho) * Metashape.Matrix.Translation(center * -1),
        Metashape.Matrix.Translation(center) * Metashape.Matrix.Rotation(R_ortho.t()),
    ]

    proj = Metashape.OrthoProjection()
    proj.type = Metashape.OrthoProjection.Type.Planar
    proj.crs = crs

    for matrix in candidates:
        proj.matrix = matrix
        if check_projection(proj, crs, center, ax, ay, az):
            break
    else:
        raise RuntimeError("no s'ha pogut orientar la projeccio planar amb el bbox")

    # Extensio: els 8 vertexs del bbox portats a coordenades de la projeccio.
    sx, sy, sz = region.size.x / 2, region.size.y / 2, region.size.z / 2
    xs, ys = [], []
    for dx in (-sx, sx):
        for dy in (-sy, sy):
            for dz in (-sz, sz):
                corner = T.mulp(region.center + region.rot * Metashape.Vector([dx, dy, dz]))
                p = proj.transform(corner, crs, proj)
                xs.append(p.x)
                ys.append(p.y)

    bbox = Metashape.BBox()
    bbox.min = Metashape.Vector([min(xs) - MARGIN, min(ys) - MARGIN])
    bbox.max = Metashape.Vector([max(xs) + MARGIN, max(ys) + MARGIN])

    width = int(round((bbox.max.x - bbox.min.x) / PIXEL_SIZE))
    height = int(round((bbox.max.y - bbox.min.y) / PIXEL_SIZE))
    return proj, bbox, width, height


def check_projection(proj, crs, center, ax, ay, az):
    """Comprova que el pla de la projeccio es realment el pla del bbox."""
    try:
        origin = proj.transform(center, crs, proj)
        px = proj.transform(center + ax, crs, proj) - origin
        py = proj.transform(center + ay, crs, proj) - origin
        pz = proj.transform(center + az, crs, proj) - origin
    except Exception:
        return False
    return (abs(abs(px.x) - 1) < 1e-3 and abs(px.z) < 1e-3 and
            abs(abs(py.y) - 1) < 1e-3 and abs(py.z) < 1e-3 and
            abs(abs(pz.z) - 1) < 1e-3)


def export_pair(chunk, proj, bbox, folder, base, width, height):
    tif_path = os.path.join(folder, base + ".tif")
    jpg_path = os.path.join(folder, base + ".jpg")

    tif_compression = Metashape.ImageCompression()
    tif_compression.tiff_compression = Metashape.ImageCompression.TiffCompressionLZW
    tif_compression.tiff_big = True
    tif_compression.tiff_tiled = True
    tif_compression.tiff_overviews = True

    common = dict(format=Metashape.RasterFormat.RasterFormatTiles,
                  source_data=Metashape.DataSource.OrthomosaicData,
                  projection=proj,
                  region=bbox,
                  resolution=PIXEL_SIZE,
                  split_in_blocks=False,
                  north_up=False,          # cal per a projeccions planars no zenitals
                  save_world=True)

    log("  exportant TIFF: " + tif_path)
    chunk.exportRaster(path=tif_path, image_format=Metashape.ImageFormat.ImageFormatTIFF,
                       save_alpha=True, white_background=False,
                       image_compression=tif_compression, **common)

    if max(width, height) > MAX_JPEG_SIDE:
        log("  JPG omes: %d x %d px supera el limit del format (%d px)" %
            (width, height, MAX_JPEG_SIDE))
        return [tif_path]

    jpg_compression = Metashape.ImageCompression()
    jpg_compression.jpeg_quality = JPEG_QUALITY

    log("  exportant JPG:  " + jpg_path)
    chunk.exportRaster(path=jpg_path, image_format=Metashape.ImageFormat.ImageFormatJPEG,
                       save_alpha=False, white_background=True,
                       image_compression=jpg_compression, **common)
    return [tif_path, jpg_path]


# ---------------------------------------------------------------- proces


def process():
    doc = Metashape.app.document

    if not doc.path:
        Metashape.app.messageBox("Guarda el projecte (.psx) abans d'executar l'script.")
        return

    chunks = [c for c in doc.chunks if c.enabled]
    if not chunks:
        Metashape.app.messageBox("No hi ha cap chunk activat al projecte.")
        return

    options = ask_options()
    if options is None:
        log("cancel·lat per l'usuari")
        return
    view_axis, plane_name, cull_faces = options
    log("pla de projeccio: %s (eix %s), backface culling: %s" %
        (plane_name, view_axis, "si" if cull_faces else "no"))

    folder = Metashape.app.getExistingDirectory("Carpeta on guardar els orthomosaics (TIFF + JPG)")
    if not folder:
        log("cancel·lat per l'usuari")
        return

    done, skipped = [], []
    for index, chunk in enumerate(chunks, 1):
        label = chunk.label or ("Chunk_%d" % index)
        log("[%d/%d] %s" % (index, len(chunks), label))

        if chunk.model is None:
            log("  omes: el chunk no te model")
            skipped.append(label + " (sense model)")
            continue

        try:
            proj, bbox, width, height = build_projection(chunk, view_axis)
            log("  extensio: %.3f x %.3f unitats  ->  %d x %d px a %.4f u/px" %
                (bbox.max.x - bbox.min.x, bbox.max.y - bbox.min.y, width, height, PIXEL_SIZE))

            chunk.buildOrthomosaic(surface_data=Metashape.DataSource.ModelData,
                                   blending_mode=BLENDING,
                                   fill_holes=FILL_HOLES,
                                   cull_faces=cull_faces,
                                   refine_seamlines=REFINE_SEAMLINES,
                                   projection=proj,
                                   region=bbox,
                                   resolution=PIXEL_SIZE,
                                   replace_asset=False)
            # v3: replace_asset=False ES INNEGOCIABLE. Amb True, generar
            # una segona vista DESTRUIA l'ortho anterior del chunk, i les
            # shapes 2D dibuixades sobre l'ortho mort quedaven orfenes i
            # es corrompien (patologia dissecada el 17/08/2026: y=0,
            # marcs incoherents). Amb False, cada vista conviu com a
            # asset propi. Nota: les shapes de MESURA ja no es dibuixen
            # sobre orthos (convencio v1.1: model 3D), pero els orthos
            # de planta segueixen sent suport tolerat i, en tot cas,
            # producte grafic que no s'ha de matar.

            base = safe_name(chunk.label, "Chunk_%d" % index) + "_" + plane_name
            done += export_pair(chunk, proj, bbox, folder, base, width, height)
        except Exception as exc:
            log("  ERROR: %s" % exc)
            skipped.append("%s (%s)" % (label, exc))

        doc.save()

    summary = "Orthomosaics generats: %d fitxer(s) a\n%s" % (len(done), folder)
    if skipped:
        summary += "\n\nNo processats:\n- " + "\n- ".join(skipped)
    log(summary.replace("\n", " | "))
    Metashape.app.messageBox(summary)


process()
