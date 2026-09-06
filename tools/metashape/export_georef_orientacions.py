# -*- coding: utf-8 -*-
"""
export_georef_orientacions.py — Agisoft Metashape Professional 2.3.0
────────────────────────────────────────────────────────────────────────────
Exporta a CSV, UNA FILA PER ESTRUCTURA, per a la importació a Access:

  · ORIENTACIÓ DE FAÇANA — dels chunks que l'usuari marca a la llista
    "Chunks d'orientació" (candidats: etiqueta amb codi PALP, amb sufix de
    variant admés). La normal exterior del pla del bounding box es deriva
    amb la mateixa funció i el mateix VIEW_AXIS que batch_orto_bbox.py.
    Com que tots els bbox conformen plans verticals, Tilt_deg és un
    control de qualitat: |tilt| gran = bbox mal orientat.

  · COORDENADES DEL PUNT DE CIRCULACIÓ — dels chunks que l'usuari marca a
    la llista "Chunks de coordenades" (candidats: qualsevol chunk amb
    almenys un marcador situat el nom del qual és un codi PALP complet —
    típicament el model del sector sencer). Cada punt s'exporta en WGS84
    (EPSG::4326) i UTM 18S (EPSG::32718) des del MATEIX punt geocèntric.

LA TRIA ÉS DE L'USUARI, no d'una heurística: les llistes mostren tots els
chunks del projecte (els inactius, marcats com a tals) i venen
preseleccionades — orientació: chunks actius amb codi EXACTE (les
variants amb sufix, com DW-S01-EA01-2018, mai no se seleccionen soles);
coordenades: chunks actius amb marcadors de codi. El resum es recalcula
a cada canvi de casella.

Si un codi té marcador a més d'un chunk marcat, s'usa el primer (ordre
del projecte) i es calcula la DIVERGÈNCIA amb el segon com a test creuat
de georeferenciació (avís per damunt del llindar).

Sector8 en codis ANGLESOS (N/NE/E/SE/S/SW/W/NW), els que emmagatzema la
BD. ARREDONIMENT (convencio del projecte): UTM i altitud al cm (2 dec.),
lat/lon a 7 decimals (~1 cm), azimut en graus ENTERS - cap decimal no
fingeix mes precisio que la declarada a Precision_m. L'script no modifica el projecte: és de només lectura.

Execució: Tools → Run Script. Per a penjar-lo del menú Extensions,
vegeu el bloc final.
────────────────────────────────────────────────────────────────────────────
"""

from __future__ import annotations

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

# ---------------------------------------------------------------- paràmetres

# Eix LOCAL del bbox que fa de direcció de vista. HA DE COINCIDIR amb el
# VIEW_AXIS de batch_orto_bbox.py: la normal de façana se'n deriva.
VIEW_AXIS_DEFAULT = "-Y"

CRS_GEO_CODE = "EPSG::4326"    # WGS84 geogràfiques
CRS_UTM_CODE = "EPSG::32718"   # UTM 18S (Leymebamba)

# Catàleg de L_SECTORS (v23). Si la BD n'afig, afegiu-los ací.
SECTOR_CODES = (
    "LP-G", "LP-N", "LP-C", "LP-U", "LP-S",
    "DW-G", "DW-S01", "DW-S02", "DW-S03", "DW-S04", "DW-S05", "DW-S06",
)

# Gramàtica del codi. Per a chunks: sufix de variant opcional.
CODE_RE = re.compile(r"^((LP|DW)-([A-Z0-9]+)-EA(\d+))(?:-(.+))?$")
# Per a marcadors: codi EXACTE, sense sufix.
MARKER_RE = re.compile(r"^(LP|DW)-[A-Z0-9]+-EA\d+$")

# Sectors de huit valors, codis EMMAGATZEMATS (anglés, com a la BD).
SECTORS8 = ("N", "NE", "E", "SE", "S", "SW", "W", "NW")

TILT_WARN_DEG = 15.0     # els plans són verticals: més inclinació és sospitosa
MIN_HORIZ_DEG = 20.0     # pla a menys de 20° de l'horitzontal -> azimut ND
DIVERGENCE_WARN_M = 0.25 # distància entre dos chunks que dispara avís

CSV_FIELDS = (
    "Code", "Lat_WGS84", "Lon_WGS84", "E_UTM32718", "N_UTM32718",
    "Alt_masl", "Precision_m", "Azimuth_deg", "Sector8", "Tilt_deg",
    "Orientation_Chunk", "Marker_Chunk", "Notes",
)


def log(msg):
    print("[exportGeoref] " + str(msg))


# ---------------------------------------------------------------- geometria

def ortho_axes(rot, axis):
    """Idèntica a batch_orto_bbox.py. Retorna (X, Y, Z) de la projecció.

    Les columnes de region.rot són els eixos locals del bbox en coordenades
    internes del chunk. Z és la normal cap a l'espectador (base dextrogira):
    és a dir, Z ÉS LA NORMAL EXTERIOR DE LA FAÇANA.
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


def _georeferenced(chunk):
    try:
        return (chunk.crs is not None and chunk.transform is not None
                and chunk.transform.matrix is not None
                and chunk.transform.scale is not None)
    except Exception:
        return False


def facade_orientation(chunk, view_axis):
    """Azimut (graus, 0=N), sector de 8 i tilt del pla del bbox.

    Retorna (azimut_o_None, sector, tilt_deg, notes:list).
    """
    notes = []
    T = chunk.transform.matrix
    crs_geo = Metashape.CoordinateSystem(CRS_GEO_CODE)

    ax, ay, az = ortho_axes(chunk.region.rot, view_axis)
    n_world = (chunk.transform.rotation * az).normalized()
    center = T.mulp(chunk.region.center)

    # Base Est-Nord-Amunt al punt de l'estructura: geocèntric -> ENU.
    R = crs_geo.localframe(center).rotation()
    n_enu = (R * n_world).normalized()

    horiz = math.hypot(n_enu.x, n_enu.y)
    tilt = math.degrees(math.atan2(n_enu.z, horiz))

    if horiz < math.sin(math.radians(MIN_HORIZ_DEG)):
        notes.append("pla quasi horitzontal: azimut indeterminat (ND)")
        return None, "ND", tilt, notes

    azim = math.degrees(math.atan2(n_enu.x, n_enu.y)) % 360.0
    sector = SECTORS8[int(round(azim / 45.0)) % 8]
    if abs(tilt) > TILT_WARN_DEG:
        notes.append("bbox inclinat %.1f graus: revisar 'rotate to view'" % tilt)
    return azim, sector, tilt, notes


def marker_geocentric(chunk, marker):
    """Posició geocèntrica del marcador (o None si no està situat)."""
    if marker.position is None:
        return None
    return chunk.transform.matrix.mulp(marker.position)


def project_point(v):
    """(lat, lon, e, n, alt) des d'un punt geocèntric únic."""
    g = Metashape.CoordinateSystem(CRS_GEO_CODE).project(v)   # x=lon y=lat z=alt
    u = Metashape.CoordinateSystem(CRS_UTM_CODE).project(v)
    return g.y, g.x, u.x, u.y, g.z


# ---------------------------------------------------------------- inventari

def scan_project():
    """Cataloga TOTS els chunks del projecte per a les dues llistes.

    Retorna (orient_candidates, coord_candidates):
      orient: [(chunk, base_code, suffix_o_None)]
      coord:  [(chunk, n_marcadors_de_codi_situats)]
    """
    doc = Metashape.app.document
    orient, coord = [], []
    for chunk in doc.chunks:
        m = CODE_RE.match((chunk.label or "").strip())
        if m:
            orient.append((chunk, m.group(1), m.group(5)))
        n = 0
        for mk in chunk.markers:
            if (MARKER_RE.match((mk.label or "").strip())
                    and mk.position is not None):
                n += 1
        if n > 0:
            coord.append((chunk, n))
    return orient, coord


class Entry(object):
    """Una estructura: orientació + coordenades, cadascuna amb errors propis."""

    def __init__(self, code):
        self.code = code
        self.auth = None            # chunk marcat amb etiqueta == codi exacte
        self.variants = []          # chunks marcats amb sufix
        self.placements = []        # [(chunk, marker)] als chunks marcats
        self.o_errors = []
        self.c_errors = []
        self.warnings = []

    @property
    def orientation_chunk(self):
        if self.auth is not None:
            return self.auth
        if len(self.variants) == 1:
            return self.variants[0]
        return None


def build_inventory(orient_chunks, coord_chunks):
    """Construeix les entrades a partir dels chunks MARCATS per l'usuari.

    orient_chunks: [(chunk, base, suffix)] seleccionats a la llista d'orientació.
    coord_chunks:  [chunk] seleccionats a la llista de coordenades.
    """
    entries = {}

    # -- orientació --
    for chunk, base, suffix in orient_chunks:
        e = entries.setdefault(base, Entry(base))
        if suffix:
            e.variants.append(chunk)
        else:
            if e.auth is not None:
                e.o_errors.append("dos chunks marcats amb l'etiqueta exacta '%s'" % base)
            e.auth = chunk

    # -- coordenades: marcadors de codi als chunks marcats, en ordre --
    for chunk in coord_chunks:
        for mk in chunk.markers:
            label = (mk.label or "").strip()
            if not MARKER_RE.match(label) or mk.position is None:
                continue
            e = entries.setdefault(label, Entry(label))
            e.placements.append((chunk, mk))

    # -- validacions per estructura --
    for e in entries.values():
        prefix = e.code.rsplit("-EA", 1)[0]
        if prefix not in SECTOR_CODES:
            e.warnings.append("prefix '%s' fora del catàleg de L_SECTORS" % prefix)

        # orientació
        if e.auth is None and len(e.variants) > 1:
            e.o_errors.append("%d variants marcades i cap chunk de codi exacte: "
                              "marca'n només una" % len(e.variants))
        elif e.auth is None and len(e.variants) == 1:
            e.warnings.append("orientació presa de la variant '%s'"
                              % e.variants[0].label)
        ock = e.orientation_chunk
        if ock is None and not e.o_errors:
            e.o_errors.append("cap chunk d'orientació marcat")
        if ock is not None:
            if not _georeferenced(ock):
                e.o_errors.append("chunk d'orientació sense georeferenciar")
            elif ock.region is None:
                e.o_errors.append("chunk d'orientació sense bounding box")

        # coordenades
        per_chunk = {}
        for chunk, mk in e.placements:
            per_chunk.setdefault(chunk.label, []).append(mk)
        for label, mks in per_chunk.items():
            if len(mks) > 1:
                e.c_errors.append("%d marcadors '%s' al chunk '%s'"
                                  % (len(mks), e.code, label))
        if not e.placements and not e.c_errors:
            e.c_errors.append("cap marcador '%s' als chunks de coordenades marcats"
                              % e.code)
        if e.placements and not _georeferenced(e.placements[0][0]):
            e.c_errors.append("el chunk del marcador ('%s') no està georeferenciat"
                              % e.placements[0][0].label)
        if len(per_chunk) > 1 and not e.c_errors:
            e.warnings.append("marcador a %d chunks: s'usa '%s'"
                              % (len(per_chunk), e.placements[0][0].label))
            v1 = marker_geocentric(*e.placements[0])
            v2 = marker_geocentric(*e.placements[1])
            if v1 is not None and v2 is not None:
                d = (v1 - v2).norm()
                if d > DIVERGENCE_WARN_M:
                    e.warnings.append("divergència de %.2f m entre '%s' i '%s'"
                                      % (d, e.placements[0][0].label,
                                         e.placements[1][0].label))

    return [entries[k] for k in sorted(entries)]


# ---------------------------------------------------------------- exportació

def export_csv(entries, path, view_axis, precision):
    rows = []
    n_coords = 0
    n_orient = 0

    for e in entries:
        row = {k: "" for k in CSV_FIELDS}
        row["Code"] = e.code
        notes = list(e.warnings)

        # -- orientació --
        if e.o_errors:
            notes.append("SENSE ORIENTACIÓ: " + "; ".join(e.o_errors))
        else:
            ock = e.orientation_chunk
            row["Orientation_Chunk"] = ock.label
            azim, sector, tilt, o_notes = facade_orientation(ock, view_axis)
            notes.extend(o_notes)
            row["Azimuth_deg"] = "" if azim is None else str(int(round(azim)) % 360)
            row["Sector8"] = sector
            row["Tilt_deg"] = "%.1f" % tilt
            n_orient += 1

        # -- coordenades --
        if e.c_errors:
            notes.append("SENSE COORDENADES: " + "; ".join(e.c_errors))
        else:
            chunk, mk = e.placements[0]
            v = marker_geocentric(chunk, mk)
            lat, lon, east, north, alt = project_point(v)
            row["Lat_WGS84"] = "%.7f" % lat
            row["Lon_WGS84"] = "%.7f" % lon
            row["E_UTM32718"] = "%.2f" % east
            row["N_UTM32718"] = "%.2f" % north
            row["Alt_masl"] = "%.2f" % alt
            row["Precision_m"] = precision
            row["Marker_Chunk"] = chunk.label
            n_coords += 1

        row["Notes"] = "; ".join(notes)
        rows.append(row)
        log("  %s -> orient=%s coords=%s"
            % (e.code, row["Sector8"] or "no", "si" if row["Lat_WGS84"] else "no"))

    with open(path, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=CSV_FIELDS)
        w.writeheader()
        w.writerows(rows)

    return len(rows), n_coords, n_orient


# ---------------------------------------------------------------- diàleg

def _mk_item(text, checked, data):
    item = QtWidgets.QListWidgetItem(text)
    item.setFlags(item.flags() | QtCore.Qt.ItemIsUserCheckable)
    item.setCheckState(QtCore.Qt.Checked if checked else QtCore.Qt.Unchecked)
    item.setData(QtCore.Qt.UserRole, data)
    return item


class ExportGeorefDialog(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Exportar coordenades i orientacions (BD)")
        self.setMinimumWidth(760)
        self._loading = False

        layout = QtWidgets.QVBoxLayout(self)
        layout.setSpacing(8)

        # ── Capçalera amb resum i botó de refresc ────────────────────────
        hdr = QtWidgets.QHBoxLayout()
        self.lbl_resum = QtWidgets.QLabel("")
        btn_refresh = QtWidgets.QPushButton("\u21ba")
        btn_refresh.setFixedWidth(28)
        btn_refresh.setToolTip("Torna a escanejar el projecte "
                               "(conserva les caselles per etiqueta)")
        btn_refresh.clicked.connect(self._rescan)
        hdr.addWidget(self.lbl_resum)
        hdr.addStretch()
        hdr.addWidget(btn_refresh)
        layout.addLayout(hdr)

        # ── Les dues llistes de selecció ─────────────────────────────────
        lists_lay = QtWidgets.QHBoxLayout()

        grp_o = QtWidgets.QGroupBox("Chunks d'orientació (pla del bbox)")
        lay_o = QtWidgets.QVBoxLayout(grp_o)
        self.lst_orient = QtWidgets.QListWidget()
        self.lst_orient.setMinimumHeight(180)
        lay_o.addWidget(self.lst_orient)
        row_o = QtWidgets.QHBoxLayout()
        b_o_all = QtWidgets.QPushButton("Tots")
        b_o_none = QtWidgets.QPushButton("Cap")
        for b in (b_o_all, b_o_none):
            b.setFixedWidth(60)
        b_o_all.clicked.connect(lambda: self._set_all(self.lst_orient, True))
        b_o_none.clicked.connect(lambda: self._set_all(self.lst_orient, False))
        row_o.addWidget(b_o_all)
        row_o.addWidget(b_o_none)
        row_o.addStretch()
        lay_o.addLayout(row_o)
        lists_lay.addWidget(grp_o)

        grp_c = QtWidgets.QGroupBox("Chunks de coordenades (marcadors de codi)")
        lay_c = QtWidgets.QVBoxLayout(grp_c)
        self.lst_coord = QtWidgets.QListWidget()
        self.lst_coord.setMinimumHeight(180)
        lay_c.addWidget(self.lst_coord)
        row_c = QtWidgets.QHBoxLayout()
        b_c_all = QtWidgets.QPushButton("Tots")
        b_c_none = QtWidgets.QPushButton("Cap")
        for b in (b_c_all, b_c_none):
            b.setFixedWidth(60)
        b_c_all.clicked.connect(lambda: self._set_all(self.lst_coord, True))
        b_c_none.clicked.connect(lambda: self._set_all(self.lst_coord, False))
        row_c.addWidget(b_c_all)
        row_c.addWidget(b_c_none)
        row_c.addStretch()
        lay_c.addLayout(row_c)
        lists_lay.addWidget(grp_c)

        layout.addLayout(lists_lay)

        self.lst_orient.itemChanged.connect(self._recompute)
        self.lst_coord.itemChanged.connect(self._recompute)

        # ── Resum d'estructures ──────────────────────────────────────────
        self.llista = QtWidgets.QTextEdit()
        self.llista.setReadOnly(True)
        self.llista.setMinimumHeight(140)
        self.llista.setMaximumHeight(240)
        layout.addWidget(self.llista)

        # ── Opcions ──────────────────────────────────────────────────────
        grp = QtWidgets.QGroupBox("Opcions")
        form = QtWidgets.QFormLayout(grp)
        form.setContentsMargins(12, 8, 12, 8)

        self.cmb_axis = QtWidgets.QComboBox()
        self.cmb_axis.addItems(["-Y", "Y", "-X", "X", "-Z", "Z"])
        self.cmb_axis.setCurrentText(VIEW_AXIS_DEFAULT)
        self.cmb_axis.setToolTip(
            "Ha de coincidir amb el VIEW_AXIS de batch_orto_bbox.py:\n"
            "la normal de façana es deriva del mateix eix."
        )
        self.cmb_axis.currentTextChanged.connect(self._recompute)
        form.addRow("Eix de vista (bbox):", self.cmb_axis)

        self.edt_precision = QtWidgets.QLineEdit("0.05")
        self.edt_precision.setToolTip(
            "Precisió estimada de la georeferenciació (m), constant per a\n"
            "totes les files (Coord_Precision_m). Buit = columna en blanc."
        )
        self.edt_precision.setFixedWidth(80)
        form.addRow("Precisió (m):", self.edt_precision)

        default_csv = os.path.join(
            os.path.dirname(Metashape.app.document.path or os.path.expanduser("~")),
            "georef_estructures_%s.csv" % datetime.date.today().strftime("%Y%m%d"))
        self.edt_csv = QtWidgets.QLineEdit(default_csv)
        btn_browse = QtWidgets.QPushButton("\u2026")
        btn_browse.setFixedWidth(32)
        btn_browse.clicked.connect(self._browse_csv)
        row_csv = QtWidgets.QHBoxLayout()
        row_csv.addWidget(self.edt_csv)
        row_csv.addWidget(btn_browse)
        form.addRow("Fitxer CSV:", row_csv)

        layout.addWidget(grp)

        # ── Nota informativa ─────────────────────────────────────────────
        nota = QtWidgets.QLabel(
            "<small>Preselecció: <b>orientació</b> = chunks actius de codi exacte "
            "(les variants amb sufix es marquen a mà, mai soles); "
            "<b>coordenades</b> = chunks actius amb marcadors de codi. "
            "Una fila per estructura; sectors en anglés (com a la BD). "
            "L'script no modifica el projecte.</small>"
        )
        nota.setWordWrap(True)
        layout.addWidget(nota)

        # ── Botons ───────────────────────────────────────────────────────
        blay = QtWidgets.QHBoxLayout()
        self.btn_export = QtWidgets.QPushButton("Exportar")
        self.btn_export.setDefault(True)
        btn_cancel = QtWidgets.QPushButton("Cancel·lar")
        blay.addWidget(self.btn_export)
        blay.addStretch()
        blay.addWidget(btn_cancel)
        layout.addLayout(blay)

        self.btn_export.clicked.connect(self._run)
        btn_cancel.clicked.connect(self.reject)

        self._rescan()

    # ── Poblament de les llistes ─────────────────────────────────────────

    def _checked_labels(self, lst):
        out = set()
        for i in range(lst.count()):
            it = lst.item(i)
            if it.checkState() == QtCore.Qt.Checked:
                out.add(it.data(QtCore.Qt.UserRole)[0].label)
        return out

    def _rescan(self):
        prev_o = self._checked_labels(self.lst_orient) if self.lst_orient.count() else None
        prev_c = self._checked_labels(self.lst_coord) if self.lst_coord.count() else None

        self._loading = True
        self.lst_orient.clear()
        self.lst_coord.clear()

        orient, coord = scan_project()

        for chunk, base, suffix in orient:
            tags = []
            if suffix:
                tags.append("variant")
            if not chunk.enabled:
                tags.append("inactiu")
            text = chunk.label + ("   [%s]" % ", ".join(tags) if tags else "")
            if prev_o is not None:
                checked = chunk.label in prev_o
            else:
                checked = chunk.enabled and suffix is None
            self.lst_orient.addItem(_mk_item(text, checked, (chunk, base, suffix)))

        for chunk, n in coord:
            tags = ["%d marcador%s" % (n, "s" if n != 1 else "")]
            if not chunk.enabled:
                tags.append("inactiu")
            text = chunk.label + "   [%s]" % ", ".join(tags)
            if prev_c is not None:
                checked = chunk.label in prev_c
            else:
                checked = chunk.enabled
            self.lst_coord.addItem(_mk_item(text, checked, (chunk, n)))

        self._loading = False
        self._recompute()

    def _set_all(self, lst, checked):
        self._loading = True
        state = QtCore.Qt.Checked if checked else QtCore.Qt.Unchecked
        for i in range(lst.count()):
            lst.item(i).setCheckState(state)
        self._loading = False
        self._recompute()

    # ── Resum ────────────────────────────────────────────────────────────

    def _selection(self):
        orient = []
        for i in range(self.lst_orient.count()):
            it = self.lst_orient.item(i)
            if it.checkState() == QtCore.Qt.Checked:
                orient.append(it.data(QtCore.Qt.UserRole))
        coord = []
        for i in range(self.lst_coord.count()):
            it = self.lst_coord.item(i)
            if it.checkState() == QtCore.Qt.Checked:
                coord.append(it.data(QtCore.Qt.UserRole)[0])
        return orient, coord

    def _recompute(self, *_args):
        if self._loading:
            return
        orient, coord = self._selection()
        self._entries = build_inventory(orient, coord)

        complete, only_orient, only_coords, blocked = [], [], [], []
        for e in self._entries:
            has_o = not e.o_errors
            has_c = not e.c_errors
            if has_o and has_c:
                complete.append(e)
            elif has_o:
                only_orient.append(e)
            elif has_c:
                only_coords.append(e)
            else:
                blocked.append(e)

        total = len(self._entries)
        if total == 0:
            self.lbl_resum.setText(
                "<span style='color:red'>Cap estructura amb la selecció actual.</span>")
            self.llista.setHtml(
                "<i style='color:gray'>Marca chunks a les dues llistes.</i>")
            self.btn_export.setText("Exportar")
            self.btn_export.setEnabled(False)
            return

        resum = "<b>%d</b> estructur%s" % (total, "es" if total != 1 else "a")
        if complete:
            resum += "&nbsp;&nbsp;<span style='color:#2ecc71'>(%d completes)</span>" % len(complete)
        if only_orient:
            resum += ("&nbsp;&nbsp;<span style='color:#3498db'>"
                      "(%d només orientació)</span>" % len(only_orient))
        if only_coords:
            resum += ("&nbsp;&nbsp;<span style='color:#9b59b6'>"
                      "(%d només coordenades)</span>" % len(only_coords))
        if blocked:
            resum += ("&nbsp;&nbsp;<span style='color:#e67e22'>"
                      "(%d bloquejades)</span>" % len(blocked))
        self.lbl_resum.setText(resum)

        lines = []
        for e in complete:
            extra = ("&nbsp;&nbsp;<span style='color:#aaa; font-size:small'>%s</span>"
                     % "; ".join(e.warnings)) if e.warnings else ""
            lines.append(
                "<span style='color:#2ecc71'>\u2713</span>&nbsp;&nbsp;"
                "<b>%s</b>&nbsp;&nbsp;<span style='color:#555; font-size:small'>"
                "coordenades + orientació</span>%s" % (e.code, extra))
        for e in only_orient:
            lines.append(
                "<span style='color:#3498db'><b>\u2192</b></span>&nbsp;&nbsp;"
                "<b>%s</b>&nbsp;&nbsp;<span style='color:#3498db; font-size:small'>"
                "només orientació — %s</span>" % (e.code, "; ".join(e.c_errors)))
        for e in only_coords:
            lines.append(
                "<span style='color:#9b59b6'><b>\u25cf</b></span>&nbsp;&nbsp;"
                "<b>%s</b>&nbsp;&nbsp;<span style='color:#9b59b6; font-size:small'>"
                "només coordenades — %s</span>" % (e.code, "; ".join(e.o_errors)))
        for e in blocked:
            lines.append(
                "<span style='color:#e67e22'>\u26a0</span>&nbsp;&nbsp;"
                "<span style='color:#888'>%s</span>&nbsp;&nbsp;"
                "<span style='color:#e67e22; font-size:small'>%s</span>"
                % (e.code, "; ".join(e.o_errors + e.c_errors)))
        self.llista.setHtml("<br>".join(lines))

        exportables = len(complete) + len(only_orient) + len(only_coords)
        self.btn_export.setText("Exportar (%d)" % exportables)
        self.btn_export.setEnabled(exportables > 0)

    # ── Accions ──────────────────────────────────────────────────────────

    def _browse_csv(self):
        path, _f = QtWidgets.QFileDialog.getSaveFileName(
            self, "Fitxer CSV de sortida", self.edt_csv.text(), "CSV (*.csv)")
        if path:
            self.edt_csv.setText(path)

    def _run(self):
        path = self.edt_csv.text().strip()
        if not path:
            QtWidgets.QMessageBox.warning(self, "Exportar",
                                          "Indica el fitxer CSV de sortida.")
            return
        precision = self.edt_precision.text().strip()

        self.btn_export.setEnabled(False)
        self.btn_export.setText("Exportant\u2026")
        try:
            n_rows, n_coords, n_orient = export_csv(
                self._entries, path, self.cmb_axis.currentText(), precision)
        except Exception as exc:
            QtWidgets.QMessageBox.critical(self, "Exportar", "ERROR: %s" % exc)
            self._recompute()
            return

        QtWidgets.QMessageBox.information(
            self, "Exportar",
            "CSV generat: %s\n\nFiles: %d\nAmb coordenades: %d\n"
            "Amb orientació: %d" % (path, n_rows, n_coords, n_orient))
        self.accept()


def show_dialog():
    dlg = ExportGeorefDialog()
    dlg.exec_()


# Per a penjar-lo del menú (estil install_menu.py d'oculat):
#
#   MENU_GEOREF = "Extensions/BD Chachapoya/Exportar coordenades i orientacions..."
#   try:
#       Metashape.app.removeMenuItem(MENU_GEOREF)
#   except Exception:
#       pass
#   Metashape.app.addMenuItem(MENU_GEOREF, show_dialog)

if __name__ == "__main__":
    show_dialog()
