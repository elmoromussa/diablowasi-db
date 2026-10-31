"""
exportar_georef_bd.py
────────────────────────────────────────────────────────────────────────────
Exporta a CSV, UNA FILA PER ESTRUCTURA, per a la importació a Access
(``access/chachapoya_Import_Georef_v24.bas``):

  · ORIENTACIÓ DE FAÇANA — dels chunks marcats a "Chunks d'orientació"
    (candidats: etiqueta amb codi d'estructura, amb sufix de variant admés).
    La normal exterior del pla del bounding box es deriva amb els mateixos
    eixos que *Generar Ortomosaics…* (vista ``eix_facana`` de bd.yaml).
  · COORDENADES DEL PUNT DE CIRCULACIÓ — dels chunks marcats a "Chunks de
    coordenades" (candidats: qualsevol chunk amb un marcador situat el nom
    del qual és un codi complet). Cada punt s'exporta en WGS84 i UTM
    (``epsg_utm`` de bd.yaml) des del MATEIX punt geocèntric.

La tria és de l'usuari; les llistes venen preseleccionades. Només lectura.

Compatible amb Metashape Professional 2.3.0
Ús: Extensions → Base de dades → Exportar georef i orientacions…
Lògica pura (inventari, validacions, files): ``pipeline/db/georef.py``.
────────────────────────────────────────────────────────────────────────────
"""

from __future__ import annotations

import datetime
import html
import os
import sys
from pathlib import Path

import Metashape

try:
    from PySide2 import QtCore, QtWidgets
except ImportError:
    from PySide6 import QtCore, QtWidgets  # type: ignore

_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from chachapoya_bd._vendor.qt_comu import informe  # noqa: E402
from chachapoya_bd import configuracio  # noqa: E402
from chachapoya_bd.db import vistes as om  # noqa: E402
from chachapoya_bd.db import codis as db_codis  # noqa: E402
from chachapoya_bd.db import georef as db  # noqa: E402
from chachapoya_bd.db.csv_out import escriure_csv  # noqa: E402


def log(msg: str) -> None:
    print("[exportGeoref] " + msg)


def _opcions() -> dict:
    """Opcions del ``bd.yaml`` que toca (el del jaciment vinculat o el del repo), via ``configuracio.py``.

    Hi afig ``fitxer_bd``, ``font_bd``, ``origen_epsg_utm`` i ``avis_vincle`` per al diàleg.
    No amaga cap error: si no hi ha cap bd.yaml, l'excepció puja fins a
    ``show_dialog`` i s'ensenya a l'usuari.
    """
    doc = Metashape.app.document
    c = configuracio.configuracio(doc.meta, doc.path)
    return dict(c.opcions, fitxer_bd=str(c.fitxer_bd), font_bd=c.font,
                origen_epsg_utm=c.origen_epsg_utm, avis_vincle=c.avis)


def _opcions_validades(titol: str) -> dict | None:
    """Opcions llestes per al diàleg, o None (ja s'ha informat l'usuari)."""
    try:
        opts = _opcions()
    except Exception as e:  # cap bd.yaml, o il·legible
        informe(None, titol, "No s'ha pogut llegir la configuració de la base de dades.", [str(e)])
        return None
    errors = db_codis.validar_opcions(opts)
    if errors:
        informe(None, titol, "Falten valors obligatoris a bd.yaml.",
                errors)
        return None
    if opts["avis_vincle"]:
        log(opts["avis_vincle"])
    log("bd.yaml: %s (%s)" % (opts["fitxer_bd"], configuracio.DESCRIPCIO_FONT[opts["font_bd"]]))
    log("codi_regex: %s" % opts["codi_regex"])
    log("epsg_utm: EPSG:%s (de %s)" % (opts["epsg_utm"], opts["font_epsg_utm"]))
    return opts


def _html_avis(opts: dict) -> str:
    """Avís de vincle trencat per al peu del diàleg (HTML), o buit."""
    return "<br><b>%s</b>" % html.escape(opts["avis_vincle"]) if opts.get("avis_vincle") else ""


def _carpeta_defecte() -> str:
    try:
        doc = Metashape.app.document
        return str(configuracio.carpeta_eixida(doc.meta, doc.path))
    except Exception:
        return os.path.expanduser("~")


# ── Metashape → dades pures ──────────────────────────────────────────────────


def _georeferenciat(chunk) -> bool:
    try:
        return (chunk.crs is not None and chunk.transform is not None
                and chunk.transform.matrix is not None
                and chunk.transform.scale is not None)
    except Exception:
        return False


def _chunk_info(chunk, codis: db_codis.Codis) -> db.ChunkInfo:
    marcadors = []
    georef = _georeferenciat(chunk)
    for mk in chunk.markers:
        if not codis.es_exacte(mk.label):
            continue
        pos = None
        if mk.position is not None and georef:
            v = chunk.transform.matrix.mulp(mk.position)
            pos = (v.x, v.y, v.z)
        elif mk.position is not None:
            pos = (mk.position.x, mk.position.y, mk.position.z)
        marcadors.append(db.Marcador(mk.label or "", pos))
    return db.ChunkInfo(label=chunk.label or "", enabled=bool(chunk.enabled),
                        georeferenciat=georef, te_region=chunk.region is not None,
                        marcadors=marcadors)


class Calculs:
    """Geometria amb l'API de Metashape (orientació i projecció de punts)."""

    def __init__(self, opts: dict):
        self.opts = opts
        self.crs_geo = Metashape.CoordinateSystem("EPSG::%d" % int(opts["epsg_geo"]))
        self.crs_utm = Metashape.CoordinateSystem("EPSG::%d" % int(opts["epsg_utm"]))
        self.per_label = {c.label: c for c in Metashape.app.document.chunks}

    def orientacio(self, info: db.ChunkInfo):
        chunk = self.per_label[info.label]
        rot = chunk.region.rot
        _ax, _ay, az = om.eixos_per_vista(rot.col(0), rot.col(1), rot.col(2),
                                          self.opts["eix_facana"])
        n_world = (chunk.transform.rotation * az).normalized()
        center = chunk.transform.matrix.mulp(chunk.region.center)
        R = self.crs_geo.localframe(center).rotation()   # geocèntric → ENU
        n_enu = (R * n_world).normalized()
        return db.orientacio_facana((n_enu.x, n_enu.y, n_enu.z),
                                    float(self.opts["min_horitz_deg"]),
                                    float(self.opts["tilt_avis_deg"]))

    def coordenades(self, info: db.ChunkInfo, mk: db.Marcador):
        v = Metashape.Vector(mk.posicio)
        g = self.crs_geo.project(v)   # x=lon y=lat z=alt
        u = self.crs_utm.project(v)
        return g.y, g.x, u.x, u.y, g.z


# ── Diàleg ───────────────────────────────────────────────────────────────────


def _mk_item(text, checked, data):
    item = QtWidgets.QListWidgetItem(text)
    item.setFlags(item.flags() | QtCore.Qt.ItemIsUserCheckable)
    item.setCheckState(QtCore.Qt.Checked if checked else QtCore.Qt.Unchecked)
    item.setData(QtCore.Qt.UserRole, data)
    return item


class ExportGeorefDialog(QtWidgets.QDialog):
    def __init__(self, opts: dict, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Exportar georef i orientacions (BD)")
        self.setMinimumWidth(760)
        self.opts = opts
        self.codis = db_codis.Codis(self.opts)
        self._loading = False
        self._entrades: list[db.Entrada] = []

        layout = QtWidgets.QVBoxLayout(self)
        layout.setSpacing(8)

        hdr = QtWidgets.QHBoxLayout()
        self.lbl_resum = QtWidgets.QLabel("")
        btn_refresh = QtWidgets.QPushButton("↺")
        btn_refresh.setFixedWidth(28)
        btn_refresh.setToolTip("Torna a escanejar el projecte (conserva les caselles per etiqueta)")
        btn_refresh.clicked.connect(self._rescan)
        hdr.addWidget(self.lbl_resum)
        hdr.addStretch()
        hdr.addWidget(btn_refresh)
        layout.addLayout(hdr)

        lists_lay = QtWidgets.QHBoxLayout()
        self.lst_orient = self._grup_llista("Chunks d'orientació (pla del bbox)", lists_lay)
        self.lst_coord = self._grup_llista("Chunks de coordenades (marcadors de codi)", lists_lay)
        layout.addLayout(lists_lay)
        self.lst_orient.itemChanged.connect(self._recompute)
        self.lst_coord.itemChanged.connect(self._recompute)

        self.llista = QtWidgets.QTextEdit()
        self.llista.setReadOnly(True)
        self.llista.setMinimumHeight(140)
        self.llista.setMaximumHeight(240)
        layout.addWidget(self.llista)

        grp = QtWidgets.QGroupBox("Opcions")
        form = QtWidgets.QFormLayout(grp)
        form.setContentsMargins(12, 8, 12, 8)
        self.lbl_crs = QtWidgets.QLabel("EPSG:%d (geo) · EPSG:%d (UTM, %s) · façana = vista «%s»"
                                        % (self.opts["epsg_geo"], self.opts["epsg_utm"],
                                           self.opts.get("origen_epsg_utm") or "?",
                                           self.opts["eix_facana"]))
        form.addRow("Sistemes:", self.lbl_crs)
        self.edt_precision = QtWidgets.QLineEdit(str(self.opts["precisio_m"]))
        self.edt_precision.setToolTip("Precisió estimada de la georeferenciació (m), constant per\n"
                                      "a totes les files (Coord_Precision_m). Buit = columna en blanc.")
        self.edt_precision.setFixedWidth(80)
        form.addRow("Precisió (m):", self.edt_precision)
        default_csv = os.path.join(_carpeta_defecte(), "georef_estructures_%s.csv"
                                   % datetime.date.today().strftime("%Y%m%d"))
        self.edt_csv = QtWidgets.QLineEdit(default_csv)
        btn_browse = QtWidgets.QPushButton("…")
        btn_browse.setFixedWidth(32)
        btn_browse.clicked.connect(self._browse_csv)
        row_csv = QtWidgets.QHBoxLayout()
        row_csv.addWidget(self.edt_csv)
        row_csv.addWidget(btn_browse)
        form.addRow("Fitxer CSV:", row_csv)
        layout.addWidget(grp)

        nota = QtWidgets.QLabel(
            "<small>Preselecció: <b>orientació</b> = chunks actius de codi exacte "
            "(les variants amb sufix es marquen a mà, mai soles); "
            "<b>coordenades</b> = chunks actius amb marcadors de codi. "
            "Una fila per estructura; sectors en anglés (com a la BD). "
            "No modifica el projecte. Configuració: <code>%s</code> (%s); EPSG UTM de <code>%s</code>.%s</small>"
            % (html.escape(self.opts["fitxer_bd"]), configuracio.DESCRIPCIO_FONT[self.opts["font_bd"]],
               html.escape(self.opts.get("font_epsg_utm") or "?"), _html_avis(self.opts)))
        nota.setWordWrap(True)
        layout.addWidget(nota)

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

    def _grup_llista(self, titol, parent_lay):
        grp = QtWidgets.QGroupBox(titol)
        lay = QtWidgets.QVBoxLayout(grp)
        lst = QtWidgets.QListWidget()
        lst.setMinimumHeight(180)
        lay.addWidget(lst)
        row = QtWidgets.QHBoxLayout()
        for text, val in (("Tots", True), ("Cap", False)):
            b = QtWidgets.QPushButton(text)
            b.setFixedWidth(60)
            b.clicked.connect(lambda _c=False, lst=lst, val=val: self._set_all(lst, val))
            row.addWidget(b)
        row.addStretch()
        lay.addLayout(row)
        parent_lay.addWidget(grp)
        return lst

    # ── llistes ──

    def _checked_labels(self, lst):
        return {lst.item(i).data(QtCore.Qt.UserRole)[0].label
                for i in range(lst.count())
                if lst.item(i).checkState() == QtCore.Qt.Checked}

    def _rescan(self):
        prev_o = self._checked_labels(self.lst_orient) if self.lst_orient.count() else None
        prev_c = self._checked_labels(self.lst_coord) if self.lst_coord.count() else None
        self._loading = True
        self.lst_orient.clear()
        self.lst_coord.clear()

        self.calc = Calculs(self.opts)
        infos = [_chunk_info(c, self.codis) for c in Metashape.app.document.chunks]
        orient, coord = db.candidats(infos, self.codis)

        for cand in orient:
            tags = []
            if cand.variant:
                tags.append("variant")
            if not cand.chunk.enabled:
                tags.append("inactiu")
            text = cand.chunk.label + ("   [%s]" % ", ".join(tags) if tags else "")
            checked = (cand.chunk.label in prev_o) if prev_o is not None \
                else db.preseleccio_orientacio(cand)
            self.lst_orient.addItem(_mk_item(text, checked, (cand.chunk, cand)))

        for chunk, n in coord:
            tags = ["%d marcador%s" % (n, "s" if n != 1 else "")]
            if not chunk.enabled:
                tags.append("inactiu")
            text = chunk.label + "   [%s]" % ", ".join(tags)
            checked = (chunk.label in prev_c) if prev_c is not None else chunk.enabled
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

    def _selection(self):
        orient = [self.lst_orient.item(i).data(QtCore.Qt.UserRole)[1]
                  for i in range(self.lst_orient.count())
                  if self.lst_orient.item(i).checkState() == QtCore.Qt.Checked]
        coord = [self.lst_coord.item(i).data(QtCore.Qt.UserRole)[0]
                 for i in range(self.lst_coord.count())
                 if self.lst_coord.item(i).checkState() == QtCore.Qt.Checked]
        return orient, coord

    # ── resum ──

    def _recompute(self, *_args):
        if self._loading:
            return
        orient, coord = self._selection()
        self._entrades = db.construir_inventari(orient, coord, self.codis,
                                                float(self.opts["divergencia_avis_m"]))
        r = db.resum_estats(self._entrades)
        total = len(self._entrades)
        if total == 0:
            self.lbl_resum.setText("<span style='color:red'>Cap estructura amb la selecció actual.</span>")
            self.llista.setHtml("<i style='color:gray'>Marca chunks a les dues llistes.</i>")
            self.btn_export.setText("Exportar")
            self.btn_export.setEnabled(False)
            return

        resum = "<b>%d</b> estructur%s" % (total, "es" if total != 1 else "a")
        for clau, color, text in (("completa", "#2ecc71", "completes"),
                                  ("orientacio", "#3498db", "només orientació"),
                                  ("coordenades", "#9b59b6", "només coordenades"),
                                  ("bloquejada", "#e67e22", "bloquejades")):
            if r[clau]:
                resum += "&nbsp;&nbsp;<span style='color:%s'>(%d %s)</span>" % (color, r[clau], text)
        self.lbl_resum.setText(resum)

        lines = []
        for e in self._entrades:
            if e.estat == "completa":
                extra = ("&nbsp;&nbsp;<span style='color:#aaa; font-size:small'>%s</span>"
                         % "; ".join(e.avisos)) if e.avisos else ""
                lines.append("<span style='color:#2ecc71'>✓</span>&nbsp;&nbsp;<b>%s</b>"
                             "&nbsp;&nbsp;<span style='color:#555; font-size:small'>"
                             "coordenades + orientació</span>%s" % (e.codi, extra))
            elif e.estat == "orientacio":
                lines.append("<span style='color:#3498db'><b>→</b></span>&nbsp;&nbsp;<b>%s</b>"
                             "&nbsp;&nbsp;<span style='color:#3498db; font-size:small'>"
                             "només orientació — %s</span>" % (e.codi, "; ".join(e.c_errors)))
            elif e.estat == "coordenades":
                lines.append("<span style='color:#9b59b6'><b>●</b></span>&nbsp;&nbsp;<b>%s</b>"
                             "&nbsp;&nbsp;<span style='color:#9b59b6; font-size:small'>"
                             "només coordenades — %s</span>" % (e.codi, "; ".join(e.o_errors)))
            else:
                lines.append("<span style='color:#e67e22'>⚠</span>&nbsp;&nbsp;"
                             "<span style='color:#888'>%s</span>&nbsp;&nbsp;"
                             "<span style='color:#e67e22; font-size:small'>%s</span>"
                             % (e.codi, "; ".join(e.o_errors + e.c_errors)))
        self.llista.setHtml("<br>".join(lines))
        self.btn_export.setText("Exportar (%d)" % r["exportables"])
        self.btn_export.setEnabled(r["exportables"] > 0)

    # ── accions ──

    def _browse_csv(self):
        path, _f = QtWidgets.QFileDialog.getSaveFileName(
            self, "Fitxer CSV de sortida", self.edt_csv.text(), "CSV (*.csv)")
        if path:
            self.edt_csv.setText(path)

    def _run(self):
        path = self.edt_csv.text().strip()
        if not path:
            QtWidgets.QMessageBox.warning(self, "Exportar", "Indica el fitxer CSV de sortida.")
            return
        precisio = self.edt_precision.text().strip()
        self.btn_export.setEnabled(False)
        self.btn_export.setText("Exportant…")
        try:
            files, n_coords, n_orient = db.files_csv(
                self._entrades, int(self.opts["epsg_utm"]), precisio,
                self.calc.orientacio, self.calc.coordenades)
            for f in files:
                log("  %s -> orient=%s coords=%s" % (f["Code"], f["Sector8"] or "no",
                                                    "si" if f["Lat_WGS84"] else "no"))
            escriure_csv(path, db.camps_csv(int(self.opts["epsg_utm"])), files)
        except Exception as exc:
            QtWidgets.QMessageBox.critical(self, "Exportar", "ERROR: %s" % exc)
            self._recompute()
            return
        QtWidgets.QMessageBox.information(
            self, "Exportar",
            "CSV generat: %s\n\nFiles: %d\nAmb coordenades: %d\nAmb orientació: %d"
            % (path, len(files), n_coords, n_orient))
        self.accept()


def show_dialog():
    if Metashape.app.document is None or not Metashape.app.document.chunks:
        Metashape.app.messageBox("Obri un projecte amb chunks abans d'exportar.")
        return
    opts = _opcions_validades("Exportar georef i orientacions")
    if opts is None:
        return
    dlg = ExportGeorefDialog(opts)
    dlg.exec_()


if __name__ == "__main__":
    show_dialog()
