"""
exportar_metriques_bd.py
────────────────────────────────────────────────────────────────────────────
Extractor de mètriques des de shapes 3D dibuixades sobre el model. Escriu 4
CSV al directori triat (una passada per tot el corpus marcat):

  metriques_estructures.csv   una fila per estructura (chunk)
  metriques_murs.csv          una fila per (estructura, mur, cos)
  metriques_portals.csv       una fila per portal, amb ràtios _cons i _orig
  metriques_shapes_audit.csv  una fila per shape: tot el calculat + QC

Els chunks candidats són els que tenen shapes; es preseleccionen els actius.
El codi de l'estructura és l'etiqueta del chunk (sense variant). Només lectura.

Compatible amb Metashape Professional 2.3.0
Ús: Extensions → Base de dades → Exportar mètriques de shapes…
Gramàtica, mètriques i QC: ``pipeline/db/shapes.py`` (amb tests).
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
    import numpy as np
except ImportError:          # numpy absent: show_dialog() avisa i no obri res
    np = None  # type: ignore

try:
    from PySide2 import QtCore, QtWidgets
except ImportError:
    from PySide6 import QtCore, QtWidgets  # type: ignore

_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from chachapoya_bd._vendor.qt_comu import informe, requereix_numpy  # noqa: E402
from chachapoya_bd import configuracio  # noqa: E402
from chachapoya_bd.db import vistes as om  # noqa: E402
from chachapoya_bd.db import codis as db_codis  # noqa: E402

try:
    from chachapoya_bd.db import shapes as db
except ImportError:
    db = None  # type: ignore

GEOCENTRIC_MIN_NORM = 1_000_000.0   # centre amb norma > 1000 km = coordenades geocèntriques


def log(msg: str) -> None:
    print("[exportMetriques] " + msg)


def _opcions() -> dict:
    """Opcions del ``bd.yaml`` que toca (el del jaciment vinculat o el del repo), via ``configuracio.py``.

    Hi afig ``fitxer_bd`` i ``font_bd`` per a la línia informativa del diàleg.
    No amaga cap error: si no hi ha cap bd.yaml, l'excepció puja fins a
    ``show_dialog`` i s'ensenya a l'usuari.
    """
    doc = Metashape.app.document
    c = configuracio.configuracio(doc.meta, doc.path)
    return dict(c.opcions, fitxer_bd=str(c.fitxer_bd), font_bd=c.font)


def _opcions_validades(titol: str) -> dict | None:
    """Opcions llestes per al diàleg, o None (ja s'ha informat l'usuari)."""
    try:
        opts = _opcions()
    except Exception as e:  # cap bd.yaml, o il·legible
        informe(None, titol, "No s'ha pogut llegir la configuració de la base de dades.", [str(e)])
        return None
    errors = db_codis.validar_opcions(opts, requereix_utm=False)   # només usa epsg_geo
    if errors:
        informe(None, titol, "Falten valors obligatoris a bd.yaml.",
                errors)
        return None
    log("bd.yaml: %s (%s)" % (opts["fitxer_bd"], configuracio.DESCRIPCIO_FONT[opts["font_bd"]]))
    log("codi_regex: %s" % opts["codi_regex"])
    log("epsg_geo: EPSG:%s" % opts["epsg_geo"])
    return opts


def _carpeta_defecte() -> str:
    try:
        doc = Metashape.app.document
        return str(configuracio.carpeta_eixida(doc.meta, doc.path))
    except Exception:
        return os.path.expanduser("~")


# ── Metashape → dades pures ──────────────────────────────────────────────────


def _np(v) -> np.ndarray:
    return np.array([v.x, v.y, v.z], dtype=float)


def _eixos_regio(chunk) -> dict[str, tuple]:
    """{pla: (X, Y, Z)} en coordenades INTERNES del chunk (Metashape.Vector)."""
    rot = chunk.region.rot
    rx, ry, rz = rot.col(0), rot.col(1), rot.col(2)
    return {pla: om.eixos_per_vista(rx, ry, rz, vista) for pla, vista in db.PLA_VISTA.items()}


def marcs_del_chunk(chunk, geo) -> dict[str, db.Marc]:
    """Marcs candidats: món (geocèntric o local) i intern (escala aplicada).

    La vertical depèn del context del chunk, no del mode de les shapes: en un
    chunk georeferenciat el món és geocèntric i (0,0,1) apunta a l'eix polar,
    no amunt (a Leymebamba ~87° d'error). Es deriva del ``localframe`` quan el
    centre viu en coordenades geocèntriques.
    """
    T = chunk.transform.matrix
    R = chunk.transform.rotation
    center_w = T.mulp(chunk.region.center)
    if center_w.norm() > GEOCENTRIC_MIN_NORM:
        UPw = ((geo.localframe(center_w).rotation().t()) * Metashape.Vector([0, 0, 1])).normalized()
    else:
        UPw = Metashape.Vector([0, 0, 1])
    eixos = _eixos_regio(chunk)

    mon = db.Marc(
        eixos={p: tuple(_np((R * a).normalized()) for a in axs) for p, axs in eixos.items()},
        centre=_np(center_w), amunt=_np(UPw))
    s = chunk.transform.scale if chunk.transform.scale else 1.0
    intern = db.Marc(
        eixos={p: tuple(_np(a.normalized()) for a in axs) for p, axs in eixos.items()},
        centre=_np(chunk.region.center * s), amunt=_np((R.t() * UPw).normalized()))
    return {"mon-local": mon, "intern": intern}


def shapes_del_chunk(chunk, geo, notes: list[str]) -> tuple[list[db.Shape], bool]:
    """Converteix les shapes del chunk a ``db.Shape``. Retorna (shapes, geografic).

    Si les coordenades són geogràfiques (lat/lon), es desprojecten amb el CRS
    de les shapes a geocèntriques (marc "mon-local"). Les 2D queden amb
    ``punts=None`` (llegades d'orthos).
    """
    items = []
    allc = []
    for sh in chunk.shapes:
        g = sh.geometry
        label = sh.label or "(sense etiqueta)"
        try:
            if g.type == Metashape.Geometry.Type.PolygonType:
                cs = list(g.coordinates[0])
                tipus = "polygon"
            else:
                cs = list(g.coordinates)
                tipus = "line"
        except Exception as exc:
            notes.append("[%s] error de geometria: %s" % (label, exc))
            continue
        if len(cs) >= 2:
            items.append((label, tipus, cs))
            allc.extend(cs)
    if not items:
        return [], False

    xs = [c.x for c in allc]
    ys = [c.y for c in allc]
    spread = max(max(xs) - min(xs), max(ys) - min(ys))
    geografic = (max(map(abs, xs)) <= 180 and max(map(abs, ys)) <= 90 and spread < 0.01)
    scrs = chunk.shapes.crs

    out = []
    for label, tipus, cs in items:
        if any(c.size < 3 for c in cs):
            out.append(db.Shape(label, tipus, None, len(cs)))
            continue
        if geografic:
            pts = [scrs.unproject(Metashape.Vector([c.x, c.y, c.z])) for c in cs]
        else:
            pts = [Metashape.Vector([c.x, c.y, c.z]) for c in cs]
        out.append(db.Shape(label, tipus, np.array([[p.x, p.y, p.z] for p in pts])))
    return out, geografic


def analitzar_chunk(chunk, geo, codis: db_codis.Codis) -> db.Resultat:
    notes: list[str] = []
    code = chunk.label or ""
    c = codis.analitzar(chunk.label)
    if c:
        code = c.codi
        if c.variant:
            notes.append("chunk variant '%s': codi %s" % (chunk.label, c.codi))
    else:
        notes.append("etiqueta de chunk fora de la gramatica de codis")

    if chunk.shapes is None or len(chunk.shapes) == 0:
        r = db.Resultat(code=code, chunk_label=chunk.label or "", notes=notes)
        r.notes.append("sense shapes")
        return r

    shapes, geografic = shapes_del_chunk(chunk, geo, notes)
    marcs = marcs_del_chunk(chunk, geo)
    if geografic:
        marcs = {"geografic": marcs["mon-local"]}
    return db.analitzar(code, chunk.label or "", shapes, marcs, notes)


# ── Diàleg ───────────────────────────────────────────────────────────────────


def _mk_item(text, checked, data):
    item = QtWidgets.QListWidgetItem(text)
    item.setFlags(item.flags() | QtCore.Qt.ItemIsUserCheckable)
    item.setCheckState(QtCore.Qt.Checked if checked else QtCore.Qt.Unchecked)
    item.setData(QtCore.Qt.UserRole, data)
    return item


class ExportMetriquesDialog(QtWidgets.QDialog):
    def __init__(self, opts: dict, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Exportar mètriques de shapes (BD)")
        self.setMinimumWidth(680)
        self.opts = opts
        self.codis = db_codis.Codis(self.opts)
        self._geo = Metashape.CoordinateSystem("EPSG::%d" % int(self.opts["epsg_geo"]))
        self._cache: dict[str, db.Resultat] = {}
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
        for text, val in (("Tots", True), ("Cap", False)):
            b = QtWidgets.QPushButton(text)
            b.setFixedWidth(60)
            b.clicked.connect(lambda _c=False, val=val: self._set_all(val))
            row_b.addWidget(b)
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
        default_dir = os.path.join(_carpeta_defecte(),
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
            "<small>4 CSV: estructures, murs (mur × cos, amb originals, prismes i índex de "
            "conservació), portals (ràtios _cons i _orig) i audit per shape amb els QC. "
            "Només lectura del projecte. Configuració: <code>%s</code> (%s); EPSG:%d (geo).</small>"
            % (html.escape(self.opts["fitxer_bd"]), configuracio.DESCRIPCIO_FONT[self.opts["font_bd"]],
               self.opts["epsg_geo"]))
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

        self.lst.itemChanged.connect(self._recompute)
        self._rescan()

    def _rescan(self):
        prev = None
        if self.lst.count():
            prev = {self.lst.item(i).data(QtCore.Qt.UserRole):
                    self.lst.item(i).checkState() == QtCore.Qt.Checked
                    for i in range(self.lst.count())}
        self._loading = True
        self.lst.clear()
        self._cache = {}
        for chunk in Metashape.app.document.chunks:
            if chunk.shapes is None or len(chunk.shapes) == 0:
                continue
            log("analitzant %s ..." % chunk.label)
            try:
                r = analitzar_chunk(chunk, self._geo, self.codis)
            except Exception as exc:
                log("  ERROR a %s: %s" % (chunk.label, exc))
                continue
            self._cache[chunk.label] = r
            tags = ["%d shapes" % r.n_shapes]
            if r.n_invalid or r.n_2d:
                tags.append("%d inv/2D" % (r.n_invalid + r.n_2d))
            if r.n_qc:
                tags.append("%d QC" % r.n_qc)
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

    def _selected(self) -> list[db.Resultat]:
        return [self._cache[self.lst.item(i).data(QtCore.Qt.UserRole)]
                for i in range(self.lst.count())
                if self.lst.item(i).checkState() == QtCore.Qt.Checked]

    def _recompute(self, *_a):
        if self._loading:
            return
        sel = self._selected()
        n = len(sel)
        tot_v = sum(r.n_valid for r in sel)
        tot_i = sum(r.n_invalid + r.n_2d for r in sel)
        tot_q = sum(r.n_qc for r in sel)
        resum = "<b>%d</b> estructur%s" % (n, "es" if n != 1 else "a")
        resum += "&nbsp;&nbsp;<span style='color:#2ecc71'>(%d shapes vàlides)</span>" % tot_v
        if tot_i:
            resum += "&nbsp;&nbsp;<span style='color:#e67e22'>(%d invàlides/2D)</span>" % tot_i
        if tot_q:
            resum += "&nbsp;&nbsp;<span style='color:#3498db'>(%d avisos QC)</span>" % tot_q
        self.lbl_resum.setText(resum)

        lines = []
        for r in sel:
            probs = r.problemes()
            icon = "\u2713" if not probs else "\u26a0"
            color = "#2ecc71" if not probs else "#e67e22"
            lines.append("<span style='color:%s'>%s</span>&nbsp;&nbsp;<b>%s</b>" % (color, icon, r.code))
            for p in probs[:6]:
                lines.append("&nbsp;&nbsp;&nbsp;<span style='color:#888; font-size:small'>%s</span>" % p)
            if len(probs) > 6:
                lines.append("&nbsp;&nbsp;&nbsp;<span style='color:#aaa; font-size:small'>"
                             "... i %d més</span>" % (len(probs) - 6))
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
        sel = self._selected()
        self.btn_export.setEnabled(False)
        self.btn_export.setText("Exportant\u2026")
        try:
            res = db.exportar(sel, folder)
        except Exception as exc:
            QtWidgets.QMessageBox.critical(self, "Exportar", "ERROR: %s" % exc)
            self._recompute()
            return
        QtWidgets.QMessageBox.information(
            self, "Exportar",
            "Exportació completada a:\n%s\n\nestructures: %d files\nmurs: %d\nportals: %d\n"
            "shapes (audit): %d" % (folder, res["estructures"][1], res["murs"][1],
                                    res["portals"][1], res["shapes"][1]))
        self.accept()


def show_dialog():
    if not requereix_numpy("Exportar mètriques de shapes"):
        return
    if Metashape.app.document is None or not Metashape.app.document.chunks:
        Metashape.app.messageBox("Obri un projecte amb chunks abans d'exportar.")
        return
    opts = _opcions_validades("Exportar mètriques de shapes")
    if opts is None:
        return
    dlg = ExportMetriquesDialog(opts)
    dlg.exec_()


if __name__ == "__main__":
    show_dialog()
