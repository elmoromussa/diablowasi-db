"""Peces Qt per als diàlegs del menú Base de dades (dins de Metashape).

Còpia **retallada** de ``oculat/metashape_addon/qt_comu.py`` (Oculat ``52863de``; vegeu ``ORIGEN.md``): només el que
usen les tres eines de ``chachapoya_bd``. El codi conservat és idèntic a l'origen.

- ``InformeDialog`` / ``informe()``: quadre de resultat amb resum HTML i detall de text.
  Els ``QMessageBox`` dins de Metashape ixen massa estrets i amb botons en anglés.
- ``ConfirmDialog``: confirmació amb text HTML i botons en català (Cancel·lar per defecte).
- ``refrescar_gui()``: refresc de la GUI entre tasques, a la manera d'Agisoft
  (``Metashape.app.update()``). MAI cridar-ho des d'un callback ``progress=``.
- ``te_numpy``, ``instal_lar_paquet``, ``requereix_numpy``: numpy per a «Exportar mètriques de shapes».
"""

from __future__ import annotations

import Metashape

try:
    from PySide2 import QtCore, QtWidgets
except ImportError:
    from PySide6 import QtCore, QtWidgets  # type: ignore


class InformeDialog(QtWidgets.QDialog):
    def __init__(self, titol: str, resum: str, detall: list[str], parent=None):
        super().__init__(parent)
        self.setWindowTitle(titol)
        self.setMinimumWidth(560)
        lay = QtWidgets.QVBoxLayout(self)
        lbl = QtWidgets.QLabel(resum)
        lbl.setWordWrap(True)
        lbl.setTextFormat(QtCore.Qt.RichText)
        lay.addWidget(lbl)
        if detall:
            txt = QtWidgets.QPlainTextEdit("\n".join(detall))
            txt.setReadOnly(True)
            txt.setMinimumHeight(180)
            lay.addWidget(txt)
        btn = QtWidgets.QPushButton("D'acord")
        btn.setDefault(True)
        btn.clicked.connect(self.accept)
        row = QtWidgets.QHBoxLayout()
        row.addStretch()
        row.addWidget(btn)
        lay.addLayout(row)


def executar(dlg) -> int:
    """``exec_()`` (PySide2, el de Metashape) o ``exec()`` (PySide6)."""
    return getattr(dlg, "exec_", dlg.exec)()


def informe(parent, titol: str, resum: str, detall: list[str] | None = None) -> None:
    executar(InformeDialog(titol, resum, detall or [], parent))


class ConfirmDialog(QtWidgets.QDialog):
    def __init__(self, titol: str, html: str, accio: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle(titol)
        self.setMinimumWidth(520)
        lay = QtWidgets.QVBoxLayout(self)
        lay.setSpacing(12)
        lbl = QtWidgets.QLabel(html)
        lbl.setWordWrap(True)
        lbl.setTextFormat(QtCore.Qt.RichText)
        lay.addWidget(lbl)
        row = QtWidgets.QHBoxLayout()
        row.addStretch()
        btn_ok = QtWidgets.QPushButton(accio)
        btn_cancel = QtWidgets.QPushButton("Cancel·lar")
        btn_cancel.setDefault(True)
        btn_ok.clicked.connect(self.accept)
        btn_cancel.clicked.connect(self.reject)
        row.addWidget(btn_ok)
        row.addWidget(btn_cancel)
        lay.addLayout(row)


def confirmar(parent, titol: str, html: str, accio: str) -> bool:
    return bool(executar(ConfirmDialog(titol, html, accio, parent)))


def refrescar_gui() -> None:
    """Només des del fil de la GUI, entre tasques de Metashape."""
    try:
        Metashape.app.update()
    except Exception:
        QtWidgets.QApplication.processEvents()


def te_numpy() -> bool:
    try:
        import numpy  # noqa: F401
        return True
    except ImportError:
        return False


def _ordre_pip(paquet: str) -> list[str]:
    import sys
    return [sys.executable, "-m", "pip", "install", "--user", paquet]


def instal_lar_paquet(paquet: str, parent=None) -> tuple[bool, str]:
    """Executa pip del Python de Metashape (a l'espai d'usuari, sense administrador).

    Manté la GUI viva mentre pip treballa (bucle amb ``refrescar_gui`` al fil principal).
    Retorna (èxit, eixida de pip).
    """
    import importlib
    import site
    import subprocess
    import time

    ordre = _ordre_pip(paquet)
    print("[oculat] " + " ".join(ordre))
    try:
        proc = subprocess.Popen(ordre, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                text=True, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except Exception as exc:
        return False, f"No s'ha pogut llançar pip: {exc}"
    while proc.poll() is None:
        refrescar_gui()
        time.sleep(0.2)
    eixida = proc.stdout.read() if proc.stdout else ""
    print(eixida)
    if proc.returncode != 0:
        return False, eixida
    # el paquet acaba de caure al user site: assegurar que és al sys.path i reintentar
    try:
        site.addsitedir(site.getusersitepackages())
    except Exception:
        pass
    importlib.invalidate_caches()
    try:
        importlib.import_module(paquet)
        return True, eixida
    except ImportError as exc:
        return False, eixida + f"\n\nInstal·lat però no importable encara ({exc}): reinicia Metashape."


def requereix_numpy(eina: str, parent=None) -> bool:
    """Si falta numpy, ofereix instal·lar-lo des de dins de Metashape (com fan els scripts d'Agisoft)."""
    if te_numpy():
        return True
    ordre = " ".join(f'"{p}"' if " " in p else p for p in _ordre_pip("numpy"))
    html = (f"<b>«{eina}» necessita el paquet <code>numpy</code></b>, que el Python de Metashape "
            f"no porta.<br><br>Es pot instal·lar ara mateix amb el <code>pip</code> de Metashape, "
            f"a la carpeta d'usuari (sense administrador). Cal connexió a internet i tarda "
            f"menys d'un minut. Després l'eina s'obri directament.<br><br>"
            f"<small>Equivalent manual, en PowerShell:<br><code>&amp; {ordre}</code></small>")
    if not confirmar(parent, "Falta numpy", html, "Instal·lar ara"):
        return False
    ok, eixida = instal_lar_paquet("numpy", parent)
    if ok:
        informe(parent, "numpy instal·lat", "<b>numpy instal·lat correctament.</b>",
                eixida.strip().splitlines()[-8:])
        return True
    informe(parent, "No s'ha pogut instal·lar numpy",
            f"<b>La instal·lació ha fallat.</b> Prova-ho a mà en PowerShell:<br><code>&amp; {ordre}</code>",
            eixida.strip().splitlines()[-25:])
    return False
