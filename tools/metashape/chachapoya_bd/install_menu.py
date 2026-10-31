"""
Menú Metashape — chachapoya_bd (exportadors cap a la base de dades del jaciment)

Instal·lació:
  1. Metashape → Tools → Run Script → tools/metashape/chachapoya_bd/install_menu.py
     (o afegeix install_menu.py com a script d'inici).

Menús:
  Extensions → Base de dades   → Exportar georef i orientacions…
                               → Exportar mètriques de shapes…
                               → Auditar shapes del chunk actiu

Configuració: bd.yaml del jaciment vinculat si n'hi ha; si no, el bd.yaml del repositori (este paquet; vegeu
bd.yaml.example).

Mateixes etiquetes que el menú d'Oculat: les tres entrades es lleven abans d'afegir-les, així que mana l'últim
que s'instal·la.
"""

from __future__ import annotations

import sys
from pathlib import Path

import Metashape

_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

MENU_BD    = "Extensions/Base de dades"

MENU_BD_GEOREF   = f"{MENU_BD}/Exportar georef i orientacions..."
MENU_BD_METRIQ   = f"{MENU_BD}/Exportar mètriques de shapes..."
MENU_BD_AUDIT    = f"{MENU_BD}/Auditar shapes del chunk actiu"


def _run_bd_georef():
    from chachapoya_bd.exportar_georef_bd import show_dialog
    show_dialog()


def _run_bd_metriques():
    from chachapoya_bd.exportar_metriques_bd import show_dialog
    show_dialog()


def _run_bd_auditar():
    from chachapoya_bd.auditar_shapes import run
    run()


MENU_ITEMS = (
    (MENU_BD_GEOREF,   _run_bd_georef),
    (MENU_BD_METRIQ,   _run_bd_metriques),
    (MENU_BD_AUDIT,    _run_bd_auditar),
)


def register_menu():
    for label, _ in MENU_ITEMS:
        try:
            Metashape.app.removeMenuItem(label)
        except Exception:
            pass
    for label, func in MENU_ITEMS:
        Metashape.app.addMenuItem(label, func)
    print("chachapoya_bd: menú registrat (Extensions → Base de dades)")
    try:
        import numpy  # noqa: F401
    except ImportError:
        print("chachapoya_bd: numpy no està instal·lat al Python de Metashape; «Exportar mètriques de shapes» "
              "oferirà instal·lar-lo en obrir-se.")


register_menu()
