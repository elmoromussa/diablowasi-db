"""``chachapoya_bd`` no depén d'Oculat ni de Pontarch en temps d'execució, i la lògica pura no depén de Metashape ni de Qt.

Anàlisi AST de tots els ``import`` (també els tardans dins de funcions), com ``test_contracte_imports.py`` de Pontarch.
"""

from __future__ import annotations

import ast
from pathlib import Path

import pytest

PAQUET = Path(__file__).resolve().parent.parent
ALIENS = {"pipeline", "metashape_addon", "pontarch", "oculat"}
GUI = {"Metashape", "PySide2", "PySide6"}


def _fitxers(*patrons: str) -> list[Path]:
    return sorted({p for pat in patrons for p in PAQUET.glob(pat) if "__pycache__" not in p.parts})


TOTS = _fitxers("**/*.py")
PURS = _fitxers("db/*.py", "configuracio.py", "_vendor/contracte/*.py")


def _id(p: Path) -> str:
    return p.relative_to(PAQUET).as_posix()


def _arrels(ruta: Path) -> list[str]:
    """Paquet arrel de cada ``import`` absolut del fitxer (els relatius queden dins del paquet)."""
    out = []
    for node in ast.walk(ast.parse(ruta.read_text(encoding="utf-8"))):
        if isinstance(node, ast.Import):
            out += [a.name.split(".")[0] for a in node.names]
        elif isinstance(node, ast.ImportFrom) and not node.level:
            out.append((node.module or "").split(".")[0])
    return out


@pytest.mark.parametrize("ruta", TOTS, ids=_id)
def test_cap_import_d_oculat_ni_de_pontarch(ruta):
    aliens = ALIENS.intersection(_arrels(ruta))
    assert not aliens, f"{_id(ruta)} importa {sorted(aliens)}"


@pytest.mark.parametrize("ruta", PURS, ids=_id)
def test_logica_pura_sense_metashape_ni_qt(ruta):
    gui = GUI.intersection(_arrels(ruta))
    assert not gui, f"{_id(ruta)} importa {sorted(gui)}"


def test_hi_ha_fitxers_a_comprovar():
    assert any(p.parent.name == "db" for p in PURS)
    assert len(TOTS) > len(PURS)
