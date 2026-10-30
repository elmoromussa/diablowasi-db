"""Els fitxers de ``_vendor/`` són els que declara ``_vendor/ORIGEN.md`` (sha256) i, amb ``OCULAT_REPO``, els d'Oculat."""

from __future__ import annotations

import hashlib
import os
import re
from pathlib import Path

import pytest

VENDOR = Path(__file__).resolve().parent.parent / "_vendor"
ORIGEN = VENDOR / "ORIGEN.md"
CONTRACTE = ("claus.py", "utm.py", "jaciment_yaml.py")

_FILA = re.compile(r"^\| `(?P<fitxer>[^`]+)` \| (?P<origen>[^|]+) \| (?P<commit>[^|]+) \| `(?P<sha>[0-9a-f]{64})` \| "
                   r"(?P<estat>[^|]+) \|$")


def _bytes(p: Path) -> bytes:
    # Finals de línia normalitzats: el sha256 d'ORIGEN.md és el del fitxer amb LF, com al repositori.
    return p.read_bytes().replace(b"\r\n", b"\n")


def _files() -> dict[str, dict[str, str]]:
    out = {}
    for linia in ORIGEN.read_text(encoding="utf-8").splitlines():
        m = _FILA.match(linia.strip())
        if m:
            out[m["fitxer"]] = {k: v.strip() for k, v in m.groupdict().items()}
    return out


def _vendoritzats() -> list[str]:
    return sorted(p.relative_to(VENDOR).as_posix() for p in VENDOR.rglob("*.py") if "__pycache__" not in p.parts)


def test_cada_fitxer_te_fila_a_origen():
    assert _vendoritzats() == sorted(_files())


@pytest.mark.parametrize("fitxer", _vendoritzats())
def test_sha256_igual_a_origen(fitxer):
    fila = _files()[fitxer]
    assert hashlib.sha256(_bytes(VENDOR / fitxer)).hexdigest() == fila["sha"]


def test_contracte_copiat_byte_a_byte_de_52863de():
    files = _files()
    for nom in CONTRACTE:
        fila = files[f"contracte/{nom}"]
        assert fila["estat"] == "idèntic"
        assert fila["commit"] == "`52863de`"
        assert fila["origen"] == f"`oculat/pipeline/contracte/{nom}`"


@pytest.mark.skipif(not os.environ.get("OCULAT_REPO"), reason="OCULAT_REPO no definida")
@pytest.mark.parametrize("nom", CONTRACTE)
def test_contracte_identic_a_oculat(nom):
    origen = Path(os.environ["OCULAT_REPO"]) / "oculat" / "pipeline" / "contracte" / nom
    assert _bytes(VENDOR / "contracte" / nom) == _bytes(origen)
