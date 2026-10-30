"""Gramàtica del codi d'estructura (bd.yaml)."""

from __future__ import annotations

import pytest

from chachapoya_bd.db import codis as m

# Dades de prova: nomenclatura d'una necròpolis amb sectors (Diablo Wasi).
OPTS_DW = {
    "codi_regex": r"^(?P<codi>(?P<sector>(LP|DW)-[A-Z0-9]+)-EA\d+)(?:-(?P<variant>.+))?$",
    "sectors": ["LP-G", "LP-N", "DW-G", "DW-S01", "DW-S02"],
    "epsg_utm": 32718,
}


def test_regex_amb_sectors():
    c = m.Codis(OPTS_DW)
    assert c.analitzar("DW-S01-EA01") == m.Codi("DW-S01-EA01", None, "DW-S01")
    assert c.analitzar("DW-S01-EA01-2018") == m.Codi("DW-S01-EA01", "2018", "DW-S01")
    assert c.analitzar("LP-G-EA12") is not None
    assert c.analitzar("EA01_250901_UE001") is None
    assert c.analitzar(None) is None


def test_exacte_i_sector():
    c = m.Codis(OPTS_DW)
    assert c.es_exacte("DW-S01-EA01") == "DW-S01-EA01"
    assert c.es_exacte("DW-S01-EA01-2018") is None
    assert c.sector_fora_cataleg("DW-S01-EA01") is None
    assert c.sector_fora_cataleg("DW-S99-EA01") == "DW-S99"


def test_regex_propia_sense_sector():
    c = m.Codis({"codi_regex": r"^(?P<codi>EA\d{2})(?:_(?P<variant>.+))?$", "sectors": []})
    assert c.analitzar("EA03_250901_UE012") == m.Codi("EA03", "250901_UE012", None)
    assert c.es_exacte("EA03") == "EA03"
    assert c.sector_fora_cataleg("EA03") is None


def test_regex_sense_grup_codi():
    with pytest.raises(ValueError):
        m.Codis({"codi_regex": r"^EA\d+$"})


def test_sense_regex_es_error():
    with pytest.raises(ValueError, match="codi_regex"):
        m.Codis()
    with pytest.raises(ValueError, match="codi_regex"):
        m.Codis({"epsg_utm": 25830})


def test_validar_opcions():
    errors = m.validar_opcions({})
    assert len(errors) == 2
    assert "codi_regex" in errors[0] and "bd.yaml" in errors[0]
    assert "epsg_utm" in errors[1] and "jaciment.yaml" in errors[1]
    assert m.validar_opcions(m.opcions(OPTS_DW)) == []
    assert m.validar_opcions(m.opcions({"codi_regex": "x"})) == [m.ERROR_EPSG_UTM]


def test_validar_opcions_sense_exigir_utm():
    # Diàleg de mètriques: la regex continua sent obligatòria, l'EPSG UTM no.
    assert m.validar_opcions(m.opcions({"codi_regex": "x"}), requereix_utm=False) == []
    assert m.validar_opcions({}, requereix_utm=False) == [m.ERROR_CODI_REGEX]


def test_epsg_utm_de_hierarchy_o_pipeline():
    o = m.opcions({}, crs_jaciment="EPSG:32718")
    assert (o["epsg_utm"], o["font_epsg_utm"]) == (32718, "jaciment.yaml")
    o = m.opcions({"epsg_utm": 25830}, crs_jaciment="EPSG:32718")
    assert (o["epsg_utm"], o["font_epsg_utm"]) == (25830, "bd.yaml")
    o = m.opcions({}, crs_jaciment=None)
    assert (o["epsg_utm"], o["font_epsg_utm"]) == (None, None)
    # CRS no UTM: utm_crs_info tornaria 25830 de reserva, però no s'accepta.
    o = m.opcions({}, crs_jaciment="EPSG:4326")
    assert (o["epsg_utm"], o["font_epsg_utm"]) == (None, None)


def test_opcions_fusiona_i_ignora_claus_desconegudes():
    o = m.opcions({"epsg_utm": 25830, "desconegut": 1, "sectors": []})
    assert o["epsg_utm"] == 25830
    assert o["epsg_geo"] == 4326
    assert o["codi_regex"] is None
    assert o["sectors"] == []
    assert "desconegut" not in o
