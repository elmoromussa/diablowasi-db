"""Configuració del menú a ``<jaciment>/bd.yaml`` (configuracio.py), llegida amb ``escalars`` del contracte."""

from __future__ import annotations

from pathlib import Path

import pytest

from chachapoya_bd import configuracio as cfg
from chachapoya_bd._vendor.contracte import jaciment_yaml
from chachapoya_bd.db import codis as db_codis

EXEMPLE = Path(cfg.__file__).resolve().parent / "bd.yaml.example"

# La regex de Diablo Wasi tal com és a bd.yaml.example, escrita ací a part per a comparar-la caràcter a caràcter.
REGEX_DW = r"^(?P<codi>(?P<sector>(LP|DW)-[A-Z0-9]+)-EA\d+)(?:-(?P<variant>.+))?$"
SECTORS_DW = ["LP-G", "LP-N", "LP-C", "LP-U", "LP-S",
              "DW-G", "DW-S01", "DW-S02", "DW-S03", "DW-S04", "DW-S05", "DW-S06"]
# jaciment.yaml sintètic: cap coordenada ni shift.
JACIMENT_32718 = 'versio: 2\nnom: "Prova"\nepsg: 32718\nshift: null\n'


def _text_exemple() -> str:
    return EXEMPLE.read_text(encoding="utf-8")


def _text_dw() -> str:
    """bd.yaml de DW: l'exemple amb les tres línies de la necròpolis descomentades (com ho fa l'usuari)."""
    linies = _text_exemple().splitlines()
    i = next(n for n, li in enumerate(linies) if li.startswith("# ── Exemple complet"))
    out = linies[:i]
    for li in linies[i:]:
        if li.startswith(("# codi_regex:", "# sectors:", "# epsg_utm:")):
            li = li[2:]
        out.append(li)
    return "\n".join(out) + "\n"


def _jaciment(tmp_path: Path, bd: str | None = None, jaciment: str | None = None) -> Path:
    carpeta = tmp_path / "jaciment"
    carpeta.mkdir()
    if bd is not None:
        (carpeta / cfg.FITXER_BD).write_text(bd, encoding="utf-8")
    if jaciment is not None:
        (carpeta / "jaciment.yaml").write_text(jaciment, encoding="utf-8")
    return carpeta


def _meta(carpeta: Path) -> dict:
    return {"jaciment/carpeta": str(carpeta), "jaciment/campanya": "2025_prova"}


# ── bd.yaml.example ──────────────────────────────────────────────────────────


def test_exemple_te_les_claus_del_bloc_database():
    d = jaciment_yaml.escalars(_text_exemple())
    assert set(d) == set(cfg.CLAUS) - {"epsg_utm"}          # epsg_utm, comentada: ve de jaciment.yaml
    assert d["codi_regex"] == r"^(?P<codi>EA\d+)(?:_(?P<variant>.+))?$"
    assert d["sectors"] == []
    assert (d["epsg_geo"], d["precisio_m"], d["eix_facana"]) == (4326, 0.05, "frontal")
    assert (d["tilt_avis_deg"], d["min_horitz_deg"], d["divergencia_avis_m"]) == (15, 20, 0.25)


def test_bd_yaml_de_dw_regex_caracter_a_caracter():
    text = _text_dw()
    # La regex del fitxer, entre les cometes simples de la línia descomentada, sense passar pel lector.
    linia = next(li for li in text.splitlines() if li.startswith("codi_regex:") and "LP|DW" in li)
    al_fitxer = linia.split("'", 1)[1].rsplit("'", 1)[0]
    assert al_fitxer == REGEX_DW
    d = jaciment_yaml.escalars(text)
    assert d["codi_regex"] == REGEX_DW
    assert list(d["codi_regex"]) == list(REGEX_DW)
    assert "\\d" in d["codi_regex"] and "(?:" in d["codi_regex"] and "(?P<sector>" in d["codi_regex"]


def test_bd_yaml_de_dw_sectors_com_a_llista():
    d = jaciment_yaml.escalars(_text_dw())
    assert d["sectors"] == SECTORS_DW
    assert all(isinstance(s, str) for s in d["sectors"])
    assert d["epsg_utm"] == 32718


def test_bd_yaml_de_dw_reconeix_els_codis():
    o = cfg.opcions_de(_text_dw())
    assert db_codis.validar_opcions(o) == []
    c = db_codis.Codis(o)
    assert c.analitzar("DW-S01-EA02") == db_codis.Codi("DW-S01-EA02", None, "DW-S01")
    assert c.analitzar("DW-S01-EA02-2025") == db_codis.Codi("DW-S01-EA02", "2025", "DW-S01")
    assert c.analitzar("EA02") is None
    assert c.sector_fora_cataleg("DW-S09-EA01") == "DW-S09"
    assert c.sector_fora_cataleg("LP-G-EA01") is None


# ── EPSG UTM: bd.yaml > jaciment.yaml ────────────────────────────────────────


def test_epsg_utm_de_bd_yaml_mana(tmp_path):
    carpeta = _jaciment(tmp_path, bd=_text_exemple() + "epsg_utm: 25830\n", jaciment=JACIMENT_32718)
    o = cfg.opcions(carpeta)
    assert (o["epsg_utm"], o["font_epsg_utm"]) == (25830, "bd.yaml")


def test_epsg_utm_de_jaciment_yaml_si_bd_yaml_no_en_te(tmp_path):
    carpeta = _jaciment(tmp_path, bd=_text_exemple(), jaciment=JACIMENT_32718)
    o = cfg.opcions(carpeta)
    assert (o["epsg_utm"], o["font_epsg_utm"]) == (32718, "jaciment.yaml")
    assert db_codis.validar_opcions(o) == []


def test_crs_de_jaciment():
    assert cfg.crs_de_jaciment(JACIMENT_32718) == "EPSG:32718"
    assert cfg.crs_de_jaciment('nom: "Prova"\n') is None
    assert cfg.crs_de_jaciment("epsg: true\n") is None
    assert cfg.crs_de_jaciment(None) is None


# ── Errors que es mostren a l'usuari ─────────────────────────────────────────


def test_sense_vincle():
    with pytest.raises(cfg.ErrorConfiguracio) as e:
        cfg.opcions_de_meta({})
    assert str(e.value) == cfg.ERROR_SENSE_VINCLE
    assert "jaciment/carpeta" in str(e.value)


def test_vincle_heretat_de_pontarch(tmp_path):
    carpeta = _jaciment(tmp_path, bd=_text_exemple())
    meta = {"pontarch/lliurament": jaciment_yaml.text_meta_lliurament(carpeta, "2025_prova")}
    assert cfg.carpeta_jaciment(meta) == carpeta


def test_carpeta_del_vincle_inexistent(tmp_path):
    with pytest.raises(cfg.ErrorConfiguracio) as e:
        cfg.opcions_de_meta(_meta(tmp_path / "no_hi_es"))
    assert str(e.value) == cfg.error_sense_carpeta(tmp_path / "no_hi_es")


def test_sense_bd_yaml(tmp_path):
    carpeta = _jaciment(tmp_path, jaciment=JACIMENT_32718)
    with pytest.raises(cfg.ErrorConfiguracio) as e:
        cfg.opcions_de_meta(_meta(carpeta))
    assert str(e.value) == cfg.error_sense_bd(carpeta)
    assert str(carpeta / "bd.yaml") in str(e.value) and "bd.yaml.example" in str(e.value)


def test_sense_codi_regex(tmp_path):
    carpeta = _jaciment(tmp_path, bd="epsg_utm: 32718\n")
    o = cfg.opcions(carpeta)
    assert db_codis.validar_opcions(o) == [db_codis.ERROR_CODI_REGEX]
    assert db_codis.validar_opcions(o, requereix_utm=False) == [db_codis.ERROR_CODI_REGEX]
    assert "codi_regex" in db_codis.ERROR_CODI_REGEX and "bd.yaml" in db_codis.ERROR_CODI_REGEX


def test_sense_epsg(tmp_path):
    carpeta = _jaciment(tmp_path, bd=_text_exemple())          # cap jaciment.yaml, cap epsg_utm
    o = cfg.opcions(carpeta)
    assert (o["epsg_utm"], o["font_epsg_utm"]) == (None, None)
    assert db_codis.validar_opcions(o) == [db_codis.ERROR_EPSG_UTM]
    assert db_codis.validar_opcions(o, requereix_utm=False) == []   # mètriques: no projecta a UTM
    assert "bd.yaml" in db_codis.ERROR_EPSG_UTM and "jaciment.yaml" in db_codis.ERROR_EPSG_UTM


# ── Carpeta d'eixida ─────────────────────────────────────────────────────────


def test_carpeta_eixida_per_defecte(tmp_path):
    carpeta = _jaciment(tmp_path, bd=_text_exemple())
    assert cfg.carpeta_eixida(carpeta) == carpeta / "BD"
    assert cfg.carpeta_eixida_de_meta(_meta(carpeta)) == carpeta / "BD"
    assert not (carpeta / "BD").exists()                       # llegir la configuració no crea res
