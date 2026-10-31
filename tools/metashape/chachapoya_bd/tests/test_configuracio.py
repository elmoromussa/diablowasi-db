"""Configuració del menú (configuracio.py): quin bd.yaml es llig, EPSG UTM i carpeta d'eixida."""

from __future__ import annotations

from pathlib import Path

import pytest

from chachapoya_bd import configuracio as cfg
from chachapoya_bd._vendor.contracte import jaciment_yaml
from chachapoya_bd.db import codis as db_codis

EXEMPLE = cfg.CARPETA_PAQUET / "bd.yaml.example"
BD_REPO = cfg.CARPETA_PAQUET / "bd.yaml"

# La regex de Diablo Wasi, escrita ací a part per a comparar-la caràcter a caràcter.
REGEX_DW = r"^(?P<codi>(?P<sector>(LP|DW)-[A-Z0-9]+)-EA\d+)(?:-(?P<variant>.+))?$"
SECTORS_DW = ["LP-G", "LP-N", "LP-C", "LP-U", "LP-S",
              "DW-G", "DW-S01", "DW-S02", "DW-S03", "DW-S04", "DW-S05", "DW-S06"]
# jaciment.yaml sintètic: cap coordenada ni shift.
JACIMENT_32718 = 'versio: 2\nnom: "Prova"\nepsg: 32718\nshift: null\n'
PSX = "projecte.psx"


def _text_exemple() -> str:
    return EXEMPLE.read_text(encoding="utf-8")


def _text_repo() -> str:
    return BD_REPO.read_text(encoding="utf-8")


def _regex_entre_cometes(linia: str) -> str:
    return linia.split("'", 1)[1].rsplit("'", 1)[0]


def _jaciment(tmp_path: Path, bd: str | None = None, jaciment: str | None = None) -> Path:
    carpeta = tmp_path / "jaciment"
    carpeta.mkdir()
    if bd is not None:
        (carpeta / cfg.FITXER_BD).write_text(bd, encoding="utf-8")
    if jaciment is not None:
        (carpeta / "jaciment.yaml").write_text(jaciment, encoding="utf-8")
    return carpeta


def _paquet(tmp_path: Path, bd: str | None = None) -> Path:
    """Carpeta de paquet sintètica, amb bd.yaml del «repositori» o sense."""
    carpeta = tmp_path / "paquet"
    carpeta.mkdir()
    if bd is not None:
        (carpeta / cfg.FITXER_BD).write_text(bd, encoding="utf-8")
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


# ── bd.yaml del repositori (Diablo Wasi) ─────────────────────────────────────


def test_bd_yaml_del_repo_regex_caracter_a_caracter():
    # La regex del fitxer i la de l'exemple comentat, entre les cometes simples, sense passar pel lector.
    al_repo = _regex_entre_cometes(next(li for li in _text_repo().splitlines() if li.startswith("codi_regex:")))
    a_exemple = _regex_entre_cometes(next(li for li in _text_exemple().splitlines()
                                          if li.startswith("# codi_regex:") and "LP|DW" in li))
    assert al_repo == a_exemple == REGEX_DW
    d = jaciment_yaml.escalars(_text_repo())
    assert list(d["codi_regex"]) == list(a_exemple)
    assert "\\d" in d["codi_regex"] and "(?:" in d["codi_regex"] and "(?P<sector>" in d["codi_regex"]


def test_bd_yaml_del_repo_sectors_com_a_llista():
    d = jaciment_yaml.escalars(_text_repo())
    assert d["sectors"] == SECTORS_DW and len(d["sectors"]) == 12
    assert all(isinstance(s, str) for s in d["sectors"])
    assert d["epsg_utm"] == 32718


def test_bd_yaml_del_repo_claus_i_valors():
    text = _text_repo()
    claus = [li.split(":", 1)[0] for li in text.splitlines() if li and not li.startswith(("#", " "))]
    assert sorted(claus) == sorted(cfg.CLAUS)                   # totes, i cap repetida
    d = jaciment_yaml.escalars(text)
    ex = jaciment_yaml.escalars(_text_exemple())
    for k in ("epsg_geo", "precisio_m", "eix_facana", "tilt_avis_deg", "min_horitz_deg", "divergencia_avis_m"):
        assert d[k] == ex[k], k


def test_bd_yaml_del_repo_reconeix_els_codis():
    o = cfg.opcions_de(_text_repo())
    assert db_codis.validar_opcions(o) == []
    c = db_codis.Codis(o)
    assert c.analitzar("DW-S01-EA02") == db_codis.Codi("DW-S01-EA02", None, "DW-S01")
    assert c.analitzar("DW-S01-EA02-2025") == db_codis.Codi("DW-S01-EA02", "2025", "DW-S01")
    assert c.analitzar("EA02") is None
    assert c.sector_fora_cataleg("DW-S09-EA01") == "DW-S09"
    assert c.sector_fora_cataleg("LP-G-EA01") is None


# ── Quin bd.yaml: el del jaciment si hi ha vincle i el fitxer hi és; si no, el del repo ──


def test_sense_vincle_llig_el_del_repo():
    c = cfg.configuracio({}, None)                               # paquet real
    assert (c.fitxer_bd, c.font) == (BD_REPO, cfg.FONT_REPO)
    assert c.opcions["codi_regex"] == REGEX_DW
    assert (c.opcions["epsg_utm"], c.opcions["font_epsg_utm"]) == (32718, str(BD_REPO))
    assert db_codis.validar_opcions(c.opcions) == []


def test_vincle_sense_bd_yaml_llig_el_del_repo(tmp_path):
    carpeta = _jaciment(tmp_path, jaciment=JACIMENT_32718)
    paquet = _paquet(tmp_path, bd=_text_repo())
    c = cfg.configuracio(_meta(carpeta), None, paquet)
    assert (c.fitxer_bd, c.font) == (paquet / "bd.yaml", cfg.FONT_REPO)
    assert c.carpeta_eixida == carpeta / "BD"                   # el vincle continua manant per a l'eixida


def test_vincle_amb_bd_yaml_llig_el_del_jaciment(tmp_path):
    carpeta = _jaciment(tmp_path, bd=_text_exemple(), jaciment=JACIMENT_32718)
    paquet = _paquet(tmp_path, bd=_text_repo())
    c = cfg.configuracio(_meta(carpeta), None, paquet)
    assert (c.fitxer_bd, c.font) == (carpeta / "bd.yaml", cfg.FONT_JACIMENT)
    assert c.opcions["codi_regex"] == r"^(?P<codi>EA\d+)(?:_(?P<variant>.+))?$"


def test_vincle_heretat_de_pontarch(tmp_path):
    carpeta = _jaciment(tmp_path, bd=_text_exemple())
    meta = {"pontarch/lliurament": jaciment_yaml.text_meta_lliurament(carpeta, "2025_prova")}
    assert cfg.carpeta_jaciment(meta) == carpeta
    assert cfg.configuracio(meta, None, _paquet(tmp_path)).font == cfg.FONT_JACIMENT


def test_carpeta_del_vincle_inexistent_llig_el_del_repo(tmp_path):
    paquet = _paquet(tmp_path, bd=_text_repo())
    meta = _meta(tmp_path / "no_hi_es")
    assert cfg.carpeta_jaciment(meta) is None
    c = cfg.configuracio(meta, None, paquet)
    assert (c.fitxer_bd, c.font) == (paquet / "bd.yaml", cfg.FONT_REPO)


def test_cap_fitxer_error_amb_les_dues_rutes(tmp_path):
    carpeta = _jaciment(tmp_path, jaciment=JACIMENT_32718)
    paquet = _paquet(tmp_path)
    with pytest.raises(cfg.ErrorConfiguracio) as e:
        cfg.configuracio(_meta(carpeta), None, paquet)
    assert str(e.value) == cfg.error_sense_bd(carpeta, paquet)
    assert str(carpeta / "bd.yaml") in str(e.value) and str(paquet / "bd.yaml") in str(e.value)
    assert "bd.yaml.example" in str(e.value)


def test_cap_fitxer_sense_vincle(tmp_path):
    paquet = _paquet(tmp_path)
    with pytest.raises(cfg.ErrorConfiguracio) as e:
        cfg.configuracio({}, None, paquet)
    assert "<jaciment>/bd.yaml" in str(e.value) and "no està vinculat" in str(e.value)
    assert str(paquet / "bd.yaml") in str(e.value)


# ── EPSG UTM: bd.yaml > jaciment.yaml (només amb vincle) ─────────────────────


def test_epsg_utm_de_bd_yaml_mana(tmp_path):
    carpeta = _jaciment(tmp_path, bd=_text_exemple() + "epsg_utm: 25830\n", jaciment=JACIMENT_32718)
    o = cfg.configuracio(_meta(carpeta), None, _paquet(tmp_path)).opcions
    assert (o["epsg_utm"], o["font_epsg_utm"]) == (25830, str(carpeta / "bd.yaml"))


def test_epsg_utm_de_jaciment_yaml_si_bd_yaml_no_en_te(tmp_path):
    carpeta = _jaciment(tmp_path, bd=_text_exemple(), jaciment=JACIMENT_32718)
    o = cfg.configuracio(_meta(carpeta), None, _paquet(tmp_path)).opcions
    assert (o["epsg_utm"], o["font_epsg_utm"]) == (32718, str(carpeta / "jaciment.yaml"))
    assert db_codis.validar_opcions(o) == []


def test_epsg_de_jaciment_yaml_nomes_amb_vincle(tmp_path):
    # Un jaciment.yaml al costat del .psx no compta si el projecte no està vinculat.
    carpeta = _jaciment(tmp_path, jaciment=JACIMENT_32718)
    paquet = _paquet(tmp_path, bd=_text_exemple())
    o = cfg.configuracio({}, carpeta / PSX, paquet).opcions
    assert (o["epsg_utm"], o["font_epsg_utm"]) == (None, None)
    o = cfg.configuracio(_meta(carpeta), carpeta / PSX, paquet).opcions
    assert (o["epsg_utm"], o["font_epsg_utm"]) == (32718, str(carpeta / "jaciment.yaml"))


def test_crs_de_jaciment():
    assert cfg.crs_de_jaciment(JACIMENT_32718) == "EPSG:32718"
    assert cfg.crs_de_jaciment('nom: "Prova"\n') is None
    assert cfg.crs_de_jaciment("epsg: true\n") is None
    assert cfg.crs_de_jaciment(None) is None


# ── Claus obligatòries ───────────────────────────────────────────────────────


def test_sense_codi_regex(tmp_path):
    carpeta = _jaciment(tmp_path, bd="epsg_utm: 32718\n")
    o = cfg.configuracio(_meta(carpeta), None, _paquet(tmp_path)).opcions
    assert db_codis.validar_opcions(o) == [db_codis.ERROR_CODI_REGEX]
    assert db_codis.validar_opcions(o, requereix_utm=False) == [db_codis.ERROR_CODI_REGEX]
    assert "codi_regex" in db_codis.ERROR_CODI_REGEX and "bd.yaml" in db_codis.ERROR_CODI_REGEX


def test_sense_epsg(tmp_path):
    carpeta = _jaciment(tmp_path, bd=_text_exemple())          # cap jaciment.yaml, cap epsg_utm
    o = cfg.configuracio(_meta(carpeta), None, _paquet(tmp_path)).opcions
    assert (o["epsg_utm"], o["font_epsg_utm"]) == (None, None)
    assert db_codis.validar_opcions(o) == [db_codis.ERROR_EPSG_UTM]
    assert db_codis.validar_opcions(o, requereix_utm=False) == []   # mètriques: no projecta a UTM
    assert "bd.yaml" in db_codis.ERROR_EPSG_UTM and "jaciment.yaml" in db_codis.ERROR_EPSG_UTM


# ── Carpeta d'eixida per defecte ─────────────────────────────────────────────


def test_carpeta_eixida_amb_vincle(tmp_path):
    carpeta = _jaciment(tmp_path, bd=_text_exemple())
    assert cfg.carpeta_eixida(_meta(carpeta), tmp_path / "altra" / PSX) == carpeta / "BD"
    assert cfg.configuracio(_meta(carpeta), None, _paquet(tmp_path)).carpeta_eixida == carpeta / "BD"
    assert not (carpeta / "BD").exists()                       # llegir la configuració no crea res


def test_carpeta_eixida_sense_vincle_carpeta_del_psx(tmp_path):
    assert cfg.carpeta_eixida({}, tmp_path / PSX) == tmp_path
    assert cfg.carpeta_eixida(_meta(tmp_path / "no_hi_es"), str(tmp_path / PSX)) == tmp_path


def test_carpeta_eixida_document_no_desat():
    assert cfg.carpeta_eixida({}, None) == Path.home()
    assert cfg.carpeta_eixida({}, "") == Path.home()
