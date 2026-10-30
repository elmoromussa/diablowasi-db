"""Coordenades i orientació de façana per estructura (pipeline/db/georef.py)."""

from __future__ import annotations

import csv

from chachapoya_bd.db import georef as g
from chachapoya_bd.db.codis import Codis
from chachapoya_bd.db.csv_out import escriure_csv

# Nomenclatura de prova (necròpolis amb sectors, com Diablo Wasi).
OPTS_DW = {
    "codi_regex": r"^(?P<codi>(?P<sector>(LP|DW)-[A-Z0-9]+)-EA\d+)(?:-(?P<variant>.+))?$",
    "sectors": ["LP-G", "DW-G", "DW-S01", "DW-S02"],
    "epsg_utm": 32718,
}


class TestOrientacio:
    def test_nord_est_i_sectors(self):
        assert g.orientacio_facana((0, 1, 0))[:2] == (0.0, "N")
        az, sec, tilt, notes = g.orientacio_facana((1, 0, 0))
        assert (round(az), sec, round(tilt), notes) == (90, "E", 0, [])
        assert g.orientacio_facana((-1, -1, 0))[1] == "SW"
        assert round(g.orientacio_facana((-1, -1, 0))[0]) == 225

    def test_pla_horitzontal_es_nd(self):
        az, sec, tilt, notes = g.orientacio_facana((0, 0.1, 1))
        assert az is None and sec == "ND" and tilt > 80 and notes

    def test_tilt_avis(self):
        _az, sec, _tilt, notes = g.orientacio_facana((0, 1, 0.5), tilt_avis_deg=15)
        assert sec == "N" and "inclinat" in notes[0]


def _chunk(label, marcadors=(), enabled=True, georef=True, region=True):
    return g.ChunkInfo(label, enabled, georef, region,
                       [g.Marcador(lab, pos) for lab, pos in marcadors])


def _inventari(chunks, codis=None, marca_tot=True):
    codis = codis or Codis(OPTS_DW)
    orient, coord = g.candidats(chunks, codis)
    if not marca_tot:
        orient = [c for c in orient if g.preseleccio_orientacio(c)]
        coord = [c for c, _n in coord if c.enabled]
    else:
        coord = [c for c, _n in coord]
    return g.construir_inventari(orient, coord, codis)


class TestInventari:
    def test_candidats_i_preseleccio(self):
        chunks = [
            _chunk("DW-S01-EA01"),
            _chunk("DW-S01-EA01-2018"),
            _chunk("DW-S01", [("DW-S01-EA01", (1, 2, 3)), ("DW-S01-EA02", (4, 5, 6)), ("PALP", (0, 0, 0))]),
            _chunk("altres"),
        ]
        orient, coord = g.candidats(chunks, Codis(OPTS_DW))
        assert [(c.codi, c.variant) for c in orient] == [("DW-S01-EA01", None), ("DW-S01-EA01", "2018")]
        assert [(c.label, n) for c, n in coord] == [("DW-S01", 2)]
        assert [g.preseleccio_orientacio(c) for c in orient] == [True, False]

    def test_entrada_completa_i_nomes_coordenades(self):
        chunks = [_chunk("DW-S01-EA01"),
                  _chunk("DW-S01", [("DW-S01-EA01", (1, 2, 3)), ("DW-S01-EA02", (4, 5, 6))])]
        ent = _inventari(chunks)
        assert [e.codi for e in ent] == ["DW-S01-EA01", "DW-S01-EA02"]
        assert ent[0].estat == "completa" and ent[0].chunk_orientacio.label == "DW-S01-EA01"
        assert ent[1].estat == "coordenades" and "cap chunk d'orientació" in ent[1].o_errors[0]
        assert g.resum_estats(ent) == {"completa": 1, "orientacio": 0, "coordenades": 1,
                                       "bloquejada": 0, "exportables": 2}

    def test_variant_sola_avisa_i_dues_variants_bloquegen(self):
        ent = _inventari([_chunk("DW-S01-EA01-2018")])
        assert ent[0].estat == "orientacio" and "variant" in ent[0].avisos[0]
        ent = _inventari([_chunk("DW-S01-EA01-2018"), _chunk("DW-S01-EA01-2019")])
        assert ent[0].estat == "bloquejada" and "2 variants" in ent[0].o_errors[0]

    def test_errors_de_georef_region_i_marcadors_duplicats(self):
        ent = _inventari([_chunk("DW-S01-EA01", georef=False),
                          _chunk("DW-S01-EA02", region=False),
                          _chunk("S", [("DW-S01-EA02", (0, 0, 0)), ("DW-S01-EA02", (1, 1, 1))])])
        assert "sense georeferenciar" in ent[0].o_errors[0]
        assert "sense bounding box" in ent[1].o_errors[0]
        assert "2 marcadors" in ent[1].c_errors[0]

    def test_divergencia_entre_chunks(self):
        ent = _inventari([_chunk("A", [("DW-S01-EA01", (0, 0, 0))]),
                          _chunk("B", [("DW-S01-EA01", (0, 0, 0.5))])])
        assert ent[0].estat == "coordenades"
        assert ent[0].colocacions[0][0].label == "A"
        assert any("divergència de 0.50 m" in a for a in ent[0].avisos)

    def test_sector_fora_cataleg(self):
        ent = _inventari([_chunk("DW-S99-EA01")])
        assert "fora del catàleg" in ent[0].avisos[0]


class TestCSV:
    def test_camps_segons_epsg(self):
        assert g.camps_csv(32718)[3:5] == ("E_UTM32718", "N_UTM32718")
        assert g.camps_csv(25830)[3:5] == ("E_UTM25830", "N_UTM25830")

    def test_files_i_escriptura(self, tmp_path):
        chunks = [_chunk("DW-S01-EA01"),
                  _chunk("DW-S01", [("DW-S01-EA01", (1, 2, 3)), ("DW-S01-EA02", (4, 5, 6))])]
        ent = _inventari(chunks)
        files, n_c, n_o = g.files_csv(
            ent, 32718, "0.05",
            calc_orientacio=lambda ch: (123.4, "SE", 2.0, []),
            calc_coordenades=lambda ch, mk: (-6.12345678, -77.98765432, 200000.123, 9300000.456, 2500.005))
        assert (n_c, n_o) == (2, 1)
        f = files[0]
        assert f["Code"] == "DW-S01-EA01" and f["Azimuth_deg"] == "123" and f["Sector8"] == "SE"
        assert f["Lat_WGS84"] == "-6.1234568" and f["E_UTM32718"] == "200000.12"
        assert f["Alt_masl"] == "2500.01" and f["Precision_m"] == "0.05"
        assert f["Orientation_Chunk"] == "DW-S01-EA01" and f["Marker_Chunk"] == "DW-S01"
        assert files[1]["Azimuth_deg"] == "" and "SENSE ORIENTACIÓ" in files[1]["Notes"]

        p = escriure_csv(tmp_path / "sub" / "georef.csv", g.camps_csv(32718), files)
        raw = p.read_bytes()
        assert raw.startswith(b"\xef\xbb\xbf")   # BOM per a Access
        with p.open(encoding="utf-8-sig", newline="") as fh:
            rows = list(csv.DictReader(fh))
        assert len(rows) == 2 and rows[0]["N_UTM32718"] == "9300000.46"
