"""Gramàtica d'etiquetes, mètriques i QC de shapes (pipeline/db/shapes.py)."""

from __future__ import annotations

import csv

import numpy as np
import pytest

from chachapoya_bd.db import vistes as om
from chachapoya_bd.db import shapes as s

# ── Gramàtica ────────────────────────────────────────────────────────────────


class TestGramatica:
    @pytest.mark.parametrize("label, esperat", [
        ("f-1.1", dict(plane="f", body="1.1", element=None, metric=None)),
        ("F-0.1-C", dict(plane="f", body="0.1", complete=True, canonical="f-0.1-c")),
        ("f-port-1.1a", dict(element="port", body="1.1", letter="a")),
        ("f-ll-1.1-res", dict(metric="ll", restitution=True)),
        ("f-ll-1.1+1.2-c", dict(also=[("f", "1.2")], complete=True)),
        ("f-ll-1.1+b-2.1-r", dict(also=[("b", "2.1")], restitution=True)),
        ("bo-roof-1.1", dict(plane="bo", element="roof")),
        ("f-corn-1.1-1.2", dict(element="corn", body="1.1", body2="1.2")),
        ("t-int-1.1", dict(plane="t", element="int")),
    ])
    def test_etiquetes_valides(self, label, esperat):
        d, err, _sug = s.parse_etiqueta(label)
        assert err is None, err
        for k, v in esperat.items():
            assert d[k] == v, k

    @pytest.mark.parametrize("label, motiu", [
        ("xyz", "fora de gramatica"),
        ("f", "cal element o cos"),
        ("f-1.1-c-", "fora de gramatica"),
        ("f-port-1.1-1.2", "parell de cossos nomes"),
        ("f-1.1+1.2", "nomes val per a polilinies"),
        ("f-ll-1.1+1.2", "requereix estatus"),
        ("f-ll-1.1+1.2-1.3-c", "no es combinable"),
        ("f-1.0", "cos 1.0: Y ha de ser >= 1"),
        ("f-corn-1.1-2.0", "cos 2.0"),
        ("f-ll-1.1+1.00-c", "cos 1.00"),
    ])
    def test_etiquetes_invalides(self, label, motiu):
        d, err, _sug = s.parse_etiqueta(label)
        assert d is None and motiu in err

    @pytest.mark.parametrize("vell, nou", [
        ("e-1.1", "f-1.1"),
        ("p-int-1.1", "t-int-1.1"),
        ("port-1.1", "f-port-1.1"),
        ("hh-1.1-c", "f-hh-1.1-c"),
        ("h-1.1-c", "f-hh-1.1-c"),
        ("f-ll-1.1-res", "f-ll-1.1-r"),
        ("f-11", "f-1.1"),
        ("f-n1.1", "f-1.1"),
        ("f-rooftop-1.1", "f-roof-1.1"),
        ("b-roof-1.1", "bo-roof-1.1"),
        ("f-ll-1.1-c+1.2", "f-ll-1.1+1.2-c"),
    ])
    def test_suggeriments(self, vell, nou):
        assert s.suggerir(vell) == nou

    def test_sense_suggeriment(self):
        assert s.suggerir("f-1.1") is None
        assert s.suggerir("blabla") is None


# ── Geometria ────────────────────────────────────────────────────────────────


def test_newell_i_shoelace():
    rect = np.array([[0, 0, 0], [2, 0, 0], [2, 0, 3], [0, 0, 3]], dtype=float)
    n = s.newell(rect)
    assert np.isclose(np.linalg.norm(n), 6.0)
    assert np.allclose(np.abs(n / np.linalg.norm(n)), [0, 1, 0])
    assert s.shoelace([(0, 0), (2, 0), (2, 3), (0, 3)]) == 6.0
    assert s.shoelace([(0, 0), (1, 1)]) == 0.0
    assert s.mediana([]) is None and s.mediana([3, 1, 2]) == 2 and s.mediana([1, 2, 3, 4]) == 2.5


# ── Anàlisi d'una estructura sintètica ───────────────────────────────────────

RX, RY, RZ = np.eye(3)


def _marc(rot=np.eye(3)) -> s.Marc:
    rx, ry, rz = rot[:, 0], rot[:, 1], rot[:, 2]
    eixos = {pla: tuple(np.asarray(a, dtype=float) for a in om.eixos_per_vista(rx, ry, rz, vista))
             for pla, vista in s.PLA_VISTA.items()}
    return s.Marc(eixos=eixos, centre=np.zeros(3), amunt=np.array([0, 0, 1.0]))


def _poly(label, pts):
    return s.Shape(label, "polygon", np.array(pts, dtype=float))


def _line(label, pts):
    return s.Shape(label, "line", np.array(pts, dtype=float))


def _estructura():
    """Façana f al pla y=0 (normal +Y), vertical +Z."""
    return [
        _poly("f-1.1", [[0, 0, 0], [2, 0, 0], [2, 0, 3], [0, 0, 3]]),          # mur cons 6 m²
        _line("f-ll-1.1-c", [[0, 0, 0], [2.5, 0, 0]]),                          # ll 2.5 (-c)
        _line("f-hh-1.1-c", [[0, 0, 0], [0, 0, 4]]),                            # hh 4
        _line("f-dd-1.1-c", [[0, 0, 0], [0, 0.6, 0]]),                          # dd 0.6 (gruix)
        _poly("f-port-1.1", [[0.5, 0, 0], [1.5, 0, 0], [1.5, 0, 2], [0.5, 0, 2]]),  # portal 2 m²
        _poly("t-int-1.1", [[0, 0, 3], [2, 0, 3], [2, 1, 3], [0, 1, 3]]),       # planta int 2 m²
    ]


class TestAnalisi:
    def test_mur_portal_i_totals(self):
        r = s.analitzar("DW-S01-EA01", "DW-S01-EA01", _estructura(), {"intern": _marc()})
        assert r.mode == "intern" and r.n_shapes == 6 and r.n_valid == 6 and r.n_qc == 0
        assert np.isclose(r.vert_cons, 6.0) and np.isclose(r.f_cons, 6.0) and np.isclose(r.t_int, 2.0)

        mur = r.murs[0]
        assert (mur["Wall"], mur["Body"], mur["N_Fragments"]) == ("f", "1.1", "1")
        assert mur["Area_Cons_m2"] == "6.000" and mur["Width_Cons_m"] == "2.000" and mur["Height_Cons_m"] == "3.000"
        assert (mur["LL_m"], mur["LL_Source"], mur["HH_m"], mur["DD_m"]) == ("2.500", "-c", "4.000", "0.600")
        assert mur["Area_Orig_m2"] == "10.000" and mur["Orig_Source"] == "rectangle(ll:-c,hh:-c)"
        assert mur["Volume_Prism_m3"] == "6.000" and mur["Index_Conserv_pct"] == "60.0"

        port = r.ports[0]
        assert port["Port"] == "1.1" and port["Area_Cons_m2"] == "2.000"
        assert port["Ratio_vs_Body_Cons_pct"] == "33.3"
        assert port["Area_Orig_m2"] == "" and port["Ratio_vs_Body_Orig_pct"] == ""   # només per a -c / -r

    def test_portal_complet_te_ratio_orig(self):
        shapes = _estructura() + [_poly("f-port-1.1-c", [[0.5, 0, 0], [1.5, 0, 0], [1.5, 0, 2], [0.5, 0, 2]])]
        r = s.analitzar("C", "C", shapes, {"intern": _marc()})
        port = r.ports[0]
        assert port["Status"] == "-c" and port["Area_Orig_m2"] == "2.000"
        assert port["Ratio_vs_Body_Orig_pct"] == "20.0"      # 2 / (2.5 × 4)

        fila = next(x for x in r.shapes if x["Label"] == "f-1.1")
        assert fila["Area_Proj_m2"] == "6.000" and fila["Obliquity_deg"] == "0.0"
        assert fila["Status"] == "cons" and fila["Geometry"] == "polygon"
        hh = next(x for x in r.shapes if x["Label"] == "f-hh-1.1-c")
        assert hh["Len_Vert_m"] == "4.000" and hh["Metric_Value_m"] == "4.000"

    def test_invalides_2d_i_suggeriment(self):
        shapes = [_poly("xyz", [[0, 0, 0], [1, 0, 0], [1, 0, 1]]),
                  _poly("e-1.1", [[0, 0, 0], [1, 0, 0], [1, 0, 1]]),
                  s.Shape("f-1.2", "polygon", None, 4)]
        r = s.analitzar("C", "C", shapes, {"intern": _marc()})
        assert (r.n_invalid, r.n_2d, r.n_valid) == (2, 1, 0)
        inv = {x["Label"]: x for x in r.shapes}
        assert inv["xyz"]["QC_Flags"].startswith("INVALIDA:")
        assert inv["e-1.1"]["Suggestion"] == "f-1.1"
        assert inv["f-1.2"]["QC_Flags"] == "2D_LLEGAT" and inv["f-1.2"]["Dims"] == "2D"

    def test_qc_direccio_prefix_i_conflicte(self):
        shapes = [
            _poly("f-1.1", [[0, 0, 0], [2, 0, 0], [2, 0, 3], [0, 0, 3]]),
            _line("f-hh-1.1-c", [[0, 0, 0], [2, 0, 0]]),        # "vertical" horitzontal → DIRECCIO
            _line("f-hh-1.1-r", [[0, 0, 0], [0, 0, 3]]),        # -r i -c alhora → CONFLICTE
            _poly("r-1.1", [[0, 0, 0], [2, 0, 0], [2, 0, 3], [0, 0, 3]]),  # normal +Y però prefix r → PREFIX_DUBTOS
            _line("f-dd-1.1-c", [[0, 0, 0], [0, 0.6, 0], [0, 0.6, 0.5]]),   # 3 punts → DD_MULTITRAM
            _line("t-wall-1.1", [[0, 0, 3], [1, 0, 3]]),        # línia de mur amb prefix t → ORFE_MUR
            _poly("f-port-2", [[0, 0, 0], [1, 0, 0], [1, 0, 1]]),
        ]
        r = s.analitzar("C", "C", shapes, {"intern": _marc()})
        flags = {x["Label"]: x["QC_Flags"] for x in r.shapes}
        assert flags["f-hh-1.1-c"].startswith("DIRECCIO:90")
        assert flags["r-1.1"] == "PREFIX_DUBTOS:f"
        assert "DD_MULTITRAM" in flags["f-dd-1.1-c"]
        assert "ORFE_MUR" in flags["t-wall-1.1"] and "LINIA_SENSE_TOKEN" in flags["t-wall-1.1"]
        assert flags["f-port-2"] == "INVALIDA:fora de gramatica"
        mur = next(m for m in r.murs if m["Body"] == "1.1")
        assert "CONFLICTE_hh_RES_C" in mur["QC_Notes"] and mur["HH_Source"] == "-c"

    def test_propagacio_i_discrepancia(self):
        shapes = [
            _line("f-ll-1.1+1.2-c", [[0, 0, 0], [2.5, 0, 0]]),
            _line("f-ll-1.2-r", [[0, 0, 0], [3.0, 0, 0]]),      # directa 3.0 vs propagada 2.5 → >2 %
        ]
        r = s.analitzar("C", "C", shapes, {"intern": _marc()})
        m12 = next(m for m in r.murs if m["Body"] == "1.2")
        assert m12["LL_m"] == "3.000" and m12["LL_Source"] == "-r"
        assert "DISCREPANCIA_ll_prop:f-1.1" in m12["QC_Notes"]
        shapes = [_line("f-ll-1.1+1.2-c", [[0, 0, 0], [2.5, 0, 0]])]
        r = s.analitzar("C", "C", shapes, {"intern": _marc()})
        m12 = next(m for m in r.murs if m["Body"] == "1.2")
        assert m12["LL_Source"] == "prop:f-1.1-c"

    def test_triar_marc_pren_el_menys_oblic(self):
        rot90 = np.array([[0, -1, 0], [1, 0, 0], [0, 0, 1]], dtype=float)   # gira 90° sobre Z
        marcs = {"mon-local": _marc(rot90), "intern": _marc()}
        assert s.triar_marc(_estructura(), marcs) == "intern"
        r = s.analitzar("C", "C", _estructura(), marcs)
        assert r.mode == "intern"
        assert s.triar_marc([_line("f-ll-1.1-c", [[0, 0, 0], [1, 0, 0]])], marcs) == "mon-local"

    def test_sense_shapes(self):
        r = s.analitzar("C", "C", [], {"intern": _marc()})
        assert r.n_shapes == 0 and "sense shapes" in r.notes


def test_exportar_4_csv(tmp_path):
    r = s.analitzar("DW-S01-EA01", "DW-S01-EA01", _estructura(), {"intern": _marc()})
    out = s.exportar([r], tmp_path / "csv")
    assert set(out) == {"estructures", "murs", "portals", "shapes"}
    assert out["estructures"][1] == 1 and out["murs"][1] == 1 and out["portals"][1] == 1 and out["shapes"][1] == 6
    for clau, (nom, camps) in s.FITXERS_CSV.items():
        with open(out[clau][0], encoding="utf-8-sig", newline="") as fh:
            rd = csv.DictReader(fh)
            assert tuple(rd.fieldnames) == camps, clau
            rows = list(rd)
        assert len(rows) == out[clau][1]
    est = list(csv.DictReader(open(out["estructures"][0], encoding="utf-8-sig")))[0]
    assert est["Frame_Mode"] == "intern" and est["Vert_Cons_m2"] == "6.000" and est["N_Valid"] == "6"
