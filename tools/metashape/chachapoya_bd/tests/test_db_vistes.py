"""Vistes del region (contracte amb els ortomosaics de Pontarch): eixos de cada vista."""

from __future__ import annotations

import numpy as np
import pytest

from chachapoya_bd.db import vistes as om

RX, RY, RZ = np.eye(3)


def _dextrogira(ax, ay, az) -> bool:
    return np.allclose(np.cross(ax, ay), az)


class TestEixos:
    def test_totes_les_vistes_son_bases_dextrogires(self):
        for v in om.VISTES:
            ax, ay, az = om.eixos_per_vista(RX, RY, RZ, v.id)
            assert _dextrogira(ax, ay, az), v.id

    def test_frontal_mira_des_de_la_facana_amb_vertical_amunt(self):
        # façana = +Y del region (convenció ea_front); vertical = +Z
        ax, ay, az = om.eixos_per_vista(RX, RY, RZ, "frontal")
        assert np.allclose(az, RY) and np.allclose(ay, RZ)

    def test_superior_es_la_planta_i_les_oposades_no_son_mirall(self):
        _, ay, az = om.eixos_per_vista(RX, RY, RZ, "superior")
        assert np.allclose(az, RZ) and np.allclose(ay, RY)
        for a, b in (("frontal", "posterior"), ("superior", "inferior"), ("dreta", "esquerra")):
            za = om.eixos_per_vista(RX, RY, RZ, a)[2]
            zb = om.eixos_per_vista(RX, RY, RZ, b)[2]
            assert np.allclose(za, -zb)
            assert _dextrogira(*om.eixos_per_vista(RX, RY, RZ, b))

    def test_eix_desconegut(self):
        with pytest.raises(ValueError):
            om.eixos_vista(RX, RY, RZ, "W")
