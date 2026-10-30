# Origen dels fitxers vendoritzats

Codi copiat d'Oculat (pipeline Metashape → 3DHOP del mateix autor) al commit `52863de` perquè `chachapoya_bd` no depenga en temps d'execució ni d'Oculat ni de Pontarch. Els fitxers no s'editen ací: si l'origen canvia, es tornen a copiar i s'actualitza esta taula.

El sha256 és el del fitxer tal com està al repositori, amb finals de línia LF (els de l'origen també ho són al seu repositori). `tests/test_vendor.py` el comprova; si existeix la variable d'entorn `OCULAT_REPO` (arrel d'una còpia d'Oculat), també comprova que els tres fitxers de contracte són idèntics als d'Oculat.

| Fitxer | Origen (Oculat) | Commit | sha256 | Estat |
|---|---|---|---|---|
| `contracte/claus.py` | `oculat/pipeline/contracte/claus.py` | `52863de` | `3d65003d4e4f6c712e5bd4b992c095cc1fa41b066af03d46df781ccbcb2d4c8b` | idèntic |
| `contracte/utm.py` | `oculat/pipeline/contracte/utm.py` | `52863de` | `b63c1825a70023d2abcf2e0290b90845775815db45c9281e448328b8454d3e07` | idèntic |
| `contracte/jaciment_yaml.py` | `oculat/pipeline/contracte/jaciment_yaml.py` | `52863de` | `3a786c9493c2c953c229407a9cafd1d2063799f4d196e7f805a06394f761f5d8` | idèntic |
| `qt_comu.py` | `oculat/metashape_addon/qt_comu.py` | `52863de` | `2467dddb5e7d1f339d8a5ae761e45c3400ef01751aee9e8863d1e994d0152896` | retallat (funcions: `InformeDialog`, `executar`, `informe`, `ConfirmDialog`, `confirmar`, `refrescar_gui`, `te_numpy`, `_ordre_pip`, `instal_lar_paquet`, `requereix_numpy`) |
| `contracte/__init__.py` | — | — | `2bb02945d31b2ef1eef26504f41617249f466ff26cfd8c27718f6b4f8e0e3de6` | nou (mínim: només docstring) |
| `__init__.py` | — | — | `2510de1b6e48b79150bd9a430f0eabc33c535391b90866170b61bb76288516ce` | nou (mínim: només docstring) |

## `qt_comu.py`: què s'ha llevat

L'origen té sha256 `fa2e49aeab822542f93b9fd06d1495afb931690ac86ceeafbef9be6fefaaf2b1`. Respecte d'ell, la còpia retallada:

- substitueix la docstring del mòdul per una que diu què és i què conté;
- lleva el bloc `_ROOT` / `sys.path` (i els imports `sys` i `Path` que només servien per a ell): el paquet es troba pel `sys.path` que posen les eines;
- lleva `ProgresConsola`, que cap de les tres eines no usa;
- lleva la reexportació final de `metashape_addon.qt_contracte` (`SelectorCrsUtm`, `BandaJaciment`, `EstatBuit`, `FormulariVincle`): són peces del vincle al jaciment, que no és cosa d'este menú.

La resta del codi és idèntica, línia per línia, a l'origen.
