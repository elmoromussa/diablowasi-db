# `chachapoya_bd`: menú *Base de dades* de Metashape

Paquet de Python per a Agisoft Metashape Professional 2.3.0 que afig el menú *Extensions → Base de dades* amb tres eines de només lectura sobre el projecte fotogramètric:

| Entrada del menú | Mòdul | Eixida |
|---|---|---|
| Exportar georef i orientacions... | `exportar_georef_bd.py` | `georef_estructures_AAAAMMDD.csv`: una fila per estructura amb les coordenades del punt de circulació (WGS84 i UTM des del mateix punt geocèntric) i l'azimut de façana. |
| Exportar mètriques de shapes... | `exportar_metriques_bd.py` | `metriques_estructures.csv`, `metriques_murs.csv`, `metriques_portals.csv` i `metriques_shapes_audit.csv`, segons la gramàtica de `docs/CONVENCIO_SHAPES_METRIQUES.md`. |
| Auditar shapes del chunk actiu | `auditar_shapes.py` | Informe a la consola de Metashape (geometria crua de cada *shape* i ortomosaics del chunk). |

El menú es va desenvolupar dins d'Oculat (pipeline Metashape → 3DHOP del mateix autor) i s'ha traslladat ací, al commit `52863de` d'Oculat, **sense redissenyar-lo**: mateixes entrades, mateixos diàlegs, mateixa lògica i mateixos CSV. Només canvien els noms dels paquets, d'on ve la configuració (`bd.yaml`, vegeu més avall) i la carpeta d'eixida per defecte (`<jaciment>/BD`). No importa res d'Oculat ni de Pontarch en temps d'execució: el que en necessita és a `_vendor/`.

Instal·lació, consumidors VBA dels CSV i relació amb els scripts històrics: vegeu `../README.md`.

## Estructura

| Ruta | Contingut |
|---|---|
| `install_menu.py` | Registra les tres entrades. S'executa des de Tools → Run Script. |
| `exportar_georef_bd.py`, `exportar_metriques_bd.py`, `auditar_shapes.py` | Diàlegs i adaptadors de l'API de Metashape. |
| `configuracio.py` | Lectura de `<jaciment>/bd.yaml`, precedència de l'EPSG UTM i carpeta d'eixida. Lògica pura. |
| `db/` | Lògica pura, sense `Metashape`: gramàtica del codi d'estructura (`codis.py`), inventari i files de georeferenciació (`georef.py`), gramàtica d'etiquetes, mètriques i controls de qualitat de les *shapes* (`shapes.py`), vistes del *region* (`vistes.py`) i escriptura dels CSV en UTF-8 amb BOM (`csv_out.py`). |
| `_vendor/` | Contracte del jaciment (`claus.py`, `utm.py`, `jaciment_yaml.py`, còpies byte a byte) i peces Qt (`qt_comu.py`, còpia retallada). Procedència i sha256 a `_vendor/ORIGEN.md`. |
| `tests/` | Tests de pytest de la lògica pura, de la configuració, dels fitxers vendoritzats i dels imports. |
| `bd.yaml.example` | Plantilla de `bd.yaml`. |

`db/vistes.py` és contracte amb Pontarch: les vistes `frontal`, `posterior`, `superior`, `inferior`, `dreta` i `esquerra` són les dels seus ortomosaics, i els plans `f b t bo r l` de les *shapes* hi corresponen. És còpia byte a byte de la d'Oculat; ids, etiquetes, eixos i la funció `eixos_vista` coincideixen amb `pontarch/comu/ortomosaics.py`.

## Configuració: `bd.yaml`

Les opcions del menú viuen en un fitxer propi de la carpeta de jaciment, `<jaciment>/bd.yaml`, al costat de `jaciment.yaml` però separat d'ell: `jaciment.yaml` és el contracte entre Oculat i Pontarch i este menú no el modifica. La carpeta de jaciment és la del vincle del projecte (claus `jaciment/carpeta` i `jaciment/campanya` del document, o la clau heretada `pontarch/lliurament`); el vincle el fa Oculat o Pontarch, no este paquet.

Les claus són les del bloc `database:` de `pipeline.yaml` d'Oculat, amb els mateixos noms, al primer nivell:

| Clau | Obligatòria | Valor per defecte | Ús |
|---|---|---|---|
| `codi_regex` | sí | — | Regex del codi d'estructura (`Code` de `T_STRUCTURES`), amb el grup `codi` i, opcionals, `variant` i `sector`. |
| `sectors` | no | `[]` | Catàleg de sectors (`L_SECTORS`); buit = no es comprova. |
| `epsg_geo` | no | `4326` | Coordenades geogràfiques del CSV. |
| `epsg_utm` | per a la georeferenciació | — | EPSG UTM de sortida. Si s'omet, es pren l'`epsg` de `jaciment.yaml`. |
| `precisio_m` | no | `0.05` | `Coord_Precision_m` per defecte. |
| `eix_facana` | no | `frontal` | Vista que fa de façana. |
| `tilt_avis_deg` | no | `15` | Inclinació del pla a partir de la qual s'avisa. |
| `min_horitz_deg` | no | `20` | Angle mínim amb l'horitzontal; per davall, azimut `ND`. |
| `divergencia_avis_m` | no | `0.25` | Distància a partir de la qual un marcador present en dos chunks dona avís. |

El fitxer es llig amb `jaciment_yaml.escalars` del contracte, sense PyYAML. Les regex s'escriuen entre cometes simples, que conserven les barres inverses tal qual. `bd.yaml.example` porta un exemple genèric actiu i, comentat, el de la necròpolis de Diablo Wasi i La Petaca (`LP|DW`, catàleg de sectors, `epsg_utm: 32718`). Quan es descomenta, substitueix les claus de dalt perquè mana l'última aparició.

**EPSG UTM**: `epsg_utm` de `bd.yaml` > `epsg` de `jaciment.yaml`. El diàleg de georeferenciació mostra d'on ve (`bd.yaml` o `jaciment.yaml`).

**Carpeta d'eixida per defecte**: `<jaciment>/BD/`, que substitueix la `paths.db_csv` d'Oculat. El selector de fitxer o carpeta dels diàlegs permet triar-ne una altra. Els CSV no es versionen en este repositori.

**Errors**: sense vincle, amb la carpeta del vincle inexistent, sense `bd.yaml` o sense alguna clau obligatòria, el diàleg no s'obri i un quadre d'informe diu què falta i on. El codi no té cap valor de reserva per a les claus obligatòries.

## Convencions

**Idioma del codi.** És l'excepció a la convenció del repositori (codi en anglés): els identificadors, comentaris i textos del codi traslladat es queden en català tal com eren a Oculat, perquè reanomenar-los seria redissenyar-lo. Les columnes dels CSV sí que són en anglés, i no canvien.

**Rutes d'Oculat a les docstrings.** El codi traslladat no s'ha retocat més enllà dels imports i de la configuració, i algunes docstrings citen rutes del repositori d'origen. `docs/CONVENCIO_SHAPES_METRIQUES.md` (§9, «Flux de treball») cita l'script històric `export_metriques_shapes.py`; l'entrada *Exportar mètriques de shapes...* n'és el successor. La correspondència és:

| Citat com a | Al repositori |
|---|---|
| `pipeline/db/<mòdul>.py` | `tools/metashape/chachapoya_bd/db/<mòdul>.py` |
| `metashape_addon/exportar_*_bd.py`, `metashape_addon/auditar_shapes.py` | `tools/metashape/chachapoya_bd/` |
| `access/chachapoya_Import_Georef_v24.bas` | `tools/migrations/chachapoya_Import_Georef_v24.bas` |
| `pipeline.yaml`, secció `database:` | `<jaciment>/bd.yaml` |
| `paths.db_csv` | `<jaciment>/BD/` |
| `CONVENCIO_SHAPES_METRIQUES.md` | `docs/CONVENCIO_SHAPES_METRIQUES.md` |
| `georef/ea_front.py` (a `db/vistes.py`) | Oculat, sense equivalent ací: referència a la convenció +Y = façana |

**Paquet amb nom propi.** Es diu `chachapoya_bd` i no `pipeline` ni `metashape_addon` perquè Oculat i este paquet poden conviure al mateix intèrpret de Metashape: amb un nom repetit, `sys.modules` serviria en silenci el codi de l'altre.

## Tests

Des de `tools/metashape/`, amb pytest i numpy:

```
python -m pytest
```

`test_db_*.py` són els tests d'Oculat de la lògica pura, amb els mateixos noms. `test_configuracio.py` cobreix `bd.yaml`, `test_vendor.py` els sha256 de `_vendor/ORIGEN.md` (i, si existeix la variable d'entorn `OCULAT_REPO` amb l'arrel d'una còpia d'Oculat, que el contracte hi és idèntic) i `test_imports.py`, per anàlisi sintàctica, que res no importa Oculat ni Pontarch i que `db/`, `configuracio.py` i el contracte no importen Metashape ni Qt. Els diàlegs només es poden provar dins de Metashape.
