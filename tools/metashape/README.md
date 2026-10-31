# Scripts de Metashape

Scripts en Python per a Agisoft Metashape Professional 2.3.0. Són la meitat d'entrada de dos pipelines la meitat de sortida dels quals és en VBA: generen els CSV que consumeixen `tools/migrations/chachapoya_Import_Georef_v24.bas` i `tools/chachapoya_CheckMetrics.bas`. S'executen des de Metashape (Tools, Run Script) sobre el projecte fotogramètric; cap dels tres no modifica el projecte excepte `batch_orto_bbox.py`, que hi afig ortomosaics.

| Script | Què fa | Consumidor |
|---|---|---|
| `batch_orto_bbox.py` | Genera un ortomosaic planar per chunk segons el pla del *bounding box* triat (front, back, top, bottom, right, left). Fixa la disciplina del bbox que comparteixen els altres dos scripts: la vista es tria amb el diàleg, mai girant el bbox. | Documentació gràfica; calibratge del bbox. |
| `export_metriques_shapes.py` | Llig les *shapes* 3D dibuixades sobre el model, en valida l'etiqueta segons la gramàtica `PLA[-MÈTRICA]-[ELEMENT]-COS...` i escriu quatre CSV: estructures, murs, portals i auditoria per *shape* amb els controls de qualitat. | `CheckMetrics` (v25), que els creua amb `T_STRUCTURES` i pobla `T_METRIC_REVIEW`. |
| `export_georef_orientacions.py` | Exporta, una fila per estructura, les coordenades del punt de circulació (WGS84 i UTM 18S des del mateix punt geocèntric) i l'azimut de façana derivat de la normal del bbox. | `ImportGeoref` (v24), que ompli les coordenades i `Facade_Azimuth_Deg`. |

Ordre de treball: primer la passada de l'extractor de mètriques amb el corpus net (zero `PREFIX_DUBTOS`, que certifica que tots els bbox estan alineats amb la façana), després l'exportador de georeferenciació i, per últim, `ImportGeoref` a Access.

Els tres scripts comparteixen la constant `VIEW_AXIS` (per defecte `-Y`): la normal de façana i el pla de projecció se'n deriven i han de coincidir.

Una eixida de l'extractor de mètriques, la passada del 30 d'agost de 2026 que va alimentar la revisió v25, és a `data/metriques_v25/`. El CSV de georeferenciació no es publica perquè conté les coordenades de les estructures.

## Menú de Metashape (`chachapoya_bd`)

El paquet `chachapoya_bd/` afig a Metashape el menú *Extensions → Base de dades*, successor dels dos exportadors d'esta carpeta. Ve d'Oculat, on es va desenvolupar a partir d'ells, i s'hi ha traslladat sense redissenyar-lo; la descripció completa (estructura, `bd.yaml`, convencions i tests) és a `chachapoya_bd/README.md`.

**Instal·lació.** Metashape → Tools → Run Script → `tools\metashape\chachapoya_bd\install_menu.py` (o afegir-lo com a script d'inici). Apareixen tres entrades:

| Entrada | Fa el que feia | Eixida |
|---|---|---|
| Exportar georef i orientacions... | `export_georef_orientacions.py` | `georef_estructures_AAAAMMDD.csv` |
| Exportar mètriques de shapes... | `export_metriques_shapes.py` | `metriques_estructures.csv`, `metriques_murs.csv`, `metriques_portals.csv`, `metriques_shapes_audit.csv` |
| Auditar shapes del chunk actiu | — (eina de depuració nova) | Informe a la consola |

Si el menú d'Oculat també està instal·lat, té les mateixes etiquetes: cada instal·lador lleva les tres entrades abans d'afegir-les, així que mana l'últim que s'executa.

**Configuració.** La regex del codi d'estructura, el catàleg de sectors i els llindars dels controls són a `bd.yaml`. El de Diablo Wasi és versionat a `chachapoya_bd/bd.yaml` i és el que s'usa per defecte, amb qualsevol projecte. Si el projecte està vinculat a una carpeta de jaciment (claus `jaciment/*` del document, que escriuen Oculat o Pontarch) i a l'arrel d'eixa carpeta hi ha un `bd.yaml` (fet a partir de `chachapoya_bd/bd.yaml.example`), mana este. El vincle no és obligatori. L'EPSG UTM es pren del `bd.yaml` triat (`epsg_utm`) o, si no hi és i hi ha vincle, de `jaciment.yaml` (`epsg`). Els CSV s'escriuen per defecte a `<jaciment>/BD/` amb vincle o a la carpeta del `.psx` sense. Cada diàleg mostra quin `bd.yaml` usa; si no en troba cap o hi falta alguna clau obligatòria, no s'obri i diu què falta.

**Consumidors VBA.**

| Entrada | CSV | Consumidor |
|---|---|---|
| Exportar georef i orientacions... | `georef_estructures_*.csv` | `tools/migrations/chachapoya_Import_Georef_v24.bas` → `ImportGeoref` (ompli `T_STRUCTURES` per `Code`: coordenades, altitud, precisió, `Facade_Azimuth_Deg`; per defecte només camps NULL) |
| Exportar mètriques de shapes... | `metriques_estructures.csv`, `metriques_murs.csv`, `metriques_portals.csv`, `metriques_shapes_audit.csv` | `tools/chachapoya_CheckMetrics.bas` → `CheckMetrics "…\metriques_shapes_audit.csv"` (creua els quatre CSV amb `T_STRUCTURES` i pobla `T_METRIC_REVIEW`; res s'escriu sense signatura) i `ApplyMetricReview` |

`CheckMetrics` troba els tres fitxers germans substituint `shapes_audit` al nom del fitxer triat: els quatre CSV han d'estar a la mateixa carpeta i amb els noms que escriu l'eina (`chachapoya_bd/db/shapes.py`, `FITXERS_CSV`). L'ordre de treball és el mateix que amb els scripts: mètriques amb el corpus net, georeferenciació i `ImportGeoref`.

**Relació amb els scripts d'esta carpeta.** Els tres scripts es conserven tal com es van usar per a les passades de v24 i v25. Les dues eines d'exportació del menú en són la continuació, amb la lògica separada en mòduls purs i provats (`chachapoya_bd/db/`) i la configuració en `bd.yaml` en lloc de constants al codi. Les columnes de `georef_estructures_*.csv` són les mateixes i en el mateix ordre que les de `export_georef_orientacions.py` (els noms d'`E_UTM`/`N_UTM` porten l'EPSG configurat). `batch_orto_bbox.py` no té successor al menú: els ortomosaics els genera Pontarch.

**Convenció de vistes.** Els scripts nomenen els plans pel signe d'un eix local del bbox (`PLANE_AXIS = {"f": "-Y", "b": "Y", "t": "Z", "bo": "-Z", "r": "X", "l": "-X"}`, `VIEW_AXIS = "-Y"`) i el menú pel nom de la vista de Pontarch (`PLA_VISTA = {"f": "frontal", ...}` a `db/shapes.py`, amb `frontal` = eix `Y` cap a l'observador a `db/vistes.py`, i `eix_facana: frontal`). Llegint el codi, la geometria és la mateixa: `ortho_axes` dels scripts i `eixos_vista` de `db/vistes.py` retornen els mateixos tres eixos (X i Y de la imatge, Z de la normal) per a `Z`, `-Z`, `X` i `-X`, i tenen intercanviades les dues entrades de l'eix Y: `ortho_axes(rot, "-Y")` = `eixos_vista(rx, ry, rz, "Y")` = (−rx, rz, ry) i `ortho_axes(rot, "Y")` = `eixos_vista(rx, ry, rz, "-Y")` = (rx, rz, −ry). Per tant, per a cada pla (`f` amb `-Y` als scripts i amb `frontal` al menú, i així els altres cinc) els dos codis donen la mateixa base de projecció i la mateixa normal de façana (+ry per a `f`), que és la que fan servir les mètriques i l'azimut. La diferència és només de nom, i només en l'eix Y: als scripts, l'etiqueta de ±Y diu la direcció de la mirada (`-Y`: l'observador és a +Y i mira cap a −Y), mentre que per a X i Z, i a `db/vistes.py` per a les sis vistes, l'etiqueta és l'eix cap a l'observador. El comentari de `batch_orto_bbox.py` que diu que «l'eix apunta cap a l'observador» és, doncs, exacte per a X i Z però no per a Y. Esta equivalència és de lectura; no s'ha comparat cap eixida a Metashape.
