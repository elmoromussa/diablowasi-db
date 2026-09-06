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
