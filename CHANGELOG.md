# Historial de versions

## Publicacions

| Versió | Data | Contingut |
|---|---|---|
| 1.2.0 | 2026-09-10 | Taules de corpus i riquesa per tipologia per al capítol 4 de la memòria (`analysis/scripts/05_taules_cap4.R`, `resultats_pas3.txt` i les taules corresponents). Cap canvi d'esquema ni de dades. |
| 1.1.0 | 2026-09-07 | Paquet d'anàlisi del TFM (`analysis/`): pipeline reproduïble d'extracció, matriu A-X i quatre scripts d'R (clustering UPGMA, co-ocurrència, graf de connexions, orientacions, riquesa, MNI, afectació, decoració, H02), amb dades derivades, registres, taules i les onze figures de la memòria. Cap canvi d'esquema ni de la base. |
| 1.0.1 | 2026-09-06 | Llicència dual (codi MIT; documentació i dades CC BY-NC 4.0), crèdits del projecte PALP, política de dades de posició, referències bibliogràfiques de les set datacions a `T_DATING`. Cap canvi d'esquema ni de dades d'anàlisi. |
| 1.0.0 | 2026-09-06 | Publicació del repositori: codi i documentació de la iteració v26, base Access v26 amb les dades, scripts de Metashape, metadades de citació. |

## Iteracions de treball

Cada fila és una etiqueta de git. La columna Delta enllaça l'especificació de canvis que va guiar la iteració; quan no n'hi ha, els canvis estan consolidats a la nota inicial de `docs/esquema.md`.

| Etiqueta | Data | Resum | Delta |
|---|---|---|---|
| v26 | 2026-09-06 | Versió definitiva: Jamb_Fabric_Reveal retirat (derivable de la parella de muntants), regla 3 i R55 reformulades, reformulació documental base/superestructura (N0/N1), estatut documental de Support_Modified, acta de 29 escriptures de dades. Pedaç de dades i migrador a base neta. | [DELTA_v25g_v26](docs/deltas/DELTA_v25g_v26.md) |
| v25g | 2026-09-03 | Tancament de worklist i bateria (566 NULLs, 34 decisions de regla), EA78 i EA79, R54 i R66 esmenades. Base tancada per a l'anàlisi: 106 EA, 91 arestes, 12 paternitats. | [DELTA_v25f_v25g](docs/deltas/DELTA_v25f_v25g.md) |
| v25f | 2026-08-31 | Només dades: primera població sistemàtica de T_CONNECTIONS (83 arestes), 9 assignacions d'ID_Parent, diedre EA76 reclassificat. | [DELTA_v25e_v25f](docs/deltas/DELTA_v25e_v25f.md) |
| v25e | 2026-08-31 | Domini de muntants amb segon membre no observable (Slab-NObs, Composite-NObs, Fabric-NObs); R69 esmenada. | [DELTA_v25d_v25e](docs/deltas/DELTA_v25d_v25e.md) |
| v25d | 2026-08-31 | Revisió de v25 (DB, formulari, documentació). | [DELTA_v24_v25](docs/deltas/DELTA_v24_v25.md) |
| v25c | 2026-08-31 | Revisió de v25 (DB, formulari). | [DELTA_v24_v25](docs/deltas/DELTA_v24_v25.md) |
| v25 | 2026-08-31 | Pipeline de coherència mètriques-BD: T_METRIC_REVIEW, mòdul CheckMetrics, consultes QRY_30 a 32. Camps Jamb_Fabric i Niche_Partition. | [DELTA_v24_v25](docs/deltas/DELTA_v24_v25.md) |
| v24 | 2026-08-30 | Facade_Azimuth_Deg, escrit només per l'importador de georeferenciació; regla 68. | [DELTA_v23_v24](docs/deltas/DELTA_v23_v24.md) |
| v23 | 2026-08-16 | Finestra del brancal (regla 55), murs de la cambra (regles 65 i 66), EA_Number emmagatzemat (regla 67), L_SUBSECTORS. | [DELTA_v22_v23](docs/deltas/DELTA_v22_v23.md) |
| v22a | 2026-08-15 | Formulari interí. | |
| v22 | 2026-08-15 | 2.Arq oberta als contextos naturals, Metrics_Available (regla 63), datació només per C14 (regla 64). | [DELTA_v21_v22](docs/deltas/DELTA_v21_v22.md) |
| v21 | 2026-08-15 | Criteri de la terrassa amb Z reservat al sistema volat (regla 62), Masonry_Present (regla 61), errates v20. | [DELTA_v20_v21](docs/deltas/DELTA_v20_v21.md) |
| v20 | 2026-08-15 | Quinze blocs tancats i auditoria completa del corpus (86 estructures). Primera Worklist. | [DELTA_v19_v20](docs/deltas/DELTA_v19_v20.md) |
| v19a | 2026-08-14 | Formulari interí. | |
| v19 | 2026-08-13 | Candidats pendents del delta anterior i troballes de la inspecció de la còpia amb dades. | |
| v18 | 2026-08-11 | Primera campanya real d'entrada de dades (35 estructures) i les seues notes. | [DELTA_v17_v18](docs/deltas/DELTA_v17_v18.md) |
| v17a | 2026-08-08 | Pedaç sobre v17 aplicat a la còpia amb dades. | [DELTA_v16_v17](docs/deltas/DELTA_v16_v17.md) |
| v17 | 2026-08-07 | Onze punts de disseny tancats a partir de la inspecció directa de la còpia amb dades. Primer manual d'ús. | [DELTA_v16_v17](docs/deltas/DELTA_v16_v17.md) |
| v16a | 2026-08-07 | Correcció: L_LOST_EVIDENCE.Name_VAL, el desplegable d'evidència sortia buit. Pedaç per a bases amb dades. | [DELTA_v15_v16](docs/deltas/DELTA_v15_v16.md) |
| v16 | 2026-08-07 | Revisió de la interfície amb les fitxes de camp a la mà. | [DELTA_v15_v16](docs/deltas/DELTA_v15_v16.md) |
| v15 | 2026-08-06 | Esquema i metodologia actualitzats. Exportació i importació de dades v11 a v15. | [DELTA_v12_v13](docs/deltas/DELTA_v12_v13.md) |
| v14 | 2026-08-06 | Etiquetes valencianes a totes les taules lookup. | [DELTA_v12_v13](docs/deltas/DELTA_v12_v13.md) |
| v13 | 2026-08-06 | Reconstrucció de zero. NULL per defecte. Domini gradual restringit als elements A-X. | [DELTA_v12_v13](docs/deltas/DELTA_v12_v13.md) |
| v12 | 2026-08-05 | Dec_Present i gating condicional de tres nivells al formulari. | [DELTA_v10_v11](docs/deltas/DELTA_v10_v11.md) |
| v11 | 2026-08-05 | Domini de cinc valors per als elements A-X, T_LOST_ELEMENTS, sistemes compositius, Record_Class, L_ELEMENTS, combos bilingües. | [DELTA_v10_v11](docs/deltas/DELTA_v10_v11.md) |
| v10 | 2026-08-03 | Consolidació del paquet. | |
| v9 | 2026-08-03 | Consolidació del paquet. | |
| v8 | 2026-08-03 | Consolidació del paquet. | |
| v6 | 2026-08-03 | Primer paquet publicat: mòduls DB i Form, esquema i metodologia. | |

## Abans de v6, sense etiqueta

| Data | Fita |
|---|---|
| 2026-08-03 | Mòduls consolidats DB i Form, en anglés i en valencià. Els generadors de primera generació queden substituïts. |
| 2026-06-24 | Formulari consolidat v2 en valencià (`chachapoya_02_Form_v2_val`). |
| 2026-06-22 a 23 | Generadors VBA de primera generació: primer en valencià (`bbdd_chachapoya`, `crear_formulari`), el mateix dia en anglés, i exportats des d'Access amb els noms `generate_db`, `generate_mainform`, `update_db`, `update_arch`, `fix_combos`, `addMissingFields`, `add_systems` i `addExtraFields`. |
| 2026-06-21 a 23 | Esquema v1 a v3 i metodologia v1 a v2, en Word. |
| 2023-03-14 | Fitxa d'estructures en Excel, precedent del disseny. |
