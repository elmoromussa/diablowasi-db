# Dades

Fitxers de treball generats durant la construcció de la base de dades. La font autoritzada de les dades és `db/chachapoya_v25g.accdb`; els CSV d'aquesta carpeta són instantànies intermèdies que es conserven perquè documenten com hi van arribar els valors.

## Procedència

Totes les dades provenen de la documentació de la necròpolis de Diablo Wasi (Leymebamba, Amazonas, Perú) feta pel projecte La Petaca (Panograma Labs i University of Central Florida, campanya de 2021) amb tècniques no invasives: fotogrametria, imatge 360 i gigapíxel. Els valors els va introduir l'autor a partir de les fitxes de camp i del model fotogramètric, seguint els criteris que descriuen `docs/esquema.md` i `docs/metodologia.md`.

Cap fitxer d'aquesta carpeta conté coordenades: les columnes de coordenades d'`EXP_STRUCTURES.csv` van buides.

## Contingut

| Carpeta | Contingut |
|---|---|
| `transfer_v11_v15/` | Exportació de les 35 estructures de la base v11 amb `tools/migrations/chachapoya_EXPORT_v11.bas` i importació a l'esquema v15 amb `chachapoya_IMPORT_v15.bas`. `EXP_STRUCTURES.csv` i `EXP_DECORATIONS.csv` són les files transferides; `EXP_ROCKART_PROPOSED.csv` i `EXP_LOST_PROPOSED.csv`, les files proposades per a revisió manual; `EXP_DROPPED.csv`, cada valor no transferit amb la regla que ho justifica. Els dos `.txt` són els registres de l'exportació i la importació. |
| `tancament_v25g/` | Fulls de tancament de NULLs que van alimentar la cascada de v25g (`tools/chachapoya_Tancament.bas`). Una fila per camp buit, amb el context, la proposta i la nota de decisió. |

## Format

- Els CSV de `transfer_v11_v15/` estan separats per punt i coma, amb els textos entre cometes. `EXP_STRUCTURES.csv` està codificat en ISO-8859-1; la resta són ASCII.
- Els CSV de `tancament_v25g/` estan separats per coma i codificats en UTF-8 amb BOM, tal com els desa Excel.
- Els codis d'estructura segueixen el patró `DW-S01-EA01` (jaciment, sector, element arqueològic).

## Condicions d'ús

Les dades es publiquen sota la mateixa llicència que la resta del repositori, [CC BY-NC 4.0](https://creativecommons.org/licenses/by-nc/4.0/deed.ca). Se'n permet la reutilització amb finalitats no comercials, amb la citació que indica el `README.md` de l'arrel.
