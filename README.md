# Base de dades arqueològica de Diablo Wasi

![Llicència: CC BY-NC 4.0](https://img.shields.io/badge/Llic%C3%A8ncia-CC_BY--NC_4.0-lightgrey.svg)
[![DOI](https://zenodo.org/badge/1358688334.svg)](https://doi.org/10.5281/zenodo.22453846)

*Design, VBA code and documentation of a relational database (Microsoft Access) for the study of the funerary structures of Diablo Wasi, a Chachapoya cliff necropolis (Leymebamba, Amazonas, Peru, 9th to 16th centuries AD). Master's thesis, Universitat d'Alacant. The interface and the documentation are in Valencian; the schema, the code and the stored values are in English.*

Disseny, codi i documentació de la base de dades relacional (Microsoft Access) per a l'estudi de les estructures funeràries de la necròpolis de penya-segat Chachapoya de Diablo Wasi (Leymebamba, Amazonas, Perú, s. IX-XVI).

L'esquema es va dissenyar per a documentar dos jaciments, La Petaca i Diablo Wasi, i conserva La Petaca com a jaciment donat d'alta amb els seus sectors. Els registres, però, són tots de Diablo Wasi: 106 elements arqueològics a la base tancada v26. La incorporació de La Petaca queda per a una versió posterior.

Treball de Fi de Màster. Màster en Arqueologia Professional i Gestió Integral del Patrimoni, Universitat d'Alacant.
Autor: Esteve Ribera-Torró. Directors: Dr. Ignasi Grau Mira (UA) i Dra. J. Marla Toyne (UCF).

## Estat del projecte

La base de dades està tancada per a la fase d'anàlisi del treball: la iteració v26 (6 de setembre de 2026) és l'última i no s'hi preveuen canvis d'esquema. El repositori es publica com a versió 1.0.0. Les etiquetes `v6` a `v26` són les iteracions de treball; les versions numerades (`1.0.0` i següents) són les publicacions. Les errates o els dubtes es poden comunicar mitjançant les *issues* del repositori.

## Estructura del repositori

| Ruta | Contingut |
|---|---|
| `src/chachapoya_DB.bas` | Mòdul VBA que construeix la base de dades completa (taules, lookups, relacions, consultes i bateria de validació) sobre un Access en blanc. |
| `src/chachapoya_Form.bas` | Mòdul VBA que construeix el formulari d'entrada `F_STRUCTURES` (interfície en valencià, valors emmagatzemats en anglés). |
| `tools/` | Utilitats d'operació: `Worklist`, `CheckMetrics`, `Relacions`, `Tancament`, `diagnostic_v13`. |
| `tools/patches/` | Pedaços que porten una base amb dades d'una versió a la següent (v16a a v26). |
| `tools/migrations/` | Migracions, exportadors i importadors entre versions (fins al migrador de v25 a v26) i l'importador de georeferenciació des de Metashape. |
| `tools/metashape/` | Scripts de Python per a Agisoft Metashape que generen els CSV de mètriques i de georeferenciació que consumeixen `CheckMetrics` i `ImportGeoref`. |
| `docs/esquema.md` | Referència tècnica de l'esquema vigent. |
| `docs/metodologia.md` | Justificació metodològica (capítol del TFM). |
| `docs/manual.md` | Manual d'ús del formulari. |
| `docs/deltas/` | Especificació de canvis de cada iteració. Són la font dels missatges de commit. |
| `docs/precedents/` | Fitxa d'estructures en Excel de 2023, antecedent del disseny. |
| `data/` | CSV de la transferència v11 a v15, mètriques extretes de Metashape per a v25 i fulls de tancament de v25g. Vegeu `data/README.md`. |
| `db/chachapoya_v26.accdb` | Base de dades tancada per a l'anàlisi, amb les dades. Única còpia Access del repositori: les bases de fites anteriors es regeneren amb el codi de cada etiqueta. |

## Versions i noms de fitxer

Els fitxers no porten el número de versió al nom: la versió la porta git. Cada iteració és una etiqueta (`v6` a `v26`) i el `CHANGELOG.md` les enumera. Per a veure què va canviar entre dues versions:

```
git diff v24 v25 -- src/chachapoya_DB.bas
git diff v24 v25 --stat
```

El prefix `chachapoya_` dels mòduls i de la base identifica el sistema de registre, dissenyat per a La Petaca i Diablo Wasi; el nom del repositori identifica el jaciment documentat. La documentació cita els fitxers amb el nom que tenien en cada paquet. Correspondència:

| Nom citat a la documentació | Fitxer del repositori |
|---|---|
| `chachapoya_DB_vN.bas` | `src/chachapoya_DB.bas` |
| `chachapoya_Form_vN_val.bas` | `src/chachapoya_Form.bas` |
| `chachapoya_Worklist_vN.bas`, `chachapoya_CheckMetrics_vN.bas`, etc. | `tools/chachapoya_*.bas` |
| `chachapoya_PATCH_vN.bas` | `tools/patches/chachapoya_PATCH_vN.bas` |
| `esquema_bbdd_estructures_vN.md` | `docs/esquema.md` |
| `tfm_metodologia_bbdd_vN.md` | `docs/metodologia.md` |
| `manual_us_bbdd_vN.md` | `docs/manual.md` |
| `DELTA_vN_vM.md` | `docs/deltas/DELTA_vN_vM.md` |

## Requisits

- Microsoft Access d'escriptori per a Windows, versió 2016 o posterior. El codi usa DAO, que Access referencia per defecte, i `Application.FileDialog` en dues utilitats.
- La base ha d'estar en una ubicació de confiança o amb les macros habilitades perquè el VBA s'execute.
- Els mòduls `.bas` s'importen amb finals de línia CRLF, tal com els desa el repositori.

## Com reconstruir la base de dades

1. Crea una base de dades Access nova i en blanc.
2. Importa `src/chachapoya_DB.bas` (Alt+F11, Fitxer, Importa fitxer) i executa `BuildDB()`.
3. Importa `src/chachapoya_Form.bas` i executa `BuildForm()`.
4. Per a portar una base amb dades d'una versió a la següent, aplica el pedaç corresponent de `tools/patches/` seguint la seqüència que indica el delta d'aquella versió. `BuildDB()` no s'ha d'executar mai sobre una base amb dades.

## Sobre aquest historial

Aquest repositori es va crear el setembre de 2026 a partir de la carpeta de treball del projecte. Els commits reprodueixen les iteracions tal com es van produir: cada versió sobreescriu els mateixos fitxers, i només els deltes i els pedaços s'acumulen perquè n'hi ha un per iteració. Les dates dels commits són les dates de modificació dels fitxers originals.

## Llicència

Aquest repositori es publica sota la llicència [Creative Commons Reconeixement-NoComercial 4.0 Internacional (CC BY-NC 4.0)](https://creativecommons.org/licenses/by-nc/4.0/deed.ca). Pots reutilitzar el codi, la documentació i les dades amb finalitats no comercials, sempre que en reconegues l'autoria. Qualsevol ús comercial requereix el permís de l'autor.

## Citació

> Ribera-Torró, E. (2026). *Base de dades arqueològica de Diablo Wasi* (versió 1.0.0) [codi i dades]. Treball de Fi de Màster, Universitat d'Alacant. Zenodo. https://doi.org/10.5281/zenodo.22453846

El DOI anterior identifica el projecte i resol sempre a l'última versió publicada; la versió 1.0.0 té el DOI propi [10.5281/zenodo.22453847](https://doi.org/10.5281/zenodo.22453847). El fitxer `CITATION.cff` conté la mateixa citació en format llegible per màquina i el botó «Cite this repository» de GitHub la mostra.
