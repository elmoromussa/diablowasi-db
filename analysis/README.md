# Paquet d'anàlisi del corpus de Diablo Wasi (BD v26)

Pipeline analític complet del TFM *Necròpolis de cingle de Diablo Wasi*
(E. Ribera-Torró, Màster MAGIP, Universitat d'Alacant, 2026). Reprodueix
totes les anàlisis estadístiques citades a la memòria a partir de la base
de dades publicada en aquest mateix repositori (`db/chachapoya_v26.accdb`).

Sessió analítica: 7 de setembre de 2026. Corpus congelat v26 (106 EA).

## Entorn

- R 4.3.3 amb `igraph` i `ggplot2` (la resta és base R)
- Python 3 amb `pandas` (només per a 01)
- `mdbtools` (només per a 00; alternativament, exporteu les taules des
  d'Access mateix)

## Execució (des d'`analysis/`)

```bash
bash scripts/00_extract_data.sh          # taules BD -> data/*.csv
python3 scripts/01_build_matriu_AX.py    # matriu A-X (rèplica de QRY_13)
Rscript scripts/02_analisi_DW_v26.R      # passada 1
Rscript scripts/03_analisi_DW_v26_pas2.R # passada 2
Rscript scripts/04_analisi_H02.R         # contingències H02
```

Les extraccions locals `data/T_*.csv` i `data/L_*.csv` no es versionen
(vegeu `data/.gitignore`); es regeneren amb el pas 00. Cap fitxer d'aquest
paquet conté coordenades: les extraccions es limiten als camps que consumeixen
els scripts i les sortides no en porten.

## Decisions analítiques (tancades el 2026-09-07)

- **D1 Població**: *Built funerary structure* + *Natural funerary context*
  amb cobertura >= 50% d'elements avaluables -> **n = 92**. Exclosos els
  panells d'art rupestre, les traces MEN i 6 EA per cobertura insuficient.
- **D2 Codificació**: lectura primària «què es va construir»
  ({1,2,3} -> 1, {0} -> 0, {9} -> NA; distàncies amb exclusió per parells).
  Sensibilitat amb la lectura «què sobreviu» ({1,2} -> 1): concordança de
  pertinença als clústers = 0,92.
- **D3 Elements**: els 21 de QRY_13 menys A (constant-absent al corpus).
  `Support_Modified` exclòs de la co-ocurrència; els sistemes H/P/U no
  entren com a columnes (v. justificació de la promoció a sistemes,
  `docs/esquema.md`).
- **D4 Mètodes**: distància de Jaccard + UPGMA (k = 5 per silueta mitjana);
  co-ocurrència element x element (índex de Jaccard); components i punts
  d'articulació amb `igraph` sobre les 91 arestes de T_CONNECTIONS;
  estadística circular bàsica per a les orientacions de façana; khi-quadrat
  de Monte Carlo (B = 20.000) i Fisher simulat per a les contingències.

## Estructura

```
analysis/
  scripts/          00..04 (pipeline complet, en ordre)
  data/             matriu_AX_cinc_valors.csv (derivada, regenerable amb 01)
                    metriques_murs_v5.csv (mètriques fotogramètriques de mur;
                    vegeu «Procedència de les dades»)
  output/
    logs/           resultats_analisi.txt, resultats_pas2.txt, resultats_H02.txt
    tables/         perfils de clúster, Jaccard d'elements, components,
                    punts d'articulació
    figures/        11 figures PDF (les citades a la memòria del TFM)
```

## Procedència de les dades

- `data/matriu_AX_cinc_valors.csv` és una derivada de `T_STRUCTURES` de la
  base v26: el pas 01 replica la lògica de `QRY_13_AX_Pattern_Export`
  (anul·lació dels elements segons el sistema contenidor). Es versiona perquè
  el pipeline es puga reprendre a partir del pas 02 sense Access ni mdbtools.
- `data/metriques_murs_v5.csv` és la cinquena passada de
  `tools/metashape/export_metriques_shapes.py` sobre el model fotogramètric,
  feta el setembre de 2026 amb el conveni de *shapes* v1.1. Substitueix
  `data/metriques_v25/metriques_murs.csv` (passada del 30 d'agost) amb les
  correccions posteriors a la revisió de v25: reetiquetatge de les *shapes*
  d'EA35 a EA36, valors revisats de S03-EA01 i mesures d'altura noves; cobreix
  29 estructures en lloc de 28. Té el mateix format (UTF-8 amb BOM, una fila
  per estructura, mur i cos) i cap coordenada. Només s'usa a la passada 2
  (esforç constructiu), on la intersecció amb la població analítica és de
  5 EA.

## Notes de reproducibilitat

- Tota l'aleatorietat està fixada amb `set.seed(26)`; els p-valors de
  Monte Carlo poden oscil·lar en l'última xifra significativa.
- Camps de la BD buits en tot el corpus i, per tant, mai invocats per cap
  script ni pel text del TFM: `Volume_m3`, `Support_Width_m`,
  `Support_Depth_m`.
- La calibració radiocarbònica NO es recalcula ací: els rangs 2-sigma
  (SHCal20, amb IntCal20 comparatiu) provenen dels certificats ORAU i
  estan emmagatzemats a `T_DATING`. El model bayesià de fases s'executa amb
  OxCal fora d'aquest pipeline i no forma part del repositori.

## Llicència i citació

Els scripts (`scripts/`) es publiquen sota la llicència MIT; les dades i les
sortides (`data/`, `output/`) sota CC BY-NC 4.0, com la documentació i la
base (vegeu `LICENSE`, `LICENSE-DOCS-DATA` i el `README.md` de l'arrel per a
la citació).
