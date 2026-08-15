# DELTA v21 -> v22 (tancat)

Estat: **decisions tancades** (sessio 2026-08-15, segona tanda:
verificacio de la BD v21 aplicada + observacions d'Esteve sobre
l'us real del formulari). Com el delta anterior, escrit
directament en forma tancada, sense ronda d'esborrany.

Verificacio previa que motiva part del delta: la BD v21 pujada
(84 estructures) estava **correctament migrada** — Masonry_Present
present, 57 derivats a 1, cap contradiccio, formulari v21
desplegat, plataformes Z-sol: cap. Dues troballes: la consulta
orfena `QRY_16g_Rules_57_60` (defecte del Rebuild v21, bloc E
d'aquest delta) i el retipat de 12.Extra encara pendent (ara
desbloquejat, no es defecte).

Deliverables:

| Fitxer | Paper |
|---|---|
| `chachapoya_PATCH_v22.bas` | Migracio in situ v21 -> v22 (`PatchV22`) |
| `chachapoya_DB_v22.bas` | Build complet + `RebuildQueriesV22` |
| `chachapoya_Form_v22_val.bas` | Formulari v22 (`BuildForm`) |
| `chachapoya_Worklist_v22.bas` | Navegador de worklist (`BuildWorklist`) |
| `DELTA_v21_v22.md` | Aquest document |
| `esquema_bbdd_estructures_v22.md` | Referencia tecnica |
| `manual_us_bbdd_v22.md` | Manual |
| `tfm_metodologia_bbdd_v22.md` | Metodologia (nota de versio) |

**Sequencia sobre la BD amb dades:**
1. `PatchV22()` (columna nova + derivacions + 0 dels panells +
   neteja de l'orfena)
2. `RebuildQueriesV22()` (mai `BuildDB()`)
3. `BuildForm()` v22
4. `BuildWorklist()` v22
5. Els judicis de maconeria dels naturals continuen a `QRY_29`
   (font «Migracio v21») — ara responibles des de 2.Arq mateix.

---

## BLOC A - La maconeria als naturals i als panells [TANCAT]

La v21 va crear una friccio que els numeros de la BD real
quantifiquen: dels **27 judicis de maconeria pendents, 23 seien
en classes amb 2.Arq tancada per classe** (19 contextos naturals,
4 panells) — la worklist demanava una resposta que el formulari
no deixava donar. El cas de camp que decideix: **una cavitat
natural amb bigues transversals i pis de lloses dins** (fotografia
d'Esteve, 2026-08-15) — un context natural POT dur fabrica i
cossos basals de veritat.

- **Contextos naturals: 2.Arq S'OBRI per a la classe.** El porter
  es `Masonry_Present`: amb 0 o 9 tot el bloc de maconeria es
  tanca i una cova neta es despatxa en segons; amb 1 s'obri el
  detall. Cap canvi de dades: els seus 19 judicis continuen a
  QRY_29, ara responibles.
- **Etiqueta**: es mante **«Maconeria present»** (decisio
  d'Esteve, amb la proposta alternativa «elements constructius»
  sobre la taula). Motiu registrat: els elements constructius
  no-de-fabrica (bigues, lloses de pis) ja tenen casa a 11.Sist,
  que per als naturals esta oberta; una etiqueta mes ampla
  solaparia les dues pestanyes i faria borrosa la regla 61.
- **Panells: el 0 es DERIVABLE de la classe** — un panell pur no
  te fabrica per definicio (si en tinguera, es reclassificaria).
  Pel criteri v17: inaplicabilitat derivable -> es resol sense
  judici manual. UPDATE unic al patch (4 files: DW-S01-EA37,
  DW-S01-EA60, DW-S01-EA61, DW-S06-EA01) + autoompliment al
  canvi de tipologia (patro EA-TER: **nomes sobre el camp encara
  buit, mai sobre un valor declarat**). 2.Arq continua tancada
  per a la classe panell.
- Un panell entrat per full de dades esquiva l'autoompliment i
  queda a la regla 17: acceptat (la 61 no dispara sense detall).

## BLOC B - Metrics_Available: el porter de 9.Metr [TANCAT]

Peticio d'Esteve: un camp porter del qual depenguen els altres —
*hi ha dades metriques disponibles?* Mateix patro que la
maconeria, amb **una divergencia deliberada i raonada**.

- **Camp nou** `Metrics_Available` BYTE, 0/1/9, default NULL.
  Lectura del 9: *no mesurable mai* (sense acces i sense
  cobertura de model), distint del NULL (*encara no decidit*).
- **Que gateja**: Length_m, Width_m, Height_m,
  Height_Above_Base_m (cota a 2.Arq: gate creuat entre
  pestanyes, precedent chNA), Dim_Method, Opening_Width/Height_m
  (combinat amb el gate de portal), Support_Width/Depth_m,
  Area_m2, Volume_m3, ID_Vol_Method, Vol_Notes.
- **Que queda FORA (decisio)**: les COORDENADES (Lat/Lon, UTM,
  altitud, precisio, metode) i Metric_Notes. Una estructura
  sense res mesurable continua tenint posicio, i tancar la
  georeferenciacio darrere d'un porter dimensional bloquejaria
  la dada mes basica del registre.
- **FORA de les arrays de la regla 17 (divergencia respecte de
  Masonry_Present, a proposit)**: la disponibilitat metrica no
  es decidible registre a registre fins que el flux d'extraccio
  (Metashape -> T_MEASUREMENTS, a l'horitzo) estiga en marxa.
  Posar-la a la 17 inundaria la worklist amb 82 files d'una
  pregunta que hui no te resposta. La regla 63 vigila el costat
  del detall (valor metric sota declaracio 0/9); el judici
  d'inventari es fara quan toque, en la segona passada.
- **Migracio derivable**: =1 on hi ha qualsevol valor
  dimensional (2 registres al corpus actual). La resta queda
  NULL i **no va a cap worklist** (vegeu el punt anterior).
- Exports: `Metrics_Available` entra a QRY_05 i a l'export
  germa, al costat de la declaracio de maconeria.
- UI: seccio «Disponibilitat» + porter al cap de 9.Metr;
  AfterUpdate -> ApplyGating; caption de full de dades.

## BLOC C - Nomes datacions C14 [TANCAT]

**Reversio explicita d'una decisio documentada** (la que deixava
els segles sempre editables perque «la font habitual de
l'atribucio cronologica no es radiocarbonica sino tipologica»).
Esteve declara el criteri del projecte: **les uniques datacions
que s'entraran son les del C14.**

- Gating: `Chrono_Start_Cent` i `Chrono_End_Cent` s'obrin nomes
  amb el senyal C14 marcat. El subformulari T_DATING ja penjava
  del C14 (sense canvi).
- **Regla 64** (espill de la 59 en factura): segles amb C14 a
  no -> avis. Vigila el full de dades; el formulari gateja.
- **Cost de dades: zero**, verificat sobre la BD real — cap
  registre amb C14 marcat i cap amb segles omplerts. Els camps
  naixen tancats pertot fins que arribe la primera datacio.
- Nota de reversio per al TFM: si mai es vol tornar a
  l'atribucio tipologica de segles, cal reobrir aquest bloc
  explicitament — la regla 64 i el gating son la materialitzacio
  del criteri «nomes C14».

## BLOC D - El codi a la capcalera [TANCAT]

Millora d'UX demanada per Esteve: el codi de l'estructura,
visible des de qualsevol pestanya durant l'entrada.

- La franja superior del formulari (titol + boto «Valida») viu
  fora del control de pestanyes: s'hi afig un quadre lligat a
  `Code`, **en negreta i bloquejat** (l'edicio te sa casa a
  1.Id; dos controls lligats al mateix camp son legals i el
  segon no roba el tabulador).
- De pas, el titol estatic deixa de dir «v11» (fossil de sis
  versions) i diu v22.

## BLOC E - Neteja de consultes llegades [TANCAT]

Defecte v21 reconegut: el bucle d'esborrat de
`CreateAllQueries`/`Rebuild` nomes coneix els noms VIGENTS, aixi
que el reanomenament de la setena parcial va deixar
`QRY_16g_Rules_57_60` orfena (inert pero confusa: obrir-la per
error mostraria una bateria de quatre regles).

- El build v22 esborra explicitament els noms llegats coneguts
  (`_57_60`, `_57_62`) abans de reconstruir.
- El patch tambe esborra la `_57_60` per si s'executa abans del
  Rebuild.
- Norma per a versions futures: **tot reanomenament de consulta
  afig el nom vell a la llista de llegats.**

---

## Regles i consultes (recompte v22)

- **Noves: 63** (detall metric sota declaracio negada), **64**
  (segles amb C14 a no).
- Cap retirada. Recompte: **62 actives, numerades fins a 64**,
  retirades 41 i 50.
- `QRY_16g_Rules_57_62` passa a **`QRY_16g_Rules_57_64`** (huit
  branques).
- **Cap consulta de revisio nova**: el judici de disponibilitat
  metrica queda fora de la worklist a proposit (bloc B), els 0
  dels panells els deriva el patch (bloc A), i les files de
  maconeria dels naturals ja viuen a `QRY_29` (ara responibles).
- Recompte de consultes del build: **41** (una BD migrada des de
  17a en porta mes, amb QRY_22 i QRY_28).

## PatchV22 (ordre intern)

1. Preflight: `Masonry_Present` existeix (marcador v21).
2. `ADD COLUMN Metrics_Available BYTE` (default NULL).
3. UPDATE derivable: `Metrics_Available = 1` on hi ha qualsevol
   valor dimensional (2 esperats).
4. UPDATE derivable: `Masonry_Present = 0` als panells amb el
   camp encara NULL (4 esperats).
5. Esborrat de `QRY_16g_Rules_57_60` si encara existeix.

## Fora d'abast, explicitament

- Cap camp pla de «dades metriques variades»: el buit que
  cobriria es el de la taula normalitzada `T_MEASUREMENTS` del
  flux Metashape (a l'horitzo), i un calaix de text ara
  competiria amb ella despres.
- Les COORDENADES fora del porter metric (decisio del bloc B).
- El judici de disponibilitat metrica fora de la regla 17 i de
  tota worklist fins a la segona passada.
- Cap canvi a QRY_23/QRY_24 (residu acceptat, delta v21 bloc 6).
