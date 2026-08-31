# DELTA v24 -> v25 (tancat)

Estat: **decisions tancades** (sessions 2026-08-31, dins de la fase
d'operació i verificació del flux Metashape ↔ Access ↔ R; paquet
generat després que la passada v3 de l'extractor — 246 shapes, 0
invàlides, 0 llegat 2D, cossos quadrats després de la correcció
dels tres `1.0` — certificara la referència de calibratge).
Mètode: decisions primer, codi després, com mana el protocol.

Deliverables:

| Fitxer | Paper |
|---|---|
| `chachapoya_PATCH_v25.bas` | Migració in situ v24 -> v25 (`PatchV25`) |
| `chachapoya_DB_v25.bas` | Build complet + `RebuildQueriesV25` (afig QRY_30/31/32) |
| `chachapoya_Form_v25_val.bas` | Formulari v25 (dos camps nous amb gating) |
| `chachapoya_Worklist_v25.bas` | Navegador (tres fonts noves) |
| `chachapoya_CheckMetrics_v25.bas` | Pipeline de coherència: `CheckMetrics` + `ApplyMetricReview` |
| `DELTA_v24_v25.md` | Aquest document |
| `esquema_bbdd_estructures_v25.md` | Referència tècnica |
| `manual_us_bbdd_v25.md` | Manual |
| `tfm_metodologia_bbdd_v25.md` | Metodologia (nota de versió) |

**Seqüència sobre la BD amb dades:**
1. `PatchV25()` (T_METRIC_REVIEW + `Jamb_Fabric` + `Niche_Partition`
   + `Term_Lit`/`Lit_Source` amb les tres concordances)
2. `RebuildQueriesV25()` (MAI `BuildDB()` sobre esta BD)
3. `BuildForm()` (amb `F_STRUCTURES` tancat)
4. `BuildWorklist()`
5. `CheckMetrics "ruta\metriques_shapes_audit_vN.csv"` amb els CSV
   **nets vigents** (els germans es deriven del nom)
6. Firmar `Decision` a `T_METRIC_REVIEW` (Accept / Keep DB /
   Investigate) i executar `ApplyMetricReview`

---

## Bloc A: el pipeline de coherència mètriques ↔ BD

### A1. Primacia direccional (la decisió mare)

Les shapes són mesura instrumental sobre el model; els camps de la
BD, observació de camp. La primacia de l'instrument s'aplica, però
**mai plana**: una shape només parla quan existeix.

| Creuament | Regla |
|---|---|
| C2 presències | Shape = evidència positiva -> proposa presència (1/2 segons estatus; 3 si només `-r`). Cap shape no diu res: **mai es proposa baixar** un camp. |
| C1 recomptes | Les instàncies dibuixades són **fita inferior**: `shapes > BD` proposa pujar; `shapes < BD` és nota d'acta (mènsules no dibuixades són legítimes). |
| C3 cossos | Sense primacia automàtica: error d'etiqueta o cos infraregistrat, cas a cas. |
| C5 mètriques | Primacia instrumental sobre camps NULL; un futur conflicte amb valor manual és fila Pending com qualsevol altra. |
| **Tot** | **Cap escriptura sense firma.** |

### A2. `T_METRIC_REVIEW`: la firma com a registre

Una fila per discrepància: estructura (FK per nom, inline),
creuament, camp, valor BD, proposta, evidència (etiquetes de les
shapes), lot CSV, dates. `Decision` ∈ {Pending, Accept, Keep DB,
Investigate}; `ApplyMetricReview` executa **només** les Accept i
estampa `Applied_On`. La decisió queda auditada i datada: d'ací a
tres mesos se sabrà per què cada camp val el que val.

Idempotència: cada passada de `CheckMetrics` esborra les seues
Pending i repobla; les firmades es conserven i **no es
ressusciten** (clau de signatura estructura+creuament+camp+proposta).

### A3. Catàleg de creuaments

- **C1** `Timber_Bracket_Count` vs instàncies `bra` (cos+lletra).
- **C2** presències d'element vs shapes: token -> lletra (fixat ací,
  convenció v1.1 §2.1) -> camp llegit **en viu** de
  `L_ELEMENTS.Field_Name`, perquè un delta de vocabulari no
  desincronitze mai el mòdul. `wall` es reparteix per pla
  (`f`->L, `r`/`l`->V; `b` sense camp, s'omet). Compatibilitats:
  numèrics {1,2,3}; `Sys_*` textuals {Present*, Attested lost}.
- **SLOT** shapes `slot` -> `Support_Modified` = 1.
- **C3a** gramàtica de cossos: `X.Y` amb `Y >= 1` (el caçador del
  `1.0`). Fila Investigate-only, sense camp ni proposta.
- **C3b** cossos citats vs `N_Basal_Bodies`/`N_Chamber_Bodies`
  (proposa pujar; C3 proposa, no jutja els zeros).
- **C5** integració de camps NULL: `Area_m2` <- suma `t-int`;
  `Opening_Width/Height_m` <- polígon `f-port` amb estatus `-c`/`-r`
  (les mètriques de disseny exigeixen denominador original,
  convenció §7); `Length/Height/Width_m` <- murs CSV (extensió
  màxima de façana / suma d'alçades per cos / extensió r-l);
  `Support_Depth_m` <- mitjana dels `dd` de `led`; `Dim_Method` =
  'Photogrammetry' quan alguna dimensió aterra (el paral·lel
  exacte d'`ID_Coord_Method` a l'importador de georeferenciació).
  Camp ja poblat amb valor diferent -> nota d'acta, mai proposta
  d'ofici.
- **C7** `N_Built_Walls` vs plans de mur dibuixats (mateixa
  asimetria que C1).
- **C4** (cobertura corpus mètric a la BD): verificat en fase de
  disseny, 31/31; no necessita codi.

### A4. Protocol de correcció per capes

Quan un creuament destape una incoherència de vocabulari, es
corregeix **a la capa d'origen**, mai aigües avall:
1. error d'etiqueta -> Metashape + reexportació (el model és la
   font de veritat de l'observació);
2. llacuna de la convenció -> delta conscient v1.2 amb changelog;
3. llacuna del vocabulari BD -> delta de lookup.

Cas fundacional: els tres `1.0` (EA06, S04-EA01, S04-EA03) van
resultar typos d'etiqueta -> capa 1, la v1.1 aguanta. Millora
recomanada al validador de l'extractor: incorporar `Y >= 1` al
patró de cossos perquè el typo salte a la validació prèvia per
chunk (canvi menor a `export_metriques_shapes_v2.py`, pendent).

### A5. Precedent EA06/EA62 (criteri cos vs. estructura)

El discriminant és l'**esdeveniment constructiu**, no el suport:
fàbrica travada amb el coronament inferior (aparell que continua)
= un projecte, un cos més; fàbrica que hi **descansa** amb
discontinuïtat = esdeveniment posterior, EA pròpia. Compartir la
diaclasa no és criteri d'identitat. EA06/EA62 es mantenen com a
dues EA; la superposició es registra a `T_CONNECTIONS` amb
direcció determinable (EA62 posterior) — or cronològic relatiu i
evidència directa de l'ús acumulatiu que defineix la categoria de
necròpolis de cingle. Pendent manual: comprovar si la parella ja
és un dels cinc registres pendents de `T_CONNECTIONS`;
`N_Chamber_Bodies` d'EA62 el caçarà C3b.

---

## Bloc B: dos camps nous d'observació

### B1. `Jamb_Fabric` (TEXT 25, 2.Arq -> pestanya 11.Sist)

Fàbrica del brancal **existent** (el cas sense brancal ja el cobria
`Jamb_Fabric_Reveal`, finestra O ∈ {0,2}):

| Valor | Criteri |
|---|---|
| Monolithic slab | Llosa vertical de llindar a dintell als **dos** brancals |
| Composite slab-masonry | Llosa al tram inferior; per damunt, fàbrica encofrada que carega sobre la llosa i **es traba amb el parament** fins al dintell |
| Fabric as jamb | Almenys un brancal sense cap llosa pròpia: parament de dalt a baix |
| ND | Indeterminat |

Anti-ambigüitat: (1) la frontera composite/monolithic és
**constructiva, no mètrica** — una falca d'anivellament no fa
composite, la trava d'aparell sí; (2) **precedència** per a
asimetries: cap llosa en un brancal -> Fabric as jamb; si no,
qualsevol composite -> Composite. Sense valor Mixed (no s'infla el
domini); el detall per-brancal, si mai cal, viu a
`T_ARCH_FEATURES`. Gating: patró de la casa via `ElemHas(Jambs)`
(NULL manté editable: el judici no s'ha fet). Worklist: QRY_31.

Valor analític: gramàtica constructiva del mòdul portal —
integració estructural entre el sistema N+O+Q->P i el mur, parent
de la peça bifuncional lint+bra de la convenció (H01, H04/H05).

### B2. `Niche_Partition` (TEXT 15, pestanya 1.Id, davall del suport)

Llosa horitzontal que parteix el ninxol natural en dos
sub-ninxols: inversió constructiva dins d'un contenidor natural
(H04/H05; dialoga amb `Support_Modified` si hi ha encaixos
retallats). Domini {Present, Absent, Attested lost, Not
observable}; Attested lost **en reserva** (cap cas d'encaixos
buits al corpus, però és el cas més fàcil d'aparéixer demà). Sense
camp de recompte: una llosa/dos compartiments en tots els casos
confirmats — cap categoria sense cas. Gating: **només NIX** (per
tipologia); CAV exclosa per decisió (una llosa dins d'una cavitat
transitable compartimenta espai, fenomen distint — es reobriria
amb el cas davant). Worklist: QRY_32. Convenció de shapes: cap
canvi (cap NIX al corpus mètric).

---

## Bloc C: concordança bibliogràfica del vocabulari

`L_ELEMENTS` guanya `Term_Lit` (terme publicat) i `Lit_Source`
(cita curta). Tres concordances confirmades sobre el text llegit
sencer:

| Element | Term_Lit | Font |
|---|---|---|
| B Base_Level | platform-base | Guengerich 2014 |
| I Interbody_Cornice | cornice | Guengerich 2014 |
| M Relief_Frieze | frieze | Guengerich 2014 |

La d'I és concordança **de posició, no sols de nom**: Guengerich
defineix la cornisa a la juntura platform-base/superstructure, la
mateixa posició arquitectònica de l'element I. Termes **sense**
correspondència de camp (documentats a la nota metodològica, no a
la taula): *superstructure* (concepte = cossos N1), *bedrock*
(suport geològic, L_SUPPORT), *sub-cornice* (Fig. 3 de Guengerich
2014, motllura secundària cornisa-fris: **cap cas al corpus**, fora
per principi, amb l'àncora bibliogràfica esperant), *ledge*
(Epstein & Toyne 2016; ja recollit com a `led` a la convenció).
Reforça OE1 pel mateix mecanisme que Aoujgal: el vocabulari
dialoga amb la terminologia publicada.

---

## Errates i esmenes de la sessió d'estrena (2026-08-31, incorporades)

La primera execució real del pipeline sobre la BD viva va destapar
tres coses; totes tres queden **resoltes dins d'aquest mateix
paquet** (fitxers republicats), i es documenten perquè la història
del delta incloga la depuració:

1. **`HasK` i l'error 457.** El test d'existència de clau feia una
   assignació *Let* (`v = col(k)`), vàlida per a ítems escalars
   però no per a les sub-col·leccions d'evidència: l'assignació
   petava, `HasK` responia fals «no existeix», `SubAdd` reafegia la
   clau i JET responia 457. Esmena: `VarType(col(k))`, que accepta
   qualsevol variant, objectes inclosos.
2. **La col·lecció `codes` iterada per valor.** Guardava `"1"` com a
   valor amb el codi com a clau; les col·leccions VBA no rendeixen
   les claus, així que el bucle principal llegia `"1"` trenta-una
   vegades i cap estructura casava amb la BD. Esmena: el codi entra
   com a valor **i** com a clau.
3. **`QRY_30a_Metric_Signing` (consulta de firma).** El disseny
   original enviava a firmar al full de dades de `T_METRIC_REVIEW`,
   on l'estructura és un FK numèric il·legible. La consulta de
   firma porta el codi de l'estructura al costat de les tres
   columnes de decisió, **editables** (JET ho permet perquè
   T_STRUCTURES hi entra per la seua clau primària), i filtra
   Pending perquè es buide a mesura que es firma. Incorporada a
   `BuildV25Queries` (i al cicle esborra-i-crea de
   `RebuildQueriesV25`); qui la tinga creada a mà des de l'Immediat
   no ha de fer res — el rebuild següent la regenera de sèrie.

Validació d'estrena: l'acta real (167 files Pending: 130 C5, 20 C2,
8 C7, 7 C3b, 1 C1, 1 SLOT; C3a a zero; 7 notes informatives) va
quadrar amb la simulació prèvia del paquet, amb les úniques
divergències explicades per l'avanç de la BD viva respecte de la
còpia simulada (EA62 i S05-EA01 ja corregits a mà). El protocol de
firma en bloc per a C5 (validació mostral d'una estructura
coneguda + UPDATE conscient sobre les files C5 Pending) queda
sancionat com a pràctica legítima: la firma és l'acte deliberat,
no el clic.

---

## Pendents que este delta NO obri

- Correcció `Y >= 1` al validador de l'extractor (canvi menor a
  `export_metriques_shapes_v2.py`, anotat a A4).
- Fase 3 del pla: menú estil «oculat» per a la família d'scripts
  de Metashape (`install_menu` al peu de l'extractor).
- Registre manual de la superposició EA06↔EA62 a `T_CONNECTIONS`
  (comprovar els cinc pendents) i la resta de la revisió de
  direccionalitat.
- El volum interior (`Volume_m3`) queda fora de C5: cap font
  directa als CSV; la conversa R->BD del delta v24 continua
  sobirana sobre la semàntica volumètrica.
- Residus de qualitat de dades v23/v24 (etiqueta «Unclassifiable»
  amb referència a EA11; files de T_LOST_ELEMENTS amb
  Element_Code en blanc).
