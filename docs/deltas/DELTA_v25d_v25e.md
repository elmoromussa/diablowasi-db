# DELTA v25d -> v25e (tancat)

Estat: **decisions tancades** (sessió 2026-08-31). El delta té dos
blocs de naturalesa distinta: **A**, la decisió que ja esperava dades
(el domini de muntants amb segon membre no observable, anunciat al
DELTA v25 com «fins la decisió v25e»); i **B**, un incident descobert
durant la revisió prèvia a la generació del paquet: el bug de locale
d'`ApplyMetricReview` que va escriure 125 valors C5 multiplicats per
100. Mètode: decisions primer, codi després; l'auditoria del bloc B es
va fer sencera sobre la BD viva abans d'escriure una línia.

**v25e no toca esquema**: cap camp nou, cap lookup nou. El domini de
`Jamb_Fabric` viu al combo del formulari i a la regla R69, i la
reparació re-deriva valors de `T_METRIC_REVIEW`. Per això el paquet és
quirúrgic i no obri v26 (que conserva el seu inquilí congelat: la
jubilació de `Jamb_Fabric_Reveal` amb migració — ara amb un motiu més,
vegeu A6).

Deliverables:

| Fitxer | Paper |
|---|---|
| `chachapoya_PATCH_v25e.bas` | `RepairC5Locale` + `SurgeryMuntantsV25e` (informe/aplicació, verificació prèvia i posterior) |
| `chachapoya_DB_v25e.bas` | Build complet amb R69 esmenada (`RebuildQueriesV25` regenera la bateria) |
| `chachapoya_CheckMetrics_v25e.bas` | Pipeline amb l'arrel del bug corregida (`Val()` a la frontera text→camp) |
| `chachapoya_Form_v25e_val.bas` | Combo de `Jamb_Fabric` amb els tres valors NObs |
| `DELTA_v25d_v25e.md` | Aquest document |
| `esquema_bbdd_estructures_v25e.md` | Referència tècnica |
| `manual_us_bbdd_v25e.md` | Manual |
| `tfm_metodologia_bbdd_v25e.md` | Metodologia (nota de versió v25e) |

**Seqüència sobre la BD amb dades (estat v25d):**
1. Reemplaçar mòduls: PATCH v25e (nou), DB v25e, CheckMetrics v25e,
   Form v25e
2. `RepairC5Locale` (informe) → `RepairC5Locale True`
   (esperat: 125 reparats, 0 divergents, 28 omesos)
3. `SurgeryMuntantsV25e` (informe, 7/7 en verd) →
   `SurgeryMuntantsV25e True`
4. `RebuildQueriesV25()` (R69 esmenada; MAI `BuildDB()`)
5. `BuildForm()` (amb `F_STRUCTURES` tancat)
6. Obrir `QRY_16_Validation_Check` i confirmar el verd

---

## Bloc A: el domini NObs de la parella de muntants

### A1. La decisió

El corpus tenia sis portals amb un únic muntant llegible, entrats
duplicant el membre observat (`Composite-Composite` amb un sol muntant
conservat, etc.) i anotats a `Systems_Notes`. La duplicació destruïa
la distinció entre «els dos costats resolen igual» i «només un costat
és llegible». El domini guanya tres valors **epistèmics, no
tipològics** — no inventen cap combinació constructiva; deixen de
perdre la meitat observada:

| Valor BD | UI | Semàntica |
|---|---|---|
| Slab-NObs | Llosa + no observable | Un muntant resol com a llosa; l'altre no és observable |
| Composite-NObs | Composta + no observable | Un muntant resol com a composta; l'altre no és observable |
| Fabric-NObs | Fàbrica + no observable | Un muntant resol com a fàbrica; l'altre no és observable |

Els tres tenen cas confirmat al corpus (principi de la casa satisfet
sense necessitat de reserva). Ordre canònic sense canvi: l'observat
primer, `NObs` sempre segon. Grafia: `-NObs` (compacte, cap dins de
TEXT(25), inequívoc en cerca exacta; es descarta reutilitzar `ND` com
a sufix per no sobrecarregar el token a dos nivells — trampa per a
patrons `Like '*ND'`).

**Terminologia confirmada (decisió en sessió): les etiquetes conserven
«llosa»**, no «brancal». Motiu: brancal = l'element singular O (la
peça), muntant = la posició; i la **composta també té brancal** (el
tram inferior de llosa és element O present, R69 ho exigeix). Etiquetar
la resolució (a) com a «brancal» suggeriria que només un tipus en té.
Les pròpies notes de camp mostren la col·lisió: «brancal de llosa de
baix a dalt», «brancal mixt (brancal + fàbrica)» — el mateix mot per a
la peça i per al tipus dins d'una frase.

### A2. ND, redefinit

Amb els valors NObs al domini, **ND s'estreteix**: ja no és «muntant
no observable» genèric sinó «**cap muntant amb resolució
determinable**» — no observable *o indeterminat*. El cas que va forçar
la precisió és DW-S04-EA12: es conserva una llosa de brancal, però
sense el dintell ni res per damunt no es pot saber si arribava dalt
(llosa) o portava fàbrica travada (composta). No obri categoria nova
perquè el fet positiu (existeix llosa) **ja viu a l'element O**
(Jambs=2): la parella només perd la resolució, que és exactament el
que ND diu. Cap fet registrat dues vegades.

### A3. R69, esmenada

La branca d'O absent de la regla nomenava «tot el que no és
Fabric-Fabric», i amb el domini ampliat hauria disparat en fals sobre
`Fabric-NObs`. Esmena:

- `Slab-*` i `Composite-*` (NObs inclosos): afirmen almenys una llosa
  → **incompatibles amb O=0**, disparen com abans.
- `Fabric-Fabric`: cap llosa enlloc → incompatible amb O∈{1,2,3},
  sense canvi.
- `Fabric-NObs`: el costat llegible no té llosa, l'ocult podria
  tindre'n → **compatible amb qualsevol O, mai dispara**. És l'única
  parella del domini sense poder de contradicció, per construcció.
- NULL, ND i O=9: mai disparen, sense canvi.

La bateria conserva el nom `QRY_16g_Rules_57_69` (esmena de regla, cap
regla nova). Simulació sobre l'estat post-cirurgia: cap alarma.

### A4. Acta de reclassificació (executada per `SurgeryMuntantsV25e`)

| Codi | Abans | Després | Fonament (nota de camp) |
|---|---|---|---|
| DW-S04-EA18 | Slab-Slab | **Slab-NObs** | «un únic muntant conservat, brancal de llosa de baix a dalt» |
| DW-S02-EA01 | Composite-Composite | **Composite-NObs** | «tan sols roman el muntant dret, tipus compost» |
| DW-S03-EA01 | Composite-Composite | **Composite-NObs** | «tan sols es conserva el muntant esquerre, composta» |
| DW-S01-EA03 | Composite-Composite | **Composite-NObs** | «tan sols es conserva el muntant dret» (tipus compost, confirmat en sessió) |
| DW-S04-EA12 | Composite-Composite | **ND** | llosa conservada sense poder determinar la resolució (vegeu A2); l'evidència positiva la porta O=2 |
| DW-S01-EA22 | Fabric-Fabric | **Fabric-NObs** | «un únic muntant conservat, de fàbrica» |

Dues decisions annexes firmades en sessió:

- **DW-S01-EA22, element O es manté en 0.** L'autòpsia de l'element és
  «avaluat i sense brancal» sobre l'observable; la parella
  `Fabric-NObs` porta l'epistemologia del muntant ocult, i
  `Facade_Observability` dona el context. Cada capa diu la seua
  veritat: la separació muntant/brancal treballant com toca.
- **DW-S02-EA02: la parella es queda ND** («es constata un portal,
  però no s'observa prou bé per saber la tipologia dels muntants»), i
  en conseqüència el `Jamb_Fabric_Reveal = 1` queda en fals — era una
  asserció positiva d'observació que la mateixa fitxa ara nega. **JFR
  passa 1 → 9** amb nota d'acta a `Systems_Notes` (pendent de
  verificació amb foto o en camp). La cirurgia ho executa amb la resta.

### A5. Worklist i formulari

`QRY_31` no canvia: pendent = NULL; `ND` i els `-NObs` són estats
resolts. El combo del formulari afig les tres parelles noves amb les
etiquetes «Llosa / Composta / Fàbrica + no observable» (redesplegat
complet de `BuildForm`, patró de la casa).

### A6. Efecte colateral sobre v26

`Jamb_Fabric_Reveal` (congelat com a candidat a jubilació) suma un
motiu: la seua semàntica («fàbrica fa de brancal») és ara derivable de
la parella *-Fabric **i** el cas EA02 demostra que els dos camps poden
contradir-se — dos llocs per a un fet és exactament la patologia que
la jubilació amb migració resoldrà.

---

## Bloc B: l'incident de locale (CDbl ×100) i la seua reparació

### B1. Diagnosi

Descobert en la revisió prèvia al paquet (la sospita inicial era una
confusió m/cm). La línia d'escriptura d'`ApplyMetricReview` v25 feia:

    rs2.Fields(fld).Value = CDbl(Replace(pv, ",", "."))

La defensa contra la coma garantia que la cadena porta **punt**, i
`CDbl` — que respecta el locale — sota configuració regional de coma
decimal llig el punt com a separador de milers: `CDbl("1.39") = 139`.
Tots els valors numèrics C5 (guardats a `Proposed_Value` amb dos
decimals via `Fmt2`) van entrar **exactament ×100**. El `Debug.Print`
de l'acta mostrava el valor bo («NULL -> 1.39») mentre la taula rebia
139: la pantalla deia la veritat i la BD no.

Nota per a la posteritat: als camps lineals el número resultant
*coincideix numèricament* amb els cm (1.81 m → 181), cosa que va fer
que l'incident es llegira inicialment com a error d'unitats; però
`Area_m2` també és ×100 (1.39 → 139), que no són cm² (serien ×10.000).
No és conversió: és el punt decimal perdut.

### B2. Abast (auditat fila a fila sobre la BD viva)

- **Afectat**: 125 valors en 6 camps — `Length_m` (28), `Height_m`
  (28), `Width_m` (16), `Area_m2` (25), `Opening_Width_m` (14),
  `Opening_Height_m` (14). Els 125 són ×100 exactes de la proposta;
  tots sobre NULL previ (`DB_Value` buit en el 100% de files C5): cap
  valor manual trepitjat. L'únic valor manual del bloc,
  `Height_Above_Base_m = 3.00`, està intacte.
- **Net**: `ImportGeoref` (usa `Val()`, locale-segur — coordenades,
  altituds i azimuts verificats en rang); totes les aplicacions no-C5
  (C1, C2, C3b, C7, SLOT: enters i textos, valor viu = proposta al
  100%); l'interior de `CheckMetrics` (col·leccions de Doubles, CSV
  llegits amb `Val()` — per això les **propostes** eren correctes);
  `Dim_Method` (text).
- **Conseqüència derivada**: qualsevol consulta o export executat
  entre l'aplicació C5 i la reparació (`QRY_05`, `QRY_06`, exports R)
  portava estos camps ×100. **Re-executar el que s'haja mirat.** Res
  publicat.

### B3. Reparació (`RepairC5Locale`)

La reparació **no divideix per 100**: re-deriva cada valor des de
`Proposed_Value` de la fila de firma — el registre d'auditoria com a
font de reparació, que és exactament el motiu pel qual
`T_METRIC_REVIEW` guarda la proposta literal. Per fila: si el valor
viu ja és la proposta, salta (idempotent); si és proposta×100, repara
i estampa «Repaired v25e (locale CDbl x100)» a `Notes` de la fila
(incident datat al mateix registre); si no quadra amb cap de les dues,
**no toca i llista** (protecció contra edicions manuals posteriors).
Mode informe primer, aplicació després. Simulat sobre l'export de la
BD v25d: 125 reparats, 0 divergents, 28 omesos.

### B4. Arrel

`CheckMetrics` v25e substitueix el `CDbl` de la frontera text→camp per
`Val(Replace(pv, ",", "."))` — `Val()` interpreta sempre el punt com a
decimal, independentment del locale. Principi que queda escrit:
**qualsevol frontera text→número en este codebase usa `Val()`, mai
`CDbl` sobre cadenes** (`CDbl` només sobre valors ja numèrics).

---

## Serrells que este delta tanca de passada

- **`QRY_16g_Rules_57_68` (consulta òrfena)**: verificat sobre la BD
  v25d que ja no existeix — la línia d'higiene es va executar. Tancat.

## Pendents que este delta NO obri (herència viva)

- Registre de la superposició EA06↔EA62 a `T_CONNECTIONS` (criteri de
  l'esdeveniment constructiu, DELTA v25 §A5) i revisió de
  direccionalitat dels cinc registres existents. **Següent de la
  sessió.**
- Delta v1.2 de la convenció de shapes (bloc acumulat al DELTA v25) i
  validador `Y >= 1` a `export_metriques_shapes_v2.py`.
- Fase 3: menú «oculat» per a la família d'scripts de Metashape.
- v26 (quan s'òbriga): jubilació de `Jamb_Fabric_Reveal` amb migració
  (vegeu A6).
- Agrupació funcional i jerarquia pare-fill (`ID_Parent`, `ID_Group`,
  `T_CONNECTIONS`) via script quirúrgic — punt 2 de l'agenda de
  sessió, decisions per obrir.
