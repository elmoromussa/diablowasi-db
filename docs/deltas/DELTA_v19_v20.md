# DELTA v19 -> v20 (definitiu, pendent de vistiplau)

Estat: **decisions tancades** (sessions 2026-08-14/15, punts 1-4 de la
primera tanda + punts A-K de la llista d'Esteve + auditoria completa
de DB_v19a amb 86 estructures). Els blocs marcats [VALIDAR] contenen
lletra menuda d'especificacio que Esteve ha de confirmar abans de
generar scripts; cap d'ells reobri decisions.

Precondicio ja complida: duplicats de codi corregits al PC d'Esteve
(DW-S01-EA46 i DW-S04-EA10). El patch ho verifica igualment i avorta
llistant-los si en troba.

Deliverables d'aquesta versio:

| Fitxer | Paper |
|---|---|
| `chachapoya_PATCH_v20.bas` | Migracio in situ v19a -> v20 (`PatchV20`) |
| `chachapoya_DB_v20.bas` | Build complet + `RebuildQueriesV20` |
| `chachapoya_Form_v20_val.bas` | Formulari v20 (`BuildForm`) |
| `DELTA_v19_v20.md` | Aquest document |
| `esquema_bbdd_estructures_v20.md` | Referencia tecnica |
| `manual_us_bbdd_v20.md` | Manual (seccions noves: dues families del zero, dues cares de R, dues preguntes, arbre i tests ampliats, glossari) |
| `tfm_metodologia_bbdd_v20.md` | Metodologia |

**Sequencia sobre la BD amb dades:**
1. `PatchV20()` (preflight de duplicats + esquema + lookups + neteges)
2. `RebuildQueriesV20()` (mai `BuildDB()` sobre BD amb dades)
3. `BuildForm()` v20
4. `QRY_25_V20_Review` (worklist de la migracio)

---

## BLOC 1 - Muret o piler transversal (D) [TANCAT]

L'etiqueta ja esta desplegada (v19a); aci la resta de la decisio.
- `L_ELEMENTS` fila D: `Name_EN` `Transverse wall` -> **`Transverse
  wall or pier`** (pier = forma, no funcio; criteri C2 v19).
  Descripcio: afegir "from a long anchoring wall to a compact pier;
  length does not change the letter" + remissio als dos tests.
- Descripcio MEN: "bracket, transverse wall, pilaster" -> "bracket,
  transverse wall or pier, pilaster".
- Manual, Nus 11, DOS TESTS NOUS de frontera:
  * Contra E: la mensula PENJA (peca encastada que vola); el piler
    S'ALCA (fabrica que recolza i descarrega cap avall).
  * Contra K: la pilastra viu AL PLA DEL MUR de facana; el piler viu
    CONTRA LA ROCA, fora de cap parament.
- Formules constants: `base: load-bearing for pier above` i
  `base: load-bearing for corbels above` (les uniques que
  sobreviuen; la d'outline mor al bloc F).
- MOR el candidat "exempcio de K per a pilastra aillada": la
  pilastra aillada contra roca es D-piler -> MEN, registrable des de
  v18. Al Pendent queda el cami (precedent d'E) per si mai apareix
  una K aillada real al pla d'un mur desaparegut.

## BLOC 2 - Retirada d'`Interbody_Cornice_Format` [TANCAT]

- PATCH: `DROP COLUMN`. Perdua anotada: DW-S01-EA15 i DW-S01-EA36
  portaven `Tabular blocks`.
- La **regla 50 es retira** amb el camp. Ix de `QRY_13`, del
  formulari (control, caption, gating, clears) i dels documents.
- El buit de graella l'ocupa el camp nou del bloc C.

## BLOC 3 - Vora superior d'un N0 sense N1 [TANCAT]

Nomes manual + esquema (criteri h). Cap camp; el format de R es del
bloc C.
- **R = remat superior construit en general**, amb les dues cares
  escrites: parament que tanca la cova per damunt del rafec (cas
  fundacional) i remat de la massa en estructures baixes. La
  diferenciacio es DERIVA del registre (comptadors i sistemes), no
  es declara.
- **Criteri de les dues preguntes** (independents, no tria): P1 hi
  ha remat? -> R = 1/2. P2 hi ha pla d'us? -> `Sys_Platform` obert +
  Z (1/2/9) + `Platform_Function`. La terrassa general respon si a
  les dues; R3 calla (Z porta el sistema tot sol, v18 C1).
- **G = progressio**: filades successives cadascuna mes enfora;
  exigeix >=2 filades en vol; una filada volada en un sol gest mai
  no es G.
- **Prohibicio del "coronament per descart"**: R afirma estructura
  completa i acabat deliberat. Banda indecidible -> 9 + nota (+ fila
  12.Extra si mereix descripcio), eixida 8bis.7; evidencia de cos
  superior perdut -> ruta EA01-A.

## BLOC 4 - Ninxols amb pintura perimetral [TANCAT]

- Cavitats naturals usades com a contenidor funerari amb pintura al
  voltant: **NIX + files a T_DECORATIONS amb posicio ROC-PER**.
  Verificat: R40 nomes vigila `Record_Class='Rock art panel'`.
- Frontera amb 8bis.9 explicita al manual: contenidor existent
  (cavitat natural) -> NIX + decoracio; cap contenidor supervivent
  amb contorn que afirma estructura -> `Unclassifiable` + evidencia
  de pigment (EA11).
- Glossari: **ninxol natural** (tipologia NIX, la cavitat com a
  contenidor) != **ninxol decoratiu** (decoracio retallada en
  fabrica). Dos usos del terme, cap contacte.

## BLOC A - Les dues families del zero [TANCAT]

Cap canvi de valors; etiquetes + capcaleres + manual.
- `DOM5` (nomes els 21 elements constructius): `Absent` ->
  **`Absent (constatat)`**; `Desaparegut` -> **`Desaparegut
  atestat`**.
- Capcalera epistemica a 9.Bio, 10.Mat i la seccio de
  pigment/revoc: "VESTIGIS CONSERVATS - el 0 afirma 'no en sobreviu
  res observable', mai 'mai no n'hi hague'".
- Manual, seccio nova "Les dues families del zero": constructius
  (0 = decisio constructiva sobre superficie llegible) vs peribles
  (0 = supervivencia, mai origen); mapa buit/0/9/3 (buit = no
  observat; 9 = no observable; la distincio que Esteve demanava ja
  existeix); **nota de la fusta**: els elements constructius de
  fusta (A, E, F, S) son peribles pero NO desapareixen sense rastre
  - deixen encaixos (interficies negatives, cataleg del bloc I) -
  aixi que el seu 0 sobre suport llegible es de la familia forta, i
  el 3 cobreix la perdua amb evidencia.

## BLOC B - Identitat de codis [TANCAT]

- **Index unic** sobre `T_STRUCTURES.Code` (patch, amb preflight que
  avorta llistant duplicats). Identitat es bloqueja, no s'alerta.
- Criteri de numeracio al manual: **maxim + 1, mai reutilitzar
  buits** (els forats son historia de campanya).
- `L_SECTORS.Sector_Code` NOVA (patch la crea i pobla) + **regla
  57**: codi d'estructura que no comenca pel `Sector_Code` del seu
  sector -> alerta (el cas EA46, capturat estructuralment).
- Numero d'EA: **cap camp** (duplicaria el codi); expressio `Val()`
  a `QRY_07` (etiquetes QGIS, ordenacio numerica) i alla on faca
  falta.
- `QRY_26_Next_EA` NOVA: maxim i seguent lliure per sector.
- Gramatica del codi formalitzada a l'esquema: JACIMENT-SECTOR-EAn.

## BLOC C - `Upper_Crown_Format` [TANCAT]

- Camp TEXT(25) nou, default NULL (camp de judici), a `T_STRUCTURES`.
- Cataleg tancat (EN emmagatzemat / VAL en pantalla):
  `Closing panel` / Tancament vertical; `Projecting course` / Filera
  en voladis; `Flush course` / Filera a ras; `ND` / Indeterminat.
- Combinacions (panell + filera volada): el dominant al camp, la
  resta a notes; tres casos i es reobri.
- Gating: visible amb `Upper_Crown IN (1, 2, 3)` (el format d'un
  coronament atestat pot ser atestable per les lloses caigudes).
- **Regla 58** (alerta): `Upper_Crown IN (1, 2)` amb format NULL.
  ND legal; la branca es pot buidar (llico R9 de naixement).
- Migracio: els 12 registres amb R present a la worklist.
- Justificacio contra el precedent ICF, escrita: interes analitic
  declarat pel dissenyador + poblacio revisable (12) + observacio no
  derivable (cadena operativa, H01).

## BLOC D - Regla 6 i col.lapsades [TANCAT]

- **Cap conversio automatica** de 0 a 9: la llegibilitat es jutja
  superficie a superficie (model: DW-S01-EA13, cinc sistemes amb
  quatre respostes epistemiques distintes).
- Text d'accio de R6 esmenat: "use 3 with evidence, 9, **or confirm
  surface legibility in Notes**" (la tercera eixida existia i no es
  nomenava). La regla continua alertant sobre zeros confirmats:
  friccio deliberada.
- Formulari: `lblColl` es mou a baix de tot (fila 30, amplada
  completa); text nou: "AVIS: estructura col.lapsada - el 0 (absent)
  exigeix superficie llegible; en dubte, 3 (amb evidencia) o 9."

## BLOC E - Gate de portal a 2.Arq [TANCAT]

- **`N_Chamber_Bodies = 0` -> es tanquen** `Access_Plane`,
  `Portal_Orientation`, `Portal_Position` i `Recessed_Frame`, siga
  quina siga la tipologia. Les condicions per tipologia (isTer)
  queden absorbides i s'eliminen. Comptador buit = camps oberts.
- **Regla 59** (alerta): camp de portal amb valor i comptador de
  cambra a 0.
- **La maconeria NO es gateja per murs** - decisio explicita amb
  motiu escrit: 13 registres del corpus amb 0 murs i aparell
  registrat (una terrassa no te murs pero te fabrica). El seu gate
  es la classe (bloc J).

## BLOC F - Bandes de contorn [TANCAT]

- `L_DEC_TYPE`, UPDATE per nom:
  * `RA U-shape geometric` -> **`Outline band, defined`** / VAL
    `Banda de contorn definida`
  * `RA U-shape organic` -> **`Outline band, amorphous`** / VAL
    `Banda de contorn amorfa`
  * Descripcions noves DESACOBLADES del veredicte: banda de pigment
    que ressegueix un contorn, trac net / trac difus; "the
    structural verdict lives in the record's typology and
    attestation, never here".
- **8bis.9 reescrita**: la geometria baixa de criteri a INDICI; el
  veredicte estructura-vs-panell es el judici d'atestacio complet a
  nivell de registre. Motivacio documentada: a La Petaca hi ha
  estructures perimetrades amb C i amb O (contraexemple de camp).
- El reanomenament (sense prefix `RA `) mou els dos tipus al costat
  no-rupestre del filtre del desplegable TOT SOL - i es correcte:
  un panell pur no contorneja res (la U com a iconografia de panell
  ja te `RA Geometric motif`). El filtre no es toca.
- **Forma U/C/O derivable dels trams** (troballa d'auditoria): cap
  formula constant; `QRY_27_Outline_Topology` NOVA deriva la
  topologia de cada fila ROC-PER dels quatre spans (1,1,1,0 = U
  invertida; 1,1,1,1 = O; dos trams = arc/C). La hipotesi
  estructura->U / ninxol->C es comprova amb forma x tipologia x
  jaciment, sense entrada nova.
- Migracio: les 2 files geometric -> defined (automatic, patch);
  les **5 organic a revisio fila a fila** (el valor vell barrejava
  corb amb mal definit).

## BLOC G - Fases constructives [TANCAT]

- **Cap default de fases** (un default escriuria el judici "he
  buscat limits i no n'hi ha" sense mirar; argument del denominador
  sobre H03).
- **`Phase_Evidence` gatejada per `Construction_Phases >= 2`**
  (inaplicabilitat derivable -> gating, criteri v17 literal).
- **Regla 60** (alerta): fases >= 2 amb evidencia NULL. ND legal.
- Neteja al patch: `Phase_Evidence = NULL WHERE Construction_Phases
  = 1` (3 files ND sense subjecte).
- Els 77 buits: passada conscient (un clic per registre), junt amb
  metriques.

## BLOC H - Formulari per classes (art rupestre) [TANCAT]

- `Body_No`: tambe `ColumnHidden` (la vista de full de dades ignora
  `Visible` - el fantasma del "no haviem quedat que...").
- Posicio de decoracio amb classe Rock art panel: desplegable
  filtrat a **ROC-PAN, autoomplit i bloquejat** (patro
  NIX-suport v18). Les altres classes conserven la llista ROC
  completa (ninxols i estructures amb pintura la necessiten).
- `Interior_Observability` amb classe Rock art panel: deshabilitat
  i buit. Neteja al patch dels 4 valors existents (DW-S06-EA01
  Partial, DW-S01-EA60 Complete, DW-S01-EA61 Complete, DW-S01-EA37
  ND) - perdua deliberada anotada.

## BLOC I - T_ARCH_FEATURES [TANCAT]

- **`DROP COLUMN Present`**: una fila existeix perque s'ha observat
  una cosa; la casella duplicava l'existencia de la fila i el seu
  default 0 feia mentir 11 de 16 files. Cap regla en depen.
  `Feature_Count` i `Material` es queden.
- Cataleg del desplegable (value list), s'AFIGEN:
  * `Socket / negative interface` / `Encaix / interficie negativa`
    (3 casos: EA14-S01, EA04, EA17)
  * `Access bench` / `Banqueta d'acces` (2 casos: EA09, EA14-S04;
    entra amb 2 perque es l'INSTRUMENT DE RECOMPTE de l'element Y,
    reservat des de v17 pendent de tres casos)
  Els 4 tipus especulatius es queden (literatura; retirar exigeix
  mes paciencia que afegir).
- Migracio: 5 files a retipar (worklist).

## BLOC J - MEN i 2.Arq [TANCAT]

- `Structural trace` IX de la condicio que tanca `pgArq`. La classe
  va canviar de significat en v20 (murets i pilers tenen fabrica) i
  el gating s'havia quedat en el vell. La faena fina la fan els
  gates dels blocs E i G. Naturals purs i panells: tancats com fins
  ara.
- Mini-passada de dades: maconeria/suport/dimensions dels 4 MEN.

## BLOC K - Validacio al punt d'entrada [TANCAT]

- **Cap codi de colors per camp**: l'obligatorietat aci es funcio de
  l'estat (classe, sistemes, comptadors, finestres); un color
  estatic mentiria i un de dinamic duplicaria la bateria en VBA de
  colors (divergencia garantida).
- **Boto "Valida aquest registre"** a la capcalera de F_STRUCTURES:
  executa `QRY_16_Validation_Check` filtrada pel codi actual i
  mostra violacions i pendents del registre en pantalla. Una sola
  font de veritat.

---

## Regles i consultes (recompte v20)

- Regla 6: text d'accio esmenat (D).
- **Noves: 57** (concordanca codi-sector), **58** (format de
  coronament), **59** (portal vs comptador), **60** (evidencia de
  fases).
- **Retirada: 50** (amb el camp ICF). La 41 continua retirada.
- Recompte: **58 actives, numerades fins a 60**, retirades 41 i 50.
- [VALIDAR] Les 57-60 van a `QRY_16f` (que passaria de 47-56 a
  47-60, 14 branques) si JET ho tolera; si la UNION peta, s'obri
  `QRY_16g_Rules_57_60`. Decisio d'implementacio.
- **`QRY_25_V20_Review` NOVA** (segueix la serie 22/23/24), cinc
  branques: (a) format de coronament als 12 R presents; (b) 5 files
  organic a retipar amb els spans davant; (c) 5 files de 12.Extra a
  retipar (3 encaixos, 2 banquetes); (d) revisio de les dues
  preguntes: TER amb plataforma oberta i R=0, i R=2 amb plataforma
  Absent; (e) els 4 MEN amb 2.Arq buida.
- **`QRY_26_Next_EA` NOVA** (bloc B).
- **`QRY_27_Outline_Topology` NOVA** (bloc F).
- `QRY_13`: ix `Interbody_Cornice_Format`. `QRY_07`: entra
  l'expressio `EA_Num`.
- Recompte de consultes: 36 + 3 = **39** (una BD migrada des de 17a
  en porta 40, amb QRY_22).

## PatchV20 (ordre intern)

1. Preflight: `Jamb_Fabric_Reveal` existeix (marcador v19);
   recompte de codis duplicats -> si n'hi ha, MsgBox amb la llista i
   Exit Sub.
2. `DROP COLUMN Interbody_Cornice_Format` (T_STRUCTURES).
3. `DROP COLUMN Present` (T_ARCH_FEATURES).
4. `ADD COLUMN Upper_Crown_Format TEXT(25)`.
5. `ADD COLUMN Sector_Code TEXT(10)` a L_SECTORS + UPDATE per nom
   de sector [VALIDAR baix].
6. `CREATE UNIQUE INDEX` sobre Code.
7. UPDATE L_ELEMENTS D (nom i descripcio, bloc 1).
8. UPDATE L_TYPOLOGY MEN (descripcio, bloc 1).
9. UPDATE L_DEC_TYPE (dos renoms + descripcions + Name_VAL, bloc F;
   les 2 files geometric del corpus queden retipades pel renom
   mateix - el valor viu al lookup, les files apunten per ID).
10. UPDATE T_STRUCTURES: Phase_Evidence=NULL on fases=1 (bloc G).
11. UPDATE T_STRUCTURES: Interior_Observability=NULL on classe PR
    (bloc H, 4 files).

## Lletra menuda [VALIDAR en bloc]

- **Sector_Code proposats**: DW-S01..S06 -> `DW-S01`..`DW-S06`;
  DW-General -> `DW-G`; LP-General/North/Central/Upper/South ->
  `LP-G`, `LP-N`, `LP-C`, `LP-U`, `LP-S`. CONSEQUENCIA: els futurs
  codis de La Petaca serien `LP-N-EA01` etc. (la regla 57 ho
  prescriu). Si preferixes numerar els sectors LP (`LP-S01`...),
  es decideix ARA, abans que hi haja estructures LP.
- **Posicio de graella**: Upper_Crown_Format al buit de l'ICF
  (15,2), al costat de la cornisa; R es queda a 16,1. Alternativa:
  moure'l a 16,2 tocant R. Proposta: 16,2 (adjacencia semantica),
  amb 15,2 buit.
- **Boto de validacio**: obri un formulari de full de dades sobre
  QRY_16 filtrat per Code (implementacio DoCmd.OpenForm amb
  WhereCondition).
- **Concordanca de recodificacio** (auditoria): DW-S04-EA02b ->
  DW-S04-EA21; DW-S04-EA12b -> DW-S04-EA20; DW-S01-EA22b ->
  DW-S01-EA71; DW-S01-EA40b -> DW-S01-EA72. VERIFICAR si hi havia
  mes parelles amb sufix (EA07b?) i completar la taula. La taula
  entra a l'esquema com a apendix historic.

## Fora del delta - tasques de dades (consolidada, codis remapejats)

1. EA11: reclassificacio completa (v19, la primera de totes).
2. DW-S04-EA15: N/O/Q -> 9,9,9 + formula de posicio a la fila LOST.
3. DW-S04-EA01: `Sys_Interface` -> Absent (unic pas que faltava de
   la via B).
4. DW-S01-EA40: D -> 2 + `Systems_Notes` `D: beam pier`.
5. DW-S01-EA38: resoldre contradiccio Sys_Base/nota.
6. DW-S01-EA72: confirmar comptador basal.
7. Connexio EA02<->EA21 (S04): tipus -> `Associated natural context`.
8. Worklist QRY_24 v19: finestra JFR (63 confirmacions rapides) +
   6 murets NULL.
9. Les cinc branques de QRY_25 (format x12, organic x5, extra x5,
   terrasses, MEN x4).
10. Passada de fases (77 clics conscients).
11. Metriques, orientacions, coordenades, datacions, individus,
    relacions i grups: el bloc de camp/laboratori que ja coneixes.

## Tancament de pendents (tots, sessio 2026-08-15)

Per encarrec d'Esteve, cap pendent no queda com a pregunta oberta:
cada un acaba RESOLT, ABSORBIT, PROGRAMAT o convertit en
CONTINGENCIA amb l'eixida pre-decidida (el disparador es l'unica
cosa que espera; la decisio ja esta presa i escrita).

### Resolts

- **T_BODIES: NO es descompon dins del TFM.** Motius: cap analisi
  feta ni prevista ha demanat atributs per cos; la doctrina de la
  unitat de registre ja descompon els cossos realment divergents en
  registres enllacats (la recodificacio a/b -> codis propis ho
  demostra funcionant); el comptador `Fabric` cobreix la resta; i
  la migracio amb 86 registres costaria setmanes d'escala TFM.
  Post-publicacio, si una analisi ho demana, el mecanisme es el
  delta normal - no un pendent.
- **Caracteritzacio de superficies ROC: es resol per la via 3D, no
  per esquema.** El flux Blender/CloudCompare (exposicio a pluja,
  deteccio de superficies horitzontals) ES l'instrument de
  caracteritzacio; cap camp nou mai. Les metriques derivades, quan
  arriben, van al disseny T_MEASUREMENTS ja esbossat.
- **Forat de R3 sobre plataforma: TEORIC.** Verificat sobre els 86:
  cap registre amb sistema plataforma obert i tots els components a
  0 (totes les plataformes aeries tenen E o F presents). Es tanca
  sense accio; si mai apareix el cas, l'eixida esta pre-decidida:
  el patro qualificador de la v19 (finestra + regla + exempcio),
  calcat de `Jamb_Fabric_Reveal`.

### Absorbits

- **Revisio de `RA Perimeter band` (QRY_22): COMPLETA.** Zero files
  amb el tipus retirat al corpus. QRY_22 passa a historial (com
  QRY_23); la revisio fina de les bandes ja viu a la branca (b) de
  QRY_25.

### Programats (milestone, no pendent)

- **Rebuild net al deposit del TFM**, amb checklist tancada ara:
  (1) exportacio CSV de totes les taules; (2) `BuildDB()` de la
  versio final sobre BD en blanc; (3) importacio; (4) comparacio
  `mdb-schema` migrada-vs-neta; (5) bateria = buida o alertes
  conegudes documentades; (6) recomptes de files identics. Entra al
  protocol de deposit, no a cap llista de dubtes.

### Contingencies tancades (decisio presa, disparador automatic)

| Contingencia | Disparador | Eixida pre-decidida |
|---|---|---|
| **Element Y (banqueta)** | 3r cas al cataleg de 12.Extra (2 documentats: DW-S04-EA09, DW-S04-EA14) | Mini-especificacio TANCADA baix: s'executa tal qual, sense reobrir res |
| K aillada real (al pla d'un mur desaparegut) | 1r cas documentat | Precedent d'E: exempcio a R1 + regla de coherencia per tipologia |
| Remat combinat (panell + filera volada) | 3r cas amb la combinacio a notes | Es decideix entre camp secundari o formula constant, amb els casos davant |
| Convergencia D/E (`D: beam pier`) | 3r cas de la formula | Es reobri la delimitacio amb corpus; la fusio segueix descartada (contraexemple EA40+K) |
| Perfil "basament residual" | Si `Unclassifiable` amb B=1/2 supera ~5 casos | Consulta de perfil per a H01; MAI valor de tipologia (conservacio != identitat) |

### Mini-especificacio de l'element Y (pre-decidida, s'executa al 3r cas)

- Lletra **Y = Access bench** / Banqueta d'acces: banc construit o
  retallat en roca adjacent a l'acces. NO es compta com a cos.
- Camp `Access_Bench` BYTE, domini de cinc valors, **sempre actiu**
  (com D, I i R: pot ser retallat en roca i el gating per Sys_Base
  fallaria en eixe cas; el seu 0 es sempre assercio).
- `L_ELEMENTS` fila Y: "Access bench - built or rock-cut bench
  adjacent to the access; may equal the basal level in height; does
  not count as a body."
- Formulari: 11.Sist, davall del conjunt basal (fila compartida amb
  D si hi ha lloc). `QRY_13` exporta la columna Y.
- Migracio del dia: les files de 12.Extra amb tipus banqueta es
  conserven com a descripcio; el camp Y es fixa a 1/2 segons estat.
- Regla de delimitacio (manual): si el banc tanca o delimita un
  interior, no es Y - torneu al test D/V.
