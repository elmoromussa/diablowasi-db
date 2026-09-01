# DELTA v25e -> v25f (tancat)

Estat: **decisions tancades** (sessió 2026-08-31, segona part). El
delta és **només de dades**: cap camp nou, cap valor de domini nou,
cap canvi de formulari. El domini tancat de `Connection_Type` (v11)
ja contenia les cinc famílies necessàries — `Shared support`,
`Abutted vertical joint`, `Superposition`, `Associated natural
context`, `Vertical association` — i el principi «cap categoria nova
sense necessitat» ha treballat a favor: la campanya sencera entra amb
el vocabulari existent.

Deliverables:

| Fitxer | Paper |
|---|---|
| `chachapoya_Relacions_v25f.bas` | `RelacionsV25f` (informe/aplicació): 9 paternitats + 1 correcció tipològica + 4 esmenes + 83 arestes |
| `DELTA_v25e_v25f.md` | Aquest document |
| `esquema_bbdd_estructures_v25f.md` | §3.6 amb la doctrina de capes v25f |
| `manual_us_bbdd_v25f.md` | «Què ha canviat en v25f» |
| `tfm_metodologia_bbdd_v25f.md` | Nota de versió «El graf abans que el grup» |

**Seqüència:** importar el mòdul → `RelacionsV25f` (informe, esperat
tot en verd amb 72 arestes verificades) → revisar l'informe →
`RelacionsV25f True` → reobrir `QRY_14_Connections_Edges` i
`QRY_16_Validation_Check`. Cap altre pas: ni `RebuildQueries` ni
`BuildForm` (res no canvia).

---

## A. La decisió de model: arestes, no grups

La primera formulació de l'inventari de camp proposava *grups*
d'estrat (`ID_Group`). El model va caure amb les dades a la mà:
DW-S01-EA02 i DW-S01-EA46 pertanyien a dos estrats alhora, i 22
estructures participaven simultàniament en cadenes horitzontals i
verticals — impossible amb un `ID_Group` únic. La decisió (proposada
per Esteve, adoptada): **tot són parells encadenats a
`T_CONNECTIONS`**; els *clusters* són components connexos del graf,
derivables a R (`igraph`) filtrant per família d'aresta. La
pertinença múltiple desapareix per construcció, cap fet es registra
dues vegades, i `T_GROUPS` queda reservada per a agrupacions
genuïnament funcionals-interpretatives futures.

**Precedència quan un parell té dues lectures** (generalitza «la
junta mana sobre el suport», v-anteriors): contacte > proximitat >
estrat compartit. Aplicació final: 10 parells d'estrat subsumits per
la seua junta adossada (derivables del component, cap pèrdua), 1
parell **mediat** (EA04↔EA05 s'adossen a través de l'arranc murari
EA77: cap aresta directa, dues d'indirectes) i 1 **reclassificat**
(EA49↔EA73, d'estrat a `Horizontal association` — no comparteixen
estrat estricte: EA49 és un ninxolet, possible forat de mènsula, a
~1.4 m de la base, i EA73 a ~2 m; debut de la família v17 al corpus).

**Usos fixats de les famílies:**
- `Shared support`: mateix estrat de suport geològic, **la distància
  no importa**; cadenes → parells consecutius en l'ordre físic.
- `Vertical association`: proximitat vertical **o diagonal** sense
  contacte (material d'OE3/H03).
- `Associated natural context` + principi firmat: **el context
  natural precedeix la construcció** → `Sequential`, `ID_Earlier` =
  el context natural (NIX en els dos casos, tipologia verificada).

## B. QC geomètric previ (cap escriptura sense passar-lo)

Sobre les coordenades UTM/altitud de la BD: els 12 estrats mostren
desnivells consecutius ≤3.7 m i **cap inversió d'ordre** al llarg del
front (l'ordre de llista = ordre físic, confirmat); les 16 cadenes de
proximitat mostren separació vertical franca (1.8–13.1 m) amb un únic
cas límit — DW-S01-EA72↔EA39, dAlt 0.7 m, més diagonal que vertical —
anotat a la seua aresta.

**Segona passada (auditoria inversa, veïns pròxims sense aresta):**
va caçar un error real — DW-S02-EA01↔EA02 estava com a adossada amb
**11.8 m** entre marcadors (precisió 5 cm): impossible; passa a
`Shared support`, la lectura d'estrat que la llista de camp també
contenia — una omissió (DW-S01-EA75↔EA22, a 2.0 m, entra com a
proximitat diagonal) i una llista de candidats que Esteve va triar un
per un: sis proximitats noves acceptades (quatre d'elles són els
ponts verticals entre els estrats S5 i S6: EA61↔EA29, EA31↔EA55,
EA33↔EA56, EA55↔EA34; més EA48↔EA19 i EA16↔EA15) i la resta
descartades deliberadament (EA49↔EA19, EA48↔EA73, EA03↔EA45,
EA07↔EA62 — esta última derivable de EA62 sobre EA06 adjacent a
EA07). **El cas EA03↔EA04 fixa doctrina**: sense contacte directe
actual però amb probable adossament en origen (EA03 conserva només el
parament lateral dret d'una façana de mausoleu, elevat sobre pedra),
l'aresta registra **l'observable** (`Vertical association`) i la
hipòtesi restituïda viatja a `Notes` amb el seu argument — la lògica
del `-r` de la convenció de shapes aplicada al graf; si el model
mostra mai contacte de filades, l'aresta es promou amb un UPDATE. I
EA03 guanya la seua relació principal: `Associated natural context`
amb EA44 (context natural quasi a sota, anterior per principi).

## C. Acta d'escriptura

### C1. Paternitat (9 assignacions; contenció física, pares = EA de ple dret)

| Pare (context) | Filles noves |
|---|---|
| DW-S01-EA76 (diedre, sostre compartit) | DW-S01-EA02, EA13, EA14, EA51, EA59 |
| DW-S01-EA42 (cova gran; EA35/EA36/EA74 ja hi eren) | DW-S01-EA34 |
| DW-S05-EA02 (cova gran) | DW-S05-EA01 |
| DW-S06-EA04 (cova gran) | DW-S06-EA01, DW-S06-EA02 |

**Correcció tipològica annexa:** DW-S01-EA76 estava tipificat 9
(*Unclassifiable*), reservat per doctrina a estructures **construïdes**
indeterminables («els contextos naturals sempre són classificables com
NIX o CAV»). Passa a **6 (CAV)**; la morfologia de diedre queda a la
fitxa.

### C2. Esmenes de files existents de T_CONNECTIONS

| ID | Parell | Abans | Després |
|---|---|---|---|
| 3 | DW-S01-EA22↔EA71 | Superposition, Undetermined | **Abutted vertical joint**, Sequential, earlier = **EA71**; nota: persisteix superposició-escaló de massa reduïda |
| 4 | DW-S01-EA40↔EA72 | Abutted, Undetermined | + Sequential, earlier = **EA40** |
| 5 | DW-S04-EA02↔EA21 | Other (see Notes) | **Associated natural context**, Sequential, earlier = **EA21** (nínxol); nota: plataforma apilada sobre el nínxol |
| 6 | DW-S04-EA12↔EA20 | Superposition | **Abutted vertical joint** («contígua i no perpendicular») |

### C3. Arestes noves (83; Confidence = High, com les existents)

- **Shared support (36, 35 de noves):** les 12 cadenes d'estrat menys
  els parells subsumits, mediats o reclassificats; inclou
  DW-S02-EA01↔EA02 (esmenada d'adossada a estrat per l'auditoria:
  11.8 m entre marcadors) i els dos parells de la **cadena estesa
  v25f.1** — EA46↔EA73 (l'estrat va directe; el ninxolet EA49 en
  penja via EA46) i EA19↔EA02 (pont entre les zones EA46–EA19 i
  EA02–EA18), que confirmen que S2, S3 i S4 formen un únic component
  d'estrat.
- **Abutted vertical joint (9 noves):** DW-S01: EA17↔EA71, EA14↔EA51,
  EA51↔EA59, EA73↔EA19, EA48↔EA47, EA04↔EA77, EA77↔EA05 (EA77:
  arranc murari nou entremig; el parell EA04↔EA05 queda mediat),
  EA45↔EA44, EA20↔EA75 (al fons de la cova).
- **Horizontal association (1, debut):** DW-S01-EA49↔EA73, amb la
  descripció del ninxolet a la nota.
- **Superposition (1):** DW-S01-EA06↔EA62, Sequential, earlier =
  **EA06** (estratigrafia firmada).
- **Associated natural context (2):** DW-S03-EA01↔EA02 (primera
  relació registrada a S03) i DW-S01-EA03↔EA44; en les dues,
  Sequential amb earlier = el context natural. **Doctrina explícita
  (petició d'Esteve):** la família cobreix **qualsevol context
  natural (CAV o NIX) funcionalment associat a qualsevol estructura
  construïda** — no només nínxol + plataforma; el corpus ja ho
  il·lustra amb tres tipologies construïdes distintes (plataforma,
  mausoleu, cambra) i els dos tipus de context (NIX i CAV).
- **Vertical association (35):** les 16 cadenes de proximitat, més
  EA75↔EA22 (omissió caçada per l'auditoria), més les sis triades del
  llistat de candidats, més EA03↔EA04 (amb la hipòtesi d'adossament
  original a la nota); inclou DW-S01-EA60↔EA39 (correcció del tecleig
  EA6039) i DW-S01-EA20↔EA22.

**Totals resultants:** `T_CONNECTIONS` = 88 files; `Sequential` amb
`ID_Earlier` = 6; `ID_Parent` poblats = 12. Regla 52 i índex
`UQ_CONN_PAIR` satisfets per construcció (normalització A<B per ID).
DW-S01-EA77 (tipologia 9, correcta: estructura construïda de funció
indeterminada — el contrast exacte amb el cas EA76) es verifica en
runtime pel mateix script.

## D. Resolucions de sessió incorporades

- `DW-S01-EA6039` → **EA39**; `DW-S01-EA3452` → **EA34** (coveta-nínxol
  amb base i repisa volada, filla d'EA42).
- DW-S01-EA75 = estructura (NIX) al fons de la cova, adossada a
  DW-S01-EA20. *Pendent menor (no bloqueja):* si eixa cova és
  documentable com a EA, EA75/EA20 podrien rebre'n la paternitat.
- DW-S04-EA02↔EA21: opció (a) — tipus `Associated natural context`
  amb la cronologia dins; l'apilament, a la nota.

## E. Pendents que este delta NO obri (herència viva)

- Delta v1.2 de la convenció de shapes; validador `Y >= 1` a
  l'extractor; menú de fase 3 (Metashape).
- v26: jubilació de `Jamb_Fabric_Reveal` amb migració.
- Worklist restant (punt 3 de l'agenda de sessió): repàs de què hi
  queda viu i triatge lots-scriptats / a mà.
- R: reconstrucció de components i graf dirigit sobre `QRY_14` (la
  matèria primera ja hi és).

## F. Entrada proposada per al quadern interpretatiu (per a la teua còpia local)

**El penya-segat com a graf.** La campanya de relacions ha convertit
l'organització del front en un objecte computable: 88 arestes en sis
famílies que separen tres lectures del mateix espai — la geològica
(estrats de suport que encadenen estructures a desenes de metres), la
constructiva (juntes i superposicions, cinc d'elles ja dirigides per
estratigrafia) i la de circulació (proximitats verticals i diagonals
sense contacte, els graons perduts de l'accés aeri). Que els conjunts
siguen components derivables i no fets registrats té una conseqüència
interpretativa: les «terrasses de mausoleus» que la vista reconeix al
front no són una categoria imposada sinó un resultat — si canvien els
criteris d'aresta, canvien els conjunts, i la hipòtesi es pot
falsejar des de la mateixa base. Connecta amb H02/H03 (OE3) i amb la
lectura del suport geològic com a decisió constructiva (H04, OE4).
