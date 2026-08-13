# DELTA v17a -> v18 (definitiu)

Estat: **tancat** (tots els blocs, 2026-08-11, "tanca tot segons recomanacions").
Fonts: `Prev17-18.md` (notes d'entrada de dades) + inspeccio directa de `DB_v17a.accdb`
(35 estructures, 62 decoracions, 27 elements desapareguts, 5 connexions) + scripts v17/v17a.

Deliverables d'aquesta versio:

| Fitxer | Paper |
|---|---|
| `chachapoya_DB_v18.bas` | Construccio completa sobre BD en blanc (`BuildDB`) **i** reconstruccio nomes de consultes (`RebuildQueriesV18`) |
| `chachapoya_Form_v18_val.bas` | `F_STRUCTURES` + subformularis v18 (`BuildForm`) |
| `chachapoya_PATCH_v18.bas` | Migracio in situ v17a -> v18 amb dades preservades (`PatchV18`) |
| `DELTA_v17_v18.md` | Aquest document |
| `esquema_bbdd_estructures_v18.md` | Referencia tecnica actualitzada |
| `manual_us_bbdd_v18.md` | Manual d'us actualitzat (criteris A5, A7, C3, E2) |
| `tfm_metodologia_bbdd_v18.md` | Document metodologic actualitzat |

**Sequencia sobre la BD amb dades (la teua):**
1. `PatchV18()` (esquema + dades + lookups + relacions)
2. `RebuildQueriesV18()` de `chachapoya_DB_v18.bas` (les 35 consultes; **mai** `BuildDB()` sobre una BD amb dades: fa DROP de `T_DECORATIONS`, `T_LOST_ELEMENTS`, `T_CONNECTIONS` i `T_ARCH_FEATURES`)
3. `BuildForm()` de `chachapoya_Form_v18_val.bas`
4. Revisar `QRY_23_V18_Review` (la worklist de la migracio)

---

## BLOC 0 - Resolt sense delta: el formulari desplegat era anterior al patch 17a

*"Falta camp notes a pestanya 1"* i *"eliminar camps geometria de 4.Dec"* ja estaven resolts
a l'esquema 17a: la BD portava les columnes correctes pero el formulari en us era el pre-17a.
El `BuildForm()` del pas 3 ho tanca sense cap canvi d'esquema.

---

## BLOC A - Gating i redundancies observacionals

### A1. `Recessed_Frame` gatejat per `N_Chamber_Bodies = 0` [TANCAT]
Inaplicabilitat **derivable** (sense cos de cambra no hi ha facana, i el marc reculat
qualifica el pla de facana) -> gating amb 0 de farciment, mai `Not applicable`.
- Formulari: `EnSrc` condicionat + farciment automatic a `N_Chamber_Bodies_AfterUpdate`.
- **Regla 47**: `Recessed_Frame = 1` amb `N_Chamber_Bodies = 0`.
- `QRY_13`: exporta `IIF(N_Chamber_Bodies=0, Null, Recessed_Frame)` — substitueix el gate
  per `Sys_Portal` que la v17 arrossegava contra la seua propia decisio (fix F3).

### A2. `Portal_Orientation` tancat amb `Access_Plane = 'Facade'` [TANCAT]
Amb l'acces a la facana, les dues orientacions son **una sola dada** registrada una vegada.
- Formulari: camp desactivat mentre el pla siga `Facade`; esdeveniment nou
  `Access_Plane_AfterUpdate`.
- `QRY_07`: columna `Portal_Orientation_Effective = IIF(Access_Plane='Facade',
  Facade_Orientation, Portal_Orientation)` — R mai no fa mitjana d'un duplicat contra un buit.
- Migracio nul.la: les dues columnes eren buides a les 35 files (pendents de bruixola).

### A3. EA-TER tanca facana i portal [TANCAT]
Derivable de la tipologia: una terrassa en repisa no te facana ni portal.
- Formulari: `isTer` (per **nom** de tipologia, `Like "EA-TER*"`) desactiva
  `Facade_Orientation`, `Access_Plane`, `Portal_Orientation`, `Portal_Position` i
  `Recessed_Frame`; `ID_Typology_AfterUpdate` declara `Sys_Portal = 'Not applicable'`
  **nomes si** el valor era el `Absent` per defecte (un valor ja declarat no es toca:
  el marca la regla).
- **Regla 48**: EA-TER amb qualsevol d'eixos camps omplits.

### A4. `Sill_Coincides_Cornice` (comparticio d'element) [TANCAT]
El cas que el gradient no pot dir: la cornisa intercos **fa de llindar**. N es honestament 0
(cap llindar diferenciat) i la posicio queda resolta igualment.
- Camp nou `Sill_Coincides_Cornice` BYTE, domini 0/1/9, **default 0 de farciment**
  (inaplicabilitat derivable -> gating; expressament **fora** de `FillThreeValueFields`,
  que governa els defaults NULL i la regla 17).
- Finestra: `Sill = 0` i `Interbody_Cornice IN (1,2,3)`. El formulari nomes l'activa dins.
- **Regla 49**: `= 1` fora de la finestra.
- `QRY_23` llista els candidats de la finestra amb el 0 de migracio per confirmar.

### A5. EA-PLA-V amb N0 = 0 i N1 = 0 es legal [MANUAL, TANCAT]
Cap canvi d'esquema: una plataforma amb tots els components en 0 i el sistema present es un
resultat sobre la inversio de treball, no una incoherencia. Criteri al manual.

### A6. "Mur de retorn" unificat [TANCAT]
Nomes etiquetes UI; els noms emmagatzemats no es toquen.
- `L_STRUCT_BODY` RTW: `Name_VAL = 'Mur de retorn (V)'` (patch, **per Code**).
- Formulari: etiqueta de `Return_Wall` a 11.Sist i valor de la llista d'`Access_Plane`.

### A7. Fons de cambra = oposat al PLA DE FACANA [MANUAL, TANCAT]
El fons es defineix contra el pla de facana (el pla exposat), **mai** contra l'acces: un
acces pel mur de retorn no gira `Rear_Closure_Type` 90 graus. Criteri al manual i a la
descripcio del camp.

### A8. "Flanc de facana" es mante [TANCAT]
Sense canvi.

---

## BLOC B - Maconeria

### B1. `'Tabular blocks'` -> `'Regular tabular blocks'` [TANCAT]
Amb l'entrada de `Large blocks`, un bloc gran TAMBE es tabular: el terme nu va deixar de
nomenar una classe. **Canvi de valor emmagatzemat**, fet una sola vegada al patch
(25 files primaries + 2 secundaries al corpus) perque exportacions v17 i v18 no barregen
grafies.

### B2. `'Large blocks'` / "Blocs de grans dimensions" [TANCAT]
Nou valor a `Stone_Format` i `Stone_Format_Secondary`: una peca que fa filada per ella
mateixa, el senyal d'inversio de treball que el test funcional existeix per a llegir.

### B3. `Interbody_Cornice_Format` [TANCAT]
Camp nou TEXT(20): `Laminar slabs / Tabular blocks / Mixed / ND`. Patro `Lintel_Material`:
actiu nomes mentre I te entitat; s'esborra amb el quick-fill de `Sys_Interface`.
- **Regla 50**: format amb `I = 0` o `I = 9`.
- Exporta a `QRY_13` gatejat per `Sys_Interface` (entra al mapa `cm` al lloc que
  `Recessed_Frame` deixa lliure) i a `QRY_05`.

---

## BLOC C - Vocabulari A-X i connexions

### C1. `Platform_Surface` = element **Z** [TANCAT]
La plataforma obte la seua superficie: el paral.lel exacte de T al rafec.
- Camp nou BYTE amb domini de 5 valors, **21e element del mapa** (`FillElementMap` m(20)),
  gatejat per `Sys_Platform`; default 0.
- `L_ELEMENTS`: fila Z (25 files; **Y queda reservada** per a la banqueta, pendent dels
  tres casos documentats).
- `Platform_Surface_Material` penja ara de Z (**regla 51**: material amb Z = 0) i de la
  condicio Isolated preexistent (R7).
- Regla 3 ampliada: `Sys_Platform` present amb E, F, G **i Z** tots 0.
- Quick-fill de `Sys_Platform` inclou Z (E continua exclos, delta v13).
- Migracio: 0 sota sistema tancat, 9 sota `Not observable`, NULL on el sistema es obert
  (worklist `QRY_23`).

### C2. `Tie_Walls` (D) desgatejat de `Sys_Base` [TANCAT]
El corpus el mostra a nivell de mur, fora de la massa basal: com I i R, queda permanentment
actiu i el seu 0 es sempre una asercio. `Sys_Base` governa A, B, C.
- `FillElementMap`: m(3,2) = "" (les regles 1-2-4-6 i l'exportacio s'ajusten soles).
- Quick-fill de `Sys_Base` ja no l'escriu; `EnSrc "Tie_Walls", True`.
- `QRY_23` llista les files amb `Sys_Base` tancat i `D = 0`: eixe 0 era farciment i ara
  cal confirmar-lo com a observacio (o passar-lo a 9/1/2/3).

### C3. `Connection_Type` + `'Associated natural context'` [TANCAT]
El cas terrassa + ninxol natural al seu extrem: **dos registres** (regla de la sequencia
constructiva minima) i una aresta que diu que van junts sense afirmar mecanisme.
Etiqueta: "Context natural associat". Criteri d'us al manual.

---

## BLOC D - Connexions i elements desapareguts

### D1. La direccio ix de la parella: `Chrono_Relation` + `ID_Earlier` [TANCAT]
La v17 normalitzava l'ordre pero la lectura seguia sent posicional, i la llista entrant
havia de GIRAR la cronologia per no mentir.
- `Chrono_Relation` passa a 3 valors: `Sequential / Contemporary / Undetermined`
  (default `Undetermined` es mante).
- Camp nou `ID_Earlier` LONG, FK a `T_STRUCTURES` (relacio `REL_STR_CONE`, flag 2 com A i B),
  resolt **per codi** al formulari.
- `BeforeUpdate`: el gir A/B ja no toca res (un ID absolut no canvia); amb `Sequential`
  exigeix `ID_Earlier` dins de la parella i el tipus direccional (regla 20); amb qualsevol
  altra relacio autoneteja `ID_Earlier`.
- `F_CONN_IN` deixa d'invertir: llig `ID_Earlier` i mostra "Aquesta es anterior/posterior".
- `QRY_14`: columna `Code_Earlier` (LEFT JOIN, 3 parentesis).
- **Regla 52** (a/b): `Sequential` sense `ID_Earlier` valid; no-`Sequential` amb `ID_Earlier`.
- Migracio: `'A is earlier'`/`'B is earlier'` (i grafies v16) -> `Sequential` + FK.
  Al corpus actual: 5 files `Undetermined`, migracio nul.la.

### D2. `ID_Position` eliminat; `Evidence_Scope` amb regla dura [TANCAT]
1 valor omplit en 27 files: la columna feia una pregunta que ningu no responia
(`Element_Code` ja situa un element; una fila Body/Whole no te posicio unica per definicio).
- Patch: valor migrat a Notes ("position: <Name>", per nom via lookup), relacio
  `REL_SB_LOST` eliminada, DROP COLUMN.
- `Evidence_Scope` es mante; etiqueta "Cos constructiu" (desfa l'ambiguitat amb el cos huma).
- **Regla 53** (mirall de la 21): `Element_Code` no nul amb abast != `Element`.
  16 files del corpus hi seuen: worklist a `QRY_23` + validacio BeforeUpdate en les dues
  direccions.

---

## BLOC E - Suport i llindar analitzable

### E1. `ID_Support` filtrat per tipologia [TANCAT]
- NIX: llista reduida a `Natural niche (<1m2)` + ND, **autoompli** (relacio 1:1 derivable).
- CAV: llista reduida a les cavitats (`Like '*cavity*'`) + ND (l'eleccio mitjana/gran es
  del investigador).
- PR: llista reduida a `Rock surface` + ND, autoompli.
- Filtre i autoompli **per nom**, mai per ID. **Regla 54** vigila les tres coherencies
  (ND sempre legal: es un dubte, no una contradiccio).

### E2. Llindar minim analitzable [MANUAL, TANCAT]
Els camps secundaris son opcionals; el minim per a entrar una fila a l'analisi A-Z es:
els 21 elements + els 6 `Sys_*` sense NULL. Criteri al manual (i la regla 17 /
`QRY_16s_Observational_Nulls` es la eina de seguiment).

---

## BLOC F - Fixes propis detectats en la revisio

- **F1.** `QRY_18_Notes_Review` llegia `E.Notes`, renomenat `Doc_Notes` en 17a: la consulta
  demanava un parametre en obrir-se. Reescrita amb els **14 camps de notes** en ordre de
  pestanya (es la consulta que converteix text lliure recurrent en camps candidats).
- **F2.** `QRY_13` no exportava `Sys_Interface`: I es gatejava per un sistema el estat del
  qual era invisible a la matriu. `sy(5)` afegit (`SYS_INTERFACE`).
- **F3.** `QRY_13` gatejava `Recessed_Frame` per `Sys_Portal` contra la decisio v17 (el marc
  va eixir del portal en v17). Ara es gateja per `N_Chamber_Bodies` (A1).
- **F4.** `QRY_23_V18_Review` nou: worklist de la migracio (Z per observar, farciments de D,
  candidats cornisa-llindar, incoherencies d'abast, `Sequential` sense anterior).

---

## Comptadors v18

| Cosa | v17a | v18 |
|---|---|---|
| Camps de `T_STRUCTURES` | 141 | **144** (+`Platform_Surface`, +`Interbody_Cornice_Format`, +`Sill_Coincides_Cornice`) |
| Camps de `T_CONNECTIONS` | 8 | **9** (+`ID_Earlier`) |
| Camps de `T_LOST_ELEMENTS` | 7 | **6** (-`ID_Position`) |
| Elements del domini de 5 valors | 20 (A-X) | **21 (A-X + Z)** |
| Files de `L_ELEMENTS` | 24 | **25** (Y reservada) |
| Relacions | 25 | **25** (-`REL_SB_LOST`, +`REL_STR_CONE`) |
| Consultes | 33 | **35** (+`QRY_16f`, +`QRY_23`) |
| Regles actives | 45 (numerades fins a 46) | **53 (numerades fins a 54; la 41 retirada)** |

## Pendent (sense canvis)

- Descomposicio de `T_BODIES` (el comptador `Fabric` continua sent la solucio interina).
- Element **Y** (banqueta): reservat, pendent de tres casos documentats.
- Caracteritzacio de superficies ROC: ajornada al creixement del corpus.
- Revisio manual de `RA Perimeter band` (`QRY_22_V17a_Review`).
- Camp de `Facade_Orientation`, coordenades i metrica: pendents de camp (bruixola/GNSS),
  no de la BD.
- **Candidat v19** (cas DW-S01-EA22a, 2026-08-13): qualificador `Jamb_Fabric_Reveal` —
  posicio del brancal resolta per la fabrica mateixa (cara terminal acabada i deliberada,
  *masonry reveal*), patro A4: domini 0/1/9, farciment 0, regla de finestra.
  - **Finestra `Jambs IN (0, 2)`**, no nomes 0: amb O = 2 (brancal a un costat i cara
    terminal a l'altre) el camp ha de viure, i eixe es el cas que resol l'ambiguitat que
    el manual declarava indecidible al Nus 1 ("asimetria de disseny o brancal perdut?").
    Amb O = 1 queda tancat: cap posicio per resoldre d'una altra manera.
  - **Exempcio a R3** per la via de l'exempcio de `Timber_Brackets` a R1: `AND
    Jamb_Fabric_Reveal <> 1`. Sense aixo, el cas legitim del manual (Nus 2: zeros de
    portal com a resultat) continua sent indistingible d'un registre a mig entrar.
    NOTA: la plataforma te el mateix forat obert (delta A5, EA-PLA-V amb tot 0 es legal
    i R3 hi dispara igual) i NO queda resolt per aquest punt.
  - **`DOM3Q`, etiquetes de qualificador** (acordat 2026-08-13): els qualificadors de
    comparticio d'element porten llista propia `No / Si / No observable` en lloc de la
    `DOM3` compartida (`Absent / Present / No observable`). Valors emmagatzemats
    IDENTICS (0/1/9): canvien nomes les etiquetes UI, i per tant regles, exportacions i
    migracio queden intactes. Motiu: la casella respon una pregunta, no declara la
    presencia d'una cosa - "Cornisa fa de llindar: Absent" no diu el que passa.
    S'aplica als DOS camps alhora: `Sill_Coincides_Cornice` (v18) i el nou. Cal escriure
    al manual que el projecte passa a tenir dos vocabularis d'etiqueta per a un mateix
    domini, o semblara arbitrari.
  - Mentrestant: `Sys_Portal = Present partial` quan sobreviu el buit i alguna vora,
    N = O = Q = 0, i formula constant a `Systems_Notes`: `jamb: masonry reveal, dressed`.
    S'acumulen casos abans d'obrir el delta (criteri dels tres casos, com la Y);
    `QRY_18_Notes_Review` els recuperara tots d'una consulta.
- **Candidat v19 — bloc `Timber_Bracket_Role`** (cas DW-S01-EA07a + 6 mes, R9 en bateria,
  2026-08-13). El camp respon UNA pregunta: la mensula formava part del sistema
  plataforma? No es la foto ("la trobe sola") ni la funcio d'us (aixo va a
  `T_ARCH_FEATURES` / notes): es el vincle amb H, agregat a nivell d'estructura.
  Quatre punts tancats:
  - **R9 accepta `ND`**: dispara nomes amb `Is Null`. `ND` es un judici ("avaluat, no
    decidible"), no una casella buida - la distincio NULL/9/0 traslladada al camp TEXT.
    Amb la condicio actual la branca no es pot buidar mai (7 de 35 files) i la bateria
    acumula soroll permanent.
  - **Etiquetes UI noves** (patro A6, valors emmagatzemats `Platform support / Isolated /
    Both / ND` INTACTES - R7, exportacions i migracio no es toquen):
    `Component de plataforma (H)` / `Sense vincle amb plataforma` / `De les dues
    classes` / `Vincle indeterminat`. Motiu: les etiquetes velles no deixaven deduir la
    pregunta del camp ("Aillada" es llegia com a descripcio de la troballa, "Tipus
    indeterminat" suggeria una tipologia de mensules inexistent, "Ambdos" no deia
    ambdos QUE - i `Both` ja es va retirar una vegada, v13, pel mateix motiu). L'eix
    ("vincle") apareix a dues etiquetes com a ancoratge deliberat.
  - **Regla nova de coherencia**: rol `Platform support` o `Both` amb `Sys_Platform IN
    ('Absent','Not applicable')` es contradiccio - si les mensules sostenien
    plataforma, el sistema es present o atestat. (La direccio inversa ja la cobreix R7.)
  - **Criteris al manual** (nus nou o ampliacio del nus 1): arbre de tres preguntes.
    (1) L'evidencia soste l'atestacio? - diverses mensules alineades al mateix nivell,
    encaixos buits, cicatrius: `Sys_Platform = Attested lost` + rol `Platform support` +
    fila a 12.Extra amb el tipus "mensules al buit" (que EXISTEIX al cataleg tancat
    precisament per a aixo). (2) Es pot afirmar que MAI no va sostenir plataforma? -
    context llegible al seu nivell i net: rol `Isolated`. Exigeix haver pogut llegir el
    context, com el 0 exigeix haver mirat. (3) Ni una cosa ni l'altra: rol `ND`,
    resposta completa i no pendent; sospita redactada a `Systems_Notes`. Frontera
    `Isolated`/`ND` = frontera 0/9. `Both` = l'estructura te mensules de les dues
    classes (p.ex. tres alineades sota plataforma + una desvinculada a un altre
    nivell); si el corpus n'acumula, la solucio de fons seria baixar el rol a nivell de
    mensula (mateix debat que `T_BODIES`), ajornat pel mateix criteri.
  - Nota: `Timber_Bracket_Count` compta el que SOBREVIU; el rol interpreta el que ERA.
    Count = 1 amb rol `Platform support` (cas 1) no es contradiccio.
- **Candidat v19 — bloc muret transversal (D) / MEN** (casos DW-S01-EA38/39/40b + EA40,
  R11 en bateria, 2026-08-13). R11 va disparar BE: un muret transversal comptat com a
  cos de cambra es una classificacio forcada, no un defecte de regla. Els registres es
  corrigen JA amb l'esquema actual: Muret transversal (D) present, N_Basal_Bodies = 0,
  N_Chamber_Bodies = 0 (aquests son exactament els casos que van motivar el delta C2).
  - **DW-S04-EA15 es un cas DIFERENT** (esmena 2026-08-13: mal agrupat inicialment amb
    els de dalt): mausoleu amb cossos de cambra REALS, mur d'acces (mur de retorn
    dret) esfondrat, cap vestigi d'element de portal. Els comptadors ES QUEDEN; el
    problema era Sys_Portal = Absent quan tocava **Attested lost** (cambra implica
    acces: ho diu la definicio mateixa de N1). Registre correcte, RESOLT amb l'esquema
    actual sense cap canvi: Sys_Portal = Attested lost; N/O/Q = 9 (el mur no hi es: no
    es pot examinar si hi havia peces diferenciades; ni 3 - caldria evidencia de peca -
    ni 0 - afirmaria que no n'hi havia); Return_Wall = 2 si l'esquerre sobreviu; UNA
    fila a T_LOST_ELEMENTS amb abast **Body (Cos constructiu)**, evidencia Murs
    truncats, codi buit (regla 53), i a Notes la formula `position: right return wall
    (presumed)` + el raonament per eliminacio i analogia; Access_Plane = Return wall
    (precedent v16: el deduit amb base es registra, el raonament a notes). MECANISME
    CLAU documentat: la cobertura de la regla 2 per a sistemes en Attested lost es fa
    per files d'abast Body/Whole (QRY_16s_Lost_Cover cobreix tot el vocabulari de
    l'estructura, sistemes inclosos) - el combo de codis exclou les lletres de sistema
    A PROPOSIT (Is_System=False). Cal una linia al MANUAL que ho diga: "per a atestar
    un sistema perdut, fila d'abast Cos constructiu o Estructura sencera; el codi
    d'element es per a peces concretes". R11 nomes dispara amb Absent i R3 nomes amb
    Present*: cap de les dues toca el registre corregit.
  Per a v19:
  - **MEN promogut**: Name emmagatzemat `MEN Isolated Bracket` -> `MEN Isolated
    Structural Element` (la descripcio EN ja ho deia; cap regla ni consulta filtra per
    MEN, es un UPDATE tipus B1). Name_VAL -> `MEN Element estructural aillat`.
    Descripcio nova ANCORADA AL VOCABULARI: "registre l'evidencia del qual es redueix a
    un o pocs elements A-Z sense estructura classificable (mensula, muret transversal,
    pilastra); la interpretacio (circulacio, estructura previa, suport) va a
    T_ARCH_FEATURES / Notes, mai a la tipologia; el nom historic ve del primer cas
    documentat, la mensula". SENSE subtipus: la identitat de l'element la diuen els
    camps de 11.Sist (D present = muret, E present = mensula) - subtipus duplicarien
    informacio i caldria una regla per vigilar la divergencia. La col.lisio del terme
    "aillat" amb el rol de mensules desapareix amb el reetiquetatge del rol (bloc
    anterior): al nivell tipologic la lectura descriptiva es la correcta.
  - **Manual, dues linies MEN**: (1) com es registra (element a 11.Sist, sistemes
    tancats, comptadors a 0 NECESSARIAMENT - sense cossos no hi ha nivell del qual
    parlar); (2) frontera amb l'estructura atestada: si l'evidencia soste una
    estructura concreta (mensules alineades amb encaixos = plataforma perduda), el
    registre NO es MEN sino l'estructura amb el sistema Attested lost + fila
    d'evidencia. MEN = quan no es pot afirmar cap estructura.
  - **L_ELEMENTS, fila D**: Name_EN `Tie walls` -> `Transverse wall` ("tie" AFIRMA la
    funcio de trava, i la funcio es precisament el que no sabem - mateix criteri que
    va triar "diedre" sobre el terme genetic). Name_VAL `Muret transversal` es queda.
    Level_Type ES QUEDA a N0 (posicio de la lletra en l'ordenacio, bloc A-H): NO
    canviar a N0-N1, que per a I significa "a la junta" i per a D significaria "a
    qualsevol dels dos" - un marcador, dos significats. La llibertat d'ancoratge va a
    la Description dita amb ALCADA, mai amb NIVELL ("may anchor at any height of the
    fabric"): "nivell" es paraula reservada (N0/N1 es defineixen pels cossos, i un
    muret no es cos de res).
  - **Justificacio C2 reescrita** (esquema i capcaleres): no "el corpus el mostra a
    nivell de mur" (llegible com "formant part dels murs de la cambra", FALS - cap D
    del corpus tanca cambra) sino "apareix desvinculat de la massa basal: sobre
    plataformes volades, com a suport adossat, o aillat".
  - **Formulari**: capcalera propia per a D just davall del conjunt basal, amb el
    DISCRIMINADOR i no la geometria (l'orientacio no distingeix D de V: amb facana
    paral.lela al farallo son geometricament identics): "MURET TRANSVERSAL (D) - NO
    TANCA CAMBRA (AIXO ES V) NI ES COMPTA COM A COS; SEMPRE ACTIU". La basal torna a
    "CONJUNT BASAL: A B C".
  - **Manual, test de delimitacio D/V**: "Muret transversal (D): fabrica perpendicular
    al farallo que no tanca cap interior ni compta com a cos. Si tanca cambra, es mur
    de retorn (V). Si soste plataforma pel davall com a peca encastada, mireu mensules
    (E) / bigues (F). Si esta sol, el registre es MEN." Implicacio de comptadors: V
    implica cos (paret de N1: cambra amb acces present o atestat, R11 vigilant); D no
    n'implica cap.
  - **Revisio unica del corpus**: rellegir amb el test de delimitacio les files amb
    Tie_Walls IN (1,2,3) i les de Return_Wall IN (1,2,3) - Esteve admet possibles
    confusions D/V entrades.
  - **EA40** (pilar adossat a la roca que soste una jacena, que soste les bigues del
    sostre): es element DINS d'estructura, no MEN. D present + formula constant a
    Systems_Notes: `D: beam pier` - UNA sola formula, sense destinacio funcional (la
    interpretacio no entra ni per la formula; que la jacena hi descansa es observacio).
    El sostre: X present amb tipus Built timber and slabs (o Mixed). No es K
    (Structural_Pilasters: integrades al pla del parament, d'altura completa - aquest
    esta adossat a la ROCA). Nota TFM: versio litica i dreta del que les mensules (E)
    fan encastades, aplicada al sostre - material H01.
  - NO es toca: el nom de camp `Tie_Walls` (renom de camp = consultes trencades,
    llico v11).
