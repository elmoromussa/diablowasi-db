# Especificacio de canvis v10 -> v11
## Base de dades `chachapoya_DB` — TFM Esteve Ribera Torro

**Data:** 2026-08-05 (rev. 3 — totes les decisions tancades)
**Base de partida:** `chachapoya_DB_v10.bas` + `chachapoya_Form_v10_val.bas`
**Estat de les dades:** 35 registres a T_STRUCTURES (tots Diablo Wasi, sectors S01–S04), 29 a T_DECORATIONS, resta de taules buides
**Objectiu:** implementar els canvis d'esquema acordats i generar `chachapoya_DB_v11.bas` i `chachapoya_Form_v11_val.bas`

---

## 0. Convencions que es mantenen sense canvis

- Valencia nomes com a etiquetes d'interficie (captions de formulari, noms de pestanya). Tots els valors emmagatzemats, noms de camp, esquema i contingut de lookups en angles.
- Codificacio CP1252, cap caracter no-ASCII al codi VBA, finals de linia CRLF.
- Cap continuacio de linia (`& _`); patro `sql = sql & "..."` en linies separades.
- Tots els Sub auxiliars `Private`; nomes `BuildDB` i `BuildForm` publics.
- `DROP TABLE` abans de `CREATE TABLE`.
- Captions de columna via DAO `TableDef.Fields.Properties("Caption")`.

---

## 1. Domini dels camps d'element arquitectonic

### 1.1 Canvi

Substitueix el domini BYTE 0/1/9 per un gradient ordinal de cinc valors, aplicat **exclusivament als camps d'element del vocabulari A–X** (veure 1.6 per a la resta).

| Valor | Caption (valencia) | Etiqueta EN | Criteri operatiu |
|---|---|---|---|
| 0 | Absent | `Absent` | La posicio es examinable i no hi ha ni element ni traca. **VALOR PER DEFECTE** |
| 1 | Present, complet | `Present complete` | Es conserva **75% o mes** de l'element |
| 2 | Present, parcial | `Present partial` | Es conserva **menys del 75%**, pero l'element es identificable in situ |
| 3 | Desaparegut | `Attested lost` | No es conserva res in situ; nomes traca fisica directa. **Exigeix registre a T_LOST_ELEMENTS** |
| 9 | No observable | `Not observable` | La posicio no es examinable: destruida o amagada |

### 1.2 Regla semantica

Cada camp d'element respon una sola pregunta: **«va tenir l'estructura aquest element, i en quin estat esta ara?»**
El camp descriu l'element, mai la relacio de l'observador amb ell.

- `9` NO significa "no observat". Significa que la posicio no es examinable.
- Si s'ha mirat la posicio i no hi ha res, el valor es `0`.
- Agregacions analitiques: «que es va construir» = 1+2+3 · «que es conserva» = 1+2 · «integre» = 1.
  Denominador de cada camp = registres amb valor <> 9.

### 1.3 Criteri operatiu del llindar 75% (per al document metodologic)

> Un element es registra com a **Present complet** quan se'n conserva el 75% o mes de la seua extensio original; com a **Present parcial** quan se'n conserva menys del 75% pero l'element continua sent identificable in situ.
>
> En **elements comptables** (mensules, bigues transversals, cantoneres, brancals) la proporcio es calcula sobre el nombre d'unitats originalment previstes, deduit de la simetria o de les traces conservades. Exemple: 5 mensules conservades de 7 encaixos documentats = 71% = parcial.
>
> En **elements continus** (paraments, cornises, socols, superficies de rafec) la proporcio s'estima sobre l'extensio lineal o superficial original, reconstruida a partir de l'amplaria de l'estructura i de les traces d'arrencament.
>
> Quan la proporcio no siga estimable amb un marge raonable, s'aplica el criteri subsidiari de **llegibilitat morfologica**: si la forma original no es reconstruible sense inferencia, es registra com a parcial.

Nota: el llindar del 75% es conservador — un element al qual falta una quarta part ja es «parcial». `Present complete` designa per tant elements quasi integres.

### 1.4 Etiqueta ND

Elimina `ND` com a etiqueta als camps d'element; se substitueix per `No observable` / `Not observable`.

Als lookups de tipus (`Lintel_Material`, `Chamber_Roof_Type`, `Mortar_Type`, `Rear_Closure_Type`, `Platform_Surface_Material`, `Masonry_Quality`, `Plaster_*`, `Pigment_*`) `ND` significa una altra cosa — l'element hi es pero no se'n determina el tipus. Reanomena-hi l'etiqueta a `Type undetermined` (caption: *Tipus indeterminat*). No canvies els valors ja introduits: nomes l'etiqueta del lookup.

### 1.5 Camps amb domini de 5 valors (20 elements)

| Codi | Camp v11 | Camp v10 |
|---|---|---|
| A | `Embedded_Base_Beams` | igual |
| B | `Base_Level` | igual |
| C | `Decorative_Socle` | igual |
| D | `Tie_Walls` | igual |
| E | `Timber_Brackets` | igual |
| F | `Transverse_Beams` | igual |
| G | `Corbelled_Courses` | igual |
| I | `Interbody_Cornice` | igual |
| J | `Corner_Quoins` | igual |
| K | `Structural_Pilasters` | igual |
| L | `Facade_Flank` | `Lateral_Wall_Faces` |
| M | `Relief_Frieze` | igual |
| N | `Sill` | igual |
| O | `Jambs` | igual |
| Q | `Lintel` | `Lintel` (era lookup de material — veure 5.4) |
| R | `Upper_Crown` | igual |
| S | `Eave_Beam` | igual |
| T | `Eave_Surface` | igual |
| V | `Return_Wall` | `Lateral_Walls` |
| X | `Chamber_Roof` | igual |

Els elements **H, P i U no figuren aci**: son sistemes i passen als camps `Sys_*` (veure seccio 4). L'element **W (`Rear_Wall`) s'elimina** (veure 5.5).

### 1.6 Camps observacionals amb domini 0/1/9

Conserven tres valors, amb l'unica modificacio de reanomenar l'etiqueta de `9` a `Not observable` (caption *No observable*):

- Atribut de portal: `Recessed_Frame` (configuracio, no element: no admet gradient)
- Conservacio: `Looting`, `Fire_Damage`, `Animal_Activity`, `Modern_Access`
- Bioarqueologia: `Human_Remains`, `Anatomical_Connection`, `Mummification`, `Funerary_Bundles`, `Dispersed_Remains`, `Flexed_Position`, `Bone_Burning`
- Materials: `Mat_Textiles`, `Mat_Wood`, `Mat_VegFiber`, `Mat_Ceramics`, `Mat_Fauna`, `Mat_DeerAntler`, `Mat_Other`
- Suport i maconeria: `Support_Modified`, `Mortar_Present`, `Chinking_Stones`
- Acabats: `Plaster_Present`, `Pigment_Present`

Motiu: «present parcial» i «desaparegut» son categories morfologiques d'element construit. Aplicades a processos (`Looting`) o a vestigis mobles (`Mat_Textiles`) forcen una semantica que no correspon.

### 1.7 Implementacio tecnica

- `DEFAULT 0` a totes les BYTE observacionals del DDL (tant les de 5 valors com les de 3).
- Els controls combo del formulari han d'obrir amb `0` seleccionat, mai en blanc.
- **Cap camp observacional pot quedar NULL.** El NULL ja no es un estat valid del domini.

---

## 2. Taula nova: `T_LOST_ELEMENTS`

Justificacio: el valor `3` es l'unic del domini que es una inferencia i no una observacio directa. Sense registre de l'evidencia no es replicable. Es a mes la base evidencial de l'OE3 (reconstruccio de la infraestructura de circulacio aeria).

```
T_LOST_ELEMENTS
  ID                 AUTOINCREMENT  PRIMARY KEY
  ID_Structure       LONG    FK -> T_STRUCTURES.ID    NOT NULL
  Element_Code       TEXT(2) FK -> L_ELEMENTS.Code    NOT NULL
  ID_Evidence_Type   LONG    FK -> L_LOST_EVIDENCE.ID NOT NULL
  Evidence_Scope     TEXT(20)   Element / Body / Whole structure
  ID_Position        LONG    FK -> L_STRUCT_BODY.ID   (opcional)
  Notes              MEMO
```

### Lookup `L_LOST_EVIDENCE`

| ID | Name | Description |
|---|---|---|
| 1 | Negative socket / impression | Encaix buit al parament |
| 2 | Beam hole | Forat de biga passant |
| 3 | Break scar | Cicatriu d'arrencament |
| 4 | Detached fragment in situ | Fragment caigut identificable al peu |
| 5 | Mortar imprint | Empremta de morter sense element |
| 6 | Corbels into void | Mensules que no sostenen res |
| 7 | Pigment on bedrock | Pigment sobre suport ara nu |
| 8 | Truncated walls | Murs escapcats en pla net |
| 9 | Other (see Notes) | |

### Regla d'abast del valor 3

S'admet `3` a qualsevol element **si i nomes si** es pot assenyalar un valor de `L_LOST_EVIDENCE` aplicable. Si no n'hi ha cap, el valor correcte es `0` o `9`. El lookup fa de filtre metodologic; no cal una llista tancada d'elements.

Els tres camps de sistema (`Sys_Platform`, `Sys_Portal`, `Sys_Eave`) admeten igualment `Attested lost` i queden subjectes a la mateixa regla.

### Migracio de `Lost_Body_Evidence`

Queda cobert per aquesta taula amb `Evidence_Scope = Body`. Migrar manualment els 25 valors i **eliminar el camp** un cop verificada la migracio.

---

## 3. Lookup nou: `L_ELEMENTS`

El vocabulari A–X existeix actualment nomes com a noms de camp; no es referenciable des d'altres taules ni exportable com a entitat. `L_STRUCT_BODY` son **posicions** (11 entrades), no elements.

```
L_ELEMENTS
  ID            AUTOINCREMENT PRIMARY KEY
  Code          TEXT(2)   A .. X
  Name_EN       TEXT(60)
  Name_VAL      TEXT(60)
  Level_Type    TEXT(10)  N0 / N0-N1 / N1 / SUP
  System        TEXT(20)  Platform / Portal / Eave / (buit si no pertany a cap sistema)
  Is_System     YESNO     TRUE per a H, P, U
  Field_Name    TEXT(40)  nom del camp corresponent a T_STRUCTURES
  Description   MEMO
```

Poblar amb els 24 elements A–X segons el vocabulari normalitzat vigent (ordre bottom-to-top), inclosos H, P i U amb `Is_System = TRUE` i `Field_Name` apuntant al camp `Sys_*` corresponent. W s'hi mante com a entrada de vocabulari encara que ja no tinga camp propi (`Field_Name` buit).

---

## 4. Sistemes constructius i gating

### 4.1 Els tres elements-sistema

El vocabulari defineix tres elements que son en realitat sistemes compositius. En v11 deixen de tenir camp propi i **son** els camps de sistema:

| Element | Camp v11 | Components | Camp v10 eliminat |
|---|---|---|---|
| H Plataforma | `Sys_Platform` | E + F + G | `Corbelled_Platform` |
| P Portal | `Sys_Portal` | N + O + Q | `Access_Opening` |
| U Rafec | `Sys_Eave` | S + T | `Eave` |

Domini d'aquests tres (sis valors):
`Present complete` / `Present partial` / `Attested lost` / `Absent` / `Not applicable` / `Not observable`
Captions: *Present complet* / *Present parcial* / *Desaparegut* / *Absent* / *No aplicable* / *No observable*

D'aquesta manera no es perd el gradient de conservacio en fusionar element i sistema.

### 4.2 Els dos camps d'agrupacio

```
Sys_Base      elements A B C D
Sys_Chamber   elements J K L M V X
```

No corresponen a cap element del vocabulari: son agrupacions per a gating. Domini de quatre valors:
`Present` / `Absent` / `Not applicable` / `Not observable`

### 4.3 Semantica analitica (critica)

| Valor del sistema | Com surten els seus camps components a l'exportacio R |
|---|---|
| `Present complete` / `Present partial` / `Present` | valors reals (0/1/2/3/9) |
| `Attested lost` | valors reals; s'espera `3` o `9` als components |
| `Absent` | `0` — l'absencia es informativa (un mausoleu sense rafec es un resultat) |
| `Not applicable` | **`NA`** — no entra a cap denominador |
| `Not observable` | `NA` |

Aquesta distincio entre *no aplicable* i *absent verificat* es la mateixa logica del 0/9 aplicada al nivell de sistema. **Sense ella, els percentatges publicats son incorrectes.**

### 4.4 Gating de nivell 1: `Record_Class` -> pestanyes

`Record_Class` (derivat de `L_TYPOLOGY`, veure seccio 6) determina quines pestanyes del formulari s'activen:

| Pestanya | Built funerary structure | Natural funerary context | Structural trace | Rock art panel |
|---|---|---|---|---|
| 1.Id | si | si | si | si |
| 2.Arq (morfologia) | si | **no** | **no** | **no** |
| 4.Dec | si | si | si | si |
| 5.Bio | si | si | **no** | **no** |
| 6.Mat | si | si | **no** | **no** |
| 7.Cron | si | si | si | si |
| 8.Cons | si | si | si | si |
| 9.Metr. | si | si | si | si |
| 11.Sist | si | parcial | parcial | **no** |
| 12.Doc | si | si | si | si |

«Parcial» vol dir que els cinc `Sys_*` es resolen igualment, pero la majoria eixiran `Not applicable`.

### 4.5 Gating de nivell 2: `Sys_*` -> grups d'elements

A la pestanya `11.Sist`, els cinc combos `Sys_*` van dalt de tot. Els grups de camps components queden desactivats (`Enabled = False`) si el sistema corresponent no es `Present*` o `Attested lost`. Cal codi al `Current` i als `AfterUpdate` dels cinc combos.

Camps subordinats que es tanquen amb `Sys_Platform`: `Timber_Brackets`, `Timber_Bracket_Count`, `Timber_Bracket_Role`, `Transverse_Beams`, `Corbelled_Courses`, `Platform_Surface_Material`, `Platform_Function`.

### 4.6 Gating de nivell 2 fora de 11.Sist

**Pestanya 6.Mat.** No cal camp nou: `ID_Material_Status` ja fa de porta.

| `ID_Material_Status` | Comportament dels camps `Mat_*` |
|---|---|
| `Absent` | tots a `0`, desactivats |
| `ND` | tots a `9`, desactivats |
| `Good` / `Fair` / `Poor` | actius, entrada detallada |

**Pestanya 5.Bio.** `Human_Remains` fa de porta per a `MNI`, `Anatomical_Connection`, `Mummification`, `Funerary_Bundles`, `Dispersed_Remains`, `Flexed_Position`, `Bone_Burning`.

### 4.7 Actualitzacio de QRY_13

L'exportacio ha d'emetre `NULL` (-> `NA` a R) per als camps de sistemes marcats `Not applicable` o `Not observable`, no el valor cru.

---

## 5. Vocabulari: reanomenaments i desdoblaments

### 5.1 Murs laterals vs paraments laterals

Els noms v10 son ambigus: tots dos contenien «lateral» referit a coses distintes. Les dades confirmen que son variables independents (les cambres tendeixen a L=1 amb V=0; els mausoleus a V=1).

| Element | Nom v11 | Caption | Definicio |
|---|---|---|---|
| V | `Return_Wall` | Mur de retorn | Mur perpendicular al pla de facana, que retorna cap a la penya |
| L | `Facade_Flank` | Ala de facana | Parament en el pla de facana, flanquejant l'obertura |

### 5.2 Portal i marc reculat

| Nom v10 | Nom v11 | Caption |
|---|---|---|
| `Access_Opening` | -> `Sys_Portal` (seccio 4.1) | Sistema portal |
| `Recessed_Portal` | `Recessed_Frame` | Marc reculat |

`Recessed_Frame` es un atribut de configuracio del portal, no un element: domini 0/1/9 (veure 1.6).

### 5.3 Separacio forma / funcio a la plataforma

`Corbelled_Platform` tenia caption *Plataforma d'acces*, barrejant morfologia verificable amb interpretacio funcional. L'OE3 existeix precisament per a determinar la funcio de circulacio; anomenar-la «acces» al camp de dades preassumeix la conclusio.

- `Sys_Platform` caption: *Plataforma en voladis* (nomes morfologia)
- Camp nou:

```
Platform_Function   TEXT(20)
   Access          acces a una estructura concreta
   Circulation     transit entre estructures
   Construction    bastida constructiva sense us posterior
   Support         base d'una estructura superior
   Multiple        mes d'una funcio documentada
   Undetermined    [DEFECTE]
```

Precedent intern: `Timber_Bracket_Role` ja separa forma i funcio per a l'element E.

### 5.4 Dintell (element Q)

`Lintel` en v10 es un lookup de material (`Stone` 10 · `Absent` 6 · `Wood` 3 · `ND` 2) que barreja presencia i tipus. Desdoblar seguint el patro ja emprat a `Chamber_Roof` / `Chamber_Roof_Type`:

```
Lintel            domini de 5 valors (element Q, component de Sys_Portal)
Lintel_Material   Stone / Wood / Mixed / Type undetermined
```

**Migracio dels 21 valors existents:** `Stone` -> `Lintel=1`, `Lintel_Material=Stone` · `Wood` -> `Lintel=1`, `Material=Wood` · `Absent` -> `Lintel=0`, `Material` buit · `ND` -> `Lintel=9`. La conversio de `Stone`/`Wood` a `1` es provisional: la revisio complet/parcial (seccio 16) ha de confirmar si algun es `2`.

### 5.5 Mur posterior (element W)

Els camps v10 estan mal utilitzats: hi ha `Rear_Wall=0` amb `Rear_Wall_Type='Built masonry'` a EA01 i EA06, que es contradictori.

- **`Rear_Wall` s'elimina** — redundant amb `Sys_Chamber`.
- **`Rear_Wall_Type` es conserva**, reanomenat a `Rear_Closure_Type`:

```
Rear_Closure_Type   Natural bedrock / Built masonry / Mixed / Type undetermined
```

Motiu de conservar-lo: distingeix 5 casos de roca natural i 3 de maconeria construida. Una cambra que aprofita la penya com a tancament posterior no s'ha construit igual que una que el basteix — distincio directament rellevant per a H01. `N_Built_Walls` no ho recupera (tambe es inconsistent en aquests casos).

### 5.6 Cornisa intercos: material CONSERVAT

`Interbody_Cornice_Material` **no s'elimina**. Encara que 19 de 20 casos son `Stone slabs`, `DW-S04-EA12a` te `Wooden beams`. Una cornisa de fusta enmig de dinou de pedra es exactament el tipus d'anomalia constructiva que interessa a H01.

**Verificacio previa:** confirmar al model 3D que EA12a no es un error d'entrada. Si ho fora, el camp seguiria justificat per a campanyes futures.

### 5.7 Ampliacio de `L_SUPPORT`

Afegir entre repisa estreta i fissura:

```
Micro-ledge (<50cm)     caption: Microrepisa (<50cm)
```

### 5.8 Ampliacio de `Pigment_Extent`

Afegir valor **`Perimeter / threshold`** (caption: *Perimetral / llindar*).

Motiu: hi ha pigment perimetral en dos contextos estructuralment distints (EA11, estructura arrasada; EA09, cavitat natural). Combinat amb `Pigment_Substrate = Bedrock` dona una consulta directa sobre una practica de marcatge del llindar independent del tipus de suport — rellevant per a la dimensio simbolica del marc teoric.

---

## 6. Tipologia: classe de registre, ND i MIX

### 6.1 Columna de classe a `L_TYPOLOGY`

Camp nou `Record_Class TEXT(30)` a `L_TYPOLOGY` (no a T_STRUCTURES: es derivable de la tipologia).

| Tipologies | Record_Class |
|---|---|
| MAU, CAM, PLA-R, PLA-V | `Built funerary structure` |
| NIX, CAV | `Natural funerary context` |
| MEN | `Structural trace` |
| PR | `Rock art panel` |

Totes les consultes analitiques han d'incorporar filtre per `Record_Class`. Motiu: dels 35 registres actuals, 5 no son estructures; qualsevol percentatge calculat sobre 35 es incorrecte.

### 6.2 Desdoblament de ND

`ND Undetermined` (ID 10) cobreix dos estats distints. Substituir per:

- `Unclassifiable` — evidencia insuficient per a classificar. **Es un resultat.**
- `Not yet classified` — estat de treball, pendent de revisio.

### 6.3 Eliminacio de MIX

`MIX Mixed` (ID 9) s'elimina. Motiu: tota EA-CAM es mixta per definicio (cavitat natural + facana construida), de manera que la categoria no distingeix res. La combinacio natural/construit ja la capturen `ID_Support` i els camps `Sys_*`.

Equivalent derivat: `ID_Typology IN (NIX, CAV) AND algun Sys_* = 'Present*'`.

---

## 7. Decoracio: registre unic a T_DECORATIONS

### 7.1 Problema

Els booleans `Dec_*` de T_STRUCTURES dupliquen la informacio de T_DECORATIONS, que si registra posicio, tipus, color i substrat. Amb nomes 35 registres ja divergeixen en **6 casos** (17%).

### 7.2 Canvi

Elimina de T_STRUCTURES:

```
Dec_Square_Niche   Dec_Relief_T      Dec_Relief_T_Inv   Dec_Relief_L
Dec_Relief_L_Inv   Dec_Zigzag        Dec_Stepped        Rock_Art
RA_Anthropomorphic RA_Zoomorphic     RA_Geometric       RA_Abstract
```

`Rock_Art` i els `RA_*` queden coberts per `Pigment_Present` + `Pigment_Substrate = Bedrock` + la fila corresponent de T_DECORATIONS. Ara mateix es informacio triplicada.

### 7.3 Que NO s'elimina

**`Relief_Frieze` (M) es un element arquitectonic, no una decoracio.** Es mante a 11.Sist dins de `Sys_Chamber`, amb el domini de cinc valors. El *motiu* que hi va representat es registra a T_DECORATIONS.

### 7.4 Formulari

La pestanya `4.Dec` passa a contenir **unicament un subformulari sobre T_DECORATIONS**. Desapareix el formulari rapid de booleans.

### 7.5 Migracio previa obligatoria

Abans d'eliminar res, comprova que cada boolea a `1` te fila corresponent a T_DECORATIONS. Casos coneguts sense fila:
`DW-S01-EA36`, `DW-S03-EA01`, `DW-S04-EA05`, `DW-S01-EA39`, `DW-S01-EA06`, `DW-S01-EA09`.

---

## 8. Camps a mantenir, eliminar o revisar

| Camp | Estat als 35 registres | Accio |
|---|---|---|
| `ID_Parent` | 35/35 NULL | **MANTENIR.** Pendent de passades posteriors. El formulari NO n'ha de forcar l'us |
| `ID_Group` | 35/35 NULL | **MANTENIR.** Idem |
| `Lost_Body_Evidence` | 25 valors | Migrar a T_LOST_ELEMENTS i eliminar (seccio 2) |
| `C14` | 35/35 False | Mantenir (esperable; encara no hi ha datacions) |
| `ChaXR_Documented` | 35/35 False | Mantenir |
| `ID_Campaign` | 35/35 NULL | El formulari l'ha d'omplir per defecte amb la campanya activa — **valor pendent de confirmar** |
| `Support_Morphology` | ja eliminat en v10 | — |

**Quasi-constants que NO s'eliminen** (son resultats, no defectes): `Interbody_Cornice` (21 presents / 1 absent), `Base_Level` (27/3).

**Nota sobre `ID_Parent` i `ID_Group`:** com que resten buits i sense validacio obligatoria, cap regla de QRY_16 pot exigir-ne el valor.

---

## 9. Taules relacionals sense us

`T_CONNECTIONS`, `T_GROUPS`, `T_ARCH_FEATURES`, `T_INDIVIDUALS`, `T_DATING`: 0 registres.

Aixo contradiu el criteri de registre ja fixat: hi ha **cinc parells a/b** entrats (EA40a/b, EA02a/b, EA12a/b, EA22a/b, EA07a/b) que son exactament el cas per al qual es van dissenyar T_CONNECTIONS i T_GROUPS. La relacio de junta vertical no travada existeix ara nomes com a convencio al camp `Code`.

### 9.1 Camp nou `Chrono_Relation` a T_CONNECTIONS

Esquema actual: `ID`, `ID_Struct_A`, `ID_Struct_B`, `Connection_Type`, `Confidence`, `Notes`.

```
Chrono_Relation   TEXT(20)
   A earlier than B     A es anterior; B s'hi adossa
   B earlier than A     B es anterior; A s'hi adossa
   Contemporary         construccio en un sol episodi (juntes travades)
   Undetermined         [DEFECTE]
```

**Criteri operatiu:** la direccio es determina per adossament — l'estructura que presenta la junta no travada contra el parament de l'altra es la posterior. Quan la relacio no siga observable a la junta, `Undetermined`. Rellevant per a H03 i H04.

Nota: el sentit de la relacio depen de l'ordre `ID_Struct_A` / `ID_Struct_B`, que **no** ha de reordenar-se un cop introduit el registre.

### 9.2 Formulari

Afegir subformulari de connexions a la pestanya principal de F_STRUCTURES, de manera que registrar la relacio siga el cami de menor resistencia.

---

## 10. Eliminacio de NULL a camps observacionals

Els 35 registres tenen entre 5 i 91 NULL en camps observacionals, amb gradient segons ordre d'entrada — es a dir, NULL = «encara no avaluat», no «no observable».

Amb `DEFAULT 0` aixo deixa de produir-se en registres nous. Per als 35 existents cal **revisio manual**: cada NULL ha de resoldre's a 0, 1, 2, 3 o 9.

> **AVIS PER A LA IMPLEMENTACIO:** no es pot fer per script. La conversio automatica de NULL a `0` afirmaria centenars d'absencies no verificades i invalidaria l'analisi. Qualsevol agent que implemente aquesta especificacio ha de deixar els NULL existents intactes.

**Prioritat:** els 14 registres amb menys de 8 NULL ja son quasi complets. Els 21 restants requereixen segona visita al model 3D.

---

## 11. Registres concrets a revisar

| Codi | Situacio | Accio acordada |
|---|---|---|
| `DW-S01-EA11` | Pigment perimetral sobre roca, cap element construit conservat | Reclassificar com a **Built funerary structure**, tipologia `Unclassifiable`. Elements a **9** (no a 0): la posicio ja no existeix. Dimensions mesurables des del perimetre de pigment (`Dim_Method = perimeter pigment outline`). Entrada a T_LOST_ELEMENTS amb `Evidence_Scope = Whole structure` |
| `DW-S01-EA09` | Coveta amb pigment perimetral i dues mensules amb lloses | Reclassificar de `MIX` a **`CAV`**. `Sys_Platform = Present complete` o `Present partial`, resta `Not applicable`. Corregir `Timber_Bracket_Role` de `Isolated` a `Platform support`. `Platform_Function = Undetermined` |
| `DW-S01-EA38` | Cos basal + cos de cambra + mur + 7 mensules sobre repla artificial | Reclassificar de `ND` a tipologia real (probablement `CAM` o `MAU`) |
| `DW-S01-EA39` | Idem, 5 mensules | Idem |
| `DW-S01-EA40b` | 2 cossos de cambra, 2 murs, MNI=1 | Reclassificar de `ND` a tipologia real |
| `DW-S01-EA10`, `DW-S01-EA21` | Mensules aillades, sense restes humanes | Confirmar `MEN` -> `Structural trace` |
| `DW-S04-EA12a` | Unica cornisa intercos de fusta | Verificar al model 3D (seccio 5.6) |

### Incoherencies detectades a corregir

- `Platform_Surface_Material = Stone` amb `Timber_Bracket_Role = Isolated` — EA10, EA21, EA09.
- `DW-S04-EA09`: plataforma present amb E, F i G tots NULL.
- `DW-S01-EA07a`: `Plaster_Present = 9` amb Color i Extent = ND.
- `Human_Remains = 1` amb `MNI` NULL — EA17, S04-EA18, S01-EA01.
- `Rear_Wall = 0` amb `Rear_Wall_Type = Built masonry` — EA01, EA06 (es resol amb 5.5).

---

## 12. Regles per a QRY_16 (validacio)

1. Component d'un sistema amb valor 1/2/3 mentre el `Sys_*` corresponent no es `Present*` ni `Attested lost`
2. Qualsevol camp d'element o de sistema amb valor `3` / `Attested lost` sense fila a T_LOST_ELEMENTS
3. `Platform_Surface_Material` no nul amb `Timber_Bracket_Role = 'Isolated'`
4. `Timber_Brackets IN (1,2)` amb `Timber_Bracket_Count` nul
5. Camps d'un sistema marcat `Not applicable` amb valor <> 0
6. `Human_Remains = 1` amb `MNI` nul
7. `ID_Arch_Status = 'Good'` amb mes del 30% dels elements a valor 2 o 3
8. Qualsevol camp observacional amb valor NULL
9. `Record_Class = 'Structural trace'` amb `Human_Remains = 1`
10. `Record_Class` incompatible amb pestanyes emplenades (p. ex. `Rock art panel` amb dades bioarqueologiques)
11. `Chrono_Relation <> 'Undetermined'` amb `Connection_Type` que no implique adossament
12. `ID_Material_Status = 'Absent'` amb algun `Mat_*` <> 0
13. `Lintel = 0` amb `Lintel_Material` no nul

---

## 13. Ordre d'implementacio recomanat

1. `L_ELEMENTS` i `L_LOST_EVIDENCE` (lookups nous, sense dependencies)
2. Ampliacions de lookup: `L_SUPPORT` (microrepisa), `L_TYPOLOGY` (classe, desdoblament ND, eliminacio MIX), `Pigment_Extent`
3. Domini de cinc valors als 20 camps d'element + domini 0/1/9 reetiquetat a la resta + `DEFAULT 0`
4. Reanomenaments: `Lateral_Walls` -> `Return_Wall`, `Lateral_Wall_Faces` -> `Facade_Flank`, `Recessed_Portal` -> `Recessed_Frame`, `Rear_Wall_Type` -> `Rear_Closure_Type`
5. Desdoblament de `Lintel` en `Lintel` + `Lintel_Material`; migracio dels 21 valors
6. Camps `Sys_*` (cinc) amb els seus dominis respectius; eliminacio de `Corbelled_Platform`, `Access_Opening`, `Eave`, `Rear_Wall`
7. `Platform_Function`
8. `T_LOST_ELEMENTS` + relacions
9. `Chrono_Relation` a T_CONNECTIONS
10. Formulari: gating de nivell 1 (`Record_Class` -> pestanyes) i nivell 2 (`Sys_*`, `ID_Material_Status`, `Human_Remains`)
11. Subformularis de connexions i de decoracio
12. QRY_13 amb semantica NA, QRY_16 amb les 13 regles
13. **Nomes despres de migracio verificada:** eliminacio dels booleans `Dec_*`, `Rock_Art`, `RA_*` i `Lost_Body_Evidence`

**Atencio:** els camps reanomenats o eliminats apareixen a QRY_13, QRY_16 i als controls del formulari. Cal actualitzar-hi totes les referencies.

---

## 14. Decisions tancades

| Questio | Resolucio |
|---|---|
| Llindar complet / parcial | **75%**, criteri dual comptable/continu, subsidiari de llegibilitat morfologica (1.3) |
| Abast del domini de 5 valors | **Nomes els 20 camps d'element A–X.** La resta mante 0/1/9 (1.5, 1.6) |
| `ID_Parent` i `ID_Group` | **Mantenir, sense forcar-ne l'us** (8) |
| `Access_Opening` | Absorbit per `Sys_Portal`; `Recessed_Portal` -> `Recessed_Frame` (4.1, 5.2) |
| `Chrono_Relation` | **Afegit** a T_CONNECTIONS (9.1) |
| `Lost_Body_Evidence` | **Migrar i eliminar** (2) |
| H, P, U com a sistemes | **Absorbits** pels camps `Sys_*` amb domini estes (4.1) |
| Murs laterals / paraments | `Return_Wall` (V) i `Facade_Flank` (L) (5.1) |
| Dintell | Desdoblat en `Lintel` + `Lintel_Material` (5.4) |
| Mur posterior | `Rear_Wall` eliminat; `Rear_Wall_Type` -> `Rear_Closure_Type` conservat (5.5) |
| Cornisa intercos | Material **conservat** (anomalia EA12a) (5.6) |
| Microrepisa | Afegida a `L_SUPPORT` amb llindar <50cm (5.7) |
| 4.Dec | **Nomes T_DECORATIONS**; `Relief_Frieze` es queda com a element (7) |
| 6.Mat | Gating per `ID_Material_Status`, sense camp nou (4.6) |

### Unica questio pendent

- Valor de la campanya activa per al defecte de `ID_Campaign` (8).

---

## 15. Abast de la revisio de dades

| Tasca | Volum | Metode |
|---|---|---|
| Resoldre NULL en camps observacionals | 21 registres | Manual, model 3D |
| Distingir complet / parcial als valors `1` actuals | ~183 caselles (84 en estructures `Good`) | Manual, model 3D |
| Assignar `Sys_*` als 35 registres | 175 valors | Manual, rapid |
| Migrar `Lintel` a `Lintel` + `Lintel_Material` | 21 valors | Script + verificacio |
| Migrar `Lost_Body_Evidence` a T_LOST_ELEMENTS | 25 valors | Manual |
| Reclassificar tipologies | 6 registres | Manual |
| Registrar connexions dels 5 parells a/b | 5 registres | Manual |
| Verificar cornisa de fusta EA12a | 1 registre | Manual, model 3D |
| Segona passada metrica | 35 registres | Pendent de Metashape / CloudCompare |
