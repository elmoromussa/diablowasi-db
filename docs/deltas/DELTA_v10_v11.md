# Especificacio de canvis v10 -> v11
## Base de dades `chachapoya_DB` — TFM Esteve Ribera Torro

**Data:** 2026-08-05 (rev. 4 — integra la revisio critica de la rev. 3; totes les decisions de la seccio 14 es mantenen intactes)
**Base de partida:** `chachapoya_DB_v10.bas` + `chachapoya_Form_v10_val.bas`
**Estat de les dades:** 35 registres a T_STRUCTURES (tots Diablo Wasi, sectors S01-S04), 29 a T_DECORATIONS, resta de taules buides
**Objectiu:** implementar els canvis d'esquema acordats i generar `chachapoya_DB_v11.bas` i `chachapoya_Form_v11_val.bas`

**Convencio de marcatge:** els passatges nous o modificats respecte a la rev. 3 van marcats **[REV4]**. Les decisions de disseny que la rev. 4 ha hagut d'introduir per a completar buits estan recollides a la seccio 14bis i pendents de vist-i-plau.

---

## 0. Convencions que es mantenen sense canvis

- Valencia nomes com a etiquetes d'interficie (captions de formulari, noms de pestanya). Tots els valors emmagatzemats, noms de camp, esquema i contingut de lookups en angles.
- Codificacio CP1252, cap caracter no-ASCII al codi VBA, finals de linia CRLF.
- Cap continuacio de linia (`& _`); patro `sql = sql & "..."` en linies separades.
- Tots els Sub auxiliars `Private`; nomes `BuildDB` i `BuildForm` publics.
- `DROP TABLE` abans de `CREATE TABLE`.
- Captions de columna via DAO `TableDef.Fields.Properties("Caption")`.
- **[REV4]** Els valors per defecte s'estableixen SEMPRE via DAO (`Field.DefaultValue`), mai amb la clausula `DEFAULT` al DDL: JET nomes accepta `DEFAULT` en DDL executat via ADO (mode ANSI-92), i tot l'script usa `db.Execute` DAO. El mecanisme de v10 (`SetByteDefaults`) es el correcte i es mante, reescrit amb els noms i valors v11.
- **[REV4]** Els combos amb domini controlat sobre camps TEXT (tipus, materials, extensions) passen a llistes de valors de DUES columnes: columna 1 (amagada) = valor emmagatzemat en angles; columna 2 = etiqueta en valencia. Es el mateix patro que PC9. Motiu: en una llista d'una columna el valor emmagatzemat ES l'etiqueta, i "reetiquetar sense canviar valors" (seccio 1.4) hi es impossible.

---

## 1. Domini dels camps d'element arquitectonic

### 1.1 Canvi

Substitueix el domini BYTE 0/1/9 per un gradient ordinal de cinc valors, aplicat **exclusivament als camps d'element del vocabulari A-X** (veure 1.6 per a la resta).

| Valor | Caption (valencia) | Etiqueta EN | Criteri operatiu |
|---|---|---|---|
| 0 | Absent | `Absent` | La posicio es examinable i no hi ha ni element ni traca. **VALOR PER DEFECTE** |
| 1 | Present, complet | `Present complete` | Es conserva **75% o mes** de l'element |
| 2 | Present, parcial | `Present partial` | Es conserva **menys del 75%**, pero l'element es identificable in situ |
| 3 | Desaparegut | `Attested lost` | No es conserva res in situ; nomes traca fisica directa. **Exigeix registre a T_LOST_ELEMENTS** |
| 9 | No observable | `Not observable` | La posicio no es examinable: destruida o amagada |

### 1.2 Regla semantica

Cada camp d'element respon una sola pregunta: **"va tenir l'estructura aquest element, i en quin estat esta ara?"**
El camp descriu l'element, mai la relacio de l'observador amb ell.

- `9` NO significa "no observat". Significa que la posicio no es examinable.
- Si s'ha mirat la posicio i no hi ha res, el valor es `0`.
- Agregacions analitiques: "que es va construir" = 1+2+3 · "que es conserva" = 1+2 · "integre" = 1.
  Denominador de cada camp = registres amb valor <> 9.
- **[REV4] Convencio tecnica del 0 sota `Not applicable`:** quan un sistema esta marcat `Not applicable`, els seus camps components queden a `0` per imperatiu tecnic (cap camp observacional pot ser NULL, i 9 tampoc es cert: la posicio no es que no siga examinable, es que la pregunta no es planteja). Aquest `0` es un farciment, no una absencia verificada, i QRY_13 el neutralitza exportant NULL (seccio 4.7). El document metodologic ha de recollir aquesta convencio explicitament.

### 1.3 Criteri operatiu del llindar 75% (per al document metodologic)

> Un element es registra com a **Present complet** quan se'n conserva el 75% o mes de la seua extensio original; com a **Present parcial** quan se'n conserva menys del 75% pero l'element continua sent identificable in situ.
>
> En **elements comptables** (mensules, bigues transversals, cantoneres, brancals) la proporcio es calcula sobre el nombre d'unitats originalment previstes, deduit de la simetria o de les traces conservades. Exemple: 5 mensules conservades de 7 encaixos documentats = 71% = parcial.
>
> En **elements continus** (paraments, cornises, socols, superficies de rafec) la proporcio s'estima sobre l'extensio lineal o superficial original, reconstruida a partir de l'amplaria de l'estructura i de les traces d'arrencament.
>
> Quan la proporcio no siga estimable amb un marge raonable, s'aplica el criteri subsidiari de **llegibilitat morfologica**: si la forma original no es reconstruible sense inferencia, es registra com a parcial.

Nota: el llindar del 75% es conservador — un element al qual falta una quarta part ja es "parcial". `Present complete` designa per tant elements quasi integres.

### 1.4 Etiqueta ND **[REV4: implementacio corregida]**

Elimina `ND` com a etiqueta als camps d'element; se substitueix per `No observable` / `Not observable`.

Als camps de tipus (`Lintel_Material`, `Chamber_Roof_Type`, `Mortar_Type`, `Rear_Closure_Type`, `Platform_Surface_Material`, `Masonry_Quality`, `Plaster_*`, `Pigment_*`) `ND` significa una altra cosa — l'element hi es pero no se'n determina el tipus.

**Implementacio corregida:** aquests camps NO son lookups de taula sino llistes de valors del formulari, on el valor emmagatzemat es l'etiqueta mateixa. La instruccio de la rev. 3 ("reanomena l'etiqueta sense canviar els valors introduits") era inaplicable tal com estava escrita i hauria produit un corpus mixt `'ND'` / `'Type undetermined'`. La solucio que compleix les dues condicions alhora:

- Convertir cada combo a llista de valors de DUES columnes (convencio de la seccio 0): valor emmagatzemat `'ND'` (columna amagada), etiqueta mostrada *Tipus indeterminat*.
- Els valors ja introduits (`'ND'`) no es toquen i les consultes continuen filtrant per `'ND'`.
- El valor emmagatzemat `'Type undetermined'` NO s'introdueix mai: nomes existeix com a etiqueta de pantalla.

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

**[REV4] Nota de sequenciacio:** al pas del domini (13.5) nomes es converteixen **19** d'aquests camps. `Lintel` es encara TEXT en aquell punt i s'incorpora al domini de 5 valors al seu propi pas de migracio (13.7).

### 1.6 Camps observacionals amb domini 0/1/9

Conserven tres valors, amb l'unica modificacio de reanomenar l'etiqueta de `9` a `Not observable` (caption *No observable*):

- Atribut de portal: `Recessed_Frame` (configuracio, no element: no admet gradient)
- Conservacio: `Looting`, `Fire_Damage`, `Animal_Activity`, `Modern_Access`
- Bioarqueologia: `Human_Remains`, `Anatomical_Connection`, `Mummification`, `Funerary_Bundles`, `Dispersed_Remains`, `Flexed_Position`, `Bone_Burning`
- Materials: `Mat_Textiles`, `Mat_Wood`, `Mat_VegFiber`, `Mat_Ceramics`, `Mat_Fauna`, `Mat_DeerAntler`, `Mat_Other`
- Suport i maconeria: `Support_Modified`, `Mortar_Present`, `Chinking_Stones`
- Acabats: `Plaster_Present`, `Pigment_Present`

Motiu: "present parcial" i "desaparegut" son categories morfologiques d'element construit. Aplicades a processos (`Looting`) o a vestigis mobles (`Mat_Textiles`) forcen una semantica que no correspon.

### 1.7 Implementacio tecnica **[REV4: mecanisme corregit i completat]**

- Defecte `0` a totes les BYTE observacionals (les de 5 valors i les de 3), establit **via DAO `Field.DefaultValue`** (reescriptura de `SetByteDefaults` amb els noms v11), no al DDL (veure seccio 0).
- Els controls combo del formulari han d'obrir amb `0` seleccionat, mai en blanc.
- **Cap camp observacional pot quedar NULL.** El NULL ja no es un estat valid del domini.
- **[REV4]** `ALTER TABLE ADD COLUMN` no ompli les files existents: qualsevol camp nou (els cinc `Sys_*`, `Platform_Function`, `Lintel_Material`, el nou `Lintel` BYTE) naix NULL als 35 registres, independentment del defecte. Aixo es el comportament desitjat: la regla 17 de QRY_16 (cap NULL observacional) fa de quadre de pendents fins que la revisio manual els resol. El defecte nomes actua sobre registres nous.
- **[REV4]** Tipus i defectes dels camps de sistema (no especificats a la rev. 3): els cinc `Sys_*` son `TEXT(20)` (el valor mes llarg, `'Present complete'`, fa 16 caracters), amb defecte `'Absent'` via DAO, coherent amb la filosofia del defecte 0 dels elements. `Chrono_Relation` te defecte `'Undetermined'` via DAO.

---

## 2. Taula nova: `T_LOST_ELEMENTS` **[REV4: Element_Code opcional]**

Justificacio: el valor `3` es l'unic del domini que es una inferencia i no una observacio directa. Sense registre de l'evidencia no es replicable. Es a mes la base evidencial de l'OE3 (reconstruccio de la infraestructura de circulacio aeria).

```
T_LOST_ELEMENTS
  ID                 AUTOINCREMENT  PRIMARY KEY
  ID_Structure       LONG    FK -> T_STRUCTURES.ID    NOT NULL
  Element_Code       TEXT(2) FK -> L_ELEMENTS.Code    (opcional; veure regla d'abast)
  ID_Evidence_Type   LONG    FK -> L_LOST_EVIDENCE.ID NOT NULL
  Evidence_Scope     TEXT(20)   Element / Body / Whole structure
  ID_Position        LONG    FK -> L_STRUCT_BODY.ID   (opcional)
  Notes              MEMO
```

**[REV4] Canvi respecte a la rev. 3:** `Element_Code` deixa de ser NOT NULL. La rev. 3 exigia files d'abast `Body` (migracio de `Lost_Body_Evidence`) i `Whole structure` (EA11) que el seu propi DDL rebutjava, perque un cos o una estructura sencera no tenen un codi d'element unic. Regla substituta, no bloquejant:

- `Element_Code` es **obligatori si i nomes si `Evidence_Scope = 'Element'`** (regla 21 de QRY_16 i validacio de formulari).
- A les files `Body` / `Whole structure`, `Element_Code` queda buit i `Notes` pot detallar els elements implicats.

**[REV4] Requisit tecnic:** la relacio `Element_Code -> L_ELEMENTS.Code` exigeix un **index unic** sobre `L_ELEMENTS.Code` (JET nomes accepta relacions contra la clau primaria o un index unic). Veure seccio 3.

### Lookup `L_LOST_EVIDENCE` **[REV4: descripcions en angles, per convencio de la seccio 0]**

| ID | Name | Description |
|---|---|---|
| 1 | Negative socket / impression | Empty socket or impression left in the wall face |
| 2 | Beam hole | Through-hole for a spanning beam |
| 3 | Break scar | Detachment scar on masonry or bedrock |
| 4 | Detached fragment in situ | Identifiable fallen fragment at the foot of the structure |
| 5 | Mortar imprint | Mortar imprint without the element it once bonded |
| 6 | Corbels into void | Corbels that no longer support anything |
| 7 | Pigment on bedrock | Pigment on bedrock now bare of the masonry that carried it |
| 8 | Truncated walls | Walls truncated on a clean plane |
| 9 | Other (see Notes) | |

### Regla d'abast del valor 3

S'admet `3` a qualsevol element **si i nomes si** es pot assenyalar un valor de `L_LOST_EVIDENCE` aplicable. Si no n'hi ha cap, el valor correcte es `0` o `9`. El lookup fa de filtre metodologic; no cal una llista tancada d'elements.

Els tres camps de sistema (`Sys_Platform`, `Sys_Portal`, `Sys_Eave`) admeten igualment `Attested lost` i queden subjectes a la mateixa regla.

**[REV4] Satisfaccio de la regla 2 de QRY_16:** un valor `3` en un camp d'element queda justificat per (a) una fila amb `Element_Code` coincident, o (b) una fila d'abast `Body` o `Whole structure` de la mateixa estructura, que cobreix en bloc els elements del cos o de l'estructura perduts.

### Migracio de `Lost_Body_Evidence` **[REV4: mapatge explicit]**

Queda cobert per aquesta taula amb `Evidence_Scope = 'Body'` i `Element_Code` buit. Mapatge de valors:

| Valor v10 de `Lost_Body_Evidence` | Fila a T_LOST_ELEMENTS |
|---|---|
| `Pigment on bedrock` | ID_Evidence_Type = 7 |
| `Truncated walls` | ID_Evidence_Type = 8 |
| `Empty beam sockets` | ID_Evidence_Type = 1 (*Negative socket / impression*) |
| `Corbels into void` | ID_Evidence_Type = 6 |
| `Detached debris` | ID_Evidence_Type = 4 |
| `None` | **cap fila** |
| `ND` | **cap fila** (el 9 dels camps d'element ja recull la no-observabilitat) |

Migrar manualment i **eliminar el camp** nomes un cop verificada la migracio (pas 13.14). **Verificacio previa:** confirmar si el recompte de "25 valors" (seccio 15) inclou o exclou `None`/`ND`; nomes els valors amb evidencia real generen fila.

---

## 3. Lookup nou: `L_ELEMENTS` **[REV4: index unic + reanomenament d'una columna]**

El vocabulari A-X existeix actualment nomes com a noms de camp; no es referenciable des d'altres taules ni exportable com a entitat. `L_STRUCT_BODY` son **posicions** (11 entrades), no elements.

```
L_ELEMENTS
  ID            AUTOINCREMENT PRIMARY KEY
  Code          TEXT(2)   A .. X   -- CONSTRAINT UQ_ELEM_CODE UNIQUE
  Name_EN       TEXT(60)
  Name_VAL      TEXT(60)
  Level_Type    TEXT(10)  N0 / N0-N1 / N1 / SUP
  Sys_Group     TEXT(20)  Platform / Portal / Eave / (buit si no pertany a cap sistema)
  Is_System     YESNO     TRUE per a H, P, U
  Field_Name    TEXT(40)  nom del camp corresponent a T_STRUCTURES
  Description   MEMO
```

**[REV4] Canvis respecte a la rev. 3:**

- `Code` porta index unic (`UQ_ELEM_CODE`): sense ell, la relacio des de `T_LOST_ELEMENTS.Element_Code` no es pot crear en JET.
- La columna `System` es reanomena **`Sys_Group`**: `System` no consta com a paraula reservada JET documentada, pero el projecte ja va patir el cas de `Level` i el terme es prou generic per a aplicar-hi la mateixa prudencia, a cost zero (la taula encara no te dades).

Poblar amb els 24 elements A-X segons el vocabulari normalitzat vigent (ordre bottom-to-top), inclosos H, P i U amb `Is_System = TRUE` i `Field_Name` apuntant al camp `Sys_*` corresponent. W s'hi mante com a entrada de vocabulari encara que ja no tinga camp propi (`Field_Name` buit).

**[REV4] Nota de sequenciacio:** `Field_Name` es pobla al pas 13.1 amb noms v11 que no existiran fins als passos 13.6-13.8. Es text pla i no trenca res, pero el pas 13.15 inclou una verificacio final: cada `Field_Name` no buit ha de correspondre a un camp real de T_STRUCTURES.

---

## 4. Sistemes constructius i gating

### 4.1 Els tres elements-sistema

El vocabulari defineix tres elements que son en realitat sistemes compositius. En v11 deixen de tenir camp propi i **son** els camps de sistema:

| Element | Camp v11 | Components | Camp v10 eliminat |
|---|---|---|---|
| H Plataforma | `Sys_Platform` | E + F + G | `Corbelled_Platform` |
| P Portal | `Sys_Portal` | N + O + Q | `Access_Opening` |
| U Rafec | `Sys_Eave` | S + T | `Eave` |

Domini d'aquests tres (sis valors), tipus `TEXT(20)`:
`Present complete` / `Present partial` / `Attested lost` / `Absent` / `Not applicable` / `Not observable`
Captions: *Present complet* / *Present parcial* / *Desaparegut* / *Absent* / *No aplicable* / *No observable*

D'aquesta manera no es perd el gradient de conservacio en fusionar element i sistema.

**[REV4] Migracio dels valors v10 (abans de cap eliminacio, pas 13.8):** els camps H, P i U de v10 contenen observacions ja fetes que no s'han de destruir. Migracio per script, deterministica i sense tocar cap NULL:

| Valor v10 (`Corbelled_Platform` / `Access_Opening` / `Eave`) | Valor al `Sys_*` corresponent |
|---|---|
| `0` | `'Absent'` |
| `9` | `'Not observable'` |
| `1` | **NULL** (pendent de revisio manual: complet / parcial / desaparegut) |
| NULL | NULL (mai convertir; seccio 10) |

`Rear_Wall` no te desti directe (`Sys_Chamber` agrega sis elements): no s'hi aplica script; l'assignacio de `Sys_Base` i `Sys_Chamber` es integrament manual. L'eliminacio dels quatre camps v10 es trasllada al pas final (13.14), **nomes despres de migracio verificada** — el mateix principi que la rev. 3 ja aplicava als booleans `Dec_*`.

### 4.2 Els dos camps d'agrupacio

```
Sys_Base      elements A B C D
Sys_Chamber   elements J K L M V X
```

No corresponen a cap element del vocabulari: son agrupacions per a gating. Tipus `TEXT(20)`, domini de quatre valors:
`Present` / `Absent` / `Not applicable` / `Not observable`

### 4.3 Semantica analitica (critica)

| Valor del sistema | Com surten els seus camps components a l'exportacio R |
|---|---|
| `Present complete` / `Present partial` / `Present` | valors reals (0/1/2/3/9) |
| `Attested lost` | valors reals; s'espera `3` o `9` als components |
| `Absent` | `0` — l'absencia es informativa (un mausoleu sense rafec es un resultat) |
| `Not applicable` | **`NA`** — no entra a cap denominador |
| `Not observable` | `NA` |

Aquesta distincio entre *no aplicable* i *absent verificat* es la mateixa logica del 0/9 aplicada al nivell de sistema. **Sense ella, els percentatges publicats son incorrectes.** Veure tambe la convencio tecnica del 0 sota `Not applicable` (1.2).

### 4.4 Gating de nivell 1: `Record_Class` -> pestanyes **[REV4: taula completa i numeracio del formulari v10]**

`Record_Class` (derivat de `L_TYPOLOGY`, veure seccio 6) determina quines pestanyes del formulari s'activen. La taula de la rev. 3 ometia 3.Acab i 12.Extra i usava una numeracio que no corresponia al formulari v10; aquesta versio usa els noms i numeros reals de les 12 pestanyes i afegeix la classe de treball `Pending classification` (seccio 6):

| Pestanya | Built funerary structure | Natural funerary context | Structural trace | Rock art panel | Pending classification |
|---|---|---|---|---|---|
| 1.Id | si | si | si | si | si |
| 2.Arq (morfologia) | si | **no** | **no** | **no** | si |
| 3.Acab (acabats) | si | si | si | **si** | si |
| 4.Dec | si | si | si | si | si |
| 5.Estat (conservacio i observabilitat) | si | si | si | si | si |
| 6.Bio | si | si | **no** | **no** | si |
| 7.Mat | si | si | **no** | **no** | si |
| 8.Cron | si | si | si | si | si |
| 9.Metr. | si | si | si | si | si |
| 10.Doc | si | si | si | si | si |
| 11.Sist | si | parcial | parcial | **no** | si |
| 12.Extra | si | si | si | si | si |

Notes:
- **3.Acab activa per a tot:** un *Rock art panel* es defineix precisament per `Pigment_*`; una traca estructural pot conservar pigment sobre la mensula.
- "Parcial" vol dir que els cinc `Sys_*` es resolen igualment, pero la majoria eixiran `Not applicable`.
- **[REV4] Valors de les pestanyes tancades:** els camps BYTE de pestanyes desactivades pel gating queden al seu defecte `0` (convencio de farciment analoga a 1.2); els camps no-BYTE (MNI, tipus) queden buits. La regla 16 de QRY_16 vigila que no s'hi introduisca contingut real (definicio d'"emplenada" a la seccio 12).

### 4.5 Gating de nivell 2: `Sys_*` -> grups d'elements **[REV4: llistes completes per als cinc sistemes]**

A la pestanya `11.Sist`, els cinc combos `Sys_*` van dalt de tot. Els grups de camps components queden desactivats (`Enabled = False`) si el sistema corresponent no es `Present*` o `Attested lost`. Cal codi al `Current` i als `AfterUpdate` dels cinc combos.

| Sistema | Camps subordinats |
|---|---|
| `Sys_Platform` | `Timber_Brackets`, `Timber_Bracket_Count`, `Timber_Bracket_Role`, `Transverse_Beams`, `Corbelled_Courses`, `Platform_Surface_Material`, `Platform_Function` |
| `Sys_Portal` | `Sill`, `Jambs`, `Lintel`, `Lintel_Material`, `Recessed_Frame` |
| `Sys_Eave` | `Eave_Beam`, `Eave_Surface` |
| `Sys_Base` | `Embedded_Base_Beams`, `Base_Level`, `Decorative_Socle`, `Tie_Walls` |
| `Sys_Chamber` | `Corner_Quoins`, `Structural_Pilasters`, `Facade_Flank`, `Relief_Frieze`, `Return_Wall`, `Chamber_Roof`, `Chamber_Roof_Type`, `Rear_Closure_Type` |

- `Opening_Width_cm` / `Opening_Height_cm` (9.Metr.) NO es gategen: la segona passada metrica es independent.
- **[REV4] Sense gating, intencionadament:** `Interbody_Cornice` (I) + `Interbody_Cornice_Material` i `Upper_Crown` (R) no pertanyen a cap sistema ni agrupacio; queden sempre actius a 11.Sist.

### 4.6 Gating de nivell 2 fora de 11.Sist

**Pestanya 7.Mat.** No cal camp nou: `ID_Material_Status` ja fa de porta.

| `ID_Material_Status` | Comportament dels camps `Mat_*` |
|---|---|
| `Absent` | tots a `0`, desactivats |
| `ND` | tots a `9`, desactivats |
| `Good` / `Fair` / `Poor` | actius, entrada detallada |

**Pestanya 6.Bio.** `Human_Remains` fa de porta per a `MNI`, `Anatomical_Connection`, `Mummification`, `Funerary_Bundles`, `Dispersed_Remains`, `Flexed_Position`, `Bone_Burning`.

### 4.7 Actualitzacio de QRY_13 **[REV4: abast precisat]**

L'exportacio ha d'emetre `NULL` (-> `NA` a R) per als camps de sistema marcats `Not applicable` o `Not observable` **i tambe per als seus camps components** (que a la taula contenen el 0 de farciment de 1.2), no el valor cru. Especificacio completa de QRY_13 v11 a la seccio 12bis.

---

## 5. Vocabulari: reanomenaments i desdoblaments

### 5.1 Murs laterals vs paraments laterals

Els noms v10 son ambigus: tots dos contenien "lateral" referit a coses distintes. Les dades confirmen que son variables independents (les cambres tendeixen a L=1 amb V=0; els mausoleus a V=1).

| Element | Nom v11 | Caption | Definicio |
|---|---|---|---|
| V | `Return_Wall` | Mur de retorn | Mur perpendicular al pla de facana, que retorna cap a la penya |
| L | `Facade_Flank` | Ala de facana | Parament en el pla de facana, flanquejant l'obertura |

**[REV4]** L'entrada corresponent de `L_STRUCT_BODY` s'actualitza en coherencia: `LWF "Lateral wall face (L)"` passa a `FFL "Facade flank (L)"` (mateix ID; nomes Code i Name).

### 5.2 Portal i marc reculat

| Nom v10 | Nom v11 | Caption |
|---|---|---|
| `Access_Opening` | -> `Sys_Portal` (seccio 4.1) | Sistema portal |
| `Recessed_Portal` | `Recessed_Frame` | Marc reculat |

`Recessed_Frame` es un atribut de configuracio del portal, no un element: domini 0/1/9 (veure 1.6).

### 5.3 Separacio forma / funcio a la plataforma

`Corbelled_Platform` tenia caption *Plataforma d'acces*, barrejant morfologia verificable amb interpretacio funcional. L'OE3 existeix precisament per a determinar la funcio de circulacio; anomenar-la "acces" al camp de dades preassumeix la conclusio.

- `Sys_Platform` caption: *Plataforma en voladis* (nomes morfologia)
- Camp nou:

```
Platform_Function   TEXT(20)
   Access          acces a una estructura concreta
   Circulation     transit entre estructures
   Construction    bastida constructiva sense us posterior
   Support         base d'una estructura superior
   Multiple        mes d'una funcio documentada
   Undetermined    [DEFECTE, via DAO]
```

Precedent intern: `Timber_Bracket_Role` ja separa forma i funcio per a l'element E.

### 5.4 Dintell (element Q) **[REV4: migracio completada amb NULLs i mecanica]**

`Lintel` en v10 es un lookup de material (`Stone` 10 · `Absent` 6 · `Wood` 3 · `ND` 2) que barreja presencia i tipus. Desdoblar seguint el patro ja emprat a `Chamber_Roof` / `Chamber_Roof_Type`:

```
Lintel            BYTE, domini de 5 valors (element Q, component de Sys_Portal)
Lintel_Material   TEXT(20)   Stone / Wood / Mixed / ND
```

(`ND` amb etiqueta *Tipus indeterminat* segons 1.4; el valor emmagatzemat es `'ND'`.)

**Mecanica de la migracio (JET no pot convertir TEXT->BYTE in situ amb `'Stone'` dins):**

1. Reanomenar `Lintel` -> `Lintel_Material` via DAO (`Field.Name`), que conserva les dades.
2. Crear el nou camp `Lintel` BYTE.
3. Poblar per script els **21 valors**:

| Valor v10 | `Lintel` (nou) | `Lintel_Material` |
|---|---|---|
| `Stone` (10) | `1` | `Stone` (es conserva) |
| `Wood` (3) | `1` | `Wood` (es conserva) |
| `Absent` (6) | `0` | NULL (esborrar) |
| `ND` (2) | `9` | NULL (esborrar: si la presencia mateixa es desconeguda, el tipus no es "indeterminat", es inaplicable) |

4. Els **~14 registres amb `Lintel` NULL en v10 queden NULL** al camp nou: es resolen manualment (regla 17 de QRY_16 en fa el seguiment), coherent amb la seccio 10.
5. Verificar: `SELECT COUNT` per valor abans i despres; total migrat = 21.

La conversio de `Stone`/`Wood` a `1` es provisional: la revisio complet/parcial (seccio 15) ha de confirmar si algun es `2`.

**[REV4] Consequencia a QRY_13:** la derivacio `AX_Q` per IIF sobre text desapareix; `AX_Q` passa a exportar el camp `Lintel` BYTE directament.

### 5.5 Mur posterior (element W)

Els camps v10 estan mal utilitzats: hi ha `Rear_Wall=0` amb `Rear_Wall_Type='Built masonry'` a EA01 i EA06, que es contradictori.

- **`Rear_Wall` s'elimina** — redundant amb `Sys_Chamber`. (Eliminacio al pas 13.14, mai abans de l'assignacio manual de `Sys_Chamber`.)
- **`Rear_Wall_Type` es conserva**, reanomenat a `Rear_Closure_Type`:

```
Rear_Closure_Type   TEXT(20)   Natural bedrock / Built masonry / Mixed / ND
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

### 5.8 Ampliacio de `Pigment_Extent` **[REV4: valor escurcat i mecanisme precisat]**

Afegir el valor **`Perimeter/threshold`** (19 caracters; el de la rev. 3, `Perimeter / threshold`, en tenia 21 i no cabia a `Pigment_Extent TEXT(20)`). Caption: *Perimetral / llindar*.

`Pigment_Extent` no es un lookup de taula: l'ampliacio es fa al RowSource del combo del formulari (llista de dues columnes segons la seccio 0).

Motiu: hi ha pigment perimetral en dos contextos estructuralment distints (EA11, estructura arrasada; EA09, cavitat natural). Combinat amb `Pigment_Substrate = Bedrock` dona una consulta directa sobre una practica de marcatge del llindar independent del tipus de suport — rellevant per a la dimensio simbolica del marc teoric.

---

## 6. Tipologia: classe de registre, ND i MIX

### 6.1 Columna de classe a `L_TYPOLOGY` **[REV4: taula completa]**

Camp nou `Record_Class TEXT(30)` a `L_TYPOLOGY` (no a T_STRUCTURES: es derivable de la tipologia).

| Tipologies | Record_Class |
|---|---|
| MAU, CAM, PLA-R, PLA-V | `Built funerary structure` |
| NIX, CAV | `Natural funerary context` |
| MEN | `Structural trace` |
| PR | `Rock art panel` |
| **[REV4]** Unclassifiable | `Built funerary structure` |
| **[REV4]** Not yet classified | `Pending classification` |

**[REV4] Criteri operatiu d'`Unclassifiable`:** es reserva per a vestigis d'estructures **construides** el tipus de les quals no es pot determinar (cas EA11: perimetre de pigment que implica una estructura desapareguda). Els contexts naturals son sempre classificables com a NIX o CAV per criteris geometrics, de manera que no hi ha "natural inclassificable". Aixo fa coherent la reclassificacio d'EA11 de la seccio 11 amb el fet que `Record_Class` viu a L_TYPOLOGY (una classe per tipologia, no per registre).

**[REV4] `Pending classification`** es una classe de treball, no un resultat: totes les pestanyes actives (4.4) i exclosa automaticament de qualsevol consulta analitica pel filtre de `Record_Class`.

Totes les consultes analitiques han d'incorporar filtre per `Record_Class`. Motiu: dels 35 registres actuals, 5 no son estructures; qualsevol percentatge calculat sobre 35 es incorrecte.

### 6.2 Desdoblament de ND **[REV4: mecanica UPDATE, no DELETE]**

`ND Undetermined` (ID 10) cobreix dos estats distints. Substituir per:

- `Unclassifiable` — evidencia insuficient per a classificar. **Es un resultat.**
- `Not yet classified` — estat de treball, pendent de revisio.

**Mecanica:** `UPDATE` in situ de la fila existent (ID 10) a `Not yet classified` — que es semanticament exacte, perque tots els registres ND actuals (EA38, EA39, EA40b...) estan pendents de revisio — mes un `INSERT` d'`Unclassifiable` com a fila nova. **Mai DELETE+INSERT:** la integritat referencial de `REL_TYP_STR` bloquejaria el DELETE mentre hi haja registres que referencien l'ID 10, i el patro `X()` de l'script silenciaria l'error.

### 6.3 Eliminacio de MIX **[REV4: precondicio d'ordre]**

`MIX Mixed` (ID 9) s'elimina. Motiu: tota EA-CAM es mixta per definicio (cavitat natural + facana construida), de manera que la categoria no distingeix res. La combinacio natural/construit ja la capturen `ID_Support` i els camps `Sys_*`.

**Precondicio:** EA09 (l'unic registre conegut amb MIX) s'ha de reclassificar a CAV **abans** del DELETE, i el DELETE ha d'anar precedit d'un `SELECT COUNT(*) FROM T_STRUCTURES WHERE ID_Typology = 9` que retorne 0. El DELETE s'executa amb `dbFailOnError` directe, **no** via `X()`: si falla, ha de fallar sorollosament.

Equivalent derivat: tipologia NIX o CAV amb algun `Sys_*` a `Present*`.

---

## 7. Decoracio: registre unic a T_DECORATIONS

### 7.1 Problema

Els booleans `Dec_*` de T_STRUCTURES dupliquen la informacio de T_DECORATIONS, que si registra posicio, tipus, color i substrat. Amb nomes 35 registres ja divergeixen en **6 casos** (17%).

### 7.2 Canvi

Elimina de T_STRUCTURES (pas 13.14, nomes despres de la migracio de 7.5):

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

### 7.6 Reconstruccio de QRY_02 **[REV4: seccio nova — la rev. 3 destruia la consulta sense substituir-la]**

QRY_02_Decoration_by_Site esta construida integrament sobre els booleans que 7.2 elimina, i es la font del khi-quadrat LP-DW. A mes, T_DECORATIONS nomes registra **presencies**: la terna Present / Absent verificat / No observable que permetia triar el denominador no hi existeix. QRY_02 v11 es reconstrueix aixi:

- **Present** (per estructura i motiu): existeix fila a T_DECORATIONS amb l'`ID_Dec_Type` corresponent. Implementacio JET: `T_STRUCTURES LEFT JOIN T_DECORATIONS` + `SUM(IIF(...))` per tipus.
- **Absent verificat:** cap fila del motiu **i** `Facade_Observability = 'Complete'`. (Criteri estricte i deliberadament conservador: una facana parcialment observable no permet afirmar l'absencia d'un motiu concret.)
- **No avaluable:** cap fila del motiu i `Facade_Observability` distint de `'Complete'` (o NULL).
- Denominador valid del test = Present + Absent verificat; el percentatge de no avaluables es reporta com a cobertura, com en v10.
- Filtre obligatori per `Record_Class` (6.1).

El criteri del denominador estricte es documenta al capitol metodologic.

---
