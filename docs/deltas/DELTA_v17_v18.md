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
