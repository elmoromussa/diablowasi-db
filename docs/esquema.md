**ESQUEMA DE LA BASE DE DADES ARQUEOLÒGICA**

*La Petaca i Diablo Wasi (Leymebamba, Amazonas, Perú)*

chachapoya_DB_v12.bas + chachapoya_Form_v12_val.bas (construcció des de zero)
migrate_v10_to_v11.bas + upgrade_form_v12.bas (actualització d'una BD amb dades)

Sub BuildDB() + Sub BuildForm() | MigrateV10toV11() + UpgradeV12() | Microsoft Access JET SQL | Esteve Ribera Torró

*Versió 12 del document — actualitzada segons el codi v12 (agost 2026). La numeració dels fitxers del projecte queda alineada a v12. Aquest document substitueix la versió 10; els canvis de la iteració v10→v11 (delta rev. 5) i de l'addendum v12 hi queden consolidats.*

# **1. Resum general**

| **Element** | **Valor** |
| --- | --- |
| Taules principals | T_STRUCTURES (120 camps), T_DATING, T_INDIVIDUALS, T_GROUPS, T_DECORATIONS, T_ARCH_FEATURES, T_CONNECTIONS, T_LOST_ELEMENTS |
| Taules lookup | L_SITES, L_SECTORS, L_TYPOLOGY, L_SUPPORT, L_STATUS, L_MATERIAL_STATUS, L_VOL_METHOD, L_COORD_METHOD, L_GROUP_TYPE, L_CAMPAIGN, L_STRUCT_BODY, L_DEC_TYPE, L_ELEMENTS, L_LOST_EVIDENCE |
| Total taules | 22 |
| Total relacions | 25 (inclou l'autoreferenciant de T_STRUCTURES, les dues de T_CONNECTIONS, la del suport secundari i les quatre de T_LOST_ELEMENTS) |
| Consultes SQL | 26 (família QRY_01–QRY_16; la bateria QRY_16 comprén cinc consultes auxiliars, tres parcials de regles i la consulta unió amb 26 regles) |
| Formulari | F_STRUCTURES — 12 pestanyes + subformularis F_DECORATIONS, F_ARCH_FEATURES, F_CONNECTIONS i F_LOST_ELEMENTS. Gating de tres nivells amb emplenat ràpid confirmat. Etiquetes UI en valencià; valors emmagatzemats en anglés |
| Dominis observacionals | 20 camps d'element A–X amb domini de cinc valors 0/1/2/3/9; 25 camps observacionals amb domini 0/1/9; 5 camps de sistema TEXT amb domini de sis valors. Valor per defecte 0 |
| Idioma BD | Anglés (noms de taules, camps i valors lookup). UI del formulari en valencià. Capçaleres de subformulari via etiquetes adjuntes; captions DAO com a reforç |
| Motor | Microsoft Access JET SQL / ACE │ cp1252 |
| Jaciments | La Petaca (WGS84: Lat -6.8311, Lon -77.8084) │ Diablo Wasi (Lat -6.8475, Lon -77.8154) |

## **1.1. Canvis de la iteració v10→v11 (delta rev. 5)**

**(a) Domini de cinc valors per als 20 elements A–X.** Els camps d'element passen de 0/1/9 a **0 = Absent / 1 = Present complet / 2 = Present parcial / 3 = Desaparegut atestat / 9 = No observable**. El gradient permet comptar per separat *què es va construir* (1+2+3), *què sobreviu* (1+2) i *què resta intacte* (1) — tres preguntes que el domini binari col·lapsava en una. El valor 3 és l'únic del domini que és una inferència i no una observació: per això exigeix evidència material registrada a T_LOST_ELEMENTS (regla 2 de QRY_16). Els altres camps observacionals **conserven el domini 0/1/9**: «present parcial» i «desaparegut atestat» són categories morfològiques d'elements construïts, i forçar-les sobre processos (Looting) o vestigis mobles (Mat_Textiles) no significaria res.

**(b) El valor per defecte passa de 9 a 0.** Sota el domini nou, un registre nou comença de «examinat, res no hi és», i el 9 es reclama explícitament quan la posició no és examinable. El canvi optimitza la velocitat d'entrada (el farciment que exigeix la regla 4 apareix sol), i el risc epistèmic que comporta —zeros deixats per inèrcia que semànticament serien absències assertades— queda mitigat per tres mecanismes: el gating del formulari (els zeros de farciment els posa el sistema conscientment, secció 11), l'avís de col·lapse (regla 6) i la llista de treball de la regla 17 sobre els NULL heretats.

**(c) Els dos zeros (rev. 5).** La distinció que fa compatibles les regles 4 i 6 de la bateria:

- **0 d'asserció** — el sistema del camp és Present*/Attested lost, o el camp no pertany a cap sistema (I, R). Significa «la posició s'ha examinat i no hi havia res»; en una estructura col·lapsada és inverificable i la regla 6 el marca.
- **0 de farciment** — el sistema és Not applicable / Not observable / Absent. No asserta res, només evita un NULL, i la regla 4 l'exigeix activament.

Sense aquesta separació, una cova col·lapsada amb Sys_Portal = Not applicable faria saltar la regla 6 exactament sobre els zeros de Llindar/Brancals/Dintell que la regla 4 demana.

**(d) H, P i U mai no van ser elements: eren sistemes.** Corbelled_Platform, Access_Opening i Eave se substitueixen per **Sys_Platform** (E+F+G), **Sys_Portal** (N+O+Q) i **Sys_Eave** (S+T), camps TEXT(20) amb domini de sis valors: *Present complete / Present partial / Attested lost / Absent / Not applicable / Not observable*. La distinció «no aplicable» vs «absent verificat» és la mateixa lògica del 0 vs 9 elevada al nivell del sistema: sense ella, els percentatges publicats serien erronis. **Sys_Base** i **Sys_Chamber** s'hi afigen com a camps d'agrupació (domini de quatre valors: Present / Absent / Not applicable / Not observable) per al gating del formulari; no tenen lletra pròpia al vocabulari. En l'anàlisi de coocurrència, la promoció elimina la correlació trivial que H mantenia amb E/F/G per construcció.

**(e) Lintel desdoblat.** El camp v10 era un lookup de material que confonia presència amb tipus. Ara: **Lintel** BYTE (element Q, domini de cinc valors) + **Lintel_Material** TEXT(20) (*Stone / Wood / Mixed / ND*), el mateix patró que ja usaven Chamber_Roof / Chamber_Roof_Type. En la migració, els valors 'Absent' i 'ND' de material es netegen a NULL: quan la presència és 0 o 9, el material no és indeterminat sinó **inaplicable**.

**(f) Booleans de decoració eliminats.** Els 12 camps Dec_*/RA_*/Rock_Art duplicaven T_DECORATIONS —que a més registra posició, tipus, color i substrat— i amb només 35 registres els dos sistemes ja discrepaven en 6 casos. QRY_02 es reconstrueix sobre T_DECORATIONS. Relief_Frieze es manté: és un element arquitectònic (M), no una decoració. *La contrapartida v12 d'aquesta eliminació és Dec_Present: vegeu 1.2(a).*

**(g) T_LOST_ELEMENTS + L_LOST_EVIDENCE.** El valor 3 exigeix poder assenyalar una evidència física: encaix negatiu, forat de biga, cicatriu de despreniment, fragment despres in situ, empremta de morter, mènsules al buit, pigment sobre roca nua, murs truncats. El lookup tancat actua com a filtre metodològic: si cap evidència del catàleg no és aplicable, el valor correcte és 0 o 9. La taula absorbeix la funció de l'antic camp Lost_Body_Evidence (una fila amb abast *Body* o *Whole structure* cobreix tot el vocabulari de l'estructura d'una vegada) i constitueix la base evidencial d'OE3.

**(h) L_ELEMENTS.** El vocabulari A–X existia només com a noms de camp: no era referenciable ni exportable. Ara és un lookup de 24 files (codi, nom EN/VAL, nivell, sistema, camp associat, descripció) amb índex únic sobre Code, cosa que permet que T_LOST_ELEMENTS hi apunte amb integritat referencial.

**(i) Record_Class a L_TYPOLOGY.** Dels 35 registres actuals, 5 no són estructures: qualsevol percentatge sobre 35 és erroni. Cada tipologia porta una classe de registre — *Built funerary structure* (EA-MAU, EA-CAM, EA-PLA-R/V, Unclassifiable), *Natural funerary context* (NIX, CAV), *Structural trace* (MEN), *Rock art panel* (PR), *Pending classification*. La classe filtra cada anàlisi i governa el gating de pestanyes. **MIX s'elimina** (tot EA-CAM és mixt per definició: no distingia res) i **ND es desdobla** en *Unclassifiable* (un resultat: evidència insuficient per a classificar una estructura construïda) i *Not yet classified* (un estat de treball, exclòs de tota consulta analítica).

**(j) Renomenaments.** Lateral_Walls → **Return_Wall** (V: mur que retorna cap al penyal, perpendicular al pla de façana) i Lateral_Wall_Faces → **Facade_Flank** (L: parament dins del pla de façana, flanquejant l'obertura): els dos noms v10 deien «lateral» sobre coses distintes, i les dades confirmen que són variables independents. Recessed_Portal → **Recessed_Frame**; Rear_Wall_Type → **Rear_Closure_Type**. **Rear_Wall (W) desapareix com a camp**, redundant amb Sys_Chamber; l'entrada W es conserva a L_ELEMENTS com a referència de vocabulari. A L_STRUCT_BODY, l'entrada LWF passa a **FFL** (Facade flank).

## **1.2. Canvis de l'addendum v12**

**(a) Dec_Present (BYTE 0/1/9, per defecte 0).** L'eliminació dels booleans de decoració (1.1.f) va deixar l'**absència** de decoració irregistrable: «cap fila a T_DECORATIONS» no distingia l'absència verificada (0) de la no observable (9), una violació de l'axioma d'observabilitat sobre el qual es construeix tot el domini. Dec_Present restaura el judici agregat; el detall continua vivint **només** a T_DECORATIONS. Les regles 22–23 mantenen les dues coses d'acord, i al formulari Dec_Present governa el subformulari de decoració.

**(b) Regles 22–26 (QRY_16c).** Coherència Dec_Present ↔ files de T_DECORATIONS (22, 23); detall de revoc o pigment registrat sota presència 0 o 9 (24, 25); Mortar_Type incompatible amb Mortar_Present = 0 (26). Són la contrapartida a escala de corpus del gating v12 del formulari: **el formulari preveu en el moment d'entrada; la bateria detecta el que hi entra pel costat** (taules obertes a mà, importacions, gating desactivat al Centre de confiança).

**(c) Gating de tres nivells + emplenat ràpid (formulari).** Vegeu la secció 11.

# **2. T_STRUCTURES (120 camps)**

## **2.1. Identificació (8)**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| ID | COUNTER PK | Clau primària autonumèrica |
| Code | TEXT(20) NN | Codi PALP: [SITE][SECTOR]-[TIPTYPE][NUM] — ex: DW-S01-EA01 |
| ID_Sector | LONG NN | FK → L_SECTORS |
| ID_Typology | LONG | FK → L_TYPOLOGY. Porta Record_Class, que filtra les anàlisis i governa el gating de pestanyes |
| ID_Support | LONG | FK → L_SUPPORT. Classe de suport geomorfològic dominant |
| ID_Support_Secondary | LONG | FK → L_SUPPORT. Segon component en suports compostos (parella ordenada en lloc del valor «Combined») |
| ID_Parent | LONG | FK autoreferenciant → T_STRUCTURES.ID. Contenció física |
| ID_Group | LONG | FK → T_GROUPS. Agrupació funcional (alineament, xarxa) |

## **2.2. Morfologia i dimensions (11)**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| N_Basal_Bodies | INTEGER | Masses superposades **sense obertura** (nivell N0) |
| N_Chamber_Bodies | INTEGER | Cambres funeràries superposades (nivell N1). Criteri: un cos és N1 si conté (o contenia) obertura d'accés |
| Floor_Plan | TEXT(20) | Planta: Rectangular / Sub-rectangular / Square / Circular / Sub-circular / Trapezoidal / Irregular / ND |
| N_Built_Walls | INTEGER | Nombre de murs construïts (0–4) |
| Length_m / Width_m / Height_m | SINGLE | Dimensions exteriors (m). Segona passada mètrica |
| Height_Above_Base_m | SINGLE | Cota de posició sobre la base del faralló (m): és una cota, no una dimensió |
| Dim_Method | TEXT(30) | Photogrammetric model / Tape measure / Laser / Estimation / perimeter pigment outline / ND |
| Opening_Width_cm / Opening_Height_cm | SINGLE | Dimensions de l'obertura d'accés (cm). **No gatejades per Sys_Portal**: la passada mètrica és un exercici separat |

*L'antic Lost_Body_Evidence desapareix: la seua informació és ara una fila de T_LOST_ELEMENTS amb abast Body (secció 3.7). Lintel es trasllada al bloc del sistema portal (2.5).*

## **2.2b. Detall del suport geològic (3) — H02**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| Support_Width_cm | SINGLE | Amplària de la repisa o suport (cm) |
| Support_Depth_cm | SINGLE | Profunditat de la repisa o cavitat (cm) |
| Support_Modified | BYTE | 0/1/9. Modificació antròpica del suport natural. Sovint oculta darrere de la maçoneria |

## **2.2c. Maçoneria i morter (6) — H01/H04, cf. Toyne i Anzellini 2017**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| Masonry_Quality | TEXT(20) | Good / Moderate / Poor / ND |
| Masonry_Type | TEXT(30) | Well-coursed / Irregular-coursed / Uncoursed / Mixed / ND |
| Mortar_Present | BYTE | 0/1/9. 0 = fàbrica en sec |
| Mortar_Type | TEXT(30) | Mud / Mud with gravel / Mud with organics / None dry-laid / ND. Amb Mortar_Present = 0 el gating l'autoassigna a None dry-laid (regla 26) |
| Chinking_Stones | BYTE | 0/1/9. Ripio / falques entre carreus. Independent del morter: una fàbrica en sec pot dur falques |
| Mortar_Notes | TEXT(150) | Notes sobre morter i juntes |

## **2.3. Sistemes constructius (6 camps de sistema)**

**Domini dels tres sistemes d'element (Sys_Platform, Sys_Portal, Sys_Eave):** *Present complete / Present partial / Attested lost / Absent / Not applicable / Not observable*. **Domini dels dos conjunts d'agrupació (Sys_Base, Sys_Chamber):** *Present / Absent / Not applicable / Not observable*. Per defecte: Absent.

| **Camp** | **Elem.** | **Governa (gating)** | **Descripció** |
| --- | --- | --- | --- |
| Sys_Platform | **H** | E, F, G + Timber_Bracket_Count/Role, Platform_Surface_Material, Platform_Function | Sistema plataforma: E+F+G → H |
| Sys_Portal | **P** | N, O, Q + Lintel_Material, Recessed_Frame | Sistema portal: N+O+Q → P |
| Sys_Eave | **U** | S, T | Sistema ràfec: S+T → U. Sempre pedra a T i U |
| Sys_Base | — | A, B, C, D | Conjunt basal (agrupació sense lletra pròpia) |
| Sys_Chamber | — | J, K, L, M, V, X + Chamber_Roof_Type, Rear_Closure_Type | Conjunt cambra (agrupació; absorbeix l'antic W) |
| Platform_Function | — | — | Access / Circulation / Construction / Support / Multiple / Undetermined. Forma separada de funció: OE3 existeix per a determinar per a què servia la plataforma, i la morfologia no ho ha de pressuposar |

*Els sistemes es resolen primer (pestanya 11.Sist els presenta dalt de tot): es decideix que hi ha plataforma abans de comptar-ne les mènsules. Un sistema obri el seu grup només en Present* o Attested lost; sota Absent / Not applicable / Not observable els components porten el 0 de farciment.*

## **2.4. Elements A–X — Nivell 0 / basament (10)**

**Domini dels camps d'element: 0 = Absent / 1 = Present complet / 2 = Present parcial / 3 = Desaparegut atestat (exigeix T_LOST_ELEMENTS) / 9 = No observable. Per defecte 0.**

| **Camp** | **Tipus** | **Elem.** | **Descripció** |
| --- | --- | --- | --- |
| Embedded_Base_Beams | BYTE | **A** | Jàcenes horitzontals empotrades dins la maçoneria del basament |
| Base_Level | BYTE | **B** | Basament com a element constructiu diferenciat (pòdium) |
| Decorative_Socle | BYTE | **C** | Tractament decoratiu del basament |
| Tie_Walls | BYTE | **D** | Murets perpendiculars al faralló (ancoratge/compartimentació) |
| Timber_Brackets | BYTE | **E** | Mènsules de fusta empotrades a la roca (component de H). Criteri: perpendicular a la façana, encastada, en voladís. Exempta de la regla 1: és l'únic element amb existència independent del seu sistema |
| Timber_Bracket_Count | INTEGER | **E** | Nombre de mènsules visibles. Obligatori amb E present (regla 8): el nombre porta l'argument de la plataforma |
| Timber_Bracket_Role | TEXT(25) | **E** | Platform support / Isolated / Both / ND. Obligatori amb E present (regla 9). Isolated bloqueja Platform_Surface_Material (regla 7) |
| Transverse_Beams | BYTE | **F** | Bigues transversals (component de H). Criteri: paral·lela a la façana, salvant llum |
| Corbelled_Courses | BYTE | **G** | Filades de pedra en voladís creixent (component de H, variant lítia) |
| Platform_Surface_Material | TEXT(20) | (H) | Material de la superfície de la plataforma: Timber / Stone / Mixed / ND |

## **2.4b. Interfície N0/N1 i coronament (3) — elements I, R (sense sistema, deliberadament)**

| **Camp** | **Tipus** | **Elem.** | **Descripció** |
| --- | --- | --- | --- |
| Interbody_Cornice | BYTE | **I** | Cornisa intercòs entre cossos superposats |
| Interbody_Cornice_Material | TEXT(20) | **I** | Stone slabs / Wooden beams / Mixed / ND |
| Upper_Crown | BYTE | **R** | Coronament superior / coping |

## **2.5. Elements A–X — Nivell 1 / cos principal i sistema portal (9)**

| **Camp** | **Tipus** | **Elem.** | **Descripció** |
| --- | --- | --- | --- |
| Corner_Quoins | BYTE | **J** | Cantoneres: pedres de major mida disposades verticalment als angles |
| Structural_Pilasters | BYTE | **K** | Pilastres estructurals integrades al pla del parament, d'altura completa |
| Facade_Flank | BYTE | **L** | Parament dins del pla de façana, flanquejant l'obertura (abans Lateral_Wall_Faces) |
| Relief_Frieze | BYTE | **M** | Fris decoratiu en baix relleu (tractament de L) |
| Sill | BYTE | **N** | Llindar (component de P) |
| Jambs | BYTE | **O** | Brancals (component de P) |
| Lintel | BYTE | **Q** | Dintell (component de P). Desdoblat en v11: presència ací, material a banda |
| Lintel_Material | TEXT(20) | (Q) | Stone / Wood / Mixed / ND. NULL quan Q és 0 o 9: inaplicable, no indeterminat (regla 10) |
| Recessed_Frame | BYTE | — | 0/1/9. Marc reculat de façana (qualificador de P; abans Recessed_Portal) |

## **2.6. Elements A–X — Cambra i zona superior (6)**

| **Camp** | **Tipus** | **Elem.** | **Descripció** |
| --- | --- | --- | --- |
| Return_Wall | BYTE | **V** | Mur perpendicular al pla de façana, retornant cap al penyal (abans Lateral_Walls) |
| Rear_Closure_Type | TEXT(20) | — | Natural bedrock / Built masonry / Mixed / ND. Com es tanca la cambra pel darrere. L'element W desapareix com a camp (absorbit per Sys_Chamber); el tipus sobreviu perquè una cambra que usa la roca com a tancament posterior no es va construir com una que alça un mur |
| Eave_Beam | BYTE | **S** | Biga de suport del ràfec (component de U; pot ser fusta) |
| Eave_Surface | BYTE | **T** | Superfície del ràfec (component de U; sempre pedra) |
| Chamber_Roof | BYTE | **X** | Cambra tancada per dalt, amb la solució que siga |
| Chamber_Roof_Type | TEXT(25) | (X) | Natural bedrock / Built masonry / Built timber and slabs / Mixed / ND. La decisió que l'anàlisi necessita: roca natural vs obra (regla 14) |

## **2.7. Tractaments superficials (7)**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| Plaster_Present | BYTE | 0/1/9. Revoc / lluït present. Amb 0 o 9, el detall es buida i es bloqueja (regla 24) |
| Plaster_Color | TEXT(20) | White / Cream / Red / Ochre / Grey / ND |
| Plaster_Extent | TEXT(20) | Full facade / Partial / Traces only / ND |
| Pigment_Present | BYTE | 0/1/9. Pigment aplicat present. Amb 0 o 9, el detall es buida i es bloqueja (regla 25) |
| **Pigment_Substrate** | TEXT(20) | Plaster / Masonry stone / Bedrock / Mixed / ND. Variable clau: distingeix pintar sobre revoc preparat de pintar sobre el parament (regles 12–13). Al formulari, l'opció Plaster desapareix de la llista quan Plaster_Present = 0 |
| Pigment_Color | TEXT(20) | Red / White / Both / Ochre / ND |
| Pigment_Extent | TEXT(20) | Whole facade / Architectural elements / Decorative motifs / **Perimeter/threshold** / Traces / ND. El valor perimetral (v11) aïlla la pràctica de marcar el llindar independentment del tipus de suport (documentada a EA11 arrasada i EA09 cavitat natural) |

## **2.7b. Paisatge i orientació (2) — H06**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| Facade_Orientation | TEXT(5) | N / NE / E / SE / S / SW / W / NW / ND |
| Visibility_Valley | TEXT(10) | High / Medium / Low / ND |

## **2.8. Decoració (1) — el detall viu a T_DECORATIONS**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| Dec_Present | BYTE | NOU v12. 0/1/9. Judici agregat de decoració. 1 exigeix files a T_DECORATIONS (regla 22); files sense 1 es marquen (regla 23). Al formulari, 0 o 9 desactiven el subformulari de decoració |

## **2.9. Estat de conservació (6)**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| ID_Arch_Status | LONG | FK → L_STATUS (Good/Fair/Pre-collapse/Collapsed/ND). Amb Collapsed, cap 0 d'asserció és vàlid als elements (regla 6); el formulari mostra l'avís en roig |
| ID_Material_Status | LONG | FK → L_MATERIAL_STATUS (Good/Fair/Poor/Absent/ND). Absent exigeix Mat_* = 0 (regla 19) i el gating ho autoemplena |
| Looting / Fire_Damage / Animal_Activity / Modern_Access | BYTE | 0/1/9. Causes, independents dels graus |

## **2.10. Bioarqueologia (8) i materials culturals (7) — BYTE 0/1/9**

Idèntics a la versió anterior: Human_Remains (governa la resta del bloc bio; amb 1, MNI és obligatori — regla 18), MNI, Anatomical_Connection, Mummification, Funerary_Bundles, Dispersed_Remains, Flexed_Position, Bone_Burning; Mat_Textiles, Mat_Wood, Mat_VegFiber, Mat_Ceramics, Mat_Fauna, Mat_DeerAntler, Mat_Other. Les classes de registre *Structural trace* i *Rock art panel* tanquen els dos blocs (regles 15–16).

## **2.11. Cronologia (3) i fases constructives (2)**

C14 (YESNO, metadada de corpus), Chrono_Start_Cent, Chrono_End_Cent; Construction_Phases, Phase_Evidence (C14 / Stratigraphy / Superposition / Mortar / ND). Les fases sostenen l'argument morfològic-estructural de H03 sense dependre de datacions absolutes; la direccionalitat entre estructures adjacents es registra ara a T_CONNECTIONS.Chrono_Relation (secció 3.6).

## **2.12. Volumetria i àrea (5), coordenades (7), documentació i observabilitat (10)**

Sense canvis d'esquema respecte de la versió anterior: Interior_Area_m2, Interior_Vol_m3, Total_Vol_m3, ID_Vol_Method, Vol_Notes; Coord_Lat/Lon_WGS84, Coord_E/N_UTM, Altitude_masl, Coord_Precision_m, ID_Coord_Method; URL_Pano, URL_Pano_2, URL_Giga, URL_3D, ChaXR_Documented, ID_Campaign, Doc_Basis, Facade_Observability, Interior_Observability, Notes. C14 i ChaXR_Documented es mantenen YESNO: metadades del corpus, mai ambigües. Els tres camps d'observabilitat fan interpretables els 9 de tot el registre i es reporten a QRY_15.

# **3. Taules secundàries**

## **3.1–3.3. T_DATING, T_INDIVIDUALS, T_GROUPS**

Sense canvis d'esquema. T_DATING (mostra, data BP, intervals 1σ/2σ, laboratori, bibliografia); T_INDIVIDUALS (edat, sexe, preservació); T_GROUPS (Group_Code, ID_Sector, ID_Group_Type, N_Members, Notes).

## **3.4. T_DECORATIONS — registre únic de decoració (v11: sense duplicat de booleans)**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| ID | COUNTER PK |  |
| ID_Structure | LONG NN | FK → T_STRUCTURES (cascade delete) |
| ID_Struct_Body | LONG | FK → L_STRUCT_BODY. Posició a la façana |
| ID_Dec_Type | LONG | FK → L_DEC_TYPE |
| Body_No | INTEGER | Número del cos constructiu |
| Color | TEXT(20) | Red / White / Both / Ochre / None / ND |
| Substrate | TEXT(20) | Plaster / Masonry stone / Bedrock / ND. Resol els casos mixtos per posició |
| Notes | TEXT(255) |  |

*Des de v11 és l'únic registre de decoració; QRY_02 s'hi reconstrueix i QRY_02a en deriva les banderes per motiu quan cal el format ample. Dec_Present (v12) és el judici agregat 0/1/9 a T_STRUCTURES.*

## **3.5. T_ARCH_FEATURES — registre flexible**

Sense canvis: Feature_Code (llista amb «Other (see Notes)» com a via d'escapament del LimitToList), Present, Feature_Count, Material, Notes.

## **3.6. T_CONNECTIONS — connexions físiques entre estructures (OE3, H03)**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| ID | COUNTER PK |  |
| ID_Struct_A / ID_Struct_B | LONG NN | FK → T_STRUCTURES.ID (extrems de l'aresta) |
| Connection_Type | TEXT(30) | Llista tancada (v11): Abutted vertical joint / Superposition / Bonded joint / Shared support / Aerial connection / Other (see Notes). La llista tancada és el que fa computable la regla 20 |
| **Chrono_Relation** | TEXT(20) | NOU v11: A earlier than B / B earlier than A / Contemporary / Undetermined. Converteix el graf no dirigit en potencialment dirigit (H03). La direcció es llig de la junta: només una junta vertical adossada o una superposició la poden portar (regla 20); una junta travada implica Contemporary (el formulari l'autoassigna). L'ordre A/B no s'ha d'invertir després de l'entrada |
| Confidence | TEXT(10) | High / Medium / Low |
| Notes | TEXT(150) |  |

*El cas d'ús paradigmàtic: un complex amb junta vertical no travada es registra com a dues entrades amb codis PALP consecutius, enllaçades ací i agrupades a T_GROUPS; si el model 3D revela filades basals travades sota la junta aparent, correspon un registre únic amb Construction_Phases = 2.*

## **3.7. T_LOST_ELEMENTS — NOVA v11 — evidència del valor 3**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| ID | COUNTER PK |  |
| ID_Structure | LONG NN | FK → T_STRUCTURES (cascade delete) |
| Element_Code | TEXT(2) | FK → L_ELEMENTS.Code. Obligatori quan l'abast és Element (regla 21); el formulari només ofereix elements reals, no sistemes |
| ID_Evidence_Type | LONG NN | FK → L_LOST_EVIDENCE |
| Evidence_Scope | TEXT(20) | Element / Body / Whole structure. Una fila amb abast Body o Whole structure cobreix tot el vocabulari de l'estructura (absorbeix l'antic Lost_Body_Evidence) |
| ID_Position | LONG | FK → L_STRUCT_BODY |
| Notes | MEMO |  |

# **4. Taules lookup — valors**

*Tots els valors emmagatzemats són en anglés (convenció fixa). Les descripcions es tradueixen només a efectes de lectura.*

## **L_TYPOLOGY (reconstruïda v11: Record_Class, sense MIX, ND desdoblat)**

| **Nom** | **Record_Class** | **Descripció** |
| --- | --- | --- |
| EA-MAU Mausoleum/Chullpa | Built funerary structure | Estructura construïda (3+ murs + sostre artificial) sobre repisa. Predominant a La Petaca |
| EA-CAM Funerary Chamber | Built funerary structure | Cavitat natural tancada per una façana construïda. Predominant a Diablo Wasi |
| EA-PLA-R Ledge Platform | Built funerary structure | Plataforma constructiva sobre repisa natural |
| EA-PLA-V Aerial Platform | Built funerary structure | Plataforma artificial sobre bigues i lloses, sense repisa de suport |
| NIX Natural Niche | Natural funerary context | Petita cavitat natural (<1 m²). Ossari o enterrament secundari |
| CAV Cave/Cavern | Natural funerary context | Gran cavitat natural (>1 m²) amb ús funerari o ritual documentat |
| PR Rock Art | Rock art panel | Motiu pictòric sobre roca, documentat de forma independent |
| MEN Isolated Bracket | Structural trace | Element estructural aïllat. Evidència de xarxa de circulació aèria perduda |
| Unclassifiable | Built funerary structure | Evidència insuficient per a classificar una estructura CONSTRUÏDA (ex: EA11, perímetre de pigment sense construcció conservada). Els contextos naturals sempre són classificables com a NIX o CAV |
| Not yet classified | Pending classification | Estat de treball: pendent de revisió manual. Exclosa de tota consulta analítica |

*MIX s'ha eliminat: tot EA-CAM és mixt per definició, de manera que no distingia res. DW-S01-EA09 (cavitat amb mènsules i plataforma) es reclassifica MIX → CAV en la migració.*

## **L_SUPPORT (9 valors, v11: micro-repisa)**

Wide natural ledge (>2m) | Narrow natural ledge (<2m) | Artificial ledge | Large cavity (>10m2) | Medium cavity (1-10m2) | Natural niche (<1m2) | Fissure/Crack | ND | **Micro-ledge (<50cm)** (NOU v11, afegit al final per a no renumerar les FK existents). Els suports compostos es registren com a parella ordenada ID_Support + ID_Support_Secondary.

## **L_ELEMENTS — NOVA v11 — el vocabulari A–X com a lookup (24 files)**

Columnes: Code (únic, UQ_ELEM_CODE), Name_EN, Name_VAL, Level_Type (N0 / N0-N1 / N1 / SUP), Sys_Group (Platform / Portal / Eave / buit), Is_System, Field_Name, Description. Les files H, P i U porten Is_System = True i Field_Name = Sys_*; la fila W es conserva com a referència de vocabulari sense camp propi. Sys_Base i Sys_Chamber no apareixen a Sys_Group: són agrupacions sense lletra.

## **L_LOST_EVIDENCE — NOVA v11 — catàleg tancat d'evidències (9 files)**

| **Nom** | **Descripció** |
| --- | --- |
| Negative socket / impression | Encaix o empremta buida al parament |
| Beam hole | Forat passant per a biga |
| Break scar | Cicatriu de despreniment sobre maçoneria o roca |
| Detached fragment in situ | Fragment caigut identificable al peu de l'estructura |
| Mortar imprint | Empremta de morter sense l'element que unia |
| Corbels into void | Mènsules que ja no sostenen res |
| Pigment on bedrock | Pigment sobre roca nua de la maçoneria que el portava |
| Truncated walls | Murs truncats en pla net |
| Other (see Notes) | Detall al camp Notes |

## **L_STATUS | L_MATERIAL_STATUS | L_GROUP_TYPE | L_CAMPAIGN**

Sense canvis (Good/Fair/Pre-collapse/Collapsed/ND; Good/Fair/Poor/Absent/ND; els set tipus d'agrupació; les quatre campanyes 2013–2023).

## **L_STRUCT_BODY (11 entrades; v11: LWF → FFL)**

BAS, SOC (N0); **FFL** Facade flank (L) — abans LWF, mateix ID, les files de T_DECORATIONS no es toquen —, PIL, QUO, JAM, OVL, COR, CRO (N1); EAV (SUP); ND. Descriu només la posició dins d'un cos; el cos el diu Body_No.

## **L_DEC_TYPE (13 entrades)**

T-shaped niche (i inv.), L-shaped niche (i inv.), Zigzag, Stepped motif, Frieze / Greca, Triangular motif, Painted band, Square niche, Plain colour field, Decapitation scene, ND. «Plain colour field» converteix T_DECORATIONS en el registre general de tractament de superfície per posició.

# **5. Relacions (25)**

Les 21 de la versió anterior més les quatre de T_LOST_ELEMENTS:

| **Nom** | **Pare** | **Filla** | **Camp fill** | **Notes** |
| --- | --- | --- | --- | --- |
| REL_STR_LOST | T_STRUCTURES | T_LOST_ELEMENTS | ID_Structure | Cascade delete |
| REL_ELEM_LOST | L_ELEMENTS (Code) | T_LOST_ELEMENTS | Element_Code | Requereix l'índex únic UQ_ELEM_CODE |
| REL_LEV_LOST | L_LOST_EVIDENCE | T_LOST_ELEMENTS | ID_Evidence_Type | Update cascade |
| REL_SB_LOST | L_STRUCT_BODY | T_LOST_ELEMENTS | ID_Position | Update cascade |

*REL_STR_SELF (autoreferenciant) i les dues de T_CONNECTIONS continuen amb dbRelationDontEnforceIntegrity (valor numèric 2).*

# **6. Consultes SQL (26)**

| **Nom** | **Descripció** | **Hip.** |
| --- | --- | --- |
| QRY_01_Typology_by_Site | Distribució de tipologies per jaciment i sector. | H01, H04 |
| QRY_02s_Decoration_Typed | Auxiliar: files de T_DECORATIONS amb tipus i posició resolts. | — |
| QRY_02a_Decoration_Flags | Banderes per motiu derivades de T_DECORATIONS (format ample per a khi-quadrat). | H01, H04 |
| QRY_02_Decoration_by_Site | Reconstruïda v11 sobre T_DECORATIONS: recomptes per motiu i jaciment. | H01, H04 |
| QRY_03 – QRY_12 | Sense canvis de funció: conservació, C14, exportació R, volumetria, exportació QGIS, fills, membres de grup, cobertura ChaXR, maçoneria, geologia-construcció. QRY_05 i QRY_07 actualitzades v11 als noms nous i als camps de sistema. | diverses |
| QRY_13_AX_Pattern_Export | Reconstruïda v11: **20 columnes AX_ + 5 columnes SYS_** amb semàntica NA (un component gatejat per un sistema Not applicable s'exporta NA, no 0, perquè el seu 0 és de farciment). La font única FillElementMap garanteix que exportació i validació no divergeixen mai. | H01, H04, H05 |
| QRY_14_Connections_Edges | Llista d'arestes amb Chrono_Relation: graf potencialment dirigit per a igraph/QGIS. | OE3, H03 |
| QRY_15_Observability_Bias | Recompte per jaciment i sector de base documental i observabilitat. | OE1 |
| QRY_16s_* (5) | Auxiliars de la bateria: Element_Values (format llarg dels 20 elements: la clau que evita ~80 branques UNION), Lost_Cover, Damage_Count, Observational_Nulls (llista de treball camp a camp), Null_Count. | — |
| QRY_16a / QRY_16b / QRY_16c | Regles 1–11, 12–21 i 22–26. Emmagatzemades en tres parts perquè una UNION única de 31 branques supera el límit «query too complex» de JET. | — |
| QRY_16_Validation_Check | Unió de les tres parts: **26 regles**. Resultat buit = corpus coherent. | — |

## **6.1. La bateria de 26 regles**

No bloquejant per disseny: una restricció dura impediria registrar una absència genuïnament observada en una estructura parcialment col·lapsada. Resum:

| # | Regla | # | Regla |
| --- | --- | --- | --- |
| 1 | Component present (1/2/3) amb sistema no present (exempts: Not observable, i Timber_Brackets) | 14 | Coberta present sense tipus |
| 2 | Valor 3 (o sistema Attested lost) sense fila d'evidència a T_LOST_ELEMENTS | 15 | Structural trace amb restes humanes |
| 3 | Sistema present amb tots els components a 0 | 16 | Contingut bio/materials en classes que tanquen els blocs |
| 4 | Component ≠ 0 sota sistema Not applicable | 17 | Camps observacionals encara NULL (llista de treball, no error) |
| 5 | Estat Good amb >30% d'elements parcials o perduts | 18 | Restes humanes sense MNI |
| 6 | 0 d'asserció en estructura Collapsed | 19 | Vestigis Absent amb Mat_* ≠ 0 |
| 7 | Material de plataforma amb rol Isolated | 20 | Cronologia direccional en connexió sense adossament |
| 8 | Mènsules presents sense recompte | 21 | Abast Element sense Element_Code |
| 9 | Mènsules presents sense rol | 22 | Dec_Present = 1 sense files de decoració (v12) |
| 10 | Dintell 0 amb material | 23 | Files de decoració sense Dec_Present = 1 (v12) |
| 11 | Cossos de cambra amb Sys_Portal Absent | 24 | Detall de revoc sota presència 0/9 (v12) |
| 12 | Pigment present sense substrat | 25 | Detall de pigment sota presència 0/9 (v12) |
| 13 | Pigment sobre revoc amb revoc absent | 26 | Mortar_Type incompatible amb Mortar_Present = 0 (v12) |

*L'exempció de Timber_Brackets a la regla 1 mereix nota: E és l'únic element del vocabulari amb existència independent del seu sistema — una mènsula aïllada no ha d'haver portat cap plataforma. Per això existeixen Timber_Bracket_Role i la tipologia MEN. La taula 4.5 llista E sota Sys_Platform per al GATING, que és una pregunta distinta de si E implica H.*

# **7. Ús estadístic dels dominis — advertència operativa**

De lectura obligada abans de qualsevol prova estadística.

## **7.1. El problema del denominador (es manté)**

Prendre el total d'estructures com a denominador infla la mostra amb casos mai avaluats, i com que la conservació difereix entre LP i DW, la inflació no és igual als dos jaciments: la prova mesuraria preservació diferencial i es llegiria com a pràctica diferencial. **Denominador vàlid = presents + absències verificades**; el total es conserva només per a reportar la cobertura, que s'ha de fer constar junt amb la n.

## **7.2. El domini de cinc valors en R**

La conversió prèvia a tota anàlisi de la matriu de QRY_13:

```r
ax <- read.csv("QRY_13_AX_Pattern_Export.csv")
axcols <- grep("^AX_", names(ax), value = TRUE)

# 9 -> NA sempre (mai 0). Els NA de la propia exportacio
# (components gatejats per un sistema Not applicable) ja arriben buits.
ax[axcols] <- lapply(ax[axcols], function(x) ifelse(x == 9, NA, x))

# Tres binaritzacions legitimes; declareu quina responeu:
built    <- lapply(ax[axcols], function(x) as.integer(x %in% c(1,2,3)))  # que es va construir
survives <- lapply(ax[axcols], function(x) as.integer(x %in% c(1,2)))    # que sobreviu
intact   <- lapply(ax[axcols], function(x) as.integer(x == 1))           # que resta intacte
```

Per a les hipòtesis constructives (H01, H04, H05) la binarització pertinent és normalment *built*: el gest constructiu existí encara que l'element haja desaparegut — és exactament el que el valor 3 registra i el que T_LOST_ELEMENTS evidencia. *Survives* i *intact* responen preguntes tafonòmiques i de conservació. Barrejar-les en una mateixa prova reintrodueix el biaix que el domini existeix per a eliminar.

## **7.3. Les columnes SYS_ i els zeros de farciment**

QRY_13 exporta els cinc sistemes com a columnes pròpies. Un component el sistema del qual és *Not applicable* arriba com a NA d'origen: el seu 0 emmagatzemat és de farciment i no ha d'entrar en cap recompte. Les dues estratègies de mostra (restricció a Facade_Observability = Complete, o distàncies binàries amb tractament de nuls) es mantenen; QRY_15 proporciona el recompte que justifica la decisió.

## **7.4. Regla general**

> Un 9 no és ni un 0 ni un 1: és una cel·la buida. Un 0 de farciment tampoc no és un 0 d'asserció: és un NA estructural. Qualsevol operació que els convertisca implícitament en dades reintrodueix el biaix al resultat.

# **8. Vocabulari arquitectònic normalitzat A–X (v11)**

24 entrades: 20 elements amb camp propi, 3 sistemes (H, P, U → Sys_*) i una entrada de referència (W, absorbida per Sys_Chamber). Els criteris d'ordenació (nivell constructiu bottom-to-top, alçada, funció) i la nota terminològica nivell/cos es mantenen: N0/N1/Superior són categories funcionals, mai numeració; un cos és N1 si conté (o contenia) obertura d'accés.

| **Ll.** | **Valencià** | **Nivell** | **Sistema** | **Camp v12** |
| --- | --- | --- | --- | --- |
| A | Jàcenes basals empotrades | N0 | — (conjunt basal) | Embedded_Base_Beams |
| B | Basament | N0 | — (conjunt basal) | Base_Level |
| C | Sòcol decoratiu | N0 | — (conjunt basal) | Decorative_Socle |
| D | Muret transversal | N0 | — (conjunt basal) | Tie_Walls |
| E | Mènsules de fusta | N0 | Component de H | Timber_Brackets |
| F | Bigues transversals | N0 | Component de H | Transverse_Beams |
| G | Filades en voladís | N0 | Component de H | Corbelled_Courses |
| **H** | **Plataforma en voladís** | N0 | **SISTEMA E+F+G** | **Sys_Platform** |
| I | Cornisa intercòs | N0/N1 | Cap (deliberadament) | Interbody_Cornice |
| J | Cantoneres | N1 | — (conjunt cambra) | Corner_Quoins |
| K | Pilastres estructurals | N1 | — (conjunt cambra) | Structural_Pilasters |
| L | Ala de façana | N1 | — (conjunt cambra) | Facade_Flank |
| M | Fris en relleu | N1 | — (conjunt cambra) | Relief_Frieze |
| N | Llindar | N1 | Component de P | Sill |
| O | Brancals | N1 | Component de P | Jambs |
| **P** | **Sistema portal** | N1 | **SISTEMA N+O+Q** | **Sys_Portal** |
| Q | Dintell | N1 | Component de P | Lintel (+ Lintel_Material) |
| R | Coronament | N1 | Cap (deliberadament) | Upper_Crown |
| S | Biga de suport del ràfec | Sup. | Component de U | Eave_Beam |
| T | Superfície del ràfec | Sup. | Component de U (PEDRA) | Eave_Surface |
| **U** | **Ràfec en voladís** | Sup. | **SISTEMA S+T (PEDRA)** | **Sys_Eave** |
| V | Mur de retorn | N1 | — (conjunt cambra) | Return_Wall |
| W | Mur posterior | N1 | Referència de vocabulari | (absorbit per Sys_Chamber; tipus a Rear_Closure_Type) |
| X | Coberta de la cambra | N1 | — (conjunt cambra) | Chamber_Roof (+ Chamber_Roof_Type) |

**Criteris operatius de replicabilitat (fixats per escrit):** pilastra (K) = integrada al pla del parament; contrafort = sobreïx del pla (cap documentat). Mènsula (E) = perpendicular a la façana, encastada, en voladís; biga (F) = paral·lela, salvant llum. Suport de plataforma vs mènsula aïllada = sosté (o sostenia) la plataforma o no. Ala de façana (L) = parament dins del pla de façana; mur de retorn (V) = perpendicular, cap al penyal — les dades confirmen que són variables independents.

# **9. Formulari i entrada de dades (v12)**

## **9.1. Estructura**

F_STRUCTURES, 12 pestanyes: 1.Id. (identificació + detall del suport), 2.Arq. (morfologia, façana/paisatge, maçoneria i morter, fases), 3.Acab. (revoc + pigment), 4.Dec. (**Dec_Present** + subformulari F_DECORATIONS), 5.Estat (conservació + base documental i observabilitat), 6.Bio., 7.Mat., 8.Cron., 9.Mètr. (dimensions, obertura, mètrica del suport, volumetria, coordenades — segona passada), 10.Doc., 11.Sist. (els cinc sistemes dalt de tot i els grups de components a continuació, amb l'avís de col·lapse), 12.Extra (subformularis F_ARCH_FEATURES, F_CONNECTIONS i **F_LOST_ELEMENTS**).

Tots els combos de domini són de dues columnes: la columna emmagatzemada (anglés) està oculta i l'etiqueta (valencià) és visible, de manera que reanomenar etiquetes mai no toca les dades. LimitToList és sempre cert (Access ho imposa amb la columna lligada oculta); la via d'escapament per a excepcions genuïnes és el valor «Other (see Notes)» + el camp Notes.

## **9.2. Gating de tres nivells**

- **Nivell 1 — Record_Class → pestanyes.** *Natural funerary context* tanca 2.Arq; *Structural trace* i *Rock art panel* tanquen a més 6.Bio i 7.Mat; *Rock art panel* tanca també 11.Sist. Un registre sense classificar ho manté tot obert. Matís deliberat: una cova (CAV) conserva 11.Sist obert — DW-S01-EA09 té mènsules i plataforma dins d'una cavitat natural. 3.Acab resta oberta per a totes les classes: un panell d'art rupestre es defineix pel seu pigment.
- **Nivell 2 — portes → grups.** Cada Sys_* obri el seu grup només en Present* o Attested lost. ID_Material_Status (Absent/ND) tanca 7.Mat; Human_Remains (0/9) tanca el detall bio; Plaster_Present, Pigment_Present i Mortar_Present tanquen els seus detalls; Dec_Present tanca el subformulari de decoració.
- **Nivell 3 — element → detall.** Timber_Bracket_Count i Timber_Bracket_Role només amb E a 1/2/3; Platform_Surface_Material bloquejat amb rol Isolated; Lintel_Material només amb Q a 1/2/3; Chamber_Roof_Type només amb X a 1/2/3; Interbody_Cornice_Material només amb I a 1/2/3. Les regles 7–10 i 14 passen de detectades a impossibles.

## **9.3. Emplenat ràpid amb confirmació**

En tancar una porta, el formulari ofereix escriure el valor coherent a tot el grup: sistema a Absent/Not applicable → components a 0 (farciment) i detall a buit; Not observable → components a 9; revoc o pigment a 0/9 → detall a buit; vestigis Absent → Mat_* a 0 (exactament el que exigeix la regla 19); vestigis ND → Mat_* a 9; restes humanes 0/9 → detall bio a 0/9 i MNI a buit; Mortar_Present 0 → Mortar_Type = None dry-laid (silenciós, regla 26). **Mai no s'escriu sense un Sí explícit, i mai en navegar entre registres** (Form_Current només activa i desactiva: l'anàleg de la regla B de la migració). L'assignació per codi no dispara AfterUpdate, de manera que no hi ha recursió.

## **9.4. Avisos i validacions en viu**

Marcar un 3 en qualsevol element mostra el recordatori de la regla 2 (evidència a 12.Extra). Amb estat Collapsed apareix l'avís en roig de la regla 6 a 11.Sist. El substrat de pigment amaga l'opció Plaster quan Plaster_Present = 0 (prevenció de R13). Als subformularis —on Enabled afectaria la columna sencera de totes les files— la lògica per fila és validació BeforeUpdate: F_CONNECTIONS bloqueja una cronologia direccional sobre una connexió sense adossament (R20) i autoassigna Contemporary a la junta travada; F_LOST_ELEMENTS exigeix el codi d'element quan l'abast és Element (R21).

## **9.5. Prevenció i detecció**

El principi de disseny: **el formulari preveu en el moment d'entrada; QRY_16 detecta a escala de corpus.** La bateria continua sent imprescindible perquè el gating es pot esquivar (taules obertes a mà, importacions, Centre de confiança desactivat), i perquè algunes regles són intrínsecament de corpus (5, 17). Executeu QRY_16_Validation_Check periòdicament durant l'entrada de dades.

## **9.6. Requisit tècnic**

La injecció dels mòduls VBA del gating i de les validacions exigeix «Confiar en l'accés al model d'objectes de projectes VBA» al Centre de confiança d'Access. Si està desactivat, el formulari es construeix igualment i tot queda editable; només falta l'automatisme, i l'script ho avisa a la finestra immediata en lloc de fallar en silenci.

# **10. Scripts i rutes d'actualització**

| **Fitxer** | **Sub públic** | **Ús** |
| --- | --- | --- |
| chachapoya_DB_v12.bas | BuildDB() | Construcció completa sobre una BD EN BLANC: 22 taules, valors per defecte (45 BYTE a 0, sistemes a Absent), lookups, 25 relacions, 26 consultes |
| chachapoya_Form_v12_val.bas | BuildForm() | Després de BuildDB(): 4 subformularis, F_STRUCTURES amb 12 pestanyes, combos de dues columnes, captions DAO, gating v12 i validacions de subformulari |
| migrate_v10_to_v11.bas | MigrateV10toV11() | Sobre la BD v10 AMB DADES: transforma l'esquema preservant cada registre (ordre estricte rename→add→populate→delete; regla B: cap NULL existent es toca; esborrats finals barrats per ALLOW_FINAL_DELETIONS + precondicions vives) |
| upgrade_form_v12.bas | UpgradeV12() | Sobre la BD v11 AMB DADES: afig Dec_Present i les regles 22–26, crea F_LOST_ELEMENTS i injecta el mateix gating v12 que BuildForm, sense reconstruir res. Idempotent |

Les restriccions tècniques de VBA/JET es mantenen: cap continuació de línia, ASCII estricte en codi (cp1252), patró sql = sql & "…", paraules reservades evitades, N−1 parèntesis per a N taules en JOIN, Private per a tot auxiliar, DROP previ per a objectes regenerables, combos configurats per ControlSource.
