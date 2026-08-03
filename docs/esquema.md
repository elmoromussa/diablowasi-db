**ESQUEMA DE LA BASE DE DADES ARQUEOLÒGICA**

*La Petaca i Diablo Wasi (Leymebamba, Amazonas, Perú)*

chachapoya_DB_v4.bas + chachapoya_Form_v3_val.bas

Sub BuildDB() + Sub BuildForm() | Microsoft Access JET SQL | Esteve Ribera Torró

*Versió 5 del document — actualitzada segons el codi v4 de la BD (agost 2026)*

# **1. Resum general**

| **Element** | **Valor** |
| --- | --- |
| Taules principals | T_STRUCTURES (124 camps), T_DATING, T_INDIVIDUALS, T_GROUPS, T_DECORATIONS, T_ARCH_FEATURES, T_CONNECTIONS |
| Taules lookup | L_SITES, L_SECTORS, L_TYPOLOGY, L_SUPPORT, L_STATUS, L_MATERIAL_STATUS, L_VOL_METHOD, L_COORD_METHOD, L_GROUP_TYPE, L_CAMPAIGN, L_STRUCT_BODY, L_DEC_TYPE |
| Total taules | 19 |
| Total relacions | 20 (inclou l'autoreferenciant de T_STRUCTURES i les dues de T_CONNECTIONS) |
| Consultes SQL | 14 (QRY_01 a QRY_14; QRY_02, QRY_05, QRY_07 i QRY_13 actualitzades en v4) |
| Formulari | F_STRUCTURES - 12 pestanyes + subformularis F_DECORATIONS i F_ARCH_FEATURES. Etiquetes UI en valencià; valors emmagatzemats en anglés |
| Idioma BD | Anglés (noms de taules, camps i valors lookup). UI del formulari en valencià. Capçaleres de subformulari via etiquetes adjuntes; captions DAO com a reforç |
| Motor | Microsoft Access JET SQL / ACE │ cp1252 |
| Jaciments | La Petaca (WGS84: Lat -6.8311, Lon -77.8084) │ Diablo Wasi (Lat -6.8475, Lon -77.8154) |

**Canvis respecte de la versió anterior de l'esquema (v4 del document, que descrivia la BD v3):**

(a) **Domini 0/1/9 per als elements A–X.** Els 21 camps de presència del vocabulari A–X passen de YESNO a BYTE amb domini tancat: 0 = Absent (decisió constructiva), 1 = Present, 9 = ND / no observable (col·lapse o biaix documental). El valor per defecte és 9, fixat via DAO: l'absència ha de registrar-se positivament, mai per omissió. Aquesta distinció separa l'elecció tècnica del biaix tafonòmic, condició necessària per a la validesa de la coocurrència, el clúster i la I de Moran (el col·lapse no és aleatori: afecta més els elements superiors R, S, T, U, X que els basals).

(b) **Vocabulari A–X complet a la BD.** S'afigen els camps Lateral_Wall_Faces (element L) i Relief_Frieze (element M, que substitueix el booleà Dec_Frieze). L'element Q (dintell) es deriva del camp de material Lintel a QRY_13 (Absent→0, ND/Null→9, resta→1). Amb això, els 24 elements són exportables com a matriu completa.

(c) **Terminologia cos/nivell.** «Nivell» queda reservat per als nivells constructius del vocabulari (N0 / N1 / Superior); «cos» designa els pisos superposats (N_Floors, Body_No, L_STRUCT_BODY). En conseqüència: Interlevel_Cornice → **Interbody_Cornice** i Cornice_Material → **Interbody_Cornice_Material** (cornisa intercòs, element I). L'element H (plataforma d'accés) es reclassifica com a **interfície N0/N1**, igual que I: tanca la seqüència de N0 i és la precondició de N1, però es distingeix d'I per funció (superfície de circulació i accés vs marcatge entre cossos), per composició (H és sistema E+F+G→H; I és standalone) i per posició en la *chaîne opératoire*.

(d) **Bloc de morter i falques** (H04, cf. Toyne i Anzellini 2017): Mortar_Present (BYTE 0/1/9), Mortar_Type, Chinking_Stones (BYTE 0/1/9), Mortar_Notes. El morter ja apareixia com a valor de Phase_Evidence però no es registrava com a atribut. No confondre amb Plastered (revoc: acabat superficial, operació distinta de la cadena).

(e) **Mètriques noves**: Opening_Width_cm i Opening_Height_cm (dimensions de l'obertura d'accés P: connecten tipologia i pràctica funerària — capacitat per a farcells i persones); Dim_Method (mètode de mesura de L/W/H); Approx_Height_m → **Height_Above_Base_m** (cota de posició sobre la base del faralló, no una dimensió de l'estructura).

(f) **T_CONNECTIONS** (nova taula) + QRY_14_Connections_Edges: registre de connexions físiques entre estructures (repisa compartida, plataforma contínua, mur o biga compartits) per a construir la matriu d'adjacència de la xarxa de circulació aèria (OE3). ID_Parent (contenció) i ID_Group (agrupació funcional) no capturaven l'adjacència física.

(g) **Formulari v3 (12 pestanyes)**: absorbeix el patch mètric (chachapoya_patch_metric.bas queda obsolet); lletres A–X corregides a la pestanya 11.Sist; camps d'art rupestre reubicats sota la seua capçalera (abans dos camps quedaven tapats pel subformulari i eren inaccessibles); etiquetes senceres (LW 1900→2600 twips; C2 4880→5600); etiquetes adjuntes als controls dels subformularis perquè la vista full de dades mostre capçaleres en valencià; combos de domini tancat per a Floor_Plan, Lintel, Access_Orientation, Dim_Method i Mortar_Type; etiquetes d'estat sense sigles (I)/(M) per a evitar col·lisió amb les lletres del vocabulari.

# **2. T_STRUCTURES (124 camps)**

## **2.1. Identificació (7)**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| ID | COUNTER PK | Clau primària autonumèrica |
| Code | TEXT(20) NN | Codi PALP: [SITE][SECTOR]-[TIPTYPE][NUM] - ex: DWS1-EF01, PTC-SS-EF18 |
| ID_Sector | LONG NN | FK -> L_SECTORS |
| ID_Typology | LONG | FK -> L_TYPOLOGY |
| ID_Support | LONG | FK -> L_SUPPORT. Suport geomorfològic |
| ID_Parent | LONG | FK autoreferenciant -> T_STRUCTURES.ID. Contenció física |
| ID_Group | LONG | FK -> T_GROUPS. Agrupació funcional (alineament, xarxa) |

## **2.2. Morfologia i dimensions (15)**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| N_Floors | INTEGER | Nombre de cossos (pisos) constructius superposats |
| Floor_Plan | TEXT(20) | Planta (llista al formulari): Rectangular / Sub-rectangular / Square / Circular / Sub-circular / Trapezoidal / Irregular / ND |
| N_Built_Walls | INTEGER | Nombre de murs construïts (0-4) |
| Length_m | SINGLE | Llarg exterior (m) |
| Width_m | SINGLE | Ample exterior (m) |
| Height_m | SINGLE | Alçada total de l'estructura (m) |
| Height_Above_Base_m | SINGLE | Cota de posició sobre la base del faralló (m). RENOMENAT (abans Approx_Height_m): és una cota, no una dimensió |
| Dim_Method | TEXT(30) | NOU v4. Mètode de mesura de les dimensions: Photogrammetric model / Tape measure / Laser / Estimation / ND |
| Opening_Width_cm | SINGLE | NOU v4. Amplada de l'obertura d'accés (cm) |
| Opening_Height_cm | SINGLE | NOU v4. Alçada de l'obertura d'accés (cm) |
| Access_Orientation | TEXT(5) | Orientació del vano (llista al formulari): N / NE / E / SE / S / SW / W / NW / ND. Si divergeix de Facade_Orientation, documentar el criteri a Notes |
| Lintel | TEXT(20) | Material del dintell (**element Q**; llista al formulari): Stone / Wood / Mixed / Absent / ND. AX_Q es deriva d'aquest camp a QRY_13 |
| Natural_Roof | YESNO | Sostre natural de roca (el faralló actua com a sostre) |
| Buttresses | YESNO | Contraforts a la façana |
| Wooden_Stakes | YESNO | Estaques o pals de fusta verticals (elements estructurals d'ancoratge) |

## **2.2b. Detall del suport geològic (4) — H02: geologia com a factor determinant**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| Support_Width_cm | SINGLE | Amplària de la repisa o suport (cm). Mesura directa o del model 3D |
| Support_Depth_cm | SINGLE | Profunditat de la repisa o cavitat (cm) |
| Support_Morphology | TEXT(30) | Morfologia del suport (llista al formulari): Flat ledge / Concave ledge / Fissure / Small cavity / Medium cavity / Large cavity / Vertical no support / ND |
| Support_Modified | YESNO | Evidència de modificació antròpica del suport natural (retall, anivellament) |

*Aquests camps quantifiquen la relació entre les dimensions del suport natural i les decisions constructives, nucli empíric de la hipòtesi H02.*

## **2.2c. Maçoneria i morter (6) — H01/H04, cf. Toyne i Anzellini 2017**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| Masonry_Quality | TEXT(20) | Qualitat d'execució (llista al formulari): Good / Moderate / Poor / ND |
| Masonry_Type | TEXT(30) | Tipus d'aparell (llista al formulari): Well-coursed / Irregular-coursed / Uncoursed / Mixed / ND |
| Mortar_Present | BYTE | NOU v4. 0=Absent (fàbrica en sec) / 1=Present / 9=ND. Per defecte 9 |
| Mortar_Type | TEXT(30) | NOU v4 (llista al formulari): Mud / Mud with gravel / Mud with organics / None dry-laid / ND |
| Chinking_Stones | BYTE | NOU v4. Ripio / pedres de falca entre carreus: 0/1/9. Discriminador de qualitat de fàbrica (T&A 2017) |
| Mortar_Notes | TEXT(150) | NOU v4. Notes sobre morter i juntes |

*El morter és evidència potencial de fases (valor «Mortar» de Phase_Evidence) i atribut de la fàbrica. No confondre amb Plastered (revoc, acabat superficial: operació distinta de la cadena operativa).*

## **2.3. Sistemes A-X — Nivell 0 / basament (10) — elements A-H**

**Convenció de domini per a tots els camps BYTE del vocabulari A-X: 0 = Absent (decisió constructiva) / 1 = Present / 9 = ND, no observable (per defecte).**

| **Camp** | **Tipus** | **Elem.** | **Descripció** |
| --- | --- | --- | --- |
| Embedded_Base_Beams | BYTE | **A** | Jàcenes horitzontals empotrades dins la maçoneria del basament |
| Base_Level | BYTE | **B** | Basament com a element constructiu diferenciat (pòdium) |
| Decorative_Socle | BYTE | **C** | Tractament decoratiu del basament (sòcol decoratiu) |
| Tie_Walls | BYTE | **D** | Murets perpendiculars al faralló (ancoratge/compartimentació) |
| Timber_Brackets | BYTE | **E** | Mènsules de fusta empotrades a la roca (component de H). Criteri operatiu: element perpendicular a la façana, encastat, en voladís |
| Timber_Bracket_Count | INTEGER | **E** | Nombre de mènsules visibles |
| Transverse_Beams | BYTE | **F** | Bigues transversals (component de H). Criteri operatiu: element paral·lel a la façana, salvant llum entre suports |
| Corbelled_Courses | BYTE | **G** | Filades de pedra en voladís creixent (component de H, variant lítia) |
| Corbel_Material | TEXT(20) | **E/G** | Material: Timber / Stone / Mixed / ND |
| Corbelled_Platform | BYTE | **H** | Plataforma volada d'accés. SISTEMA: E+F+G -> H. Nivell: interfície N0/N1 (tanca N0, precondició de N1) |

## **2.3b. Interfície N0/N1 (2) — element I**

| **Camp** | **Tipus** | **Elem.** | **Descripció** |
| --- | --- | --- | --- |
| Interbody_Cornice | BYTE | **I** | Cornisa intercòs entre cossos superposats. RENOMENAT (abans Interlevel_Cornice): «cos» per als pisos, «nivell» per a N0/N1/Sup |
| Interbody_Cornice_Material | TEXT(20) | **I** | Material (llista al formulari): Stone slabs / Wooden beams / Mixed / ND. RENOMENAT (abans Cornice_Material) |

## **2.4. Sistemes A-X — Nivell 1 / cos principal (12) — elements J-P, R, V, W**

| **Camp** | **Tipus** | **Elem.** | **Descripció** |
| --- | --- | --- | --- |
| Corner_Quoins | BYTE | **J** | Cantoneres: pedres de major mida disposades verticalment als angles |
| Structural_Pilasters | BYTE | **K** | Pilastres estructurals que flanquegen la façana (altura completa) |
| Lateral_Wall_Faces | BYTE | **L** | NOU v4. Paraments laterals del cos principal (fora del marc del portal) |
| Relief_Frieze | BYTE | **M** | NOU v4. Fris decoratiu en baix relleu (tractament de L). Substitueix el booleà Dec_Frieze |
| Sill | BYTE | **N** | Llindar / threshold (component de P) |
| Jambs | BYTE | **O** | Brancals: elements verticals formals del portal (component de P) |
| Access_Opening | BYTE | **P** | Obertura d'accés present. SISTEMA: N+O+Q -> P |
| Recessed_Portal | BYTE | **-** | Portal retranquejat de façana (qualificador de P) |
| Upper_Crown | BYTE | **R** | Coronament superior / coping |
| Lateral_Walls | BYTE | **V** | Murs laterals que donen profunditat a la cambra funerària |
| Rear_Wall | BYTE | **W** | Mur posterior present |
| Rear_Wall_Built | BYTE | **W** | Qualificador de W: construït (1) vs roca natural (0) / ND (9) |

## **2.5. Sistemes A-X — Zona superior (4) — elements S, T, U, X**

| **Camp** | **Tipus** | **Elem.** | **Descripció** |
| --- | --- | --- | --- |
| Eave_Beam | BYTE | **S** | Biga de suport del ràfec (component de U; pot ser fusta) |
| Eave_Surface | BYTE | **T** | Superfície del ràfec (component de U; sempre pedra - lloses) |
| Eave | BYTE | **U** | Ràfec-voladís / visera. SISTEMA: S+T -> U. Sempre pedra |
| Chamber_Roof | BYTE | **X** | Coberta construïda de la cambra (distinta de Natural_Roof i Eave) |

## **2.6. Acabats superficials (4)**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| Plastered | YESNO | Arrebossat / lluïment present |
| Plaster_Color | TEXT(20) | Color de l'arrebossat (text lliure) |
| Rock_Painting | YESNO | Pintura sobre la roca present |
| Rock_Paint_Color | TEXT(20) | Color de la pintura sobre roca (text lliure) |

## **2.6b. Paisatge i orientació (2) — observacional; anàlisi QGIS posterior**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| Facade_Orientation | TEXT(5) | Orientació de la façana (pot diferir d'Access_Orientation): N / NE / E / SE / S / SW / W / NW / ND |
| Visibility_Valley | TEXT(10) | Visibilitat des del fons de la vall (estimació de camp): High / Medium / Low / ND |

*Registre observacional de camp que prepara l'anàlisi formal de visibilitat i orientació en QGIS (H06: visibilitat i marcatge territorial).*

## **2.7. Decoració - camps booleans de resum (13) - per a khi-quadrat i filtratge ràpid**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| Dec_Square_Niche | YESNO | Nínxol quadrat en sèrie (hornacinas cuadradas) |
| Dec_Relief_T | YESNO | Relleu en T |
| Dec_Relief_T_Inv | YESNO | Relleu en T invertida |
| Dec_Relief_L | YESNO | Relleu en L |
| Dec_Relief_L_Inv | YESNO | Relleu en L invertida |
| Dec_Zigzag | YESNO | Motiu en zig-zag o chevron |
| Dec_Stepped | YESNO | Motiu escalonat (inusual a DW i LP; documentat a DWS1) |
| Rock_Art | YESNO | Pintura rupestre present (associada a l'estructura) |
| RA_Anthropomorphic | YESNO | Motiu antropomorf a la pintura rupestre |
| RA_Zoomorphic | YESNO | Motiu zoomorf |
| RA_Geometric | YESNO | Motiu geomètric |
| RA_Abstract | YESNO | Motiu abstracte |
| RA_Decap_Scene | YESNO | Escena de decapitació documentada |

*Dec_Frieze ha estat eliminat en v4: el fris és l'element M del vocabulari (Relief_Frieze, secció 2.4). Els camps booleans de decoració a T_STRUCTURES són per a filtratge ràpid; el registre detallat per posició a la façana (cos constructiu + tipus) és a T_DECORATIONS. Els camps RA_* qualifiquen l'art rupestre, no el baix relleu: al formulari estan sota la capçalera «Art rupestre associat».*

## **2.8. Estat de conservació (6)**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| ID_Arch_Status | LONG | FK -> L_STATUS. Estat de l'estructura construïda (Good/Fair/Pre-collapse/Collapsed/ND) |
| ID_Material_Status | LONG | FK -> L_MATERIAL_STATUS. Grau de preservació dels vestigis mobles (Good/Fair/Poor/Absent/ND) |
| Looting | YESNO | Evidència de saqueig (causa, independent del grau de conservació) |
| Fire_Damage | YESNO | Evidència d'incendi |
| Animal_Activity | YESNO | Activitat animal documentada |
| Modern_Access | YESNO | Evidència d'accés modern no autoritzat |

*Al formulari, les etiquetes són «Estat estructura» i «Estat vestigis mobles» (sense sigles (I)/(M), reservades al vocabulari A-X).*

## **2.9. Bioarqueologia (8)**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| Human_Remains | YESNO | Presència de restes humanes documentades |
| MNI | INTEGER | Nombre Mínim d'Individus (Minimum Number of Individuals) |
| Anatomical_Connection | YESNO | Restes en connexió anatòmica |
| Mummification | YESNO | Evidència de momificació (conservació de teixits tous) |
| Funerary_Bundles | YESNO | Fardells funeraris (embolcall tèxtil) |
| Dispersed_Remains | YESNO | Restes disperses (posició secundària) |
| Flexed_Position | YESNO | Posició flexionada documentada |
| Bone_Burning | YESNO | Cremació d'ossos |

## **2.10. Materials culturals (7)**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| Mat_Textiles | YESNO | Tèxtils presents |
| Mat_Wood | YESNO | Fusta cultural (artefactes, no estructural) |
| Mat_VegFiber | YESNO | Fibra vegetal |
| Mat_Ceramics | YESNO | Ceràmica (inusual en tombes aèries de DW i LP) |
| Mat_Fauna | YESNO | Restes de fauna |
| Mat_DeerAntler | YESNO | Banya de cérvol |
| Mat_Other | YESNO | Altres materials culturals |

## **2.11. Cronologia (3)**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| C14 | YESNO | Datació radiocarbònica disponible (detall a T_DATING) |
| Chrono_Start_Cent | INTEGER | Segle d'inici d.n.e. (ex: 10 = s. X). Estimació sense C14 o rang calibrat |
| Chrono_End_Cent | INTEGER | Segle de fi d.n.e. |

## **2.11b. Fases constructives (2) — H03: saturació acumulativa / H04: seqüència operativa**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| Construction_Phases | INTEGER | Nombre de fases constructives identificades a l'estructura |
| Phase_Evidence | TEXT(30) | Evidència de la identificació de fases (llista al formulari): C14 / Stratigraphy / Superposition / Mortar / ND |

*Permeten registrar creixement per fases sense dependre de datacions absolutes: l'argument morfològic-estructural (superposicions, juntes, canvis d'aparell) sosté H03 en absència de C14.*

## **2.12. Volumetria i àrea (5)**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| Interior_Area_m2 | SINGLE | Àrea de sòl interior (m2). Per a cambres irregulars, extreta del model 3D |
| Interior_Vol_m3 | SINGLE | Volum interior útil (m3) |
| Total_Vol_m3 | SINGLE | Volum total incloent parets (m3) |
| ID_Vol_Method | LONG | FK -> L_VOL_METHOD. Mètode de càlcul |
| Vol_Notes | TEXT(200) | Notes sobre el càlcul volumètric |

## **2.13. Coordenades espacials (7)**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| Coord_Lat_WGS84 | DOUBLE | Latitud WGS84 (graus decimals, negatiu = S) |
| Coord_Lon_WGS84 | DOUBLE | Longitud WGS84 (graus decimals, negatiu = W) |
| Coord_E_UTM | DOUBLE | Coordenada Est UTM Zona 18S / WGS84 (m) |
| Coord_N_UTM | DOUBLE | Coordenada Nord UTM Zona 18S / WGS84 (m) |
| Altitude_masl | SINGLE | Altitud sobre el nivell del mar (m) |
| Coord_Precision_m | SINGLE | Precisió estimada de les coordenades (m) |
| ID_Coord_Method | LONG | FK -> L_COORD_METHOD |

## **2.14. Documentació digital (7)**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| URL_Pano | TEXT(255) | URL de la primera escena panoràmica 360 a Chacha XR |
| URL_Pano_2 | TEXT(255) | URL de la segona escena (per a estructures amb 2+ nodes) |
| URL_Giga | TEXT(255) | URL de la gigafoto |
| URL_3D | TEXT(255) | URL del model 3D fotogramètric |
| ChaXR_Documented | YESNO | Estructura publicada a la plataforma Chacha XR |
| ID_Campaign | LONG | FK -> L_CAMPAIGN. Campanya en què es va documentar sistemàticament |
| Notes | MEMO | Notes lliures (text llarg) |

# **3. Taules secundàries**

## **3.1. T_DATING**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| ID | COUNTER PK |  |
| ID_Structure | LONG NN | FK -> T_STRUCTURES (cascade delete) |
| Sample_Type | TEXT(30) | Tipus de mostra: Charcoal / Bone / Wood / Textile / Other / ND |
| Date_BP | LONG | Data radiocarbònica en anys BP (Before Present) |
| Sigma1_Start | INTEGER | Interval calibrat 1 sigma - inici (cal d.n.e.) |
| Sigma1_End | INTEGER | Interval calibrat 1 sigma - fi (cal d.n.e.) |
| Sigma2_Start | INTEGER | Interval calibrat 2 sigma - inici |
| Sigma2_End | INTEGER | Interval calibrat 2 sigma - fi |
| Lab_Reference | TEXT(50) | Referència del laboratori (ex: Beta-123456) |
| Bibliog_Reference | MEMO | Referència bibliogràfica de la publicació de la data |

## **3.2. T_INDIVIDUALS**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| ID | COUNTER PK |  |
| ID_Structure | LONG NN | FK -> T_STRUCTURES |
| Individual_No | INTEGER | Número d'individu dins de l'estructura |
| Age_Category | TEXT(20) | Infant / Juvenile / Adult / Old adult / ND |
| Sex_Category | TEXT(20) | Male / Female / Indeterminate / ND |
| Preservation | TEXT(20) | Good / Fair / Poor / ND |
| Notes | MEMO | Notes sobre l'individu |

## **3.3. T_GROUPS**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| ID | COUNTER PK |  |
| Group_Code | TEXT(20) NN | Codi del conjunt (ex: DWS1-C01, PTC-S-C03) |
| ID_Sector | LONG NN | FK -> L_SECTORS |
| ID_Group_Type | LONG NN | FK -> L_GROUP_TYPE |
| N_Members | INTEGER | Nombre d'estructures membres del conjunt |
| Notes | MEMO |  |

## **3.4. T_DECORATIONS - Registre detallat de decoració per estructura i posició**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| ID | COUNTER PK |  |
| ID_Structure | LONG NN | FK -> T_STRUCTURES (cascade delete) |
| ID_Struct_Body | LONG | FK -> L_STRUCT_BODY. On està la decoració a la façana |
| ID_Dec_Type | LONG | FK -> L_DEC_TYPE. Tipus de decoració |
| Body_No | INTEGER | Número del cos constructiu (0=basament, 1=primer cos, 2=segon cos) |
| Color | TEXT(20) | Red / White / Both / None / ND |
| Notes | TEXT(255) |  |

## **3.5. T_ARCH_FEATURES - Registre flexible d'elements no previstos a l'esquema**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| ID | COUNTER PK |  |
| ID_Structure | LONG NN | FK -> T_STRUCTURES (cascade delete) |
| Feature_Code | TEXT(40) | Nom de l'element (text lliure o desplegable amb suggeriments: Pigment trace / Unusual bond / Textile fixation / Wooden peg / Other) |
| Present | YESNO |  |
| Feature_Count | INTEGER | Nombre d'unitats si escau |
| Material | TEXT(20) | Stone / Timber / Mixed / ND |
| Notes | TEXT(255) |  |

*En v4, els antics «elements addicionals» del formulari (A, D, J, V, W, X) tenen camp propi a la pestanya 11.Sist; T_ARCH_FEATURES queda reservada per a elements realment no previstos.*

## **3.6. T_CONNECTIONS - NOVA v4 - Connexions físiques entre estructures (OE3)**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| ID | COUNTER PK |  |
| ID_Struct_A | LONG NN | FK -> T_STRUCTURES.ID (extrem A de l'aresta) |
| ID_Struct_B | LONG NN | FK -> T_STRUCTURES.ID (extrem B de l'aresta) |
| Connection_Type | TEXT(30) | Shared ledge / Continuous platform / Shared wall / Shared beam / Intervisibility only / ND |
| Confidence | TEXT(10) | Certain / Probable / Possible |
| Notes | TEXT(150) |  |

*Cada registre és una aresta no dirigida de la xarxa de circulació aèria. Convenció d'entrada: ID_Struct_A < ID_Struct_B (evita arestes duplicades). ID_Parent registra contenció i ID_Group agrupació funcional; cap dels dos captura l'adjacència física que reconstrueix la circulació — d'ací aquesta taula. Exportació a igraph (R) o QGIS via QRY_14.*

# **4. Taules lookup - valors**

*Tots els valors emmagatzemats són en anglés (convenció fixa del projecte). Les descripcions següents es tradueixen només a efectes de lectura del document.*

## **L_TYPOLOGY**

| **Codi** | **Descripció** |
| --- | --- |
| EA-MAU Mausoleum/Chullpa | Estructura construïda (3+ murs + sostre artificial) sobre repisa. 1-3 pisos. Predominant a La Petaca. |
| EA-CAM Funerary Chamber | Cavitat natural tancada per 1 façana construïda. Predominant a Diablo Wasi. |
| EA-PLA-R Ledge Platform | Plataforma constructiva sobre repisa natural. Funció: trànsit o base per a mausoleus. |
| EA-PLA-V Aerial Platform | Plataforma artificial sobre bigues i lloses, sense repisa natural de suport. |
| NIX Natural Niche | Petita cavitat natural (<1m2). Ossari o enterrament secundari. |
| CAV Cave/Cavern | Gran cavitat natural (>1m2) amb ús funerari o ritual documentat. |
| PR Rock Art | Motiu pictòric sobre roca, documentat de forma independent. |
| MEN Isolated Bracket | Element estructural aïllat. Evidència de xarxa de circulació aèria perduda. |
| MIX Mixed | Combinació de dues o més categories anteriors. |
| ND Undetermined | Informació insuficient per a classificar. |

## **L_SUPPORT**

| **Valor** |
| --- |
| Wide natural ledge (>2m) |
| Narrow natural ledge (<2m) |
| Artificial ledge |
| Large cavity (>10m2) |
| Medium cavity (1-10m2) |
| Natural niche (<1m2) |
| Fissure/Crack |
| Combined |
| ND |

## **L_STATUS (estructura) | L_MATERIAL_STATUS (vestigis mobles)**

| **L_STATUS** | **L_MATERIAL_STATUS** |
| --- | --- |
| Good | Good - Material remains well preserved and identifiable |
| Fair | Fair - Partially preserved: some elements present and identifiable |
| Pre-collapse | Poor - Fragmentary: barely identifiable, heavily degraded or scattered |
| Collapsed | Absent - No movable remains documented (cause in boolean fields) |
| ND | ND - Not determined: not assessed or structure not accessible |

## **L_GROUP_TYPE**

| **Valor** | **Descripció** |
| --- | --- |
| Vertical alignment | Estructures en la mateixa vertical del faralló (clivella o estrats successius) |
| Ledge cluster | Múltiples EA compartint una mateixa repisa horitzontal |
| Platform with brackets | Plataforma aèria + mènsules que la componen o la flanquegen |
| Cave cluster | Cova + estructures construïdes al seu interior |
| Circulation network | MEN i EA-PLA que reconstrueixen una circulació aèria perduda |
| Rock art cluster | PR + EA adjacent visualment o funcionalment vinculada |
| Functional group | Qualsevol altra agrupació intrasector amb coherència funcional |

## **L_STRUCT_BODY (posició de la decoració a la façana)**

| **Code** | **Name** | **Level** | **Descripció** |
| --- | --- | --- | --- |
| N0-SOC | Socle (N0) | 0 | Sòcol decoratiu - element basal nivell 0 |
| N1-SPA | Spandrel (N1) | 1 | Parament lateral del cos principal (fora del marc del portal) |
| N1-JAM | Jamb (N1) | 1 | Brancal del portal - cos principal |
| N1-OVL | Over-lintel (N1) | 1 | Zona per damunt del dintell - àrea del fris decoratiu |
| N1-COR | Interbody cornice | 1 | Zona de la cornisa intercòs entre cossos superposats (v4: terminologia actualitzada) |
| N2-SPA | Spandrel (N2) | 2 | Parament lateral del cos superior (estructures de 2 cossos) |
| N2-JAM | Jamb (N2) | 2 | Brancal del portal - cos superior |
| ND | Not determined | -1 | Posició no determinada |

## **L_DEC_TYPE**

| **Nom** | **Descripció** |
| --- | --- |
| T-shaped niche | Nínxol o baix relleu en forma de T. Element vertical + horitzontal. |
| T-shaped niche inv. | T invertida. |
| L-shaped niche | Nínxol o baix relleu en forma de L. |
| L-shaped niche inv. | L invertida. |
| Zigzag | Motiu en zig-zag o chevron. |
| Stepped motif | Motiu escalonat. Documentat a DWS1 (inusual a la cultura Chachapoya). |
| Frieze / Greca | Fris o greca repetitiva. Comú a La Petaca. |
| Triangular motif | Patró triangular / chevron pintat. Documentat a DW en zona sobre el dintell. |
| Painted band | Banda horitzontal pintada (roja/blanca). Zona de la cornisa intercòs. |
| Square niche | Nínxols quadrats en sèrie (hornacinas cuadradas). |
| ND | Tipus de decoració no determinat. |

## **L_CAMPAIGN**

| **Code** | **Campaign_Name** | **Descripció** |
| --- | --- | --- |
| 2013 | PALP I | Primera campanya. Prospecció i documentació inicial de La Petaca. Dra. J. Marla Toyne (UCF). |
| 2016 | PALP II | Ampliació a Diablo Wasi. Primera documentació sistemàtica de DW. |
| 2021 | La Petaca Project | Documentació no invasiva integral. Fotogrametria, 360, gigafotos. Panograma Labs/UCF. |
| 2023 | PALP IV | Campanya d'excavació arqueològica i reconstruccions 3D detallades. |

# **5. Relacions (20)**

| **Nom** | **Taula pare** | **Camp pare** | **Taula filla** | **Camp fill** | **Cascade** | **Integritat** |
| --- | --- | --- | --- | --- | --- | --- |
| REL_SIT_SEC | L_SITES | ID | L_SECTORS | ID_Site | Delete | Sí |
| REL_SEC_STR | L_SECTORS | ID | T_STRUCTURES | ID_Sector | Delete | Sí |
| REL_SEC_GRP | L_SECTORS | ID | T_GROUPS | ID_Sector | Delete | Sí |
| REL_TYP_STR | L_TYPOLOGY | ID | T_STRUCTURES | ID_Typology | Update | Sí |
| REL_SUP_STR | L_SUPPORT | ID | T_STRUCTURES | ID_Support | Update | Sí |
| REL_STR_SELF | T_STRUCTURES | ID | T_STRUCTURES | ID_Parent | Update | NO |
| REL_STA_STR | L_STATUS | ID | T_STRUCTURES | ID_Arch_Status | Update | Sí |
| REL_MATSTA_STR | L_MATERIAL_STATUS | ID | T_STRUCTURES | ID_Material_Status | Update | Sí |
| REL_VM_STR | L_VOL_METHOD | ID | T_STRUCTURES | ID_Vol_Method | Update | Sí |
| REL_CM_STR | L_COORD_METHOD | ID | T_STRUCTURES | ID_Coord_Method | Update | Sí |
| REL_GT_GRP | L_GROUP_TYPE | ID | T_GROUPS | ID_Group_Type | Update | Sí |
| REL_GRP_STR | T_GROUPS | ID | T_STRUCTURES | ID_Group | Update | Sí |
| REL_STR_DAT | T_STRUCTURES | ID | T_DATING | ID_Structure | Delete | Sí |
| REL_STR_IND | T_STRUCTURES | ID | T_INDIVIDUALS | ID_Structure | Delete | Sí |
| REL_STR_DEC | T_STRUCTURES | ID | T_DECORATIONS | ID_Structure | Delete | Sí |
| REL_SB_DEC | L_STRUCT_BODY | ID | T_DECORATIONS | ID_Struct_Body | Update | Sí |
| REL_DT_DEC | L_DEC_TYPE | ID | T_DECORATIONS | ID_Dec_Type | Update | Sí |
| REL_STR_AFEAT | T_STRUCTURES | ID | T_ARCH_FEATURES | ID_Structure | Delete | Sí |
| REL_STR_CONA | T_STRUCTURES | ID | T_CONNECTIONS | ID_Struct_A | - | NO |
| REL_STR_CONB | T_STRUCTURES | ID | T_CONNECTIONS | ID_Struct_B | - | NO |

*REL_STR_SELF (autoreferenciant) i les dues relacions de T_CONNECTIONS (doble referència a T_STRUCTURES) usen dbRelationDontEnforceIntegrity (valor numèric 2) per a evitar conflictes. La resta usen dbRelationUpdateCascade i, si Cascade=Delete, dbRelationDeleteCascade.*

# **6. Consultes SQL (14)**

| **Nom** | **Descripció** | **Hipòtesis** |
| --- | --- | --- |
| QRY_01_Typology_by_Site | Distribució de tipologies per jaciment i sector (COUNT per grup). | H01, H04 |
| QRY_02_Decoration_by_Site | Presència de cada motiu decoratiu per jaciment. Font per a khi-quadrat. Actualitzada v4: N_Frieze es compta des de Relief_Frieze (element M, valor 1). | H01, H04 |
| QRY_03_Conservation_by_Sector | Distribució dual d'estat (estructura + vestigis mobles) per sector. | General |
| QRY_04_C14_Structures | Estructures amb datació C14 ordenades cronològicament. | H04 |
| QRY_05_Export_RStats | Exportació plana completa per a R/SPSS. Actualitzada v4: Relief_Frieze, Interbody_Cornice, Lateral_Wall_Faces, bloc de morter, dimensions d'obertura i Height_Above_Base_m. | Totes |
| QRY_06_Volumetry_by_Typology | Volumetria i àrea interiors per tipologia (mitjana, mín., màx.). | OE2, H01 |
| QRY_07_Export_QGIS | Exportació espacial per a QGIS. Actualitzada v4: Height_Above_Base_m. LEFT JOIN per a conservar estructures sense estat. | OE3, H02, H03, H06 |
| QRY_08_Children_of_Parent | Elements continguts per una estructura pare (paràmetre: [Parent ID?]). | H01, H02 |
| QRY_09_Group_Members | Membres d'un conjunt funcional ordenats per altitud (paràm.: [Group code?]). | H03, OE3 |
| QRY_10_ChaXR_Coverage | Cobertura Chacha XR vs total per jaciment i campanya. | Metodologia |
| QRY_11_Masonry_by_Site | Qualitat i tipus de maçoneria per jaciment i tipologia (COUNT per grup; exclou registres sense Masonry_Quality). | H01, H04 |
| QRY_12_Geology_Construction | Encreuament suport geològic - tipologia: morfologia i modificació del suport, amb mitjanes d'amplària i profunditat (cm). | H02 |
| QRY_13_AX_Pattern_Export | Reescrita v4. Matriu A-X completa: 24 columnes AX_A..AX_X en ordre alfabètic amb valors 0/1/9. AX_Q derivada del camp Lintel (Absent→0, ND/Null→9, resta→1). Inclou qualificadors (comptatge de mènsules, materials, morter) i coordenades UTM. En R, filtrar o ponderar les cel·les amb valor 9 abans de calcular phi/Jaccard, clúster jeràrquic, I de Moran o AC. LEFT JOIN amb L_TYPOLOGY per a incloure registres parcials. | H01, H04, H05 |
| QRY_14_Connections_Edges | NOVA v4. Llista d'arestes de T_CONNECTIONS amb els codis i les coordenades UTM/altitud dels dos extrems: entrada directa per a igraph (R) o generació de línies en QGIS (xarxa de circulació aèria). | OE3, H03 |

# **7. Vocabulari arquitectònic normalitzat A-X**

24 elements organitzats per nivell constructiu (bottom-to-top) i funció (estructural -> decoratiu). Tres sistemes compositius: E+F+G->H (plataforma), N+O+Q->P (portal), S+T->U (ràfec).

**Nota terminològica (v4):** «nivell» designa exclusivament els nivells constructius del vocabulari (N0 / N1 / Superior); «cos» designa els pisos superposats de l'estructura (N_Floors, Body_No, L_STRUCT_BODY). La cornisa I és per tant «cornisa intercòs». H i I comparteixen posició d'interfície N0/N1 però es distingeixen per funció (H: superfície de circulació i accés; I: marcatge i separació entre cossos), per composició (H és sistema; I és standalone) i per seqüència operativa (H tanca N0 i precondiciona N1; I apareix dins de l'alçat N1, només amb 2+ cossos).

| **Ll.** | **Valencià** | **Anglés (BD)** | **Nivell** | **Sistema** | **Camp BD** |
| --- | --- | --- | --- | --- | --- |
| A | Jàcenes basals empotrades | Embedded base beams | N0 | Standalone | Embedded_Base_Beams |
| B | Basament | Base level / podium | N0 | Standalone | Base_Level |
| C | Sòcol decoratiu | Decorative socle | N0 | Standalone (tract. de B) | Decorative_Socle |
| D | Muret transversal | Tie wall | N0 | Standalone | Tie_Walls |
| E | Mènsules (fusta) | Timber corbels | N0 | Component de H | Timber_Brackets |
| F | Bigues transversals | Transverse beams | N0 | Component de H | Transverse_Beams |
| G | Filades en voladís | Corbelled masonry courses | N0 | Component de H (lítia) | Corbelled_Courses |
| **H** | **Plataforma volada d'accés** | Corbelled access platform | **N0/N1 (interfície)** | **SISTEMA: E+F+G -> H** | Corbelled_Platform |
| I | Cornisa intercòs | Interbody cornice | N0/N1 (interfície) | Standalone (límit cossos) | Interbody_Cornice |
| J | Cantoneres | Corner quoins | N1 | Standalone | Corner_Quoins |
| K | Pilastres estructurals | Structural pilasters | N1 | Standalone | Structural_Pilasters |
| L | Paraments laterals | Lateral wall faces | N1 | Standalone | Lateral_Wall_Faces |
| M | Fris decoratiu en baix relleu | Bas-relief decorative frieze | N1 | Standalone (tract. de L) | Relief_Frieze |
| N | Llindar | Sill / threshold | N1 | Component de P | Sill |
| O | Brancals | Jambs | N1 | Component de P | Jambs |
| **P** | **Obertura d'accés** | Access opening | N1 | **SISTEMA: N+O+Q -> P** | Access_Opening |
| Q | Dintell | Lintel | N1 | Component de P | Lintel (material; AX_Q derivada a QRY_13) |
| R | Coronament | Upper crown / coping | N1 | Standalone | Upper_Crown |
| S | Biga de suport del ràfec | Eave-supporting beam | Sup. | Component de U | Eave_Beam |
| T | Superfície del ràfec | Eave surface | Sup. | Comp. de U (SEMPRE PEDRA) | Eave_Surface |
| **U** | **Ràfec-voladís / Visera** | Eave / roof overhang | Sup. | **SISTEMA: S+T -> U (PEDRA)** | Eave |
| V | Murs laterals | Lateral walls | N1 | Standalone | Lateral_Walls |
| W | Mur posterior | Rear wall | N1 | Standalone (Rear_Wall_Built) | Rear_Wall |
| X | Coberta de la cambra | Chamber roof | Sup. | Standalone | Chamber_Roof |

**Criteri operatiu E vs F (reproduïbilitat):** E (mènsula) = element perpendicular a la façana, encastat a la roca, treballant en voladís; F (biga transversal) = element paral·lel a la façana, salvant llum entre suports. Fixar aquest criteri per escrit garanteix la consistència del registre entre estructures i observadors.

# **8. Formulari i visualització de dades**

El formulari es genera amb chachapoya_Form_v3_val.bas i segueix la convenció bilingüe fixa del projecte: **totes les etiquetes visibles de la interfície (pestanyes, camps, capçaleres de subformularis) són en valencià, mentre que tots els valors emmagatzemats (llistes de valors dels ComboBox, continguts de les taules lookup, dominis 0/1/9 amb etiquetes Absent/Present/ND) romanen en anglés**, per a garantir la reproduïbilitat de les exportacions analítiques.

Les 12 pestanyes de F_STRUCTURES són: 1.Id. (identificació + morfologia del suport geològic), 2.Arq. (morfologia general + façana/paisatge + maçoneria i morter + fases), 3.Acab., 4.Dec. (booleans de baix relleu + art rupestre + subformulari F_DECORATIONS), 5.Estat, 6.Bio., 7.Mat., 8.Cron., 9.Metr. (dimensions + obertura + mètrica del suport + volumetria + coordenades), 10.Doc. (URLs i notes), 11.Sist. (els 24 elements A-X amb combos 0/1/9) i 12.Extra (subformulari F_ARCH_FEATURES).

**Capçaleres de columna dels subformularis:** en vista full de dades, Access mostra com a capçalera la llegenda de l'etiqueta **adjunta** al control; si el control no té etiqueta adjunta, mostra el nom del control (el nom de camp anglés). En v3, les etiquetes dels subformularis es creen amb el control com a pare (patró: primer el control amb nom = camp, després CreateControl amb el nom del control com a quart paràmetre), cosa que resol les capçaleres en valencià. Les captions DAO a nivell de TableDef (SetFieldCaptionsVal) es mantenen com a reforç per a l'obertura directa de T_DECORATIONS, T_ARCH_FEATURES i T_CONNECTIONS en vista de taula.

**Entrada de dades a T_CONNECTIONS:** sense formulari propi (volum baix d'arestes); s'ompli en vista de taula amb les capçaleres DAO en valencià. Convenció: ID_Struct_A < ID_Struct_B.
