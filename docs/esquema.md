**ESQUEMA DE LA BASE DE DADES ARQUEOLOGICA**

*La Petaca i Diablo Wasi (Leymebamba, Amazonas, Peru)*

chachapoya_01_DB.bas + chachapoya_02_Form.bas

Sub BuildDB() + Sub BuildForm() | Microsoft Access JET SQL

## **1. Resum general**

| **Element** | **Valor** |
|---|---|
| Taules principals | T_STRUCTURES (106 camps), T_DATING, T_INDIVIDUALS, T_GROUPS, T_DECORATIONS, T_ARCH_FEATURES |
| Taules lookup | L_SITES, L_SECTORS, L_TYPOLOGY, L_SUPPORT, L_STATUS, L_MATERIAL_STATUS, L_VOL_METHOD, L_COORD_METHOD, L_GROUP_TYPE, L_CAMPAIGN, L_STRUCT_BODY, L_DEC_TYPE |
| Total taules | 18 |
| Total relacions | 18 (inclou autoreferenciant T_STRUCTURES i relacions de T_DECORATIONS) |
| Consultes SQL | 10 (QRY_01 a QRY_10) |
| Formulari | F_STRUCTURES - 11 pestanyes + subformularis F_DECORATIONS i F_ARCH_FEATURES |
| Idioma BD | Angles (nom de taules, camps i valors lookup) |
| Motor | Microsoft Access JET SQL / ACE \| cp1252 |
| Jaciments | La Petaca (WGS84: Lat -6.8311, Lon -77.8084) \| Diablo Wasi (Lat -6.8475, Lon -77.8154) |

## **2. T_STRUCTURES (106 camps)**

#### **Identificacio (7)**

| **Camp** | **Tipus** | **Descripcio** |
|---|---|---|
| ID | COUNTER PK | Clau primaria autonumerica |
| Code | TEXT(20) NN | Codi PALP: [SITE][SECTOR]-[TIPTYPE][NUM] - ex: DWS1-EF01, PTC-SS-EF18 |
| ID_Sector | LONG NN | FK -> L_SECTORS |
| ID_Typology | LONG | FK -> L_TYPOLOGY |
| ID_Support | LONG | FK -> L_SUPPORT. Suport geomorfologic |
| ID_Parent | LONG | FK autoreferenciant -> T_STRUCTURES.ID. Contenicio fisica |
| ID_Group | LONG | FK -> T_GROUPS. Agrupacio funcional (alineament, xarxa) |

#### **Morfologia (13)**

| **Camp** | **Tipus** | **Descripcio** |
|---|---|---|
| N_Floors | INTEGER | Nombre de pisos constructius |
| Floor_Plan | TEXT(20) | Planta: Rectangular / Square / Trapezoidal / Irregular / ND |
| N_Built_Walls | INTEGER | Nombre de murs construits (0-4) |
| Length_m | SINGLE | Llarg exterior (m) |
| Width_m | SINGLE | Ample exterior (m) |
| Height_m | SINGLE | Alcada total de l'estructura (m) |
| Approx_Height_m | SINGLE | Alcada aproximada sobre el sol del faralló (m) |
| Access_Orientation | TEXT(5) | Orientacio del vano: N / S / E / W / NE / NW / SE / SW / ND |
| Lintel | TEXT(20) | Material del dintell del vano principal: Stone / Wood / Mixed / Absent / ND |
| Natural_Roof | YESNO | Techo natural de roca (faralló actua com a sostre) |
| Buttresses | YESNO | Contraforts a la facana |
| Interlevel_Cornice | YESNO | Cornisa intranivell (element I del vocabulari A-X) present |
| Wooden_Stakes | YESNO | Estacas o pals de fusta verticals (elements estructurals d'ancoratge) |

#### **Sistemes arquitectonics - Nivell 0 / Basament (10) - vocabulari A-H**

| **Camp** | **Tipus** | **Elem.** | **Descripcio** |
|---|---|---|---|
| Base_Level | YESNO | B | Basament com a element constructiu diferenciat (podium) |
| Decorative_Socle | YESNO | C | Tractament decoratiu del basament (socol decoratiu) |
| Corbelled_Platform | YESNO | H | Plataforma volada d'acces. Sistema: E+F+G->H |
| Timber_Brackets | YESNO | E | Mensules de fusta empotrades a la roca (component de H) |
| Timber_Bracket_Count | INTEGER | E | Nombre de mensules visibles |
| Transverse_Beams | YESNO | F | Bigues transversals (component de H) |
| Corbelled_Courses | YESNO | G | Filades de pedra en voladis creixent (component de H, variant litia) |
| Corbel_Material | TEXT(20) | E/G | Material: Timber / Stone / Mixed / ND |
| Embedded_Base_Beams | YESNO | A | Jacenes horitzontals empotrades dins la maconeria del basament |
| Tie_Walls | YESNO | D | Murets perpendiculars al faralló (ancoratge/compartimentacio) |

#### **Sistemes arquitectonics - Nivell 1 / Cos principal (9) - vocabulari J-W**

| **Camp** | **Tipus** | **Elem.** | **Descripcio** |
|---|---|---|---|
| Access_Opening | YESNO | P | Obertura d'acces present. Sistema: N+O+Q->P |
| Sill | YESNO | N | Llindar / threshold (component de P) |
| Jambs | YESNO | O | Brancals: elements verticals formals del portal (component de P) |
| Lateral_Walls | YESNO | V | Murs laterals que donen profunditat a la cambra funeraria |
| Rear_Wall | YESNO | W | Mur posterior present |
| Rear_Wall_Built | YESNO | W | Mur posterior construït (True) vs roca natural (False) |
| Recessed_Portal | YESNO | - | Portal retraqueig de facana (el vano esta enfonssat respecte al pla exterior) |
| Structural_Pilasters | YESNO | K | Pilastres estructurals que flanquegen la facana (altura completa) |
| Cornice_Material | TEXT(20) | I | Material de la cornisa intranivell: Stone slabs / Wooden beams / Mixed / ND |

#### **Sistemes - Zona superior / Coberta (6) - vocabulari R,S,T,U,X**

| **Camp** | **Tipus** | **Elem.** | **Descripcio** |
|---|---|---|---|
| Eave | YESNO | U | Rafec-voladis / visera present. Sistema: S+T->U. Sempre pedra |
| Eave_Beam | YESNO | S | Biga de suport del rafec (component de U; pot ser fusta) |
| Eave_Surface | YESNO | T | Superficie del rafec (component de U; sempre pedra - lloses) |
| Upper_Crown | YESNO | R | Coronament superior / coping |
| Chamber_Roof | YESNO | X | Coberta construida de la cambra (distinta de Natural_Roof i Eave) |
| Corner_Quoins | YESNO | J | Cantoneres: pedres de major mida disposades verticalment als angles |

#### **Acabats superficials (4)**

| **Camp** | **Tipus** | **Descripcio** |
|---|---|---|
| Plastered | YESNO | Arrebossat / enlluiment present |
| Plaster_Color | TEXT(20) | Color de l'arrebossat (text lliure) |
| Rock_Painting | YESNO | Pintura sobre la roca present |
| Rock_Paint_Color | TEXT(20) | Color de la pintura sobre roca (text lliure) |

#### **Decoracio - camps booleans de resum (14) - per a chi-quadrat i filtratge rapid**

| **Camp** | **Tipus** | **Descripcio** |
|---|---|---|
| Dec_Square_Niche | YESNO | Nixol quadrat en serie (hornacinas cuadradas) |
| Dec_Relief_T | YESNO | Relleu en T |
| Dec_Relief_T_Inv | YESNO | Relleu en T invertida |
| Dec_Relief_L | YESNO | Relleu en L |
| Dec_Relief_L_Inv | YESNO | Relleu en L invertida |
| Dec_Zigzag | YESNO | Motiu en zigzag o chevron |
| Dec_Stepped | YESNO | Motiu escalonat (inusual a DW i LP; documentat a DWS1) |
| Dec_Frieze | YESNO | Fris / greca |
| Rock_Art | YESNO | Pintura rupestre present (associada a l'estructura) |
| RA_Anthropomorphic | YESNO | Motiu antropomorf a la pintura rupestre |
| RA_Zoomorphic | YESNO | Motiu zoomorf |
| RA_Geometric | YESNO | Motiu geometric |
| RA_Abstract | YESNO | Motiu abstracte |
| RA_Decap_Scene | YESNO | Escena de decapitacio documentada |

*Nota: els camps booleans de decoracio a T_STRUCTURES son per a filtratge rapid. El registre detallat per posicio a la facana (cos constructiu + tipus) es a T_DECORATIONS.*

#### **Estat de conservacio (6)**

| **Camp** | **Tipus** | **Descripcio** |
|---|---|---|
| ID_Arch_Status | LONG | FK -> L_STATUS. Estat de l'arquitectura construida (Good/Fair/Pre-collapse/Collapsed/ND) |
| ID_Material_Status | LONG | FK -> L_MATERIAL_STATUS. Grau preservacio vestigis mobles (Good/Fair/Poor/Absent/ND) |
| Looting | YESNO | Evidencia de saqueig (causa, independent del grau de conservacio) |
| Fire_Damage | YESNO | Evidencia d'incendi |
| Animal_Activity | YESNO | Activitat animal documentada |
| Modern_Access | YESNO | Evidencia d'acces modern no autoritzat |

#### **Bioarqueologia (8)**

| **Camp** | **Tipus** | **Descripcio** |
|---|---|---|
| Human_Remains | YESNO | Presencia de restes humanes documentades |
| MNI | INTEGER | Nombre Minim d'Individus (Minimum Number of Individuals) |
| Anatomical_Connection | YESNO | Restes en connexio anatomica |
| Mummification | YESNO | Evidencia de momificacio (conservacio de teixits tous) |
| Funerary_Bundles | YESNO | Fardells funeraris (envolcall textil) |
| Dispersed_Remains | YESNO | Restes disperses (posicio secundaria) |
| Flexed_Position | YESNO | Posicio flexionada documentada |
| Bone_Burning | YESNO | Cremacio d'ossos |

#### **Materials culturals (7)**

| **Camp** | **Tipus** | **Descripcio** |
|---|---|---|
| Mat_Textiles | YESNO | Textils presents |
| Mat_Wood | YESNO | Fusta cultural (artefactes, no estructural) |
| Mat_VegFiber | YESNO | Fibra vegetal |
| Mat_Ceramics | YESNO | Ceramica (inusual en tombes aeries de DW i LP) |
| Mat_Fauna | YESNO | Restes de fauna |
| Mat_DeerAntler | YESNO | Banya de cervol |
| Mat_Other | YESNO | Altres materials culturals |

#### **Cronologia (3)**

| **Camp** | **Tipus** | **Descripcio** |
|---|---|---|
| C14 | YESNO | Datacio radiocarbonica disponible (detall a T_DATING) |
| Chrono_Start_Cent | INTEGER | Segle d'inici d.n.e. (ex: 10 = s. X). Estimacio sense C14 o rang calibrat |
| Chrono_End_Cent | INTEGER | Segle de fi d.n.e. |

#### **Volumetria i area (5)**

| **Camp** | **Tipus** | **Descripcio** |
|---|---|---|
| Interior_Area_m2 | SINGLE | Area de sol interior (m2). Per a cambres irregulars, extret del model 3D |
| Interior_Vol_m3 | SINGLE | Volum interior util (m3) |
| Total_Vol_m3 | SINGLE | Volum total incloent parets (m3) |
| ID_Vol_Method | LONG | FK -> L_VOL_METHOD. Metode de calcul |
| Vol_Notes | TEXT(200) | Notes sobre el calcul volumetric |

#### **Coordenades espacials (7)**

| **Camp** | **Tipus** | **Descripcio** |
|---|---|---|
| Coord_Lat_WGS84 | DOUBLE | Latitud WGS84 (graus decimals, negatiu = S) |
| Coord_Lon_WGS84 | DOUBLE | Longitud WGS84 (graus decimals, negatiu = W) |
| Coord_E_UTM | DOUBLE | Coordenada Est UTM Zona 18S / WGS84 (m) |
| Coord_N_UTM | DOUBLE | Coordenada Nord UTM Zona 18S / WGS84 (m) |
| Altitude_masl | SINGLE | Altitud sobre el nivell del mar (m) |
| Coord_Precision_m | SINGLE | Precisio estimada de les coordenades (m) |
| ID_Coord_Method | LONG | FK -> L_COORD_METHOD |

#### **Documentacio digital (7)**

| **Camp** | **Tipus** | **Descripcio** |
|---|---|---|
| URL_Pano | TEXT(255) | URL de la primera escena panoramica 360 a Chacha XR |
| URL_Pano_2 | TEXT(255) | URL de la segona escena (per a estructures amb 2+ nodes) |
| URL_Giga | TEXT(255) | URL de la gigafoto |
| URL_3D | TEXT(255) | URL del model 3D fotogrametric |
| ChaXR_Documented | YESNO | Estructura publicada a la plataforma Chacha XR |
| ID_Campaign | LONG | FK -> L_CAMPAIGN. Campanya en que es va documentar sistemàticament |
| Notes | MEMO | Notes lliures (text llarg) |

## **3. Taules secundaries**

### **3.1. T_DATING**

| **Camp** | **Tipus** | **Descripcio** |
|---|---|---|
| ID | COUNTER PK |  |
| ID_Structure | LONG NN | FK -> T_STRUCTURES (cascade delete) |
| Sample_Type | TEXT(30) | Tipus de mostra: Charcoal / Bone / Wood / Textile / Other / ND |
| Date_BP | LONG | Data radiocarbonica en anys BP (Before Present) |
| Sigma1_Start | INTEGER | Interval calibrat 1 sigma - inici (cal d.n.e.) |
| Sigma1_End | INTEGER | Interval calibrat 1 sigma - fi (cal d.n.e.) |
| Sigma2_Start | INTEGER | Interval calibrat 2 sigma - inici |
| Sigma2_End | INTEGER | Interval calibrat 2 sigma - fi |
| Lab_Reference | TEXT(50) | Referencia del laboratori (ex: Beta-123456) |
| Bibliog_Reference | MEMO | Referencia bibliografica de la publicacio de la data |

### **3.2. T_INDIVIDUALS**

| **Camp** | **Tipus** | **Descripcio** |
|---|---|---|
| ID | COUNTER PK |  |
| ID_Structure | LONG NN | FK -> T_STRUCTURES |
| Individual_No | INTEGER | Numero d'individu dins de l'estructura |
| Age_Category | TEXT(20) | Infant / Juvenile / Adult / Old adult / ND |
| Sex_Category | TEXT(20) | Male / Female / Indeterminate / ND |
| Preservation | TEXT(20) | Good / Fair / Poor / ND |
| Notes | MEMO | Notes sobre l'individu |

### **3.3. T_GROUPS**

| **Camp** | **Tipus** | **Descripcio** |
|---|---|---|
| ID | COUNTER PK |  |
| Group_Code | TEXT(20) NN | Codi del conjunt (ex: DWS1-C01, PTC-S-C03) |
| ID_Sector | LONG NN | FK -> L_SECTORS |
| ID_Group_Type | LONG NN | FK -> L_GROUP_TYPE |
| N_Members | INTEGER | Nombre d'estructures membres del conjunt |
| Notes | MEMO |  |

### **3.4. T_DECORATIONS - Registre detallat de decoracio per estructura i posicio**

| **Camp** | **Tipus** | **Descripcio** |
|---|---|---|
| ID | COUNTER PK |  |
| ID_Structure | LONG NN | FK -> T_STRUCTURES (cascade delete) |
| ID_Struct_Body | LONG | FK -> L_STRUCT_BODY. On esta la decoracio a la facana |
| ID_Dec_Type | LONG | FK -> L_DEC_TYPE. Tipus de decoracio |
| Body_No | INTEGER | Numero del cos constructiu (0=basament, 1=primer cos, 2=segon cos) |
| Color | TEXT(20) | Red / White / Both / None / ND |
| Notes | TEXT(255) |  |

### **3.5. T_ARCH_FEATURES - Registre flexible d'elements no previstos al schema**

| **Camp** | **Tipus** | **Descripcio** |
|---|---|---|
| ID | COUNTER PK |  |
| ID_Structure | LONG NN | FK -> T_STRUCTURES (cascade delete) |
| Feature_Code | TEXT(40) | Nom de l'element (text lliure o desplegable amb suggeriments) |
| Present | YESNO |  |
| Feature_Count | INTEGER | Nombre d'unitats si aplica |
| Material | TEXT(20) | Stone / Timber / Mixed / ND |
| Notes | TEXT(255) |  |

## **4. Taules lookup - valors**

### **L_TYPOLOGY**

| **Codi** | **Descripcio** |
|---|---|
| EA-MAU Mausoleum/Chullpa | Estructura construida (3+ murs + sostre artificial) sobre repisa. 1-3 pisos. Predominant a La Petaca. |
| EA-CAM Funerary Chamber | Cavitat natural tancada per 1 facana construida. Predominant a Diablo Wasi. |
| EA-PLA-R Ledge Platform | Plataforma constructiva sobre repisa natural. Funcio: transit o base per a mausoleus. |
| EA-PLA-V Aerial Platform | Plataforma artificial sobre bigues i lloses, sense repisa natural de suport. |
| NIX Natural Niche | Petita cavitat natural (<1m2). Ossari o enterrament secundari. |
| CAV Cave/Cavern | Gran cavitat natural (>1m2) amb us funerari o ritual documentat. |
| PR Rock Art | Motiu pictoric sobre roca, documentat de forma independent. |
| MEN Isolated Bracket | Element estructural aillat. Evidencia de xarxa de circulacio aeria perduda. |
| MIX Mixed | Combinacio de dues o mes categories anteriors. |
| ND Undetermined | Informacio insuficient per classificar. |

### **L_SUPPORT**

| **Valor** |
|---|
| Wide natural ledge (>2m) |
| Narrow natural ledge (<2m) |
| Artificial ledge |
| Large cavity (>10m2) |
| Medium cavity (1-10m2) |
| Natural niche (<1m2) |
| Fissure/Crack |
| Combined |
| ND |

### **L_STATUS (arquitectura) | L_MATERIAL_STATUS (vestigis mobles)**

| **L_STATUS** | **L_MATERIAL_STATUS** |
|---|---|
| Good | Good - Material remains well preserved and identifiable |
| Fair | Fair - Partially preserved: some elements present and identifiable |
| Pre-collapse | Poor - Fragmentary: barely identifiable, heavily degraded or scattered |
| Collapsed | Absent - No movable remains documented (cause in boolean fields) |
| ND | ND - Not determined: not assessed or structure not accessible |

### **L_GROUP_TYPE**

| **Valor** | **Descripcio** |
|---|---|
| Vertical alignment | Estructures en la mateixa vertical del faralló (grieta o estrats successius) |
| Ledge cluster | Multiples EAs compartint una mateixa repisa horitzontal |
| Platform with brackets | Repisa volada + mensules que la composen o flanquegen |
| Cave cluster | Cova + estructures construides al seu interior |
| Circulation network | MENs i EA-PLAs que reconstrueixen circulacio aeria perduda |
| Rock art cluster | PR + EA adjacent visualment o funcionalment vinculada |
| Functional group | Qualsevol altra agrupacio intra-sector amb coherencia funcional |

### **L_STRUCT_BODY (posicio de la decoracio a la facana)**

| **Code** | **Name** | **Level** | **Descripcio** |
|---|---|---|---|
| N0-SOC | Socle (N0) | 0 | Socol decoratiu - element basal nivell 0 |
| N1-SPA | Spandrel (N1) | 1 | Parament lateral del cos principal (fora del marc del portal) |
| N1-JAM | Jamb (N1) | 1 | Brancal del portal - cos principal |
| N1-OVL | Over-lintel (N1) | 1 | Zona per damunt del dintell - area del fris decoratiu |
| N1-COR | Interlevel cornice | 1 | Zona de la cornisa intranivell entre cossos |
| N2-SPA | Spandrel (N2) | 2 | Parament lateral del cos superior (estructures de 2 cossos) |
| N2-JAM | Jamb (N2) | 2 | Brancal del portal - cos superior |
| ND | Not determined | -1 | Posicio no determinada |

### **L_DEC_TYPE**

| **Nom** | **Descripcio** |
|---|---|
| T-shaped niche | Nixol o baix-relleu en forma de T. Element vertical + horitzontal. |
| T-shaped niche inv. | T invertida. |
| L-shaped niche | Nixol o baix-relleu en forma de L. |
| L-shaped niche inv. | L invertida. |
| Zigzag | Motiu en zigzag o chevron. |
| Stepped motif | Motiu escalonat. Documentat a DWS1 (inusual a la cultura Chachapoya). |
| Frieze / Greca | Fris o greca repetitiva. Comun a La Petaca. |
| Triangular motif | Patro triangular / chevron pintat. Documentat a DW zona sobrediatell. |
| Painted band | Banda horitzontal pintada (roja/blanca). Associada a la cornisa intranivell. |
| Square niche | Nixols quadrats en serie (hornacinas cuadradas). |
| ND | Tipus de decoracio no determinat. |

### **L_CAMPAIGN**

| **Code** | **Campaign_Name** | **Descripcio** |
|---|---|---|
| 2013 | PALP I | Primera campanya. Prospeccio i doc. inicial de La Petaca. Dr. J. Marla Toyne (UCF). |
| 2016 | PALP II | Ampliacio a Diablo Wasi. Primera doc. sistematica de DW. |
| 2021 | La Petaca Project | Doc. no invasiva integral. Fotogrametria, 360, gigafotos. Panograma Labs/UCF. |
| 2023 | PALP IV | Campanya d'excavacio arqueologica i reconstruccions 3D detallades. |

## **5. Relacions (18)**

| **Nom** | **Taula pare** | **Camp pare** | **Taula filla** | **Camp fill** | **Cascade** | **Integritat** |
|---|---|---|---|---|---|---|
| REL_SIT_SEC | L_SITES | ID | L_SECTORS | ID_Site | Delete | Enforced |
| REL_SEC_STR | L_SECTORS | ID | T_STRUCTURES | ID_Sector | Delete | Enforced |
| REL_SEC_GRP | L_SECTORS | ID | T_GROUPS | ID_Sector | Delete | Enforced |
| REL_TYP_STR | L_TYPOLOGY | ID | T_STRUCTURES | ID_Typology | Update | Enforced |
| REL_SUP_STR | L_SUPPORT | ID | T_STRUCTURES | ID_Support | Update | Enforced |
| REL_STR_SELF | T_STRUCTURES | ID | T_STRUCTURES | ID_Parent | Update | NOT enforced |
| REL_STA_STR | L_STATUS | ID | T_STRUCTURES | ID_Arch_Status | Update | Enforced |
| REL_MATSTA_STR | L_MATERIAL_STATUS | ID | T_STRUCTURES | ID_Material_Status | Update | Enforced |
| REL_VM_STR | L_VOL_METHOD | ID | T_STRUCTURES | ID_Vol_Method | Update | Enforced |
| REL_CM_STR | L_COORD_METHOD | ID | T_STRUCTURES | ID_Coord_Method | Update | Enforced |
| REL_GT_GRP | L_GROUP_TYPE | ID | T_GROUPS | ID_Group_Type | Update | Enforced |
| REL_GRP_STR | T_GROUPS | ID | T_STRUCTURES | ID_Group | Update | Enforced |
| REL_STR_DAT | T_STRUCTURES | ID | T_DATING | ID_Structure | Delete | Enforced |
| REL_STR_IND | T_STRUCTURES | ID | T_INDIVIDUALS | ID_Structure | Delete | Enforced |
| REL_STR_DEC | T_STRUCTURES | ID | T_DECORATIONS | ID_Structure | Delete | Enforced |
| REL_SB_DEC | L_STRUCT_BODY | ID | T_DECORATIONS | ID_Struct_Body | Update | Enforced |
| REL_DT_DEC | L_DEC_TYPE | ID | T_DECORATIONS | ID_Dec_Type | Update | Enforced |
| REL_STR_AFEAT | T_STRUCTURES | ID | T_ARCH_FEATURES | ID_Structure | Delete | Enforced |

*Nota: REL_STR_SELF (T_STRUCTURES autoreferenciant) usa dbRelationDontEnforceIntegrity (valor 2) per evitar conflictes en cascada. La resta de relacions usen dbRelationUpdateCascade i, si 'Delete'=si, dbRelationDeleteCascade.*

## **6. Consultes SQL (10)**

| **Nom** | **Descripcio** | **Hipotesis** |
|---|---|---|
| QRY_01_Typology_by_Site | Distribucio de tipologies per jaciment i sector (COUNT per grup). | H01, H04 |
| QRY_02_Decoration_by_Site | Presencia de cada motiu decoratiu per jaciment. Font per a chi-quadrat. | H01, H04 |
| QRY_03_Conservation_by_Sector | Distribucio dual d'estat (arquitectura + vestigis mobles) per sector. | General |
| QRY_04_C14_Structures | Estructures amb datacio C14 ordenades cronologicament. | H04 |
| QRY_05_Export_RStats | Exportacio plana completa per a R/SPSS (55+ variables llegibles). | Tots |
| QRY_06_Volumetry_by_Typology | Volumetria i area interiors per tipologia (mitja, min, max). | OE2, H01 |
| QRY_07_Export_QGIS | Exportacio espacial per a QGIS (coordenades + atributs clau). | OE2, H02, H03 |
| QRY_08_Children_of_Parent | Elements continguts per una estructura pare (parametre: [Parent ID?]). | H01, H02 |
| QRY_09_Group_Members | Membres d'un conjunt funcional ordenats per altitud (param: [Group code?]). | H03, OE2 |
| QRY_10_ChaXR_Coverage | Cobertura Chacha XR vs total per jaciment i campanya. | Metodologia |

## **7. Vocabulari arquitectonic normalitzat A-X**

24 elements organitzats per nivell constructiu (bottom-to-top) i funcio (estructural -> decoratiu). Tres sistemes compositius: E+F+G->H (plataforma), N+O+Q->P (portal), S+T->U (rafec).

| **Ll.** | **Catala** | **Angles (BD)** | **Nivell** | **Sistema** |
|---|---|---|---|---|
| A | Jacenes basals empotrades | Embedded base beams | N0 | Standalone |
| B | Basament | Base level / podium | N0 | Standalone |
| C | Socol decoratiu | Decorative socle | N0 | Standalone (tractament de B) |
| D | Muret transversal | Tie wall | N0 | Standalone |
| E | Mensules (fusta) | Timber corbels | N0 | Component de H |
| F | Bigues transversals | Transverse beams | N0 | Component de H |
| G | Filades en voladis | Corbelled masonry courses | N0 | Component de H (variant litia) |
| **H** | Plataforma volada d'acces | Corbelled access platform | N0 | SISTEMA: E+F+G -> H |
| I | Cornisa intranivell | Interlevel cornice | N0/N1 | Standalone (limit entre cossos) |
| J | Cantoneres | Corner quoins | N1 | Standalone |
| K | Pilastres estructurals | Structural pilasters | N1 | Standalone |
| L | Paraments laterals | Lateral wall faces | N1 | Standalone |
| M | Fris decoratiu en baix-relleu | Bas-relief decorative frieze | N1 | Standalone (tractament de L) |
| N | Llindar | Sill / threshold | N1 | Component de P |
| O | Brancals | Jambs | N1 | Component de P |
| **P** | Obertura d'acces | Access opening | N1 | SISTEMA: N+O+Q -> P |
| Q | Dintell | Lintel | N1 | Component de P |
| R | Coronament | Upper crown / coping | N1 | Standalone |
| S | Biga de suport del rafec | Eave-supporting beam | Sup. | Component de U |
| T | Superficie del rafec | Eave surface | Sup. | Component de U (SEMPRE PEDRA) |
| **U** | Rafec-voladis / Visera | Eave / roof overhang | Sup. | SISTEMA: S+T -> U (SEMPRE PEDRA) |
| V | Murs laterals | Lateral walls | N1 | Standalone |
| W | Mur posterior | Rear wall | N1 | Standalone (Rear_Wall_Built = construït vs roca) |
| X | Coberta de la cambra | Chamber roof | Sup. | Standalone (distint de Natural_Roof i Eave) |
