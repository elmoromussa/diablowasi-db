**ESQUEMA DE LA BASE DE DADES ARQUEOLÒGICA**

*La Petaca i Diablo Wasi (Leymebamba, Amazonas, Perú)*

chachapoya_DB_v3.bas + chachapoya_Form_v2_val.bas

Sub BuildDB() + Sub BuildForm() | Microsoft Access JET SQL | Esteve Ribera Torró

*Versió 4 del document — actualitzada segons el codi v3 de la BD (agost 2026)*

## **1. Resum general**

| **Element** | **Valor** |
|---|---|
| Taules principals | T_STRUCTURES (116 camps), T_DATING, T_INDIVIDUALS, T_GROUPS, T_DECORATIONS, T_ARCH_FEATURES |
| Taules lookup | L_SITES, L_SECTORS, L_TYPOLOGY, L_SUPPORT, L_STATUS, L_MATERIAL_STATUS, L_VOL_METHOD, L_COORD_METHOD, L_GROUP_TYPE, L_CAMPAIGN, L_STRUCT_BODY, L_DEC_TYPE |
| Total taules | 18 |
| Total relacions | 18 (inclou l'autoreferenciant de T_STRUCTURES i les relacions de T_DECORATIONS) |
| Consultes SQL | 13 (QRY_01 a QRY_13; QRY_05 i QRY_07 actualitzades amb els camps nous de v2) |
| Formulari | F_STRUCTURES - 11 pestanyes + subformularis F_DECORATIONS i F_ARCH_FEATURES. Etiquetes UI en valencià; valors emmagatzemats en anglés |
| Idioma BD | Anglés (noms de taules, camps i valors lookup). UI del formulari en valencià. Captions de columnes via DAO |
| Motor | Microsoft Access JET SQL / ACE \| cp1252 |
| Jaciments | La Petaca (WGS84: Lat -6.8311, Lon -77.8084) \| Diablo Wasi (Lat -6.8475, Lon -77.8154) |

**Canvis respecte de la versió anterior de l'esquema (v3 del document, que descrivia la BD v1): **(a) T_STRUCTURES passa de 106 a 116 camps amb quatre grups nous: detall del suport geològic (4 camps, H02), qualitat de la maçoneria (2 camps, H01/H04), paisatge i orientació de façana (2 camps, preparació QGIS) i fases constructives (2 camps, H03/H04); (b) tres consultes noves: QRY_11_Masonry_by_Site, QRY_12_Geology_Construction i QRY_13_AX_Pattern_Export; (c) el formulari passa a la versió valenciana amb captions DAO per a la vista full de dades.

## **2. T_STRUCTURES (116 camps)**

### **2.1. Identificació (7)**

| **Camp** | **Tipus** | **Descripció** |
|---|---|---|
| ID | COUNTER PK | Clau primària autonumèrica |
| Code | TEXT(20) NN | Codi PALP: [SITE][SECTOR]-[TIPTYPE][NUM] - ex: DWS1-EF01, PTC-SS-EF18 |
| ID_Sector | LONG NN | FK -> L_SECTORS |
| ID_Typology | LONG | FK -> L_TYPOLOGY |
| ID_Support | LONG | FK -> L_SUPPORT. Suport geomorfològic |
| ID_Parent | LONG | FK autoreferenciant -> T_STRUCTURES.ID. Contenció física |
| ID_Group | LONG | FK -> T_GROUPS. Agrupació funcional (alineament, xarxa) |

### **2.2. Morfologia (13)**

| **Camp** | **Tipus** | **Descripció** |
|---|---|---|
| N_Floors | INTEGER | Nombre de pisos constructius |
| Floor_Plan | TEXT(20) | Planta: Rectangular / Square / Trapezoidal / Irregular / ND |
| N_Built_Walls | INTEGER | Nombre de murs construïts (0-4) |
| Length_m | SINGLE | Llarg exterior (m) |
| Width_m | SINGLE | Ample exterior (m) |
| Height_m | SINGLE | Alçada total de l'estructura (m) |
| Approx_Height_m | SINGLE | Alçada aproximada sobre el sòl del faralló (m) |
| Access_Orientation | TEXT(5) | Orientació del vano: N / S / E / W / NE / NW / SE / SW / ND |
| Lintel | TEXT(20) | Material del dintell del vano principal: Stone / Wood / Mixed / Absent / ND |
| Natural_Roof | YESNO | Sostre natural de roca (el faralló actua com a sostre) |
| Buttresses | YESNO | Contraforts a la façana |
| Interlevel_Cornice | YESNO | Cornisa intranivell (element I del vocabulari A-X) present |
| Wooden_Stakes | YESNO | Estaques o pals de fusta verticals (elements estructurals d'ancoratge) |

### **2.2b. Detall del suport geològic (4) — NOU v2 | H02: geologia com a factor determinant**

| **Camp** | **Tipus** | **Descripció** |
|---|---|---|
| Support_Width_cm | SINGLE | Amplària de la repisa o suport (cm). Mesura directa o del model 3D |
| Support_Depth_cm | SINGLE | Profunditat de la repisa o cavitat (cm) |
| Support_Morphology | TEXT(30) | Morfologia del suport (llista al formulari): Flat ledge / Concave ledge / Fissure / Small cavity / Medium cavity / Large cavity / Vertical no support / ND |
| Support_Modified | YESNO | Evidència de modificació antròpica del suport natural (retall, anivellament) |

*Aquests camps quantifiquen la relació entre les dimensions del suport natural i les decisions constructives, nucli empíric de la hipòtesi H02.*

### **2.2c. Qualitat de la maçoneria (2) — NOU v2 | H01/H04, cf. Toyne i Anzellini 2017**

| **Camp** | **Tipus** | **Descripció** |
|---|---|---|
| Masonry_Quality | TEXT(20) | Qualitat d'execució (llista al formulari): Good / Moderate / Poor / ND |
| Masonry_Type | TEXT(30) | Tipus d'aparell (llista al formulari): Well-coursed / Irregular-coursed / Uncoursed / Mixed / ND |

*Permeten testar si la inversió tècnica varia per tipologia, sector o jaciment, seguint la línia de variabilitat i identitat social de Toyne i Anzellini (2017).*

### **2.3. Sistemes arquitectònics - Nivell 0 / Basament (10) - vocabulari A-H**

| **Camp** | **Tipus** | **Elem.** | **Descripció** |
|---|---|---|---|
| Base_Level | YESNO | **B** | Basament com a element constructiu diferenciat (pòdium) |
| Decorative_Socle | YESNO | **C** | Tractament decoratiu del basament (sòcol decoratiu) |
| Corbelled_Platform | YESNO | **H** | Plataforma volada d'accés. Sistema: E+F+G -> H |
| Timber_Brackets | YESNO | **E** | Mènsules de fusta empotrades a la roca (component de H) |
| Timber_Bracket_Count | INTEGER | **E** | Nombre de mènsules visibles |
| Transverse_Beams | YESNO | **F** | Bigues transversals (component de H) |
| Corbelled_Courses | YESNO | **G** | Filades de pedra en voladís creixent (component de H, variant lítia) |
| Corbel_Material | TEXT(20) | **E/G** | Material: Timber / Stone / Mixed / ND |
| Embedded_Base_Beams | YESNO | **A** | Jàcenes horitzontals empotrades dins la maçoneria del basament |
| Tie_Walls | YESNO | **D** | Murets perpendiculars al faralló (ancoratge/compartimentació) |

### **2.4. Sistemes arquitectònics - Nivell 1 / Cos principal (9) - vocabulari J-W**

| **Camp** | **Tipus** | **Elem.** | **Descripció** |
|---|---|---|---|
| Access_Opening | YESNO | **P** | Obertura d'accés present. Sistema: N+O+Q -> P |
| Sill | YESNO | **N** | Llindar / threshold (component de P) |
| Jambs | YESNO | **O** | Brancals: elements verticals formals del portal (component de P) |
| Lateral_Walls | YESNO | **V** | Murs laterals que donen profunditat a la cambra funerària |
| Rear_Wall | YESNO | **W** | Mur posterior present |
| Rear_Wall_Built | YESNO | **W** | Mur posterior construït (True) vs roca natural (False) |
| Recessed_Portal | YESNO | **-** | Portal retranquejat de façana (el vano està enfonsat respecte del pla exterior) |
| Structural_Pilasters | YESNO | **K** | Pilastres estructurals que flanquegen la façana (altura completa) |
| Cornice_Material | TEXT(20) | **I** | Material de la cornisa intranivell: Stone slabs / Wooden beams / Mixed / ND |

### **2.5. Sistemes - Zona superior / Coberta (6) - vocabulari R, S, T, U, X**

| **Camp** | **Tipus** | **Elem.** | **Descripció** |
|---|---|---|---|
| Eave | YESNO | **U** | Ràfec-voladís / visera present. Sistema: S+T -> U. Sempre pedra |
| Eave_Beam | YESNO | **S** | Biga de suport del ràfec (component de U; pot ser fusta) |
| Eave_Surface | YESNO | **T** | Superfície del ràfec (component de U; sempre pedra - lloses) |
| Upper_Crown | YESNO | **R** | Coronament superior / coping |
| Chamber_Roof | YESNO | **X** | Coberta construïda de la cambra (distinta de Natural_Roof i Eave) |
| Corner_Quoins | YESNO | **J** | Cantoneres: pedres de major mida disposades verticalment als angles |

### **2.6. Acabats superficials (4)**

| **Camp** | **Tipus** | **Descripció** |
|---|---|---|
| Plastered | YESNO | Arrebossat / lluïment present |
| Plaster_Color | TEXT(20) | Color de l'arrebossat (text lliure) |
| Rock_Painting | YESNO | Pintura sobre la roca present |
| Rock_Paint_Color | TEXT(20) | Color de la pintura sobre roca (text lliure) |

### **2.6b. Paisatge i orientació (2) — NOU v2 | observacional; anàlisi QGIS posterior**

| **Camp** | **Tipus** | **Descripció** |
|---|---|---|
| Facade_Orientation | TEXT(5) | Orientació de la façana (pot diferir d'Access_Orientation): N / S / E / W / NE / NW / SE / SW / ND |
| Visibility_Valley | TEXT(10) | Visibilitat des del fons de la vall (estimació de camp): High / Medium / Low / ND |

*Registre observacional de camp que prepara l'anàlisi formal de visibilitat i orientació en QGIS (H06: visibilitat i marcatge territorial).*

### **2.7. Decoració - camps booleans de resum (14) - per a khi-quadrat i filtratge ràpid**

| **Camp** | **Tipus** | **Descripció** |
|---|---|---|
| Dec_Square_Niche | YESNO | Nínxol quadrat en sèrie (hornacinas cuadradas) |
| Dec_Relief_T | YESNO | Relleu en T |
| Dec_Relief_T_Inv | YESNO | Relleu en T invertida |
| Dec_Relief_L | YESNO | Relleu en L |
| Dec_Relief_L_Inv | YESNO | Relleu en L invertida |
| Dec_Zigzag | YESNO | Motiu en zig-zag o chevron |
| Dec_Stepped | YESNO | Motiu escalonat (inusual a DW i LP; documentat a DWS1) |
| Dec_Frieze | YESNO | Fris / greca |
| Rock_Art | YESNO | Pintura rupestre present (associada a l'estructura) |
| RA_Anthropomorphic | YESNO | Motiu antropomorf a la pintura rupestre |
| RA_Zoomorphic | YESNO | Motiu zoomorf |
| RA_Geometric | YESNO | Motiu geomètric |
| RA_Abstract | YESNO | Motiu abstracte |
| RA_Decap_Scene | YESNO | Escena de decapitació documentada |

*Els camps booleans de decoració a T_STRUCTURES són per a filtratge ràpid. El registre detallat per posició a la façana (cos constructiu + tipus) és a T_DECORATIONS.*

### **2.8. Estat de conservació (6)**

| **Camp** | **Tipus** | **Descripció** |
|---|---|---|
| ID_Arch_Status | LONG | FK -> L_STATUS. Estat de l'arquitectura construïda (Good/Fair/Pre-collapse/Collapsed/ND) |
| ID_Material_Status | LONG | FK -> L_MATERIAL_STATUS. Grau de preservació dels vestigis mobles (Good/Fair/Poor/Absent/ND) |
| Looting | YESNO | Evidència de saqueig (causa, independent del grau de conservació) |
| Fire_Damage | YESNO | Evidència d'incendi |
| Animal_Activity | YESNO | Activitat animal documentada |
| Modern_Access | YESNO | Evidència d'accés modern no autoritzat |

### **2.9. Bioarqueologia (8)**

| **Camp** | **Tipus** | **Descripció** |
|---|---|---|
| Human_Remains | YESNO | Presència de restes humanes documentades |
| MNI | INTEGER | Nombre Mínim d'Individus (Minimum Number of Individuals) |
| Anatomical_Connection | YESNO | Restes en connexió anatòmica |
| Mummification | YESNO | Evidència de momificació (conservació de teixits tous) |
| Funerary_Bundles | YESNO | Fardells funeraris (embolcall tèxtil) |
| Dispersed_Remains | YESNO | Restes disperses (posició secundària) |
| Flexed_Position | YESNO | Posició flexionada documentada |
| Bone_Burning | YESNO | Cremació d'ossos |

### **2.10. Materials culturals (7)**

| **Camp** | **Tipus** | **Descripció** |
|---|---|---|
| Mat_Textiles | YESNO | Tèxtils presents |
| Mat_Wood | YESNO | Fusta cultural (artefactes, no estructural) |
| Mat_VegFiber | YESNO | Fibra vegetal |
| Mat_Ceramics | YESNO | Ceràmica (inusual en tombes aèries de DW i LP) |
| Mat_Fauna | YESNO | Restes de fauna |
| Mat_DeerAntler | YESNO | Banya de cérvol |
| Mat_Other | YESNO | Altres materials culturals |

### **2.11. Cronologia (3)**

| **Camp** | **Tipus** | **Descripció** |
|---|---|---|
| C14 | YESNO | Datació radiocarbònica disponible (detall a T_DATING) |
| Chrono_Start_Cent | INTEGER | Segle d'inici d.n.e. (ex: 10 = s. X). Estimació sense C14 o rang calibrat |
| Chrono_End_Cent | INTEGER | Segle de fi d.n.e. |

### **2.11b. Fases constructives (2) — NOU v2 | H03: saturació acumulativa / H04: seqüència operativa**

| **Camp** | **Tipus** | **Descripció** |
|---|---|---|
| Construction_Phases | INTEGER | Nombre de fases constructives identificades a l'estructura |
| Phase_Evidence | TEXT(30) | Evidència de la identificació de fases (llista al formulari): C14 / Stratigraphy / Superposition / Mortar / ND |

*Permeten registrar creixement per fases sense dependre de datacions absolutes: l'argument morfològic-estructural (superposicions, juntes, canvis d'aparell) sosté H03 en absència de C14.*

### **2.12. Volumetria i àrea (5)**

| **Camp** | **Tipus** | **Descripció** |
|---|---|---|
| Interior_Area_m2 | SINGLE | Àrea de sòl interior (m2). Per a cambres irregulars, extreta del model 3D |
| Interior_Vol_m3 | SINGLE | Volum interior útil (m3) |
| Total_Vol_m3 | SINGLE | Volum total incloent parets (m3) |
| ID_Vol_Method | LONG | FK -> L_VOL_METHOD. Mètode de càlcul |
| Vol_Notes | TEXT(200) | Notes sobre el càlcul volumètric |

### **2.13. Coordenades espacials (7)**

| **Camp** | **Tipus** | **Descripció** |
|---|---|---|
| Coord_Lat_WGS84 | DOUBLE | Latitud WGS84 (graus decimals, negatiu = S) |
| Coord_Lon_WGS84 | DOUBLE | Longitud WGS84 (graus decimals, negatiu = W) |
| Coord_E_UTM | DOUBLE | Coordenada Est UTM Zona 18S / WGS84 (m) |
| Coord_N_UTM | DOUBLE | Coordenada Nord UTM Zona 18S / WGS84 (m) |
| Altitude_masl | SINGLE | Altitud sobre el nivell del mar (m) |
| Coord_Precision_m | SINGLE | Precisió estimada de les coordenades (m) |
| ID_Coord_Method | LONG | FK -> L_COORD_METHOD |

### **2.14. Documentació digital (7)**

| **Camp** | **Tipus** | **Descripció** |
|---|---|---|
| URL_Pano | TEXT(255) | URL de la primera escena panoràmica 360 a Chacha XR |
| URL_Pano_2 | TEXT(255) | URL de la segona escena (per a estructures amb 2+ nodes) |
| URL_Giga | TEXT(255) | URL de la gigafoto |
| URL_3D | TEXT(255) | URL del model 3D fotogramètric |
| ChaXR_Documented | YESNO | Estructura publicada a la plataforma Chacha XR |
| ID_Campaign | LONG | FK -> L_CAMPAIGN. Campanya en què es va documentar sistemàticament |
| Notes | MEMO | Notes lliures (text llarg) |

## **3. Taules secundàries**

### **3.1. T_DATING**

| **Camp** | **Tipus** | **Descripció** |
|---|---|---|
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

### **3.2. T_INDIVIDUALS**

| **Camp** | **Tipus** | **Descripció** |
|---|---|---|
| ID | COUNTER PK |  |
| ID_Structure | LONG NN | FK -> T_STRUCTURES |
| Individual_No | INTEGER | Número d'individu dins de l'estructura |
| Age_Category | TEXT(20) | Infant / Juvenile / Adult / Old adult / ND |
| Sex_Category | TEXT(20) | Male / Female / Indeterminate / ND |
| Preservation | TEXT(20) | Good / Fair / Poor / ND |
| Notes | MEMO | Notes sobre l'individu |

### **3.3. T_GROUPS**

| **Camp** | **Tipus** | **Descripció** |
|---|---|---|
| ID | COUNTER PK |  |
| Group_Code | TEXT(20) NN | Codi del conjunt (ex: DWS1-C01, PTC-S-C03) |
| ID_Sector | LONG NN | FK -> L_SECTORS |
| ID_Group_Type | LONG NN | FK -> L_GROUP_TYPE |
| N_Members | INTEGER | Nombre d'estructures membres del conjunt |
| Notes | MEMO |  |

### **3.4. T_DECORATIONS - Registre detallat de decoració per estructura i posició**

| **Camp** | **Tipus** | **Descripció** |
|---|---|---|
| ID | COUNTER PK |  |
| ID_Structure | LONG NN | FK -> T_STRUCTURES (cascade delete) |
| ID_Struct_Body | LONG | FK -> L_STRUCT_BODY. On està la decoració a la façana |
| ID_Dec_Type | LONG | FK -> L_DEC_TYPE. Tipus de decoració |
| Body_No | INTEGER | Número del cos constructiu (0=basament, 1=primer cos, 2=segon cos) |
| Color | TEXT(20) | Red / White / Both / None / ND |
| Notes | TEXT(255) |  |

### **3.5. T_ARCH_FEATURES - Registre flexible d'elements no previstos a l'esquema**

| **Camp** | **Tipus** | **Descripció** |
|---|---|---|
| ID | COUNTER PK |  |
| ID_Structure | LONG NN | FK -> T_STRUCTURES (cascade delete) |
| Feature_Code | TEXT(40) | Nom de l'element (text lliure o desplegable amb suggeriments) |
| Present | YESNO |  |
| Feature_Count | INTEGER | Nombre d'unitats si escau |
| Material | TEXT(20) | Stone / Timber / Mixed / ND |
| Notes | TEXT(255) |  |

## **4. Taules lookup - valors**

*Tots els valors emmagatzemats són en anglés (convenció fixa del projecte). Les descripcions següents es tradueixen només a efectes de lectura del document.*

### **L_TYPOLOGY**

| **Codi** | **Descripció** |
|---|---|
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

| **Valor** | **Descripció** |
|---|---|
| Vertical alignment | Estructures en la mateixa vertical del faralló (clivella o estrats successius) |
| Ledge cluster | Múltiples EA compartint una mateixa repisa horitzontal |
| Platform with brackets | Plataforma aèria + mènsules que la componen o la flanquegen |
| Cave cluster | Cova + estructures construïdes al seu interior |
| Circulation network | MEN i EA-PLA que reconstrueixen una circulació aèria perduda |
| Rock art cluster | PR + EA adjacent visualment o funcionalment vinculada |
| Functional group | Qualsevol altra agrupació intrasector amb coherència funcional |

### **L_STRUCT_BODY (posició de la decoració a la façana)**

| **Code** | **Name** | **Level** | **Descripció** |
|---|---|---|---|
| N0-SOC | Socle (N0) | 0 | Sòcol decoratiu - element basal nivell 0 |
| N1-SPA | Spandrel (N1) | 1 | Parament lateral del cos principal (fora del marc del portal) |
| N1-JAM | Jamb (N1) | 1 | Brancal del portal - cos principal |
| N1-OVL | Over-lintel (N1) | 1 | Zona per damunt del dintell - àrea del fris decoratiu |
| N1-COR | Interlevel cornice | 1 | Zona de la cornisa intranivell entre cossos |
| N2-SPA | Spandrel (N2) | 2 | Parament lateral del cos superior (estructures de 2 cossos) |
| N2-JAM | Jamb (N2) | 2 | Brancal del portal - cos superior |
| ND | Not determined | -1 | Posició no determinada |

### **L_DEC_TYPE**

| **Nom** | **Descripció** |
|---|---|
| T-shaped niche | Nínxol o baix relleu en forma de T. Element vertical + horitzontal. |
| T-shaped niche inv. | T invertida. |
| L-shaped niche | Nínxol o baix relleu en forma de L. |
| L-shaped niche inv. | L invertida. |
| Zigzag | Motiu en zig-zag o chevron. |
| Stepped motif | Motiu escalonat. Documentat a DWS1 (inusual a la cultura Chachapoya). |
| Frieze / Greca | Fris o greca repetitiva. Comú a La Petaca. |
| Triangular motif | Patró triangular / chevron pintat. Documentat a DW en zona sobre el dintell. |
| Painted band | Banda horitzontal pintada (roja/blanca). Associada a la cornisa intranivell. |
| Square niche | Nínxols quadrats en sèrie (hornacinas cuadradas). |
| ND | Tipus de decoració no determinat. |

### **L_CAMPAIGN**

| **Code** | **Campaign_Name** | **Descripció** |
|---|---|---|
| 2013 | PALP I | Primera campanya. Prospecció i documentació inicial de La Petaca. Dra. J. Marla Toyne (UCF). |
| 2016 | PALP II | Ampliació a Diablo Wasi. Primera documentació sistemàtica de DW. |
| 2021 | La Petaca Project | Documentació no invasiva integral. Fotogrametria, 360, gigafotos. Panograma Labs/UCF. |
| 2023 | PALP IV | Campanya d'excavació arqueològica i reconstruccions 3D detallades. |

## **5. Relacions (18)**

| **Nom** | **Taula pare** | **Camp pare** | **Taula filla** | **Camp fill** | **Cascade** | **Integritat** |
|---|---|---|---|---|---|---|
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

*REL_STR_SELF (T_STRUCTURES autoreferenciant) usa dbRelationDontEnforceIntegrity (valor numèric 2) per a evitar conflictes en cascada. La resta de relacions usen dbRelationUpdateCascade i, si Cascade=Delete, dbRelationDeleteCascade.*

## **6. Consultes SQL (13)**

| **Nom** | **Descripció** | **Hipòtesis** |
|---|---|---|
| QRY_01_Typology_by_Site | Distribució de tipologies per jaciment i sector (COUNT per grup). | H01, H04 |
| QRY_02_Decoration_by_Site | Presència de cada motiu decoratiu per jaciment. Font per a khi-quadrat. | H01, H04 |
| QRY_03_Conservation_by_Sector | Distribució dual d'estat (arquitectura + vestigis mobles) per sector. | General |
| QRY_04_C14_Structures | Estructures amb datació C14 ordenades cronològicament. | H04 |
| QRY_05_Export_RStats | Exportació plana completa per a R/SPSS. Actualitzada en v2 amb els 10 camps nous. | Totes |
| QRY_06_Volumetry_by_Typology | Volumetria i àrea interiors per tipologia (mitjana, mín., màx.). | OE2, H01 |
| QRY_07_Export_QGIS | Exportació espacial per a QGIS. Actualitzada en v2 amb Facade_Orientation, Visibility_Valley i Masonry_Quality. LEFT JOIN per a conservar estructures sense estat. | OE3, H02, H03, H06 |
| QRY_08_Children_of_Parent | Elements continguts per una estructura pare (paràmetre: [Parent ID?]). | H01, H02 |
| QRY_09_Group_Members | Membres d'un conjunt funcional ordenats per altitud (paràm.: [Group code?]). | H03, OE3 |
| QRY_10_ChaXR_Coverage | Cobertura Chacha XR vs total per jaciment i campanya. | Metodologia |
| QRY_11_Masonry_by_Site | NOVA v2. Qualitat i tipus de maçoneria per jaciment i tipologia (COUNT per grup; exclou registres sense Masonry_Quality). | H01, H04 |
| QRY_12_Geology_Construction | NOVA v2. Encreuament suport geològic - tipologia: morfologia i modificació del suport, amb mitjanes d'amplària i profunditat (cm). | H02 |
| QRY_13_AX_Pattern_Export | NOVA v3. Exportació neta del vector A-X per estructura (amb coordenades UTM i grup) per a anàlisi de patrons en R: coocurrència (phi/Jaccard), clúster jeràrquic, I de Moran i anàlisi de correspondències. LEFT JOIN amb L_TYPOLOGY per a incloure registres parcials sense tipologia. | H01, H04, H05 |

## **7. Vocabulari arquitectònic normalitzat A-X**

24 elements organitzats per nivell constructiu (bottom-to-top) i funció (estructural -> decoratiu). Tres sistemes compositius: E+F+G->H (plataforma), N+O+Q->P (portal), S+T->U (ràfec).

| **Ll.** | **Valencià** | **Anglés (BD)** | **Nivell** | **Sistema** |
|---|---|---|---|---|
| A | Jàcenes basals empotrades | Embedded base beams | N0 | Standalone |
| B | Basament | Base level / podium | N0 | Standalone |
| C | Sòcol decoratiu | Decorative socle | N0 | Standalone (tract. de B) |
| D | Muret transversal | Tie wall | N0 | Standalone |
| E | Mènsules (fusta) | Timber corbels | N0 | Component de H |
| F | Bigues transversals | Transverse beams | N0 | Component de H |
| G | Filades en voladís | Corbelled masonry courses | N0 | Component de H (lítia) |
| **H** | **Plataforma volada d'accés** | Corbelled access platform | N0 | **SISTEMA: E+F+G -> H** |
| I | Cornisa intranivell | Interlevel cornice | N0/N1 | Standalone (límit cossos) |
| J | Cantoneres | Corner quoins | N1 | Standalone |
| K | Pilastres estructurals | Structural pilasters | N1 | Standalone |
| L | Paraments laterals | Lateral wall faces | N1 | Standalone |
| M | Fris decoratiu en baix relleu | Bas-relief decorative frieze | N1 | Standalone (tract. de L) |
| N | Llindar | Sill / threshold | N1 | Component de P |
| O | Brancals | Jambs | N1 | Component de P |
| **P** | **Obertura d'accés** | Access opening | N1 | **SISTEMA: N+O+Q -> P** |
| Q | Dintell | Lintel | N1 | Component de P |
| R | Coronament | Upper crown / coping | N1 | Standalone |
| S | Biga de suport del ràfec | Eave-supporting beam | Sup. | Component de U |
| T | Superfície del ràfec | Eave surface | Sup. | Comp. de U (SEMPRE PEDRA) |
| **U** | **Ràfec-voladís / Visera** | Eave / roof overhang | Sup. | **SISTEMA: S+T -> U (PEDRA)** |
| V | Murs laterals | Lateral walls | N1 | Standalone |
| W | Mur posterior | Rear wall | N1 | Standalone (Rear_Wall_Built) |
| X | Coberta de la cambra | Chamber roof | Sup. | Standalone |

## **8. Formulari i visualització de dades**

El formulari es genera amb chachapoya_Form_v2_val.bas i segueix la convenció bilingüe fixa del projecte: **totes les etiquetes visibles de la interfície (pestanyes, camps, capçaleres de subformularis) són en valencià, mentre que tots els valors emmagatzemats (llistes de valors dels ComboBox, continguts de les taules lookup) romanen en anglés**, per a garantir la reproduïbilitat de les exportacions analítiques.

Les 11 pestanyes de F_STRUCTURES són: 1.Id. (identificació + detall del suport geològic), 2.Arq. (morfologia + paisatge/orientació + maçoneria + fases), 3.Acab., 4.Dec. (booleans + subformulari F_DECORATIONS), 5.Estat, 6.Bio., 7.Mat., 8.Cron., 9.SIG, 10.Sist. (sistemes A-X) i 11.Extra (subformulari F_ARCH_FEATURES).

**Captions DAO: **la vista full de dades (datasheet) dels subformularis mostra per defecte els noms de camp anglesos de la TableDef, no les etiquetes del formulari. La rutina SetFieldCaptionsVal resol aquest comportament establint la propietat Caption de cada camp a nivell de TableDef via DAO (amb CreateProperty si la propietat no existeix), de manera que les capçaleres de columna es mostren en valencià sense alterar els noms de camp reals.
