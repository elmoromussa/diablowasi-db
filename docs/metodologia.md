**Universitat d'Alacant**

Master en Arqueologia Professional i Gestio Integral del Patrimoni

**DISSENY D'UNA BASE DE DADES ARQUEOLOGICA**

per a l'estudi de les necropolis de penya-segat de La Petaca i Diablo Wasi

*(Leymebamba, Amazonas, Peru, s. IX-XVI)*

Autor: Esteve Ribera Torro

Directors: Dr. Ignasi Grau Mira . Dra. J. Marla Toyne (UCF)

Curs academic 2024-2025

## **Resum**

El present document descriu el disseny, la justificacio i la implementacio de la base de dades relacional destinada a la documentacio sistematica i l'analisi estadistica de les estructures funeraries de les necropolis de penya-segat de La Petaca i Diablo Wasi (Leymebamba, Departament d'Amazonas, Peru, s. IX-XVI d.n.e.). La base de dades constitueix l'eix vertebrador de la metodologia del Treball de Fi de Master, en tant que permet centralitzar les variables arqueologiques, establir relacions jerarquiques entre elements, exportar dades per a l'analisi estadistica i vincular el registre arqueologic amb el Sistema d'Informacio Geografica (SIG) implementat en QGIS.

L'esquema segueix els principis de la tercera forma normal (3FN) i inclou 18 taules, una taula principal (T_STRUCTURES) amb 103 camps organitzats en 14 categories tematiques, 18 relacions, 10 consultes SQL i un formulari d'entrada de dades amb 11 pestanyes. Un dels avancos metodologics centrals es la normalitzacio d'un vocabulari arquitectonic bilingue (catala/angles) de 24 elements (A-X), derivat de l'analisi fotogrametrica 3D de les estructures de Diablo Wasi, que permet el registre sistematic dels sistemes constructius i la seua analisi comparativa entre jaciments i sectors.

## **1. Introduccio i objectius**

L'estudi de les necropolis de penya-segat Chachapoya planteja reptes documentals inedits en l'arqueologia andina. La verticalitat dels jaciments, l'heterogeneitat tipologica de les estructures i la multiplicitat de variables arquitectoniques, decoratives, bioarqueologiques i cronologiques exigeixen una eina de gestio de dades que vaja mes enlla d'un full de calcul pla. La base de dades relacional respon a aquesta necessitat, centralitzant en un sistema unic tota la informacio necessaria per a testar les hipotesis de recerca del TFM.

La BD serveix com a infraestructura metodologica per a la consecucio de tres objectius especifics:

- OE1. Documentar i classificar les estrategies constructives i els sistemes arquitectonics de les estructures funeraries.
- OE2. Analitzar volums, arees i relacions espacials entre estructures per sector i jaciment.
- OE3. Interpretar la relacio entre el suport geomorfologic, les decisions constructives i l'us funerari.
| **Hipotesi** | **Descripcio** | **Camps principals** |
|---|---|---|
| H01 | Ingenieria funeraria planificada | N_Floors, Corbelled_Platform, Structural_Pilasters, Jambs, Recessed_Portal, T_DECORATIONS |
| H02 | Geologia com a factor determinant | ID_Support, Natural_Roof, Altitude_masl, Approx_Height_m, Base_Level, Tie_Walls |
| H03 | Saturacio espacial acumulativa | T_GROUPS, ID_Group, QRY_07 (exportacio QGIS) |
| H04 | Sequencia operativa consistent | ID_Typology, T_DECORATIONS, T_DATING (C14), Chrono_Start/End_Cent |

*Taula 1. Hipotesis del TFM i camps de la base de dades que les operacionalitzen.*

## **2. Fonamentacio metodologica**

### **2.1. Idioma de la base de dades**

La base de dades esta completament en angles (noms de taules, camps i valors lookup). Aquesta decisio respon a tres factors: (a) la publicacio derivada del projecte (Ribera-Torro et al. 2026, Open Archaeology) usa terminologia anglesa per a elements i camps; (b) la co-directora del TFM, la Dra. J. Marla Toyne (UCF), treballa en angles; i (c) l'exportacio a R/SPSS amb capcaleres en angles es l'estandard en la publicacio cientifica internacional. Els documents del TFM (en Valencia), les presentacions a congressos (en angles o Valencia) i les presentacions al Peru (en castella) mantenen la seua llengua corresponent.

### **2.2. Marc disciplinar**

El disseny de les variables arqueologiques s'ha fonamentat en dues disciplines complementaries. L'Arqueologia de l'Arquitectura aporta el marc per a la caracteritzacio constructiva (materials, tecniques, decoracio) i la tipologia de les estructures. L'Arqueologia del Paisatge proporciona el marc per a l'analisi de la distribucio espacial i les relacions amb la geologia i l'entorn natural. La integracio de la BD amb QGIS permet implementar analisis de densitat, visibilitat i distribucio vertical directament relacionades amb les hipotesis H02 i H03.

### **2.3. Procediment de disseny**

El disseny segueix un proces iteratiu en quatre etapes:

- Inventari de variables a partir de la documentacio existent: analisi de les fitxes de camp, els papers publicats (Toyne i Anzellini 2017; Epstein i Toyne 2016; Toyne et al. 2018), el TFM Chacha XR (Ribera-Torro 2023) i l'article derivat (Ribera-Torro et al. 2026).
- Modelatge entitat-relacio (ER): identificacio de les entitats principals i les relacions entre elles.
- Normalitzacio fins a la 3FN: eliminacio de la redundancia i creacio de taules de consulta (lookup) per a tots els camps categorics.
- Validacio amb les estructures documentades de Diablo Wasi Sector 1 i Sector 4, incloent l'analisi fotogrametrica 3D dels models Metashape.
## **3. Vocabulari arquitectonic normalitzat (A-X)**

Un dels avancos metodologics centrals de la base de dades es la normalitzacio bilingue (catala/angles) d'un vocabulari de 24 elements arquitectonics (A-X), derivat de l'analisi sistematica dels models fotogrametrics 3D de Diablo Wasi. Aquest vocabulari permet el registre no ambiguu dels elements constructius observables en camp o en model 3D, i estableix la base per a l'analisi comparativa entre estructures, sectors i jaciments.

#### **Criteris d'ordenacio**

La codificacio segueix tres criteris jerarquitzats: (1) NIVELL CONSTRUCTIU: de N0 (basament) cap a N1 (cos principal) i zona superior; (2) ALCADA dins de cada nivell: de baix cap a dalt; (3) FUNCIO: primer l'element estructural portant, despres el tractament decoratiu.

#### **Sistemes compositius**

Tres elements del vocabulari son sistemes compositius, es a dir, el resultat de la combinacio obligatoria d'altres elements:

- Sistema E+F+G -> H: Mensules (E) + Bigues transversals (F) + Filades en voladis (G) generen la Plataforma volada d'acces (H). E i G son variants del mateix principi (mensules de fusta vs filades de pedra).
- Sistema N+O+Q -> P: Llindar (N) + Brancals (O) + Dintell (Q) conformen l'Obertura d'acces (P). Els tres elements emmarquen el buit.
- Sistema S+T -> U: Biga de suport (S) + Superficie del rafec (T, sempre pedra) generen el Rafec-voladis (U). La fusta pot apareixer a S (biga de suport) pero T i U son sempre lapidis.
| **Lletra** | **Catala** | **Angles (BD)** | **Nivell** | **Sistema / Notes** |
|---|---|---|---|---|
| A | Jacenes basals empotrades | Embedded base beams | N0 | Standalone, bigues horitzontals dins la maconeria del basament |
| B | Basament | Base level / podium | N0 | Standalone, cos constructiu basal |
| C | Socol decoratiu | Decorative socle | N0 | Standalone, tractament decoratiu de B |
| D | Muret transversal | Tie wall | N0 | Standalone, mur perpendicular al faralló d'ancoratge |
| E | Mensules (fusta) | Timber corbels | N0 | Component de H (sistema E+F+G -> H) |
| F | Bigues transversals | Transverse beams | N0 | Component de H |
| G | Filades en voladis (pedra) | Corbelled masonry courses | N0 | Component de H (variant litia de E) |
| H | Plataforma volada d'acces | Corbelled access platform | N0 | SISTEMA: E+F+G -> H |
| I | Cornisa intranivell | Interlevel cornice | N0/N1 | Standalone, element liminal entre cossos |
| J | Cantoneres | Corner quoins | N1 | Standalone, pedres verticals de major mida als angles |
| K | Pilastres estructurals | Structural pilasters | N1 | Standalone, elements verticals que flanquegen la facana |
| L | Paraments laterals | Lateral wall faces | N1 | Standalone, superficies observables del cos principal |
| M | Fris decoratiu en baix-relleu | Bas-relief decorative frieze | N1 | Standalone, tractament de L (detall a T_DECORATIONS) |
| N | Llindar | Sill / threshold | N1 | Component de P |
| O | Brancals | Jambs | N1 | Component de P |
| P | Obertura d'acces | Access opening | N1 | SISTEMA: N+O+Q -> P |
| Q | Dintell | Lintel | N1 | Component de P |
| R | Coronament | Upper crown / coping | N1 | Standalone, element de remat superior del cos |
| S | Biga de suport del rafec | Eave-supporting beam | Sup. | Component de U |
| T | Superficie del rafec | Eave surface | Sup. | Component de U (sempre pedra) |
| U | Rafec-voladis / Visera | Eave / roof overhang | Sup. | SISTEMA: S+T -> U (sempre pedra) |
| V | Murs laterals | Lateral walls | N1 | Standalone, murs que donen profunditat a la cambra |
| W | Mur posterior | Rear wall | N1 | Standalone, mur del fons (construït o roca natural) |
| X | Coberta de la cambra | Chamber roof | Sup. | Standalone, tancament superior de l'espai funerari construït |

*Taula 2. Vocabulari arquitectonic normalitzat A-X. Tres sistemes compositius (H, P, U) en negreta.*

*Nota sobre el material: la Superficie del rafec (T) i el Rafec-voladis (U) son sempre de pedra. La fusta pot apareixer unicament a S (biga de suport). El Mur posterior (W) pot ser construït (Rear_Wall_Built = True) o el propi faralló de roca (Rear_Wall_Built = False, relacionat amb ID_Support).*

## **4. Arquitectura de la base de dades**

### **4.1. Taules i relacions (18 taules, 18 relacions)**

| **Cat.** | **Taula** | **Contingut i funcio** |
|---|---|---|
| Principal | T_STRUCTURES | Fitxa completa (103 camps). Font principal per a l'analisi estadistica. |
| Secundaria | T_DATING | Datacions radiocarboniques (C14). N:1 amb T_STRUCTURES. |
| Secundaria | T_INDIVIDUALS | Dades individuals d'individus (edat, sexe, preservacio). N:1. |
| Secundaria | T_GROUPS | Agrupacions funcionals (alineaments, xarxes, conjunts). N:1. |
| Secundaria | T_DECORATIONS | Decoracio per estructura i cos constructiu. N:1. Clau per a H01 i H04. |
| Secundaria | T_ARCH_FEATURES | Registre flexible d'elements no previstos al schema. N:1. |
| Lookup | L_SITES | Jaciments: La Petaca, Diablo Wasi. Coordenades de referencia. |
| Lookup | L_SECTORS | Sectors (LP: N, Central, Superior, S; DW: Sectors 1-6). |
| Lookup | L_TYPOLOGY | Tipologia (EA-MAU, EA-CAM, EA-PLA-R/V, NIX, CAV, PR, MEN, MIX, ND). |
| Lookup | L_SUPPORT | Suport geomorfologic (repisa natural amplia/estreta, artificial, cavitat...). |
| Lookup | L_STATUS | Estat conservacio arquitectura (Good/Fair/Pre-collapse/Collapsed/ND). |
| Lookup | L_MATERIAL_STATUS | Grau preservacio vestigis mobles (Good/Fair/Poor/Absent/ND). |
| Lookup | L_VOL_METHOD | Metode de calcul volumetric (L x W x H / Fotogrametric / Estimacio / ND). |
| Lookup | L_COORD_METHOD | Metode d'obtencio coordenades (Drone RTK, Fotogrametria, GPS mobil...). |
| Lookup | L_GROUP_TYPE | Tipus d'agrupacio funcional (Vertical alignment, Circulation network...). |
| Lookup | L_CAMPAIGN | Campanyes: 2013 PALP I, 2016 PALP II, 2021 La Petaca Project, 2023 PALP IV. |
| Lookup | L_STRUCT_BODY | Posicio de la decoracio a la facana (N0-SOC, N1-JAM, N1-OVL, N1-COR...). |
| Lookup | L_DEC_TYPE | Tipus de decoracio (T-shaped niche, Triangular motif, Painted band...). |

*Taula 3. Les 18 taules de la base de dades amb la seua funcio.*

## **5. Variables arqueologiques de T_STRUCTURES (103 camps)**

La taula T_STRUCTURES concentra 103 camps en 14 categories tematiques. La taula seguent resumeix les categories i el nombre de camps per categoria:

| **Categoria** | **Camps principals** | **N** |
|---|---|---|
| Identificacio | Code, ID_Sector, ID_Typology, ID_Support, ID_Parent, ID_Group | 7 |
| Morfologia | N_Floors, Floor_Plan, N_Built_Walls, Length/Width/Height_m, Approx_Height_m, Access_Orientation, Lintel, Natural_Roof, Buttresses, Interlevel_Cornice, Wooden_Stakes | 13 |
| Sistemes N0 (vocab. A-H) | Base_Level (B), Decorative_Socle (C), Tie_Walls (D), Timber_Brackets (E), Transverse_Beams (F), Corbelled_Courses (G), Corbelled_Platform (H), Embedded_Base_Beams (A), Corbel_Material, Timber_Bracket_Count | 10 |
| Sistemes N1 (vocab. J-W) | Corner_Quoins (J), Structural_Pilasters (K), Jambs (O), Sill (N), Access_Opening (P), Recessed_Portal, Lateral_Walls (V), Rear_Wall (W), Rear_Wall_Built | 9 |
| Cornisa i coberta (vocab. I, S-X) | Interlevel_Cornice (I), Cornice_Material, Eave_Beam (S), Eave_Surface (T), Eave (U), Upper_Crown (R), Chamber_Roof (X) | 7 |
| Acabats superficials | Plastered, Plaster_Color, Rock_Painting, Rock_Paint_Color | 4 |
| Decoracio (booleans) | Dec_Square_Niche, Dec_Relief_T/T_Inv/L/L_Inv, Dec_Zigzag, Dec_Stepped, Dec_Frieze, Rock_Art, RA_Anthropomorphic/Zoomorphic/Geometric/Abstract/Decap_Scene | 14 |
| Estat de conservacio | ID_Arch_Status, ID_Material_Status, Looting, Fire_Damage, Animal_Activity, Modern_Access | 6 |
| Bioarqueologia | Human_Remains, MNI, Anatomical_Connection, Mummification, Funerary_Bundles, Dispersed_Remains, Flexed_Position, Bone_Burning | 8 |
| Materials culturals | Mat_Textiles, Mat_Wood, Mat_VegFiber, Mat_Ceramics, Mat_Fauna, Mat_DeerAntler, Mat_Other | 7 |
| Cronologia | C14, Chrono_Start_Cent, Chrono_End_Cent | 3 |
| Volumetria i area | Interior_Area_m2, Interior_Vol_m3, Total_Vol_m3, ID_Vol_Method, Vol_Notes | 5 |
| Coordenades espacials | Coord_Lat/Lon_WGS84, Coord_E/N_UTM, Altitude_masl, Coord_Precision_m, ID_Coord_Method | 7 |
| Documentacio digital | URL_Pano, URL_Pano_2, URL_Giga, URL_3D, ChaXR_Documented, ID_Campaign, Notes | 7 |

*Taula 4. Categories tematiques de T_STRUCTURES amb nombre de camps per categoria (total: 103).*

### **5.1. Estat de conservacio dual**

La BD distingeix entre dos dimensions de l'estat de conservacio: ID_Arch_Status (-> L_STATUS) mesura el grau de preservacio de l'arquitectura construida (escala: Good/Fair/Pre-collapse/Collapsed/ND), mentre que ID_Material_Status (-> L_MATERIAL_STATUS) mesura el grau de preservacio dels vestigis mobles (escala: Good/Fair/Poor/Absent/ND). La causa del deteriorament queda registrada pels camps booleans Looting, Fire_Damage, Animal_Activity i Modern_Access, que son independents dels graus de conservacio. El creuament entre grau i causa permet reconstruir la historia postdeposicional de cada estructura.

### **5.2. T_DECORATIONS: decoracio per cos constructiu**

La taula T_DECORATIONS permet el registre detallat de decoracions per estructura, complementant els camps booleans de T_STRUCTURES. Cada registre inclou: ID_Struct_Body (-> L_STRUCT_BODY: N0-SOC, N1-JAM, N1-OVL, N1-COR, N2-SPA, N2-JAM, ND), ID_Dec_Type (-> L_DEC_TYPE), Body_No (numero de cos constructiu), Color (Red/White/Both/None/ND) i Notes. Aquesta estructura permet analitzar si, per exemple, els motius en T es concentren als brancals (N1-JAM) i el fris triangular al sobrediatell (N1-OVL), com s'observa als models 3D de DW Sector 1.

Els camps booleans de decoracio a T_STRUCTURES (Dec_Square_Niche, Dec_Relief_T...) es mantenen per al filtratge rapid en consultes estadistiques (chi-quadrat). T_DECORATIONS s'usa per al registre posicional i tipologic detallat.

### **5.3. T_ARCH_FEATURES: registre flexible**

La taula T_ARCH_FEATURES permet registrar qualsevol element constructiu no previst al schema principal sense modificar l'estructura de la BD (Feature_Code, Present, Feature_Count, Material, Notes). Opera com a 'safety net' per a elements nous que puguen aparèixer en campanyes futures, evitant la proliferacio indefinida de nous camps booleans a T_STRUCTURES.

## **6. Estrategia de jerarquia i agrupacio**

La BD implementa dos mecanismes complementaris i independents per a gestionar les relacions entre elements:

- ID_Parent (autoreferenciant T_STRUCTURES.ID): contenicio fisica o estructural. Exemple: EA-MAU construït dins d'una CAV apunta a la cova com a pare. Permet jerarquies de N nivells.
- ID_Group (-> T_GROUPS): agrupacio funcional. Exemple: 4 EA-MAU en alineament vertical pertenyen al mateix T_GROUPS sense que cap continga l'altre. T_GROUPS classifica l'agrupacio per ID_Group_Type (L_GROUP_TYPE: Vertical alignment, Ledge cluster, Platform with brackets, Cave cluster, Circulation network, Rock art cluster, Functional group).

*Regla practica: usar ID_Parent quan un element no te sentit arqueologic independent del seu contenidor (un EA-MAU dins una CAV). Usar ID_Group quan els elements son arqueologicament independents pero formen una unitat d'estudi (alineament de mensules que reconstrueixen una xarxa de circulacio perduda).*

## **7. Integracio espacial amb QGIS**

El camp Coord_Lat_WGS84 / Coord_Lon_WGS84 permet carregar T_STRUCTURES com a capa de punts a QGIS (Layer -> Add Delimited Text Layer). Les coordenades individuals s'obtenen preferentment dels models Metashape georeferencitats (File -> Export -> Export Markers en format CSV, coordenades UTM zona 18S / WGS84). La consulta QRY_07_Export_QGIS exporta els camps mes rellevants per a la visualitzacio cartografica: tipologia, estat de conservacio dual, volumetria, cronologia, MNI i documentacio digital.

La naturalesa vertical dels farallons requereix treballar en dos nivells: l'ortofoto vertical de cada sector (exportada des de Metashape) per a relacions relatives entre estructures, i les coordenades absolutes (UTM + altitud) per a la distribucio espacial general del jaciment.

## **8. Analisi volumetrica i areal**

La BD registra tres mesures de volumetria: Interior_Area_m2 (area de sol interior, m2), Interior_Vol_m3 (volum interior util, m3) i Total_Vol_m3 (volum total incloent parets). Per a estructures de geometria regular (EA-MAU), el volum interior es calcula com (Length_m - 2e) x (Width_m - 2e) x Height_m, on e es l'espessor de parets (~0.20-0.30 m documentat a DW). Per a cavitats irregulars (EA-CAM, CAV), el volum s'extreu directament del model fotogrametric 3D via MeshLab o CloudCompare. El camp ID_Vol_Method registra el procediment per a ponderar la comparabilitat entre mesures.

La correlacio entre Interior_Area_m2 / Interior_Vol_m3 i MNI permet avaluar si les tombes de major capacitat allotjaven mes individus (H01), mentre que la variacio del volum per tipologia i jaciment (QRY_06) caracteritza les estrategies d'adaptacio al suport geomorfologic (H02).

## **9. Consultes SQL per a l'analisi estadistica**

La BD inclou 10 consultes SQL predissenyades que cobreixen els principals analisiss del TFM:

| **Consulta** | **Funcio analitica** | **Hipotesis** |
|---|---|---|
| QRY_01_Typology_by_Site | Distribucio de tipologies per jaciment i sector. | H01, H04 |
| QRY_02_Decoration_by_Site | Presencia/absencia de cada motiu decoratiu per jaciment (font per a chi-quadrat). | H01, H04 |
| QRY_03_Conservation_by_Sector | Distribucio de l'estat de conservacio arquitectura + vestigis mobles per sector. | General |
| QRY_04_C14_Structures | Estructures amb datacio C14 ordenades cronologicament. | H04 |
| QRY_05_Export_RStats | Exportacio plana completa per a R/SPSS (55+ variables per estructura). | Tots |
| QRY_06_Volumetry_by_Typology | Volumetria i area interiors per tipologia i jaciment (mitja, min, max). | OE2, H01 |
| QRY_07_Export_QGIS | Exportacio espacial per a QGIS (coordenades + atributs clau). | OE2, H02, H03 |
| QRY_08_Children_of_Parent | Elements fills d'un element pare (contingut d'una cova, per exemple). | H01, H02 |
| QRY_09_Group_Members | Membres d'un conjunt funcional ordenats per altitud. | H03, OE2 |
| QRY_10_ChaXR_Coverage | Cobertura Chacha XR vs total per jaciment i campanya. | Metodologia |

*Taula 5. Les 10 consultes SQL predissenyades i la seua vinculacio amb els objectius del TFM.*

## **10. Formulari principal F_STRUCTURES**

El formulari F_STRUCTURES centralitza l'entrada de dades amb 11 pestanyes tematiques, dos subformularis vinculats i tots els camps FK configurats com a ComboBox (desplegables):

| **Pestanya** | **Contingut** |
|---|---|
| 1. Ident. | Code, ID_Sector, ID_Typology, ID_Support, ID_Parent, ID_Group |
| 2. Arq. | Morfologia: N_Floors, Floor_Plan, dimensions, Access_Orientation, Lintel, Natural_Roof, Buttresses, Interlevel_Cornice, Wooden_Stakes |
| 3. Acab. | Acabats: Plastered, Plaster_Color, Rock_Painting, Rock_Paint_Color |
| 4. Dec. | 14 camps booleans de decoracio + subformulari F_DECORATIONS (T_DECORATIONS per cos i tipus) |
| 5. Estat | Conservacio: ID_Arch_Status, ID_Material_Status, Looting, Fire_Damage, Animal_Activity, Modern_Access |
| 6. Bio. | Bioarqueologia: Human_Remains, MNI, Anatomical_Connection, Mummification, Funerary_Bundles, Dispersed_Remains, Flexed_Position, Bone_Burning |
| 7. Mat. | Materials: Mat_Textiles, Mat_Wood, Mat_VegFiber, Mat_Ceramics, Mat_Fauna, Mat_DeerAntler, Mat_Other |
| 8. Cron. | Cronologia + Volumetria: C14, Chrono_Start/End_Cent, ID_Campaign, Interior_Area_m2, Interior_Vol_m3, Total_Vol_m3, ID_Vol_Method, Vol_Notes |
| 9. SIG | Coordenades WGS84 + UTM, Altitude_masl, ID_Coord_Method, URLs (panorames/gigafotos/models 3D), ChaXR_Documented, Notes |
| 10. Systems | Sistemes A-X: NIVELL 0 (B,C,D,E+F+G->H), OBERTURA (N+O+Q->P), FACANA/SUPERIOR (I,J,K,R,S+T->U,V,W,X) |
| 11. Extra | Elements confirmats extres (A: jac. basals, V: murs lat., W: mur post.) + subformulari F_ARCH_FEATURES (T_ARCH_FEATURES) |

*Taula 6. Estructura del formulari F_STRUCTURES amb 11 pestanyes tematiques.*

## **11. Implementacio tecnica**

### **11.1. Scripts VBA**

La BD s'implementa en Microsoft Access via dos scripts VBA consolidats:

- chachapoya_01_DB.bas - Sub BuildDB(): crea les 18 taules en ordre de dependencia, pobla els 12 lookups, estableix les 18 relacions i genera les 10 consultes. S'executa sobre una BD en blanc.
- chachapoya_02_Form.bas - Sub BuildForm(): crea els subformularis F_DECORATIONS i F_ARCH_FEATURES, construeix F_STRUCTURES amb 11 pestanyes i configura tots els ComboBox per ControlSource. S'executa despres de BuildDB().
### **11.2. Restriccions tecniques de VBA/Access**

Els scripts apliquen les seguents restriccions tecniques del motor JET SQL i del VBA d'Access:

- Cap continuacio de linia (& _): tot el SQL llarg usa el patro sql = sql & '...' en linies separades.
- Codificacio cp1252 (Windows-1252) sense caracters no-ASCII en cadenes de codi.
- Paraules reservades evitades: Level -> Body_Level (LEVEL es paraula reservada en JET SQL).
- Patro tmpName = f.Name per a CreateControl sobre formularis nous, seguit de DoCmd.Save acForm, tmpName i DoCmd.Rename FRM, acForm, tmpName.
- Per a N taules en JOIN necessitem N-1 parentesis d'obertura al FROM en JET SQL.
- La relacio autoreferenciant T_STRUCTURES.ID -> T_STRUCTURES.ID_Parent usa el flag dbRelationDontEnforceIntegrity (valor numeric 2) per evitar conflictes en cascada.
### **11.3. Exportacio per a analisi estadistica**

L'exportacio per a R o SPSS es realitza via QRY_05_Export_RStats (Dades externes -> Exportar -> Text/Excel). Tots els camps categorics de la consulta estan en format llegible (no com a IDs numerics) gracies als JOINs amb les taules lookup. En R, el paquet readxl o read.csv permet la carrega directa.

## **12. Limitacions i perspectives**

La limitacio principal es la manca de dades de camp per a moltes de les estructures de La Petaca Sector Sud (>=96 estructures identificades pero no totes amb fitxa completa). El camp ID_Campaign permet identificar quines estructures han estat documentades i en quina campanya, facilitant la deteccio de biaixos documentals.

La taula T_ARCH_FEATURES aporta una solucio al problema de l'escalabilitat del schema: els elements constructius nous que puguen aparèixer en campanyes futures es poden registrar com a entrades flexibles sense modificar l'estructura de la BD.

La seccio de vocabulari arquitectonic A-X es el primer intent de normalitzacio terminologica per a les estructures funeraries Chachapoya en la bibliografia. La seua aplicacio a La Petaca i la validacio creuada amb la literatura existent (Toyne i Anzellini 2017; Epstein i Toyne 2016; Toyne et al. 2018) representa una de les aportacions originals del present TFM.

De cara a publicacions futures, es planteja la migracio a GeoPackage (.gpkg) per a la integracio nativa amb QGIS, mantenint Access com a interficie d'entrada de dades amb sincronitzacions periodiques.

## **13. References**

Epstein, L.; Toyne, J.M. (2016). When Space Is Limited: A Spatial Exploration of Pre-Hispanic Chachapoya Mortuary and Ritual Microlandscape. A: Osterholtz, A.J. (ed.), Theoretical Approaches to Analysis and Interpretation of Commingled Human Remains. Springer, Switzerland, pp. 97-124.

Ribera-Torro, E. (2023). Chacha XR. Una experiencia immersiva de no-ficcio per l'arqueologia Chachapoya. Treball de Fi de Master, Master Universitari en Arts Visuals i Multimedia, Universitat Politecnica de Valencia.

Ribera-Torro, E.; Toyne, J.M.; Del Aguila, R.; Ribera, J.A.; Galexner, J.; Anzellini, A.; Pans, M. (2026). Extended Reality on Chachapoya Cliffside Necropolises: From Digital Documentation to Public Engagement. Open Archaeology, 12(1). DOI 10.1515/opar-2025-0071.

Toyne, J.M.; Anzellini, A. (2017). Sociedad, identidad y variedad en los mausoleos de La Petaca, Chachapoyas. Boletin de Arqueologia PUCP, 23, 231-257.

Toyne, J.M.; Anzellini, A.; Epstein Miculas, L.; Mejias Pitti, I.; Puig Castell, J.; Guinot Castello, S. (2018). Going Vertical: Using Vertical Progression Techniques to Explore a Cliff Necropolis in Late Precolumbian Chachapoyas, Peru. Advances in Archaeological Practice, 6. DOI 10.1017/aap.2018.31.
