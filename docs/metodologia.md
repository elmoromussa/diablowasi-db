**Universitat d'Alacant**

Departament de Prehistoria, Arqueologia, H. Antiga, Filologia Grega i Filologia Llatina

**DISSENY D'UNA BASE DE DADES ARQUEOLOGICA**

per a l'estudi de les necropolis de penya-segat de La Petaca i Diablo Wasi

*(Leymebamba, Amazonas, Peru)*

*Versio 2.0 - amb terminologia arquitectonica normalitzada*

Treball de Fi de Master - Master en Arqueologia Professional i Gestio Integral del Patrimoni

Autor: Esteve Ribera Torro

Directors: Dr. Ignasi Grau Mira . Dra. J. Marla Toyne

Curs academic 2024-2025

## **Resum**

El present document descriu, en la seua versio 2.0 actualitzada, el disseny, la justificacio i la implementacio de la base de dades relacional destinada a la documentacio sistematica i l'analisi estadistica de les estructures funeraries de les necropolis de penya-segat de La Petaca i Diablo Wasi (Leymebamba, Departament d'Amazonas, Peru). Respecte a la versio inicial, aquesta segona versio incorpora: (a) migracioompleta a l'angles com a idioma de la BD per a garantir la projecciointernacional i la col·laboracio amb la Universitat de Florida Central (UCF); (b) un vocabulari arquitectonic normalitzat (elements A-P) derivat de l'analisi fotogrametrica 3D dels models de Diablo Wasi; (c) el desdoblament de l'estat de conservacio en dos camps independents per a l'arquitectura (ID_Arch_Status) i per als vestigis mobles (ID_Material_Status); (d) la creacio de la taula T_DECORATIONS per al registre detallat de decoracions per cos constructiu; i (e) la taula flexible T_ARCH_FEATURES per a la documentacio d'elements constructius no catalogats previament.

L'esquema final inclou 18 taules, 97 camps a la taula principal T_STRUCTURES, 18 relacions, 10 consultes SQL i un formulari principal amb 11 pestanyes tematiques, incloent subformularis vinculats a T_DECORATIONS i T_ARCH_FEATURES.

## **1. Introduccio**

L'estudi de les necropolis de penya-segat Chachapoya planteja reptes documentals i analitics inedits. La verticalitat dels jaciments, l'heterogeneitat tipologica i la multiplicitat de variables exigeixen una eina de gestio de dades que vaja mes enlla d'un full de calcul pla. La base de dades relacional respon a aquesta necessitat, i en la seua versio 2.0 incorpora un nivell addicional de precisio arquitectonica basat en l'analisi directa dels models fotogrametrics 3D de les estructures de Diablo Wasi.

En el context del present TFM, la BD serveix com a infraestructura metodologica per a la consecucio dels tres objectius especifics:

- Documentar i classificar les estrategies constructives de les estructures funeraries.
- Analitzar volums, arees i relacions espacials entre estructures.
- Interpretar la relacio entre geologia, construccio i us funerari.

La taula seguent sintetitza la correspondencia actualitzada entre hipotesis i camps de la BD:

| **Hipotesi** | **Descripcio sintetica** | **Camps i taules rellevants** |
|---|---|---|
| H01 | Ingenieria funeraria planificada | N_Floors, Floor_Plan, Structural_Pilasters, Access_Opening, Sill, Corbelled_Platform, Recessed_Portal, T_DECORATIONS |
| H02 | Geologia com a factor determinant | ID_Support, Natural_Roof, Altitude_masl, Approx_Height_m, Base_Level |
| H03 | Saturacio espacial acumulativa | T_GROUPS, ID_Group, distribucio espacial via QGIS (QRY_07) |
| H04 | Sequencia operativa consistent | ID_Typology, Dec_*, T_DECORATIONS, T_DATING (C14) |

*Taula 1. Correspondencia entre les hipotesis del TFM i els camps de la base de dades v2.0.*

## **2. Fonamentacio metodologica**

### **2.1. Idioma de la base de dades**

La versio 2.0 adopta l'angles com a idioma exclusiu de la BD (noms de taules, camps i valors lookup). Aquesta decisio respon a tres factors: (a) l'article publicat a Open Archaeology (Ribera-Torro et al. 2026) ja usa terminologia anglesa per als elements i camps; (b) la co-directora del TFM, la Dra. J. Marla Toyne (UCF), treballa en angles i pot necessitar acces o revisio de la BD; (c) l'exportacio de les consultes (QRY_05) a R o SPSS amb capcaleres en angles es l'estandard de publicacio cientifica. Els documents del TFM, les presentacions a congressos en catala i les presentacions al Peru en castella mantenen la seua llengua corresponent.

### **2.2. Terminologia arquitectonica normalitzada (elements A-P)**

Un dels avancos metodologics mes significatius de la versio 2.0 es la normalitzacio bilingue (catala/angles) del vocabulari arquitectonic, derivada de l'analisi fotogrametrica 3D dels models de Diablo Wasi Sectors 1 i 4. Aquesta terminologia cobreix els elements constructius que la Taula 2 sintetitza, organitzats per sistema i nivell:

| **Codi** | **Catala** | **Angles (BD)** | **Nivell** | **Sistema** |
|---|---|---|---|---|
| A | Brancals | Jambs | N1 | Portal: A+B+P -> O |
| B | Dintell | Lintel | N1 | Portal |
| C | Pilastres estructurals | Structural pilasters | N1 | Facana |
| D | Paraments laterals | Lateral wall faces | N1 | Facana |
| E | Fris decoratiu en baix-relleu | Bas-relief decorative frieze | N1 | Decoracio |
| F | Biga de suport del rafec | Eave-supporting beam | Sup. | Coberta: F -> G |
| G | Rafec / voladis de coberta | Eave / roof overhang | Sup. | Coberta |
| H | Coronament | Upper crown / coping | Sup. | Coronament |
| I | Basament | Base level / podium | N0 | Base |
| J | Socol decoratiu | Decorative socle | N0 | Base: J <- I |
| K | Mensules (fusta) | Timber corbels | N0 | Plataforma: K+L+N -> M |
| L | Bigues transversals | Transverse beams | N0 | Plataforma |
| M | Plataforma volada d'acces | Corbelled access platform | N0 | Plataforma |
| N | Filades en voladis | Corbelled masonry courses | N0 | Plataforma (variant pedra) |
| O | Obertura d'acces | Access opening | N1 | Portal |
| P | Llindar | Sill / threshold | N1 | Portal |

*Taula 2. Vocabulari arquitectonic normalitzat (elements A-P) amb codis, termes bilingues i organitzacio per sistemes. Els tres sistemes principals son: K+L+N->M (plataforma), A+B+P->O (portal), I conte J.*

*Nota metodologica: la presencia de cornisa intranivell (Q) entre el Nivell 0 (I, J, K, L, M, N) i el Nivell 1 (A-H, O) permet identificar estructures de 2 cossos constructius. La distincia Nivell 0 / Nivell 1 es la base de la taula L_STRUCT_BODY per al registre posicional de decoracions a T_DECORATIONS.*

### **2.3. Desdoblament de l'estat de conservacio**

La versio 2.0 desdobla el camp ID_Status en dos camps independents amb logiques avaluatives especifiques:

- ID_Arch_Status (-> L_STATUS): estat de l'arquitectura construida (Good / Fair / Pre-collapse / Collapsed / ND). Recull el grau de preservacio de l'estructura immobil.
- ID_Material_Status (-> L_MATERIAL_STATUS): estat dels vestigis mobles (Good / Fair / Poor / Absent / ND). Recull el grau de preservacio del registre arqueologic moble, independentment de la causa (que queda registrada pels camps booleans Looting, Animal_Activity, Fire_Damage, Modern_Access).

Aquesta distinccio es arqueologicament rellevant perque permet identificar, per exemple, estructures amb arquitectura ben preservada (Good) pero sense vestigis mobles (Absent) per saqueig, o estructures parcialment col·lapsades (Pre-collapse) que han preservat momies intactes (Good).

## **3. Arquitectura de la base de dades**

### **3.1. Taules i relacions (versio 2.0)**

La BD v2.0 consta de 18 taules organitzades en quatre categories. Respecte a la versio inicial (13 taules), s'han afegit: L_MATERIAL_STATUS, L_STRUCT_BODY, L_DEC_TYPE, T_DECORATIONS i T_ARCH_FEATURES.

| **Categoria** | **Taula** | **Contingut i funcio** |
|---|---|---|
| Taula principal | T_STRUCTURES | Fitxa completa (97 camps). Font principal per a l'analisi estadistica. |
| Secundaries | T_DATING | Datacions radiocarboniques (C14). Relacio N:1. |
| Secundaries | T_INDIVIDUALS | Dades individuals d'individus quan existeix informacio desagregada. |
| Secundaries | T_GROUPS | Agrupacions funcionals (alineaments, xarxes de circulacio, conjunts). |
| Secundaries | T_DECORATIONS | Registre detallat de decoracions per estructura, cos i tipus (N:1 amb T_STRUCTURES). |
| Secundaries | T_ARCH_FEATURES | Registre flexible d'elements constructius no previstos al schema principal. |
| Lookup | L_SITES | Jaciments (La Petaca, Diablo Wasi) amb coordenades de referencia. |
| Lookup | L_SECTORS | Sectors per jaciment (LP: N, Central, Superior, S; DW: Sectors 1-6). |
| Lookup | L_TYPOLOGY | Tipologia de l'element (EA-MAU, EA-CAM, EA-PLA-R, EA-PLA-V, NIX, CAV, PR, MEN...). |
| Lookup | L_SUPPORT | Suport geomorfologic (repisa natural amplia/estreta, artificial, cavitat, grieta...). |
| Lookup | L_STATUS | Estat conservacio arquitectura (Good / Fair / Pre-collapse / Collapsed / ND). |
| Lookup | L_MATERIAL_STATUS | Grau preservacio vestigis mobles (Good / Fair / Poor / Absent / ND). |
| Lookup | L_VOL_METHOD | Metode de calcul del volum interior (L x W x H, model fotogrametric, estimacio, ND). |
| Lookup | L_COORD_METHOD | Metode d'obtencio de coordenades (drone RTK, fotogrametria, GPS mobil...). |
| Lookup | L_GROUP_TYPE | Tipus d'agrupacio funcional (alineament vertical, xarxa de circulacio...). |
| Lookup | L_CAMPAIGN | Any de campanya (2013 PALP I / 2016 PALP II / 2021 La Petaca Project / 2023 PALP IV). |
| Lookup | L_STRUCT_BODY | Posicio de la decoracio a la facana (N0-SOC, N1-JAM, N1-OVL, N1-COR, N2-SPA...). |
| Lookup | L_DEC_TYPE | Tipus de decoracio (T-shaped niche, Triangular motif, Painted band, Frieze...). |

*Taula 3. Les 18 taules de la base de dades v2.0 i la seua funcio.*

## **4. Variables arqueologiques de T_STRUCTURES (97 camps)**

### **4.1. Identificacio, tipologia i suport geomorfologic**

Cada registre s'identifica amb un codi estructural (Code) seguint la convencio PALP: [JACIMENT][SECTOR]-[TIPUS][NUM]. La tipologia (ID_Typology -> L_TYPOLOGY) inclou 10 valors que cobreixen totes les casuistiques: EA-MAU Mausoleum/Chullpa, EA-CAM Funerary Chamber, EA-PLA-R Ledge Platform, EA-PLA-V Aerial Platform, NIX Natural Niche, CAV Cave/Cavern, PR Rock Art, MEN Isolated Bracket, MIX Mixed, ND Undetermined.

### **4.2. Sistemes arquitectonics (Nivell 0 i Nivell 1)**

La versio 2.0 organitza els camps arquitectonics en dos nivells constructius explicits:

NIVELL 0 (Base/Podium): Base_Level, Decorative_Socle, Corbelled_Platform, Timber_Brackets, Timber_Bracket_Count, Transverse_Beams, Corbelled_Courses, Corbel_Material, Embedded_Base_Beams, Tie_Walls. El sistema K+L+N->M captura la cadena constructiva: corbells de fusta (K, Timber_Brackets) o filades en voladis de pedra (N, Corbelled_Courses) + bigues transversals (L, Transverse_Beams) -> plataforma volada (M, Corbelled_Platform).

NIVELL 1 (Cos principal): Access_Opening, Sill, Recessed_Portal, Structural_Pilasters, Cornice_Material, Eave, Eave_Beam, Upper_Crown, Corner_Quoins. El sistema A+B+P->O captura el portal: brancals (A, Jambs), dintell (B, Lintel), llindar (P, Sill) -> obertura d'acces (O, Access_Opening). El camp Recessed_Portal (YESNO) registra el retranqueig de facana documentat als models 3D de DW.

### **4.3. Decoracio (camps booleans + T_DECORATIONS)**

La v2.0 manté els camps booleans de decoracio a T_STRUCTURES (Dec_Square_Niche, Dec_Relief_T, Dec_Relief_T_Inv, Dec_Relief_L, Dec_Relief_L_Inv, Dec_Zigzag, Dec_Stepped, Dec_Frieze, Rock_Art i cinc camps RA_*) per a filtratge rapid i calcul de chi-quadrat. Paral·lelament, la nova taula T_DECORATIONS permet un registre detallat per estructura: cada fila recull el cos constructiu on apareix la decoracio (ID_Struct_Body -> L_STRUCT_BODY: N0-SOC, N1-JAM, N1-OVL, N1-COR, N2-SPA, N2-JAM), el tipus (ID_Dec_Type -> L_DEC_TYPE), el numero de cos (Body_No), el color i notes. Aquesta estructura permet analitzar si, per exemple, els motius en T es concentren als brancals (N1-JAM) i el fris triangular al sobredialintel (N1-OVL), com s'observa als models 3D de DW Sector 1.

### **4.4. Estat de conservacio dual**

La distincio entre ID_Arch_Status i ID_Material_Status (seccio 2.3) es complementa pels camps booleans de causa: Looting, Fire_Damage, Animal_Activity, Modern_Access. L'analisi creuada entre el grau de preservacio i la causa permet reconstruir la historia postdeposicional de cada estructura.

### **4.5. Bioarqueologia i materials**

MNI (Minimum Number of Individuals, nota: NMI en l'anterior versio), Mummification, Funerary_Bundles, Dispersed_Remains, Flexed_Position, Bone_Burning. Materials: Mat_Textiles, Mat_Wood, Mat_VegFiber, Mat_Ceramics, Mat_Fauna, Mat_DeerAntler, Mat_Other.

### **4.6. Volumetria i aree**

La versio 2.0 afegeix el camp Interior_Area_m2 (SINGLE) per a la superfice de sola interior, complementant Interior_Vol_m3 i Total_Vol_m3. Per a mausoleus, Interior_Area_m2 = (Length_m - 2e) x (Width_m - 2e) on e = espessor de parets. Per a cavitats irregulars, extret del model 3D via MeshLab. La consulta QRY_06 inclou ara Mean_Area juntament amb Mean_Vol.

### **4.7. T_ARCH_FEATURES: registre flexible**

La nova taula T_ARCH_FEATURES permet registrar qualsevol element constructiu no previst al schema principal sense modificar l'estructura de la BD. Cada fila inclou: Feature_Code (text lliure o desplegable amb suggeriments), Present (YESNO), Feature_Count, Material i Notes. Esta taula opera com a safety net per a elements nous que puguen aparèixer en campanyes futures.

## **5. Estrategia de jerarquia i agrupacio**

La BD manté els dos mecanismes complementaris de la v1.0:

- ID_Parent (autoreferenciant T_STRUCTURES.ID): contenicio fisica o estructural. Exemple: EA-MAU dins d'una CAV apunta a la cova com a parent.
- ID_Group (-> T_GROUPS): agrupacio funcional independent. Exemple: 4 EA-MAU en alineament vertical comporten el mateix grup, independentment de si alguna esta dins d'una cova.

T_GROUPS inclou un camp ID_Group_Type (-> L_GROUP_TYPE amb 7 valors: Vertical alignment / Ledge cluster / Platform with brackets / Cave cluster / Circulation network / Rock art cluster / Functional group) que connecta directament amb les hipotesis H03 (saturacio espacial) i OE2 (relacions espacials).

## **6. Integracio espacial amb QGIS**

El flux de treball recomanat es el descrit a la v1.0 (Metashape -> CSV coordenades UTM 18S -> Access -> QRY_07_Export_QGIS -> QGIS capa de punts). La v2.0 actualitza QRY_07 per incloure Arch_Status, Material_Status i Interior_Area_m2 a l'exportacio.

*Nota: la nova QRY_07 usa LEFT JOIN per a L_STATUS i L_MATERIAL_STATUS (en lloc d'INNER JOIN) per tal que les estructures sense estat assignat no queden excloses de l'exportacio espacial, que es critica per a la visualitzacio completa al mapa.*

## **7. Consultes SQL actualitzades**

Les 10 consultes SQL s'han actualitzat per a incorporar els camps nous. Les modificacions mes significatives:

| **Consulta** | **Canvis principals en v2.0** |
|---|---|
| QRY_03_Conservation_by_Sector | Ara usa ID_Arch_Status + ID_Material_Status (via LEFT JOIN amb L_STATUS i L_MATERIAL_STATUS). Permet veure la combinacio dels dos estats per sector. |
| QRY_05_Export_RStats | Afegits 15 nous camps arquitectonics (Base_Level, Corbelled_Platform, Structural_Pilasters...) i Interior_Area_m2. Ara exporta un total de 55 variables per a R/SPSS. |
| QRY_06_Volumetry_by_Typology | Afegit AVG(Interior_Area_m2) AS Mean_Area juntament amb les estadistiques de volum. |
| QRY_07_Export_QGIS | Afegits Arch_Status, Material_Status, Interior_Area_m2 via LEFT JOIN. Mantε WHERE Coord_Lat_WGS84 IS NOT NULL. |

*Taula 4. Canvis principals a les consultes SQL en la versio 2.0.*

## **8. Formulari principal F_STRUCTURES (11 pestanyes)**

La versio 2.0 reestructura el formulari en 11 pestanyes tematiques, mes 2 subformularis vinculats:

| **Pestanya** | **Contingut principal** |
|---|---|
| 1.Ident. | Code, ID_Sector, ID_Typology, ID_Support, ID_Parent, ID_Group |
| 2.Arq. | N_Floors, Floor_Plan, N_Built_Walls, Length/Width/Height_m, Approx_Height_m, Lintel, Natural_Roof, Buttresses, Interlevel_Cornice, Wooden_Stakes |
| 3.Acab. | Plastered, Plaster_Color, Rock_Painting, Rock_Paint_Color |
| 4.Dec. | 14 camps de decoracio booleans + subformulari F_DECORATIONS (T_DECORATIONS per cos constructiu i tipus) |
| 5.Estat | ID_Arch_Status, ID_Material_Status, Looting, Fire_Damage, Animal_Activity, Modern_Access |
| 6.Bio. | Human_Remains, MNI, Anatomical_Connection, Mummification, Funerary_Bundles, Dispersed_Remains, Flexed_Position, Bone_Burning |
| 7.Mat. | Mat_Textiles, Mat_Wood, Mat_VegFiber, Mat_Ceramics, Mat_Fauna, Mat_DeerAntler, Mat_Other |
| 8.Cron. | C14, Chrono_Start/End_Cent, ID_Campaign + Interior_Area_m2, Interior_Vol_m3, Total_Vol_m3, ID_Vol_Method, Vol_Notes |
| 9.SIG | Coordenades WGS84 + UTM, Altitude_masl, ID_Coord_Method + URLs panorames/gigafotos/models 3D, ChaXR_Documented, Notes |
| 10.Systems | Sistemes arquitectonics: NIVELL 0 (I, J, K+L+N->M), OBERTURA (A+B+P->O), FACANA+SUPERIOR (C, Q, F, G, H) |
| 11.Extra | Embedded_Base_Beams, Tie_Walls, Corner_Quoins + subformulari F_ARCH_FEATURES (T_ARCH_FEATURES) |

*Taula 5. Estructura del formulari F_STRUCTURES amb 11 pestanyes tematiques.*

Tots els camps FK estan configurats com a ComboBox vinculats a les seues taules lookup per ControlSource (no per nom del control, la qual cosa el fa robust davant canvis de disseny). Els 12 combos configurats son: ID_Sector, ID_Typology, ID_Support, ID_Arch_Status, ID_Material_Status, ID_Vol_Method, ID_Coord_Method, ID_Campaign, ID_Group, ID_Parent, ID_Struct_Body (a F_DECORATIONS) i ID_Dec_Type (a F_DECORATIONS).

## **9. Scripts VBA consolidats**

La versio 2.0 consolida tots els scripts incremental executats durant el disseny en dos unics fitxers:

- chachapoya_01_DB.bas (629 linies) - Sub BuildDB(): crea les 18 taules, pobla els 12 lookups, estableix les 18 relacions i genera les 10 consultes. S'executa sobre una BD en blanc.
- chachapoya_02_Form.bas (497 linies) - Sub BuildForm(): crea els subformularis F_DECORATIONS i F_ARCH_FEATURES, el formulari principal F_STRUCTURES amb 11 pestanyes i configura tots els combos. S'executa despres de BuildDB().

Ambdos scripts segueixen les restriccions tecnicas de VBA/Access JET SQL: (a) cap continuacio de linia (& _), s'usa el patro sql = sql & '...' per a SQL llarga; (b) codificacio cp1252 (Windows-1252); (c) cap caracter no-ASCII en cadenes de codi; (d) paraules reservades en ACCESS entre claudators o evitades (Level -> Body_Level); (e) el patro tmpName = f.Name per a CreateControl sobre formularis nous; (f) DoCmd.Save acForm, tmpName + DoCmd.Rename FRM, acForm, tmpName per al reanomenament.

## **10. Limitacions i perspectives**

La limitacio principal es la manca de dades de camp per a moltes estructures de La Petaca Sector Sud. La taula T_ARCH_FEATURES aporta una solucio parcial a l'escalabilitat del schema: en lloc d'afegir nous camps booleans per a cada element constructiu nou que pugui aparèixer en campanyes futures, l'usuari pot registrar-los com a entrades flexibles sense modificar l'estructura.

La distincio Nivell 0 / Nivell 1 implementada a L_STRUCT_BODY permet abordar per primera vegada l'analisi de la distribucio de decoracio per posicio a la facana, relacionada directament amb H01 (ingenieria funeraria planificada) i H04 (sequencia operativa consistent).

De cara a publicacions futures, es planteja la migracio a GeoPackage (.gpkg) per a la integracio nativa amb QGIS, mantenint Access com a interficie d'entrada de dades.

## **11. References**

Epstein, L.; Toyne, J.M. (2016). When Space Is Limited: A Spatial Exploration of Pre-Hispanic Chachapoya Mortuary and Ritual Microlandscape. In Osterholtz, A.J. (ed.), Theoretical Approaches to Analysis and Interpretation of Commingled Human Remains. Springer, Switzerland, pp. 97-124.

Ribera-Torro, E. (2023). Chacha XR. Una experiencia immersiva de no-ficcio per l'arqueologia Chachapoya. Treball de Fi de Master, Master Universitari en Arts Visuals i Multimedia, Universitat Politecnica de Valencia.

Ribera-Torro, E.; Toyne, J.M.; Del Aguila, R.; Ribera, J.A.; Galexner, J.; Anzellini, A.; Pans, M. (2026). Extended Reality on Chachapoya Cliffside Necropolises: From Digital Documentation to Public Engagement. Open Archaeology, 12(1). DOI 10.1515/opar-2025-0071.

Toyne, J.M.; Anzellini, A. (2017). Sociedad, identidad y variedad en los mausoleos de La Petaca, Chachapoyas. Boletin de Arqueologia PUCP, 23, 231-257.

Toyne, J.M.; Anzellini, A.; Epstein Miculas, L.; Mejias Pitti, I.; Puig Castell, J.; Guinot Castello, S. (2018). Going Vertical: Using Vertical Progression Techniques to Explore a Cliff Necropolis in Late Precolumbian Chachapoyas, Peru. Advances in Archaeological Practice, 6. DOI 10.1017/aap.2018.31.
