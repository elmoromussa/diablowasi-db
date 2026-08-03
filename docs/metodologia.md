**Universitat d'Alacant**

Màster en Arqueologia Professional i Gestió Integral del Patrimoni

**DISSENY D'UNA BASE DE DADES ARQUEOLÒGICA**

per a l'estudi de les necròpolis de penya-segat de La Petaca i Diablo Wasi

*(Leymebamba, Amazonas, Perú, s. IX-XVI)*

Autor: Esteve Ribera Torró

Directors: Dr. Ignasi Grau Mira (UA) i Dra. J. Marla Toyne (UCF)

Curs acadèmic 2024-2025

*Versió 2 del document — actualitzada segons la BD v3 (agost 2026)*

# **Resum**

El present document descriu el disseny, la justificació i la implementació de la base de dades relacional destinada a la documentació sistemàtica i l'anàlisi estadística de les estructures funeràries de les necròpolis de penya-segat de La Petaca i Diablo Wasi (Leymebamba, Departament d'Amazonas, Perú, s. IX-XVI d.n.e.). La base de dades constitueix l'eix vertebrador de la metodologia del Treball de Fi de Màster, en tant que permet centralitzar les variables arqueològiques, establir relacions jeràrquiques entre elements, exportar dades per a l'anàlisi estadística i vincular el registre arqueològic amb el Sistema d'Informació Geogràfica (SIG) implementat en QGIS.

L'esquema segueix els principis de la tercera forma normal (3FN) i inclou 19 taules, una taula principal (T_STRUCTURES) amb 124 camps organitzats en 20 categories temàtiques, 20 relacions, 14 consultes SQL i un formulari d'entrada de dades amb 12 pestanyes i interfície en valencià. Un dels avanços metodològics centrals és la normalització d'un vocabulari arquitectònic bilingüe (valencià/anglés) de 24 elements (A-X), derivat de l'anàlisi fotogramètrica 3D de les estructures de Diablo Wasi, que permet el registre sistemàtic dels sistemes constructius i la seua anàlisi comparativa entre jaciments i sectors. La versió actual de la BD (v4) consolida l'esquema amb tres avanços metodològics: (a) el domini 0/1/9 per als elements A-X (0=Absent, 1=Present, 9=ND/no observable, per defecte 9), que separa la decisió constructiva del biaix tafonòmic i documental, condició de validesa de la coocurrència, el clúster jeràrquic i l'autocorrelació espacial; (b) la matriu A-X completa de 24 columnes (amb els elements L i M com a camps propis i Q derivada del material del dintell) exportada per QRY_13; i (c) la taula T_CONNECTIONS d'arestes físiques entre estructures (repisa compartida, plataforma contínua, mur o biga compartits), que habilita la reconstrucció formal de la xarxa de circulació aèria (OE3) via QRY_14.

# **1. Introducció i objectius**

L'estudi de les necròpolis de penya-segat Chachapoya planteja reptes documentals inèdits en l'arqueologia andina. La verticalitat dels jaciments, l'heterogeneïtat tipològica de les estructures i la multiplicitat de variables arquitectòniques, decoratives, bioarqueològiques i cronològiques exigeixen una eina de gestió de dades que vaja més enllà d'un full de càlcul pla. La base de dades relacional respon a aquesta necessitat, centralitzant en un sistema únic tota la informació necessària per a testar les hipòtesis de recerca del TFM.

La BD serveix com a infraestructura metodològica per a la consecució dels quatre objectius específics del TFM:

- **OE1.** Construir una infraestructura documental sistemàtica amb un vocabulari arquitectònic normalitzat, integrada amb el SIG.

- **OE2.** Caracteritzar la variabilitat tipològica i els sistemes constructius de les estructures funeràries.

- **OE3.** Analitzar la distribució espacial, la volumetria i les relacions funcionals entre estructures, incloent-hi la reconstrucció de la xarxa de circulació aèria.

- **OE4.** Interpretar la interacció entre el suport geomorfològic, les decisions constructives i les pràctiques funeràries.

El marc de recerca revisat articula sis hipòtesis (H01-H06). H01-H04 procedeixen del marc original, amb H03 reformulada com a argument morfològic-estructural (testable sense dependre de datacions absolutes); H05 i H06 són hipòtesis noves que completen el triangle natura-tècnica-símbol de l'objectiu general. La taula següent resumeix la seua operacionalització en camps de la BD:

| **Hip.** | **Descripció** | **Camps i taules principals que l'operacionalitzen** |
| --- | --- | --- |
| H01 | Enginyeria funerària planificada | N_Floors, Corbelled_Platform, Structural_Pilasters, Jambs, Recessed_Portal, Masonry_Quality, Masonry_Type, vector A-X (QRY_13), T_DECORATIONS |
| H02 | Geologia com a factor determinant | ID_Support, Support_Width_cm, Support_Depth_cm, Support_Morphology, Support_Modified, Natural_Roof, Altitude_masl, Height_Above_Base_m, Base_Level, Tie_Walls (QRY_12) |
| H03 | Saturació espacial acumulativa (argument morfològic-estructural) | Construction_Phases, Phase_Evidence, T_GROUPS, ID_Group, ID_Parent, QRY_07 (exportació QGIS) |
| H04 | Seqüència operativa consistent (chaîne opératoire) | ID_Typology, vector A-X (QRY_13), Construction_Phases, T_DECORATIONS, T_DATING (C14), Chrono_Start/End_Cent |
| H05 | Tradició constructiva compartida entre LP i DW | Vector A-X per estructura (QRY_13): coocurrència, clúster de Jaccard, anàlisi de correspondències; Masonry_Quality, Masonry_Type |
| H06 | Visibilitat i marcatge territorial | Facade_Orientation, Visibility_Valley, Coord_E/N_UTM, Altitude_masl (QRY_07 -> anàlisi de visibilitat en QGIS) |

*Taula 1. Hipòtesis del TFM i camps de la base de dades que les operacionalitzen (marc revisat, sis hipòtesis).*

# **2. Fonamentació metodològica**

## **2.1. Convenció d'idiomes**

La base de dades està completament en anglés (noms de taules, camps i **valors emmagatzemats**, incloent-hi les llistes de valors dels ComboBox i el contingut de les taules lookup). Aquesta decisió respon a tres factors: (a) la publicació derivada del projecte (Ribera-Torró et al. 2026, Open Archaeology) usa terminologia anglesa per a elements i camps; (b) la codirectora del TFM, la Dra. J. Marla Toyne (UCF), treballa en anglés; i (c) l'exportació a R/SPSS amb capçaleres i valors en anglés és l'estàndard en la publicació científica internacional i garanteix la reproduïbilitat de les anàlisis.

La interfície d'entrada de dades, en canvi, és en valencià: totes les etiquetes visibles del formulari (pestanyes, camps, capçaleres) es mostren en valencià, sense que això afecte mai els valors emmagatzemats. Aquesta separació estricta *interfície en valencià / dades en anglés* és una restricció de disseny fixa del projecte. Els documents del TFM (en valencià), les presentacions a congressos (en anglés o valencià) i les presentacions al Perú (en castellà) mantenen la seua llengua corresponent.

## **2.2. Marc disciplinar**

El disseny de les variables arqueològiques s'ha fonamentat en dues disciplines complementàries. L'Arqueologia de l'Arquitectura aporta el marc per a la caracterització constructiva (materials, tècniques, decoració) i la tipologia de les estructures. L'Arqueologia del Paisatge proporciona el marc per a l'anàlisi de la distribució espacial i les relacions amb la geologia i l'entorn natural. La integració de la BD amb QGIS permet implementar anàlisis de densitat, visibilitat i distribució vertical directament relacionades amb les hipòtesis H02, H03 i H06. H02 actua com a eix de connexió entre les dues disciplines: és alhora factor constructiu (eix arquitectònic) i element configurador del microentorn funerari (eix paisatgístic).

## **2.3. Procediment de disseny**

El disseny segueix un procés iteratiu en cinc etapes:

- Inventari de variables a partir de la documentació existent: anàlisi de les fitxes de camp, els treballs publicats (Toyne i Anzellini 2017; Epstein i Toyne 2016; Toyne et al. 2018), el TFM Chacha XR (Ribera-Torró 2023) i l'article derivat (Ribera-Torró et al. 2026).

- Modelatge entitat-relació (ER): identificació de les entitats principals i les relacions entre elles.

- Normalització fins a la 3FN: eliminació de la redundància i creació de taules de consulta (lookup) per a tots els camps categòrics.

- Validació amb les estructures documentades de Diablo Wasi Sector 1 i Sector 4, incloent-hi l'anàlisi fotogramètrica 3D dels models Metashape.

- Revisió contra el marc de recerca (v2/v3): avaluació de la cobertura de cada hipòtesi pels camps existents, que va conduir a la incorporació de deu camps nous en quatre grups temàtics i tres consultes noves (vegeu la secció 5.4).

# **3. Vocabulari arquitectònic normalitzat (A-X)**

Un dels avanços metodològics centrals de la base de dades és la normalització bilingüe (valencià/anglés) d'un vocabulari de 24 elements arquitectònics (A-X), derivat de l'anàlisi sistemàtica dels models fotogramètrics 3D de Diablo Wasi. Aquest vocabulari permet el registre no ambigu dels elements constructius observables en camp o en model 3D, i estableix la base per a l'anàlisi comparativa entre estructures, sectors i jaciments.

## **Criteris d'ordenació**

La codificació segueix tres criteris jerarquitzats: (1) NIVELL CONSTRUCTIU: de N0 (basament) cap a N1 (cos principal) i zona superior; (2) ALÇADA dins de cada nivell: de baix cap a dalt; (3) FUNCIÓ: primer l'element estructural portant, després el tractament decoratiu.

## **Sistemes compositius**

Tres elements del vocabulari són sistemes compositius, és a dir, el resultat de la combinació obligatòria d'altres elements:

- Sistema E+F+G -> H: Mènsules (E) + Bigues transversals (F) + Filades en voladís (G) generen la Plataforma volada d'accés (H). E i G són variants del mateix principi (mènsules de fusta vs filades de pedra).

- Sistema N+O+Q -> P: Llindar (N) + Brancals (O) + Dintell (Q) conformen l'Obertura d'accés (P). Els tres elements emmarquen el buit.

- Sistema S+T -> U: Biga de suport (S) + Superfície del ràfec (T, sempre pedra) generen el Ràfec-voladís (U). La fusta pot aparéixer a S (biga de suport), però T i U són sempre lapidis.

| **Ll.** | **Valencià** | **Anglés (BD)** | **Nivell** | **Sistema / Notes** |
| --- | --- | --- | --- | --- |
| A | Jàcenes basals empotrades | Embedded base beams | N0 | Standalone, dins la maçoneria del basament |
| B | Basament | Base level / podium | N0 | Standalone, cos constructiu basal |
| C | Sòcol decoratiu | Decorative socle | N0 | Standalone, tractament decoratiu de B |
| D | Muret transversal | Tie wall | N0 | Standalone, mur perpendicular d'ancoratge |
| E | Mènsules (fusta) | Timber corbels | N0 | Component de H (sistema E+F+G -> H) |
| F | Bigues transversals | Transverse beams | N0 | Component de H |
| G | Filades en voladís (pedra) | Corbelled masonry courses | N0 | Component de H (variant lítia de E) |
| **H** | **Plataforma volada d'accés** | Corbelled access platform | N0 | **SISTEMA: E+F+G -> H** |
| I | Cornisa intercòs | Interbody cornice | N0/N1 (interfície) | Standalone, element liminal entre cossos superposats |
| J | Cantoneres | Corner quoins | N1 | Standalone, pedres verticals als angles |
| K | Pilastres estructurals | Structural pilasters | N1 | Standalone, flanquegen la façana |
| L | Paraments laterals | Lateral wall faces | N1 | Standalone, superfícies del cos principal |
| M | Fris decoratiu en baix relleu | Bas-relief decorative frieze | N1 | Standalone, tractament de L (T_DECORATIONS) |
| N | Llindar | Sill / threshold | N1 | Component de P |
| O | Brancals | Jambs | N1 | Component de P |
| **P** | **Obertura d'accés** | Access opening | N1 | **SISTEMA: N+O+Q -> P** |
| Q | Dintell | Lintel | N1 | Component de P |
| R | Coronament | Upper crown / coping | N1 | Standalone, remat superior del cos |
| S | Biga de suport del ràfec | Eave-supporting beam | Sup. | Component de U |
| T | Superfície del ràfec | Eave surface | Sup. | Component de U (sempre pedra) |
| **U** | **Ràfec-voladís / Visera** | Eave / roof overhang | Sup. | **SISTEMA: S+T -> U (sempre pedra)** |
| V | Murs laterals | Lateral walls | N1 | Standalone, donen profunditat a la cambra |
| W | Mur posterior | Rear wall | N1 | Standalone (construït o roca natural) |
| X | Coberta de la cambra | Chamber roof | Sup. | Standalone, tancament superior construït |

*Taula 2. Vocabulari arquitectònic normalitzat A-X. Tres sistemes compositius (H, P, U) en negreta. La Superfície del ràfec (T) i el Ràfec-voladís (U) són sempre de pedra; la fusta pot aparéixer únicament a S. El Mur posterior (W) pot ser construït (Rear_Wall_Built = True) o el mateix faralló (False, en relació amb ID_Support).*

# **4. Arquitectura de la base de dades**

## **4.1. Taules i relacions (19 taules, 20 relacions)**

| **Cat.** | **Taula** | **Contingut i funció** |
| --- | --- | --- |
| Principal | T_STRUCTURES | Fitxa completa (124 camps). Font principal per a l'anàlisi estadística. |
| Principal | T_CONNECTIONS | NOVA v4. Arestes físiques entre estructures (repisa, plataforma, mur o biga compartits) per a l'anàlisi de xarxa de la circulació aèria (OE3). |
| Secundària | T_DATING | Datacions radiocarbòniques (C14). N:1 amb T_STRUCTURES. |
| Secundària | T_INDIVIDUALS | Dades individuals (edat, sexe, preservació). N:1. |
| Secundària | T_GROUPS | Agrupacions funcionals (alineaments, xarxes, conjunts). N:1. |
| Secundària | T_DECORATIONS | Decoració per estructura i cos constructiu. N:1. Clau per a H01 i H04. |
| Secundària | T_ARCH_FEATURES | Registre flexible d'elements no previstos a l'esquema. N:1. |
| Lookup | L_SITES | Jaciments: La Petaca, Diablo Wasi. Coordenades de referència. |
| Lookup | L_SECTORS | Sectors (LP: General, North, Central, Upper, South; DW: General + Sectors 1-6). |
| Lookup | L_TYPOLOGY | Tipologia (EA-MAU, EA-CAM, EA-PLA-R/V, NIX, CAV, PR, MEN, MIX, ND). |
| Lookup | L_SUPPORT | Suport geomorfològic (repisa natural àmplia/estreta, artificial, cavitat...). |
| Lookup | L_STATUS | Estat de conservació de l'arquitectura (Good/Fair/Pre-collapse/Collapsed/ND). |
| Lookup | L_MATERIAL_STATUS | Grau de preservació dels vestigis mobles (Good/Fair/Poor/Absent/ND). |
| Lookup | L_VOL_METHOD | Mètode de càlcul volumètric (L x W x H / Fotogramètric / Estimació / ND). |
| Lookup | L_COORD_METHOD | Mètode d'obtenció de coordenades (Drone RTK, Fotogrametria, GPS mòbil...). |
| Lookup | L_GROUP_TYPE | Tipus d'agrupació funcional (Vertical alignment, Circulation network...). |
| Lookup | L_CAMPAIGN | Campanyes: 2013 PALP I, 2016 PALP II, 2021 La Petaca Project, 2023 PALP IV. |
| Lookup | L_STRUCT_BODY | Posició de la decoració a la façana (N0-SOC, N1-JAM, N1-OVL, N1-COR...). |
| Lookup | L_DEC_TYPE | Tipus de decoració (T-shaped niche, Triangular motif, Painted band...). |

*Taula 3. Les 19 taules de la base de dades amb la seua funció.*

# **5. Variables arqueològiques de T_STRUCTURES (124 camps)**

La taula T_STRUCTURES concentra 124 camps en 20 categories temàtiques. Els camps de presència del vocabulari A-X són de tipus BYTE amb domini 0=Absent / 1=Present / 9=ND (per defecte 9: l'absència ha de registrar-se positivament, mai per omissió), de manera que la matriu analítica distingeix la decisió constructiva del biaix de conservació o de documentació. La taula següent resumeix les categories i el nombre de camps per categoria:

| **Categoria** | **Camps principals** | **N** |
| --- | --- | --- |
| Identificació | Code, ID_Sector, ID_Typology, ID_Support, ID_Parent, ID_Group | 7 |
| Morfologia i dimensions | N_Floors, Floor_Plan, N_Built_Walls, Length/Width/Height_m, Height_Above_Base_m, Dim_Method, Opening_Width/Height_cm, Access_Orientation, Lintel (Q), Natural_Roof, Buttresses, Wooden_Stakes | 15 |
| Detall del suport geològic (NOU v2) | Support_Width_cm, Support_Depth_cm, Support_Morphology, Support_Modified | 4 |
| Qualitat de la maçoneria (NOU v2) | Masonry_Quality, Masonry_Type | 2 |
| Sistemes N0 (vocab. A-H) | Base_Level (B), Decorative_Socle (C), Tie_Walls (D), Timber_Brackets (E), Timber_Bracket_Count, Transverse_Beams (F), Corbelled_Courses (G), Corbelled_Platform (H), Embedded_Base_Beams (A), Corbel_Material | 10 |
| Interfície N0/N1 (vocab. I) | Interbody_Cornice (I), Interbody_Cornice_Material | 2 |
| Sistemes N1 (vocab. J-W) | Corner_Quoins (J), Structural_Pilasters (K), Lateral_Wall_Faces (L), Relief_Frieze (M), Sill (N), Jambs (O), Access_Opening (P), Recessed_Portal, Upper_Crown (R), Lateral_Walls (V), Rear_Wall (W), Rear_Wall_Built | 12 |
| Maçoneria i morter | Masonry_Quality, Masonry_Type, Mortar_Present, Mortar_Type, Chinking_Stones, Mortar_Notes | 6 |
| Zona superior (vocab. J, R-U, X) | Eave (U), Eave_Beam (S), Eave_Surface (T), Upper_Crown (R), Chamber_Roof (X), Corner_Quoins (J) | 6 |
| Acabats superficials | Plastered, Plaster_Color, Rock_Painting, Rock_Paint_Color | 4 |
| Paisatge i orientació (NOU v2) | Facade_Orientation, Visibility_Valley | 2 |
| Decoració (booleans) | Dec_Square_Niche, Dec_Relief_T/T_Inv/L/L_Inv, Dec_Zigzag, Dec_Stepped, Rock_Art, RA_Anthropomorphic/Zoomorphic/Geometric/Abstract/Decap_Scene (el fris és l'element M: Relief_Frieze) | 13 |
| Estat de conservació | ID_Arch_Status, ID_Material_Status, Looting, Fire_Damage, Animal_Activity, Modern_Access | 6 |
| Bioarqueologia | Human_Remains, MNI, Anatomical_Connection, Mummification, Funerary_Bundles, Dispersed_Remains, Flexed_Position, Bone_Burning | 8 |
| Materials culturals | Mat_Textiles, Mat_Wood, Mat_VegFiber, Mat_Ceramics, Mat_Fauna, Mat_DeerAntler, Mat_Other | 7 |
| Cronologia | C14, Chrono_Start_Cent, Chrono_End_Cent | 3 |
| Fases constructives (NOU v2) | Construction_Phases, Phase_Evidence | 2 |
| Volumetria i àrea | Interior_Area_m2, Interior_Vol_m3, Total_Vol_m3, ID_Vol_Method, Vol_Notes | 5 |
| Coordenades espacials | Coord_Lat/Lon_WGS84, Coord_E/N_UTM, Altitude_masl, Coord_Precision_m, ID_Coord_Method | 7 |
| Documentació digital | URL_Pano, URL_Pano_2, URL_Giga, URL_3D, ChaXR_Documented, ID_Campaign, Notes | 7 |

*Taula 4. Categories temàtiques de T_STRUCTURES amb nombre de camps per categoria (total: 124).*

## **5.1. Estat de conservació dual**

La BD distingeix entre dues dimensions de l'estat de conservació: ID_Arch_Status (-> L_STATUS) mesura el grau de preservació de l'arquitectura construïda (escala: Good/Fair/Pre-collapse/Collapsed/ND), mentre que ID_Material_Status (-> L_MATERIAL_STATUS) mesura el grau de preservació dels vestigis mobles (escala: Good/Fair/Poor/Absent/ND). La causa del deteriorament queda registrada pels camps booleans Looting, Fire_Damage, Animal_Activity i Modern_Access, que són independents dels graus de conservació. L'encreuament entre grau i causa permet reconstruir la història postdeposicional de cada estructura.

## **5.2. T_DECORATIONS: decoració per cos constructiu**

La taula T_DECORATIONS permet el registre detallat de decoracions per estructura, complementant els camps booleans de T_STRUCTURES. Cada registre inclou: ID_Struct_Body (-> L_STRUCT_BODY: N0-SOC, N1-JAM, N1-OVL, N1-COR, N2-SPA, N2-JAM, ND), ID_Dec_Type (-> L_DEC_TYPE), Body_No (número de cos constructiu), Color (Red/White/Both/None/ND) i Notes. Aquesta estructura permet analitzar si, per exemple, els motius en T es concentren als brancals (N1-JAM) i el fris triangular sobre el dintell (N1-OVL), com s'observa als models 3D de DW Sector 1.

Els camps booleans de decoració a T_STRUCTURES (Dec_Square_Niche, Dec_Relief_T...) es mantenen per al filtratge ràpid en consultes estadístiques (khi-quadrat). T_DECORATIONS s'usa per al registre posicional i tipològic detallat.

## **5.3. T_ARCH_FEATURES: registre flexible**

La taula T_ARCH_FEATURES permet registrar qualsevol element constructiu no previst a l'esquema principal sense modificar l'estructura de la BD (Feature_Code, Present, Feature_Count, Material, Notes). Opera com a xarxa de seguretat per a elements nous que puguen aparéixer en campanyes futures, evitant la proliferació indefinida de nous camps booleans a T_STRUCTURES.

## **5.4. Camps incorporats en la revisió v2: justificació**

La revisió del marc de recerca (sis hipòtesis, quatre objectius) va evidenciar buits de cobertura en l'esquema original, resolts amb deu camps nous en quatre grups:

- **Detall del suport geològic (4 camps, H02):** Support_Width_cm, Support_Depth_cm, Support_Morphology i Support_Modified quantifiquen la relació entre les dimensions i la morfologia del suport natural i les decisions constructives. Sense aquestes mesures, H02 només es podia testar amb la categoria genèrica d'ID_Support.

- **Qualitat de la maçoneria (2 camps, H01/H04):** Masonry_Quality (Good/Moderate/Poor/ND) i Masonry_Type (Well-coursed/Irregular-coursed/Uncoursed/Mixed/ND) capturen la inversió tècnica diferencial, seguint la línia de variabilitat i identitat social de Toyne i Anzellini (2017).

- **Paisatge i orientació (2 camps, H06 i preparació QGIS):** Facade_Orientation i Visibility_Valley registren observacions de camp que alimenten l'anàlisi formal de visibilitat i orientació en QGIS.

- **Fases constructives (2 camps, H03/H04):** Construction_Phases i Phase_Evidence (C14/Stratigraphy/Superposition/Mortar/ND) permeten sostenir l'argument morfològic-estructural de la saturació acumulativa sense dependre de datacions absolutes.

# **6. Estratègia de jerarquia i agrupació**

La BD implementa dos mecanismes complementaris i independents per a gestionar les relacions entre elements:

- ID_Parent (autoreferenciant T_STRUCTURES.ID): contenció física o estructural. Exemple: un EA-MAU construït dins d'una CAV apunta a la cova com a pare. Permet jerarquies de N nivells.

- ID_Group (-> T_GROUPS): agrupació funcional. Exemple: 4 EA-MAU en alineament vertical pertanyen al mateix T_GROUPS sense que cap continga l'altre. T_GROUPS classifica l'agrupació per ID_Group_Type (L_GROUP_TYPE: Vertical alignment, Ledge cluster, Platform with brackets, Cave cluster, Circulation network, Rock art cluster, Functional group).

*Regla pràctica: usar ID_Parent quan un element no té sentit arqueològic independent del seu contenidor (un EA-MAU dins una CAV). Usar ID_Group quan els elements són arqueològicament independents però formen una unitat d**'**estudi (alineament de mènsules que reconstrueixen una xarxa de circulació perduda).*

# **7. Integració espacial amb QGIS**

Els camps Coord_Lat_WGS84 / Coord_Lon_WGS84 permeten carregar T_STRUCTURES com a capa de punts a QGIS (Layer -> Add Delimited Text Layer). Les coordenades individuals s'obtenen preferentment dels models Metashape georeferenciats (File -> Export -> Export Markers en format CSV, coordenades UTM zona 18S / WGS84, EPSG:32718). La consulta QRY_07_Export_QGIS exporta els camps més rellevants per a la visualització cartogràfica: tipologia, estat de conservació dual, volumetria, cronologia, MNI, documentació digital i, des de la v2, l'orientació de façana, la visibilitat des de la vall i la qualitat de la maçoneria, que preparen l'anàlisi de H06.

La naturalesa vertical dels farallons requereix treballar en dos nivells: l'ortofoto vertical de cada sector (exportada des de Metashape) per a les relacions relatives entre estructures, i les coordenades absolutes (UTM + altitud) per a la distribució espacial general del jaciment.

# **8. Anàlisi volumètrica i areal**

La BD registra tres mesures de volumetria: Interior_Area_m2 (àrea de sòl interior, m2), Interior_Vol_m3 (volum interior útil, m3) i Total_Vol_m3 (volum total incloent-hi parets). Per a estructures de geometria regular (EA-MAU), el volum interior es calcula com (Length_m - 2e) x (Width_m - 2e) x Height_m, on e és el gruix de parets (~0.20-0.30 m documentat a DW). Per a cavitats irregulars (EA-CAM, CAV), el volum s'extrau directament del model fotogramètric 3D via MeshLab o CloudCompare. El camp ID_Vol_Method registra el procediment per a ponderar la comparabilitat entre mesures.

La correlació entre Interior_Area_m2 / Interior_Vol_m3 i MNI permet avaluar si les tombes de major capacitat allotjaven més individus (H01), mentre que la variació del volum per tipologia i jaciment (QRY_06) caracteritza les estratègies d'adaptació al suport geomorfològic (H02).

# **9. Consultes SQL per a l'anàlisi estadística**

La BD inclou 14 consultes SQL predissenyades que cobreixen les principals anàlisis del TFM:

| **Consulta** | **Funció analítica** | **Hipòtesis** |
| --- | --- | --- |
| QRY_01_Typology_by_Site | Distribució de tipologies per jaciment i sector. | H01, H04 |
| QRY_02_Decoration_by_Site | Presència/absència de cada motiu decoratiu per jaciment (font per a khi-quadrat). | H01, H04 |
| QRY_03_Conservation_by_Sector | Distribució de l'estat de conservació arquitectura + vestigis mobles per sector. | General |
| QRY_04_C14_Structures | Estructures amb datació C14 ordenades cronològicament. | H04 |
| QRY_05_Export_RStats | Exportació plana completa per a R/SPSS. Actualitzada v2 amb els 10 camps nous. | Totes |
| QRY_06_Volumetry_by_Typology | Volumetria i àrea interiors per tipologia i jaciment (mitjana, mín., màx.). | OE2, H01 |
| QRY_07_Export_QGIS | Exportació espacial per a QGIS (coordenades + atributs clau). Actualitzada v2. | OE3, H02, H03, H06 |
| QRY_08_Children_of_Parent | Elements fills d'un element pare (contingut d'una cova, per exemple). | H01, H02 |
| QRY_09_Group_Members | Membres d'un conjunt funcional ordenats per altitud. | H03, OE3 |
| QRY_10_ChaXR_Coverage | Cobertura Chacha XR vs total per jaciment i campanya. | Metodologia |
| QRY_11_Masonry_by_Site | NOVA v2. Qualitat i tipus de maçoneria per jaciment i tipologia. | H01, H04 |
| QRY_12_Geology_Construction | NOVA v2. Encreuament suport geològic - tipologia amb mitjanes dimensionals. | H02 |
| QRY_13_AX_Pattern_Export | Reescrita v4. Matriu A-X completa de 24 columnes (AX_A..AX_X, valors 0/1/9; AX_Q derivada de Lintel) per a l'anàlisi de patrons en R. | H01, H04, H05 |
| QRY_14_Connections_Edges | NOVA v4. Llista d'arestes de T_CONNECTIONS amb coordenades UTM dels dos extrems, per a igraph (R) o línies en QGIS. | OE3, H03 |

*Taula 5. Les 14 consultes SQL predissenyades i la seua vinculació amb els objectius i hipòtesis del TFM.*

## **9.1. Estratègia d'anàlisi de patrons constructius (QRY_13)**

L'atomització de les estructures en el vocabulari A-X converteix cada estructura en un vector de 24 posicions amb valors 0/1/9 (absència, presència, no observable), complementat amb Masonry_Quality, Masonry_Type, la morfologia del suport i les coordenades UTM. La consulta QRY_13_AX_Pattern_Export exporta aquest vector de forma neta (LEFT JOIN amb L_TYPOLOGY per a no excloure registres parcials), i habilita quatre famílies d'anàlisi en R:

- **Coocurrència d'elements:** matrius de coeficients phi i de Jaccard entre parells d'elements A-X, que revelen quins elements tendeixen a aparéixer junts (evidència de sistemes i de seqüències operatives consistents: H01, H04).

- **Clúster jeràrquic:** dendrograma sobre la matriu de distàncies de Jaccard entre estructures, per a identificar famílies constructives dins i entre jaciments (H05: tradició compartida).

- **Anàlisi de correspondències:** projecció d'elements i estructures en un espai reduït, amb LP i DW com a punts suplementaris, per a visualitzar la proximitat entre repertoris constructius.

- **Autocorrelació espacial (I de Moran):** sobre les coordenades UTM, per a determinar si els patrons constructius tenen estructura geogràfica (concentracions sectorials de combinacions específiques) o es distribueixen de manera homogènia (H05, H03).

Aquesta estratègia analítica és el fil conductor que connecta la caracterització constructiva (OE2), la planificació (H01), la seqüència operativa (H04) i la tradició compartida (H05), i dota de contingut empíric el debat sobre transmissió cultural i comunitats de pràctica.

# **10. Formulari principal F_STRUCTURES**

El formulari F_STRUCTURES (chachapoya_Form_v3_val.bas) centralitza l'entrada de dades amb 12 pestanyes temàtiques, dos subformularis vinculats i tots els camps FK configurats com a ComboBox. Els camps A-X es mostren com a combos de dues columnes (valor 0/1/9 emmagatzemat; etiqueta Absent/Present/ND visible). Tota la interfície és en valencià; tots els valors emmagatzemats són en anglés (secció 2.1).

| **Pestanya** | **Contingut** |
| --- | --- |
| 1. Id. | Code, ID_Sector, ID_Typology, ID_Support, ID_Parent, ID_Group + detall del suport geològic (Support_Morphology, Support_Modified, Support_Width_cm, Support_Depth_cm) |
| 2. Arq. | Morfologia general: N_Floors, Floor_Plan, N_Built_Walls, Height_Above_Base_m, Access_Orientation, Lintel (Q), Natural_Roof, Buttresses, Wooden_Stakes + Facade_Orientation, Visibility_Valley + Masonry_Quality, Masonry_Type, bloc de morter + Construction_Phases, Phase_Evidence |
| 3. Acab. | Acabats: Plastered, Plaster_Color, Rock_Painting, Rock_Paint_Color |
| 4. Dec. | 14 camps booleans de decoració + subformulari F_DECORATIONS (T_DECORATIONS per cos i tipus) |
| 5. Estat | Conservació: ID_Arch_Status, ID_Material_Status, Looting, Fire_Damage, Animal_Activity, Modern_Access |
| 6. Bio. | Bioarqueologia: Human_Remains, MNI, Anatomical_Connection, Mummification, Funerary_Bundles, Dispersed_Remains, Flexed_Position, Bone_Burning |
| 7. Mat. | Materials: Mat_Textiles, Mat_Wood, Mat_VegFiber, Mat_Ceramics, Mat_Fauna, Mat_DeerAntler, Mat_Other |
| 8. Cron. | Cronologia + Volumetria: C14, Chrono_Start/End_Cent, ID_Campaign, Interior_Area_m2, Interior_Vol_m3, Total_Vol_m3, ID_Vol_Method, Vol_Notes |
| 9. SIG | Coordenades WGS84 + UTM, Altitude_masl, ID_Coord_Method, URL (panorames/gigafotos/models 3D), ChaXR_Documented, Notes |
| 10. Sist. | Sistemes A-X: NIVELL 0 (B, C, D, E+F+G->H), OBERTURA (N+O+Q->P), FAÇANA/SUPERIOR (I, J, K, R, S+T->U, V, W, X) |
| 11. Extra | Elements confirmats extres (A: jàcenes basals, V: murs laterals, W: mur posterior) + subformulari F_ARCH_FEATURES (T_ARCH_FEATURES) |

*Taula 6. Estructura del formulari F_STRUCTURES amb 12 pestanyes temàtiques: 1.Id., 2.Arq., 3.Acab., 4.Dec., 5.Estat, 6.Bio., 7.Mat., 8.Cron., 9.Metr., 10.Doc., 11.Sist. (24 elements A-X) i 12.Extra.*

## **10.1. Captions DAO per a la vista full de dades**

La vista full de dades dels subformularis mostra per defecte els noms de camp de la TableDef (en anglés), no les etiquetes del formulari. La rutina SetFieldCaptionsVal resol aquest comportament establint la propietat Caption de cada camp a nivell de TableDef via DAO (amb CreateProperty si la propietat no existeix prèviament). D'aquesta manera, les capçaleres de columna es mostren en valencià sense alterar els noms de camp reals ni els valors emmagatzemats.

# **11. Implementació tècnica**

## **11.1. Scripts VBA**

La BD s'implementa en Microsoft Access via dos scripts VBA consolidats:

- chachapoya_DB_v4.bas - Sub BuildDB(): crea les 19 taules en ordre de dependència, fixa el valor per defecte 9 dels 27 camps BYTE via DAO, pobla els 12 lookups, estableix les 20 relacions i genera les 14 consultes. S'executa sobre una BD en blanc. És idempotent: elimina i regenera les consultes existents i comprova l'existència de taules abans de crear-les.

- chachapoya_Form_v3_val.bas - Sub BuildForm(): crea els subformularis F_DECORATIONS i F_ARCH_FEATURES amb etiquetes adjuntes als controls (capçaleres de columna en valencià en vista full de dades), construeix F_STRUCTURES amb 12 pestanyes (absorbeix l'antic patch mètric, ara obsolet), configura tots els ComboBox per ControlSource i estableix les captions DAO de reforç. S'executa després de BuildDB().

## **11.2. Restriccions tècniques de VBA/Access**

Els scripts apliquen les següents restriccions tècniques del motor JET SQL i del VBA d'Access:

- Cap continuació de línia (& _): tot el SQL llarg usa el patró sql = sql & "..." en línies separades.

- Codificació cp1252 (Windows-1252) sense caràcters no-ASCII en cadenes de codi.

- Paraules reservades evitades: Level -> Body_Level (LEVEL és paraula reservada en JET SQL).

- Patró tmpName = f.Name per a CreateControl sobre formularis nous, seguit de DoCmd.Save acForm, tmpName i DoCmd.Rename FRM, acForm, tmpName.

- Per a N taules en JOIN calen N-1 parèntesis d'obertura al FROM en JET SQL.

- La relació autoreferenciant T_STRUCTURES.ID -> T_STRUCTURES.ID_Parent usa el flag dbRelationDontEnforceIntegrity (valor numèric 2) per a evitar conflictes en cascada.

- Tots els procediments auxiliars són Private Sub; només BuildDB i BuildForm són públics.

- DROP TABLE / QueryDefs.Delete abans de la creació per a objectes que poden existir de sessions anteriors.

- Els ComboBox es configuren per ControlSource (no pel nom del control).

## **11.3. Exportació per a l'anàlisi estadística**

L'exportació per a R o SPSS es realitza via QRY_05_Export_RStats (exportació plana completa) i QRY_13_AX_Pattern_Export (vector A-X per a l'anàlisi de patrons), des de Dades externes -> Exportar -> Text/Excel. Tots els camps categòrics de les consultes estan en format llegible (no com a ID numèrics) gràcies als JOIN amb les taules lookup. En R, els paquets readxl o read.csv permeten la càrrega directa.

# **12. Limitacions i perspectives**

La limitació principal és la manca de dades de camp per a moltes de les estructures de La Petaca Sector Sud (>=96 estructures identificades, però no totes amb fitxa completa). El camp ID_Campaign permet identificar quines estructures han estat documentades i en quina campanya, facilitant la detecció de biaixos documentals. La consulta QRY_13 usa LEFT JOIN precisament per a no excloure de l'anàlisi de patrons els registres parcials sense tipologia assignada.

La taula T_ARCH_FEATURES aporta una solució al problema de l'escalabilitat de l'esquema: els elements constructius nous que puguen aparéixer en campanyes futures es poden registrar com a entrades flexibles sense modificar l'estructura de la BD. De la mateixa manera, els tipus de decoració nous s'afigen directament a L_DEC_TYPE sense reconstruir la base.

El vocabulari arquitectònic A-X és el primer intent de normalització terminològica per a les estructures funeràries Chachapoya en la bibliografia. La seua aplicació a La Petaca i la validació creuada amb la literatura existent (Toyne i Anzellini 2017; Epstein i Toyne 2016; Toyne et al. 2018) representa una de les aportacions originals del present TFM.

De cara a publicacions futures, es planteja la migració a GeoPackage (.gpkg) per a la integració nativa amb QGIS, mantenint Access com a interfície d'entrada de dades amb sincronitzacions periòdiques.

# **13. Referències**

Epstein, L.; Toyne, J.M. (2016). When Space Is Limited: A Spatial Exploration of Pre-Hispanic Chachapoya Mortuary and Ritual Microlandscape. A: Osterholtz, A.J. (ed.), Theoretical Approaches to Analysis and Interpretation of Commingled Human Remains. Springer, Switzerland, pp. 97-124.

Ribera-Torró, E. (2023). Chacha XR. Una experiència immersiva de no-ficció per a l'arqueologia Chachapoya. Treball de Fi de Màster, Màster Universitari en Arts Visuals i Multimèdia, Universitat Politècnica de València.

Ribera-Torró, E.; Toyne, J.M.; Del Aguila, R.; Ribera, J.A.; Galexner, J.; Anzellini, A.; Pans, M. (2026). Extended Reality on Chachapoya Cliffside Necropolises: From Digital Documentation to Public Engagement. Open Archaeology, 12(1). DOI 10.1515/opar-2025-0071.

Toyne, J.M.; Anzellini, A. (2017). Sociedad, identidad y variedad en los mausoleos de La Petaca, Chachapoyas. Boletín de Arqueología PUCP, 23, 231-257.

Toyne, J.M.; Anzellini, A.; Epstein Miculas, L.; Mejías Pitti, I.; Puig Castell, J.; Guinot Castelló, S. (2018). Going Vertical: Using Vertical Progression Techniques to Explore a Cliff Necropolis in Late Precolumbian Chachapoyas, Peru. Advances in Archaeological Practice, 6. DOI 10.1017/aap.2018.31.