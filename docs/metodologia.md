**Universitat d'Alacant**

Màster en Arqueologia Professional i Gestió Integral del Patrimoni

**DISSENY D'UNA BASE DE DADES ARQUEOLÒGICA**

per a l'estudi de les necròpolis de penya-segat de La Petaca i Diablo Wasi

*(Leymebamba, Amazonas, Perú, s. IX-XVI)*

Autor: Esteve Ribera Torró

Directors: Dr. Ignasi Grau Mira (UA) i Dra. J. Marla Toyne (UCF)

Curs acadèmic 2024-2025

*Versió 3 del document — actualitzada segons la BD v12 (agost 2026). Consolida la iteració v10→v11 (delta rev. 5) i l'addendum v12.*

# **Resum**

El present document descriu el disseny, la justificació i la implementació de la base de dades relacional destinada a la documentació sistemàtica i l'anàlisi estadística de les estructures funeràries de les necròpolis de penya-segat de La Petaca i Diablo Wasi (Leymebamba, Departament d'Amazonas, Perú, s. IX-XVI d.n.e.). La base de dades constitueix l'eix vertebrador de la metodologia del Treball de Fi de Màster, en tant que permet centralitzar les variables arqueològiques, establir relacions jeràrquiques entre elements, exportar dades per a l'anàlisi estadística i vincular el registre arqueològic amb el Sistema d'Informació Geogràfica (SIG) implementat en QGIS.

L'esquema segueix els principis de la tercera forma normal (3FN) i inclou 22 taules, una taula principal (T_STRUCTURES) amb 120 camps, 25 relacions, la família de 26 consultes SQL i un formulari d'entrada de dades amb 12 pestanyes, quatre subformularis i interfície en valencià. Un dels avanços metodològics centrals és la normalització d'un vocabulari arquitectònic bilingüe (valencià/anglés) de 24 entrades (A-X), derivat de l'anàlisi fotogramètrica 3D de les estructures de Diablo Wasi, formalitzat des de la v11 com a taula lookup pròpia (L_ELEMENTS).

La versió actual consolida l'esquema al voltant de dos principis metodològics. El primer, ja establit en versions anteriors: **el registre ha de distingir sempre l'absència verificada de la manca d'observació**. El segon, incorporat amb la iteració v10→v11: **el registre ha de distingir el que es va construir del que sobreviu**. Els 20 camps d'element del vocabulari A-X adopten per això un domini de cinc valors (0 = Absent / 1 = Present complet / 2 = Present parcial / 3 = Desaparegut atestat / 9 = No observable), on el valor 3 —l'únic del domini que és una inferència i no una observació— exigeix evidència material registrada a la taula nova T_LOST_ELEMENTS. Els tres antics «elements» H, P i U es reconeixen com el que sempre van ser, sistemes compositius, i esdevenen camps d'estat propis (Sys_Platform, Sys_Portal, Sys_Eave) amb un domini de sis valors que separa «no aplicable» d'«absent verificat». La classe de registre (Record_Class, a L_TYPOLOGY) filtra tota anàlisi: dels 35 registres actuals, cinc no són estructures, i qualsevol percentatge calculat sobre 35 seria erroni.

L'addendum v12 tanca el cercle en el punt d'entrada: el camp Dec_Present restaura el judici agregat 0/1/9 sobre la decoració que l'eliminació dels booleans havia deixat orfe; la bateria de validació QRY_16 passa de 21 a 26 regles; i el formulari incorpora un gating condicional de tres nivells amb emplenat ràpid confirmat, de manera que la majoria de les incoherències que la bateria detectava a posteriori esdevenen impossibles d'introduir. El principi de disseny resultant: el formulari preveu en el moment d'entrada; la bateria detecta a escala de corpus.

# **1. Introducció i objectius**

L'estudi de les necròpolis de penya-segat Chachapoya planteja reptes documentals inèdits en l'arqueologia andina. La verticalitat dels jaciments, l'heterogeneïtat tipològica de les estructures i la multiplicitat de variables arquitectòniques, decoratives, bioarqueològiques i cronològiques exigeixen una eina de gestió de dades que vaja més enllà d'un full de càlcul pla. La base de dades relacional respon a aquesta necessitat, centralitzant en un sistema únic tota la informació necessària per a testar les hipòtesis de recerca del TFM.

La BD serveix com a infraestructura metodològica per a la consecució dels quatre objectius específics del TFM:

- **OE1.** Construir una infraestructura documental sistemàtica amb un vocabulari arquitectònic normalitzat, integrada amb el SIG.

- **OE2.** Caracteritzar la variabilitat tipològica i els sistemes constructius de les estructures funeràries.

- **OE3.** Analitzar la distribució espacial, la volumetria i les relacions funcionals entre estructures, incloent-hi la reconstrucció de la xarxa de circulació aèria.

- **OE4.** Interpretar la interacció entre el suport geomorfològic, les decisions constructives i les pràctiques funeràries.

El marc de recerca revisat articula sis hipòtesis (H01-H06). H01-H04 procedeixen del marc original, amb H03 reformulada com a argument morfològic-estructural (testable sense dependre de datacions absolutes); H05 i H06 completen el triangle natura-tècnica-símbol de l'objectiu general. La taula següent resumeix la seua operacionalització en camps de la BD (noms actualitzats a v12):

| **Hip.** | **Descripció** | **Camps i taules principals que l'operacionalitzen** |
| --- | --- | --- |
| H01 | Enginyeria funerària planificada | N_Basal_Bodies, N_Chamber_Bodies, Sys_Platform, Sys_Portal, Structural_Pilasters, Jambs, Recessed_Frame, Masonry_Quality, Masonry_Type, matriu AX+SYS (QRY_13), T_DECORATIONS |
| H02 | Geologia com a factor determinant | ID_Support, ID_Support_Secondary, Support_Width_cm, Support_Depth_cm, Support_Modified, Altitude_masl, Height_Above_Base_m, Base_Level, Tie_Walls, Rear_Closure_Type, Chamber_Roof_Type (QRY_12) |
| H03 | Saturació espacial acumulativa (argument morfològic-estructural) | Construction_Phases, Phase_Evidence, T_CONNECTIONS.Chrono_Relation (graf dirigit, QRY_14), T_GROUPS, ID_Group, ID_Parent, QRY_07 (exportació QGIS) |
| H04 | Seqüència operativa consistent (chaîne opératoire) | ID_Typology + Record_Class, matriu AX+SYS (QRY_13), T_LOST_ELEMENTS (el gest constructiu del valor 3), Construction_Phases, T_DECORATIONS, T_DATING (C14), Chrono_Start/End_Cent |
| H05 | Tradició constructiva compartida entre LP i DW | Matriu AX+SYS per estructura (QRY_13): coocurrència, clúster de Jaccard, anàlisi de correspondències; Masonry_Quality, Masonry_Type |
| H06 | Visibilitat i marcatge territorial | Facade_Orientation, Visibility_Valley, Pigment_Extent (valor perimetral), Coord_E/N_UTM, Altitude_masl (QRY_07 → anàlisi de visibilitat en QGIS) |

*Taula 1. Hipòtesis del TFM i camps de la base de dades que les operacionalitzen (marc revisat, sis hipòtesis; noms de camp v12).*

# **2. Fonamentació metodològica**

## **2.1. Convenció d'idiomes**

La base de dades està completament en anglés (noms de taules, camps i **valors emmagatzemats**, incloent-hi les llistes de valors dels ComboBox i el contingut de les taules lookup). Aquesta decisió respon a tres factors: (a) la publicació derivada del projecte (Ribera-Torró et al. 2026, Open Archaeology) usa terminologia anglesa per a elements i camps; (b) la codirectora del TFM, la Dra. J. Marla Toyne (UCF), treballa en anglés; i (c) l'exportació a R/SPSS amb capçaleres i valors en anglés és l'estàndard en la publicació científica internacional i garanteix la reproduïbilitat de les anàlisis.

La interfície d'entrada de dades, en canvi, és en valencià: totes les etiquetes visibles del formulari es mostren en valencià, sense que això afecte mai els valors emmagatzemats. Des de la v11 aquesta separació és tècnicament infrangible: tots els combos de domini són de dues columnes, amb la columna emmagatzemada (anglés) oculta i l'etiqueta (valencià) visible, de manera que reanomenar una etiqueta no pot tocar mai una dada. Els documents del TFM (en valencià), les presentacions a congressos (en anglés o valencià) i les presentacions al Perú (en castellà) mantenen la seua llengua corresponent.

## **2.2. Marc disciplinar**

El disseny de les variables arqueològiques s'ha fonamentat en dues disciplines complementàries. L'Arqueologia de l'Arquitectura aporta el marc per a la caracterització constructiva (materials, tècniques, decoració) i la tipologia de les estructures. L'Arqueologia del Paisatge proporciona el marc per a l'anàlisi de la distribució espacial i les relacions amb la geologia i l'entorn natural. La integració de la BD amb QGIS permet implementar anàlisis de densitat, visibilitat i distribució vertical directament relacionades amb les hipòtesis H02, H03 i H06. H02 actua com a eix de connexió entre les dues disciplines: és alhora factor constructiu (eix arquitectònic) i element configurador del microentorn funerari (eix paisatgístic). La chaîne opératoire, finalment, dona el marc per al gradient del domini de cinc valors i per a la taula T_LOST_ELEMENTS: un element desaparegut continua testimoniant un gest constructiu, i el registre l'ha de poder comptar com a tal.

## **2.3. Procediment de disseny**

El disseny segueix un procés iteratiu documentat versió a versió:

- Inventari de variables a partir de la documentació existent: fitxes de camp, treballs publicats (Toyne i Anzellini 2017; Epstein i Toyne 2016; Toyne et al. 2018), el TFM Chacha XR (Ribera-Torró 2023) i l'article derivat (Ribera-Torró et al. 2026).

- Modelatge entitat-relació i normalització fins a la 3FN, amb taules de consulta per a tots els camps categòrics.

- Validació amb les estructures documentades de Diablo Wasi, incloent-hi l'anàlisi fotogramètrica 3D dels models Metashape.

- Revisió contra el marc de recerca: cada camp queda etiquetat amb la hipòtesi que serveix.

- **Confrontació amb el corpus real (v10→v11):** l'entrada dels primers 35 registres va posar a prova l'esquema i va motivar la revisió més profunda del projecte, especificada en un document delta (rev. 5) i executada per un script de migració amb regles explícites de preservació de dades (secció 11.1).

# **3. Vocabulari arquitectònic normalitzat (A-X)**

Un dels avanços metodològics centrals de la base de dades és la normalització bilingüe (valencià/anglés) d'un vocabulari de 24 entrades arquitectòniques (A-X), derivat de l'anàlisi sistemàtica dels models fotogramètrics 3D de Diablo Wasi. Des de la v11 el vocabulari és una taula lookup pròpia (L_ELEMENTS: codi, noms, nivell, sistema, camp associat, descripció), cosa que el fa referenciable amb integritat —T_LOST_ELEMENTS hi apunta— i exportable com a metadada del projecte.

## **Criteris d'ordenació**

La codificació segueix tres criteris jerarquitzats: (1) NIVELL CONSTRUCTIU: de N0 (basament) cap a N1 (cos principal) i zona superior; (2) ALÇADA dins de cada nivell: de baix cap a dalt; (3) FUNCIÓ: primer l'element estructural portant, després el tractament decoratiu.

## **Elements i sistemes: la correcció conceptual de la v11**

Tres entrades del vocabulari són sistemes compositius, és a dir, el resultat de la combinació obligatòria d'altres elements: E+F+G → H (plataforma), N+O+Q → P (portal) i S+T → U (ràfec, sempre lapidi a T i U). La v10 els registrava com si foren elements més, amb un camp binari cadascun. La v11 corregeix la categoria: **H, P i U no tenen existència material pròpia** —són un judici sobre la combinació dels seus components— i per això esdevenen camps d'estat (Sys_Platform, Sys_Portal, Sys_Eave) amb un domini de sis valors: *Present complete / Present partial / Attested lost / Absent / Not applicable / Not observable*.

La correcció té dues conseqüències analítiques. La primera: la distinció entre «no aplicable» i «absent verificat» —la mateixa lògica del 0 vs 9 elevada al nivell del sistema— fa que els percentatges de presència siguen calculables sobre el denominador correcte. La segona: en l'anàlisi de coocurrència, H correlacionava trivialment amb E, F i G perquè n'era la suma; amb la matriu de QRY_13 reorganitzada en 20 columnes d'element i 5 de sistema, el doble comptatge desapareix abans d'arribar a R.

S'hi afigen dos camps d'agrupació sense lletra pròpia, Sys_Base (A-D) i Sys_Chamber (J, K, L, M, V, X), amb domini de quatre valors, que serveixen el gating del formulari i el mateix problema del denominador a escala de conjunt. La cornisa intercòs (I) i el coronament (R) no pertanyen deliberadament a cap sistema: veure un d'aquests elements sense poder resoldre cap sistema no és una contradicció sinó observabilitat parcial, la norma en un penya-segat.

L'element W (mur posterior) desapareix com a camp, absorbit per Sys_Chamber; en sobreviu el tipus (Rear_Closure_Type), perquè una cambra que usa la roca com a tancament posterior no es va construir com una que alça un mur. L'entrada W es conserva a L_ELEMENTS com a referència de vocabulari. Dos renomenaments resolen una col·lisió terminològica que les dades van demostrar perillosa: els antics Lateral_Walls i Lateral_Wall_Faces deien «lateral» sobre variables independents; ara són Return_Wall (V: mur perpendicular al pla de façana, retornant cap al penyal) i Facade_Flank (L: parament dins del pla de façana, flanquejant l'obertura).

La taula completa del vocabulari, amb la correspondència de camps v12, es troba a l'esquema tècnic (esquema_bbdd_estructures_v12.md, secció 8). Els criteris operatius de replicabilitat queden fixats per escrit: pilastra integrada al pla vs contrafort que en sobreïx; mènsula perpendicular i en voladís vs biga paral·lela salvant llum; suport de plataforma vs mènsula aïllada; ala de façana vs mur de retorn.

# **4. Arquitectura de la base de dades**

## **4.1. Taules i relacions (22 taules, 25 relacions)**

| **Cat.** | **Taula** | **Contingut i funció** |
| --- | --- | --- |
| Principal | T_STRUCTURES | Fitxa completa (120 camps). Font principal per a l'anàlisi estadística. |
| Principal | T_CONNECTIONS | Arestes físiques entre estructures. Des de v11, amb llista tancada de tipus de junta i el camp Chrono_Relation, que converteix el graf en potencialment dirigit (OE3, H03). |
| Principal | T_LOST_ELEMENTS | NOVA v11. L'evidència material darrere de cada valor 3 (desaparegut atestat): tipus d'evidència (lookup tancat), abast (element / cos / estructura sencera) i posició. Base evidencial d'OE3. |
| Secundària | T_DATING | Datacions radiocarbòniques (C14). N:1 amb T_STRUCTURES. |
| Secundària | T_INDIVIDUALS | Dades individuals (edat, sexe, preservació). N:1. |
| Secundària | T_GROUPS | Agrupacions funcionals (alineaments, xarxes, conjunts). N:1. |
| Secundària | T_DECORATIONS | Registre únic de decoració per estructura, posició i cos. Des de v11 sense duplicat de booleans. Clau per a H01 i H04. |
| Secundària | T_ARCH_FEATURES | Registre flexible d'elements no previstos a l'esquema. N:1. |
| Lookup | L_SITES, L_SECTORS | Jaciments i sectors. |
| Lookup | L_TYPOLOGY | Reconstruïda v11: cada tipologia porta Record_Class; MIX eliminada; ND desdoblada en Unclassifiable i Not yet classified. |
| Lookup | L_SUPPORT | Suport geomorfològic (9 valors; v11 afig Micro-ledge <50cm). |
| Lookup | L_ELEMENTS | NOVA v11. El vocabulari A-X com a taula (24 files, índex únic sobre Code). |
| Lookup | L_LOST_EVIDENCE | NOVA v11. Catàleg tancat de 9 evidències de pèrdua: filtre metodològic del valor 3. |
| Lookup | L_STATUS, L_MATERIAL_STATUS, L_VOL_METHOD, L_COORD_METHOD, L_GROUP_TYPE, L_CAMPAIGN, L_STRUCT_BODY, L_DEC_TYPE | Sense canvis de funció (L_STRUCT_BODY: entrada LWF renomenada FFL seguint el vocabulari). |

*Taula 3. Les 22 taules de la base de dades amb la seua funció. Les 25 relacions comprenen les 21 anteriors més les quatre de T_LOST_ELEMENTS (estructura, element, evidència, posició).*

## **4.2. Classes de registre: no tot registre és una estructura**

El corpus conté entrades que no són estructures funeràries construïdes: cavitats naturals, mènsules aïllades (traces estructurals de la xarxa de circulació), panells d'art rupestre. La v10 les tipificava però no les classificava, i qualsevol percentatge calculat sobre el total les incloïa indegudament. Record_Class —una classe per tipologia, mai per registre— resol el problema en el punt correcte de la normalització: *Built funerary structure* (EA-MAU, EA-CAM, EA-PLA-R/V, Unclassifiable), *Natural funerary context* (NIX, CAV), *Structural trace* (MEN), *Rock art panel* (PR) i *Pending classification* (estat de treball, exclòs de tota consulta analítica). La classe filtra les anàlisis i governa el primer nivell del gating del formulari (secció 10.2). El desdoblament de l'antic «ND» distingeix un veredicte epistèmic (Unclassifiable: evidència insuficient per a classificar una estructura construïda, com EA11, un perímetre de pigment sense construcció conservada) d'un simple estat del flux de treball (Not yet classified).

# **5. Variables arqueològiques de T_STRUCTURES (120 camps)**

La taula T_STRUCTURES concentra 120 camps. Els dominis observacionals són ara tres:

- **20 camps d'element A-X, domini de cinc valors** (0/1/2/3/9): la novetat central de la v11 (secció 5.1).
- **25 camps observacionals, domini 0/1/9** (alteracions, bioarqueologia, materials, morter, revoc, pigment, marc reculat, suport modificat i, des de v12, Dec_Present): «present parcial» i «desaparegut atestat» són categories morfològiques d'elements construïts, i forçar-les sobre processos o vestigis mobles no significaria res.
- **5 camps de sistema, domini textual** de sis valors (Sys_Platform/Portal/Eave) o quatre (Sys_Base/Chamber).

El valor per defecte de tots els camps BYTE passa de 9 a 0 en la v11: sota el domini nou, un registre nou comença de «examinat, res no hi és», i el 9 —que mai no va significar «no mirat» sinó «la posició no és examinable»— es reclama explícitament i a consciència. El canvi optimitza la velocitat d'entrada, i el risc epistèmic que comporta (zeros deixats per inèrcia) queda mitigat per la distinció entre els dos zeros (secció 5.2), pel gating del formulari (secció 10) i per la regla 17 de la bateria, que converteix els NULL heretats en llista de treball.

Només C14 i ChaXR_Documented es mantenen YESNO: no són observacions sobre l'estructura sinó metadades sobre el corpus propi, i el seu FALSE no és mai ambigu.

## **5.1. El domini de cinc valors: construït, sobreviu, intacte**

El domini binari de presència col·lapsava tres preguntes en una. El gradient 0/1/2/3/9 les separa: *què es va construir* (1+2+3), *què sobreviu* (1+2) i *què resta intacte* (1). Per a les hipòtesis constructives (H01, H04, H05) la magnitud pertinent és normalment la primera: el gest constructiu existí encara que l'element haja desaparegut. Per a les preguntes tafonòmiques i de conservació, les altres dues. El valor 3 és l'únic del domini que és una inferència i no una observació; per això la regla 2 de la bateria exigeix que cada 3 tinga una fila d'evidència a T_LOST_ELEMENTS, amb el tipus triat d'un catàleg tancat (encaix negatiu, forat de biga, cicatriu de despreniment, fragment despres in situ, empremta de morter, mènsules al buit, pigment sobre roca nua, murs truncats). Si cap evidència del catàleg no és aplicable, el valor correcte és 0 o 9: el lookup actua com a filtre metodològic de la inferència. Una fila amb abast *Body* o *Whole structure* cobreix tot el vocabulari de l'estructura d'una vegada, cosa que absorbeix l'antic camp Lost_Body_Evidence i en generalitza la funció.

## **5.2. Els dos zeros**

La coexistència del gating per sistemes i de la regla del col·lapse exigia distingir dos zeros semànticament diferents. El **0 d'asserció** —el sistema del camp és present o desaparegut atestat, o el camp no pertany a cap sistema— significa «la posició s'ha examinat i no hi havia res», i en una estructura col·lapsada és inverificable (regla 6). El **0 de farciment** —el sistema és no aplicable, no observable o absent— no asserta res: només evita un NULL, i la regla 4 l'exigeix activament sota un sistema no aplicable. Sense aquesta separació, una cova col·lapsada amb el sistema portal no aplicable faria saltar la regla del col·lapse exactament sobre els zeros que la regla de farciment demana. En l'exportació analítica, el 0 de farciment arriba com a NA d'origen (secció 9.2).

## **5.3. Estat de conservació dual**

Es manté la distinció entre ID_Arch_Status (arquitectura: Good/Fair/Pre-collapse/Collapsed/ND) i ID_Material_Status (vestigis mobles: Good/Fair/Poor/Absent/ND), amb les causes (Looting, Fire_Damage, Animal_Activity, Modern_Access) registrades a banda. La v11 hi afig dues connexions amb la resta de l'esquema: l'estat Collapsed invalida els zeros d'asserció (regla 6, amb avís en viu al formulari), i l'estat Good amb més del 30% d'elements parcials o perduts es marca com a incoherència a reconciliar (regla 5).

## **5.4. T_DECORATIONS com a registre únic i Dec_Present com a judici agregat**

La v11 elimina els 12 booleans de decoració de T_STRUCTURES: duplicaven T_DECORATIONS —que a més registra posició, tipus, color i substrat— i amb només 35 registres els dos sistemes ja discrepaven en sis casos. QRY_02 es reconstrueix sobre la taula de detall.

L'eliminació, però, va obrir un buit que l'addendum v12 tanca: sense cap camp agregat, «cap fila a T_DECORATIONS» no distingia l'absència verificada de decoració (0) de la façana no observable (9) — una violació de l'axioma fundacional del projecte. **Dec_Present** (BYTE 0/1/9) restaura el judici agregat; el detall continua vivint només a la taula filla, i les regles 22-23 mantenen les dues coses d'acord en les dues direccions. Al formulari, Dec_Present governa el subformulari de decoració.

## **5.5. Presència i tipus: el patró general**

Diversos parells de camps segueixen el mateix patró presència + tipus, amb la regla que el tipus és inaplicable —NULL, no «ND»— quan la presència és 0 o 9: Lintel + Lintel_Material (desdoblat en v11: el camp antic era un lookup de material que confonia presència amb tipus), Chamber_Roof + Chamber_Roof_Type, Plaster_Present + color i extensió, Pigment_Present + substrat, color i extensió, Mortar_Present + Mortar_Type. Les regles 10, 12-14 i 24-26 vigilen el patró, i el gating v12 el fa complir en entrada.

# **6. Estratègia de jerarquia, agrupació i connexió**

La BD implementa tres mecanismes complementaris i independents:

- **ID_Parent** (autoreferenciant): contenció física. Un EA-MAU construït dins d'una CAV apunta a la cova com a pare.
- **ID_Group** (→ T_GROUPS): agrupació funcional. Quatre EA-MAU en alineament vertical pertanyen al mateix conjunt sense que cap continga l'altre.
- **T_CONNECTIONS**: adjacència física entre parells d'estructures, que cap dels dos mecanismes anteriors captura. Des de v11, la llista tancada de tipus de junta (junta vertical adossada, superposició, junta travada, suport compartit, connexió aèria) i el camp **Chrono_Relation** converteixen el graf en potencialment dirigit: la direcció cronològica es llig de la junta —l'estructura que mostra la junta no travada contra el parament de l'altra és la posterior— i només una junta vertical adossada o una superposició la poden portar (regla 20); una junta travada implica contemporaneïtat. Aquest registre és el suport empíric relacional de H03.

*Regla pràctica per al cas paradigmàtic: un complex amb junta vertical no travada es registra com a dues entrades amb codis PALP consecutius, enllaçades a T_CONNECTIONS i agrupades a T_GROUPS; si el model 3D revela filades basals travades sota la junta aparent, correspon un registre únic amb Construction_Phases = 2.*

# **7. Integració espacial amb QGIS**

Sense canvis de procediment: T_STRUCTURES es carrega com a capa de punts (coordenades preferentment dels models Metashape georeferenciats, UTM 18S / WGS84, EPSG:32718) via QRY_07_Export_QGIS, i QRY_14 exporta les arestes de T_CONNECTIONS amb les coordenades dels dos extrems per a generar línies de la xarxa de circulació. La naturalesa vertical dels farallons requereix treballar en dos nivells: l'ortofoto vertical de cada sector per a les relacions relatives, i les coordenades absolutes per a la distribució general.

# **8. Anàlisi volumètrica i areal**

Sense canvis: Interior_Area_m2, Interior_Vol_m3 i Total_Vol_m3, amb ID_Vol_Method com a traçabilitat del procediment (càlcul geomètric per a EA-MAU regulars; extracció del model 3D via MeshLab o CloudCompare per a cavitats irregulars). La correlació amb MNI avalua la relació capacitat-ocupació (H01), i la variació per tipologia i jaciment (QRY_06) caracteritza les estratègies d'adaptació al suport (H02).

# **9. Consultes SQL per a l'anàlisi estadística**

La família de 26 consultes conserva les funcions de la versió anterior (distribucions, exportacions R i QGIS, jerarquies, cobertura, maçoneria, geologia) amb tres reconstruccions majors:

- **QRY_02** es reconstrueix sobre T_DECORATIONS (amb QRY_02s com a auxiliar i QRY_02a com a taula de banderes per motiu per al khi-quadrat).
- **QRY_13_AX_Pattern_Export** exporta ara **20 columnes AX_ + 5 columnes SYS_** amb semàntica NA: un component gatejat per un sistema no aplicable s'exporta com a NA, no com a 0, perquè el seu 0 emmagatzemat és de farciment. La definició element-sistema viu en una única font de veritat compartida amb la bateria de validació, de manera que exportació i validació no poden divergir mai.
- **QRY_14** incorpora Chrono_Relation (graf dirigit per a igraph/QGIS).

## **9.1. Estratègia d'anàlisi de patrons constructius (QRY_13)**

L'atomització en el vocabulari A-X converteix cada estructura en un vector d'elements i sistemes, complementat amb la maçoneria, el suport i les coordenades UTM. Les quatre famílies d'anàlisi en R es mantenen: coocurrència (phi, Jaccard) per a H01 i H04; clúster jeràrquic per a les famílies constructives de H05; anàlisi de correspondències amb LP i DW com a punts suplementaris; i autocorrelació espacial (I de Moran) per a l'estructura geogràfica dels patrons (H05, H03). Aquesta estratègia és el fil que connecta OE2 amb el debat sobre transmissió cultural i comunitats de pràctica.

## **9.2. Tractament estadístic dels dominis — advertència operativa**

L'advertència del denominador es manté íntegra: **prendre el total d'estructures com a denominador** infla la mostra amb casos mai avaluats, i com que la conservació difereix entre jaciments, la prova mesuraria preservació diferencial i es llegiria com a pràctica diferencial. Denominador vàlid = presents + absències verificades; la cobertura es reporta sempre junt amb la n.

El domini de cinc valors hi afig una segona obligació: **declarar la binarització**. Abans de qualsevol càlcul, els 9 es converteixen a NA (mai a 0), i després es tria conscientment la variable que respon la pregunta: *construït* (1+2+3), *sobreviu* (1+2) o *intacte* (1). Per a les hipòtesis constructives, normalment la primera. Els NA d'origen de l'exportació (zeros de farciment sota sistemes no aplicables) ja arriben buits i no s'han de reomplir. Les dues estratègies de mostra continuen sent legítimes i declarables: restricció a Facade_Observability = Complete, o distàncies binàries amb tractament de nuls; QRY_15 proporciona el recompte que justifica la decisió.

> Regla general: un 9 no és ni un 0 ni un 1, sinó una cel·la buida; i un 0 de farciment no és un 0 d'asserció, sinó un NA estructural. Qualsevol operació que els convertisca implícitament en dades reintrodueix el biaix de conservació al resultat.

## **9.3. Nivells constructius, cossos superposats i cossos perduts**

Es manté el criteri fixat: N0/N1/Superior són **categories funcionals, no un sistema de numeració** (mai N2, N3), i un cos pertany a N1 si conté o contenia una obertura d'accés. El recompte viu als camps N_Basal_Bodies i N_Chamber_Bodies, i la matriu es manté plana (una fila per estructura). La novetat v11 és el destí de l'evidència de cossos perduts: l'antic camp Lost_Body_Evidence desapareix i la seua informació esdevé una fila de T_LOST_ELEMENTS amb abast *Body*, que cobreix el vocabulari sencer del cos desaparegut — el cas diagnòstic continuen sent les bandes verticals de pigment sobre la roca nua, el fantasma d'un parament que la pintura ha sobreviscut. La conseqüència analítica és ara més fina: els elements del cos desaparegut ja no s'han de marcar 9 en bloc, sinó que poden marcar-se 3 amb l'evidència que ho justifica, cosa que els reincorpora al recompte de *què es va construir*.

## **9.4. Validació de coherència del registre (QRY_16, 26 regles)**

La bateria passa de deu (v10) a vint-i-sis regles, emmagatzemades en tres consultes parcials —una unió única de 31 branques supera el límit «query too complex» de JET— i unides a QRY_16_Validation_Check. Continua sent deliberadament **un informe i no una restricció de taula**: una regla dura impediria registrar una absència genuïnament observada en una estructura parcialment col·lapsada. Un resultat buit indica corpus coherent; la regla 17, que llista els camps observacionals encara NULL dels registres heretats, no és una queixa sinó la llista de treball de la revisió manual. El detall de les 26 regles es troba a l'esquema tècnic (secció 6.1); en síntesi cobreixen la coherència component-sistema (1-4), la coherència estat-elements (5-6), els parells presència-detall (7-10, 12-14, 24-26), la coherència classe-contingut (15-16), la bioarqueologia i els materials (18-19), les connexions i l'evidència de pèrdua (20-21) i la decoració (22-23).

# **10. Formulari principal F_STRUCTURES (v12)**

El formulari centralitza l'entrada de dades amb 12 pestanyes temàtiques, quatre subformularis vinculats (F_DECORATIONS, F_ARCH_FEATURES, F_CONNECTIONS, F_LOST_ELEMENTS) i tots els camps FK i de domini configurats com a ComboBox de dues columnes (valor anglés ocult, etiqueta valenciana visible). La pestanya 11.Sist es reorganitza al voltant dels sistemes: els cinc combos Sys_* dalt de tot i cada grup de components a continuació, perquè aquest és l'ordre real de les decisions — es decideix que hi ha plataforma abans de comptar-ne les mènsules.

## **10.1. Gating condicional de tres nivells**

El formulari activa i desactiva seccions segons el que ja s'ha registrat, de manera que la majoria de les incoherències que la bateria detectava esdevenen impossibles d'introduir:

- **Nivell 1 — classe de registre → pestanyes.** Derivat de Record_Class via ID_Typology: un context funerari natural tanca 2.Arq; una traça estructural i un panell d'art rupestre tanquen a més 6.Bio i 7.Mat; el panell tanca també 11.Sist. Un registre sense classificar ho manté tot obert (tancar pestanyes d'un registre de classe desconeguda amagaria les dades que cal per a classificar-lo). Matís deliberat: una cova conserva 11.Sist obert — DW-S01-EA09 té mènsules i plataforma dins d'una cavitat natural — i 3.Acab resta oberta per a totes les classes, perquè un panell d'art rupestre es defineix pel seu pigment.
- **Nivell 2 — portes → grups.** Cada Sys_* obri el seu grup de components només quan és present o desaparegut atestat; ID_Material_Status tanca els materials; Human_Remains tanca el detall bioarqueològic; Plaster_Present, Pigment_Present i Mortar_Present tanquen els seus detalls; Dec_Present tanca el subformulari de decoració.
- **Nivell 3 — element → detall.** El nombre i el rol de les mènsules només amb E present; el material de plataforma bloquejat amb rol aïllat; el material del dintell només amb Q present; el tipus de coberta només amb X present; el material de la cornisa només amb I present.

## **10.2. Emplenat ràpid amb confirmació**

Tancar una porta ofereix escriure el valor coherent a tot el grup depenent: components a 0 de farciment (sistema absent o no aplicable) o a 9 (no observable), detalls a buit, materials a 0 quan els vestigis són absents (exactament el que exigeix la regla 19). **Mai no s'escriu res sense un Sí explícit, i mai en navegar entre registres**: l'esdeveniment de navegació només activa i desactiva controls, l'anàleg formal de la regla B de la migració (cap valor emmagatzemat es toca sense intervenció deliberada). Completen el sistema tres avisos en viu: el recordatori d'evidència en marcar un 3 (regla 2), l'avís en roig d'estructura col·lapsada (regla 6) i l'ocultació del substrat «Plaster» quan el revoc és absent verificat (regla 13). Als subformularis, on la desactivació afectaria totes les files alhora, la lògica per fila és validació en desar: F_CONNECTIONS bloqueja una cronologia direccional sobre una junta que no la pot portar (regla 20) i autoassigna la contemporaneïtat de la junta travada; F_LOST_ELEMENTS exigeix el codi d'element quan l'abast és Element (regla 21).

El principi de disseny resultant: **el formulari preveu; la bateria detecta.** Cap dels dos substitueix l'altre — el gating es pot esquivar (taules obertes a mà, importacions, Centre de confiança desactivat) i algunes regles són intrínsecament de corpus.

## **10.3. Captions DAO per a la vista full de dades**

Es manté el mecanisme dual: etiquetes adjuntes als controls dels subformularis (capçaleres de columna en valencià) i captions DAO a nivell de TableDef com a reforç per a l'obertura directa de les taules.

# **11. Implementació tècnica**

## **11.1. Scripts VBA i rutes d'actualització**

El paquet v12 comprén quatre scripts amb dues rutes d'ús excloents:

- **Construcció des de zero** (BD en blanc): chachapoya_DB_v12.bas → BuildDB() crea les 22 taules, fixa els valors per defecte (45 camps BYTE a 0; sistemes a Absent), pobla els lookups, estableix les 25 relacions i genera les 26 consultes. A continuació, chachapoya_Form_v12_val.bas → BuildForm() crea els quatre subformularis, F_STRUCTURES amb 12 pestanyes, els combos de dues columnes, les captions DAO, el mòdul de gating i les validacions de subformulari.
- **Actualització amb dades** (BD v10 amb registres): migrate_v10_to_v11.bas → MigrateV10toV11() transforma l'esquema preservant cada registre, i upgrade_form_v12.bas → UpgradeV12() afig les peces v12 (Dec_Present, regles 22-26, F_LOST_ELEMENTS) i injecta el mateix gating que BuildForm. Tots dos són idempotents i re-executables.

L'script de migració mereix nota metodològica pròpia, perquè les seues regles són citables com a protocol de preservació de dades: ordre estricte *rename → add → populate → delete* (cap camp s'elimina abans que el seu substitut estiga poblat i verificat); **regla B**: cap NULL existent es toca mai, i els valors per defecte només afecten files futures; els esborrats finals estan doblement barrats per un commutador explícit i per precondicions comprovades en viu contra les dades, de manera que amb el commutador desactivat cada execució informa del progrés de la revisió manual en lloc d'esborrar; i cada bloc imprimeix recomptes d'afectació, cosa que converteix el log en pista d'auditoria de la migració.

## **11.2. Restriccions tècniques de VBA/Access**

Es mantenen les restriccions documentades: cap continuació de línia (patró sql = sql & "..."), codificació cp1252 sense caràcters no-ASCII al codi, paraules reservades evitades, patró de creació i renomenament de formularis, N-1 parèntesis per a N taules en JOIN, dbRelationDontEnforceIntegrity (valor 2) per a l'autoreferenciant i les dues de T_CONNECTIONS, procediments auxiliars Private, DROP previ per a objectes regenerables, i combos configurats per ControlSource. S'hi afigen les apreses en la iteració v11-v12: JET no pot convertir TEXT a BYTE in situ (el desdoblament del dintell exigeix renomenar per DAO, crear el camp nou i poblar-lo per UPDATE); un renomenament DAO no reescriu el SQL de les consultes existents (les consultes afectades queden trencades fins que es reconstrueixen, i l'ordre dels blocs ho preveu); una UNION de més de ~30 branques supera el límit «query too complex» (la bateria s'emmagatzema en tres parts); la injecció de mòduls de formulari exigeix l'accés al model d'objectes VBA al Centre de confiança (si està desactivat, tot es construeix igualment i l'script avisa); i amb la columna lligada oculta, Access imposa LimitToList (la via d'escapament és el valor «Other (see Notes)»).

## **11.3. Exportació per a l'anàlisi estadística**

Sense canvis de procediment: QRY_05 (exportació plana) i QRY_13 (matriu AX+SYS) via Dades externes → Exportar, amb els camps categòrics en format llegible gràcies als JOIN amb els lookups. La conversió 9 → NA i la declaració de la binarització (secció 9.2) són pas previ obligat en R.

# **12. Limitacions i perspectives**

La limitació principal continua sent la cobertura desigual del corpus (La Petaca Sector Sud amb ≥96 estructures identificades, no totes amb fitxa completa). Els camps de base documental i observabilitat, la regla 17 (llista de treball dels NULL) i QRY_15 converteixen aquesta limitació en una variable mesurada en lloc d'un buit silenciós. La resolució manual dels registres heretats de la migració (assignació dels cinc sistemes, reavaluació dels camps al domini nou, Dec_Present) és la tasca oberta immediata, i el mesurador de progrés dels blocs finals de la migració la fa auditable.

T_ARCH_FEATURES i L_DEC_TYPE mantenen l'escalabilitat de l'esquema sense modificar-ne l'estructura. El vocabulari A-X, ara formalitzat com a taula amb la distinció explícita element/sistema, és el primer intent de normalització terminològica per a les estructures funeràries Chachapoya en la bibliografia, i la seua aplicació creuada a La Petaca i Diablo Wasi és una de les aportacions originals del TFM.

De cara a publicacions futures, es manté la perspectiva de migració a GeoPackage (.gpkg) per a la integració nativa amb QGIS, amb Access com a interfície d'entrada.

# **13. Referències**

Epstein, L.; Toyne, J.M. (2016). When Space Is Limited: A Spatial Exploration of Pre-Hispanic Chachapoya Mortuary and Ritual Microlandscape. A: Osterholtz, A.J. (ed.), Theoretical Approaches to Analysis and Interpretation of Commingled Human Remains. Springer, Switzerland, pp. 97-124.

Ribera-Torró, E. (2023). Chacha XR. Una experiència immersiva de no-ficció per a l'arqueologia Chachapoya. Treball de Fi de Màster, Màster Universitari en Arts Visuals i Multimèdia, Universitat Politècnica de València.

Ribera-Torró, E.; Toyne, J.M.; Del Aguila, R.; Ribera, J.A.; Galexner, J.; Anzellini, A.; Pans, M. (2026). Extended Reality on Chachapoya Cliffside Necropolises: From Digital Documentation to Public Engagement. Open Archaeology, 12(1). DOI 10.1515/opar-2025-0071.

Toyne, J.M.; Anzellini, A. (2017). Sociedad, identidad y variedad en los mausoleos de La Petaca, Chachapoyas. Boletín de Arqueología PUCP, 23, 231-257.

Toyne, J.M.; Anzellini, A.; Epstein Miculas, L.; Mejías Pitti, I.; Puig Castell, J.; Guinot Castelló, S. (2018). Going Vertical: Using Vertical Progression Techniques to Explore a Cliff Necropolis in Late Precolumbian Chachapoyas, Peru. Advances in Archaeological Practice, 6. DOI 10.1017/aap.2018.31.
