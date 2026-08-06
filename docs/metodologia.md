**Universitat d'Alacant**

Màster en Arqueologia Professional i Gestió Integral del Patrimoni

**DISSENY D'UNA BASE DE DADES ARQUEOLÒGICA**

per a l'estudi de les necròpolis de penya-segat de La Petaca i Diablo Wasi

*(Leymebamba, Amazonas, Perú, s. IX-XVI)*

Autor: Esteve Ribera Torró

Directors: Dr. Ignasi Grau Mira (UA) i Dra. J. Marla Toyne (UCF)

Curs acadèmic 2024-2025

*Versió 6 del document — actualitzada segons la BD v14 (agost 2026). Consolida les iteracions v10→v11, v12, v13 i v14.*

# **Resum**

El present document descriu el disseny, la justificació i la implementació de la base de dades relacional destinada a la documentació sistemàtica i l'anàlisi estadística de les estructures funeràries de les necròpolis de penya-segat de La Petaca i Diablo Wasi (Leymebamba, Departament d'Amazonas, Perú, s. IX-XVI d.n.e.). La base de dades constitueix l'eix vertebrador de la metodologia del Treball de Fi de Màster, en tant que permet centralitzar les variables arqueològiques, establir relacions jeràrquiques entre elements, exportar dades per a l'anàlisi estadística i vincular el registre arqueològic amb el Sistema d'Informació Geogràfica (SIG) implementat en QGIS.

L'esquema segueix els principis de la tercera forma normal (3FN) i inclou 22 taules, una taula principal (T_STRUCTURES) amb 127 camps, 25 relacions, la família de 28 consultes SQL i un formulari d'entrada de dades amb 12 pestanyes, sis subformularis i interfície en valencià. Un dels avanços metodològics centrals és la normalització d'un vocabulari arquitectònic bilingüe (valencià/anglés) de 24 entrades (A-X), derivat de l'anàlisi fotogramètrica 3D de les estructures de Diablo Wasi, formalitzat des de la v11 com a taula lookup pròpia (L_ELEMENTS).

La versió actual consolida l'esquema al voltant de dos principis metodològics. El primer, ja establit en versions anteriors: **el registre ha de distingir sempre l'absència verificada de la manca d'observació**. El segon, incorporat amb la iteració v10→v11: **el registre ha de distingir el que es va construir del que sobreviu**. Els 20 camps d'element del vocabulari A-X adopten per això un domini de cinc valors (0 = Absent / 1 = Present complet / 2 = Present parcial / 3 = Desaparegut atestat / 9 = No observable), on el valor 3 —l'únic del domini que és una inferència i no una observació— exigeix evidència material registrada a la taula nova T_LOST_ELEMENTS. Els tres antics «elements» H, P i U es reconeixen com el que sempre van ser, sistemes compositius, i esdevenen camps d'estat propis (Sys_Platform, Sys_Portal, Sys_Eave) amb un domini de sis valors que separa «no aplicable» d'«absent verificat». La classe de registre (Record_Class, a L_TYPOLOGY) filtra tota anàlisi: dels 35 registres actuals, cinc no són estructures, i qualsevol percentatge calculat sobre 35 seria erroni.

L'addendum v12 tanca el cercle en el punt d'entrada: el camp Dec_Present restaura el judici agregat sobre la decoració que l'eliminació dels booleans havia deixat orfe, i el formulari incorpora un gating condicional de tres nivells amb emplenat ràpid confirmat, de manera que la majoria de les incoherències que la bateria detectava a posteriori esdevenen impossibles d'introduir. El principi de disseny resultant: el formulari preveu en el moment d'entrada; la bateria detecta a escala de corpus.

La iteració v13, finalment, afina la semàntica dels dominis a partir de la confrontació amb el corpus real. Tres decisions la resumeixen. Primera: **el valor per defecte passa a NULL** en tot camp que cap regla de coherència no obligue a emplenar, perquè un zero fals asserta de més —infla el denominador amb absències que ningú no ha verificat— mentre que un buit només diu que el treball està pendent; això preserva tres estats distingibles (no avaluat / avaluat i no examinable / avaluat i absent) que la uniformitat anterior col·lapsava. Segona: **el domini de cinc valors queda restringit als vint elements del vocabulari arquitectònic i a cap altre camp**. Aquesta restricció és, ella mateixa, el resultat d'una correcció instructiva. Una primera formulació va estendre el gradient a les «capes aplicades a la fàbrica» —morter, revoc, pigment—, però la confrontació amb casos reals va mostrar que el criteri estava mal enunciat: el que habilita el valor «desaparegut atestat» no és que la cosa siga una capa, sinó que **l'evidència de la pèrdua siga de naturalesa distinta de la cosa perduda** (un encaix buit no és una mènsula), i el que habilita «present parcial» és **poder inferir l'extensió original**. El pigment falla els dos criteris, perquè l'única evidència de pigment és pigment; el morter els falla per una raó distinta, en no ser una capa que es perd per zones sinó un atribut de la tècnica de fàbrica. La regla definitiva distingeix **discret de continu**: el que es compta element a element admet gradació de conservació; el que s'estén sobre superfícies, no. Tercera: **el suport geològic distingeix la diàclasi de la junta d'estratificació**, dues formes que el vocabulari havia fos en una casella i que ofereixen lògiques estructurals oposades —confinament contra encastament—, cosa que converteix el ritme estratigràfic del faralló en variable explicativa de dues decisions constructives simultànies. A això s'hi afig la integració de la pintura rupestre associada en la mateixa taula que la decoració arquitectònica, distingida per la posició i no per una taula pròpia.

A diferència de les iteracions anteriors, **la v13 no migra la base sinó que es construeix de zero i els registres es reintrodueixen**. La decisió elimina l'aparat de migració i, sobretot, els valors heretats de defectes de versions prèvies que un diagnòstic sistemàtic va destapar: sota v13, cada valor emmagatzemat és un judici deliberat, que és la precondició per a publicar qualsevol percentatge.

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
| H02 | Geologia com a factor determinant | ID_Support + Support_Mode (v13: diàclasi vs junta d'estratificació), ID_Support_Secondary, Support_Width_cm, Support_Depth_cm, Support_Modified, Altitude_masl, Height_Above_Base_m, Base_Level, Tie_Walls, Rear_Closure_Type, Chamber_Roof_Type (QRY_12) |
| H03 | Saturació espacial acumulativa (argument morfològic-estructural) | Construction_Phases, Phase_Evidence, T_CONNECTIONS.Chrono_Relation (graf dirigit, QRY_14), T_GROUPS, ID_Group, ID_Parent, QRY_07 (exportació QGIS) |
| H04 | Seqüència operativa consistent (chaîne opératoire) | ID_Typology + Record_Class, matriu AX+SYS (QRY_13), T_LOST_ELEMENTS (el gest constructiu del valor 3), Construction_Phases, T_DECORATIONS, T_DATING (C14), Chrono_Start/End_Cent |
| H05 | Tradició constructiva compartida entre LP i DW | Matriu AX+SYS per estructura (QRY_13): coocurrència, clúster de Jaccard, anàlisi de correspondències; Masonry_Quality, Masonry_Type |
| H06 | Visibilitat i marcatge territorial | Facade_Orientation, Visibility_Valley, Pigment_Extent (valor perimetral), RockArt_Present i files ROC de T_DECORATIONS (QRY_17), Coord_E/N_UTM, Altitude_masl (QRY_07 → anàlisi de visibilitat en QGIS) |

*Taula 1. Hipòtesis del TFM i camps de la base de dades que les operacionalitzen (marc revisat, sis hipòtesis; noms de camp v12).*

# **2. Fonamentació metodològica**

## **2.1. Convenció d'idiomes**

La base de dades està completament en anglés (noms de taules, camps i **valors emmagatzemats**, incloent-hi les llistes de valors dels ComboBox i el contingut de les taules lookup). Aquesta decisió respon a tres factors: (a) la publicació derivada del projecte (Ribera-Torró et al. 2026, Open Archaeology) usa terminologia anglesa per a elements i camps; (b) la codirectora del TFM, la Dra. J. Marla Toyne (UCF), treballa en anglés; i (c) l'exportació a R/SPSS amb capçaleres i valors en anglés és l'estàndard en la publicació científica internacional i garanteix la reproduïbilitat de les anàlisis.

La interfície d'entrada de dades, en canvi, és en valencià: totes les etiquetes visibles del formulari es mostren en valencià, sense que això afecte mai els valors emmagatzemats. Des de la v11 aquesta separació és tècnicament infrangible: tots els combos de domini són de dues columnes, amb la columna emmagatzemada (anglés) oculta i l'etiqueta (valencià) visible, de manera que reanomenar una etiqueta no pot tocar mai una dada.

La v14 completa aquesta separació, que fins llavors era **parcial sense que resultara evident**. El formulari tenia dues menes de desplegable i només una estava localitzada: els que porten les dues columnes escrites al codi es mostraven en valencià, mentre que els que llegeixen d'una taula de consulta mostraven el terme anglés emmagatzemat. El resultat era mig formulari en cada idioma. La correcció aplica el mateix patró als segons, afegint a cada taula de consulta una columna d'etiqueta valenciana al costat del terme de referència anglés; ací el valor emmagatzemat és un identificador numèric, cosa que fa la separació encara més robusta.

Aquesta arquitectura deixa oberta, sense cost, l'ampliació del vocabulari controlat al castellà —llengua de treball del projecte al Perú—, que seria afegir una tercera columna d'etiquetes. **El que no seria barat, i convé que quede raonat, és un selector d'idioma en temps d'execució**: al formulari conviuen quatre menes de text —desplegables de taula, desplegables de llista de valors, etiquetes dels controls i noms de les pestanyes— i les columnes de la taula de consulta només afecten la primera. Commutar-la sola reproduiria exactament el defecte que aquesta revisió corregeix. La via consistent, si algun dia fa falta, és parametritzar l'script de construcció del formulari amb un codi d'idioma i generar un formulari sencer i independent per a cada llengua. Els documents del TFM (en valencià), les presentacions a congressos (en anglés o valencià) i les presentacions al Perú (en castellà) mantenen la seua llengua corresponent.

## **2.2. Marc disciplinar**

El disseny de les variables arqueològiques s'ha fonamentat en dues disciplines complementàries. L'Arqueologia de l'Arquitectura aporta el marc per a la caracterització constructiva (materials, tècniques, decoració) i la tipologia de les estructures. L'Arqueologia del Paisatge proporciona el marc per a l'anàlisi de la distribució espacial i les relacions amb la geologia i l'entorn natural. La integració de la BD amb QGIS permet implementar anàlisis de densitat, visibilitat i distribució vertical directament relacionades amb les hipòtesis H02, H03 i H06. H02 actua com a eix de connexió entre les dues disciplines: és alhora factor constructiu (eix arquitectònic) i element configurador del microentorn funerari (eix paisatgístic). La chaîne opératoire, finalment, dona el marc per al gradient del domini de cinc valors i per a la taula T_LOST_ELEMENTS: un element desaparegut continua testimoniant un gest constructiu, i el registre l'ha de poder comptar com a tal.

## **2.3. Procediment de disseny**

El disseny segueix un procés iteratiu documentat versió a versió:

- Inventari de variables a partir de la documentació existent: fitxes de camp, treballs publicats (Toyne i Anzellini 2017; Epstein i Toyne 2016; Toyne et al. 2018), el TFM Chacha XR (Ribera-Torró 2023) i l'article derivat (Ribera-Torró et al. 2026).

- Modelatge entitat-relació i normalització fins a la 3FN, amb taules de consulta per a tots els camps categòrics.

- Validació amb les estructures documentades de Diablo Wasi, incloent-hi l'anàlisi fotogramètrica 3D dels models Metashape.

- Revisió contra el marc de recerca: cada camp queda etiquetat amb la hipòtesi que serveix.

- **Confrontació amb el corpus real (v10→v11):** l'entrada dels primers 35 registres va posar a prova l'esquema i va motivar la revisió més profunda del projecte, especificada en un document delta (rev. 5) i executada per un script de migració amb regles explícites de preservació de dades (secció 11.1).

- **Diagnòstic sistemàtic previ a la decisió (v12→v13):** abans d'implementar cap canvi, un script de només lectura va respondre nou comprovacions sobre el corpus (variància real dels camps, distribució dels valors per defecte, freqüència de cada forma de suport, casos que justificarien un canvi de domini). Diverses decisions de disseny se'n van derivar directament, i almenys una es va **descartar** per falta de base empírica. Aquest pas explicita un principi que les versions anteriors aplicaven de manera implícita: **cap camp nou ni cap canvi de domini s'introdueix sense comprovar què diuen les dades existents**, perquè afinar la codificació d'una variable sense variància és treball perdut i multiplicar categories sense casos genera soroll.

# **3. Vocabulari arquitectònic normalitzat (A-X)**

Un dels avanços metodològics centrals de la base de dades és la normalització bilingüe (valencià/anglés) d'un vocabulari de 24 entrades arquitectòniques (A-X), derivat de l'anàlisi sistemàtica dels models fotogramètrics 3D de Diablo Wasi. Des de la v11 el vocabulari és una taula lookup pròpia (L_ELEMENTS: codi, noms, nivell, sistema, camp associat, descripció), cosa que el fa referenciable amb integritat —T_LOST_ELEMENTS hi apunta— i exportable com a metadada del projecte.

## **Criteris d'ordenació**

La codificació segueix tres criteris jerarquitzats: (1) NIVELL CONSTRUCTIU: de N0 (basament) cap a N1 (cos principal) i zona superior; (2) ALÇADA dins de cada nivell: de baix cap a dalt; (3) FUNCIÓ: primer l'element estructural portant, després el tractament decoratiu.

## **Elements i sistemes: la correcció conceptual de la v11**

Tres entrades del vocabulari són sistemes compositius, és a dir, el resultat de la combinació obligatòria d'altres elements: E+F+G → H (plataforma), N+O+Q → P (portal) i S+T → U (ràfec, sempre lapidi a T i U). La v10 els registrava com si foren elements més, amb un camp binari cadascun. La v11 corregeix la categoria: **H, P i U no tenen existència material pròpia** —són un judici sobre la combinació dels seus components— i per això esdevenen camps d'estat (Sys_Platform, Sys_Portal, Sys_Eave) amb un domini de sis valors: *Present complete / Present partial / Attested lost / Absent / Not applicable / Not observable*.

Un cas límit resolt en la v14 delimita bé l'abast del concepte de sistema. El marc reculat de façana estava registrat com a qualificador del sistema portal, però el recul afecta el **pla del parament sencer** i el portal hi queda inscrit, de manera que la dependència era inversa i el gating el bloquejava precisament allà on hi ha recul sense portal. Es va valorar promoure la façana a sistema propi i **es va descartar**: els cinc sistemes existents són combinacions de components identificables o camps d'agrupació per al control d'entrada, mentre que una façana no és una combinació de res — és el pla on tota la resta passa, i **no podria valdre mai «absent»**. Un sistema que no pot ser absent no fa el que fan els sistemes, que és precisament permetre distingir l'absència verificada de la inaplicabilitat. El camp passa, doncs, al bloc de la façana com el que és: un qualificador del pla.

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

# **5. Variables arqueològiques de T_STRUCTURES (128 camps)**

La taula T_STRUCTURES concentra 128 camps. Els dominis observacionals són tres:

- **20 camps amb domini de cinc valors** (0/1/2/3/9): els elements A-X, i només ells (secció 5.2).
- **26 camps amb domini 0/1/9**: atributs de tècnica, capes contínues, processos, bioarqueologia, materials culturals, qualificadors i judicis agregats.
- **5 camps de sistema, domini textual** de sis valors (Sys_Platform/Portal/Eave) o quatre (Sys_Base/Chamber).

El **valor per defecte** és una decisió que ha evolucionat en tres passos, i val la pena seguir-la perquè il·lustra com el disseny s'ha anat ajustant al que el registre realment necessita. La v10 usava 9 (no observable) per a tots els camps booleans convertits; la v11 el va canviar a 0 perquè el gating dels sistemes constructius **exigeix** un zero de farciment als components d'un sistema no aplicable; la v13 conserva eixe 0 **només allà on la regla el necessita** —els 20 camps d'element— i el substitueix per **NULL a la resta**.

El raonament és una asimetria d'errors. Un 0 fals **asserta de més**: infla el denominador amb absències que ningú no ha verificat, i com que el biaix correlaciona amb la qualitat d'observació —que difereix entre els dos jaciments— reintrodueix exactament l'artefacte que tot el sistema de dominis existeix per a evitar. Un buit fals només ajorna feina. I NULL i no 9, perquè el 9 no és una casella buida sinó una afirmació: «aquesta posició no es pot examinar». Usar-lo com a defecte faria un camp intacte indistingible d'un de marcat a consciència, i desactivaria la funció de la regla 17, que és llistar el que encara no s'ha avaluat. El diagnòstic previ a la v13 ho va confirmar empíricament: una vintena de valors 9 heretats del defecte v10 en un mateix camp, cap d'ells un judici real.

El resultat és que un camp observacional té **tres estats distingibles**: buit (no avaluat encara), 9 (avaluat i no examinable) i 0 (avaluat i absent). Els dos primers són `NA` en l'anàlisi, però només el segon és un judici, i la seua freqüència mesura l'observabilitat del corpus.

Només C14 i ChaXR_Documented es mantenen YESNO: no són observacions sobre l'estructura sinó metadades sobre el corpus propi, i el seu FALSE no és mai ambigu.

## **5.1. El domini de cinc valors: construït, sobreviu, intacte**

El domini binari de presència col·lapsava tres preguntes en una. El gradient 0/1/2/3/9 les separa: *què es va construir* (1+2+3), *què sobreviu* (1+2) i *què resta intacte* (1). Per a les hipòtesis constructives (H01, H04, H05) la magnitud pertinent és normalment la primera: el gest constructiu existí encara que l'element haja desaparegut. Per a les preguntes tafonòmiques i de conservació, les altres dues. El valor 3 és l'únic del domini que és una inferència i no una observació; per això la regla 2 de la bateria exigeix que cada 3 tinga una fila d'evidència a T_LOST_ELEMENTS, amb el tipus triat d'un catàleg tancat (encaix negatiu, forat de biga, cicatriu de despreniment, fragment despres in situ, empremta de morter, mènsules al buit, pigment sobre roca nua, murs truncats). Si cap evidència del catàleg no és aplicable, el valor correcte és 0 o 9: el lookup actua com a filtre metodològic de la inferència. Una fila amb abast *Body* o *Whole structure* cobreix tot el vocabulari de l'estructura d'una vegada, cosa que absorbeix l'antic camp Lost_Body_Evidence i en generalitza la funció.

## **5.2. Els dos zeros, i les capes aplicades**

La coexistència del gating per sistemes i de la regla del col·lapse exigia distingir dos zeros semànticament diferents. El **0 d'asserció** —el sistema del camp és present o desaparegut atestat, o el camp no pertany a cap sistema— significa «la posició s'ha examinat i no hi havia res», i en una estructura col·lapsada és inverificable (regla 6). El **0 de farciment** —el sistema és no aplicable, no observable o absent— no asserta res: només evita un NULL, i la regla 4 l'exigeix activament sota un sistema no aplicable. Sense aquesta separació, una cova col·lapsada amb el sistema portal no aplicable faria saltar la regla del col·lapse exactament sobre els zeros que la regla de farciment demana. En l'exportació analítica, el 0 de farciment arriba com a NA d'origen (secció 9.2).

Als camps de tres valors, el 0 fon dues coses: «mai no en va tindre» i «en tenia i s'ha perdut». Es va valorar estendre el gradient de cinc valors a les capes aplicades a la fàbrica —morter, revoc, pigment— per a resoldre-ho, i el cas semblava sòlid: el catàleg d'evidències de pèrdua ja contenia «empremta de morter», és a dir, preveia una situació que el domini del camp no podia expressar.

**L'extensió es va descartar en confrontar-la amb casos reals**, i el procés val la pena de consignar-se perquè va obligar a formular millor el criteri. El que habilita el valor «desaparegut atestat» no és que la cosa siga una capa, sinó que **l'evidència de la pèrdua siga de naturalesa distinta de la cosa perduda**: un encaix buit no és una mènsula, i per això funciona. El pigment falla aquest criteri de manera irreparable, perquè **l'única evidència de pigment és pigment** — si en queda rastre és present, i si no en queda no es pot atestar res, de manera que «desaparegut» i «absent» es confondrien necessàriament. El que habilita «present parcial», al seu torn, és **poder inferir l'extensió original**, cosa impossible en una superfície pintada; i en tot cas els camps d'extensió ja registraven la conservació parcial amb el valor «traces», que el nou valor duplicava. El morter falla per una raó distinta: **no és una capa que es perd per zones sinó un atribut de la tècnica de fàbrica** — un mur en sec no ho és a trossos, i el que varia amb la conservació és la visibilitat, que ja registren els camps d'observabilitat.

La regla definitiva distingeix, doncs, **discret de continu**: cinc valors per als vint elements del vocabulari, que són unitats constructives comptables amb integritat física i extensió inferible; tres valors per a tota la resta. Això no és uniformitat, però és regla — que és el que fa l'esquema replicable per un tercer.

## **5.3. Estat de conservació dual**

Es manté la distinció entre ID_Arch_Status (arquitectura: Good/Fair/Pre-collapse/Collapsed/ND) i ID_Material_Status (vestigis mobles: Good/Fair/Poor/ND), amb les causes (Looting, Fire_Damage, Animal_Activity, Modern_Access) registrades a banda. La v11 hi afig dues connexions amb la resta de l'esquema: l'estat Collapsed invalida els zeros d'asserció (regla 6, amb avís en viu al formulari), i l'estat Good amb més del 30% d'elements parcials o perduts es marca com a incoherència a reconciliar (regla 5).

La v14 corregeix un defecte d'aquesta parella que hi era des del principi. L'escala dels vestigis mobles barrejava dues variables: *bo*, *regular* i *dolent* són graus de conservació, però *absent* és una afirmació de presència — **el mateix error que el camp del dintell tenia en la v10**, quan un lookup de material confonia presència amb tipus. El desdoblament és idèntic al que ja es va aplicar allà: un camp de presència amb el domini de tres valors, anàleg al de les restes humanes, que actua com a porta de tot el bloc, i una escala de conservació que ja només parla de conservació. Convé retindre l'asimetria que això revela: **l'estat arquitectònic no s'ha de desdoblar igual, perquè una estructura sempre existeix — és el registre mateix**. Els vestigis mobles poden no existir, i per això necessiten un camp que l'arquitectura no necessita.

## **5.4. T_DECORATIONS com a registre únic i Dec_Present com a judici agregat**

La v11 elimina els 12 booleans de decoració de T_STRUCTURES: duplicaven T_DECORATIONS —que a més registra posició, tipus, color i substrat— i amb només 35 registres els dos sistemes ja discrepaven en sis casos. QRY_02 es reconstrueix sobre la taula de detall.

L'eliminació, però, va obrir un buit que l'addendum v12 tanca: sense cap camp agregat, «cap fila a T_DECORATIONS» no distingia l'absència verificada de decoració (0) de la façana no observable (9) — una violació de l'axioma fundacional del projecte. **Dec_Present** restaura el judici agregat; el detall continua vivint només a la taula filla, i les regles 22-23 mantenen les dues coses d'acord en les dues direccions. Al formulari, Dec_Present governa el subformulari de decoració.

La v13 estén aquest mateix disseny a la **pintura rupestre associada**. El problema era el mateix de sempre en una forma nova: el pigment aplicat sobre la penya al voltant d'una estructura no tenia on registrar-se sense duplicar maquinària. La solució adoptada no crea cap taula: l'art rupestre associat esdevé **files de la mateixa taula** que la decoració arquitectònica, distingides per la posició. L'únic obstacle real era que el catàleg de posicions només contenia posicions arquitectòniques, de manera que una pintura sobre la roca acabava classificada com a «posició no determinada» —que significa una cosa ben distinta de «posició no arquitectònica»: el mateix patró d'error que la distinció entre absència i manca d'observació existeix per a evitar. Tres entrades noves de tipus rupestre ho resolen, i el tipus de nivell d'aquestes entrades passa a ser el **discriminador analític** entre els dos conjunts: separar-los és una condició en una consulta, i obtindre'ls units no exigeix cap unió.

`RockArt_Present` és el germà de `Dec_Present`, un judici agregat per a cada meitat de la taula, coherent amb les dues subseccions en què es divideix la pestanya de decoració del formulari.

**Criteri d'associació.** Quan una pintura és fila d'una estructura i quan és registre independent es decideix amb quatre tests en ordre de prioritat: continuïtat física amb la fàbrica (aleshores és pigment de l'estructura), compartir suport geològic, emmarcar un element de l'estructura, i en darrer terme fitxa pròpia vinculada per agrupació. Es va descartar un llindar mètric perquè la distància és de segona passada i l'associació s'ha de decidir en entrar el registre, i perquè en una paret vertical la distància euclidiana ignora el que estructura aquests jaciments: una marca a més d'un metre però dins de la mateixa escletxa està relacionada; una a mig metre separada per un banc que trenca la continuïtat, no necessàriament.

## **5.5. Presència i tipus: el patró general**

Diversos parells de camps segueixen el mateix patró presència + tipus, amb la regla que el tipus és inaplicable —NULL, no «ND»— quan la presència és 0 o 9: Lintel + Lintel_Material (desdoblat en v11: el camp antic era un lookup de material que confonia presència amb tipus), Chamber_Roof + Chamber_Roof_Type, Plaster_Present + color i extensió, Pigment_Present + substrat, color i extensió, Mortar_Present + Mortar_Type. Les regles 10, 12-14 i 24-26 vigilen el patró, i el gating v12 el fa complir en entrada.

# **6. Estratègia de jerarquia, agrupació i connexió**

La BD implementa tres mecanismes complementaris i independents:

- **ID_Parent** (autoreferenciant): contenció física. Un EA-MAU construït dins d'una CAV apunta a la cova com a pare.
- **ID_Group** (→ T_GROUPS): agrupació funcional. Quatre EA-MAU en alineament vertical pertanyen al mateix conjunt sense que cap continga l'altre.
- **T_CONNECTIONS**: adjacència física entre parells d'estructures, que cap dels dos mecanismes anteriors captura. Des de v11, la llista tancada de tipus de junta (junta vertical adossada, superposició, junta travada, suport compartit, connexió aèria) i el camp **Chrono_Relation** converteixen el graf en potencialment dirigit: la direcció cronològica es llig de la junta —l'estructura que mostra la junta no travada contra el parament de l'altra és la posterior— i només una junta vertical adossada o una superposició la poden portar (regla 20); una junta travada implica contemporaneïtat. Aquest registre és el suport empíric relacional de H03.

La v14 hi afig dues peces. La primera és un tipus de connexió per a l'**associació de verticalitat** entre estructures, que registra una relació observable —proximitat, posició relativa, morfologia del faralló— sense afirmar-ne el mecanisme: la hipòtesi que una mènsula servira de politja o de suport d'escala continua sent especulativa i es registra com a element personalitzat amb els seus indicis, però l'associació sí que s'observa i és material directe per a OE3. Quan concorren dues relacions —una junta vertical entre les fàbriques sobre una base compartida— el criteri fixat és que **la junta mana sobre el suport**, perquè la junta porta la direcció cronològica i el suport compartit no.

La segona peça resol un problema de registre que la nomenclatura ja intuïa. Quan dues unitats constructives coherents s'assenten sobre una mateixa base, es registren com a estructures separades amb codis consecutius (Xa, Xb) precisament per a no col·lapsar atributs que difereixen; però la relació entre elles vivia **només a la cadena del codi**, com a hàbit textual, i comptar quantes unitats supraestructurals hi ha exigia coincidències de text. Un tipus d'agrupació nou la fa consultable, amb un criteri formulat de manera prou ampla: **dues o més unitats constructives coherents amb atributs propis, no necessàriament dues cambres** — un mausoleu amb una plataforma en voladís qualifica, i és un cas freqüent a La Petaca. La divisió del treball entre els dos mecanismes és neta: la connexió és binària i descriu la junta, l'agrupació és n-ària i descriu la unitat.

*Regla pràctica per al cas paradigmàtic: un complex amb junta vertical no travada es registra com a dues entrades amb codis PALP consecutius, enllaçades a T_CONNECTIONS i agrupades a T_GROUPS; si el model 3D revela filades basals travades sota la junta aparent, correspon un registre únic amb Construction_Phases = 2.*

# **6bis. El suport geològic com a variable explicativa (v13)**

La v13 desdobla el que fins ara era un únic valor de suport, «fissura», en dues formes que la geologia del faralló produeix per mecanismes distints i que ofereixen al constructor lògiques oposades. La **diàclasi** és una fractura vertical que travessa els bancs: aporta dos paraments laterals, constreny la planta i s'omple de material per a generar el nivell basal, de manera que la roca actua com a encofrat — la seua lògica és el **confinament**. La **junta d'estratificació** és el rebaix horitzontal que deixa l'erosió diferencial d'un estrat tou entre bancs competents: aporta una ranura contínua on empotrar jàcenes o mènsules, amb treball a tallant i moment en lloc de compressió — la seua lògica és l'**encastament**. Que totes dues compartiren casella era un accident del vocabulari, no una afinitat real: en català i en anglés, «fissura» serveix per a les dues.

El guany per a H02 és considerable, perquè permet formular la hipòtesi en termes molt més forts que una preferència per determinades repises. El **ritme estratigràfic** del faralló —l'alternança de bancs competents i estrats erosionables— genera *simultàniament* les superfícies on es reposa i les ranures on s'ancora: una sola propietat geològica que condiciona dues decisions constructives distintes. Aquesta és la geologia operant a dues escales alhora, i és empíricament demostrable amb el registre.

El mode mecànic de cada forma **no es registra com a camp** sinó que s'escriu a la descripció del lookup. La raó és de disseny: es dedueix deterministament del nom de la forma, de manera que un camp per registre seria derivable —no informació, sinó una oportunitat de contradicció— i, amb el corpus actual, quasi constant per forma. El que sí que era imprescindible és **escriure'l**, perquè la deducció siga legítima i replicable per un tercer.

La v14 hi afig una tercera forma que no és una subdivisió sinó una **negació**: el terreny a la base del cingle, per a les estructures que no estan penjades. El seu interés és que **fa explícit un supòsit que les altres deu formes compartien en silenci** —que l'estructura està elevada sobre un buit—, i el seu criteri operatiu és per això negatiu i comprovable: no hi ha buit a sota. És, a més, l'única forma sobre sediment i no sobre roca, cosa que no és un matís: sobre roca l'estructura no assenta i no cal fonamentació; sobre sediment sí, i això podria explicar la presència o absència de les jàcenes basals encastades i dels murets transversals.

Convé retindre un solapament que el registre ha de resoldre a consciència: **micro-repisa i junta d'estratificació són el mateix fenomen vist des de cares oposades.** Quan un estrat tou s'erosiona, el banc inferior queda volat (i s'hi pot reposar) i alhora queda una ranura (on es pot encastar). Quina de les dues coses és depén exclusivament del que va fer el constructor, no de la roca. Sense el criteri escrit, dos observadors codificarien la mateixa paret de manera distinta.

**Criteri de registre.** `ID_Support` és la forma que **rep la càrrega** —siga per repòs o per confinament—, i `ID_Support_Secondary` la que **estabilitza o allotja**. La prova operativa és contrafactual: *què cauria si eixa forma no hi fora?* Els elements construïts són transparents: cal seguir la càrrega cap avall a través de la maçoneria fins a trobar roca, tenint present que la cadena pot acabar **lateralment** i no per sota — un basament dins d'una diàclasi transmet la càrrega a les parets de l'escletxa. Es va descartar explícitament la regla «primari = nivell basal, secundari = cos de cambra»: quan la cambra s'allotja directament a la roca, el suport primari és a cota del cos principal per definició.

# **7. Integració espacial amb QGIS**

Sense canvis de procediment: T_STRUCTURES es carrega com a capa de punts (coordenades preferentment dels models Metashape georeferenciats, UTM 18S / WGS84, EPSG:32718) via QRY_07_Export_QGIS, i QRY_14 exporta les arestes de T_CONNECTIONS amb les coordenades dels dos extrems per a generar línies de la xarxa de circulació. La naturalesa vertical dels farallons requereix treballar en dos nivells: l'ortofoto vertical de cada sector per a les relacions relatives, i les coordenades absolutes per a la distribució general.

# **8. Anàlisi volumètrica i areal**

La v14 redueix el registre volumètric a **una superfície i un volum**, amb el mètode de càlcul com a traçabilitat del procediment (càlcul geomètric per a mausoleus regulars; extracció del model 3D via MeshLab o CloudCompare per a cavitats irregulars) i un camp de notes per al detall.

La simplificació corregeix un problema conceptual que la parella anterior arrossegava: **volum interior i volum total no significaven el mateix segons la tipologia**. En un mausoleu, la diferència entre tots dos *és* la fàbrica construïda, i per tant una mesura de l'esforç; però en una cambra allotjada dins d'una cavitat, l'«interior» és espai natural que ningú no va excavar i el «total» inclou roca mare, de manera que la resta no designa res comparable. Una mitjana calculada sobre una sola columna amb les dues tipologies barrejades hauria produït una xifra sense significat — el mateix gènere d'error que el tractament dels dominis observacionals evita en altres punts de l'esquema.

El camp únic registra ara **l'espai funerari**, que sí que és comparable entre totes les tipologies i és el que es relaciona amb el nombre mínim d'individus. Els desglossaments que en algun cas concret siguen possibles —volum per cos superposat, superfície de plataforma, volum de fàbrica— es consignen a les notes: són casos puntuals, i un camp propi quedaria buit en la immensa majoria de registres. Si l'anàlisi mostra que són freqüents, aleshores es formalitzaran, seguint el criteri general del projecte de no crear categories sense casos que les sostinguen.

La correlació amb el MNI avalua la relació capacitat-ocupació (H01), i la variació per tipologia i jaciment caracteritza les estratègies d'adaptació al suport (H02).

# **9. Consultes SQL per a l'anàlisi estadística**

La família de 26 consultes conserva les funcions de la versió anterior (distribucions, exportacions R i QGIS, jerarquies, cobertura, maçoneria, geologia) amb tres reconstruccions majors:

- **QRY_02** es reconstrueix sobre T_DECORATIONS (amb QRY_02s com a auxiliar i QRY_02a com a taula de banderes per motiu per al khi-quadrat).
- **QRY_13_AX_Pattern_Export** exporta ara **20 columnes AX_ + 5 columnes SYS_** amb semàntica NA: un component gatejat per un sistema no aplicable s'exporta com a NA, no com a 0, perquè el seu 0 emmagatzemat és de farciment. La definició element-sistema viu en una única font de veritat compartida amb la bateria de validació, de manera que exportació i validació no poden divergir mai.
- **QRY_14** incorpora Chrono_Relation (graf dirigit per a igraph/QGIS).

## **9.1. Estratègia d'anàlisi de patrons constructius (QRY_13)**

L'atomització en el vocabulari A-X converteix cada estructura en un vector d'elements i sistemes, complementat amb la maçoneria, el suport i les coordenades UTM. Les quatre famílies d'anàlisi en R es mantenen: coocurrència (phi, Jaccard) per a H01 i H04; clúster jeràrquic per a les famílies constructives de H05; anàlisi de correspondències amb LP i DW com a punts suplementaris; i autocorrelació espacial (I de Moran) per a l'estructura geogràfica dels patrons (H05, H03). Aquesta estratègia és el fil que connecta OE2 amb el debat sobre transmissió cultural i comunitats de pràctica.

## **9.2. Tractament estadístic dels dominis — advertència operativa**

L'advertència del denominador es manté íntegra: **prendre el total d'estructures com a denominador** infla la mostra amb casos mai avaluats, i com que la conservació difereix entre jaciments, la prova mesuraria preservació diferencial i es llegiria com a pràctica diferencial. Denominador vàlid = presents + absències verificades; la cobertura es reporta sempre junt amb la n.

El domini de cinc valors hi afig una segona obligació, i la v13 l'estén a les capes aplicades: **declarar la binarització**. Abans de qualsevol càlcul, els 9 es converteixen a NA (mai a 0), i després es tria conscientment la variable que respon la pregunta: *construït* (1+2+3), *sobreviu* (1+2) o *intacte* (1). Per a les hipòtesis constructives, normalment la primera. Els NA d'origen de l'exportació (zeros de farciment sota sistemes no aplicables) ja arriben buits i no s'han de reomplir. Les dues estratègies de mostra continuen sent legítimes i declarables: restricció a Facade_Observability = Complete, o distàncies binàries amb tractament de nuls; QRY_15 proporciona el recompte que justifica la decisió.

La v13 hi afig una tercera precaució, derivada del defecte NULL: **un buit no és un 9**. Tots dos són `NA` en el càlcul, però només el 9 és un judici. La proporció de 9 sobre el total avaluat mesura l'observabilitat del corpus i s'ha de reportar; la de buits només indica quant treball queda pendent, i és la regla 17 qui la llista.

> Regla general: un 9 no és ni un 0 ni un 1, sinó una cel·la buida; i un 0 de farciment no és un 0 d'asserció, sinó un NA estructural. Qualsevol operació que els convertisca implícitament en dades reintrodueix el biaix de conservació al resultat.

## **9.3. Nivells constructius, cossos superposats i cossos perduts**

Es manté el criteri fixat: N0/N1/Superior són **categories funcionals, no un sistema de numeració** (mai N2, N3), i un cos pertany a N1 si conté o contenia una obertura d'accés. El recompte viu als camps N_Basal_Bodies i N_Chamber_Bodies, i la matriu es manté plana (una fila per estructura). La novetat v11 és el destí de l'evidència de cossos perduts: l'antic camp Lost_Body_Evidence desapareix i la seua informació esdevé una fila de T_LOST_ELEMENTS amb abast *Body*, que cobreix el vocabulari sencer del cos desaparegut — el cas diagnòstic continuen sent les bandes verticals de pigment sobre la roca nua, el fantasma d'un parament que la pintura ha sobreviscut. La conseqüència analítica és ara més fina: els elements del cos desaparegut ja no s'han de marcar 9 en bloc, sinó que poden marcar-se 3 amb l'evidència que ho justifica, cosa que els reincorpora al recompte de *què es va construir*.

## **9.4. Validació de coherència del registre (QRY_16, 31 regles)**

La bateria passa de deu (v10) a trenta-una regles, emmagatzemades en tres consultes parcials —una unió única de 31 branques supera el límit «query too complex» de JET— i unides a QRY_16_Validation_Check. Continua sent deliberadament **un informe i no una restricció de taula**: una regla dura impediria registrar una absència genuïnament observada en una estructura parcialment col·lapsada. Un resultat buit indica corpus coherent; la regla 17, que llista els camps observacionals encara NULL dels registres heretats, no és una queixa sinó la llista de treball de la revisió manual. El detall de les 31 regles es troba a l'esquema tècnic (secció 6.1); en síntesi cobreixen la coherència component-sistema (1-4), la coherència estat-elements (5-6), els parells presència-detall (7-10, 12-14, 24-26), la coherència classe-contingut (15-16), la bioarqueologia i els materials (18-19), les connexions i l'evidència de pèrdua (20-21), la decoració arquitectònica (22-23), el recompte de cossos basals (27), la pintura rupestre (28-29) i l'alerta de cos perdut (30).

Aquesta darrera mereix una nota, perquè il·lustra la diferència entre una regla i un bloqueig. La combinació «pigment sobre la penya seguint elements arquitectònics» sembla contradictòria i podria semblar candidata a impedir-se al formulari; però és precisament el **cas diagnòstic de cos desaparegut** —bandes verticals de pigment sobre la roca alineades amb els brancals del cos inferior, on la pintura ha sobreviscut al parament que la sostenia—, de manera que bloquejar-la impediria registrar l'evidència que sosté el valor «desaparegut atestat». La regla informa i suggereix obrir el registre d'evidència; no impedeix res.

# **10. Formulari principal F_STRUCTURES (v13)**

El formulari centralitza l'entrada de dades amb 12 pestanyes temàtiques, cinc subformularis vinculats (F_DECORATIONS, F_ROCKART, F_ARCH_FEATURES, F_CONNECTIONS, F_LOST_ELEMENTS) i tots els camps FK i de domini configurats com a ComboBox de dues columnes (valor anglés ocult, etiqueta valenciana visible). La pestanya 11.Sist es reorganitza al voltant dels sistemes: els cinc combos Sys_* dalt de tot i cada grup de components a continuació, perquè aquest és l'ordre real de les decisions — es decideix que hi ha plataforma abans de comptar-ne les mènsules.

## **10.1. Gating condicional de tres nivells**

El formulari activa i desactiva seccions segons el que ja s'ha registrat, de manera que la majoria de les incoherències que la bateria detectava esdevenen impossibles d'introduir:

- **Nivell 1 — classe de registre → pestanyes.** Derivat de Record_Class via ID_Typology: un context funerari natural tanca 2.Arq; una traça estructural i un panell d'art rupestre tanquen a més 6.Bio i 7.Mat; el panell tanca també 11.Sist. Un registre sense classificar ho manté tot obert (tancar pestanyes d'un registre de classe desconeguda amagaria les dades que cal per a classificar-lo). Matís deliberat: una cova conserva 11.Sist obert — DW-S01-EA09 té mènsules i plataforma dins d'una cavitat natural — i 3.Acab resta oberta per a totes les classes, perquè un panell d'art rupestre es defineix pel seu pigment.
- **Nivell 2 — portes → grups.** Cada Sys_* obri el seu grup de components només quan és present o desaparegut atestat; ID_Material_Status tanca els materials; Human_Remains tanca el detall bioarqueològic; Plaster_Present, Pigment_Present i Mortar_Present tanquen els seus detalls; Dec_Present i RockArt_Present governen cadascun el seu subformulari. Amb el domini de cinc valors, els valors 1, 2 i 3 **obrin** —una capa parcial o desapareguda pot tindre color, extensió i motius documentats— i només 0 i 9 tanquen.
- **Nivell 3 — element → detall.** El nombre i el rol de les mènsules només amb E present; el material de plataforma bloquejat amb rol aïllat; el material del dintell només amb Q present; el tipus de coberta només amb X present; el material de la cornisa només amb I present.

**Una correcció de la v13 que il·lustra el límit del gating.** La v12 tractava l'element E (mènsules de fusta) com un component més del sistema de plataforma, de manera que amb el sistema absent el camp es bloquejava i l'emplenat ràpid oferia posar-lo a zero. Això és **activament fals** en un cas real i freqüent: una cambra funerària amb una mènsula lateral que no sosté cap plataforma. E és l'únic element del vocabulari amb existència independent del seu sistema —una mènsula aïllada no ha d'haver portat mai res, i és precisament la raó de ser del camp de rol i de la tipologia MEN—, i la bateria de validació **ja l'exemptava** de la regla corresponent. El gating havia aplicat la pertinença del sistema sense recollir una exempció que la validació ja feia: una incoherència interna del paquet, no un criteri nou.

La lliçó de disseny és que **el gating no pot ser més restrictiu que la validació**: si una regla admet una combinació, el formulari no la pot impedir. En cas contrari es força l'observador a falsejar el registre o a esquivar la interfície, que són les dues coses que el gating existeix per a evitar.

## **10.2. Emplenat ràpid amb confirmació**

Tancar una porta ofereix escriure el valor coherent a tot el grup depenent: components a 0 de farciment (sistema absent o no aplicable) o a 9 (no observable), detalls a buit, materials a 0 quan els vestigis són absents (exactament el que exigeix la regla 19). **Mai no s'escriu res sense un Sí explícit, i mai en navegar entre registres**: l'esdeveniment de navegació només activa i desactiva controls, l'anàleg formal de la regla B de la migració (cap valor emmagatzemat es toca sense intervenció deliberada). Completen el sistema tres avisos en viu: el recordatori d'evidència en marcar un 3 (regla 2), l'avís en roig d'estructura col·lapsada (regla 6) i l'ocultació del substrat «Plaster» quan el revoc és absent verificat (regla 13). Als subformularis, on la desactivació afectaria totes les files alhora, la lògica per fila és validació en desar: F_CONNECTIONS bloqueja una cronologia direccional sobre una junta que no la pot portar (regla 20) i autoassigna la contemporaneïtat de la junta travada; F_LOST_ELEMENTS exigeix el codi d'element quan l'abast és Element (regla 21).

El principi de disseny resultant: **el formulari preveu; la bateria detecta.** Cap dels dos substitueix l'altre — el gating es pot esquivar (taules obertes a mà, importacions, Centre de confiança desactivat) i algunes regles són intrínsecament de corpus.

## **10.3. Camps de notes per pestanya**

La v14 incorpora sis camps de text lliure, un a cada pestanya que conté **judicis interpretatius que el domini no pot expressar**: morfologia i fases, acabats, conservació, bioarqueologia, materials culturals i sistemes constructius. No en tenen totes: les pestanyes mètrica i de documentació ja disposen de camps de notes específics, i les taules filles porten els seus propis.

La decisió de col·locar-los **a la pestanya que els correspon i no en una pestanya de repàs** respon a com s'escriu realment una nota: es redacta quan apareix el cas, i el cas apareix emplenant eixa pestanya. Una pestanya única de comentaris hauria estat més còmoda de llegir però hauria recollit menys notes. La vista de conjunt existeix igualment com a consulta de només lectura.

Val la pena consignar l'advertència que acompanya aquesta decisió, perquè afecta la qualitat analítica del registre: **les notes són el forat negre de la base de dades**. Tot el que s'hi escriu deixa de ser analitzable, i la temptació d'anotar en lloc de decidir el valor del camp és considerable. El criteri correctiu és el mateix que s'ha aplicat en altres punts del disseny: **si el mateix tipus d'observació apareix repetidament en text lliure, això és el senyal que cal un camp**, no que calga més text.

## **10.4. Captions DAO per a la vista full de dades**

Es manté el mecanisme dual: etiquetes adjuntes als controls dels subformularis (capçaleres de columna en valencià) i captions DAO a nivell de TableDef com a reforç per a l'obertura directa de les taules.

# **11. Implementació tècnica**

## **11.1. Scripts VBA i rutes d'actualització**

El paquet v13 comprén tres scripts, amb una única ruta d'ús:

- **chachapoya_DB_v14.bas → BuildDB()**, sobre una base en blanc: crea les 22 taules, fixa els valors per defecte (0 als 20 camps d'element, buit a la resta; sistemes a «Absent»), pobla els lookups, estableix les 25 relacions i genera les 28 consultes.
- **chachapoya_Form_v14_val.bas → BuildForm()**, a continuació: crea els sis subformularis, F_STRUCTURES amb 12 pestanyes, els combos de dues columnes, les captions DAO, el mòdul de gating i les validacions de subformulari.
- **diagnostic_v13.bas → RunDiagnostic()**, de només lectura: comprovacions de coherència i de variància sobre el corpus, sense modificar cap dada ni cap esquema.

**Sobre la decisió de no migrar.** Les iteracions anteriors van transformar la base preservant els registres, amb un protocol explícit que val la pena consignar perquè és citable com a bona pràctica: ordre estricte *renomenar → afegir → poblar → eliminar*, de manera que cap camp desapareix abans que el seu substitut estiga poblat i verificat; cap valor existent es toca mai, i els valors per defecte només afecten files futures; els esborrats finals doblement barrats per un commutador explícit i per precondicions comprovades en viu contra les dades; i recomptes d'afectació a cada bloc, cosa que converteix el registre d'execució en pista d'auditoria.

Per a la v13, en canvi, s'ha optat per **reintroduir els registres sobre una base construïda de zero**. La raó no és de comoditat sinó epistèmica: el diagnòstic previ va mostrar que una part substancial dels valors del corpus no eren judicis de l'investigador sinó **herència de valors per defecte de versions anteriors** —una vintena de «no observable» en un sol camp, procedents del defecte de la v10—, i eixos valors, en no ser buits, quedaven fora de la llista de treball i entraven als recomptes com si fossen observacions. Cap migració els podia distingir dels judicis reals. Reintroduint, cada valor del corpus passa a ser deliberat, que és la condició prèvia per a publicar qualsevol percentatge.

## **11.2. Restriccions tècniques de VBA/Access**

Es mantenen les restriccions documentades: cap continuació de línia (patró sql = sql & "..."), codificació cp1252 sense caràcters no-ASCII al codi, paraules reservades evitades, patró de creació i renomenament de formularis, N-1 parèntesis per a N taules en JOIN, dbRelationDontEnforceIntegrity (valor 2) per a l'autoreferenciant i les dues de T_CONNECTIONS, procediments auxiliars Private, DROP previ per a objectes regenerables, i combos configurats per ControlSource. S'hi afigen les apreses en les iteracions v11-v13: JET no pot convertir TEXT a BYTE in situ (el desdoblament del dintell exigeix renomenar per DAO, crear el camp nou i poblar-lo per UPDATE); un renomenament DAO no reescriu el SQL de les consultes existents (les consultes afectades queden trencades fins que es reconstrueixen, i l'ordre dels blocs ho preveu); una UNION de més de ~30 branques supera el límit «query too complex» (la bateria s'emmagatzema en tres parts); la injecció de mòduls de formulari exigeix l'accés al model d'objectes VBA al Centre de confiança (si està desactivat, tot es construeix igualment i l'script avisa); i amb la columna lligada oculta, Access imposa LimitToList (la via d'escapament és el valor «Other (see Notes)»).

## **11.3. Exportació per a l'anàlisi estadística**

Sense canvis de procediment: QRY_05 (exportació plana) i QRY_13 (matriu AX+SYS) via Dades externes → Exportar, amb els camps categòrics en format llegible gràcies als JOIN amb els lookups. La conversió 9 → NA i la declaració de la binarització (secció 9.2) són pas previ obligat en R.

# **12. Limitacions i perspectives**

La limitació principal continua sent la cobertura desigual del corpus (La Petaca Sector Sud amb ≥96 estructures identificades, no totes amb fitxa completa). Els camps de base documental i observabilitat, la regla 17 (llista de treball dels NULL) i QRY_15 converteixen aquesta limitació en una variable mesurada en lloc d'un buit silenciós. La reintroducció dels 35 registres sobre l'esquema v13 és la tasca oberta immediata; la regla 17 en mesura el progrés camp a camp. Alguns registres exigiran, a més, revisió sobre el model 3D: els que portaven el valor de suport ara desdoblat, i aquells en què la parella de suports apareix en ordres oposats i cal determinar amb el criteri nou quina forma rep la càrrega i quina estabilitza.

T_ARCH_FEATURES i L_DEC_TYPE mantenen l'escalabilitat de l'esquema sense modificar-ne l'estructura. El vocabulari A-X, ara formalitzat com a taula amb la distinció explícita element/sistema, és el primer intent de normalització terminològica per a les estructures funeràries Chachapoya en la bibliografia, i la seua aplicació creuada a La Petaca i Diablo Wasi és una de les aportacions originals del TFM.

De cara a publicacions futures, es manté la perspectiva de migració a GeoPackage (.gpkg) per a la integració nativa amb QGIS, amb Access com a interfície d'entrada.

# **13. Referències**

Epstein, L.; Toyne, J.M. (2016). When Space Is Limited: A Spatial Exploration of Pre-Hispanic Chachapoya Mortuary and Ritual Microlandscape. A: Osterholtz, A.J. (ed.), Theoretical Approaches to Analysis and Interpretation of Commingled Human Remains. Springer, Switzerland, pp. 97-124.

Ribera-Torró, E. (2023). Chacha XR. Una experiència immersiva de no-ficció per a l'arqueologia Chachapoya. Treball de Fi de Màster, Màster Universitari en Arts Visuals i Multimèdia, Universitat Politècnica de València.

Ribera-Torró, E.; Toyne, J.M.; Del Aguila, R.; Ribera, J.A.; Galexner, J.; Anzellini, A.; Pans, M. (2026). Extended Reality on Chachapoya Cliffside Necropolises: From Digital Documentation to Public Engagement. Open Archaeology, 12(1). DOI 10.1515/opar-2025-0071.

Toyne, J.M.; Anzellini, A. (2017). Sociedad, identidad y variedad en los mausoleos de La Petaca, Chachapoyas. Boletín de Arqueología PUCP, 23, 231-257.

Toyne, J.M.; Anzellini, A.; Epstein Miculas, L.; Mejías Pitti, I.; Puig Castell, J.; Guinot Castelló, S. (2018). Going Vertical: Using Vertical Progression Techniques to Explore a Cliff Necropolis in Late Precolumbian Chachapoyas, Peru. Advances in Archaeological Practice, 6. DOI 10.1017/aap.2018.31.
