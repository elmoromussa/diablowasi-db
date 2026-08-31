**Universitat d'Alacant**

Màster en Arqueologia Professional i Gestió Integral del Patrimoni

**DISSENY D'UNA BASE DE DADES ARQUEOLÒGICA**

per a l'estudi de les necròpolis de penya-segat de La Petaca i Diablo Wasi

*(Leymebamba, Amazonas, Perú, s. IX-XVI)*

Autor: Esteve Ribera Torró

Directors: Dr. Ignasi Grau Mira (UA) i Dra. J. Marla Toyne (UCF)

Curs acadèmic 2024-2025

*Versió 8 del document — actualitzada segons la BD v17 (agost 2026). Consolida les iteracions v10→v11, v12, v13, v14, v16 i v17.*

*Nota de la versió 23 del paquet: aquest text consolida el disseny fins a la v17; les decisions v18→v23 (superfície de plataforma i la seua reserva al sistema volat, format del coronament, declaracions de fàbrica i de disponibilitat mètrica, bandes de contorn, índex únic de codi, datació només per C14, regles 47–67 amb la 55 revisada, i dues reversions explícites documentades: datació només per C14 i EA emmagatzemat) estan documentades als deltes corresponents (`DELTA_v17a_v18` … `DELTA_v22_v23`) i a l'esquema tècnic `esquema_bbdd_estructures_v23.md`, que és la referència vigent. La justificació metodològica d'ací — dominis, NULL significatiu, gating, unitat de registre, chaîne opératoire — continua sent vàlida i és la que aquelles decisions apliquen.*

# **Resum**

El present document descriu el disseny, la justificació i la implementació de la base de dades relacional destinada a la documentació sistemàtica i l'anàlisi estadística de les estructures funeràries de les necròpolis de penya-segat de La Petaca i Diablo Wasi (Leymebamba, Departament d'Amazonas, Perú, s. IX-XVI d.n.e.). La base de dades constitueix l'eix vertebrador de la metodologia del Treball de Fi de Màster, en tant que permet centralitzar les variables arqueològiques, establir relacions jeràrquiques entre elements, exportar dades per a l'anàlisi estadística i vincular el registre arqueològic amb el Sistema d'Informació Geogràfica (SIG) implementat en QGIS.

L'esquema segueix els principis de la tercera forma normal (3FN) i inclou 22 taules, una taula principal (T_STRUCTURES) amb 136 camps, 25 relacions, la família de 33 consultes SQL i un formulari d'entrada de dades amb 12 pestanyes, set subformularis i interfície en valencià. Un dels avanços metodològics centrals és la normalització d'un vocabulari arquitectònic bilingüe (valencià/anglés) de 24 entrades (A-X), derivat de l'anàlisi fotogramètrica 3D de les estructures de Diablo Wasi, formalitzat des de la v11 com a taula lookup pròpia (L_ELEMENTS).

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

# **3. Vocabulari arquitectònic normalitzat (A-Z)**

Un dels avanços metodològics centrals de la base de dades és la normalització bilingüe (valencià/anglés) d'un vocabulari de 25 entrades arquitectòniques (A-Z; la lletra Y queda reservada per a la massa adossada o banqueta, pendent de casos documentats), derivat de l'anàlisi sistemàtica dels models fotogramètrics 3D de Diablo Wasi. Des de la v11 el vocabulari és una taula lookup pròpia (L_ELEMENTS: codi, noms, nivell, sistema, camp associat, descripció), cosa que el fa referenciable amb integritat —T_LOST_ELEMENTS hi apunta— i exportable com a metadada del projecte.

## **Criteris d'ordenació**

La codificació segueix tres criteris jerarquitzats: (1) NIVELL CONSTRUCTIU: de N0 (basament) cap a N1 (cos principal) i zona superior; (2) ALÇADA dins de cada nivell: de baix cap a dalt; (3) FUNCIÓ: primer l'element estructural portant, després el tractament decoratiu.

## **Elements i sistemes: la correcció conceptual de la v11**

Tres entrades del vocabulari són sistemes compositius, és a dir, el resultat de la combinació obligatòria d'altres elements: E+F+G → H (plataforma), N+O+Q → P (portal) i S+T → U (ràfec, sempre lapidi a T i U). La v10 els registrava com si foren elements més, amb un camp binari cadascun. La v11 corregeix la categoria: **H, P i U no tenen existència material pròpia** —són un judici sobre la combinació dels seus components— i per això esdevenen camps d'estat (Sys_Platform, Sys_Portal, Sys_Eave) amb un domini de sis valors: *Present complete / Present partial / Attested lost / Absent / Not applicable / Not observable*.

Un cas límit resolt en la v14 delimita bé l'abast del concepte de sistema. El marc reculat de façana estava registrat com a qualificador del sistema portal, però el recul afecta el **pla del parament sencer** i el portal hi queda inscrit, de manera que la dependència era inversa i el gating el bloquejava precisament allà on hi ha recul sense portal. Es va valorar promoure la façana a sistema propi i **es va descartar**: els cinc sistemes existents són combinacions de components identificables o camps d'agrupació per al control d'entrada, mentre que una façana no és una combinació de res — és el pla on tota la resta passa, i **no podria valdre mai «absent»**. Un sistema que no pot ser absent no fa el que fan els sistemes, que és precisament permetre distingir l'absència verificada de la inaplicabilitat. El camp passa, doncs, al bloc de la façana com el que és: un qualificador del pla.

La correcció té dues conseqüències analítiques. La primera: la distinció entre «no aplicable» i «absent verificat» —la mateixa lògica del 0 vs 9 elevada al nivell del sistema— fa que els percentatges de presència siguen calculables sobre el denominador correcte. La segona: en l'anàlisi de coocurrència, H correlacionava trivialment amb E, F i G perquè n'era la suma; amb la matriu de QRY_13 reorganitzada en columnes d'element i de sistema (21 i 6 des de la v18), el doble comptatge desapareix abans d'arribar a R.

S'hi afigen dos camps d'agrupació sense lletra pròpia, Sys_Base i Sys_Chamber, amb domini de quatre valors, que serveixen el gating del formulari i el mateix problema del denominador a escala de conjunt (des de la v18, Sys_Base agrupa A, B i C: el muret transversal D en va eixir quan el corpus el va mostrar desvinculat de la massa basal — sobre plataformes volades, com a suport adossat, o aïllat). El coronament (R) i, des de la v18, el muret transversal (D) no pertanyen deliberadament a cap sistema: veure un d'aquests elements sense poder resoldre cap sistema no és una contradicció sinó observabilitat parcial, la norma en un penya-segat. La cornisa intercòs (I) va rebre en v17 el seu propi conjunt (Sys_Interface), l'únic que passa el test que va rebutjar un sistema de façana: una interfície pot ser absent.

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

El corpus conté entrades que no són estructures funeràries construïdes: cavitats naturals, mènsules aïllades (traces estructurals de la xarxa de circulació), panells d'art rupestre. La v10 les tipificava però no les classificava, i qualsevol percentatge calculat sobre el total les incloïa indegudament. Record_Class —una classe per tipologia, mai per registre— resol el problema en el punt correcte de la normalització: *Built funerary structure* (EA-MAU, EA-CAM, EA-TER, EA-PLA-V, Unclassifiable), *Natural funerary context* (NIX, CAV), *Structural trace* (MEN), *Rock art panel* (PR) i *Pending classification* (estat de treball, exclòs de tota consulta analítica). La classe filtra les anàlisis i governa el primer nivell del gating del formulari (secció 10.2). El desdoblament de l'antic «ND» distingeix un veredicte epistèmic (Unclassifiable: evidència insuficient per a classificar una estructura construïda, com EA11, un perímetre de pigment sense construcció conservada) d'un simple estat del flux de treball (Not yet classified).

# **5. Variables arqueològiques de T_STRUCTURES (145 camps)**

La taula T_STRUCTURES concentra 145 camps (v20: retirat el format de cornisa, incorporat el format del coronament). Els dominis observacionals són tres:

- **21 camps amb domini de cinc valors** (0/1/2/3/9): els elements A-X més la superfície de plataforma Z (v18), i només ells (secció 5.2).
- **31 camps amb domini 0/1/9**: atributs de tècnica, capes contínues, processos, bioarqueologia, materials culturals, qualificadors i judicis agregats.
- **6 camps de sistema, domini textual** de sis valors (Sys_Platform/Portal/Eave) o quatre (Sys_Base/Chamber/Interface).

El **valor per defecte** és una decisió que ha evolucionat en tres passos, i val la pena seguir-la perquè il·lustra com el disseny s'ha anat ajustant al que el registre realment necessita. La v10 usava 9 (no observable) per a tots els camps booleans convertits; la v11 el va canviar a 0 perquè el gating dels sistemes constructius **exigeix** un zero de farciment als components d'un sistema no aplicable; la v13 conserva eixe 0 **només allà on la regla el necessita** —els camps d'element governats per un sistema, 21 des de la v18— i el substitueix per **NULL a la resta**. (La v18 hi afig una segona excepció puntual amb la mateixa lògica, i la v19 una tercera del mateix tipus: `Sill_Coincides_Cornice` i `Jamb_Fabric_Reveal`, els dos **qualificadors de compartició d'element**, la inaplicabilitat dels quals és derivable, porten el 0 de farciment i no NULL — vegeu 5.11 i 5.12.)

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

## **5.4bis. Un camp que servia dues meitats sense compartir-ne el referent (v17)**

La v16 va afegir a la taula de decoració un camp de posició relativa —esquerra, dreta, damunt, davall— pensat per a servir alhora les files arquitectòniques i les rupestres. L'economia era aparent, i el motiu que fallara il·lustra un tipus d'error de disseny que no es detecta llegint l'esquema sinó **intentant omplir-lo**: la pregunta «esquerra respecte de què?» no tenia resposta.

No la tenia perquè les dues meitats de la taula no comparteixen referent. A les files arquitectòniques, *esquerra* designa **quina instància d'un parell** porta el motiu —quin dels dos brancals, quin dels dos flancs—: una relació **identificacional**, interna a l'estructura, que tria un element entre un conjunt. A les files rupestres designa **on està la pintura** en relació amb el volum construït: una relació **topològica** entre dos objectes separats. Que totes dues es diguen amb les mateixes paraules és una coincidència del llenguatge natural, no una propietat compartida; i un camp que codifica dues relacions lògiques distintes no codifica cap de les dues.

Les conseqüències eren comprovables sobre el corpus. Els valors verticals no tenien referent possible a la meitat arquitectònica, on l'eix vertical el diuen ja la posició i el número de cos. El valor «totes dues» significava coses diferents a cada meitat i la documentació només en definia una. I el pla de referència declarat —el pla d'accés— era el que no tocava per a l'art rupestre, que es veu des de la vall i no des de l'entrada: en tota estructura amb l'accés desviat de la façana, la convenció **invertia** esquerra i dreta, precisament en el cas d'interés.

La solució adoptada desdobla i restringeix, com ja s'havia fet amb el format i el treball de la pedra. La meitat arquitectònica **no rep reemplaçament**: cap de les cinquanta-sis files del corpus usava la lateralitat sobre posicions arquitectòniques, i crear un camp sense ocurrències hauria contradit el mateix criteri que va limitar les posicions rupestres a quatre. La meitat rupestre rep **quatre camps de tram** —esquerra, damunt, dreta, davall—, cadascun amb el domini observacional de tres valors.

Quatre camps i no una llista tancada de cobertures, i la raó no és combinatòria sinó epistèmica. Un valor únic no pot distingir *la banda no cobria la base* de *el tram de base no és observable*, i eixa distinció és exactament la que separa una banda en U invertida d'un anell tancat mal conservat. Al corpus hi ha casos on el tram superior queda tapat per la visera rocosa i pel cos que hi ha damunt: marcar-lo absent afirmaria una cosa que ningú no ha pogut mirar. Les etiquetes de lectura —U invertida, envoltant, flanquejant— **no s'emmagatzemen**: es deriven per consulta, de manera que una U amb la base no observable apareix com a *candidata* a envoltant en lloc de quedar afirmada com a U. És el mateix patró de judici agregat que separa la presència de decoració de les files que la detallen.

Una quarta propietat, que no és ni posició ni tipus, va resultar ser la més carregada arqueològicament: **la geometria del traç**. Una banda que fa angles rectes dibuixa un rectangle i per tant **afirma un referent construït**; una banda corba no afirma res i pot estar resseguint un accident de la roca. Aquesta distinció és l'única evidència disponible per a decidir si un cas de pintura perimetral sense fàbrica conservada és el rastre d'una estructura o un motiu rupestre, i per tant decideix en quina meitat del corpus va el registre. Registra l'observació; l'afirmació *ací hi havia una estructura* viu a la taula d'elements desapareguts, la mateixa separació entre el que es veu i el que s'infereix que ja governava l'associació de verticalitat.

**Un detall de procediment que val la pena consignar.** El camp de geometria del traç existeix perquè una fila del corpus portava la paraula «geomètrica-ortogonal» escrita a mà al camp de notes. La consulta de repàs de notes es va crear en la v14 amb l'argument explícit que *quan la mateixa observació apareix repetidament en text lliure, això és el senyal que hauria de ser un camp*. Ací es va complir, i el mecanisme va funcionar: la solució venia de les dades pròpies i no d'una revisió abstracta de l'esquema.

## **5.5. Format i treball de la pedra: la selecció com a primer gest (v16)**

`Masonry_Type` registrava com s'apila la pedra i `Masonry_Quality` un judici agregat d'execució. Cap dels dos registrava **quina pedra és**, que és el primer gest de la cadena operativa i, per tant, matèria directa de H04 i el descriptor amb més capacitat de discriminar tradicions entre els dos jaciments (H05).

El vocabulari que la revisió de camp va proposar inicialment —pedres irregulars, laminars, blocs semiescairats, carreus, lloses de gran format— **barrejava tres eixos**: morfologia natural, grau de treball i dimensió. «Irregular» i «laminar» no s'oposen a «carreu», perquè una pedra laminar pot estar escairada o no; en una llista única, un mur de lloses laminars ben escairades no té valor possible i el registre es decideix per hàbit de l'observador, que és exactament el que trenca la replicabilitat entre observadors. La solució és desdoblar en dos camps ortogonals, i els termes proposats es reconstrueixen tots com a combinacions.

El primer (`Stone_Format`) descriu el producte de la fractura natural de la roca amb un test deliberadament **funcional** —quantes peces fan una filada— perquè es llig directament a la façana i no exigeix mesurar res. El segon (`Stone_Working`) és **ordinal** i registra evidència positiva de treball, de manera que el valor baix no afirma «ningú no la va tocar» sinó «no hi ha evidència que la tocaren»: una precisió que importa perquè un gres estratificat fractura en cares indistingibles del carejat, i confondre-les inflaria l'aparent inversió de treball.

Cap dels dos porta valor «no observable», i el criteri val per a tot camp de domini textual de l'esquema: **el 9 existeix per a protegir el 0**, i aquests camps no en tenen —un parament sempre està fet d'alguna pedra amb algun grau de treball. Els dos motius pels quals la decisió pot quedar oberta, «no s'hi veia» i «no era decidible», comparteixen el valor `ND` però no es col·lapsen: els separa l'encreuament amb `Doc_Basis` i `Facade_Observability`, que ja registren la visibilitat. És el mateix raonament que va deixar `Mortar_Present` sense promocionar a cinc valors.

## **5.6. El criteri de diferenciació i el zero com a resultat (v16)**

La pregunta que va originar aquest criteri era estreta —quan es diu que un portal té llindar i brancals?— i la resposta va acabar fixant el criteri de presència de **tot** el vocabulari A–X:

> Un element A–X és present quan hi ha un component **físicament diferenciat** del parament contigu, per format de pedra, dimensió, material o tractament.

Una obertura resolta deixant simplement un buit al mur té, doncs, llindar, brancals i dintell a zero, **i el sistema portal present igualment**, perquè l'obertura hi és. No és una contradicció sinó la divisió del treball entre sistema i components: el sistema registra *que existeix obertura d'accés*, els components *si cada posició es va resoldre amb un element distint*. La coherència amb el disseny previ és exacta — la regla 11 ja feia incompatible tindre cossos de cambra amb el sistema portal absent, i les dimensions de l'obertura deliberadament no estan gatejades pel sistema.

L'interés arqueològic està en els zeros que el criteri produeix. Que a Diablo Wasi l'obertura es resolga sovint sense llindar ni brancals diferenciats és una afirmació sobre la **inversió de treball** i sobre la seqüència operativa, i per tant material de H01 i H04. Amb el criteri contrari —«si hi ha obertura, hi ha brancals»— la variable desapareixeria i tots els portals serien iguals. El criteri no és nou de fons: el basament ja el portava implícitament en la seua definició com a «element constructiu diferenciat»; el que fa la v16 és generalitzar-lo i escriure'l.

La conseqüència operativa és que la transferència a v16 porta aquests tres camps a NULL en tot el corpus. Un valor emmagatzemat sota el criteri antic no és un judici sota el nou, i tractar-lo com si ho fos reintroduiria per la porta del darrere el problema que la reintroducció manual de la v13 va resoldre.

## **5.7. La façana com a pla exposat, i el que se'n deriva (v16)**

Fins a la v15 el mot «façana» feia dues faenes alhora: designava el **pla exposat** —el que mira a la vall, el que mesuren l'orientació i la visibilitat, el que sosté tota H06— i el **pla compositiu**, el que porta l'obertura i organitza els elements que la flanquegen. Normalment coincideixen i el mot funciona sense friccions.

Una estructura documentada els separa: mausoleu de dos cossos basals i una cambra, amb l'obertura al mur estret perpendicular al farallò i la decoració principal al mur llarg paral·lel. Calia decidir quin dels dos sentits es queda amb el mot. Fixar la façana per l'obertura hauria deixat l'orientació de façana sense apuntar a la vall i la visibilitat mesurant un pla que no es veu: **tota una hipòtesi degradada per a salvar la definició d'un element**. Fixar-la pel pla exposat costa una línia de definició —el flanc de façana deixa de dependre de l'obertura— i no toca res més.

La divergència, però, no s'ha d'absorbir sinó **registrar**: que l'accés no estiga al pla exposat diu que la circulació mana sobre l'exhibició, que s'entra per on es camina i s'exhibeix cap a on es mira. Això és H06 i H01 alhora, i fins ara no quedava enlloc. Dos camps nous ho recullen, i no es dupliquen: el **pla d'accés** és relacional i intrínsec, comparable entre estructures amb orientacions absolutes distintes; l'**orientació del portal** és absoluta i entra a QGIS per a creuar-se amb la topografia i les rutes de circulació. Cap dels dos es dedueix de l'altre sense conéixer la geometria del cas, i la seua **divergència** passa a ser una variable derivada calculable: es podrà preguntar si les estructures amb accés desviat comparteixen tipologia, sector, forma de suport o amplària de repisa.

En la v16 el pla d'accés feia a més una segona faena: declarava el pla de referència de la lateralitat de la decoració. **La v17 li'n lleva**, i el motiu està a 5.4bis: eixa lateralitat es llegia des del pla equivocat. Els camps de tram que la substitueixen prenen com a referència el **pla exposat**, sempre i sense excepció, que és el que la definició de façana de la v16 havia fixat i el que fa que la convenció es mantinga estable quan l'accés divergeix. El pla d'accés conserva només la primera faena, que era la que justificava crear-lo.

La mateixa definició de façana com a pla exposat resol també, per derivació directa, una pregunta de gating que la v17 es va plantejar: què fer amb els camps de façana i paisatge en estructures que només tenen cossos basals i cap cambra. El pla d'accés i l'orientació del portal deixen de tindre sentit —sense cambra no hi ha interior on entrar—, però l'orientació de façana i la visibilitat de la vall **s'han de mantindre obertes**, perquè una massa basal té pla exposat i mira cap a algun lloc. Desactivar-les eliminaria del càlcul precisament les estructures sense cambra, que són el **grup de control natural** de la hipòtesi sobre exhibició: la pregunta és si s'orienten igual o distint que les que en tenen. Una decisió de definició presa per raons analítiques continua produint conseqüències correctes dues versions després, cosa que és, ella mateixa, un indici que la definició era la bona.

## **5.8. Presència i tipus: el patró general**

Diversos parells de camps segueixen el mateix patró presència + tipus, amb la regla que el tipus és inaplicable —NULL, no «ND»— quan la presència és 0 o 9: Lintel + Lintel_Material (desdoblat en v11: el camp antic era un lookup de material que confonia presència amb tipus), Chamber_Roof + Chamber_Roof_Type, Plaster_Present + color i extensió, Pigment_Present + substrat, color i extensió, Mortar_Present + Mortar_Type. Les regles 10, 12-14 i 24-26 vigilen el patró, i el gating v12 el fa complir en entrada.

## **5.9. Registrar la divergència en lloc de reestructurar el model (v17)**

La qüestió ajornada des de la iteració anterior era la **cardinalitat**: descompondre cada estructura en els seus cossos constructius, amb el vocabulari A–X penjant de cadascun, en lloc d'un sol joc de camps per registre. Dues línies d'evidència independents hi apuntaven —la descripció dels sistemes per subnivells i, sobretot, el fet que amb un sol joc de camps una estructura amb el cos basal de blocs irregulars i la cambra de blocs tabulars només es puga descriure com a «mixta», que **destrueix exactament la informació d'interés**: que hi ha dues fàbriques i per tant possiblement dues mans, dues fases o dues intencions.

Que dues línies independents demanaren la mateixa taula era el senyal més fort que s'havia tingut que la descomposició fora la resposta correcta i no una sobreenginyeria. Tres consideracions la van ajornar una segona vegada, i convé escriure-les perquè la decisió es puga revisar amb els mateixos arguments.

**La primera és de suficiència de dades.** La xifra que ha de decidir-ho —quants registres tenen cossos amb fàbrica realment divergent— no existeix ni pot existir a la base, perquè cap camp la recull; i els camps de format i treball de la pedra, que són els que la produirien, estaven emplenats en quatre registres de trenta-sis. Dissenyar l'estructura d'unes dades encara no recollides és la situació on més fàcil és construir una cosa elegant que després no encaixa amb el que es troba.

**La segona és de disseny d'anàlisi, i és la decisiva.** Si els cossos passen a ser files pròpies, canvia la **unitat d'anàlisi** de les proves estadístiques: la matriu de patrons constructius passaria d'unes trenta-sis files a unes seixanta. Això sembla un guany de mostra i no ho és, perquè **dos cossos de la mateixa estructura no són observacions independents** — els va fer la mateixa gent, el mateix dia, amb la mateixa pedra. Tractar-los com a casos separats en una anàlisi de conglomerats o de correspondències infla la mostra amb repeticions i produeix agrupacions que reflecteixen quants cossos té cada estructura, no quines pràctiques constructives s'assemblen. La descomposició continua sent vàlida com a eina descriptiva, però **no resol l'anàlisi**, i part de l'atractiu de fer-la ara era eixe.

**La tercera és d'oportunitat**, i no s'amaga: és el canvi més gran que el projecte hauria escomés —esquema, formulari, matriu, exportació i criteri d'unitat de registre alhora— i arriba en el moment de tancar el treball.

La sortida adoptada no és tornar a ajornar sense més: és **instal·lar la mesura que falta**. Dos camps registren si la fàbrica és uniforme o divergent i, quan ho és, si la divergència es dona **entre cossos o dins d'un mateix cos**. Aquesta segona distinció és tot el sentit de l'operació, perquè si la divergència resulta estar sobretot dins d'un cos, la descomposició per cossos no resoldria res —el mateix «mixt» reapareixeria un nivell més avall—; i un tercer camp diu si la divergència entre cossos és base contra cambra, que probablement és funcional, o cambra contra cambra, que és on hi ha dues mans o dues intencions.

L'efecte és doble. Es deixa de destruir informació immediatament, perquè el fenomen passa a ser una variable en lloc d'una impressió; i **la segona passada produirà la xifra decisiva com a subproducte**, sense cap treball addicional, quan s'òmpliga el format i el treball de la pedra dels trenta-sis registres. La decisió sobre la descomposició es prendrà llavors amb un nombre i no amb una intuïció, que és el que ara mateix no era possible.

Un darrer punt de frontera, perquè el camp nou té un veí perillós: un canvi de fàbrica a mitja alçària és sovint **precisament l'evidència d'una fase constructiva**. La separació s'estableix amb la mateixa lògica que ja governava l'associació de verticalitat: *la divergència de fàbrica és una observació —la pedra canvia, i això es veu—; la fase constructiva és una interpretació —hi va haver dos moments, i això s'argumenta*. Pot haver-hi canvi de fàbrica sense cap fase afirmada, per un canvi de proveïment o per dos operaris el mateix dia; quan sí que se sosté la fase, l'argument va al camp d'evidència, i una regla nova detecta les fases declarades que no en tenen cap.

## **5.10. Quan la inspecció de les dades corregeix el disseny (v17)**

Aquesta iteració és la primera del projecte en què el disseny es va contrastar **directament contra la còpia de treball** —trenta-sis registres, cinquanta-sis files de decoració— i no només contra la documentació. El resultat va modificar substancialment el que s'havia previst, i el patró mereix consignar-se com a procediment.

Tres punts que s'havien plantejat com a **preguntes obertes de disseny** eren, en realitat, **incoherències ja presents a les dades**. El cas més il·lustratiu: la descripció de l'entrada tipològica «no classificable» nomenava explícitament una estructura concreta com a exemple d'estructura construïda de la qual només resta el perímetre pintat; eixa mateixa estructura estava desada com a registre d'art rupestre, i portava a més una fila de decoració en una posició —«perimetral»— que és lògicament impossible en un registre d'art rupestre, perquè significa *ressegueix el contorn de l'estructura* i un panell aïllat no en té cap. La decisió, doncs, estava presa i escrita al vocabulari; el que faltava era una regla que detectara la contradicció, que la v17 afig en les dues direccions.

De manera semblant, la duplicació de connexions no era un risc a previndre sinó un fet: les dues úniques files de connexió del corpus eren **la mateixa connexió** entrada des de cada extrem. I la causa no era descuit de l'usuari sinó de la interfície, que només mostrava les connexions registrades des d'un dels dos costats: amb eixe disseny, duplicar era el comportament per defecte i no l'excepció.

Dos camps que s'havien discutit en abstracte van quedar decidits per l'ús real sense necessitat d'argumentar-los. El de posició relativa el portaven set files de cinquanta-sis, **totes rupestres i cap arquitectònica**, cosa que va resoldre sola la pregunta de si calia conservar la lateralitat arquitectònica. Un tipus decoratiu que semblava solapar-se amb un element del vocabulari tenia **zero usos**, de manera que retirar-lo no costava cap migració. I un camp tècnic present en vint-i-sis registres i absent en cap va quedar identificat com a variable sense variància, incapaç d'aparéixer en cap resultat: es conserva com a descriptor, però amb la conseqüència escrita.

La lliçó de procediment és simple i probablement general: **una revisió de disseny feta només sobre l'esquema veu el que el model permet dir, no el que s'hi ha dit**. Les dues coses divergeixen ràpidament, i la divergència és informativa en totes dues direccions — camps que ningú no usa i camps que s'usen per a dir una cosa distinta de la que declaren.

## **5.11. El primer delta nascut de l'entrada de dades (v18)**

La iteració v17a→v18 és la primera especificada íntegrament des de l'experiència d'entrada: trenta-cinc estructures entrades amb el formulari en ús real, i un delta on cada punt cita la nota de camp que el va motivar. Tres decisions tenen valor metodològic més enllà del canvi concret.

**La compartició d'element com a categoria de registre.** El corpus va presentar un cas que el gradient de cinc valors no pot dir: un portal sense llindar diferenciat **perquè la cornisa intercòs fa eixa funció**. Registrar N = 1 mentiria (no hi ha peça diferenciada); registrar N = 0 a seques perdria com es va resoldre la posició. La solució és un qualificador de compartició (`Sill_Coincides_Cornice`): N diu la veritat sobre la diferenciació, i el qualificador diu que la funció la cobreix una altra peça del vocabulari. El patró és generalitzable —qualsevol element pot resoldre's per delegació en un altre— i s'ha implementat amb la mateixa disciplina del gating: el camp només viu dins de la seua finestra d'aplicabilitat (N = 0 amb cornisa amb entitat), fora porta el zero de farciment, i una regla vigila la finestra.

**La direcció d'una relació pertany a la identitat, no a la posició.** Les connexions entre estructures van necessitar dues iteracions per a arribar a la forma estable. La v17 va normalitzar la parella (ID menor sempre al primer extrem, índex únic contra la duplicació) però va mantenir la direcció cronològica com a valor posicional («A és anterior»), cosa que obligava el formulari a girar la cronologia en normalitzar i la vista entrant a invertir-la en mostrar. La v18 substitueix el valor posicional per una clau forana que **nomena l'estructura anterior** (`ID_Earlier`): el gir de normalització deixa de tocar res, la vista mostra el fet emmagatzemat, i el graf de H03 esdevé genuïnament dirigit per identitat. La seqüència il·lustra un principi de modelatge: quan una dada obliga a mantenir codi de compensació en dos llocs (girar en escriure, invertir en llegir), la dada està mal situada.

**El vocabulari creix per paral·lelisme, no per acumulació.** L'element Z (superfície de plataforma) no entra perquè aparega un tret nou sinó perquè el vocabulari ja contenia el seu paral·lel exacte —la superfície del ràfec (T)— i l'asimetria era un buit del model, no del registre: la plataforma tenia suports (E, F, G) i material de superfície, però no la superfície mateixa. La lletra Y, en canvi, queda **reservada** per a la massa adossada (banqueta) fins que hi haja tres casos documentats: el mateix criteri empíric que en v13 va descartar promoure un camp sense variància.

El delta complet, amb els comptadors de la seua versió (144 camps, 21 elements, 6 sistemes, 53 regles actives, 35 consultes) i la seqüència de migració in situ —patch d'esquema amb dades preservades, regeneració només de consultes, redesplegament del formulari, llista de treball de revisió—, és a `DELTA_v17_v18.md`.

## **5.12. La iteració v19: qualificadors, judicis i termes que no afirmen**

La v18→v19 és un delta menut en superfície (un camp, dues regles, tres actualitzacions de lookup) però amb tres decisions que consoliden criteris transversals del disseny. Els comptadors passen a **145 camps, 55 regles actives (numerades fins a 56) i 36 consultes**; el detall és a `DELTA_v18_v19.md`.

**Primera: la compartició d'element es converteix en patró.** `Jamb_Fabric_Reveal` registra la posició del brancal resolta per la fàbrica mateixa — una cara terminal acabada i deliberada (*masonry reveal*) — i replica exactament el mecanisme de `Sill_Coincides_Cornice`: domini 0/1/9 amb zero de farciment, finestra d'aplicabilitat vigilada per regla (la 55), i exempció de la regla de coherència que altrament llegiria el cas legítim com a registre a mig entrar (la branca de portal de la regla 3). El que en v18 era una excepció puntual és ara una **classe de camp**: el qualificador que diu *com* es va resoldre una posició quan cap element diferenciat la resol. La conseqüència d'interfície és deliberada: aquests camps responen una pregunta, no declaren una presència, i per això porten etiquetes pròpies (*No / Sí / No observable*) sobre els mateixos valors emmagatzemats — dos vocabularis d'etiqueta per a un únic domini, documentats perquè no semblen arbitraris.

**Segona: un «no decidible» és una resposta, no una casella buida.** La regla 9 exigia rol a tota mènsula present i marcava també el valor `ND`, de manera que la branca no es podia buidar mai: set de trenta-cinc files eren soroll permanent. La v19 restringeix la regla a `Is Null` i trasllada al camp de text la mateixa distinció NULL/9/0 que governa els camps numèrics — buit = no avaluat, `ND` = avaluat i no decidible, valor = avaluat i decidit. La coherència que la condició antiga fingia cobrir la vigila ara una regla pròpia (la 56: rol de suport amb el sistema plataforma negat), i el criteri d'assignació és un arbre de tres preguntes escrit al manual, amb la frontera entre «sense vincle» i «indeterminat» calcada de la frontera 0/9.

**Tercera: els termes de referència no han d'afirmar el que no sabem.** El terme anglés de D passa de *Tie walls* a *Transverse wall* perquè «tie» afirma la funció de trava — i la funció és precisament la incògnita, el mateix criteri que va fer triar «díedre» (descriptiu) sobre el terme genètic per als suports. La tipologia MEN es generalitza (*Isolated Structural Element*) amb la descripció ancorada al vocabulari A–Z i sense subtipus, perquè la identitat de l'element ja la diuen els camps d'element mateixos i un subtipus duplicaria informació que caldria vigilar. I la correcció dels zeros de D que el patch v18 havia deixat com a assercions (eren farciment sota el gating v17) tanca el cercle del principi de defaults: **un judici que ningú no ha fet no pot quedar escrit com si algú l'haguera fet** — els sis casos tornen a NULL i entren a la llista de treball de la regla 17.

## **5.13. La iteració v20: identitat, famílies del zero i veredictes desacoblats**

La v19→v20 és el delta més ample del projecte (quinze blocs: quatre decisions de casos límit i onze punts plantejats pel dissenyador després d'entrar el corpus complet, 86 estructures), i alhora el més barat en esquema: un camp retirat, un d'incorporat, una columna de lookup i un índex. La desproporció és el resultat metodològic: **quan els dubtes d'entrada es resolen majoritàriament amb maquinària existent, les decisions estructurals anteriors estaven ben preses** — el que faltava era vocabulari compartit entre l'observador i la base (etiquetes que no induïsquen a error, criteris escrits on el cap els perd). Els comptadors passen a **145 camps, 58 regles actives (numerades fins a 60; retirades 41 i 50) i 40 consultes**; el detall és a `DELTA_v19_v20.md`. Quatre criteris transversals nous:

**Primer: la identitat es bloqueja al motor, no s'alerta a la bateria.** L'auditoria va trobar dues parelles d'estructures reals competint pel mateix codi — col·lisió de numeració de camp que cap regla no hauria evitat, perquè la bateria s'obri quan es vol i el duplicat corromp cada referència creuada des del minut zero. Un codi duplicat no és cap judici transitori: és identitat, com la clau primària, i per això la resposta és un **índex únic** (rebuig en teclejar) i no una regla. La bateria conserva el que sí que és judici: la **concordança** entre el codi i el sector del desplegable (regla 57), que és exactament la costura per on es va colar el cas real.

**Segon: el zero té dues famílies, i la interfície les separa per construcció.** En elements constructius, el 0 afirma una decisió constructiva sobre superfície llegible; en vestigis peribles, afirma només supervivència observable — la màxima «l'absència d'evidència no és evidència d'absència» val per al que pot desaparéixer sense rastre, i per això els camps de vestigis mai no han afirmat origen. La v20 ho fa visible al punt de decisió: el domini dels elements diu *Absent (constatat)* i *Desaparegut atestat*, i les pestanyes de vestigis porten la bandera epistèmica. Els elements constructius de fusta pertanyen a la família forta perquè la fusta encastada deixa encaixos. La fusió dels valors 0 i 9 que es va considerar («absent/no verificable») es va descartar amb l'argument del denominador: sense la distinció, la variabilitat constructiva i la conservació es tornen indistingibles — la confusió contra la qual està dissenyada mitja base.

**Tercer: l'observació no porta el veredicte.** Les bandes pintades de contorn duien noms que afirmaven la classificació del registre (geomètrica = estructura); el corpus de La Petaca ho va refutar — estructures perimetrades amb C i amb O. La v20 desacobla: el tipus diu el **traç**, els quatre trams diuen la **cobertura**, una consulta **deriva** la topologia, i el veredicte estructura-o-panell viu on sempre ha viscut la classificació: la tipologia del registre i l'atestació. El mateix moviment governa el coronament: quina cara corona (cambra o massa basal) es deriva del registre i no es declara, i només la **forma** — l'única cosa no derivable — es pregunta.

**Quart: els gates es pengen de comptadors, no de tipologies.** Els camps de portal es tanquen amb el comptador de cossos de cambra a zero (absorbint les llistes de tipologies, que calia mantenir); l'evidència de fase s'obri amb el comptador de fases; i cap dels dos comptadors té valor per defecte, perquè «una fase» i «cap cambra» són judicis. El contraexemple que fixa el límit del criteri: la maçoneria **no** es gateja pels murs, perquè tretze registres sense cap mur porten fàbrica real — el gate llig les dades, però només quan les dades realment impliquen la inaplicabilitat.

El delta tanca a més tots els pendents històrics del projecte (descomposició per cossos: fora de l'abast del TFM; caracterització de superfícies: via 3D; revisió v17a: completa; rebuild net: milestone del depòsit) i converteix les contingències en disparadors amb l'eixida pre-decidida — la mini-especificació de l'element Y (banqueta) està escrita i s'executa sola al tercer cas documentat.

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

## **9.1bis. Quan un estat es fa passar per un element: la correcció de X (v16)**

La matriu de patrons constructius pressuposa que cada columna A–X registra **una decisió constructiva**. La coberta de cambra la trencava. Definida com «cambra tancada per dalt, amb la solució que siga», no era un element sinó **un estat**: amb el tipus «penya natural» no s'havia construït absolutament res —la visera rocosa tanca la cambra i el constructor no hi va posar cap peça— i tanmateix la columna entrava a la matriu exactament igual que un dintell o una cantonera.

En un corpus on molts mausoleus aprofiten viseres naturals, l'efecte no és menor: la coocurrència de Jaccard i l'anàlisi de correspondències **compten geologia com si fora construcció** i inflen artificialment la similitud entre estructures que no comparteixen cap decisió constructiva. És la mateixa família d'error que ja es va corregir separant l'estat arquitectònic de l'estat dels vestigis mobles, i desdoblant el dintell en v10: dues variables allotjades dins d'una mateixa escala.

La correcció no exigeix cap camp nou. La columna exporta **NA** quan el tancament és roca natural, i NA i no zero, perquè no es tracta d'una absència sinó d'una **solució no construïda**: un zero seria tan fals com un u, i el NA no entra en cap denominador, que és el comportament que la matriu necessita. On hi ha obra, encara que siga parcial —el cas mixt, freqüent, d'obra a l'entrada i roca al fons—, la columna val u, perquè hi ha gest constructiu.

La revisió de camp va afegir un tercer cas que la definició no distingia: mausoleus **sense sostre** sota una visera separada de l'estructura, a vegades per metres. Amb la definició antiga eren indistingibles d'una cambra encaixada sota una visera ajustada, i tanmateix el constructor no havia fet res i l'espai no està tancat. El criteri que ho resol és binari i sense gradient —**sense contacte, no hi ha sostre**—, i el que hi ha per damunt és suport geològic, no coberta.

**Conseqüència sobre els resultats ja obtinguts:** el filtre canvia valors de la matriu, de manera que qualsevol prova de similitud, clúster o correspondències executada abans de la v16 s'ha de refer. La comparació de les dues matrius, abans i després, forma part del protocol de verificació de la transferència.

## **9.1ter. Discriminar filades en voladís de cornisa intercòs (v16)**

Els dos elements ocupen la mateixa franja física —la transició entre cossos— i es confonen amb facilitat quan a penes queden vestigis de la plataforma volada. La confusió no és innòcua: una plataforma arrasada codificada com a cornisa desplaça l'estructura de categoria en tota anàlisi de sistemes.

La distinció és categorial abans que morfològica: **la plataforma és una superfície** (horitzontal, practicable, amb profunditat útil, i té sistema propi), **la cornisa és una junta** (lineal, sense profunditat útil, no practicable, i deliberadament sense sistema). No són excloents: coexisteixen quan hi ha plataforma al cos basal i cornisa marcant la junta amb la cambra, i fer-les excloents crearia un artefacte.

Quatre tests en ordre de prioritat resolen la major part dels casos: evidència positiva de mènsules o bigues (que sobreviu millor que la fàbrica, perquè la roca conserva l'encaix); existència d'un cos construït damunt (sense el qual una cornisa *intercòs* és una contradicció de termes, i que la regla 39 automatitza); progressivitat del voladís; i continuïtat al llarg de la façana. Quan cap no decideix —pocs casos al corpus—, la sortida honesta és el valor «no observable» més una fitxa a la taula d'elements arquitectònics amb el detall morfològic, perquè l'alternativa, triar una de les dues per convenció, fabrica dades.

La retirada del camp de material de la cornisa en v16 contribueix a aquesta discriminació més del que sembla. El camp oferia «bigues de fusta» com a valor, i eixe valor no descrivia una cornisa de fusta sinó **una plataforma mal classificada**: un element horitzontal de fusta volat entre dos cossos és una mènsula o una biga transversal, per definició. En incorporar la clàusula a la definició de l'element —una cornisa és sempre de pedra—, una descripció feble es converteix en un test d'identificació.

## **9.2. Tractament estadístic dels dominis — advertència operativa**

L'advertència del denominador es manté íntegra: **prendre el total d'estructures com a denominador** infla la mostra amb casos mai avaluats, i com que la conservació difereix entre jaciments, la prova mesuraria preservació diferencial i es llegiria com a pràctica diferencial. Denominador vàlid = presents + absències verificades; la cobertura es reporta sempre junt amb la n.

El domini de cinc valors hi afig una segona obligació, i la v13 l'estén a les capes aplicades: **declarar la binarització**. Abans de qualsevol càlcul, els 9 es converteixen a NA (mai a 0), i després es tria conscientment la variable que respon la pregunta: *construït* (1+2+3), *sobreviu* (1+2) o *intacte* (1). Per a les hipòtesis constructives, normalment la primera. Els NA d'origen de l'exportació (zeros de farciment sota sistemes no aplicables) ja arriben buits i no s'han de reomplir. Les dues estratègies de mostra continuen sent legítimes i declarables: restricció a Facade_Observability = Complete, o distàncies binàries amb tractament de nuls; QRY_15 proporciona el recompte que justifica la decisió.

La v13 hi afig una tercera precaució, derivada del defecte NULL: **un buit no és un 9**. Tots dos són `NA` en el càlcul, però només el 9 és un judici. La proporció de 9 sobre el total avaluat mesura l'observabilitat del corpus i s'ha de reportar; la de buits només indica quant treball queda pendent, i és la regla 17 qui la llista.

> Regla general: un 9 no és ni un 0 ni un 1, sinó una cel·la buida; i un 0 de farciment no és un 0 d'asserció, sinó un NA estructural. Qualsevol operació que els convertisca implícitament en dades reintrodueix el biaix de conservació al resultat.

## **9.3. Nivells constructius, cossos superposats i cossos perduts**

Es manté el criteri fixat: N0/N1/Superior són **categories funcionals, no un sistema de numeració** (mai N2, N3), i un cos pertany a N1 si conté o contenia una obertura d'accés. El recompte viu als camps N_Basal_Bodies i N_Chamber_Bodies, i la matriu es manté plana (una fila per estructura). La novetat v11 és el destí de l'evidència de cossos perduts: l'antic camp Lost_Body_Evidence desapareix i la seua informació esdevé una fila de T_LOST_ELEMENTS amb abast *Body*, que cobreix el vocabulari sencer del cos desaparegut — el cas diagnòstic continuen sent les bandes verticals de pigment sobre la roca nua, el fantasma d'un parament que la pintura ha sobreviscut. La conseqüència analítica és ara més fina: els elements del cos desaparegut ja no s'han de marcar 9 en bloc, sinó que poden marcar-se 3 amb l'evidència que ho justifica, cosa que els reincorpora al recompte de *què es va construir*.

## **9.4. Validació de coherència del registre (QRY_16, 58 regles actives en v20)**

La bateria passa de deu (v10) a quaranta-sis regles, emmagatzemades en **cinc** consultes parcials —una unió única de totes les branques supera el límit «query too complex» de JET— i unides a QRY_16_Validation_Check. Continua sent deliberadament **un informe i no una restricció de taula**: una regla dura impediria registrar una absència genuïnament observada en una estructura parcialment col·lapsada. Un resultat buit indica corpus coherent; la regla 17, que llista els camps observacionals encara NULL dels registres heretats, no és una queixa sinó la llista de treball de la revisió manual. El detall de les 46 regles es troba a l'esquema tècnic (secció 6.1); en síntesi cobreixen la coherència component-sistema (1-4), la coherència estat-elements (5-6), els parells presència-detall (7-10, 12-14, 24-26), la coherència classe-contingut (15-16), la bioarqueologia i els materials (18-19), les connexions i l'evidència de pèrdua (20-21), la decoració arquitectònica (22-23), el recompte de cossos basals (27), la pintura rupestre (28-29), l'alerta de cos perdut (30), la plausibilitat mètrica (31) i, des de la v16, la parella de format de pedra (32-33), la coherència entre decoració i element decorat (34-35), la coherència entre pla d'accés i orientacions (36), les files rupestres (37-38) i la discriminació entre cornisa i filada en voladís (39); i, des de la v17, la coherència entre classe de registre i posició de decoració (40), la geometria del traç fora de context (41), la coherència entre fris i motiu constructiu (42), les dues direccions que faltaven sobre els materials culturals (43-44), la divergència de fàbrica sense prou cossos (45) i les fases declarades sense evidència (46).

Aquesta darrera mereix una nota, perquè il·lustra la diferència entre una regla i un bloqueig. La combinació «pigment sobre la penya seguint elements arquitectònics» sembla contradictòria i podria semblar candidata a impedir-se al formulari; però és precisament el **cas diagnòstic de cos desaparegut** —bandes verticals de pigment sobre la roca alineades amb els brancals del cos inferior, on la pintura ha sobreviscut al parament que la sostenia—, de manera que bloquejar-la impediria registrar l'evidència que sosté el valor «desaparegut atestat». La regla informa i suggereix obrir el registre d'evidència; no impedeix res.

**Una lliçó sobre l'error silenciat.** En preparar la v16 es va detectar que la consulta d'unió de la bateria **no existia** a la base v15. El procediment que crea les consultes rebia un nom per a la tercera parcial i un altre per a la referència dins de la unió, de manera que la creació d'aquesta fallava; i el mateix procediment, per disseny, degrada qualsevol fracàs de creació a un missatge a la finestra d'immediat, per a no avortar la construcció de vint-i-huit consultes per culpa d'una. La construcció acabava amb aspecte d'èxit. Les tres parcials existien i s'obrien per separat, cosa que explica que passara desapercebut: **el que faltava era justament la porta d'entrada**, i la bateria de trenta-una regles mai no es va executar sencera sobre el corpus.

La correcció és trivial; la lliçó no. Degradar un error a avís continua sent la decisió correcta en aquest context, però un avís que ningú no llegeix és equivalent al silenci. La v16 hi afig un **recompte de consultes creades contra les previstes**, informat a la línia final del registre d'execució, on ja s'informa del recompte de taules. El principi és general i val la pena consignar-lo: *un error que s'ha decidit no fer fatal ha de deixar rastre en el lloc on algú mira, no només en el lloc on s'ha produït*.

# **10. Formulari principal F_STRUCTURES (v17)**

El formulari centralitza l'entrada de dades amb 12 pestanyes temàtiques, set subformularis vinculats (F_DECORATIONS, F_ROCKART, F_DATING, F_ARCH_FEATURES, F_CONNECTIONS, F_CONN_IN i F_LOST_ELEMENTS) i tots els camps FK i de domini configurats com a ComboBox de dues columnes (valor anglés ocult, etiqueta valenciana visible). La pestanya 11.Sist es reorganitza al voltant dels sistemes: els combos Sys_* dalt de tot i cada grup de components a continuació, perquè aquest és l'ordre real de les decisions — es decideix que hi ha plataforma abans de comptar-ne les mènsules.

**L'ordre de les preguntes és una variable de disseny, no una qüestió estètica.** La v17 en corregeix un cas que va aparéixer omplint fitxes: la pregunta que obri el bloc de materials culturals —n'hi ha?— vivia en una pestanya i la pregunta sobre l'estat d'eixos materials en una altra, i la segona anava abans. No era només incòmode: el gating opera dins d'una pestanya, de manera que declarar que no hi ha materials i haver-ne registrat l'estat era una contradicció **indetectable**. Amb el trasllat, cada pestanya passa a tractar un sol objecte —la construcció d'una banda, el seu contingut de l'altra— i el bloc de materials adopta exactament la mateixa forma que el bioarqueològic: una pregunta que l'obri i el governa sencer. El cost és perdre la contigüitat visual entre l'estat de l'arquitectura i el dels vestigis mobles, i es considera un guany: la proximitat suggeria que eren dos graus de la mateixa cosa quan un parla de l'edifici i l'altre del que hi ha dins.

**La visibilitat com a condició de la integritat.** El mateix contrast amb les dades va revelar que setze files de decoració de cinquanta-sis **no es podien veure des del formulari**: el filtre que reparteix la taula entre les dues subseccions eliminava les files sense posició assignada, que eren precisament les que la migració anterior havia deixat marcades amb la nota «completar posició». Existien, comptaven a les consultes i a les regles, i completar-les era exactament el que la interfície impedia. La correcció és d'una línia; el principi que se'n deriva no: **una dada que existeix i no es pot veure és pitjor que una dada que falta**, perquè participa dels recomptes sense poder ser revisada. Cap decisió d'interfície que oculte controls s'hauria d'aplicar sense haver comprovat abans que no oculta files.

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

El paquet v17 comprén cinc scripts, amb una única ruta d'ús:

- **chachapoya_DB_v17.bas → BuildDB()**, sobre una base en blanc: crea les 22 taules, fixa els valors per defecte (0 als 20 camps d'element, buit a la resta; sistemes a «Absent»), pobla els lookups, estableix les 25 relacions i l'índex únic sobre la parella de connexió, i genera les 33 consultes, informant del recompte obtingut contra el previst.
- **chachapoya_Form_v17_val.bas → BuildForm()**, a continuació: crea els set subformularis, F_STRUCTURES amb 12 pestanyes, els combos de dues columnes, les captions DAO, el mòdul de gating i les validacions de subformulari.
- **chachapoya_EXPORT_v16.bas → ExportAll()**, sobre la base d'origen: escriu vuit fitxers CSV inspeccionables.
- **chachapoya_IMPORT_v17.bas → ImportAll()**, sobre la base nova i buida: llig els CSV, normalitza les parelles de connexió girant-ne la cronologia i escriu els registres.
- **diagnostic_v13.bas → RunDiagnostic()**, de només lectura: comprovacions de coherència i de variància sobre el corpus, sense modificar cap dada ni cap esquema.

**Sobre la decisió de no migrar.** Les iteracions anteriors van transformar la base preservant els registres, amb un protocol explícit que val la pena consignar perquè és citable com a bona pràctica: ordre estricte *renomenar → afegir → poblar → eliminar*, de manera que cap camp desapareix abans que el seu substitut estiga poblat i verificat; cap valor existent es toca mai, i els valors per defecte només afecten files futures; els esborrats finals doblement barrats per un commutador explícit i per precondicions comprovades en viu contra les dades; i recomptes d'afectació a cada bloc, cosa que converteix el registre d'execució en pista d'auditoria.

**Sobre la transferència en dos passos.** La v13 va reintroduir els registres a mà; des de la v16 no es repeteix, però tampoc no es torna a la migració *in situ*. La via adoptada és exportar a un fitxer de text, revisar-lo i importar-lo a una base construïda de zero, i el fitxer intermedi **és el motiu del disseny, no un tecnicisme**: és l'únic moment en què una assignació errònia encara es pot detectar barata.

**Sobre l'abast desigual de la revisió (v17).** Aquesta transferència il·lustra que la regla conservadora —el que ha canviat de criteri arriba buit— **no és un reflex sinó un judici cas per cas**, i que aplicar-la de manera indiscriminada té un cost. Ací cap camp de l'estructura no ha canviat de criteri, de manera que no hi ha ni una sola columna de rescat i tot el que és nou arriba buit per absència. La revisió manual es concentra en set files de decoració, i el motiu que no es puguen convertir automàticament és precís: el valor «totes dues» s'estava usant en tres d'elles per a dir una banda en U invertida —tres trams—, de manera que traduir-lo mecànicament afirmaria que el tram superior és absent precisament on hi és. Un zero equivocat és pitjor que un buit, perquè el zero és un judici.

El cas contrari apareix a les connexions, i convé no confondre'ls. La seua cronologia també canvia de redacció, però no de significat, i la conversió és per tant determinista: l'importador la fa sol, inclosa la inversió quan normalitza la parella. Enfosquir-les hauria fabricat feina de revisió a partir d'un canvi de nom. **El criteri és si el judici que sosté el valor continua sent vàlid, no si el text que el nomena ha canviat.**

Tres decisions el governen. La primera: **cap identificador travessa la frontera**. Totes les claus alienes s'exporten com el valor llegible que apunten i es tornen a resoldre a l'arribada. No és pulcritud —en la transferència anterior els identificadors d'una taula de vocabulari s'havien reorganitzat entre versions, de manera que claus numèricament idèntiques apuntaven a posicions **diferents**, i dinou de vint-i-nou files de decoració haurien quedat silenciosament mal etiquetades. Una clau errònia no dona cap error ni mostra cap símptoma: només canvia el que diuen les dades.

La segona: **totes les claus es resolen abans d'obrir el registre nou**. Resoldre-les després converteix un diagnòstic net —«aquest valor de vocabulari no existeix»— en una fallada d'escriptura amb el registre a mig fer, que no diu res sobre què anava malament.

La tercera és la que connecta amb l'epistemologia de la v13: **els camps el criteri dels quals ha canviat arriben buits**. Un valor transferit sense revisió afirmaria un judici que ningú no ha fet sota el criteri nou, i tornaria a barrejar herència amb observació. Però buidar-los en silenci llançaria la lectura antiga, que és el millor punt de partida de la revisió; per això cada camp afectat s'exporta dues vegades, la columna viva en blanc i una columna informativa amb el que deia la versió anterior. **El fitxer intermedi es converteix així en l'espai de treball de la revisió**: es mira la fotografia i, si la lectura antiga aguanta sota el criteri nou, es copia de tornada. El que es copia és un judici; el que queda buit queda honestament buit, i una consulta específica (QRY_19) el llista.

Val la pena consignar que això té un cost real: la revisió afecta el portal de tots els registres amb obertura, la coberta de tots els que tenen sostre de roca o mixt, i els dos camps nous de pedra a tot el corpus. És feina de fitxa i de fotografia, no de codi, i era una raó suficient per a **no acumular-hi** la descomposició per subnivells.

Per a la v13, en canvi, s'havia optat per **reintroduir els registres sobre una base construïda de zero**. La raó no és de comoditat sinó epistèmica: el diagnòstic previ va mostrar que una part substancial dels valors del corpus no eren judicis de l'investigador sinó **herència de valors per defecte de versions anteriors** —una vintena de «no observable» en un sol camp, procedents del defecte de la v10—, i eixos valors, en no ser buits, quedaven fora de la llista de treball i entraven als recomptes com si fossen observacions. Cap migració els podia distingir dels judicis reals. Reintroduint, cada valor del corpus passa a ser deliberat, que és la condició prèvia per a publicar qualsevol percentatge.

## **11.2. Restriccions tècniques de VBA/Access**

Es mantenen les restriccions documentades: cap continuació de línia (patró sql = sql & "..."), codificació cp1252 sense caràcters no-ASCII al codi, paraules reservades evitades, patró de creació i renomenament de formularis, N-1 parèntesis per a N taules en JOIN, dbRelationDontEnforceIntegrity (valor 2) per a l'autoreferenciant i les dues de T_CONNECTIONS, procediments auxiliars Private, DROP previ per a objectes regenerables, i combos configurats per ControlSource. S'hi afigen les apreses en les iteracions v11-v13: JET no pot convertir TEXT a BYTE in situ (el desdoblament del dintell exigeix renomenar per DAO, crear el camp nou i poblar-lo per UPDATE); un renomenament DAO no reescriu el SQL de les consultes existents (les consultes afectades queden trencades fins que es reconstrueixen, i l'ordre dels blocs ho preveu); una UNION de més de ~30 branques supera el límit «query too complex» (la bateria s'emmagatzema en quatre parts des de la v16); la injecció de mòduls de formulari exigeix l'accés al model d'objectes VBA al Centre de confiança (si està desactivat, tot es construeix igualment i l'script avisa); i amb la columna lligada oculta, Access imposa LimitToList (la via d'escapament és el valor «Other (see Notes)»).

## **11.3. Exportació per a l'anàlisi estadística**

Sense canvis de procediment: QRY_05 (exportació plana) i QRY_13 (matriu AX+SYS) via Dades externes → Exportar, amb els camps categòrics en format llegible gràcies als JOIN amb els lookups. La conversió 9 → NA i la declaració de la binarització (secció 9.2) són pas previ obligat en R.

# **12. Limitacions i perspectives**

La limitació principal continua sent la cobertura desigual del corpus (La Petaca Sector Sud amb ≥96 estructures identificades, no totes amb fitxa completa). Els camps de base documental i observabilitat, la regla 17 (llista de treball dels NULL) i QRY_15 converteixen aquesta limitació en una variable mesurada en lloc d'un buit silenciós. L'entrada dels 35 registres ja s'ha fet (v17a) i la migració v18 s'aplica in situ; la revisió manual associada —la regla 17, `QRY_21`, `QRY_22`, `QRY_23` i `QRY_24` en mesuren el progrés camp a camp— és la tasca oberta immediata. Alguns registres exigiran, a més, revisió sobre el model 3D: els que portaven el valor de suport ara desdoblat, i aquells en què la parella de suports apareix en ordres oposats i cal determinar amb el criteri nou quina forma rep la càrrega i quina estabilitza.

T_ARCH_FEATURES i L_DEC_TYPE mantenen l'escalabilitat de l'esquema sense modificar-ne l'estructura. El vocabulari A-X, ara formalitzat com a taula amb la distinció explícita element/sistema, és el primer intent de normalització terminològica per a les estructures funeràries Chachapoya en la bibliografia, i la seua aplicació creuada a La Petaca i Diablo Wasi és una de les aportacions originals del TFM.

**La línia de treball oberta per la v16 — i tancada per la v20 —** és la descomposició per subnivells: descriure per separat cada cos basal i cada cos de cambra en les estructures que en tenen més d'un, amb els intranivells derivats de la seua successió. No és una qüestió de vocabulari sinó de **cardinalitat** —exigiria una taula de cossos amb els elements A–X penjant de cadascun, i tocaria l'esquema, el formulari, la matriu de patrons i el criteri d'unitat de registre—, i s'ha ajornat deliberadament perquè la dada que l'ha de decidir encara no existeix: quants registres tenen cossos amb atributs realment divergents. Si són pocs, la solució ja és a l'esquema (registres separats vinculats per T_CONNECTIONS); si són molts, cal la taula. Entrar les dades amb l'esquema v16 és el que produirà eixa xifra, i la decisió es podrà prendre en fred. **La xifra va arribar amb el corpus complet (v20): cap anàlisi feta ni prevista ha demanat atributs per cos, i els casos realment divergents es resolen amb registres separats vinculats (la recodificació dels antics sufixos a/b en codis propis ho demostra funcionant). La descomposició queda formalment fora de l'abast del TFM; si una anàlisi post-publicació la demana, el mecanisme és el delta ordinari.**

Alguns camps queden identificats com a candidats a baixar a nivell de cos si eixa descomposició prospera: el format i el treball de la pedra, en particular, perquè un cos basal de pedra irregular sota una cambra de blocs semiescairats és un cas real i és precisament el tipus d'evidència que separa fases o mans.

De cara a publicacions futures, es manté la perspectiva de migració a GeoPackage (.gpkg) per a la integració nativa amb QGIS, amb Access com a interfície d'entrada.

# **13. Referències**

Epstein, L.; Toyne, J.M. (2016). When Space Is Limited: A Spatial Exploration of Pre-Hispanic Chachapoya Mortuary and Ritual Microlandscape. A: Osterholtz, A.J. (ed.), Theoretical Approaches to Analysis and Interpretation of Commingled Human Remains. Springer, Switzerland, pp. 97-124.

Ribera-Torró, E. (2023). Chacha XR. Una experiència immersiva de no-ficció per a l'arqueologia Chachapoya. Treball de Fi de Màster, Màster Universitari en Arts Visuals i Multimèdia, Universitat Politècnica de València.

Ribera-Torró, E.; Toyne, J.M.; Del Aguila, R.; Ribera, J.A.; Galexner, J.; Anzellini, A.; Pans, M. (2026). Extended Reality on Chachapoya Cliffside Necropolises: From Digital Documentation to Public Engagement. Open Archaeology, 12(1). DOI 10.1515/opar-2025-0071.

Toyne, J.M.; Anzellini, A. (2017). Sociedad, identidad y variedad en los mausoleos de La Petaca, Chachapoyas. Boletín de Arqueología PUCP, 23, 231-257.

Toyne, J.M.; Anzellini, A.; Epstein Miculas, L.; Mejías Pitti, I.; Puig Castell, J.; Guinot Castelló, S. (2018). Going Vertical: Using Vertical Progression Techniques to Explore a Cliff Necropolis in Late Precolumbian Chachapoyas, Peru. Advances in Archaeological Practice, 6. DOI 10.1017/aap.2018.31.


---

# Nota de versió v24: el doble registre de l'orientació

La v24 introdueix un sol camp, però el seu interés metodològic
supera la seua mida: `Facade_Azimuth_Deg` materialitza la decisió
de mantindre **dos registres del mateix fet amb estatuts
epistemològics diferents**, en lloc de substituir l'un per
l'altre.

**El problema.** L'orientació de la façana existia com a
observació de camp (`Facade_Orientation`, sector de 8): dada
primària, presa davant de l'estructura, amb la resolució i la
fal·libilitat de l'observació directa. La fotogrametria ofereix
ara una mesura instrumental de resolució de grau, derivada de la
normal del *bounding box* sobre els models de sector
georeferenciats (EPSG:32718, azimut per `atan2(ΔE, ΔN)`). La
temptació òbvia — reemplaçar el sector per l'azimut — hauria
esborrat la dada observacional i, amb ella, la possibilitat de
contrastar els dos canals.

**La decisió.** Els dos camps conviuen amb papers explícits:
l'observacional resta com a dada primària de camp (editable al
formulari); l'instrumental entra només per importació i es mostra
bloquejat. La **regla 68** formalitza la relació com a *alerta de
concordança*: divergència circular superior a 67,5° (un sector i
mig) implica que un dels dos canals falla — i cap dels dos té
prioritat automàtica: l'alerta demana diagnòstic, no
sobreescriptura.

**El prerequisit instrumental.** El desenvolupament del pipeline
de mètriques per *shapes* (agost 2026) va demostrar empíricament
que el *bounding box* és l'instrument compartit de tres productes
— azimuts, ortomosaics i projeccions mètriques — i que rotar-lo
per a obtindre vistes zenitals el descalibra per als tres alhora.
D'ací dues regles operatives incorporades al protocol: les vistes
es seleccionen per paràmetre (`VIEW_AXIS`), mai rotant el
*bounding box*; i l'exportació d'azimuts es regenera **després**
d'una passada de l'extractor de mètriques amb zero avisos
d'obliqüitat, que actua com a **certificat de calibratge** de
l'instrument per a tot el corpus. La cadena de custòdia de la
dada instrumental queda així documentada de l'instrument al camp.

**Coherència amb el marc general.** El camp queda fora de les
llistes de treball pel mateix criteri que `Metrics_Available`
(v22): la completesa d'un camp instrumental no és una tasca
d'entrada sinó l'estat d'un flux extern. I el parell
observacional/instrumental replica, a escala de camp, el principi
que governa tot el disseny: les fonts no es barregen, es
confronten — la mateixa lògica que separa l'estat conservat del
restituït en el registre mètric, o l'observació de la inferència
en el domini dels elements A–X.


----

## Nota de versió v25 — Coherència mètriques↔BD, dos camps d'observació i concordança bibliogràfica

**Epistemologia de la reconciliació.** El delta v25 tanca el circuit
obert pel v24: si aquell va portar l'instrument a la BD (l'azimut
fotogramètric, amb la regla 68 com a frontissa amb l'observació de
camp), aquest porta la *reconciliació sistemàtica* entre el corpus
mètric de shapes i el registre estructurat. El principi rector és la
**primacia direccional**: la shape és evidència positiva (pot pujar
una presència o un recompte, mai baixar-los, perquè no dibuixar no
és negar), els recomptes dibuixats són fita inferior, i la mesura
instrumental ompli el buit però no esborra mai una observació sense
firma. La firma mateixa esdevé registre (`T_METRIC_REVIEW`): cada
discrepància, la seua evidència, la decisió presa i la data. És la
traducció a taula del lema operatiu del projecte — *els QC
aconsellen, l'investigador firma* — i converteix la reconciliació en
un procés auditable i repetible, no en una sessió d'edicions
disperses.

**Protocol de correcció per capes.** Les incoherències de vocabulari
es corregeixen a la capa d'origen (etiqueta al model → Metashape i
reexportació; llacuna expressiva → delta conscient de la convenció;
llacuna del lookup → delta de BD), mai aigües avall. El cas
fundacional — tres cossos `1.0`, gramaticalment impossibles
(`X.Y, Y≥1`) — va resultar ser typos d'etiqueta: capa 1, la
convenció v1.1 aguanta intacta.

**El criteri cos/estructura (precedent EA06–EA62).** La revisió va
forçar a explicitar un discriminant que estava implícit: un cos
pertany al mateix **esdeveniment constructiu** (fàbrica travada,
aparell que continua); una estructura que *descansa* sobre el
coronament d'una altra amb discontinuïtat és un esdeveniment
posterior i mereix EA pròpia, amb la superposició registrada a
`T_CONNECTIONS` com a dada de cronologia relativa — evidència
directa de l'ús acumulatiu a llarg termini que defineix la
necròpolis de cingle com a categoria. Compartir el suport geològic
no és criteri d'identitat: tot el sector comparteix el cingle.

**Dos camps nascuts de l'observació, amb el principi de sempre.**
`Jamb_Fabric` (els muntants del portal — cada costat de l'obertura
resolt com a llosa, composta o fàbrica — registrats com a parella
completa sense col·lapsar, amb el composite llosa-més-fàbrica-travada
com a cas central i, s'ha vist en classificar, majoritari) i `Niche_Partition`
(la llosa que parteix el nínxol natural) entren perquè tenen casos
confirmats i nombrosos; el que no en té, no entra: ni recompte de
compartiments (cap cas triple), ni valor Mixed (la parella completa el
fa innecessari, i la coherència amb l'element O la vigila la regla
de bateria R69), ni extensió a CAV (fenomen distint, es reobriria
amb el cas davant). Tots dos il·luminen la mateixa tesi que el
clúster CAM–MAU: una tradició constructiva única modulada pel que
la geologia ofereix (H04/H05), amb el mòdul portal com a lloc
d'integració entre sistemes (H01).

**Concordança bibliogràfica (OE1).** `L_ELEMENTS` incorpora
`Term_Lit`/`Lit_Source`: platform-base↔B, cornice↔I i frieze↔M
segons Guengerich (2014), amb la d'I com a concordança **de
posició** (la juntura platform-base/superstructure és exactament la
posició de l'element I). Es documenta també el que **no** mapeja,
perquè el rigor inclou els buits: *superstructure* correspon al
concepte de cossos N1, no a un element; *bedrock* al suport
geològic (L_SUPPORT); *ledge* (Epstein & Toyne 2016) ja viu a la
convenció com a `led`; i *sub-cornice* (Fig. 3 de Guengerich 2014)
queda fora per absència de cas al corpus — amb l'àncora
bibliogràfica esperant si mai n'apareix un. El mecanisme és el
mateix que valida el vocabulari contra Aoujgal: mostrar que l'A–X
no és idiosincràtic sinó que dialoga, terme a terme, amb la
terminologia publicada i amb els corpus comparatius.
