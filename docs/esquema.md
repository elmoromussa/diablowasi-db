**ESQUEMA DE LA BASE DE DADES ARQUEOLÒGICA**

*La Petaca i Diablo Wasi (Leymebamba, Amazonas, Perú)*

chachapoya_DB_v25.bas + chachapoya_Form_v25_val.bas (v25 sense patch: `RebuildQueriesV25()` + `BuildForm()`; chachapoya_Worklist_v25.bas com a navegador; chachapoya_CheckMetrics_v25.bas com a pipeline de coherència mètriques↔BD)

Sub BuildDB() + Sub BuildForm() | Microsoft Access JET SQL | Esteve Ribera Torró

*Versió 25 del document — actualitzada segons el codi v25 (agost 2026). Consolida el delta v24→v25 (tres blocs: **A**, el pipeline de coherència mètriques↔BD — taula `T_METRIC_REVIEW` amb FK inline, mòdul `CheckMetrics`/`ApplyMetricReview`, primacia direccional de les shapes amb firma manual obligatòria, creuaments C1/C2/SLOT/C3a/C3b/C5/C7, consultes QRY_30–32 i tres fonts noves a la worklist; **B**, dos camps d'observació nous — `Jamb_Fabric` (fàbrica del brancal existent: Monolithic slab / Composite slab-masonry / Fabric as jamb / ND, frontera constructiva i precedència per a asimetries, gatejat per Brancals) i `Niche_Partition` (llosa que parteix el ninxol natural: Present / Absent / Attested lost / Not observable, gatejat a NIX, CAV exclosa); **C**, la concordança bibliogràfica del vocabulari — `Term_Lit`/`Lit_Source` a `L_ELEMENTS` amb platform-base↔B, cornice↔I i frieze↔M segons Guengerich 2014, concordança de posició en el cas d'I). Text heretat: Consolida el delta v23→v24 (bloc únic: `Facade_Azimuth_Deg`, l'azimut fotogramètric del pla exposat — INTEGER 0–359, escrit només per l'importador de georeferenciació, mai a mà; regla 68 d'ALERTA sobre la divergència circular amb `Facade_Orientation` > 67,5°; fora de les worklists pel precedent `Metrics_Available`; al formulari bloquejat a 2.Arq (7,2) amb `Visibility_Valley` desplaçada a (9,2); la passada neta de l'extractor de mètriques com a certificat de calibratge del bbox). Text heretat: Consolida el delta v22→v23 (set blocs: finestra del brancal amb l'eix del portal i regla 55 revisada, amb el residu del bloc 6 v21 rediagnosticat; codi del registre al títol de la finestra, sense sufix; Planta amb «Triangular»; murs redefinits com a MURS DE LA CAMBRA amb regles 65-66 i 0 derivable; `EA_Number` emmagatzemat — reversió explícita de la decisió v20 «mai a banda», amb la regla 67 de concordança; `L_SUBSECTORS` i `ID_Subsector`; Bedrock i banqueta/esglaó a 12.Extra) i l'interí v22a. Text heretat: Consolida el delta v21→v22 (cinc blocs: 2.Arq oberta als contextos naturals amb la declaració de fàbrica de porter i el 0 derivable als panells; el porter de disponibilitat mètrica `Metrics_Available` amb regla 63, fora de les llistes de treball a propòsit; datació només per C14 amb regla 64, reversió explícita; el codi a la capçalera del formulari; neteja de consultes llegades). Text heretat: Consolida el delta v20→v21 (sis blocs: criteri de la terrassa revisat amb Z reservat al sistema volat i regla 62; la declaració de fàbrica `Masonry_Present` amb regla 61; dues errates v20 corregides al formulari — bandes de contorn i catàleg de 12.Extra —; etiquetes «(opc.)»; residus de worklist documentats). Text heretat: Consolida el delta v19→v20 (§1.9: quinze blocs tancats en dues tandes de sessions + auditoria completa del corpus de 86 estructures). Text heretat: Consolida el delta v18→v19 (§1.8), tancat sobre els candidats que la secció Pendent del delta anterior ja especificava més les troballes de la inspecció directa de la còpia amb dades. Text heretat de la v18: Consolida el delta v17a→v18 (§1.7), nascut de la primera campanya real d'entrada de dades: 35 estructures entrades i les notes de camp que l'entrada va generar. Text heretat de la v17: Consolida el delta v16→v17, que naix de dues fonts alhora: les observacions recollides emplenant registres i, per primera vegada, **la inspecció directa de la còpia local** (36 registres, 56 files de decoració). La segona en va canviar el resultat: tres punts plantejats com a preguntes obertes eren **incoherències ja presents a les dades**, i dos camps discutits en abstracte tenien un ús real que decidia la discussió sense necessitat d'argumentar-la.*

**v25: migració in situ, no reconstrucció.** `PatchV25()` afig `T_METRIC_REVIEW` (FK inline a `T_STRUCTURES`, captions DAO), `Jamb_Fabric` i `Niche_Partition` a `T_STRUCTURES`, i `Term_Lit`/`Lit_Source` a `L_ELEMENTS` amb les tres concordances (guardades: només ompli NULL); `RebuildQueriesV25()`, `BuildForm()` v25 i `BuildWorklist()` fan la resta, i `CheckMetrics` pobla la revisió amb els CSV nets vigents. El paràgraf següent descriu el mecanisme v22 i es conserva com a referència. **v22 (i anteriors): migració in situ, no reconstrucció.** `PatchV22()` afig el porter mètric amb les dades dins (columna + derivació), deriva el 0 de fàbrica dels panells i neteja la consulta òrfena; `RebuildQueriesV22()`, `BuildForm()` v22 i `BuildWorklist()` fan la resta. El paràgraf següent descriu el mecanisme v21 i es conserva com a referència. `PatchV21()` afig la declaració de fàbrica amb les dades dins (columna nova + derivació del valor 1 on hi ha detall de maçoneria), `RebuildQueriesV21()` regenera només les consultes, `BuildForm()` v21 redesplega el formulari i `BuildWorklist()` el navegador; el revert de les plataformes Z-sol del criteri antic de terrassa és **fila a fila via `QRY_29`, mai UPDATE cec**. El paràgraf següent descriu el mecanisme v20 i es conserva com a referència. `PatchV20()` modifica l'esquema amb les dades dins (retirades, columna nova, `Sector_Code`, índex únic, lookups per nom, neteges), `RebuildQueriesV20()` regenera només les consultes i `BuildForm()` v20 redesplega el formulari. **El patch v20 avorta si troba codis duplicats**: la identitat es corregeix abans de tot. `BuildDB()` **no s'ha d'executar mai** sobre una base amb dades: fa DROP de `T_DECORATIONS`, `T_LOST_ELEMENTS`, `T_CONNECTIONS` i `T_ARCH_FEATURES`. El paràgraf següent descriu el mecanisme de les versions anteriors i es conserva com a referència.

**Construcció de zero, transferència en dos passos (referència v17).** La base v17 es construeix des de zero amb els scripts canònics i les dades hi arriben per `chachapoya_EXPORT_v16.bas` → CSV inspeccionable i editable → `chachapoya_IMPORT_v17.bas`. **Cap valor arriba sense haver pogut ser mirat**, i els camps el criteri dels quals ha canviat arriben NULL: un valor transferit sense revisió afirmaria un judici que ningú no ha fet sota el criteri nou. La precondició per a publicar qualsevol percentatge continua sent la mateixa — cada valor emmagatzemat, un judici deliberat.

*Aquesta transferència és més lleugera que l'anterior i convé dir per què: **cap camp d'estructura no ha canviat de criteri**. Tot el que és nou a la v17 és nou de zero, de manera que arriba NULL per absència i no cal enfosquir res. La v15→v16 va necessitar sis columnes de rescat; aquesta, cap. La revisió es concentra en **set files de decoració**, i el motiu que no es puguen convertir automàticament està a 1.6(a).*

# **1. Resum general**

| **Element** | **Valor** |
| --- | --- |
| Taules principals | T_STRUCTURES (145 camps), T_DATING, T_INDIVIDUALS, T_GROUPS, T_DECORATIONS, T_ARCH_FEATURES, T_CONNECTIONS, T_LOST_ELEMENTS |
| Taules lookup | L_SITES, L_SECTORS, L_TYPOLOGY, L_SUPPORT, L_STATUS, L_MATERIAL_STATUS, L_VOL_METHOD, L_COORD_METHOD, L_GROUP_TYPE, L_CAMPAIGN, L_STRUCT_BODY, L_DEC_TYPE, L_ELEMENTS, L_LOST_EVIDENCE |
| Total taules | 22 |
| Total relacions | 25 (inclou l'autoreferenciant de T_STRUCTURES, les dues de T_CONNECTIONS, la del suport secundari i les quatre de T_LOST_ELEMENTS) |
| Consultes SQL | 40 al script v20 (família QRY_01–QRY_27; la bateria QRY_16 comprén cinc consultes auxiliars, **set** parcials de regles i la consulta unió amb **58 regles actives**). Una BD migrada des de 17a en porta 41, amb `QRY_22_V17a_Review` |
| Formulari | F_STRUCTURES — 12 pestanyes + subformularis F_DECORATIONS, F_ROCKART, **F_DATING**, F_ARCH_FEATURES, F_CONNECTIONS, **F_CONN_IN** (només lectura) i F_LOST_ELEMENTS. Gating de tres nivells amb emplenat ràpid confirmat. Etiquetes UI en valencià; valors emmagatzemats en anglés |
| Dominis observacionals | **21 camps amb domini de cinc valors** 0/1/2/3/9: **els elements A–X + Z, i només ells**; **31 camps amb domini 0/1/9**: atributs de tècnica, capes contínues, processos, vestigis mobles, qualificadors, judicis agregats i **els quatre trams del contorn a `T_DECORATIONS`**; 6 camps de sistema TEXT amb domini de sis o quatre valors |
| Valors per defecte | **Dos, deliberadament**: 0 als 21 camps d'element governats per un `Sys_*` (la regla 4 hi exigeix el zero de farciment) **i als dos qualificadors de compartició (`Sill_Coincides_Cornice`, `Jamb_Fabric_Reveal`)** (gating derivable, §2.5); **NULL a la resta** — buit = no avaluat encara, 9 = avaluat i no examinable, 0 = avaluat i absent |
| Idioma BD | Anglés: noms de taules, camps i **valors emmagatzemats**. UI del formulari en valencià, servida per les columnes `Name_VAL` dels lookups i per la columna d'etiqueta de les llistes de valors — en tots dos casos la columna emmagatzemada està oculta, de manera que **reanomenar una etiqueta no pot tocar cap dada**. Capçaleres de subformulari via etiquetes adjuntes; captions DAO com a reforç |
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

**(i) Record_Class a L_TYPOLOGY.** Dels 35 registres actuals, 5 no són estructures: qualsevol percentatge sobre 35 és erroni. Cada tipologia porta una classe de registre — *Built funerary structure* (EA-MAU, EA-CAM, EA-TER, EA-PLA-V, Unclassifiable), *Natural funerary context* (NIX, CAV), *Structural trace* (MEN), *Rock art panel* (PR), *Pending classification*. La classe filtra cada anàlisi i governa el gating de pestanyes. **MIX s'elimina** (tot EA-CAM és mixt per definició: no distingia res) i **ND es desdobla** en *Unclassifiable* (un resultat: evidència insuficient per a classificar una estructura construïda) i *Not yet classified* (un estat de treball, exclòs de tota consulta analítica).

**(j) Renomenaments.** Lateral_Walls → **Return_Wall** (V: mur que retorna cap al penyal, perpendicular al pla de façana) i Lateral_Wall_Faces → **Facade_Flank** (L: parament dins del pla de façana, flanquejant l'obertura): els dos noms v10 deien «lateral» sobre coses distintes, i les dades confirmen que són variables independents. Recessed_Portal → **Recessed_Frame**; Rear_Wall_Type → **Rear_Closure_Type**. **Rear_Wall (W) desapareix com a camp**, redundant amb Sys_Chamber; l'entrada W es conserva a L_ELEMENTS com a referència de vocabulari. A L_STRUCT_BODY, l'entrada LWF passa a **FFL** (Facade flank).

## **1.2. Canvis de l'addendum v12**

**(a) Dec_Present (BYTE 0/1/9, per defecte 0).** L'eliminació dels booleans de decoració (1.1.f) va deixar l'**absència** de decoració irregistrable: «cap fila a T_DECORATIONS» no distingia l'absència verificada (0) de la no observable (9), una violació de l'axioma d'observabilitat sobre el qual es construeix tot el domini. Dec_Present restaura el judici agregat; el detall continua vivint **només** a T_DECORATIONS. Les regles 22–23 mantenen les dues coses d'acord, i al formulari Dec_Present governa el subformulari de decoració.

**(b) Regles 22–26 (QRY_16c).** Coherència Dec_Present ↔ files de T_DECORATIONS (22, 23); detall de revoc o pigment registrat sota presència 0 o 9 (24, 25); Mortar_Type incompatible amb Mortar_Present = 0 (26). Són la contrapartida a escala de corpus del gating v12 del formulari: **el formulari preveu en el moment d'entrada; la bateria detecta el que hi entra pel costat** (taules obertes a mà, importacions, gating desactivat al Centre de confiança).

**(c) Gating de tres nivells + emplenat ràpid (formulari).** Vegeu la secció 11.

## **1.3. Canvis de la iteració v12→v13**

**(a) `L_SUPPORT`: desdoblament de «Fissure/Crack».** El valor únic fonia dos fenòmens geològics amb lògiques estructurals **oposades**: la **diàclasi** (escletxa vertical) *conté* — aporta dos paraments laterals, constreny la planta i s'omple per a generar el nivell basal, de manera que la roca actua com a encofrat; la **junta d'estratificació** (rebaix horitzontal, erosió diferencial d'un estrat tou entre bancs competents) *ancora* — aporta una ranura contínua on empotrar jàcenes (A) o mènsules (E), amb treball a tallant i moment en lloc de compressió. Compartir casella era un accident del vocabulari («fissura» serveix per a totes dues en català i en anglés), no una afinitat real.

El **mode mecànic** s'escriu a la columna `Description` del lookup i no com a camp de `T_STRUCTURES`: es dedueix deterministament del nom de la forma, de manera que un camp per registre seria derivable — no informació, sinó una oportunitat de contradicció, i quasi constant per forma sobre 35 registres. Escriure'l és el que fa la deducció legítima: fixa el criteri per a l'observador i deixa la derivació disponible per a R com a variable agregada.

*Solapament a tindre present:* micro-repisa i junta d'estratificació són **el mateix fenomen vist des de cares oposades**. Quan un estrat tou s'erosiona, el banc inferior queda volat (micro-repisa: s'hi reposa) i alhora queda una ranura (rebaix: s'hi encasta). Quina de les dues coses és depén exclusivament del que va fer el constructor, no de la roca — d'ací el criteri escrit a cada `Description`.

**(b) Valor per defecte NULL on cap gating no exigeix un zero.** El defecte 0 de la v11 existia perquè **la regla 4 exigeix el zero de farciment** als components d'un sistema no aplicable. Els camps de fora de qualsevol sistema el van heretar per uniformitat, sense raó pròpia. Un **0 fals asserta de més**: infla el denominador amb absències que ningú no ha verificat, i com que el biaix correlaciona amb la qualitat d'observació —que difereix entre LP i DW— reintrodueix exactament l'artefacte que tot el domini existeix per a evitar.

**NULL i no 9**, deliberadament: el 9 no és una casella buida sinó una afirmació («aquesta posició no es pot examinar»). Com a defecte faria un camp intacte indistingible d'un de marcat a consciència i desactivaria la regla 17, la funció de la qual és precisament llistar el que encara no s'ha avaluat. El diagnòstic previ ho va confirmar sobre el corpus v12: 20 nous a `Fire_Damage` heretats del defecte v10, cap d'ells un judici real.

**(c) Domini de cinc valors estés a les capes aplicades.** `Mortar_Present`, `Plaster_Present`, `Pigment_Present`, `Dec_Present` i `RockArt_Present` passen a 0/1/2/3/9. **No és una promoció general:** són els únics camps de tres valors que registren una **capa aplicada a la fàbrica**, de manera que tenen integritat física, es degraden gradualment i deixen rastre en desaparéixer. `L_LOST_EVIDENCE` ja portava *Mortar imprint* — el catàleg d'evidències preveia un cas que el domini del camp no podia expressar.

**Criteri general d'assignació de domini:**

> **Cinc valors** per a tot allò que és **fàbrica o capa aplicada sobre fàbrica**.
> **Tres valors** per a **processos, vestigis mobles i qualificadors**: sense integritat física pròpia, «parcial» i «desaparegut» no hi signifiquen res comprovable.

Prova d'absurd de la promoció general: `Looting` = 2 («saqueig parcial») — el saqueig no té integritat física; `Fire_Damage` = 3 («dany per foc desaparegut») — sense sentit; `Mat_Textiles` = 3 exigiria evidència material d'un tèxtil desaparegut, categoria inexistent.

**(d) Art rupestre associat integrat a `T_DECORATIONS`.** No hi ha segona taula: en duplicaria l'estructura, les relacions i les regles de validació, i obligaria a una UNION cada vegada que es volguera el conjunt. L'únic obstacle real era la **posició** — `L_STRUCT_BODY` només contenia posicions arquitectòniques, de manera que una pintura sobre la penya acabava en `ND`, que significa «posició no determinada» i no «posició no arquitectònica»: dues coses molt distintes col·lapsades en un valor, el mateix patró d'error que el domini 0/1/9 existeix per a evitar.

Tres entrades amb **`Level_Type = ROC`** ho resolen, i eixa columna passa a ser el **discriminador analític**: separar decoració arquitectònica de pintura rupestre és `WHERE Level_Type = 'ROC'` o el seu complementari, sense cap camp nou a la taula de decoracions.

**(e) `RockArt_Present`, `Color_Secondary`, `Sort_Order`.** El primer és el germà de `Dec_Present`, un judici agregat per a cada meitat de la taula, coherent amb les dues subseccions del formulari. El segon resol els casos bicroms amb els colors **junts o contigus dins d'un mateix motiu** (camp clar amb vora roja), on partir-los en dues files inventaria dos motius on n'hi ha un; segueix el patró de parella ordenada de `ID_Support` + `ID_Support_Secondary` i permet retirar el valor `Both`. El tercer restaura l'ordre constructiu bottom-to-top de la llista de posicions, que l'ordenació per `Level_Type` + `Name` havia deixat alfabètica dins de cada nivell i amb el ràfec al final.

**(f) Regles 27–31 i correcció del gating de l'element E.** Vegeu 6.1 i 9.2.

## **1.4. Canvis de la revisió final (v14)**

**(a) El domini de cinc valors es restringeix als 20 elements A–X.** El criteri de la v13 («capes aplicades a la fàbrica») era mal formulat. El que de veres habilita cada valor és:

> El **valor 3** exigeix que **l'evidència de la pèrdua siga de naturalesa distinta de la cosa perduda**: un encaix buit no és una mènsula, una empremta de morter no és morter.
> El **valor 2** exigeix **poder inferir l'extensió original**: en un mur es veu fins on arribava; en una superfície pintada, no.

Aplicat, el criteri exclou els cinc camps promocionats. **El pigment falla els dos:** l'única evidència de pigment és pigment, de manera que 3 i 0 es confonen necessàriament, i «present complet» no és afirmable mai. El **revoc**, igual — i a més `Plaster_Extent` i `Pigment_Extent` ja porten *Traces*, que era el que el 2 duplicava. El **morter** no és una capa que es perd per zones sinó un **atribut de la tècnica de fàbrica**: un mur en sec no ho és a trossos, i el que varia amb la conservació és la visibilitat, que registren `Facade_Observability` i `Mortar_Notes`. `Dec_Present` i `RockArt_Present` són **judicis agregats**, i cada motiu ja té la seua fila amb el seu estat.

**Criteri general definitiu — la distinció és discret vs continu:**

> **Cinc valors: els 20 elements A–X, i només ells.** Unitats constructives discretes, amb integritat física i extensió inferible.
> **Tres valors: tota la resta.** Atributs de tècnica, capes contínues, processos, vestigis mobles i judicis agregats.

**(b) `Cultural_Materials_Present`.** `ID_Material_Status` barrejava dues variables en una escala: *Good/Fair/Poor* són graus de conservació, però *Absent* és una afirmació de presència — **el mateix error que `Lintel` tenia en v10**. Es desdobla: camp de presència 0/1/9 anàleg a `Human_Remains`, i estat de conservació sense el valor *Absent*. La porta tanca tot 7.Mat inclòs l'estat. *Asimetria a retindre: `ID_Arch_Status` no es desdobla igual, perquè una estructura sempre existeix — és el registre mateix.*

**(c) `Doc_Basis` com a escala ordinal.** *Direct access / Close-range (<5m) / Medium-range (5-30m) / Long-range (>30m) / ND*. Els llindars van dins de l'etiqueta perquè el criteri s'aplique sense consultar cap manual. **La frontera entre els dos primers és física, no mètrica:** amb cordes s'està evidentment a menys de 5 m, de manera que el que distingeix l'accés directe és el **contacte**. La tècnica de captura no es registra ací: ja la descriuen les URL de 10.Doc.

**(d) `C14` governa el subformulari de datacions.** `F_DATING` es crea en v14: `T_DATING` existia des de la v1 però no tenia subformulari, i les datacions s'havien d'entrar obrint la taula. **Els segles queden sempre editables**, perquè la font habitual de l'atribució cronològica no és radiocarbònica sinó tipològica.

**(e) Totes les dimensions lineals en metres**, amb dos decimals forçats al control. No és cosmètic: una consulta pot sumar o comparar camps sense conversions, i dividir per 100 a mà és on s'introdueixen errors silenciosos. La **regla 31** marca qualsevol dimensió superior a 20 m, que és l'error d'entrar 62 pensant en centímetres.

**(f) `Ground` a `L_SUPPORT`.** Fa explícit un supòsit que les altres deu formes compartien en silenci: que l'estructura està **elevada sobre un buit**. `Ground` n'és la negació, amb criteri operatiu negatiu i comprovable —no hi ha buit a sota— i és **l'única forma sobre sediment i no sobre roca**: sobre roca l'estructura no assenta i no cal fonamentació; sobre sediment sí.

**(g) `Vertical association` i `Supra-structural unit`.** El primer és un tipus de connexió nou per a registrar la relació de verticalitat observable entre estructures sense afirmar-ne el mecanisme (la hipòtesi de politja o de suport d'escala continua a `T_ARCH_FEATURES`). El segon és un tipus de grup que treu la convenció Xa/Xb de la cadena del codi i la fa comptable.

**(i) Un volum i una superfície.** `Interior_Vol_m3` i `Total_Vol_m3` **no volien dir el mateix segons la tipologia**, cosa pitjor que no tindre'ls: en un mausoleu la diferència entre els dos *és* la fàbrica construïda, però en una cambra dins d'una cavitat l'«interior» és espai natural que ningú no va excavar i el «total» inclou roca, de manera que la resta no significa res comparable — i una mitjana sobre les dues tipologies hauria donat una xifra sense sentit. Es substitueixen per **`Area_m2`** i **`Volume_m3`**, on el volum és **l'espai funerari**: comparable entre totes les tipologies i el que es relaciona amb el MNI (H01). El desglossament fi —volum per cos, superfície de plataforma, volum construït— va a `Vol_Notes`.

**(j2) `Recessed_Frame` passa a 2.Arq.** És un **qualificador del pla de façana**, no del portal: el recul afecta el parament sencer i el portal hi queda inscrit. Gatejat darrere de `Sys_Portal` es bloquejava precisament allà on hi ha recul però no portal. **No es promociona a un `Sys_Facade`**: els cinc sistemes són combinacions de components identificables (E+F+G, N+O+Q, S+T) o camps d'agrupació, i una façana no és una combinació de res — és el pla on tota la resta passa, i no podria valdre *Absent* mai. Un sistema que no pot ser absent no fa el que fan els sistemes.

**(k) `Name_VAL` a tots els lookups.** El formulari mostrava uns desplegables en valencià i altres en anglés: no era un error puntual sinó **dos mecanismes, un d'incomplet**. Els combos de llista de valors porten les dues columnes a la cadena del codi; els de taula llegien `Name`. La correcció aplica el patró que ja existia —columna emmagatzemada oculta, etiqueta visible— als deu lookups que faltaven. Vegeu 4.0.

**(h) Sis camps de notes per pestanya** (`Arch_Notes`, `Finish_Notes`, `Condition_Notes`, `Bio_Notes`, `Materials_Notes`, `Systems_Notes`) i `QRY_18_Notes_Review`. `Color_Secondary` perd el valor `ND`.

## **1.5. Canvis de la iteració v15→v16**

*Aquesta iteració no naix d'una auditoria d'esquema sinó de **casos reals que la interfície no permetia registrar o registrava malament**. Això explica el seu perfil: molt de criteri operatiu, poca maquinària nova — cinc camps més, un menys, cap taula.*

**(a) Format i treball de la pedra.** `Masonry_Type` deia com s'apila la pedra i `Masonry_Quality` un judici d'execució; cap dels dos deia **quina pedra és**, que és el primer gest de la cadena operativa i el descriptor amb més capacitat de discriminar tradicions entre LP i DW. El vocabulari proposat inicialment barrejava **tres eixos** —morfologia natural (irregular, laminar), grau de treball (semiescairat, carreu) i dimensió (gran format)—, i en una llista única un mur de lloses laminars ben escairades no tenia valor possible. Es desdobla en `Stone_Format` (parella ordenada, sense `Mixed`) i `Stone_Working` (ordinal, amb `Mixed` i `ND` fora d'escala). *Large-format slabs* es descarta: el gran format viu als elements singulars, que ja tenen camp de material.

**(b) La façana és el pla exposat.** Una estructura del corpus té l'obertura al mur estret perpendicular al farallò i la decoració al mur llarg paral·lel, de manera que «façana» com a **pla exposat** i com a **pla compositiu** es van separar. Fixar-la per l'obertura costaria que `Facade_Orientation` deixara d'apuntar a la vall i que `Visibility_Valley` mesurara un pla que no es veu: tota H06 degradada per a salvar la definició d'un element. Es fixa pel pla exposat, L es redefineix sense referència a l'obertura, i la divergència es registra a **`Access_Plane`** — que no és una anomalia a absorbir sinó una variable: diu que la circulació mana sobre l'exhibició. **`Portal_Orientation`** no la duplica: el pla és relacional i intrínsec, l'orientació és absoluta i va a QGIS, i cap de les dues es dedueix de l'altra.

**(c) `Position_Relative` a `T_DECORATIONS`.** `Body_No` indexa el cos constructiu, que una pintura sobre la penya no té. No se substitueix per esquerra/dreta —són eixos distints i les files arquitectòniques continuen necessitant l'índex de cos—: s'hi afig un segon camp que serveix **les dues meitats** de 4.Dec i resol de passada la lateralitat de `FFL`, `JAM`, `PRT` i `RTW` amb un sol mecanisme. La convenció és indispensable: **esquerra i dreta des de l'observador situat davant del pla**, i el pla el declara `Access_Plane`.

**(d) `L_STRUCT_BODY`: −1 i +5.** `SOC` es retira perquè era circular — C es defineix com el *tractament decoratiu* del basament, de manera que «una decoració al sòcol» és «una decoració al tractament decoratiu del basament». Es van afegir `RTW`, `PRT`, `SIL` i `LIN`: v15 no tenia ni llindar ni dintell, i `Over-lintel` és la zona del fris (M), no el dintell, de manera que **un portal pintat s'havia de registrar contra un element que no era el pintat**. `PRT` existeix perquè una banda contínua que emmarca l'obertura és **un sol gest, no tres**. `RTW` respon un cas documentat: decoració al mur lateral i no a la façana, que és decoració no dirigida a l'espectador de la vall.

**(e) Les posicions ROC es redefineixen per contacte geomètric.** La v15 enviava «els tests 2 o 3» a `ROC` i el test 3 també a `ROC-PER`, de manera que qualsevol pintura que emmarcara una obertura **casava amb dues entrades** i la tria quedava a l'atzar de l'observador. El criteri nou és purament geomètric i s'aplica en ordre de més contacte a menys: **`ROC-OVL` (solapada) → `ROC-PER` (perimetral) → `ROC` (pròxima)**. La solapada és la categoria analíticament més rica de les quatre.

**(f) `Interbody_Cornice_Material` retirat.** No perquè de fet siguen sempre de pedra sinó perquè **no podrien no ser-ho**: un element horitzontal de fusta volat entre dos cossos és una mènsula (E) o una biga transversal (F), de manera que el valor *Wooden beams* descrivia **una plataforma mal classificada**. T i U ja porten «sempre pedra» a la definició sense camp de material: la cornisa era l'excepció incoherent.

**(g) X és un estat, no un element.** Definit com «cambra tancada per dalt amb la solució que siga», `Chamber_Roof` entrava a la matriu A–X igual si s'havia alçat obra que si una visera rocosa feia la faena. Amb molts sostres de penya natural al corpus, **Jaccard i les correspondències comptaven geologia com si fora construcció**. Definició nova: *sense contacte, no hi ha sostre*; i a `QRY_13`, X exporta **NA** quan el tipus és penya natural — NA i no 0, perquè no és una absència sinó una solució no construïda.

**(h) Criteri de diferenciació, generalitzat.** Un element A–X és present quan hi ha un component **físicament diferenciat** del parament contigu. Una obertura resolta deixant un buit al mur és N=0, O=0, Q=0 — i `Sys_Portal` present igualment. Vegeu 8bis.6.

**(i) Regles 32–39 i una correcció de bug.** `QRY_16c` es creava amb un nom (`..._22_30`) i es referenciava amb un altre (`..._22_31`), de manera que **la unió de la bateria no arribava a existir** i `MkQuery` degradava el fracàs a un missatge de depuració que ningú no llegia. Corregit, unió sobre quatre parcials, i `BuildDB` compta ara les consultes creades contra les previstes.

**(j) Ajornat a v17.** La descomposició per subnivells (N0, N1.1, N1.2 amb intranivells derivats) **no és vocabulari sinó cardinalitat**: exigiria una taula `T_BODIES` amb els elements A–X penjant de cada cos. La dada que ho ha de decidir —quants registres tenen cossos amb atributs realment divergents— encara no existeix.

## **1.6. Canvis de la iteració v16→v17**

*Onze punts de disseny, tres correccions de dades i un pedaç de formulari. La descomposició per subnivells (`T_BODIES`), ajornada des del delta anterior, **es torna a ajornar amb una diferència essencial: amb un comptador posat** (vegeu 1.6(h) i 8bis.10).*

**(a) `Position_Relative` eliminat; quatre camps de tram el substitueixen a la meitat rupestre.** El camp creat en v16 servia les dues meitats de 4.Dec, i **les dues meitats no comparteixen referent**: a les files arquitectòniques *Esquerra* designava **quina instància d'un parell** portava el motiu —una relació identificacional, interna a l'estructura—; a les files ROC designava **on està la pintura** en relació amb el volum construït —una relació topològica entre dos objectes separats. Que compartisquen les paraules *esquerra* i *dreta* és una coincidència del llenguatge, no una propietat comuna. Tres conseqüències eren errors vius: `Above` i `Below` no tenien referent possible a la meitat arquitectònica, on l'eix vertical el diuen ja `ID_Struct_Body` i `Body_No`; `Both` significava coses distintes a cada meitat i la documentació només en definia una; i el pla de referència declarat (`Access_Plane`) **invertia esquerra i dreta** en tota estructura amb l'accés desviat de la façana. A més, `PRT` no havia de ser lateralitzable: existeix perquè una banda contínua és *un sol gest, no tres*, de manera que un `PRT` amb costat es contradiu a si mateix.

*Evidència del corpus: 7 files de 56 portaven valor, **totes ROC, cap arquitectònica**, i en tres d'elles `Both` s'estava usant per a dir una **U invertida** amb la descripció refugiada a `Notes`. El camp no és que li faltara un valor: n'estava dient un altre. Aquest és també el motiu que la transferència no puga convertir-lo automàticament: traduir `Both` a esquerra+dreta afirmaria que el tram superior és **absent** precisament a les files on hi és, i un 0 és un judici mentre que un buit no ho és.*

**A la meitat arquitectònica no se substitueix per res.** Cap fila del corpus usa la lateralitat sobre posicions arquitectòniques, i la revisió va confirmar que `QUO` i `PIL` no la demanen —cosa que resol de passada el problema de les pilastres múltiples, on *Esquerra/Dreta* hauria deixat de funcionar. Crear un camp sense cap ocurrència en 56 files contradiria el criteri que va limitar les posicions ROC a quatre. La lateralitat va a `Notes` fins que aparega un cas real.

**(b) `Outline_Geometry`.** La banda perimetral en U invertida no és una propietat sinó quatre: que ressegueix el contorn (ja ho diu `ROC-PER`), que sembla una banda (ja ho diu el tipus), **quina part del contorn cobreix** (els quatre trams) i **amb quina geometria de traç** (aquest camp). No es filtra per tipus de decoració perquè les tres files que descriuen una U invertida porten **tres tipus diferents**, de manera que qualsevol filtre les deixaria fora precisament a elles. No és un descriptor accessori: un contorn ortogonal **afirma un referent construït rectangular** i un de corb no afirma res, i és per tant l'única evidència disponible per a decidir el punt (d).

**(c) Vocabulari de suports: `Rock dihedral`, `Rock surface` i un mode nou.** Vegeu 8bis.1 i el catàleg `L_SUPPORT`. El **díedre** és el buit que deixa un bloc en tascó despresa al llarg de dues discontinuïtats que s'encreuen; s'adopta el terme descriptiu i no el genètic, coherent amb el criteri geomètric de les posicions ROC. La **superfície de roca** entra amb `Support_Mode = Substrate`, **un mode nou**, perquè cap dels tres existents descriu el que fa: una superfície que sosté pigment no és descans gravitacional, ni confinament, ni encastament, i la categoria pròpia la manté fora de tot recompte de suport constructiu. Criteri escrit, **orientatiu i no validat**: el suport principal respon a les forces verticals, el secundari a les horitzontals.

**(d) L'estructura desapareguda és un registre d'estructura.** No era una pregunta oberta sinó una **incoherència viva**: la descripció de la tipologia `Unclassifiable` nomenava DW-S01-EA11 com a exemple d'estructura construïda no classificable, i el registre estava desat com a `PR Rock Art`, amb una fila en posició `ROC-PER` — que no pot existir en un registre d'art rupestre, perquè significa «ressegueix el contorn de l'estructura» i un panell aïllat no en té cap. Vegeu 8bis.9 per a la composició del registre i el criteri de frontera.

**(e) Connexions: l'ordre deixa de significar res.** El corpus contenia dues files i **eren la mateixa connexió**, entrada des de cada extrem. La causa era estructural i no descuit: `F_CONNECTIONS` s'enllaçava només per `ID_Struct_A`, de manera que des de B no es veia res. `Chrono_Relation` passa a **designar quina estructura és l'anterior**, el formulari normalitza la parella —l'ID menor sempre a A— **girant també la cronologia** quan les inverteix, i un índex únic fa la duplicació impossible en lloc d'una cosa a vigilar. Desapareix la instrucció «no invertir l'ordre després de l'entrada». S'afig `Horizontal association`, simètrica de la vertical.

**(f) `ID_Material_Status` baixa a 7.Mat.** La porta (`Cultural_Materials_Present`) vivia a 7.Mat i el camp que governa a 5.Estat, de manera que es preguntava **en quin estat estan abans de si n'hi ha** — i la contradicció era indetectable, perquè el gating de nivell 2 opera dins d'una pestanya. 5.Estat queda per a l'estructura construïda; 7.Mat, per al seu contingut, amb la mateixa forma que 6.Bio. Les tres capes **no són redundants**: la presència és contestable des de trenta metres, el tipus exigeix identificar-los i l'estat exigeix veure'ls prou bé. Regles 43 i 44.

**(g) Frontera del fris i tres tipus retirats.** `Relief_Frieze` (M) és **l'afirmació agregada** que hi ha fris; les files diuen **quin motiu és**. No són alternatives: es posen els dos. El precedent del sòcol no s'aplica directament, perquè els frisos chachapoya **són** la fàbrica —el ziga-zaga i l'escalonat es fan col·locant les pedres endins i enfora—, de manera que «la fàbrica a l'element i el motiu a la fila» perdria el repertori que H01 i H04 comparen. Regla 42, unidireccional com la 34, i mira `OVL` **i** `FFL`. Es retiren de `L_DEC_TYPE`, tots tres sense cap ús al corpus: *Frieze / Greca* (l'única entrada que no nomena cap motiu — diu el que ja diu M, i `ND` cobreix el cas no resolt), *Triangular motif* (mai observat com a distint del ziga-zaga) i *Decapitation scene* (un sol cas, i **art rupestre**, no decoració arquitectònica: va a `RA Anthropomorphic` amb la lectura a `Notes`).

**(h) Gating, «no aplicable» i el sisé sistema.** `Access_Plane` i `Portal_Orientation` es desactiven **quan `Sys_Chamber` es declara** Absent o No aplicable —mai amb el camp buit, que vol dir «encara no avaluat»—, i es gategen per la **cambra** i no pel portal, perquè gatejar-los pel portal trencaria la decisió v16 que una obertura arrasada conserva orientació coneguda. `Facade_Orientation` i `Visibility_Valley` **es queden obertes**: v16 va definir la façana com el **pla exposat**, deslligat de l'obertura, i una massa basal en té i mira cap a algun lloc; tancar-les eliminaria de H06 el seu grup de control natural. Criteri general: **deduïble → gating; judici → valor «no aplicable»**. `Sys_Interface` s'afig com a sisé sistema i passa el test que va rebutjar un `Sys_Facade` en v16 —una façana no pot ser absent mai, una interfície sí—; governa **només** `Interbody_Cornice` (I), perquè amb un sol cos no hi ha cornisa entre cossos però el coronament (R) continua existint.

**(i) `Fabric`: el comptador que substitueix `T_BODIES`.** Vegeu 8bis.10.

**(j) Punts heretats tancats.** `Portal_Position` (descentrar un portal és una decisió de planificació, no un accident de fàbrica: H01 i H04). `Chinking_Stones` es manté **sense gradació**, amb la nota que amb 26 presents i cap absent no pot aparéixer en cap resultat mentre no hi haja variància. El llindar mètric de volada per a la distinció G/I **ja estava tancat** en v16 (no se'n fixa cap; la profunditat es mesura sobre el model a la segona passada): era una decisió presa i no escrita al lloc que toca. Coordenades pròpies a `T_DECORATIONS`: descartades. Caracterització de la superfície rocosa a les files ROC: **descartada** —«no observable» ja és una observació completa— i la variant útil (descriure-la com a *objecte* en registres d'art rupestre aïllat) no té cap cas al corpus.

**(k) Correccions que no són de disseny.** Vegeu 9.3.

## **1.7. Canvis de la iteració v17a→v18**

*El primer delta nascut íntegrament de l'entrada de dades real: 35 estructures entrades, i cada punt ve d'una nota escrita davant d'un registre concret. El detall complet, amb la discussió de cada decisió, és a `DELTA_v17_v18.md`; ací, el resum executiu.*

**(a) Tres camps nous a `T_STRUCTURES`.** `Platform_Surface` (element **Z**, §2.4): la plataforma obté la seua superfície, el paral·lel exacte de T al ràfec — 21é camp del domini de cinc valors, gatejat per `Sys_Platform`. `Interbody_Cornice_Format` (§2.4b): laminar contra tabular en la cornisa, el patró de `Lintel_Material`. `Sill_Coincides_Cornice` (§2.5): **compartició d'element** — la cornisa intercòs que fa de llindar, el cas que el gradient no pot dir.

**(b) `Tie_Walls` (D) desgatejat de `Sys_Base`.** Apareix desvinculat de la massa basal — sobre plataformes volades, com a suport adossat, o aïllat: com I i R, queda permanentment actiu i el seu 0 és sempre una asserció. `Sys_Base` governa A, B, C. *(Justificació reescrita en v19, delta C3: la formulació anterior, «a nivell de mur», es podia llegir com «formant part dels murs de la cambra», i cap D del corpus tanca cambra.)*

**(c) Gating nou derivat de l'entrada.** `Recessed_Frame` es gateja per `N_Chamber_Bodies = 0` (sense cos de cambra no hi ha façana — regla 47); `Portal_Orientation` es tanca amb `Access_Plane = 'Facade'` (una sola dada, dos camps: `QRY_07` exporta la columna efectiva); la tipologia `EA-TER` tanca façana i portal sencers (regla 48); i `ID_Support` es filtra per tipologia, amb autoompliment de NIX i PR (regla 54).

**(d) La direcció de les connexions ix de la parella.** `Chrono_Relation` passa a tres valors (*Sequential / Contemporary / Undetermined*) i el camp nou **`ID_Earlier`** nomena l'estructura anterior per FK. El gir de normalització ja no toca la cronologia, i `F_CONN_IN` deixa d'invertir res: llig el fet emmagatzemat (regla 52). Tipus nou `Associated natural context` per al cas terrassa + nínxol natural.

**(e) `T_LOST_ELEMENTS.ID_Position` eliminat.** Un valor omplit en 27 files: la columna feia una pregunta que ningú no responia. El valor migra a `Notes`; `Evidence_Scope` es queda sota la regla 53 (mirall de la 21), que el corpus violava en 16 files.

**(f) Maçoneria.** `'Tabular blocks'` → `'Regular tabular blocks'` (canvi de valor emmagatzemat, migrat una sola vegada al patch) i **`'Large blocks'`** entra a les dues llistes de format: una peça que fa filada per ella mateixa, el senyal d'inversió de treball que el test funcional llig.

**(g) Correccions pròpies.** `QRY_18` llegia `E.Notes`, desaparegut en 17a (ara escombra els 14 camps de notes); `QRY_13` no exportava `Sys_Interface` i gatejava `Recessed_Frame` per `Sys_Portal` contra la decisió v17. `QRY_23_V18_Review` és la llista de treball de la migració.

**(h) Criteris de manual, sense esquema.** EA-PLA-V amb tots els components a 0 és legal; el fons de cambra es defineix contra el **pla de façana**, mai contra l'accés; llindar mínim analitzable = els 21 elements + els 6 `Sys_*` sense NULL.

## **1.8. Canvis de la iteració v18→v19**

*Delta tancat sobre els tres candidats que la secció Pendent del delta anterior ja especificava (criteri dels tres casos) més l'acumulat de les sessions d'entrada. El detall complet és a `DELTA_v18_v19.md`; ací, el resum executiu.*

**(a) `Jamb_Fabric_Reveal` (§2.5b).** Segon qualificador de compartició: la posició del brancal resolta per la fàbrica mateixa — cara terminal acabada i deliberada (*masonry reveal*), sense element diferenciat. Finestra `Jambs ∈ {0, 2}` (amb O = 2 resol el cas que el manual declarava indecidible al Nus 1); la regla 55 vigila la finestra i la branca de portal de la regla 3 queda exempta mentre el qualificador està declarat. El forat equivalent de la plataforma **es queda obert a propòsit** (v18, criteri (h)): fals positiu conegut i acceptat.

**(b) `DOM3Q`.** Els dos qualificadors de compartició porten etiquetes de pregunta (*No / Sí / No observable*) en lloc de les de presència; valors emmagatzemats idèntics (0/1/9). El projecte té des d'ara **dos vocabularis d'etiqueta per a un mateix domini**: DOM3 per a presències, DOM3Q per a preguntes.

**(c) `Timber_Bracket_Role`.** La regla 9 dispara només amb `Is Null`: **ND és un judici** («avaluat, no decidible»), no una casella buida — la distinció NULL/9/0 traslladada al camp de text. La regla 56 nova vigila la coherència (rol de suport amb `Sys_Platform` negat; la direcció inversa ja era R7); les etiquetes UI es reescriuen al voltant de l'única pregunta del camp — el **vincle** amb el sistema plataforma — amb els valors emmagatzemats intactes. L'arbre de tres preguntes és al manual.

**(d) MEN promogut i D reanomenat.** La tipologia passa a `MEN Isolated Structural Element`, amb descripció ancorada al vocabulari A–Z i sense subtipus (la identitat de l'element la diuen els camps de 11.Sist: D present = muret, E present = mènsula). El terme de referència de D passa a **`Transverse wall`** («tie» afirma la funció de trava, i la funció és precisament el que no sabem); el nom de camp `Tie_Walls` **no es toca** (renom de camp = consultes trencades, lliçó v11). La llibertat d'ancoratge es diu amb **alçada**, mai amb *nivell*.

**(e) Correcció dels zeros de D.** El patch v18 va deixar com a assercions els zeros que sota el gating v17 eren farciment: el `PatchV19` els retorna a NULL (6 files al corpus) i `QRY_24` els llista per al judici real. Els zeros sota `Sys_Base = Present` no es toquen automàticament: el camp hi era obert i es confirmen amb el test de delimitació.

**(f) Manual.** Test de delimitació D/V; dues línies MEN (com es registra; frontera amb l'estructura atestada); mecanisme d'atestació de sistemes perduts per files d'abast *Body/Whole*; i la línia que faltava sobre comptadors: **inclouen els cossos en *Attested lost* amb fila d'evidència**.

## **1.9. Canvis de la iteració v19→v20**

*Quinze blocs tancats (quatre de la primera tanda + els onze punts A–K) més el tancament de tots els pendents històrics. Detall complet a `DELTA_v19_v20.md`; ací el resum executiu.*

**(a) Identitat de codis a nivell de motor.** Índex únic sobre `Code` (bloqueja, no alerta: la identitat no és cap judici); `Sector_Code` nova a `L_SECTORS` i **regla 57** de concordança codi–sector (el cas EA46, capturat estructuralment); número d'EA per expressió de consulta (mai camp: duplicaria el codi); `QRY_26_Next_EA` fa visible el criteri **màxim+1, mai reutilitzar buits**. La gramàtica del codi (JACIMENT-SECTOR-EAn) queda formalitzada a §2.1.

**(b) Les dues famílies del zero (manual + etiquetes).** En elements constructius el 0 afirma una decisió constructiva sobre superfície llegible; en vestigis peribles afirma només supervivència observable. `DOM5` ho diu ara al desplegable (*Absent (constatat)*, *Desaparegut atestat*) i les pestanyes de vestigis porten la bandera epistèmica. Els elements de fusta (A, E, F, S) pertanyen a la família forta: la fusta encastada deixa encaixos.

**(c) `Upper_Crown_Format`.** La forma del remat (tancament vertical / filera en voladís / filera a ras / ND), regla 58 amb semàntica de llista de treball. Quina cara de R corona es **deriva** del registre i per això no es pregunta. Substitueix (i aprén de) el `Interbody_Cornice_Format` retirat amb la regla 50: aquest naix amb interés analític declarat, població revisable (12) i observació no derivable.

**(d) Bandes de contorn desacoblades del veredicte.** Els dos tipus passen a `Outline band, defined / amorphous` (el traç); la topologia U/C/O es **deriva dels quatre trams** (`QRY_27_Outline_Topology`); el veredicte estructura-vs-panell és el judici d'atestació complet a nivell de registre — a La Petaca hi ha estructures perimetrades amb C i amb O, així que cap propietat de la banda no el decideix sola (§8bis.9 reescrita). **v21 (bloc 3, errata v20): el filtre estàtic dels combos deixava les bandes inabastables des de la meitat rupestre** — el raonament v20 confonia la classe panell amb tota la meitat ROC. Correcció al formulari: el combo rupestre les incorpora, l'arquitectònic les exclou, i la classe *Rock art panel* restringeix la llista a RA* pel patró per-classe del bloc H. Cap valor emmagatzemat canvia.

**(e) Gates per comptador.** Els quatre camps de portal de 2.Arq es tanquen amb `N_Chamber_Bodies = 0` (regla 59), absorbint les condicions per tipologia; l'evidència de fase s'obri només amb fases ≥ 2 (regla 60), **sense cap default de fases** (el 1 és un judici); la maçoneria **no** es gateja per murs — 13 registres sense murs porten fàbrica real — **però sí (v21) per la declaració de fàbrica**: `Masonry_Present` a 0 o 9 tanca el bloc sencer, presència mana sobre el detall. La classe *Structural trace* obri 2.Arq (un muret té aparell); la classe rupestre tanca l'observabilitat interior i autoompli la posició a *Panell*.

**(f) Regla 6 i col·lapsades.** Cap conversió automàtica 0→9: la llegibilitat es jutja superfície a superfície (registre model: DW-S01-EA13). El text d'acció anomena la tercera eixida (confirmar la llegibilitat a notes). L'avís del formulari baixa a baix de tot.

**(g) `T_ARCH_FEATURES`.** Fora la casella *Present* (duplicava l'existència de la fila i el seu 0 feia mentir 11 de 16 files); el catàleg guanya *Encaix / interfície negativa* (3 casos) i *Banqueta d'accés* (2 casos — l'instrument de recompte de l'element Y, reservat des de v17).

**(h) Validació al punt d'entrada.** Botó *Valida aquest registre* sobre la bateria filtrada pel codi actual (`F_VALIDATION`). Cap codi de colors per camp: duplicaria les 58 regles en lògica de colors.

**(i) Tancament de tots els pendents.** `T_BODIES` no es descompon dins del TFM; les superfícies ROC es caracteritzen per la via 3D; el forat de R3 sobre plataforma és teòric (0 casos en 86); la revisió v17a està completa; el rebuild net passa a milestone del depòsit amb checklist; l'element Y té la mini-especificació pre-decidida (2 de 3 casos). Les recodificacions a/b → codis propis queden a la taula de concordança de l'apèndix.

# **2. T_STRUCTURES (145 camps)**

## **2.1. Identificació (8)**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| ID | COUNTER PK | Clau primària autonumèrica |
| Code | TEXT(20) NN | Codi PALP: `JACIMENT-SECTOR-EAn` — ex: DW-S01-EA01. **v20 (bloc B): gramàtica formalitzada i vigilada.** Índex únic (el motor rebutja duplicats en teclejar-los); el prefix ha de coincidir amb el `Sector_Code` del sector triat (regla 57); numeració **màxim + 1, mai reutilitzar buits** — els forats de la seqüència són història de campanya (`QRY_26_Next_EA` diu el següent lliure). ~~El número d'EA no s'emmagatzema mai a banda~~ **REVERTIT v23 (bloc 5)**: vegeu `EA_Number` |
| **EA_Number** | INTEGER | **NOU v23 (bloc 5). REVERSIÓ EXPLÍCITA de la decisió v20** («mai a banda»): el número d'EA s'emmagatzema, **derivat del codi** — ordenable i enllaçable amb el flux mètric de Metashape, el cas d'ús que la decisió vella no tenia davant. El patch el deriva per a tot el corpus; `Code_AfterUpdate` el recalcula a cada edició; el control d'1.Id està **bloquejat**. La **regla 67** respon la preocupació fundacional: dispara si falta o no concorda amb el codi |
| **ID_Subsector** | LONG | **NOU v23 (bloc 6).** FK → **L_SUBSECTORS** (nova: vocabulari posicional tancat, transversal als sectors — Upper/Lower/North/Central/South amb etiquetes valencianes). **NULL legítim sempre** («cap valor» és estat final) → etiqueta «(opc.)» pel criteri v21; fora de tota worklist. Relació 26 (`REL_SSEC_STR`) |
| ID_Sector | LONG NN | FK → L_SECTORS (que porta `Sector_Code` des de v20: LP-G/N/C/U/S, DW-G, DW-S01…DW-S06 — els futurs codis de La Petaca seran `LP-N-EA01` etc.). El subsector (v23) és camp propi, no part del codi |
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
| Floor_Plan | TEXT(20) | Planta: Rectangular / Sub-rectangular / Square / Circular / Sub-circular / Trapezoidal / **Triangular (v23)** / Irregular / ND |
| N_Built_Walls | INTEGER | **REDEFINIT v23 (bloc 4): murs DE LA CAMBRA** (façana + retorns; els murets basals D no s'hi compten mai). La pregunta ja no és *quants murs hi ha* sinó *quants murs té la cambra*. Cambra absent/NA → **0 derivable** (autoomplert, camp tancat); cambra present/atestada → obert i **mínim 1** (regla 65: una cambra té almenys la façana); cambra no observable/no avaluada → NULL. La regla 66 denuncia murs comptats sense cambra. Etiqueta: «Núm. murs (façana + retorn)» |
| **Metrics_Available** | BYTE | **NOU v22 (bloc B). EL PORTER MÈTRIC.** 0/1/9, default NULL — *hi ha dades mètriques disponibles per a aquest registre?* El 9 vol dir *no mesurable mai* (sense accés ni cobertura de model); el NULL, *encara no decidit*. **Gateja** dimensions, cota sobre la base, mètode, obertura (combinat amb el gate de portal), suport, volumetria i notes de volum; les **coordenades i `Metric_Notes` queden fora** — una estructura sense res mesurable continua tenint posició. **Fora de la regla 17 i de tota worklist a propòsit** (divergència raonada respecte de `Masonry_Present`): la disponibilitat no és decidible registre a registre fins que el flux d'extracció Metashape estiga en marxa. La regla 63 vigila el costat del detall |
| Length_m / Width_m / Height_m | SINGLE | Dimensions exteriors (m). Segona passada mètrica. **v22: gatejades pel porter mètric** |
| Height_Above_Base_m | SINGLE | Cota de posició sobre la base del faralló (m): és una cota, no una dimensió |
| Dim_Method | TEXT(30) | Photogrammetric model / Tape measure / Laser / Estimation / perimeter pigment outline / ND |
| Opening_Width_m / Opening_Height_m | SINGLE | Dimensions de l'obertura d'accés (m). **No gatejades per Sys_Portal**: la passada mètrica és un exercici separat |

*L'antic Lost_Body_Evidence desapareix: la seua informació és ara una fila de T_LOST_ELEMENTS amb abast Body (secció 3.7). Lintel es trasllada al bloc del sistema portal (2.5).*

## **2.2b. Detall del suport geològic (3) — H02**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| Support_Width_m | SINGLE | Amplària de la repisa o suport (m) |
| Support_Depth_m | SINGLE | Profunditat de la repisa o cavitat (m) |
| Support_Modified | BYTE | 0/1/9. Modificació antròpica del suport natural. Sovint oculta darrere de la maçoneria |

## **2.2c. Maçoneria, morter i fàbrica (11) — H01/H04, cf. Toyne i Anzellini 2017**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| **Masonry_Present** | BYTE | **NOU v21 (bloc 2). LA DECLARACIÓ DE FÀBRICA.** 0/1/9, default NULL (camp de judici, tres estats epistèmics). La classe no la deriva — una mènsula MEN no té fàbrica, un muret sí (DW-S01-EA55) — així que pel criteri v17 el judici té camp propi, família de `Plaster_Present`/`Pigment_Present`. **Gateja el bloc sencer** (qualitat, formats, treball, aparell, morter, ripio, notes i `Fabric`); regla 61 vigila el full de dades; entra a la família de capes i la regla 17 la llista. Revoc i pigment NO s'hi encadenen (decisió v21: es reobri amb contraexemple). **v22 (bloc A): fa de porter de 2.Arq per als contextos naturals** — la pestanya s'obri per a la classe (una cavitat pot dur fàbrica i cossos basals: cas de camp fotografiat) i amb 0/9 el bloc es tanca sencer; **als panells el 0 és derivable de la classe** i s'ompli sol (patch + canvi de tipologia), mai sobre un valor declarat |
| **Stone_Format** | TEXT(30) | **NOU v16; llista revisada v18.** Irregular stones / **Regular tabular blocks** / **Large blocks** / Laminar slabs / ND. Morfologia de la peça, producte de la fractura natural de la roca. **Abast: el parament**, no els elements singulars. *Regular* va entrar al valor emmagatzemat quan va arribar *Large blocks*: un bloc gran també és tabular, i el terme nu va deixar de nomenar una classe |
| **Stone_Format_Secondary** | TEXT(30) | **NOU v16; llista revisada v18.** Mateixos valors sense `ND`. Segon format convivint al parament (parella ordenada com `ID_Support`) |
| **Stone_Working** | TEXT(20) | **NOU v16.** Unworked / Semi-dressed / Dressed / Mixed / ND. **Ordinal**: `Mixed` i `ND` són fora d'escala i cauen de correspondències i regressions, com els *Not applicable* |
| Masonry_Quality | TEXT(20) | Good / Moderate / Poor / ND |
| Masonry_Type | TEXT(30) | Well-coursed / Irregular-coursed / Uncoursed / Mixed / ND |
| Mortar_Present | BYTE | 0/1/9. 0 = fàbrica en sec verificada. **No promocionat (v14)**: no és una capa que es perd per zones sinó un atribut de la tècnica — un mur en sec no ho és a trossos. El que varia amb la conservació és la visibilitat, registrada a `Facade_Observability` i `Mortar_Notes` |
| Mortar_Type | TEXT(30) | Mud / Mud with gravel / Mud with organics / None dry-laid / ND. Amb Mortar_Present = 0 el gating l'autoassigna a None dry-laid (regla 26) |
| Chinking_Stones | BYTE | 0/1/9. Ripio / falques entre carreus. Independent del morter: una fàbrica en sec pot dur falques. **No promocionat a cinc valors (v13)**: al corpus v12 presentava 26 presències i cap absència verificada, de manera que no té variància i no pot alimentar cap prova. Es manté com a descriptor tècnic (T&A 2017); la seua universalitat és ella mateixa un resultat sobre la tradició constructiva de DW. **Revisat v17**: 26 presències i cap absència sobre 36 registres — es manté **sense gradació**, amb la conseqüència anotada que no pot aparéixer en cap resultat mentre no hi haja variància |
| Mortar_Notes | TEXT(150) | Notes sobre morter i juntes |
| **Fabric** | TEXT(25) | **NOU v17.** Single / Between bodies only / Within a body / Not observable. **v21: gatejat per `Masonry_Present`** — la nota v17 continua vàlida (no es gateja PER COSSOS); el seu gate és la declaració de fàbrica mateixa. Vegeu 8bis.10 |
*Els tres camps de pedra van **davant** de `Masonry_Type` al formulari: es tria la pedra abans d'apilar-la, i eixe és l'ordre de la cadena operativa.*

***`Mixed` canvia de sentit en v17 i cal escriure-ho.*** *A `Masonry_Type` i `Stone_Working`, `Mixed` diu que **el registre té més d'un valor**; `Fabric` diu **si eixa pluralitat coincideix amb la divisió en cossos**. Sense la distinció escrita, tornarien a ser dos camps que semblen dir el mateix.*

*Test operatiu de `Stone_Format`, deliberadament **funcional** perquè es llig a la façana sense mesurar res: **irregular** = la peça no té dues cares planes subparal·leles dominants; **tabular regular** = dues cares planes i **alçada de filada pròpia** (una pedra, una filada); **de grans dimensions** (v18) = **una sola peça fa la filada per ella mateixa**, sense companyes al mateix nivell — el senyal d'inversió de treball més fort que el test pot llegir; **laminar** = dues cares planes però **cal apilar-ne diverses per fer una filada**. Desempat quantitatiu per als dubtosos: gruix ≤ ¼ de la dimensió major → laminar. **Dominància per superfície de parament, no per nombre de peces**: amb lloses menudes i blocs grans, comptar peces inverteix el resultat.*

*`Mixed` **no existeix** a `Stone_Format`: seria un tercer camí per a dir el que ja diu la parella, i el projecte ja ha retirat dues vegades el valor que la parella substitueix (`Combined` en v8, `Both` en v13). El secundari no porta `ND` — buit vol dir «cap segon format» (regles 32–33).*

*`Stone_Working` registra **evidència positiva de treball**: traces d'eina, aristes vives regulars, cares que trenquen el pla de fractura natural. Sense eixes traces la cara plana s'atribueix a la fractura i el valor és `Unworked`, que així no afirma «ningú no la va tocar» sinó «no hi ha evidència que la tocaren»; `ND` queda per al dubte genuí — un gres estratificat fractura en cares indistingibles del carejat.*

*Cap dels tres porta valor «no observable», com cap altre camp de domini TEXT de l'esquema: **el 9 existeix per a protegir el 0, i ací no hi ha 0 a protegir** — un parament sempre està fet d'alguna pedra amb algun grau de treball. Els dos motius de no-decisió no es col·lapsen tot i compartir `ND`: «no s'hi veia» i «no era decidible» els separa l'encreuament amb `Doc_Basis` i `Facade_Observability`, que ja registren la visibilitat. És el mateix criteri que va deixar `Mortar_Present` sense promocionar.*

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| **Recessed_Frame** | BYTE | 0/1/9. Marc reculat de façana. **Traslladat ací en v14** (abans al bloc del portal): és un qualificador del **pla de façana** i no del portal — el recul afecta el parament sencer i el portal hi queda inscrit. **v18: gatejat per `N_Chamber_Bodies = 0`** — sense cos de cambra no hi ha façana a qualificar (regla 47), i `QRY_13` l'exporta nul en eixe cas. La tipologia `EA-TER` també el tanca (regla 48) |

## **2.3. Sistemes constructius (7 camps de sistema)**

**Domini dels tres sistemes d'element (Sys_Platform, Sys_Portal, Sys_Eave):** *Present complete / Present partial / Attested lost / Absent / Not applicable / Not observable*. **Domini dels tres conjunts d'agrupació (Sys_Base, Sys_Chamber, Sys_Interface):** *Present / Absent / Not applicable / Not observable*. Per defecte: Absent.

| **Camp** | **Elem.** | **Governa (gating)** | **Descripció** |
| --- | --- | --- | --- |
| Sys_Platform | **H** | E, F, G, **Z (v18)** + Timber_Bracket_Count/Role, Platform_Surface_Material, Platform_Function | Sistema plataforma: E+F+G+Z → H |
| Sys_Portal | **P** | N, O, Q + Lintel_Material | Sistema portal: N+O+Q → P |
| Sys_Eave | **U** | S, T | Sistema ràfec: S+T → U. Sempre pedra a T i U |
| Sys_Base | — | A, B, C | Conjunt basal (agrupació sense lletra pròpia). **v18: D ix del conjunt** — el corpus el mostra a nivell de mur, i com I i R queda permanentment actiu |
| Sys_Chamber | — | J, K, L, M, V, X + Chamber_Roof_Type, Rear_Closure_Type, **`Access_Plane` i `Portal_Orientation` (v17)** | Conjunt cambra (agrupació; absorbeix l'antic W) |
| **Sys_Interface** | — | **I** (el format de cornisa, v18–v19, es retirà en v20) | **NOU v17.** Conjunt interfície. Passa el test que va rebutjar un `Sys_Facade` en v16: una façana no pot ser absent mai, una interfície sí — una estructura d'un sol cos no en té. **`Upper_Crown` (R) queda fora**: amb un sol cos no hi ha cornisa entre cossos, però l'estructura té part de dalt igualment |
| Platform_Function | — | — | Access / Circulation / Construction / Support / Multiple / Undetermined. Forma separada de funció: OE3 existeix per a determinar per a què servia la plataforma, i la morfologia no ho ha de pressuposar |

*Els sistemes es resolen primer (pestanya 11.Sist els presenta dalt de tot): es decideix que hi ha plataforma abans de comptar-ne les mènsules. Un sistema obri el seu grup només en Present* o Attested lost; sota Absent / Not applicable / Not observable els components porten el 0 de farciment.*

## **2.4. Elements A–Z — Nivell 0 / basament (11)**

**Domini dels camps d'element: 0 = Absent / 1 = Present complet / 2 = Present parcial / 3 = Desaparegut atestat (exigeix T_LOST_ELEMENTS) / 9 = No observable. Per defecte 0.** *v18: el vocabulari es tanca en A–Z amb l'entrada de Z; la lletra Y queda **reservada** per a la massa adossada (banqueta), pendent dels seus tres casos documentats.*

| **Camp** | **Tipus** | **Elem.** | **Descripció** |
| --- | --- | --- | --- |
| Embedded_Base_Beams | BYTE | **A** | Jàcenes horitzontals empotrades dins la maçoneria del basament |
| Base_Level | BYTE | **B** | Basament com a element constructiu diferenciat (pòdium) |
| Decorative_Socle | BYTE | **C** | Tractament decoratiu del basament |
| Tie_Walls | BYTE | **D** | Murets perpendiculars al faralló (ancoratge/compartimentació). **v18: desgatejat de `Sys_Base`** — sempre actiu i el seu 0 és sempre una asserció. **v19: terme de referència `Transverse wall`** («tie» afirma la funció, que és el que no sabem); ancoren a qualsevol **alçada** de la fàbrica (mai «nivell»: N0/N1 es defineixen pels cossos i un muret no és cos de res); no tanca cap interior ni es compta com a cos — si tanca cambra és V (test de delimitació al manual). El `PatchV19` retorna a NULL els zeros de farciment heretats del gating v17. **v20 (bloc 1): terme eixamplat a `Transverse wall or pier`** — del muret llarg d'ancoratge al piler compacte contra la roca; la longitud no canvia la lletra. Tests de frontera al Nus 11 del manual: contra E (la mènsula penja, el piler s'alça) i contra K (la pilastra viu al pla del mur, el piler contra la roca) |
| Timber_Brackets | BYTE | **E** | Mènsules de fusta empotrades a la roca (component de H). Criteri: perpendicular a la façana, encastada, en voladís. Exempta de la regla 1: és l'únic element amb existència independent del seu sistema |
| Timber_Bracket_Count | INTEGER | **E** | Nombre de mènsules visibles. Obligatori amb E present (regla 8): el nombre porta l'argument de la plataforma |
| Timber_Bracket_Role | TEXT(25) | **E** | Platform support / Isolated / Both / ND. Obligatori amb E present (regla 9, **v19: dispara només amb NULL — ND és un judici complet**, arbre de tres preguntes al manual). Isolated bloqueja Platform_Surface_Material (regla 7); rol de suport amb `Sys_Platform` negat és la regla 56. **v19: etiquetes UI reescrites al voltant del vincle amb H**, valors emmagatzemats intactes |
| Transverse_Beams | BYTE | **F** | Bigues transversals (component de H). Criteri: paral·lela a la façana, salvant llum |
| Corbelled_Courses | BYTE | **G** | Filades de pedra en voladís creixent (component de H, variant lítia) |
| **Platform_Surface** | BYTE | **Z** | **NOU v18; RESERVA v21 (bloc 1).** Superfície transitable acabada de la plataforma — el paral·lel exacte de T al ràfec. **Z queda reservat a la superfície del sistema VOLAT** (lloses sobre mènsules E o sobre biga F): el cap de la massa basal d'una terrassa NO és Z i no obri `Sys_Platform` — la tipologia EA-TER ja respon la pregunta d'ús, i l'única pregunta real sobre un N0 sense N1 és el remat (R). Component de H, gatejat per `Sys_Platform`; pot sostenir el sistema tot sol (regla 3) **només en el cas v18 C1**: suports irresolubles (E/F/G a 9). Amb E, F i G confirmats a 0, la regla 62 denuncia el cap de massa disfressat de plataforma |
| Platform_Surface_Material | TEXT(20) | (Z) | Material de la superfície: Timber / Stone / Mixed / ND. **v18: penja de Z** (regla 51: sense superfície, sense material), a més del bloqueig per rol *Isolated* (regla 7) |

## **2.4b. Interfície N0/N1 i coronament (3) — element I sota `Sys_Interface` (v17); element R sense sistema, deliberadament**

| **Camp** | **Tipus** | **Elem.** | **Descripció** |
| --- | --- | --- | --- |
| Interbody_Cornice | BYTE | **I** | Filada volada **de pedra** que marca la junta entre dos cossos superposats. **Si l'element volat és de fusta, no és una cornisa**: és E o F, i el que hi ha és plataforma |
| Upper_Crown | BYTE | **R** | **Coronament = remat superior construït en general (v20, bloc 3)**, amb dues cares: el parament que tanca la cova per damunt del ràfec (cas fundacional, estructures amb cambra) i el remat de la massa en estructures baixes (lloses terminals). Quina cara és **es deriva** del registre (comptadors i sistemes). Sempre actiu; el «coronament per descart» està prohibit al manual: R afirma estructura completa i acabat deliberat |
| **Upper_Crown_Format** | TEXT(25) | (R) | **NOU v20 (bloc C).** La forma del remat: `Closing panel` (tancament vertical) / `Projecting course` (filera en voladís) / `Flush course` (filera a ras) / `ND`. Default NULL (camp de judici); visible amb R ∈ {1, 2, 3}; la **regla 58** vigila els dos valors presents. Naix amb interés analític declarat i 12 registres a revisar — la lliçó del format de cornisa retirat, apresa |

*`Interbody_Cornice_Material` **es retira en v16**. No perquè de fet les cornises siguen sempre de pedra sinó perquè no podrien no ser-ho: el valor *Wooden beams* no descrivia una cornisa de fusta sinó **una plataforma mal classificada**. T i U ja porten «SEMPRE PEDRA» dins de la definició del vocabulari i no tenen camp de material; la cornisa era l'excepció incoherent. El guany no és estalviar un camp sinó que **una descripció feble es converteix en un test d'identificació** — el que fa falta per a discriminar G de I (8bis.7, regla 39).*

## **2.5. Elements A–Z — Nivell 1 / cos principal i sistema portal (10)**

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
| **Sill_Coincides_Cornice** | BYTE | (N/I) | **NOU v18. Compartició d'element**, el cas que el gradient no pot dir: la cornisa intercòs que **fa de llindar**. N és honestament 0 (cap llindar diferenciat) i la posició queda resolta igualment. Domini 0/1/9 amb **0 de farciment per defecte** (inaplicabilitat derivable → gating, deliberadament fora de la llista de defaults NULL i de la regla 17). Finestra: `Sill = 0` i `Interbody_Cornice ∈ {1,2,3}`; la regla 49 vigila la finestra i `QRY_23` en llista els candidats de migració |
| **Jamb_Fabric_Reveal** | BYTE | (O) | **NOU v19. Segon qualificador de compartició**: la posició del brancal resolta per la fàbrica mateixa — cara terminal acabada i deliberada (*masonry reveal*), sense element diferenciat. Domini 0/1/9 amb **0 de farciment per defecte** (mateix tractament que el qualificador anterior, fora de la llista de defaults NULL). Finestra **(revisada v23)**: `Jambs ∈ {0, 2}` **i `Sys_Portal` Present\*** — una esqueixada pressuposa una obertura, i sense portal el qualificador no té subjecte (l'eix que faltava, trobat treballant la worklist: 53 de les 61 files de la branca de QRY_24 eren inaplicables). Amb O = 2 resol el cas indecidible del Nus 1; amb O = 1 no queda cap posició per resoldre. La regla 55 (revisada) vigila els dos eixos, la branca de portal de la regla 3 s'hi exempta, i `QRY_24` llista els candidats de migració només amb portal present. Etiquetes DOM3Q |

## **2.6. Elements A–X — Cambra i zona superior (6)**

| **Camp** | **Tipus** | **Elem.** | **Descripció** |
| --- | --- | --- | --- |
| Return_Wall | BYTE | **V** | Mur perpendicular al pla de façana, retornant cap al penyal (abans Lateral_Walls) |
| Rear_Closure_Type | TEXT(20) | — | Natural bedrock / Built masonry / Mixed / ND. Com es tanca la cambra pel darrere. L'element W desapareix com a camp (absorbit per Sys_Chamber); el tipus sobreviu perquè una cambra que usa la roca com a tancament posterior no es va construir com una que alça un mur |
| Eave_Beam | BYTE | **S** | Biga de suport del ràfec (component de U; pot ser fusta) |
| Eave_Surface | BYTE | **T** | Superfície del ràfec (component de U; sempre pedra) |
| Chamber_Roof | BYTE | **X** | Tancament efectiu de l'espai funerari per dalt, per obra o per roca **en contacte amb la fàbrica**. **Sense contacte, no hi ha sostre**: una visera que passa per damunt sense tocar no tanca res, i això és `ID_Support` |
| Chamber_Roof_Type | TEXT(25) | (X) | Natural bedrock / Built masonry / Built timber and slabs / Mixed / ND. La decisió que l'anàlisi necessita: roca natural vs obra (regla 14). **`Mixed` = obra i roca repartint-se el tancament**, típicament obra a l'entrada i roca al fons |

*L'**X redefinit en v16** corregeix un error de categoria: definit com «tancada per dalt amb la solució que siga», X no era un element sinó **un estat**, i entrava a la matriu A–X igual si s'havia construït alguna cosa que si no. Amb molts sostres de penya natural al corpus, la coocurrència de Jaccard i les correspondències **comptaven geologia com si fora construcció** i inflaven la similitud entre estructures que no comparteixen cap decisió constructiva — la mateixa família d'error que `ID_Material_Status` i que `Lintel` en v10, dues variables dins d'una escala.*

*El **test és binari i sense gradient**: 20 cm de buit és buit igual que 3 m. Un mausoleu amb la visera separada és `Chamber_Roof = 0` i la visera va a `ID_Support` i a `Systems_Notes`; **no es crea camp** per a l'abric superior, perquè els casos són pocs.*

*A `QRY_13`, X exporta **NA** amb tipus `Natural bedrock` i **1** amb `Mixed` (hi ha obra). NA i no 0: no és una absència sinó una solució no construïda, i un 0 seria tan fals com un 1. **X = 0 continua exportant 0**, perquè sota el patró presència+tipus el tipus és NULL quan la presència és 0, de manera que l'excepció només pot disparar sobre 1, 2 o 3. `Rear_Closure_Type` té la mateixa estructura i el mateix tractament.*

## **2.7. Tractaments superficials (7)**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| Plaster_Present | BYTE | 0/1/9. Revoc / lluït. La conservació parcial la registra `Plaster_Extent` amb *Traces only*, no un valor del domini. Amb 0 o 9, el detall es buida i es bloqueja (regla 24) |
| Plaster_Color | TEXT(20) | White / Cream / Red / Ochre / Grey / ND |
| Plaster_Extent | TEXT(20) | Full facade / Partial / Traces only / ND |
| Pigment_Present | BYTE | 0/1/9. Pigment aplicat. **No promocionat (v14)**: l'única evidència de pigment és pigment, de manera que 3 i 0 es confondrien, i «present complet» no és afirmable. Amb 0 o 9, el detall es buida i es bloqueja (regla 25) |
| **Pigment_Substrate** | TEXT(20) | Plaster / Masonry stone / Bedrock / Mixed / ND. Variable clau: distingeix pintar sobre revoc preparat de pintar sobre el parament (regles 12–13). Al formulari, l'opció Plaster desapareix de la llista quan Plaster_Present = 0 |
| Pigment_Color | TEXT(20) | Red / White / Both / Ochre / ND |
| Pigment_Extent | TEXT(20) | Whole facade / Architectural elements / Decorative motifs / **Perimeter/threshold** / Traces / ND. El valor perimetral (v11) aïlla la pràctica de marcar el llindar independentment del tipus de suport (documentada a EA11 arrasada i EA09 cavitat natural) |

*La combinació `Pigment_Substrate = Bedrock` amb `Pigment_Extent = Architectural elements` **no es bloqueja**: és precisament el cas diagnòstic de cos perdut — bandes verticals de pigment sobre la roca alineades amb els brancals del cos inferior, on l'arquitectura hi era i la pintura li ha sobreviscut. Bloquejar-la impediria registrar l'evidència que sosté un valor 3 amb `Pigment on bedrock`. La regla 30 hi avisa sense impedir res.*

*`Pigment_Substrate = Bedrock` **es manté** i no queda redundant amb el registre de l'art rupestre: és un camp de l'**estructura** i registra sobre què s'aplica el pigment que forma part d'ella. El cas paradigmàtic —el fons rocós d'una cambra que aprofita una cavitat— significa que la roca *era* la superfície acabada prevista, sense operació preparatòria de revoc. La frontera la fixa el test 1 del criteri d'associació (secció 8bis).*

## **2.7b. Paisatge i orientació (4) — H06**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| Facade_Orientation | TEXT(5) | N / NE / E / SE / S / SW / W / NW / ND. **Orientació del pla exposat** |
| **Facade_Azimuth_Deg** | INTEGER | **NOU v24.** Azimut fotogramètric del pla exposat (0–359, nord geogràfic), calculat de la normal del bbox sobre els models de sector georeferenciats i escrit **només** per `chachapoya_Import_Georef_v24.bas` — mai a mà. **No substitueix** `Facade_Orientation`: el sector de 8 resta observacional; l'azimut és instrumental. La **regla 68 (ALERTA)** dispara si divergeixen més de 67,5° circulars (bbox rotat després del calibratge, o observació errada). **Fora de les worklists** (precedent `Metrics_Available`). Al formulari, bloquejat a 2.Arq |
| Visibility_Valley | TEXT(10) | High / Medium / Low / ND |
| **Access_Plane** | TEXT(20) | Facade / Return wall / Rear / ND. Per quin pla de l'estructura s'entra. **v17: gatejat per `Sys_Chamber`**. Etiqueta v18: *Mur de retorn* (A6). **v18: `EA-TER` el tanca** (regla 48) |
| **Portal_Orientation** | TEXT(5) | Mateixos valors que `Facade_Orientation`. Cap a on mira l'obertura. **v17: gatejat per `Sys_Chamber`**, no per `Sys_Portal`. **v18: tancat també amb `Access_Plane = 'Facade'`** — amb l'accés a la façana les dues orientacions són una sola dada, registrada una vegada a `Facade_Orientation`; `QRY_07` exporta `Portal_Orientation_Effective` perquè R mai no faça mitjana d'un duplicat contra un buit. `EA-TER` també el tanca |
| **Portal_Position** | TEXT(20) | **NOU v17.** Centred / Off-centre left / Off-centre right / NA / ND. On se situa el portal dins del pla de façana. Gatejat per `Sys_Portal`. Mateixa convenció d'observador que els trams de `T_DECORATIONS`: **des de davant del pla exposat** |

*La **façana és el pla exposat**, no el pla de l'obertura. Normalment coincideixen i el mot funciona; en una estructura documentada divergeixen —obertura al mur estret perpendicular al farallò, decoració al mur llarg paral·lel— i el mot es trenca. Fixar-la per l'obertura costaria que `Facade_Orientation` deixara d'apuntar a la vall i que `Visibility_Valley` mesurara un pla que no es veu: **tota H06 degradada per a salvar la definició d'un element**. Fixar-la pel pla exposat costa una línia: L és el parament dins del pla de façana, i quan l'obertura hi és, els paraments que la flanquegen; quan no, el parament corregut.*

*`Access_Plane` i `Portal_Orientation` **no es dupliquen**: el pla és **relacional i intrínsec** (com l'arquitectura organitza exhibició i accés, comparable entre orientacions absolutes distintes); l'orientació és **absoluta** i va a QGIS contra la topografia i les rutes de circulació. Cap de les dues es dedueix de l'altra sense conéixer la geometria del cas.*

*`Portal_Orientation` **no es gateja darrere de `Sys_Portal`**, pel mateix motiu que `Recessed_Frame` no ho està: una obertura arrasada té orientació coneguda, i un portal a *Attested lost* amb orientació registrada és precisament un cas informatiu.*

***Gating v17.*** *`Access_Plane` i `Portal_Orientation` es desactiven **quan `Sys_Chamber` es declara** `Absent` o `Not applicable`: sense cambra no hi ha interior on entrar i per tant ni pla d'accés ni portal. Dues precisions que importen. **Es gategen per la cambra i no pel portal**, perquè gatejar-los pel portal trencaria la decisió del paràgraf anterior — amb `Sys_Chamber` present i el portal arrasat, tots dos camps continuen oberts. I **només amb declaració explícita**: amb el camp buit no es desactiva res, perquè buit vol dir «encara no s'ha avaluat» i desactivar per no haver arribat a la pregunta és una manera de perdre dades sense adonar-se'n.*

*`Facade_Orientation` i `Visibility_Valley` **no es desactiven mai**. La façana és el **pla exposat**, deslligat de l'obertura per decisió de v16, i un cos basal sol o una plataforma en tenen i miren cap a algun lloc. Desactivar-los eliminaria del càlcul precisament les estructures sense cambra, que són **el grup de control natural de H06**: la pregunta és si s'orienten igual o distint que les que en tenen.*

*El **guany analític** és que la divergència entre els dos camps d'orientació passa a ser calculable, i per tant preguntable: si les estructures amb accés desviat del pla exposat comparteixen tipologia, sector, forma de suport o amplària de repisa. Amb un sol camp d'orientació no es podia ni formular. La regla 36 vigila la coherència quan `Access_Plane = Facade`.*

***Gating v18: la redundància es tanca en compte de vigilar-se.*** *Amb `Access_Plane = 'Facade'` els dos camps d'orientació diuen la mateixa cosa, i la v17 ho resolia amb un avís (regla 36) que exigia entrar la dada dues vegades coherentment. La v18 tanca `Portal_Orientation` en eixe cas: la dada s'entra **una vegada**, a `Facade_Orientation`, i l'anàlisi la recupera de `Portal_Orientation_Effective` a `QRY_07`. La regla 36 continua vigilant les dades heretades. I la tipologia **`EA-TER` tanca el bloc sencer** — façana, pla d'accés, orientacions, posició i marc: una terrassa en repisa no té ni façana ni portal, inaplicabilitat derivable de la tipologia (regla 48). En triar-la, el formulari declara `Sys_Portal = 'Not applicable'` només si el valor era l'`Absent` per defecte: un valor ja declarat no es toca, es marca.*

***v17: la segona faena d'`Access_Plane` ha desaparegut.*** *En v16 declarava també el pla de referència de `Position_Relative`; eixe camp s'ha retirat, i els quatre trams que el substitueixen es llegeixen **des del pla exposat**, no des d'aquest. La convenció v16 invertia esquerra i dreta precisament en les estructures amb l'accés desviat de la façana, que són el cas d'interés. `Access_Plane` conserva només la seua primera faena.*

## **2.8. Decoració i art rupestre (2) — el detall viu a T_DECORATIONS**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| Dec_Present | BYTE | 0/1/9. Judici agregat de decoració **arquitectònica**. Governa les files amb `Level_Type <> 'ROC'` (regles 22–23) |
| RockArt_Present | BYTE | 0/1/9. Judici agregat de pintura rupestre **associada**. Governa les files amb `Level_Type = 'ROC'` (regles 28–29) |

*Dos camps i no un: `Dec_Present` no pot respondre alhora «aquesta estructura estava decorada?» i «hi ha pintura rupestre associada?». Amb les dues subseccions del formulari separades, un judici agregat damunt de cada una és el que l'usuari veu i el que cada regla pot comprovar. Són judicis agregats i per això mantenen tres valors: cada motiu ja té la seua fila amb el seu propi estat.*

## **2.9. Estat de conservació (6)**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| ID_Arch_Status | LONG | FK → L_STATUS (Good/Fair/Pre-collapse/Collapsed/ND). Amb Collapsed, cap 0 d'asserció és vàlid als elements (regla 6); el formulari mostra l'avís en roig |
| Looting / Fire_Damage / Animal_Activity / Modern_Access | BYTE | 0/1/9. Causes, independents dels graus |

***v17: `Cultural_Materials_Present` i `ID_Material_Status` viuen ara a 2.10 (pestanya 7.Mat).*** *La porta estava a 7.Mat i el camp que governa ací, de manera que es preguntava en quin estat estan els vestigis **abans** de si n'hi ha — i la contradicció era indetectable, perquè el gating de nivell 2 opera dins d'una pestanya. Amb el trasllat, **5.Estat queda per a l'estructura construïda** (com està, què l'ha malmesa, com de bé s'ha pogut observar) i **7.Mat per al seu contingut**, amb exactament la mateixa forma que 6.Bio: una pregunta que obri el bloc i el governa sencer. El que es perd és la contigüitat visual de la conservació dual, i es considera un guany: la proximitat feia pensar que `ID_Arch_Status` i `ID_Material_Status` eren dos graus de la mateixa cosa, quan un parla de l'edifici i l'altre del seu contingut. L'asimetria de v14 es manté: `ID_Arch_Status` no es desdobla, perquè una estructura sempre existeix — és el registre mateix.*

## **2.10. Bioarqueologia (8) i materials culturals (7 + 2) — BYTE 0/1/9**

Bio, idèntic a la versió anterior: Human_Remains (governa la resta del bloc; amb 1, MNI és obligatori — regla 18), MNI, Anatomical_Connection, Mummification, Funerary_Bundles, Dispersed_Remains, Flexed_Position, Bone_Burning.

Materials: **`Cultural_Materials_Present`** (BYTE 0/1/9, porta del bloc, anàleg exacte de `Human_Remains`), els set `Mat_*` (Mat_Textiles, Mat_Wood, Mat_VegFiber, Mat_Ceramics, Mat_Fauna, Mat_DeerAntler, Mat_Other) i **`ID_Material_Status`** (LONG → L_MATERIAL_STATUS: Good/Fair/Poor/ND, sense *Absent* des de v14; vestigis reduïts a pols o fragments irreconeixibles són presència amb `Poor`, i l'absència es reserva per a cambra buida). **Els dos primers arriben de 5.Estat en v17**, i l'estat queda gatejat per la porta.

Les classes de registre *Structural trace* i *Rock art panel* tanquen els dos blocs (regles 15–16).

***Les tres capes de material cultural no són redundants***, i convé escriure per què abans que algú les «simplifique»: responen preguntes amb **llindars d'evidència distints**. *N'hi ha?* és contestable **des de trenta metres** —es veu que dins hi ha coses sense poder dir què són—; *de quins tipus?* exigeix identificar-los; *en quin estat?* exigeix veure'ls prou bé per a jutjar-ho. El cas que ho demostra: una estructura observada de lluny on es veu clarament que hi ha material però no se'n distingeix cap categoria val `1` amb els set a `9`. **Si la presència es deduïra dels set, eixe registre passaria a dir «no hi ha materials», que és fals.** És el mateix patró de judici agregat que `Dec_Present` i `RockArt_Present`. Regles 19, 43 i 44.

## **2.11. Cronologia (3) i fases constructives (2)**

C14 (YESNO, metadada de corpus), Chrono_Start_Cent, Chrono_End_Cent; Construction_Phases, Phase_Evidence (C14 / Stratigraphy / Superposition / Mortar / ND). Les fases sostenen l'argument morfològic-estructural de H03 sense dependre de datacions absolutes; la direccionalitat entre estructures adjacents es registra ara a T_CONNECTIONS.Chrono_Relation (secció 3.6). **v20 (bloc G): cap valor per defecte de fases** — «una fase» és el judici *he buscat límits i no n'he trobat*, i un default l'escriuria sense mirar (argument del denominador sobre H03). **`Phase_Evidence` es gateja per `Construction_Phases ≥ 2`** (inaplicabilitat derivable → gating): amb una sola fase l'evidència no té subjecte. La regla 60 vigila el costat obert; ND és resposta completa. **v22 (bloc C): NOMÉS DATACIONS C14** — reversió explícita de la decisió que deixava els segles sempre editables per a l'atribució tipològica: el projecte data només per radiocarboni, així que `Chrono_Start_Cent` i `Chrono_End_Cent` es gategen pel senyal C14 i la **regla 64** vigila el full de dades (segles amb C14 a no). Cost de dades a la migració: zero — cap registre portava segles ni C14.

## **2.11b. Camps de notes (7 — un per pestanya des de v17a)**

*Cinc pestanyes no en tenien —identificació, decoració, cronologia, mètrica i extra—, de manera que el que no cabia en un camp d'eixes pestanyes no tenia on anar. **`QRY_18_Notes_Review` existeix per a convertir text lliure repetit en camps nous, i només ho pot fer allà on el text lliure és possible**: `Outline_Geometry` va nàixer exactament així, d'una paraula escrita a mà en un camp de notes. El camp general `Notes` passa a `Doc_Notes` perquè les set seguisquen un sol patró en lloc de sis més una excepció.*

`Id_Notes` · `Arch_Notes` · `Finish_Notes` · `Dec_Notes` · `Condition_Notes` · `Bio_Notes` · `Materials_Notes` · `Chrono_Notes` · `Metric_Notes` · `Doc_Notes` · `Systems_Notes` · `Extra_Notes`, més els dos subordinats `Mortar_Notes` i `Vol_Notes`.

*Les de 4.Dec són del **conjunt decoratiu**, no d'un motiu: cada fila ja porta les seues.*

| **Camp** | **Pestanya** |
| --- | --- |
| Arch_Notes | 2.Arq — morfologia, maçoneria, fases |
| Finish_Notes | 3.Acab — revoc i pigment |
| Condition_Notes | 5.Estat — conservació |
| Bio_Notes | 6.Bio |
| Materials_Notes | 7.Mat |
| Systems_Notes | 11.Sist — sistemes i elements |

*Una pestanya mereix camp de notes quan conté **judicis interpretatius que el domini no pot expressar**. No en tenen totes: 9.Mètr ja té `Vol_Notes` i `Dim_Method`, la maçoneria té `Mortar_Notes`, i les taules filles porten el seu propi `Notes`.*

*Van **a la pestanya que els correspon i no a una pestanya de repàs**: una nota s'escriu quan apareix el cas, i el cas apareix emplenant eixa pestanya. La vista de conjunt és `QRY_18_Notes_Review`, de només lectura.*

> **Advertència operativa.** Les notes són el forat negre del registre: tot el que hi va deixa de ser analitzable, i la temptació d'escriure-hi en lloc de decidir el valor del camp és forta. **Si el mateix tipus d'observació apareix repetidament a les notes, això és el senyal que cal un camp** — el mateix criteri aplicat a la hipòtesi de politja i a la relació cromàtica.

## **2.12. Volumetria i superfície (4), coordenades (7), documentació i observabilitat (10)**

**`Area_m2`**, **`Volume_m3`**, `ID_Vol_Method` i `Vol_Notes` (v14: un sol volum i una sola superfície, vegeu 1.4.i). El mètode registra **com** s'ha calculat i les notes **què** s'hi ha inclòs, que és on va la precisió quan un cas la demana; Coord_Lat/Lon_WGS84, Coord_E/N_UTM, Altitude_masl, Coord_Precision_m, ID_Coord_Method; URL_Pano, URL_Pano_2, URL_Giga, URL_3D, ChaXR_Documented, ID_Campaign, Doc_Basis, Facade_Observability, Interior_Observability, Notes. C14 i ChaXR_Documented es mantenen YESNO: metadades del corpus, mai ambigües. Els tres camps d'observabilitat fan interpretables els 9 de tot el registre i es reporten a QRY_15.

**`Doc_Basis` (v14) és una escala ordinal de qualitat documental:** *Direct access* / *Close-range (<5m)* / *Medium-range (5-30m)* / *Long-range (>30m)* / *ND*. Els llindars van dins de l'etiqueta perquè el criteri s'aplique sense consultar cap manual. **La frontera entre els dos primers valors és física i no mètrica**: amb cordes s'està evidentment a menys de 5 m, de manera que el que distingeix l'accés directe és el **contacte** — el que de veres canvia què es pot registrar. La tècnica de captura (gigafoto, dron, 360°, fotogrametria) no es duplica ací: ja la descriuen les URL d'aquesta mateixa secció.

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
| Body_No | INTEGER | Número del cos constructiu. **Obligatòriament buit a les files ROC** (regla 37): una pintura sobre la penya no està en cap cos. **v17: el control s'oculta** a la meitat rupestre en lloc de quedar desactivat |
| **Span_Left** | BYTE | **NOU v17.** 0/1/9. Tram esquerre del contorn. **Només files ROC**; NULL per defecte |
| **Span_Above** | BYTE | **NOU v17.** 0/1/9. Tram del coronament |
| **Span_Right** | BYTE | **NOU v17.** 0/1/9. Tram dret |
| **Span_Below** | BYTE | **NOU v17.** 0/1/9. Tram de la base |

| Color | TEXT(20) | Red / White / Ochre / None / ND. **`Both` retirat en v13** |
| Color_Secondary | TEXT(20) | Color acompanyant: *Red / White / Ochre*, **sense `ND`** (v14). Buit en els casos monocroms, que són la majoria. Buit o ple **és** la distinció monocrom/policrom: cap camp derivat |
| Substrate | TEXT(20) | Plaster / Masonry stone / Bedrock / ND. Resol els casos mixtos per posició |
| Notes | TEXT(255) |  |

*Des de v11 és l'únic registre de decoració; QRY_02 s'hi reconstrueix i QRY_02a en deriva les banderes per motiu quan cal el format ample.*

*Des de v13 conté **també la pintura rupestre associada**, distingida per `Level_Type` de la posició: `ROC` per a la rupestre, la resta per a l'arquitectònica. Una taula, dos conjunts, un discriminador — separar-los és un `WHERE`, i obtindre el conjunt sencer no exigeix cap unió.*

***v17: `Position_Relative` s'elimina.*** *Servia les dues meitats de la taula i **les dues meitats no comparteixen referent**. A les files arquitectòniques *Esquerra* designava **quina instància d'un parell** portava el motiu: una relació **identificacional**, interna a l'estructura. A les files ROC designava **on està la pintura** respecte del volum construït: una relació **topològica** entre dos objectes separats. Que compartisquen les paraules *esquerra* i *dreta* és una coincidència del llenguatge, no una propietat comuna, i per això la pregunta «esquerra respecte de què?» no tenia resposta: canviava segons la meitat, i el camp no ho sabia.*

***Els quatre trams.*** *`Span_Left`, `Span_Above`, `Span_Right` i `Span_Below`, domini 0/1/9, **NULL per defecte** (la regla 4 de farciment amb zero val només per als 20 elements A–X). Registren **quins trams del contorn de l'estructura —present o desapareguda— ocupa aquest motiu**. Habilitats a `ROC-OVL`, `ROC-PER` i `ROC`; desactivats a `ROC-PAN`, que penja d'un registre PR i per definició no té estructura de referència, i a tota la meitat arquitectònica.*

***Per què quatre camps i no una llista de cobertures.*** *Quatre trams donen quinze combinacions i el corpus ja n'ha produït quatre distintes, inclosa una de **base + dreta**. Però l'argument decisiu no és combinatori sinó **epistèmic**: un valor únic no pot distingir «la banda no cobria la base» de «el tram de base no és observable», i eixa distinció és exactament la que separa una **U invertida** d'un **anell tancat mal conservat**. El corpus conté casos on el tram superior queda tapat per la visera rocosa i pel cos superior: marcar-lo com a absent afirmaria una cosa que ningú no ha pogut mirar.*

***U invertida, envoltant i flanquejant no s'emmagatzemen: es deriven*** *a `QRY_20_RockArt_Span`, amb el mateix patró de judici agregat que `Dec_Present`, i amb l'avantatge que una U amb un `9` a la base hi apareix com a **candidata a envoltant** en lloc de quedar afirmada com a U.*

***Convenció d'observador, corregida respecte de v16: esquerra i dreta des de l'observador situat davant del PLA EXPOSAT***, mai des d'`Access_Plane`. La convenció anterior **invertia** esquerra i dreta en tota estructura amb l'accés desviat de la façana, que és precisament el cas d'interés.

***`Outline_Geometry`.*** *No es filtra per tipus de decoració perquè les tres files del corpus que descriuen una U invertida porten **tres tipus diferents** (`RA Amorphous stain`, `RA Perimeter band`, `Painted band`), de manera que qualsevol filtre les deixaria fora precisament a elles. No és un descriptor accessori: un contorn ortogonal **afirma un referent construït rectangular** i un de corb no afirma res —pot estar resseguint un rebaix natural—, i és per tant el camp que decideix si un cas de pintura perimetral sense fàbrica va a la meitat construïda o a la rupestre (8bis.9). Registra **l'observació**: l'afirmació «ací hi havia una estructura» viu a `T_LOST_ELEMENTS`, la mateixa separació que fa `Vertical association`.*

***Eix descartat explícitament***, perquè no torne: **la posició dins d'un element** —«al terç esquerre del dintell»— no s'obri. No s'ha identificat cap ús analític, `Notes` ho cobreix, i obrir-lo multiplicaria combinacions sense cap hipòtesi que les demane.

***A la meitat arquitectònica no hi ha reemplaçament.*** *Cap fila del corpus usa la lateralitat sobre posicions arquitectòniques; `QUO` i `PIL` no la demanen, cosa que resol de passada el problema de les pilastres múltiples. El cas que semblava exigir-la —decoració al pla de façana amb accés lateral— és una propietat de **l'estructura** i `Access_Plane` ja el registra. Fins que aparega una decoració realment asimètrica, va a `Notes`.*

*`Substrate` a les files ROC deixa de ser una tria i passa a ser un **verificador**: *Plaster* i *Masonry stone* són impossibles per definició del criteri d'associació —si el pigment toca la fàbrica, allò és pigment de l'estructura i va a l'altra meitat—, de manera que la llista del formulari es redueix a `Bedrock` i `Mixed` (aquest darrer només per a `ROC-OVL`, que per definició abasta fàbrica i roca) i les regles 38a–38b detecten les files mal encaminades. Una incoherència abans silenciosa passa a ser un error detectat.*

*La parella `Color` + `Color_Secondary` substitueix el valor `Both` amb avantatge, perquè diu quin color domina. Es va adoptar per als casos bicroms amb els colors **junts o contigus dins d'un mateix motiu** (un camp clar amb vora d'un altre to), on descompondre en dues files inventaria dos motius on n'hi ha un i duplicaria la posició. Quan els colors ocupen posicions distintes, en canvi, la descomposició en files continua sent la via correcta: són dues observacions.*

## **3.5. T_ARCH_FEATURES — registre flexible**

Feature_Code (llista amb «Other (see Notes)» com a via d'escapament del LimitToList), Feature_Count, Material, Notes. **v20 (bloc I): la casella `Present` es retira** — una fila existeix perquè s'ha observat una cosa; la casella duplicava l'existència de la fila i el seu default 0 feia mentir 11 de 16 files del corpus. **El catàleg s'ancora a la realitat**: s'hi afigen `Socket / negative interface` (encaix / interfície negativa — 3 casos) i `Access bench` (banqueta d'accés — 2 casos, **l'instrument de recompte de l'element Y**, reservat des de v17 amb mini-especificació pre-decidida al delta); els quatre tipus especulatius de la literatura es conserven — retirar valors de catàleg exigeix més paciència que afegir-ne. **v21 (bloc 4, errata v20): el catàleg és una value list del formulari i el Form v20 no la va rebre** — les dues entrades noves eren inabastables i el retipat de la worklist no es podia fer. Corregit al Form v21; cap canvi de taula.

## **3.6. T_CONNECTIONS — connexions físiques entre estructures (OE3, H03)**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| ID | COUNTER PK |  |
| ID_Struct_A / ID_Struct_B | LONG NN | FK → T_STRUCTURES.ID (extrems de l'aresta) |
| Connection_Type | TEXT(30) | Llista tancada: Abutted vertical joint / Superposition / Bonded joint / Shared support / **Vertical association** (v14) / **Horizontal association** (v17) / **Associated natural context** (v18) / Aerial connection / Other (see Notes). La llista tancada és el que fa computable la regla 20 |
| **Chrono_Relation** | TEXT(20) | **Redefinit v18**: **Sequential / Contemporary / Undetermined**. Diu només **si hi ha direcció**; qui és l'anterior ho diu `ID_Earlier`. La direcció es llig de la junta: només una junta vertical adossada o una superposició la poden portar (regla 20); una junta travada implica Contemporary (el formulari l'autoassigna) |
| **ID_Earlier** | LONG | **NOU v18.** FK → T_STRUCTURES.ID (relació `REL_STR_CONE`, sense integritat forçada com A i B). **Nomena l'estructura anterior per identitat**, no per posició; el formulari la resol **per codi**. Obligatori i restringit a la parella amb `Sequential`; autonetejat en qualsevol altra relació (regla 52). Converteix el graf en genuïnament dirigit (H03) |
| **UQ_CONN_PAIR** | ÍNDEX ÚNIC | **NOU v17.** Sobre (`ID_Struct_A`, `ID_Struct_B`). Amb l'ordre normalitzat, fa **impossible** la parella duplicada |
| Confidence | TEXT(10) | High / Medium / Low |
| Notes | TEXT(150) |  |

*`Vertical association` (v14) registra una relació de verticalitat **observable** entre estructures —proximitat, posició relativa, morfologia del faralló— sense afirmar-ne el mecanisme: la hipòtesi de politja o de suport d'escala continua sent especulativa i va a `T_ARCH_FEATURES` amb els indicis a `Notes`. `Confidence` en registra la seguretat. Material per a OE3.*

***Criteri quan concorren dues relacions*** (junta vertical entre fàbriques sobre una base compartida): **la junta mana sobre el suport**, perquè la junta porta la direcció cronològica (regla 20) i el suport compartit no.*

***v17→v18: la direcció ix de la parella en dos passos.*** *En v16 la direcció vivia a la posició —«A anterior a B»— i l'esquema havia de portar la instrucció «no invertir l'ordre». La v17 va normalitzar la parella (ID menor a A, índex únic) però la lectura seguia sent posicional, de manera que el gir havia de **girar també la cronologia** i la llista entrant havia d'**invertir-la en mostrar-la** per a no mentir. La v18 tanca el problema d'arrel: `ID_Earlier` nomena l'anterior **per identitat**, el gir de normalització ja no toca res (un ID absolut no canvia de valor perquè canvie de columna), i `F_CONN_IN` mostra el fet emmagatzemat en lloc d'un càlcul. Cada pas va eliminar una norma que calia recordar.*

***No era hipotètic.*** *La còpia local contenia dues files de connexió i **eren la mateixa connexió**, entrada des de cada extrem amb valors idèntics. La causa era estructural i no descuit: `F_CONNECTIONS` s'enllaçava amb `LinkChildFields = "ID_Struct_A"`, de manera que des de B no es veia res del que s'havia registrat des d'A, i amb eixa interfície **duplicar era el comportament per defecte**. L'altra meitat de la solució és `F_CONN_IN`, la llista de connexions entrants, de només lectura. **v18: ja no gira res** — llig `ID_Earlier` i diu «aquesta és anterior» o «aquesta és posterior» segons quin extrem siga l'anterior emmagatzemat.*

*`Horizontal association` (v17) és la simètrica de la vertical: registra una alineació **observable** sobre la mateixa cornisa natural o repisa, sense contacte físic i sense afirmar-ne el mecanisme.*

*`Associated natural context` (v18) cobreix el cas terrassa + nínxol natural al seu extrem: **dos registres** (regla de la seqüència constructiva mínima) i una aresta que diu que van junts **sense afirmar cap mecanisme constructiu** — no hi ha junta, ni suport compartit, ni alineació: hi ha un context natural que forma part funcional del conjunt.*

*El cas d'ús paradigmàtic: un complex amb junta vertical no travada es registra com a dues entrades amb codis PALP consecutius, enllaçades ací i agrupades a T_GROUPS; si el model 3D revela filades basals travades sota la junta aparent, correspon un registre únic amb Construction_Phases = 2.*

## **3.7. T_LOST_ELEMENTS — NOVA v11 — evidència del valor 3**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| ID | COUNTER PK |  |
| ID_Structure | LONG NN | FK → T_STRUCTURES (cascade delete) |
| Element_Code | TEXT(2) | FK → L_ELEMENTS.Code. Obligatori quan l'abast és Element (regla 21); el formulari només ofereix elements reals, no sistemes |
| ID_Evidence_Type | LONG NN | FK → L_LOST_EVIDENCE |
| Evidence_Scope | TEXT(20) | Element / Body / Whole structure. Una fila amb abast Body o Whole structure cobreix tot el vocabulari de l'estructura (absorbeix l'antic Lost_Body_Evidence). **v18: regla 53** (mirall de la 21): una fila que nomena un element ha de portar abast *Element*; per a un cos o l'estructura sencera el codi queda buit. Etiqueta UI: *Cos constructiu*, per desfer l'ambigüitat amb el cos humà |
| Notes | MEMO | **v18: hereta l'únic valor que `ID_Position` va arribar a portar** («position: ...») |

***v18: `ID_Position` eliminat.*** *Un valor omplit en 27 files deia tot el que calia dir: la columna feia una pregunta que ningú no responia, perquè `Element_Code` ja situa un element i una fila d'abast Body o Whole no té posició única per definició. El patch va migrar el valor a `Notes` (per nom, via lookup) i va eliminar la relació `REL_SB_LOST`.*

# **4. Taules lookup — valors**

## **4.0. `Name` i `Name_VAL`: la convenció d'idiomes, completa**

Els lookups porten des de v14 **dues columnes de terme**:

| Columna | Funció |
| --- | --- |
| `Name` | **Terme de referència, en anglés.** El que va a exportacions i publicació, i el que llegeixen les regles de QRY_16 que busquen per nom (`Record_Class` via tipologia, `L_STATUS` per a l'avís de col·lapse) |
| `Name_VAL` | **Etiqueta d'interfície, en valencià.** El que mostren els desplegables del formulari |

*El valor emmagatzemat a `T_STRUCTURES` continua sent l'`ID` numèric, mai el text. Això fa la separació encara més robusta que a les llistes de valors: **reanomenar una etiqueta no pot tocar cap dada**, i les files sense traduir cauen a l'anglés en lloc de quedar buides.*

*Les traduccions es poblen per `UPDATE ... WHERE Name = ...` en un sol bloc llegible del codi, no reescrivint cada `INSERT`. **Afegir el castellà seria copiar eixe bloc.** El que no és barat és un selector d'idioma en viu: al formulari hi ha quatre menes de text —combos de taula, combos de llista de valors, etiquetes dels controls i noms de pestanyes— i les columnes del lookup només afecten la primera; commutar-la sola tornaria a deixar mig formulari en cada idioma. La via neta, si algun dia cal, és parametritzar `BuildForm` amb un codi d'idioma i generar un formulari sencer i independent.*

Les llistes següents donen els valors emmagatzemats (`Name`). Les descripcions es tradueixen només a efectes de lectura.

## **L_TYPOLOGY (reconstruïda v11; v17: `EA-PLA-R` reanomenada `EA-TER`)**

***El vocabulari de «plataforma», fixat en v17.*** *La paraula servia tres objectes distints: l'element H, la tipologia `EA-PLA-V` i la tipologia `EA-PLA-R`. Els dos primers **són la mateixa cosa a dues escales** —una superfície volada sobre el buit, a tallant i moment, sola o dins d'una estructura més complexa— i conserven el terme. El tercer és un objecte diferent: una massa construïda que **anivella una repisa**, i passa a dir-se **terrassa**.*

*Es va descartar «cos basal» com a substitut perquè ja el porten `N_Basal_Bodies` i la posició decorativa `BAS`: nomenaria alhora una unitat de registre sencera i un component d'una unitat de registre. També hi pesa que la definició de la tipologia diu «funció: trànsit o base de mausoleu», de manera que un nom que afirmara «basal» resoldria per decret el que `Platform_Function` ha de resoldre cas per cas. I «banqueta» estava ocupada pel massís adossat de 8bis.8.*

| Terme | Reservat per a |
| --- | --- |
| **Plataforma** | Superfície volada sobre el buit: element H i tipologia `EA-PLA-V` |
| **Terrassa** | Massa construïda que anivella una repisa: `EA-TER` |
| **Banqueta** | Massís adossat que no sosté res: `T_ARCH_FEATURES` (8bis.8) |
| **Cos basal** | Component d'una estructura: `N_Basal_Bodies` i posició `BAS` |

| **Nom** | **Record_Class** | **Descripció** |
| --- | --- | --- |
| EA-MAU Mausoleum/Chullpa | Built funerary structure | Estructura construïda (3+ murs + sostre artificial) sobre repisa. Predominant a La Petaca |
| EA-CAM Funerary Chamber | Built funerary structure | Cavitat natural tancada per una façana construïda. Predominant a Diablo Wasi |
| EA-TER Ledge Terrace | Built funerary structure | Massa construïda que anivella una repisa natural |
| EA-PLA-V Aerial Platform | Built funerary structure | Plataforma artificial sobre bigues i lloses, sense repisa de suport |
| NIX Natural Niche | Natural funerary context | Petita cavitat natural (<1 m²). Ossari o enterrament secundari |
| CAV Cave/Cavern | Natural funerary context | Gran cavitat natural (>1 m²) amb ús funerari o ritual documentat |
| PR Rock Art | Rock art panel | Motiu pictòric sobre roca, documentat de forma independent |
| MEN Isolated Structural Element (**v19**) | Structural trace | Registre l'evidència del qual es redueix a un o pocs elements A–Z sense estructura classificable (mènsula, muret transversal, pilastra). La interpretació (circulació, estructura prèvia, suport) va a T_ARCH_FEATURES / Notes, mai a la tipologia; el nom històric ve del primer cas documentat, la mènsula. Sense subtipus: la identitat de l'element la diuen els camps de 11.Sist. Name_VAL: *MEN Element estructural aïllat* |
| Unclassifiable | Built funerary structure | Evidència insuficient per a classificar una estructura CONSTRUÏDA (ex: EA11, perímetre de pigment sense construcció conservada). Els contextos naturals sempre són classificables com a NIX o CAV |
| Not yet classified | Pending classification | Estat de treball: pendent de revisió manual. Exclosa de tota consulta analítica |

*MIX s'ha eliminat: tot EA-CAM és mixt per definició, de manera que no distingia res. DW-S01-EA09 (cavitat amb mènsules i plataforma) es reclassifica MIX → CAV en la migració.*

## **L_SUPPORT (13 valors; v13: desdoblament de Fissure/Crack, v14: Ground, v17: díedre i superfície de roca)**

Columnes: `Name`, **`Support_Mode`** (NOU v13) i **`Description`** (NOU v13), que fixa el criteri operatiu de cada forma.

| **Valor** | **Support_Mode** | **Criteri** |
| --- | --- | --- |
| Wide natural ledge (>2m) | Gravitational rest | Repisa de més de 2 m. L'estructura hi reposa en compressió |
| Narrow natural ledge (<2m) | Gravitational rest | Repisa de menys de 2 m |
| Artificial ledge | Gravitational rest | Repisa construïda o eixamplada per retall. Encreuar amb `Support_Modified` |
| Large cavity (>10m2) | Gravitational rest | Cavitat de més de 10 m². La cambra reposa sobre el sòl |
| Medium cavity (1-10m2) | Gravitational rest | Cavitat d'1–10 m² |
| Natural niche (<1m2) | Gravitational rest | Cavitat de menys d'1 m² |
| Micro-ledge (<50cm) | Gravitational rest | Banc volat de menys de 50 cm. **Criteri: l'estructura hi REPOSA. Si s'ENCASTA a la ranura, és junta d'estratificació** |
| **Vertical cleft** | **Confinement** | Diàclasi que travessa els bancs. Aporta dos paraments laterals, constreny la planta i s'omple per a generar el nivell basal: la roca actua com a encofrat |
| **Bedding-plane recess** | **Embedment** | Rebaix horitzontal per erosió diferencial d'un estrat tou entre bancs competents. Aporta una ranura contínua per a empotrar jàcenes (A) o mènsules (E): treball a tallant i moment, no a compressió |
| **Ground** | **Gravitational rest** | Superfície de terreny a la base del cingle o al peu del vessant. L'estructura assenta sobre sòl, **no en posició elevada sobre un buit**. Única forma sobre sediment i no sobre roca |
| **Rock dihedral** | **Confinement** | **NOU v17.** Angle entrant entre dos plans de roca aproximadament ortogonals, normalment amb sostre. **Natural, no retallat**: encreuar amb `Support_Modified` |
| **Rock surface** | **Substrate** | **NOU v17.** Superfície de roca exposada que sosté pigment, sense cap estructura que hi repose ni s'hi recolze. **Només per a registres PR** |
| ND | ND | Forma de suport no determinada |

*El valor `Combined` es va retirar en v8; els suports compostos es registren com a parella ordenada `ID_Support` + `ID_Support_Secondary` (criteri a la secció 8bis).*

***El díedre (v17).*** *Geològicament és el **negatiu d'un bloc en tascó** despresa al llarg de dues discontinuïtats que s'encreuen, amb el sostre triangular com a tercer pla de fractura. No hi ha terme arqueològic consagrat; s'adopta **díedre**, el de la descripció de parets, perquè nomena **el que s'observa** i no la gènesi — coherent amb el criteri purament geomètric adoptat en v16 per a les posicions ROC.*

***Frontera amb les cavitats: geomètrica, no mètrica.*** *Una **cavitat** és un buit excavat cap endins de la paret: s'hi entra per una boca i la roca envolta per darrere i pels costats. Un **díedre** és un angle entrant entre dos plans: **no té boca** i està obert en dues direccions. Test de camp: **quantes parets estalvia la roca?** Al díedre l'estructura no construeix els dos murs que la roca li dona; a la cavitat construeix la façana i prou. El sostre triangular és una descripció habitual, **no una condició**: el que compta és que la roca estalvie dues parets i una coberta, no si el sostre és un pla de fractura o un d'estratificació.*

***El mode `Substrate` és nou i és necessari.*** *Cap dels tres existents descriu el que fa una superfície que sosté pigment: no és descans gravitacional, ni confinament, ni encastament. La categoria pròpia la manté **fora de tot recompte de suport constructiu**, que és el que es vol.*

***Criteri principal / secundari (v17), orientatiu i deliberadament no validat:*** **el suport principal respon a les forces verticals** (aguanta el pes), **el secundari a les horitzontals** (estabilitza). No s'afig cap comprovació a la bateria: tres dels quatre suports secundaris actualment emplenats són repises o cavitats —formes que aguanten pes— i poden ser perfectament correctes, perquè una estructura pot descansar alhora sobre una repisa i sobre el sòl d'una cavitat. Una regla estricta les marcaria com a errònies sense motiu. Amb aquest criteri, el díedre serà quasi sempre secundari.*

## **L_ELEMENTS — NOVA v11 — el vocabulari A–Z com a lookup (25 files)**

Columnes: Code (únic, UQ_ELEM_CODE), Name_EN, Name_VAL, Level_Type (N0 / N0-N1 / N1 / SUP), Sys_Group (Platform / Portal / Eave / buit), Is_System, Field_Name, Description. Les files H, P i U porten Is_System = True i Field_Name = Sys_*; la fila W es conserva com a referència de vocabulari sense camp propi. Sys_Base i Sys_Chamber no apareixen a Sys_Group: són agrupacions sense lletra. **v18: entra la fila Z** (*Platform surface / Superfície de plataforma*, N0, grup Platform, Field_Name `Platform_Surface`) i la descripció de D registra el seu desgatejament. **La lletra Y queda reservada** per a la massa adossada (banqueta), pendent dels seus tres casos documentats.

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

L_STATUS: Good/Fair/Pre-collapse/Collapsed/ND. **L_MATERIAL_STATUS (v14): Good/Fair/Poor/ND**, sense *Absent* — l'absència la registra `Cultural_Materials_Present`. Les quatre campanyes 2013–2023.

**L_GROUP_TYPE (8 tipus, v14: +1):** alineament vertical, conjunt de repisa, plataforma amb mènsules, conjunt de cova, xarxa de circulació, conjunt d'art rupestre, **unitat supraestructural** i grup funcional.

*La **unitat supraestructural** treu la convenció d'anomenar Xa / Xb les unitats d'un mateix conjunt de la cadena del codi, on vivia com a hàbit textual sense ser consultable. Criteri, formulat ample: **dues o més unitats constructives coherents amb atributs propis sobre una base compartida — no necessàriament dues cambres**; un mausoleu més una plataforma en voladís qualifica, i és un cas freqüent. Divisió del treball amb `T_CONNECTIONS`: la connexió és **binària** i descriu la junta (i porta la direcció cronològica); el grup és **n-ari** i descriu la unitat — tres cossos donen tres connexions però un sol grup.*

## **L_STRUCT_BODY (18 entrades; v16: −`SOC`, +4 arquitectòniques, ROC redefinides)**

| **Code** | **Name** | **Name_VAL** | **Level_Type** | **Sort_Order** |
| --- | --- | --- | --- | --- |
| BAS | Basal body (B/C) | Cos basal (B/C) | N0 | 10 |
| COR | Interbody cornice (I) | Cornisa intercòs (I) | N1 | 30 |
| QUO | Corner quoin (J) | Cantonera (J) | N1 | 40 |
| PIL | Pilaster (K) | Pilastra (K) | N1 | 50 |
| FFL | Facade flank (L) | **Flanc de façana (L)** | N1 | 60 |
| **RTW** | **Return wall (V)** | **Mur de retorn (V)** | N1 | **62** |
| **PRT** | **Portal system (P)** | **Sistema portal (P)** | N1 | **64** |
| **SIL** | **Sill (N)** | **Llindar (N)** | N1 | **65** |
| JAM | Jamb (O) | Brancal (O) | N1 | 70 |
| **LIN** | **Lintel (Q)** | **Dintell (Q)** | N1 | **75** |
| OVL | Over-lintel (M) | Sobre-dintell (M) | N1 | 80 |
| CRO | Upper crown (R) | Coronament (R) | N1 | 90 |
| EAV | Eave (U) | Ràfec (U) | SUP | 100 |
| **ROC-OVL** | **Overlapping** | **Solapada** | ROC | **200** |
| **ROC-PER** | **Perimeter** | **Perimetral** | ROC | 210 |
| **ROC** | **Proximate** | **Pròxima** | ROC | 220 |
| ROC-PAN | Panel (no architectural ref.) | Panell (sense ref. arquitectònica) | ROC | 230 |
| ND | Not determined | No determinada | ND | 999 |

*Descriu **només la posició dins d'un cos**; de quin cos es tracta ho diu `T_DECORATIONS.Body_No`.*

**`SOC` retirat (v16).** C es defineix com el **tractament decoratiu** del basament, de manera que «una decoració situada al sòcol» significa «una decoració situada al tractament decoratiu del basament»: **la posició pressuposa el que la fila hauria de documentar**, i l'operació inversa tampoc no funciona — registrar un motiu al basament ja el converteix en sòcol. Es va valorar mantindre les dues entrades repartint-les per tipus de motiu (sòcol per als relleus, basament per a les pintures) i **es descarta**: faria que el camp Posició codificara el tipus, que és el que registra `ID_Dec_Type`, i un fris pintat en relleu quedaria sense valor possible.

> **Frontera fixada:** **C = tractament plàstic** (filada volada, motllura, canvi d'aparell o de format de pedra que individualitza la massa basal). **Fila `BAS` = motiu aplicat** (pigment, ninxol, relleu) sobre el cos basal.

*Un basament amb banda pintada i sense motllura és `Decorative_Socle = 0` **més** una fila `BAS`. Amb la definició v15 això no es podia decidir. La regla 34 és **unidireccional**: C i les files són independents —pot haver-hi motllura sense motiu i motiu sense motllura—, i només un relleu al cos basal amb C = 0 és incoherent.*

**Quatre posicions arquitectòniques noves (v16).** La llista v15 tenia `JAM` i `OVL` però **ni llindar ni dintell**, i `Over-lintel` no és el dintell sinó la zona per damunt, l'àrea de fris (element M): **un portal pintat s'havia de registrar contra un element que no era el pintat**. `PRT` respon la pregunta que ve darrere: una banda contínua que emmarca tota l'obertura és **un sol gest, no tres**, i partir-la en tres files inventaria tres motius on n'hi ha un — el mateix error que va justificar `Color_Secondary`. Precedent estructural: el vocabulari ja tracta N+O+Q com un **sistema** perquè els components actuen junts.

> **Criteri d'ús:** posició elemental (`SIL` / `JAM` / `LIN`) quan el motiu **s'interromp** entre elements o en decora un de sol; `PRT` quan el tractament és **continu al voltant de l'obertura** i no s'atribueix a cap component. **En cas de dubte, elemental**: sempre es pot agregar després, mai desagregar.

*`RTW` respon un cas documentat: decoració al mur lateral **i no a la façana**. Això no és un detall menor — és decoració no dirigida a l'espectador de la vall, i toca H06 de front. Sense l'entrada, eixos casos es perdien o es registraven com a façana, que és pitjor.*

*Nota terminològica que ha de quedar escrita: **llindar** (peça inferior, element N, *sill*) i **llinda** (peça superior, element Q, *lintel*) són falsos amics quasi homògrafs i font directa d'error de registre. **A la interfície s'usa sempre `Dintell` per a Q i mai `llinda`.***

**Les posicions ROC, redefinides per contacte geomètric (v16).** La v15 deia de `ROC` «usar quan s'apliquen els tests 2 o 3» i de `ROC-PER` «test 3»: **els dos tests 3 xocaven**, de manera que qualsevol pintura que emmarcara una obertura complia la condició de les dues entrades i la tria quedava a l'atzar de l'observador. El criteri nou és **purament geomètric** —es resol mirant, sense consultar `ID_Support`— i s'aplica **en ordre, de més contacte a menys**:

> **`ROC-OVL` · Solapada** — un mateix motiu té part sobre la fàbrica i part sobre la roca. `Substrate = Mixed`.
> **`ROC-PER` · Perimetral** — ressegueix el contorn de l'estructura, **present o desapareguda**, en contacte tangencial.
> **`ROC` · Pròxima** — no toca la fàbrica en cap punt, però es troba **dins del mateix accident del farallò** que allotja l'estructura: la mateixa repisa, cavitat, escletxa o rebaix. **Si cal travessar una vora de banc, una cornisa natural o un buit per a arribar-hi, és fitxa PR.**
> **`ROC-PAN` · Panell** — files penjades d'un registre PR. **No es tria mai des d'un registre d'estructura.**

*El **límit exterior de la pròxima** és indispensable: «no toca l'estructura» no acaba enlloc —tota pintura del farallò no la toca— i sense límit `ROC` s'empassaria el que ha de ser fitxa PR. El límit és **no mètric**, coherent amb el descart del llindar d'un metre: en paret vertical la distància euclidiana ignora el que estructura aquests jaciments, i el criteri reutilitza una decisió ja presa (`ID_Support`).*

*La clàusula **«present o desapareguda»** de la perimetral és el que fa funcionar la categoria: bandes verticals sobre roca nua alineades amb els brancals d'un cos que ja no hi és són `ROC-PER` sense contacte possible, perquè el contorn que ressegueixen és el que hi havia. La definició diu **contorn**, no **contacte**, perquè el contacte és una observació de l'estat actual. Encreua amb `T_LOST_ELEMENTS` i amb l'avís de la regla 30.*

*`ROC-OVL` **contradiu el test 1 vigent** i el reformula: **el contacte parcial no basta per a fer una pintura arquitectònica; el que la fa arquitectònica és estar sencera sobre la fàbrica**. Una fila, posició `ROC-OVL`, i el costat fàbrica el registra `Pigment_Substrate = Mixed` a 3.Acab. Partir-la en dues files es descarta: inventaria dos motius on n'hi ha un.*

*És, a més, **la categoria analíticament més interessant de les quatre**: un motiu que travessa la junta entre roca i maçoneria demostra que qui pintava tractava les dues superfícies com una sola, i que la pintura és posterior a la fàbrica en eixe punt. Toca H04 i H06. En v15 eixos casos es dissolien dins de `Pigment_Substrate = Mixed` sense deixar rastre com a motiu.*

*`ROC-PER` (posició) i `RA Perimeter band` (tipus) **no són el mateix camp** i no s'han de pressuposar junts: el tipus diu **què sembla**, la posició diu **on està**. Una banda perimetral pot ser pròxima; un motiu antropomorf pot ser perimetral.*

*Amb `Position_Relative` operatiu **no cal cap posició ROC més**: els casos que semblaven demanar-ne —pintura per damunt, al costat, davall— es diuen amb `ROC` + `Damunt` / `Esquerra` / `Davall`. Quatre entrades i un camp cobreixen més combinacions que sis entrades, i no s'inventa res sobre els dos casos documentats al corpus.*

*`Level_Type = ROC` continua sent el **discriminador analític** entre decoració arquitectònica i pintura rupestre, i `ND` continua significant «posició no determinada»: **no s'ha d'usar mai per a posicions no arquitectòniques**. Confondre-ho és el mateix error que confondre 0 amb 9.*

## **L_DEC_TYPE (17 entrades; v17: tres retirats; v17a: meitat rupestre refeta)**

**Repertori arquitectònic (10):** T-shaped niche (i inv.), L-shaped niche (i inv.), Zigzag, Stepped motif, Painted band, Square niche, Plain colour field, ND. «Plain colour field» converteix T_DECORATIONS en el registre general de tractament de superfície per posició.

***Tres entrades retirades en v17, totes sense cap ús a les 56 files del corpus:***

- **`Frieze / Greca`** — era **l'única entrada de la llista que no nomenava cap motiu**: deia que hi ha fris, que és exactament el que ja diu `Relief_Frieze` (M). Els altres —ziga-zaga, escalonat— sí que especifiquen, i `ND` cobreix el fris amb motiu no resolt. Mantindre dos valors per a la mateixa situació és el que va causar el problema de les posicions ROC en v15. Si la **greca** com a motiu andí específic mereix entrada algun dia, serà una entrada **nova** amb definició de motiu.
- **`Triangular motif`** — mai observat al corpus com a distint del ziga-zaga.
- **`Decapitation scene`** — un sol cas documentat, i **art rupestre, no decoració arquitectònica**. Va a `RA Anthropomorphic` amb la lectura a `Notes`, que és el que la seua pròpia descripció ja instruïa. Una entrada sencera de la llista arquitectònica per a un motiu rupestre invertia la proporció entre vocabulari i corpus.

***El repertori arquitectònic restant és CONSTRUCTIU.*** *Ziga-zaga, escalonat, T, T invertida, L, L invertida i nínxols quadrats es fan **col·locant les pedres endins i enfora**, i tots set es llegeixen com a **tipus de fris**; que a més estiguen pintats és una altra pregunta, i per això la fila porta `Color`. `Painted band` i `Plain colour field` queden fora d'eixe conjunt: **s'apliquen sobre la superfície en lloc de modelar-la**, i una banda pintada al parament no és evidència de fris en relleu. La regla 42 cobreix els set primers i no els dos darrers.*

**Repertori rupestre (5 amb prefix RA, refet en v17a i podat en v20):** RA Anthropomorphic, RA Zoomorphic, RA Geometric motif, RA Amorphous stain, RA Pigment traces. **Les bandes de contorn (v20, bloc F) perden el prefix i passen al costat arquitectònic del filtre**: `Outline band, defined` (banda de contorn definida) i `Outline band, amorphous` (amorfa) — el tipus diu el **traç**, la topologia U/C/O es deriva dels trams (`QRY_27`), i el veredicte estructural viu a la tipologia del registre i l'atestació, mai a la banda. Una U com a iconografia de panell és `RA Geometric motif`.

***Per què es va refer.*** *Emplenant registres, la llista resultava confusa, i el motiu era estructural i no de redacció: **barrejava eixos**. `Painted band` nomenava una **forma** i `Perimeter band` una **posició**, de manera que una banda perimetral era també una banda pintada; `Abstract` i `Amorphous stain` nomenaven totes dues un **grau de llegibilitat** sense dir on era la frontera. Una llista les entrades de la qual responen preguntes distintes no es pot aplicar de manera consistent per molt bé que es definisca cada entrada.*

*La llista nova respon **una sola pregunta: quin motiu és**. `Plain colour field` i `Painted band` surten de la meitat rupestre; una banda sobre penya que no siga una U és `RA Geometric motif`.*

***Les dues formes en U absorbeixen `Outline_Geometry`***, que es retira (vegeu 8bis.9). El camp només discriminava dins de la banda perimetral, de manera que com a camp general quedava buit a la majoria de files i hi feia una pregunta sense sentit. Plegat dins del tipus, la mateixa informació no costa cap columna i **no pot quedar en contradicció amb ell**.

***Les dues fronteres que donaven problemes, escrites:***

> **Taca amorfa** — superfície de pigment definida, amb **vora reconeixible**, però sense motiu identificable.
> **Traces de pigment** — restes disperses o massa degradades per a dir si van formar cap motiu.
> **Motiu geomètric** — línies, bandes o figures amb organització regular reconeixible.

*El criteri que separa les dues primeres és si el pigment té **vora llegible**, no si sembla significatiu: eixa era exactament la confusió que produïa el parell `Abstract` / `Amorphous`.*

*Els dos repertoris **no són compartits** — es va comprovar contra el corpus—, de manera que les entrades rupestres s'afigen en lloc de fondre's. Alguns tipus genèrics (`Plain colour field`, `Painted band`, `ND`) serveixen les dues bandes i apareixen a les dues llistes del formulari; per això el discriminador entre conjunts és `Level_Type` i no `ID_Dec_Type`.*

*Les dues formes en U mereixen atenció: marques que emmarquen el contorn d'una estructura, present o desapareguda. Quan la fàbrica ja no hi és, encreuar-les amb `T_LOST_ELEMENTS` — i la **geomètrica** és la que sosté l'afirmació que allí hi va haver construcció.*

# **5. Relacions (25)**

Les 21 de la versió anterior més les quatre de T_LOST_ELEMENTS:

| **Nom** | **Pare** | **Filla** | **Camp fill** | **Notes** |
| --- | --- | --- | --- | --- |
| REL_STR_LOST | T_STRUCTURES | T_LOST_ELEMENTS | ID_Structure | Cascade delete |
| REL_ELEM_LOST | L_ELEMENTS (Code) | T_LOST_ELEMENTS | Element_Code | Requereix l'índex únic UQ_ELEM_CODE |
| REL_LEV_LOST | L_LOST_EVIDENCE | T_LOST_ELEMENTS | ID_Evidence_Type | Update cascade |
| ~~REL_SB_LOST~~ | — | — | — | **Eliminada v18** amb `ID_Position` |
| **REL_STR_CONE** | T_STRUCTURES | T_CONNECTIONS | ID_Earlier | **NOVA v18.** Tercera aresta de connexions cap a estructures; sense integritat forçada (flag 2), com A i B |

*REL_STR_SELF (autoreferenciant) i les ara **tres** de T_CONNECTIONS continuen amb dbRelationDontEnforceIntegrity (valor numèric 2). El total es manté en 25: −`REL_SB_LOST`, +`REL_STR_CONE`.*

# **6. Consultes SQL (36)**

| **Nom** | **Descripció** | **Hip.** |
| --- | --- | --- |
| QRY_01_Typology_by_Site | Distribució de tipologies per jaciment i sector. | H01, H04 |
| QRY_02s_Decoration_Typed | Auxiliar: files de T_DECORATIONS amb tipus i posició resolts. | — |
| QRY_02a_Decoration_Flags | Banderes per motiu derivades de T_DECORATIONS (format ample per a khi-quadrat). | H01, H04 |
| QRY_02_Decoration_by_Site | Reconstruïda v11 sobre T_DECORATIONS: recomptes per motiu i jaciment. | H01, H04 |
| QRY_03 – QRY_12 | Sense canvis de funció: conservació, C14, exportació R, volumetria, exportació QGIS, fills, membres de grup, cobertura ChaXR, maçoneria, geologia-construcció. QRY_05 i QRY_07 actualitzades v11 als noms nous i als camps de sistema. | diverses |
| QRY_13_AX_Pattern_Export | Reconstruïda v11; **v18: 21 columnes AX_ (entra `AX_Z`) + 6 columnes SYS_ (entra `SYS_INTERFACE`, que fins ara gatejava I sense exportar el seu propi estat)**. `Recessed_Frame` s'hi exporta gatejat per `N_Chamber_Bodies` i el format de cornisa per `Sys_Interface`. Semàntica v11 original: **20 columnes AX_ + 5 columnes SYS_** amb semàntica NA (un component gatejat per un sistema Not applicable s'exporta NA, no 0, perquè el seu 0 és de farciment). La font única FillElementMap garanteix que exportació i validació no divergeixen mai. **v16: X exporta NA quan `Chamber_Roof_Type = Natural bedrock`** i la columna `Interbody_Cornice_Material` desapareix. **Advertència: amb molts sostres de penya natural al corpus, el filtre canvia els resultats i qualsevol prova de Jaccard, clúster o correspondències ja executada s'ha de refer.** | H01, H04, H05 |
| QRY_14_Connections_Edges | Llista d'arestes amb Chrono_Relation: graf dirigit per a igraph/QGIS. **v18: columna `Code_Earlier`** resolta per LEFT JOIN sobre `ID_Earlier` — la direcció viatja per nom d'estructura, no per posició. | OE3, H03 |
| QRY_15_Observability_Bias | Recompte per jaciment i sector de base documental i observabilitat. | OE1 |
| QRY_16s_* (5) | Auxiliars de la bateria: Element_Values (format llarg dels **21** elements: la clau que evita ~80 branques UNION), Lost_Cover, Damage_Count, Observational_Nulls (llista de treball camp a camp), Null_Count. | — |
| QRY_16a / QRY_16b / QRY_16c / QRY_16d / QRY_16e / **QRY_16f (v18, ampliada v19)** | Regles 1–11, 12–21, 22–31, 32–39, 40–46 i **47–56 (sense la 50, retirada en v20 amb el seu camp)**. Emmagatzemades en parts perquè una UNION única de totes les branques supera el límit «query too complex» de JET. **v24: la setena parcial es diu `QRY_16g_Rules_57_68`** (dotze branques; la 68 és l'ALERTA d'azimut). **v20 afig `QRY_16g_Rules_57_60`** amb les quatre regles noves: 57 (concordança codi–sector), 58 (format del coronament), 59 (camps de portal amb comptador de cambra a 0), 60 (fases ≥ 2 sense evidència). **`QRY_16d` porta nou branques per a huit números**: la regla 38 es desdobla, perquè el substrat que ha de portar una fila ROC depén de quina posició ROC és | — |
| QRY_16_Validation_Check | Unió de les **set** parts: **65 regles actives (numerades fins a 67; retirades 41 i 50)**. Resultat buit = corpus coherent. **Correcció de bug v16**: `QRY_16c` es creava amb el nom `..._22_30` i es referenciava com a `..._22_31`, de manera que aquesta unió **no arribava a existir** — i `MkQuery` degrada eixe fracàs a un `Debug.Print` que ningú no llegia. Les tres parcials existien i s'obrien per separat, que és per què va passar desapercebut | — |
| **QRY_19_V16_Review** | **NOVA v16.** Llista de treball de la revisió manual que la transferència v15→v16 exigeix: elements de portal a rejudicar, cobertes de cambra pendents del test de contacte, format i treball de la pedra per entrar, i files de decoració en posició `JAM` o `OVL` que podrien ser ara `PRT`, `SIL` o `LIN` | — |
| **QRY_17_RockArt_All** | **NOVA v13.** Totes les pintures del corpus amb una columna `Context` (*Associated* / *Isolated*). No necessita cap UNION: les associades i les aïllades són **totes dues files de `T_DECORATIONS`** —les primeres penjades d'una estructura, les segones d'un registre PR, que és també un registre de `T_STRUCTURES`—, de manera que és una sola consulta amb JOIN a `L_TYPOLOGY` per a saber quina mena de pare té cada fila. | OE1, H06 |
| **QRY_18_Notes_Review** | **NOVA v14; reescrita v18** — llegia `E.Notes`, renomenat `Doc_Notes` en 17a, i demanava un paràmetre en obrir-se; ara escombra els **14 camps de notes** en ordre de pestanya. Tots els camps de notes de tots els registres, costat per costat, filtrant els buits. Serveix per al repàs i, sobretot, per a **detectar patrons**: quan la mateixa observació apareix repetidament en text lliure, això és el senyal que hauria de ser un camp. *Va funcionar: `Outline_Geometry` existeix en v17 perquè una fila del corpus portava la paraula «(geomètrica-ortogonal)» escrita a mà a `Notes`.* | — |
| **QRY_20_RockArt_Span** | **NOVA v17.** Deriva les etiquetes de cobertura —*Surrounding*, *Inverted U*, *Flanking*, *Partial*— dels quatre camps de tram, i marca com a **infralegida** (`Under_Read`) tota fila amb algun `9`. Les etiquetes **no s'emmagatzemen mai**: així una U invertida amb la base no observable apareix com a **candidata** a envoltant en lloc de quedar afirmada com a U. | OE1, H06 |
| **QRY_22_V17a_Review** | **NOVA v17a.** Files de decoració **sense tipus**: les que eren `RA Perimeter band` i que el pedaç va deixar deliberadament sense decidir, perquè triar entre U geomètrica i U orgànica és una lectura de la fotografia. Porta els quatre trams i les notes, on el pedaç ha bolcat el valor antic de geometria. | — |
| **QRY_21_V17_Review** | **NOVA v17.** Llista de treball de la transferència v16→v17: files ROC amb els trams per omplir, files de decoració **sense cap posició** (invisibles sota el formulari v16), connexions amb cronologia per rellegir, estructures multicòs amb `Fabric` buida i fases declarades sense evidència. | — |
| **QRY_23_V18_Review** | **NOVA v18.** Llista de treball de la migració v17a→v18: Z per observar on el sistema plataforma és obert, zeros de D a confirmar com a observació (eren farciment sota `Sys_Base` tancat), candidats de coincidència cornisa-llindar, incoherències d'abast a `T_LOST_ELEMENTS` i `Sequential` sense anterior nomenat. Buida en una construcció de zero, per construcció. | — |
| **QRY_24_V19_Review** | **NOVA v19.** Llista de treball de la migració v18→v19: candidats de la finestra del qualificador de fàbrica (Jambs a 0 o 2 amb el 0 de migració), zeros de D retornats a NULL pel patch (judici pendent) i rols de mènsula encara NULL (arbre de tres preguntes). Duplica R9 a propòsit — la bateria informa d'un error, açò és una llista de treball. Buida en una construcció de zero, per construcció. | — |
| **QRY_25_V20_Review** | **NOVA v20; revisada v21.** Llista de treball de la migració v19a→v20, ara **quatre branques**: format del coronament als R presents; bandes *amorfes* a confirmar amb els trams davant; files de 12.Extra a retipar (encaixos i banquetes); i els MEN amb la passada de 2.Arq pendent — **branca que en v21 respecta la declaració de fàbrica** (dispara amb `Masonry_Present` NULL, o a 1 amb la qualitat buida: una mènsula amb declaració 0 té la pestanya resolta). **La branca de les «dues preguntes» a les terrasses es retira en v21 amb el criteri que la motivava** (bloc 1). Buida en una construcció de zero. | — |
| **QRY_26_Next_EA** | **NOVA v20 (bloc B).** Per sector: recompte, màxim EA i següent lliure. El criteri màxim+1 fet visible en el moment d'entrar; els buits de la seqüència són història de campanya, mai espai lliure. | — |
| **QRY_27_Outline_Topology** | **NOVA v20 (bloc F).** La topologia de cada banda de contorn **derivada dels quatre trams**: 1,1,1,0 és la U invertida; els quatre, una O; parcials, arc/C. Cap fórmula constant, cap entrada nova — la hipòtesi estructura→U / nínxol→C es comprova creuant açò amb tipologia i jaciment. | — |
| **QRY_29_V21_Review** | **NOVA v21.** Llista de treball de la migració v20→v21, dues branques: **plataformes Z-sol** del criteri antic de terrassa (Present* amb E, F i G confirmats a 0) a revertir **fila a fila** — duplica la regla 62 a propòsit, la bateria denuncia i açò és worklist; i **declaracions de fàbrica NULL** que el patch no ha pogut derivar — duplica la 17 a propòsit, és el judici que la migració obri. `QRY_28` és del navegador de worklist, d'ací el salt de numeració. Buida en una construcció de zero. | — |

## **6.1. La bateria de regles (65 actives, numerades fins a 67; retirades 41 i 50)**

No bloquejant per disseny: una restricció dura impediria registrar una absència genuïnament observada en una estructura parcialment col·lapsada. Resum:

| # | Regla | # | Regla |
| --- | --- | --- | --- |
| 1 | Component present (1/2/3) amb sistema no present (exempts: Not observable, i Timber_Brackets) | 14 | Coberta present sense tipus |
| 2 | Valor 3 (o sistema Attested lost) sense fila d'evidència a T_LOST_ELEMENTS | 15 | Structural trace amb restes humanes |
| 3 | Sistema present amb tots els components a 0 (**v19: branca de portal exempta amb el qualificador de fàbrica declarat**) | 16 | Contingut bio/materials en classes que tanquen els blocs |
| 4 | Component ≠ 0 sota sistema Not applicable | 17 | Camps observacionals encara NULL (llista de treball, no error) |
| 5 | Estat Good amb >30% d'elements parcials o perduts | 18 | Restes humanes sense MNI |
| 6 | 0 d'asserció en estructura Collapsed | 19 | `Cultural_Materials_Present` = 0 amb Mat_* ≠ 0 |
| 7 | Material de plataforma amb rol Isolated | 20 | Cronologia direccional en connexió sense adossament |
| 8 | Mènsules presents sense recompte | 21 | Abast Element sense Element_Code |
| 9 | Mènsules presents sense rol (**v19: només amb NULL; `ND` és judici**) | 22 | `Dec_Present` sense files no-ROC de decoració (v12, restringida en v13) |
| 10 | Dintell 0 amb material | 23 | Files no-ROC de decoració sense `Dec_Present` (v12, restringida en v13) |
| 11 | Cossos de cambra amb Sys_Portal Absent | 24 | Detall de revoc sota presència 0/9 (v12) |
| 12 | Pigment present sense substrat | 25 | Detall de pigment sota presència 0/9 (v12) |
| 13 | Pigment sobre revoc amb revoc absent | 26 | Mortar_Type incompatible amb Mortar_Present = 0 (v12) |
| — | — | 27 | Cossos basals comptats amb `Base_Level` (B) = 0 (v13) |
| — | — | 28 | `RockArt_Present` declarat sense cap fila ROC (v13) |
| — | — | 29 | Files ROC sense `RockArt_Present` (v13) |
| — | — | 30 | Pigment sobre penya seguint elements arquitectònics: possible cos perdut (v13, **avís**) |
| — | — | 31 | Dimensió lineal implausible (>20 m): comprovar la unitat (v14) |
| — | — | **32** | **`Stone_Format_Secondary` amb format dominant buit o `ND` (v16)** |
| — | — | **33** | **`Stone_Format_Secondary` igual al primari (v16)** |
| — | — | **34** | **Motiu en relleu al cos basal amb `Decorative_Socle` = 0 (v16, unidireccional)** |
| — | — | **35** | **Decoració en posició `LIN` o `SIL` amb l'element corresponent a 0 (v16)** |
| — | — | **36** | **`Access_Plane` = Facade amb les dues orientacions divergents (v16, avís)** |
| — | — | **37** | **Fila ROC amb `Body_No` emplenat (v16)** |
| — | — | **38** | **Substrat incompatible amb la posició ROC (v16, dues branques: no-`Bedrock` a ROC/PER/PAN, no-`Mixed` a `ROC-OVL`)** |
| — | — | **39** | **Cornisa intercòs amb menys de dos cossos (v16, avís): revisar si és filada en voladís (G)** |
| — | — | **40** | **Classe de registre i posició de decoració incompatibles (v17, dues branques): registre rupestre amb posició de contorn, i registre d'estructura amb posició `ROC-PAN`** |
| — | — | **41** | ~~`Outline_Geometry` en fila no-ROC~~ — **retirada en v17a amb el camp que vigilava. Els altres números NO es renumeren** |
| — | — | **42** | **Motiu constructiu a `OVL` o `FFL` amb `Relief_Frieze` = 0 (v17, unidireccional)** |
| — | — | **43** | **`Cultural_Materials_Present` = 1 amb els set `Mat_*` a 0 (v17)** |
| — | — | **44** | **`ID_Material_Status` emplenat amb `Cultural_Materials_Present` = 0 (v17)** |
| — | — | **45** | **Fàbrica declarada múltiple *per cossos* amb menys de dos cossos (v17)** |
| — | — | **46** | **Dues o més fases constructives sense `Phase_Evidence` (v17)** |
| — | — | **47** | **Marc reculat amb `N_Chamber_Bodies = 0` (v18): sense cos de cambra no hi ha façana a qualificar** |
| — | — | **48** | **`EA-TER` amb dades de façana o portal (v18): una terrassa no en té; els camps es netegen o el registre es reclassifica** |
| — | — | **49** | **`Sill_Coincides_Cornice = 1` fora de la finestra `Sill = 0` amb cornisa amb entitat (v18)** |
| — | — | **50** | **Format de cornisa amb I = 0 o 9 (v18, forma de la 10)** |
| — | — | **51** | **Material de superfície de plataforma amb Z = 0 (v18, forma de la 10); la direcció *Isolated* continua a la 7** |
| — | — | **52** | **Connexions (v18, dues branques com la 38): `Sequential` sense `ID_Earlier` vàlid dins de la parella; no-`Sequential` amb `ID_Earlier`** |
| — | — | **53** | **Fila d'element desaparegut amb codi i abast no-*Element* (v18, mirall de la 21). El corpus 17a en tenia 16** |
| — | — | **54** | **Tipologia natural amb suport incoherent (v18): NIX fora del nínxol natural, CAV fora de les cavitats, PR fora de la superfície de roca; `ND` sempre legal — és un dubte, no una contradicció** |
| — | — | **55** | **`Jamb_Fabric_Reveal = 1` fora de la finestra (v19, forma de la 49; REVISADA v23): la finestra és `Jambs ∈ {0, 2}` i `Sys_Portal` Present\* — sense obertura, l'esqueixada no té subjecte. Amb el portal Absent solapa l'exempció de la regla 3: totes dues diuen la veritat** |
| — | — | **56** | **Rol de mènsules `Platform support` o `Both` amb `Sys_Platform` a `Absent` o `Not applicable` (v19): si sostenien plataforma, el sistema és present o atestat; la direcció inversa ja la cobreix la 7** |
| — | — | **57** | **Codi d'estructura que no comença pel codi del seu sector (v20, bloc B): la concordança que el cas EA46 va trencar en silenci** |
| — | — | **58** | **Coronament present sense format (v20, bloc C): semàntica de worklist, ND és resposta completa; els atestats (3) queden fora — el format POT ser atestable, no exigible** |
| — | — | **59** | **Camp adjacent al portal amb valor i `N_Chamber_Bodies = 0` (v20, bloc E): sense cambra no hi ha accés** |
| — | — | **60** | **Dues o més fases sense evidència registrada (v20, bloc G): ND és legal, la branca es pot buidar** |
| — | — | **61** | **Detall de maçoneria amb valor mentre `Masonry_Present` és 0 o 9 (v21, bloc 2): espill de les 24/25 sobre la declaració de fàbrica. El NULL és faena de la 17, no d'aquesta** |
| — | — | **62** | **`Sys_Platform` obert amb E, F i G tots confirmats a 0 (v21, bloc 1): el cap de la massa no és una plataforma — Z queda reservat al sistema volat. Els 9 NO la disparen: el cas v18 C1 (suports irresolubles) es conserva intacte** |
| — | — | **63** | **Camp mètric amb valor mentre `Metrics_Available` és 0 o 9 (v22, bloc B): espill de la 61 sobre el porter mètric. El NULL és legítim ací i NO és faena de la 17: el camp viu fora de les llistes a propòsit** |
| — | — | **64** | **Segles de cronologia amb el senyal C14 a no (v22, bloc C): el projecte data només per radiocarboni — un segle sense C14 darrere no té font. ND no aplica: el camp o té font o es buida** |
| — | — | **65** | **Cambra present o atestada amb el recompte de murs a 0 (v23, bloc 4): una cambra té almenys la façana. El NULL no dispara: una cambra atestada de murs incomptables és NULL, no 0** |
| — | — | **66** | **Murs comptats sense cambra present ni atestada (v23, bloc 4): o l'estat de la cambra es revisa (Atestada perduda?), o el recompte baixa al 0 derivable — els murs comptats no eren murs de cambra. Les files que dispara SÓN la revisió del canvi de semàntica: cap UPDATE cec** |
| — | — | **67** | **`EA_Number` absent o divergent del codi (v23, bloc 5): el número és derivat — reteclejar el codi al formulari el recalcula. Aquesta regla és la resposta a la preocupació que fonamentava la decisió v20 de no emmagatzemar-lo** |

*L'exempció de Timber_Brackets a la regla 1 mereix nota: E és l'únic element del vocabulari amb existència independent del seu sistema — una mènsula aïllada no ha d'haver portat cap plataforma. Per això existeixen Timber_Bracket_Role i la tipologia MEN. La taula 4.5 llista E sota Sys_Platform per al GATING, que és una pregunta distinta de si E implica H. **En v13 el gating del formulari recull finalment aquesta exempció** (secció 9.2).*

*Les regles **22 i 23 queden restringides a files no-ROC** en v13, i les 28–29 en són la contrapartida per a les files ROC. Sense la restricció dispararien creuadament: una estructura amb pintura rupestre però sense decoració arquitectònica faria saltar la 23 contra `Dec_Present`.*

*La **regla 27** deriva del criteri que `N_Basal_Bodies` compta masses de maçoneria i no plataformes (secció 8bis). Sobre el corpus v12 va disparar en 3 de 31 registres amb cossos basals: taxa prou baixa per a ser senyal i no soroll, però convé comprovar si el registre no està simplement a mig entrar abans de corregir-lo.*

*La **regla 31** és barata i captura precisament l'error que la conversió a metres introdueix: entrar 62 pensant en centímetres donaria una repisa de 62 m, i cap altra regla no ho detectaria.*

*Les **regles 32–33** repeteixen sobre la parella de format de pedra les dues úniques maneres com una parella ordenada pot estar mal formada, ja vigilades sobre `ID_Support`: un secundari orfe i una parella que diu dues vegades el mateix.*

*La **regla 38 dispara només amb un desajust registrat, no amb NULL**: una fila ROC sense substrat és feina pendent, no un error, i eixe cas ja el llista la regla 17. El que la regla detecta és una fila que **havia d'anar a l'altra meitat** de la taula sota el test 1 del criteri d'associació.*

*La **regla 39 és el test T2 de la discriminació G/I automatitzat** (8bis.7), i captura precisament el cas que més preocupa en camp: **la plataforma arrasada codificada com a cornisa**. Una cornisa intercòs sense segon cos és una contradicció de termes.*

*La **regla 55 és la forma de la 49** sobre el segon qualificador de compartició: un valor de farciment promogut a afirmació fora de la seua finestra. La **regla 56** tanca el forat que la condició antiga de la 9 fingia cobrir: si les mènsules sostenien plataforma, `Sys_Platform` no pot negar-la — present o atestat. La modificació de la **regla 9** (només `Is Null`) i l'exempció de la branca de portal de la **regla 3** són les dues cares del mateix criteri: **una resposta completa no és soroll**. `ND` al rol és un judici, i el portal tot a zeros amb el qualificador declarat és el cas legítim del Nus 2 del manual, no un registre a mig entrar.*

*La **regla 30 és un avís, no una prohibició**: la combinació que assenyala és el cas diagnòstic de cos perdut, i bloquejar-la impediria registrar l'evidència que la sosté. Funciona com l'avís de col·lapse — informa, no impedeix.*

*La **regla 40 és la que hauria detectat el cas EA11**. Un registre de classe *Rock art panel* no té contorn d'estructura a resseguir, de manera que les tres posicions de contacte hi són impossibles; i `ROC-PAN` està reservada a les files que pengen d'un registre PR i no ha d'aparéixer mai en un registre d'estructura. El corpus contenia exactament això: un registre PR amb una fila `ROC-PER`, mentre la descripció de la tipologia `Unclassifiable` ja nomenava eixa mateixa estructura com a exemple seu.*

*La **regla 42 és unidireccional, com la 34**, i per la mateixa raó: pot haver-hi fris amb el motiu no resolt —això és un registre honest, no incomplet—, però **un motiu constructiu al parament de façana amb M = 0 és una contradicció**. Cobreix els set motius constructius (ziga-zaga, escalonat, T i T invertida, L i L invertida, nínxols quadrats) i no `Painted band` ni `Plain colour field`, que s'apliquen sobre la superfície en lloc de modelar-la. Mira **`OVL` i `FFL`**, no només la primera: M és un tractament del parament de façana, i un fris pot recórrer-lo sense passar per damunt del portal. Calibratge sobre el corpus v16: 14 files són a `OVL` o `FFL` amb motiu constructiu; **la regla dispara en una**, i deu més tenen M encara NULL, que és feina de la regla 17.*

*Les **regles 43 i 44** cobreixen les dues direccions que la 19 deixava obertes. La 44 quasi desapareix sola amb el trasllat d'`ID_Material_Status` a 7.Mat i el seu gating, però es manté per a detectar dades heretades.*

*La **regla 46 és el que sosté la frontera entre divergència i fase** (8bis.10). Una interpretació sense el camp d'evidència emplenat és exactament el que la frontera existeix per a evitar. Quan es va especificar v17, els quatre registres que declaraven dues fases **no en tenien cap evidència entrada**.*

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

QRY_13 exporta els sis sistemes com a columnes pròpies. Un component el sistema del qual és *Not applicable* arriba com a NA d'origen: el seu 0 emmagatzemat és de farciment i no ha d'entrar en cap recompte. Les dues estratègies de mostra (restricció a Facade_Observability = Complete, o distàncies binàries amb tractament de nuls) es mantenen; QRY_15 proporciona el recompte que justifica la decisió.

## **7.4. Els tres estats del buit (v13)**

Amb el defecte NULL (1.3.b), un camp observacional té **tres estats distingibles**, i confondre'ls és la manera més fàcil de tornar a introduir el biaix que tot l'esquema evita:

| Estat | Significat | En R |
| --- | --- | --- |
| NULL | No avaluat encara | `NA`. Fora de tot denominador |
| 9 | Avaluat: la posició no és examinable | `NA`. Fora de tot denominador, però **és un judici**, i la seua freqüència mesura l'observabilitat |
| 0 | Avaluat: absent | Dada. Part del denominador |

*La diferència entre NULL i 9 no afecta els càlculs —tots dos són NA— però sí la **interpretació de la cobertura**: la proporció de 9 sobre el total avaluat és una mesura de l'observabilitat del corpus i s'ha de reportar; la de NULL només diu quant treball queda per fer. La regla 17 llista els segons.*

## **7.5. Regla general**

> Un 9 no és ni un 0 ni un 1: és una cel·la buida. Un 0 de farciment tampoc no és un 0 d'asserció: és un NA estructural. Un NULL no és un 9: és treball pendent, no un judici. Qualsevol operació que els convertisca implícitament en dades reintrodueix el biaix al resultat.

# **8. Vocabulari arquitectònic normalitzat A–X (v11)**

24 entrades: 20 elements amb camp propi, 3 sistemes (H, P, U → Sys_*) i una entrada de referència (W, absorbida per Sys_Chamber). Els criteris d'ordenació (nivell constructiu bottom-to-top, alçada, funció) i la nota terminològica nivell/cos es mantenen: N0/N1/Superior són categories funcionals, mai numeració; un cos és N1 si conté (o contenia) obertura d'accés.

| **Ll.** | **Valencià** | **Nivell** | **Sistema** | **Camp v12** |
| --- | --- | --- | --- | --- |
| A | Jàcenes basals empotrades | N0 | — (conjunt basal) | Embedded_Base_Beams |
| B | Basament | N0 | — (conjunt basal) | Base_Level |
| C | Sòcol decoratiu | N0 | — (conjunt basal) | Decorative_Socle |
| D | Muret transversal (*Transverse wall*, v19) | N0 | — (conjunt basal, desgatejat v18) | Tie_Walls |
| E | Mènsules de fusta | N0 | Component de H | Timber_Brackets |
| F | Bigues transversals | N0 | Component de H | Transverse_Beams |
| G | Filades en voladís | N0 | Component de H | Corbelled_Courses |
| **H** | **Plataforma en voladís** | N0 | **SISTEMA E+F+G** | **Sys_Platform** |
| I | Cornisa intercòs | N0/N1 | Cap (deliberadament) | Interbody_Cornice (**PEDRA**) |
| J | Cantoneres | N1 | — (conjunt cambra) | Corner_Quoins |
| K | Pilastres estructurals | N1 | — (conjunt cambra) | Structural_Pilasters |
| L | Flanc de façana | N1 | — (conjunt cambra) | Facade_Flank |
| M | Fris en relleu | N1 | — (conjunt cambra) | Relief_Frieze |
| N | Llindar | N1 | Component de P | Sill |
| O | Brancals | N1 | Component de P | Jambs |
| **P** | **Sistema portal** | N1 | **SISTEMA N+O+Q** | **Sys_Portal** |
| Q | Dintell | N1 | Component de P | Lintel (+ Lintel_Material) |
| R | Coronament | N1 | Cap (deliberadament) | Upper_Crown |
| S | Biga de suport del ràfec | Sup. | Component de U | Eave_Beam |
| T | Superfície del ràfec | Sup. | Component de U (PEDRA) | Eave_Surface |
| **U** | **Ràfec en voladís** | Sup. | **SISTEMA S+T (PEDRA)** | **Sys_Eave** |
| V | Mur lateral de cambra | N1 | — (conjunt cambra) | Return_Wall |
| W | Mur posterior | N1 | Referència de vocabulari | (absorbit per Sys_Chamber; tipus a Rear_Closure_Type) |
| X | Tancament superior de la cambra | N1 | — (conjunt cambra) | Chamber_Roof (+ Chamber_Roof_Type) |

**Criteris operatius de replicabilitat (fixats per escrit):** pilastra (K) = integrada al pla del parament; contrafort = sobreïx del pla (cap documentat). Mènsula (E) = perpendicular a la façana, encastada, en voladís; biga (F) = paral·lela, salvant llum. Suport de plataforma vs mènsula aïllada = sosté (o sostenia) la plataforma o no. Flanc de façana (L) = parament dins del pla de façana; mur lateral de cambra (V) = perpendicular, en retorn cap al penyal — les dades confirmen que són variables independents. **Cornisa (I) = filada volada de pedra: si l'element volat és de fusta, és E o F i el que hi ha és plataforma.** **X = tancament efectiu en contacte amb la fàbrica: sense contacte, no hi ha sostre.**

*Nomenclatura valenciana revisada en v16: **flanc de façana** (abans «ala de façana», calc de *flank* que no deia res en camp), **mur lateral de cambra** (abans «mur de retorn», que nomenava l'operació geomètrica i no l'element; *en retorn* es conserva com a glossa tècnica) i **fons de cambra** per a `Rear_Closure_Type` (abans «tancament posterior», que en arqueologia funerària s'entén com el segellat de la tomba i era font d'error silenciós de registre). **Cap nom de camp anglés canvia**: `Facade_Flank`, `Return_Wall` i `Rear_Closure_Type` són els publicats.*

*Nota de corpus: **no s'ha documentat cap llosa ni muret de tancament d'obertura**. Es registra com a **resultat i no com a silenci** — si a DW/LP no hi ha segellat de l'accés, això diu alguna cosa sobre l'accés recurrent a la cambra i sobre la relació entre H i P.*

# **8bis. Criteris operatius fixats en v13–v14**

## **8bis.0. Assignació de domini**

> **Cinc valors (0/1/2/3/9): els 20 elements A–X, i només ells.**
> El **valor 3** exigeix que l'evidència de la pèrdua siga **de naturalesa distinta de la cosa perduda** (un encaix no és una mènsula). El **valor 2** exigeix **poder inferir l'extensió original**.
>
> **Tres valors (0/1/9): tota la resta.** Atributs de tècnica, capes contínues, processos, vestigis mobles i judicis agregats. La distinció de fons és **discret vs continu**: el que es compta element a element admet gradació de conservació; el que s'estén sobre superfícies, no.

Criteris de registre que no es dedueixen de l'esquema i que han de quedar escrits per a garantir la replicabilitat entre observadors i campanyes.

## **8bis.1. Suport geològic: què registra `ID_Support`**

> **`ID_Support`** = la forma del relleu que **rep la càrrega** de l'estructura, siga per repòs (compressió sobre un pla) o per **confinament** (transmissió lateral a les parets).
> **Prova operativa:** *què cauria si eixa forma no hi fora?*
>
> **`ID_Support_Secondary`** = la forma que **estabilitza o allotja** sense rebre la càrrega principal: defineix l'espai, protegeix o contribueix a la fixació.

**Transparència dels elements construïts:** per a identificar el primari cal seguir la cadena de càrrega cap avall a través de la maçoneria fins a trobar roca. La cadena acaba **on la maçoneria toca roca**, que pot ser per sota (repòs) o **lateralment** (confinament) — un basament construït dins d'una diàclasi transmet la càrrega a les parets de l'escletxa, no al sòl que hi haja davall.

*Per què la prova contrafactual i no «per on entra la càrrega»:* la primera es resol mirant; la segona obliga a decidir prèviament quina és la cadena, que és on és fàcil equivocar-se.

**El que NO és criteri:** «primari = N0, secundari = N1». Els casos reals la contradiuen en les dues direccions, i hi ha una raó estructural: quan la cambra s'allotja directament a la roca, el suport primari és a cota de N1 per definició. Nota orientativa admissible, no norma: *el primari sol situar-se a cota inferior al secundari — conseqüència de la gravetat — però no necessàriament al nivell N0.*

**Ambigüitat cavitat / diàclasi eixamplada:** si l'allotjament té sostre rocós propi i profunditat cap a dins → cavitat; si és l'eixamplament d'un pla de fractura que continua amunt i avall → diàclasi. Indici corroborador: `Chamber_Roof_Type = Natural bedrock`. Si al model no es resol, registrar les dues amb la parella ordenada i explicar-ho a `Notes`.

**Casos de referència** (Diablo Wasi):

| Cas | Primari (rep càrrega) | Secundari (estabilitza / allotja) |
| --- | --- | --- |
| Basament construït dins d'una diàclasi + cambra en cavitat preexistent | Diàclasi (confina el basament) | Cavitat (allotja la cambra) |
| Torre sobre micro-repisa + cambra empotrada entre dos cossos de pedra | Micro-repisa (repòs) | Diàclasi (empotra la cambra) |

*El repertori geològic és el mateix i el que varia és **quina forma rep i quina estabilitza**. La diàclasi apareix en tots dos casos amb funció distinta: és la forma més versàtil del repertori.*

## **8bis.2. La plataforma (H) no compta com a cos basal**

> **`N_Basal_Bodies` compta masses de maçoneria** (elements de tipus B), **no plataformes**.

Un cos és N1 si té o tenia obertura d'accés, i N0 si no en té. Però H no és un cos: és un **sistema**, un element d'interfície que tanca N0 i precondiciona N1, no una massa de maçoneria superposada. Si comptara com a cos basal, un mausoleu sobre plataforma de mènsules i un mausoleu sobre pòdium de maçoneria donarien tots dos `2, 1` — i són precisament dues solucions constructives que cal poder distingir. Ho comprova la regla 27.

## **8bis.3. Mènsula aïllada dins d'una estructura**

Una mènsula que no forma part del sistema de plataforma es registra igualment: `Timber_Brackets` = 1 (o 2 si està trencada), `Timber_Bracket_Count`, `Timber_Bracket_Role` = *Isolated*, amb `Sys_Platform` = *Absent* o *Not observable*.

> **`Role = Isolated`** és una **absència verificada** (s'ha mirat al voltant i no hi ha plataforma). **`ND`** és la manca d'observació. És la mateixa distinció 0/9 aplicada al camp de rol: marcar `ND` per prudència quan s'ha verificat fa perdre el cas de tot denominador.

*`Isolated` dins d'una EA-CAM no vol dir «mènsula aïllada al penyal» sinó «mènsula que no forma part del sistema H d'aquesta estructura» — la contrapartida intraestructura de la tipologia MEN. La tipologia de l'estructura no canvia. Registrar-ho a les dues escales permet preguntar després si les mènsules sense plataforma es concentren en sectors concrets, que és material per a OE3.*

**Hipòtesis funcionals** (per exemple, que una mènsula fera de politja) **no van a `Timber_Bracket_Role`**: barrejarien una observació verificable amb una inferència, i a partir d'ací seria impossible saber si un `Isolated` significa «he comprovat» o «no se m'ha acudit cap hipòtesi». Van a `T_ARCH_FEATURES` amb `Feature_Code = Other (see Notes)` i els indicis a `Notes`.

## **8bis.4. Art rupestre: quan és fila i quan és fitxa (revisat v16)**

Quatre preguntes **geomètriques** en ordre. La primera que es compleix decideix, i l'ordre **és** la prioritat: ja no cal cap nota sobre què mana quan dues condicions concorren.

> 1. **El motiu és sencer sobre la fàbrica?** → **pigment de l'estructura** (`Pigment_*`; detall per posició a `T_DECORATIONS` amb posició arquitectònica).
> 2. **Té part sobre la fàbrica i part sobre la roca?** → **`ROC-OVL`, solapada**. Una fila, `Substrate = Mixed`, i el costat fàbrica registrat a 3.Acab.
> 3. **Ressegueix el contorn de l'estructura —present o desapareguda— en contacte tangencial?** → **`ROC-PER`, perimetral**.
> 4. **No toca la fàbrica enlloc, però és dins del mateix accident del farallò que l'allotja?** → **`ROC`, pròxima**. Si cal travessar una vora de banc, una cornisa natural o un buit per a arribar-hi, → **fitxa PR pròpia**, vinculada per `T_GROUPS` (*Rock art cluster*).

*Per què no un llindar mètric:* la distància és de segona passada i l'associació s'ha de decidir en entrar el registre, de manera que un criteri mètric acabaria aplicant-se a ull. A més, en una paret vertical 1 m no significa el mateix a tot arreu: una marca a 1,20 m però dins de la mateixa escletxa està relacionada; una a 60 cm separada per un banc que trenca la continuïtat, no necessàriament. **Sí que val com a desempat** per a decidir l'agrupació de les fitxes PR — no la naturalesa del registre.

*El criteri v15 tenia un **encavalcament real**: enviava «els tests 2 o 3» a `ROC` i el test 3 també a `ROC-PER`, de manera que una pintura que emmarcara una obertura casava amb dues entrades. La reformulació geomètrica el resol i, de passada, fa que **la posició codifique quina pregunta va decidir**: la traçabilitat que abans s'havia d'anotar a `Notes` és ara estructural i consultable.*

*Els casos dubtosos tenen sortida sense forçar: fitxa PR + agrupació no perd res — el conjunt continua sent consultable per QRY_17 i la pintura conserva coordenades, orientació i visibilitat pròpies.*

## **8bis.6. Criteri de diferenciació (NOU v16) — val per a tot el vocabulari A–X**

> Un element A–X **és present** quan hi ha un component **físicament diferenciat** del parament contigu — per format de pedra, dimensió, material o tractament. **Si la vora de l'obertura és maçoneria contínua amb el mur, sense cap peça distinta, l'element és absent (0).**

Aplicat al portal: una obertura resolta deixant simplement un buit al mur és `Sill = 0`, `Jambs = 0`, `Lintel = 0` — **i `Sys_Portal` = *Present complete* igualment**, perquè l'obertura hi és. No és cap contradicció sinó la divisió del treball entre sistema i components:

| | Registra |
| --- | --- |
| `Sys_Portal` | Que **existeix obertura d'accés** |
| N, O, Q | Si **cada posició es va resoldre amb un element diferenciat** |

*Coherent amb el que ja hi havia: la regla 11 fa incompatible tindre cossos cambra amb `Sys_Portal = Absent` —un cos és N1 si conté o contenia obertura— i `Opening_Width_m` / `Opening_Height_m` deliberadament **no** estan gatejades pel sistema.*

*El **0 resultant és un resultat de primera magnitud**. Que a DW l'obertura es resolga sovint sense llindar ni brancals diferenciats és una afirmació sobre la inversió de treball i sobre la seqüència operativa — H01 i H04 directament. Amb el criteri contrari («si hi ha obertura, hi ha brancals») la variable desapareixeria i tots els portals serien iguals.*

*El criteri **no és nou de fons**: B ja el portava implícitament («basament com a element constructiu diferenciat»). Ací es fa general i explícit. Conseqüència operativa: **la transferència v15→v16 porta N, O i Q a NULL**, perquè un valor emmagatzemat sota el criteri antic no és un judici sota el nou.*

**Portal asimètric.** Un brancal diferenciat i l'altre no: sense evidència de pèrdua no es pot decidir entre disseny asimètric i brancal perdut, i el domini no té valor per a «asimètric per disseny». Amb menys de tres casos al corpus **no es crea camp**: valor **2** més nota amb fórmula constant a `Systems_Notes` (`asymmetric: left only`). Si el patró es repeteix, eixe serà el senyal que en cal un — el mateix criteri aplicat a la hipòtesi de politja i a la relació cromàtica.

## **8bis.7. Filades en voladís (G) contra cornisa intercòs (I) (NOU v16)**

Els dos ocupen la mateixa franja física i per això es confonen, sobretot quan a penes queden vestigis de la plataforma. **La cornisa no exigeix absència de plataforma**: poden coexistir, i ho fan quan hi ha plataforma volada al cos basal i cornisa marcant la junta amb la cambra. Fer-les excloents crearia un artefacte.

Diferència categorial: **la plataforma és una superfície** (horitzontal, practicable, amb profunditat útil, i té sistema propi); **la cornisa és una junta** (lineal, sense profunditat útil, no practicable, i no té sistema deliberadament). H és **N0** i no intranivell: és la culminació del cos basal, es construeix des de baix i pot existir sense res damunt; I és **N0/N1** perquè exigeix que hi haja dos cossos.

Quatre tests en ordre de prioritat:

| # | Pregunta | Resposta |
| --- | --- | --- |
| **T1** | Hi ha evidència positiva d'E o F? (encaixos, forats passants, fusta conservada, mènsules al buit) | **Sí → plataforma.** És l'evidència que sobreviu millor: la roca conserva l'encaix molt després que la fàbrica caiga |
| **T2** | Hi ha cos construït damunt? | **No → no pot ser cornisa.** Sense segon cos, ni present ni atestat, la filada volada és **G**. Automatitzat a la regla 39 |
| **T3** | El voladís és progressiu o d'un sol gest? | Filades successives cadascuna més enfora → **G**. Una o dues filades que sobreïxen per igual, amb el mur reprenent la vertical damunt → **I** |
| **T4** | Continuïtat | Recorre tota l'amplària de la façana marcant la junta → **I**. Es concentra on hi havia superfície practicable, sovint desbordant l'amplària del cos → **G** |

**Quan cap test decideix.** El domini de cinc valors no té valor per a «observat però no identificable». Sortida: `Interbody_Cornice = 9` i `Sys_Platform = Not observable`, amb `Systems_Notes` descrivint el que es veu **i fila a `T_ARCH_FEATURES`** amb el detall morfològic. No és elegant, però l'alternativa —triar una de les dues per convenció— fabrica dades, i `T_ARCH_FEATURES` existeix precisament perquè el que no encaixa quede registrat en lloc de perdre's. Sobre el corpus són pocs casos.

*Solució de segona passada: la profunditat de volada és mesurable amb Metashape sobre el model. Registrada a `Systems_Notes` amb fórmula constant, un llindar es podrà aplicar retroactivament quan el corpus diga on és la frontera real. **Ara no es fixa cap llindar mètric**: seria inventar-lo.*

## **8bis.9. L'estructura desapareguda que només conserva la pintura perimetral (NOU v17)**

**No era una pregunta oberta sinó una incoherència viva.** La descripció de l'entrada `Unclassifiable` de `L_TYPOLOGY` diu literalment *«evidència insuficient per a classificar una estructura CONSTRUÏDA — p. ex. EA11: perímetre de pigment sense construcció conservada»*, amb `Record_Class = Built funerary structure`. Però DW-S01-EA11 estava desat com a `PR Rock Art`, i la seua única fila de decoració —la que porta la nota *«aparentment en forma de U invertida (geomètrica-ortogonal)»*— en posició `ROC-PER`. Doble contradicció: el vocabulari deia una cosa i el registre una altra, i **`ROC-PER` no pot existir en un registre d'art rupestre**, perquè significa «ressegueix el contorn de l'estructura» i un panell aïllat no en té cap.

> **És un registre d'ESTRUCTURA**, amb tipologia `Unclassifiable`, el vocabulari A–X a **9** (no a 0), una fila a `T_LOST_ELEMENTS` amb `Evidence_Scope = Whole structure` i `ID_Evidence_Type = Pigment on bedrock`, i la fila `ROC-PER` amb els quatre trams i `Outline_Geometry` emplenats.

Tres raons:

**(a) Unitat de registre.** Allí hi havia una construcció. Registrada com a art rupestre, desapareix del recompte d'estructures — i els recomptes per sector i per jaciment són exactament el que mesuren H02 i H03. Excloure les desaparegudes fa que es compte **preservació** i es llija com a **densitat constructiva**: el problema del denominador de 7.1, agreujat, perquè no és un valor absent sinó una **fila** absent.

**(b) La maquinària existia i no s'usava.** L'abast `Whole structure` i el tipus d'evidència *pigment sobre penya ara nua de la fàbrica que el portava* no tenien cap fila al corpus. Es van crear per a aquest cas.

**(c) L'alternativa perd informació i no en guanya cap.** Com a art rupestre, la banda queda registrada; que ressegueix un contorn rectangular alineat amb elements arquitectònics, no.

**Criteri de frontera per als casos futurs**, perquè la decisió no depenga del dia:

> **(Reescrita en v20, bloc F.)** La geometria del contorn és un **indici**, no el criteri: a La Petaca hi ha estructures perimetrades amb bandes en C i fins i tot en O, i una corba de traç impecable segueix sense afirmar arquitectura. El veredicte estructura-vs-panell és el **judici d'atestació complet a nivell de registre** — alineació amb restes constructives, relació amb la cavitat, encaixos, el conjunt — i es declara on sempre: la **tipologia** del registre més les files d'atestació. La banda mateixa només diu el que es veu: el traç (tipus *definida/amorfa*) i els trams coberts (dels quals `QRY_27` deriva la forma).

*Aquesta és la faena analítica de la distinció entre les dues formes en U. **En v17 vivia a un camp propi, `Outline_Geometry`, retirat en v17a**: només discriminava dins de la banda perimetral, i plegat dins del tipus fa la mateixa faena sense una columna que quedava buida a la majoria de files. La regla 40 vigila les dues direccions de l'error.*

## **8bis.10. La fàbrica com a comptador que substitueix `T_BODIES` (NOU v17)**

**El problema.** `Stone_Format`, `Stone_Working`, `Masonry_Type`, `Masonry_Quality` i `Mortar_*` són un joc per registre. Si el cos basal és de blocs irregulars sense treballar i la cambra de blocs tabulars semiescairats, només es pot dir `Mixed` — i eixe `Mixed` **destrueix exactament la informació d'interés**: que hi ha dues fàbriques i per tant possiblement dues mans, dues fases o dues intencions. Toca H01, H03 i H04 alhora.

**Els números que hi ha.** De 36 registres, 33 tenen el recompte de cossos entrat, i **24 (73 %) tenen més d'un cos**. Ara bé, la immensa majoria són un cos basal més una cambra, que és el mausoleu estàndard; els casos amb pila vertical real són **quatre registres amb dues cambres superposades** i **quatre amb dues fases declarades**.

**Els números que NO hi ha, i per què no.** La xifra que ha de decidir `T_BODIES` —quants registres tenen cossos amb fàbrica realment distinta— **no està a la base i no hi pot estar**, perquè cap camp la recull. I hi ha una cosa pitjor que la falta de dades: **`Stone_Format` i `Stone_Working` només estan emplenats en 4 registres de 36**, i `Mixed` no s'ha usat mai. Decidir ara com estructurar unes dades encara no recollides és la situació on més fàcil és construir una cosa elegant que després no encaixa.

**L'argument que decideix: la unitat d'anàlisi.** Si els cossos passen a ser files pròpies, `QRY_13` passa d'unes 36 files a unes 60. Sembla un guany i no ho és: **dos cossos de la mateixa estructura no són observacions independents** — els va fer la mateixa gent, el mateix dia, amb la mateixa pedra. Ficar-los com a files separades a un clúster o a una anàlisi de correspondències infla la mostra amb repeticions i produeix agrupacions que reflecteixen **quantes cossos té cada estructura**, no quines pràctiques constructives s'assemblen. Això no invalida `T_BODIES` com a eina descriptiva, però sí que vol dir que **no resol l'anàlisi**, i part de l'atractiu era eixe.

> **LA UNITAT PRINCIPAL D'ANÀLISI CONTINUA SENT L'ESTRUCTURA. `T_BODIES` s'ajorna, però amb un comptador posat.**

*No és ajornar el problema una segona vegada: és **instal·lar la mesura que falta** perquè la decisió es prenga amb una xifra i no amb una impressió. Quan la segona passada haja omplit el format i el treball de la pedra dels 36 registres, la xifra existirà com a **subproducte**, sense cap treball addicional; ara caldria endevinar-la mirant fotos.*

**El camp.** `Fabric` (*Single / Between bodies only / Within a body / Not observable*). Un sol camp, quatre valors.

***Els tres valors plurals són excloents, i el criteri que els fa excloents importa.*** *Un primer disseny els va definir per **on** hi ha divergència, i se solapaven: un cos plural per dins també és plural respecte del veí. El criteri correcte no és on hi és, sinó **si la divisió coincideix amb els cossos o els travessa**. `Between bodies only` = cada cos és uniforme per dins però difereixen entre ells. `Within a body` = almenys un cos té més d'una fàbrica per dins.*

*I eixa distinció és tot el sentit de l'operació: **`Between bodies only` és exactament el cas que `T_BODIES` resoldria, i `Within a body` és el cas que NO resoldria** — el mateix `Mixed` reapareixeria un nivell més avall.*

*El camp de localització de la divergència (base contra cambra, entre cambres) es va descartar: amb dos o tres casos previstos al corpus, un segon camp no aportava prou i el detall va a `Arch_Notes`.*

***El camp no es gateja.*** *Un primer intent el tancava amb menys de dos cossos, i era un error de fons: **una estructura d'un sol cos pot tindre dues fàbriques** —eixe és el valor `Within a body`—, de manera que el bloqueig tancava justament el cas que el camp existeix per a recollir. L'única combinació impossible és `Between bodies only` amb un sol cos, i la vigila la **regla 45**, que per això marca només eixe valor i no els dos plurals.*

*El detall qualitatiu —si el que canvia és el format, el treball, l'aparell o el morter— va a `Arch_Notes`, que `QRY_18` ja recupera. Quan arribe el moment de decidir, seran quatre o quinze casos i es llegiran d'una tirada.*

**Frontera amb les fases constructives.** Un canvi de fàbrica a mitja alçària d'una cambra és sovint **precisament l'evidència d'una fase**, de manera que cal fixar-ho amb la mateixa lògica que `Vertical association`:

> **La divergència de fàbrica és una OBSERVACIÓ**: la pedra canvia, i això es veu.
> **La fase constructiva és una INTERPRETACIÓ**: hi va haver dos moments de construcció, i això s'argumenta.

*Pot haver-hi canvi de fàbrica sense afirmar cap fase —un canvi de proveïment, o dos paletes el mateix dia—; quan sí que se sosté la fase, l'argument va a `Phase_Evidence`. Regles 45 i 46.*

**Camps candidats si `T_BODIES` prospera, congelats de moment:** `Stone_Format`, `Stone_Format_Secondary`, `Stone_Working`, `Masonry_Type`, `Masonry_Quality`, `Mortar_*`, `Chamber_Roof` i el criteri de numeració de `Body_No`.

## **8bis.8. Massís adossat («banqueta») — element documentat i no promocionat (v16)**

Massa construïda que **reposa sobre el suport natural adossada a l'estructura, sense sostindre cap cos** i sense continuïtat de fàbrica amb ella; la cara superior és una superfície practicable a cota intermèdia. **No és B** (un basament sosté el que hi ha damunt), **no és H** (H vola sobre el buit, a tallant i moment; això reposa en compressió), **no és I** (una junta entre cossos no és un cos) i **no és registre independent** (no té coherència constructiva pròpia).

Es registra a **`T_ARCH_FEATURES`** amb `Feature_Type = Abutted mass` —literalment idèntic a totes les files, o la consulta de recompte no les trobarà— i la funció a `Notes` amb fórmula constant (`function: access`). **Llindar de revisió: 3 casos**; si s'arriba, es planteja l'element Y al delta següent.

*El nom «escaló d'accés» o «plataforma d'accés» **es descarta**: pressuposa la funció, que és exactament l'error que `Platform_Function` existeix per a evitar. «Plataforma» col·lisionaria a més amb H.*

## **8bis.5. Límits de la lectura fotogràfica**

En una fotografia frontal, **un pla que avança i un pla que retrocedeix es confonen**: la vora d'una escletxa es llig com un banc que sobreïx, i un ressalt inferior com una escletxa que flanqueja. Això afecta directament la codificació de `ID_Support`, i els criteris de 8bis.1 (prova contrafactual, cavitat vs diàclasi) **exigeixen visió tridimensional o accés directe**. Les assignacions de suport fetes només sobre fotografia han de quedar marcades com a tals a `Doc_Basis`.

# **9. Formulari i entrada de dades (v16)**

## **9.1. Estructura**

F_STRUCTURES, 12 pestanyes: 1.Id. (identificació + detall del suport), 2.Arq. (morfologia, façana/paisatge, maçoneria i morter, fases), 3.Acab. (revoc + pigment), 4.Dec. (**dues subseccions**: decoració arquitectònica amb `Dec_Present` + F_DECORATIONS, i pintura rupestre amb `RockArt_Present` + F_ROCKART), 5.Estat (conservació + base documental i observabilitat), 6.Bio., 7.Mat. (amb `Cultural_Materials_Present` com a porta), 8.Cron. (amb el subformulari F_DATING governat per `C14`), 9.Mètr. (dimensions, obertura, mètrica del suport, volumetria, coordenades — segona passada), 10.Doc., 11.Sist. (els cinc sistemes dalt de tot i els grups de components a continuació, amb l'avís de col·lapse), 12.Extra (subformularis F_ARCH_FEATURES, F_CONNECTIONS i **F_LOST_ELEMENTS**).

Tots els combos de domini són de dues columnes: la columna emmagatzemada (anglés) està oculta i l'etiqueta (valencià) és visible, de manera que reanomenar etiquetes mai no toca les dades. LimitToList és sempre cert (Access ho imposa amb la columna lligada oculta); la via d'escapament per a excepcions genuïnes és el valor «Other (see Notes)» + el camp Notes.

## **9.2. Gating de tres nivells**

- **Nivell 1 — Record_Class → pestanyes.** *Natural funerary context* tanca 2.Arq; *Structural trace* i *Rock art panel* tanquen a més 6.Bio i 7.Mat; *Rock art panel* tanca també 11.Sist. Un registre sense classificar ho manté tot obert. Matís deliberat: una cova (CAV) conserva 11.Sist obert — DW-S01-EA09 té mènsules i plataforma dins d'una cavitat natural. 3.Acab resta oberta per a totes les classes: un panell d'art rupestre es defineix pel seu pigment.
- **Nivell 2 — portes → grups.** Cada Sys_* obri el seu grup només en Present* o Attested lost. `Cultural_Materials_Present` tanca tot 7.Mat inclòs l'estat de conservació; `Human_Remains` tanca el detall bio; `Plaster_Present`, `Pigment_Present` i `Mortar_Present` tanquen els seus detalls; `Dec_Present` i `RockArt_Present` governen cadascun el seu subformulari; `C14` governa el de datacions (llig un booleà, no un byte: és l'única porta amb eixa forma).
- **Nivell 3 — element → detall.** Timber_Bracket_Count i Timber_Bracket_Role només amb E a 1/2/3; Platform_Surface_Material bloquejat amb rol Isolated; Lintel_Material només amb Q a 1/2/3; Chamber_Roof_Type només amb X a 1/2/3. Les regles 7–10 i 14 passen de detectades a impossibles. **v16: el nivell 3 del material de la cornisa desapareix amb el camp** — una cornisa és sempre de pedra per definició.

**Correcció v13 — l'element E queda fora del gating del seu sistema.** La v12 el tractava com un component més de `Sys_Platform`, de manera que amb el sistema absent el camp es bloquejava i l'emplenat ràpid oferia posar-lo a 0: **activament fals** en el cas d'una cambra funerària amb una mènsula lateral que no sosté cap plataforma. E és l'únic element del vocabulari amb existència independent del seu sistema, i la regla 1 de la bateria ja l'exemptava — el gating v12 aplicava la pertinença de la taula 4.5 sense recollir una exempció que la validació ja feia, una incoherència interna del paquet i no un criteri nou.

En v13: `Timber_Brackets` queda **sempre editable**; `Timber_Bracket_Count` i `Timber_Bracket_Role` depenen **d'E** i no del sistema; i l'emplenat ràpid de `Sys_Platform` **exclou E** dels components que posa a 0 o 9 (F i G s'hi mantenen: són molt més difícils de llegir com a cosa distinta de components de plataforma).

## **9.3. Emplenat ràpid amb confirmació**

En tancar una porta, el formulari ofereix escriure el valor coherent a tot el grup: sistema a Absent/Not applicable → components a 0 (farciment) i detall a buit; Not observable → components a 9; revoc o pigment a **0 o 9** → detall a buit (amb el domini de cinc valors, els valors 2 i 3 **obrin** el detall: una capa parcial o desapareguda pot tindre color i extensió documentats); vestigis Absent → Mat_* a 0 (exactament el que exigeix la regla 19); vestigis ND → Mat_* a 9; restes humanes 0/9 → detall bio a 0/9 i MNI a buit; Mortar_Present 0 → Mortar_Type = None dry-laid (silenciós, regla 26). **Mai no s'escriu sense un Sí explícit, i mai en navegar entre registres** (Form_Current només activa i desactiva: l'anàleg de la regla B de la migració). L'assignació per codi no dispara AfterUpdate, de manera que no hi ha recursió.

## **9.4. Avisos i validacions en viu**

**v20:** l'avís de col·lapse baixa a baix de tot de 11.Sist (a la fila 3 tapava el desplegable del Sistema interfície) amb el text alineat amb la regla 6 esmenada — que ara anomena la **tercera eixida**: *confirmar la llegibilitat de la superfície a notes*. Cap conversió automàtica 0→9 sobre col·lapsades: la llegibilitat es jutja superfície a superfície (registre model: DW-S01-EA13, cinc sistemes amb quatre respostes epistèmiques distintes). I el botó **«Valida aquest registre»** obri la bateria filtrada pel codi actual (`F_VALIDATION` sobre `QRY_16_Validation_Check`): una sola font de veritat, cap lògica duplicada — el motiu pel qual es va descartar el codi de colors per camp.

Marcar un 3 en qualsevol element mostra el recordatori de la regla 2 (evidència a 12.Extra). Amb estat Collapsed apareix l'avís en roig de la regla 6 a 11.Sist. El substrat de pigment amaga l'opció Plaster quan Plaster_Present = 0 (prevenció de R13). Als subformularis —on Enabled afectaria la columna sencera de totes les files— la lògica per fila és validació BeforeUpdate: F_CONNECTIONS bloqueja una cronologia direccional sobre una connexió sense adossament (R20) i autoassigna Contemporary a la junta travada; F_LOST_ELEMENTS exigeix el codi d'element quan l'abast és Element (R21).

## **9.5. Prevenció i detecció**

El principi de disseny: **el formulari preveu en el moment d'entrada; QRY_16 detecta a escala de corpus.** La bateria continua sent imprescindible perquè el gating es pot esquivar (taules obertes a mà, importacions, Centre de confiança desactivat), i perquè algunes regles són intrínsecament de corpus (5, 17). Executeu QRY_16_Validation_Check periòdicament durant l'entrada de dades.

## **9.6. Requisit tècnic**

La injecció dels mòduls VBA del gating i de les validacions exigeix «Confiar en l'accés al model d'objectes de projectes VBA» al Centre de confiança d'Access. Si està desactivat, el formulari es construeix igualment i tot queda editable; només falta l'automatisme, i l'script ho avisa a la finestra immediata en lloc de fallar en silenci.

# **10. Scripts i rutes d'actualització**

| **Fitxer** | **Sub públic** | **Ús** |
| --- | --- | --- |
| chachapoya_DB_v17.bas | BuildDB() | Construcció completa sobre una BD EN BLANC: 22 taules, valors per defecte (0 als 20 elements, NULL a la resta; sistemes a Absent), lookups, 25 relacions, **33 consultes**, i l'índex únic `UQ_CONN_PAIR`. Informa del **recompte de consultes creades contra les previstes** |
| chachapoya_Form_v17_val.bas | BuildForm() | Després de BuildDB(): **7 subformularis** (F_CONN_IN inclòs), F_STRUCTURES amb 12 pestanyes, combos de dues columnes, captions DAO, gating i validacions de subformulari |
| **chachapoya_EXPORT_v16.bas** | ExportAll() | **S'executa sobre la BD v16 d'origen.** Escriu 8 CSV inspeccionables i editables. **Cap identificador travessa la frontera**: totes les claus alienes ixen com el valor llegible que apunten |
| **chachapoya_IMPORT_v17.bas** | ImportAll() | **S'executa sobre una v17 acabada de construir i buida.** Es nega a arrancar si `T_STRUCTURES` ja té files. Resol totes les claus **abans** de `rs.AddNew`, **normalitza les parelles de connexió girant la cronologia amb elles**, en deduplica les col·lisions i informa per fila |
| **chachapoya_PATCH_v17a.bas** | PatchV17a() | **S'executa sobre una v17 AMB DADES.** Afig els cinc camps de notes, reanomena `Notes` → `Doc_Notes` (renom de TableDef: **conserva els valors**), bolca `Outline_Geometry` a `Notes` de cada fila i elimina la columna, refà la meitat rupestre de `L_DEC_TYPE` **repuntant les files abans d'esborrar cap entrada**, i retira la regla 41. Deixa **sense tipus** les files de banda perimetral i les llista a `QRY_22` |
| diagnostic_v13.bas | RunDiagnostic() | **Només lectura.** Comprovacions de coherència i de variància sobre el corpus; no modifica cap dada ni cap esquema. Es pot executar tantes vegades com calga |

*Els scripts de migració de versions anteriors queden fora del paquet: la base es construeix sempre de zero i les dades hi arriben per la parella export/import, amb un fitxer inspeccionable entremig.*

**El CSV és l'espai de treball de la revisió, no només un pas intermedi.**

*En la transferència v16→v17 el pes de la revisió es reparteix molt distint que en l'anterior. **Cap camp de `T_STRUCTURES` no ha canviat de criteri**, de manera que no hi ha ni una sola columna `OLD_` ni cap valor enfosquit: els camps nous de v17 no tenen columna d'origen i arriben NULL per absència, que és el correcte. Tota la revisió es concentra en **`T_DECORATIONS`**, i concretament en les **set files** que portaven `Position_Relative`: els quatre trams ixen en blanc, el valor antic viatja a `OLD_Position_Relative`, i s'omplin a mà contra la fotografia.*

***Per què eixes set no es poden convertir automàticament.*** *`Both` no vol dir «esquerra i dreta»: en tres de les set s'està usant, amb la descripció aparcada a `Notes`, per a dir una **U invertida** — esquerra, damunt i dreta. Traduir `Both` a `Span_Left = 1` i `Span_Right = 1` afirmaria que el tram del coronament és **absent** precisament a les files on hi és. Un 0 equivocat és pitjor que un buit, perquè el 0 és un judici i el buit no.*

***Les connexions són el cas contrari, i convé no confondre'ls.*** *La seua cronologia també canvia de redacció, però **no de significat**: «A anterior a B» i «A és l'anterior» diuen el mateix del mateix lloc. La conversió és determinista i l'importador la fa sol, **inclosa la inversió quan normalitza la parella**. Enfosquir-les hauria fabricat feina de revisió a partir d'un canvi de nom, que és el contrari del que serveix la regla conservadora. El que sí que demana ull humà és la **parella duplicada**: després de normalitzar, les dues files del corpus col·lideixen, i l'importador conserva la primera i **informa de en què discrepen** en lloc de triar en silenci.*

*El que s'escriga a les columnes vives és un judici; el que quede en blanc queda honestament buit i `QRY_21_V17_Review` el llistarà.*

**Ordre d'execució:** `ExportAll()` sobre la còpia local real → revisió dels CSV → `BuildDB()` i `BuildForm()` sobre una base en blanc → `ImportAll()` → finestra d'immediat → `QRY_16_Validation_Check` → `QRY_21_V17_Review` (i `QRY_19_V16_Review`, que continua vigent: el format i el treball de la pedra segueixen sense entrar en 32 registres).

## **10.1. Correccions de dades pendents (no són de disseny)**

*Aquestes tres no depenen del delta i s'han d'aplicar sobre la còpia local abans o durant la transferència.*

**(a) Files de decoració sense posició — 16 de 56.** El `RecordSource` dels dos subformularis de 4.Dec feia `INNER JOIN L_STRUCT_BODY`, i un INNER JOIN elimina les files amb `ID_Struct_Body` nul. Quinze eren propostes automàtiques de la migració v11 des dels booleans `RA_*`, totes amb la nota «Completar posició»; una era una fila mig entrada. **Existien, comptaven a les consultes i a les regles 22–23 i 28–29, i no es podien ni veure ni completar des de la interfície** — la nota que porten demanava completar la posició, i completar-la era exactament el que el formulari impedia. **Resolt en v17 amb `LEFT JOIN`**, amb les files sense posició assignades a la meitat arquitectònica, on `ND` ja viu; `QRY_21` les llista.

**(b) Tipologia de DW-S01-EA11.** Passa de `PR Rock Art` a `Unclassifiable`, amb la composició de registre de 8bis.9.

**(c) Registre de prova.** El registre 36 té `Code = 'c'` i cap tipologia: fila a eliminar.

Les restriccions tècniques de VBA/JET es mantenen: cap continuació de línia, ASCII estricte en codi (cp1252), patró sql = sql & "…", paraules reservades evitades, N−1 parèntesis per a N taules en JOIN, Private per a tot auxiliar, DROP previ per a objectes regenerables, combos configurats per ControlSource.

## **10.2. Apèndix: concordança de recodificacions (auditoria v20)**

*Els registres amb sufix a/b es van recodificar com a codis propis durant l'entrada del corpus complet; les connexions preserven les parelles. Les referències dels deltes i documents anteriors es llegeixen amb aquesta taula.*

| Codi antic | Codi actual | Parella preservada a T_CONNECTIONS |
| --- | --- | --- |
| DW-S04-EA02a / EA02b | DW-S04-EA02 / **DW-S04-EA21** | EA02 ↔ EA21 |
| DW-S04-EA12a / EA12b | DW-S04-EA12 / **DW-S04-EA20** | EA12 ↔ EA20 |
| DW-S01-EA22a / EA22b | DW-S01-EA22 / **DW-S01-EA71** | EA22 ↔ EA71 |
| DW-S01-EA40a / EA40b | DW-S01-EA40 / **DW-S01-EA72** | EA40 ↔ EA72 |

*(Pendent de la verificació d'Esteve: si existia cap altra parella amb sufix — p. ex. EA07b — completar la taula.)*


----

## Annex v25 — Referència tècnica del delta

### T_METRIC_REVIEW (taula nova)

| Camp | Tipus | Notes |
|---|---|---|
| ID | COUNTER PK | |
| ID_Structure | LONG NOT NULL | FK inline `FK_MREV_STRUCT` → T_STRUCTURES.ID |
| Check_Code | TEXT(4) NOT NULL | C1 / C2 / SLOT / C3a / C3b / C5 / C7 |
| Field_Name | TEXT(40) | buit a les files Investigate-only (C3a) |
| DB_Value | TEXT(60) | valor actual (buit si NULL) |
| Proposed_Value | TEXT(60) | proposta derivada del CSV |
| Evidence | MEMO | etiquetes de les shapes o derivació |
| CSV_Batch | TEXT(80) | nom del fitxer audit de la passada |
| Detected_On | DATETIME | |
| Decision | TEXT(12) | Pending / Accept / Keep DB / Investigate |
| Decided_On | DATETIME | l'estampa l'investigador en firmar |
| Applied_On | DATETIME | l'estampa `ApplyMetricReview` |
| Notes | MEMO | |

Cicle de vida: `CheckMetrics` esborra les Pending pròpies i repobla
(idempotent); les firmades es conserven i no es ressusciten (clau
estructura+creuament+camp+proposta); `ApplyMetricReview` executa
només Accept amb `Applied_On` NULL i estampa la data. Conversió de
tipus per `TableDefs` (numèric vs text) en el moment d'aplicar.

### Camps nous de T_STRUCTURES

| Camp | Tipus | Domini | Gating |
|---|---|---|---|
| Jamb_Fabric | TEXT(25) | Monolithic slab / Composite slab-masonry / Fabric as jamb / ND | `ElemHas(Jambs)` (patró de la casa: NULL editable) |
| Niche_Partition | TEXT(15) | Present / Absent / Attested lost / Not observable | tipologia NIX (per `L_TYPOLOGY.Name Like 'NIX*'`) |

El domini de cinc valors 0/1/2/3/9 **continua exclusiu** dels camps
A–X: els dos camps nous són textuals a l'estil `Rear_Closure_Type`.

### Columnes noves de L_ELEMENTS

`Term_Lit TEXT(30)` i `Lit_Source TEXT(40)`, poblades per a B
(platform-base), I (cornice) i M (frieze), font Guengerich 2014.
Buides on no hi ha equivalent publicat.

### Consultes noves

| Consulta | Contingut |
|---|---|
| QRY_30_Metric_Review | Files Pending de T_METRIC_REVIEW amb el codi de l'estructura (per a la worklist; la firma es fa al full de dades de la taula) |
| QRY_31_JambFabric_Pending | Estructures amb Jambs ∈ {1,2,3} i Jamb_Fabric NULL |
| QRY_32_NichePartition_Pending | NIX amb Niche_Partition NULL |

`QRY_28_Worklist` incorpora les tres com a fonts «Revisio metrica»
i «Classificacio»; el salt de pestanya del doble clic coneix les
paraules clau noves (Metrica→9.Metr, Jamb_Fabric→11.Sist,
Niche_Partition→1.Id).

### Mapa token→lletra del pipeline (fixat a CheckMetrics; lletra→camp es llig en viu de L_ELEMENTS)

bra→E · tbeam→F · ebeam→S · bbeam→A · corn→I · crown→R · sill→N ·
jamb→O · lint→Q · pil→K · tie→D · roof→X · plat→Z · port→P ·
eave→U · wall: pla f→L, plans r/l→V, pla b sense camp (s'omet) ·
slot→Support_Modified · int→Area_m2 · led→Support_Depth_m.
