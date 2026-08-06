**ESQUEMA DE LA BASE DE DADES ARQUEOLÒGICA**

*La Petaca i Diablo Wasi (Leymebamba, Amazonas, Perú)*

chachapoya_DB_v14.bas + chachapoya_Form_v14_val.bas

Sub BuildDB() + Sub BuildForm() | Microsoft Access JET SQL | Esteve Ribera Torró

*Versió 15 del document — actualitzada segons el codi v14 (agost 2026). Substitueix la versió 12 i consolida el delta v12→v13 amb les seues set revisions.*

**Sense ruta de migració.** A diferència de les iteracions anteriors, la v13 no transforma la base existent: els 35 registres es reintrodueixen a mà sobre una base construïda de zero. La decisió elimina tot l'aparat de migració i, sobretot, elimina els valors heretats de defectes de versions anteriors — sota v13, **cada valor emmagatzemat és un judici deliberat**, que és la precondició per a publicar qualsevol percentatge.

# **1. Resum general**

| **Element** | **Valor** |
| --- | --- |
| Taules principals | T_STRUCTURES (127 camps), T_DATING, T_INDIVIDUALS, T_GROUPS, T_DECORATIONS, T_ARCH_FEATURES, T_CONNECTIONS, T_LOST_ELEMENTS |
| Taules lookup | L_SITES, L_SECTORS, L_TYPOLOGY, L_SUPPORT, L_STATUS, L_MATERIAL_STATUS, L_VOL_METHOD, L_COORD_METHOD, L_GROUP_TYPE, L_CAMPAIGN, L_STRUCT_BODY, L_DEC_TYPE, L_ELEMENTS, L_LOST_EVIDENCE |
| Total taules | 22 |
| Total relacions | 25 (inclou l'autoreferenciant de T_STRUCTURES, les dues de T_CONNECTIONS, la del suport secundari i les quatre de T_LOST_ELEMENTS) |
| Consultes SQL | 28 (família QRY_01–QRY_18; la bateria QRY_16 comprén cinc consultes auxiliars, tres parcials de regles i la consulta unió amb **31 regles**) |
| Formulari | F_STRUCTURES — 12 pestanyes + subformularis F_DECORATIONS, F_ROCKART, **F_DATING**, F_ARCH_FEATURES, F_CONNECTIONS i F_LOST_ELEMENTS. Gating de tres nivells amb emplenat ràpid confirmat. Etiquetes UI en valencià; valors emmagatzemats en anglés |
| Dominis observacionals | **20 camps amb domini de cinc valors** 0/1/2/3/9: **els elements A–X, i només ells**; **26 camps amb domini 0/1/9**: atributs de tècnica, capes contínues, processos, vestigis mobles, qualificadors i judicis agregats; 5 camps de sistema TEXT amb domini de sis o quatre valors |
| Valors per defecte | **Dos, deliberadament**: 0 als 20 camps d'element governats per un `Sys_*` (la regla 4 hi exigeix el zero de farciment); **NULL a la resta** — buit = no avaluat encara, 9 = avaluat i no examinable, 0 = avaluat i absent |
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

**(i) Record_Class a L_TYPOLOGY.** Dels 35 registres actuals, 5 no són estructures: qualsevol percentatge sobre 35 és erroni. Cada tipologia porta una classe de registre — *Built funerary structure* (EA-MAU, EA-CAM, EA-PLA-R/V, Unclassifiable), *Natural funerary context* (NIX, CAV), *Structural trace* (MEN), *Rock art panel* (PR), *Pending classification*. La classe filtra cada anàlisi i governa el gating de pestanyes. **MIX s'elimina** (tot EA-CAM és mixt per definició: no distingia res) i **ND es desdobla** en *Unclassifiable* (un resultat: evidència insuficient per a classificar una estructura construïda) i *Not yet classified* (un estat de treball, exclòs de tota consulta analítica).

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

# **2. T_STRUCTURES (128 camps)**

## **2.1. Identificació (8)**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| ID | COUNTER PK | Clau primària autonumèrica |
| Code | TEXT(20) NN | Codi PALP: [SITE][SECTOR]-[TIPTYPE][NUM] — ex: DW-S01-EA01 |
| ID_Sector | LONG NN | FK → L_SECTORS |
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
| Floor_Plan | TEXT(20) | Planta: Rectangular / Sub-rectangular / Square / Circular / Sub-circular / Trapezoidal / Irregular / ND |
| N_Built_Walls | INTEGER | Nombre de murs construïts (0–4) |
| Length_m / Width_m / Height_m | SINGLE | Dimensions exteriors (m). Segona passada mètrica |
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

## **2.2c. Maçoneria i morter (6) — H01/H04, cf. Toyne i Anzellini 2017**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| Masonry_Quality | TEXT(20) | Good / Moderate / Poor / ND |
| Masonry_Type | TEXT(30) | Well-coursed / Irregular-coursed / Uncoursed / Mixed / ND |
| Mortar_Present | BYTE | 0/1/9. 0 = fàbrica en sec verificada. **No promocionat (v14)**: no és una capa que es perd per zones sinó un atribut de la tècnica — un mur en sec no ho és a trossos. El que varia amb la conservació és la visibilitat, registrada a `Facade_Observability` i `Mortar_Notes` |
| Mortar_Type | TEXT(30) | Mud / Mud with gravel / Mud with organics / None dry-laid / ND. Amb Mortar_Present = 0 el gating l'autoassigna a None dry-laid (regla 26) |
| Chinking_Stones | BYTE | 0/1/9. Ripio / falques entre carreus. Independent del morter: una fàbrica en sec pot dur falques. **No promocionat a cinc valors (v13)**: al corpus v12 presentava 26 presències i cap absència verificada, de manera que no té variància i no pot alimentar cap prova. Es manté com a descriptor tècnic (T&A 2017); la seua universalitat és ella mateixa un resultat sobre la tradició constructiva de DW |
| Mortar_Notes | TEXT(150) | Notes sobre morter i juntes |
| **Recessed_Frame** | BYTE | 0/1/9. Marc reculat de façana. **Traslladat ací en v14** (abans al bloc del portal): és un qualificador del **pla de façana** i no del portal — el recul afecta el parament sencer i el portal hi queda inscrit |

## **2.3. Sistemes constructius (6 camps de sistema)**

**Domini dels tres sistemes d'element (Sys_Platform, Sys_Portal, Sys_Eave):** *Present complete / Present partial / Attested lost / Absent / Not applicable / Not observable*. **Domini dels dos conjunts d'agrupació (Sys_Base, Sys_Chamber):** *Present / Absent / Not applicable / Not observable*. Per defecte: Absent.

| **Camp** | **Elem.** | **Governa (gating)** | **Descripció** |
| --- | --- | --- | --- |
| Sys_Platform | **H** | E, F, G + Timber_Bracket_Count/Role, Platform_Surface_Material, Platform_Function | Sistema plataforma: E+F+G → H |
| Sys_Portal | **P** | N, O, Q + Lintel_Material | Sistema portal: N+O+Q → P |
| Sys_Eave | **U** | S, T | Sistema ràfec: S+T → U. Sempre pedra a T i U |
| Sys_Base | — | A, B, C, D | Conjunt basal (agrupació sense lletra pròpia) |
| Sys_Chamber | — | J, K, L, M, V, X + Chamber_Roof_Type, Rear_Closure_Type | Conjunt cambra (agrupació; absorbeix l'antic W) |
| Platform_Function | — | — | Access / Circulation / Construction / Support / Multiple / Undetermined. Forma separada de funció: OE3 existeix per a determinar per a què servia la plataforma, i la morfologia no ho ha de pressuposar |

*Els sistemes es resolen primer (pestanya 11.Sist els presenta dalt de tot): es decideix que hi ha plataforma abans de comptar-ne les mènsules. Un sistema obri el seu grup només en Present* o Attested lost; sota Absent / Not applicable / Not observable els components porten el 0 de farciment.*

## **2.4. Elements A–X — Nivell 0 / basament (10)**

**Domini dels camps d'element: 0 = Absent / 1 = Present complet / 2 = Present parcial / 3 = Desaparegut atestat (exigeix T_LOST_ELEMENTS) / 9 = No observable. Per defecte 0.**

| **Camp** | **Tipus** | **Elem.** | **Descripció** |
| --- | --- | --- | --- |
| Embedded_Base_Beams | BYTE | **A** | Jàcenes horitzontals empotrades dins la maçoneria del basament |
| Base_Level | BYTE | **B** | Basament com a element constructiu diferenciat (pòdium) |
| Decorative_Socle | BYTE | **C** | Tractament decoratiu del basament |
| Tie_Walls | BYTE | **D** | Murets perpendiculars al faralló (ancoratge/compartimentació) |
| Timber_Brackets | BYTE | **E** | Mènsules de fusta empotrades a la roca (component de H). Criteri: perpendicular a la façana, encastada, en voladís. Exempta de la regla 1: és l'únic element amb existència independent del seu sistema |
| Timber_Bracket_Count | INTEGER | **E** | Nombre de mènsules visibles. Obligatori amb E present (regla 8): el nombre porta l'argument de la plataforma |
| Timber_Bracket_Role | TEXT(25) | **E** | Platform support / Isolated / Both / ND. Obligatori amb E present (regla 9). Isolated bloqueja Platform_Surface_Material (regla 7) |
| Transverse_Beams | BYTE | **F** | Bigues transversals (component de H). Criteri: paral·lela a la façana, salvant llum |
| Corbelled_Courses | BYTE | **G** | Filades de pedra en voladís creixent (component de H, variant lítia) |
| Platform_Surface_Material | TEXT(20) | (H) | Material de la superfície de la plataforma: Timber / Stone / Mixed / ND |

## **2.4b. Interfície N0/N1 i coronament (3) — elements I, R (sense sistema, deliberadament)**

| **Camp** | **Tipus** | **Elem.** | **Descripció** |
| --- | --- | --- | --- |
| Interbody_Cornice | BYTE | **I** | Cornisa intercòs entre cossos superposats |
| Interbody_Cornice_Material | TEXT(20) | **I** | Stone slabs / Wooden beams / Mixed / ND |
| Upper_Crown | BYTE | **R** | Coronament superior / coping |

## **2.5. Elements A–X — Nivell 1 / cos principal i sistema portal (9)**

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

## **2.6. Elements A–X — Cambra i zona superior (6)**

| **Camp** | **Tipus** | **Elem.** | **Descripció** |
| --- | --- | --- | --- |
| Return_Wall | BYTE | **V** | Mur perpendicular al pla de façana, retornant cap al penyal (abans Lateral_Walls) |
| Rear_Closure_Type | TEXT(20) | — | Natural bedrock / Built masonry / Mixed / ND. Com es tanca la cambra pel darrere. L'element W desapareix com a camp (absorbit per Sys_Chamber); el tipus sobreviu perquè una cambra que usa la roca com a tancament posterior no es va construir com una que alça un mur |
| Eave_Beam | BYTE | **S** | Biga de suport del ràfec (component de U; pot ser fusta) |
| Eave_Surface | BYTE | **T** | Superfície del ràfec (component de U; sempre pedra) |
| Chamber_Roof | BYTE | **X** | Cambra tancada per dalt, amb la solució que siga |
| Chamber_Roof_Type | TEXT(25) | (X) | Natural bedrock / Built masonry / Built timber and slabs / Mixed / ND. La decisió que l'anàlisi necessita: roca natural vs obra (regla 14) |

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

## **2.7b. Paisatge i orientació (2) — H06**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| Facade_Orientation | TEXT(5) | N / NE / E / SE / S / SW / W / NW / ND |
| Visibility_Valley | TEXT(10) | High / Medium / Low / ND |

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
| **Cultural_Materials_Present** | BYTE | **NOU v14**. 0/1/9. Porta de 7.Mat, anàleg exacte de `Human_Remains`. 0 o 9 bloquegen tot el bloc **inclòs l'estat de conservació**, i el quick-fill emplena els `Mat_*` (regla 19) |
| ID_Material_Status | LONG | FK → L_MATERIAL_STATUS (**Good/Fair/Poor/ND**, sense *Absent* des de v14). Criteri: vestigis reduïts a pols o fragments irreconeixibles són `Present = 1` + `Poor`; `Present = 0` es reserva per a cambra buida |
| Looting / Fire_Damage / Animal_Activity / Modern_Access | BYTE | 0/1/9. Causes, independents dels graus |

## **2.10. Bioarqueologia (8) i materials culturals (7) — BYTE 0/1/9**

Idèntics a la versió anterior: Human_Remains (governa la resta del bloc bio; amb 1, MNI és obligatori — regla 18), MNI, Anatomical_Connection, Mummification, Funerary_Bundles, Dispersed_Remains, Flexed_Position, Bone_Burning; Mat_Textiles, Mat_Wood, Mat_VegFiber, Mat_Ceramics, Mat_Fauna, Mat_DeerAntler, Mat_Other. Les classes de registre *Structural trace* i *Rock art panel* tanquen els dos blocs (regles 15–16).

## **2.11. Cronologia (3) i fases constructives (2)**

C14 (YESNO, metadada de corpus), Chrono_Start_Cent, Chrono_End_Cent; Construction_Phases, Phase_Evidence (C14 / Stratigraphy / Superposition / Mortar / ND). Les fases sostenen l'argument morfològic-estructural de H03 sense dependre de datacions absolutes; la direccionalitat entre estructures adjacents es registra ara a T_CONNECTIONS.Chrono_Relation (secció 3.6).

## **2.11b. Camps de notes (6, NOU v14)**

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
| Body_No | INTEGER | Número del cos constructiu |
| Color | TEXT(20) | Red / White / Ochre / None / ND. **`Both` retirat en v13** |
| Color_Secondary | TEXT(20) | Color acompanyant: *Red / White / Ochre*, **sense `ND`** (v14). Buit en els casos monocroms, que són la majoria. Buit o ple **és** la distinció monocrom/policrom: cap camp derivat |
| Substrate | TEXT(20) | Plaster / Masonry stone / Bedrock / ND. Resol els casos mixtos per posició |
| Notes | TEXT(255) |  |

*Des de v11 és l'únic registre de decoració; QRY_02 s'hi reconstrueix i QRY_02a en deriva les banderes per motiu quan cal el format ample.*

*Des de v13 conté **també la pintura rupestre associada**, distingida per `Level_Type` de la posició: `ROC` per a la rupestre, la resta per a l'arquitectònica. Una taula, dos conjunts, un discriminador — separar-los és un `WHERE`, i obtindre el conjunt sencer no exigeix cap unió.*

*La parella `Color` + `Color_Secondary` substitueix el valor `Both` amb avantatge, perquè diu quin color domina. Es va adoptar per als casos bicroms amb els colors **junts o contigus dins d'un mateix motiu** (un camp clar amb vora d'un altre to), on descompondre en dues files inventaria dos motius on n'hi ha un i duplicaria la posició. Quan els colors ocupen posicions distintes, en canvi, la descomposició en files continua sent la via correcta: són dues observacions.*

## **3.5. T_ARCH_FEATURES — registre flexible**

Sense canvis: Feature_Code (llista amb «Other (see Notes)» com a via d'escapament del LimitToList), Present, Feature_Count, Material, Notes.

## **3.6. T_CONNECTIONS — connexions físiques entre estructures (OE3, H03)**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| ID | COUNTER PK |  |
| ID_Struct_A / ID_Struct_B | LONG NN | FK → T_STRUCTURES.ID (extrems de l'aresta) |
| Connection_Type | TEXT(30) | Llista tancada: Abutted vertical joint / Superposition / Bonded joint / Shared support / **Vertical association** (v14) / Aerial connection / Other (see Notes). La llista tancada és el que fa computable la regla 20 |
| **Chrono_Relation** | TEXT(20) | NOU v11: A earlier than B / B earlier than A / Contemporary / Undetermined. Converteix el graf no dirigit en potencialment dirigit (H03). La direcció es llig de la junta: només una junta vertical adossada o una superposició la poden portar (regla 20); una junta travada implica Contemporary (el formulari l'autoassigna). L'ordre A/B no s'ha d'invertir després de l'entrada |
| Confidence | TEXT(10) | High / Medium / Low |
| Notes | TEXT(150) |  |

*`Vertical association` (v14) registra una relació de verticalitat **observable** entre estructures —proximitat, posició relativa, morfologia del faralló— sense afirmar-ne el mecanisme: la hipòtesi de politja o de suport d'escala continua sent especulativa i va a `T_ARCH_FEATURES` amb els indicis a `Notes`. `Confidence` en registra la seguretat. Material per a OE3.*

***Criteri quan concorren dues relacions*** (junta vertical entre fàbriques sobre una base compartida): **la junta mana sobre el suport**, perquè la junta porta la direcció cronològica (regla 20) i el suport compartit no.*

*El cas d'ús paradigmàtic: un complex amb junta vertical no travada es registra com a dues entrades amb codis PALP consecutius, enllaçades ací i agrupades a T_GROUPS; si el model 3D revela filades basals travades sota la junta aparent, correspon un registre únic amb Construction_Phases = 2.*

## **3.7. T_LOST_ELEMENTS — NOVA v11 — evidència del valor 3**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| ID | COUNTER PK |  |
| ID_Structure | LONG NN | FK → T_STRUCTURES (cascade delete) |
| Element_Code | TEXT(2) | FK → L_ELEMENTS.Code. Obligatori quan l'abast és Element (regla 21); el formulari només ofereix elements reals, no sistemes |
| ID_Evidence_Type | LONG NN | FK → L_LOST_EVIDENCE |
| Evidence_Scope | TEXT(20) | Element / Body / Whole structure. Una fila amb abast Body o Whole structure cobreix tot el vocabulari de l'estructura (absorbeix l'antic Lost_Body_Evidence) |
| ID_Position | LONG | FK → L_STRUCT_BODY |
| Notes | MEMO |  |

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

## **L_TYPOLOGY (reconstruïda v11: Record_Class, sense MIX, ND desdoblat)**

| **Nom** | **Record_Class** | **Descripció** |
| --- | --- | --- |
| EA-MAU Mausoleum/Chullpa | Built funerary structure | Estructura construïda (3+ murs + sostre artificial) sobre repisa. Predominant a La Petaca |
| EA-CAM Funerary Chamber | Built funerary structure | Cavitat natural tancada per una façana construïda. Predominant a Diablo Wasi |
| EA-PLA-R Ledge Platform | Built funerary structure | Plataforma constructiva sobre repisa natural |
| EA-PLA-V Aerial Platform | Built funerary structure | Plataforma artificial sobre bigues i lloses, sense repisa de suport |
| NIX Natural Niche | Natural funerary context | Petita cavitat natural (<1 m²). Ossari o enterrament secundari |
| CAV Cave/Cavern | Natural funerary context | Gran cavitat natural (>1 m²) amb ús funerari o ritual documentat |
| PR Rock Art | Rock art panel | Motiu pictòric sobre roca, documentat de forma independent |
| MEN Isolated Bracket | Structural trace | Element estructural aïllat. Evidència de xarxa de circulació aèria perduda |
| Unclassifiable | Built funerary structure | Evidència insuficient per a classificar una estructura CONSTRUÏDA (ex: EA11, perímetre de pigment sense construcció conservada). Els contextos naturals sempre són classificables com a NIX o CAV |
| Not yet classified | Pending classification | Estat de treball: pendent de revisió manual. Exclosa de tota consulta analítica |

*MIX s'ha eliminat: tot EA-CAM és mixt per definició, de manera que no distingia res. DW-S01-EA09 (cavitat amb mènsules i plataforma) es reclassifica MIX → CAV en la migració.*

## **L_SUPPORT (11 valors; v13: desdoblament de Fissure/Crack, v14: Ground)**

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
| ND | ND | Forma de suport no determinada |

*El valor `Combined` es va retirar en v8; els suports compostos es registren com a parella ordenada `ID_Support` + `ID_Support_Secondary` (criteri a la secció 8bis).*

## **L_ELEMENTS — NOVA v11 — el vocabulari A–X com a lookup (24 files)**

Columnes: Code (únic, UQ_ELEM_CODE), Name_EN, Name_VAL, Level_Type (N0 / N0-N1 / N1 / SUP), Sys_Group (Platform / Portal / Eave / buit), Is_System, Field_Name, Description. Les files H, P i U porten Is_System = True i Field_Name = Sys_*; la fila W es conserva com a referència de vocabulari sense camp propi. Sys_Base i Sys_Chamber no apareixen a Sys_Group: són agrupacions sense lletra.

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

## **L_STRUCT_BODY (14 entrades; v13: tres posicions ROC i `Sort_Order`)**

| **Code** | **Name** | **Level_Type** | **Sort_Order** |
| --- | --- | --- | --- |
| BAS | Base level (B) | N0 | 10 |
| SOC | Socle (C) | N0 | 20 |
| COR | Interbody cornice (I) | N1 | 30 |
| QUO | Corner quoin (J) | N1 | 40 |
| PIL | Pilaster (K) | N1 | 50 |
| FFL | Facade flank (L) | N1 | 60 |
| JAM | Jamb (O) | N1 | 70 |
| OVL | Over-lintel | N1 | 80 |
| CRO | Upper crown (R) | N1 | 90 |
| EAV | Eave (U) | SUP | 100 |
| **ROC** | **Adjacent bedrock** | **ROC** | 200 |
| **ROC-PER** | **Opening perimeter (rock)** | **ROC** | 210 |
| **ROC-PAN** | **Panel (no architectural ref.)** | **ROC** | 220 |
| ND | Not determined | ND | 999 |

*Descriu **només la posició dins d'un cos**; de quin cos es tracta ho diu `T_DECORATIONS.Body_No`.*

*`Level_Type = ROC` és el **discriminador analític** entre decoració arquitectònica i pintura rupestre. `ROC-PAN` existeix per a les files que pengen d'un registre d'art rupestre **aïllat** (tipologia PR), que és també un registre de `T_STRUCTURES` i per tant descriu els seus motius per aquesta mateixa taula: sense eixa entrada, tornarien a caure en `ND`.*

*`ND` significa «posició no determinada» i **no s'ha d'usar mai per a posicions no arquitectòniques**: per a això hi ha les entrades ROC. Confondre-ho és el mateix error que confondre 0 amb 9.*

*`Sort_Order` (v13) restaura l'ordre constructiu bottom-to-top: l'ordenació per `Level_Type` + `Name` deixava el ràfec al final (Level_Type = SUP) i cada nivell en ordre alfabètic.*

## **L_DEC_TYPE (19 entrades; v13: sis tipus rupestres)**

**Repertori arquitectònic (13):** T-shaped niche (i inv.), L-shaped niche (i inv.), Zigzag, Stepped motif, Frieze / Greca, Triangular motif, Painted band, Square niche, Plain colour field, Decapitation scene, ND. «Plain colour field» converteix T_DECORATIONS en el registre general de tractament de superfície per posició.

**Repertori rupestre (6, NOU v13):** RA Anthropomorphic, RA Zoomorphic, RA Geometric, RA Abstract, RA Amorphous stain, RA Perimeter band.

*Els dos repertoris **no són compartits** — es va comprovar contra el corpus—, de manera que les entrades rupestres s'afigen en lloc de fondre's. Alguns tipus genèrics (`Plain colour field`, `Painted band`, `ND`) serveixen les dues bandes i apareixen a les dues llistes del formulari; per això el discriminador entre conjunts és `Level_Type` i no `ID_Dec_Type`.*

*`RA Perimeter band` mereix atenció: marques que emmarquen una obertura, un llindar o el contorn d'una estructura. Quan la fàbrica ja no hi és, encreuar-la amb `T_LOST_ELEMENTS`.*

# **5. Relacions (25)**

Les 21 de la versió anterior més les quatre de T_LOST_ELEMENTS:

| **Nom** | **Pare** | **Filla** | **Camp fill** | **Notes** |
| --- | --- | --- | --- | --- |
| REL_STR_LOST | T_STRUCTURES | T_LOST_ELEMENTS | ID_Structure | Cascade delete |
| REL_ELEM_LOST | L_ELEMENTS (Code) | T_LOST_ELEMENTS | Element_Code | Requereix l'índex únic UQ_ELEM_CODE |
| REL_LEV_LOST | L_LOST_EVIDENCE | T_LOST_ELEMENTS | ID_Evidence_Type | Update cascade |
| REL_SB_LOST | L_STRUCT_BODY | T_LOST_ELEMENTS | ID_Position | Update cascade |

*REL_STR_SELF (autoreferenciant) i les dues de T_CONNECTIONS continuen amb dbRelationDontEnforceIntegrity (valor numèric 2).*

# **6. Consultes SQL (28)**

| **Nom** | **Descripció** | **Hip.** |
| --- | --- | --- |
| QRY_01_Typology_by_Site | Distribució de tipologies per jaciment i sector. | H01, H04 |
| QRY_02s_Decoration_Typed | Auxiliar: files de T_DECORATIONS amb tipus i posició resolts. | — |
| QRY_02a_Decoration_Flags | Banderes per motiu derivades de T_DECORATIONS (format ample per a khi-quadrat). | H01, H04 |
| QRY_02_Decoration_by_Site | Reconstruïda v11 sobre T_DECORATIONS: recomptes per motiu i jaciment. | H01, H04 |
| QRY_03 – QRY_12 | Sense canvis de funció: conservació, C14, exportació R, volumetria, exportació QGIS, fills, membres de grup, cobertura ChaXR, maçoneria, geologia-construcció. QRY_05 i QRY_07 actualitzades v11 als noms nous i als camps de sistema. | diverses |
| QRY_13_AX_Pattern_Export | Reconstruïda v11: **20 columnes AX_ + 5 columnes SYS_** amb semàntica NA (un component gatejat per un sistema Not applicable s'exporta NA, no 0, perquè el seu 0 és de farciment). La font única FillElementMap garanteix que exportació i validació no divergeixen mai. | H01, H04, H05 |
| QRY_14_Connections_Edges | Llista d'arestes amb Chrono_Relation: graf potencialment dirigit per a igraph/QGIS. | OE3, H03 |
| QRY_15_Observability_Bias | Recompte per jaciment i sector de base documental i observabilitat. | OE1 |
| QRY_16s_* (5) | Auxiliars de la bateria: Element_Values (format llarg dels 20 elements: la clau que evita ~80 branques UNION), Lost_Cover, Damage_Count, Observational_Nulls (llista de treball camp a camp), Null_Count. | — |
| QRY_16a / QRY_16b / QRY_16c | Regles 1–11, 12–21 i **22–31**. Emmagatzemades en tres parts perquè una UNION única de totes les branques supera el límit «query too complex» de JET. | — |
| QRY_16_Validation_Check | Unió de les tres parts: **31 regles**. Resultat buit = corpus coherent. | — |
| **QRY_17_RockArt_All** | **NOVA v13.** Totes les pintures del corpus amb una columna `Context` (*Associated* / *Isolated*). No necessita cap UNION: les associades i les aïllades són **totes dues files de `T_DECORATIONS`** —les primeres penjades d'una estructura, les segones d'un registre PR, que és també un registre de `T_STRUCTURES`—, de manera que és una sola consulta amb JOIN a `L_TYPOLOGY` per a saber quina mena de pare té cada fila. | OE1, H06 |
| **QRY_18_Notes_Review** | **NOVA v14.** Tots els camps de notes de tots els registres, costat per costat, filtrant els buits. Serveix per al repàs i, sobretot, per a **detectar patrons**: quan la mateixa observació apareix repetidament en text lliure, això és el senyal que hauria de ser un camp. | — |

## **6.1. La bateria de 26 regles**

No bloquejant per disseny: una restricció dura impediria registrar una absència genuïnament observada en una estructura parcialment col·lapsada. Resum:

| # | Regla | # | Regla |
| --- | --- | --- | --- |
| 1 | Component present (1/2/3) amb sistema no present (exempts: Not observable, i Timber_Brackets) | 14 | Coberta present sense tipus |
| 2 | Valor 3 (o sistema Attested lost) sense fila d'evidència a T_LOST_ELEMENTS | 15 | Structural trace amb restes humanes |
| 3 | Sistema present amb tots els components a 0 | 16 | Contingut bio/materials en classes que tanquen els blocs |
| 4 | Component ≠ 0 sota sistema Not applicable | 17 | Camps observacionals encara NULL (llista de treball, no error) |
| 5 | Estat Good amb >30% d'elements parcials o perduts | 18 | Restes humanes sense MNI |
| 6 | 0 d'asserció en estructura Collapsed | 19 | `Cultural_Materials_Present` = 0 amb Mat_* ≠ 0 |
| 7 | Material de plataforma amb rol Isolated | 20 | Cronologia direccional en connexió sense adossament |
| 8 | Mènsules presents sense recompte | 21 | Abast Element sense Element_Code |
| 9 | Mènsules presents sense rol | 22 | `Dec_Present` sense files no-ROC de decoració (v12, restringida en v13) |
| 10 | Dintell 0 amb material | 23 | Files no-ROC de decoració sense `Dec_Present` (v12, restringida en v13) |
| 11 | Cossos de cambra amb Sys_Portal Absent | 24 | Detall de revoc sota presència 0/9 (v12) |
| 12 | Pigment present sense substrat | 25 | Detall de pigment sota presència 0/9 (v12) |
| 13 | Pigment sobre revoc amb revoc absent | 26 | Mortar_Type incompatible amb Mortar_Present = 0 (v12) |
| — | — | 27 | Cossos basals comptats amb `Base_Level` (B) = 0 (v13) |
| — | — | 28 | `RockArt_Present` declarat sense cap fila ROC (v13) |
| — | — | 29 | Files ROC sense `RockArt_Present` (v13) |
| — | — | 30 | Pigment sobre penya seguint elements arquitectònics: possible cos perdut (v13, **avís**) |
| — | — | 31 | Dimensió lineal implausible (>20 m): comprovar la unitat (v14) |

*L'exempció de Timber_Brackets a la regla 1 mereix nota: E és l'únic element del vocabulari amb existència independent del seu sistema — una mènsula aïllada no ha d'haver portat cap plataforma. Per això existeixen Timber_Bracket_Role i la tipologia MEN. La taula 4.5 llista E sota Sys_Platform per al GATING, que és una pregunta distinta de si E implica H. **En v13 el gating del formulari recull finalment aquesta exempció** (secció 9.2).*

*Les regles **22 i 23 queden restringides a files no-ROC** en v13, i les 28–29 en són la contrapartida per a les files ROC. Sense la restricció dispararien creuadament: una estructura amb pintura rupestre però sense decoració arquitectònica faria saltar la 23 contra `Dec_Present`.*

*La **regla 27** deriva del criteri que `N_Basal_Bodies` compta masses de maçoneria i no plataformes (secció 8bis). Sobre el corpus v12 va disparar en 3 de 31 registres amb cossos basals: taxa prou baixa per a ser senyal i no soroll, però convé comprovar si el registre no està simplement a mig entrar abans de corregir-lo.*

*La **regla 31** és barata i captura precisament l'error que la conversió a metres introdueix: entrar 62 pensant en centímetres donaria una repisa de 62 m, i cap altra regla no ho detectaria.*

*La **regla 30 és un avís, no una prohibició**: la combinació que assenyala és el cas diagnòstic de cos perdut, i bloquejar-la impediria registrar l'evidència que la sosté. Funciona com l'avís de col·lapse — informa, no impedeix.*

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

QRY_13 exporta els cinc sistemes com a columnes pròpies. Un component el sistema del qual és *Not applicable* arriba com a NA d'origen: el seu 0 emmagatzemat és de farciment i no ha d'entrar en cap recompte. Les dues estratègies de mostra (restricció a Facade_Observability = Complete, o distàncies binàries amb tractament de nuls) es mantenen; QRY_15 proporciona el recompte que justifica la decisió.

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
| D | Muret transversal | N0 | — (conjunt basal) | Tie_Walls |
| E | Mènsules de fusta | N0 | Component de H | Timber_Brackets |
| F | Bigues transversals | N0 | Component de H | Transverse_Beams |
| G | Filades en voladís | N0 | Component de H | Corbelled_Courses |
| **H** | **Plataforma en voladís** | N0 | **SISTEMA E+F+G** | **Sys_Platform** |
| I | Cornisa intercòs | N0/N1 | Cap (deliberadament) | Interbody_Cornice |
| J | Cantoneres | N1 | — (conjunt cambra) | Corner_Quoins |
| K | Pilastres estructurals | N1 | — (conjunt cambra) | Structural_Pilasters |
| L | Ala de façana | N1 | — (conjunt cambra) | Facade_Flank |
| M | Fris en relleu | N1 | — (conjunt cambra) | Relief_Frieze |
| N | Llindar | N1 | Component de P | Sill |
| O | Brancals | N1 | Component de P | Jambs |
| **P** | **Sistema portal** | N1 | **SISTEMA N+O+Q** | **Sys_Portal** |
| Q | Dintell | N1 | Component de P | Lintel (+ Lintel_Material) |
| R | Coronament | N1 | Cap (deliberadament) | Upper_Crown |
| S | Biga de suport del ràfec | Sup. | Component de U | Eave_Beam |
| T | Superfície del ràfec | Sup. | Component de U (PEDRA) | Eave_Surface |
| **U** | **Ràfec en voladís** | Sup. | **SISTEMA S+T (PEDRA)** | **Sys_Eave** |
| V | Mur de retorn | N1 | — (conjunt cambra) | Return_Wall |
| W | Mur posterior | N1 | Referència de vocabulari | (absorbit per Sys_Chamber; tipus a Rear_Closure_Type) |
| X | Coberta de la cambra | N1 | — (conjunt cambra) | Chamber_Roof (+ Chamber_Roof_Type) |

**Criteris operatius de replicabilitat (fixats per escrit):** pilastra (K) = integrada al pla del parament; contrafort = sobreïx del pla (cap documentat). Mènsula (E) = perpendicular a la façana, encastada, en voladís; biga (F) = paral·lela, salvant llum. Suport de plataforma vs mènsula aïllada = sosté (o sostenia) la plataforma o no. Ala de façana (L) = parament dins del pla de façana; mur de retorn (V) = perpendicular, cap al penyal — les dades confirmen que són variables independents.

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

## **8bis.4. Art rupestre: quan és fila i quan és fitxa**

Quatre tests en ordre de prioritat. **El primer que es compleix decideix.**

> 1. **Continuïtat física amb la fàbrica** — el pigment toca la maçoneria, la ressegueix o continua sobre ella → **és pigment de l'estructura** (`Pigment_*`; detall per posició a `T_DECORATIONS`).
> 2. **Comparteix suport geològic** — és dins de la mateixa escletxa, repisa, cavitat o rebaix que allotja l'estructura → **associada** (fila ROC).
> 3. **Emmarca un element de l'estructura** — obertura, llindar, perímetre, encara que hi haja discontinuïtat → **associada** (fila ROC).
> 4. **Cap de les anteriors** → **fitxa PR pròpia**, vinculada per `T_GROUPS` (*Rock art cluster*) si hi ha proximitat visual.

*Per què no un llindar mètric:* la distància és de segona passada i l'associació s'ha de decidir en entrar el registre, de manera que un criteri mètric acabaria aplicant-se a ull. A més, en una paret vertical 1 m no significa el mateix a tot arreu: una marca a 1,20 m però dins de la mateixa escletxa està relacionada; una a 60 cm separada per un banc que trenca la continuïtat, no necessàriament. La distància euclidiana ignora el que estructura aquests jaciments. **Sí que val com a desempat del test 4**, per a decidir l'agrupació — no la naturalesa del registre.

*El test 2 reutilitza una decisió ja presa (`ID_Support`), de manera que no cal mesurar res.* **Anotar a `Notes` quin test s'ha aplicat**, perquè un tercer puga verificar la decisió en lloc de refer-la.

*Els casos dubtosos tenen sortida sense forçar: fitxa PR + agrupació no perd res — el conjunt continua sent consultable per QRY_17 i la pintura conserva coordenades, orientació i visibilitat pròpies.*

## **8bis.5. Límits de la lectura fotogràfica**

En una fotografia frontal, **un pla que avança i un pla que retrocedeix es confonen**: la vora d'una escletxa es llig com un banc que sobreïx, i un ressalt inferior com una escletxa que flanqueja. Això afecta directament la codificació de `ID_Support`, i els criteris de 8bis.1 (prova contrafactual, cavitat vs diàclasi) **exigeixen visió tridimensional o accés directe**. Les assignacions de suport fetes només sobre fotografia han de quedar marcades com a tals a `Doc_Basis`.

# **9. Formulari i entrada de dades (v14)**

## **9.1. Estructura**

F_STRUCTURES, 12 pestanyes: 1.Id. (identificació + detall del suport), 2.Arq. (morfologia, façana/paisatge, maçoneria i morter, fases), 3.Acab. (revoc + pigment), 4.Dec. (**dues subseccions**: decoració arquitectònica amb `Dec_Present` + F_DECORATIONS, i pintura rupestre amb `RockArt_Present` + F_ROCKART), 5.Estat (conservació + base documental i observabilitat), 6.Bio., 7.Mat. (amb `Cultural_Materials_Present` com a porta), 8.Cron. (amb el subformulari F_DATING governat per `C14`), 9.Mètr. (dimensions, obertura, mètrica del suport, volumetria, coordenades — segona passada), 10.Doc., 11.Sist. (els cinc sistemes dalt de tot i els grups de components a continuació, amb l'avís de col·lapse), 12.Extra (subformularis F_ARCH_FEATURES, F_CONNECTIONS i **F_LOST_ELEMENTS**).

Tots els combos de domini són de dues columnes: la columna emmagatzemada (anglés) està oculta i l'etiqueta (valencià) és visible, de manera que reanomenar etiquetes mai no toca les dades. LimitToList és sempre cert (Access ho imposa amb la columna lligada oculta); la via d'escapament per a excepcions genuïnes és el valor «Other (see Notes)» + el camp Notes.

## **9.2. Gating de tres nivells**

- **Nivell 1 — Record_Class → pestanyes.** *Natural funerary context* tanca 2.Arq; *Structural trace* i *Rock art panel* tanquen a més 6.Bio i 7.Mat; *Rock art panel* tanca també 11.Sist. Un registre sense classificar ho manté tot obert. Matís deliberat: una cova (CAV) conserva 11.Sist obert — DW-S01-EA09 té mènsules i plataforma dins d'una cavitat natural. 3.Acab resta oberta per a totes les classes: un panell d'art rupestre es defineix pel seu pigment.
- **Nivell 2 — portes → grups.** Cada Sys_* obri el seu grup només en Present* o Attested lost. `Cultural_Materials_Present` tanca tot 7.Mat inclòs l'estat de conservació; `Human_Remains` tanca el detall bio; `Plaster_Present`, `Pigment_Present` i `Mortar_Present` tanquen els seus detalls; `Dec_Present` i `RockArt_Present` governen cadascun el seu subformulari; `C14` governa el de datacions (llig un booleà, no un byte: és l'única porta amb eixa forma).
- **Nivell 3 — element → detall.** Timber_Bracket_Count i Timber_Bracket_Role només amb E a 1/2/3; Platform_Surface_Material bloquejat amb rol Isolated; Lintel_Material només amb Q a 1/2/3; Chamber_Roof_Type només amb X a 1/2/3; Interbody_Cornice_Material només amb I a 1/2/3. Les regles 7–10 i 14 passen de detectades a impossibles.

**Correcció v13 — l'element E queda fora del gating del seu sistema.** La v12 el tractava com un component més de `Sys_Platform`, de manera que amb el sistema absent el camp es bloquejava i l'emplenat ràpid oferia posar-lo a 0: **activament fals** en el cas d'una cambra funerària amb una mènsula lateral que no sosté cap plataforma. E és l'únic element del vocabulari amb existència independent del seu sistema, i la regla 1 de la bateria ja l'exemptava — el gating v12 aplicava la pertinença de la taula 4.5 sense recollir una exempció que la validació ja feia, una incoherència interna del paquet i no un criteri nou.

En v13: `Timber_Brackets` queda **sempre editable**; `Timber_Bracket_Count` i `Timber_Bracket_Role` depenen **d'E** i no del sistema; i l'emplenat ràpid de `Sys_Platform` **exclou E** dels components que posa a 0 o 9 (F i G s'hi mantenen: són molt més difícils de llegir com a cosa distinta de components de plataforma).

## **9.3. Emplenat ràpid amb confirmació**

En tancar una porta, el formulari ofereix escriure el valor coherent a tot el grup: sistema a Absent/Not applicable → components a 0 (farciment) i detall a buit; Not observable → components a 9; revoc o pigment a **0 o 9** → detall a buit (amb el domini de cinc valors, els valors 2 i 3 **obrin** el detall: una capa parcial o desapareguda pot tindre color i extensió documentats); vestigis Absent → Mat_* a 0 (exactament el que exigeix la regla 19); vestigis ND → Mat_* a 9; restes humanes 0/9 → detall bio a 0/9 i MNI a buit; Mortar_Present 0 → Mortar_Type = None dry-laid (silenciós, regla 26). **Mai no s'escriu sense un Sí explícit, i mai en navegar entre registres** (Form_Current només activa i desactiva: l'anàleg de la regla B de la migració). L'assignació per codi no dispara AfterUpdate, de manera que no hi ha recursió.

## **9.4. Avisos i validacions en viu**

Marcar un 3 en qualsevol element mostra el recordatori de la regla 2 (evidència a 12.Extra). Amb estat Collapsed apareix l'avís en roig de la regla 6 a 11.Sist. El substrat de pigment amaga l'opció Plaster quan Plaster_Present = 0 (prevenció de R13). Als subformularis —on Enabled afectaria la columna sencera de totes les files— la lògica per fila és validació BeforeUpdate: F_CONNECTIONS bloqueja una cronologia direccional sobre una connexió sense adossament (R20) i autoassigna Contemporary a la junta travada; F_LOST_ELEMENTS exigeix el codi d'element quan l'abast és Element (R21).

## **9.5. Prevenció i detecció**

El principi de disseny: **el formulari preveu en el moment d'entrada; QRY_16 detecta a escala de corpus.** La bateria continua sent imprescindible perquè el gating es pot esquivar (taules obertes a mà, importacions, Centre de confiança desactivat), i perquè algunes regles són intrínsecament de corpus (5, 17). Executeu QRY_16_Validation_Check periòdicament durant l'entrada de dades.

## **9.6. Requisit tècnic**

La injecció dels mòduls VBA del gating i de les validacions exigeix «Confiar en l'accés al model d'objectes de projectes VBA» al Centre de confiança d'Access. Si està desactivat, el formulari es construeix igualment i tot queda editable; només falta l'automatisme, i l'script ho avisa a la finestra immediata en lloc de fallar en silenci.

# **10. Scripts i rutes d'actualització**

| **Fitxer** | **Sub públic** | **Ús** |
| --- | --- | --- |
| chachapoya_DB_v14.bas | BuildDB() | Construcció completa sobre una BD EN BLANC: 22 taules, valors per defecte (0 als 20 elements, NULL a la resta; sistemes a Absent), lookups, 25 relacions, 28 consultes |
| chachapoya_Form_v14_val.bas | BuildForm() | Després de BuildDB(): 6 subformularis, F_STRUCTURES amb 12 pestanyes, combos de dues columnes, captions DAO, gating v14 i validacions de subformulari |
| diagnostic_v13.bas | RunDiagnostic() | **Només lectura.** Comprovacions de coherència i de variància sobre el corpus; no modifica cap dada ni cap esquema. Es pot executar tantes vegades com calga |

*Els scripts de migració de versions anteriors (`migrate_v10_to_v11.bas`, `upgrade_form_v12.bas`) queden fora del paquet v13: la base es construeix de zero i els registres es reintrodueixen.*

Les restriccions tècniques de VBA/JET es mantenen: cap continuació de línia, ASCII estricte en codi (cp1252), patró sql = sql & "…", paraules reservades evitades, N−1 parèntesis per a N taules en JOIN, Private per a tot auxiliar, DROP previ per a objectes regenerables, combos configurats per ControlSource.
