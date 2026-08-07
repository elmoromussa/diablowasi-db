**ESQUEMA DE LA BASE DE DADES ARQUEOLÒGICA**

*La Petaca i Diablo Wasi (Leymebamba, Amazonas, Perú)*

chachapoya_DB_v16.bas + chachapoya_Form_v16_val.bas

Sub BuildDB() + Sub BuildForm() | Microsoft Access JET SQL | Esteve Ribera Torró

*Versió 16 del document — actualitzada segons el codi v16 (agost 2026). Consolida el delta v15→v16, que recull la revisió de la interfície amb les fitxes de camp a la mà.*

**Construcció de zero, transferència en dos passos.** La base v16 es construeix des de zero amb els scripts canònics i les dades hi arriben per `chachapoya_EXPORT_v15.bas` → CSV inspeccionable i editable → `chachapoya_IMPORT_v16.bas`. **Cap valor arriba sense haver pogut ser mirat**, i els camps el criteri dels quals ha canviat en v16 arriben NULL: un valor transferit sense revisió afirmaria un judici que ningú no ha fet sota el criteri nou. La precondició per a publicar qualsevol percentatge continua sent la mateixa — cada valor emmagatzemat, un judici deliberat.

# **1. Resum general**

| **Element** | **Valor** |
| --- | --- |
| Taules principals | T_STRUCTURES (132 camps), T_DATING, T_INDIVIDUALS, T_GROUPS, T_DECORATIONS, T_ARCH_FEATURES, T_CONNECTIONS, T_LOST_ELEMENTS |
| Taules lookup | L_SITES, L_SECTORS, L_TYPOLOGY, L_SUPPORT, L_STATUS, L_MATERIAL_STATUS, L_VOL_METHOD, L_COORD_METHOD, L_GROUP_TYPE, L_CAMPAIGN, L_STRUCT_BODY, L_DEC_TYPE, L_ELEMENTS, L_LOST_EVIDENCE |
| Total taules | 22 |
| Total relacions | 25 (inclou l'autoreferenciant de T_STRUCTURES, les dues de T_CONNECTIONS, la del suport secundari i les quatre de T_LOST_ELEMENTS) |
| Consultes SQL | 30 (família QRY_01–QRY_19; la bateria QRY_16 comprén cinc consultes auxiliars, **quatre** parcials de regles i la consulta unió amb **39 regles**) |
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

# **2. T_STRUCTURES (132 camps)**

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

## **2.2c. Maçoneria i morter (9) — H01/H04, cf. Toyne i Anzellini 2017**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| **Stone_Format** | TEXT(30) | **NOU v16.** Irregular blocks / Tabular blocks / Laminar slabs / ND. Morfologia de la peça, producte de la fractura natural de la roca. **Abast: el parament**, no els elements singulars |
| **Stone_Format_Secondary** | TEXT(30) | **NOU v16.** Mateixos valors sense `ND`. Segon format convivint al parament (parella ordenada com `ID_Support`) |
| **Stone_Working** | TEXT(20) | **NOU v16.** Unworked / Semi-dressed / Dressed / Mixed / ND. **Ordinal**: `Mixed` i `ND` són fora d'escala i cauen de correspondències i regressions, com els *Not applicable* |
| Masonry_Quality | TEXT(20) | Good / Moderate / Poor / ND |
| Masonry_Type | TEXT(30) | Well-coursed / Irregular-coursed / Uncoursed / Mixed / ND |
| Mortar_Present | BYTE | 0/1/9. 0 = fàbrica en sec verificada. **No promocionat (v14)**: no és una capa que es perd per zones sinó un atribut de la tècnica — un mur en sec no ho és a trossos. El que varia amb la conservació és la visibilitat, registrada a `Facade_Observability` i `Mortar_Notes` |
| Mortar_Type | TEXT(30) | Mud / Mud with gravel / Mud with organics / None dry-laid / ND. Amb Mortar_Present = 0 el gating l'autoassigna a None dry-laid (regla 26) |
| Chinking_Stones | BYTE | 0/1/9. Ripio / falques entre carreus. Independent del morter: una fàbrica en sec pot dur falques. **No promocionat a cinc valors (v13)**: al corpus v12 presentava 26 presències i cap absència verificada, de manera que no té variància i no pot alimentar cap prova. Es manté com a descriptor tècnic (T&A 2017); la seua universalitat és ella mateixa un resultat sobre la tradició constructiva de DW |
| Mortar_Notes | TEXT(150) | Notes sobre morter i juntes |
*Els tres camps de pedra van **davant** de `Masonry_Type` al formulari: es tria la pedra abans d'apilar-la, i eixe és l'ordre de la cadena operativa.*

*Test operatiu de `Stone_Format`, deliberadament **funcional** perquè es llig a la façana sense mesurar res: **irregular** = la peça no té dues cares planes subparal·leles dominants; **tabular** = dues cares planes i **alçada de filada pròpia** (una pedra, una filada); **laminar** = dues cares planes però **cal apilar-ne diverses per fer una filada**. Desempat quantitatiu per als dubtosos: gruix ≤ ¼ de la dimensió major → laminar. **Dominància per superfície de parament, no per nombre de peces**: amb lloses menudes i blocs grans, comptar peces inverteix el resultat.*

*`Mixed` **no existeix** a `Stone_Format`: seria un tercer camí per a dir el que ja diu la parella, i el projecte ja ha retirat dues vegades el valor que la parella substitueix (`Combined` en v8, `Both` en v13). El secundari no porta `ND` — buit vol dir «cap segon format» (regles 32–33).*

*`Stone_Working` registra **evidència positiva de treball**: traces d'eina, aristes vives regulars, cares que trenquen el pla de fractura natural. Sense eixes traces la cara plana s'atribueix a la fractura i el valor és `Unworked`, que així no afirma «ningú no la va tocar» sinó «no hi ha evidència que la tocaren»; `ND` queda per al dubte genuí — un gres estratificat fractura en cares indistingibles del carejat.*

*Cap dels tres porta valor «no observable», com cap altre camp de domini TEXT de l'esquema: **el 9 existeix per a protegir el 0, i ací no hi ha 0 a protegir** — un parament sempre està fet d'alguna pedra amb algun grau de treball. Els dos motius de no-decisió no es col·lapsen tot i compartir `ND`: «no s'hi veia» i «no era decidible» els separa l'encreuament amb `Doc_Basis` i `Facade_Observability`, que ja registren la visibilitat. És el mateix criteri que va deixar `Mortar_Present` sense promocionar.*

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
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

## **2.4b. Interfície N0/N1 i coronament (2) — elements I, R (sense sistema, deliberadament)**

| **Camp** | **Tipus** | **Elem.** | **Descripció** |
| --- | --- | --- | --- |
| Interbody_Cornice | BYTE | **I** | Filada volada **de pedra** que marca la junta entre dos cossos superposats. **Si l'element volat és de fusta, no és una cornisa**: és E o F, i el que hi ha és plataforma |
| Upper_Crown | BYTE | **R** | Coronament superior / coping |

*`Interbody_Cornice_Material` **es retira en v16**. No perquè de fet les cornises siguen sempre de pedra sinó perquè no podrien no ser-ho: el valor *Wooden beams* no descrivia una cornisa de fusta sinó **una plataforma mal classificada**. T i U ja porten «SEMPRE PEDRA» dins de la definició del vocabulari i no tenen camp de material; la cornisa era l'excepció incoherent. El guany no és estalviar un camp sinó que **una descripció feble es converteix en un test d'identificació** — el que fa falta per a discriminar G de I (8bis.7, regla 39).*

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
| Visibility_Valley | TEXT(10) | High / Medium / Low / ND |
| **Access_Plane** | TEXT(20) | **NOU v16.** Facade / Return wall / Rear / ND. Per quin pla de l'estructura s'entra |
| **Portal_Orientation** | TEXT(5) | **NOU v16.** Mateixos valors que `Facade_Orientation`. Cap a on mira l'obertura. **Sense gating** |

*La **façana és el pla exposat**, no el pla de l'obertura. Normalment coincideixen i el mot funciona; en una estructura documentada divergeixen —obertura al mur estret perpendicular al farallò, decoració al mur llarg paral·lel— i el mot es trenca. Fixar-la per l'obertura costaria que `Facade_Orientation` deixara d'apuntar a la vall i que `Visibility_Valley` mesurara un pla que no es veu: **tota H06 degradada per a salvar la definició d'un element**. Fixar-la pel pla exposat costa una línia: L és el parament dins del pla de façana, i quan l'obertura hi és, els paraments que la flanquegen; quan no, el parament corregut.*

*`Access_Plane` i `Portal_Orientation` **no es dupliquen**: el pla és **relacional i intrínsec** (com l'arquitectura organitza exhibició i accés, comparable entre orientacions absolutes distintes); l'orientació és **absoluta** i va a QGIS contra la topografia i les rutes de circulació. Cap de les dues es dedueix de l'altra sense conéixer la geometria del cas.*

*`Portal_Orientation` **no es gateja darrere de `Sys_Portal`**, pel mateix motiu que `Recessed_Frame` no ho està: una obertura arrasada té orientació coneguda, i un portal a *Attested lost* amb orientació registrada és precisament un cas informatiu.*

*El **guany analític** és que la divergència entre els dos camps d'orientació passa a ser calculable, i per tant preguntable: si les estructures amb accés desviat del pla exposat comparteixen tipologia, sector, forma de suport o amplària de repisa. Amb un sol camp d'orientació no es podia ni formular. La regla 36 vigila la coherència quan `Access_Plane = Facade`.*

*`Access_Plane` fa a més una faena que no és de paisatge: **declara el pla de referència** sense el qual `Position_Relative` de `T_DECORATIONS` no significa res.*

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
| Body_No | INTEGER | Número del cos constructiu. **Obligatòriament buit a les files ROC** (regla 37): una pintura sobre la penya no està en cap cos |
| **Position_Relative** | TEXT(20) | **NOU v16.** Left / Right / Both / Above / Below / ND. Compartit per les dues meitats de la taula |
| Color | TEXT(20) | Red / White / Ochre / None / ND. **`Both` retirat en v13** |
| Color_Secondary | TEXT(20) | Color acompanyant: *Red / White / Ochre*, **sense `ND`** (v14). Buit en els casos monocroms, que són la majoria. Buit o ple **és** la distinció monocrom/policrom: cap camp derivat |
| Substrate | TEXT(20) | Plaster / Masonry stone / Bedrock / ND. Resol els casos mixtos per posició |
| Notes | TEXT(255) |  |

*Des de v11 és l'únic registre de decoració; QRY_02 s'hi reconstrueix i QRY_02a en deriva les banderes per motiu quan cal el format ample.*

*Des de v13 conté **també la pintura rupestre associada**, distingida per `Level_Type` de la posició: `ROC` per a la rupestre, la resta per a l'arquitectònica. Una taula, dos conjunts, un discriminador — separar-los és un `WHERE`, i obtindre el conjunt sencer no exigeix cap unió.*

*`Position_Relative` **no substitueix `Body_No`**: són eixos distints —índex vertical de cos contra posició relativa— i les files arquitectòniques continuen necessitant l'índex, encara més si prospera la descomposició per subnivells. **Convenció indispensable, o el camp és soroll: esquerra i dreta des del punt de vista de l'observador situat davant del pla**, mai des de l'estructura; el pla de referència el declara `Access_Plane`. `Both` es conserva per al tractament bilateral simètric, perquè la simetria és una afirmació arqueològica i partir-la en dues files inventaria dos motius on n'hi ha un. Amb aquest camp queda resolta també la lateralitat de `FFL`, `JAM`, `PRT` i `RTW` sense duplicar cap entrada de lookup per costat.*

*`Substrate` a les files ROC deixa de ser una tria i passa a ser un **verificador**: *Plaster* i *Masonry stone* són impossibles per definició del criteri d'associació —si el pigment toca la fàbrica, allò és pigment de l'estructura i va a l'altra meitat—, de manera que la llista del formulari es redueix a `Bedrock` i `Mixed` (aquest darrer només per a `ROC-OVL`, que per definició abasta fàbrica i roca) i les regles 38a–38b detecten les files mal encaminades. Una incoherència abans silenciosa passa a ser un error detectat.*

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

## **L_STRUCT_BODY (18 entrades; v16: −`SOC`, +4 arquitectòniques, ROC redefinides)**

| **Code** | **Name** | **Name_VAL** | **Level_Type** | **Sort_Order** |
| --- | --- | --- | --- | --- |
| BAS | Basal body (B/C) | Cos basal (B/C) | N0 | 10 |
| COR | Interbody cornice (I) | Cornisa intercòs (I) | N1 | 30 |
| QUO | Corner quoin (J) | Cantonera (J) | N1 | 40 |
| PIL | Pilaster (K) | Pilastra (K) | N1 | 50 |
| FFL | Facade flank (L) | **Flanc de façana (L)** | N1 | 60 |
| **RTW** | **Return wall (V)** | **Mur lateral de cambra (V)** | N1 | **62** |
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

# **6. Consultes SQL (30)**

| **Nom** | **Descripció** | **Hip.** |
| --- | --- | --- |
| QRY_01_Typology_by_Site | Distribució de tipologies per jaciment i sector. | H01, H04 |
| QRY_02s_Decoration_Typed | Auxiliar: files de T_DECORATIONS amb tipus i posició resolts. | — |
| QRY_02a_Decoration_Flags | Banderes per motiu derivades de T_DECORATIONS (format ample per a khi-quadrat). | H01, H04 |
| QRY_02_Decoration_by_Site | Reconstruïda v11 sobre T_DECORATIONS: recomptes per motiu i jaciment. | H01, H04 |
| QRY_03 – QRY_12 | Sense canvis de funció: conservació, C14, exportació R, volumetria, exportació QGIS, fills, membres de grup, cobertura ChaXR, maçoneria, geologia-construcció. QRY_05 i QRY_07 actualitzades v11 als noms nous i als camps de sistema. | diverses |
| QRY_13_AX_Pattern_Export | Reconstruïda v11: **20 columnes AX_ + 5 columnes SYS_** amb semàntica NA (un component gatejat per un sistema Not applicable s'exporta NA, no 0, perquè el seu 0 és de farciment). La font única FillElementMap garanteix que exportació i validació no divergeixen mai. **v16: X exporta NA quan `Chamber_Roof_Type = Natural bedrock`** i la columna `Interbody_Cornice_Material` desapareix. **Advertència: amb molts sostres de penya natural al corpus, el filtre canvia els resultats i qualsevol prova de Jaccard, clúster o correspondències ja executada s'ha de refer.** | H01, H04, H05 |
| QRY_14_Connections_Edges | Llista d'arestes amb Chrono_Relation: graf potencialment dirigit per a igraph/QGIS. | OE3, H03 |
| QRY_15_Observability_Bias | Recompte per jaciment i sector de base documental i observabilitat. | OE1 |
| QRY_16s_* (5) | Auxiliars de la bateria: Element_Values (format llarg dels 20 elements: la clau que evita ~80 branques UNION), Lost_Cover, Damage_Count, Observational_Nulls (llista de treball camp a camp), Null_Count. | — |
| QRY_16a / QRY_16b / QRY_16c / **QRY_16d** | Regles 1–11, 12–21, 22–31 i **32–39**. Emmagatzemades en quatre parts perquè una UNION única de totes les branques supera el límit «query too complex» de JET. **`QRY_16d` porta nou branques per a huit números**: la regla 38 es desdobla, perquè el substrat que ha de portar una fila ROC depén de quina posició ROC és | — |
| QRY_16_Validation_Check | Unió de les quatre parts: **39 regles**. Resultat buit = corpus coherent. **Correcció de bug v16**: `QRY_16c` es creava amb el nom `..._22_30` i es referenciava com a `..._22_31`, de manera que aquesta unió **no arribava a existir** — i `MkQuery` degrada eixe fracàs a un `Debug.Print` que ningú no llegia. Les tres parcials existien i s'obrien per separat, que és per què va passar desapercebut | — |
| **QRY_19_V16_Review** | **NOVA v16.** Llista de treball de la revisió manual que la transferència v15→v16 exigeix: elements de portal a rejudicar, cobertes de cambra pendents del test de contacte, format i treball de la pedra per entrar, i files de decoració en posició `JAM` o `OVL` que podrien ser ara `PRT`, `SIL` o `LIN` | — |
| **QRY_17_RockArt_All** | **NOVA v13.** Totes les pintures del corpus amb una columna `Context` (*Associated* / *Isolated*). No necessita cap UNION: les associades i les aïllades són **totes dues files de `T_DECORATIONS`** —les primeres penjades d'una estructura, les segones d'un registre PR, que és també un registre de `T_STRUCTURES`—, de manera que és una sola consulta amb JOIN a `L_TYPOLOGY` per a saber quina mena de pare té cada fila. | OE1, H06 |
| **QRY_18_Notes_Review** | **NOVA v14.** Tots els camps de notes de tots els registres, costat per costat, filtrant els buits. Serveix per al repàs i, sobretot, per a **detectar patrons**: quan la mateixa observació apareix repetidament en text lliure, això és el senyal que hauria de ser un camp. | — |

## **6.1. La bateria de 39 regles**

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
| — | — | **32** | **`Stone_Format_Secondary` amb format dominant buit o `ND` (v16)** |
| — | — | **33** | **`Stone_Format_Secondary` igual al primari (v16)** |
| — | — | **34** | **Motiu en relleu al cos basal amb `Decorative_Socle` = 0 (v16, unidireccional)** |
| — | — | **35** | **Decoració en posició `LIN` o `SIL` amb l'element corresponent a 0 (v16)** |
| — | — | **36** | **`Access_Plane` = Facade amb les dues orientacions divergents (v16, avís)** |
| — | — | **37** | **Fila ROC amb `Body_No` emplenat (v16)** |
| — | — | **38** | **Substrat incompatible amb la posició ROC (v16, dues branques: no-`Bedrock` a ROC/PER/PAN, no-`Mixed` a `ROC-OVL`)** |
| — | — | **39** | **Cornisa intercòs amb menys de dos cossos (v16, avís): revisar si és filada en voladís (G)** |

*L'exempció de Timber_Brackets a la regla 1 mereix nota: E és l'únic element del vocabulari amb existència independent del seu sistema — una mènsula aïllada no ha d'haver portat cap plataforma. Per això existeixen Timber_Bracket_Role i la tipologia MEN. La taula 4.5 llista E sota Sys_Platform per al GATING, que és una pregunta distinta de si E implica H. **En v13 el gating del formulari recull finalment aquesta exempció** (secció 9.2).*

*Les regles **22 i 23 queden restringides a files no-ROC** en v13, i les 28–29 en són la contrapartida per a les files ROC. Sense la restricció dispararien creuadament: una estructura amb pintura rupestre però sense decoració arquitectònica faria saltar la 23 contra `Dec_Present`.*

*La **regla 27** deriva del criteri que `N_Basal_Bodies` compta masses de maçoneria i no plataformes (secció 8bis). Sobre el corpus v12 va disparar en 3 de 31 registres amb cossos basals: taxa prou baixa per a ser senyal i no soroll, però convé comprovar si el registre no està simplement a mig entrar abans de corregir-lo.*

*La **regla 31** és barata i captura precisament l'error que la conversió a metres introdueix: entrar 62 pensant en centímetres donaria una repisa de 62 m, i cap altra regla no ho detectaria.*

*Les **regles 32–33** repeteixen sobre la parella de format de pedra les dues úniques maneres com una parella ordenada pot estar mal formada, ja vigilades sobre `ID_Support`: un secundari orfe i una parella que diu dues vegades el mateix.*

*La **regla 38 dispara només amb un desajust registrat, no amb NULL**: una fila ROC sense substrat és feina pendent, no un error, i eixe cas ja el llista la regla 17. El que la regla detecta és una fila que **havia d'anar a l'altra meitat** de la taula sota el test 1 del criteri d'associació.*

*La **regla 39 és el test T2 de la discriminació G/I automatitzat** (8bis.7), i captura precisament el cas que més preocupa en camp: **la plataforma arrasada codificada com a cornisa**. Una cornisa intercòs sense segon cos és una contradicció de termes.*

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

Marcar un 3 en qualsevol element mostra el recordatori de la regla 2 (evidència a 12.Extra). Amb estat Collapsed apareix l'avís en roig de la regla 6 a 11.Sist. El substrat de pigment amaga l'opció Plaster quan Plaster_Present = 0 (prevenció de R13). Als subformularis —on Enabled afectaria la columna sencera de totes les files— la lògica per fila és validació BeforeUpdate: F_CONNECTIONS bloqueja una cronologia direccional sobre una connexió sense adossament (R20) i autoassigna Contemporary a la junta travada; F_LOST_ELEMENTS exigeix el codi d'element quan l'abast és Element (R21).

## **9.5. Prevenció i detecció**

El principi de disseny: **el formulari preveu en el moment d'entrada; QRY_16 detecta a escala de corpus.** La bateria continua sent imprescindible perquè el gating es pot esquivar (taules obertes a mà, importacions, Centre de confiança desactivat), i perquè algunes regles són intrínsecament de corpus (5, 17). Executeu QRY_16_Validation_Check periòdicament durant l'entrada de dades.

## **9.6. Requisit tècnic**

La injecció dels mòduls VBA del gating i de les validacions exigeix «Confiar en l'accés al model d'objectes de projectes VBA» al Centre de confiança d'Access. Si està desactivat, el formulari es construeix igualment i tot queda editable; només falta l'automatisme, i l'script ho avisa a la finestra immediata en lloc de fallar en silenci.

# **10. Scripts i rutes d'actualització**

| **Fitxer** | **Sub públic** | **Ús** |
| --- | --- | --- |
| chachapoya_DB_v16.bas | BuildDB() | Construcció completa sobre una BD EN BLANC: 22 taules, valors per defecte (0 als 20 elements, NULL a la resta; sistemes a Absent), lookups, 25 relacions, **30 consultes**. Informa del **recompte de consultes creades contra les previstes** |
| chachapoya_Form_v16_val.bas | BuildForm() | Després de BuildDB(): 6 subformularis, F_STRUCTURES amb 12 pestanyes, combos de dues columnes, captions DAO, gating i validacions de subformulari |
| **chachapoya_EXPORT_v15.bas** | ExportAll() | **S'executa sobre la BD v15 d'origen.** Escriu 8 CSV inspeccionables i editables. **Cap identificador travessa la frontera**: totes les claus alienes ixen com el valor llegible que apunten. Els camps el criteri dels quals ha canviat ixen en blanc i es conserven en columnes `OLD_` |
| **chachapoya_IMPORT_v16.bas** | ImportAll() | **S'executa sobre una v16 acabada de construir i buida.** Es nega a arrancar si `T_STRUCTURES` ja té files. Remapa `SOC` → `BAS` per nom, resol totes les claus **abans** de `rs.AddNew`, i informa per fila del que no ha pogut escriure |
| diagnostic_v13.bas | RunDiagnostic() | **Només lectura.** Comprovacions de coherència i de variància sobre el corpus; no modifica cap dada ni cap esquema. Es pot executar tantes vegades com calga |

*Els scripts de migració de versions anteriors queden fora del paquet: la base es construeix sempre de zero i les dades hi arriben per la parella export/import, amb un fitxer inspeccionable entremig.*

**El CSV és l'espai de treball de la revisió, no només un pas intermedi.** `Sill`, `Jambs` i `Lintel` ixen en blanc i el valor v15 va a `OLD_Sill`, `OLD_Jambs`, `OLD_Lintel`; `Chamber_Roof` i el seu tipus ixen en blanc **només** on el tipus era penya natural o mixt, perquè un sostre d'obra es va construir igual sota qualsevol criteri. El que es copie de tornada a la columna viva és un judici; el que quede en blanc queda honestament buit i `QRY_19_V16_Review` el llistarà.

**Ordre d'execució:** `ExportAll()` sobre la còpia local real → revisió dels CSV → `BuildDB()` i `BuildForm()` sobre una base en blanc → `ImportAll()` → finestra d'immediat → `QRY_16_Validation_Check` → `QRY_19_V16_Review` → **reexecució de la matriu A–X i comparació amb els resultats anteriors**, que el filtre NA sobre els sostres naturals modifica.

Les restriccions tècniques de VBA/JET es mantenen: cap continuació de línia, ASCII estricte en codi (cp1252), patró sql = sql & "…", paraules reservades evitades, N−1 parèntesis per a N taules en JOIN, Private per a tot auxiliar, DROP previ per a objectes regenerables, combos configurats per ControlSource.
