**DELTA v12 → v13**

*Base de dades arqueològica — La Petaca i Diablo Wasi (PALP)*

Esteve Ribera Torró | TFM Arqueologia UA

**Document viu.** S'hi anoten les decisions d'actualització a mesura que es prenen, amb la justificació que les sosté i les passes d'implementació que impliquen. Quan el conjunt estiga tancat, aquest document és l'especificació des de la qual es generen els scripts v13 i s'actualitzen `esquema_bbdd_estructures` i `tfm_metodologia_bbdd`.

**Estat:** IMPLEMENTAT (revisió 7, agost 2026). Paquet v14.

**Estat anterior:** IMPLEMENTAT (revisió 5). La revisió 5 (secció 12) corregeix decisions de la revisió 3 a partir de contraexemples de camp.

**Estat anterior:** IMPLEMENTAT (revisió 4). El paquet v13 recull totes les decisions marcades `[DECIDIT]`. Els punts `[OBERT]` passen al futur delta v13→v14.

**CANVI D'ESTRATÈGIA (rev. 4): els 35 registres es reintrodueixen a mà sobre una BD construïda de zero.** Això elimina tot l'aparat de migració (regla B, UPDATE+INSERT segur per a no trencar FK, reassignació dels sis registres de D1, normalització del valor solt `Arch. elements`, migració de `Both`) i, sobretot, elimina els 9 heretats del defecte v10 que el diagnòstic va destapar. Sota v13, cada valor emmagatzemat és un judici deliberat — la precondició per a publicar qualsevol percentatge.

**Convenció d'estat de cada punt:**
`[DECIDIT]` acordat, pendent d'implementar · `[OBERT]` en discussió · `[PENDENT DADES]` depén d'una comprovació sobre el corpus · `[FET]` implementat i verificat

---

# **0. Resum**

**Revisió 1** — suport geològic, gating d'E, recompte de cossos.

| # | Punt | Estat | Abast |
| --- | --- | --- | --- |
| 1 | Correcció del gating de l'element E | `[DECIDIT]` | Formulari (bug) |
| 2 | Desdoblament de `Fissure/Crack` a `L_SUPPORT` | `[PENDENT DADES]` | Lookup + documentació |
| 3 | Criteri de `ID_Support` / `ID_Support_Secondary` | `[DECIDIT]` | Documentació |
| 4 | H no compta com a cos basal | `[DECIDIT]` | Documentació + regla nova |
| 5 | Hipòtesis funcionals de mènsules (politja) | `[DECIDIT]` | Documentació (sense canvi d'esquema) |

**Revisió 2** — art rupestre associat i registre de la decoració (secció 7).

| # | Punt | Estat | Abast |
| --- | --- | --- | --- |
| 7.1 | Art rupestre associat integrat a `T_DECORATIONS` | `[DECIDIT]` | Cap canvi de taula |
| 7.2 | `L_STRUCT_BODY`: posicions `ROC` + `Sort_Order` | `[PENDENT DADES]` | Lookup |
| 7.3 | Criteri d'associació (quatre tests) | `[DECIDIT]` | Documentació |
| 7.4 | Dues subseccions a 4.Dec sobre una sola taula | `[DECIDIT]` | Formulari |
| 7.5 | `RockArt_Present` germà de `Dec_Present` | `[DECIDIT]` | Esquema + formulari + regles |
| 7.6 | `Color_Secondary`, retirada de `Both` | `[DECIDIT]` | Esquema + formulari |
| 7.7 | `QRY_17_RockArt_All` amb columna de context | `[DECIDIT]` | Consulta nova |
| 7.8 | `Pigment_Substrate = Bedrock` es manté | `[DECIDIT]` | Documentació |
| 7.9 | Relació cromàtica entre colors contigus | `[OBERT]` | Ajornat: recollir a `Notes` |

**Revisió 3** — semàntica del zero, valors per defecte i abast del domini de cinc valors (secció 8).

| # | Punt | Estat | Abast |
| --- | --- | --- | --- |
| 8.1 | Semàntica de l'«absent» documentada per domini | `[DECIDIT]` | Documentació |
| 8.2 | Valor per defecte dels camps sense gating | `[DECIDIT]` | Esquema (DAO) |
| 8.3 | Promoció a cinc valors de les capes aplicades | `[DECIDIT]` | Esquema + formulari + consultes |
| 8.4 | Criteri general d'assignació de domini | `[DECIDIT]` | Documentació |
| 8.5 | Solapament `Plaster_Present` = 2 amb `Traces only` | `[DECIDIT]` | Documentació o lookup |
| 8.6 | Variància real de `Chinking_Stones` | `[PENDENT DADES]` | Possible retirada |

---

# **1. Correcció del gating de l'element E** `[DECIDIT]`

## Problema

El gating v12 tracta `Timber_Brackets` (E) com un component més de `Sys_Platform`: amb el sistema a *Absent*, *Not applicable* o *Not observable*, el camp es desactiva i l'emplenat ràpid ofereix posar-lo a 0.

Això és **actiu­ment fals** en un cas real i freqüent: una cambra funerària (EA-CAM) amb una mènsula lateral que no sosté cap plataforma. El sistema plataforma és genuïnament absent i la mènsula existeix. El formulari, tal com està, obliga a negar-la o a entrar-la en un ordre concret per a esquivar el bloqueig.

## Justificació

E és **l'únic element del vocabulari amb existència independent del seu sistema**. La bateria de validació ja ho reconeix: la regla 1 exempta explícitament `Timber_Brackets`, i l'exempció està raonada al codi — una mènsula aïllada no ha d'haver portat mai una plataforma, i és exactament la raó de ser de `Timber_Bracket_Role` i de la tipologia MEN. La taula 4.5 llista E sota `Sys_Platform` **per al gating**, que és una pregunta distinta de si E implica H.

El gating v12 va aplicar la pertinença de la taula 4.5 sense recollir l'exempció que la validació ja feia: una incoherència interna del paquet, no un criteri nou.

## Decisió

1. `Timber_Brackets` queda **sempre editable** dins del grup de la plataforma, independentment del valor de `Sys_Platform`.
2. `Timber_Bracket_Count` i `Timber_Bracket_Role` passen a dependre **de E** (nivell 3), no del sistema: editables si E val 1, 2 o 3; bloquejats si val 0 o 9.
3. `Platform_Surface_Material` i `Platform_Function` continuen depenent del sistema (són atributs de H, no de E), i el material continua bloquejat amb `Role = Isolated` (regla 7).
4. L'emplenat ràpid de `QuickFillSys` per a `Sys_Platform` **exclou E** de la llista de components que posa a 0 o 9. F i G s'hi mantenen: són molt més difícils de llegir com a cosa distinta de components de plataforma.

## Implementació

- `BuildGatingV12` (present en dos fitxers: `chachapoya_Form_v12_val.bas` i `upgrade_form_v12.bas`): moure `EnSrc "Timber_Brackets"` fora del bloc governat per `SysOpen(Me!Sys_Platform)`; condicionar `Timber_Bracket_Count` i `Timber_Bracket_Role` a `ElemHas("Timber_Brackets")`.
- `QuickFillSys`, cas `Sys_Platform`: llevar `Timber_Brackets` de la cadena `comps`.
- Cap canvi d'esquema, cap dada afectada.

## Registre de referència (per a la documentació)

Cambra funerària amb mènsula lateral sense plataforma:

`Sys_Platform` = Absent (verificat) o Not observable · `Timber_Brackets` = 1 · `Timber_Bracket_Count` = 1 · `Timber_Bracket_Role` = Isolated (si s'ha comprovat que no sosté ni sostenia res) o ND (si no és resoluble) · `Transverse_Beams`, `Corbelled_Courses` = 0 o 9, coherents amb el sistema.

*Nota sobre el rol:* `Isolated` és una **absència verificada** (s'ha mirat i no hi ha plataforma) i `ND` és la manca d'observació — la mateixa distinció 0/9 aplicada al camp de rol. Marcar `ND` per prudència quan s'ha verificat fa perdre el cas de tot denominador i el deixa a la llista de treball de la regla 9.

---

# **2. Desdoblament de `Fissure/Crack` a `L_SUPPORT`** `[PENDENT DADES]`

## Problema

El valor `Fissure/Crack` fon dos fenòmens geològics distints que ofereixen lògiques estructurals oposades:

- **Diàclasi (escletxa vertical):** fractura que travessa els bancs. Aporta dos paraments laterals, constreny la planta i s'omple de material per a generar el nivell basal. Actua per **confinament**: la roca fa d'encofrat.
- **Junta d'estratificació (rebaix horitzontal):** erosió diferencial dels estrats prims de llims o margues entre bancs calcaris competents. Aporta una ranura contínua on empotrar jàcenes (A) o mènsules (E). Actua per **encastament**: treball a tallant i moment, no a compressió.

Que compartesquen casella és un accident del vocabulari (en català i en anglés, «fissura» serveix per a totes dues), no una afinitat real. Un conté, l'altre ancora.

## Decisió provisional

Substituir el valor únic per dos:

| Valor emmagatzemat | Valencià | Mode mecànic |
| --- | --- | --- |
| `Vertical cleft` | Diàclasi / escletxa vertical | Confinament |
| `Bedding-plane recess` | Junta d'estratificació | Encastament |

*Sobre els noms:* s'evita `joint` en anglés perquè `Connection_Type` ja el gasta per a juntes de fàbrica, i la col·lisió seria confusa en l'exportació.

## Camp `Support_Mode`: descartat

Es va valorar afegir un camp `Support_Mode` (*Gravitational rest / Embedment / Confinement / Mixed / ND*) i **s'ha descartat**: el mode mecànic és deduïble del nom de la forma (una junta d'estratificació implica encastament; una micro-repisa implica repòs), de manera que seria un camp derivable — no informació, sinó una oportunitat de contradicció. Amb 35 registres seria a més quasi constant per valor de forma.

**Contrapartida obligatòria:** la deducció ha de quedar **escrita, no intuïda**. La columna `Description` de `L_SUPPORT` ha de recollir el mode mecànic de cada forma, cosa que fixa el criteri per a l'observador i deixa la derivació disponible per a R quan es vulga com a variable agregada.

Modes per forma, per a eixa columna:

- **Repòs gravitacional:** repisa ampla, repisa estreta, micro-repisa, sòl de cavitat, repisa artificial.
- **Encastament:** junta d'estratificació.
- **Confinament:** diàclasi.
- **Mixt:** no és un valor; es registra amb la parella `ID_Support` + `ID_Support_Secondary`.

## Comprovació prèvia `[PENDENT DADES]`

Abans d'implementar, revisar els registres amb `ID_Support = Fissure/Crack` i determinar quants són verticals i quants horitzontals.

- Si **tots** són d'una mena, el canvi es redueix a renomenar la fila existent i afegir l'altra al final: cap reassignació.
- Si n'hi ha de les dues, cal reassignar-los un a un abans de tancar el punt.

**Restricció tècnica:** no esborrar mai la fila existent si hi ha registres que hi apunten (trencaria les FK de `REL_SUP_STR` i `REL_SUP2_STR`). El patró correcte és el ja usat amb el desdoblament d'`ND` a `L_TYPOLOGY` en la v11: `UPDATE` en el lloc de la fila existent + `INSERT` de la nova al final. Mateix ID, `ID_Support` intacte.

## Comprovacions analítiques derivades `[PENDENT DADES]`

Dues consultes que val la pena fer un cop desdoblat, perquè poden ser material de resultats:

1. **Quantes vegades la diàclasi apareix com a primària i quantes com a secundària.** Si resulta que quasi sempre és secundària, això diu alguna cosa sobre com s'usa (mai com a suport principal, sempre com a estabilitzador).
2. **Amb quina altra forma s'aparella cada vegada** (diàclasi + cavitat, micro-repisa + diàclasi...). Aquestes parelles són, de fet, **tipus de solució constructiva**, i podrien ser la classificació que ix de H02.

## Valor per a H02

És l'argument de H02 més fort disponible: el **ritme estratigràfic** del faralló —alternança de bancs competents i estrats erosionables— genera *simultàniament* les repises on es construeix i les ranures on s'ancora. Una sola propietat geològica que condiciona dues decisions constructives distintes és molt més demostratiu que constatar una preferència per repises amples. Si en algun moment es registra la litologia o la potència dels bancs, es tanca amb la variable independent.

*Solapament a tindre present:* micro-repisa i junta d'estratificació són **el mateix fenomen vist des de cares oposades**. Quan un estrat tou s'erosiona, el banc inferior queda volat (micro-repisa: s'hi reposa) i alhora queda una ranura (encaix: s'hi encasta). Quina de les dues coses és depén exclusivament del que va fer el constructor, no de la roca. Sense el criteri escrit, dos observadors codificaran la mateixa paret de manera distinta.

---

# **3. Criteri de `ID_Support` i `ID_Support_Secondary`** `[DECIDIT]`

## Problema

El criteri operatiu no estava fixat. La interpretació implícita («suport de la part basal») falla en dos casos freqüents: les estructures sense cos basal (EA-PLA-V, cambres allotjades directament en cavitat) i les que reben la càrrega per confinament lateral en lloc de per compressió.

## Criteri adoptat

> **`ID_Support`** = la forma del relleu que **rep la càrrega** de l'estructura, siga per repòs (compressió sobre un pla) o per **confinament** (transmissió lateral a les parets).
> **Prova operativa:** *què cauria si eixa forma no hi fora?*
>
> **`ID_Support_Secondary`** = la forma que **estabilitza o allotja** sense rebre la càrrega principal: defineix l'espai, protegeix o contribueix a la fixació.

**Transparència dels elements construïts:** per a identificar el primari cal seguir la cadena de càrrega cap avall a través de la maçoneria fins a trobar roca. La cadena acaba **on la maçoneria toca roca**, que pot ser per sota (repòs) o lateralment (confinament).

*Per què la prova contrafactual i no «per on entra la càrrega»:* la primera es resol mirant; la segona obliga a decidir prèviament quina és la cadena, que és justament on és fàcil equivocar-se. Cas real d'error durant la discussió: tractar un basament dins d'una diàclasi com a «transparent» i seguir la càrrega fins al sòl de la cavitat, quan la maçoneria transmet la càrrega **a les parets de l'escletxa** i no al sòl.

## El que NO és criteri

S'ha valorat i **descartat** la regla «primari = N0, secundari = N1». Els casos reals la contradiuen en les dues direccions, i hi ha una raó estructural: quan la cambra s'allotja directament a la roca, el suport primari està a cota de N1 per definició. Confondre els dos eixos reintroduiria per la porta de darrere la barreja nivell/cos que la v9 va separar.

**Nota orientativa admissible** (no norma): *el suport primari sol situar-se a cota inferior al secundari —conseqüència de la gravetat—, però no necessàriament al nivell N0.*

## Casos de referència documentats

Tots dos de Diablo Wasi, llegits sobre model 3D i fotografia:

| Cas | Primari (rep càrrega) | Secundari (estabilitza / allotja) |
| --- | --- | --- |
| Basament construït dins d'una diàclasi + cambra en cavitat preexistent | **Diàclasi** (confina el basament) | **Cavitat** (allotja la cambra) |
| Torre sobre micro-repisa + cambra empotrada entre dos cossos de pedra | **Micro-repisa** (repòs) | **Diàclasi** (empotra la cambra) |

**El que aquests dos casos il·lustren, i que convé traslladar a la memòria:** el repertori geològic és el mateix i el que varia és **quina forma rep i quina estabilitza**. La diàclasi apareix en tots dos casos amb funció distinta — és la forma més versàtil del repertori.

## Ambigüitat cavitat / diàclasi eixamplada `[DECIDIT]`

En aquesta geologia, l'allotjament d'una cambra pot llegir-se com a cavitat o com a escletxa en la seua màxima amplitud. Criteri:

> **Si l'allotjament té sostre rocós propi i profunditat cap a dins → cavitat. Si és l'eixamplament d'un pla de fractura que continua amunt i avall → diàclasi.**
> Indici corroborador: `Chamber_Roof_Type = Natural bedrock`.
> Si al model no es resol, registrar les dues amb la parella ordenada i explicar-ho a `Notes`.

## Confinament fabricat `[OBERT]`

Observació de camp pendent de resoldre: en un dels casos, la cambra queda empotrada **entre dos cossos de maçoneria**, no entre dos paraments rocosos. Això és una estructura que *fabrica el seu propi confinament* — la solució de l'escletxa reproduïda constructivament allà on la roca no la donava.

No és un valor de `L_SUPPORT` (no és geologia). Per ara es registra a `T_ARCH_FEATURES` o `Notes`, amb `N_Built_Walls` reflectint-ho. **A decidir:** si apareix en diversos registres, mereix camp propi o entrada de vocabulari? És observació de primer ordre per a H01 i H04.

## Perspectiva analítica (sense canvi d'esquema)

La relació entre **forma de suport i nivell constructiu** és una variable analitzable que cap camp captura directament, però que es deriva creuant `ID_Support`, `ID_Support_Secondary`, `N_Basal_Bodies` i `Chamber_Roof_Type` en R. Amb 35 registres no dona per a res; amb el corpus de La Petaca sencer podria ser una de les taules de resultats de H02.

---

# **4. La plataforma (H) no compta com a cos basal** `[DECIDIT]`

## Problema

Cas real: un mausoleu que s'erigeix sobre una plataforma de mènsules encastades a una junta d'estratificació. Compta la plataforma com a cos basal?

## Decisió

**No.** `N_Basal_Bodies` compta **masses de maçoneria** (elements de tipus B), no plataformes.

## Justificació

El criteri de nivell ja establit diu que un **cos** és N1 si té o tenia obertura d'accés, i N0 si no en té. Però H no és un cos: és un **sistema**, un element d'interfície que tanca N0 i precondiciona N1. No és una massa de maçoneria superposada.

Si H comptara com a cos basal, un mausoleu sobre plataforma i un mausoleu sobre pòdium de maçoneria donarien tots dos `2, 1` — i són precisament dues solucions constructives que cal poder distingir.

## Registre de referència

Mausoleu sobre plataforma de mènsules: `N_Basal_Bodies` = 0 · `N_Chamber_Bodies` = 1 · `Sys_Platform` = Present · `Timber_Bracket_Role` = Platform support · `ID_Support` = junta d'estratificació.

La plataforma queda registrada al vocabulari A–X, que és on li correspon, i no infla el recompte de cossos.

## Regla de validació nova (proposta) `[OBERT]`

**R27:** `N_Basal_Bodies` > 0 amb `Base_Level` (B) = 0 → un cos basal és una massa de maçoneria; si no hi ha B, revisar si el que s'ha comptat és una plataforma.

*A valorar:* pot generar falsos positius en basaments que no s'han registrat com a B diferenciat. Comprovar contra el corpus abans d'incorporar-la.

---

# **5. Hipòtesis funcionals de mènsules** `[DECIDIT]`

## Problema

Sospita fundada que una mènsula aïllada tinguera funció de **politja** (hissat vertical) i no de suport de plataforma. On es registra?

## Decisió

**No es toca `Timber_Bracket_Role`.** El seu domini (*Platform support / Isolated / Both / ND*) respon una única pregunta comprovable al model 3D: *aquesta mènsula forma part del sistema H d'aquesta estructura?* Una politja no ho és, de manera que el valor continua sent `Isolated`.

Afegir un valor «Pulley» barrejaria en un mateix camp una observació verificable amb una inferència, i a partir d'ací seria impossible saber si un `Isolated` d'un altre registre significa «he comprovat que no sosté res» o «no se m'ha acudit cap hipòtesi». És el mateix error que la v10 va corregir eliminant `Wooden_Stakes`: **un rol distint no ha de generar un element distint, ni tampoc un valor de domini distint quan el que canvia és la interpretació.**

La hipòtesi funcional va a `T_ARCH_FEATURES`: `Feature_Code` = *Other (see Notes)*, `Material` = Timber, i a `Notes` la hipòtesi amb els indicis. Queda recuperable i consultable sense contaminar cap denominador ni cap vector A–X.

## Criteris diagnòstics a consignar a les notes

Perquè la hipòtesi siga contrastable quan hi haja més casos:

- Solc, desgast o polit a la cara superior o al vol (una corda deixa rastre).
- Posició respecte de l'obertura d'accés: una politja de hissat estaria sobre l'eix vertical de l'entrada o desplaçada per damunt, no en un lateral qualsevol.
- Alçada relativa i orientació del vol.
- Secció més robusta del que demanaria un suport passiu.
- Si n'hi ha una de sola o formen parella.

## Implicació per a OE3 `[OBERT]`

Si la hipòtesi s'aguanta en diversos casos, implicaria una infraestructura de circulació **vertical** (hissat de fardells o de persones) al costat de la circulació aèria **horitzontal** que reconstrueixen mènsules i plataformes. Són dues coses distintes i el registre actual no les separa. **A decidir amb més casos:** camp propi, valor de `Platform_Function` reconvertit, o tipus nou a `L_GROUP_TYPE`.

---

# **6. Nota metodològica: límits de la lectura fotogràfica** `[DECIDIT]`

Durant aquesta discussió s'han produït errors de lectura reals en interpretar fotografies frontals: **un pla que avança i un pla que retrocedeix es confonen en dues dimensions.** Concretament, la vora d'una escletxa es va llegir com un banc que sobreïx, i un ressalt inferior com una escletxa que flanqueja.

Això no és anecdòtic: és exactament la classe d'error que els models 3D resolen i les fotografies no, i afecta directament la codificació de `ID_Support`. Reforça la funció de `Doc_Basis`, `Facade_Observability` i `Interior_Observability`, i justifica que les assignacions de suport fetes només sobre fotografia queden marcades com a tals.

**A valorar** `[OBERT]`: els criteris de la secció 3 (prova contrafactual, cavitat vs diàclasi) exigeixen **visió tridimensional o accés directe**. Convé una convenció que impedesca assignar suport primari a partir d'una fotografia frontal, o que almenys ho registre. La via de menys fricció seria una regla de validació que avise quan `ID_Support` estiga assignat amb `Doc_Basis = Ground photography` o `Distant photogrammetry`.

---

# **7. Art rupestre associat i registre de la decoració** `[REVISIÓ 2]`

## 7.0. Plantejament

Els jaciments presenten pigment aplicat sobre la penya en situacions molt diverses: motius sobre roca al costat d'estructures, taques amorfes i bandes que emmarquen obertures o llindars, camps de color amb vora d'un altre to, i panells sense cap estructura al voltant. La v12 no els podia registrar bé: els booleans `RA_*` es van eliminar en la v11 (correctament, duplicaven `T_DECORATIONS`) però el criteri que els substituïa apuntava a camps que ja no existien.

**Constatació de partida:** el repertori iconogràfic rupestre **no és compartit** amb el de la decoració arquitectònica. Són dos repertoris distints sobre dos suports distints, i el disseny ho ha de reflectir sense duplicar maquinària.

## 7.1. Integració a `T_DECORATIONS`, no taula nova `[DECIDIT]`

L'art rupestre **associat a una estructura** es registra com a files de `T_DECORATIONS`, la mateixa taula que la decoració arquitectònica.

**Per què no una taula pròpia:** duplicaria l'estructura (cada camp futur s'hauria d'afegir dues vegades i divergirien), duplicaria les relacions i les regles de validació, i obligaria a una UNION cada vegada que es vulga el conjunt. La taula ja té el que cal: `ID_Structure` per al vincle, `Substrate` amb valor `Bedrock`, `Color`, `ID_Dec_Type` i `Notes`.

**Per què no fitxa PR per a tot** (alternativa valorada i descartada): trencaria la sèrie en dos orígens amb numeradors i denominadors distints, i faria difusa la frontera en els casos ambigus. La fitxa PR es reserva per a l'art rupestre **no associable** (secció 7.3).

**L'únic obstacle real era la posició**, i és reparable: `ID_Struct_Body` apunta a `L_STRUCT_BODY`, que només conté posicions arquitectòniques. Una pintura sobre la penya acabaria en `ND` — que significa «posició no determinada», no «posició no arquitectònica»: dues coses molt distintes col·lapsades en un valor, el mateix patró d'error que el domini 0/1/9 existeix per a evitar. Ho resol el punt següent.

## 7.2. `L_STRUCT_BODY`: posicions `ROC` i ordre explícit `[PENDENT DADES]`

Entrades noves amb **`Level_Type = ROC`**. Aquesta columna és el **discriminador analític** entre els dos conjunts: separar decoració arquitectònica de pintura rupestre passa a ser `WHERE Level_Type = 'ROC'` o el seu contrari, sense cap camp nou a `T_DECORATIONS`.

Entrades previstes (a confirmar contra els casos reals):

| Code | Name | Level_Type | Ús |
| --- | --- | --- | --- |
| ROC | Adjacent bedrock | ROC | Penya adjacent a l'estructura |
| ROC-PER | Opening perimeter (bedrock) | ROC | Marques que emmarquen una obertura o llindar natural |
| ROC-PAN | Panel (no architectural reference) | ROC | Per a les files penjades d'un registre PR aïllat (secció 7.7) |

**Sense l'entrada ROC-PAN, les pintures aïllades tornarien a caure en `ND`**, que és exactament el problema que es vol evitar.

**Correcció col·lateral detectada:** `L_STRUCT_BODY` s'ordena per `Level_Type, Name`, de manera que `Eave (U)` apareix al final de la llista (Level_Type = SUP) i dins de cada nivell l'ordre és alfabètic, no constructiu. Afegir una columna **`Sort_Order`** i ordenar-hi les llistes desplegables, per a recuperar la lògica bottom-to-top del vocabulari.

`L_DEC_TYPE` creix amb els tipus del repertori rupestre. `[PENDENT DADES]`: el conjunt concret depén dels casos documentats.

## 7.3. Criteri d'associació: quatre tests en ordre de prioritat `[DECIDIT]`

Decideix si una pintura és fila de `T_DECORATIONS` d'una estructura o registre PR propi. **El primer test que es compleix decideix.**

> 1. **Continuïtat física amb la fàbrica** — el pigment toca la maçoneria, la ressegueix o continua sobre ella. → **És pigment de l'estructura** (`Pigment_*` a 3.Acab; detall per posició a `T_DECORATIONS`).
> 2. **Comparteix suport geològic** — es troba dins de la mateixa escletxa, repisa, cavitat o rebaix que allotja l'estructura (el que ja registra `ID_Support`). → **Associada** (fila ROC).
> 3. **Emmarca un element de l'estructura** — obertura, llindar, perímetre, encara que hi haja discontinuïtat. → **Associada** (fila ROC).
> 4. **Cap de les anteriors** → **fitxa PR pròpia**, vinculada per `T_GROUPS` (*Rock art cluster*) si hi ha proximitat visual.

**Per què no un llindar mètric** (es va valorar 1 m i es descarta): la distància és de segona passada i l'associació s'ha de decidir en entrar el registre, de manera que un criteri mètric acabaria aplicant-se «a ull». A més, en una paret vertical 1 m no significa el mateix a tot arreu: una marca a 1,20 m però dins de la mateixa escletxa està relacionada; una a 60 cm separada per un banc que trenca la continuïtat, no necessàriament. La distància euclidiana ignora el que estructura aquests jaciments.

**Virtut del test 2:** reutilitza una decisió ja presa en entrar el registre (`ID_Support`), de manera que no cal mesurar res i és replicable per un tercer.

**La distància sí que val com a desempat del test 4**, per a decidir l'agrupació — no la naturalesa del registre.

**Traçabilitat:** anotar a `Notes` quin test s'ha aplicat, perquè un tercer puga verificar la decisió en lloc de refer-la.

**Els casos dubtosos tenen sortida sense forçar:** fitxa PR + agrupació no perd res — el conjunt continua sent consultable i la pintura conserva coordenades, orientació i visibilitat pròpies. És l'opció conservadora quan el criteri no resol.

`[PENDENT DADES]` **Validació del criteri:** aplicar els quatre tests a cinc o sis casos reals ja documentats. Si algun no cau clarament en cap, eixe cas dirà què falta.

## 7.4. Formulari: dues subseccions sobre una sola taula `[DECIDIT]`

La pestanya 4.Dec presenta **dos subformularis**, tots dos sobre `T_DECORATIONS`:

| Subsecció | Filtre | Combos restringits a |
| --- | --- | --- |
| Decoració arquitectònica | `Level_Type <> 'ROC'` | posicions arquitectòniques + tipus arquitectònics |
| Pintura rupestre | `Level_Type = 'ROC'` | posicions ROC + tipus rupestres |

**Una taula, dues vistes.** L'usuari no veu mai una llista barrejada (entrada més ràpida i menys errors), i no es manté res dues vegades. La separació analítica la dona `Level_Type`, no la separació física de les dades.

## 7.5. `RockArt_Present` germà de `Dec_Present` `[DECIDIT]`

`Dec_Present` no pot respondre alhora «aquesta estructura estava decorada?» i «hi ha pintura rupestre associada?». Amb les subseccions separades, la solució coherent amb el que l'usuari veu és **un combo damunt de cada subsecció**, cadascun governant el seu subformulari:

- **`Dec_Present`** BYTE 0/1/9 — decoració **arquitectònica** (files no-ROC).
- **`RockArt_Present`** BYTE 0/1/9 — pintura rupestre **associada** (files ROC). Camp nou, per defecte 0.

Regles de coherència noves, calcades de les 22-23 però restringides per `Level_Type`:

- **R28:** `RockArt_Present` = 1 sense cap fila ROC.
- **R29:** files ROC existents amb `RockArt_Present` ≠ 1.

I les regles 22-23 s'han de **restringir a files no-ROC**, o dispararan creuadament. `[IMPLEMENTACIÓ]` cal afegir el JOIN amb `L_STRUCT_BODY` a les dues.

## 7.6. `Color_Secondary` i retirada de `Both` `[DECIDIT]`

**Problema:** `Color` és una llista tancada i `Both` només cobreix roig+blanc. Els casos bicroms documentats tenen els colors **junts o contigus dins d'un mateix motiu** (per exemple, camp clar amb vora roja), de manera que descompondre en dues files no serveix: inventaria dos motius on n'hi ha un i duplicaria la posició.

**Decisió**, seguint el patró de parella ordenada ja usat a `ID_Support` + `ID_Support_Secondary`:

- `Color` = color **dominant** (el que ocupa més superfície o defineix el motiu)
- `Color_Secondary` = color acompanyant. Buit en els casos monocroms, que són la majoria (roig)

**`Both` es retira**: la parella ordenada el substitueix amb avantatge, perquè diu quin domina. `[IMPLEMENTACIÓ]` migrar els registres existents amb `Both` a la parella segons el model; si no és resoluble, `Color = ND` amb nota.

**Descartat:** camp multivalor d'Access — trenca l'exportació neta a R, que és el sentit de tot l'esquema.

## 7.7. `QRY_17_RockArt_All`: les dues meitats en una consulta `[DECIDIT]`

Les pintures associades i les aïllades queden en llocs distints, però **totes dues acaben sent files de `T_DECORATIONS`**: les associades penjades d'una estructura, les aïllades penjades d'un registre PR (que és també un registre de `T_STRUCTURES`, amb les seues pròpies files de motiu i color).

Per tant la unió **no necessita cap UNION**: és una consulta sobre `T_DECORATIONS` amb JOIN a `L_TYPOLOGY` per a saber si el pare és PR o una estructura, i a `L_STRUCT_BODY` per al filtre ROC. Columna derivada:

**`Context`** = *Isolated* (el pare té tipologia PR) / *Associated* (el pare és una estructura i la fila és ROC).

**El context és una variable analitzable, no una molèstia administrativa:** permet preguntar si els motius, els colors o la posició respecte del llindar difereixen segons si hi ha estructura o no.

**Limitació coneguda i acceptada:** les coordenades de les pintures associades són les de l'estructura, no les de la pintura. Irrellevant a l'escala del faralló (distribució per sectors, densitat, visibilitat); si en algun moment cal posició fina, la solució barata és afegir coordenades opcionals a `T_DECORATIONS`. **No es fa ara.**

## 7.8. `Pigment_Substrate = Bedrock` es manté `[DECIDIT]`

Es va plantejar si aquest valor quedava redundant amb el registre de l'art rupestre. **No hi ha solapament:** `Pigment_Substrate` és un camp de **l'estructura** i registra sobre què s'aplica el pigment que forma part d'ella. El cas paradigmàtic —el fons rocós d'una cambra que aprofita una cavitat— significa que la roca *era* la superfície acabada prevista, sense operació preparatòria de revoc. És la distinció de cadena operativa que justifica el camp, i res més no la registra.

**Criteri de frontera:** pigment sobre la penya que forma part de la superfície acabada de l'estructura → `Pigment_Substrate = Bedrock`. Pigment sobre la penya que no forma part de l'estructura → fila ROC. El test 1 de la secció 7.3 ho decideix.

## 7.9. Avís en lloc de bloqueig: substrat Bedrock + extensió arquitectònica `[DECIDIT]`

Es va plantejar bloquejar les combinacions «substrat = Penya» amb valors de `Pigment_Extent` referits a elements arquitectònics. **No s'ha de bloquejar:** eixa combinació és precisament el **cas diagnòstic de cos perdut** — bandes verticals de pigment sobre la roca alineades amb els brancals del cos inferior, on l'arquitectura hi era i la pintura li ha sobreviscut. Bloquejar-ho impediria registrar l'evidència que sosté un valor 3 amb `Pigment on bedrock` a `T_LOST_ELEMENTS`.

**Igualment, «façana sencera» amb substrat Bedrock és legítim:** una cambra que aprofita una cavitat pot tindre pigment cobrint tot el front.

**R30 (proposta):** amb `Pigment_Substrate = Bedrock` i `Pigment_Extent` = elements arquitectònics, avisar que això pot ser un cos desaparegut i suggerir obrir fila a `T_LOST_ELEMENTS`. **Informa, no impedeix** — com l'avís de col·lapse.

## 7.10. Relació cromàtica entre colors contigus `[OBERT]`

Un camp clar amb vora d'un altre color no és el mateix que dues bandes juxtaposades ni que un motiu entrellaçat. Si el patró és recurrent, és informació de **tècnica** (una vora aplicada després del camp implica dues operacions) que connecta amb la chaîne opératoire.

**No es crea camp encara.** Amb pocs casos bicroms, seria quasi buit i les seues categories s'estarien inventant sense base. **Recollir a `Notes` amb fórmula consistent** (per exemple «camp X amb vora Y») i formalitzar-ho quan les dades donen el vocabulari — el mateix criteri aplicat a la hipòtesi de politja (secció 5).

---

# **7bis. Resultats del diagnòstic previ** `[REVISIÓ 4]`

Executat sobre el corpus v12 (35 registres) amb `diagnostic_v13.bas`. Es reprodueixen només les lectures que van decidir alguna cosa.

## Decisions que es tanquen

| Comprovació | Resultat | Decisió |
| --- | --- | --- |
| **D7** `Chinking_Stones` | 26 presents, **0 absències verificades**, 1 no observable | **Variància zero.** El camp no pot alimentar cap prova. Es manté per decisió de l'investigador com a descriptor tècnic (T&A 2017) i perquè la universalitat és ella mateixa un resultat sobre la tradició constructiva de DW. **No es promociona a cinc valors**: afinar el domini d'un camp sense variància seria treball inútil |
| **D6** `Color = Both` | **0 files** | Retirada gratuïta, sense migració. `Color_Secondary` es crea igualment per al cas bicrom documentat (camp clar amb vora roja), encara no registrat |
| **D4** R27 | 3 casos (EA36, EA38, EA39) sobre 31 amb cossos basals | Taxa prou baixa: **R27 és senyal, no soroll**. S'implementa. *Advertència:* EA38 i EA39 estaven entre els «Not yet classified», de manera que la incoherència pot ser simplement que estaven a mig entrar |
| **D9** `T_LOST_ELEMENTS` | **0 files** a tota la taula | La promoció de les capes aplicades no té encara cap cas que la reclame. S'implementa igualment: és barata i el catàleg d'evidències ja la preveia |
| **D5** Art rupestre | **0 registres PR**, 2 estructures amb pigment sobre penya (EA11, EA39), 1 fila amb substrat Bedrock | Corpus mínim. S'implementa el conjunt sencer perquè, reintroduint dades, el cost marginal és baix i estalvia una v14 quan es documente La Petaca |

## Troballes no previstes

**(a) Els 9 heretats del defecte v10 són «no avaluat» disfressat.** `Fire_Damage`: 20 nous, 0 presents. `Modern_Access`: 17 nous, cap 0. Aquests registres es van entrar sota la v10, **on el defecte era 9**, de manera que la majoria d'eixos nous no són judicis deliberats sinó camps que ningú no ha tocat. Com que no són NULL, queden **fora de la llista de treball de la regla 17** i entren als recomptes com si fossen observacions.

Amb la reintroducció (rev. 4) el problema desapareix per complet. Els 0 eren pocs a tot arreu (màxim 3), o siga que el defecte 0 de la v11 encara no havia generat zeros falsos.

**(b) Valor fora de llista a `Pigment_Extent`.** Apareix `Arch. elements`, però el domini emmagatzema `Architectural elements`. Valor heretat que la migració no va normalitzar; amb `LimitToList = True` el registre es veuria buit al formulari. Resolt per la reintroducció.

**(c) La mateixa parella de suports en ordres oposats.** D3 mostra `Medium cavity + Fissure/Crack` (2 casos) i `Fissure/Crack + Medium cavity` (2 casos). Amb el criteri nou de primari/secundari (secció 3), aquests quatre casos **s'han de revisar al model en reintroduir-los**: o bé són genuïnament distints (en uns la cavitat rep càrrega, en altres l'escletxa), i llavors és una troballa sobre la versatilitat de la diàclasi, o bé el criteri no s'havia aplicat consistentment.

**(d) Formes de suport sense cap ús:** `Wide natural ledge`, `Natural niche` i `Micro-ledge` tenen zero registres. Es mantenen al lookup (el corpus de La Petaca pot usar-les). També hi ha un `ID_Support_Secondary = ND` que hauria de ser buit: ND com a secundari no vol dir res.

## Els sis registres pendents de revisió (D1)

`Fissure/Crack` com a **primari**: DW-S01-EA01, DW-S01-EA06, DW-S04-EA12b.
`Fissure/Crack` com a **secundari**: DW-S04-EA04, DW-S04-EA05, DW-S04-EA18.

En reintroduir-los, decidir per a cadascun si és **diàclasi** (vertical, confinament) o **junta d'estratificació** (horitzontal, encastament).

---

# **8. Semàntica del zero, valors per defecte i abast del domini de cinc valors** `[REVISIÓ 3]`

## 8.0. Plantejament

Diversos camps de tres valors tenen per defecte 0 (Absent) sense que això estiga justificat pel seu ús: `Mortar_Present`, `Chinking_Stones`, `Pigment_Present`, `Plaster_Present`, `Dec_Present`, `Looting`, `Fire_Damage`, `Animal_Activity`, `Modern_Access`. En tots ells, **no observar una cosa no implica que no hi siga**.

## 8.1. Què significa «absent»: una asimetria no documentada `[DECIDIT]`

| Domini | Significat del 0 | Com s'expressa la pèrdua |
| --- | --- | --- |
| Cinc valors (elements A–X) | No va existir mai, o no hi ha rastre que permeta afirmar-ho | Valor 3 + fila a `T_LOST_ELEMENTS` |
| Tres valors | **Fon dues coses**: «mai no en va tindre» i «en tenia i s'ha perdut» | No es pot expressar |

Per als **processos** (`Looting`, `Fire_Damage`, `Animal_Activity`, `Modern_Access`) la fusió és irrellevant: si no hi ha rastre, no hi ha hagut procés. Per a les **capes aplicades** (`Mortar_Present`, `Plaster_Present`, `Pigment_Present`, `Dec_Present`) és un problema real, perquè totes són coses que s'esborren: un morter rentat per l'aigua deixa una junta que sembla en sec; un revoc despres deixa una façana que sembla no haver-ne tingut mai.

**Indici que el problema és real:** `L_LOST_EVIDENCE` ja conté el valor *Mortar imprint* («empremta de morter sense l'element que unia»). El catàleg d'evidències preveia un cas que el domini del camp no pot expressar. Ho resol el punt 8.3.

**A documentar:** la taula anterior ha d'anar a l'esquema tècnic. La semàntica del 0 no és la mateixa als dos dominis, i això no estava escrit enlloc.

## 8.2. Valor per defecte dels camps sense gating `[DECIDIT]`

**Per què el defecte 0 no els correspon:** la v11 va adoptar el defecte 0 perquè **la regla 4 exigeix el zero de farciment** als components d'un sistema no aplicable — és a dir, per necessitat del gating dels elements A–X. Els camps llistats a 8.0 no pertanyen a cap sistema i no tenen cap farciment que satisfer: van heretar el defecte per uniformitat, no per raó pròpia.

**Asimetria dels errors**, que decideix la direcció del canvi:

- Un **0 fals afirma de més**: infla el denominador amb absències que ningú no ha verificat. Com que el biaix correlaciona amb la qualitat d'observació —que difereix entre LP i DW—, reintrodueix exactament l'artefacte que tot el domini existeix per a evitar.
- Un **9 fals afirma de menys**: perd casos, però no torça la direcció del resultat i és recuperable tornant a mirar.

**Decisió: defecte NULL** per a tots els camps de 8.0, no 9.

Raó: el 9 **no és una casella buida**, és una afirmació — «aquesta posició no es pot examinar», tal com es va fixar en canviar l'etiqueta de «ND» a «No observable». Si el 9 fos el defecte, un camp que ningú no ha tocat quedaria indistingible d'un camp marcat deliberadament com a inexaminable, i això **desactivaria la regla 17**, que existeix precisament per a llistar el que encara no s'ha avaluat.

Amb defecte NULL es conserven tres estats separats:

| Valor | Significat |
| --- | --- |
| NULL | No avaluat encara (llista de treball, regla 17) |
| 9 | Avaluat: no examinable |
| 0 | Avaluat: absent |

**Cost acceptat:** la regla 17 donarà llistes llargues durant l'entrada. Eixa llista llarga és la veritat, i és la mateixa mecànica que ja es gestiona amb els 35 registres heretats de la migració.

**Alternativa si es prefereix evitar NULLs** per comoditat d'entrada: defecte 9, clarament millor que 0. La pèrdua és només la funció de llista de treball.

**El defecte 0 es manté** allà on el gating el necessita: els components d'element A–X governats per un `Sys_*`.

## 8.3. Promoció a cinc valors de les capes aplicades `[DECIDIT]`

**No es fa de manera uniforme.** Estendre els cinc valors a tots els camps de tres seria una simetria estètica pagada amb la coherència del domini. La v11 ja ho va decidir així i el raonament es manté: «present parcial» i «desaparegut atestat» són **categories morfològiques d'elements construïts**.

Prova d'absurd: `Looting` = 2 («saqueig parcial») — el saqueig no té integritat física; `Fire_Damage` = 3 («dany per foc desaparegut») — sense sentit; `Mat_Textiles` = 3 exigiria evidència material d'un tèxtil desaparegut, categoria que no existeix.

**Es promocionen només els camps que registren una capa aplicada a la superfície de la fàbrica**, que són els únics del grup amb integritat física, degradació gradual i rastre en desaparéixer:

| Camp | 2 = present parcial | 3 = desaparegut atestat |
| --- | --- | --- |
| `Mortar_Present` | Morter conservat en zones | *Mortar imprint* a `L_LOST_EVIDENCE` |
| `Plaster_Present` | Revoc conservat en zones | Empremta o adherències |
| `Pigment_Present` | Pigment conservat parcialment | Rastre sobre substrat despres |
| `Dec_Present` | Decoració parcialment conservada | Evidència de decoració perduda |
| `RockArt_Present` (secció 7.5) | Ídem | Ídem |

**No es promocionen:** alteracions (processos), bioarqueologia i materials culturals (vestigis mobles, no fàbrica), `Support_Modified`, `Recessed_Frame` i `Chinking_Stones` (vegeu 8.6).

**Nota d'implementació:** el canvi és barat en dades — 0, 1 i 9 continuen significant el mateix i cap valor es migra — però toca `FillThreeValueFields` i `FillElementMap`, que són **font única de veritat** compartida pels defectes, QRY_13 i QRY_16. Caldrà una **tercera llista**: aquests camps tindran cinc valors sense ser elements A–X ni estar gatejats per cap sistema. Cal revisar que les regles que iteren sobre elements no els incloguen per error, i que el formulari els assigne el control de cinc valors (PC5) en lloc del de tres (PC9).

## 8.4. Criteri general d'assignació de domini `[DECIDIT]`

> **Cinc valors (0/1/2/3/9)** per a tot allò que és **fàbrica o capa aplicada sobre fàbrica**: té integritat física, es degrada gradualment i deixa rastre en desaparéixer.
> **Tres valors (0/1/9)** per a **processos, vestigis mobles i qualificadors**: no tenen integritat física pròpia, de manera que «parcial» i «desaparegut» no hi signifiquen res comprovable.

Això no és uniformitat, però és **regla**, que és el que fa l'esquema replicable per un tercer.

## 8.5. Solapament de `Plaster_Present` = 2 amb `Traces only` `[DECIDIT]`

En promocionar `Plaster_Present`, el valor 2 (present parcial) i el valor *Traces only* de `Plaster_Extent` es trepitgen parcialment.

**Criteri:** la **presència** registra l'estat de conservació de la capa; l'**extensió** registra la superfície que cobria originalment. Amb això, «revoc en traces sobre tota la façana» és `Plaster_Present = 2` + `Plaster_Extent = Full facade`, que és més informatiu que qualsevol dels dos sols.

**Alternativa si el criteri resulta difícil de sostindre en camp:** retirar *Traces only* de `Plaster_Extent`, ja que el valor 2 el cobriria.

## 8.6. Variància real de `Chinking_Stones` `[PENDENT DADES]`

Observació de camp: el ripio apareix pràcticament a tots els murs de maçoneria. Un camp quasi constant **no discrimina res** i no aporta a cap prova estadística.

**Comprovar la distribució real abans de decidir.** Si el seu valor com a indicador de qualitat de fàbrica (Toyne i Anzellini 2017) no es materialitza en aquest corpus, cal decidir entre retirar-lo, mantindre'l com a descriptor sense funció analítica, o substituir-lo per una variable més fina (densitat, regularitat, posició del ripio).

Mentre no es resolga, **no es promociona a cinc valors**: afinar el domini d'un camp que potser cal retirar seria treball inútil.

---

# **9. Ordre d'implementació previst**

Seguint el patró de la migració v10→v11 (blocs independentment verificables, ordre estricte, cap esborrat abans de la verificació):

| Pas | Contingut | Depén de |
| --- | --- | --- |
| 1 | Correcció del gating d'E (secció 1) | Res. Es pot aplicar ja: cap dada afectada |
| 2 | Comprovació del corpus: registres amb `Fissure/Crack` | Revisió manual |
| 3 | Desdoblament de `L_SUPPORT` + columna `Description` amb modes (secció 2) | Pas 2 |
| 4 | Reassignació dels registres afectats | Pas 3 |
| 5 | Regla R27 si es confirma útil (secció 4) | Comprovació contra corpus |
| 6 | Regla d'avís de suport assignat per fotografia (secció 6) | Decisió |
| 7 | Validació del criteri d'associació contra 5-6 casos reals (7.3) | Revisió manual |
| 8 | `L_STRUCT_BODY`: entrades ROC + `Sort_Order`; `L_DEC_TYPE`: tipus rupestres (7.2) | Pas 7 |
| 9 | `RockArt_Present` + `Color_Secondary` a `T_STRUCTURES` / `T_DECORATIONS` (7.5, 7.6) | Pas 8 |
| 10 | Migració dels registres amb `Color = Both` a la parella ordenada (7.6) | Pas 9 |
| 11 | Formulari: dues subseccions a 4.Dec amb combos filtrats (7.4) | Passos 8-9 |
| 12 | Regles R28-R30 i restricció de R22-R23 a files no-ROC (7.5, 7.9) | Pas 9 |
| 13 | `QRY_17_RockArt_All` (7.7) | Pas 8 |
| 14 | Defecte NULL als camps sense gating (8.2) | Res: només DAO, cap dada afectada |
| 15 | Promoció a cinc valors de les capes aplicades (8.3) | Pas 14. Toca la font única de veritat: revisar QRY_13 i QRY_16 |
| 16 | Controls PC5 per als camps promocionats al formulari (8.3) | Pas 15 |
| 17 | Decisió sobre `Chinking_Stones` (8.6) | Comprovació de variància |
| 18 | Actualització de `esquema_bbdd_estructures_v13.md` i `tfm_metodologia_bbdd_v13.md` | Passos 1-17 tancats |

**Recordatori de restriccions:** el gating viu duplicat en dos fitxers (`chachapoya_Form_v12_val.bas` i `upgrade_form_v12.bas`); qualsevol canvi s'ha d'aplicar als dos o el paquet diverge segons la ruta d'instal·lació. Es manté la regla B: cap valor emmagatzemat es toca sense verificació prèvia i comptador d'afectació al log.

---

# **10. Punts oberts pendents de decisió**

- **Confinament fabricat** (secció 3): camp propi si apareix repetidament?
- **Circulació vertical vs horitzontal** (secció 5): com es distingeixen a OE3?
- **R27** (secció 4): útil o generadora de falsos positius?
- **Avís de suport per fotografia** (secció 6): regla de validació o convenció escrita?
- **Litologia i potència dels bancs** (secció 2): val la pena registrar-les com a variable independent per a tancar l'argument de H02?
- **Relació cromàtica** (7.10): ajornat fins a tindre prou casos bicroms.
- **Conjunt concret d'entrades ROC i de tipus rupestres** (7.2): depén dels casos documentats.
- **Coordenades pròpies a `T_DECORATIONS`** (7.7): només si cal posició fina de les pintures associades.
- **`Chinking_Stones`** (8.6): retirar, mantindre com a descriptor o substituir per una variable més fina?
- **`Traces only` a `Plaster_Extent`** (8.5): es manté amb el criteri de conservació vs superfície original, o es retira?

---

# **11. Comprovacions pendents sobre el corpus**

Llista consolidada del que cal mirar abans de tancar el delta. Totes són sobre 35 registres, de manera que són qüestió de minuts, i cadascuna desbloqueja una decisió:

| Comprovació | Desbloqueja |
| --- | --- |
| Registres amb `ID_Support = Fissure/Crack`: quants verticals, quants horitzontals | Si el desdoblament és un simple renomenament o exigeix reassignació (2) |
| Freqüència de la diàclasi com a primària vs secundària | Lectura sobre com s'usa; possible material de resultats (2) |
| Parelles de formes de suport (diàclasi+cavitat, micro-repisa+diàclasi...) | Possible classificació de solucions constructives per a H02 (2) |
| `N_Basal_Bodies` > 0 amb `Base_Level` = 0 | Si R27 és útil o genera falsos positius (4) |
| Aplicar els quatre tests d'associació a 5-6 casos reals | Validació del criteri; detecta què hi falta (7.3) |
| Casos de pintura rupestre documentats: quines posicions calen realment | Conjunt d'entrades ROC (7.2) |
| Registres amb `Color = Both`: són resolubles a parella dominant/secundari? | Migració de 7.6 |
| Distribució real de `Chinking_Stones`: quants 0, 1 i 9 | Si el camp discrimina alguna cosa (8.6) |
| Casos de revoc o morter amb empremta sense la capa: n'hi ha? | Confirma el valor del 3 a les capes aplicades (8.3) |


---

# **12. Correccions de la revisió 5** `[IMPLEMENTAT]`

Sis punts, cinc dels quals **corregeixen o simplifiquen decisions de les revisions anteriors** a partir de contraexemples de camp. Es documenten com a correccions i no com a novetats, perquè el valor està en per què la formulació anterior fallava.

## 12.1. Retirada del domini de cinc valors a les capes aplicades `[CORREGEIX 8.3]`

**El criteri de 8.3 era mal formulat.** Deia «capes aplicades a la fàbrica», però el que de veres habilita el valor 3 és una altra cosa:

> **El valor 3 exigeix que l'evidència de la pèrdua siga de naturalesa DISTINTA de la cosa perduda.** Un encaix buit no és una mènsula; una empremta de morter no és morter. Per això funcionen.
>
> **El valor 2 exigeix poder inferir l'extensió original.** En un mur es veu fins on arribava; en una superfície pintada, no.

Aplicat als cinc camps promocionats, el criteri els exclou tots:

| Camp | Per què torna a tres valors |
| --- | --- |
| `Pigment_Present` | **L'única evidència de pigment és pigment.** Si en queda rastre és present; si no en queda, no es pot atestar res. El 3 i el 0 es confonen necessàriament. I «present complet» no és afirmable mai |
| `Plaster_Present` | Igual. A més, `Plaster_Extent` **ja té *Traces only*** i `Pigment_Extent` **ja té *Traces***: el 2 duplicava el que el camp d'extensió registrava |
| `Mortar_Present` | El morter **no és una capa que es perd per zones: és un atribut de la tècnica de fàbrica**. Un mur en sec no ho és a trossos. El que varia amb la conservació no és la presència sinó la visibilitat, i això ja ho registren `Facade_Observability` i `Mortar_Notes`. Un `Mortar_Present = 2` descriuria l'estat del model 3D, no l'estructura |
| `Dec_Present` / `RockArt_Present` | Són **judicis agregats**: cada motiu ja té la seua fila amb el seu estat. El 2 i el 3 duplicarien informació que viu al detall |

**Criteri general reformulat.** La distinció clau no és «fàbrica vs no fàbrica» sinó **discret vs continu**:

> **Cinc valors (0/1/2/3/9): els 20 elements A–X, i només ells.** Unitats constructives discretes, amb integritat física i extensió inferible, l'absència de les quals pot deixar evidència d'una altra naturalesa.
> **Tres valors (0/1/9): tota la resta.** Atributs de tècnica, capes contínues, processos, vestigis mobles i judicis agregats.

**Implicacions:**
- `FillLayerFields` **desapareix** del codi, i amb ella la tercera llista que era el punt delicat del paquet v13. La font única de veritat torna a ser binària: elements (cinc valors, defecte 0) i la resta (tres valors, defecte NULL).
- Els cinc camps tornen a PC9 al formulari.
- Les regles 22-23 i 28-29 tornen a comprovar `=1` en lloc de `In (1,2)`.
- **El punt 8.5 es tanca sense necessitat de criteri:** el solapament entre `Plaster_Present = 2` i *Traces only* desapareix perquè el 2 ja no existeix. Els camps d'extensió queden com a únic registre de la conservació parcial.
- La fila *Mortar imprint* de `L_LOST_EVIDENCE` es queda sense camp que la reclame. **Es manté al catàleg**: pot documentar la pèrdua d'un element travat amb morter.
- `Mortar_Notes` guanya importància: hi va tota la matisació sobre visibilitat i distribució que el domini ja no expressa.

## 12.2. `Doc_Basis` com a escala ordinal de qualitat documental `[SUBSTITUEIX el camp v11]`

El camp anterior barrejava mode d'accés i tècnica de captura en una llista sense ordre (*Direct access*, *Rope access*, *Ground photography*, *Distant photogrammetry*...). Passa a ser una **escala ordinal única**, on els valors s'exclouen i el primer és el grau màxim:

| Valor | Criteri |
| --- | --- |
| Direct access | **S'ha arribat físicament a l'estructura**, amb possibilitat de tocar i mesurar |
| Close-range (<5m) | Documentació a menys de 5 m **sense arribar-hi** (dron proper) |
| Medium-range (5-30m) | Entre 5 i 30 m |
| Long-range (>30m) | Més de 30 m (gigafoto amb teleobjectiu des de terra) |
| ND | No determinat |

**Els llindars van dins de l'etiqueta**, no només a la documentació: és el que fa que el criteri s'aplique igual sense consultar cap manual (mateix patró que «Wide natural ledge (>2m)»).

**La frontera entre els dos primers és física, no mètrica:** amb cordes s'està evidentment a menys de 5 m, de manera que el que distingeix *Direct access* és el **contacte**, no la distància — que és el que de veres canvia el registre.

**No es desdobla la tècnica de captura** (gigafoto, dron, 360°, fotogrametria): ja queda descrita a la pestanya 10.Doc per les URL. Pèrdua acceptada: una estructura visitada amb cordes però amb documentació gràfica només des de terra tria el grau màxim i matisa a `Notes`.

## 12.3. `Cultural_Materials_Present` `[CORREGEIX el disseny de 7.Mat]`

`ID_Material_Status` barrejava dues variables en una escala: *Good / Fair / Poor* són graus de conservació, però *Absent* és una afirmació de presència. **És el mateix error que `Lintel` tenia en v10** (un lookup de material que confonia presència amb tipus), i es resol igual:

- **`Cultural_Materials_Present`** (BYTE 0/1/9) — camp general, anàleg a `Human_Remains`.
- **`ID_Material_Status`** — *Good / Fair / Poor / ND*, **sense *Absent***.

Gating: 0 o 9 bloquegen **tot** 7.Mat inclòs l'estat de conservació, i el quick-fill ofereix posar els `Mat_*` a 0 o 9. La regla 19 apunta ara al camp nou, més directe que un `DLookup` sobre el nom del lookup.

**Criteri de frontera:** vestigis reduïts a pols o fragments irreconeixibles són `Present = 1` + `Status = Poor`. `Present = 0` es reserva per a cambra buida.

**`ID_Arch_Status` NO es desdobla igual:** una estructura sempre existeix — és el registre mateix. Els vestigis mobles poden no existir, i per això necessiten un camp de presència que l'arquitectura no necessita. La simetria aparent entre els dos camps d'estat és enganyosa.

## 12.4. `C14` governa el subformulari de datacions `[NOU]`

Amb `C14 = No`, el subformulari de `T_DATING` es desactiva (mateix patró que `Dec_Present`).

**Els segles queden sempre editables.** `Chrono_Start_Cent` i `Chrono_End_Cent` són l'atribució cronològica de l'estructura, i la font habitual **no és radiocarbònica** sinó tipològica o per analogia. Bloquejar-los amb `C14 = No` impediria registrar l'atribució de la immensa majoria del corpus, que és on més es necessita.

**`Chrono_Basis` valorat i descartat:** seria quasi constant (*Typological* per a tot excepte tres casos) i un camp sense variància no discrimina res — mateix criteri aplicat a `Chinking_Stones` (8.6).

**`C14` es manté YESNO.** No és una observació sobre l'estructura sinó una metadada del corpus: «existeix una mostra datada?». «No observable» no hi voldria dir res. Implicació tècnica: el gating d'aquesta porta llig un booleà, no un byte.

## 12.5. Totes les mètriques lineals en metres `[CORREGEIX v6]`

`Support_Width_cm` → `Support_Width_m`, `Support_Depth_cm` → `Support_Depth_m`, `Opening_Width_cm` → `Opening_Width_m`, `Opening_Height_cm` → `Opening_Height_m`.

**El guany no és estètic:** una consulta pot sumar o comparar camps sense conversions. Calcular quina fracció de l'amplària de la repisa ocupa l'estructura exigia dividir per 100 a mà, i eixe és el tipus d'operació on s'introdueixen errors silenciosos.

**Riscos i mitigacions:**
- Els controls es formaten amb 2 decimals perquè l'usuari veja el que ha escrit (0,62 m i no 0,6).
- Les etiquetes diuen la unitat.
- **Regla R31 (rang):** qualsevol dimensió lineal > 20 m és implausible. Barata, i captura precisament l'error d'entrar 62 pensant en centímetres — que cap altra regla no detectaria.

`Coord_Precision_m` ja estava en metres i no canvia.

## 12.6. `Ground` a `L_SUPPORT` `[NOU]`

Valor nou per a estructures **a la base del cingle, no penjades**.

> **`Ground`** — *Gravitational rest*. «Superfície de terreny a la base del cingle o al peu del vessant. L'estructura assenta sobre sòl, no en posició elevada sobre un buit.»

**Fa explícit un supòsit de fons que mai s'havia escrit:** les altres deu formes comparteixen una condició implícita —l'estructura està elevada sobre un buit—, i `Ground` n'és la negació. El criteri operatiu és negatiu i comprovable: **no hi ha buit a sota**.

**És l'única forma sobre sediment i no sobre roca**, i això no és un matís: sobre roca l'estructura no assenta i no cal fonamentació; sobre sediment sí. Probablement explica la presència o absència de jàcenes basals (A) i murets transversals (D).

**Orientacions d'entrada** (no gatejades, per si un cas les contradiu):
- `Height_Above_Base_m` hauria de valdre 0 o quasi. Candidata a regla si el cas es repeteix.
- `Sys_Platform` hauria de ser normalment *Not applicable* i no *Absent*: mènsules, bigues, filades en voladís i ràfec existeixen per a resoldre problemes de la verticalitat.

---

# **13. Punts oberts després de la revisió 5** *(superat per la secció 14)*

# **14. Tancament dels punts oberts** `[REVISIÓ 6]`

Repàs de tots els punts que les revisions 1-5 havien deixat oberts. **Cap no queda pendent de decisió**; els que continuen sense implementar-se ho fan amb un motiu escrit i un criteri de reobertura.

## 14.0. Ja resolts per decisions posteriors (neteja del registre)

| Punt | Estat real |
| --- | --- |
| **R27** (secció 4) | Implementada. El diagnòstic va donar 3 de 31: senyal, no soroll |
| **Entrades ROC i tipus rupestres** (7.2) | Implementades a la revisió 4: 3 posicions ROC + 6 tipus `RA *` |
| **`Chinking_Stones`** (8.6) | Es manté a 0/1/9 com a descriptor tècnic. No promocionat |
| **`Traces only` a `Plaster_Extent`** (8.5) | **Es tanca sol.** Amb la retirada de `Plaster_Present = 2` (12.1) el solapament desapareix: el valor d'extensió queda com a únic registre de la conservació parcial, que és el que ha de ser |

## 14.1. Descartats amb motiu `[DECIDIT]`

**Avís de suport assignat per fotografia** (secció 6). La majoria de registres tenen model 3D i accés directe, de manera que la regla dispararia només en els pocs casos que ja se saben febles. `Doc_Basis` registra la informació i `QRY_15` la reporta: una regla que marca el que ja es veu a la consulta és soroll.

**Litologia i potència dels bancs** (secció 2). En un mateix faralló la litologia és quasi constant i la potència exigiria mesures de camp no disponibles. Com a camp seria quasi buit — el mateix criteri aplicat a `Chrono_Basis` (12.4) i a `Chinking_Stones` (8.6). **Va a la descripció geològica general de la memòria**, no a la base.

**Coordenades pròpies a `T_DECORATIONS`** (7.7). No, tret que aparega la necessitat de posició fina. Les coordenades del pare basten a escala de faralló.

**Camp derivat monocrom/policrom** (7.10). Innecessari: `Color_Secondary` buit o ple **ja és** la distinció monocrom/policrom. Un camp derivat seria redundant.

## 14.2. `Vertical association` a `Connection_Type` `[IMPLEMENTAT]`

La hipòtesi de politja continua sent especulativa —cap cas amb evidència irrefutable, i les alternatives (suport d'escala, ancoratge) són igual de plausibles—, però **l'associació de verticalitat entre estructures sí que és observable**: proximitat, posició relativa, morfologia del faralló.

Valor nou a la llista de tipus de connexió. `Confidence` (High/Medium/Low), que ja existeix, en registra la seguretat. La hipòtesi mecànica concreta continua anant a `T_ARCH_FEATURES` amb els indicis a `Notes` (secció 5).

**Criteri quan concorren dues relacions** (junta vertical entre fàbriques *i* base compartida): **la junta mana sobre el suport**, perquè la junta porta la direcció cronològica (regla 20) i el suport compartit no.

## 14.3. `Supra-structural unit` a `L_GROUP_TYPE` `[IMPLEMENTAT]`

**El problema:** la convenció d'anomenar Xa / Xb les unitats d'un mateix conjunt vivia **només a la cadena del codi**, com a hàbit textual. Res no feia la relació consultable: comptar unitats supraestructurals exigia coincidències de text sobre `Code`, que és fràgil.

**La solució** no necessita cap camp nou: un tipus de grup a `T_GROUPS`, que ja té `Group_Code`, `N_Members` i `ID_Group_Type`.

**Divisió del treball amb `T_CONNECTIONS`**, que convé tindre clara:

| | Cardinalitat | Descriu |
| --- | --- | --- |
| `T_CONNECTIONS` | **Binària** | La junta: adossada, travada, superposició. **Porta la direcció cronològica** |
| `T_GROUPS` | **N-ària** | La unitat. Tres cossos Xa/Xb/Xc donen **tres connexions però un sol grup** |

**Criteri de registre, formulat ample:** dues o més **unitats constructives coherents amb atributs propis** sobre una base compartida — **no necessàriament dues cambres**. Un mausoleu més una plataforma en voladís qualifica, i és de fet un cas freqüent. Això és consistent amb el principi d'unitat mínima de registre ja fixat: els atributs que difereixen entre episodis constructius no s'han de col·lapsar en un sol registre.

*Freqüència observada: casos puntuals a Diablo Wasi, molt més freqüent a La Petaca — de manera que el camp guanyarà valor a mesura que s'incorpore aquell corpus.*

## 14.4. `Color_Secondary` sense `ND` `[IMPLEMENTAT]`

Un color secundari indeterminat no diu res útil. Si es veu que hi ha un segon color però no se'n distingeix el to, el registre honest és deixar-lo buit i explicar-ho a `Notes`. Així el camp només afirma «hi ha segon color i és aquest».

*Casos bicroms documentats: dos tons de roig en art rupestre, cercles amb perímetre roig i emplenat blanc, franges blanca i roja en mausoleus. La varietat existeix, però la distinció monocrom/policrom ja captura el gruix de la informació.*

## 14.5. Ajornats amb criteri de reobertura `[OBERT]`

Tots comparteixen la mateixa lògica: **amb un o dos casos no es poden inventar categories**, i `Notes` o `T_ARCH_FEATURES` els conserven fins llavors.

| Punt | Es reobri quan |
| --- | --- |
| **Mecanisme de les mènsules aïllades** (politja, suport d'escala, ancoratge) | Aparega un cas amb solc, desgast o posició diagnòstica |
| **Relació cromàtica** (vora vs juxtaposició vs entrellaçat) | Hi haja prou casos per a que les categories les donen les dades |
| **Naturalesa de les estructures sobre `Ground`** | Es documenten. **Decidit: es queden com a valor de suport, no com a tipologia** — són funeràries com les altres |
| **Regla d'alçada per a `Ground`** | Hi haja casos que la justifiquen |

---

# **15. Comprovacions per a la reintroducció**

El que cal mirar al model mentre s'entren els registres:

- Els **sis registres amb `Fissure/Crack`** (DW-S01-EA01, EA06, S04-EA12b com a primari; S04-EA04, EA05, EA18 com a secundari): decidir diàclasi o junta d'estratificació.
- Els **quatre casos de parella invertida** (cavitat+escletxa en un ordre, escletxa+cavitat en l'altre): aplicar el criteri de la secció 3 i veure si la diferència és real.
- L'`ID_Support_Secondary = ND` que hauria de ser buit.
- Les **parelles Xa/Xb**: crear la fila de `T_GROUPS` corresponent.
- Els **casos candidats a `Vertical association`**.


---

# **16. Revisió 7** `[IMPLEMENTAT]`

## 16.1. Un volum i una superfície `[SUBSTITUEIX]`

`Interior_Vol_m3` i `Total_Vol_m3` **no volien dir el mateix segons la tipologia**, cosa pitjor que no tindre'ls. En un mausoleu la diferència entre els dos **és la fàbrica construïda**; en una cambra dins d'una cavitat, l'«interior» és espai natural que ningú no va excavar i el «total» inclou roca, de manera que la resta no significa res comparable. Fer la mitjana d'una columna sobre les dues tipologies hauria donat una xifra sense sentit.

**`Area_m2` + `Volume_m3`.** El volum és **l'espai funerari**: comparable entre totes les tipologies i el que es relaciona amb el MNI (H01).

El desglossament fi —volum per cos (N0/N1), superfície de plataforma, volum construït— **va a `Vol_Notes`**: són casos puntuals i un camp quedaria buit en la majoria de registres. Si resulten freqüents, aleshores es formalitza — mateix criteri que la hipòtesi de politja i la relació cromàtica.

## 16.2. `Recessed_Frame` passa a 2.Arq `[CORREGEIX]`

És un **qualificador del pla de façana**, no del portal: el recul afecta el parament sencer i el portal hi queda inscrit. Gatejat darrere de `Sys_Portal` es bloquejava precisament on hi ha recul però no portal. Passa al costat de `Masonry_Type` i `Facade_Orientation`, i surt de la cadena de components de l'emplenat ràpid.

**No es promociona a `Sys_Facade`**, tot i que la pregunta era raonable. Els cinc sistemes actuals són **combinacions de components identificables** (E+F+G, N+O+Q, S+T) o camps d'agrupació per al gating. Una façana no és una combinació de res — és el pla on tota la resta passa —, i **no podria valdre *Absent* mai**. Un sistema que no pot ser absent no fa el que fan els sistemes.

## 16.3. `Name_VAL` a tots els lookups `[CORREGEIX una asimetria]`

**El problema observat:** el formulari mostrava uns desplegables en valencià i altres en anglés. No era un error puntual sinó **dos mecanismes, un d'incomplet**: els combos de llista de valors (`PC5`, `PC9`, `PCV`) porten les dues columnes a la cadena del codi i es veien en valencià; els combos de taula llegien `Name` i es veien en anglés. Mig formulari en cada idioma.

**La solució és el patró que ja existia**, aplicat de manera completa: columna emmagatzemada oculta, etiqueta visible. Ací l'emmagatzemat és un **ID numèric**, cosa que ho fa encara més segur que a les llistes de valors — reanomenar una etiqueta no pot tocar cap dada.

Columna `Name_VAL` a `L_TYPOLOGY`, `L_STATUS`, `L_MATERIAL_STATUS`, `L_SUPPORT`, `L_DEC_TYPE`, `L_STRUCT_BODY`, `L_GROUP_TYPE`, `L_LOST_EVIDENCE`, `L_VOL_METHOD` i `L_COORD_METHOD`. `L_ELEMENTS` ja en tenia.

**`Name` continua sent el terme de referència:** exportacions, publicació, i les regles de QRY_16 que hi busquen per nom (`Record_Class` via tipologia, `L_STATUS` per a l'avís de col·lapse) continuen llegint `Name`.

Poblat per `UPDATE ... WHERE Name = ...` en un sol bloc llegible, no reescrivint cada `INSERT`. Les files sense traducció cauen a l'anglés en lloc de quedar buides.

**Sobre el castellà** (valorat, ajornat): afegir `Name_ES` seria copiar aquest bloc. El que **no** és barat és un selector d'idioma en viu, i convé que quede escrit per què: al formulari hi ha quatre menes de text i les traduccions del lookup només afecten una. Els combos de llista de valors tenen les etiquetes **al codi**; les etiquetes dels controls i els noms de pestanyes estan fixats al disseny. Commutar només els combos de taula deixaria mig formulari en cada idioma — el mateix defecte que aquesta revisió corregeix. La via neta, si algun dia cal, és **parametritzar `BuildForm`** amb un codi d'idioma i generar un formulari sencer i independent.
