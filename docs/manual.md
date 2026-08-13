**MANUAL D'ÚS DE LA BASE DE DADES**

*Estructures funeràries de La Petaca i Diablo Wasi*

Base de dades v19 | Formulari `F_STRUCTURES` | Esteve Ribera Torró

*Aquest manual serveix per a **emplenar** la base de dades. No explica per què està dissenyada així: això és a `esquema_bbdd_estructures_v19.md` (referència tècnica) i a `tfm_metodologia_bbdd_v19.md` (justificació). Ací només hi ha el que cal per a decidir què escrius.*

> **Què ha canviat des de la v16, en una ullada.** La posició relativa de 4.Dec ha desaparegut i a la meitat rupestre la substitueixen **quatre caselles de tram** i la **geometria del traç** (nus 8). L'estat dels vestigis mobles ha baixat de 5.Estat a **7.Mat**, davall de la pregunta que l'obri. A 11.Sist hi ha un **sisé sistema**, el d'interfície. **Totes les pestanyes tenen ara camp de notes**, i la llista de tipus d'art rupestre s'ha refet (nus 8). A 2.Arq hi ha tres camps nous: **posició del portal** i **fàbrica** (nus 9). I a 12.Extra ara veus **les connexions que t'han registrat des d'altres estructures**.
>
> A més: el **marc reculat** ha passat a «Morfologia general», l'**evidència de fase** té nou valors en compte de cinc, la tipologia `EA-PLA-R Plataforma en repisa` ara es diu **`EA-TER Terrassa en repisa`**.

> **Què canvia en v18, en una ullada.** A 11.Sist la plataforma té un component nou, la **superfície (Z)**, i el material en penja; el **muret transversal (D) ja no es bloqueja mai** — el seu 0 és sempre una observació teua; la cornisa intercòs té **format de pedra** propi; i hi ha una casella nova, **«Cornisa fa de llindar»**, que només s'obri quan pot tenir sentit (nus 2). A 2.Arq el **marc reculat es bloqueja sol si no hi ha cos de cambra**, l'**orientació del portal es bloqueja quan el pla d'accés és la façana** (és la mateixa dada: s'entra una vegada), i el format de pedra distingeix ara **blocs tabulars regulars** de **blocs de grans dimensions** (nus 4). La tipologia **EA-TER tanca sola** tot el bloc de façana i portal, i el **suport es filtra per tipologia** (els nínxols i les superfícies de roca s'autoomplin). A 12.Extra la cronologia de connexions es diu ara amb **dos camps**: la relació (*Seqüencial / Contemporanis / Indeterminat*) i **quina estructura és l'anterior, pel seu codi** — i la llista de només lectura ja no gira res. La **posició** dels elements desapareguts ha desaparegut (va a notes). Hi ha un tipus de connexió nou, **Context natural associat**. I la llista de treball de la migració és **`QRY_23_V18_Review`**.

> **Què canvia en v19, en una ullada.** A 11.Sist el **muret transversal (D) té capçalera pròpia** amb el criteri que el separa del mur de retorn (V): *no tanca cambra ni es compta com a cos* (nus 11). Hi ha una casella nova al portal, **«Brancal resolt en fàbrica»**, que només s'obri amb els brancals a 0 o 2 (nus 2). Les dues caselles de compartició —esta i «Cornisa fa de llindar»— responen ara **No / Sí / No observable**: són preguntes, no presències. Les **etiquetes del rol de mènsules** s'han reescrit al voltant de l'única pregunta del camp, el vincle amb la plataforma, i **«Vincle indeterminat» (ND) és una resposta completa** — la bateria ja no la marca (nus 10). La tipologia MEN es diu ara **Element estructural aïllat**: serveix per a qualsevol element solt, no sols mènsules. I la llista de treball de la migració és **`QRY_24_V19_Review`**.

---

# **0. Abans de començar**

## Què necessites obert

| | |
| --- | --- |
| La base de dades | Formulari `F_STRUCTURES` |
| La fitxa de camp | `fichas_estructuras.xlsx` |
| La documentació gràfica | Fotografies, gigafotos, panoràmiques, model 3D si n'hi ha |
| Aquest manual | Sobretot la **part 3** |

**Regla prèvia:** si un valor no el pots justificar assenyalant alguna cosa a la fitxa o a la imatge, no l'escrigues. Deixa'l buit.

## Què és un registre

Un registre és **una seqüència constructiva coherent**, no «una cosa que es veu al farallò».

Si dues masses adossades tenen tipologia distinta, o accés independent, o conservació molt diferent, **són dos registres** vinculats per `T_CONNECTIONS` (pestanya 12.Extra), no un registre amb valors mesclats. Fondre-les en un fa que cap dels dos valors siga cert.

`Construction_Phases` (2.Arq) és per a fases **dins d'un mateix registre**, no per a estructures distintes.

## Les classes de registre

La tipologia que tries a **1.Id** obri o tanca pestanyes senceres. No és un error del formulari: és el disseny.

| Classe | Què tanca |
| --- | --- |
| Estructura construïda | Res |
| Context funerari natural | 2.Arq |
| Traça estructural | 2.Arq, 6.Bio, 7.Mat |
| Panell d'art rupestre (PR) | 2.Arq, 6.Bio, 7.Mat, 11.Sist |

Si una pestanya està grisa, comprova primer la tipologia.

---

# **1. Les tres regles d'or**

*Si només llegeixes una pàgina d'aquest manual, que siga aquesta.*

## Regla 1 — Buit, 0 i 9 són tres coses diferents

| Escriu | Quan | Vol dir |
| --- | --- | --- |
| **(buit)** | Encara no ho has mirat | «Feina pendent» |
| **9** | Ho has mirat i **no es pot examinar** | «Judici: aquí no puc veure-ho» |
| **0** | Ho has mirat i **no hi és** | «Judici: aquí no hi ha res» |

**El 0 és una afirmació, no un valor de farciment.** Si no estàs segur que l'element falte, no poses 0: deixa-ho buit o posa 9.

**El 9 també és una afirmació.** Vol dir «he anat a mirar i la façana està tapada / arrasada / massa lluny». No vol dir «no ho sé».

*Per què importa: els percentatges del TFM es calculen sobre els 0 i els 1. Un 0 posat per omplir infla el denominador amb absències que ningú no ha comprovat.*

## Regla 2 — Present vol dir diferenciat

Un element A–X **hi és** quan hi ha una peça **físicament distinta** del mur del costat: per format de pedra, per dimensió, per material o per tractament.

Si la vora de l'obertura és maçoneria contínua amb el mur, sense cap peça distinta → **l'element és 0**.

Això val per a tot el vocabulari, no només per al portal.

## Regla 3 — Quan dubtes, tria l'opció que es pot desfer

- Entre **una fitxa i dues** → dues, vinculades. Fusionar després és fàcil; separar, no.
- Entre **posició elemental i posició de sistema** a 4.Dec → elemental. Agregar després es pot; desagregar, no.
- Entre **fila associada i fitxa pròpia** en art rupestre → fitxa pròpia amb agrupació. No perds res.
- Entre **un valor dubtós i el buit** → buit.

---

# **2. Recorregut per la fitxa**

*Ordre recomanat d'entrada. No és l'ordre de les pestanyes: és l'ordre en què les decisions depenen unes d'altres.*

## Pas 1 — 1.Id (identificació)

Codi, jaciment, sector, **tipologia** i suport geològic.

La tipologia va primer perquè decideix quines pestanyes s'obrin. El suport (`ID_Support`) també importa més del que sembla: és el que després decidirà si una pintura sobre la penya està «pròxima» o és fitxa a banda.

Si el suport té dues formes, usa la **parella ordenada**: la principal és **la que rep la càrrega**, la secundària la que estabilitza. No hi ha valor «mixt» — és deliberat.

**El suport es filtra per la tipologia (v18).** Amb un nínxol natural (`NIX`) o un panell rupestre (`PR`) el camp **s'ompli sol** — la relació és 1:1 i no cal que la respongues; amb una cavitat (`CAV`) la llista es redueix a les cavitats i tu tries entre mitjana i gran. `Indeterminat` és sempre legal: és un dubte, no una contradicció (regla 54).

## Pas 2 — 2.Arq (morfologia, maçoneria, façana, fases)

**Cossos basals (N0) i cossos cambra (N1).** Compta **masses de maçoneria**. Una plataforma **no** és un cos.

**Façana i paisatge.** Ací hi ha tres camps que es confonen fàcilment:

| Camp | Què és |
| --- | --- |
| Orientació façana | Cap a on mira el **pla exposat** (el que es veu des de la vall) |
| Pla d'accés | Per quin pla de l'estructura **s'entra**: façana / mur lateral / fons |
| Orientació portal | Cap a on mira **l'obertura** |

Normalment el pla d'accés és «façana», i en eixe cas **l'orientació del portal es bloqueja sola (v18)**: és la mateixa dada que l'orientació de façana, i s'entra **una vegada**. Quan l'accés no és per la façana —obertura al mur estret perpendicular al farallò— és quan aquests camps guanyen el seu sou, i llavors sí que els emplenes els dos.

**El fons és el contrari de la façana, no de l'entrada (v18).** Tot el que es diu «posterior» o «fons» —el pla d'accés `Fons`, el tancament posterior de la cambra— es defineix **contra el pla de façana** (el pla exposat), mai contra l'accés: si s'entra pel mur de retorn, el fons continua sent el que era.

**Si no hi ha cambra**, el pla d'accés i l'orientació del portal es bloquegen sols: sense interior no hi ha per on entrar. L'orientació de façana i la visibilitat de la vall **es queden obertes**, i les has d'emplenar igual: un cos basal sol també té pla exposat i també mira cap a algun lloc.

*Es bloquegen només quan **declares** que no hi ha cambra. Si el camp està buit no es bloqueja res, perquè buit vol dir «encara no ho he mirat».*

**El marc reculat es bloqueja sol si no hi ha cos de cambra (v18)**: un marc reculat qualifica el pla de façana, i sense cos de cambra no hi ha façana a qualificar. Quan poses els cossos de cambra a 0, la casella rep el seu 0 tècnic sense preguntar-te res.

**Si la tipologia és `EA-TER Terrassa en repisa`, tot aquest bloc es tanca sol (v18)**: una terrassa no té ni façana ni portal. El sistema portal es declara *No aplicable* automàticament — però només si encara portava l'*Absent* per defecte; un valor que hages declarat tu no es toca mai (la regla 48 t'avisarà si queda incoherent).

**Posició del portal** (nou v17): `Centrat` / `Descentrat a l'esquerra` / `Descentrat a la dreta`. On se situa el portal dins del pla de façana. **Esquerra i dreta des de tu, mirant la façana de front** — la mateixa convenció que els trams del nus 8.

**Maçoneria.** Vegeu el **nus 4**. I si la fàbrica no és igual per tot arreu, el **nus 9**.

## Pas 3 — 11.Sist (sistemes i elements A–X)

Va abans que 3.Acab i 4.Dec perquè els acabats i la decoració es descriuen **sobre** elements que has d'haver identificat.

**Primer els sis sistemes, després els components.** El sistema mana: si el poses a *No aplicable*, els seus components es bloquegen i es reompliran de 0 automàticament (te'n demanarà confirmació).

| Sistema | Valors |
| --- | --- |
| Conjunt basal (A–D), Conjunt cambra, **Interfície** | Present / Absent / No aplicable / No observable |
| Plataforma, Portal, Ràfec | Present complet / Present parcial / Desaparegut atestat / Absent / No aplicable / No observable |

**La plataforma té un component nou (v18): la superfície (Z).** El que trepitges — el paral·lel exacte de la superfície del ràfec (T). Va amb la mateixa escala de cinc valors que la resta, i **el material de superfície ara en penja**: sense superfície registrada, el material es bloqueja (regla 51). Una plataforma pot sostenir-se només amb Z: si veus la superfície però no pots resoldre què l'aguanta, Z amb valor i E-F-G segons el que veges.

**El muret transversal (D) ja no es bloqueja mai (v18) i té capçalera pròpia (v19).** Ha eixit del conjunt basal — apareix desvinculat de la massa basal: sobre plataformes volades, com a suport adossat, o aïllat — i funciona com la cornisa (I) i el coronament (R): sempre actiu. La capçalera del formulari porta el **discriminador**, no la geometria: *no tanca cambra (això és V) ni es compta com a cos* — el test complet és al **nus 11**. **Conseqüència que has de recordar: el seu 0 és sempre una observació teua**, mai un farciment automàtic. Els zeros de farciment heretats de la v17 han tornat a buit amb el patch v19: `QRY_24_V19_Review` te'ls llista per al judici real.

**Una plataforma amb tots els components a 0 és legal (v18).** Una `EA-PLA-V` amb el sistema present i E, F, G i Z tots a 0 no és cap incoherència: és el resultat que la plataforma es va resoldre sense cap d'eixos elements diferenciats, i diu alguna cosa sobre la inversió de treball. No «arregles» eixos zeros.

**El sistema d'interfície és nou (v17)** i governa **només la cornisa intercòs (I)**. Si l'estructura té un sol cos, no hi pot haver cornisa entre cossos: posa'l a `Absent` des del principi i t'estalvies la pregunta. **El coronament (R) no en depén** i continua sempre actiu — amb un sol cos l'estructura té part de dalt igualment.

**La cornisa té format de pedra propi (v18):** lloses laminars contra blocs tabulars, el mateix parell del parament aplicat a la peça de la cornisa. Només s'obri mentre la cornisa té entitat (regla 50).

**La diferència entre *Absent* i *No aplicable*:** *Absent* és un resultat (no hi havia plataforma, i això vol dir alguna cosa). *No aplicable* vol dir que la pregunta no té sentit en aquest registre. *Absent* compta com a 0 a les anàlisis; *No aplicable* no compta.

**Excepció que has de recordar:** les **mènsules (E)** no es bloquegen mai, encara que la plataforma siga absent. Una mènsula aïllada no ha d'haver portat cap plataforma.

Nusos que et trobaràs ací: **1** (valors dels elements), **2** (portal i les dues caselles de compartició), **3** (cornisa vs voladís), **5** (coberta), **10** (rol de les mènsules), **11** (muret transversal, mur de retorn i MEN).

## Pas 4 — 3.Acab (revoc i pigment)

Presència primer, detall després. Si el revoc és 0 o 9, els camps de detall es bloquegen.

**El pigment d'aquesta pestanya és el pigment de l'estructura** — sobre la fàbrica. La pintura sobre la penya va a 4.Dec. Vegeu el **nus 6**.

## Pas 5 — 4.Dec (decoració i pintura rupestre)

Dues subseccions: decoració arquitectònica a dalt, pintura rupestre a baix. **Són la mateixa taula**, però **cada meitat té els seus camps**: la de dalt porta el número de cos, la de baix porta els quatre trams i la geometria del traç.

Una fila per motiu. Nusos **6**, **7** i **8**.

## Pas 6 — 5.Estat, 6.Bio, 7.Mat

Conservació dual: `Estat arquitectònic` (l'obra) i `Estat dels vestigis mobles` (el contingut). Són independents: una estructura pot estar sencera i saquejada, o arruïnada amb el contingut protegit.

**Han canviat de lloc (v17):** l'estat dels vestigis mobles **ja no és a 5.Estat, sinó a 7.Mat**, just davall de la pregunta «Materials culturals». Ara les tres preguntes van en l'ordre que toca:

| A 7.Mat, per ordre | Què et demana |
| --- | --- |
| **Materials culturals** | *N'hi ha?* Es pot contestar des de lluny: veus que dins hi ha coses |
| Els set tipus | *De quins?* Cal identificar-los |
| **Estat dels vestigis** | *En quin estat?* Cal veure'ls prou bé |

*No són tres maneres de dir el mateix.* Si de lluny veus que hi ha material però no distingeixes de quin tipus és, la resposta correcta és **`Materials culturals = 1` i els set tipus a `9`**. Si dedueixes la primera dels set, eixe registre passaria a dir «no hi ha materials», que és fals.

Ara **5.Estat parla només de l'obra** (com està, què l'ha malmesa, com de bé s'ha vist) i **7.Mat només del contingut**, amb la mateixa forma que 6.Bio.

A 6.Bio i 7.Mat el domini és 0/1/9 i la regla 1 val igual.

## Pas 7 — 8.Cron, 9.Metr, 10.Doc

**9.Metr es queda buida en la primera passada**, i és correcte: les dimensions, la volumetria i les coordenades depenen de Metashape, CloudCompare i QGIS. No cal forçar-les. `QRY_16s_Observational_Nulls` (i la regla 17) et donaran la llista del que falta, camp a camp i registre a registre.

A 10.Doc, `Doc_Basis` i `Facade_Observability` **no són burocràcia**: són el que després permetrà distingir «no s'hi veia» de «no era decidible». Emplena'ls sempre.

## Pas 8 — 12.Extra

Connexions amb altres estructures, elements arquitectònics no previstos i elements desapareguts.

**Si has posat cap element a valor 3 (desaparegut atestat), ací has d'obrir la fila d'evidència.** Sense ella, el 3 no és defensable i la regla 2 el marcarà.

### Connexions (canvia en v17 i en v18)

Ara hi ha **dues llistes**:

| Llista | Què hi ha |
| --- | --- |
| Connexions registrades **des d'aquesta estructura** | On escrius |
| Connexions registrades **des d'altres estructures** | Només lectura. El que t'han registrat a tu |

**Mira la segona abans d'afegir res.** Si la connexió ja hi és, no la tornes a entrar: el programa no t'ho permetrà, perquè la mateixa parella no es pot registrar dues vegades.

A la llista de només lectura, la cronologia es mostra **des del teu punt de vista**: si l'anterior és l'altra estructura, ací llegiràs «aquesta és posterior».

**La cronologia es diu ara amb dos camps (v18).** La **relació** diu només si hi ha direcció: `Seqüencial` / `Contemporanis` / `Indeterminat`. I si és seqüencial, el camp **`Anterior`** diu **quina estructura és l'anterior, pel seu codi** — has de triar una de les dues de la parella, i el formulari no et deixa guardar sense fer-ho (regla 52). L'ordre en què escrius les estructures ja no importa gens: el programa normalitza la parella i **no ha de girar res**, perquè l'anterior està nomenada per identitat i no per posició.

*Recorda: la direcció cronològica només es pot llegir d'una **junta vertical adossada** o d'una **superposició**. Amb qualsevol altre tipus, `Indeterminat` (regla 20).*

**Tipus nou: `Context natural associat` (v18).** El cas que el va motivar: una terrassa amb un nínxol natural al seu extrem. El registre correcte és **dues fitxes** — la terrassa i el nínxol, cadascú amb la seua tipologia, seguint la regla de la seqüència constructiva mínima — i **una connexió d'este tipus** entre elles, que diu que van junts **sense afirmar cap mecanisme constructiu**: no hi ha junta, ni suport compartit, ni alineació; hi ha un accident natural que forma part funcional del conjunt.

**Elements desapareguts: la posició ha desaparegut (v18).** Si cal dir on era, va a les **notes** de la fila. I l'abast es diu ara `Cos constructiu` en compte de `Cos` — perquè en un context funerari «cos» era ambigu de mala manera. Regla nova (53): **si nomenes un element, l'abast ha de ser *Element***; per a un cos o l'estructura sencera, deixa el codi buit.

---

# **3. Els onze nusos**

*Els punts on el registre s'encalla de veres. Cadascun és un arbre de decisió.*

---

## Nus 1 — Quin valor pose a un element A–X

Els vint-i-un camps d'element (A–X més la superfície de plataforma, Z) van amb aquesta escala:

| | |
| --- | --- |
| **0** Absent | He mirat i no hi ha element diferenciat |
| **1** Present complet | Hi és en tota la seua extensió esperada |
| **2** Present parcial | Hi és però incomplet, i puc inferir-ne l'extensió original |
| **3** Desaparegut atestat | No hi és, però **hi ha evidència que hi va ser** |
| **9** No observable | No es pot examinar |

**El 3 és l'únic valor que és una inferència.** Exigeix evidència **de naturalesa distinta de la cosa perduda**: un encaix buit, una empremta, una cicatriu de despreniment, un forat de biga. «Sembla que hi devia haver alguna cosa» no és evidència.

Cada 3 necessita fila a **12.Extra → elements desapareguts**, amb el tipus tret del catàleg tancat. **Si cap tipus del catàleg no encaixa, el valor correcte no és 3: és 0 o 9.**

**I si el que s'ha perdut és un sistema sencer (v19, criteri escrit).** Per a atestar un **sistema** perdut —el portal sencer, la cambra sencera— la fila d'evidència porta abast **Cos constructiu** o **Estructura sencera** i el **codi d'element buit**: el codi és per a peces concretes, i el desplegable exclou les lletres de sistema a propòsit. Una sola fila d'abast ampli cobreix tot el vocabulari del cos perdut.

**Els comptadors de cossos inclouen els perduts atestats (v19, criteri escrit).** `N_Basal_Bodies` i `N_Chamber_Bodies` compten **el que es va construir**: un cos en *Desaparegut atestat* amb la seua fila d'evidència **compta**. Excloure'l faria que els recomptes mesuraren preservació i es llegiren com a densitat constructiva.

### Cas freqüent: element asimètric

Un brancal de pedra laminar i l'altre de maçoneria corrent. ¿Disseny asimètric o brancal perdut?

Sense evidència de pèrdua **no ho pots decidir**, i l'escala no té valor per a «asimètric per disseny». Escriu:

- Valor **2**
- A `Notes sistemes`: `asymmetric: left only` (o `right only`), **exactament així**

*La fórmula constant importa: si el patró es repeteix, una consulta podrà trobar tots els casos i decidirem si mereix camp propi.*

*Per al cas concret del **brancal**, la v19 resol part de l'ambigüitat: si el costat sense brancal acaba en una **cara terminal treballada i deliberada**, ja no és «no ho puc decidir» — és un brancal resolt en fàbrica, i té casella pròpia (nus 2).*

---

## Nus 2 — Quan el portal té llindar, brancals i dintell

Aquesta és la que més canvia respecte de com es feia abans.

> **Un element del portal hi és quan hi ha una peça diferenciada del mur.** Si l'obertura és un buit deixat a la maçoneria, sense peces distintes → `Llindar = 0`, `Brancals = 0`, `Dintell = 0`.

**I el sistema portal continua sent «present».** No és contradictori:

| | Registra |
| --- | --- |
| Sistema portal | Que **hi ha obertura** |
| N, O, Q | Si **cada posició es va resoldre amb una peça distinta** |

**Aquests zeros són el resultat, no un buit.** Que a Diablo Wasi les obertures es resolguen sovint sense llindar ni brancals diferenciats diu alguna cosa sobre com es treballava. Si poses 1 «perquè hi ha obertura», eixa informació desapareix i tots els portals queden iguals.

### Test ràpid

> Tape mentalment l'obertura. **Es distingeix encara alguna peça del mur del voltant?**
> Sí → l'element hi és (1 o 2). No → 0.

### La cornisa que fa de llindar (v18)

Cas real del corpus: el portal no té llindar diferenciat **perquè la cornisa intercòs ja fa eixa faena** — l'obertura descansa directament sobre la cornisa. El registre honest és:

- `Llindar (N) = 0` — no hi ha peça diferenciada del llindar, i el 0 és el resultat
- `Cornisa intercòs (I)` amb el seu valor — la peça existeix i és una cornisa
- **`Cornisa fa de llindar` = Sí** — la casella nova, que només s'obri exactament en aquest cas (N a 0 i cornisa amb entitat)

Així no perds la informació de com es va resoldre la posició del llindar, i N continua dient la veritat: no hi ha llindar **diferenciat**. Si la casella està bloquejada és que la finestra no es dona; no la busques (regla 49).

### El brancal resolt en fàbrica (v19)

L'altre cas de compartició, ara amb casella pròpia: la posició del brancal no té peça diferenciada **perquè la fàbrica mateixa la resol** — el mur acaba en una **cara terminal acabada i deliberada** (*masonry reveal*), treballada com a vora de l'obertura. El registre honest és:

- `Brancals (O) = 0` (cap costat amb peça) o `= 2` (un costat amb brancal, l'altre resolt en fàbrica)
- **`Brancal resolt en fàbrica` = Sí** — la casella nova, que només s'obri amb O a 0 o 2 (regla 55). Amb O = 1 està bloquejada: no queda cap posició per resoldre d'una altra manera

La fórmula constant `jamb: masonry reveal, dressed` que escrivies a `Notes sistemes` **ja no cal per als casos nous** — la casella la substitueix; les notes antigues són la pista per a repassar els candidats que `QRY_24` et llista.

*Les dues caselles de compartició responen **No / Sí / No observable** (v19): són **preguntes** («la cornisa fa de llindar?», «la fàbrica resol el brancal?»), no presències — «Absent» no deia el que passa. Els valors guardats són els mateixos de sempre.*

---

## Nus 3 — Filades en voladís (G) o cornisa intercòs (I)

El nus més difícil, sobretot quan a penes queda vestigi de la plataforma.

**Poden coexistir.** No és un o l'altre per força: hi ha estructures amb plataforma volada al cos basal **i** cornisa marcant la junta amb la cambra.

Diferència de fons:

| | |
| --- | --- |
| **Plataforma (H)** | Una **superfície**: horitzontal, practicable, amb profunditat útil |
| **Cornisa (I)** | Una **junta**: lineal, sense profunditat útil, no practicable |

### Quatre preguntes, en ordre. La primera que respon, decideix.

**1. Hi ha encaixos, forats passants, fusta conservada o mènsules al buit?**
→ Sí: **plataforma**. És l'evidència més forta i la que sobreviu millor — la roca conserva l'encaix molt després que la fàbrica caiga. Obri fila a 12.Extra.

**2. Hi ha cos construït damunt (present o atestat)?**
→ No: **no pot ser cornisa**. Una cornisa *intercòs* sense segon cos és una contradicció. És **G**.

**3. El voladís és progressiu o d'un sol gest?**
→ Filades successives, cadascuna més enfora, acumulant volada: **G**.
→ Una o dues filades que sobreïxen per igual i el mur reprén la vertical damunt: **I**.

**4. Continuïtat?**
→ Recorre tota l'amplària de la façana marcant la junta: **I**.
→ Es concentra on hi havia superfície practicable, sovint desbordant l'amplària del cos: **G**.

### I una regla que t'estalvia el dubte

> **Una cornisa és sempre de pedra.** Si l'element volat és de fusta, no és cornisa: és mènsula (E) o biga transversal (F), i el que tens és **plataforma**.

### Si cap pregunta no decideix

Passa, i és poc freqüent. Escriu:

- `Cornisa intercòs = 9`
- `Sistema plataforma = No observable`
- Descripció a `Notes sistemes`
- **Fila a 12.Extra → elements arquitectònics** amb el detall morfològic

No és elegant, però triar-ne una per costum inventa dades. La profunditat de volada es podrà mesurar amb Metashape més endavant i aleshores es podrà decidir en bloc.

---

## Nus 4 — Format i treball de la pedra

Dos camps, i **no descriuen el mateix**.

### `Format pedra` — quina forma té la peça

Descriu **el parament**, no el dintell ni la superfície de ràfec ni el fons de cambra, que tenen camp de material propi.

| Valor | Test |
| --- | --- |
| **Pedres irregulars** | La peça no té dues cares planes subparal·leles. Assenta com puga i la junta l'ajusta el ripio |
| **Blocs tabulars regulars** | Dues cares planes **i alçada de filada pròpia**: una pedra, una filada |
| **Blocs de grans dimensions** (v18) | **Una sola peça fa la filada per ella mateixa**, sense companyes al mateix nivell. És el senyal d'inversió de treball més fort que el test pot llegir |
| **Lloses laminars** | Dues cares planes, però **cal apilar-ne diverses per fer una filada** |

> **El test és: quantes peces fan una filada.** Es llig directament a la façana i no cal mesurar res.
> Si dubtes entre tabular i laminar: gruix ≤ ¼ de la dimensió major → laminar.
> *Per què «regulars» (v18): un bloc gran també és tabular, i el terme nu havia deixat de nomenar una classe. Els valors antics es van migrar automàticament al patch.*

**Si al parament conviuen dos formats**, usa la parella. El **dominant és el que ocupa més superfície de mur**, no el que té més peces — amb lloses menudes i blocs grans, comptar peces inverteix el resultat.

No hi ha valor «mixt»: és deliberat, la parella ja ho diu.

### `Treball pedra` — com de treballada està

| | |
| --- | --- |
| **Sense treballar** | No hi ha traces de treball |
| **Semiescairada** | Treball parcial |
| **Escairada** | Treball complet |
| **Mixt** | Conviuen graus clarament distints |
| **Indeterminat** | No decidible |

> **Registra evidència positiva de treball:** traces d'eina, aristes vives regulars, cares que trenquen el pla de fractura natural.
> **Si no hi ha eixes traces, la cara plana és de la fractura → «sense treballar».**

Això és important perquè un gres estratificat fractura en cares que **semblen** carejades. «Sense treballar» no afirma que ningú no la tocara: afirma que no hi ha evidència que la tocaren.

**Quan les cares planes són compatibles amb les dues coses i no hi ha traça diagnòstica → `Indeterminat`.**

*Cap dels dos camps té «no observable». Si no s'hi veia, posa `Indeterminat` i que `Doc_Basis` i `Facade_Observability` (10.Doc) expliquen per què. Per això has d'emplenar-los.*

---

## Nus 5 — Coberta de cambra (X)

> **Sense contacte, no hi ha sostre.**

X registra que l'espai funerari està **tancat per dalt**, per obra o per roca, **en contacte amb la fàbrica**.

| Cas | X | Tipus |
| --- | --- | --- |
| Obra de maçoneria damunt | 1 | Obra de maçoneria |
| Bigues i lloses | 1 | Fusta i lloses |
| Visera rocosa **tocant** la coronació | 1 | Penya natural |
| Obra a l'entrada, roca al fons | 1 | **Mixt** |
| Visera que passa per damunt **sense tocar** (20 cm o 3 m, igual) | **0** | — |

L'últim cas: **no és un sostre**. Que hi haja abric ho registra `ID_Support` a 1.Id, i pots descriure-ho a `Notes sistemes`. El test és binari, sense gradient: buit és buit.

**Cambres superposades:** el sostre de la cambra de baix **és** el paviment de la de dalt. Compta com a X de la de baix, amb el tipus que corresponga a l'obra.

*El mateix criteri val per a `Fons de cambra (W)`: penya natural no és fàbrica.*

---

## Nus 6 — Pintura: de l'estructura o rupestre?

Quatre preguntes **geomètriques**, en ordre. La primera que respon, decideix.

### 1. El motiu està **sencer** sobre la fàbrica?

→ **Pigment de l'estructura.** Va a **3.Acab** (`Pigment present`, extensió, color) i, si vols detallar-ne la posició, fila a la **subsecció de decoració arquitectònica** de 4.Dec.

### 2. Té **part sobre la fàbrica i part sobre la roca**?

→ **Fila a la subsecció rupestre, posició `Solapada`.** Substrat = `Mixt`. A 3.Acab, `Pigment present = 1` amb substrat `Mixt`.

**Una fila, no dues.** És un sol motiu.

*Aquest és el cas més interessant de tots: un motiu que travessa la junta demostra que qui pintava tractava roca i mur com una sola superfície, i que la pintura és posterior a la fàbrica en eixe punt.*

### 3. **Ressegueix el contorn** de l'estructura, en contacte tangencial?

→ **`Perimetral`.**

Val també si **el cos ja no hi és**: bandes verticals sobre roca nua alineades amb on hi havia els brancals són perimetrals, encara que ara no toquen res. Diu *contorn*, no *contacte*. Si és aquest cas, obri fila a 12.Extra → elements desapareguts.

**I si el que ha desaparegut és l'estructura sencera?** Si a la penya només queda la banda pintada i cap fàbrica, la pregunta és si allí hi va haver una construcció:

> **Traç ortogonal**, o alineat amb elements reconeixibles (brancals, base, coronament) → **era una estructura**. Fes-ne un **registre d'estructura** amb tipologia `No classificable`, tot el vocabulari A–X a **9**, i una fila a 12.Extra amb abast `Estructura sencera` i evidència `Pigment sobre penya`. La banda va com a fila `Perimetral`.
> **Traç corb o irregular**, sense alinear-se amb res → **és art rupestre**. Fitxa PR.

*Per què estructura i no art rupestre quan el traç és ortogonal: si la registres com a art rupestre, desapareix del recompte d'estructures del sector — i els recomptes per sector són el que mesuren les hipòtesis sobre distribució. Comptaries preservació creient que comptes construcció.*

### 4. **No toca res**, però és dins del mateix accident del farallò?

→ **`Pròxima`.** Mateixa repisa, mateixa cavitat, mateixa escletxa, mateix rebaix.

**Si cal travessar una vora de banc, una cornisa natural o un buit per a arribar-hi → no és pròxima: és fitxa PR pròpia**, vinculada per `T_GROUPS` com a *Rock art cluster*.

### Recorda

- `Panell` **no es tria mai** des d'un registre d'estructura. És per a les files d'un registre PR.
- El substrat de la meitat rupestre només accepta **Penya** i **Mixt** (aquest, només amb posició Solapada). Si necessites «revoc» o «pedra de parament», la fila està a la meitat equivocada.
- **Si la vora que separaria no la veus en tres dimensions** (una foto frontal confon un pla que avança amb un que retrocedeix), la sortida conservadora és **fitxa PR amb agrupació**. No perds res i es pot desfer.

---

## Nus 7 — Quina posició pose a una decoració

Una fila per motiu. La posició diu **on està**; el tipus diu **què és**. No barreges els dos.

### Posicions arquitectòniques

`Cos basal (B/C)` · `Cornisa intercòs (I)` · `Cantonera (J)` · `Pilastra (K)` · `Flanc de façana (L)` · `Mur lateral de cambra (V)` · `Sistema portal (P)` · `Llindar (N)` · `Brancal (O)` · `Dintell (Q)` · `Sobre-dintell (M)` · `Coronament (R)` · `Ràfec (U)`

Hi ha també `No determinada`, i vol dir exactament això: **la posició arquitectònica no s'ha pogut determinar**. No l'uses mai per a una pintura sobre la penya — per a això hi ha les quatre posicions rupestres (nus 6). Confondre-ho és el mateix error que confondre 0 amb 9.

### Tres confusions que has d'evitar

**Sobre-dintell no és el dintell.** És la zona **per damunt**, l'àrea de fris. El dintell és `Dintell (Q)`.

**Llindar és la peça de baix** (la que es trepitja). La de dalt és el **dintell**. Sona semblant a *llinda* i per això a la interfície no s'usa mai eixa paraula.

**Flanc de façana (L) i mur lateral de cambra (V):**

| | Pla | Gira? |
| --- | --- | --- |
| **L** Flanc de façana | Dins del pla de façana | No |
| **V** Mur lateral de cambra | Perpendicular | Sí, cap al penyal |

### Portal pintat: una fila o tres?

> **Continu al voltant de l'obertura**, sense poder dir on comença i acaba cada component → **`Sistema portal (P)`**, una fila.
> **S'interromp entre elements**, o només en decora un → posició **elemental**, una fila per element.
> **En cas de dubte: elemental.**

### Cos basal

Hi ha **una sola** posició per al cos basal. No busques «sòcol»: es va retirar perquè era circular.

- Un **relleu** al cos basal → fila `Cos basal` **i** `Sòcol decoratiu (C) = 1` a 11.Sist (un relleu **és** un tractament plàstic)
- Una **banda pintada** sense motllura → fila `Cos basal` i `Sòcol decoratiu = 0`

### La posició relativa ja no existeix (v17)

Si vens de la v16: el camp `Posició relativa` **s'ha llevat**. A la meitat arquitectònica no el substitueix res —si un brancal està decorat i l'altre no, anota-ho a `Notes` de la fila—, i a la meitat rupestre el substitueixen els quatre trams del **nus 8**.

*Es va llevar perquè feia dues faenes distintes: a la meitat arquitectònica «esquerra» volia dir **quin dels dos** brancals, i a la rupestre **on està la pintura** respecte de l'estructura. No és el mateix, i per això la pregunta «esquerra respecte de què?» no tenia resposta.*

---

## Nus 8 — Els quatre trams i la geometria del traç

**Només a la meitat rupestre.** Quatre caselles, una per cada costat del contorn de l'estructura:

| Casella | Vol dir |
| --- | --- |
| **Esquerra** | El motiu ocupa el costat esquerre |
| **Damunt** | Ocupa la part de dalt |
| **Dreta** | Ocupa el costat dret |
| **Davall** | Ocupa la part de baix |

Cadascuna amb els tres valors de sempre: **0** absent, **1** present, **9** no observable.

> **Esquerra i dreta des de tu, mirant la façana de front** — el pla exposat, el que es veu des de la vall. **Mai des del pla d'accés.**

### La casella que més importa és el 9

Si el tram de dalt està tapat per la visera de roca, o pel cos superior, o en ombra: **9, no 0**.

*És la diferència entre una U invertida i un anell tancat mal conservat. Si hi poses 0, estàs afirmant que allí no hi havia pintura — i no ho has pogut mirar.*

### No busques «U invertida» a cap llista

**No cal.** Marca els trams i el programa ho dedueix sol:

| Trams marcats | El que llegiràs a `QRY_20` |
| --- | --- |
| Esquerra + Damunt + Dreta | **U invertida** |
| Els quatre | **Envoltant** |
| Esquerra + Dreta | **Flanquejant** |
| Un de sol | **Parcial** |

I si algun tram val 9, la fila apareix marcada com a **infralegida**: vol dir que la forma que veus pot no ser la que hi havia.

### El tipus de motiu

La llista rupestre (set valors, refeta en v17a):

| Valor | Quan |
| --- | --- |
| `Antropomorf` | Figura humana |
| `Zoomorf` | Figura animal |
| `Forma U geomètrica` | Banda en U invertida amb **angles rectes i traç net** |
| `Forma U orgànica` | Banda en U invertida **corba o mal definida** |
| `Motiu geomètric` | Línies, bandes o figures amb organització regular |
| `Taca amorfa` | Pigment amb **vora reconeixible** però sense motiu identificable |
| `Traces de pigment` | Restes disperses o massa degradades per a dir res |

**Les dues confusions que aquesta llista arregla:**

> **Taca amorfa contra traces**: la pregunta és si el pigment té **vora llegible**, no si sembla significatiu.
> **Les dues U**: no és una qüestió estètica. La geomètrica dibuixa un rectangle i per tant **diu que allí hi havia una construcció** (nus 6); l'orgànica pot estar seguint un rebaix natural.

*Si vens de la v17: el camp separat `Geometria del traç` ha desaparegut. El que deia ara ho diuen les dues formes en U.*

---

## Nus 9 — La fàbrica no és igual per tot arreu

**A 2.Arq. Sempre actiu**, també amb un sol cos: una estructura d'un sol cos pot tindre dues fàbriques.

Quan el cos basal està fet de pedres irregulars i la cambra de blocs tabulars ben escairats, els camps de maçoneria només et deixen dir `Mixt` — i eixe `Mixt` **esborra justament el que interessa**: que hi ha dues fàbriques, i per tant potser dues mans, dues fases o dues intencions.

**Fàbrica:**

| Valor | Quan |
| --- | --- |
| `Única` | Una sola fàbrica a tot el registre |
| `Múltiple per cossos` | Cada cos és uniforme per dins, però difereixen entre ells |
| `Múltiple dins d'un cos` | Almenys un cos té més d'una fàbrica per dins |
| `No observable` | No s'hi veu prou |

**La pregunta no és on hi ha el canvi, sinó si el canvi segueix els cossos o els travessa.** Un cos que ja té dues fàbriques per dins també serà distint del seu veí; si preguntàrem «on», les respostes se solaparien. Per això els dos valors plurals **no són una parella simètrica**: el segon no és el contrari del primer.

*Amb un sol cos, l'únic valor plural possible és `Múltiple dins d'un cos`. Si tries `Múltiple per cossos` amb un sol cos, la regla 45 t'ho marcarà.*

*I la distinció importa: el segon valor és exactament el cas que una descomposició per cossos resoldria; el tercer és el que no resoldria.*

### El que canvia exactament, a les notes

Si el que varia és el format, el treball, l'aparell o el morter, **escriu-ho a `Notes` d'arquitectura**. La consulta de repàs de notes ho recuperarà tot junt.

### No confongues divergència amb fase

> **Fàbrica múltiple = el que veus.** La pedra canvia.
> **Fase constructiva = el que argumentes.** Hi va haver dos moments de construcció.

---

## Nus 10 — El rol de les mènsules (E)

El camp respon **una** pregunta: *la mènsula formava part del sistema plataforma?* No és la foto («la trobe sola») ni la funció d'ús (això va a 12.Extra o a notes): és el **vincle amb H**, agregat a nivell d'estructura.

### Tres preguntes, en ordre. La primera que respon amb un sí, decideix.

**1. L'evidència sosté l'atestació?** Diverses mènsules alineades al mateix nivell, encaixos buits, cicatrius → `Sistema plataforma = Desaparegut atestat` + rol **Component de plataforma (H)** + fila a 12.Extra amb el tipus **«Mènsules al buit»** (existeix al catàleg precisament per a això).

**2. Pots afirmar que MAI no va sostenir plataforma?** Context llegible al seu nivell i net de tot vestigi → rol **Sense vincle amb plataforma**. Exigeix haver pogut llegir el context, com el 0 exigeix haver mirat.

**3. Ni una cosa ni l'altra** → rol **Vincle indeterminat (ND)**. És una **resposta completa, no una pendent** — la bateria ja no la marca (v19). La sospita, redactada a `Notes sistemes`.

*La frontera entre «sense vincle» i «indeterminat» és la mateixa frontera 0/9 de sempre.*

**«De les dues classes» (Both)**: l'estructura té mènsules de les dues menes — per exemple tres alineades sota la plataforma i una desvinculada a un altre nivell.

**El recompte i el rol no s'han de quadrar.** `Núm. mènsules` compta **el que sobreviu**; el rol interpreta **el que era**. Una sola mènsula supervivent amb rol «Component de plataforma» (cas 1) no és cap contradicció.

*Coherència vigilada (regla 56): rol de suport amb el sistema plataforma a `Absent` o `No aplicable` és contradicció — si sostenien plataforma, el sistema és present o atestat.*

---

## Nus 11 — Muret transversal (D), mur de retorn (V) i el registre MEN

Amb la façana paral·lela al faralló, D i V són **geomètricament idèntics**: tots dos perpendiculars a la roca. El que els separa no és l'orientació sinó **què fan**.

### Test de delimitació

> **Muret transversal (D)**: fàbrica perpendicular al faralló que **no tanca cap interior ni es compta com a cos**.
> **Si tanca cambra → és mur de retorn (V).**
> **Si sosté plataforma pel davall com a peça encastada → mireu mènsules (E) / bigues (F).**
> **Si està sol → el registre és MEN.**

**Implicació de comptadors**: V implica cos (és paret de N1 — cambra amb accés present o atestat, i la regla 11 ho vigila); **D no n'implica cap**. Un muret transversal amb els dos comptadors a 0 és un registre correcte.

**D pot ancorar a qualsevol alçada de la fàbrica** — sobre una plataforma volada, com a suport adossat que sosté una jàssera, o aïllat. Per això no penja de cap sistema i el seu 0 és sempre una observació teua.

### El registre MEN (v19: Element estructural aïllat)

**Com es registra**: l'element a 11.Sist (D present = muret, E present = mènsula — la tipologia no porta subtipus perquè els camps ja ho diuen), sistemes tancats, i **comptadors a 0 necessàriament**: sense cossos no hi ha nivell del qual parlar.

**Frontera amb l'estructura atestada**: si l'evidència sosté una **estructura concreta** —mènsules alineades amb encaixos = plataforma perduda—, el registre **no és MEN** sinó l'estructura amb el sistema *Desaparegut atestat* i la seua fila d'evidència. **MEN = quan no es pot afirmar cap estructura.** La interpretació (circulació, estructura prèvia, suport) va a 12.Extra o a notes, mai a la tipologia.


Pot haver-hi canvi de fàbrica sense cap fase: un canvi de proveïment, o dos paletes el mateix dia. **Si declares dues fases, has d'emplenar `Evidència fase`** — si no, la regla 46 t'ho marcarà, i amb raó.

---

# **4. Quan no ho saps**

*La font d'error més comuna no és equivocar-se: és **triar el valor que sembla més informatiu** quan no en tens prou.*

## Taula de decisió

| Situació | Escriu |
| --- | --- |
| No hi he arribat encara | **Buit** |
| Hi he anat, però la façana està tapada / arrasada / massa lluny | **9** o `ND` |
| Ho veig bé, però la distinció no és decidible | **9** o `ND`, i explica a Notes |
| Ho veig bé i no hi ha element diferenciat | **0** |
| Crec que hi era però ja no hi és, **i tinc evidència** | **3** + fila a 12.Extra |
| Crec que hi era però ja no hi és, **sense evidència** | **0** o **9** — mai 3 |

## Les tres coses que no has de fer

**No òmpligues per completar la fitxa.** Una fitxa amb buits és una fitxa honesta. `QRY_16` amb la regla 17 et donarà la llista del que falta, camp a camp.

**No uses 9 com a «no ho sé».** El 9 diu «he anat i no es pot veure». Si no hi has anat, és buit.

**No inventes un 3.** Si cap tipus d'evidència del catàleg tancat no encaixa, el valor és 0 o 9. El catàleg funciona com a filtre de la inferència: si no hi caps, la inferència no està prou sostinguda.

## Notes amb fórmula constant

Quan una observació no té camp, va a Notes — però **escrita sempre igual**, o després no es podrà trobar:

| Cas | Escriu literalment |
| --- | --- |
| Portal asimètric | `asymmetric: left only` |
| Banqueta adossada (12.Extra, `Feature_Type`) | `Abutted mass`, i a Notes `function: access` |
| Abric superior sense contacte | `overhead shelter, no contact` |

*Si el mateix cas apareix tres vegades o més, això és el senyal que necessita camp propi. Amb la fórmula constant, una consulta ho detecta; amb text lliure, no.*

---

# **5. Abans de tancar la fitxa**

## Comprovació ràpida

- [ ] Tipologia posada (decideix quines pestanyes s'obrin)
- [ ] Els sis sistemes de 11.Sist tenen valor
- [ ] **Els 21 elements sense cap buit** — és el llindar mínim perquè la fitxa entre a l'anàlisi A–Z (v18): un element NULL trau l'estructura sencera de la matriu
- [ ] Cada element a **3** té la seua fila a 12.Extra
- [ ] `Doc_Basis` i `Facade_Observability` emplenats (10.Doc)
- [ ] `Pla d'accés` emplenat, i `Orientació portal` **si el pla no és la façana** (2.Arq; amb accés per la façana el camp queda bloquejat i la dada viu a `Orientació façana`)
- [ ] Cada fila rupestre té els quatre trams i la geometria del traç
- [ ] Si hi ha dos cossos o més, `Fàbrica` emplenada
- [ ] Si has declarat dues fases, `Evidència fase` emplenada
- [ ] Cap 0 que no puges justificar

## Les dues consultes

**`QRY_16_Validation_Check`** — 53 regles de coherència actives (numerades fins a 54). **Resultat buit vol dir corpus coherent.**

No és una llista d'errors: la **regla 17** hi apareix a propòsit per a llistar els camps encara buits. És la teua llista de feina, no una queixa.

**`QRY_19_V16_Review`** — la feina que va deixar oberta la transferència a v16 i que continua vigent: portals a rejudicar, cobertes pendents del test de contacte, format de pedra per entrar (encara falta en 32 registres), i files de decoració que podrien haver de canviar de posició.

**`QRY_21_V17_Review`** — la que ha deixat oberta la transferència a v17: files rupestres amb els trams per omplir, files de decoració **sense cap posició**, connexions per rellegir, estructures de dos cossos o més sense divergència declarada, i fases sense evidència.

**`QRY_23_V18_Review`** — la de la migració a v18: superfícies de plataforma (Z) per observar on el sistema és obert, candidats de «cornisa fa de llindar», files d'elements desapareguts amb abast incoherent i seqüencials sense anterior nomenada.

**`QRY_24_V19_Review`** — la de la migració a v19: candidats de «brancal resolt en fàbrica» (brancals a 0 o 2 amb el 0 de migració per confirmar o apujar), zeros del muret transversal (D) retornats a buit pel patch (el judici, ara sí, l'has de fer tu — el 9 és resposta legítima), i mènsules sense rol per passar per l'arbre del nus 10.

**`QRY_20_RockArt_Span`** — no és de revisió sinó de lectura: totes les files rupestres amb la forma del contorn ja calculada.

---

# **6. El vocabulari A–Z d'una ullada**

*De baix a dalt. Els tres sistemes van marcats. La lletra **Y** està reservada (banqueta, pendent de casos); **Z** entra en v18.*

## Nivell 0 — cos basal

| | | |
| --- | --- | --- |
| **A** | Jàsseres basals | Bigues encastades a la base |
| **B** | Basament | Massa basal diferenciada |
| **C** | Sòcol decoratiu | Tractament plàstic del basament |
| **D** | Muret transversal | *Transverse wall* (v19). No tanca cambra ni compta com a cos — nus 11. *No es bloqueja mai (v18)* |
| **E** | Mènsules | *No es bloqueja mai pel sistema* |
| **F** | Bigues transversals | |
| **G** | Filades en voladís | ⚠ vegeu nus 3 |
| **Z** | Superfície de plataforma | *Nou v18.* El que trepitges; paral·lel de T |
| **H** | **Plataforma** | **SISTEMA** = E + F + G + Z |

## Interfície N0/N1

| | | |
| --- | --- | --- |
| **I** | Cornisa intercòs | **Sempre pedra.** ⚠ vegeu nus 3 |

## Nivell 1 — cos cambra

| | | |
| --- | --- | --- |
| **J** | Cantoneres | |
| **K** | Pilastres estructurals | |
| **L** | Flanc de façana | Dins del pla de façana |
| **M** | Fris decoratiu | La zona sobre-dintell |
| **N** | Llindar | La peça **de baix** |
| **O** | Brancals | |
| **P** | **Obertura d'accés** | **SISTEMA** = N + O + Q |
| **Q** | Dintell | La peça **de dalt** |
| **R** | Coronament | |
| **V** | Mur de retorn | Perpendicular, gira cap al penyal |
| **W** | Fons de cambra | **Sempre el contrari de la façana**, mai de l'entrada (v18) |

## Zona superior

| | | |
| --- | --- | --- |
| **S** | Biga de suport del ràfec | Pot ser fusta |
| **T** | Superfície del ràfec | **Sempre pedra** |
| **U** | **Ràfec-voladís** | **SISTEMA** = S + T. **Sempre pedra** |
| **X** | Tancament superior de cambra | ⚠ vegeu nus 5 |

---

*Dubtes sobre criteris no coberts ací: `esquema_bbdd_estructures_v19.md`, secció 8bis.*
