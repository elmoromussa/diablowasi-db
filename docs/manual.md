**MANUAL D'ÚS DE LA BASE DE DADES**

*Estructures funeràries de La Petaca i Diablo Wasi*

Base de dades v17 | Formulari `F_STRUCTURES` | Esteve Ribera Torró

*Aquest manual serveix per a **emplenar** la base de dades. No explica per què està dissenyada així: això és a `esquema_bbdd_estructures_v17.md` (referència tècnica) i a `tfm_metodologia_bbdd_v17.md` (justificació). Ací només hi ha el que cal per a decidir què escrius.*

> **Què ha canviat des de la v16, en una ullada.** La posició relativa de 4.Dec ha desaparegut i a la meitat rupestre la substitueixen **quatre caselles de tram** i la **geometria del traç** (nus 8). L'estat dels vestigis mobles ha baixat de 5.Estat a **7.Mat**, davall de la pregunta que l'obri. A 11.Sist hi ha un **sisé sistema**, el d'interfície. A 2.Arq hi ha tres camps nous: **posició del portal** i **fàbrica** (nus 9). I a 12.Extra ara veus **les connexions que t'han registrat des d'altres estructures**.
>
> A més: el **marc reculat** ha passat a «Morfologia general», l'**evidència de fase** té nou valors en compte de cinc, la tipologia `EA-PLA-R Plataforma en repisa` ara es diu **`EA-TER Terrassa en repisa`**, i hi ha un **botó** dalt a la dreta per a omplir amb 0 els camps buits de la pestanya on estàs.

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

## Pas 2 — 2.Arq (morfologia, maçoneria, façana, fases)

**Cossos basals (N0) i cossos cambra (N1).** Compta **masses de maçoneria**. Una plataforma **no** és un cos.

**Façana i paisatge.** Ací hi ha tres camps que es confonen fàcilment:

| Camp | Què és |
| --- | --- |
| Orientació façana | Cap a on mira el **pla exposat** (el que es veu des de la vall) |
| Pla d'accés | Per quin pla de l'estructura **s'entra**: façana / mur lateral / fons |
| Orientació portal | Cap a on mira **l'obertura** |

Normalment el pla d'accés és «façana» i les dues orientacions coincideixen. Quan no —obertura al mur estret perpendicular al farallò— és quan aquests camps guanyen el seu sou. Emplena'ls els dos sempre.

**Si no hi ha cambra**, el pla d'accés i l'orientació del portal es bloquegen sols: sense interior no hi ha per on entrar. L'orientació de façana i la visibilitat de la vall **es queden obertes**, i les has d'emplenar igual: un cos basal sol també té pla exposat i també mira cap a algun lloc.

*Es bloquegen només quan **declares** que no hi ha cambra. Si el camp està buit no es bloqueja res, perquè buit vol dir «encara no ho he mirat».*

**Posició del portal** (nou v17): `Centrat` / `Descentrat a l'esquerra` / `Descentrat a la dreta`. On se situa el portal dins del pla de façana. **Esquerra i dreta des de tu, mirant la façana de front** — la mateixa convenció que els trams del nus 8.

**Maçoneria.** Vegeu el **nus 4**. I si la fàbrica no és igual per tot arreu, el **nus 9**.

## Pas 3 — 11.Sist (sistemes i elements A–X)

Va abans que 3.Acab i 4.Dec perquè els acabats i la decoració es descriuen **sobre** elements que has d'haver identificat.

**Primer els sis sistemes, després els components.** El sistema mana: si el poses a *No aplicable*, els seus components es bloquegen i es reompliran de 0 automàticament (te'n demanarà confirmació).

| Sistema | Valors |
| --- | --- |
| Conjunt basal (A–D), Conjunt cambra, **Interfície** | Present / Absent / No aplicable / No observable |
| Plataforma, Portal, Ràfec | Present complet / Present parcial / Desaparegut atestat / Absent / No aplicable / No observable |

**El sistema d'interfície és nou (v17)** i governa **només la cornisa intercòs (I)**. Si l'estructura té un sol cos, no hi pot haver cornisa entre cossos: posa'l a `Absent` des del principi i t'estalvies la pregunta. **El coronament (R) no en depén** i continua sempre actiu — amb un sol cos l'estructura té part de dalt igualment.

**La diferència entre *Absent* i *No aplicable*:** *Absent* és un resultat (no hi havia plataforma, i això vol dir alguna cosa). *No aplicable* vol dir que la pregunta no té sentit en aquest registre. *Absent* compta com a 0 a les anàlisis; *No aplicable* no compta.

**Excepció que has de recordar:** les **mènsules (E)** no es bloquegen mai, encara que la plataforma siga absent. Una mènsula aïllada no ha d'haver portat cap plataforma.

Nusos que et trobaràs ací: **1** (valors dels elements), **2** (portal), **3** (cornisa vs voladís), **5** (coberta).

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

## El botó d'emplenat ràpid

Dalt a la dreta: **«Omplir amb 0 els buits d'aquesta pestanya»**.

Posa a 0 tots els camps 0/1/9 que estiguen buits **a la pestanya on estàs**, després de dir-te quants són i de demanar-te confirmació. No toca els bloquejats ni els que ja tenen valor.

> **Fes-ho quan hages acabat de mirar la pestanya, no abans.** El 0 vol dir «he mirat i no hi ha»; si el poses sense mirar, després no hi haurà manera de saber quins vas comprovar de veritat.

## Pas 8 — 12.Extra

Connexions amb altres estructures, elements arquitectònics no previstos i elements desapareguts.

**Si has posat cap element a valor 3 (desaparegut atestat), ací has d'obrir la fila d'evidència.** Sense ella, el 3 no és defensable i la regla 2 el marcarà.

### Connexions (canvia en v17)

Ara hi ha **dues llistes**:

| Llista | Què hi ha |
| --- | --- |
| Connexions registrades **des d'aquesta estructura** | On escrius |
| Connexions registrades **des d'altres estructures** | Només lectura. El que t'han registrat a tu |

**Mira la segona abans d'afegir res.** Si la connexió ja hi és, no la tornes a entrar: el programa no t'ho permetrà, perquè la mateixa parella no es pot registrar dues vegades.

A la llista de només lectura, la cronologia es mostra **des del teu punt de vista**: si des de l'altra estructura han dit que ella és l'anterior, ací llegiràs «aquesta és posterior».

**La cronologia canvia de redacció.** Ara diu quina de les dues estructures és l'anterior: `A és anterior` / `B és anterior` / `Contemporanis` / `Indeterminat`. Si escrius les estructures en un ordre i el programa te les gira, **gira també la cronologia**, de manera que el sentit es manté. Ja no cal recordar cap norma sobre l'ordre.

*Recorda: la direcció cronològica només es pot llegir d'una **junta vertical adossada** o d'una **superposició**. Amb qualsevol altre tipus, `Indeterminat`.*

---

# **3. Els nou nusos**

*Els punts on el registre s'encalla de veres. Cadascun és un arbre de decisió.*

---

## Nus 1 — Quin valor pose a un element A–X

Els vint camps d'element van amb aquesta escala:

| | |
| --- | --- |
| **0** Absent | He mirat i no hi ha element diferenciat |
| **1** Present complet | Hi és en tota la seua extensió esperada |
| **2** Present parcial | Hi és però incomplet, i puc inferir-ne l'extensió original |
| **3** Desaparegut atestat | No hi és, però **hi ha evidència que hi va ser** |
| **9** No observable | No es pot examinar |

**El 3 és l'únic valor que és una inferència.** Exigeix evidència **de naturalesa distinta de la cosa perduda**: un encaix buit, una empremta, una cicatriu de despreniment, un forat de biga. «Sembla que hi devia haver alguna cosa» no és evidència.

Cada 3 necessita fila a **12.Extra → elements desapareguts**, amb el tipus tret del catàleg tancat. **Si cap tipus del catàleg no encaixa, el valor correcte no és 3: és 0 o 9.**

### Cas freqüent: element asimètric

Un brancal de pedra laminar i l'altre de maçoneria corrent. ¿Disseny asimètric o brancal perdut?

Sense evidència de pèrdua **no ho pots decidir**, i l'escala no té valor per a «asimètric per disseny». Escriu:

- Valor **2**
- A `Notes sistemes`: `asymmetric: left only` (o `right only`), **exactament així**

*La fórmula constant importa: si el patró es repeteix, una consulta podrà trobar tots els casos i decidirem si mereix camp propi.*

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
| **Blocs irregulars** | La peça no té dues cares planes subparal·leles. Assenta com puga i la junta l'ajusta el ripio |
| **Blocs tabulars** | Dues cares planes **i alçada de filada pròpia**: una pedra, una filada |
| **Lloses laminars** | Dues cares planes, però **cal apilar-ne diverses per fer una filada** |

> **El test és: quantes peces fan una filada.** Es llig directament a la façana i no cal mesurar res.
> Si dubtes entre tabular i laminar: gruix ≤ ¼ de la dimensió major → laminar.

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

### Geometria del traç

`Ortogonal` / `Corb` / `Irregular` / `Indeterminada`. **Emplena'l sempre que hi haja una banda que ressegueix un contorn.**

> **Ortogonal** = angles rectes, traç net. Dibuixa un rectangle.
> **Corb** = línies corbes, contorn menys definit.
> **Irregular** = ni una cosa ni l'altra; taques sense forma reconeixible.

*Sembla un detall estètic i no ho és: **és el que decidirà si allí hi havia una estructura** (nus 6). Una banda que fa angles rectes està dibuixant una planta rectangular; una de corba pot estar seguint un rebaix natural de la penya.*

---

## Nus 9 — La fàbrica no és igual per tot arreu

**A 2.Arq, i només si l'estructura té dos cossos o més** (si en té un, es bloqueja sol: no hi ha res a comparar).

Quan el cos basal està fet de pedres irregulars i la cambra de blocs tabulars ben escairats, els camps de maçoneria només et deixen dir `Mixt` — i eixe `Mixt` **esborra justament el que interessa**: que hi ha dues fàbriques, i per tant potser dues mans, dues fases o dues intencions.

**Fàbrica:**

| Valor | Quan |
| --- | --- |
| `Única` | Una sola fàbrica a tot el registre |
| `Múltiple, coincident amb els cossos` | Cada cos és uniforme per dins, però difereixen entre ells |
| `Múltiple, dins d'un cos` | Almenys un cos té més d'una fàbrica per dins |
| `No observable` | No s'hi veu prou |

**La pregunta no és on hi ha el canvi, sinó si el canvi segueix els cossos o els travessa.** Un cos que ja té dues fàbriques per dins també serà distint del seu veí; si preguntàrem «on», les respostes se solaparien.

*I la distinció importa: el segon valor és exactament el cas que una descomposició per cossos resoldria; el tercer és el que no resoldria.*

### El que canvia exactament, a les notes

Si el que varia és el format, el treball, l'aparell o el morter, **escriu-ho a `Notes` d'arquitectura**. La consulta de repàs de notes ho recuperarà tot junt.

### No confongues divergència amb fase

> **Fàbrica múltiple = el que veus.** La pedra canvia.
> **Fase constructiva = el que argumentes.** Hi va haver dos moments de construcció.

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
- [ ] Cada element a **3** té la seua fila a 12.Extra
- [ ] `Doc_Basis` i `Facade_Observability` emplenats (10.Doc)
- [ ] `Pla d'accés` i `Orientació portal` emplenats (2.Arq)
- [ ] Cada fila rupestre té els quatre trams i la geometria del traç
- [ ] Si hi ha dos cossos o més, `Fàbrica` emplenada
- [ ] Si has declarat dues fases, `Evidència fase` emplenada
- [ ] Cap 0 que no puges justificar

## Les dues consultes

**`QRY_16_Validation_Check`** — 46 regles de coherència. **Resultat buit vol dir corpus coherent.**

No és una llista d'errors: la **regla 17** hi apareix a propòsit per a llistar els camps encara buits. És la teua llista de feina, no una queixa.

**`QRY_19_V16_Review`** — la feina que va deixar oberta la transferència a v16 i que continua vigent: portals a rejudicar, cobertes pendents del test de contacte, format de pedra per entrar (encara falta en 32 registres), i files de decoració que podrien haver de canviar de posició.

**`QRY_21_V17_Review`** — la que ha deixat oberta la transferència a v17: files rupestres amb els trams per omplir, files de decoració **sense cap posició**, connexions per rellegir, estructures de dos cossos o més sense divergència declarada, i fases sense evidència.

**`QRY_20_RockArt_Span`** — no és de revisió sinó de lectura: totes les files rupestres amb la forma del contorn ja calculada.

---

# **6. El vocabulari A–X d'una ullada**

*De baix a dalt. Els tres sistemes van marcats.*

## Nivell 0 — cos basal

| | | |
| --- | --- | --- |
| **A** | Jàsseres basals | Bigues encastades a la base |
| **B** | Basament | Massa basal diferenciada |
| **C** | Sòcol decoratiu | Tractament plàstic del basament |
| **D** | Muret transversal | Mur de trava |
| **E** | Mènsules | *No es bloqueja mai pel sistema* |
| **F** | Bigues transversals | |
| **G** | Filades en voladís | ⚠ vegeu nus 3 |
| **H** | **Plataforma** | **SISTEMA** = E + F + G |

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
| **V** | Mur lateral de cambra | Perpendicular, gira cap al penyal |
| **W** | Fons de cambra | |

## Zona superior

| | | |
| --- | --- | --- |
| **S** | Biga de suport del ràfec | Pot ser fusta |
| **T** | Superfície del ràfec | **Sempre pedra** |
| **U** | **Ràfec-voladís** | **SISTEMA** = S + T. **Sempre pedra** |
| **X** | Tancament superior de cambra | ⚠ vegeu nus 5 |

---

*Dubtes sobre criteris no coberts ací: `esquema_bbdd_estructures_v16.md`, secció 8bis.*
