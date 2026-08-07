**DELTA v15 → v16**

*Base de dades arqueològica — La Petaca i Diablo Wasi (PALP)*

Esteve Ribera Torró | TFM Arqueologia UA

**Document viu.** S'hi anoten les decisions d'actualització a mesura que es prenen, amb la justificació que les sosté i les passes d'implementació que impliquen. Quan el conjunt estiga tancat, aquest document és l'especificació des de la qual es generen els scripts v16 i s'actualitzen `esquema_bbdd_estructures` i `tfm_metodologia_bbdd`.

**Estat:** DECIDIT, pendent d'implementació (agost 2026).

**Origen:** revisió de la interfície v15 amb les fitxes de camp a la mà. Els quatre punts revisats no procedeixen de cap auditoria d'esquema sinó de **casos reals que la interfície no permetia registrar o registrava malament**. Això explica el perfil dels canvis: molt de criteri operatiu, poc de maquinària nova.

**Abast deliberadament limitat.** El plantejament de descomposició per subnivells (N0, N1.1, N1.2 amb intranivells derivats) **s'ajorna al delta v16→v17**: no és una qüestió de vocabulari sinó de cardinalitat —implicaria una taula `T_BODIES` amb els elements A–X penjant de cada cos— i la dada que ha de decidir-lo (quants registres tenen cossos amb atributs realment divergents) encara no existeix. Vegeu la secció 12.

**Convenció d'estat de cada punt:**
`[DECIDIT]` acordat, pendent d'implementar · `[OBERT]` en discussió · `[PENDENT DADES]` depén d'una comprovació sobre el corpus · `[FET]` implementat i verificat

---

# **0. Resum**

**Revisió 1** — format i treball de la pedra (secció 1).

| # | Punt | Estat | Abast |
| --- | --- | --- | --- |
| 1.1 | `Stone_Format` + `Stone_Format_Secondary` | `[DECIDIT]` | Esquema + formulari |
| 1.2 | `Stone_Working` com a escala ordinal | `[DECIDIT]` | Esquema + formulari |
| 1.3 | `Chinking_Stones` es manté intacte, fora de la llista nova | `[DECIDIT]` | Documentació |
| 1.4 | Criteri d'evidència positiva de treball | `[DECIDIT]` | Documentació |
| 1.5 | Sense valor «no observable»: `ND` i encreuament amb `Doc_Basis` | `[DECIDIT]` | Documentació |
| 1.6 | `Dressed` es manté al domini | `[DECIDIT]` | Cap canvi; comprovar variància |

**Revisió 2** — posicions de decoració i pla de façana (seccions 2 i 3).

| # | Punt | Estat | Abast |
| --- | --- | --- | --- |
| 2.1 | Retirada de `SOC`; `BAS` → *Cos basal (B/C)* | `[DECIDIT]` | Lookup + remapatge |
| 2.2 | `FFL` → *Flanc de façana (L)*, redefinit pel pla exposat | `[DECIDIT]` | Lookup + esquema |
| 2.3 | V → *Mur lateral de cambra*, *de retorn* com a glossa | `[DECIDIT]` | Lookup + formulari |
| 2.4 | W → *Fons de cambra* | `[DECIDIT]` | Formulari |
| 2.5 | Quatre posicions noves: `RTW`, `PRT`, `SIL`, `LIN` | `[DECIDIT]` | Lookup |
| 2.6 | *Llinda* proscrita a la interfície; Q sempre *Dintell* | `[DECIDIT]` | Convenció |
| 2.7 | Absència de lloses de tancament d'obertura | `[DECIDIT]` | Resultat per al TFM |
| 3.1 | Façana = pla exposat, no pla de l'obertura | `[DECIDIT]` | Esquema |
| 3.2 | `Access_Plane` | `[DECIDIT]` | Esquema + formulari |
| 3.3 | `Portal_Orientation` | `[DECIDIT]` | Esquema + formulari |
| 3.4 | Banqueta d'accés → `T_ARCH_FEATURES` (`Abutted mass`) | `[DECIDIT]` | Convenció d'entrada |

**Revisió 3** — pintura rupestre (secció 4).

| # | Punt | Estat | Abast |
| --- | --- | --- | --- |
| 4.1 | `Position_Relative`, compartit per les dues meitats de 4.Dec | `[DECIDIT]` | Esquema + formulari |
| 4.2 | Lateralitat esquerra/dreta resolta pel mateix camp | `[DECIDIT]` | Documentació |
| 4.3 | `Body_No` restringit a les files no-ROC | `[DECIDIT]` | Convenció + regla |
| 4.4 | Posicions ROC redefinides per contacte geomètric | `[DECIDIT]` | Lookup + esquema |
| 4.5 | `ROC-OVL` (solapada): entrada nova | `[DECIDIT]` | Lookup |
| 4.6 | `Substrate` restringit a `Bedrock` / `Mixed` a la meitat ROC | `[DECIDIT]` | Formulari + regles |
| 4.7 | Vocabulari ROC mínim: cap posició més | `[DECIDIT]` | Cap canvi |

**Revisió 4** — sistemes constructius i criteris de discriminació (secció 5).

| # | Punt | Estat | Abast |
| --- | --- | --- | --- |
| 5.1 | Retirada de `Interbody_Cornice_Material`; I és pedra per definició | `[DECIDIT]` | Esquema + formulari + gating |
| 5.2 | X = tancament efectiu; sense contacte, X = 0 | `[DECIDIT]` | Esquema |
| 5.3 | `Natural bedrock` computa NA a la matriu A–X | `[DECIDIT]` | QRY_13 |
| 5.4 | Criteri de diferenciació generalitzat a tot el vocabulari | `[DECIDIT]` | Esquema + revisió manual |
| 5.5 | `Sys_Portal` vs N/O/Q: divisió del treball | `[DECIDIT]` | Documentació |
| 5.6 | Portal asimètric: valor 2 + nota amb fórmula constant | `[DECIDIT]` | Convenció |
| 5.7 | H és N0, no intranivell | `[DECIDIT]` | Documentació |
| 5.8 | Quatre tests G vs I i sortida per 9 + `T_ARCH_FEATURES` | `[DECIDIT]` | Documentació + regla |

---

# **1. Format i treball de la pedra** `[DECIDIT]`

## 1.1. Problema

`Masonry_Type` registra **com s'apila** la pedra (regularitat de les filades) i `Masonry_Quality` un judici agregat d'execució. Cap dels dos registra **quina pedra és**, que és el primer gest de la seqüència operativa: la selecció i preparació de la matèria primera. És material directe de H04 i el descriptor amb més capacitat de discriminar tradicions entre LP i DW (H05).

## 1.2. Per què dos camps i no un

El vocabulari inicialment proposat —*pedres irregulars, laminars, blocs semiescairats, carreus, lloses de gran format, pedra de rebliment*— barrejava **tres eixos**: morfologia natural (irregular, laminar), grau de treball (semiescairat, carreu) i dimensió (gran format). En una llista única, un mur de lloses laminars ben escairades no té valor possible, i el registre es decideix per hàbit de l'observador: precisament el que trenca la replicabilitat.

Es desdobla en dos camps ortogonals. Els termes proposats es reconstrueixen tots com a combinacions, i se n'expressen de noves que la llista única no permetia.

## 1.3. `Stone_Format` — morfologia de la peça

Descriu el producte de la fractura natural de la roca. **Abast: el parament de maçoneria**, no els elements singulars — el gran format viu al dintell, la superfície de ràfec, el tancament posterior i la coberta, que ja tenen camp de material propi. Per això *Large-format slabs* es descarta com a valor.

| Valor | Test |
| --- | --- |
| Irregular blocks | La peça no té dues cares planes subparal·leles dominants. Assenta com puga i la junta l'ajusta el ripio |
| Tabular blocks | Dues cares planes subparal·leles **i alçada de filada pròpia**: una pedra, una filada |
| Laminar slabs | Dues cares planes, **però cal apilar-ne diverses per a fer una filada**: el gruix és clarament subordinat |
| ND | No determinat |

*Test primari deliberadament funcional —**quantes peces fan una filada**— perquè es llig directament a la façana i no exigeix mesurar res. Desempat quantitatiu per als casos dubtosos: gruix ≤ ¼ de la dimensió major → laminar.*

*`Laminar slabs` es manté al parament encara que el mateix format aparega en frisos i relleus: a les filades en voladís (G) i a la cornisa intercòs (I) fa treball estructural. El que va a `T_DECORATIONS` és el motiu; el que va ací és de què està fet el mur que el porta.*

**Parella ordenada.** Els paraments amb dos formats convivint es registren amb `Stone_Format` + `Stone_Format_Secondary`, seguint el patró de `ID_Support` + `ID_Support_Secondary` i de `Color` + `Color_Secondary`. **`Mixed` no existeix com a valor**: seria un tercer camí per a dir el mateix, i el projecte ja ha retirat dues vegades el valor que la parella substitueix (`Combined` de `L_SUPPORT` en v8, `Both` de `Color` en v13). El camp secundari no porta `ND` (buit ja vol dir «cap segon format»).

**Criteri de dominància:** el format dominant és el de **major superfície de parament**, no el de major nombre de peces. Amb lloses laminars menudes i blocs tabulars grans, comptar peces inverteix el resultat.

## 1.4. `Stone_Working` — grau de treball

*Unworked / Semi-dressed / Dressed / Mixed / ND*

**Escala ordinal** `Unworked < Semi-dressed < Dressed`, utilitzable com a proxy d'inversió de treball en correspondències. **`Mixed` i `ND` són valors fora d'escala** i s'exclouen d'eixes anàlisis, igual que els `Not applicable`.

**Criteri d'evidència positiva** (indispensable, perquè la fractura natural d'un gres estratificat produeix cares indistingibles del carejat):

> `Stone_Working` registra **evidència positiva de treball**: traces d'eina, aristes vives regulars, cares que trenquen el pla de fractura natural de la roca. En absència d'eixes traces, la cara plana s'atribueix a la fractura i el valor és `Unworked`. Quan les cares planes són compatibles amb ambdues coses i no hi ha traces diagnòstiques, el valor és `ND`.

Amb això `Unworked` deixa de ser una afirmació arriscada —no diu «ningú no la va tocar» sinó «no hi ha evidència que la tocaren»— i `ND` queda reservat per al dubte genuí.

**Lectura:** `Stone_Working` qualifica el **format dominant**. Si el secundari té un treball notòriament distint (cas real: blocs tabulars sense treballar amb brancals semiescairats), això va a `Arch_Notes`.

## 1.5. Sense valor «no observable» `[DECIDIT]`

Les dues llistes tanquen amb `ND` i **no porten un 9 a banda**. El 9 existeix per a protegir el 0: separa «he mirat i no hi és» de «no he pogut mirar». Aquests camps **no tenen zero** —un parament sempre està fet d'alguna pedra amb algun grau de treball—, de manera que el 9 no fa cap treball. Cap camp de domini TEXT de l'esquema en porta.

Els tres estats distingibles es mantenen igualment: **NULL** = encara no valorat · **ND** = valorat i no determinable · valor = valorat.

**Els dos motius de no-decisió no es col·lapsen**, tot i compartir `ND`:

| Situació | Es recupera per |
| --- | --- |
| No es veu prou (distància, revoc, vegetació) | `ND` + `Doc_Basis` / `Facade_Observability` |
| Es veu bé però **no és decidible** (la fractura imita el carejat) | `ND` + `Doc_Basis = Direct access` i façana observable |

La informació que els separa **ja està registrada en un altre camp** i la consulta els distingeix amb un filtre. És el mateix criteri que va deixar `Mortar_Present` sense promocionar: el que varia amb la conservació és la visibilitat, i eixa ja té camps propis.

**Cas informatiu a retindre:** parament amb revoc conservat a tota la façana → `Plaster_Present = 1` + `Plaster_Extent = Full facade` + `Stone_Format = ND`. No és una llacuna: diu que la fàbrica no estava pensada per a veure's.

## 1.6. Variància de `Dressed` `[DECIDIT]`

**El valor es manté al domini.** Pot resultar quasi buit, com va passar amb `Chinking_Stones`, i no és motiu per a llevar-lo: si surt zero, l'absència de carreu escairat a tot DW *és un resultat* sobre la tradició constructiva, i el valor ha d'existir perquè eixa afirmació siga defensable. Comprovar-ho al diagnòstic previ abans de donar el camp per bo analíticament.

---

# **2. Posicions de decoració** `[DECIDIT]`

## 2.1. Retirada de `SOC`

C (`Decorative_Socle`) es defineix com el **tractament decoratiu** del basament. Per tant «una decoració situada al sòcol» significa «una decoració situada al tractament decoratiu del basament»: **la posició pressuposa el que la fila hauria de documentar**. I l'operació inversa tampoc no funciona — registrar un motiu al basament ja el converteix en sòcol.

`SOC` es retira i `BAS` passa a **`Cos basal (B/C)` / `Basal body (B/C)`**. L'element C sobreviu intacte com a camp a 11.Sist; el que desapareix és la posició duplicada.

**Es descarta l'alternativa** d'usar `SOC` per als relleus i `BAS` per a les pintures: faria que el camp Posició codificara **el tipus de motiu**, que és el que registra `ID_Dec_Type`. Dos eixos barrejats, i un fris pintat en relleu quedaria sense valor possible.

**Frontera fixada per escrit:**

> **C = tractament plàstic**: filada volada, motllura, canvi d'aparell o de format de pedra que individualitza la massa basal.
> **Fila a `T_DECORATIONS` en posició `BAS` = motiu aplicat** (pigment, ninxol, relleu) sobre el cos basal.

Un basament amb banda pintada i sense motllura és `Decorative_Socle = 0` + una fila `BAS`. Amb la definició v15 això no es podia decidir.

## 2.2. Renomenaments de terminologia

| Element | v15 (VAL) | v16 (VAL) | Motiu |
| --- | --- | --- | --- |
| L | Ala de façana | **Flanc de façana** | *Ala* és calc de *flank* i no diu res en camp. *Flanc* és terme arquitectònic correcte i transparent |
| V | Mur de retorn | **Mur lateral de cambra** | *Retorn* descriu l'operació geomètrica (el mur gira), no l'element. «Lateral» ha quedat lliure en passar L a *flanc* |
| W | Tancament posterior | **Fons de cambra** | En arqueologia funerària *tancament* s'entén com el segellat de la tomba: risc d'error silenciós de registre |

**No es pot usar «lateral» per a L.** La v14 va renombrar `Lateral_Walls` → `Return_Wall` i `Lateral_Wall_Faces` → `Facade_Flank` precisament perquè els dos noms deien «lateral» de coses distintes. Reintroduir-ho desfaria la correcció.

*«Mur de retorn» és terme català correcte i es conserva com a **glossa tècnica** a la descripció del lookup i a l'esquema —«mur lateral de cambra, en retorn perpendicular al pla de façana»—, que és on serveix per a la publicació.*

Cap nom de camp anglés canvia: `Facade_Flank`, `Return_Wall` i `Rear_Closure_Type` es mantenen.

**Criteri de distinció L / V, fixat:**

| | Pla | Gira? |
| --- | --- | --- |
| **L** Flanc de façana | Dins del pla de façana | No |
| **V** Mur lateral de cambra | Perpendicular | Sí, cap al penyal |

## 2.3. Posicions noves

La llista v15 tenia `Brancal (O)` i `Sobre-dintell`, però **no llindar (N) ni dintell (Q)**. I `Sobre-dintell` no és el dintell: és la zona per damunt, l'àrea de fris (element M). Un portal pintat s'havia de registrar amb un element que no era el que estava pintat.

| Codi | VAL | EN | Level | Sort |
| --- | --- | --- | --- | --- |
| `RTW` | Mur lateral de cambra (V) | Return wall (V) | N1 | 62 |
| `PRT` | Sistema portal (P) | Portal system (P) | N1 | 64 |
| `SIL` | Llindar (N) | Sill (N) | N1 | 65 |
| `LIN` | Dintell (Q) | Lintel (Q) | N1 | 75 |

**`PRT` respon la pregunta que ve darrere de la pregunta:** una banda contínua que emmarca tota l'obertura **és un sol gest, no tres**. Partir-la en tres files inventaria tres motius on n'hi ha un — el mateix error que va justificar `Color_Secondary`. Precedent estructural: el vocabulari ja tracta N+O+Q com un **sistema** perquè els components actuen junts, i `Pigment_Extent` ja porta *Perimeter/threshold* amb la mateixa idea a escala d'estructura.

**Criteri d'ús:**

> Posició elemental (`SIL` / `JAM` / `LIN`) quan el motiu **s'interromp** entre elements o en decora un de sol. Posició `PRT` quan el tractament és **continu al voltant de l'obertura** i no s'atribueix a cap component. **En cas de dubte, elemental**: sempre es pot agregar després, mai desagregar.

**`RTW` respon un cas documentat:** hi ha estructures amb decoració al mur lateral **i no a la façana**. Això no és un detall menor — és decoració no dirigida a l'espectador de la vall, i toca H06 de front. Sense l'entrada, eixos casos es perdien o es registraven com a façana, que és pitjor.

**Descartades:** posició de plataforma (cap cas de pigment documentat).

## 2.4. Convenció terminològica: *llinda* proscrita

**Llindar** = peça inferior del marc (element **N**, *sill*). **Llinda** o **dintell** = peça horitzontal superior (element **Q**, *lintel*). Són falsos amics quasi homògrafs i font directa d'error de registre.

**A la interfície s'usa sempre `Dintell` per a Q i mai `llinda`.** Cal que quede escrit a l'esquema.

## 2.5. Lloses de tancament d'obertura: absència verificada

No se n'ha documentat cap al corpus. **Es registra com a resultat, no com a silenci:** si a DW/LP no hi ha segellat de l'obertura, això diu alguna cosa sobre l'accés recurrent a la cambra i sobre la relació entre H i P. Mateix criteri aplicat a `Chinking_Stones`.

## 2.6. Llista `L_STRUCT_BODY` resultant

18 entrades (14 − 1 + 5, comptant `ROC-OVL` de la secció 4):

`Cos basal (B/C)` 10 · `Cornisa intercòs (I)` 30 · `Cantonera (J)` 40 · `Pilastra (K)` 50 · `Flanc de façana (L)` 60 · **`Mur lateral de cambra (V)` 62** · **`Sistema portal (P)` 64** · **`Llindar (N)` 65** · `Brancal (O)` 70 · **`Dintell (Q)` 75** · `Sobre-dintell (M)` 80 · `Coronament (R)` 90 · `Ràfec (U)` 100 · **`Solapada` 200** · `Perimetral` 210 · `Pròxima` 220 · `Panell` 230 · `ND` 999

---

# **3. Façana, pla d'accés i orientació del portal** `[DECIDIT]`

## 3.1. El problema: «façana» feia dues faenes

A l'esquema v15 el mot significava alhora el **pla exposat** (el que mira a la vall: `Facade_Orientation`, `Visibility_Valley`, tota H06) i el **pla compositiu** (el que porta l'obertura: definició de L, N, O, Q, M). Normalment coincideixen. En un cas documentat **divergeixen**: mausoleu amb l'obertura al mur estret perpendicular al farallò i la decoració principal al mur llarg paral·lel.

## 3.2. Decisió: la façana és el pla exposat

Fixar-la per l'obertura tindria un cost desproporcionat — `Facade_Orientation` deixaria d'apuntar cap a la vall i `Visibility_Valley` mesuraria un pla que no es veu: H06 sencera es degradaria per a salvar la definició d'un element. Fixar-la pel pla exposat costa **una línia de definició**:

> **L (Flanc de façana)** = parament de maçoneria dins del pla de façana. Quan l'obertura és al pla de façana, L són els paraments que la flanquegen; quan no, L és el parament corregut.

L deixa de dependre de l'obertura. `RTW`, `SIL`, `JAM`, `LIN` i `PRT` continuen sent posicions vàlides visquen al pla que visquen: **una posició és una posició**.

En el cas documentat: mur llarg paral·lel al farallò = **façana** (L, decorada, orientada a la vall); mur estret perpendicular = **mur lateral de cambra (V)**, que és qui porta el portal.

## 3.3. `Access_Plane` — camp nou

```
Access_Plane  TEXT(20)  Facade / Return wall / Rear / ND
```

VAL: *Pla d'accés* — `Façana` / `Mur lateral` / `Fons` / `Indeterminat`. A 2.Arq, al costat de `Facade_Orientation`. Sense gating, tres valors.

**La divergència no és una anomalia a absorbir: és una variable.** Que l'accés no estiga al pla exposat diu que la circulació mana sobre l'exhibició — s'entra per on es camina, per la repisa, i s'exhibeix cap a on es mira. És H06 i H01 alhora, i en v15 no quedava enlloc.

**Efecte col·lateral necessari:** declara el **pla de referència** de la lateralitat esquerra/dreta (secció 4.2), que sense ell no significa res.

## 3.4. `Portal_Orientation` — camp nou

```
Portal_Orientation  TEXT(5)  N / NE / E / SE / S / SW / W / NW / ND
```

VAL: *Orientació portal*. A 2.Arq, immediatament davall de `Facade_Orientation`.

**No duplica `Access_Plane`:** el pla d'accés és **relacional i intrínsec** (com l'arquitectura organitza exhibició i accés, comparable entre orientacions absolutes distintes); l'orientació del portal és **absoluta** (entra a QGIS, es creua amb topografia i rutes de circulació). Cap de les dues es dedueix de l'altra sense conéixer la geometria del cas.

**Sense gating.** És temptador penjar-lo de `Sys_Portal`, però es va rebutjar per al mateix cas amb `Recessed_Frame` i el motiu val igual: una obertura arrasada té orientació coneguda encara que el sistema no siga *Present*. Un portal a *Attested lost* amb orientació registrada és precisament un cas informatiu.

**Guany analític:** la **divergència** entre `Facade_Orientation` i `Portal_Orientation` passa a ser calculable. Es podrà preguntar si les estructures amb accés desviat comparteixen tipologia, sector, tipus de suport o amplària de repisa. Amb un sol camp d'orientació no es podia formular.

## 3.5. Banqueta d'accés → `T_ARCH_FEATURES` `[DECIDIT]`

Element documentat i no descrit: massa construïda que **reposa sobre el suport natural adossada a l'estructura, sense sostindre cap cos** i sense continuïtat de fàbrica amb ella; la cara superior és una superfície practicable a cota intermèdia.

**No és B** (un basament sosté el que hi ha damunt), **no és H** (H vola sobre el buit, a tallant i moment; això reposa en compressió), **no és I** (I és una junta entre cossos, no un cos), **i no és registre independent** (no té coherència constructiva pròpia).

**Decisió: `T_ARCH_FEATURES`, no element nou.** No toca el vocabulari A–X ja publicat (Ribera-Torró et al. 2026) i acumula evidència.

**Convencions d'entrada, perquè la promoció futura siga possible sense revisar fitxes:**

- `Feature_Type` = **`Abutted mass`**, literalment idèntic a totes les files, o la consulta de recompte no les trobarà.
- Funció a `Notes` amb fórmula constant: `function: access` / `function: levelling` / `function: circulation`.
- **Llindar de revisió: 3 casos.** Si s'arriba, es planteja l'element Y (`Abutted_Mass` + `Abutted_Mass_Function`) al delta següent. És el criteri ja escrit a l'esquema: *si el mateix tipus d'observació apareix repetidament, cal un camp*.

*El nom «escaló d'accés» o «plataforma d'accés» es descarta: pressuposa la funció, que és exactament l'error que `Platform_Function` existeix per a evitar (OE3 existeix per a determinar per a què servia). «Plataforma» col·lisionaria a més amb H.*

---

# **4. Pintura rupestre** `[DECIDIT]`

**Advertiment previ.** El diagnòstic sobre el corpus v12 va donar **0 registres PR i 2 estructures amb pigment sobre penya** (EA11, EA39). Les decisions d'aquesta secció es prenen sobre dos casos, i per això el criteri rector és **no inflar el vocabulari**: afegir posicions que ningú no ha documentat és inventar categories.

## 4.1. `Position_Relative` — camp nou a `T_DECORATIONS`

```
Position_Relative  TEXT(20)  Left / Right / Both / Above / Below / ND
```

VAL: *Posició relativa* — `Esquerra` / `Dreta` / `Ambdós` / `Damunt` / `Davall` / `Indeterminada`.

**Compartit per les dues meitats de 4.Dec.** Un motiu al flanc esquerre, un brancal pintat dels dos, una pintura sobre la penya per damunt de l'estructura: el mateix mecanisme. Les dues meitats ja comparteixen taula.

**Convenció de referència, indispensable o el camp no és replicable:**

> Esquerra i dreta **des del punt de vista de l'observador situat davant del pla**, no des de l'estructura. El pla de referència és el que declara `Access_Plane`; quan la posició és de façana, el pla és la façana.

Sense això, dos observadors codifiquen invertit i el camp és soroll.

**`Both`** es conserva per al tractament bilateral simètric: la simetria és una afirmació arqueològica, i partir-la en dues files inventaria dos motius.

## 4.2. Lateralitat: resolta ací

La qüestió esquerra/dreta plantejada a la secció 2 (per a `FFL`, `JAM`, `PRT` i `RTW`) **no necessita cap mecanisme propi**: la resol `Position_Relative` amb la convenció de l'observador. Un camp, les dues meitats de la taula, i cap entrada de lookup duplicada per costat.

## 4.3. `Body_No` restringit a les files no-ROC

`Body_No` numera el cos constructiu. Una pintura sobre la penya **no està en cap cos**, de manera que a les files ROC el camp o queda buit sense criteri o porta un número que no significa res.

**No se substitueix per la posició relativa:** són dos eixos distints (índex vertical de cos vs posició relativa), i a les files arquitectòniques `Body_No` fa falta — i en farà més si prospera la descomposició per subnivells.

**Decisió:** `Body_No` es manté, **obligatòriament buit a les files ROC** (regla R37).

*El **criteri de numeració** de `Body_No` no es toca ací: és exactament el que el delta v16→v17 ha de fixar.*

## 4.4. Posicions ROC redefinides per contacte geomètric

**Problema v15:** la descripció al codi deia de `ROC` *«use when tests 2 or 3 apply»* i de `ROC-PER` *«test 3»*. **Els dos tests 3 xocaven**: qualsevol pintura que emmarcara una obertura complia la condició de les dues entrades i la tria quedava a l'atzar de l'observador.

**Solució: criteri purament geomètric**, resoluble mirant, sense consultar `ID_Support` ni raonar sobre accidents del farallò.

| Codi | VAL | EN | Criteri | Sort |
| --- | --- | --- | --- | --- |
| `ROC-OVL` | **Solapada** | Overlapping | Un mateix motiu té part sobre la fàbrica i part sobre la roca | 200 |
| `ROC-PER` | **Perimetral** | Perimeter | Ressegueix el contorn de l'estructura —**present o desapareguda**— en contacte tangencial | 210 |
| `ROC` | **Pròxima** | Proximate | No toca la fàbrica en cap punt, però és **dins del mateix accident del farallò** que allotja l'estructura | 220 |
| `ROC-PAN` | Panell | Panel | Files penjades d'un registre PR: no hi ha estructura de la qual dir res | 230 |

**Ordre d'aplicació: `ROC-OVL` → `ROC-PER` → `ROC`**, de més contacte a menys. La primera que es compleix decideix. Els codis `ROC` i `ROC-PER` es conserven (canvien de nom i definició, no d'identificador): la migració és mínima.

**Límit exterior de *Pròxima*, indispensable.** *No toca l'estructura* no acaba enlloc — tota pintura del farallò no toca l'estructura — i sense límit `ROC` s'empassaria el que ha de ser fitxa PR:

> **Pròxima** — el pigment no toca la fàbrica en cap punt, però es troba **dins del mateix accident del farallò que allotja l'estructura**: la mateixa repisa, cavitat, escletxa o rebaix. **Si cal travessar una vora de banc, una cornisa natural o un buit per a arribar-hi, no és pròxima: és fitxa PR.**

*El límit és no mètric: el delta v12→v15 ja va descartar el llindar d'un metre perquè en paret vertical la distància euclidiana no significa el mateix a tot arreu, i perquè l'associació es decideix en entrar el registre. Reutilitza `ID_Support`, una decisió ja presa.*

**`ROC-PAN` no es tria mai des d'un registre d'estructura.** Si apareix dins d'una estructura, alguna cosa s'ha encaminat malament.

**Clàusula «present o desapareguda» de la perimetral.** És el que fa funcionar la categoria: bandes verticals sobre roca nua alineades amb els brancals d'un cos que ja no hi és són `ROC-PER` sense contacte possible, perquè el contorn que ressegueixen és el que hi havia. Encreua amb `T_LOST_ELEMENTS` (`Pigment on bedrock`) i amb l'avís de la regla 30. **La definició diu *contorn*, no *contacte***, perquè el contacte és una observació de l'estat actual.

## 4.5. On va la fila solapada

`ROC-OVL` **contradiu el test 1 vigent** («si el pigment toca la fàbrica, és pigment de l'estructura»). Es resol així:

> **Una fila, posició `ROC-OVL`.** El costat fàbrica el registra 3.Acab: `Pigment_Present = 1` amb `Pigment_Substrate = Mixed`.
> **Test 1 reformulat:** el contacte **parcial** no basta per a fer una pintura arquitectònica; el que la fa arquitectònica és estar **sencera** sobre la fàbrica.

*Es descarta partir-la en dues files, una a cada meitat: inventaria dos motius on n'hi ha un — l'error que va justificar `Color_Secondary`.*

**`ROC-OVL` és la categoria analíticament més interessant de les quatre.** Un motiu que travessa la junta entre roca i maçoneria demostra que qui pintava tractava les dues superfícies com una sola, i que la pintura és posterior a la fàbrica en eixe punt. Toca H04 i H06. En v15 eixos casos es dissolien dins de `Pigment_Substrate = Mixed` sense deixar rastre com a motiu.

## 4.6. `Substrate` a la meitat ROC

**Objecció de partida, correcta:** *Revoc* i *Pedra de parament* són impossibles en una fila ROC per definició del criteri d'associació. Una fila ROC amb substrat de parament és una fila que hauria d'haver anat a l'altra meitat.

**No es lleva el camp: es converteix en verificador.** Al subformulari de pintura rupestre la llista es restringeix a **`Bedrock` / `Mixed`** (`Mixed` només per a `ROC-OVL`). A la meitat arquitectònica, `Plaster / Masonry stone / Bedrock / ND` es mantenen intactes.

Una incoherència abans silenciosa passa a ser un error detectat (regles R38a i R38b).

*Es va valorar substituir-lo per una caracterització de la superfície rocosa (fractura fresca / pàtina / alteració / abric) i **es descarta**: seria un camp nou sobre dos casos. Es replanteja si el corpus creix.*

## 4.7. Cap posició ROC més `[DECIDIT]`

Amb `Position_Relative` operatiu, els casos que semblaven demanar entrada pròpia —pintura per damunt, al costat, davall— es diuen amb `ROC` + `Damunt` / `Esquerra` / `Davall`. **Quatre entrades i un camp cobreixen més combinacions que sis entrades**, i no s'inventa res sobre dos casos.

*Nota: `ROC-PER` (posició) i `RA Perimeter band` (tipus) **no són el mateix camp** i no s'han de pressuposar junts. El tipus diu **què sembla** (una banda contínua), la posició diu **on està**. Una banda perimetral pot ser pròxima; un motiu antropomorf pot ser perimetral.*

---

# **5. Sistemes constructius i criteris de discriminació** `[DECIDIT]`

## 5.1. Retirada de `Interbody_Cornice_Material`

**No és que de fet siguen sempre de pedra: és que no podrien no ser-ho.** Un element horitzontal de fusta que sobreïx entre dos cossos no és una cornisa — és una mènsula (E, perpendicular, encastada, en voladís) o una biga transversal (F, paral·lela, salvant llum), i els dos ja tenen camp i són components de la plataforma. El valor *Wooden beams* no descrivia una cornisa de fusta: **descrivia una plataforma mal classificada**.

**Comprovació sobre el corpus: totes les cornises registrades són de pedra, cap de fusta ni mixta.** La retirada és neta i no exigeix revisió prèvia.

**Definició de I actualitzada:**

> **I · Cornisa intercòs** — filada volada **de pedra** que marca la junta entre dos cossos superposats. **Si l'element volat és de fusta, no és una cornisa**: és E o F, i el que hi ha és plataforma.

*Precedent exacte: T i U ja porten «SEMPRE PEDRA» dins de la definició del vocabulari i **no tenen camp de material**. La cornisa era l'excepció incoherent.*

**El guany no és estalviar un camp:** una descripció feble es converteix en un **test d'identificació**, i eixe test és el que fa falta a la secció 5.8. Si algun cas rar apareix, `Systems_Notes` el recull i el criteri es reobri amb evidència.

**Efectes:** un camp menys; un nivell de gating menys (nivell 3, «material de la cornisa només amb I present»); la regla del patró presència+tipus corresponent desapareix.

## 5.2. X redefinit: tancament efectiu

**Problema.** X, definit com *«cambra tancada per dalt, amb la solució que siga»*, **no és un element sinó un estat**. Amb `Chamber_Roof_Type = Natural bedrock` no s'ha construït res, però `Chamber_Roof = 1` entrava a la matriu A–X igual que un dintell.

**El corpus ho agreuja:** hi ha **molts** sostres de penya natural. Amb això, la coocurrència de Jaccard i l'anàlisi de correspondències comptaven geologia com si fora construcció, inflant la similitud entre estructures que no comparteixen cap decisió constructiva. Mateixa família d'error que `ID_Material_Status` i que `Lintel` en v10: dues variables dins d'una escala.

**Tres casos que el corpus presenta i que v15 no distingia:**

| Cas | Què va fer el constructor | Espai funerari |
| --- | --- | --- |
| Cambra sota visera ajustada | Aprofitar la roca com a tancament | Tancat |
| Bigues i lloses a l'entrada, roca al fons | Construir **on la roca no arribava** | Tancat, en dues operacions |
| Mausoleu amb la visera a metres de distància | **Res**: la roca no tanca res | **Obert per dalt** |

**Definició v16:**

> **X = 1** quan l'espai funerari està tancat per dalt, per obra o per roca **en contacte amb la fàbrica**.
> **X = 0** quan no ho està, **inclosa la roca que passa per damunt sense contacte**. Que hi haja abric és `ID_Support`, no X.
>
> **Test: si no hi ha contacte, no hi ha sostre.** Binari, sense gradient: 20 cm de buit és buit, igual que 3 m.

**`Chamber_Roof_Type = Mixed` precisat:** obra i roca natural **repartint-se el tancament**, típicament obra a l'entrada i roca al fons. És el segon cas de la taula i és freqüent.

**Cas que la definició ha de resoldre explícitament:** en cossos cambra superposats, el sostre de la cambra inferior **és** el paviment de la superior. Compta com a X de la inferior, amb el tipus que corresponga a l'obra. *Enllaça amb el delta següent: amb subnivells, X passa a ser per cos.*

**Mausoleu obert per dalt sota abric separat:** `Chamber_Roof = 0` + descripció a `Systems_Notes`. **No es crea camp** (`Overhead_Shelter` es va valorar i es descarta): són pocs casos.

*`Rear_Closure_Type` té la mateixa estructura i el mateix tractament: `Natural bedrock` no és fàbrica.*

## 5.3. `Natural bedrock` computa NA a la matriu A–X

| `Chamber_Roof_Type` | Computa a QRY_13 |
| --- | --- |
| `Built masonry`, `Built timber and slabs` | **1** |
| `Mixed` | **1** — hi ha obra |
| `Natural bedrock` | **NA** — tancat sense construir |
| X = 0 | **0** — absència real |

**NA i no 0:** no és una absència, és una solució no construïda. Un 0 seria tan fals com un 1. `Not applicable` no entra en cap denominador, que és exactament el comportament que cal.

> **Advertiment d'implementació.** Amb molts sostres de penya natural al corpus, **el filtre canvia els resultats de la matriu A–X**. Qualsevol prova de Jaccard, clúster o correspondències ja executada s'ha de refer. Comprovar quants registres canvien de valor abans de donar per bona cap anàlisi anterior.

## 5.4. Criteri de diferenciació, generalitzat a tot el vocabulari

Aquesta és la decisió de més abast del delta. Neix d'una pregunta sobre el portal —quan es diu que hi ha llindar i brancals?— i la resposta fixa el criteri de **tots** els elements A–X.

> Un element A–X **és present** quan hi ha un component **físicament diferenciat** del parament contigu — per format de pedra, dimensió, material o tractament. **Si la vora de l'obertura és maçoneria contínua amb el mur, sense cap peça distinta, l'element és absent (0).**

**El criteri no és nou:** B ja el portava implícitament (*basament com a element constructiu diferenciat*). Ací es fa general i explícit.

**El 0 resultant és un resultat de primera magnitud.** Que a DW l'obertura es resolga sovint sense llindar ni brancals diferenciats és una afirmació sobre la inversió de treball i sobre la seqüència operativa — H01 i H04 directament. Amb el criteri contrari («si hi ha obertura, hi ha brancals») la variable desapareixeria i tots els portals serien iguals.

**Complet vs parcial:**

- **1 · Complet** — element diferenciat present en tota la seua extensió esperada: els dos brancals, el llindar sencer
- **2 · Parcial** — present però incomplet, amb l'extensió original inferible
- **3 · Desaparegut atestat** — evidència **de naturalesa distinta de la cosa perduda**: encaix buit, empremta, cicatriu

## 5.5. `Sys_Portal` vs N/O/Q: divisió del treball

Una obertura resolta deixant simplement un buit al mur és `Sill = 0`, `Jambs = 0`, `Lintel = 0` — **i `Sys_Portal` = *Present complete* igualment**, perquè l'obertura hi és.

| | Registra |
| --- | --- |
| `Sys_Portal` | Que **existeix obertura d'accés** |
| N, O, Q | Si **cada posició es va resoldre amb un element diferenciat** |

Coherent amb el que ja hi ha: R11 fa incompatible tindre cossos cambra amb `Sys_Portal = Absent` (un cos és N1 si conté o contenia obertura), i `Opening_Width_m` / `Opening_Height_m` deliberadament **no** estan gatejades pel sistema.

## 5.6. Portal asimètric `[DECIDIT]`

Un portal amb un brancal de pedra laminar i l'altre de maçoneria corrent: ¿disseny asimètric o brancal perdut? Sense evidència de pèrdua no es pot decidir, i el domini no té valor per a «asimètric per disseny».

**Comprovació sobre el corpus: menys de tres casos.** No es crea camp.

**Convenció:** valor **2** (parcial) + nota amb **fórmula constant** a `Systems_Notes`: `asymmetric: left only` / `asymmetric: right only`. Si el patró resulta recurrent, això serà el senyal que cal un camp — el criteri ja aplicat a la hipòtesi de politja i a la relació cromàtica.

## 5.7. H és N0, no intranivell

| Element | Nivell | Per què |
| --- | --- | --- |
| **H** Plataforma | **N0** | Culminació del cos basal. Es construeix des de baix (E+F+G) i pot existir sense res damunt |
| **I** Cornisa | **N0/N1** | És la **junta** entre dos cossos. Exigeix que n'hi haja dos |

Ocupen la mateixa franja física —tots dos a la transició—, i per això es confonen. La diferència és categorial:

> **La plataforma és una superfície.** Horitzontal, practicable, amb profunditat útil, i té sistema propi.
> **La cornisa és una junta.** Lineal, sense profunditat útil, no practicable, i no té sistema deliberadament.

Recordatori de la decisió ja tancada al delta anterior: **H no compta com a cos basal**. Una plataforma no és un cos.

## 5.8. G vs I: quatre tests

**La cornisa no exigeix absència de plataforma:** poden coexistir, i ho fan quan hi ha plataforma volada al cos basal i cornisa marcant la junta amb la cambra. Fer-les excloents crearia un artefacte.

**Tests en ordre de prioritat. El primer que decideix, decideix.**

| # | Pregunta | Resposta |
| --- | --- | --- |
| **T1** | Hi ha evidència positiva d'E o F? (encaixos, forats passants, fusta conservada, mènsules al buit) | **Sí → plataforma.** És l'evidència que sobreviu millor: la roca conserva l'encaix molt després que la fàbrica caiga. `T_LOST_ELEMENTS` amb *Corbels into void* o *Beam hole* sosté `Corbelled_Courses = 3` o `Sys_Platform = Attested lost` |
| **T2** | Hi ha cos construït damunt? | **No → no pot ser cornisa.** Una cornisa intercòs sense segon cos és una contradicció de termes. Si damunt no hi ha res, ni present ni atestat, la filada volada és **G** |
| **T3** | El voladís és progressiu o d'un sol gest? | Filades successives cadascuna més enfora, acumulant volada → **G**. Una o dues filades que sobreïxen per igual, amb el mur reprenent la vertical damunt → **I** |
| **T4** | Continuïtat | Recorre tota l'amplària de la façana marcant la junta → **I**. Es concentra on hi havia superfície practicable, sovint desbordant l'amplària del cos → **G** |

**Quan cap test decideix.** El domini de cinc valors no té valor per a «observat però no identificable». Sortida:

> `Interbody_Cornice = 9` i `Sys_Platform = Not observable`, amb `Systems_Notes` descrivint el que es veu, **i fila a `T_ARCH_FEATURES`** amb el detall morfològic.

No és elegant, però l'alternativa —triar una de les dues per convenció— fabrica dades. `T_ARCH_FEATURES` existeix precisament perquè el que no encaixa quede registrat en lloc de perdre's.

**Comprovació sobre el corpus: pocs casos ambigus.** Els quatre tests i la sortida per 9 són suficients; no cal replantejar la distinció.

**Solució de segona passada:** la profunditat de volada és mesurable amb Metashape sobre el model. Registrada a `Systems_Notes` amb fórmula constant, un llindar es podrà aplicar retroactivament quan el corpus diga on és la frontera real. **Ara no es fixa cap llindar mètric**: seria inventar-lo.

---

# **6. Resum de canvis d'esquema**

## 6.1. `T_STRUCTURES` — 128 → 132 camps

**Nous (5):**

| Camp | Tipus | Secció | Pestanya |
| --- | --- | --- | --- |
| `Stone_Format` | TEXT(30) | 2.2c Maçoneria | 2.Arq |
| `Stone_Format_Secondary` | TEXT(30) | 2.2c Maçoneria | 2.Arq |
| `Stone_Working` | TEXT(20) | 2.2c Maçoneria | 2.Arq |
| `Access_Plane` | TEXT(20) | 2.7b Paisatge i orientació | 2.Arq |
| `Portal_Orientation` | TEXT(5) | 2.7b Paisatge i orientació | 2.Arq |

**Retirat (1):** `Interbody_Cornice_Material` TEXT(20).

**Per defecte NULL** a tots cinc (regla 8.2: només els 20 elements A–X porten el 0 de farciment).

## 6.2. `T_DECORATIONS` — +1 camp

`Position_Relative` TEXT(20), per defecte NULL.

## 6.3. `L_STRUCT_BODY` — 14 → 18 entrades

−1 (`SOC`), +5 (`RTW`, `PRT`, `SIL`, `LIN`, `ROC-OVL`), 3 renoms VAL, 2 redefinicions ROC, `Sort_Order` reassignat al bloc ROC.

## 6.4. Res més

Cap taula nova, cap relació nova, cap nom de camp anglés alterat, cap valor de `L_DEC_TYPE` tocat.

---

# **7. Regles noves (R32–R39)**

La bateria passa de 31 a 39 regles (40 branques: R38 en porta dues).

| # | Condició | Tipus | Acció suggerida |
| --- | --- | --- | --- |
| **R32** | `Stone_Format_Secondary` amb `Stone_Format` nul o `ND` | Error | Parella òrfena: fixa el format dominant o buida el secundari |
| **R33** | `Stone_Format_Secondary` = `Stone_Format` | Error | La parella exigeix dos valors distints |
| **R34** | Fila `BAS` amb tipus de relleu (fris, motiu escalonat, triangular, ninxols) i `Decorative_Socle` = 0 | Avís | Un relleu al cos basal **és** un tractament plàstic |
| **R35** | Fila `LIN` amb `Lintel` = 0, o fila `SIL` amb `Sill` = 0 | Avís | Reconcilia la posició de la decoració amb l'element |
| **R36** | `Access_Plane` = `Facade` i `Portal_Orientation` ≠ `Facade_Orientation` | Avís | Si l'accés és al pla exposat, les dues orientacions han de coincidir |
| **R37** | Fila amb `Level_Type` = `ROC` i `Body_No` no nul | Error | Les files rupestres no estan en cap cos constructiu |
| **R38a** | `ROC`, `ROC-PER` o `ROC-PAN` amb `Substrate` ≠ `Bedrock` | Error | Revisa el test 1: probablement és pigment de l'estructura |
| **R38b** | `ROC-OVL` amb `Substrate` ≠ `Mixed` | Error | Una fila solapada té substrat mixt per definició |
| **R39** | `Interbody_Cornice` ∈ (1,2,3) amb `N_Basal_Bodies + N_Chamber_Bodies` < 2 | Avís | Una cornisa intercòs exigeix dos cossos: revisa si és una filada en voladís (G) |

*R34 és **unidireccional**. Una versió bidireccional (C ↔ fila `BAS`) seria errònia: C i les files són independents — pot haver-hi motllura sense motiu i motiu sense motllura. Només és incoherent el relleu sense C.*

*R39 és T2 automatitzat i captura el cas que més preocupa: la plataforma arrasada codificada com a cornisa.*

> **Restricció tècnica.** Una unió de 40 branques supera el límit «query too complex» de JET. Cal **una quarta consulta parcial `QRY_16d`**, unida a `QRY_16_Validation_Check` com les tres existents.

## 7bis. Bug heretat de v15 detectat i corregit `[DECIDIT]`

En preparar la implementació s'ha detectat una **incoherència de noms a `chachapoya_DB_v15.bas`** que deixa la bateria de validació sense punt d'entrada:

| Lloc | Nom usat |
| --- | --- |
| Creacio de la tercera parcial (linia 2050) | `QRY_16c_Rules_22_30` |
| Array de baixa `qn(24)` (linia 1416) | `QRY_16c_Rules_22_31` |
| `FROM` de `QRY_16_Validation_Check` (linia 2054) | `QRY_16c_Rules_22_31` |

**Conseqüència:** `CreateQueryDef` de `QRY_16_Validation_Check` falla perquè la font no existeix. `MkQuery` captura l'error i **només l'escriu a la finestra d'immediat** (`*** FAILED to create ...`), de manera que `BuildDB` acaba amb aspecte d'exit i **la consulta de validacio no arriba a existir**. Les tres parcials si que es creen i es poden obrir per separat; el que falta es la unio.

**Comprovacio feta:** la resta de noms de consulta —28 creades contra 28 declarades— son coherents. Aquesta es l'unica discrepancia.

**Correccio a v16:** el nom de creacio passa a `QRY_16c_Rules_22_31`, coherent amb l'array i amb la unio. La unio passa a quatre parcials amb `QRY_16d_Rules_32_39`.

**Leccio de disseny:** `MkQuery` degrada un error a missatge de depuracio, que es el comportament correcte per a no avortar una construccio de 28 consultes per una sola, pero deixa el fracas invisible si ningu no mira la finestra d'immediat. **`BuildDB` ha de portar un comptador de consultes creades i comparar-lo amb el previst**, i informar-ho a la linia final del log com ja fa amb les taules.

---

# **8. Formulari**

## 8.1. 2.Arq

Secció «Maçoneria i morter», **davant de `Masonry_Type`** (es tria la pedra abans d'apilar-la, i eixe és l'ordre de la cadena operativa):

- *Format pedra* — `Stone_Format`
- *Format secundari* — `Stone_Format_Secondary`
- *Treball pedra* — `Stone_Working`

Secció «Façana i paisatge», sota `Facade_Orientation`:

- *Pla d'accés* — `Access_Plane`
- *Orientació portal* — `Portal_Orientation`

Combos de llista de valors amb dues columnes (anglés ocult, valencià visible), com `Masonry_Type`. **Cap lookup nou.**

## 8.2. 4.Dec

- *Posició relativa* (`Position_Relative`) als **dos** subformularis
- Meitat rupestre: `Substrate` restringit a `Bedrock` / `Mixed`
- Meitat rupestre: `Body_No` desactivat
- Etiquetes de posició actualitzades a les dues meitats

## 8.3. 11.Sist

- *Mat. cornisa (I)* eliminat, amb el seu gating de nivell 3
- Etiqueta *Tancament posterior* → **Fons de cambra (W)**
- Etiqueta *Coberta cambra (X)* conservada; el criteri de contacte va a la descripció i a l'esquema

## 8.4. Recordatori de restricció

El gating viu duplicat en dos fitxers (`chachapoya_Form_v15_val.bas` i qualsevol script d'actualització). Qualsevol canvi s'ha d'aplicar als dos o el paquet diverge segons la ruta d'instal·lació.

---

# **9. Consultes**

| Consulta | Canvi |
| --- | --- |
| `QRY_13` (matriu AX+SYS) | Filtre NA per a `Chamber_Roof` amb `Chamber_Roof_Type = Natural bedrock`; columna `Interbody_Cornice_Material` retirada |
| `QRY_16c_Rules_22_31` | **Correccio de nom** (secció 7bis): es creava com a `..._22_30` |
| `QRY_16d_Rules_32_39` | **Nova**: quarta consulta parcial de la bateria (R32–R39) |
| `QRY_16_Validation_Check` | Uneix ara quatre parcials |
| `QRY_17_RockArt_All` | Afig `Position_Relative`; el filtre ROC continua sent `Level_Type = 'ROC'` i ara abraça `ROC-OVL` |
| `QRY_07_Export_QGIS` | Afig `Access_Plane` i `Portal_Orientation` (H06) |
| `QRY_19_V16_Review` | **Nova**: llista de treball de la revisió manual (secció 10) |

---

# **10. Transferència v15 → v16 i revisió manual**

**Estratègia habitual:** v16 construïda **de zero** amb els `.bas` canònics, i transferència en dos passos amb CSV inspeccionable entremig (`chachapoya_EXPORT_v15.bas` → CSV editable → `chachapoya_IMPORT_v16.bas`). **No un pedaç sobre v15.**

## 10.1. El que la transferència ha de resoldre

**(a) Remapatge de `SOC` → `BAS`** a `T_DECORATIONS`, **abans** de retirar l'entrada. És exactament el parany de FK silenciosa que va aparéixer al v11→v15 amb la reorganització d'IDs de `L_STRUCT_BODY`. **Remapar per nom, no per ID.**

**(b) Camps que arriben NULL** pel principi conservador — un valor transferit sense revisió afirmaria un judici que ningú no ha fet:

| Camp | Motiu |
| --- | --- |
| `Sill`, `Jambs`, `Lintel` | El criteri de diferenciació (5.4) canvia què compta com a present. **Molts passaran d'1 a 0** |
| `Chamber_Roof` i `Chamber_Roof_Type` **només** on el tipus és `Natural bedrock` o `Mixed` | El test de contacte (5.2) exigeix rejudici. On el tipus és obra, no cal |

*La resta d'elements A–X es transfereixen intactes: cap altre canvia de definició.*

**(c) Files de decoració a revisar, no a buidar.** Les files amb posició `JAM` o `OVL` poden ser portals que ara serien `PRT`, `SIL` o `LIN`. Buidar-les perdria tot el registre; es **llisten** a `QRY_19_V16_Review` per a revisió humana.

**(d) `Position_Relative` arriba buit** a totes les files: camp nou, cap dada d'origen.

## 10.2. Cost real de la revisió

Sobre 35 registres, la revisió manual afecta:

- Portal (N, O, Q) de **tots** els registres amb `Sys_Portal` present
- Coberta de cambra dels registres amb sostre de roca o mixt (**molts**)
- Files de decoració en posició `JAM` / `OVL`
- Format i treball de la pedra: **camps nous, entrada des de zero a tot el corpus**

**Això no és feina de codi: és feina de fitxa i fotografia.** Convé saber-ho abans de començar, i és una raó addicional per a no acumular-hi el punt 5.

---

# **11. Ordre d'implementació previst**

| Pas | Contingut | Depén de |
| --- | --- | --- |
| 1 | Revisió i tancament d'aquest delta | — |
| 2 | `chachapoya_DB_v16.bas`: esquema, lookups, relacions | Pas 1 |
| 3 | Correcció del nom de `QRY_16c`; `QRY_16d`; bateria a quatre parcials; comptador de consultes a `BuildDB` | Pas 2 |
| 4 | `QRY_13` amb el filtre NA; `QRY_07` i `QRY_17` ampliades | Pas 2 |
| 5 | `QRY_19_V16_Review` | Pas 2 |
| 6 | `chachapoya_Form_v16_val.bas`: 2.Arq, 4.Dec, 11.Sist, gating | Pas 2 |
| 7 | `chachapoya_EXPORT_v15.bas` amb els camps de revisió a NULL | Pas 2 |
| 8 | `chachapoya_IMPORT_v16.bas` amb remapatge `SOC` → `BAS` per nom | Pas 7 |
| 9 | Transferència sobre la còpia local real i verificació de recomptes | Passos 6–8 |
| 10 | Revisió manual guiada per `QRY_19_V16_Review` | Pas 9 |
| 11 | Reexecució de la matriu A–X i comparació amb els resultats anteriors | Pas 10 |
| 12 | Actualització de `esquema_bbdd_estructures_v16.md` i `tfm_metodologia_bbdd_v16.md` | Passos 1–11 |
| 13 | Diagnòstic per a alimentar el delta v16→v17 (secció 13) | Pas 10 |

**Regla mantinguda:** cap valor emmagatzemat es toca sense verificació prèvia i comptador d'afectació al log.

---

# **12. Punts oberts, ajornats al delta v16 → v17**

**El punt gran:**

- **Descomposició per subnivells.** Descriure per separat cada cos basal i cada cos cambra (N0, N1.1, N1.2, amb intranivells N0–N1.1 i N1.1–N1.2 derivats), carregant dinàmicament els subnivells declarats a `N_Basal_Bodies` i `N_Chamber_Bodies`. **No és vocabulari: és cardinalitat** — implicaria una taula `T_BODIES` amb els elements A–X penjant de cada cos, i tocaria esquema, formulari, matriu A–X, `QRY_13` i el criteri d'unitat de registre. Decideix-lo la dada de la secció 13. Si són pocs els registres amb cossos d'atributs divergents, la solució ja existeix (registres separats + `T_CONNECTIONS`); si són molts, cal `T_BODIES`.
- Conseqüència sobre camps ja tancats: `Stone_Format` i `Stone_Working` són candidats a baixar a nivell de cos; el criteri de numeració de `Body_No` es fixa allí; X per cos.

**Heretats de deltes anteriors:**

- `Chinking_Stones` (v13, 8.6): retirar, mantindre com a descriptor o substituir per una variable més fina
- `Traces only` a `Plaster_Extent` (v13, 8.5)
- Relació cromàtica entre colors contigus (v13, 7.10)
- Confinament fabricat: camp propi si apareix repetidament?
- Circulació vertical vs horitzontal a OE3
- Litologia i potència dels bancs com a variable independent per a H02
- Coordenades pròpies a `T_DECORATIONS` (només si cal posició fina de les pintures associades)

**Nous d'aquest delta:**

- **Element Y (`Abutted_Mass`)** si es documenten 3 casos o més de banqueta adossada
- **Caracterització de la superfície rocosa** a les files ROC, si el corpus de pintura creix
- **Camp d'asimetria de portal**, si el patró resulta recurrent
- **Llindar mètric de volada** per a la distinció G/I, quan Metashape done prou mesures

---

# **13. Comprovacions pendents sobre el corpus**

Totes sobre 35 registres: qüestió de minuts, i cadascuna desbloqueja una decisió.

| Comprovació | Desbloqueja |
| --- | --- |
| Registres amb `Chamber_Roof_Type = Natural bedrock`: recompte exacte | Magnitud del canvi a la matriu A–X (5.3) i abast de la revisió manual |
| Distribució de `Stone_Working` un cop entrat: quants `Dressed` | Si el camp discrimina o la seua universalitat és el resultat (1.6) |
| Files de `T_DECORATIONS` amb `ID_Struct_Body = SOC` | Volum del remapatge (10.1a) |
| Files amb posició `JAM` o `OVL` que siguen portals | Volum de la revisió de decoració (10.1c) |
| Registres amb `Sys_Portal` present i `Sill`/`Jambs` = 1 | Volum del rejudici pel criteri de diferenciació (5.4) |
| **Registres amb dos o més cossos d'atributs divergents** | **Decideix el punt 5: `T_BODIES` o registres separats** |
| Casos de banqueta adossada (`Abutted mass`) | Element Y al delta següent (3.5) |
| Comparació matriu A–X abans/després del filtre NA | Validesa de les proves estadístiques ja executades |
