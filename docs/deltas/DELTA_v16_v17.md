**DELTA v16 → v17**

*Base de dades arqueològica — La Petaca i Diablo Wasi (PALP)*

Esteve Ribera Torró | TFM Arqueologia UA

**Document viu.** S'hi anoten les decisions d'actualització a mesura que es prenen, amb la justificació que les sosté i les passes d'implementació que impliquen. Quan el conjunt estiga tancat, aquest document és l'especificació des de la qual es generen els scripts v17 i s'actualitzen `esquema_bbdd_estructures`, `tfm_metodologia_bbdd` i `manual_us_bbdd`.

**Estat:** DECIDIT, pendent d'implementació (agost 2026).

**Origen.** A diferència del delta anterior, aquest naix de **dues fonts alhora**: les observacions recollides emplenant registres (`Prev16-17.md`) i, per primera vegada, la **inspecció directa de la còpia local** (`DB_v16.accdb`, 36 registres, 56 files de decoració). La segona ha canviat el resultat de manera substancial. Tres punts que es plantejaven com a preguntes de disseny obertes resultaven ser **incoherències ja presents a les dades**, i dos camps que es discutien en abstracte tenien un ús real que decidia la discussió sense necessitat d'argumentar-la.

**Abast.** Onze punts de disseny tancats, tres correccions de dades i un pedaç de formulari que no és de disseny. La descomposició per subnivells (`T_BODIES`), ajornada des del delta anterior, **es torna a ajornar amb una diferència essencial: amb un comptador posat**. Vegeu la secció 6.

**Convenció d'estat de cada punt:**
`[DECIDIT]` acordat, pendent d'implementar · `[OBERT]` en discussió · `[PENDENT DADES]` depén d'una comprovació sobre el corpus · `[FET]` implementat i verificat

---

# **0. Resum**

**Bloc A** — marc de referència de la decoració (secció 1).

| # | Punt | Estat | Abast |
| --- | --- | --- | --- |
| A1.1 | `Position_Relative` eliminat | `[DECIDIT]` | Esquema + formulari + migració |
| A1.2 | Meitat arquitectònica sense reemplaçament | `[DECIDIT]` | Documentació |
| A1.3 | Quatre camps de tram a la meitat rupestre | `[DECIDIT]` | Esquema + formulari |
| A1.4 | Convenció d'observador corregida: pla exposat | `[DECIDIT]` | Documentació |
| A2.1 | `Outline_Geometry` | `[DECIDIT]` | Esquema + formulari |
| A2.2 | U invertida, envoltant i flanquejant derivades | `[DECIDIT]` | Consulta |

**Bloc B** — vocabulari (secció 2).

| # | Punt | Estat | Abast |
| --- | --- | --- | --- |
| B1.1 | `Rock dihedral`: entrada nova a `L_SUPPORT` | `[DECIDIT]` | Lookup |
| B1.2 | `Rock surface` + mode `Substrate` | `[DECIDIT]` | Lookup |
| B1.3 | Criteri principal/secundari: verticals/horitzontals | `[DECIDIT]` | Documentació |
| B2 | `Body_No` ocult a la meitat rupestre | `[DECIDIT]` | Formulari (condicionat a P1) |
| B3.1 | Dues etiquetes d'element corregides | `[DECIDIT]` | Formulari |
| B3.2 | Singular/plural entre pestanyes: es manté | `[DECIDIT]` | Documentació |

**Bloc C** — unitat de registre (secció 3).

| # | Punt | Estat | Abast |
| --- | --- | --- | --- |
| C1.1 | Estructura desapareguda = registre d'estructura | `[DECIDIT]` | Convenció + correcció de dades |
| C1.2 | Criteri de frontera: geometria del traç | `[DECIDIT]` | Documentació |
| C1.3 | Regla 40 | `[DECIDIT]` | Bateria |

**Bloc D** — coherència del model (secció 4).

| # | Punt | Estat | Abast |
| --- | --- | --- | --- |
| D1.1 | Llista de connexions entrants, de només lectura | `[DECIDIT]` | Formulari + consulta |
| D1.2 | La cronologia designa l'estructura anterior | `[DECIDIT]` | Esquema + migració |
| D1.3 | Ordre normalitzat + índex únic sobre el parell | `[DECIDIT]` | Esquema |
| D2 | `ID_Material_Status` es trasllada a 7.Mat | `[DECIDIT]` | Formulari + gating |
| D3 | Tres capes de materials culturals: justificades | `[DECIDIT]` | Documentació + regles |
| D4.1 | Frontera fris: element agregat, files de detall | `[DECIDIT]` | Documentació + regla |
| D4.2 | Tipus `Frieze / Greca` retirat | `[DECIDIT]` | Lookup |
| D5 | Judicis agregats confirmats | `[DECIDIT]` | Documentació |

**Bloc E** — gating i dominis (secció 5).

| # | Punt | Estat | Abast |
| --- | --- | --- | --- |
| E1.1 | Pla d'accés i orientació de portal, gatejats | `[DECIDIT]` | Formulari |
| E1.2 | Orientació de façana i visibilitat: es queden obertes | `[DECIDIT]` | Cap canvi |
| E2 | Criteri general: deduïble → gating; judici → «no aplicable» | `[DECIDIT]` | Documentació |
| E3 | `Sys_Interface` | `[DECIDIT]` | Esquema + formulari |

**Bloc F** — cardinalitat (secció 6).

| # | Punt | Estat | Abast |
| --- | --- | --- | --- |
| F1 | `T_BODIES` ajornada; unitat d'anàlisi = estructura | `[DECIDIT]` | Decisió de disseny |
| F2.1 | `Fabric_Divergence` | `[DECIDIT]` | Esquema + formulari |
| F2.2 | `Divergence_Location` | `[DECIDIT]` | Esquema + formulari |
| F2.3 | Frontera divergència/fase: observació vs interpretació | `[DECIDIT]` | Documentació |

**Punts heretats** (secció 7) i **correccions de dades** (secció 8).

---

# **1. Bloc A — el marc de referència de la decoració**

## **1.1. El diagnòstic**

`Position_Relative` es va crear en v16 amb el domini *Left / Right / Both / Above / Below / ND*, compartit per les dues meitats de 4.Dec. La justificació d'aleshores era l'economia: un sol camp que resolia la lateralitat dels elements aparellats (`FFL`, `JAM`, `PRT`, `RTW`) i alhora la posició de la pintura rupestre respecte de l'estructura.

**L'economia era falsa perquè les dues faenes no comparteixen referent.** A la meitat arquitectònica, *Esquerra* designa **quina instància d'un parell** porta el motiu: una relació identificacional, interna a l'estructura. A la meitat rupestre, *Esquerra* designa **on està la pintura en relació amb el volum construït**: una relació topològica entre dos objectes separats. Que les dues faces compartisquen les paraules *esquerra* i *dreta* és una coincidència del llenguatge, no una propietat comuna.

Quatre conseqüències, totes elles errors vius a v16 i no incomoditats:

**(a) `Above` i `Below` no signifiquen res a la meitat arquitectònica.** L'eix vertical de la decoració arquitectònica el diuen `ID_Struct_Body` (ordenat *bottom-to-top* per `Sort_Order`) i `Body_No`. «Un relleu al dintell, posició *Damunt*» no té referent possible: damunt del dintell és `OVL`, damunt del cos és `Body_No`+1. Són valors que només poden produir registres inútils o contradictoris amb la posició.

**(b) `Both` significa dues coses distintes.** A la meitat arquitectònica és *tractament bilateral simètric* —una afirmació arqueològica sobre un parell d'elements—; a la rupestre significaria *als dos costats de l'estructura*, una descripció topogràfica d'una banda que no té parell. La documentació v16 només defineix la primera.

**(c) El pla de referència declarat és el que no toca per a l'art rupestre.** La convenció v16 remetia a `Access_Plane`. Per a una pintura sobre la penya la referència correcta és **el pla exposat** —la façana tal com la va redefinir la v16b—, perquè la pintura es veu des de la vall o des de l'aproximació, no des de l'accés. En una estructura amb `Access_Plane = Return wall`, la convenció vigent **inverteix o gira** l'esquerra i la dreta de les files ROC. Bug latent que només es manifesta en el cas divergent, que és justament el cas d'interés.

**(d) `PRT` no havia de ser lateralitzable**, i això es dedueix d'una decisió ja tancada. `PRT` existeix perquè una banda contínua al voltant de l'obertura és *un sol gest, no tres*. Un `PRT` amb costat *Esquerra* és una contradicció interna: si el tractament té costat, és elemental, i el criteri v16 ja diu «en cas de dubte, elemental» (`JAM`).

## **1.2. El que deien les dades**

7 files de 56 portaven valor, **totes ROC, cap arquitectònica**.

| Valor | n | Contingut real |
| --- | --- | --- |
| `Both` | 3 | Notes: *banda perimetral esquerra, damunt, dreta, en forma de U invertida* |
| `Above` | 2 | — |
| `Below` | 1 | — |
| `Right` | 1 | Fila 12, `ROC-OVL`, zoomorf bicrom |

`Both` **ja s'estava forçant en producció** per a dir la U invertida, amb la descripció completa refugiada a `Notes`. El camp no era que li faltara un valor: n'estava dient un altre.

## **1.3. A1 — decisió**

**`Position_Relative` s'elimina de `T_DECORATIONS`.**

**A la meitat arquitectònica no se substitueix per res.** Cap de les 56 files fa servir la lateralitat sobre posicions arquitectòniques, i la revisió del corpus va confirmar que `QUO` i `PIL` no la demanen — cosa que resol de passada el problema de les pilastres múltiples, on *Esquerra/Dreta* hauria deixat de funcionar. El cas que semblava exigir-la —decoració al pla de façana amb accés lateral— és una propietat de **l'estructura**, no d'una fila de decoració, i `Access_Plane` ja el registra des de v16. Si apareix una decoració realment asimètrica (un brancal decorat i l'altre no), la convenció és anotar-ho a `Notes` de la fila amb fórmula constant i reobrir el punt amb el cas damunt la taula.

*Crear un camp de lateralitat sense cap ocurrència en 56 files contradiria el criteri que el mateix projecte va aplicar per a limitar les posicions ROC a quatre: no s'inventen categories sense casos.*

**A la meitat rupestre el substitueixen quatre camps de tram.**

| Camp | Etiqueta | Tipus | Domini |
| --- | --- | --- | --- |
| `Span_Left` | Esquerra | BYTE | 0 absent / 1 present / 9 no observable |
| `Span_Above` | Damunt | BYTE | ídem |
| `Span_Right` | Dreta | BYTE | ídem |
| `Span_Below` | Davall | BYTE | ídem |

> **Definició:** quins trams del contorn de l'estructura —**present o desapareguda**— ocupa aquest motiu.
> **Habilitats** amb `ID_Struct_Body` ∈ {`ROC-OVL`, `ROC-PER`, `ROC`}. Desactivats amb `ROC-PAN`, que penja d'un registre PR i per definició no té estructura de referència, i a tota la meitat arquitectònica.
> **Per defecte NULL**, mai 0: la regla 4 de farciment amb zero s'aplica exclusivament als 20 elements A–X.

**Per què quatre camps i no una llista de cobertures.** Quatre trams donen quinze combinacions, i el corpus ja n'ha produït almenys tres distintes (*damunt*, *davall*, i la U invertida), més un cas de **davall + dreta** identificat en la revisió. Cap llista tancada raonable les cobreix. Però l'argument decisiu no és combinatori sinó **epistèmic**: un valor únic no pot distingir *la banda no cobria la base* de *el tram de base no és observable*, i eixa distinció és exactament la que separa una U invertida d'un anell tancat mal conservat. El corpus conté casos on el tram superior queda tapat per la visera rocosa i pel cos superior: marcar-lo com a absent afirmaria una cosa que ningú no ha pogut mirar.

**Convenció d'observador, corregida.** **Des de davant del pla exposat** (la façana com a pla exposat, definició v16b), **mai des de `Access_Plane`**. Això corregeix la convenció v16, que girava l'esquerra i la dreta a les estructures d'accés desviat.

## **1.4. A2 — la banda en U invertida**

La U invertida no és una propietat sinó quatre, i dues ja tenien camp:

| Propietat | On viu | Estat abans de v17 |
| --- | --- | --- |
| Ressegueix el contorn de l'estructura | `ID_Struct_Body` = `ROC-PER` | Existia (v16) |
| Sembla una banda perimetral | `ID_Dec_Type` = `RA Perimeter band` | Existia (v13) |
| Quina part del contorn cobreix | Els quatre camps de tram | **Nou (A1.3)** |
| Amb quina geometria de traç | `Outline_Geometry` | **Nou (A2.1)** |

**`Outline_Geometry`** — TEXT(15). Etiqueta: *Geometria del traç*. Domini: `Orthogonal` / `Curvilinear` / `Irregular` / `ND`.

> **Visible a totes les files ROC i enlloc més.** No es gateja per tipus de decoració: les tres files del corpus que descriuen una U invertida estan registrades amb **tres tipus diferents** (*RA Amorphous stain*, *RA Perimeter band*, *Painted band*), de manera que qualsevol filtre per tipus les deixaria fora precisament a elles.

**Per què és un camp i no una nota.** La distinció no és estètica. Una banda que fa angles rectes dibuixa un rectangle, és a dir, **afirma un referent arquitectònic construït**; una banda corba no afirma res d'això i pot estar resseguint un rebaix natural de la penya. És **l'única evidència disponible** per a decidir el cas del bloc C. A més, la fila 56 del corpus ja porta la paraula *(geomètrica-ortogonal)* escrita a mà a `Notes`, que és exactament el senyal que `QRY_18_Notes_Review` existeix per a detectar: una observació repetida en text lliure ha de ser un camp.

*Registra **l'observació**, no la inferència. L'afirmació «ací hi havia una estructura» viu a `T_LOST_ELEMENTS`, igual que `Vertical association` registra la verticalitat sense afirmar-ne el mecanisme.*

**U invertida, envoltant i flanquejant no s'emmagatzemen: es deriven.** Són etiquetes de lectura calculades des dels quatre trams, seguint el mateix patró de judici agregat que `Dec_Present`. L'avantatge és que una U invertida amb la base a 9 apareix com a **candidata a envoltant** en lloc de quedar afirmada com a U.

## **1.5. Eix descartat explícitament**

**La posició dins d'un element** —«al terç esquerre del dintell»— **no s'obri**. No s'ha identificat cap ús analític, `Notes` ho cobreix, i obrir-lo multiplicaria combinacions sense cap hipòtesi que les demane. Queda escrit perquè no torne.

## **1.6. Conseqüència sobre `Access_Plane`**

`Access_Plane` tenia dues faenes documentades: dir per quin pla s'entra, i **declarar el pla de referència** de `Position_Relative`. La segona desapareix amb el camp. El camp es queda intacte per la primera, però **la documentació v16 (seccions 1.5(c), 2.x i 3.4) s'ha de corregir**, perquè afirma que governa una cosa que ja no existeix.

## **1.7. Migració**

7 files. El tram declarat passa a `1`; els altres tres queden **NULL per a revisió**, segons el principi conservador: un valor transferit sense revisió afirmaria falsament ser un judici. Les tres files amb `Both` i nota de U invertida es completen a mà des de la nota.

---

# **2. Bloc B — vocabulari**

## **2.1. B1 — el díedre rocós**

**Cas:** l'estructura aprofita un angle entrant de la roca, format per dues parets aproximadament ortogonals amb sostre triangular. Confirmat com a **forma natural aprofitada**, no retallada.

**Terminologia.** No hi ha terme arqueològic consagrat. La descripció geològica sí que és precisa: el buit correspon al negatiu d'un **bloc en tascó** despresa al llarg de dues discontinuïtats que s'encreuen (*wedge failure*), amb el sostre triangular com a tercer pla de fractura. S'adopta **díedre** —terme d'escalada i de descripció de parets, immediat en castellà i valencià— perquè descriu **el que s'observa** i no la gènesi, coherent amb el criteri geomètric adoptat per a les posicions ROC en v16.

| Codi intern | `Name` | `Name_VAL` | `Support_Mode` |
| --- | --- | --- | --- |
| — | `Rock dihedral` | Díedre rocós (angle amb sostre) | `Confinement` |

> **Criteri de frontera amb les cavitats** (mida no, geometria sí):
> **Cavitat** — buit excavat cap endins de la paret. S'hi entra per una boca; la roca envolta per darrere i pels costats.
> **Díedre** — angle entrant entre dos plans de roca. **No hi ha boca**: està obert en dues direccions.
> **Test de camp:** quantes parets estalvia la roca? Al díedre, l'estructura no construeix els dos murs que la roca li dona; a la cavitat, construeix la façana i prou.

*El sostre triangular queda com a **descripció habitual però no obligatòria**. El que compta arqueològicament és que la roca estalvia dues parets i una coberta, no si el sostre és un pla de fractura o un d'estratificació.*

**El pla vertical aïllat es descarta**: revisat, era també un díedre. La diàclasi ja implica confinament per tots dos costats i es queda intacta.

## **2.2. B1b — la superfície de roca**

**Cas:** panells d'art rupestre **sense estructura associada** (registres de tipologia PR). Múltiples casos al corpus, encara no entrats.

| Codi intern | `Name` | `Name_VAL` | `Support_Mode` |
| --- | --- | --- | --- |
| — | `Rock surface` | Superfície de roca | **`Substrate`** *(mode nou)* |

*El mode és nou perquè cap dels tres existents descriu el que fa: una superfície que sosté pigment **no** és descans gravitacional, ni confinament, ni encastament. La categoria pròpia la separa netament de les deu formes que sostenen construcció, i evita que aparega als recomptes de suport constructiu.*

## **2.3. B1c — criteri principal/secundari**

> **El suport principal respon a les forces verticals** (aguanta el pes). **El secundari respon a les horitzontals** (estabilitza).

**És una orientació, no una regla de validació.** Tres dels quatre suports secundaris actualment emplenats són repises o cavitats, formes que aguanten pes, i poden ser perfectament correctes: una estructura pot descansar alhora sobre una repisa i sobre el sòl d'una cavitat. Una regla estricta les marcaria com a errònies sense motiu. No s'afig cap comprovació a la bateria.

*Amb aquest criteri, el díedre serà quasi sempre secundari: el que sosté l'estructura continua sent la repisa o el sòl on descansa.*

## **2.4. B2 — `Body_No` ocult a la meitat rupestre**

`Body_No` passa d'**estar desactivat i visible** a **estar ocult** al subformulari `F_ROCKART`.

> **Condició d'implementació:** el pedaç que recupera les 16 files invisibles (secció 8, P1) va **abans o alhora**, mai després. Amagar controls sense haver resolt que hi ha files inaccessibles multiplica el risc de tornar a tindre dades que existeixen i no es veuen.

La regla 37 continua vigilant el que hi entre per un altre costat.

## **2.5. B3 — etiquetes d'element**

El renomenament de v16 es va aplicar a `L_STRUCT_BODY` (4.Dec) i no a les etiquetes d'element de 11.Sist. Comprovat contra el codi: dues línies.

| Camp | Etiqueta 11.Sist actual | Etiqueta v17 | 4.Dec (ja correcta) |
| --- | --- | --- | --- |
| `Facade_Flank` | Ala de façana (L) | **Flanc de façana (L)** | Flanc de façana (L) |
| `Return_Wall` | Mur de retorn (V) | **Mur lateral de cambra (V)** | Mur lateral de cambra (V) |

`Rear_Closure_Type` ja diu *Fons de cambra (W)*: correcte, sense canvi.

**El singular/plural entre pestanyes es manté, amb la justificació escrita** perquè ningú no ho «corregisca»:

| Camp | 11.Sist | 4.Dec |
| --- | --- | --- |
| `Jambs` | Brancals (O) | Brancal (O) |
| `Corner_Quoins` | Cantoneres (J) | Cantonera (J) |
| `Structural_Pilasters` | Pilastres estruct. (K) | Pilastra (K) |

> A 11.Sist es registra **l'element com a conjunt** (els dos brancals de l'estructura); a 4.Dec es registra **un motiu en una posició concreta** (un brancal). No és una inconsistència: és la diferència entre comptar un element i localitzar-hi un motiu.

**Nota terminològica nova, germana de la de *llindar*/*llinda*:** **flanc** (L) i **mur lateral** (V) no són sinònims. El flanc està **dins** del pla de façana; el mur lateral **gira** cap al cingle. Si les etiquetes es tornen a tocar, eixa és la distinció que s'ha de preservar.

---

# **3. Bloc C — unitat de registre**

## **3.1. El cas**

Estructura desapareguda de la qual només resta la pintura perimetral en U invertida sobre la roca.

## **3.2. El que hi havia: una decisió presa dues vegades en sentits contraris**

L'entrada `Unclassifiable` de `L_TYPOLOGY` porta a la descripció: *«evidència insuficient per a classificar una estructura CONSTRUÏDA (p. ex. EA11: perímetre de pigment sense construcció conservada)»*, amb `Record_Class = Built funerary structure`.

**Però DW-S01-EA11 està registrat amb `ID_Typology = 7` (`PR Rock Art`, classe *Rock art panel*).** I la seua única fila de decoració —la que porta la nota *aparentment en forma de U invertida (geomètrica-ortogonal)*— està en posició `ROC-PER`.

Doble incoherència: el vocabulari diu una cosa i el registre una altra; i **`ROC-PER` no pot existir en un registre d'art rupestre**, perquè significa «ressegueix el contorn de l'estructura» i un panell aïllat no en té cap — li correspondria `ROC-PAN`.

## **3.3. C1 — decisió**

**És un registre d'estructura**, amb tipologia `Unclassifiable`.

**Composició del registre:**

1. `T_STRUCTURES` — tipologia `Unclassifiable`, vocabulari constructiu A–X a **9** (no observable), no a 0.
2. `T_LOST_ELEMENTS` — una fila amb `Evidence_Scope = Whole structure` i `ID_Evidence_Type = Pigment on bedrock` (entrada 7 de `L_LOST_EVIDENCE`: *pigment sobre penya ara nua de la fàbrica que el portava*).
3. `T_DECORATIONS` — la fila `ROC-PER` existent, amb els quatre trams i `Outline_Geometry` emplenats.

**Tres raons:**

**(a) Unitat de registre.** Allí hi havia una construcció. Registrada com a art rupestre, desapareix del recompte d'estructures — i els recomptes per sector i jaciment són exactament el que mesuren H02 i H03. Excloure les desaparegudes fa que es compte **preservació** i es llija com a **densitat constructiva**: el mateix problema del denominador de la secció 7.1 de l'esquema, però agreujat, perquè no és un valor absent sinó una fila absent.

**(b) La maquinària existia i no s'usava.** `Evidence_Scope = Whole structure` i el tipus d'evidència *pigment sobre penya* estan tots dos sense cap fila al corpus. Es van crear per a aquest cas.

**(c) L'alternativa perd informació i no en guanya.** Com a art rupestre, la banda queda registrada; que ressegueix un contorn rectangular alineat amb elements arquitectònics, no.

## **3.4. C1b — criteri de frontera**

> **Estructura** quan el contorn pintat és **ortogonal** o **s'alinea amb elements arquitectònics reconeixibles** (brancals, base, coronament).
> **Art rupestre** quan el traç és **corb o irregular** i no s'alinea amb res.

*Aquesta és la faena analítica d'`Outline_Geometry`: no és un descriptor accessori, és el camp que decideix en quina meitat del corpus va el registre.*

## **3.5. Regla nova**

**Regla 40** — registre de `Record_Class = Rock art panel` amb files en posició `ROC-OVL`, `ROC-PER` o `ROC`; **i el recíproc**, registre d'estructura amb files `ROC-PAN`. És l'error que ha passat a EA11 i que ara mateix no detecta res.

---

# **4. Bloc D — coherència del model**

## **4.1. D1 — connexions**

**El que hi havia:** dues files a `T_CONNECTIONS`, i **són la mateixa connexió** — EA07a→EA07b i EA07b→EA07a, amb valors idèntics. Taxa de duplicació: 2 de 2.

**La causa no és descuit.** `F_CONNECTIONS` s'enllaça amb `LinkChildFields = "ID_Struct_A"`: des de B no es veu res del que s'ha registrat des d'A. Amb la interfície v16, duplicar és el comportament per defecte.

### **D1.1 — llista de connexions entrants**

Un segon subformulari a `pgExtra`, **de només lectura**, amb les connexions on l'estructura actual és `ID_Struct_B`.

> **Requisit crític:** la relació cronològica s'ha de mostrar **girada**. Si des d'EA07a s'ha declarat «aquesta és anterior a l'altra», des d'EA07b ha de llegir-se «aquesta és posterior». Sense la inversió, la llista **menteix** cada vegada que hi haja direcció cronològica, i és pitjor que no tindre-la. Com que és de només lectura, es basa en una consulta amb `IIf` i no cal que siga actualitzable.

### **D1.2 — la cronologia designa l'estructura anterior**

`Chrono_Relation` deixa de dir *A earlier than B* / *B earlier than A* i passa a **designar quina de les dues estructures és l'anterior**, conservant `Contemporary` i `Undetermined`.

**Conseqüències:**

- **L'ordre A/B deixa de significar res**, i per tant es pot normalitzar.
- **Desapareix la norma «no invertir l'ordre després de l'entrada»** (secció 9.1 de l'esquema v16): una regla menys que recordar.
- La direcció queda **explícita**, no amagada en la posició. `QRY_14_Connections_Edges` continua exportant un graf dirigit, amb la direcció llegida d'un camp en lloc d'una convenció posicional.
- La **regla 20** (cronologia direccional només amb junta vertical adossada o superposició) continua vigent sense canvis de lògica.

### **D1.3 — ordre normalitzat i índex únic**

Amb l'ordre lliure de significat, `BeforeUpdate` força **sempre l'ID menor a `ID_Struct_A`**, i un **índex únic sobre el parell** fa la duplicació estructuralment impossible en lloc d'una cosa a vigilar.

**Migració:** 2 files. Es refan a mà.

## **4.2. D2 — trasllat de `ID_Material_Status`**

**El problema:** la pregunta que obri el tema dels materials (`Cultural_Materials_Present`) està a 7.Mat; la pregunta sobre l'estat d'eixos materials (`ID_Material_Status`) està a 5.Estat. **La porta i el que hi ha darrere viuen a pestanyes distintes, i la porta va després.**

No és només ordre d'entrada: ara mateix es pot declarar que no hi ha materials a 7.Mat i haver-ne registrat l'estat a 5.Estat, i **res no ho detecta**, perquè el gating de nivell 2 opera dins de cada pestanya.

**Decisió:** `ID_Material_Status` es trasllada a **7.Mat**, immediatament davall de `Cultural_Materials_Present`, i **queda gatejat per ell**.

| Pestanya | Passa a tractar |
| --- | --- |
| **5.Estat** | L'**estructura construïda**: com està, què l'ha malmesa, com de bé s'ha pogut observar |
| **7.Mat** | El **contingut**: si n'hi ha, de quins tipus, en quin estat |

*7.Mat adquireix així exactament la mateixa forma que 6.Bio: una pregunta que obri el bloc i el governa sencer.*

**El que es perd** és la contigüitat visual de la conservació dual (`ID_Arch_Status` al costat de `ID_Material_Status`). Es considera un guany: la proximitat feia pensar que eren dos graus de la mateixa cosa, quan una parla de l'edifici i l'altra del seu contingut. **L'asimetria de v16 es manté**: `ID_Arch_Status` no es desdobla, perquè una estructura sempre existeix — és el registre mateix.

## **4.3. D3 — les tres capes de materials culturals**

**No són redundants.** Les tres responen preguntes amb **llindars d'evidència distints**:

| Capa | Pregunta | Exigeix |
| --- | --- | --- |
| `Cultural_Materials_Present` | N'hi ha? | Veure que dins hi ha coses. Contestable **des de 30 m** |
| Els set `Mat_*` | De quins tipus? | Identificar-los |
| `ID_Material_Status` | En quin estat? | Veure'ls prou bé per a jutjar-ho |

**El cas que ho demostra:** una estructura observada de lluny on es veu clarament que hi ha material però no se'n distingeix cap categoria. La pregunta general val `1` i les set van a `9`. **Si la presència es dedueix dels set, eixe registre passa a dir «no hi ha materials», que és fals.**

És el mateix patró de **judici agregat** que `Dec_Present` i `RockArt_Present`. Amb això queden contestats D3 i D5 alhora.

**Comprovacions noves** (l'existent, regla 19, cobria només un sentit):

- **Regla 43** — `Cultural_Materials_Present` = 1 amb els set `Mat_*` a 0.
- **Regla 44** — `ID_Material_Status` emplenat amb `Cultural_Materials_Present` = 0. *(Quasi desapareix sola amb el trasllat i el gating, però es manté per a detectar dades heretades.)*

## **4.4. D4 — la frontera del fris**

**El problema:** `Relief_Frieze` (element M) i una fila de `T_DECORATIONS` de tipus *Frieze / Greca* en posició `OVL` diuen aparentment el mateix, sense criteri escrit.

**El precedent no s'aplica directament.** Al sòcol es va separar **el modelat de la massa** (motllura, filada volada) del **motiu aplicat**. Ací això no funciona: els frisos chachapoya **són** la fàbrica —el zig-zag i l'escalonat es fan col·locant les pedres endins i enfora—, de manera que «el que és fàbrica va a l'element i el que és motiu va a la fila» **perdria el repertori de motius**, que és la comparació entre jaciments que sostenta H01 i H04.

**Frontera fixada:**

> **L'element `Relief_Frieze` diu que hi ha fris**, com a afirmació agregada. **Les files diuen quin motiu és**, de quin color i sobre quin substrat.
> **No són alternatives: es posen els dos.** Un fris vist de lluny pot quedar com a element present amb cap fila de detall, i això és un registre honest, no incomplet.

*Mateix patró que `Dec_Present` + files, i que `Cultural_Materials_Present` + els set `Mat_*`.*

**Regla 42 (unidireccional, com la 34)** — fila de motiu en relleu amb `Relief_Frieze` = 0. **Ha de mirar `OVL` i `FFL`**, no només la primera: l'element està definit com a tractament del **parament de façana**, i un fris en relleu pot recórrer el parament sense passar per damunt del portal.

### **D4.2 — `Frieze / Greca` retirat de `L_DEC_TYPE`**

**Zero usos en 56 files.** És l'únic tipus de la llista que **no diu quin motiu és**: diu que hi ha un fris, que és exactament el que ja diu l'element. Els altres —zig-zag, escalonat, triangular— sí que especifiquen. El valor `ND` ja cobreix el cas del fris amb motiu no resolt, i mantindre dos valors per a la mateixa situació és el que va causar el problema de les posicions ROC en v15.

*Migració nul·la. Si en el futur la **greca** com a motiu específic (el meandre escalonat andí) mereix entrada pròpia, serà una entrada **nova** amb definició de motiu, no aquesta.*

## **4.5. D5 — judicis agregats: confirmats**

Queda escrit, perquè no es reobri:

| Parella | No és duplicació perquè |
| --- | --- |
| `Dec_Present` / `RockArt_Present` vs les files | L'afirmació general sobreviu encara que no s'haja pogut entrar cap detall. Vigilada per les regles 22–23 i 28–29 |
| `Cultural_Materials_Present` vs els set `Mat_*` | Llindars d'evidència distints (secció 4.3) |
| `Relief_Frieze` vs les files `OVL`/`FFL` | Ídem (secció 4.4) |
| `N_Basal_Bodies` / `N_Chamber_Bodies` vs `Sys_Base` / `Sys_Chamber` | Quantitat contra existència del conjunt com a unitat descriptiva. Vigilada per la regla 11 |

---

# **5. Bloc E — gating i dominis**

## **5.1. E1 — façana i paisatge sense cos cambra**

**`Facade_Orientation` i `Visibility_Valley` es queden obertes.**

El motiu és una decisió ja presa: en v16 la façana es va definir com a **pla exposat**, explícitament deslligada de l'obertura, perquè fixar-la per l'obertura hauria fet que l'orientació deixara d'apuntar a la vall. **Un cos basal sol, o una plataforma, té pla exposat i mira cap a algun lloc.** Desactivar-lo eliminaria del càlcul precisament les estructures sense cambra, que són el grup de control natural de H06: la pregunta és si s'orienten igual o distint que les que en tenen.

**`Access_Plane` i `Portal_Orientation` es desactiven.** Sense cambra no hi ha interior on entrar, i per tant ni pla d'accés ni portal.

**Condició de gating — dues precisions que importen:**

> **Es gategen per `Sys_Chamber`, no per `Sys_Portal`.** Gatejar-los pel portal trencaria una decisió de v16: un portal arrasat continua tenint orientació coneguda, i per això es va deixar obert expressament. Amb `Sys_Chamber`, si hi ha cambra però el portal està arrasat, els dos camps segueixen oberts.
> **Només amb declaració explícita** (`Absent` o `Not applicable`), **mai amb NULL.** NULL significa «encara no s'ha avaluat», i desactivar camps per no haver arribat a la pregunta és una manera de perdre dades sense adonar-se'n.

## **5.2. E2 — criteri general del «no aplicable»**

Hi ha **dos mecanismes** que fan la mateixa faena i convé no confondre'ls: **desactivar un camp** (derivat, no deixa rastre) i **el valor «no aplicable»** (declarat, queda emmagatzemat i s'exporta). Un camp desactivat es veu buit, igual que un camp no avaluat.

> **Si la inaplicabilitat es pot deduir d'un altre camp, es desactiva.**
> **Si depén d'un judici que ningú no pot deduir, cal el valor «no aplicable».**

*Corol·lari: afegir «no aplicable» a llistes que ja es desactiven soles seria soroll. La revisió cas per cas queda oberta a mesura que apareguen; el criteri les resol en un minut cadascuna.*

## **5.3. E3 — `Sys_Interface`**

**S'afig un sisé sistema** al bloc «Sistemes constructius — resoldre primer», amb el domini de grup (`Present` / `Absent` / `Not applicable` / `Not observable`).

**Passa el test que va rebutjar un `Sys_Facade` en v16.** L'argument d'aleshores era que una façana **no pot ser absent mai** — és el pla on passa tota la resta, i un sistema que no pot ser absent no fa el que fan els sistemes. La interfície **sí que pot ser absent**: una estructura d'un sol cos no en té, i això és una absència real i comprovable.

> **Governa només `Interbody_Cornice` (I). `Upper_Crown` (R) continua permanentment actiu.**
> Amb un sol cos no hi ha cornisa entre cossos, **però el coronament continua existint**: l'estructura té una part de dalt igualment. Gatejar-los junts tancaria una pregunta que sí que té resposta.

*Se superposa parcialment amb la regla 39 (cornisa intercòs amb menys de dos cossos), però no la duplica: la regla és un avís **després** de l'error, i el sistema permet declarar-ho **abans**, que és el que es demanava.*

---

# **6. Bloc F — cardinalitat**

## **6.1. Els números**

De 36 registres, **33 tenen el recompte de cossos entrat**. D'eixos, **24 (73 %) tenen més d'un cos**. El cas no és marginal: és el normal.

Ara bé, la immensa majoria són **un cos basal més una cambra**, que és el mausoleu estàndard. Els casos amb pila vertical real són pocs:

- **4 registres amb dues cambres superposades**: EA02, EA12, EA40a, EA40b.
- **4 registres amb dues fases constructives** declarades.

## **6.2. El que no es pot comptar, i per què**

**La xifra que ha de decidir `T_BODIES` —quants registres tenen cossos amb fàbrica realment distinta— no està a la base i no hi pot estar**, perquè cap camp la recull. Eixe és el problema mateix.

I hi ha una cosa pitjor que la falta de dades: **`Stone_Format` i `Stone_Working` només estan emplenats en 4 registres de 36.** Són camps nous de v16 i la segona passada està pendent. **El valor `Mixed` no s'ha usat ni una sola vegada.** Decidir ara com estructurar unes dades encara no recollides és la situació on més fàcil és construir una cosa elegant que després no encaixe.

## **6.3. L'argument que decideix: la unitat d'anàlisi**

Si els cossos passen a ser files pròpies, canvia **la unitat d'anàlisi**. La matriu de `QRY_13` passaria de ~36 files a ~60.

**Sembla un guany i no ho és: dos cossos de la mateixa estructura no són observacions independents.** Els va fer la mateixa gent, el mateix dia, amb la mateixa pedra. Ficar-los com a files separades a un clúster o a una anàlisi de correspondències infla la mostra amb repeticions i produeix agrupacions que reflecteixen **quines estructures tenen més cossos**, no quines pràctiques constructives s'assemblen.

*Això no invalida `T_BODIES` com a eina descriptiva, però sí que vol dir que **no resol l'anàlisi**, i part de l'atractiu era eixe.*

## **6.4. F1 — decisió**

> **La unitat principal d'anàlisi continua sent l'ESTRUCTURA.**
> **`T_BODIES` s'ajorna, però amb un comptador posat.**

No és ajornar el problema una segona vegada: és **instal·lar la mesura que falta** perquè la decisió es prenga amb una xifra i no amb una impressió. Quan la segona passada haja omplit el format i el treball de la pedra dels 36 registres, la xifra existirà com a **subproducte**, sense cap treball addicional. Ara caldria endevinar-la mirant fotos.

## **6.5. F2 — els camps de divergència**

**`Fabric_Divergence`** — TEXT(20). Etiqueta: *Divergència de fàbrica*. A 2.Arq, davall del bloc de maçoneria i morter.

| Valor | Etiqueta |
| --- | --- |
| `Uniform` | Uniforme |
| `Between bodies` | Divergent entre cossos |
| `Within a body` | Divergent dins d'un cos |
| `Both` | Divergent en tots dos sentits |
| `Not observable` | No observable |

**Per què cinc valors i no dos.** La distinció **entre cossos** / **dins d'un cos** és la que decideix la pregunta de fons: si la major part de la divergència resulta ser **dins** d'un mateix cos, `T_BODIES` no resol res — caldria continuar dient «mixt», només un nivell més avall. Sense aquesta distinció, el comptador no compta el que ha de comptar.

**Es desactiva** quan l'estructura té un sol cos: no hi ha res a comparar. *(Criteri E2: deduïble d'un altre camp → gating, no valor «no aplicable».)*

**`Divergence_Location`** — TEXT(20). Etiqueta: *Ubicació de la divergència*. **Només actiu amb `Fabric_Divergence` ∈ {`Between bodies`, `Both`}.**

| Valor | Etiqueta |
| --- | --- |
| `Base vs chamber` | Base contra cambra |
| `Between chambers` | Entre cambres |
| `Both` | En tots dos llocs |

**Per què cal.** Sense això la xifra final seria ambigua. Quinze casos de **base contra cambra** probablement són funcionals —el cos basal no es veu, no es decora i aguanta pes, i que estiga fet més bast diu poc—; quinze casos de **cambra contra cambra** són dues mans, dues fases o dues intencions, que és el que interessa a H01, H03 i H04.

*El detall qualitatiu —si el que canvia és el format, el treball, l'aparell o el morter— va a `Arch_Notes`, que `QRY_18_Notes_Review` ja recupera. Quan arribe el moment de decidir `T_BODIES`, seran quatre o quinze casos i es llegiran d'una tirada.*

## **6.6. F2.3 — frontera amb les fases constructives**

Un canvi de fàbrica a mitja alçària d'una cambra és, sovint, **precisament l'evidència d'una fase**. Frontera fixada, amb la mateixa lògica que `Vertical association`:

> **La divergència de fàbrica és una OBSERVACIÓ**: la pedra canvia, i això es veu.
> **La fase constructiva és una INTERPRETACIÓ**: hi va haver dos moments de construcció, i això s'argumenta.

Pot haver-hi canvi de fàbrica sense afirmar cap fase —un canvi de proveïment, o dos paletes el mateix dia—; quan sí que se sosté la fase, l'argument va a `Phase_Evidence`.

## **6.7. Redefinició de `Mixed`**

`Mixed` no desapareix de `Masonry_Type` ni de `Stone_Working`, però **canvia de sentit i cal escriure-ho**:

> **`Mixed`** diu que **el registre té més d'un valor**.
> **`Fabric_Divergence`** diu **si eixa barreja coincideix amb la divisió en cossos**.

## **6.8. Regles noves**

- **Regla 45** — `Fabric_Divergence` ∈ {`Between bodies`, `Both`} amb `N_Basal_Bodies + N_Chamber_Bodies` < 2.
- **Regla 46** — `Construction_Phases` ≥ 2 amb `Phase_Evidence` buit o `ND`. *Els 4 registres que declaren dues fases no tenen cap evidència entrada; el camp està buit o a `ND` a tota la base. Sense aquesta comprovació, la frontera de 6.6 no es pot sostindre.*

## **6.9. F3 — camps candidats, congelats**

`Stone_Format`, `Stone_Format_Secondary`, `Stone_Working`, `Masonry_Type`, `Masonry_Quality`, `Mortar_*`, `Chamber_Roof` i el criteri de numeració de `Body_No` continuen a nivell de registre. La llista queda anotada per si `T_BODIES` prospera; **no es toca res ara.**

---

# **7. Punts oberts heretats**

## **7.1. Tancats en aquest delta**

**`Chinking_Stones` es manté, sense gradació.** `[DECIDIT]`
*Nota per al TFM: 26 registres el tenen present i **cap** l'té absent. Una variable sense variància no pot aparéixer en cap resultat — no discrimina jaciments, ni tipologies, ni fases. Es conserva com a descriptor de fàbrica, però **no compta com a variable analítica** mentre no aparega algun cas sense falca.*

**Llindar mètric de volada per a la distinció G/I: ja estava tancat.** `[DECIDIT — documentació]`
No és una decisió nova sinó una decisió presa i no escrita al lloc que toca. La distinció és **plataforma** (superfície practicable amb profunditat útil) contra **cornisa** (junta lineal sense profunditat útil), resolta amb quatre proves en ordre: encaixos o fusta conservada → cos construït damunt → voladís progressiu o d'un sol gest → continuïtat al llarg de la façana. **No es fixa cap llindar mètric**: seria inventar-lo. La profunditat es mesura sobre el model a la segona passada, i el llindar s'aplicarà retroactivament quan el corpus diga on és la frontera real.

**Coordenades pròpies a `T_DECORATIONS`: descartades.** `[DECIDIT]`

**Camp d'asimetria de portal: resolt amb `Portal_Position`.** `[DECIDIT]`
TEXT(20) a 2.Arq, gatejat per `Sys_Portal`. Domini: `Centred` / `Off-centre left` / `Off-centre right` / `NA` / `ND`. Etiquetes: *Centrat* / *Descentrat a l'esquerra* / *Descentrat a la dreta*. Mateixa convenció d'observador que el bloc A: **des de davant del pla exposat**.
*Descentrar un portal és una decisió de planificació, no un accident de fàbrica: toca H01 i H04.*

**Connexió horitzontal entre estructures.** `[DECIDIT]`
S'afig `Horizontal association` a `Connection_Type`, simètrica de `Vertical association` (v14): registra una relació d'horitzontalitat **observable** entre dues estructures —alineació sobre la mateixa cornisa natural o repisa, sense contacte físic— sense afirmar-ne el mecanisme. `Confidence` en registra la seguretat.

**Superfície de roca com a suport: incorporada.** Vegeu 2.2. `[DECIDIT]`

## **7.2. Descartats per manca de casos**

**Caracterització de la superfície rocosa a les files ROC.** Es va proposar un camp per a documentar **per què** un tram no és observable, i **es retira**: «no observable» ja és una observació completa, el motiu cap a `Notes`, i `T_DECORATIONS` rep ja cinc camps nous en aquest delta.
*La proposta alternativa —descriure la superfície com a **objecte** en registres d'art rupestre aïllat— és millor plantejada, però **no hi ha cap cas al corpus**: l'únic registre classificat com a PR és EA11, que la secció 3 reclassifica com a estructura. Quan aparega el primer panell real, el disseny es farà mirant-lo.*

**Relació cromàtica entre colors contigus.** Una sola fila bicroma a tota la base. Amb un cas no es dissenya vocabulari.

**Element Y (`Abutted_Mass`).** La condició establida era **tres casos de banqueta adossada**. Encara no han arribat. Continua a `T_ARCH_FEATURES`.

## **7.3. Ajornats a la segona passada**

| Punt | Depén de |
| --- | --- |
| Litologia i potència dels bancs (H02) | Treball de camp i models |
| Confinament fabricat com a camp propi | Casos i models |
| Circulació vertical vs horitzontal (OE3) | És pregunta d'anàlisi, no de registre. No bloqueja res |
| «Traces only» a `Plaster_Extent` | **Cap canvi**: en ús a 3 registres, sense conflicte detectat |

## **7.4. Nota operativa: els grups no s'estan usant**

**`T_GROUPS` està completament buida.** Zero files.

Les vuit entrades de `L_GROUP_TYPE` cobreixen ja el que es demanava en aquest delta —`Vertical alignment` per a la verticalitat, `Ledge cluster` per a l'horitzontalitat compartida, `Circulation network` per als camins, `Supra-structural unit` per als conjunts Xa/Xb—, i cap s'ha fet servir. **No és un problema d'esquema sinó d'entrada de dades**, i queda anotat ací perquè no es dissenye maquinària nova per a una que existeix i està aturada.

---

# **8. Correccions que no són de disseny**

Aquestes **no depenen del delta** i s'han d'aplicar sobre la còpia local com a pedaç `v16b`, amb prioritat sobre la implementació de v17.

## **P1 — 16 files de decoració invisibles al formulari** `[CRÍTIC]`

El `RecordSource` dels dos subformularis de 4.Dec fa `INNER JOIN L_STRUCT_BODY`, i un INNER JOIN elimina les files amb `ID_Struct_Body` nul. **A la còpia local n'hi ha 16 de 56:**

- **15** són propostes automàtiques de la migració v11 des dels booleans `RA_*` (estructures 2, 3, 5, 6, 8, 9, 10, 11, 16, 18, 19, 24, 25, 29), totes amb la nota «Completar posició».
- **1** és la fila 57 (estructura 7, *L-shaped niche*), mig entrada.

Són files que existeixen, compten a les consultes i a les regles 22–23 i 28–29, i **no es poden ni veure ni completar des de la interfície**. La nota que porten demana completar la posició, i completar-la és exactament el que el formulari impedeix.

**Solució:** `LEFT JOIN`, amb les files sense posició assignades a la meitat arquitectònica (on `ND` ja viu), o bé un subformulari de treball dedicat.

## **P2 — tipologia d'EA11**

`DW-S01-EA11` passa de `PR Rock Art` a `Unclassifiable`, amb la composició de registre de la secció 3.3. *(Corregit per l'investigador.)*

## **P3 — registre de prova**

El registre 36 té `Code = 'c'` i cap tipologia. Fila de prova a eliminar.

---

# **9. Inventari de canvis**

## **9.1. `T_STRUCTURES`** (132 → 137 camps)

| Camp | Operació | Pestanya |
| --- | --- | --- |
| `Fabric_Divergence` | NOU TEXT(20) | 2.Arq |
| `Divergence_Location` | NOU TEXT(20) | 2.Arq |
| `Portal_Position` | NOU TEXT(20) | 2.Arq |
| `Sys_Interface` | NOU TEXT(20) | 11.Sist |
| `ID_Material_Status` | TRASLLAT 5.Estat → 7.Mat | 7.Mat |

## **9.2. `T_DECORATIONS`** (10 → 14 camps)

| Camp | Operació |
| --- | --- |
| `Position_Relative` | **ELIMINAT** |
| `Span_Left`, `Span_Above`, `Span_Right`, `Span_Below` | NOUS BYTE |
| `Outline_Geometry` | NOU TEXT(15) |

## **9.3. `T_CONNECTIONS`**

| Element | Operació |
| --- | --- |
| `Chrono_Relation` | REDEFINIT: designa l'estructura anterior |
| Índex únic sobre (`ID_Struct_A`, `ID_Struct_B`) | NOU |
| `Connection_Type` | + `Horizontal association` |

## **9.4. Lookups**

| Taula | Operació |
| --- | --- |
| `L_SUPPORT` | + `Rock dihedral` (mode `Confinement`) |
| `L_SUPPORT` | + `Rock surface` (mode `Substrate`, **mode nou**) |
| `L_DEC_TYPE` | − `Frieze / Greca` |

## **9.5. Formulari**

- 4.Dec: `Position_Relative` fora; quatre trams + geometria del traç a la meitat rupestre; `Body_No` ocult a la meitat rupestre.
- 11.Sist: dues etiquetes corregides; `Sys_Interface` al bloc de sistemes; gating d'`Interbody_Cornice`.
- 2.Arq: `Fabric_Divergence`, `Divergence_Location`, `Portal_Position`; gating d'`Access_Plane` i `Portal_Orientation` per `Sys_Chamber`.
- 5.Estat / 7.Mat: trasllat d'`ID_Material_Status` i gating per `Cultural_Materials_Present`.
- 12.Extra: subformulari de connexions entrants, de només lectura, amb cronologia invertida; normalització d'ordre a `BeforeUpdate`.

## **9.6. Bateria: 39 → 46 regles**

| # | Regla |
| --- | --- |
| **40** | Registre `Rock art panel` amb files `ROC-OVL`/`ROC-PER`/`ROC`, i el recíproc amb `ROC-PAN` |
| **41** | `Outline_Geometry` emplenat en fila no-ROC |
| **42** | Motiu en relleu a `OVL` o `FFL` amb `Relief_Frieze` = 0 (avís, unidireccional) |
| **43** | `Cultural_Materials_Present` = 1 amb els set `Mat_*` a 0 |
| **44** | `ID_Material_Status` emplenat amb `Cultural_Materials_Present` = 0 |
| **45** | `Fabric_Divergence` entre cossos amb menys de dos cossos |
| **46** | `Construction_Phases` ≥ 2 sense `Phase_Evidence` |

## **9.7. Consultes**

- `QRY_13_AX_Pattern_Export`: sense canvis de lògica (F3 congelat).
- `QRY_14_Connections_Edges`: direcció llegida del camp, no de la posició.
- **`QRY_20_RockArt_Span`** (nova): deriva *U invertida*, *envoltant*, *flanquejant* i tram únic des dels quatre camps de tram, amb marca de candidatura quan hi ha algun `9`.
- **`QRY_21_V17_Review`** (nova): llista de treball de la migració — les 7 files amb trams a completar, les 2 connexions a refer, els 16 registres de decoració orfes, i els 4 registres amb fases sense evidència.

---

# **10. Ordre d'implementació**

1. **Pedaç `v16b`** sobre la còpia local: P1, P2, P3. *(No depén de res d'aquest delta.)*
2. `chachapoya_DB_v17.bas` — construcció des de zero.
3. `chachapoya_Form_v17_val.bas` — sobre la base construïda.
4. Parell de transferència `EXPORT_v16` / `IMPORT_v17`, amb els CSV intermedis inspeccionables.
5. Revisió manual: `QRY_21_V17_Review`.
6. Actualització d'`esquema_bbdd_estructures_v17.md`, `tfm_metodologia_bbdd_v17.md` i `manual_us_bbdd_v17.md`.

*Recordatori de la restricció que va fer falta en el delta anterior: qualsevol camp que exigisca judici humà arriba a la transferència com a **NULL**. Un valor transferit sense revisió afirmaria falsament ser un judici.*


---

# **11. Addendum de revisió (post-implementació, sobre la v17 neta)**

*Vuit punts detectats provant la base construïda abans d'importar. No obrin cap decisió del delta: la corregeixen o la completen.*

## **11.1. El valor per defecte 0 no s'estén** `[DECIDIT]`

Es va plantejar posar `Absent` per defecte allà on el valor existisca, per agilitzar l'entrada. **Es descarta per als camps 0/1/9**: trencaria la distinció buit / 0 / 9 que sosté tot l'esquema, perquè un 0 per defecte és un judici que ningú no ha fet i després no hi ha manera de saber quins camps s'han mirat de veritat.

**Substitut que dona la mateixa velocitat sense el fals judici:** un botó al capçal del formulari que posa a 0 els camps 0/1/9 **buits i habilitats de la pestanya activa**, després de comptar-los i demanar confirmació. S'usa en acabar de mirar una pestanya, i llavors el 0 sí que és un judici.

*Implementació: identifica els camps pel `RowSource` del combo i no per una llista de noms, que quedaria desfasada al pròxim delta.*

## **11.2. `Phase_Evidence`: vocabulari ampliat** `[DECIDIT]`

La llista v16 (C14 / Estratigrafia / Superposició / Morter / ND) **no tenia valor per a les dues observacions més freqüents en camp** —la junta vertical adossada i el canvi de fàbrica—, de manera que la regla 46 es podia complir només formalment.

| Valor | Etiqueta |
| --- | --- |
| `Abutted vertical joint` | Junta vertical adossada |
| `Superposition` | Superposició |
| `Fabric change` | Canvi de fàbrica (format, treball o aparell) |
| `Mortar difference` | Diferència de morter |
| `Blocked or altered opening` | Obertura tapiada o modificada |
| `Added mass or annex` | Massa afegida o annex |
| `Radiocarbon` | C14 |
| `Stratigraphy` | Estratigrafia |
| `ND` | Indeterminada |

## **11.3. `Recessed_Frame` a «Morfologia general»** `[DECIDIT]`

Estava al bloc de fases constructives, amb el qual no té relació. És un qualificador del pla de façana (rev. 7) i va amb la morfologia.

## **11.4. `Fabric`: un camp en lloc de dos, i els valors corregits** `[CORRECCIÓ de F2]`

`Fabric_Divergence` + `Divergence_Location` es fonen en **`Fabric`** TEXT(25).

**I els valors del disseny original se solapaven.** Estaven definits per **on** hi ha divergència, i un cos plural per dins també és plural respecte del veí, de manera que «dins d'un cos» implicava «entre cossos». El criteri correcte no és on, sinó **si la divisió coincideix amb els cossos o els travessa**:

| Valor | Etiqueta | Vol dir |
| --- | --- | --- |
| `Single` | Única | Una sola fàbrica a tot el registre |
| `Between bodies only` | Múltiple, coincident amb els cossos | Cada cos uniforme per dins, però difereixen entre ells |
| `Within a body` | Múltiple, dins d'un cos | Almenys un cos té més d'una fàbrica per dins |
| `Not observable` | No observable | — |

*Ara són excloents, i mantenen el que justificava el camp: `Between bodies only` és exactament el cas que `T_BODIES` resoldria, i `Within a body` el cas que no resoldria.*

El camp de localització es descarta: amb dos o tres casos previstos, el detall va a `Arch_Notes`.

## **11.5. `Stone_Format`: `Irregular blocks` → `Irregular stones`** `[DECIDIT]`

Etiqueta «Pedres irregulars». Canvia també el valor emmagatzemat, perquè la base està buida i és el moment; una peça irregular no és un bloc.

## **11.6. El gating nou tanca només amb «No aplicable»** `[CORRECCIÓ de E1]`

`Access_Plane`, `Portal_Orientation` i `Portal_Position` apareixien bloquejats en tot registre nou. **La causa era de disseny meu**: els vaig tancar amb `Absent`, i `Absent` és el **valor per defecte** dels sistemes, de manera que no pot valdre com a declaració de l'usuari.

> **Regla general per a tot el gating nou de v17: tanca amb `Not applicable`, mai amb `Absent`.** `No aplicable` sempre és deliberat; `Absent` pot ser només el valor amb què va nàixer el registre.

*Això no toca el gating de nivell 2 heretat de v12 sobre els components A–X, on `Absent` sí que tanca: allí és el flux previst —es resol el sistema abans que els components— i està assumit des de fa versions.*

## **11.7. La 9.Metr deixa de suposar mausoleu** `[DECIDIT]`

L'esquema mètric val per a mausoleus i cambres, però no per a coves, terrasses, nínxols o panells. La discriminació **no es fa per tipologia sinó pels sistemes**, que és on ja viu la informació:

| Bloc | Es tanca quan |
| --- | --- |
| Dimensions de l'obertura | `Sys_Portal` = `Not applicable` |
| Volumetria i superfície | `Sys_Chamber` = `Not applicable` |
| Dimensions, mètrica del suport, coordenades | Mai |

*No toca la decisió v16 que les dimensions d'obertura no depenen de llindar, brancals ni dintell: una obertura sense elements diferenciats continua tenint amplària.*

## **11.8. El vocabulari de «plataforma»** `[DECIDIT]`

La paraula servia tres objectes: l'element H, la tipologia `EA-PLA-V` i la tipologia `EA-PLA-R`. **Els dos primers són la mateixa cosa a dues escales** —una superfície volada sobre el buit, sola o dins d'una estructura més complexa— i conserven el terme. El tercer és un objecte distint i passa a **`EA-TER Ledge Terrace` / «EA-TER Terrassa en repisa»**.

*Es descarta «cos basal» com a substitut: ja el porten `N_Basal_Bodies` i la posició decorativa `BAS`, de manera que nomenaria alhora una unitat de registre i un component d'una unitat de registre — la confusió que aquest delta ha estat desfent. A més, la definició de la tipologia diu «funció: trànsit o base de mausoleu», i un nom que afirmara «basal» resoldria per decret el que `Platform_Function` ha de resoldre cas per cas. «Banqueta» estava ocupada pel massís adossat de 8bis.8.*

| Terme | Reservat per a |
| --- | --- |
| **Plataforma** | Superfície volada sobre el buit: element H i `EA-PLA-V` |
| **Terrassa** | Massa construïda que anivella una repisa: `EA-TER` |
| **Banqueta** | Massís adossat que no sosté res: `T_ARCH_FEATURES` |
| **Cos basal** | Component d'una estructura: `N_Basal_Bodies` i posició `BAS` |

## **11.9. Amplades i col·lisions del formulari** `[DECIDIT]`

Capçalera de secció i camps compartien fila al bloc de fàbrica (error d'implementació). Corregit, i amb ell una passada general: amplada d'etiqueta 2600→3000, de control 2200→2600, i inici de la segona columna 5600→6200. Les etiquetes llargues es tallaven i els combos no mostraven els valors sencers.

*Migració: `EA-PLA-R Ledge Platform` → `EA-TER Ledge Terrace` la fa l'importador automàticament. És un canvi de nom i no de definició, de manera que no genera feina de revisió.*
