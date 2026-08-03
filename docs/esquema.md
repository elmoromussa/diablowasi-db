**ESQUEMA DE LA BASE DE DADES ARQUEOLÒGICA**

*La Petaca i Diablo Wasi (Leymebamba, Amazonas, Perú)*

chachapoya_DB_v9.bas + chachapoya_Form_v9_val.bas

Sub BuildDB() + Sub BuildForm() | Microsoft Access JET SQL | Esteve Ribera Torró

*Versió 9 del document — actualitzada segons el codi v9 de la BD (agost 2026). La numeració dels quatre fitxers del projecte queda alineada a v9.*

# **1. Resum general**

| **Element** | **Valor** |
| --- | --- |
| Taules principals | T_STRUCTURES (131 camps), T_DATING, T_INDIVIDUALS, T_GROUPS, T_DECORATIONS, T_ARCH_FEATURES, T_CONNECTIONS |
| Taules lookup | L_SITES, L_SECTORS, L_TYPOLOGY, L_SUPPORT, L_STATUS, L_MATERIAL_STATUS, L_VOL_METHOD, L_COORD_METHOD, L_GROUP_TYPE, L_CAMPAIGN, L_STRUCT_BODY, L_DEC_TYPE |
| Total taules | 19 |
| Total relacions | 21 (inclou l'autoreferenciant de T_STRUCTURES, les dues de T_CONNECTIONS i la del suport secundari) |
| Consultes SQL | 16 (QRY_01 a QRY_16; QRY_16_Validation_Check nova en v8) |
| Formulari | F_STRUCTURES - 12 pestanyes + subformularis F_DECORATIONS i F_ARCH_FEATURES. Etiquetes UI en valencià; valors emmagatzemats en anglés. 62 camps amb combos de domini 0/1/9 |
| Idioma BD | Anglés (noms de taules, camps i valors lookup). UI del formulari en valencià. Capçaleres de subformulari via etiquetes adjuntes; captions DAO com a reforç |
| Motor | Microsoft Access JET SQL / ACE │ cp1252 |
| Jaciments | La Petaca (WGS84: Lat -6.8311, Lon -77.8084) │ Diablo Wasi (Lat -6.8475, Lon -77.8154) |

**Canvis respecte de la versió 5 del document (que descrivia la BD v4):**

**(a) El domini 0/1/9 s'estén a tots els camps observacionals — 62 camps BYTE.** El principi que ho ordena: **YESNO és l'únic tipus d'Access que no admet Null**, de manera que un FALSE confon «he verificat que no hi és» amb «no ho he pogut observar». Tots els camps SINGLE, INTEGER i TEXT ja podien expressar ND deixant-se buits; el problema era exclusiu dels booleans. El criteri d'auditoria ha estat una pregunta única per camp: *quan val FALSE, pot voler dir «no ho he pogut mirar»?* Si sí, passa a BYTE amb 0 = Absent (decisió constructiva verificada), 1 = Present, 9 = ND / no observable, per defecte 9.

Camps convertits en v7, a més del vocabulari A–X ja convertit en v4: morfologia (Buttresses, Wooden_Stakes, Support_Modified; Natural_Roof també, però eliminat en v8 — punt j), decoració en baix relleu (7), art rupestre (6), alteracions (4), bioarqueologia (7) i materials culturals (7).

Dos blocs mereixen justificació explícita. **Bioarqueologia i materials culturals** són els casos on el FALSE hauria estat més sovint fals: en necròpolis de penya-segat l'interior s'observa per una obertura, sovint des d'un dron, i es pot veure un farcell funerari sense poder determinar la posició flexionada. **Decoració** és el cas analíticament més sensible: aquests camps alimenten QRY_02, la font del khi-quadrat LP vs DW. Si la conservació de façanes difereix entre jaciments —i difereix—, sense el 9 la prova mesuraria preservació diferencial i el resultat es llegiria com a pràctica decorativa diferencial.

**Es mantenen com a YESNO, deliberadament: C14 i ChaXR_Documented.** No són observacions sobre l'estructura sinó metadades sobre el corpus propi: sempre se sap si hi ha una datació o si s'ha publicat una escena. El seu FALSE no és mai ambigu.

Nota sobre la càrrega d'entrada: la conversió **redueix** feina de camp. Amb booleans calia recórrer conscientment cada casella per confirmar que el FALSE era real; amb el defecte 9 només es toca allò que s'ha observat.

**(b) Tractaments superficials: operació separada de substrat.** Els quatre camps antics (Plastered, Plaster_Color, Rock_Painting, Rock_Paint_Color) barrejaven operació i suport. Ara són set: Plaster_Present, Plaster_Color, Plaster_Extent, Pigment_Present, **Pigment_Substrate**, Pigment_Color, Pigment_Extent. El camp clau és Pigment_Substrate (Plaster / Masonry stone / Bedrock / Mixed / ND), i té tres justificacions:

- **Cadena operativa.** Pintar sobre lluït exigeix una operació preparatòria addicional; pintar sobre la pedra significa que el parament de maçoneria *era* la superfície acabada prevista. Són dues seqüències amb inversió de treball distinta que, sense aquest camp, col·lapsen en un mateix «pintat» (H04).
- **H01.** Si el pigment sobre pedra s'aplica selectivament a elements arquitectònics concrets, això és evidència d'esquema policrom integrat amb la maçoneria i no d'un afegit posterior. Ho captura Pigment_Extent.
- **Biaix de conservació.** El lluït es desprén molt més fàcilment que el pigment absorbit en un gres porós. Sense registrar el substrat, una preservació diferencial es llegiria com una pràctica diferencial.

T_DECORATIONS guanya el camp **Substrate** per als casos mixtos, freqüents: fris pintat sobre lluït i brancals pintats sobre la pedra en la mateixa estructura. A nivell d'estructura es marca `Mixed`; el detall es resol per posició al subformulari, seguint el patró dual ja establit (booleans de resum + registre detallat).

Frontera amb l'art rupestre: amb Pigment_Substrate = Bedrock cal criteri explícit. **Tractament cromàtic del suport de l'estructura → Pigment; motiu figuratiu o geomètric autònom sobre el penyal → Rock_Art.**

**(c) Base documental i observabilitat — Doc_Basis, Facade_Observability, Interior_Observability + QRY_15.** El domini 0/1/9 registra *on* són els buits; aquests tres camps els fan interpretables. Permeten restringir l'anàlisi en R («només estructures amb façana completa») i **defensar la mostra en lloc de disculpar-la**, cosa que és directament OE1.

**(d) ID_Support_Secondary.** Els suports compostos es registren com a parella ordenada (dominant + secundari, tots dos FK a L_SUPPORT) en lloc del valor opac «Combined». **La llista L_SUPPORT es manté intacta**; «Combined» queda reservat per a suports amb tres components o no descomponibles. Motiu del canvi: «Combined» era analíticament mut — impedia saber *què* es combinava.

**(e) Support_Morphology reduït a descriptors morfològics purs en v7 i eliminat del tot en v8** (punt i). Abans repetia gairebé el mateix vocabulari que L_SUPPORT (cavitats per mida, fissura), i el solapament ja havia produït una contradicció en el registre de prova (suport = Medium cavity amb morfologia = Fissure). Ara **L_SUPPORT conserva la classe** de suport i Support_Morphology en descriu **la forma**.

**(f) Rear_Wall_Built (BYTE) → Rear_Wall_Type TEXT(20)** (Built masonry / Natural bedrock / Mixed / ND). Sota el domini 0/1/9 el valor 0 significa «absent» a tot arreu; ací significava «roca natural», que és un valor, no una absència. L'element W conserva la presència a Rear_Wall.

**(g) Corbel_Material → Platform_Surface_Material.** E és fusta i G és pedra per definició, de manera que el camp antic era parcialment derivable i admetia contradiccions (valor «Stone» amb E=1 i G=0). Ara registra només el material de la superfície de la plataforma H, que és l'única part no derivable.

**(h) N_Floors → N_Bodies** i **L_STRUCT_BODY.Body_Level → Body_No**, completant la separació terminològica iniciada en v4: «nivell» = N0/N1/Superior (nivells constructius del vocabulari); «cos» = pisos superposats. El nom Body_Level barrejava precisament els dos termes que s'havien separat.

**(i) Redundàncies eliminades i regles implementades en v8.** Tres dels quatre punts que la v7 deixava oberts queden resolts:

- **Support_Morphology eliminat.** Era redundant amb la parella ID_Support + ID_Support_Secondary, que ja descriu la classe i la composició del suport. El solapament havia produït una contradicció al registre de prova. QRY_12 es refà sobre la parella. *Pèrdua acceptada:* el perfil de la repisa (pla, còncau, esglaonat) ja no es codifica; si més endavant es vol recuperar com a variable, el lloc natural és T_ARCH_FEATURES, no un camp de T_STRUCTURES.
- **Access_Orientation eliminat i fos en Facade_Orientation.** En una estructura de penya-segat el pla de la façana i la direcció d'accés són la mateixa variable. Una sola entrada al formulari, cap divergència que arbitrar.
- **Natural_Roof eliminat, absorbit per Chamber_Roof_Type.** Vegeu el punt (j).
- **Regla del col·lapse implementada** com a QRY_16 (punt k).

**(j) Coberta de la cambra: element X + tipus.** La coexistència de Natural_Roof i Chamber_Roof no era pròpiament una contradicció —una cavitat amb sostre rocós i una coberta construïda al front poden conviure— però sí una duplicació: dos camps per a una sola pregunta, *com es tanca la cambra per dalt?* La solució adoptada replica exactament el patró ja establit per al mur posterior (W + Rear_Wall_Type):

- **Chamber_Roof (X)**, BYTE 0/1/9: la cambra està tancada per dalt, amb la solució que siga.
- **Chamber_Roof_Type**, TEXT(25): Natural bedrock / Built masonry / Built timber and slabs / Mixed / ND.

*Advertència sobre la semàntica de la matriu.* Amb aquest canvi, AX_X passa a significar «cambra tancada per dalt», cosa que pot ser un donat geològic i no una decisió constructiva. Això és acceptable —i coherent amb AX_W, que ja funcionava així— sempre que l'anàlisi ho tinga en compte. QRY_13 exporta Chamber_Roof_Type precisament perquè en R es puga derivar una variant «només construït» d'X quan l'anàlisi tracte específicament de decisions constructives:

```r
ax$AX_X_built <- ifelse(ax$Chamber_Roof_Type %in% c("Built masonry",
                        "Built timber and slabs","Mixed"), 1,
                 ifelse(ax$AX_X == 9, 9, 0))
```

Val la pena notar que triar una cavitat *perquè* té sostre rocós és també una decisió, i pertinent per a H02; simplement no és del mateix ordre que bastir una coberta.

**(k) QRY_16_Validation_Check — bateria de validació.** Set regles de coherència en una consulta UNION. Un resultat buit significa que el corpus és coherent. És deliberadament un **informe i no una restricció de taula**: una regla dura impediria registrar una absència genuïnament observada en una estructura parcialment col·lapsada, i una regla de validació a nivell de TableDef hauria de codificar en dur l'ID numèric de l'estat «Collapsed», cosa fràgil. Les regles:

| # | Regla | Motiu |
| --- | --- | --- |
| 1 | Element A–X amb valor 0 en una estructura amb ID_Arch_Status = Collapsed | En una estructura col·lapsada l'absència no es pot verificar: ha de ser 9 |
| 2 | Pigment_Present = 1 amb Pigment_Substrate buit o ND | El substrat és precisament la variable que justifica el desdoblament |
| 3 | Pigment_Substrate = Plaster amb Plaster_Present = 0 | Contradicció directa |
| 4 | Chamber_Roof = 1 amb Chamber_Roof_Type buit o ND | Sense el tipus, X no distingeix roca natural d'obra |
| 5 | Rear_Wall = 1 amb Rear_Wall_Type buit o ND | Idem per a W |
| 6 | Corbelled_Platform (H) = 1 sense cap component E, F ni G | El sistema E+F+G → H exigeix almenys un suport |
| 7 | Access_Opening (P) = 1 sense N, O ni Q | El sistema N+O+Q → P exigeix almenys un component |
| 8 | Lost_Body_Evidence informada amb elements de zona superior (R, S, T, U, X) codificats 0 | Els elements d'un cos desaparegut no són observables: han de ser 9 |
| 9 | N_Chamber_Bodies > 0 amb Access_Opening = 0 | Un cos només compta com a N1 si té (o tenia) obertura d'accés |

**(l) Nivells constructius: categories fixes, repetició en comptadors.** L'observació de partida: hi ha mausoleus amb **dues masses basals sota una sola cambra**, i n'hi ha amb cossos desapareguts. La temptació és relaxar la definició de nivell i obrir N2, N3, etc.

**No s'ha fet, i deliberadament.** N0 / N1 / Superior no són comptadors sinó **categories funcionals**: què fa la maçoneria (alçar i suportar / contindre la cambra i l'obertura / tancar i protegir). Si s'obri «N2» com a categoria nova, N2 significaria «segona massa basal» en una estructura i «segona cambra» en una altra, i la matriu A–X deixaria de ser comparable — que és precisament la propietat que la fa útil.

**Criteri operatiu per a assignar un cos a un nivell** (replicable per un tercer):

> Un cos és **N1** si conté (o contenia) una obertura d'accés. Si no en té, és **N0**, per alta i acurada que siga la seua fàbrica.

El que sí que estava infraespecificat era el **recompte**. `N_Bodies` («nombre de cossos superposats») no permetia respondre si una estructura amb dues masses basals sota una cambra valia 3 o 1: el camp era ambigu i, per tant, no agregable. Es desdobla en dos comptadors, separant l'eix de la categoria (fix, tres valors) de l'eix de la repetició (lliure):

| Camp | Contingut |
| --- | --- |
| `N_Basal_Bodies` | Masses superposades sense obertura (N0) |
| `N_Chamber_Bodies` | Cambres funeràries superposades (N1) |

Dos cossos basals amb una cambra → `2, 1`. Una torre de dues cambres → `1, 2`. El cas de «tres nivells» es codifica sense inventar categories.

**La matriu A–X es manté plana**, una fila per estructura. La pregunta que respon la coocurrència és «aquesta estructura fa servir l'element K en algun punt?», no «en quin cos». Una matriu per cos multiplicaria les files, trencaria la comparabilitat amb la classificació d'estructures de Toyne i Anzellini (2017) i no tindria prou casos amb múltiples cossos per a sostindre's estadísticament.

**(m) Evidència de cos perdut — `Lost_Body_Evidence`.** Valors: *Pigment on bedrock / Truncated walls / Empty beam sockets / Corbels into void / Detached debris / None / ND*.

El cas diagnòstic són les **bandes verticals de pigment aplicades directament sobre la roca per damunt del cos conservat**, alineades amb els brancals inferiors: no són decoració del penyal sinó el fantasma d'un cos desaparegut — la pintura ha sobreviscut al parament que la sostenia. Notablement, això ja quedava parcialment capturat per `Pigment_Substrate = Bedrock`, encara que no s'havia previst amb aquesta funció.

Dues conseqüències que van més enllà del registre descriptiu:

- **Serveix H03** (saturació acumulativa): permet distingir una estructura d'un cos d'una estructura de dos cossos mutilada, que fins ara es confonien.
- **Fixa quan els elements A–X han de valer 9 i no 0.** Si hi ha evidència de cos perdut, els elements que hi pertanyien són no observables, no absents. És la mateixa lògica de la regla del col·lapse, aplicada **verticalment**. Ho comprova la regla 8 de QRY_16.

*S'ha descartat un comptador `N_Lost_Bodies`: a la pràctica el valor seria sempre 0 o 1, i el camp d'evidència ja el determina.*

**Punt que continua obert:** la triple codificació dels motius decoratius. Un fris de zigzag s'enregistra a Relief_Frieze (M), a Dec_Zigzag i a una fila de T_DECORATIONS. Cada un respon una pregunta distinta i no se n'elimina cap; la regla de treball és que *tota fila de T_DECORATIONS implica el Dec_\* corresponent a 1*. No s'ha afegit a QRY_16 perquè la comprovació genèrica en SQL exigiria una regla per motiu.

# **2. T_STRUCTURES (131 camps)**

## **2.1. Identificació (8)**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| ID | COUNTER PK | Clau primària autonumèrica |
| Code | TEXT(20) NN | Codi PALP: [SITE][SECTOR]-[TIPTYPE][NUM] - ex: DWS1-EF01, PTC-SS-EF18 |
| ID_Sector | LONG NN | FK -> L_SECTORS |
| ID_Typology | LONG | FK -> L_TYPOLOGY |
| ID_Support | LONG | FK -> L_SUPPORT. Classe de suport geomorfològic dominant |
| ID_Support_Secondary | LONG | NOU v7. FK -> L_SUPPORT. Segon component en suports compostos (parella ordenada en lloc del valor «Combined») |
| ID_Parent | LONG | FK autoreferenciant -> T_STRUCTURES.ID. Contenció física |
| ID_Group | LONG | FK -> T_GROUPS. Agrupació funcional (alineament, xarxa) |

## **2.2. Morfologia i dimensions (15)**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| N_Basal_Bodies | INTEGER | NOU v9. Masses superposades **sense obertura** (nivell N0) |
| N_Chamber_Bodies | INTEGER | NOU v9. Cambres funeràries superposades (nivell N1) |
| Lost_Body_Evidence | TEXT(40) | NOU v9. Pigment on bedrock / Truncated walls / Empty beam sockets / Corbels into void / Detached debris / None / ND |
| Floor_Plan | TEXT(20) | Planta (llista al formulari): Rectangular / Sub-rectangular / Square / Circular / Sub-circular / Trapezoidal / Irregular / ND |
| N_Built_Walls | INTEGER | Nombre de murs construïts (0-4) |
| Length_m | SINGLE | Llarg exterior (m) |
| Width_m | SINGLE | Ample exterior (m) |
| Height_m | SINGLE | Alçada total de l'estructura (m) |
| Height_Above_Base_m | SINGLE | Cota de posició sobre la base del faralló (m). RENOMENAT (abans Approx_Height_m): és una cota, no una dimensió |
| Dim_Method | TEXT(30) | NOU v4. Mètode de mesura de les dimensions: Photogrammetric model / Tape measure / Laser / Estimation / ND |
| Opening_Width_cm | SINGLE | NOU v4. Amplada de l'obertura d'accés (cm) |
| Opening_Height_cm | SINGLE | NOU v4. Alçada de l'obertura d'accés (cm) |
| Lintel | TEXT(20) | Material del dintell (**element Q**; llista al formulari): Stone / Wood / Mixed / Absent / ND. AX_Q es deriva d'aquest camp a QRY_13 |
| Buttresses | BYTE | 0/1/9. Contraforts a la façana |
| Wooden_Stakes | BYTE | 0/1/9. Estaques o pals de fusta verticals (ancoratge). La fusta té un biaix de conservació més fort que la pedra |

## **2.2b. Detall del suport geològic (3) — H02: geologia com a factor determinant**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| Support_Width_cm | SINGLE | Amplària de la repisa o suport (cm). Mesura directa o del model 3D |
| Support_Depth_cm | SINGLE | Profunditat de la repisa o cavitat (cm) |
| Support_Modified | BYTE | 0/1/9. Modificació antròpica del suport natural (retall, anivellament). Sovint queda oculta darrere de la maçoneria: no és que no hi siga, és que no es veu |

*Aquests camps quantifiquen la relació entre les dimensions del suport natural i les decisions constructives, nucli empíric de la hipòtesi H02. La classe i la composició del suport es descriuen amb la parella ID_Support + ID_Support_Secondary (secció 2.1): en v8 s'elimina Support_Morphology, que repetia aquesta informació i havia produït una contradicció al registre de prova (suport = Medium cavity amb morfologia = Fissure).*

## **2.2c. Maçoneria i morter (6) — H01/H04, cf. Toyne i Anzellini 2017**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| Masonry_Quality | TEXT(20) | Qualitat d'execució (llista al formulari): Good / Moderate / Poor / ND |
| Masonry_Type | TEXT(30) | Tipus d'aparell (llista al formulari): Well-coursed / Irregular-coursed / Uncoursed / Mixed / ND |
| Mortar_Present | BYTE | NOU v4. 0=Absent (fàbrica en sec) / 1=Present / 9=ND. Per defecte 9 |
| Mortar_Type | TEXT(30) | NOU v4 (llista al formulari): Mud / Mud with gravel / Mud with organics / None dry-laid / ND |
| Chinking_Stones | BYTE | NOU v4. Ripio / pedres de falca entre carreus: 0/1/9. Discriminador de qualitat de fàbrica (T&A 2017) |
| Mortar_Notes | TEXT(150) | NOU v4. Notes sobre morter i juntes |

*El morter és evidència potencial de fases (valor «Mortar» de Phase_Evidence) i atribut de la fàbrica. No confondre amb Plaster_Present (revoc, acabat superficial: operació distinta de la cadena operativa).*

## **2.3. Sistemes A-X — Nivell 0 / basament (10) — elements A-H**

**Convenció de domini per a tots els camps BYTE del vocabulari A-X: 0 = Absent (decisió constructiva) / 1 = Present / 9 = ND, no observable (per defecte).**

| **Camp** | **Tipus** | **Elem.** | **Descripció** |
| --- | --- | --- | --- |
| Embedded_Base_Beams | BYTE | **A** | Jàcenes horitzontals empotrades dins la maçoneria del basament |
| Base_Level | BYTE | **B** | Basament com a element constructiu diferenciat (pòdium) |
| Decorative_Socle | BYTE | **C** | Tractament decoratiu del basament (sòcol decoratiu) |
| Tie_Walls | BYTE | **D** | Murets perpendiculars al faralló (ancoratge/compartimentació) |
| Timber_Brackets | BYTE | **E** | Mènsules de fusta empotrades a la roca (component de H). Criteri operatiu: element perpendicular a la façana, encastat, en voladís |
| Timber_Bracket_Count | INTEGER | **E** | Nombre de mènsules visibles |
| Transverse_Beams | BYTE | **F** | Bigues transversals (component de H). Criteri operatiu: element paral·lel a la façana, salvant llum entre suports |
| Corbelled_Courses | BYTE | **G** | Filades de pedra en voladís creixent (component de H, variant lítia) |
| Platform_Surface_Material | TEXT(20) | **H** | RENOMENAT v7 (abans Corbel_Material). Material de la superfície de la plataforma: Timber / Stone / Mixed / ND. E és fusta i G és pedra per definició, de manera que el camp antic era parcialment derivable i admetia contradiccions |
| Corbelled_Platform | BYTE | **H** | Plataforma volada d'accés. SISTEMA: E+F+G -> H. Nivell: interfície N0/N1 (tanca N0, precondició de N1) |

## **2.3b. Interfície N0/N1 (2) — element I**

| **Camp** | **Tipus** | **Elem.** | **Descripció** |
| --- | --- | --- | --- |
| Interbody_Cornice | BYTE | **I** | Cornisa intercòs entre cossos superposats. RENOMENAT (abans Interlevel_Cornice): «cos» per als pisos, «nivell» per a N0/N1/Sup |
| Interbody_Cornice_Material | TEXT(20) | **I** | Material (llista al formulari): Stone slabs / Wooden beams / Mixed / ND. RENOMENAT (abans Cornice_Material) |

## **2.4. Sistemes A-X — Nivell 1 / cos principal (12) — elements J-P, R, V, W**

| **Camp** | **Tipus** | **Elem.** | **Descripció** |
| --- | --- | --- | --- |
| Corner_Quoins | BYTE | **J** | Cantoneres: pedres de major mida disposades verticalment als angles |
| Structural_Pilasters | BYTE | **K** | Pilastres estructurals que flanquegen la façana (altura completa) |
| Lateral_Wall_Faces | BYTE | **L** | NOU v4. Paraments laterals del cos principal (fora del marc del portal) |
| Relief_Frieze | BYTE | **M** | NOU v4. Fris decoratiu en baix relleu (tractament de L). Substitueix el booleà Dec_Frieze |
| Sill | BYTE | **N** | Llindar / threshold (component de P) |
| Jambs | BYTE | **O** | Brancals: elements verticals formals del portal (component de P) |
| Access_Opening | BYTE | **P** | Obertura d'accés present. SISTEMA: N+O+Q -> P |
| Recessed_Portal | BYTE | **-** | Portal retranquejat de façana (qualificador de P) |
| Upper_Crown | BYTE | **R** | Coronament superior / coping |
| Lateral_Walls | BYTE | **V** | Murs laterals que donen profunditat a la cambra funerària |
| Rear_Wall | BYTE | **W** | Mur posterior present |
| Rear_Wall_Type | TEXT(20) | **W** | RENOMENAT v7 (abans Rear_Wall_Built): Built masonry / Natural bedrock / Mixed / ND. Sota el domini 0/1/9 el valor 0 significa «absent» a tot arreu; ací significava «roca natural», que és un valor, no una absència |

## **2.5. Sistemes A-X — Zona superior (5) — elements S, T, U, X**

| **Camp** | **Tipus** | **Elem.** | **Descripció** |
| --- | --- | --- | --- |
| Eave_Beam | BYTE | **S** | Biga de suport del ràfec (component de U; pot ser fusta) |
| Eave_Surface | BYTE | **T** | Superfície del ràfec (component de U; sempre pedra - lloses) |
| Eave | BYTE | **U** | Ràfec-voladís / visera. SISTEMA: S+T -> U. Sempre pedra |
| Chamber_Roof | BYTE | **X** | 0/1/9. Cambra tancada per dalt (qualsevol solució). Distint d'Eave (U) |
| Chamber_Roof_Type | TEXT(25) | **X** | NOU v8: Natural bedrock / Built masonry / Built timber and slabs / Mixed / ND. Absorbeix l'antic camp Natural_Roof |

## **2.6. Tractaments superficials (7) — v7: operació separada de substrat**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| Plaster_Present | BYTE | 0/1/9. Revoc / lluït present |
| Plaster_Color | TEXT(20) | White / Cream / Red / Ochre / Grey / ND |
| Plaster_Extent | TEXT(20) | Full facade / Partial / Traces only / ND |
| Pigment_Present | BYTE | 0/1/9. Pigment aplicat present |
| **Pigment_Substrate** | TEXT(20) | **Camp clau v7**: Plaster / Masonry stone / Bedrock / Mixed / ND. Distingeix pintar sobre revoc preparat de pintar directament sobre el parament de maçoneria |
| Pigment_Color | TEXT(20) | Red / White / Both / Ochre / ND |
| Pigment_Extent | TEXT(20) | Whole facade / Architectural elements / Decorative motifs / Traces / ND |

*Revocar i pintar són dues operacions distintes de la cadena operativa. Quan el pigment s'aplica directament sobre la pedra, el parament de maçoneria era la superfície acabada prevista: no hi ha operació preparatòria. Els casos mixtos dins d'una mateixa estructura es marquen `Mixed` ací i es detallen per posició al camp Substrate de T_DECORATIONS.*

## **2.6b. Paisatge i orientació (2) — observacional; anàlisi QGIS posterior**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| Facade_Orientation | TEXT(5) | Orientació de la façana i de l'accés: N / NE / E / SE / S / SW / W / NW / ND. En v8 absorbeix Access_Orientation: en una estructura de penya-segat són la mateixa variable |
| Visibility_Valley | TEXT(10) | Visibilitat des del fons de la vall (estimació de camp): High / Medium / Low / ND |

*Registre observacional de camp que prepara l'anàlisi formal de visibilitat i orientació en QGIS (H06: visibilitat i marcatge territorial).*

## **2.7. Decoració - camps de resum, BYTE 0/1/9 (13) - detall a T_DECORATIONS**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| Dec_Square_Niche | BYTE | Nínxol quadrat en sèrie (hornacinas cuadradas) |
| Dec_Relief_T | BYTE | Relleu en T |
| Dec_Relief_T_Inv | BYTE | Relleu en T invertida |
| Dec_Relief_L | BYTE | Relleu en L |
| Dec_Relief_L_Inv | BYTE | Relleu en L invertida |
| Dec_Zigzag | BYTE | Motiu en zig-zag o chevron |
| Dec_Stepped | BYTE | Motiu escalonat (inusual a DW i LP; documentat a DWS1) |
| Rock_Art | BYTE | Pintura rupestre present (associada a l'estructura) |
| RA_Anthropomorphic | BYTE | Motiu antropomorf a la pintura rupestre |
| RA_Zoomorphic | BYTE | Motiu zoomorf |
| RA_Geometric | BYTE | Motiu geomètric |
| RA_Abstract | BYTE | Motiu abstracte |
| RA_Decap_Scene | BYTE | Escena de decapitació documentada |

*Convertits a BYTE 0/1/9 en v7. Aquests camps alimenten QRY_02, la font del khi-quadrat LP vs DW: si la conservació de façanes difereix entre jaciments —i difereix—, sense el valor 9 la prova mesuraria preservació diferencial i el resultat es llegiria com a pràctica decorativa diferencial.*

*Dec_Frieze es va eliminar en v4: el fris és l'element M del vocabulari (Relief_Frieze, secció 2.4). Els camps RA_\* qualifiquen l'art rupestre, no el baix relleu: al formulari estan sota la capçalera «Art rupestre associat (sobre el penyal)».*

## **2.8. Estat de conservació (6)**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| ID_Arch_Status | LONG | FK -> L_STATUS. Estat de l'estructura construïda (Good/Fair/Pre-collapse/Collapsed/ND) |
| ID_Material_Status | LONG | FK -> L_MATERIAL_STATUS. Grau de preservació dels vestigis mobles (Good/Fair/Poor/Absent/ND) |
| Looting | BYTE | 0/1/9. Evidència de saqueig (causa, independent del grau de conservació) |
| Fire_Damage | BYTE | 0/1/9. Evidència d'incendi |
| Animal_Activity | BYTE | 0/1/9. Activitat animal documentada |
| Modern_Access | BYTE | 0/1/9. Evidència d'accés modern no autoritzat |

*Al formulari, les etiquetes són «Estat estructura» i «Estat vestigis mobles» (sense sigles (I)/(M), reservades al vocabulari A-X). Les quatre alteracions passen a BYTE en v7: si no hi ha hagut accés a l'estructura, no s'ha pogut buscar l'evidència.*

*Regla de validació: si ID_Arch_Status = Collapsed, els elements A-X haurien de valer 9, no 0.*

## **2.9. Bioarqueologia (8) — BYTE 0/1/9**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| Human_Remains | BYTE | 0/1/9. Presència de restes humanes documentades |
| MNI | INTEGER | Nombre Mínim d'Individus (Minimum Number of Individuals) |
| Anatomical_Connection | BYTE | 0/1/9. Restes en connexió anatòmica |
| Mummification | BYTE | 0/1/9. Evidència de momificació (conservació de teixits tous) |
| Funerary_Bundles | BYTE | 0/1/9. Fardells funeraris (embolcall tèxtil) |
| Dispersed_Remains | BYTE | 0/1/9. Restes disperses (posició secundària) |
| Flexed_Position | BYTE | 0/1/9. Posició flexionada documentada |
| Bone_Burning | BYTE | 0/1/9. Cremació d'ossos |

## **2.10. Materials culturals (7) — BYTE 0/1/9**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| Mat_Textiles | BYTE | 0/1/9. Tèxtils presents |
| Mat_Wood | BYTE | 0/1/9. Fusta cultural (artefactes, no estructural) |
| Mat_VegFiber | BYTE | 0/1/9. Fibra vegetal |
| Mat_Ceramics | BYTE | 0/1/9. Ceràmica (inusual en tombes aèries de DW i LP) |
| Mat_Fauna | BYTE | 0/1/9. Restes de fauna |
| Mat_DeerAntler | BYTE | 0/1/9. Banya de cérvol |
| Mat_Other | BYTE | 0/1/9. Altres materials culturals |

*Bioarqueologia i materials són els blocs on el FALSE hauria estat més sovint fals: en necròpolis de penya-segat l'interior s'observa per una obertura, sovint des d'un dron, i es pot veure un farcell funerari sense poder determinar la posició flexionada. Interior_Observability (secció 2.14) documenta fins a quin punt aquests camps són avaluables en cada estructura.*

## **2.11. Cronologia (3)**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| C14 | YESNO | Datació radiocarbònica disponible (detall a T_DATING) |
| Chrono_Start_Cent | INTEGER | Segle d'inici d.n.e. (ex: 10 = s. X). Estimació sense C14 o rang calibrat |
| Chrono_End_Cent | INTEGER | Segle de fi d.n.e. |

## **2.11b. Fases constructives (2) — H03: saturació acumulativa / H04: seqüència operativa**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| Construction_Phases | INTEGER | Nombre de fases constructives identificades a l'estructura |
| Phase_Evidence | TEXT(30) | Evidència de la identificació de fases (llista al formulari): C14 / Stratigraphy / Superposition / Mortar / ND |

*Permeten registrar creixement per fases sense dependre de datacions absolutes: l'argument morfològic-estructural (superposicions, juntes, canvis d'aparell) sosté H03 en absència de C14.*

## **2.12. Volumetria i àrea (5)**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| Interior_Area_m2 | SINGLE | Àrea de sòl interior (m2). Per a cambres irregulars, extreta del model 3D |
| Interior_Vol_m3 | SINGLE | Volum interior útil (m3) |
| Total_Vol_m3 | SINGLE | Volum total incloent parets (m3) |
| ID_Vol_Method | LONG | FK -> L_VOL_METHOD. Mètode de càlcul |
| Vol_Notes | TEXT(200) | Notes sobre el càlcul volumètric |

## **2.13. Coordenades espacials (7)**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| Coord_Lat_WGS84 | DOUBLE | Latitud WGS84 (graus decimals, negatiu = S) |
| Coord_Lon_WGS84 | DOUBLE | Longitud WGS84 (graus decimals, negatiu = W) |
| Coord_E_UTM | DOUBLE | Coordenada Est UTM Zona 18S / WGS84 (m) |
| Coord_N_UTM | DOUBLE | Coordenada Nord UTM Zona 18S / WGS84 (m) |
| Altitude_masl | SINGLE | Altitud sobre el nivell del mar (m) |
| Coord_Precision_m | SINGLE | Precisió estimada de les coordenades (m) |
| ID_Coord_Method | LONG | FK -> L_COORD_METHOD |

## **2.14. Documentació digital i observabilitat (10)**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| URL_Pano | TEXT(255) | URL de la primera escena panoràmica 360 a Chacha XR |
| URL_Pano_2 | TEXT(255) | URL de la segona escena (per a estructures amb 2+ nodes) |
| URL_Giga | TEXT(255) | URL de la gigafoto |
| URL_3D | TEXT(255) | URL del model 3D fotogramètric |
| ChaXR_Documented | YESNO | Estructura publicada a la plataforma Chacha XR |
| ID_Campaign | LONG | FK -> L_CAMPAIGN. Campanya en què es va documentar sistemàticament |
| Doc_Basis | TEXT(30) | NOU v7. Base documental del registre: Direct access / Close-range photogrammetry / Distant photogrammetry / Ground photography / Published source / ND |
| Facade_Observability | TEXT(20) | NOU v7. Complete / Partial / Poor / ND |
| Interior_Observability | TEXT(20) | NOU v7. Complete / Partial / None / ND |
| Notes | MEMO | Notes lliures (text llarg) |

*C14 i ChaXR_Documented es mantenen com a YESNO: no són observacions sobre l'estructura sinó metadades sobre el corpus propi, i el seu FALSE no és mai ambigu. Els tres camps nous fan interpretables els valors 9 de tot el registre: permeten restringir l'anàlisi en R (per exemple, només estructures amb Facade_Observability = Complete) i defensar la mostra en lloc de disculpar-la. Es reporten a QRY_15.*

# **3. Taules secundàries**

## **3.1. T_DATING**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| ID | COUNTER PK |  |
| ID_Structure | LONG NN | FK -> T_STRUCTURES (cascade delete) |
| Sample_Type | TEXT(30) | Tipus de mostra: Charcoal / Bone / Wood / Textile / Other / ND |
| Date_BP | LONG | Data radiocarbònica en anys BP (Before Present) |
| Sigma1_Start | INTEGER | Interval calibrat 1 sigma - inici (cal d.n.e.) |
| Sigma1_End | INTEGER | Interval calibrat 1 sigma - fi (cal d.n.e.) |
| Sigma2_Start | INTEGER | Interval calibrat 2 sigma - inici |
| Sigma2_End | INTEGER | Interval calibrat 2 sigma - fi |
| Lab_Reference | TEXT(50) | Referència del laboratori (ex: Beta-123456) |
| Bibliog_Reference | MEMO | Referència bibliogràfica de la publicació de la data |

## **3.2. T_INDIVIDUALS**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| ID | COUNTER PK |  |
| ID_Structure | LONG NN | FK -> T_STRUCTURES |
| Individual_No | INTEGER | Número d'individu dins de l'estructura |
| Age_Category | TEXT(20) | Infant / Juvenile / Adult / Old adult / ND |
| Sex_Category | TEXT(20) | Male / Female / Indeterminate / ND |
| Preservation | TEXT(20) | Good / Fair / Poor / ND |
| Notes | MEMO | Notes sobre l'individu |

## **3.3. T_GROUPS**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| ID | COUNTER PK |  |
| Group_Code | TEXT(20) NN | Codi del conjunt (ex: DWS1-C01, PTC-S-C03) |
| ID_Sector | LONG NN | FK -> L_SECTORS |
| ID_Group_Type | LONG NN | FK -> L_GROUP_TYPE |
| N_Members | INTEGER | Nombre d'estructures membres del conjunt |
| Notes | MEMO |  |

## **3.4. T_DECORATIONS - Registre detallat de decoració per estructura i posició**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| ID | COUNTER PK |  |
| ID_Structure | LONG NN | FK -> T_STRUCTURES (cascade delete) |
| ID_Struct_Body | LONG | FK -> L_STRUCT_BODY. On està la decoració a la façana |
| ID_Dec_Type | LONG | FK -> L_DEC_TYPE. Tipus de decoració |
| Body_No | INTEGER | Número del cos constructiu (0=basament, 1=primer cos, 2=segon cos) |
| Color | TEXT(20) | Red / White / Both / Ochre / None / ND |
| Substrate | TEXT(20) | NOU v7. Plaster / Masonry stone / Bedrock / ND. Resol els casos mixtos: fris sobre revoc i brancals sobre pedra en la mateixa estructura |
| Notes | TEXT(255) |  |

## **3.5. T_ARCH_FEATURES - Registre flexible d'elements no previstos a l'esquema**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| ID | COUNTER PK |  |
| ID_Structure | LONG NN | FK -> T_STRUCTURES (cascade delete) |
| Feature_Code | TEXT(40) | Nom de l'element (text lliure o desplegable amb suggeriments: Pigment trace / Unusual bond / Textile fixation / Wooden peg / Other) |
| Present | YESNO |  |
| Feature_Count | INTEGER | Nombre d'unitats si escau |
| Material | TEXT(20) | Stone / Timber / Mixed / ND |
| Notes | TEXT(255) |  |

*En v4, els antics «elements addicionals» del formulari (A, D, J, V, W, X) tenen camp propi a la pestanya 11.Sist; T_ARCH_FEATURES queda reservada per a elements realment no previstos.*

## **3.6. T_CONNECTIONS - NOVA v4 - Connexions físiques entre estructures (OE3)**

| **Camp** | **Tipus** | **Descripció** |
| --- | --- | --- |
| ID | COUNTER PK |  |
| ID_Struct_A | LONG NN | FK -> T_STRUCTURES.ID (extrem A de l'aresta) |
| ID_Struct_B | LONG NN | FK -> T_STRUCTURES.ID (extrem B de l'aresta) |
| Connection_Type | TEXT(30) | Shared ledge / Continuous platform / Shared wall / Shared beam / Intervisibility only / ND |
| Confidence | TEXT(10) | Certain / Probable / Possible |
| Notes | TEXT(150) |  |

*Cada registre és una aresta no dirigida de la xarxa de circulació aèria. Convenció d'entrada: ID_Struct_A < ID_Struct_B (evita arestes duplicades). ID_Parent registra contenció i ID_Group agrupació funcional; cap dels dos captura l'adjacència física que reconstrueix la circulació — d'ací aquesta taula. Exportació a igraph (R) o QGIS via QRY_14.*

# **4. Taules lookup - valors**

*Tots els valors emmagatzemats són en anglés (convenció fixa del projecte). Les descripcions següents es tradueixen només a efectes de lectura del document.*

## **L_TYPOLOGY**

| **Codi** | **Descripció** |
| --- | --- |
| EA-MAU Mausoleum/Chullpa | Estructura construïda (3+ murs + sostre artificial) sobre repisa. 1-3 pisos. Predominant a La Petaca. |
| EA-CAM Funerary Chamber | Cavitat natural tancada per 1 façana construïda. Predominant a Diablo Wasi. |
| EA-PLA-R Ledge Platform | Plataforma constructiva sobre repisa natural. Funció: trànsit o base per a mausoleus. |
| EA-PLA-V Aerial Platform | Plataforma artificial sobre bigues i lloses, sense repisa natural de suport. |
| NIX Natural Niche | Petita cavitat natural (<1m2). Ossari o enterrament secundari. |
| CAV Cave/Cavern | Gran cavitat natural (>1m2) amb ús funerari o ritual documentat. |
| PR Rock Art | Motiu pictòric sobre roca, documentat de forma independent. |
| MEN Isolated Bracket | Element estructural aïllat. Evidència de xarxa de circulació aèria perduda. |
| MIX Mixed | Combinació de dues o més categories anteriors. |
| ND Undetermined | Informació insuficient per a classificar. |

## **L_SUPPORT**

| **Valor** |
| --- |
| Wide natural ledge (>2m) |
| Narrow natural ledge (<2m) |
| Artificial ledge |
| Large cavity (>10m2) |
| Medium cavity (1-10m2) |
| Natural niche (<1m2) |
| Fissure/Crack |
| Combined |
| ND |

*La llista es manté intacta en v7. Els suports compostos es registren preferentment com a parella ordenada ID_Support (dominant) + ID_Support_Secondary; «Combined» queda reservat per a suports amb tres components o no descomponibles.*

## **L_STATUS (estructura) | L_MATERIAL_STATUS (vestigis mobles)**

| **L_STATUS** | **L_MATERIAL_STATUS** |
| --- | --- |
| Good | Good - Material remains well preserved and identifiable |
| Fair | Fair - Partially preserved: some elements present and identifiable |
| Pre-collapse | Poor - Fragmentary: barely identifiable, heavily degraded or scattered |
| Collapsed | Absent - No movable remains documented (cause in boolean fields) |
| ND | ND - Not determined: not assessed or structure not accessible |

## **L_GROUP_TYPE**

| **Valor** | **Descripció** |
| --- | --- |
| Vertical alignment | Estructures en la mateixa vertical del faralló (clivella o estrats successius) |
| Ledge cluster | Múltiples EA compartint una mateixa repisa horitzontal |
| Platform with brackets | Plataforma aèria + mènsules que la componen o la flanquegen |
| Cave cluster | Cova + estructures construïdes al seu interior |
| Circulation network | MEN i EA-PLA que reconstrueixen una circulació aèria perduda |
| Rock art cluster | PR + EA adjacent visualment o funcionalment vinculada |
| Functional group | Qualsevol altra agrupació intrasector amb coherència funcional |

## **L_STRUCT_BODY (posició de la decoració a la façana)**

| **Code** | **Name** | **Level_Type** | **Descripció** |
| --- | --- | --- | --- |
| SOC | Socle | N0 | Sòcol decoratiu - tractament de la massa basal |
| SPA | Spandrel | N1 | Parament lateral, fora del marc del portal |
| JAM | Jamb | N1 | Brancal del portal - element vertical del marc de l'obertura |
| OVL | Over-lintel | N1 | Zona per damunt del dintell - àrea del fris decoratiu |
| COR | Interbody cornice | N1 | Zona de la cornisa entre cossos superposats |
| ND | Not determined | ND | Posició no determinada |

*Reformulada en v9: aquesta taula descriu **només la posició dins d'un cos**; de quin cos es tracta ho diu `T_DECORATIONS.Body_No`. Les entrades antigues N1-SPA i N2-SPA descrivien la mateixa posició en cossos distints, de manera que una estructura de tres cossos hauria exigit inventar N3-SPA, i així indefinidament. Separant posició (ací) d'índex de cos (Body_No), l'esquema escala a qualsevol nombre de cossos sense tocar el lookup. `Level_Type` indica a quin nivell funcional pertany la posició, mai un número de cos.*

## **L_DEC_TYPE**

| **Nom** | **Descripció** |
| --- | --- |
| T-shaped niche | Nínxol o baix relleu en forma de T. Element vertical + horitzontal. |
| T-shaped niche inv. | T invertida. |
| L-shaped niche | Nínxol o baix relleu en forma de L. |
| L-shaped niche inv. | L invertida. |
| Zigzag | Motiu en zig-zag o chevron. |
| Stepped motif | Motiu escalonat. Documentat a DWS1 (inusual a la cultura Chachapoya). |
| Frieze / Greca | Fris o greca repetitiva. Comú a La Petaca. |
| Triangular motif | Patró triangular / chevron pintat. Documentat a DW en zona sobre el dintell. |
| Painted band | Banda horitzontal pintada (roja/blanca). Zona de la cornisa intercòs. |
| Square niche | Nínxols quadrats en sèrie (hornacinas cuadradas). |
| ND | Tipus de decoració no determinat. |

## **L_CAMPAIGN**

| **Code** | **Campaign_Name** | **Descripció** |
| --- | --- | --- |
| 2013 | PALP I | Primera campanya. Prospecció i documentació inicial de La Petaca. Dra. J. Marla Toyne (UCF). |
| 2016 | PALP II | Ampliació a Diablo Wasi. Primera documentació sistemàtica de DW. |
| 2021 | La Petaca Project | Documentació no invasiva integral. Fotogrametria, 360, gigafotos. Panograma Labs/UCF. |
| 2023 | PALP IV | Campanya d'excavació arqueològica i reconstruccions 3D detallades. |

# **5. Relacions (21)**

| **Nom** | **Taula pare** | **Camp pare** | **Taula filla** | **Camp fill** | **Cascade** | **Integritat** |
| --- | --- | --- | --- | --- | --- | --- |
| REL_SIT_SEC | L_SITES | ID | L_SECTORS | ID_Site | Delete | Sí |
| REL_SEC_STR | L_SECTORS | ID | T_STRUCTURES | ID_Sector | Delete | Sí |
| REL_SEC_GRP | L_SECTORS | ID | T_GROUPS | ID_Sector | Delete | Sí |
| REL_TYP_STR | L_TYPOLOGY | ID | T_STRUCTURES | ID_Typology | Update | Sí |
| REL_SUP_STR | L_SUPPORT | ID | T_STRUCTURES | ID_Support | Update | Sí |
| REL_STR_SELF | T_STRUCTURES | ID | T_STRUCTURES | ID_Parent | Update | NO |
| REL_STA_STR | L_STATUS | ID | T_STRUCTURES | ID_Arch_Status | Update | Sí |
| REL_MATSTA_STR | L_MATERIAL_STATUS | ID | T_STRUCTURES | ID_Material_Status | Update | Sí |
| REL_VM_STR | L_VOL_METHOD | ID | T_STRUCTURES | ID_Vol_Method | Update | Sí |
| REL_CM_STR | L_COORD_METHOD | ID | T_STRUCTURES | ID_Coord_Method | Update | Sí |
| REL_GT_GRP | L_GROUP_TYPE | ID | T_GROUPS | ID_Group_Type | Update | Sí |
| REL_GRP_STR | T_GROUPS | ID | T_STRUCTURES | ID_Group | Update | Sí |
| REL_STR_DAT | T_STRUCTURES | ID | T_DATING | ID_Structure | Delete | Sí |
| REL_STR_IND | T_STRUCTURES | ID | T_INDIVIDUALS | ID_Structure | Delete | Sí |
| REL_STR_DEC | T_STRUCTURES | ID | T_DECORATIONS | ID_Structure | Delete | Sí |
| REL_SB_DEC | L_STRUCT_BODY | ID | T_DECORATIONS | ID_Struct_Body | Update | Sí |
| REL_DT_DEC | L_DEC_TYPE | ID | T_DECORATIONS | ID_Dec_Type | Update | Sí |
| REL_STR_AFEAT | T_STRUCTURES | ID | T_ARCH_FEATURES | ID_Structure | Delete | Sí |
| REL_STR_CONA | T_STRUCTURES | ID | T_CONNECTIONS | ID_Struct_A | - | NO |
| REL_STR_CONB | T_STRUCTURES | ID | T_CONNECTIONS | ID_Struct_B | - | NO |
| REL_SUP2_STR | L_SUPPORT | ID | T_STRUCTURES | ID_Support_Secondary | Update | Sí |

*REL_STR_SELF (autoreferenciant) i les dues relacions de T_CONNECTIONS (doble referència a T_STRUCTURES) usen dbRelationDontEnforceIntegrity (valor numèric 2) per a evitar conflictes. La resta usen dbRelationUpdateCascade i, si Cascade=Delete, dbRelationDeleteCascade.*

# **6. Consultes SQL (16)**

| **Nom** | **Descripció** | **Hipòtesis** |
| --- | --- | --- |
| QRY_01_Typology_by_Site | Distribució de tipologies per jaciment i sector (COUNT per grup). | H01, H04 |
| QRY_02_Decoration_by_Site | Presència de cada motiu decoratiu per jaciment. Font per a khi-quadrat. Actualitzada v7: tots els comptadors avaluen `= 1`, ja que els camps són BYTE 0/1/9. Els registres amb 9 no compten com a absència, però tampoc s'exclouen del total: cal filtrar-los en R o encreuar amb QRY_15. | H01, H04 |
| QRY_03_Conservation_by_Sector | Distribució dual d'estat (estructura + vestigis mobles) per sector. | General |
| QRY_04_C14_Structures | Estructures amb datació C14 ordenades cronològicament. | H04 |
| QRY_05_Export_RStats | Exportació plana completa per a R/SPSS. Actualitzada v9: comptadors de cossos i Lost_Body_Evidence; bloc de tractaments superficials (revoc + pigment + substrat) i els tres camps d'observabilitat. | Totes |
| QRY_06_Volumetry_by_Typology | Volumetria i àrea interiors per tipologia (mitjana, mín., màx.). | OE2, H01 |
| QRY_07_Export_QGIS | Exportació espacial per a QGIS. Actualitzada v9: comptadors de cossos, Lost_Body_Evidence, Doc_Basis i Facade_Observability (permeten simbolitzar el biaix documental sobre el mapa). LEFT JOIN per a conservar estructures sense estat. | OE3, H02, H03, H06 |
| QRY_08_Children_of_Parent | Elements continguts per una estructura pare (paràmetre: [Parent ID?]). | H01, H02 |
| QRY_09_Group_Members | Membres d'un conjunt funcional ordenats per altitud (paràm.: [Group code?]). | H03, OE3 |
| QRY_10_ChaXR_Coverage | Cobertura Chacha XR vs total per jaciment i campanya. | Metodologia |
| QRY_11_Masonry_by_Site | Qualitat i tipus de maçoneria per jaciment i tipologia (COUNT per grup; exclou registres sense Masonry_Quality). | H01, H04 |
| QRY_12_Geology_Construction | Refeta v8. Encreuament suport-tipologia amb la parella completa (classe dominant + secundària), modificació antròpica i mitjanes d'amplària i profunditat (cm). Doble LEFT JOIN sobre L_SUPPORT. | H02 |
| QRY_13_AX_Pattern_Export | Matriu A-X completa: 24 columnes AX_A..AX_X en ordre alfabètic amb valors 0/1/9. AX_Q derivada del camp Lintel (Absent→0, ND/Null→9, resta→1). Ampliada v7 amb Platform_Surface_Material, Rear_Wall_Type, el bloc de pigment i els tres camps d'observabilitat. En R, filtrar o ponderar les cel·les amb valor 9 abans de calcular phi/Jaccard, clúster jeràrquic, I de Moran o AC. LEFT JOIN amb L_TYPOLOGY per a incloure registres parcials. | H01, H04, H05 |
| QRY_14_Connections_Edges | Llista d'arestes de T_CONNECTIONS amb els codis i les coordenades UTM/altitud dels dos extrems: entrada directa per a igraph (R) o generació de línies en QGIS (xarxa de circulació aèria). | OE3, H03 |
| QRY_15_Observability_Bias | Recompte per jaciment i sector creuant Doc_Basis, Facade_Observability i Interior_Observability. Quantifica on són els buits del registre i permet justificar els subconjunts analítics. | OE1, metodologia |
| QRY_16_Validation_Check | Bateria de validació de coherència (UNION de nou regles en v9). Resultat buit = corpus coherent. Vegeu la secció 9.5. | Metodologia |

# **9. Ús estadístic del domini 0/1/9 — advertència operativa**

Aquesta secció és de lectura obligada abans de qualsevol prova estadística sobre camps de domini 0/1/9. Documenta un error fàcil de cometre i difícil de detectar un cop comés.

## **9.1. El problema del denominador**

Els comptadors de QRY_02 avaluen `= 1`, de manera que un valor 9 (no observable) **no compta com a presència**. Fins ací, correcte. El problema és el denominador: si es pren `Total` —que compta totes les estructures del jaciment, inclosos els 9— i s'alimenta una taula de contingència amb `N_motiu` i `Total`, el resultat és doblement erroni:

1. El denominador s'infla amb estructures que mai no van ser avaluades, cosa que **subestima la freqüència real** del motiu.
2. I, sobretot, **aquesta inflació no és igual als dos jaciments**. La conservació de façanes difereix entre La Petaca i Diablo Wasi, de manera que la proporció de 9 també difereix. La prova mesuraria llavors preservació diferencial i el resultat es llegiria com a pràctica decorativa diferencial: exactament l'artefacte que el domini 0/1/9 existeix per a evitar.

En altres paraules: **conservar els 9 al denominador anul·la el benefici de tota la conversió**. El camp registra correctament la incertesa, però l'anàlisi la torna a esborrar.

## **9.2. La solució: denominador explícit**

QRY_02 exporta, per a cada motiu i jaciment, **tres comptadors** en lloc d'un:

| Columna | Significat | Ús |
| --- | --- | --- |
| `N_motiu` | Estructures amb el motiu **present** (valor 1) | Numerador |
| `N_motiu_Absent` | Estructures amb absència **verificada** (valor 0) | Part del denominador |
| `N_motiu_ND` | Estructures **no observables** (valor 9) | Excloses de la prova |
| `Total` | Totes les estructures del jaciment | **Mai** com a denominador |

**Denominador vàlid = `N_motiu` + `N_motiu_Absent`.** El camp `Total` es conserva únicament per a reportar la cobertura: la ràtio `N_motiu_ND / Total` és la fracció del corpus exclosa de la prova, i **s'ha de fer constar sempre que es reporte el resultat**, igual que es reporta la n.

Exemple en R a partir de QRY_02 exportada:

```r
q2 <- read.csv("QRY_02_Decoration_by_Site.csv")

# Taula de contingencia per al fris, nomes amb observacions valides
tab <- with(q2, rbind(
  LaPetaca   = c(N_Frieze[Site_Name=="La Petaca"],
                 N_Frieze_Absent[Site_Name=="La Petaca"]),
  DiabloWasi = c(N_Frieze[Site_Name=="Diablo Wasi"],
                 N_Frieze_Absent[Site_Name=="Diablo Wasi"])))
colnames(tab) <- c("Present","Absent")

chisq.test(tab)          # o fisher.test(tab) si alguna cel.la < 5

# Cobertura, a reportar junt amb el resultat
q2$coverage <- 1 - q2$N_Frieze_ND / q2$Total
```

## **9.3. La mateixa regla per a la matriu A–X (QRY_13)**

El principi és idèntic per a les 24 columnes `AX_A`…`AX_X`. Abans de calcular coeficients phi o Jaccard, clúster jeràrquic, anàlisi de correspondències o I de Moran, els valors 9 s'han de convertir a `NA` —mai a 0—:

```r
ax <- read.csv("QRY_13_AX_Pattern_Export.csv")
axcols <- grep("^AX_", names(ax), value = TRUE)
ax[axcols] <- lapply(ax[axcols], function(x) ifelse(x == 9, NA, x))
```

A partir d'ací hi ha dues estratègies legítimes, i cal declarar quina s'ha fet servir:

- **Restricció de la mostra.** Analitzar només les estructures amb `Facade_Observability == "Complete"` (camp exportat per QRY_13). Mostra menor però homogènia.
- **Distàncies amb tractament de nuls.** Calcular la matriu amb `dist(..., method = "binary")`, que ignora els parells amb `NA`. Conserva la mostra però amb pesos desiguals entre parells.

En tots dos casos, QRY_15_Observability_Bias proporciona el recompte per jaciment i sector que justifica la decisió.

## **9.4. Regla general**

> Un valor 9 no és ni un 0 ni un 1: és una **cel·la buida**. Qualsevol operació que el convertisca implícitament en 0 —sumar, comptar el total, fer la mitjana sense `na.rm`— reintrodueix el biaix de conservació al resultat.

# **10. Vocabulari arquitectònic normalitzat A-X**

24 elements organitzats per nivell constructiu (bottom-to-top) i funció (estructural -> decoratiu). Tres sistemes compositius: E+F+G->H (plataforma), N+O+Q->P (portal), S+T->U (ràfec).

**Nota terminològica i criteri operatiu:** «nivell» designa exclusivament els nivells constructius del vocabulari (N0 / N1 / Superior), que són **categories funcionals i no un sistema de numeració**: no s'amplien mai amb N2, N3. «Cos» designa els pisos superposats de l'estructura, i la seua repetició es compta a `N_Basal_Bodies` i `N_Chamber_Bodies`. El criteri per a assignar un cos a un nivell és unívoc: **un cos és N1 si conté (o contenia) una obertura d'accés; si no en té, és N0**, per acurada que siga la seua fàbrica. (Camps relacionats: Body_No, L_STRUCT_BODY). La cornisa I és per tant «cornisa intercòs». H i I comparteixen posició d'interfície N0/N1 però es distingeixen per funció (H: superfície de circulació i accés; I: marcatge i separació entre cossos), per composició (H és sistema; I és standalone) i per seqüència operativa (H tanca N0 i precondiciona N1; I apareix dins de l'alçat N1, només amb 2+ cossos).

| **Ll.** | **Valencià** | **Anglés (BD)** | **Nivell** | **Sistema** | **Camp BD** |
| --- | --- | --- | --- | --- | --- |
| A | Jàcenes basals empotrades | Embedded base beams | N0 | Standalone | Embedded_Base_Beams |
| B | Basament | Base level / podium | N0 | Standalone | Base_Level |
| C | Sòcol decoratiu | Decorative socle | N0 | Standalone (tract. de B) | Decorative_Socle |
| D | Muret transversal | Tie wall | N0 | Standalone | Tie_Walls |
| E | Mènsules (fusta) | Timber corbels | N0 | Component de H | Timber_Brackets |
| F | Bigues transversals | Transverse beams | N0 | Component de H | Transverse_Beams |
| G | Filades en voladís | Corbelled masonry courses | N0 | Component de H (lítia) | Corbelled_Courses |
| **H** | **Plataforma volada d'accés** | Corbelled access platform | **N0/N1 (interfície)** | **SISTEMA: E+F+G -> H** | Corbelled_Platform (+ Platform_Surface_Material) |
| I | Cornisa intercòs | Interbody cornice | N0/N1 (interfície) | Standalone (límit cossos) | Interbody_Cornice |
| J | Cantoneres | Corner quoins | N1 | Standalone | Corner_Quoins |
| K | Pilastres estructurals | Structural pilasters | N1 | Standalone | Structural_Pilasters |
| L | Paraments laterals | Lateral wall faces | N1 | Standalone | Lateral_Wall_Faces |
| M | Fris decoratiu en baix relleu | Bas-relief decorative frieze | N1 | Standalone (tract. de L) | Relief_Frieze |
| N | Llindar | Sill / threshold | N1 | Component de P | Sill |
| O | Brancals | Jambs | N1 | Component de P | Jambs |
| **P** | **Obertura d'accés** | Access opening | N1 | **SISTEMA: N+O+Q -> P** | Access_Opening |
| Q | Dintell | Lintel | N1 | Component de P | Lintel (material; AX_Q derivada a QRY_13) |
| R | Coronament | Upper crown / coping | N1 | Standalone | Upper_Crown |
| S | Biga de suport del ràfec | Eave-supporting beam | Sup. | Component de U | Eave_Beam |
| T | Superfície del ràfec | Eave surface | Sup. | Comp. de U (SEMPRE PEDRA) | Eave_Surface |
| **U** | **Ràfec-voladís / Visera** | Eave / roof overhang | Sup. | **SISTEMA: S+T -> U (PEDRA)** | Eave |
| V | Murs laterals | Lateral walls | N1 | Standalone | Lateral_Walls |
| W | Mur posterior | Rear wall | N1 | Standalone (+ Rear_Wall_Type) | Rear_Wall |
| X | Coberta de la cambra | Chamber roof | Sup. | Standalone | Chamber_Roof |

**Criteri operatiu E vs F (reproduïbilitat):** E (mènsula) = element perpendicular a la façana, encastat a la roca, treballant en voladís; F (biga transversal) = element paral·lel a la façana, salvant llum entre suports. Fixar aquest criteri per escrit garanteix la consistència del registre entre estructures i observadors.

# **11. Formulari i visualització de dades**

El formulari es genera amb chachapoya_Form_v9_val.bas i segueix la convenció bilingüe fixa del projecte: **totes les etiquetes visibles de la interfície (pestanyes, camps, capçaleres de subformularis) són en valencià, mentre que tots els valors emmagatzemats (llistes de valors dels ComboBox, continguts de les taules lookup, dominis 0/1/9 amb etiquetes Absent/Present/ND) romanen en anglés**, per a garantir la reproduïbilitat de les exportacions analítiques.

Les 12 pestanyes de F_STRUCTURES són: 1.Id. (identificació + morfologia del suport geològic), 2.Arq. (morfologia general + tipus de coberta + façana/paisatge + maçoneria i morter + fases), 3.Acab. (revoc + pigment i substrat), 4.Dec. (baix relleu + art rupestre + subformulari F_DECORATIONS), 5.Estat (conservació + base documental i observabilitat), 6.Bio., 7.Mat., 8.Cron., 9.Metr. (dimensions + obertura + mètrica del suport + volumetria + coordenades), 10.Doc. (URLs i notes), 11.Sist. (els 24 elements A-X amb combos 0/1/9) i 12.Extra (subformulari F_ARCH_FEATURES).

**Capçaleres de columna dels subformularis:** en vista full de dades, Access mostra com a capçalera la llegenda de l'etiqueta **adjunta** al control; si el control no té etiqueta adjunta, mostra el nom del control (el nom de camp anglés). En v3, les etiquetes dels subformularis es creen amb el control com a pare (patró: primer el control amb nom = camp, després CreateControl amb el nom del control com a quart paràmetre), cosa que resol les capçaleres en valencià. Les captions DAO a nivell de TableDef (SetFieldCaptionsVal) es mantenen com a reforç per a l'obertura directa de T_DECORATIONS, T_ARCH_FEATURES i T_CONNECTIONS en vista de taula.

**Entrada de dades a T_CONNECTIONS:** sense formulari propi (volum baix d'arestes); s'ompli en vista de taula amb les capçaleres DAO en valencià. Convenció: ID_Struct_A < ID_Struct_B.
