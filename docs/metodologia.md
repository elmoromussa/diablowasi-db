**Universitat d'Alacant**

Departament de Prehistòria, Arqueologia, H.ª Antiga, Filologia Grega i Filologia Llatina

**DISSENY D'UNA BASE DE DADES ARQUEOLÓGICA**

per a l'estudi de les necròpolis de penya-segat de La Petaca i Diablo Wasi

*(Leymebamba, Amazonas, Perú)*

Treball de Fi de Màster

Màster en Arqueologia Professional i Gestió Integral del Patrimoni

Autor: Esteve Ribera Torró

Directors: Dr. Ignasi Grau Mira · Dra. J. Marla Toyne

Curs acadèmic 2024-2025

## **Resum**

El present document descriu el disseny, la justificació i la implementació d'una base de dades relacional destinada a la documentació sistemàtica i l'anàlisi estadística de les estructures funeràries de les necròpolis de penya-segat de La Petaca i Diablo Wasi (Leymebamba, Departament d'Amazonas, Perú). La base de dades constitueix l'eix vertebrador de la metodologia del Treball de Fi de Màster, en tant que permet centralitzar les variables arqueológiques, establir relacions jeràrquiques entre elements, exportar dades per a l'anàlisi estadística i vincular el registre arqueológic amb el Sistema d'Informació Geogràfica (SIG) implementat en QGIS.

L'esquema adoptat segueix els principis de la tercera forma normal (3FN) de la teoria de bases de dades relacionals i inclou una taula principal (T_ESTRUCTURES) amb 65 camps, tres taules secundàries (T_DATACIONS, T_INDIVIDUS, T_CONJUNTS) i nou taules de consulta (lookup). El sistema implementa dos mecanismes complementaris de gestió de la jerarquia entre elements: contenció física (camp autoreferenciant ID_Estructura_Parent) i agrupació funcional (taula T_CONJUNTS), que permeten representar la complexa casuística espacial de les necròpolis de penya-segat (mènsules aïllades, estructures dins cavernes, alineaments verticals, xarxes de circulació perdudes).

## **1. Introducció**

L'estudi de les necròpolis de penya-segat Chachapoya planteja reptes documentals i analítics inèdits en l'arqueologia andina. La verticalitat dels jaciments, l'heterogeneïtat tipológica de les estructures —que combinen mausoleus construïts, cambres funeràries en cavitats naturals, plataformes volades, nínxols, mènsules aïllades i pintures rupestres— i la multiplicitat de variables arquitectòniques, decoratives, bioarqueológiques i cronológiques exigeixen una eina de gestió de dades que vagi més enllà d'un full de càlcul pla. La base de dades relacional resposta a aquesta necessitat.

En el context del present TFM, la base de dades serveix com a infraestructura metodológica central per a la consecució dels tres objectius específics declarats al projecte de recerca:

- Documentar i classificar les estratègies constructives de les estructures funeràries.
- Analitzar volums i relacions espacials entre estructures.
- Interpretar la relació entre geologia, construcció i ús funerari.

Igualment, la base de dades és la font de dades sobre la qual s'operacionalitzen les quatre hipòtesis de treball del TFM. La Taula 1 sintetitza la correspondència entre hipòtesis i camps de la base de dades.

| **Hipòtesi** | **Descripció sintètica** | **Camps i taules rellevants** |
|---|---|---|
| H01 | Ingenieria funerària planificada | N_Pisos, Planta, N_Murs_Construits, Orientacio_Vano, Dintel, Cornisa, Estacas_Fusta |
| H02 | Geologia com a factor determinant | ID_Suport, Techo_Natural, Altura_Aprox_m, Altitud_msnm, Coord_* |
| H03 | Saturació espacial acumulativa | T_CONJUNTS, ID_Conjunt, Campanya_Documentacio, distribució espacial via QGIS |
| H04 | Seqüència operativa consistent | ID_Tipologia, Dec_*, Campanya_Documentacio, T_DATACIONS (C14) |

*Taula 1. Correspondència entre les hipòtesis del TFM i els camps de la base de dades.*

## **2. Fonamentació metodológica**

### **2.1. La base de dades relacional en arqueologia de l'arquitectura**

La gestió sistemàtica de variables en arqueologia de l'arquitectura requereix eines que superen les limitacions dels fulls de càlcul plans. Les bases de dades relacionals permeten: (a) evitar la redundància d'informació mitjançant la normalització, (b) assegurar la integritat referencial entre entitats relacionades, (c) executar consultes analítiques complexes en SQL, i (d) exportar dades en formats compatibles amb programari estadístic (R, SPSS) i SIG (QGIS).

Per a l'estudi de necròpolis de penya-segat, la complexitat espacial afegeix una dificultat específica: els elements arqueológics no s'organitzen en un pla XY convencional sinó en un faralló tridimensional on múltiples estructures comparteixen coordenades XY però difereixen en la coordenada Z (altitud). La gestió d'aquesta complexitat 3D s'ha integrat al disseny de la base de dades a través de camps de coordenades tridimensionals (Coord_E_UTM, Coord_N_UTM, Altitud_msnm) i d'un sistema de jerarquies que permet representar les relacions de contenció i agrupació entre elements a escales múltiples.

### **2.2. Marc disciplinar**

El disseny de les variables arqueológiques s'ha fonamentat en dues disciplines:

L'Arqueologia de l'Arquitectura aporta el marc per a la caracterització constructiva (materials, tècniques, fase constructiva, decoració) i la tipologia de les estructures. Seguint la proposta de Toyne i Anzellini (2017) per als mausoleus de La Petaca, s'ha adoptat una classificació tipológica que distingeix entre estructures amb funció mortuòria primària (EA-MAU, EA-CAM), elements d'infraestructura (EA-PLA-R, EA-PLA-V, MEN) i elements simbòlics o rituals (PR, NIX, CAV). Aquesta classificació permet aplicar estadística descriptiva i inferencial a les estratègies constructives (H01) i a la seqüència d'ús del jaciment (H04).

L'Arqueologia del Paisatge proporciona el marc per a l'anàlisi de la distribució espacial de les estructures i les seues relacions amb la geologia i l'entorn natural. La integració de la base de dades amb QGIS permet implementar anàlisis de densitat (Kernel Density Estimation), visibilitat (viewshed analysis) i distribució vertical, directament relacionades amb les hipòtesis H02 i H03.

### **2.3. Procediment de disseny**

El disseny de la base de dades s'ha dut a terme seguint un procés iteratiu de quatre etapes:

- Inventari de variables a partir de la documentació existent: anàlisi de les fitxes d'estructures (fichas_estructuras.xlsx), els papers publicats (Toyne i Anzellini 2017; Epstein i Toyne 2016; Toyne et al. 2018), el TFM Chacha XR (Ribera-Torró 2023) i l'article derivat (Ribera-Torró et al. 2026).
- Modelatge entitat-relació (ER): identificació de les entitats principals (estructura, datació, individu, conjunt funcional), els seus atributs i les relacions entre elles.
- Normalització fins a la 3FN: eliminació de la redundància i creació de taules de consulta (lookup) per a tots els camps categòrics.
- Validació amb dades reals: prova de consistència amb els registres existents de DW i LP, ajust de les llistes de valors permesos i afegit de camps derivats de la revisió de la documentació.
## **3. Arquitectura de la base de dades**

### **3.1. Taules i relacions**

La base de dades consta de 13 taules organitzades en tres categories funcionals (Taula 2). La taula T_ESTRUCTURES és el nucli del sistema: cada fila correspon a un element arqueológic documentat, independentment del seu tipus (mausoleu, cambra, mènsula, pintura rupestre, etc.). Totes les anàlisis estadístiques principals s'executen sobre aquesta taula o sobre vistes derivades d'ella.

| **Categoria** | **Taula** | **Contingut i funció** |
|---|---|---|
| Principal | T_ESTRUCTURES | Fitxa completa de cada element arqueológic (65 camps). Font principal per a l'anàlisi estadística. |
| Secundàries | T_DATACIONS | Datacions radiocarbòniques (C14) associades a cada estructura. Relació N:1. |
| Secundàries | T_INDIVIDUS | Dades individuals de restes esquelètiques quan es disposa d'informació desagregada. |
| Secundàries | T_CONJUNTS | Agrupacions funcionals d'elements (alineaments, xarxes de circulació, conjunts de repisa). |
| Lookup | L_SITIOS | Jaciments: La Petaca, Diablo Wasi. Inclou coordenades WGS84 de referència. |
| Lookup | L_SECTORES | Sectors per jaciment (LP: Nord, Central, Superior, Sud; DW: Sectors 1-6). |
| Lookup | L_TIPOLOGIA | Tipologia de l'element (EA-MAU, EA-CAM, EA-PLA-R, EA-PLA-V, NIX, CAV, PR, MEN...). |
| Lookup | L_SUPORT | Suport geomorfológic (repisa natural àmplia/estreta, artificial, cavitat, grieta...). |
| Lookup | L_ESTAT | Estat de conservació (bo, regular, pre-col·lapse, col·lapsat, ND). |
| Lookup | L_METODE_VOLUM | Mètode de càlcul del volum interior de l'estructura. |
| Lookup | L_COORD_METODE | Mètode d'obtenció de les coordenades espacials (drone RTK, fotogrametria...). |
| Lookup | L_TIPUS_CONJUNT | Tipus d'agrupació funcional (alineament vertical, xarxa de circulació...). |
| Lookup | L_CAMPANYA | Any de campanya arqueológica (2013, 2016, 2021, 2023). |

*Taula 2. Resum de les 13 taules de la base de dades i la seua funció.*

### **3.2. Relacions entre taules**

Les relacions entre taules segueixen el model relacional estàndard, implementades com a claus foranies (Foreign Keys, FK). En la implementació a Microsoft Access, totes les relacions han d'activar la integritat referencial per evitar orfes (estructures sense sector vàlid, datacions sense estructura, etc.). Les relacions principals són:

- L_SITIOS → L_SECTORES: un jaciment pot tenir N sectors.
- L_SECTORES → T_ESTRUCTURES: un sector pot contenir N estructures.
- L_TIPOLOGIA, L_SUPORT, L_ESTAT → T_ESTRUCTURES: relació lookup 1:N.
- T_ESTRUCTURES → T_DATACIONS: una estructura pot tenir N datacions C14.
- T_ESTRUCTURES → T_INDIVIDUS: una estructura pot contenir N individus documentats.
- T_CONJUNTS → T_ESTRUCTURES: un conjunt pot agrupar N estructures.
- T_ESTRUCTURES → T_ESTRUCTURES (autoreferenciant): una estructura pot ser filla d'una altra.

*Nota implementació: per crear la relació autoreferenciant a Access, cal afegir T_ESTRUCTURES dues vegades a la vista Relaciones. La primera instància actua com a pare i la segona com a fill. Activar sempre la integritat referencial i la actualización en cascada.*

## **4. Variables arqueológiques de T_ESTRUCTURES**

La taula T_ESTRUCTURES concentra 65 camps distribuïts en 13 categories temàtiques. La selecció de variables respon a dos criteris: la rellevància per a les hipòtesis de treball del TFM i la possibilitat de ser observada directament en camp o inferida a partir dels models fotogramètrics 3D i les ortofotos verticals. S'ha donat prioritat a les variables susceptibles d'anàlisi estadística (variables categóriques codificades com a lookup i variables binàries Yes/No per a decoració, materials i restes), ja que permeten l'ús de proves de χ², correlació de Spearman i anàlisi de correspondències (AFC).

### **4.1. Identificació, tipologia i suport geomorfológic**

Cada registre s'identifica mitjançant un codi estructural (Codi) que segueix la convenció establerta en el Proyecto Arqueológico Las Peñas: [JACIMENT][SECTOR]-[TIPUS][NUM]. Exemples: DWS1-EF01 (Diablo Wasi Sector 1, Estructura Funeraria 01), PTC-SN-EF18 (La Petaca, Sector Nord, EF18). El codi coincideix amb la nomenclatura de Chacha XR i els models Metashape.

La tipologia de l'element (camp ID_Tipologia) és la variable categórica de primer ordre. La taula L_TIPOLOGIA inclou deu valors que cobreixen totes les casuístiques identificades als dos jaciments:

| **Codi** | **Nom complet** | **Caracterització** |
|---|---|---|
| EA-MAU | Mausoleu / Chullpa | Estructura construïda (3+ murs + sostre artificial) sobre repisa. 1-3 pisos. Predominant a La Petaca. |
| EA-CAM | Cambra funerària | Cavitat natural tancada per 1 façana construïda. Predominant a Diablo Wasi. |
| EA-PLA-R | Plataforma sobre repisa | Plataforma constructiva sobre repisa natural. Funció: trànsit o base per a mausoleus. |
| EA-PLA-V | Plataforma volada | Plataforma artificial sobre fustes i lloses, sense repisa natural de suport. |
| NIX | Nínxol natural | Petit nínxol (<1m²). Funció: ossari o enterrament secundari. |
| CAV | Caverna / Cova | Gran cavitat natural (>1m²) amb ús funerari o ritual documentat. |
| PR | Pintura rupestre | Motiu pictòric sobre roca, documentat de forma independent. |
| MEN | Mènsula aïllada | Element estructural aïllat sense estructura conservada. Evidència de circulació perduda. |
| MIX | Mixt | Combinació de dues o més categories anteriors. |
| ND | No determinat | Informació insuficient per classificar. |

*Taula 3. Valors de L_TIPOLOGIA amb codis i caracterització.*

La tipologia de l'element està directament relacionada amb les hipòtesis H01 i H04: la distribució de tipologies per sector i la seua evolució temporal (camp Campanya_Documentacio, T_DATACIONS) permeten avaluar si hi ha una seqüència constructiva consistent (H04) i si la planificació arquitectónica respon a patrons reconeixibles (H01).

El suport geomorfológic (ID_Suport, linked a L_SUPORT) és la variable que operacionalitza la hipòtesi H02. Recull el tipus de base geológica que els constructors Chachapoya van seleccionar i adaptar: repisa natural àmplia o estreta, repisa artificial, cavitat (gran/mitjana/nínxol), grieta o combinat. La correlació entre suport geomorfológic i tipologia arquitectónica (test de χ²) permetrà avaluar en quina mesura la geologia predetermina la solució constructiva adoptada.

### **4.2. Variables arquitectòniques**

Les variables arquitectòniques cobreixen les dimensions i la morfologia de les estructures, essencials per a l'anàlisi volumètrica i la comparació tipológica. El nombre de pisos constructius (N_Pisos), la planta (Planta: rectangular, quadrada, trapezoïdal, irregular), el nombre de murs construïts (N_Murs_Construits: 0-4) i l'orientació de l'accés (Orientacio_Vano: N, S, E, O i variants) aporten dades quantitatives i categóriques sobre la planificació constructiva.

L'orientació del vano —l'apertura d'accés de la cambra— és particularment rellevant per a H01 i H02. A Diablo Wasi Sector 4, l'observació de camp (Ribera-Torró 2023, Annex VII) documenta que els mausoleus (EA-MAU) presenten obertures orientades al migjorn, mentre que les cambres funeràries (EA-CAM) apunten cap a ponent. La base de dades permet confirmar o matisar aquesta pauta amb dades sistemàtiques de tots els sectors.

Les dimensions (Largo_m, Ancho_m, Alto_m) i l'alçada aproximada sobre el sòl del faralló (Altura_Aprox_m) permeten correlacionar la mida de les estructures amb el nombre mínim d'individus (NMI), avaluant si les tombes de major capacitat volumètrica allotjaven de forma consistent un nombre superior d'inhumats.

### **4.3. Acabats superficials i decoració arquitectònica**

Les variables decoratives s'han codificat com a camps binaris (Yes/No) independents per a cada motiu, en lloc d'un únic camp de text. Aquesta decisió metodológica és essencial per a l'anàlisi estadística: permet construir matrius de presència/absència de decoració per a l'Anàlisi de Correspondències (AFC), calcular freqüències per jaciment i sector (test de χ²) i aplicar algoritmes de clúster per identificar agrupacions tipológiques.

Els camps de decoració en baix-relleu (Dec_Nicho_Quadrat, Dec_Relieve_T, Dec_Relieve_T_Inv, Dec_Relieve_L, Dec_Relieve_L_Inv, Dec_Zigzag, Dec_Escalonat, Dec_Fris_Greca) recullen tots els motius documentats als dos jaciments i en la bibliografia de referència (Toyne i Anzellini 2017). La presència del motiu escalonat a DW Sector 1 —inèdit en la cultura Chachapoya— es recull al camp Dec_Escalonat i constitueix un element d'especial interès analític (Ribera-Torró 2023: Annex VII).

Els camps de pintura rupestre (Pintura_Rupestre, PR_Antropomorfa, PR_Zoomorfa, PR_Geometrica, PR_Abstracta, PR_Escena_Decap) permeten analitzar la distribució espacial dels motius pictòrics en relació amb les estructures arquitectóniques adjacents, especialment rellevant per als conjunts DWS1-P11 i PTC-SS-EF18.

### **4.4. Estat de conservació i alteracions**

L'estat general de conservació (ID_Estat) s'expressa en una escala ordinal de quatre valors (bo, regular, pre-col·lapse, col·lapsat) que permet anàlisi de distribució i correlació amb la posició al faralló. Els camps d'alteració (Saqueig, Incendi, Activitat_Animal, Evidencia_Acces_Modern) registren de forma diferenciada els agents d'alteració, distingint entre saqueig (presumptament antic o colonial), incendi, activitat faunística i accés modern no autoritzat —documentat, per exemple, per la presència d'un mosquetó d'escalada a DWS1-EF17 (Ribera-Torró 2023: Annex VII).

### **4.5. Context bioarqueológic i materials culturals**

Les variables bioarqueológiques recullen la presència i les característiques de les restes humanes: NMI (Nombre Mínim d'Individus), connexió anatòmica, momificació, fardells funeraris, restes disperses, posició flexionada i cremació d'ossos. La distinció entre momificació (conservació de teixits tous) i fardells funeraris (envolcall tèxtil, sense implicar necessàriament momificació) és metodológicament important per a la interpretació de les pràctiques funeràries. L'estructura DWS1-EF40 exemplifica la coexistència de tots dos: cossos momificats dins fardells tèxtils, en estat de preservació excepcional però en risc per l'exposició a la intempèrie (Ribera-Torró 2023: Annex VII).

Els materials culturals s'han codificat com a camps binaris independents (Mat_Textils, Mat_Fusta_Cultural, Mat_Fibra_Vegetal, Mat_Ceramica, Mat_Fauna, Mat_Banya_Cervol, Mat_Altres). La ceràmica és particularment significativa perquè és excepcional en les tombes aèries de DW i LP, però present a la cova subterrània del Sector 3 de Diablo Wasi —la qual cosa podria indicar un ús diferencial de l'espai— (Ribera-Torró 2023: Annex VII).

### **4.6. Coordenades espacials i volumetria**

Els camps de coordenades espacials (Coord_Lat_WGS84, Coord_Lon_WGS84, Coord_E_UTM, Coord_N_UTM, Altitud_msnm) vinculen cada registre de la base de dades amb el model fotogramètric georeferenciat de Metashape, seguint el flux de treball descrit a la secció 6. La font preferent de coordenades per a les estructures individuals és l'exportació de marcadors des de Metashape, que proporciona posicions en UTM zona 18S (WGS84) directament mesurades sobre el model 3D.

L'anàlisi volumètrica (Volum_Interior_m3, Volum_Total_m3, Metode_Volum) és un dels objectius específics del TFM (OE2: 'Analitzar volums i relacions espacials'). El volum interior es calcula de forma diferenciada per a estructures de geometria regular (mètode L×A×H ajustat per espessor de parets) i per a cavitats irregulars (extracció directa del model 3D via MeshLab o CloudCompare). El camp Metode_Volum registra el procediment emprat i permet ponderar la comparabilitat entre mesures.

### **4.7. Cronologia i campanyes de documentació**

Les datacions radiocarbòniques es gestionen a través de la taula secundària T_DATACIONS (N:1 amb T_ESTRUCTURES), que permet registrar múltiples dates C14 per estructura amb tots els paràmetres tècnics: data BP, intervals calibrats a 1σ i 2σ, referència de laboratori i mostra datada. El camp resumen Crono_Segle_Ini / Crono_Segle_Fi a T_ESTRUCTURES permet aproximacions ràpides per a anàlisi cronológica sense recórrer al join amb T_DATACIONS.

El camp Campanya_Documentacio registra l'any de la primera campanya que va documentar sistemàticament cada estructura (2013: PALP I; 2016: PALP II; 2021: La Petaca Project / documentació no invasiva integral; 2023: PALP IV / excavació i reconstruccions 3D). Permet analitzar l'evolució de la cobertura documental i identificar possibles biaixos d'arxiu en l'anàlisi estadística.

## **5. Estratègia de jerarquia i clustering**

### **5.1. Dos mecanismes complementaris**

Les necròpolis de penya-segat presenten una complexa estructura d'organització espacial que no pot ser representada per una arquitectura de base de dades plana. Per exemple: un mausoleu construït a l'interior d'una cova gran (CAV) és físicament contingut per aquesta, però alhora pot pertànyer a un alineament vertical de quatre estructures al llarg d'una grieta. La base de dades gestiona aquestes dues dimensions de la jerarquia amb mecanismes independents i complementaris.

### **5.2. Mecanisme A: contenció (ID_Estructura_Parent)**

El camp ID_Estructura_Parent és un camp autoreferenciant (FK que apunta al propi camp T_ESTRUCTURES.ID). Representa la relació d'inclusió física o estructural: A conté B, o A és un component estructural de B. Permet jerarquies de N nivells sense límit teòric.

Exemples d'ús en el jaciment:

- CAV (DWS1-EF35, gran cova) → conté → EA-MAU (mausoleu construït a l'interior).
- EA-PLA-V (plataforma volada) → conté com a components → MEN (mènsules que la componen).
- CAV (Sector 3) → conté → NIX (nínxol d'ossari secundari).
- PR (DWS1-P11) → associada (parent) → EA (estructura desapareguda per sobre o sota).
### **5.3. Mecanisme B: agrupació funcional (T_CONJUNTS)**

La taula T_CONJUNTS registra agrupacions d'elements arqueológics que formen una unitat analítica però cap dels quals conté l'altre. Cada estructura apunta al seu conjunt via el camp ID_Conjunt. La taula T_CONJUNTS inclou un Tipus_Conjunt (linked a L_TIPUS_CONJUNT) que classifica l'agrupació:

| **Tipus_Conjunt** | **Casuística** | **Rellevància per al TFM** |
|---|---|---|
| Alineament vertical | Estructures en una mateixa vertical sobre grieta o estrats successius. | H03 (saturació espacial): documenta l'ocupació acumulativa en l'eix vertical. |
| Conjunt de repisa | Múltiples EAs compartint una mateixa repisa horitzontal. | H01 (planificació): ¿la distribució sobre una repisa és aleatòria o planificada? |
| Plataforma amb mènsules | Repisa volada + les mènsules que la composen o flanquegen. | H01 + H02: relació entre infraestructura i topografia. |
| Xarxa de circulació | MENs i EA-PLAs que reconstrueixen trams de circulació aèria perduda. | OE2 (relacions espacials): base del Mapa d'Infraestructura Perduda. |
| Conjunt pintural | PR + EA adjacent visualment o funcionalment vinculada. | H04 (seqüència operativa): ¿la pintura precedeix o segueix la construcció? |

*Taula 4. Tipus d'agrupació funcional (L_TIPUS_CONJUNT) i rellevància per a les hipòtesis del TFM.*

### **5.4. Combinació dels dos mecanismes**

Una estructura pot tenir simultàniament un pare (ID_Estructura_Parent) i un conjunt (ID_Conjunt), ja que responen a dimensions analítiques independents. Exemple: un mausoleu (EA-MAU) construït a l'interior d'una cova (CAV, parent) i que forma part, juntament amb tres mausoleus més, d'un alineament vertical (T_CONJUNTS, conjunt). El pare expressa la contenció física; el conjunt, l'agrupació tipológica i l'anàlisi de la distribució espacial.

*Regla pràctica: usar ID_Estructura_Parent quan un element no té sentit arqueológic independent del seu contenidor. Usar ID_Conjunt quan els elements són arqueológicament independents però formen una unitat d'estudi. En cas de dubte, preferir ID_Conjunt —és menys restrictiu i no implica dependència física.*

## **6. Integració espacial amb QGIS**

### **6.1. El repte del faralló vertical**

La naturalesa vertical dels farallons on s'emplacen La Petaca i Diablo Wasi planteja un repte inusual per als SIG convencionals: múltiples estructures comparteixen pràcticament les mateixes coordenades X i Y però difereixen substancialment en la coordenada Z (altitud). El treball en QGIS s'articula, per tant, en dos nivells complementaris.

A nivell de visualització i anàlisi relativa, s'importa a QGIS l'ortofoto vertical de cada sector exportada des de Metashape (o les gigafotos ja disponibles) com a raster de referència. Sobre aquest raster s'efectua la digitalització de contorns i la identificació de relacions espacials entre estructures. Aquest és el nivell de treball habitual i permet mesurar distàncies entre estructures sobre el pla del faralló, identificar alineaments i reconstruir xarxes de circulació a partir de les mènsules aïllades.

A nivell de posicionament absolut, les coordenades UTM 18S (WGS84) i l'altitud de cada estructura s'obtenen directament dels models fotogramètrics georeferencitats de Metashape, a través de l'exportació de marcadors (File → Export → Export Markers). Aquestes coordenades es transfereixen als camps Coord_E_UTM, Coord_N_UTM i Altitud_msnm de la base de dades.

### **6.2. Flux de treball recomanat**

- Metashape: col·locar marcadors a l'accés de cada estructura en el model georeferenciat. Exportar com a CSV (coordenades UTM 18S + altitud).
- Access: importar el CSV i omplir els camps de coordenades a T_ESTRUCTURES. Registrar el mètode a Coord_Metode.
- QGIS (QRY_07): executar la consulta d'exportació espacial i exportar a CSV o Excel.
- QGIS: carregar el CSV com a capa de punts (Layer → Add Delimited Text Layer; X = Coord_Lon_WGS84, Y = Coord_Lat_WGS84, CRS = EPSG:4326). Reprojectar a UTM 18S (EPSG:32718) per a mesures de distàncies.
- Anàlisi: afegir simbologia per tipologia, anàlisi de densitat (Kernel), visibilitat i distribució vertical (Altitud_msnm vs tipus d'estructura).

*Nota a llarg termini: si la integració SIG es torna central al projecte, es recomana migrar la base de dades a format GeoPackage (.gpkg), que és natiu de QGIS i suporta geometries 3D (punts amb cota Z). Permet emmagatzemar múltiples capes i atributs relacionals en un únic fitxer, mantenint Access com a interfície d'entrada de dades amb sincronitzacions periòdiques.*

## **7. Anàlisi volumètrica**

L'anàlisi volumètrica, corresponent a l'objectiu específic OE2 del TFM, se centra en el camp Volum_Interior_m3 de T_ESTRUCTURES. La distinció entre volum interior (espai útil funerari) i volum total (inclou gruix de parets) és metodológicament crítica per a la comparació entre estructures de tipologies diverses.

El procediment difereix per a estructures de geometria regular i irregular:

- Geometria regular (EA-MAU, EA-PLA): Volum = (Largo_m - 2×e) × (Ancho_m - 2×e) × Alto_m, on e és l'espessor de les parets (~0.20-0.30 m documentat a DW). Metode_Volum = 'Càlcul L×A×H'.
- Geometria irregular (EA-CAM, CAV, NIX): extracció directa del model fotogramètric 3D via MeshLab (Filters → Measure → Compute Geometric Measures) o CloudCompare. Metode_Volum = 'Model fotogramètric'.

L'anàlisi estadística del volum interior permet abordar: la distribució del volum per tipologia i jaciment (Kruskal-Wallis), la correlació entre volum i NMI (Spearman), i la variació del volum al llarg del temps (correlació Crono_Segle_Ini ~ Volum_Interior_m3).

## **8. Consultes SQL per a l'anàlisi estadística**

La base de dades inclou deu consultes SQL predissenyades (QRY_01 a QRY_10) que cobreixen els principals anàlisis contemplats al TFM. A continuació es resumeixen les més rellevants des del punt de vista de les hipòtesis de recerca.

| **Consulta** | **Funció analítica** | **Hipòtesis relacionades** |
|---|---|---|
| QRY_01 | Distribució de tipologies per jaciment i sector. | H01, H04 |
| QRY_02 | Taula de presència de decoració (χ² LP vs DW). | H01, H04 |
| QRY_03 | Distribució de l'estat de conservació per sector. | General |
| QRY_04 | Estructures datades per C14 ordenades cronológicament. | H04 |
| QRY_05 | Exportació plana per a R / SPSS (inclou volum i coordenades). | Tots |
| QRY_06 | Volum interior mig per tipologia i jaciment. | OE2, H01 |
| QRY_07 | Exportació espacial per a QGIS (coordenades + atributs). | OE2, H02, H03 |
| QRY_08 | Contingut d'un element pare (elements fills d'una cova, per exemple). | H01, H02 |
| QRY_09 | Membres d'un conjunt funcional (per xarxa de circulació o alineament). | H03, OE2 |
| QRY_10 | Cobertura Chacha XR vs. total per jaciment i campanya. | Metodologia |

*Taula 5. Resum de les deu consultes SQL predissenyades i la seua vinculació amb les hipòtesis del TFM.*

L'exportació per a anàlisi estadística externa es realitza via QRY_05 (exportació plana completa) a format Excel o CSV des de Access (External Data → Export → Text File), i posteriorment s'importa a R o SPSS. En R, el paquet readxl (Excel) o read.csv (CSV) permet la càrrega directa. Tots els camps categòrics de la consulta estan en format de text llegible (no com a IDs numèrics), gràcies als JOINs amb les taules lookup integrats a QRY_05.

## **9. Vincle amb la documentació digital (Chacha XR)**

El camp Documentat_ChaXR estableix un pont entre la base de dades arqueológica i la plataforma multimèdia pública Chacha XR (Ribera-Torró et al. 2026). Chacha XR és un projecte de documentació i difusió digital desenvolupat entre 2021 i 2023 que integra models fotogramètrics 3D, panorames 360°, gigafotos i vídeo 360° dels jaciments en una experiència immersiva multiplaforma (Ribera-Torró 2023).

La codificació d'aquest camp permet analitzar la representativitat de la mostra publicada respecte al total documentat. Des de la perspectiva de l'anàlisi estadística, és important avaluar si les estructures accessibles a Chacha XR constitueixen una mostra biaixada del total (per exemple, si es van seleccionar preferentment estructures de tipologies concretes o en millor estat de conservació), ja que aquest biaix podria condicionar la interpretació dels resultats disponibles a la plataforma pública.

Els camps URL_Pano, URL_Pano_2, URL_Giga i URL_3D permeten accedir directament, des de la base de dades, als recursos digitals de cada estructura. Algunes estructures compten amb múltiples escenes panoràmiques (ex.: DWS1-EF40 té 3 nodes independents a l'app; DWS1-EF02 té una escena exterior i una interior), per la qual cosa el camp URL_Pano_2 recull la URL de la segona escena. Si en el futur el nombre d'arxius per estructura creix de forma generalitzada, es recomana implementar la taula T_MEDIA (descrita a l'esquema complet de la BD) per gestionar múltiples recursos per estructura.

## **10. Implementació a Microsoft Access**

L'elecció de Microsoft Access com a motor de la base de dades respon a criteris de disponibilitat (inclòs en la llicència Microsoft 365), facilitat d'entrada de dades a través de formularis visuals i compatibilitat amb les eines d'exportació a Excel i CSV. Access utilitza el motor Jet SQL / ACE, amb algunes limitacions respecte a SQL Server (absència de subconsultes correlacionades, sintaxi específica per a joins múltiples), que s'han tingut en compte en la redacció de les consultes.

Per a la implementació es recomana seguir l'ordre: (1) crear les taules lookup, (2) crear les taules secundàries (T_DATACIONS, T_INDIVIDUS, T_CONJUNTS), (3) crear T_ESTRUCTURES, (4) definir les relacions a Herramientas de base de datos → Relaciones, activant integritat referencial i actualización en cascada. Tots els camps categòrics de T_ESTRUCTURES han de representar-se com a ComboBox en els formularis, amb les taules lookup corresponents com a font de dades —la qual cosa elimina els errors de tipografia i unifica el vocabulari.

Les persones que entren dades han d'adoptar la convenció de codificació de 'ND' (No Determinat) per als camps categòrics en els quals no hi ha informació disponible, i NULL (camp buit) per als camps numèrics (NMI, dimensions, volum) quan la dada és absent. Aquesta distinció entre 'no s'ha pogut determinar' i 'no existeix' és crucial per a la interpretació estadística.

## **11. Limitacions i perspectives de futur**

La principal limitació de la base de dades en la seua fase actual és la manca de dades per a moltes de les estructures de La Petaca Sector Sud (≥96 estructures identificades però no totes amb fitxa completa). L'ompliment sistemàtic de T_ESTRUCTURES amb dades de camp és, per tant, una tasca pendent que condicionarà l'abast de les anàlisis estadístiques del TFM. El camp Campanya_Documentacio permet traçar de forma transparent quines estructures han estat documentades i en quina campanya, facilitant la detecció de buits documentals.

Quant a les coordenades espacials, la dependència de les coordenades individuals de cada estructura respecte als models Metashape georeferencitats significa que, per a les estructures documentades en campanyes anteriors a 2021 (quan no s'aplicaven sistemàticament tècniques de georreferenciació), les coordenades han de ser estimades o derivades de fonts secundàries (GPS de camp, ortofoto general). El camp Coord_Precisio_m permet registrar i ponderar aquesta incertesa.

De cara a la fase final del TFM i a publicacions futures, es plantegen tres extensions de la base de dades: (a) la implementació de la taula T_MEDIA per a la gestió de múltiples arxius digitals per estructura, (b) la migració a GeoPackage per a una integració nativa amb QGIS, i (c) l'afegit d'una taula T_CAMPANYES amb informació sistematitzada de cada campanya PALP (directors, finançament, nombre d'estructures documentades, publicacions derivades).

## **12. Referències**

Epstein, L.; Toyne, J.M. (2016). When Space Is Limited: A Spatial Exploration of Pre-Hispanic Chachapoya Mortuary and Ritual Microlandscape. In Osterholtz, A.J. (ed.), Theoretical Approaches to Analysis and Interpretation of Commingled Human Remains. Springer, Switzerland, pp. 97-124. DOI 10.1007/978-3-319-22554-8_6.

Ribera-Torró, E. (2023). Chacha XR. Una experiència immersiva de no-ficció per l'arqueologia Chachapoya. Treball de Fi de Màster, Màster Universitari en Arts Visuals i Multimèdia, Universitat Politècnica de València.

Ribera-Torró, E.; Toyne, J.M.; Del Águila, R.; Ribera, J.A.; Galexner, J.; Anzellini, A.; Pans, M. (2026). Extended Reality on Chachapoya Cliffside Necropolises: From Digital Documentation to Public Engagement. Open Archaeology, 12(1). DOI 10.1515/opar-2025-0071.

Toyne, J.M.; Anzellini, A. (2017). Sociedad, identidad y variedad en los mausoleos de La Petaca, Chachapoyas. Boletín de Arqueología PUCP, 23, 231-257.

Toyne, J.M.; Anzellini, A.; Epstein Mičulka, L.; Mejías Pitti, I.; Puig Castell, J.; Guinot Castelló, S. (2018). Going Vertical: Using Vertical Progression Techniques to Explore a Cliff Necropolis in Late Precolumbian Chachapoyas, Peru. Advances in Archaeological Practice, 6. DOI 10.1017/aap.2018.31.
