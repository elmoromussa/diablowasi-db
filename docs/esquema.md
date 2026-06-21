**ESQUEMA DE BASE DE DADES**

Estructures funeràries — La Petaca & Diablo Wasi

Base de dades Microsoft Access per a anàlisi estadística

Panograma Labs · Projecte La Petaca & Diablo Wasi

## 1. Introducció i objectius

Aquest document especifica l'esquema de base de dades relacional per a Microsoft Access destinat a la documentació sistemàtica i l'anàlisi estadística de les estructures funeràries dels jaciments de La Petaca i Diablo Wasi (Amazonas, Perú). L'esquema ha estat dissenyat a partir de l'anàlisi de la documentació existent (fitxes d'estructures, borrador d'article i papers publicats) i cobreix les variables arquitectòniques, decoratives, bioarqueológiques i cronológiques necessàries per a l'estudi comparatiu entre jaciments.

Objectius principals: (1) centralitzar la informació de totes les estructures en una única base de dades estructurada; (2) facilitar l'exportació de dades per a anàlisis estadístiques en programes externs (Excel, SPSS, R); (3) permetre consultes comparatives entre sectors i jaciments.

## 2. Estructura general de la base de dades

La base de dades es compon de tres tipus de taules:

(A) Taules de dades principals — contenen les fitxes arqueológiques pròpiament dites.

(B) Taules de datació — registren les dates de radiocarboni associades a cada estructura.

(C) Taules de consulta (Lookup) — llistes tancades de valors permesos per als camps categòrics, que eviten errors d'entrada i unifiquen el vocabulari.

**Resum de taules**

| **Taula** | **Tipus** | **Contingut** |
|---|---|---|
| **T_ESTRUCTURES** | Principal | Fitxa completa de cada estructura funerària (una fila per estructura). Font principal per a estadística. |
| **T_DATACIONS** | Secundari | Dates de radiocarboni (C14) associades a cada estructura. Relació N:1 amb T_ESTRUCTURES. |
| **T_INDIVIDUS** | Secundari | Dades individuals de restes esquelètiques (quan es disposa d'informació per individu). |
| **L_SITIOS** | Lookup | Llista de jaciments: La Petaca, Diablo Wasi. |
| **L_SECTORES** | Lookup | Llista de sectors per jaciment (Nord, Central, Sur, Sector 1-6...). |
| **L_TIPOLOGIA** | Lookup | Tipologia funcional/arquitectònica de cada estructura. |
| **L_SUPORT** | Lookup | Tipus de suport geomorfològic de l'estructura. |
| **L_ESTAT** | Lookup | Estat de conservació (bo, regular, pre-col·lapse, col·lapsat, ND). |

## 3. T_ESTRUCTURES — Taula principal

Taula central de la base de dades. Cada fila correspon a una estructura funerària documentada (mausoleu, cambra funerària, plataforma, pintura rupestre, etc.). Totes les anàlisis estadístiques principals s'executen sobre aquesta taula o sobre vistes/consultes derivades d'ella.

**3.1 Identificació i localització**

| **Camp / Field** | **Tipus Access** | **Obligatori** | **Descripció i valors permesos** |
|---|---|---|---|
| ID | AutoNumber | Sí (PK) | Clau primària, generada automàticament per Access. |
| Codi | Text (20) | Sí | Codi original de l'estructura (ex: 'PTC-SN-EF01', 'DWS1-EF12'). |
| ID_Sector | Number (Long) | Sí (FK) | Referència a L_SECTORES. Inclou implícitament el jaciment. |

**3.2 Tipologia arquitectònica**

| **Camp / Field** | **Tipus Access** | **Obligatori** | **Descripció i valors permesos** |
|---|---|---|---|
| ID_Tipologia | Number (Long) | Sí (FK) | Referència a L_TIPOLOGIA. Valors: Mausoleu/Chullpa, Cambra funerària, Plataforma volada, Plataforma sobre repisa, Cova natural, Pintura rupestre, Mixt, ND. |
| ID_Suport | Number (Long) | Sí (FK) | Referència a L_SUPORT. Tipus de base geomorfològica. |
| N_Pisos | Number (Integer) | No | Nombre de pisos constructius (0, 1, 2, 3). 0 = no aplicable. |
| Planta | Text (20) | No | Planta arquitectònica: Rectangular / Quadrada / Trapezoidal / Irregular / No aplica / ND. |
| N_Murs_Construits | Number (Byte) | No | Nombre de murs construïts: 0, 1, 2, 3, 4. (Les cambres funeràries solen tenir 1; els mausoleus, 3.) |
| Orientacio_Vano | Text (5) | No | Orientació del vano/accés: N, S, E, O, NE, NO, SE, SO, No aplica, ND. |
| Dintel | Text (20) | No | Material del dintell: Pedra / Fusta / Sense dintell / ND. |
| Techo_Natural | Yes/No | No | Presència de tecto/protecció natural de roca. Cert = protegit. |
| Contraforts | Yes/No | No | Presència de contraforts a la façana. |
| Cornisa_Entre_Pisos | Yes/No | No | Cornisa de lloses planes entre pisos. |
| Estacas_Fusta | Yes/No | No | Presència d'estacas o pals de fusta (elements estructurals o d'ancoratge). |

**3.3 Dimensions i posició**

| **Camp / Field** | **Tipus Access** | **Obligatori** | **Descripció i valors permesos** |
|---|---|---|---|
| Largo_m | Number (Single) | No | Longitud de la base en metres (dimensió horitzontal major). |
| Ancho_m | Number (Single) | No | Amplada de la base en metres. |
| Alto_m | Number (Single) | No | Alçada total de l'estructura en metres. |
| Altura_Aprox_m | Number (Single) | No | Alçada aproximada sobre el nivell del sòl (metres). Pot ser estimació. |

**3.4 Acabats superficials i pintura**

| **Camp / Field** | **Tipus Access** | **Obligatori** | **Descripció i valors permesos** |
|---|---|---|---|
| Arrebossat | Yes/No | No | Presència d'arrebossat / enlluïment en murs. |
| Color_Arrebossat | Text (20) | No | Color del arrebossat: Blanc / Roig / Ambdós / ND. Actiu si Arrebossat = Sí. |
| Pintura_Sobre_Roca | Yes/No | No | Pintura aplicada directament sobre la roca del faralló (forma de C, U, semicercle o gran taca). |
| Color_Pintura_Roca | Text (20) | No | Color de la pintura sobre roca: Roig / Blanc / Ambdós / ND. |

**3.5 Decoració arquitectònica en baix-relleu**

| **Camp / Field** | **Tipus Access** | **Obligatori** | **Descripció i valors permesos** |
|---|---|---|---|
| Dec_Nicho_Quadrat | Yes/No | No | Nínxols quadrats en sèrie (disposats en una o dues files horitzontals). |
| Dec_Relieve_T | Yes/No | No | Relleu en forma de T. |
| Dec_Relieve_T_Inv | Yes/No | No | Relleu en forma de T invertida. |
| Dec_Relieve_L | Yes/No | No | Relleu en forma de L. |
| Dec_Relieve_L_Inv | Yes/No | No | Relleu en forma de L invertida. |
| Dec_Zigzag | Yes/No | No | Fris o motiu en zigzag. |
| Dec_Escalonat | Yes/No | No | Motiu escalonat (símbol andí, rar a DW i LP). |
| Dec_Fris_Greca | Yes/No | No | Greca o fris de nínxols (especialment a La Petaca). |

**3.6 Pintura rupestre associada**

| **Camp / Field** | **Tipus Access** | **Obligatori** | **Descripció i valors permesos** |
|---|---|---|---|
| Pintura_Rupestre | Yes/No | No | Presència de pintura rupestre associada a l'estructura o en el seu entorn immediat. |
| PR_Antropomorfa | Yes/No | No | Figura antropomorfa (ex: personatge decapitador, braços estesos). |
| PR_Zoomorfa | Yes/No | No | Figura zoomorfa (ex: cèrvol/venado, aus). |
| PR_Geometrica | Yes/No | No | Motiu geomètric (ex: cercle, L invertida). |
| PR_Abstracta | Yes/No | No | Taques o restes de pintura sense forma identificable. |
| PR_Escena_Decap | Yes/No | No | Escena de decapitació (present a PTC SS EF18). |

**3.7 Estat de conservació i alteracions**

| **Camp / Field** | **Tipus Access** | **Obligatori** | **Descripció i valors permesos** |
|---|---|---|---|
| ID_Estat | Number (Long) | Sí (FK) | Referència a L_ESTAT. Estat general de conservació. |
| Saqueig | Yes/No | No | Evidència de saqueig (forats a les parets, restes humanes disperses, absència de béns culturals...). |
| Incendi | Yes/No | No | Evidència d'incendi o cremació del context (no dels ossos; veure Cremacio). |
| Activitat_Animal | Yes/No | No | Evidència d'activitat animal: nius, excrements, ossos d'aus, etc. |

**3.8 Restes humanes i context bioarqueológic**

| **Camp / Field** | **Tipus Access** | **Obligatori** | **Descripció i valors permesos** |
|---|---|---|---|
| Restes_Humanes | Yes/No | No | Presència de restes humanes (mínimament uns pocs ossos). |
| NMI | Number (Integer) | No | Nombre Mínim d'Individus estimat. 0 = cap/indeterminat. |
| Connexio_Anatomica | Yes/No | No | Restes en connexió anatòmica (total o parcial). |
| Momificacio | Yes/No | No | Evidència de momificació (teixits tous, pell, cabells conservats). |
| Fardells_Funeraris | Yes/No | No | Presència de fardells / bults funeraris (envolcalls tèxtils, posició fetal). |
| Restes_Disperses | Yes/No | No | Ossos dispersos, desarticulats o en posició secundària. |
| Posicio_Flexionada | Yes/No | No | Individus en posició flexionada (fetal). |
| Cremacio_Ossos | Yes/No | No | Evidència de cremació sobre les restes esquelètiques. |

**3.9 Materials culturals associats**

| **Camp / Field** | **Tipus Access** | **Obligatori** | **Descripció i valors permesos** |
|---|---|---|---|
| Mat_Textils | Yes/No | No | Tèxtils (fragments de roba, cordills, soguilles). |
| Mat_Fusta_Cultural | Yes/No | No | Fusta treballada amb funció cultural (pals de l'estructura exclosos). |
| Mat_Fibra_Vegetal | Yes/No | No | Fibra vegetal (cordell, cistelleria, etc.). |
| Mat_Ceramica | Yes/No | No | Ceràmica (rar en tombes aèries; present en coves). |
| Mat_Fauna | Yes/No | No | Restes de fauna (plomes, ossos d'animals, etc.). |
| Mat_Banya_Cervol | Yes/No | No | Banya de cèrvol incrustada en la paret (element cosmológic chachapoya). |
| Mat_Altres | Yes/No | No | Altres materials no categoritzats anteriorment. Detallar a Notes. |

**3.10 Cronologia**

| **Camp / Field** | **Tipus Access** | **Obligatori** | **Descripció i valors permesos** |
|---|---|---|---|
| C14 | Yes/No | No | Existeix datació de radiocarboni per a aquesta estructura. |
| Crono_Segle_Ini | Number (Integer) | No | Inici del rang cronològic en segles d.n.e. (enter positiu; ex: 9 = s. IX). Si és a.n.e., usar negatiu. |
| Crono_Segle_Fi | Number (Integer) | No | Fi del rang cronològic en segles d.n.e. (ex: 11 = s. XI). |

**3.11 Documentació digital i notes**

| **Camp / Field** | **Tipus Access** | **Obligatori** | **Descripció i valors permesos** |
|---|---|---|---|
| URL_Pano | Hyperlink | No | Adreça URL al panorama 360° associat (xr.chachapoya.org). |
| URL_Giga | Hyperlink | No | Adreça URL a la imatge gigapíxel associada. |
| URL_3D | Hyperlink | No | Adreça URL al model 3D a Sketchfab. |
| Notes | Memo (Long Text) | No | Observacions, incidències, dubtes de classificació o informació addicional. |

## 4. T_DATACIONS — Datacions de radiocarboni

Taula relacionada (N:1 amb T_ESTRUCTURES). Una estructura pot tenir múltiples dates C14 (de mostres de fusta, os o carbó). Permet recollir tota la informació tècnica de cada datació.

**Camps de T_DATACIONS**

| **Camp / Field** | **Tipus Access** | **Obligatori** | **Descripció i valors permesos** |
|---|---|---|---|
| ID | AutoNumber | Sí (PK) | Clau primària. |
| ID_Estructura | Number (Long) | Sí (FK) | Referència a T_ESTRUCTURES.ID. |
| Mostra_Tipus | Text (30) | No | Tipus de mostra datada: fusta, os, carbó, teixit, etc. |
| Data_BP | Number (Long) | No | Data convencional en anys BP (Before Present, 1950). |
| Sigma1_Ini | Number (Integer) | No | Inici de l'interval calibrat a 1 sigma (cal d.n.e.). |
| Sigma1_Fi | Number (Integer) | No | Fi de l'interval calibrat a 1 sigma (cal d.n.e.). |
| Sigma2_Ini | Number (Integer) | No | Inici de l'interval calibrat a 2 sigma (cal d.n.e.). |
| Sigma2_Fi | Number (Integer) | No | Fi de l'interval calibrat a 2 sigma (cal d.n.e.). |
| Ref_Laboratori | Text (50) | No | Referència de laboratori (ex: 'Beta-123456', 'AA-78901'). |
| Referencia_Bibliografica | Memo | No | Referència bibliogràfica on es publica la data. |

## 5. T_INDIVIDUS — Dades individuals

Taula opcional (N:1 amb T_ESTRUCTURES). A utilitzar quan es disposa d'informació desagregada per individu (edat, sexe, estat de preservació). Complementa la informació agregada del camp NMI a T_ESTRUCTURES.

**Camps de T_INDIVIDUS**

| **Camp / Field** | **Tipus Access** | **Obligatori** | **Descripció i valors permesos** |
|---|---|---|---|
| ID | AutoNumber | Sí (PK) | Clau primària. |
| ID_Estructura | Number (Long) | Sí (FK) | Referència a T_ESTRUCTURES.ID. |
| Num_Individu | Number (Integer) | No | Número d'individu dins de l'estructura (1, 2, 3...). |
| Edat_Categoria | Text (20) | No | Categoria d'edat: Infantil (<12a) / Juvenil (12-20a) / Adult (20-50a) / Senil (>50a) / ND. |
| Sexe_Categoria | Text (20) | No | Sexe estimat: Femení / Masculí / Indeterminat / ND. |
| Estat_Preservacio | Text (20) | No | Estat general de preservació: Bo / Regular / Fragmentari / ND. |
| Notes | Memo | No | Observacions sobre l'individu (patologies, traces, etc.). |

## 6. Taules de consulta (Lookup)

Les taules de consulta contenen llistes tancades de valors permesos. A Access s'utilitzen com a font de dades per a camps de llista desplegable (Combo Box), garantint la uniformitat terminològica i evitant errors d'entrada. Cada taula té com a mínim un camp ID (AutoNumber, PK) i un camp Nom (Text, valor visible).

### 6.1 L_SITIOS

**Valors de L_SITIOS**

| **Valor** | **Descripció** |
|---|---|
| La Petaca | Necròpolis de La Petaca. ~12.000 m² de roca exposada, 4 sectors. |
| Diablo Wasi | Necròpolis de Diablo Wasi. 6 sectors, predomini de cavitats. |

### 6.2 L_SECTORES

**Valors de L_SECTORES (relacionat amb L_SITIOS)**

| **Valor** | **Descripció** |
|---|---|
| LP - General | La Petaca — Sector General (context de tot el jaciment). |
| LP - Nord | La Petaca — Sector Nord (mida menor, 1-2 estructures documentades). |
| LP - Central | La Petaca — Sector Central. |
| LP - Superior | La Petaca — Sector Superior. |
| LP - Sud | La Petaca — Sector Sud (el més gran, ≥96 estructures documentades). |
| DW - General | Diablo Wasi — Context general. |
| DW - Sector 1 | Diablo Wasi — Sector 1 (el més important, ~40 contextos funeraris). |
| DW - Sector 2 | Diablo Wasi — Sector 2. |
| DW - Sector 3 | Diablo Wasi — Sector 3 (inclou cova subterrània). |
| DW - Sector 4 | Diablo Wasi — Sector 4 (~10 contextos funeraris). |
| DW - Sector 5 | Diablo Wasi — Sector 5. |
| DW - Sector 6 | Diablo Wasi — Sector 6. |

### 6.3 L_TIPOLOGIA

**Valors de L_TIPOLOGIA**

| **Valor** | **Descripció** |
|---|---|
| Mausoleu / Chullpa | Estructura construïda (3+ murs + sostre artificial) sobre repisa. Generalment 2 pisos. Predominant a LP. |
| Cambra funerària | Estructura que aprofita una cavitat preexistent, tancada per una única façana construïda. |
| Plataforma sobre repisa | Plataforma constructiva sobre repisa natural, sense estructura funerària visible (destruïda o de trànsit). |
| Plataforma volada | Plataforma artificial sobre fustes/lloses, sense repisa natural de suport. |
| Cova natural | Espai cavernós natural sense estructura construïda, amb ús funerari o ritual. |
| Pintura rupestre | Motiu pictòric sobre roca, sense estructura arquitectònica associada. |
| Mixt | Estructura que combina característiques de dues o més categories anteriors. |
| ND | No determinat — informació insuficient per a classificar. |

### 6.4 L_SUPORT

**Valors de L_SUPORT**

| **Valor** | **Descripció** |
|---|---|
| Repisa natural àmplia (>2m) | Repisa de més de 2 metres de profunditat, permet circulació i construcció. |
| Repisa natural estreta (<2m) | Repisa de menys de 2 metres, base justa per a l'estructura. |
| Repisa artificial | Plataforma construïda sobre fustes i lloses. |
| Cavitat gran (>10m²) | Caverna de gran extensió, generalment al peu del faralló. |
| Cavitat mitjana (1-10m²) | Cavitat de mida mitjana, ampliament aprofitada a DW. |
| Nínxol natural (<1m²) | Petit nínxol o recessió, funció probable d'ossari o enterrament secundari. |
| Grieta | Fissura natural del faralló emplenada per crear una plataforma. |
| Combinat | Combinació de dues o més morfologies anteriors. |
| ND | No determinat. |

### 6.5 L_ESTAT

**Valors de L_ESTAT (estat de conservació)**

| **Valor** | **Descripció** |
|---|---|
| Bo | Estructura íntegra o gairebé íntegra, parets dempeus, sostre conservat. |
| Regular | Estructura parcialment deteriorada però reconeixible en la seva major part. |
| Pre-col·lapse | Estructura en risc imminent de col·lapse: parets inclinades, grans fissures, elements despresos. |
| Col·lapsat | Estructura derrumbada o fortament fragmentada. |
| ND | No determinat — informació insuficient. |

## 7. Consultes SQL recomanades per a anàlisi estadística

Les consultes següents s'han de crear com a objectes de Consulta a Access (Query Design / SQL View). A continuació es mostren en sintaxi SQL estàndard compatible amb Access (JET SQL). Un cop creades, les dades es poden exportar a Excel o CSV per a anàlisi en R, SPSS o Python.

**QRY_01 — Taula de contingència: Tipologia per jaciment**

SELECT L_SITIOS.Nom_Sitio, L_TIPOLOGIA.Nom AS Tipologia, COUNT(T_ESTRUCTURES.ID) AS N
FROM ((T_ESTRUCTURES
  INNER JOIN L_SECTORES ON T_ESTRUCTURES.ID_Sector = L_SECTORES.ID)
  INNER JOIN L_SITIOS ON L_SECTORES.ID_Sitio = L_SITIOS.ID)
  INNER JOIN L_TIPOLOGIA ON T_ESTRUCTURES.ID_Tipologia = L_TIPOLOGIA.ID
GROUP BY L_SITIOS.Nom_Sitio, L_TIPOLOGIA.Nom
ORDER BY L_SITIOS.Nom_Sitio, L_TIPOLOGIA.Nom;

**QRY_02 — Presència de decoració arquitectònica per jaciment (per a χ²)**

SELECT L_SITIOS.Nom_Sitio,
  SUM(IIF(Dec_Nicho_Quadrat=True,1,0)) AS N_Nicho,
  SUM(IIF(Dec_Relieve_T=True,1,0)) AS N_T,
  SUM(IIF(Dec_Zigzag=True,1,0)) AS N_Zigzag,
  SUM(IIF(Dec_Escalonat=True,1,0)) AS N_Escalonat,
  SUM(IIF(Pintura_Rupestre=True,1,0)) AS N_Pintura,
  COUNT(T_ESTRUCTURES.ID) AS Total
FROM (T_ESTRUCTURES
  INNER JOIN L_SECTORES ON T_ESTRUCTURES.ID_Sector = L_SECTORES.ID)
  INNER JOIN L_SITIOS ON L_SECTORES.ID_Sitio = L_SITIOS.ID
GROUP BY L_SITIOS.Nom_Sitio;

**QRY_03 — Estat de conservació per sector**

SELECT L_SITIOS.Nom_Sitio, L_SECTORES.Nom_Sector, L_ESTAT.Nom AS Estat,
  COUNT(*) AS N
FROM (((T_ESTRUCTURES
  INNER JOIN L_SECTORES ON T_ESTRUCTURES.ID_Sector = L_SECTORES.ID)
  INNER JOIN L_SITIOS ON L_SECTORES.ID_Sitio = L_SITIOS.ID)
  INNER JOIN L_ESTAT ON T_ESTRUCTURES.ID_Estat = L_ESTAT.ID)
GROUP BY L_SITIOS.Nom_Sitio, L_SECTORES.Nom_Sector, L_ESTAT.Nom
ORDER BY L_SITIOS.Nom_Sitio, L_SECTORES.Nom_Sector;

**QRY_04 — Cronologia de les estructures datades per C14**

SELECT T_ESTRUCTURES.Codi, L_SITIOS.Nom_Sitio, L_SECTORES.Nom_Sector,
  L_TIPOLOGIA.Nom AS Tipologia, T_ESTRUCTURES.Crono_Segle_Ini,
  T_ESTRUCTURES.Crono_Segle_Fi,
  T_ESTRUCTURES.Crono_Segle_Fi - T_ESTRUCTURES.Crono_Segle_Ini AS Rang_Segles
FROM (((T_ESTRUCTURES
  INNER JOIN L_SECTORES ON T_ESTRUCTURES.ID_Sector = L_SECTORES.ID)
  INNER JOIN L_SITIOS ON L_SECTORES.ID_Sitio = L_SITIOS.ID)
  INNER JOIN L_TIPOLOGIA ON T_ESTRUCTURES.ID_Tipologia = L_TIPOLOGIA.ID)
WHERE T_ESTRUCTURES.C14 = True
ORDER BY T_ESTRUCTURES.Crono_Segle_Ini;

**QRY_05 — Exportació plana (flat) per a R/SPSS**

SELECT T_ESTRUCTURES.ID, T_ESTRUCTURES.Codi,
  L_SITIOS.Nom_Sitio AS Jaciment, L_SECTORES.Nom_Sector AS Sector,
  L_TIPOLOGIA.Nom AS Tipologia, L_SUPORT.Nom AS Suport,
  T_ESTRUCTURES.N_Pisos, T_ESTRUCTURES.Planta, T_ESTRUCTURES.N_Murs_Construits,
  T_ESTRUCTURES.Orientacio_Vano, T_ESTRUCTURES.Dintel,
  T_ESTRUCTURES.Techo_Natural, T_ESTRUCTURES.Contraforts,
  T_ESTRUCTURES.Arrebossat, T_ESTRUCTURES.Pintura_Sobre_Roca,
  T_ESTRUCTURES.Dec_Nicho_Quadrat, T_ESTRUCTURES.Dec_Relieve_T,
  T_ESTRUCTURES.Dec_Zigzag, T_ESTRUCTURES.Dec_Escalonat,
  T_ESTRUCTURES.Pintura_Rupestre,
  L_ESTAT.Nom AS Estat_Conservacio,
  T_ESTRUCTURES.Saqueig, T_ESTRUCTURES.Incendi, T_ESTRUCTURES.Activitat_Animal,
  T_ESTRUCTURES.Restes_Humanes, T_ESTRUCTURES.NMI,
  T_ESTRUCTURES.Momificacio, T_ESTRUCTURES.Fardells_Funeraris,
  T_ESTRUCTURES.Cremacio_Ossos,
  T_ESTRUCTURES.Mat_Textils, T_ESTRUCTURES.Mat_Ceramica,
  T_ESTRUCTURES.C14, T_ESTRUCTURES.Crono_Segle_Ini, T_ESTRUCTURES.Crono_Segle_Fi
FROM ((((T_ESTRUCTURES
  INNER JOIN L_SECTORES ON T_ESTRUCTURES.ID_Sector = L_SECTORES.ID)
  INNER JOIN L_SITIOS ON L_SECTORES.ID_Sitio = L_SITIOS.ID)
  INNER JOIN L_TIPOLOGIA ON T_ESTRUCTURES.ID_Tipologia = L_TIPOLOGIA.ID)
  INNER JOIN L_SUPORT ON T_ESTRUCTURES.ID_Suport = L_SUPORT.ID)
  INNER JOIN L_ESTAT ON T_ESTRUCTURES.ID_Estat = L_ESTAT.ID;

## 8. Notes d'implementació a Microsoft Access

A. Relacions. Crear totes les relacions a l'eina Relaciones d'Access (Herramientas de base de datos > Relaciones). Activar la integritat referencial i la actualización en cascada.

B. Formularis. Es recomana crear un formulari principal F_ESTRUCTURES (basat en T_ESTRUCTURES) amb subformularis incrustats per a T_DATACIONS i T_INDIVIDUS. Tots els camps de FK (ID_Sector, ID_Tipologia, ID_Suport, ID_Estat) han d'estar representats com a Combo Box amb les taules Lookup corresponents.

C. Camps Yes/No. A Access els camps Yes/No apareixen com a caselles de verificació (checkboxes) als formularis, cosa que agilita l'entrada de dades. Per a les consultes estadístiques, Access els converteix automàticament a -1 (Sí) i 0 (No); la funció IIF() de les consultes ho gestiona.

D. Exportació. Qualsevol consulta (Query) pot exportar-se directament a Excel (Format: Excel Workbook) o a CSV per a ser processada amb R (readxl, haven) o SPSS (File > Open > Data).

E. Codificació de la cronologia. Els camps Crono_Segle_Ini i Crono_Segle_Fi emmagatzemen enters positius per a segles d.n.e. (9 = s. IX, 16 = s. XVI) i enters negatius per a a.n.e. (si calgués). Això permet calcular l'amplada del rang cronològic directament (Crono_Segle_Fi - Crono_Segle_Ini) i ordenar cronológicament.

F. Valors 'ND' (No determinat). Tot camp categòric hauria d'incloure l'opció 'ND' a la seva taula Lookup per diferenciar 'dada absent' de 'dada negativa'. Per a camps numèrics (NMI, dimensions), el valor nul (Null / 0) indicarà absència de dada.

## 9. Anàlisis estadístics suggerits

La base de dades permetrà realitzar, entre d'altres, les anàlisis següents:

1. Freqüències i descriptiva bàsica. Distribució de tipologies per jaciment i sector. Nombre d'estructures amb decoració arquitectònica. Freqüència relativa de cada tipus de decoració (χ² per comparar La Petaca vs Diablo Wasi). Distribució de NMI per estructura.

2. Associació entre variables. Test de χ² (o Fisher exacte): Tipologia × Decoració; Jaciment × Estat de conservació; Tipologia × Presència de saqueig. Correlació de Spearman: N_Pisos ~ NMI; Alçada ~ Estat de conservació.

3. Anàlisi cronológica. Distribució temporal de les estructures datades (diagrama de Gantt o histograma de segles). Comparació de tipologies o decoracions per períodes cronológics. Tests de Kruskal-Wallis per a variables ordinals entre períodes.

4. Anàlisi de correspondències (AFC/CA). Matriu de presència/absència de decoracions (files = estructures, columnes = tipus de decoració). Permet identificar associacions entre jaciments, tipologies i elements decoratius.

5. Anàlisi de clúster. Agrupació d'estructures per semblança en el conjunt de variables binàries (decoració, materials, restes). Identificar patrons o tradicions constructives no evidents a simple vista.

— Fi del document —
