# Base de dades arqueològica de Diablo Wasi

Disseny, codi i documentació de la base de dades relacional (Microsoft Access) per a l'estudi de les estructures funeràries de la necròpolis de penya-segat Chachapoya de Diablo Wasi (Leymebamba, Amazonas, Perú, s. IX-XVI).

L'esquema es va dissenyar per a documentar dos jaciments, La Petaca i Diablo Wasi, i conserva La Petaca com a jaciment donat d'alta amb els seus sectors. Els registres, però, són tots de Diablo Wasi: 106 elements arqueològics a la base tancada v25g. La incorporació de La Petaca queda per a una versió posterior.

Treball de Fi de Màster. Màster en Arqueologia Professional i Gestió Integral del Patrimoni, Universitat d'Alacant.
Autor: Esteve Ribera Torró. Directors: Dr. Ignasi Grau Mira (UA) i Dra. J. Marla Toyne (UCF).

## Estructura del repositori

| Ruta | Contingut |
|---|---|
| `src/chachapoya_DB.bas` | Mòdul VBA que construeix la base de dades completa (taules, lookups, relacions, consultes i bateria de validació) sobre un Access en blanc. |
| `src/chachapoya_Form.bas` | Mòdul VBA que construeix el formulari d'entrada `F_STRUCTURES` (interfície en valencià, valors emmagatzemats en anglés). |
| `tools/` | Utilitats d'operació: `Worklist`, `CheckMetrics`, `Relacions`, `Tancament`, `diagnostic_v13`. |
| `tools/patches/` | Pedaços que porten una base amb dades d'una versió a la següent (v16a a v25e). |
| `tools/migrations/` | Migracions, exportadors i importadors entre versions i l'importador de georeferenciació des de Metashape. |
| `docs/esquema.md` | Referència tècnica de l'esquema vigent. |
| `docs/metodologia.md` | Justificació metodològica (capítol del TFM). |
| `docs/manual.md` | Manual d'ús del formulari. |
| `docs/deltas/` | Especificació de canvis de cada iteració. Són la font dels missatges de commit. |
| `docs/notes/` | Notes de revisió de l'autor que van donar entrada a cada delta. |
| `docs/precedents/` | Fitxa d'estructures en Excel de 2023, antecedent del disseny. |
| `data/transfer_v11_v15/` | CSV i logs de la transferència de dades entre l'esquema v11 i el v15. |
| `data/tancament_v25g/` | Fulls de tancament de NULLs que van alimentar la cascada v25g. |
| `db/chachapoya_v25g.accdb` | Base de dades tancada per a l'anàlisi. Única còpia Access del repositori. |

Els mòduls conserven el prefix `chachapoya_` amb què es van escriure. Els fitxers no porten el número de versió al nom: la versió la porta git. Cada paquet publicat és una etiqueta (`v6` a `v25g`) i el `CHANGELOG.md` les enumera. Per a veure què va canviar entre dues versions:

```
git diff v24 v25 -- src/chachapoya_DB.bas
git diff v24 v25 --stat
```

## Com reconstruir la base de dades

1. Crea una base de dades Access nova i en blanc.
2. Importa `src/chachapoya_DB.bas` (Alt+F11, Fitxer, Importa fitxer) i executa `BuildDB()`.
3. Importa `src/chachapoya_Form.bas` i executa `BuildForm()`.
4. Per a portar una base amb dades d'una versió a la següent, aplica el pedaç corresponent de `tools/patches/` seguint la seqüència que indica el delta d'aquella versió.

## Sobre aquest historial

Aquest repositori és un **historial reconstruït**. Es va crear el setembre de 2026 a partir de la carpeta de treball del projecte i de la carpeta de descàrregues on havien quedat els fitxers de les primeres setmanes. Les dates dels commits són les dates de modificació dels fitxers originals, no la data en què es van fer els commits. Els commits imiten el que hauria estat treballar amb git des del principi: cada versió sobreescriu els mateixos fitxers, i només els deltes, els pedaços i les notes s'acumulen perquè per naturalesa n'hi ha un per iteració.

Un commit apareix fora d'ordre a propòsit: els últims documents en Word (metodologia v2 i esquema v4) es van tornar a desar el 3 d'agost a les 14:25, després del paquet v6 de les 13:37, però el seu contingut el precedeix.

Anomalies de la carpeta d'origen que convé conéixer:

- **No hi ha cap paquet v7.** La sèrie de mòduls consolidats salta de DB v3 i Form v2 (3 d'agost al matí) al paquet v6.
- **El paquet v6 és una renumeració.** Els fitxers de la carpeta `v6` es diuen v6 però són, segons les seues pròpies capçaleres, DB v4, Form v3, esquema v5 i metodologia v3. El zip conserva els noms originals.
- **Falten dos mòduls antics.** Els esquemes v3 i v4 en Word citen `chachapoya_01_DB.bas` i `chachapoya_01_DB_v2.bas`, mòduls de BD que acompanyaven el formulari consolidat v2 del 24 de juny. No s'han trobat ni a la carpeta de treball ni a descàrregues.
- **El paquet v15 no canvia el codi.** L'esquema i la metodologia passen a v15, però el mòdul `chachapoya_DB_v15.bas` és el codi v14 amb una esmena menor, i el formulari és idèntic al v14. El delta `DELTA_v12_v15.md` és una còpia del `DELTA_v12_v13.md`.
- **Les carpetes `16b` i `17d`** contenen el codi que es declara v16a i v17a respectivament.
- **El fitxer `Prev18-19.md`** és posterior a v19 i el delta v19 a v20 el cita com «la llista A-K»; ací és `docs/notes/revisio_v19_v20.md`.
- **`DELTA_v17_v18_v19.md`** a l'arrel és una còpia idèntica del `DELTA_v17_v18.md` ampliat.
- **`chachapoya_Tancament_v25g_v2.bas`** no és el mòdul de tancament sinó el build script de la base amb la tercera esmena de v25g. Ací és un commit sobre `src/chachapoya_DB.bas`.
- **Els comentaris de capçalera dels mòduls VBA** no sempre s'actualitzen amb la versió: el mòdul DB de v18 a v23 encara diu v16/v17 al primer bloc. La versió real la dona l'etiqueta de git.
- **La numeració interna de la metodologia** («Versió 8 del document, segons la BD v17») és independent de la del paquet i es va aturar deliberadament a la v17; el text ho explica a la nota inicial.

## Llicència

Aquest repositori es publica sota la llicència [Creative Commons Reconeixement-NoComercial 4.0 Internacional (CC BY-NC 4.0)](https://creativecommons.org/licenses/by-nc/4.0/deed.ca). Pots reutilitzar el codi, la documentació i les dades amb finalitats no comercials, sempre que en reconeguis l'autoria. Qualsevol ús comercial requereix el permís de l'autor. Citació suggerida:

> Ribera Torró, E. (2026). *Base de dades arqueològica de Diablo Wasi* [codi i dades]. Treball de Fi de Màster, Universitat d'Alacant. https://github.com/elmoromussa/diablowasi-db

## Què s'ha deixat fora

- Les 60 còpies i backups intermedis de la base Access (327 MB en total). Només s'inclou la base tancada v25g. Les bases de fites anteriors es poden regenerar amb el codi de cada etiqueta.
- Els 26 zips de paquet: el seu contingut és a l'historial.
- Les carpetes i fitxers duplicats byte a byte: incorporats com a revisions del paquet corresponent.
- El repositori git embrionari dins de `scripts_DB11`, sense cap commit.
- Els gràfics de calibratge C14 (`C14_plots.zip`): són resultats d'anàlisi, no part de la base.
- Tres PDF de referència (l'article del congrés UA i un TFM de la UPV amb annexos) per drets d'autor i mida. Se citen a `docs/metodologia.md`.
