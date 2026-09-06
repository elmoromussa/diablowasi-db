# Context per a les sessions de treball

Base de dades arqueològica en Microsoft Access per a un Treball de Fi de Màster. Autor únic: Esteve Ribera-Torró. El repositori és públic i acadèmic; tot el que s'hi afig ha de ser presentable a un tribunal.

## Historial

- L'historial anterior al setembre de 2026 es va reconstruir a partir de la carpeta de treball, amb les dates originals dels fitxers. **No s'ha de reescriure**: ni rebase, ni amend, ni moure etiquetes. Tot canvi va en un commit nou a sobre de `main`.
- Les etiquetes d'iteració (`v6` a `v26`) són anotades i marquen els paquets de treball; les versions numerades (`1.0.0` i següents) marquen les publicacions. `CHANGELOG.md` les enumera i s'actualitza a mà.
- El repositori presenta el resultat del treball, no el procés de reconstrucció. El `README.md` i el `CHANGELOG.md` descriuen les versions pel que aporten, no per com estaven organitzats els fitxers d'origen.

## Convencions

- **Fitxers sense versió al nom.** `src/chachapoya_DB.bas`, `src/chachapoya_Form.bas`, `docs/esquema.md`, `docs/metodologia.md` i `docs/manual.md` es sobreescriuen a cada versió; la versió la dona l'etiqueta de git. Els deltes, els pedaços i les notes sí que porten versió perquè n'hi ha un per iteració.
- **Els deltes són la font dels canvis.** Cada iteració naix d'un `docs/deltas/DELTA_vN_vM.md`, que és l'especificació; els commits i el CHANGELOG el resumeixen. Per a entendre un canvi, llig primer el delta.
- **Idiomes.** Codi, noms de taula i camp, i valors emmagatzemats en anglés. Interfície del formulari, comentaris del codi, documentació, missatges de commit i converses en valencià.
- **Codificació.** Els `.bas` són ASCII pur amb finals de línia CRLF: els accents s'ometen a propòsit als comentaris i les cadenes perquè l'editor VBA d'Access no gestiona bé UTF-8. La documentació Markdown és UTF-8.
- **Noms antics.** La documentació cita els fitxers amb el nom que tenien al paquet (`chachapoya_DB_v25.bas`, `esquema_bbdd_estructures_v25.md`). No s'han de corregir: la taula de correspondència és al `README.md`.
- **Els documents històrics no s'editen.** `docs/`, `data/` i els mòduls de `tools/` són tal com es van lliurar. Les correccions es fan al README o en fitxers nous.

## Treballar amb el VBA sense Access

No hi ha Access a l'entorn de treball i el codi no es pot executar ni compilar. El que sí que es pot fer:

- Revisar per lectura, amb atenció a `Option Explicit`, a les declaracions `Dim` i a les cadenes SQL construïdes per concatenació.
- Comparar versions amb `git diff vN vM -- src/chachapoya_DB.bas` per a veure exactament què va canviar.
- Comprovar la coherència entre el codi i `docs/esquema.md`: noms de taula, camps, regles numerades (R1 a R69) i consultes (`QRY_NN`).
- No proposar canvis al codi que no es puguen verificar amb la lectura; si cal executar alguna cosa, deixar-ho indicat perquè l'autor ho prove a Access.

Els scripts de `tools/metashape/` depenen de l'API de Python d'Agisoft Metashape i tampoc no es poden executar ací: s'apliquen els mateixos criteris de revisió per lectura.

La base `db/chachapoya_v26.accdb` és un binari: no es pot inspeccionar des d'ací. Les dades es consulten a través dels CSV de `data/` o dels informes del codi.
