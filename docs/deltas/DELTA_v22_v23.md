# DELTA v22 -> v23 (tancat)

Estat: **decisions tancades** (sessio 2026-08-15/16, tres tandes:
troballa del brancal treballant la worklist; les huit correccions
d'Esteve; les tres respostes que les tancaven — murs de cambra,
Num. EA emmagatzemat, subsector amb cataleg). A diferencia dels
dos deltes anteriors, aquest SI ha passat per ronda de preguntes
abans de generar codi: el metode torna a ser decisions primer.

Deliverables:

| Fitxer | Paper |
|---|---|
| `chachapoya_PATCH_v23.bas` | Migracio in situ v22 -> v23 (`PatchV23`) |
| `chachapoya_DB_v23.bas` | Build complet + `RebuildQueriesV23` |
| `chachapoya_Form_v23_val.bas` | Formulari v23 (incorpora v22a) |
| `chachapoya_Worklist_v23.bas` | Navegador de worklist |
| `DELTA_v22_v23.md` | Aquest document |
| `esquema_bbdd_estructures_v23.md` | Referencia tecnica |
| `manual_us_bbdd_v23.md` | Manual |
| `tfm_metodologia_bbdd_v23.md` | Metodologia (nota de versio) |

**Sequencia sobre la BD amb dades:**
1. `PatchV23()` (EA_Number + derivacio; L_SUBSECTORS + columna +
   relacio; migracions dels murs de cambra)
2. `RebuildQueriesV23()` (mai `BuildDB()`)
3. `BuildForm()` v23 (amb F_STRUCTURES TANCAT abans)
4. `BuildWorklist()` v23
5. `QRY_16` (Bateria): la regla 66 llista els recomptes de murs a
   revisar fila a fila; la 55 revisada redueix el residu del
   brancal a les 8 files de veritat.

---

## BLOC 1 - La finestra del brancal-en-fabrica guanya l'eix
## del portal [TANCAT]

Troballa d'Esteve treballant la worklist: «Fabrica fa de brancal»
(11.Sist) estava actiu sense mirar mai el sistema portal — el
gate v19 vivia nomes sobre Brancals (0 o 2), i una esqueixada
pressuposa una obertura.

- Radiografia del corpus (84): el camp era actiu sobre 36
  registres amb portal Absent, 16 No aplicable, 6 No observable,
  3 Atestat perdut. La branca de QRY_24 llistava **61 files, de
  les quals nomes 8 tenen portal Present***.
- **El «residu conegut» del bloc 6 del delta v21 estava mal
  diagnosticat i es rectifica**: 53 files eren inaplicables, no
  zeros confirmats. Les 8 bones mereixen una segona passada.
- Finestra revisada: `Jambs IN (0,2)` **AND `Sys_Portal Like
  'Present*'`**. Tres materialitzacions: gate del formulari,
  **regla 55 REVISADA** (cap regla nova; amb portal Absent
  solapa l'exempcio de la regla 3 i totes dues diuen la
  veritat), i branca de QRY_24 restringida.
- **Atestat perdut fora de la finestra** (lletra del criteri
  d'Esteve); es reobri amb contraexemple de camp.
- Cost de dades: zero (els quatre valors declarats viuen sobre
  portals presents).

## BLOC 2 - El codi al titol de la finestra [TANCAT]
## (incorpora l'interi v22a)

- `Form_Current` i `Code_AfterUpdate` escriuen el codi al
  Caption del formulari — la pestanya de document d'Access, que
  cap desplacament ni pestanya interior tapa.
- **Sense sufix** (decisio d'Esteve): la pestanya diu nomes el
  codi («(nou registre)» quan encara no n'hi ha). El Caption
  estatic de base queda en «Registre d'estructura».
- La franja interior amb el codi en negreta (bloc D v22) es
  mante; el fossil «v11» del Caption vell queda documentat al
  v22a.

## BLOC 3 - Planta: entra «Triangular» [TANCAT]

- Value list de `Floor_Plan` al formulari: `Triangular` entre
  Trapezoidal i Irregular. Valor emmagatzemat en angles, com
  sempre. Cap canvi d'esquema (camp TEXT).

## BLOC 4 - Els murs son DE LA CAMBRA [TANCAT]

Reformulacio d'Esteve: la pregunta ja no es *quants murs hi ha?*
sino **quants murs te la cambra?** (facana + retorns; els murets
basals D no s'hi compten mai — la confusio que l'etiqueta vella
convidava a fer).

- **Etiqueta**: «Num. murs (facana + retorn)». Caption de full
  de dades alineada.
- **Cambra absent o no aplicable -> 0 DERIVABLE**: el camp es
  tanca i s'ompli sol (nomes sobre buit; el patch fa la passada
  inicial: ~24 buits -> 0; els 17 zeros existents ja eren
  correctes amb el significat nou).
- **Cambra present o atestada -> camp obert i minim 1** (una
  cambra te almenys la facana): **regla 65**. El corpus compleix
  (30 registres amb cambra, tots 1-3).
- **Cambra no observable o no avaluada -> buit**: no pots
  comptar els murs d'una cambra que no saps si hi es (2 zeros
  del corpus migren a NULL amb el patch).
- **Regla 66**: murs comptats sense cambra — o l'estat de la
  cambra es revisa (Atestada perduda?) o el recompte baixa al 0
  derivable. **Les 11 files del corpus que la disparen SON la
  revisio que el canvi de semantica obri: cap UPDATE cec.**
- Autoompliment del 0 al canvi de `Sys_Chamber` (patro EA-TER:
  nomes sobre camp buit, mai sobre un valor declarat).

## BLOC 5 - Num. EA emmagatzemat [TANCAT — REVERSIO v20]

**Reversio explicita d'una decisio documentada**: l'esquema v20
deia «el numero d'EA no s'emmagatzema mai a banda» (QRY_07
l'extrau per expressio). Esteve tria el camp real: **ordenable i
enllacable amb el flux metric de Metashape**, que es exactament
el cas d'us que la decisio vella no tenia davant.

- Camp `EA_Number` INTEGER, **derivat**: el patch l'ompli per
  als 84 registres; `Code_AfterUpdate` el recalcula a cada
  edicio del codi; el control d'1.Id esta **BLOQUEJAT** (no
  s'edita a ma).
- **Regla 67** respon la preocupacio que fonamentava la decisio
  vella (la divergencia silenciosa): dispara si el numero falta
  o no concorda amb el codi.
- Entra als dos exports d'estructures (QRY_05 i germa), al
  costat del codi. QRY_07 conserva la seua expressio (funciona
  igual; unificar-la es cosmetic i queda fora d'abast).

## BLOC 6 - Subsector amb cataleg [TANCAT]

- **`L_SUBSECTORS`** nova (Name / Name_VAL), vocabulari tancat
  del dissenyador, **transversal als sectors** (posicional, cap
  filtre per sector): Upper «Superior (sup.)», Lower «Inferior
  (inf.)», North «Nord (N)», Central «Central (C)», South «Sud
  (S)».
- `ID_Subsector` LONG a T_STRUCTURES, **NULL legitim sempre**
  («cap valor» es un estat final) -> etiqueta **«Subsector
  (opc.)»** pel criteri del bloc 5 v21. Relacio 26
  (`REL_SSEC_STR`).
- UI a 1.Id: el subsector al costat del sector. **El suport
  secundari baixa a la seccio «Detall suport geologic»**, que es
  sa casa, i deixa lloc al Num. EA i al subsector.
- Fora de tota worklist: un NULL de subsector no es faena
  pendent.

## BLOC 7 - 12.Extra: Bedrock i banqueta/esglao [TANCAT]

- `Material` guanya **`Bedrock`** («Penya (cingle)») — un encaix
  negatiu pot viure a la paret del farallo mateix. Valor
  emmagatzemat coherent amb `Pigment_Substrate = Bedrock`.
- **`Access bench` -> `Access bench / step`** («Banqueta /
  esglao»). **Cost zero verificat**: cap fila retipada encara
  (les 16 continuen com a Other), aixi que el valor vell no viu
  en cap registre.

---

## Regles i consultes (recompte v23)

- **Noves: 65** (cambra present amb recompte 0), **66** (murs
  sense cambra), **67** (Num. EA absent o divergent). La **55 es
  REVISA** (segon eix a la finestra).
- Recompte: **65 actives, numerades fins a 67**, retirades 41 i
  50. `QRY_16g_Rules_57_64` -> **`QRY_16g_Rules_57_67`** (onze
  branques; el precedent de 16f son catorze). Llegats coneguts:
  `_57_60`, `_57_62`, `_57_64`.
- `QRY_24`: branca del brancal restringida a portal Present*
  (61 -> 8 files).
- Exports QRY_05 i germa: entra `E.EA_Number`.
- Relacions: 25 -> **26** (`REL_SSEC_STR`).
- Recompte de consultes del build: **41**, sense canvis.

## PatchV23 (ordre intern)

1. Preflight: `Metrics_Available` (marcador v22).
2. `ADD EA_Number` + UPDATE derivable des del codi (84).
3. `L_SUBSECTORS` + 5 valors + `ADD ID_Subsector` + relacio.
4. Murs de cambra: NULL -> 0 on cambra Absent/NA (~24); 0 ->
   NULL on cambra 9/no avaluada (~2). **Les 11 files amb murs
   sense cambra NO es toquen**: la regla 66 les llista.

## Fora d'abast, explicitament

- QRY_07 conserva l'expressio de l'EA (unificar amb el camp es
  cosmetic; es fara si mai molesta).
- El subsector fora dels exports fins que s'use (el JOIN
  s'afegira aleshores, amb el nom i no l'ID, com sempre).
- Atestat perdut fora de la finestra del brancal (bloc 1).
- Cap canvi a QRY_23 (cornisa-llindar): el seu residu de zeros
  confirmats es el cas genui de friccio deliberada.
