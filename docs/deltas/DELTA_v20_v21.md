# DELTA v20 -> v21 (tancat)

Estat: **decisions tancades** (sessio 2026-08-15, revisio de la
worklist v20 amb la BD poblada). Per acord explicit, aquesta versio
NO ha passat per ronda d'esborrany: les decisions es van tancar en
conversa i el delta s'escriu directament en forma tancada com a
registre del paquet. Cap lletra menuda pendent de validar.

Origen dels blocs: les cinc qüestions de la revisio de worklist
d'Esteve (2026-08-15) mes la proposta A (gate de maconeria).

Deliverables d'aquesta versio:

| Fitxer | Paper |
|---|---|
| `chachapoya_PATCH_v21.bas` | Migracio in situ v20 -> v21 (`PatchV21`) |
| `chachapoya_DB_v21.bas` | Build complet + `RebuildQueriesV21` |
| `chachapoya_Form_v21_val.bas` | Formulari v21 (`BuildForm`) |
| `chachapoya_Worklist_v21.bas` | Navegador de worklist (`BuildWorklist`) |
| `DELTA_v20_v21.md` | Aquest document |
| `esquema_bbdd_estructures_v21.md` | Referencia tecnica |
| `manual_us_bbdd_v21.md` | Manual (nus 12 reescrit; pas de maconeria) |
| `tfm_metodologia_bbdd_v21.md` | Metodologia (nota de versio) |

**Sequencia sobre la BD amb dades:**
1. `PatchV21()` (columna nova + derivacio + informe de reverts)
2. `RebuildQueriesV21()` (mai `BuildDB()` sobre BD amb dades)
3. `BuildForm()` v21
4. `BuildWorklist()` v21
5. `QRY_29_V21_Review` (worklist de la migracio)

---

## BLOC 1 - Criteri de la terrassa revisat [TANCAT]

**Reobri el bloc 3 de v20 a peticio explicita del dissenyador**
(l'unica via que el metode permet). La decisio nova:

- **Z queda RESERVAT a la superficie del sistema volat** - les
  lloses recolzades sobre mensules (E) o sobre biga transversal
  (F). El cap de la massa basal d'una terrassa NO es Z i NO obri
  `Sys_Platform`.
- La "segona pregunta" (pla d'us) del criteri v20 desapareix:
  **una terrassa ja respon la pregunta d'us per la seua propia
  tipologia** - una massa construida amb el cap utilitzable es el
  que EA-TER significa. No cal escriure-ho en cap camp mes.
- Sobre un N0 sense N1 queda **una sola pregunta real: el remat**
  (R present/absent/atestat/9, amb el seu format del bloc C v20).
- La clausula v18 C1 ("Z pot sostenir el sistema tot sol") **es
  conserva intacta per al cas que la va motivar**: una superficie
  amb els suports IRRESOLUBLES (E/F/G a 9) continua sent
  plataforma. El que es tanca es el cas E/F/G confirmats a 0.
- **Regla 62** (alerta): `Sys_Platform Like 'Present*'` amb E, F i
  G tots a 0 confirmat. Els 9 NO la disparen (cas v18 C1).
- La branca **'Terrace two questions' de QRY_25 es retira** amb el
  criteri que la motivava.
- Migracio: **cap UPDATE automatic** - una llosa recolzada sobre
  pilers cauria a la xarxa. El patch IMPRIMEIX els codis Z-sol i
  `QRY_29` els llista; el revert esperat es `Sys_Platform ->
  Absent`, Z -> 0 de farciment, material i funcio de plataforma
  buidats, i el remat (si n'hi ha) a R.
- Manual: nus 12 REESCRIT (vegeu manual v21); la prohibicio del
  "coronament per descart", la definicio de G com a progressio i
  les dues cares de R **no es toquen** - eren l'altra meitat del
  bloc 3 v20 i continuen valides.
- Nota de metode: la resta del raonament v20 (H generalitzat com a
  "pla practicable") es va presentar, es va discutir, i el
  dissenyador el va rebutjar amb el criteri de camp per davant:
  Z es la tapa del sistema volat. Registrat perque no es reobri
  sense contraexemple nou.

## BLOC 2 - Masonry_Present: la declaracio de fabrica [TANCAT]

Cas motivador: DW-S01-EA55 (mensula aillada MEN, cap fabrica) -
el contraexemple que demostra que **la classe no deriva la
presencia de fabrica** (MEN conte membres amb fabrica -murets,
pilers- i sense -mensules-). Pel criteri v17 literal:
inaplicabilitat no derivable = judici = camp propi que gateja.

- **Camp nou** `Masonry_Present` BYTE, domini 0/1/9, **default
  NULL** (camp de judici; tres estats epistemics preservats).
  Familia de `Plaster_Present` / `Pigment_Present`, NO `Sys_`
  (reservat a les cinc combinacions d'elements).
- **Nom**: `Masonry_Present` i no `Fabric_Present` - `Fabric` ja
  existeix (divergencia, v17) i la col.lisio seria un problema.
- **Gateja el bloc sencer de 2c**: Masonry_Quality, Stone_Format
  (+Secondary), Stone_Working, Masonry_Type, Mortar_Present
  (+Type, Chinking_Stones, Mortar_Notes) i **Fabric**. La nota
  v17 sobre Fabric continua valida: no es gateja PER COSSOS; el
  seu gate es la declaracio de fabrica mateixa.
- **Plaster i Pigment NO s'encadenen** (decisio): son declaracions
  propies i encadenar-les suposaria que no hi ha revoc sense
  fabrica, cosa no comprovada contra el corpus. Es reobri amb un
  contraexemple, no abans.
- **Entra a la familia de capes** (array g): regla 17 la llista.
  El recompte de camps observacionals no gatejats passa de 28 a
  29.
- **Regla 61** (alerta): detall de maconeria amb valor mentre la
  declaracio es 0 o 9 (espill de R24/R25). NULL es faena pendent
  de la 17, no d'aquesta.
- **Migracio derivable** (patch): qualsevol registre amb algun
  detall de maconeria omplert -> `Masonry_Present = 1` automatic
  (el detall no es podia haver registrat sobre cap fabrica). La
  resta queda NULL i va a `QRY_29` per a la passada conscient.
- **La branca MEN de QRY_25 respecta el gate**: dispara amb la
  declaracio NULL, o amb declaracio 1 i qualitat NULL. Una
  mensula amb `Masonry_Present = 0` te la pestanya resolta i no
  viu a la llista per sempre.
- Exports: `Masonry_Present` entra a QRY_05 i al bloc de
  maconeria de l'export corresponent.
- UI: "Maconeria present" encapcala el bloc de maconeria de
  2.Arq (fila propia); AfterUpdate -> ApplyGating; caption de
  full de dades.

## BLOC 3 - Errata v20: bandes de contorn inabastables [TANCAT]

**Errata de v20, trobada per Esteve amb la BD davant** (captura:
fila ROC-PER amb el desplegable oferint nomes tipus RA).

- Diagnostic: el raonament v20 ("un panell pur no contorneja
  res") **confonia la classe panell amb tota la meitat ROC del
  filtre**. F_ROCKART serveix TOTES les files ROC (ROC, ROC-PER,
  ROC-OVL) de qualsevol registre; en perdre el prefix `RA `, les
  bandes van caure al costat arquitectonic i ROC-PER - exactament
  on viuen - no les podia triar. Simetricament, la meitat
  arquitectonica les oferia on no toquen.
- Correccio (nomes formulari, cap valor emmagatzemat canvia):
  * Combo rupestre: `Name Like 'RA *' OR Name='ND' OR Name Like
    'Outline band*'`.
  * Combo arquitectonic: `Name Not Like 'RA *' AND Name Not Like
    'Outline band*'`.
  * Classe Rock art panel: la llista de tipus es restringeix a
    RA*+ND **al handler per classe** (patro bloc H v20) - la
    intencio correcta del raonament v20, posada on toca.
- La resta del bloc F v20 (tipus sense veredicte, topologia
  derivada dels trams, QRY_27) no es toca.

## BLOC 4 - Errata v20: cataleg de 12.Extra no desplegat [TANCAT]

- Diagnostic: el bloc I v20 va decidir afegir `Socket / negative
  interface` i `Access bench` al cataleg, pero **el cataleg es
  una value list del formulari** i el PATCH (que nomes toca
  taules) no podia desplegar-lo; el Form v20 no es va
  actualitzar. Les dues entrades eren inabastables i el retipat
  de la worklist no es podia fer.
- Correccio: les dues entrades entren a la value list de
  F_ARCH_FEATURES (EN;VAL). Els 4 tipus especulatius es queden.
- La branca de retipat de QRY_25 s'autoneteja: en retipar les 5
  files, la condicio (`Other` + patrons de notes) deixa de
  complir-se.

## BLOC 5 - Etiquetes '(opc.)' als camps sempre-opcionals [TANCAT]

- **Criteri escrit**: porta la marca el camp el NULL del qual es
  un estat final legitim en QUALSEVOL estat del registre. Els
  camps condicionals (Plaster_Color, Mortar_Type...) NO hi
  entren: la seua opcionalitat depen de l'estat i la gestiona el
  gating.
- Conjunt v21: `ID_Support_Secondary` ("Suport secundari
  (opc.)"), `Stone_Format_Secondary` ("Format secundari (opc.)"),
  `Color_Secondary` ("Color sec. (opc.)"). Captions de full de
  dades alineades.
- **Cap color**: coherent amb el bloc K v20 - alli es rebutjava
  el color perque l'obligatorietat es funcio de l'estat; aci la
  marca es estatica i honesta perque l'opcionalitat no ho es.

## BLOC 6 - Residus de worklist, documentats [TANCAT]

- Les branques 'Fabric reveal (window)' (QRY_24) i 'Cornice as
  sill' (QRY_23) disparen sobre el valor 0, i **un 0 confirmat es
  indistingible del 0 de migracio**: la fila no es pot buidar
  quan la resposta honesta es 0. Es el precedent de la friccio
  deliberada de la regla 6.
- Decisio: **cap canvi de consulta**. Les worklists de migracio
  son instruments d'una passada; feta la passada, les files
  supervivents amb 0 confirmat son residu documentat, no faena
  pendent. La passada s'anota aci: **feta el 2026-08-15** sobre
  la BD de 86 estructures; les files que hi queden despres
  d'aquesta data son 0 confirmats.
- El navegador de worklist (v21) porta l'avis del residu al
  MsgBox i a la capcalera del modul.
- Al Pendent (per si el residu arriba a molestar de veritat):
  restringir la branca del brancal als candidats amb la formula
  constant a Systems_Notes, o un marcador de revisio. Cap de les
  dues coses es fa ara.

---

## Regles i consultes (recompte v21)

- **Noves: 61** (detall de maconeria sota declaracio negada),
  **62** (plataforma amb els tres suports confirmats a 0).
- Cap retirada. Recompte: **60 actives, numerades fins a 62**,
  retirades 41 i 50.
- `QRY_16g_Rules_57_60` passa a **`QRY_16g_Rules_57_62`** (sis
  branques, lluny del sostre de 16f).
- **`QRY_29_V21_Review` NOVA** (QRY_28 es del navegador de
  worklist, d'aci el salt de numeracio), dues branques:
  (a) plataformes Z-sol a revertir fila a fila (duplica la 62 a
  proposit: la bateria denuncia, aixo es worklist);
  (b) declaracions de fabrica NULL (duplica la 17 a proposit: es
  el judici que la migracio obri).
- `QRY_25`: la branca de terrasses IX; la branca MEN respecta el
  gate de fabrica.
- `QRY_05` (i l'export germa): entra `Masonry_Present`.
- Recompte de consultes del build: 40 + 1 = **41** (una BD
  migrada des de 17a en porta mes, amb QRY_22 i QRY_28).

## PatchV21 (ordre intern)

1. Preflight: `Upper_Crown_Format` existeix (marcador v20).
2. `ADD COLUMN Masonry_Present BYTE` (default NULL).
3. UPDATE derivable: `Masonry_Present = 1` on hi ha qualsevol
   detall de maconeria; recompte imprés.
4. Informe (SENSE update): codis amb `Sys_Platform Present*` i
   E=F=G=0, per al revert manual via QRY_29.

## Fora d'abast, explicitament

- Cap reobertura de la resta del bloc 3 v20 (dues cares de R,
  G com a progressio, prohibicio del coronament per descart).
- Cap canvi a QRY_23/QRY_24 (bloc 6: residu acceptat).
- Cap encadenament de Plaster/Pigment a la declaracio de fabrica.
