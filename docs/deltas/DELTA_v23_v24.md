# DELTA v23 -> v24 (tancat)

Estat: **decisions tancades** (sessio 2026-08-17, dins del bloc de
georeferenciacio i orientacions del pipeline Metashape; paquet
generat el 2026-08-21, despres que la passada neta de l'extractor
de metriques certificara el calibratge dels bbox de tot el corpus).
Metode: decisions primer, codi despres, com mana el protocol.

Deliverables:

| Fitxer | Paper |
|---|---|
| `chachapoya_PATCH_v24.bas` | Migracio in situ v23 -> v24 (`PatchV24`) |
| `chachapoya_DB_v24.bas` | Build complet + `RebuildQueriesV24` |
| `chachapoya_Form_v24_val.bas` | Formulari v24 (incorpora v23a) |
| `chachapoya_Worklist_v24.bas` | Navegador (SENSE canvis; nomes versio) |
| `chachapoya_Import_Georef_v24.bas` | Importador de coordenades i azimuts (lliurat amb el bloc de georeferenciacio) |
| `DELTA_v23_v24.md` | Aquest document |
| `esquema_bbdd_estructures_v24.md` | Referencia tecnica |
| `manual_us_bbdd_v24.md` | Manual |
| `tfm_metodologia_bbdd_v24.md` | Metodologia (nota de versio) |

**Sequencia sobre la BD amb dades:**
1. `PatchV24()` (afig `Facade_Azimuth_Deg`)
2. `RebuildQueriesV24()` (MAI `BuildDB()` sobre esta BD)
3. `BuildForm()` (amb `F_STRUCTURES` tancat)
4. `BuildWorklist()`
5. `ImportGeoref ruta_csv` amb el CSV **regenerat** de
   `export_georef_orientacions.py`

---

## Bloc unic: l'azimut fotogrametric de facana

### W1. Camp `Facade_Azimuth_Deg` (INTEGER 0-359, NULL)

Azimut del pla exposat, calculat sobre els models de sector
georeferenciats (EPSG:32718) a partir de la normal del bounding
box (`ortho_axes`, eix "-Y", el mateix marc que genera els
orthomosaics frontals), i escrit **exclusivament** per
l'importador de georeferenciacio. Mai a ma.

**NO substitueix `Facade_Orientation`.** Els dos camps son
registres del mateix fet amb fonts diferents i es mantenen
separats a proposit:

| | `Facade_Orientation` | `Facade_Azimuth_Deg` |
|---|---|---|
| Font | observacio de camp (ulls, brúixola) | instrument (bbox sobre model georeferenciat) |
| Resolucio | sector de 8 (45 graus) | grau |
| Escriptura | formulari, entrada manual | importador, mai a ma |
| Paper | dada primaria observacional | dada instrumental d'alta resolucio per a H06/OE3 |

Cada un audita l'altre: la regla 68 es la frontissa.

### W2. Regla 68 (ALERTA, no error)

Divergencia circular entre azimut i sector observat > 67.5 graus
(un sector i mig). Dos registres honests no poden separar-se
tant: o el bbox va rotar despres del calibratge (l'instrument
menteix: regenerar l'exportacio) o l'observacio de camp va errar
(revisitar-la). ND i NULL no la disparen mai.

**Prerequisit operatiu descobert durant el desenvolupament del
pipeline de metriques**: el bbox es l'instrument compartit
d'azimuts, orthomosaics i projeccions de shapes. Girar-lo per a
generar vistes zenitals el descalibra per als tres usos alhora
(cas EA01b, documentat). Regles derivades:
- les vistes d'ortho es trien amb `VIEW_AXIS`, MAI girant el bbox;
- la passada de l'extractor de metriques amb zero `PREFIX_DUBTOS`
  es el **certificat de calibratge** que fa fiable l'azimut;
- l'exportacio de georeferenciacio es regenera DESPRES de la
  passada neta, no abans.

### W3. Fora de les worklists (per disseny)

Precedent `Metrics_Available` (v22): un camp instrumental no esta
mai "pendent d'entrada" - esta pendent d'una importacio. El
formulari el mostra bloquejat a 2.Arq, cel-la (7,2), al costat de
l'orientacio observacional; `Visibility_Valley` baixa a la (9,2).

---

## Pendents que este delta NO obri

- Importacio selectiva de metriques de shapes a la BD
  (`t-int` -> `Interior_Area_m2`, volums amb `ID_Vol_Method`,
  `led` -> `Support_Depth_cm`): delta futur, quan el corpus
  metric estiga consolidat i la conversa R->BD decidisca quines
  columnes mereixen camp.
- Cap canvi de vocabulari A-X ni de dominis.
