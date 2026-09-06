# DELTA v25g → v26 — la versió definitiva pre-depòsit

*Sessió de decisions: 2026-09-06. Totes les decisions tancades en conversa i firmades abans de generar cap línia de codi. Corpus verificat abans de proposar res: 106 EA, 91 arestes, 7 datacions, 40 evidències de pèrdua, 84 decoracions, 19 elements arquitectònics, 183 files de revisió mètrica.*

---

## 1. Les cinc decisions congelades

### (a) Jubilació de `Jamb_Fabric_Reveal` — amb certificat empíric

La taula creuada `Jamb_Fabric` × `Jamb_Fabric_Reveal` sobre els 22 registres amb parella informada va donar **concordança 22/22, zero contradiccions**, amb la regla de derivació:

1. Parella conté *Fabric* → derivat **1** (la presència de *Fabric* domina, fins i tot amb *NObs* a l'altre membre)
2. Si no, parella conté *NObs* o és *ND* → derivat **9**
3. Resta (i parella buida) → derivat **0**

Els 84 registres amb parella buida portaven tots el 0 de farciment, coherent. El cas EA02 (DW-S01) — que en el seu moment va contradir parella i qualificador — havia demostrat la **patologia de dos-llocs-per-a-un-fet**; la derivabilitat completa n'és el certificat de jubilació. Matís semàntic que la verificació va fer aflorar: els 4 casos amb `Jambs=0` i parella *Fabric-\** (S01-EA06, EA15, EA22; S04-EA14) són exactament el «brancal per revelació del parament» — la parella ho captura completament, la jubilació no perd res.

**Execució**: la regla 3 (branca portal) i la R55 es reformulen sobre l'expressió derivada inline (`Jamb_Fabric Like '*Fabric*'` i la seua negació amb guarda de NULL); la branca JFR de QRY_24 es retira (la branca de completesa de la parella — `Jamb_Fabric Is Null` amb portal Present\* — ja vigila el pendent); el camp desapareix del rebuild i no viatja en la migració. **Micro-decisió (a-i)**: DW-S02-EA02 conserva la parella `ND` amb `Jambs=0` — *ND* és una afirmació d'indeterminació observada, no una absència de registre.

**Criteri metodològic que en queda** (a la metodologia v26): un qualificador viu mentre diu una cosa que cap altre camp diu; quan esdevé derivable, es jubila i la derivació es documenta. La classe de camp sobreviu amb un membre (`Sill_Coincides_Cornice`).

### (b) Reformulació base/superestructura (Guengerich 2014) — documental, no estructural

Perfil dels tres casos confirmats, verificat sobre la BD:

| EA | Cossos basals | Cambra | Aparell superior |
|---|---|---|---|
| DW-S01-EA07 | 2 | 0 | ràfec complet, plataforma parcial, R=2, G=2 |
| DW-S01-EA08 | 1 | 0 | coronament R=1 |
| DW-S01-EA72 | 1 | 0 | plataforma i ràfec parcials |

Redefinició: **N0 = base** (massa constructiva de suport, sense espai interior accessible); **N1 = superestructura** (allò que la base sosté — típicament una cambra, també un aparell obert de coronament o circulació). La doctrina «ràfec sobre cos basal» passa d'excepció a **teorema del marc**. Cap camp es toca: al corpus no existeix cap *cos* de superestructura comptable que no siga cambra (la superestructura d'EA07/08/72 s'expressa en elements U/R/Z, no en cossos) — sense cas confirmat, cap camp nou.

**Pol invers i simetria**: EA-PLA-V (EA20) és superestructura **sense base** — el penya-segat fa de base com el nínxol fa de cambra a EA16. Que la roca puga substituir qualsevol dels dos termes de la parella de Guengerich és l'argument del cingle com a participant estructural, i que l'arquitectura funerària de cingle compartisca la gramàtica base/superestructura de la residencial alimenta la hipòtesi de seqüència constructiva compartida.

### (c) G posicional — la família del voladís

G queda component del sistema plataforma (E+F+G→H). Tres raons: (1) el vocabulari ja reparteix la mateixa tècnica en tres posicions amb tres noms — **G plataforma / I juntura / R-*Projecting course* coronament** — i fer G transversal trencaria la lectura sistèmica (regla 62 inclosa); (2) el precedent de D va eixamplar morfologia i alliberar alçada, mai donar segona funció sistèmica; (3) *un gest tècnic, tres feines* és un argument de chaîne opératoire visible només si les posicions es mantenen separades.

Verificació dels dos registres vius, amb imatge: **EA01** — filades de la plataforma (semblen cornisa intranivell, treballen de suport del tauler: el test G/I guanya la pota de la FUNCIÓ, cas fotogràfic al manual); **EA07** — les filades rematen la plataforma i es recolzen sobre la biga transversal F: **la seqüència compositiva F→G→H llegida físicament a la fàbrica**. Cap dada es toca. *Sub-cornice* (Guengerich Fig. 3) continua sense cas; destí `T_ARCH_FEATURES` si mai n'apareix un, fins als seus tres casos.

### (d) `Support_Modified` — manteniment amb estatut documental

Doctrina v25g verificada neta sobre dades: els 8 zeros són tots NIX; els 6 positius (EA15, EA31, S04-EA12/13/14, S04-EA20) es reparteixen en cinc tipologies. El 9 dominant (92/106) és la doctrina de tres estats funcionant, no el camp fracassant. **Estatut**: camp documental, no analític — exclòs de la matriu de co-ocurrència i del clustering; els positius es discuteixen qualitativament com a traça directa del treball de la roca (OE4; a EA20, els encaixos de mènsula: el penya-segat retallat per a rebre l'estructura — contra una lectura plana del determinisme geològic).

### (e) Menors: escombratge final

Verificat sobre la BD: tipologia 10 residual — cap; sistemes NULL fora d'EA11 — cap; notes de treball a T_STRUCTURES — cap. Queda una família: **cinc placeholders v11** a `T_LOST_ELEMENTS` («Proposta automatica... Completar element.»), redactats amb textos derivats estrictament del tipus d'evidència (vegeu §2.F). Les dues files descriptives de DW-S02-EA01 es queden.

---

## 2. Acta de dades (l'única escriptura de la versió)

**Comptabilitat exacta: 24 valors en 15 files de `T_STRUCTURES` + 5 notes a `T_LOST_ELEMENTS` = 29 escriptures en 20 files.** (Correcció sobre l'estimació de sessió «21 valors»: el desglossament d'EA11 puja a 7 valors.) Cap alta, cap baixa: els recomptes de taula queden intactes (106/91/7/40).

| Bloc | Files | Escriptura | Fonament |
|---|---|---|---|
| A | 11 EA amb `Sys_Chamber=Absent` (S01: EA37, EA60, EA61, EA71, EA76–79; S05-EA02; S06-EA01, S06-EA04) | `N_Chamber_Bodies` NULL→0 | Autoomplert derivable, mateixa doctrina que `N_Built_Walls` (absent → 0) |
| B | DW-S01-EA29 | `N_Chamber_Bodies` 1→0, `N_Basal_Bodies` 0→1 | El cos existent és una base de nivellament: comptador a la columna bona |
| C | DW-S04-EA20 | `N_Chamber_Bodies` 1→0, `Sys_Base` Absent→**Not applicable** | La plataforma no és cos de cambra; una PLA-V no pot tindre base per definició tipològica — NA és constricció del tipus, no decisió constructiva (conseqüència analítica volguda: les consultes de «sense base» l'exclouen) |
| D | DW-S01-EA35 | comptadors 0→NULL | Pigment perimetral: el 0 assereix «avaluat i zero» on res no és observable |
| E | DW-S01-EA11 | `ID_Typology` MAU→**Unclassifiable** (resolt per nom), 4 sistemes NULL→*Not observable*, comptadors 0→NULL | La mateixa `L_TYPOLOGY` cita EA11 com a exemple del tipus 9 — la dada havia regressat; els sistemes passen de pendent a judici emès |
| F | 5 files de pèrdua (EA09, EA21, EA20 tipus 6; EA10 tipus 7; EA12 tipus 8) | Notes redactades | Textos derivats del tipus d'evidència, en ASCII (diacrítics polibles després via formulari) |

Textos del bloc F: tipus 6 → «Mensules romanents al buit; l'element sostingut (tauler de plataforma) perdut.» (EA20, variant: «Tauler de lloses parcialment perdut; les mensules romanen al buit.»); tipus 7 → «Pigment sobre penya nua; la fabrica que el portava, desapareguda.»; tipus 8 → «Murs truncats en pla net; filades o cossos superiors perduts.»

---

## 3. Paquet i flux d'aplicació

| Fitxer | Paper |
|---|---|
| `chachapoya_PATCH_v26.bas` | **Pas 1.** Sobre una CÒPIA de DB_v25h: les 29 escriptures, amb tota precondició verificada abans d'escriure res (una fallada = avortament amb informe i zero escriptures; transacció única). No toca esquema. |
| `chachapoya_DB_v26.bas` | **Pas 2.** `BuildDB()` sobre base EN BLANC: esquema v26 complet — sense JFR, regla 3 i R55 sobre l'expressió derivada, QRY_24 sense la branca JFR, comentaris doctrinals (reformulació, estatut, retirada) al mateix codi. |
| `chachapoya_Migrate_v26.bas` | **Pas 3.** Dins la v26 nova: selector de fitxer en temps d'execució (patró `ImportGeoref`), verificació que les taules de dades destí són buides, alineació de TOTS els lookups per (ID, columna clau real de cada taula) contra la còpia apanyada **amb fase de sincronització** (§5.3), importació en ordre de dependències amb llistes de camps construïdes des de l'esquema destí (la columna jubilada no viatja, els ID viatgen explícits), recomptes contra origen i cinc *spot checks* de les decisions v26 a l'arribada. |
| `chachapoya_Form_v26_val.bas` | **Pas 4.** `BuildForm()`: el control «Fàbrica fa de brancal» retirat amb el camp; la graella 20,2 queda lliure. |
| `esquema_bbdd_estructures_v26.md` | Referència tècnica: preàmbul v26, fila JFR marcada jubilada amb la derivació canònica, files de comptadors reformulades, estatut de `Support_Modified`, regla 55 reformulada, test G/I amb la pota de funció. |
| `manual_us_bbdd_v26.md` | Manual d'entrada + **Addenda v26**: les tres doctrines i la retirada, en clau operativa. |
| `tfm_metodologia_bbdd_v26.md` | Justificació: el patró de qualificadors guanya el seu criteri de jubilació; la reformulació N0/N1 amb el rendiment per a la hipòtesi de seqüència compartida; nota d'estatut a H02. |

**Verificació final esperada** (després del pas 4): bateria `QRY_16_Validation_Check` a **zero**; recomptes 106/91/7/40; els cinc spot checks del migrador en verd.

## 4. El que v26 deixa deliberadament fora

La lletra **Y** (banqueta d'accés) continua esperant el tercer cas (2 de 3). *Sub-cornice* continua esperant el primer. La descomposició per subnivells (`T_BODIES`) continua ajornada amb el comptador posat. Cap d'aquestes és feina de la versió definitiva: són línies del capítol de futur.

---

## 5. Acta d'estrena del rebuild (mateixa sessió, cadena executada d'extrem a extrem)

La cadena PATCH → BuildDB → Migrate → BuildForm → bateria es va executar sencera el mateix dia de les decisions, i **la primera execució real del camí de rebuild net va fer exactament la faena per a la qual el pla el preveia: traure els bugs latents abans del depòsit.** Tots resolts dins del mateix paquet (fitxers republicats), documentats perquè la història del delta incloga la depuració:

**5.1. Quatre serrells latents del codi.** (i) *L'apòstrof al literal JET*: les notes del bloc F amb «l'element» trencaven la cadena SQL — funció `Esc()` (cometes simples doblades) al PATCH. (ii) *L'ordre de `T_METRIC_REVIEW`* (error 3371): el bloc amb l'única FK inline del script vivia ABANS de la creació de `T_STRUCTURES`; mai no havia disparat perquè v25 va entrar sempre per patch sobre una base que ja tenia la taula principal — bloc reubicat després de `T_STRUCTURES`, amb comentari de correcció in situ. (iii) *El comptador de consultes*: la constant d'esperades s'havia quedat en 41 mentre v25 afegia QRY_30/30a/31/32 — apujada a 45, i el llindar d'alarma de 35 a 45. (iv) *El MsgBox fòssil*: el quadre final anunciava «DATABASE v23... 144 fields... 40 queries» amb contingut congelat en l'era v14, mentre el log deia la veritat — capçalera actualitzada als números reals (v26, 22 taules, 25 relacions, 45 consultes, 119 camps).

**5.2. El migrador guanya tres fases.** Selector de fitxer en temps d'execució amb guardes (cancel·lació neta; tria accidental de la base actual detectada); fase 0 «esquema construït?» que converteix el críptic 3078 en una instrucció llegible; i la comprovació de taules buides moguda ABANS de tocar cap lookup.

**5.3. La doctrina de lookups — el guardià que va salvar el corpus.** L'alineació va detectar `L_DEC_TYPE`: destí 17, origen 17, **casades 12**. No era deriva de noms: eren els FORATS d'autonumber de la història viva (13, 14 i 16 desapareguts en altes i baixes de la sèrie de patches). Un build fresc insereix seqüencial 1–17; les 84 files de `T_DECORATIONS` apunten als ID vius (15, 17–20) — sense el guardià, cada decoració d'eixe tram s'hauria **remapat silenciosament a un tipus equivocat**. Doctrina resultant: *els lookups vius són DADES amb cicatrius; el script els bootstrapa per a bases noves (transferibilitat: Aoujgal), però en migració la FONT és l'autoritat i els forats d'ID mai no es renumeren.* Implementació: fase de sincronització — el lookup desalineat es repobla des de la font amb ID explícits, es reverifica i queda reportat a l'informe («id holes preserved»).

**5.4. R59 després del zero derivable: la bateria fent el seu ofici.** Amb els 11 comptadors NULL→0, R59 va poder comparar per primera vegada i va destapar dues incoherències que el NULL emmascarava: **DW-S05-EA02** amb `Recessed_Frame=9` sota portal *Absent* (un «no examinable» sense subjecte → 0 derivable) i **DW-S01-EA77** amb `Access_Plane`/`Portal_Orientation`/`Portal_Position` en ND sota portal *Absent* (el ND-en-bloc de la passada de deteriorament reclamava judicis sobre preguntes sense subjecte → NULL). **Esmena post-migració firmada: 4 valors en 2 files.** Comptabilitat total de la transició v25g→v26: **33 valors en 22 files** (29 del PATCH + 4 de l'esmena R59).

**5.5. Certificat de tancament.** Sobre la BD v26 lliurada: bateria `QRY_16_Validation_Check` a **zero**; verificació externa independent (pipeline mdbtools) de 21 comprovacions en verd — recomptes complets (106/91/7/40/84/19/183), `Jamb_Fabric_Reveal` fora de l'esquema amb la parella al seu lloc, les 29 escriptures del PATCH a l'arribada, l'esmena R59 amb el predicat reimplementat a zero sobre les 106 files, els forats de `L_DEC_TYPE` preservats i les 84 decoracions resolent el seu tipus. **Residu documentat, no de v26**: una fila de decoració totalment buida (ID 58, DW-S01-EA01, present idèntica a la font — probablement un clic perdut al subformulari), sense informació falsa i fora de l'abast de tota regla; neteja opcional d'una línia anotada.

---

*Següent parada: quadern interpretatiu (Entrada 3, «El penya-segat com a graf», entrada de cronologia de la fibra del morter) i pipeline R.*
