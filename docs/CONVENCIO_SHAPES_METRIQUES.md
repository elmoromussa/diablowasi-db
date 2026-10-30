# Convenció de registre mètric per shapes 3D
### Sistema de documentació dimensional sobre models fotogramètrics — PALP / La Petaca i Diablo Wasi

**Versió 1.1 — 20 d'agost de 2026** (v1.0 congelada el 17/08; revisió: elements plat/slot/sill, lletres = instàncies, fragments per repetició d'etiqueta, propagació entre murs `+PLA-COS`; sufix de reconstrucció fixat en `-r`)
Validada empíricament en set iteracions de prova sobre quatre estructures de DW-S01 (EA02, EA04, EA12, EA15) i calibrada contra els seus propis casos de fallada.

---

## 1. Principis

1. **Es dibuixa l'observació, es deriva la resta.** Les shapes registren geometria observada (o inferida amb estatus declarat); tota mètrica composta (sumes, ràtios, rectangles, prismes, índexs) es deriva per script, mai es dibuixa.
2. **Suport de digitalització: el model 3D.** Les shapes es dibuixen sobre el model en vista Model. La digitalització sobre orthomosaics queda descartada per a geometria de mesura: Metashape hi emmagatzema els vèrtexs en 2D i la informació perpendicular al pla de l'ortho es perd de manera irrecuperable (verificat en tres configuracions diferents). Els orthomosaics queden com a producte gràfic.
3. **Els vèrtexs 3D són l'observació; la projecció és anàlisi.** L'extractor projecta cada shape sobre el pla que li pertoca en temps de càlcul. Canviar el criteri de projecció no exigeix redibuixar res.
4. **Cap valor sense procedència.** Tot valor derivat porta declarat l'origen (observació directa, restitució, propagació) al CSV.
5. **Els controls de qualitat són consellers, no jutges.** Un avís assenyala una anomalia geomètrica; la decisió final és de l'investigador, que la documenta.
## 2. Gramàtica de les etiquetes

```
PLA[-MÈTRICA]-[ELEMENT]-COS[+COS...][a-z][-COS][-res|-c]
```

| Component | Valors | Notes |
|---|---|---|
| PLA | `f b t bo r l` | front, back, top, bottom, right, left (= eixos del bounding box) |
| MÈTRICA | `ll hh dd` | només polilínies |
| ELEMENT | `port int corn crown led bra wall und cav roof eave plat slot sill jamb lint tie pil bbeam tbeam ebeam ext` | opcional |
| COS | `0.1 1.1 1.2 …` | = nomenclatura de L_STRUCT_BODY |
| `+[PLA-]COS…` | llista d'aplicabilitat | només polilínies amb estatus declarat; sense pla = mateix mur, amb pla = un altre mur (ex. `f-hh-wall-1.1+r-1.1+l-1.1-c`) |
| `a-z` | lletra d'**instància** | entitats distintes: 2n portal, 3a mènsula, 2n mur coplanari (expedient dimensional propi). Els **fragments** d'una mateixa cara es dibuixen com a shapes separades amb **la mateixa etiqueta** (s'agreguen) |
| `-COS` (segon) | parell d'interfície | només `corn`, `led` o pla nu (ex. `t-corn-0.1-1.1`) |
| `-r` / `-c` | estatus | reconstruït / conservat complet (`-res` acceptat com a sinònim, normalitzat a `-r`) |

Exemples canònics: `f-1.1` · `f-1.1a`+`f-1.1b` · `f-port-1.1-c` · `t-corn-0.1-1.1` · `t-int` · `f-ll-wall-0.1+1.1+1.2-c` · `r-dd-wall-1.1-c` · `f-hh-wall-1.2-r`.

### 2.1 Harmonització amb el vocabulari A–X i la base de dades

| Element | Correspondència |
|---|---|
| `port` | P — Access_Opening |
| `int` | alimenta `Interior_Area_m2`: **espai d'ús funerari de tota tipologia**, construït o natural |
| `corn` | I — Interlevel_Cornice |
| `crown` | R — Upper_Crown |
| `roof` | X — Chamber_Roof: `t-roof` superfície exterior, `bo-roof` intradós (sostre interior) |
| `wall` + prefix | `f` façana · `r`/`l` = V Lateral_Walls · `b` = W Rear_Wall |
| `bra` | E — Timber_Brackets (dialoga amb Timber_Bracket_Count) |
| `led` | alimenta `Support_Depth_cm` (suport geològic) |
| `cav` | cavitat natural: `f-cav` boca, `t-cav` planta del contenidor geològic (≠ `t-int` quan no coincideixen) |
| `und` | conservat no atribuïble a cap cos (propi de la convenció) |
| `eave` | **reservat** per a T/U (ràfec) — no reutilitzar `corn` |
| `plat` | H — Corbelled_Platform (superfície de plataforma: `t-plat`) |
| `slot` | encaix negatiu de mènsules: dialoga amb E i amb `Support_Modified` (`f-slot`, `f-hh-slot`) |
| `sill` | N — Sill (llindar: `f-sill`, `f-hh-sill`) |
| `jamb` | O — Jambs (brancals; instàncies per lletra: `f-jamb-1.1a`/`b`) |
| `tie` | D — Tie_Walls (muret transversal: cares a `r-`/`l-`; **atribuït al cos sobre el qual arranca**; la seua `hh` documenta si abasta més nivells). Frontera element/cos: si forma part de la seqüència constructiva (sosté el cos següent) és un **cos** de L_STRUCT_BODY, no un `tie` |
| `pil` | K — Structural_Pilasters (pilastra o pilaret de suport, **adossat a fàbrica o a roca** — la funció mana, el suport no; `hh` alçada, `ll` amplada, `dd` ressalt). Discriminant amb `tie`: membre **puntual** vertical (amplada ≈ gruix ≪ alçada, funció de sostindre) → `pil`; pany **planar** (llargària ≫ gruix, funció de travar) → `tie` |
| `lint` | Q — Lintel (dintell, siga llosa o biga encastada: la codificació és **funcional**, no material; `f-ll-lint` sobrepassant el portal → generositat d'encastament) |
| *peça bifuncional* | una biga que fa de dintell **i** es prolonga en voladís tipus mènsula es registra com a **dues funcions**: `lint` (tram sobre l'obertura, fins a la cara del suport; porta el `dd` de la peça) + `bra` (braç en voladís, instància pròpia, **sense** `dd` duplicat). El sistema registra funcions, no peces; una peça compartida entre N+O+Q→P i E+F+G→H documenta l'**acoblament dels dos sistemes compositius** (H01) |
| `bbeam` | A — Embedded_Base_Beams (jàcenes basals: `f-ll-bbeam`, `f-dd-bbeam` = diàmetre) |
| `tbeam` | F — Transverse_Beams (bigues transversals del sistema E+F+G→H) |
| `ebeam` | S — Eave_Beam (biga de suport del ràfec, sistema S+T→U) |
| `ext` | contorn exterior en planta (`t-ext`, empremta): amb `t-int` deriva el gruix mitjà de mur |

Formes esborrany normalitzades automàticament: `cov`→`cav`, `clo`→`crown`, `t-roofin`→`bo-roof`, `t-roofout`→`t-roof`, `sup`→`plat`, `strat`→`slot`, `llind`→`sill`, `jac`→`bbeam`, `rooftop`→`roof`, `dint`→`lint`, `bran`→`jamb`, `pilar`/`pilastre`→`pil`, `b-roof`/`b-eave`→`bo-`; les etiquetes que comencen per mètrica o element reben `f-` per defecte com a proposta. **`men` no es normalitza automàticament** (ambigu entre mènsula→`bra` i muret transversal→`tie`): decisió manual cas a cas.

Les etiquetes es validen per expressió regular amb normalització (insensible a majúscules; les formes llegades — `e-`, `p-`, `n0.1`, `01`, sufix davant de la llista `+` — es detecten i es proposa el canònic; `-res` es normalitza a `-r`).

## 3. Doble semàntica del prefix

- **Polígons**: el prefix declara el **pla de projecció** («des de quina de les sis vistes del bbox es veu de cara»). Es verifica per la normal del polígon (QC d'obliqüitat).
- **Polilínies**: el prefix declara la **pertinença** («de quin mur és la mesura»), amb independència d'on o des d'on es va traçar. Es verifica per proximitat del punt mitjà als polígons dels murs (QC de pertinença).
La vista de dibuix no declara res: els vèrtexs 3D no la recorden i el sistema no la necessita.

**El valor d'una polilínia no depén mai de cap pla de projecció**: `ll` i `dd` són longituds 3D absolutes i `hh` es descompon contra la vertical del món (propietat física, no del bbox). El prefix, en les línies, només **adscriu** la mesura al seu expedient (mur dret, mur frontal, interior...); els QC vigilen eixa adscripció amb la geometria, però mai alteren el valor. En els polígons, en canvi, el pla sí que computa — una àrea necessita pla per a ser mètrica comparable; una longitud, no.

## 4. Naturalesa de cada token mètric

| Token | Dimensió | Càlcul | Trams |
|---|---|---|---|
| `ll` | recorregut al llarg d'un element | longitud 3D total | multi-tram |
| `hh` | alçada | component vertical | multi-tram |
| `dd` | espan entre dues superfícies (gruix, profunditat) | **corda 3D absoluta** | **un sol tram** |

El `dd` no es projecta: la perpendicularitat és responsabilitat del traçat, vigilada pel QC de direcció, que compara amb la **normal real** del polígon del mur més pròxim (immune a murs atalussats).

## 5. Tres estatus epistemològics

| Sufix | Significat | Paral·lel al domini A–X |
|---|---|---|
| *(cap)* | conservat, integritat no jutjada | valor 2 (present parcial) |
| `-c` | conservat **i complet**: val com a original | valor 1 (present complet) |
| `-r` | original **inferit** (reconstrucció) | valor 3 (attested lost) |

- Cada element es dibuixa **una sola vegada**; el sufix declara l'estatus.
- La **restitució per polilínies** és la via ordinària: dimensions inferides ancorades en arrancaments, empremtes o forats de biga (`f-ll-wall-1.1-res` × `f-hh-wall-1.1-res` = rectangle; × `dd` = prisma). El polígon `-r` és la via excepcional, quan hi ha superfície de fons clicable rere el buit.
- El model geomètric simple (rectangle/prisma) **declara** el nivell d'inferència; quan un volum restituït entre a la base de dades, `ID_Vol_Method` en documenta l'origen.
## 6. Jerarquia d'evidència i propagació

Per a cada dimensió de cada (mur, cos):

```
-c directa  >  -r directa  >  propagada (llista +)
observació     inferència local   inferència per continuïtat
```

- La llista `+` propaga una dimensió entre **cossos del mateix mur** (ítem sense pla) o **cap a altres murs** (ítem amb pla: `+r-1.1`); el sufix descriu només l'àncora; als destins el valor arriba sempre com a inferència per continuïtat (font `prop:PLA-COS-sufix`).
- Conflicte `-r`∧`-c` de la mateixa dimensió → avís; mana l'observació.
- Mesura directa i propagada coexistents amb discrepància > 2 % → avís (o la continuïtat declarada no era tal, o una mesura té un problema).
- La propagació **entre murs** es va ajornar en la v1.0 i s'ha incorporat en esta revisió (sintaxi `+PLA-COS`) després que el corpus la demanara de manera recurrent.
## 7. Mètriques conservades vs. de disseny

- Tota mètrica dibuixada és de l'**estat conservat** i s'etiqueta `_cons`.
- Les ràtios de **disseny** (p. ex. portal/mur) només es calculen amb denominador d'estatus `-c` o `-r` (`_orig`); la ràtio canònica de portal és **portal / mur frontal del cos propi**.
- L'**índex de conservació** = àrea conservada / àrea original, per (mur, cos).
## 8. Controls de qualitat

| QC | Vigila | Llindar |
|---|---|---|
| Obliqüitat / prefix dubtós | normal del polígon vs pla declarat | > 35° i millor ajust en un altre pla |
| Pertinença | punt mitjà de la línia vs polígons dels murs | > 1,5× la distància al mur més pròxim + 10 cm |
| Direcció | corda de la línia vs el que dicten prefix, token i **família de l'element** (els projectants `bra`/`tbeam`/`bbeam` corren per la normal; el seu `dd` és un diàmetre, perpendicular a l'eix) | > 30° (∥) / < 60° (⊥) |
| Conflicte `-r`∧`-c` | doble estatus d'una mateixa dimensió | — |
| Discrepància directa/propagada | valors coexistents | > 2 % |
| Orfandat | línia d'element de mur amb prefix no-mur | — |

A més: exclusió (amb avís) de shapes 2D llegades d'orthos, i validació d'etiquetes amb proposta de canònic.

## 9. Flux de treball

1. Dibuixar sobre el model 3D (vista Model, *Draw Polygon* / *Draw Polyline*), etiquetant segons la gramàtica. Per a vèrtexs sobre buits amb fons: vista ortogràfica alineada a la façana.
2. Executar `export_metriques_shapes.py` (Tools → Run Script): validació prèvia per chunk amb els QC en pantalla.
3. Revisar avisos; corregir etiquetes o traçats si cal; re-executar.
4. Exportar: quatre CSV (`estructures`, `murs`, `portals`, `shapes_audit`) — el darrer és l'acta d'auditoria completa, shape a shape.
5. Anàlisi en R sobre `metriques_murs.csv` i `metriques_portals.csv`; integració selectiva a la base de dades (camps `Interior_Area_m2`, volums amb `ID_Vol_Method`, `Support_Depth_cm`) en el delta corresponent.
## 10. Marc de referència i autocalibratge

Les shapes 3D es guarden en el CRS de shapes del chunk (geogràfic en chunks georeferenciats; marc local d'orientació interna i escala mètrica en chunks locals). L'extractor detecta el cas per la magnitud de les coordenades i, en el cas cartesià, **autocalibra** el marc provant els dos candidats (món-local / intern) i triant el d'obliqüitat mediana mínima. El bounding box és l'instrument de projecció: la seua alineació amb el pla exposat és prerequisit compartit amb la generació d'orthomosaics i amb el càlcul d'azimuts de façana.
