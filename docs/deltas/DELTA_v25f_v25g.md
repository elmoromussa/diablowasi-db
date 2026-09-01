# DELTA v25f -> v25g (tancament de worklist i bateria)

Sessió 2026-09-01. Este delta tanca els 566 NULLs observacionals i les
34 decisions de regla via **cascada de respostes arrel** (qüestionari
P0–P8 + decisions 4.1–4.23), incorpora **DW-S01-EA78 i EA79** (nínxols
nous amb tres arestes `Associated natural context`), i esmena **dues
regles**: R54 (CAV admet suport *Rock dihedral* — el diedre és
cavitat, firmat) i R66 (`Not applicable` deixa de disparar — una
estructura que tanca un espai natural té murs construïts legítims
sense cambra; cas EA16).

Deliverables: `chachapoya_Tancament_v25g.bas` (cascada + cirurgies,
informe/aplicació, només escriu sobre NULL o valors pre-verificats),
`chachapoya_DB_v25g.bas` (R54+R66), este DELTA.

**Seqüència:** importar els dos mòduls (el DB substitueix el v25e) →
`TancamentV25g` (informe: revisar els [DEFAULT]) → `TancamentV25g
True` → `RebuildQueriesV25` → `QRY_16_Validation_Check`.

## Els sis [DEFAULT] triats per l'assistent (revisables a l'informe)

1. **EA07 `Sys_Eave` → 'Present complete'** (podria ser partial): cos
   basal coronat per superfície en voladís sobre biga — doctrina
   firmada: **el mòdul S+T→U és independent de la cambra**; un cos
   basal pot portar ràfec.
2. **EA16 `Sys_Base` → 'Present complete'** (amb B=1, decisió B1).
3. **EA72 `Sys_Platform` → 'Present partial'**: les bigues F
   conservades com a sòl d'un nivell superior (hipòtesis porxo /
   pre-assecat / circulació, a `Systems_Notes`).
4. **JFR → 9 a S01-EA03 i S04-EA18**: amb el segon muntant no
   observable, la fàbrica-fa-de-brancal és indeterminable (coherent
   amb el precedent EA02 i amb la jubilació del camp prevista en v26).
5. **EA77 morter/falques → 9**: deteriorament extrem, no avaluable.
6. **EA78/EA79 perfil 0** («avaluat i absent») als camps de bateria;
   si algun interior no era observable, canvia'l a 9 al formulari.

## Decisions destacades de l'acta

- **Bloc cambra**: B a EA07, EA71, EA72, S04-EA01 (L i murs a 0);
  **EA16** resolt amb `Sys_Chamber='Not applicable'` — el nínxol fa
  d'espai i els murs de tancament es queden (d'ací l'esmena R66).
- **Portals atestats**: EA62, EA74, S02-EA04 → `Attested lost` (cossos
  es queden).
- **EA04**: el recompte basal era l'error (N1 arranca del terreny) →
  0, sense tocar B ni obrir plataforma.
- **S03-EA03**: el pilaret exempt de la boca de la cova no és element
  K → K=0 i alta a `T_ARCH_FEATURES` («Freestanding pillar», amb el
  petit basament descrit a la nota). Primera entrada del catàleg
  d'elements imprevistos — la taula fent exactament la seua faena.
- **EA73** (col·lapsada): R i D → 9. **EA06**: coronament *Projecting
  course*. **v20**: EA09 i EA28 → *amorphous*.
- **Suport modificat**: pendents → 9, i conversió 0→9 a les
  construïdes (el dubte sempre hi és); els 0 es conserven només a
  NIX/CAV. La possible retirada del camp queda **anotada per a v26**,
  no s'obri ara.
- **Plataformes atestades per mènsules** (EA38, EA39, EA40, S04-EA03):
  `Platform_Surface=3` i `Sys_Platform` alçat a *Attested lost* on
  estava Absent — el valor 3 fent el seu ofici exacte (foto d'EA38/39
  a l'acta de sessió).

## Nota conceptual per a v26 / discussió TFM (d'Esteve, cas EA72)

L'esquema N0/N1 com a *base/cambra* mostra el seu límit en estructures
sense cambra tancada però amb dos nivells funcionals (porxo,
circulació): una generalització **base/superestructura** seria més
genèrica i útil. Queda registrada com a qüestió oberta de vocabulari —
no toca esquema ara.

## Estat esperat després de l'aplicació

Bateria QRY_16 a zero o quasi (queden només les files que l'informe
lliste com a `[RESIDU]`, si cap). Corpus: 106 EA, 91 arestes, 12
paternitats. **La BD queda tancada per a l'anàlisi** — següent parada:
pipeline R sobre QRY_13 i QRY_14.
