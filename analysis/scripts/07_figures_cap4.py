#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
07_figures_cap4.py - Figures gràfiques del capítol 4 de la memòria del TFM
«Construir a l'abisme» (FIG-4.2 a 4.9, 4.11, 4.17, 4.18 i 4.19)

Versió per al paquet d'anàlisi del generador executat el 12/09/2026 per a la
memòria (tfm-diablowasi/figures/fonts/cap4/figures_cap4.py), amb la passada
d'estil del 18/09/2026: el full d'estil comú (mides finals, tipografia,
paleta de tipologies, grisos, noms, format numèric) viu a estil_figures.py
i ací només hi ha les dades i la construcció de cada figura. Cap dada, xifra,
tipus de gràfic ni codificació descrita als peus de la memòria ha canviat.

Entrada:  data/T_STRUCTURES.csv, data/L_TYPOLOGY.csv, data/L_SUPPORT.csv,
          data/T_DECORATIONS.csv, data/T_CONNECTIONS.csv        (pas 00)
          data/matriu_AX_cinc_valors.csv                        (pas 01)
          output/tables/taula_perfil_clusters.csv,
          output/tables/taula_jaccard_elements.csv              (pas 02)
          output/tables/taula_4_1_sector_x_tipologia_106.csv,
          output/tables/riquesa_per_EA.csv                      (pas 05)
          output/tables/metriques_murs_per_EA.csv,
          output/tables/metriques_portals_per_EA.csv            (pas 06)
Sortida:  output/figures/memoria/FIG-4.Y_descripcio.svg i .png (12 + 12 fitxers)

No s'hi generen: FIG-4.1 i 4.10 (cartografia, QGIS), FIG-4.12 a 4.15
(cronologia, OxCal), FIG-4.16 (esquema de motius, dibuix sense dades) ni les
fotografies. Les xifres escrites a mà (grups de la FIG-4.2b, muntants de la
4.5b, «17 al 100 %» de la 4.4b, mitjana i R de la 4.11, recomptes de la 4.18
i de la 4.19) són les publicades a la memòria: estil_figures.literal() les
contrasta amb les dades quan és possible i anota els desajustos al final,
sense modificar-les.

Executar des d'analysis/:  python scripts/07_figures_cap4.py
"""
import sys
import os
import warnings
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
# avís de compatibilitat de versions scipy/NumPy de l'entorn (no afecta el càlcul de la disposició del graf)
warnings.filterwarnings('ignore', message='A NumPy version')
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.patches import PathPatch, Patch
from matplotlib.path import Path
from matplotlib.lines import Line2D
from matplotlib.ticker import MultipleLocator
import matplotlib.patheffects as pe
import estil_figures as E
from estil_figures import (PALETA as COL, TNOM, TNOM_CURT, TNOM_2L, TORD, TYPMAP, GRIS, NEGRE, TEXT, EIX,
                           ELEMENT_NOM, ORDRE_ELEMENTS, LIMITS_BLOCS, ELEMENTS_FUSTA, FAMILIES_DEC, FAMCOL,
                           coma, pct, literal, nota, titol_panell, desa)

print('tipografia:', E.FONT)
pal = E.comprova_paleta()
if pal:
    print('paleta (dE minim CAM02-UCS):', {k: (round(v[0], 1), v[1]) for k, v in pal.items()})
else:
    print('paleta: colorspacious no instal·lat; comprovació omesa')

T = 'output/tables/'

# ---------- entrada: extractes de la BD (pas 00)
_S = pd.read_csv('data/T_STRUCTURES.csv'); _typ = pd.read_csv('data/L_TYPOLOGY.csv'); _sup = pd.read_csv('data/L_SUPPORT.csv')
S = _S[['ID', 'Code', 'ID_Typology', 'ID_Support', 'ID_Support_Secondary', 'Facade_Azimuth_Deg', 'MNI']].copy()
S['Typ'] = S.ID_Typology.map(dict(zip(_typ.ID, _typ.Name)))
S['Sup'] = S.ID_Support.map(dict(zip(_sup.ID, _sup.Name))); S['Sup2'] = S.ID_Support_Secondary.map(dict(zip(_sup.ID, _sup.Name)))
D = pd.read_csv('data/T_DECORATIONS.csv')[['ID_Structure', 'ID_Struct_Body', 'ID_Dec_Type']]
C = pd.read_csv('data/T_CONNECTIONS.csv')[['ID_Struct_A', 'ID_Struct_B']]
S['T'] = S.Typ.map(TYPMAP)

# ======================================================================
# FIG-4.2 - corpus: barres apilades per sector + anell construir/ocupar
# ======================================================================
t106 = pd.read_csv(T + 'taula_4_1_sector_x_tipologia_106.csv', index_col=0)
t106.columns = [TYPMAP.get(c.strip(), c.strip()) for c in t106.columns]; t106.index = [str(i).replace('DW-', '') for i in t106.index]
sec = [s for s in t106.index if s.startswith('S')]
fig, (a1, a2) = E.nova_figura('FIG-4.2', ncols=2, width_ratios=[1.15, 1], wspace=0.02)
left = np.zeros(len(sec))
for t in TORD:
    v = t106.loc[sec, t].values.astype(float)
    a1.barh(sec, v, left=left, color=COL[t], edgecolor='white', linewidth=0.6, label=TNOM[t], height=0.72)
    for i, (x, w) in enumerate(zip(left, v)):
        if w >= 4:
            a1.text(x + w / 2, i, int(w), ha='center', va='center', fontsize=E.MIDA_NOTA, color=E.text_sobre(COL[t]))
    left += v
for i, s in enumerate(sec):
    a1.text(left[i] + 1.2, i, f'{int(left[i])}', va='center', fontsize=E.MIDA_BASE, color=EIX)
a1.invert_yaxis(); a1.set_xlim(0, 80); a1.set_xlabel('elements arqueològics')
a1.xaxis.set_major_locator(MultipleLocator(20))
titol_panell(a1, 'a', 'Tipologies per sector (n = 106)')
a1.legend(loc='center left', bbox_to_anchor=(1.0, 0.5), ncol=1, handlelength=1.0, handleheight=0.9, labelspacing=0.45,
          borderaxespad=0.3)
# (b) anell en la rampa neutra; els quatre grups (54/35/9/8) són les xifres publicades
tot = t106.loc[sec]
grups = [('arquitectura\nconstruïda', literal('FIG-4.2b construïda', 54, int(tot[['MAU', 'CAM', 'TER', 'PLA-V']].values.sum()))),
         ('contextos naturals\nocupats', literal('FIG-4.2b naturals', 35, int(tot[['NIX', 'CAV']].values.sum()))),
         ('no classificable', literal('FIG-4.2b NC', 9, int(tot[['NC']].values.sum()))),
         ('traces i panells', literal('FIG-4.2b traces', 8, int(tot[['MEN', 'PR']].values.sum())))]
n106 = literal('FIG-4.2b total', 106, sum(g[1] for g in grups))
w, _ = a2.pie([g[1] for g in grups], colors=GRIS, startangle=90, counterclock=False,
              wedgeprops=dict(width=0.38, edgecolor='white', linewidth=1.2), radius=1.0)
# etiquetes fora de l'anell, amb línia guia curta; posició (x, y, alineació) per grup
posl = [(1.18, -0.05, 'left'), (-1.18, -0.72, 'right'), (-1.18, 0.78, 'right'), (-0.42, 1.40, 'right')]
for g, wd, (lx, ly, ha) in zip(grups, w, posl):
    ang = np.deg2rad((wd.theta1 + wd.theta2) / 2)
    px, py = 1.04 * np.cos(ang), 1.04 * np.sin(ang)
    a2.plot([px, lx + (-0.04 if ha == 'left' else 0.04)], [py, ly], color=GRIS[1], lw=0.5, zorder=1)
    a2.text(lx, ly, f'{g[0]}\n{g[1]} ({pct(g[1] / n106 * 100)})', ha=ha, va='center', fontsize=E.MIDA_NOTA, linespacing=1.15)
a2.set_xlim(-2.45, 2.35); a2.set_ylim(-1.4, 1.7); a2.set_aspect('equal', adjustable='box', anchor='C')
a2.text(0, 0, f'{n106}\nEA', ha='center', va='center', fontsize=10, fontweight='bold')
titol_panell(a2, 'b', 'Construir o ocupar')
desa(fig, 'FIG-4.2_corpus_sector_tipologia')

# ======================================================================
# FIG-4.3 - suport: (a) tipologia -> primari ; (b) primari -> secundari (92)
# ======================================================================
pop = set(pd.read_csv(T + 'riquesa_per_EA.csv').Code)
S92 = S[S.Code.isin(pop)].copy()
S92['P'] = S92.Sup.map(E.SUPN); S92['Q'] = S92.Sup2.map(E.SUPN).fillna(E.SENSE_SECUNDARI)


def alluvial(ax, left_keys, right_keys, ct, lcol, lname, lab_l, lab_r, gap=2.2, sep_lab=4.6):
    """Diagrama de bandes; l'amplària de cada banda és el nombre d'elements.
    Les etiquetes veïnes s'espaien (sep_lab, en unitats) amb una línia guia."""
    ax.axis('off')
    lcount = ct.sum(1); rcount = ct.sum(0)

    def stack(keys, counts):
        pos = {}; y = 0
        for k in keys:
            pos[k] = (y, y + counts.get(k, 0)); y += counts.get(k, 0) + gap
        return pos, y - gap
    lpos, lh = stack(left_keys, lcount); rpos, rh = stack(right_keys, rcount); H = max(lh, rh)
    loff = (H - lh) / 2; roff = (H - rh) / 2; bw = 0.05
    ax.set_xlim(-0.05, 1 + bw + 0.05); ax.set_ylim(-1, H + 7.5)

    def etiquetes(pos, off, x_node, x_lab, ha, fmt):
        ks = [k for k, (y0, y1) in pos.items() if y1 > y0]
        yc = np.array([(pos[k][0] + pos[k][1]) / 2 + off for k in ks])
        yl = E.espaia_etiquetes(yc, sep_lab, ymin=0, ymax=H)
        for k, y0, y1 in zip(ks, yc, yl):
            ax.text(x_lab, y1, fmt(k, int(pos[k][1] - pos[k][0])), ha=ha, va='center', fontsize=E.MIDA_NOTA)
            if abs(y1 - y0) > 0.3:
                ax.plot([x_node, x_lab + (0.012 if ha == 'right' else -0.012)], [y0, y1], color=GRIS[1], lw=0.4, zorder=1,
                        clip_on=False)
    for k, (y0, y1) in lpos.items():
        if y1 > y0:
            ax.add_patch(plt.Rectangle((0, y0 + loff), bw, y1 - y0, color=lcol(k), lw=0))
    for k, (y0, y1) in rpos.items():
        if y1 > y0:
            ax.add_patch(plt.Rectangle((1, y0 + roff), bw, y1 - y0, color=GRIS[1], lw=0))
    etiquetes(lpos, loff, -0.005, -0.03, 'right', lambda k, n: f'{lname(k)}  {n}')
    etiquetes(rpos, roff, 1 + bw + 0.005, 1 + bw + 0.03, 'left', lambda k, n: f'{n}  {k}')
    lcur = {k: lpos[k][0] + loff for k in lpos}; rcur = {k: rpos[k][0] + roff for k in rpos}
    for l in left_keys:
        for r in right_keys:
            v = ct.loc[l, r] if (l in ct.index and r in ct.columns) else 0
            if not v:
                continue
            y0l = lcur[l]; y1l = y0l + v; y0r = rcur[r]; y1r = y0r + v; lcur[l] = y1l; rcur[r] = y1r
            verts = [(bw, y0l), (0.5, y0l), (0.5, y0r), (1, y0r), (1, y1r), (0.5, y1r), (0.5, y1l), (bw, y1l), (bw, y0l)]
            codes = [Path.MOVETO, Path.CURVE4, Path.CURVE4, Path.CURVE4, Path.LINETO, Path.CURVE4, Path.CURVE4, Path.CURVE4, Path.CLOSEPOLY]
            ax.add_patch(PathPatch(Path(verts, codes), facecolor=lcol(l), alpha=0.55, lw=0))
    # capçaleres de columna: la de l'esquerra acaba al node esquerre, la de la dreta comença al node dret
    ax.text(bw, H + 3.2, lab_l, fontsize=E.MIDA_NOTA, ha='right', va='bottom', color=EIX)
    ax.text(1, H + 3.2, lab_r, fontsize=E.MIDA_NOTA, ha='left', va='bottom', color=EIX)


lord = [t for t in TORD[::-1] if t in set(S92['T'])]
pord = ['repisa ampla', 'repisa estreta', 'microrepisa', "junta d'estratificació", 'terreny', 'esquerda vertical', 'diedre',
        'cavitat gran', 'cavitat mitjana', 'nínxol natural'][::-1]
pord = [p for p in pord if p in set(S92.P)]
qord = ['diedre', 'paret de roca', 'esquerda vertical', "junta d'estratificació", 'terreny', 'microrepisa', 'repisa ampla',
        'cavitat gran', 'cavitat mitjana', E.SENSE_SECUNDARI][::-1]
qord = [q for q in qord if q in set(S92.Q)]
# (b): rampa de grisos per forma del relleu (de dalt a baix del panell), sense colors de tipologia
PGRIS = {p: E.CMAP_SEQ(v) for p, v in zip(pord[::-1], np.linspace(0.25, 0.8, len(pord)))}
fig, (a1, a2) = E.nova_figura('FIG-4.3', ncols=2, wspace=0.06)
alluvial(a1, lord, pord, pd.crosstab(S92['T'], S92.P), lambda k: COL[k], lambda k: TNOM_CURT[k],
         'tipologia', 'suport primari (rep la càrrega)')
titol_panell(a1, 'a', 'Tipologia → suport primari (n = 92)')
alluvial(a2, pord, qord, pd.crosstab(S92.P, S92.Q), lambda k: PGRIS[k], lambda k: k,
         'suport primari', 'suport secundari (estabilitza o allotja)')
titol_panell(a2, 'b', 'Suport primari → secundari (n = 92)')
desa(fig, 'FIG-4.3_suport_geologic')

# ======================================================================
# FIG-4.4 - murs: (a) àrea conservada per tipologia ; (b) índex de conservació
# ======================================================================
W = pd.read_csv(T + 'metriques_murs_per_EA.csv'); W['T'] = W.Tipologia.map(TYPMAP)
fig, (a1, a2) = E.nova_figura('FIG-4.4', ncols=2, width_ratios=[1.6, 1], wspace=0.08)
order = ['CAM', 'MAU', 'TER', 'NIX', 'NC']
a1.set_yticks(range(len(order))); a1.set_yticklabels([TNOM[t] for t in order]); a1.set_ylim(len(order) - 0.5, -0.6); a1.set_xlim(0, 7.2)
a1.set_xlabel('àrea de mur conservada per estructura (m²)'); E.marques_enteres(a1, 'x')
titol_panell(a1, 'a', f'Mur conservat per tipologia (n = {len(W)})')
for i, t in enumerate(order):
    d = W[W['T'] == t].Area_Cons_m2.values
    a1.scatter(d, np.full(len(d), i) + E.eixam(a1, d, 5.6, orient='h'), s=26, color=COL[t], edgecolor='white', linewidth=0.5, zorder=3)
    if len(d) >= 3:
        a1.plot([np.median(d)] * 2, [i - 0.3, i + 0.3], color=NEGRE, lw=1.2, zorder=4)
    a1.annotate(f'n = {len(d)}', (1.0, i), xycoords=('axes fraction', 'data'), xytext=(3, 0), textcoords='offset points',
                ha='left', va='center', fontsize=E.MIDA_NOTA, color=EIX)
# (b) sense cap text damunt ni dins de les barres (comentari del tutor sobre aquest panell)
ic = W.Index_conserv_pct.dropna().sort_values(ascending=False).values
n100 = int((ic >= 99.9).sum())
a2.bar(range(len(ic)), ic, color=[GRIS[0] if v >= 99.9 else '#b5b5b5' for v in ic], width=0.8)
a2.set_ylim(0, 118); a2.set_yticks(range(0, 101, 20)); a2.set_xticks([]); a2.set_xlim(-0.6, len(ic) - 0.4)
a2.set_ylabel("% de l'àrea original conservada"); a2.set_xlabel('estructures,\nordenades per conservació')
a2.spines['bottom'].set_visible(False)
titol_panell(a2, 'b', f'Conservació (n = {len(ic)})')
n100_pub = literal('FIG-4.4b al 100 %', 17, n100)
E.clau_horitzontal(a2, -0.4, n100 - 0.6, 106, 3.5, text=f'{n100_pub} estructures al 100{E.ESPAI_FI}%', dy=2.5)
resta = ic[ic < 99.9]
# només el màxim i el mínim (les etiquetes de totes es tocarien); el màxim, alineat a l'esquerra de la seua barra
# perquè no trepitge la barra fosca veïna, que és més alta
a2.text(n100 - 0.45, resta[0] + 2.5, f'{resta[0]:.0f}', ha='left', va='bottom', fontsize=E.MIDA_MIN, color=TEXT)
a2.text(n100 + len(resta) - 1, resta[-1] + 2.5, f'{resta[-1]:.0f}', ha='center', va='bottom', fontsize=E.MIDA_MIN, color=TEXT)
desa(fig, 'FIG-4.4_murs_conservacio')

# ======================================================================
# FIG-4.5 - portals: (a) llum ; (b) muntants
# ======================================================================
P = pd.read_csv(T + 'metriques_portals_per_EA.csv'); P['T'] = P.Tipologia.map(TYPMAP); Pm = P.dropna(subset=['Amplaria_m'])
fig, (a1, a2) = E.nova_figura('FIG-4.5', ncols=2, width_ratios=[2.4, 1], wspace=0.02)
for t in ['CAM', 'MAU']:
    d = Pm[Pm['T'] == t]
    a1.scatter(d.Amplaria_m, d.Alcaria_m, s=36, color=COL[t], edgecolor='white', linewidth=0.6, zorder=3)
# etiquetes directes al costat d'un punt de cada tipologia, en espai lliure (sense creuar la diagonal)
for t, code, portal, dx, dy, ha, va in [('CAM', 'DW-S01-EA02', '1.1 (N1.1)', 0, 6, 'center', 'bottom'),
                                        ('MAU', 'DW-S01-EA40', '1.1', 6, 0, 'left', 'center')]:
    r = Pm[(Pm.Code == code) & (Pm.Portal == portal)].iloc[0]
    a1.annotate(TNOM[t], (r.Amplaria_m, r.Alcaria_m), xytext=(dx, dy), textcoords='offset points', ha=ha, va=va,
                fontsize=E.MIDA_NOTA, color=TEXT)
a1.plot([0.3, 0.85], [0.3, 0.85], color=GRIS[1], lw=0.8, ls='--', zorder=1)
a1.annotate('obertura quadrada', (0.45, 0.45), xytext=(3, -3), textcoords='offset points', rotation=45, rotation_mode='anchor',
            ha='center', va='top', fontsize=E.MIDA_NOTA, color=EIX)
a1.set_xlim(0.3, 0.85); a1.set_ylim(0.4, 0.75); a1.set_xlabel('amplària (m)'); a1.set_ylabel('alçària (m)')
a1.set_aspect('equal', adjustable='box', anchor='E'); E.eix_coma(a1, 'both')  # l'alçària mana: els dos panells queden alineats
titol_panell(a1, 'a', f'Llum dels portals (n = {len(Pm)})')
# (b) les xifres de muntants (19/9/7 de 35) són les publicades a la memòria
jam = [('composta', 19), ('fàbrica', 9), ('llosa', 7)]
n35 = literal('FIG-4.5b muntants', 35, sum(j[1] for j in jam))
a2.barh([j[0] for j in jam], [j[1] for j in jam], color=GRIS[0], height=0.6)
for i, j in enumerate(jam):
    a2.text(j[1] + 0.5, i, f'{j[1]}  ({pct(j[1] / n35 * 100)})', va='center', fontsize=E.MIDA_BASE)
a2.invert_yaxis(); a2.set_xlim(0, 26); a2.set_xlabel(f'muntants observats (n = {n35})'); E.marques_enteres(a2, 'x')
titol_panell(a2, 'b', 'Com es resol el muntant')
desa(fig, 'FIG-4.5_portals_muntants')

# ======================================================================
# FIG-4.6 - complexitat constructiva per tipologia (n = 92)
# ======================================================================
R = pd.read_csv(T + 'riquesa_per_EA.csv'); R['Tipologia'] = R.Tipologia.map(TYPMAP)
fig, ax = E.nova_figura('FIG-4.6'); order = ['CAM', 'MAU', 'PLA-V', 'TER', 'NC', 'NIX', 'CAV']  # per complexitat decreixent
ax.set_xlim(-0.55, len(order) - 0.45); ax.set_ylim(-0.7, 16.6)
ax.set_xticks(range(len(order))); ax.set_xticklabels([f'{TNOM_2L[t]}\nn = {int((R.Tipologia == t).sum())}' for t in order], fontsize=E.MIDA_MARQUES)
ax.set_ylabel("complexitat constructiva\n(nombre d'elements construïts)"); E.marques_enteres(ax, 'y'); ax.set_yticks(range(0, 16, 5))
ax.tick_params(axis='x', length=0)
for i, t in enumerate(order):
    d = R[R.Tipologia == t].Riquesa.values
    ax.scatter(np.full(len(d), i) + E.eixam(ax, d, 3.6, orient='v'), d, s=9, color=COL[t], edgecolor='white', linewidth=0.3, zorder=3)
    q1, med, q3 = np.percentile(d, [25, 50, 75])
    ax.plot([i - 0.32, i + 0.32], [med, med], color=NEGRE, lw=1.3, zorder=4); ax.plot([i, i], [q1, q3], color=NEGRE, lw=0.7, zorder=2, alpha=0.5)
nota(ax, 0.995, 0.80, 'barra: mediana · línia: rang interquartílic', transform=ax.transAxes, ha='right', va='top', color=EIX)
for x0, x1, lab in [(-0.4, 1.4, 'tombes amb portal'), (1.6, 3.4, 'plataforma volada i terrassa'), (3.6, 6.4, 'contextos naturals i no classificables')]:
    ax.plot([x0, x1], [15.6, 15.6], color=GRIS[1], lw=0.6); nota(ax, (x0 + x1) / 2, 15.8, lab, ha='center', va='bottom', color=EIX)
desa(fig, 'FIG-4.6_complexitat_tipologia')

# ======================================================================
# FIG-4.7 - famílies: matriu de presència per tipologia + perfil de les famílies
# ======================================================================
M = pd.read_csv('data/matriu_AX_cinc_valors.csv'); M = M[M.Code.isin(pop)].set_index('Code')
els = ORDRE_ELEMENTS
B = M[els].isin([1, 2, 3]).astype(int)
tmap = dict(zip(S.Code, S['T'])); B['typ'] = [tmap[c] for c in B.index]; B['rq'] = B[els].sum(1)
order = ['CAM', 'MAU', 'PLA-V', 'TER', 'NC', 'NIX', 'CAV']  # per complexitat decreixent
B['ti'] = B['typ'].map({t: i for i, t in enumerate(order)}); B = B.sort_values(['ti', 'rq'], ascending=[True, False])
fig, (ax, ax2) = E.nova_figura('FIG-4.7', ncols=2, width_ratios=[1.5, 1], wspace=0.04, sharey=True)
n = len(B); mat = B[els].values.T  # elements x estructures
ax.imshow(mat, cmap=E.LinearSegmentedColormap.from_list('pa', ['#ececec', GRIS[0]]), vmin=0, vmax=1, aspect='auto', interpolation='nearest')
ax.set_xticks(np.arange(-0.5, n, 1), minor=True); ax.set_yticks(np.arange(-0.5, len(els), 1), minor=True)
ax.grid(which='minor', color='white', lw=0.35); ax.tick_params(which='minor', length=0)
for b in LIMITS_BLOCS:
    ax.axhline(b - 0.5, color='white', lw=2.0, zorder=3)
strip_y, strip_h = -1.6, 0.8
for i in range(n):
    ax.add_patch(plt.Rectangle((i - 0.5, strip_y), 1, strip_h, color=COL[B['typ'].iloc[i]], lw=0, clip_on=False))
ax.set_yticks(range(len(els))); ax.set_yticklabels([E.etiqueta_element(e) for e in els], fontsize=E.MIDA_NOTA); ax.set_xticks([]); ax.tick_params(length=0)
ax.set_ylim(len(els) - 0.5, -4.6)
for sp in ax.spines.values():
    sp.set_visible(False)
# separadors i etiquetes de grup de tipologia (nom complet o abreujat, en color de text; les curtes, en una fila superior)
bounds = B.groupby('ti').size().cumsum().values; prev = 0
dalt = {'PLA-V', 'NC'}
for k, b in enumerate(bounds):
    t = order[k]; xc = (prev + b) / 2 - 0.5
    ax.axvline(b - 0.5, color='white', lw=1.5, zorder=3)
    if t in dalt:
        ax.text(xc, -3.1, TNOM_2L[t], ha='center', va='bottom', fontsize=E.MIDA_MIN, color=TEXT, linespacing=1.0)
        ax.plot([xc, xc], [-3.0, -1.7], color=GRIS[1], lw=0.4, clip_on=False)
    else:
        ax.text(xc, -1.8, TNOM_2L[t], ha='center', va='bottom', fontsize=E.MIDA_MIN, color=TEXT, linespacing=1.0)
    prev = b
ax.set_title("(a) Elements construïts per element\narqueològic (92 × 20)", loc='left', fontsize=E.MIDA_TITOL, fontweight='bold')
PF = pd.read_csv(T + 'taula_perfil_clusters.csv', index_col=0)[els]
fam_order = [2, 3, 1, 5, 4]; famlab = {2: '2  tombes amb portal (32)', 3: '3  plataforma (5)', 1: '1  basament (22)', 5: '5  murets (2)', 4: '4  no construït (31)'}
PFo = PF.loc[fam_order]
ax2.imshow(PFo.values.T, cmap=E.CMAP_SEQ, vmin=0, vmax=1, aspect='auto')
ax2.set_xticks(np.arange(-0.5, len(fam_order), 1), minor=True); ax2.set_yticks(np.arange(-0.5, len(els), 1), minor=True)
ax2.grid(which='minor', color='white', lw=0.35); ax2.tick_params(which='minor', length=0)
for b in LIMITS_BLOCS:
    ax2.axhline(b - 0.5, color='white', lw=2.0, zorder=3)
ax2.set_xticks(range(len(fam_order))); ax2.set_xticklabels([famlab[f] for f in fam_order], fontsize=E.MIDA_NOTA, rotation=35, ha='right', rotation_mode='anchor')
ax2.tick_params(length=0, labelleft=False)
for i in range(PFo.shape[0]):
    for j in range(PFo.shape[1]):
        v = PFo.values[i, j]
        if v >= 0.05:
            ax2.text(i, j, f'{v * 100:.0f}', ha='center', va='center', fontsize=E.MIDA_MIN, color=E.text_sobre_seq(v))
ax2.set_title('(b) Perfil de les famílies (%)', loc='left', fontsize=E.MIDA_TITOL, fontweight='bold')
for sp in ax2.spines.values():
    sp.set_visible(False)
desa(fig, 'FIG-4.7_families_constructives')

# ======================================================================
# FIG-4.8 - co-ocurrència entre elements (índex de Jaccard)
# ======================================================================
J = pd.read_csv(T + 'taula_jaccard_elements.csv', index_col=0).loc[els, els]
fig, ax = E.nova_figura('FIG-4.8'); n = len(J)
mask = np.triu(np.ones_like(J.values, dtype=bool))
vals = np.where(mask, np.nan, J.values)
im = ax.imshow(vals, cmap=E.CMAP_SEQ, vmin=0, vmax=1)
ax.set_xticks(np.arange(-0.5, n, 1), minor=True); ax.set_yticks(np.arange(-0.5, n, 1), minor=True)
ax.grid(which='minor', color='white', lw=0.35); ax.tick_params(which='minor', length=0)
for b in LIMITS_BLOCS:
    ax.axhline(b - 0.5, color='white', lw=2.0, zorder=3); ax.axvline(b - 0.5, color='white', lw=2.0, zorder=3)
ax.set_xticks(range(n)); ax.set_xticklabels(J.columns); ax.set_yticks(range(n)); ax.set_yticklabels([E.etiqueta_element(e) for e in J.index], fontsize=E.MIDA_NOTA)
ax.tick_params(length=0)
for i in range(n):
    for j in range(i):
        v = J.values[i, j]
        if v >= 0.5:
            ax.text(j, i, coma(v, 2), ha='center', va='center', fontsize=E.MIDA_MIN, color=E.text_sobre_seq(v, 0.65))
for s in ax.spines.values():
    s.set_visible(False)
cb = fig.colorbar(im, ax=ax, shrink=0.45, pad=0.02, format=E.FORMAT_COMA); cb.outline.set_visible(False)
cb.set_label('índex de Jaccard', fontsize=E.MIDA_BASE); cb.ax.tick_params(labelsize=E.MIDA_MARQUES, length=0)
desa(fig, 'FIG-4.8_coocurrencia_jaccard')

# ======================================================================
# FIG-4.9 - xarxa de vincles físics: components i punts d'articulació
# ======================================================================
import networkx as nx
codemap = dict(zip(S.ID, S.Code))
G = nx.Graph()
G.add_nodes_from(sorted(S.Code))  # ordre de nodes fix: disposició reproduïble
G.add_edges_from(sorted((min(a, b), max(a, b)) for a, b in ((codemap[r.ID_Struct_A], codemap[r.ID_Struct_B]) for _, r in C.iterrows())))
art = set(nx.articulation_points(G)); comps = sorted(nx.connected_components(G), key=lambda c: (-len(c), min(c)))
fig, ax = E.nova_figura('FIG-4.9'); ax.axis('off')
pos = {}
big = [c for c in comps if len(c) > 1]; iso = [c for c in comps if len(c) == 1]
layouts = []
for c in big:
    # subgraf construït explícitament amb els nodes ordenats: G.subgraph() itera un conjunt i l'ordre (i la disposició) canviaria d'una execució a l'altra
    sub = nx.Graph(); sub.add_nodes_from(sorted(c)); sub.add_edges_from(sorted(G.subgraph(c).edges()))
    p = nx.kamada_kawai_layout(sub); p = {k: np.array(v) for k, v in p.items()}
    xs = np.array([v[0] for v in p.values()]); ys = np.array([v[1] for v in p.values()])
    scale = np.sqrt(len(c)) * 0.55; w = (xs.max() - xs.min()) * scale + 0.6; h = (ys.max() - ys.min()) * scale + 0.6
    layouts.append((p, scale, w, h, xs.min(), ys.min()))
X = 0; Y = 0; maxh = 0; ample = 0
for i, (p, scale, w, h, xmin, ymin) in enumerate(layouts):
    if i == 3:
        ample = X - 0.4; X = 0; Y -= maxh + 0.5; maxh = 0
    for k, v in p.items():
        pos[k] = np.array([X + (v[0] - xmin) * scale + 0.3, Y - (v[1] - ymin) * scale - 0.3])
    X += w + 0.4; maxh = max(maxh, h)
Y -= maxh + 1.0
for k, x in zip((min(c) for c in iso), np.linspace(0.3, ample - 0.3, len(iso))):
    pos[k] = np.array([x, Y])
nx.draw_networkx_edges(G, pos, ax=ax, edge_color='#bbb', width=0.8)
nodes = list(G.nodes)
nc = [COL[tmap[c]] for c in nodes]; sz = [64 if c in art else 26 for c in nodes]; ec = [NEGRE if c in art else 'white' for c in nodes]
nx.draw_networkx_nodes(G, pos, ax=ax, nodelist=nodes, node_color=nc, node_size=sz, edgecolors=ec, linewidths=[1.2 if c in art else 0.5 for c in nodes])
# codis dels nodes destacats: fora del node, amb halo blanc, en la direcció (de huit) que menys nodes i arestes trepitja;
# en cas d'empat, la que s'allunya del centre del component (tot determinista)
centres = {k: np.mean([pos[m] for m in c], axis=0) for c in big for k in c}
allx = [v[0] for v in pos.values()]; ally = [v[1] for v in pos.values()]
ax.set_xlim(min(allx) - 0.35, max(allx) + 0.35); ax.set_ylim(min(ally) - 0.35, max(ally) + 0.9)
fig.canvas.draw()
PX = {k: ax.transData.transform(v) for k, v in pos.items()}
RX = {k: np.sqrt(64 if k in art else 26) / 2 / 72 * fig.dpi for k in nodes}  # radi de cada node en píxels
fs_px = E.MIDA_MIN / 72 * fig.dpi
mostres = np.concatenate([np.linspace(PX[a], PX[b], 14)[1:-1] for a, b in G.edges])
DIRS = [np.array(v) / np.linalg.norm(v) for v in [(1, 0), (-1, 0), (0, 1), (0, -1), (1, 1), (-1, 1), (1, -1), (-1, -1)]]
for c in ['DW-S01-EA02', 'DW-S01-EA33', 'DW-S01-EA22', 'DW-S01-EA55', 'DW-S04-EA11', 'DW-S01-EA39', 'DW-S04-EA20']:
    text = c.replace('DW-', ''); w = 0.62 * fs_px * len(text); h = 1.1 * fs_px
    r = RX[c] + 2.5 / 72 * fig.dpi
    fora = pos[c] - centres[c]; fora = fora / (np.linalg.norm(fora) + 1e-9)
    millor = None
    for d in DIRS:
        anc = PX[c] + d * r
        x0 = anc[0] if d[0] > 0.3 else (anc[0] - w if d[0] < -0.3 else anc[0] - w / 2)
        y0 = anc[1] if d[1] > 0.3 else (anc[1] - h if d[1] < -0.3 else anc[1] - h / 2)
        caixa = (x0 - 1, y0 - 1, x0 + w + 1, y0 + h + 1)
        dins = lambda p, m=0: caixa[0] - m <= p[0] <= caixa[2] + m and caixa[1] - m <= p[1] <= caixa[3] + m
        cost = sum(dins(PX[k], RX[k] + 1) for k in nodes if k != c) * 10 + sum(dins(p) for p in mostres)
        clau = (cost, -float(d @ fora))
        if millor is None or clau < millor[0]:
            millor = (clau, d)
    d = millor[1]
    ax.annotate(text, pos[c], xytext=(d[0] * r / fig.dpi * 72, d[1] * r / fig.dpi * 72), textcoords='offset points', fontsize=E.MIDA_MIN,
                fontweight='bold', ha='left' if d[0] > 0.3 else ('right' if d[0] < -0.3 else 'center'),
                va='bottom' if d[1] > 0.3 else ('top' if d[1] < -0.3 else 'center'),
                color=TEXT, path_effects=[pe.withStroke(linewidth=2, foreground='white')], zorder=6)
for i, c in enumerate(big[:3]):
    xs = [pos[k][0] for k in c]; ys = [pos[k][1] for k in c]; nota(ax, np.mean(xs), max(ys) + 0.35, f'component de {len(c)}', ha='center', color=EIX)
nota(ax, 0, Y + 0.45, f'{len(iso)} elements aïllats', color=EIX, va='bottom')
ax.legend(handles=[Patch(color=COL[t], label=TNOM[t]) for t in TORD] +
          [Line2D([], [], marker='o', color='w', markerfacecolor='white', markeredgecolor=NEGRE, markeredgewidth=1.2, markersize=7,
                  label="punt d'articulació (contorn negre)")],
          ncol=4, loc='upper center', bbox_to_anchor=(0.5, 0.0), handlelength=1.0, handleheight=0.9, columnspacing=1.2, borderaxespad=0.2)
desa(fig, 'FIG-4.9_xarxa_components')
print('components', [len(c) for c in comps][:9], 'art', len(art))

# ======================================================================
# FIG-4.11 - rosa d'orientacions de façana (n = 34)
# ======================================================================
az = S.Facade_Azimuth_Deg.dropna().values.astype(float); azr = np.deg2rad(az)
mitj = float(np.degrees(np.arctan2(np.mean(np.sin(azr)), np.mean(np.cos(azr)))) % 360)
Rlen = float(np.hypot(np.mean(np.cos(azr)), np.mean(np.sin(azr))))
mitj_pub = literal('FIG-4.11 mitjana', 305, int(round(mitj))); R_pub = literal('FIG-4.11 R', 0.86, round(Rlen, 2))
n34 = literal('FIG-4.11 n', 34, len(az))
fig = plt.figure(figsize=(E.MIDES['FIG-4.11'][0] * E.CM, E.MIDES['FIG-4.11'][1] * E.CM), layout='constrained')
ax = fig.add_subplot(111, projection='polar')
ax.set_theta_zero_location('N'); ax.set_theta_direction(-1)
bins = np.deg2rad(np.arange(0, 361, 20)); h, _ = np.histogram(azr, bins=bins)
# radi proporcional a l'arrel del recompte: l'àrea del sector és proporcional a la freqüència
ax.bar(bins[:-1] + np.deg2rad(10), np.sqrt(h), width=np.deg2rad(20), color=GRIS[0], edgecolor='white', linewidth=0.5, bottom=0)
ax.set_xticks(np.deg2rad([0, 90, 180, 270])); ax.set_xticklabels(['N', 'E', 'S', 'O'])
ax.set_rgrids(np.sqrt([5, 10]), labels=['5', '10'], fontsize=E.MIDA_NOTA, color=EIX); ax.set_rlabel_position(135)
ax.set_rlim(0, np.sqrt(h.max()) + 0.35); ax.grid(color='#ddd', lw=0.5); ax.spines['polar'].set_color('#ddd')
ax.plot([np.deg2rad(mitj)] * 2, [0, np.sqrt(h.max()) + 0.3], color=NEGRE, lw=1.8, zorder=5)
nota(ax, 0.5, -0.11, f'n = {n34} · mitjana {mitj_pub}° · R = {coma(R_pub, 2)}', transform=ax.transAxes, ha='center', va='top')
desa(fig, 'FIG-4.11_rosa_orientacions')

# ======================================================================
# FIG-4.17 - decoració: (a) posició per família ; (b) tipologies decorades
# ======================================================================
D['fam'] = D.ID_Dec_Type.map(E.FAM_TIPUS); D['posn'] = D.ID_Struct_Body.map(E.POSICIO_DEC)
fig, (a1, a2) = E.nova_figura('FIG-4.17', ncols=2, width_ratios=[1, 2.2], wspace=0.06)
porder = ['cos basal', 'pilastres', 'flancs de façana', 'murs de retorn', 'dintell', 'sobre el dintell (fris)', 'roca (superposada)',
          'roca (perímetre)', 'roca (pròxima)', 'panell independent']
ctp = pd.crosstab(D.posn, D.fam).reindex(porder).fillna(0)
left = np.zeros(len(porder))
for f in FAMILIES_DEC:
    v = ctp[f].values; a1.barh(porder, v, left=left, color=FAMCOL[f], edgecolor='white', linewidth=0.5, label=f, height=0.72); left += v
a1.invert_yaxis(); a1.set_xlabel(f'registres decoratius (n = {int(ctp.values.sum())})'); a1.set_xlim(0, 20); a1.xaxis.set_major_locator(MultipleLocator(5))
a1.tick_params(axis='y', length=0)
titol_panell(a1, 'a', 'On es col·loca cada família')
a1.axhline(5.5, color=GRIS[1], lw=0.6, ls=':')
# el filet i les dues notes, a l'altura de files amb barres curtes (dintell, roca superposada) per no trepitjar cap barra
nota(a1, 19.8, 4.0, 'sobre la fàbrica ↑', color=EIX, ha='right', va='center'); nota(a1, 19.8, 6.0, 'sobre la roca ↓', color=EIX, ha='right', va='center')
st = S.set_index('ID'); D['T'] = D.ID_Structure.map(st['T'])
dec = D.groupby(['T', 'fam']).ID_Structure.nunique().unstack().fillna(0)
tot = S['T'].value_counts()
torder = ['MAU', 'CAM', 'TER', 'PLA-V', 'NIX', 'CAV', 'NC']  # construïdes, naturals i, a la fi, no classificables
anyd = D.groupby('T').ID_Structure.nunique()
x = np.arange(len(torder)); wbar = 0.26
for k, f in enumerate(FAMILIES_DEC):
    a2.bar(x + (k - 1) * wbar, [dec.loc[t, f] / tot[t] * 100 if t in dec.index else 0 for t in torder], width=wbar, color=FAMCOL[f], label=f)
a2.set_xticks(x); a2.set_xticklabels([f'{TNOM_2L[t]}\n{int(anyd.get(t, 0))}/{tot[t]}' for t in torder], fontsize=E.MIDA_NOTA, linespacing=1.05)
a2.set_ylabel("% d'estructures de la tipologia"); a2.set_ylim(0, 100); a2.set_xlim(-0.5, len(torder) - 0.5); a2.tick_params(axis='x', length=0)
titol_panell(a2, 'b', 'Quines estructures reben decoració')
fig.legend(handles=[Patch(color=FAMCOL[f], label=f) for f in FAMILIES_DEC], loc='outside lower center', ncol=3, handlelength=1.0,
           handleheight=0.9, columnspacing=1.5)
desa(fig, 'FIG-4.17_decoracio_families')

# ======================================================================
# FIG-4.18 - estat de conservació de cada element construït
# ======================================================================
# (lletra, asterisc, sencer, parcial, perdut amb evidència): recomptes publicats a la memòria
af = [('Z', True, 0, 4, 8), ('E', False, 3, 15, 0), ('J', False, 1, 4, 0), ('D', False, 2, 6, 0), ('R', False, 4, 11, 0),
      ('X', False, 3, 2, 3), ('F', True, 3, 3, 6), ('M', False, 8, 10, 0), ('V', False, 8, 9, 0), ('L', False, 15, 15, 1),
      ('G', False, 1, 1, 0), ('K', False, 5, 5, 0), ('T', False, 5, 4, 0), ('B', False, 31, 22, 0), ('O', False, 10, 7, 0),
      ('I', False, 19, 11, 0), ('C', False, 5, 1, 0), ('N', False, 7, 1, 0), ('Q', True, 16, 0, 12), ('S', False, 2, 0, 0)]
fig, ax = E.nova_figura('FIG-4.18')
labs = [E.etiqueta_element(a[0], ' *' if a[1] else '') for a in af]
n = np.array([a[2] + a[3] + a[4] for a in af]); c = np.array([a[2] for a in af]) / n; p = np.array([a[3] for a in af]) / n; l = np.array([a[4] for a in af]) / n
y = np.arange(len(af))
ax.barh(y, c * 100, color=E.ESTAT['sencer'], label='sencer', height=0.72)
ax.barh(y, p * 100, left=c * 100, color=E.ESTAT['parcial'], label='parcial', height=0.72)
ax.barh(y, l * 100, left=(c + p) * 100, facecolor='white', edgecolor=GRIS[0], linewidth=0.5, hatch=E.TRAMA, label='perdut amb evidència', height=0.72)
for i in range(len(af)):
    nota(ax, 102, i, f'n = {n[i]}', va='center', color=EIX)
ax.set_yticks(y); ax.set_yticklabels(labs, fontsize=E.MIDA_MARQUES); ax.invert_yaxis(); ax.set_xlim(0, 100); ax.set_xlabel('% dels casos construïts')
ax.tick_params(axis='y', length=0); ax.set_ylim(len(af) - 0.5, -0.6)
for t in ax.get_yticklabels():
    if t.get_text().split()[0] in ELEMENTS_FUSTA:
        t.set_fontweight('bold'); t.set_color(TEXT)
ax.legend(handles=[E.patch_estat(s, label=lab) for s, lab in [('sencer', 'sencer'), ('parcial', 'parcial'), ('perdut', 'perdut amb evidència')]],
          ncol=3, loc='upper left', bbox_to_anchor=(0.0, -0.13), handlelength=1.2, handleheight=0.9, columnspacing=1.5, borderaxespad=0)
nota(ax, 1.0, -0.155, 'en negreta, elements de fusta', transform=ax.transAxes, ha='right', va='top', color=EIX)
desa(fig, 'FIG-4.18_conservacio_elements')

# ======================================================================
# FIG-4.19 - nombre mínim d'individus per estructura
# ======================================================================
EXCAVADES = ['DW-S01-EA17', 'DW-S01-EA01']  # les dues cambres funeràries excavades el 2023
m = S[S.MNI.notna()].copy(); m['exc'] = m.Code.isin(EXCAVADES)
m = m.sort_values(['MNI', 'Code'], ascending=[False, True]).reset_index(drop=True)
fig, ax = E.nova_figura('FIG-4.19')
ax.bar(range(len(m)), m.MNI, color=[COL[t] for t in m['T']], edgecolor=[NEGRE if e else 'white' for e in m.exc],
       linewidth=[1.2 if e else 0.5 for e in m.exc], width=0.78, zorder=3)
ax.set_xticks(range(len(m))); ax.set_xticklabels([c.replace('DW-', '') for c in m.Code], rotation=90, fontsize=E.MIDA_MIN)
ax.tick_params(axis='x', length=0); ax.set_xlim(-0.6, len(m) - 0.4)
ax.set_ylabel("nombre mínim d'individus (MNI)"); ax.set_ylim(0, 62); ax.set_yticks(range(0, 61, 20))
for i, r in m.iterrows():
    if r.MNI >= 5:
        nota(ax, i, r.MNI + 1.2, int(r.MNI), ha='center', va='bottom')
# clau i anotació directa de les dues excavades (les dues primeres barres)
iex = [i for i, e in enumerate(m.exc) if e]
E.clau_horitzontal(ax, min(iex) - 0.4, max(iex) + 0.4, 53, 2.5)
nota(ax, min(iex) - 0.4, 55.5, 'excavades el 2023', ha='left', va='bottom')
nota(ax, 0.985, 0.93, '44 EA sense restes observades · 37 EA amb interior no observable',  # xifres publicades (no deriven de S)
     transform=ax.transAxes, ha='right', va='top', color=EIX)
presents = [t for t in TORD if t in set(m['T'])]
ax.legend(handles=[Patch(color=COL[t], label=TNOM[t]) for t in presents], loc='upper left', bbox_to_anchor=(1.0, 1.0), ncol=1,
          handlelength=1.0, handleheight=0.9, labelspacing=0.45, borderaxespad=0.3)
desa(fig, 'FIG-4.19_mni_estructures')

# ---------- comprovacions finals
faltes = E.comprova_textos()
print('textos vetats als SVG:', 'cap' if not faltes else faltes)
if E.DESAJUSTOS:
    print('xifres publicades que no coincideixen amb el càlcul (es conserva la publicada):')
    for d in E.DESAJUSTOS:
        print('  -', d)
print('done')
