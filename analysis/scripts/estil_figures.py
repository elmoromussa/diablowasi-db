#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
estil_figures.py - Full d'estil comú de les figures gràfiques del capítol 4
de la memòria del TFM «Construir a l'abisme» (importat per 07_figures_cap4.py).

Principi rector: el color és de la tipologia; tota la resta es dibuixa en
grisos, amb trama o amb un únic accent negre. Un color, un significat, en
totes les figures. Cap color, mida de lletra ni etiqueta de tipologia o
d'element viu dins de l'script 07: tot és ací.

Continguts:
  - mides finals d'impressió (cm) de cada figura, sense bbox_inches='tight'
  - tipografia i rcParams (SVG amb text editable, hashsalt fix)
  - paleta de tipologies (única paleta cromàtica) i comprova_paleta()
  - rampa neutra, trama i mapa seqüencial de grisos
  - diccionaris de noms: tipologies, elements del vocabulari (TAULA-3.1),
    famílies decoratives, formes del relleu
  - format numèric amb coma decimal i espai fi abans de «%»
  - eixam determinista (sense atzar) per als diagrames de punts
  - desa(): SVG sense data a les metadades + PNG a 600 ppp
  - literal(): xifres publicades a la memòria, contrastades amb les dades
"""
import io
import os
import re
import matplotlib
matplotlib.use('Agg')
import numpy as np
import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.colors import LinearSegmentedColormap
from matplotlib.patches import Patch
from matplotlib.ticker import FuncFormatter, MaxNLocator

# ---------------------------------------------------------------- mides (cm)
CM = 1 / 2.54
MIDES = {  # amplària x alçària en cm; totes 16,0 llevat de la rosa (9,0)
    'FIG-4.2': (16.0, 7.8), 'FIG-4.3': (16.0, 8.3), 'FIG-4.4': (16.0, 7.4),
    'FIG-4.5': (16.0, 7.9), 'FIG-4.6': (16.0, 7.3), 'FIG-4.7': (16.0, 15.8),
    'FIG-4.8': (16.0, 12.8), 'FIG-4.9': (16.0, 12.7), 'FIG-4.11': (9.0, 10.0),
    'FIG-4.17': (16.0, 7.1), 'FIG-4.18': (16.0, 9.7), 'FIG-4.19': (16.0, 6.0),
}
PPP_PNG = 600

# ---------------------------------------------------------------- tipografia
for _f in ['/usr/share/texmf/fonts/opentype/public/tex-gyre/texgyreheros-regular.otf',
           '/usr/share/texmf/fonts/opentype/public/tex-gyre/texgyreheros-bold.otf',
           '/usr/share/texmf/fonts/opentype/public/tex-gyre/texgyreheros-italic.otf']:
    try:
        font_manager.fontManager.addfont(_f)
    except Exception:
        pass
_installed = {f.name for f in font_manager.fontManager.ttflist}
FONT = next((f for f in ['TeX Gyre Heros', 'Arial', 'DejaVu Sans'] if f in _installed), 'sans-serif')

MIDA_BASE, MIDA_MARQUES, MIDA_NOTA, MIDA_MIN, MIDA_TITOL = 8, 7.5, 7, 6.5, 8.5
TEXT, EIX, NEGRE = '#222', '#555', '#111'
GRIS = ['#3d3d3d', '#8a8a8a', '#c4c4c4', '#e6e6e6']   # rampa neutra de quatre passos
ESTAT = {'sencer': '#3d3d3d', 'parcial': '#b5b5b5', 'perdut': 'white'}  # perdut: trama
TRAMA = '////'

plt.rcParams.update({
    'font.family': FONT, 'font.size': MIDA_BASE,
    'axes.titlesize': MIDA_TITOL, 'axes.titleweight': 'bold', 'axes.titlelocation': 'left',
    'axes.titlepad': 6, 'axes.labelsize': MIDA_BASE,
    'xtick.labelsize': MIDA_MARQUES, 'ytick.labelsize': MIDA_MARQUES,
    'legend.fontsize': MIDA_NOTA, 'legend.frameon': False, 'legend.handlelength': 1.0,
    'axes.spines.top': False, 'axes.spines.right': False,
    'axes.edgecolor': EIX, 'axes.linewidth': 0.6, 'xtick.color': EIX, 'ytick.color': EIX,
    'xtick.major.width': 0.6, 'ytick.major.width': 0.6, 'xtick.major.size': 2.5, 'ytick.major.size': 2.5,
    'axes.labelcolor': TEXT, 'text.color': TEXT,
    'hatch.color': GRIS[0], 'hatch.linewidth': 0.5,
    'svg.fonttype': 'none', 'svg.hashsalt': 'diablowasi',
    'figure.dpi': 100, 'savefig.facecolor': 'white',
})

# ---------------------------------------------------------------- tipologies
TORD = ['MAU', 'CAM', 'TER', 'PLA-V', 'NC', 'NIX', 'CAV', 'MEN', 'PR']
PALETA = {'MAU': '#2a6f97', 'CAM': '#c8553d', 'TER': '#2bb3b2', 'PLA-V': '#cd6bdf',
          'NIX': '#e6a13a', 'CAV': '#bab096', 'NC': '#dcdcdc', 'MEN': '#454545', 'PR': '#f5d97b'}
TNOM = {'MAU': 'mausoleu', 'CAM': 'cambra funerària', 'TER': 'terrassa de repisa',
        'PLA-V': 'plataforma volada', 'NIX': 'nínxol natural', 'CAV': 'cavitat',
        'NC': 'element no classificable', 'MEN': "traça d'element estructural",
        'PR': "panell d'art rupestre"}
TNOM_CURT = dict(TNOM, TER='terrassa', NC='no classificable', MEN='traça estructural', PR='art rupestre')
TNOM_2L = dict(TNOM, CAM='cambra\nfunerària', TER='terrassa\nde repisa', **{'PLA-V': 'plataforma\nvolada'},
               NIX='nínxol\nnatural', NC='no\nclassificable')
TYPMAP = {'EA-MAU Mausoleum/Chullpa': 'MAU', 'EA-CAM Funerary Chamber': 'CAM', 'EA-TER Ledge Terrace': 'TER',
          'EA-PLA-V Aerial Platform': 'PLA-V', 'NIX Natural Niche': 'NIX', 'CAV Cave/Cavern': 'CAV',
          'MEN Isolated Structural Element': 'MEN', 'PR Rock Art': 'PR', 'Unclassifiable': 'NC'}


def comprova_paleta(minim=12.0):
    """Distància mínima CAM02-UCS entre parells de la paleta en visió normal,
    deuteranopia, protanopia i tritanopia (colorspacious, dependència de
    desenvolupament). Exigeix >= `minim`. Retorna {condició: (ΔE, parell)}."""
    try:
        from colorspacious import cspace_convert
    except ImportError:
        return None
    ks = list(PALETA)
    X = np.array([[int(PALETA[k][i:i + 2], 16) / 255 for i in (1, 3, 5)] for k in ks])
    res = {}
    for nom, cvd in [('normal', None), ('deuteranopia', 'deuteranomaly'),
                     ('protanopia', 'protanomaly'), ('tritanopia', 'tritanomaly')]:
        Y = X if cvd is None else np.clip(
            cspace_convert(X, {'name': 'sRGB1+CVD', 'cvd_type': cvd, 'severity': 100}, 'sRGB1'), 0, 1)
        U = cspace_convert(Y, 'sRGB1', 'CAM02-UCS')
        best = (np.inf, None)
        for i in range(len(ks)):
            for j in range(i + 1, len(ks)):
                d = float(np.linalg.norm(U[i] - U[j]))
                if d < best[0]:
                    best = (d, (ks[i], ks[j]))
        res[nom] = best
    pitjor = min(v[0] for v in res.values())
    if pitjor < minim:
        raise ValueError(f'paleta de tipologies: dE minim {pitjor:.1f} < {minim}: {res}')
    return res


def luminancia(hexcol):
    r, g, b = [int(hexcol[i:i + 2], 16) / 255 for i in (1, 3, 5)]
    lin = [c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4 for c in (r, g, b)]
    return 0.2126 * lin[0] + 0.7152 * lin[1] + 0.0722 * lin[2]


def text_sobre(hexcol):
    """Color de text llegible sobre un fons: blanc sobre fosc, TEXT sobre clar."""
    return 'white' if luminancia(hexcol) < 0.3 else TEXT


# ---------------------------------------------------------------- grisos
CMAP_SEQ = LinearSegmentedColormap.from_list('greys_trunc', plt.get_cmap('Greys')(np.linspace(0.05, 0.85, 256)))


def text_sobre_seq(v, llindar=0.62):
    """Text de cel·la sobre CMAP_SEQ: blanc a partir de `llindar` (0-1)."""
    return 'white' if v >= llindar else TEXT


# ---------------------------------------------------------------- elements (TAULA-3.1)
ELEMENT_NOM = {'B': 'nivell de basament', 'C': 'sòcol decoratiu', 'D': 'murets transversals de trava',
               'E': 'mènsules de fusta', 'F': 'bigues transversals', 'G': 'filades en voladís',
               'Z': 'superfície de plataforma', 'I': 'cornisa entre nivells', 'J': 'cantoneres',
               'K': 'pilastres estructurals', 'L': 'flancs de façana', 'M': 'fris en baix relleu',
               'V': 'murs de retorn', 'R': 'coronament', 'X': 'coberta de cambra', 'N': 'llindar',
               'O': 'brancals', 'Q': 'dintell', 'S': 'biga de suport del ràfec', 'T': 'superfície de ràfec'}
# Ordre del vocabulari, de baix a dalt de l'edifici, en sis blocs (decisió d'Esteve, 18/09/2026)
BLOCS_ELEMENTS = [('base', ['B', 'C', 'D']), ('plataforma volada', ['E', 'F', 'G', 'Z']),
                  ('cornisa entre nivells', ['I']), ('cos de la cambra funerària', ['J', 'K', 'L', 'M', 'V', 'R', 'X']),
                  ('portal', ['N', 'O', 'Q']), ('ràfec', ['S', 'T'])]
ORDRE_ELEMENTS = [e for _, els in BLOCS_ELEMENTS for e in els]
# índexs (entre files) on comença cada bloc llevat del primer: per a dibuixar els filets
LIMITS_BLOCS = list(np.cumsum([len(els) for _, els in BLOCS_ELEMENTS])[:-1])
ELEMENTS_FUSTA = {'E', 'F', 'S'}


def etiqueta_element(lletra, sufix=''):
    return f'{lletra}  {ELEMENT_NOM[lletra]}{sufix}'


# ---------------------------------------------------------------- famílies decoratives i suport
FAMILIES_DEC = ['baix relleu', 'pintura', 'art rupestre']
FAMCOL = {'baix relleu': GRIS[0], 'pintura': GRIS[1], 'art rupestre': GRIS[2]}
FAM_TIPUS = {1: 'baix relleu', 2: 'baix relleu', 3: 'baix relleu', 4: 'baix relleu', 5: 'baix relleu',
             6: 'baix relleu', 8: 'baix relleu', 7: 'pintura', 9: 'pintura', 18: 'pintura', 19: 'pintura',
             11: 'art rupestre', 12: 'art rupestre', 15: 'art rupestre', 17: 'art rupestre',
             20: 'art rupestre', 10: 'ND'}
POSICIO_DEC = {1: 'cos basal', 4: 'pilastres', 5: 'flancs de façana', 6: 'murs de retorn', 10: 'dintell',
               11: 'sobre el dintell (fris)', 14: 'roca (superposada)', 15: 'roca (perímetre)',
               16: 'roca (pròxima)', 17: 'panell independent'}
SUPN = {'Narrow natural ledge (<2m)': 'repisa estreta', 'Micro-ledge (<50cm)': 'microrepisa',
        'Medium cavity (1-10m2)': 'cavitat mitjana', 'Large cavity (>10m2)': 'cavitat gran',
        'Natural niche (<1m2)': 'nínxol natural', 'Vertical cleft': 'esquerda vertical',
        'Bedding-plane recess': "junta d'estratificació", 'Ground': 'terreny', 'Rock dihedral': 'diedre',
        'Wide natural ledge (>2m)': 'repisa ampla', 'Rock surface': 'paret de roca'}
SENSE_SECUNDARI = 'cap (una sola forma)'

# ---------------------------------------------------------------- nombres
ESPAI_FI = ' '


def coma(x, dec=None):
    """Nombre amb coma decimal, sense dependre del locale."""
    if dec is None:
        s = f'{x:g}'
    else:
        s = f'{x:.{dec}f}'
    return s.replace('.', ',')


def pct(x, dec=0):
    """Percentatge amb coma decimal i espai fi abans de «%»."""
    return f'{coma(x, dec)}{ESPAI_FI}%'


FORMAT_COMA = FuncFormatter(lambda x, _: coma(x))


def eix_coma(ax, quin='both', dec=None):
    f = FuncFormatter(lambda x, _: coma(x, dec))
    if quin in ('x', 'both'):
        ax.xaxis.set_major_formatter(f)
    if quin in ('y', 'both'):
        ax.yaxis.set_major_formatter(f)


def marques_enteres(ax, quin='y'):
    """Marques enteres en un eix de recompte."""
    if quin in ('x', 'both'):
        ax.xaxis.set_major_locator(MaxNLocator(integer=True))
    if quin in ('y', 'both'):
        ax.yaxis.set_major_locator(MaxNLocator(integer=True))


# ---------------------------------------------------------------- figures
def nova_figura(nom, ncols=1, nrows=1, width_ratios=None, height_ratios=None, wspace=None, hspace=None,
                w_pad=0.04, h_pad=0.04, **kw):
    """Figura a mida final exacta (cm de MIDES) amb layout='constrained'."""
    w, h = MIDES[nom]
    gs = {}
    if width_ratios:
        gs['width_ratios'] = width_ratios
    if height_ratios:
        gs['height_ratios'] = height_ratios
    fig, axs = plt.subplots(nrows, ncols, figsize=(w * CM, h * CM), layout='constrained', gridspec_kw=gs, **kw)
    eng = fig.get_layout_engine()
    opts = dict(w_pad=w_pad, h_pad=h_pad)
    if wspace is not None:
        opts['wspace'] = wspace
    if hspace is not None:
        opts['hspace'] = hspace
    eng.set(**opts)
    return fig, axs


def titol_panell(ax, lletra, text):
    ax.set_title(f'({lletra}) {text}', loc='left', fontsize=MIDA_TITOL, fontweight='bold')


def nota(ax, x, y, text, **kw):
    """Anotació menuda (7 pt) en color de text."""
    kw.setdefault('fontsize', MIDA_NOTA)
    kw.setdefault('color', TEXT)
    return ax.text(x, y, text, **kw)


def clau_horitzontal(ax, x0, x1, y, alt, text=None, color=TEXT, lw=0.6, fontsize=MIDA_NOTA, dy=None, **kw):
    """Clau fina (línia amb dos extrems cap avall) que abraça de x0 a x1 a l'altura y;
    `alt` és la llargària dels extrems en unitats de dades; text centrat damunt."""
    ax.plot([x0, x0, x1, x1], [y - alt, y, y, y - alt], color=color, lw=lw, solid_capstyle='butt',
            clip_on=False, zorder=5)
    if text:
        ax.text((x0 + x1) / 2, y + (dy if dy is not None else alt * 0.6), text, ha='center', va='bottom',
                fontsize=fontsize, color=color, **kw)


def eixam(ax, valors, d_pt, orient='v'):
    """Eixam determinista i simètric: desplaçaments (en unitats de dades de l'eix
    transversal) perquè cap punt de diàmetre `d_pt` punts trepitge un altre.
    Els punts es col·loquen per ordre de valor; els que xocarien s'aparten
    alternativament a un costat i a l'altre (0, +d, -d, +2d, -2d ...).
    Cal que els límits de l'eix estiguen fixats abans de cridar-la."""
    fig = ax.figure
    fig.canvas.draw()  # executa el layout perquè transData siga el definitiu
    p0 = ax.transData.transform((0, 0))
    sx = abs((ax.transData.transform((1, 0)) - p0)[0])
    sy = abs((ax.transData.transform((0, 1)) - p0)[1])
    d = d_pt / 72 * fig.dpi
    s_val, s_off = (sy, sx) if orient == 'v' else (sx, sy)
    valors = np.asarray(valors, float)
    ordre = np.argsort(valors, kind='stable')
    col = np.zeros(len(valors))
    placed = []
    for i in ordre:
        u = valors[i] * s_val
        prop = [(pu, po) for pu, po in placed if abs(pu - u) < d]
        k = 0
        while True:
            cand = 0.0 if k == 0 else ((k + 1) // 2) * d * (1 if k % 2 else -1)
            if all((pu - u) ** 2 + (po - cand) ** 2 >= (d * 0.999) ** 2 for pu, po in prop):
                break
            k += 1
        placed.append((u, cand))
        col[i] = cand / s_off
    return col


def espaia_etiquetes(ys, sep, ymin=None, ymax=None):
    """Desplaça etiquetes veïnes perquè cap parell quede a menys de `sep`;
    conserva l'ordre i centra el paquet quan cal."""
    ys = np.asarray(ys, float)
    ordre = np.argsort(ys, kind='stable')
    y = ys[ordre].copy()
    for _ in range(50):
        mogut = False
        for i in range(1, len(y)):
            if y[i] - y[i - 1] < sep - 1e-9:
                falta = sep - (y[i] - y[i - 1])
                y[i - 1] -= falta / 2
                y[i] += falta / 2
                mogut = True
        if not mogut:
            break
    if ymin is not None and y[0] < ymin:
        y += ymin - y[0]
    if ymax is not None and y[-1] > ymax:
        y -= y[-1] - ymax
    out = np.empty_like(y)
    out[ordre] = y
    return out


def patch_estat(estat, **kw):
    """Handle de llegenda per als tres estats de conservació (FIG-4.18)."""
    if estat == 'perdut':
        return Patch(facecolor='white', edgecolor=GRIS[0], hatch=TRAMA, linewidth=0.5, **kw)
    return Patch(facecolor=ESTAT[estat], edgecolor='none', **kw)


# ---------------------------------------------------------------- literals publicats
DESAJUSTOS = []


def literal(nom, publicat, calculat):
    """Retorna sempre la xifra publicada a la memòria; si el càlcul sobre les
    dades no coincideix, ho anota a DESAJUSTOS (sense modificar la xifra)."""
    try:
        assert calculat == publicat, f'{nom}: publicat {publicat}, calculat {calculat}'
    except AssertionError as e:
        DESAJUSTOS.append(str(e))
    return publicat


# ---------------------------------------------------------------- eixides
OUT = 'output/figures/memoria/'
QUANTITZA_PNG = True


def desa(fig, nom):
    """SVG (sense data a les metadades) i PNG a 600 ppp, fons blanc."""
    os.makedirs(OUT, exist_ok=True)
    fig.savefig(OUT + nom + '.svg', facecolor='white', metadata={'Date': None})
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=PPP_PNG, facecolor='white')
    plt.close(fig)
    buf.seek(0)
    try:
        from PIL import Image
        im = Image.open(buf).convert('RGB')
        if QUANTITZA_PNG:
            im = im.quantize(colors=256, method=Image.Quantize.MEDIANCUT, dither=Image.Dither.NONE)
        im.save(OUT + nom + '.png', optimize=True, dpi=(PPP_PNG, PPP_PNG))
    except ImportError:
        with open(OUT + nom + '.png', 'wb') as f:
            f.write(buf.getvalue())
    print('  ', OUT + nom + '.svg', '+ .png')


# ---------------------------------------------------------------- comprovacions de text
PARAULES_PROHIBIDES = ['front', 'riquesa', 'atestat']


def textos_svg(cami):
    """Cadenes de text d'un SVG de matplotlib (svg.fonttype = none)."""
    s = open(cami, encoding='utf-8').read()
    ts = re.findall(r'<text[^>]*>(.*?)</text>', s, flags=re.S)
    out = []
    for t in ts:
        t = re.sub(r'<[^>]+>', '', t)
        t = t.replace('&#x27;', "'").replace('&quot;', '"').replace('&amp;', '&').replace('&lt;', '<').replace('&gt;', '>')
        out.append(t.strip())
    return out


def comprova_textos(dirsvg=OUT):
    """Cerca en tots els SVG les paraules i formats vetats. Retorna la llista de faltes."""
    faltes = []
    for f in sorted(os.listdir(dirsvg)):
        if not f.endswith('.svg'):
            continue
        ts = textos_svg(os.path.join(dirsvg, f))
        for i, t in enumerate(ts):
            tl = t.lower()
            seguent = ts[i + 1].lower() if i + 1 < len(ts) else ''  # les línies d'un text partit són <text> consecutius
            for p in PARAULES_PROHIBIDES:
                if re.search(r'\b' + p, tl):
                    faltes.append((f, t, p))
            if re.search(r'\brelleu\b', tl) and 'baix relleu' not in tl and 'formes del relleu' not in tl:
                faltes.append((f, t, 'relleu sense baix'))
            if re.search(r'\bcambra\b', tl) and not re.search(r'cambra\s+funerària', tl) and 'coberta de cambra' not in tl \
                    and not (tl.endswith('cambra') and seguent.startswith('funerària')):
                faltes.append((f, t, 'cambra sense funerària'))
            if re.search(r'\d\.\d', t):
                faltes.append((f, t, 'punt decimal'))
    return faltes
