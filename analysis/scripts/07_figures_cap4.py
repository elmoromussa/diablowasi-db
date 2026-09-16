#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
07_figures_cap4.py - Figures gràfiques del capítol 4 de la memòria del TFM
«Construir a l'abisme» (FIG-4.2 a 4.9, 4.11, 4.17, 4.18 i 4.19)

Versió per al paquet d'anàlisi del generador executat el 12/09/2026 per a la
memòria (tfm-diablowasi/figures/fonts/cap4/figures_cap4.py). El codi de les
figures és el mateix; només canvia l'entrada, que ací es llig íntegrament
dels CSV d'aquest paquet en lloc d'extractes intermedis de la BD:

Entrada:  data/T_STRUCTURES.csv, data/L_TYPOLOGY.csv, data/L_SUPPORT.csv,
          data/T_DECORATIONS.csv, data/T_CONNECTIONS.csv        (pas 00)
          data/matriu_AX_cinc_valors.csv                        (pas 01)
          output/tables/taula_perfil_clusters.csv,
          output/tables/taula_jaccard_elements.csv              (pas 02)
          output/tables/taula_4_1_sector_x_tipologia_106.csv,
          output/tables/riquesa_per_EA.csv                      (pas 05)
          output/tables/metriques_murs_per_EA.csv,
          output/tables/metriques_portals_per_EA.csv            (pas 06)
Sortida:  output/figures/memoria/FIG-4.Y_descripcio.svg (12 fitxers)

No s'hi generen: FIG-4.1 i 4.10 (cartografia, QGIS), FIG-4.12 a 4.15
(cronologia, OxCal), FIG-4.16 (esquema de motius, dibuix sense dades) ni les
fotografies. Diferències respecte del generador de la memòria: només SVG (la
versió d'inserció PNG viu a la memòria), sense marca de temps al fitxer, i
tipografia Arial quan TeX Gyre Heros no és instal·lada (mateixa mètrica
Helvetica).

Executar des d'analysis/:  python3 scripts/07_figures_cap4.py
"""
import os
import pandas as pd, numpy as np, matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import PathPatch, Patch
from matplotlib.path import Path
from matplotlib import font_manager
for f in ['/usr/share/texmf/fonts/opentype/public/tex-gyre/texgyreheros-regular.otf','/usr/share/texmf/fonts/opentype/public/tex-gyre/texgyreheros-bold.otf','/usr/share/texmf/fonts/opentype/public/tex-gyre/texgyreheros-italic.otf']:
    try: font_manager.fontManager.addfont(f)
    except Exception: pass
_installed = {f.name for f in font_manager.fontManager.ttflist}
FONT = next((f for f in ['TeX Gyre Heros', 'Arial', 'DejaVu Sans'] if f in _installed), 'sans-serif')
print('tipografia:', FONT)
plt.rcParams.update({'font.family':FONT,'font.size':9,'axes.spines.top':False,'axes.spines.right':False,
    'axes.edgecolor':'#444','axes.linewidth':0.6,'xtick.color':'#444','ytick.color':'#444','axes.labelcolor':'#222',
    'text.color':'#222','axes.titlesize':10,'axes.titleweight':'bold','axes.titlelocation':'left','legend.frameon':False,
    'savefig.dpi':300,'figure.dpi':100})
OUT='output/figures/memoria/'; os.makedirs(OUT,exist_ok=True)
def save(fig,name):
    fig.savefig(OUT+name+'.svg',bbox_inches='tight',facecolor='white',metadata={'Date':None}); plt.close(fig)
    print('  ', OUT+name+'.svg')

T='output/tables/'
TYPMAP={'EA-MAU Mausoleum/Chullpa':'MAU','EA-CAM Funerary Chamber':'CAM','EA-TER Ledge Terrace':'TER','EA-PLA-V Aerial Platform':'PLA-V','NIX Natural Niche':'NIX','CAV Cave/Cavern':'CAV','MEN Isolated Structural Element':'MEN','PR Rock Art':'PR','Unclassifiable':'NC'}

# ---------- entrada: extractes de la BD (pas 00) en lloc de S.pkl, D.pkl i C.pkl
_S=pd.read_csv('data/T_STRUCTURES.csv'); _typ=pd.read_csv('data/L_TYPOLOGY.csv'); _sup=pd.read_csv('data/L_SUPPORT.csv')
S=_S[['ID','Code','ID_Typology','ID_Support','ID_Support_Secondary','Facade_Azimuth_Deg','MNI']].copy()
S['Typ']=S.ID_Typology.map(dict(zip(_typ.ID,_typ.Name)))
S['Sup']=S.ID_Support.map(dict(zip(_sup.ID,_sup.Name))); S['Sup2']=S.ID_Support_Secondary.map(dict(zip(_sup.ID,_sup.Name)))
D=pd.read_csv('data/T_DECORATIONS.csv')[['ID_Structure','ID_Struct_Body','ID_Dec_Type']]
C=pd.read_csv('data/T_CONNECTIONS.csv')[['ID_Struct_A','ID_Struct_B']]
S['T']=S.Typ.map(TYPMAP)
TORD=['MAU','CAM','TER','PLA-V','NC','NIX','CAV','MEN','PR']
COL={'MAU':'#2a6f97','CAM':'#c8553d','TER':'#8cb369','PLA-V':'#7b5ea7','NIX':'#e6a13a','CAV':'#9aa5b1','NC':'#cfcfcf','MEN':'#5c5c5c','PR':'#f2d16b'}
TNAME={'MAU':'mausoleu','CAM':'cambra','TER':'terrassa','PLA-V':'plataforma volada','NIX':'nínxol natural','CAV':'cavitat','NC':'no classificable','MEN':'traça estructural','PR':'panell d\'art rupestre'}
GREY='#8a8a8a'

# ---------- FIG corpus: stacked bars per sector + donut
# (la taula del pas 05 porta sectors DW-Sxx i tipologies amb el nom complet)
t106=pd.read_csv(T+'taula_4_1_sector_x_tipologia_106.csv',index_col=0)
t106.columns=[TYPMAP.get(c.strip(),c.strip()) for c in t106.columns]; t106.index=[str(i).replace('DW-','') for i in t106.index]
sec=[s for s in t106.index if s.startswith('S')]
fig,(a1,a2)=plt.subplots(1,2,figsize=(7.2,3.0),gridspec_kw={'width_ratios':[1.7,1]})
left=np.zeros(len(sec))
for t in TORD:
    v=t106.loc[sec,t].values.astype(float)
    a1.barh(sec,v,left=left,color=COL[t],edgecolor='white',linewidth=0.6,label=t)
    for i,(x,w) in enumerate(zip(left,v)):
        if w>=4: a1.text(x+w/2,i,int(w),ha='center',va='center',fontsize=7,color='white' if t not in('NC','PR','CAV','NIX') else '#222')
    left+=v
for i,s in enumerate(sec): a1.text(left[i]+1,i,f'{int(left[i])}',va='center',fontsize=8,color='#444')
a1.invert_yaxis(); a1.set_xlim(0,80); a1.set_xlabel('elements arqueològics'); a1.set_title('a. Tipologies per sector (n = 106)')
a1.legend(ncol=3,fontsize=7,loc='lower right',handlelength=1,columnspacing=0.8)
groups=[('arquitectura\nconstruïda',54,'#2a6f97'),('contextos naturals\nocupats',35,'#e6a13a'),('no classificable',9,'#cfcfcf'),('traces i panells',8,'#5c5c5c')]
w,_=a2.pie([g[1] for g in groups],colors=[g[2] for g in groups],startangle=90,counterclock=False,wedgeprops=dict(width=0.38,edgecolor='white',linewidth=1.5))
for g,wd in zip(groups,w):
    ang=np.deg2rad((wd.theta1+wd.theta2)/2); r=1.22
    a2.text(r*np.cos(ang),r*np.sin(ang),f'{g[0]}\n{g[1]} ({g[1]/106*100:.0f} %)',ha='center',va='center',fontsize=7)
a2.set_title('b. Construir o ocupar')
a2.text(0,0,'106\nEA',ha='center',va='center',fontsize=10,fontweight='bold')
save(fig,'FIG-4.2_corpus_sector_tipologia')

# ---------- FIG suport: (a) tipologia -> primari ; (b) primari -> secundari (92)
pop=set(pd.read_csv(T+'riquesa_per_EA.csv').Code)
S92=S[S.Code.isin(pop)].copy()
SUPN={'Narrow natural ledge (<2m)':'repisa estreta','Micro-ledge (<50cm)':'microrepisa','Medium cavity (1-10m2)':'cavitat mitjana','Large cavity (>10m2)':'cavitat gran','Natural niche (<1m2)':'nínxol natural','Vertical cleft':'esquerda vertical','Bedding-plane recess':"junta d'estratificació",'Ground':'terreny','Rock dihedral':'diedre','Wide natural ledge (>2m)':'repisa ampla','Rock surface':'paret de roca'}
S92['P']=S92.Sup.map(SUPN); S92['Q']=S92.Sup2.map(SUPN).fillna('cap (una sola forma)')
PCOL={'repisa ampla':'#9dc3d6','repisa estreta':'#2a6f97','microrepisa':'#6aa5c9',"junta d'estratificació":'#7b5ea7','terreny':'#8d8d8d','esquerda vertical':'#b07aa1','diedre':'#c6a0b6','cavitat gran':'#c8553d','cavitat mitjana':'#e08a76','nínxol natural':'#e6a13a','paret de roca':'#bbbbbb','cap (una sola forma)':'#e5e5e5'}
def alluvial(ax,left_keys,right_keys,ct,lcol,title,lab_l,lab_r,gap=2.2):
    ax.axis('off')
    lcount=ct.sum(1); rcount=ct.sum(0)
    def stack(keys,counts):
        pos={};y=0
        for k in keys:
            pos[k]=(y,y+counts.get(k,0)); y+=counts.get(k,0)+gap
        return pos,y-gap
    lpos,lh=stack(left_keys,lcount); rpos,rh=stack(right_keys,rcount); H=max(lh,rh)
    loff=(H-lh)/2; roff=(H-rh)/2; bw=0.05; ax.set_xlim(-0.75,1.85); ax.set_ylim(-1,H+6)
    for k,(y0,y1) in lpos.items():
        if y1==y0: continue
        ax.add_patch(plt.Rectangle((0,y0+loff),bw,y1-y0,color=lcol(k),lw=0)); ax.text(-0.03,(y0+y1)/2+loff,f'{k}  {int(y1-y0)}',ha='right',va='center',fontsize=7.5)
    for k,(y0,y1) in rpos.items():
        if y1==y0: continue
        ax.add_patch(plt.Rectangle((1,y0+roff),bw,y1-y0,color='#aaa',lw=0)); ax.text(1+bw+0.03,(y0+y1)/2+roff,f'{int(y1-y0)}  {k}',ha='left',va='center',fontsize=7.5)
    lcur={k:lpos[k][0]+loff for k in lpos}; rcur={k:rpos[k][0]+roff for k in rpos}
    for l in left_keys:
        for r in right_keys:
            v=ct.loc[l,r] if (l in ct.index and r in ct.columns) else 0
            if not v: continue
            y0l=lcur[l]; y1l=y0l+v; y0r=rcur[r]; y1r=y0r+v; lcur[l]=y1l; rcur[r]=y1r
            verts=[(bw,y0l),(0.5,y0l),(0.5,y0r),(1,y0r),(1,y1r),(0.5,y1r),(0.5,y1l),(bw,y1l),(bw,y0l)]
            codes=[Path.MOVETO,Path.CURVE4,Path.CURVE4,Path.CURVE4,Path.LINETO,Path.CURVE4,Path.CURVE4,Path.CURVE4,Path.CLOSEPOLY]
            ax.add_patch(PathPatch(Path(verts,codes),facecolor=lcol(l),alpha=0.55,lw=0))
    ax.text(-0.75,H+4.6,title,fontsize=9,fontweight='bold',ha='left'); ax.text(0,H+1.2,lab_l,fontsize=7.5,ha='left',color='#444'); ax.text(1+bw,H+1.2,lab_r,fontsize=7.5,ha='left',color='#444')
lord=[t for t in TORD[::-1] if t in set(S92['T'])]
pord=['repisa ampla','repisa estreta','microrepisa',"junta d'estratificació",'terreny','esquerda vertical','diedre','cavitat gran','cavitat mitjana','nínxol natural'][::-1]
pord=[p for p in pord if p in set(S92.P)]
qord=['diedre','paret de roca','esquerda vertical',"junta d'estratificació",'terreny','microrepisa','repisa ampla','cavitat gran','cavitat mitjana','cap (una sola forma)'][::-1]
qord=[q for q in qord if q in set(S92.Q)]
fig,(a1,a2)=plt.subplots(1,2,figsize=(7.4,4.4))
alluvial(a1,lord,pord,pd.crosstab(S92['T'],S92.P),lambda k:COL[k],'a. Tipologia → suport primari','tipologia','suport primari (rep la càrrega)')
alluvial(a2,pord,qord,pd.crosstab(S92.P,S92.Q),lambda k:PCOL[k],'b. Suport primari → suport secundari','suport primari','suport secundari (estabilitza o allotja)')
plt.subplots_adjust(wspace=0.9)
save(fig,'FIG-4.3_suport_geologic')

# ---------- FIG murs
W=pd.read_csv(T+'metriques_murs_per_EA.csv'); W['T']=W.Tipologia.map(TYPMAP)
fig,(a1,a2)=plt.subplots(1,2,figsize=(7.2,3.0),gridspec_kw={'width_ratios':[1.6,1]})
order=['CAM','MAU','TER','NIX','NC']; rng=np.random.default_rng(3)
for i,t in enumerate(order):
    d=W[W['T']==t].Area_Cons_m2.values
    a1.scatter(d,np.full(len(d),i)+rng.uniform(-0.18,0.18,len(d)),s=28,color=COL[t],edgecolor='white',linewidth=0.5,zorder=3)
    if len(d)>=3: a1.plot([np.median(d)]*2,[i-0.3,i+0.3],color='#222',lw=1.2,zorder=4)
    a1.text(7.3,i,f'n = {len(d)}',va='center',fontsize=7,color='#555')
a1.set_yticks(range(len(order))); a1.set_yticklabels([TNAME[t] for t in order]); a1.invert_yaxis(); a1.set_xlim(0,8.2)
a1.set_xlabel('àrea de mur conservada per estructura (m²)'); a1.set_title('a. Mur conservat per tipologia (n = 28)')
ic=W.Index_conserv_pct.dropna().sort_values(ascending=False).values
a2.bar(range(len(ic)),ic,color=['#2a6f97' if v>=99.9 else '#9aa5b1' for v in ic],width=0.8)
a2.axhline(100,color='#ccc',lw=0.6); a2.set_ylim(0,105); a2.set_xticks([]); a2.set_ylabel('% de l\'àrea original conservada')
a2.set_title(f'b. Conservació (n = {len(ic)})'); a2.text(len(ic)-0.5,20,'17 al 100 %',ha='right',fontsize=7,color='#2a6f97')
save(fig,'FIG-4.4_murs_conservacio')

# ---------- FIG portals
P=pd.read_csv(T+'metriques_portals_per_EA.csv'); P['T']=P.Tipologia.map(TYPMAP); Pm=P.dropna(subset=['Amplaria_m'])
fig,(a1,a2)=plt.subplots(1,2,figsize=(7.2,3.0),gridspec_kw={'width_ratios':[1.15,1]})
for t in ['CAM','MAU']:
    d=Pm[Pm['T']==t]; a1.scatter(d.Amplaria_m,d.Alcaria_m,s=40,color=COL[t],edgecolor='white',linewidth=0.6,label=TNAME[t],zorder=3)
a1.plot([0.3,0.85],[0.3,0.85],color='#bbb',lw=0.8,ls='--'); a1.text(0.72,0.735,'quadrat',fontsize=7,color='#888',ha='right')
a1.set_xlim(0.3,0.85); a1.set_ylim(0.4,0.75); a1.set_xlabel('amplària (m)'); a1.set_ylabel('alçària (m)'); a1.set_aspect('equal')
a1.set_title(f'a. Llum dels portals (n = {len(Pm)})'); a1.legend(fontsize=7,loc='lower right')
a1.axvline(0.53,color='#ddd',lw=0.6); a1.axhline(0.58,color='#ddd',lw=0.6)
jam=[('composta',19,'#2a6f97'),('fàbrica',9,'#8cb369'),('llosa',7,'#e6a13a')]
a2.barh([j[0] for j in jam],[j[1] for j in jam],color=[j[2] for j in jam],height=0.6)
for i,j in enumerate(jam): a2.text(j[1]+0.4,i,f'{j[1]}  ({j[1]/35*100:.0f} %)',va='center',fontsize=8)
a2.invert_yaxis(); a2.set_xlim(0,26); a2.set_xlabel('muntants observats (n = 35)'); a2.set_title('b. Com es resol el muntant')
save(fig,'FIG-4.5_portals_muntants')

# ---------- FIG riquesa per tipologia
# (la taula del pas 05 porta la tipologia amb el nom complet)
R=pd.read_csv(T+'riquesa_per_EA.csv'); R['Tipologia']=R.Tipologia.map(TYPMAP)
fig,ax=plt.subplots(figsize=(7.2,3.0)); order=['CAM','MAU','PLA-V','TER','NC','NIX','CAV']
for i,t in enumerate(order):
    d=R[R.Tipologia==t].Riquesa.values
    ax.scatter(np.full(len(d),i)+rng.uniform(-0.22,0.22,len(d)),d,s=26,color=COL[t],edgecolor='white',linewidth=0.5,zorder=3,alpha=0.9)
    q1,med,q3=np.percentile(d,[25,50,75]); ax.plot([i-0.32,i+0.32],[med,med],color='#222',lw=1.4,zorder=4); ax.plot([i,i],[q1,q3],color='#222',lw=0.8,zorder=2,alpha=0.5)
    ax.text(i,-1.6,f'n = {len(d)}',ha='center',fontsize=7,color='#555')
ax.set_xticks(range(len(order))); ax.set_xticklabels([TNAME[t] for t in order],fontsize=8); ax.set_ylim(-2.2,16); ax.set_ylabel('elements construïts per estructura')
ax.set_title('Riquesa constructiva per tipologia (n = 92; barra = mediana, línia = rang interquartílic)')
for x0,x1,lab in [(-0.4,1.4,'cambra i mausoleu'),(1.6,3.4,'plataforma i terrassa'),(3.6,6.4,'contextos naturals i NC')]:
    ax.plot([x0,x1],[15.3,15.3],color='#999',lw=0.8); ax.text((x0+x1)/2,15.5,lab,ha='center',fontsize=7,color='#666')
save(fig,'FIG-4.6_complexitat_tipologia')

# ---------- FIG families: presence matrix by typology + profile
M=pd.read_csv('data/matriu_AX_cinc_valors.csv'); M=M[M.Code.isin(pop)].set_index('Code')
els=['B','C','D','E','F','G','I','J','K','L','M','N','O','Q','R','S','T','V','X','Z']
B=M[els].isin([1,2,3]).astype(int)
tmap=dict(zip(S.Code,S['T'])); B['typ']=[tmap[c] for c in B.index]; B['rq']=B[els].sum(1)
order=['CAM','MAU','PLA-V','TER','NC','NIX','CAV']
B['ti']=B['typ'].map({t:i for i,t in enumerate(order)}); B=B.sort_values(['ti','rq'],ascending=[True,False])
fig=plt.figure(figsize=(7.4,6.2)); gs=fig.add_gridspec(1,2,width_ratios=[1.5,1],wspace=0.3)
ax=fig.add_subplot(gs[0]); n=len(B)
mat=B[els].values.T  # elements x structures
ax.imshow(np.where(mat==1,1,np.nan),cmap='Greys',vmin=0,vmax=1.4,aspect='auto',interpolation='nearest')
for i in range(n):
    ax.add_patch(plt.Rectangle((i-0.5,-1.5),1,0.8,color=COL[B['typ'].iloc[i]],lw=0,clip_on=False))
ax.set_yticks(range(len(els))); ax.set_yticklabels(els,fontsize=7); ax.set_xticks([]); ax.tick_params(length=0)
ax.set_ylim(len(els)-0.5,-3.0)
for sp in ax.spines.values(): sp.set_visible(False)
# typology group separators and labels
bounds=B.groupby('ti').size().cumsum().values; prev=0
for k,b in enumerate(bounds):
    ax.axvline(b-0.5,color='white',lw=1.5); ax.text((prev+b)/2-0.5,-2.0,order[k],ha='center',va='bottom',fontsize=6.5,rotation=0,color=COL[order[k]]); prev=b
ax.set_title('a. Elements construïts per estructura (92 × 20), agrupades per tipologia',fontsize=9,loc='left',pad=14)
ax.text(-0.02,1.01,'',transform=ax.transAxes)
ax2=fig.add_subplot(gs[1]); PF=pd.read_csv(T+'taula_perfil_clusters.csv',index_col=0)
fam_order=[2,3,1,5,4]; famlab={2:'2  tombes amb portal (32)',3:'3  plataforma (5)',1:'1  basament (22)',5:'5  murets (2)',4:'4  no construït (31)'}
PFo=PF.loc[fam_order]
ax2.imshow(PFo.values.T,cmap='Blues',vmin=0,vmax=1,aspect='auto')
ax2.set_yticks(range(len(PF.columns))); ax2.set_yticklabels(PF.columns,fontsize=7); ax2.set_xticks(range(len(fam_order))); ax2.set_xticklabels([famlab[f] for f in fam_order],fontsize=7,rotation=35,ha='right')
for i in range(PFo.shape[0]):
    for j in range(PFo.shape[1]):
        v=PFo.values[i,j]
        if v>=0.05: ax2.text(i,j,f'{v*100:.0f}',ha='center',va='center',fontsize=6.5,color='white' if v>0.55 else '#333')
ax2.set_title('b. Perfil de les famílies (%)',fontsize=9,loc='left'); ax2.tick_params(length=0)
for sp in ax2.spines.values(): sp.set_visible(False)
save(fig,'FIG-4.7_families_constructives')

# ---------- FIG coocurrencia
J=pd.read_csv(T+'taula_jaccard_elements.csv',index_col=0)
fig,ax=plt.subplots(figsize=(5.6,5.2)); n=len(J)
mask=np.triu(np.ones_like(J.values,dtype=bool))
vals=np.where(mask,np.nan,J.values)
im=ax.imshow(vals,cmap='Oranges',vmin=0,vmax=1)
ax.set_xticks(range(n)); ax.set_xticklabels(J.columns); ax.set_yticks(range(n)); ax.set_yticklabels(J.index); ax.tick_params(length=0)
for i in range(n):
    for j in range(i):
        v=J.values[i,j]
        if v>=0.5: ax.text(j,i,f'{v:.2f}',ha='center',va='center',fontsize=6,color='white' if v>0.7 else '#222')
for s in ax.spines.values(): s.set_visible(False)
ax.set_title('Co-ocurrència entre elements (índex de Jaccard); valors ≥ 0,50 etiquetats')
cb=fig.colorbar(im,ax=ax,shrink=0.5,pad=0.02); cb.outline.set_visible(False); cb.set_label('índex de Jaccard',fontsize=8)
save(fig,'FIG-4.8_coocurrencia_jaccard')

# ---------- FIG graf
import networkx as nx
codemap=dict(zip(S.ID,S.Code))
G=nx.Graph()
for c in S.Code: G.add_node(c)
for _,r in C.iterrows(): G.add_edge(codemap[r.ID_Struct_A],codemap[r.ID_Struct_B])
art=set(nx.articulation_points(G)); comps=sorted(nx.connected_components(G),key=len,reverse=True)
fig,ax=plt.subplots(figsize=(7.2,5.0)); ax.axis('off')
pos={};
# lay out components in a grid-ish packing
big=[c for c in comps if len(c)>1]; iso=[c for c in comps if len(c)==1]
xoff=0; y=0; row_h=0; cursor=[0,0]
layouts=[]
for c in big:
    sub=G.subgraph(c); p=nx.kamada_kawai_layout(sub); p={k:np.array(v) for k,v in p.items()}
    xs=np.array([v[0] for v in p.values()]); ys=np.array([v[1] for v in p.values()])
    scale=np.sqrt(len(c))*0.55; w=(xs.max()-xs.min())*scale+0.6; h=(ys.max()-ys.min())*scale+0.6
    layouts.append((p,scale,w,h,xs.min(),ys.min()))
# pack: first row big three, second row small
X=0;Y=0;maxh=0
for i,(p,scale,w,h,xmin,ymin) in enumerate(layouts):
    if i==3: X=0; Y-= maxh+0.5; maxh=0
    for k,v in p.items(): pos[k]=np.array([X+(v[0]-xmin)*scale+0.3, Y-(v[1]-ymin)*scale-0.3])
    X+=w+0.4; maxh=max(maxh,h)
# isolated: bottom row
Y-=maxh+0.6; X=0
for c in iso:
    (k,)=tuple(c); pos[k]=np.array([X,Y]); X+=0.9
nx.draw_networkx_edges(G,pos,ax=ax,edge_color='#bbb',width=0.8)
nc=[COL[tmap[c]] for c in G.nodes]; sz=[70 if c in art else 28 for c in G.nodes]; ec=['#222' if c in art else 'white' for c in G.nodes]
nx.draw_networkx_nodes(G,pos,ax=ax,node_color=nc,node_size=sz,edgecolors=ec,linewidths=[1.2 if c in art else 0.5 for c in G.nodes])
for c in ['DW-S01-EA02','DW-S01-EA33','DW-S01-EA22','DW-S01-EA55','DW-S04-EA11','DW-S01-EA39','DW-S04-EA20']:
    ax.annotate(c.replace('DW-',''),pos[c],xytext=(4,4),textcoords='offset points',fontsize=6.5,fontweight='bold')
for i,c in enumerate(big[:3]):
    xs=[pos[k][0] for k in c]; ys=[pos[k][1] for k in c]; ax.text(np.mean(xs),max(ys)+0.35,f'component de {len(c)}',ha='center',fontsize=7.5,color='#555')
ax.text(0,Y-0.6,f'{len(iso)} estructures aïllades',fontsize=7.5,color='#555')
ax.legend(handles=[Patch(color=COL[t],label=t) for t in ['MAU','CAM','TER','PLA-V','NIX','CAV','NC','MEN','PR']]+[plt.Line2D([],[],marker='o',color='w',markerfacecolor='white',markeredgecolor='#222',markersize=8,label="punt d'articulació")],ncol=5,fontsize=7,loc='upper center',bbox_to_anchor=(0.5,0.0),handlelength=1)
ax.set_title('El front com a xarxa: 106 estructures, 91 vincles físics, 30 components')
save(fig,'FIG-4.9_xarxa_components')
print('components',[len(c) for c in comps][:9],'art',len(art))

# ---------- FIG rosa
az=S.Facade_Azimuth_Deg.dropna().values.astype(float)
fig=plt.figure(figsize=(3.4,3.4)); ax=fig.add_subplot(111,projection='polar')
ax.set_theta_zero_location('N'); ax.set_theta_direction(-1)
bins=np.deg2rad(np.arange(0,361,20)); h,_=np.histogram(np.deg2rad(az),bins=bins)
ax.bar(bins[:-1]+np.deg2rad(10),h,width=np.deg2rad(20),color='#2a6f97',alpha=0.85,edgecolor='white')
ax.set_xticks(np.deg2rad([0,90,180,270])); ax.set_xticklabels(['N','E','S','O']); ax.set_yticks([5,10]); ax.set_yticklabels(['5','10'],fontsize=7); ax.grid(color='#ddd',lw=0.5)
ax.plot([np.deg2rad(305.4)]*2,[0,h.max()+1],color='#c8553d',lw=1.5)
ax.set_title('Orientació de les façanes (n = 34)\nmitjana 305°, R = 0,86',fontsize=8,fontweight='normal',loc='center',pad=12)
save(fig,'FIG-4.11_rosa_orientacions')

# ---------- FIG decoracio
fam={1:'relleu',2:'relleu',3:'relleu',4:'relleu',5:'relleu',6:'relleu',8:'relleu',7:'pintura',9:'pintura',18:'pintura',19:'pintura',11:'art rupestre',12:'art rupestre',15:'art rupestre',17:'art rupestre',20:'art rupestre',10:'ND'}
D['fam']=D.ID_Dec_Type.map(fam)
posmap={1:'cos basal',4:'pilastres',5:'flancs de façana',6:'murs de retorn',10:'dintell',11:'sobre el dintell (fris)',14:'roca (superposada)',15:'roca (perímetre)',16:'roca (pròxima)',17:'panell independent'}
D['posn']=D.ID_Struct_Body.map(posmap)
FAMCOL={'relleu':'#2a6f97','pintura':'#c8553d','art rupestre':'#e6a13a'}
fig,(a1,a2)=plt.subplots(1,2,figsize=(7.2,3.2),gridspec_kw={'width_ratios':[1.3,1]})
porder=['cos basal','pilastres','flancs de façana','murs de retorn','dintell','sobre el dintell (fris)','roca (superposada)','roca (perímetre)','roca (pròxima)','panell independent']
ctp=pd.crosstab(D.posn,D.fam).reindex(porder).fillna(0)
left=np.zeros(len(porder))
for f in ['relleu','pintura','art rupestre']:
    v=ctp[f].values; a1.barh(porder,v,left=left,color=FAMCOL[f],edgecolor='white',label=f); left+=v
a1.invert_yaxis(); a1.set_xlabel('registres decoratius'); a1.set_title('a. On es col·loca cada família decorativa')
a1.axhline(5.5,color='#999',lw=0.6,ls=':'); a1.text(21.5,5.25,'sobre la fàbrica ↑',fontsize=7,color='#666',ha='right'); a1.text(21.5,5.9,'sobre la roca ↓',fontsize=7,color='#666',ha='right',va='top'); a1.set_xlim(0,22)
a1.legend(fontsize=7,loc='lower right')
st=S.set_index('ID'); D['T']=D.ID_Structure.map(st['T'])
dec=D.groupby(['T','fam']).ID_Structure.nunique().unstack().fillna(0)
tot=S['T'].value_counts()
torder=['CAM','MAU','TER','PLA-V','NIX','CAV','NC']
anyd=D.groupby('T').ID_Structure.nunique()
x=np.arange(len(torder)); wbar=0.26
for k,f in enumerate(['relleu','pintura','art rupestre']):
    a2.bar(x+(k-1)*wbar,[dec.loc[t,f]/tot[t]*100 if t in dec.index else 0 for t in torder],width=wbar,color=FAMCOL[f],label=f)
a2.set_xticks(x); a2.set_xticklabels(torder,fontsize=8); a2.set_ylabel('% d\'estructures de la tipologia'); a2.set_title('b. Quines estructures reben decoració')
a2.set_xticklabels([f'{t}\n{int(anyd.get(t,0))}/{tot[t]}' for t in torder],fontsize=7)
a2.set_ylim(0,100)
save(fig,'FIG-4.17_decoracio_families')

# ---------- FIG afectacio
af=[('Z','superfície de plataforma *',0,4,8),('E','mènsules de fusta',3,15,0),('J','cantoneres',1,4,0),('D','murets de trava',2,6,0),('R','coronament',4,11,0),('X','coberta de cambra',3,2,3),('F','bigues transversals *',3,3,6),('M','fris',8,10,0),('V','murs de retorn',8,9,0),('L','flancs de façana',15,15,1),('G','filades en voladís',1,1,0),('K','pilastres',5,5,0),('T','superfície de ràfec',5,4,0),('B','basament',31,22,0),('O','brancals',10,7,0),('I','cornisa',19,11,0),('C','sòcol',5,1,0),('N','llindar',7,1,0),('Q','dintell *',16,0,12),('S','biga de ràfec',2,0,0)]
fig,ax=plt.subplots(figsize=(7.2,4.4))
labs=[f'{a[0]}  {a[1]}' for a in af]; n=np.array([a[2]+a[3]+a[4] for a in af]); c=np.array([a[2] for a in af])/n; p=np.array([a[3] for a in af])/n; l=np.array([a[4] for a in af])/n
y=np.arange(len(af))
ax.barh(y,c*100,color='#2a6f97',label='sencer'); ax.barh(y,p*100,left=c*100,color='#9fc5e8',label='parcial'); ax.barh(y,l*100,left=(c+p)*100,color='#c8553d',label='perdut (atestat)')
for i,a in enumerate(af): ax.text(101,i,f'n = {n[i]}',va='center',fontsize=7,color='#555')
wood={'E','F','S'}
for i,a in enumerate(af):
    if a[0] in wood: ax.get_yticklabels
ax.set_yticks(y); ax.set_yticklabels(labs,fontsize=8); ax.invert_yaxis(); ax.set_xlim(0,112); ax.set_xlabel('% dels casos construïts')
for t in ax.get_yticklabels():
    if t.get_text().split()[0] in wood: t.set_color('#7a4b1e'); t.set_fontweight('bold')
ax.legend(ncol=3,fontsize=7,loc='upper center',bbox_to_anchor=(0.5,-0.12)); ax.set_title('Estat de conservació de cada element construït (elements de fusta en negreta)')
save(fig,'FIG-4.18_conservacio_elements')

# ---------- FIG MNI
m=S[S.MNI.notna()].copy(); m['exc']=m.Code.isin(['DW-S01-EA17','DW-S01-EA01'])
m=m.sort_values('MNI',ascending=False).reset_index(drop=True)
fig,ax=plt.subplots(figsize=(7.2,2.6))
ax.bar(range(len(m)),m.MNI,color=[COL[t] for t in m['T']],edgecolor=['#222' if e else 'white' for e in m.exc],linewidth=1.0)
ax.set_xticks(range(len(m))); ax.set_xticklabels([c.replace('DW-','') for c in m.Code],rotation=90,fontsize=6.5)
ax.set_ylabel('MNI'); ax.set_title('Nombre mínim d\'individus per estructura (25 EA amb restes comptables; contorn negre = excavades el 2023)')
for i,r in m.iterrows():
    if r.MNI>=5: ax.text(i,r.MNI+0.8,int(r.MNI),ha='center',fontsize=7)
ax.text(24,30,'44 EA sense restes observades\n37 EA amb interior no observable',ha='right',fontsize=7,color='#555')
save(fig,'FIG-4.19_mni_estructures')
print('done')
