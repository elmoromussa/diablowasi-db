#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
06_metriques_cap4.py - Metriques del construit per EA (cap. 4.2 de la memoria)

Deriva les dues taules de metriques que la memoria cita a l'apartat 4.2:

  output/tables/metriques_murs_per_EA.csv
      Una fila per EA amb metriques de mur (28 EA): plans de mur mesurats,
      cossos amb area conservada, area conservada total, cossos amb poligon
      complet / restituit / sense referencia d'original, area original de
      referencia i index de conservacio (area conservada / original, nomes
      sobre els cossos amb referencia).
  output/tables/metriques_portals_per_EA.csv
      Una fila per portal present (complet o parcial) segons Sys_Portal
      (22 EA, 23 portals): parella de muntants, posicio, amplaria i alcaria
      de l'obertura, ratio amplaria/alcaria i llum (A x H).

Entrada:  data/metriques_murs_v5.csv          (versionada; shapes de Metashape,
                                               conveni v1.1)
          data/T_STRUCTURES.csv, data/L_TYPOLOGY.csv   (pas 00)
          output/tables/riquesa_per_EA.csv    (pas 05; riquesa constructiva)
          ../data/metriques_v25/metriques_shapes_audit.csv
                                              (nomes per al segon portal de
                                               DW-S01-EA02, no incorporat a la BD)
Sortida:  output/tables/metriques_murs_per_EA.csv
          output/tables/metriques_portals_per_EA.csv
          output/logs/resultats_metriques_cap4.txt

Executar des d'analysis/:  python3 scripts/06_metriques_cap4.py
"""
import pandas as pd
import numpy as np

S = pd.read_csv("data/T_STRUCTURES.csv")
# Opening_* son SINGLE a Access: es normalitzen a mil.limetres
for c in ["Opening_Width_m", "Opening_Height_m"]:
    S[c] = S[c].round(3)
typ = pd.read_csv("data/L_TYPOLOGY.csv")
S = S.merge(typ[["ID", "Name"]].rename(columns={"ID": "ID_Typology",
                                                "Name": "Tipologia"}),
            on="ID_Typology", how="left")
R = pd.read_csv("output/tables/riquesa_per_EA.csv")
MU = pd.read_csv("data/metriques_murs_v5.csv", encoding="utf-8-sig")
AU = pd.read_csv("../data/metriques_v25/metriques_shapes_audit.csv",
                 encoding="utf-8-sig")

log = []


def out(s=""):
    print(s)
    log.append(s)


# ---------------------------------------------------------------- murs
# Cada fila de metriques_murs_v5 es un cos de mur (Code, Wall, Body) o una
# instancia d'un cos fragmentat (Instance a/b). Area_Cons_m2 es l'area
# conservada mesurada; Area_Complete_m2 / Area_Restituted_m2 marquen si el
# poligon es complet o restituit; Area_Orig_m2 es la referencia d'original
# (nomes quan es pot establir).
def agg_murs(g):
    cons = g.Area_Cons_m2.notna()
    ref = g.Area_Orig_m2.notna()
    orig = g.loc[ref, "Area_Orig_m2"].sum() if ref.any() else np.nan
    # La referencia d'original val per a tot el cos (pla + cos): les
    # instancies a/b d'un cos fragmentat hereten la del cos restituit.
    cos_ref = g[["Wall", "Body"]].apply(tuple, axis=1).isin(
        g.loc[ref, ["Wall", "Body"]].apply(tuple, axis=1))
    # index: area conservada dels cossos amb referencia / area original
    cons_ref = g.loc[cos_ref & cons, "Area_Cons_m2"].sum()
    idx = round(cons_ref / orig * 100, 3) if ref.any() and orig > 0 else np.nan
    n_cons = int(cons.sum())
    n_comp = int(g.Area_Complete_m2.notna().sum())
    n_rest = int(g.Area_Restituted_m2.notna().sum())
    return pd.Series({
        "Plans_mur": "".join(sorted(g.Wall.unique())),
        "N_cossos_amb_area": n_cons,
        "Area_Cons_m2": round(g.Area_Cons_m2.sum(), 3),
        "n_complets": n_comp,
        "n_restituits": n_rest,
        # resta, com a la memoria; 0 quan un cos es alhora complet i
        # restituit (DW-S04-EA03, f 1.1)
        "n_sense_ref": max(0, n_cons - n_comp - n_rest),
        "Area_Orig_ref_m2": round(orig, 3) if ref.any() else np.nan,
        "Index_conserv_pct": idx,
    })


murs = MU.groupby("Code").apply(agg_murs, include_groups=False).reset_index()
murs = murs.merge(S[["Code", "Tipologia"]], on="Code", how="left")
murs = murs.merge(R[["Code", "Riquesa"]], on="Code", how="left")
murs = murs[["Code", "Tipologia", "Riquesa", "Plans_mur", "N_cossos_amb_area",
             "Area_Cons_m2", "n_complets", "n_restituits", "n_sense_ref",
             "Area_Orig_ref_m2", "Index_conserv_pct"]].sort_values("Code")
murs.to_csv("output/tables/metriques_murs_per_EA.csv", index=False)

out("=== Metriques de mur per EA (metriques_murs_v5.csv) ===")
out(f"EA amb metriques de mur: {len(murs)} "
    f"(fora de la poblacio de 92: {int(murs.Riquesa.isna().sum())})")
out("Per tipologia: " + ", ".join(f"{k} {v}" for k, v in
                                  murs.Tipologia.value_counts().items()))
out(f"Cossos de mur amb area conservada: {int(MU.Area_Cons_m2.notna().sum())} "
    f"sobre {MU.Wall.nunique()} plans ("
    + ", ".join(f"{k} {v}" for k, v in MU.Wall.value_counts().items()) + ")")
out(f"Poligons complets {int(MU.Area_Complete_m2.notna().sum())}, restituits "
    f"{int(MU.Area_Restituted_m2.notna().sum())}, sense referencia d'original "
    f"{int((MU.Area_Cons_m2.notna() & MU.Area_Orig_m2.isna()).sum())}")
a = murs.Area_Cons_m2
out(f"Area conservada per EA: min {a.min():.2f}, mediana {a.median():.2f}, "
    f"max {a.max():.2f} m2 ({murs.loc[a.idxmax(), 'Code']}); total {a.sum():.2f} m2")
for t, g in murs.groupby("Tipologia"):
    out(f"  {t}: n = {len(g)}, {g.Area_Cons_m2.min():.1f}-{g.Area_Cons_m2.max():.1f} m2, "
        f"mediana {g.Area_Cons_m2.median():.1f}")
ic = murs.dropna(subset=["Index_conserv_pct"])
out(f"Index de conservacio computable en {len(ic)} EA; al 100 %: "
    f"{int((ic.Index_conserv_pct >= 99.9).sum())}; minims: "
    + ", ".join(f"{r.Code} {r.Index_conserv_pct:.0f} %" for r in
                ic.nsmallest(2, "Index_conserv_pct").itertuples()))
out(f"Volume_Prism_m3 informat en {int(MU.Volume_Prism_m3.notna().sum())} cossos "
    f"de {MU.loc[MU.Volume_Prism_m3.notna(), 'Code'].nunique()} EA: no reportable")

# ------------------------------------------------------------- portals
FONT_BD = "T_STRUCTURES.Opening_Width_m/Opening_Height_m (v26)"
pres = S[S.Sys_Portal.isin(["Present complete", "Present partial"])].copy()
rows = []
for r in pres.sort_values("Code").itertuples():
    amb_dada = pd.notna(r.Opening_Width_m) and pd.notna(r.Opening_Height_m)
    rows.append({"Code": r.Code, "Portal": "1.1", "Tipologia": r.Tipologia,
                 "Estat_portal": r.Sys_Portal, "Jamb_Fabric": r.Jamb_Fabric,
                 "Posicio": r.Portal_Position, "Amplaria_m": r.Opening_Width_m,
                 "Alcaria_m": r.Opening_Height_m,
                 "Font": FONT_BD if amb_dada else "sense dada a la BD"})

# Segon portal: la BD guarda una sola obertura per EA. L'unica EA amb dos
# portals es DW-S01-EA02 (nivells N1.1 i N1.2, Systems_Notes v25g); la llum
# del segon prove de la shape f-port-1.2-c de l'auditoria de Metashape del
# 30/08/2026 i la parella de muntants, de Systems_Notes ("llosa + composta").
extra = AU[(AU.Element == "port") & (AU.Plane == "f") & (AU.Valid == "si")
           & (AU.Body != 1.1) & AU.Code.isin(pres.Code)]
NOTES = {("DW-S01-EA02", 1.2): "Slab-Composite (Systems_Notes)"}
for e in extra.itertuples():
    base = next(x for x in rows if x["Code"] == e.Code)
    base["Portal"] = "1.1 (N1.1)"
    rows.append({"Code": e.Code, "Portal": f"{e.Body} (N{e.Body})",
                 "Tipologia": base["Tipologia"],
                 "Estat_portal": base["Estat_portal"],
                 "Jamb_Fabric": NOTES.get((e.Code, e.Body), "ND"),
                 "Posicio": base["Posicio"], "Amplaria_m": e.Width_m,
                 "Alcaria_m": e.Height_m,
                 "Font": f"metriques_shapes_audit.csv 30/08/2026, shape "
                         f"{e.Label} (no a la BD)"})

por = pd.DataFrame(rows).sort_values(["Code", "Portal"]).reset_index(drop=True)
por["Ratio_A_H"] = (por.Amplaria_m / por.Alcaria_m).round(2)
por["Llum_m2"] = (por.Amplaria_m * por.Alcaria_m).round(3)
por.to_csv("output/tables/metriques_portals_per_EA.csv", index=False)

out()
out("=== Portals (T_STRUCTURES v26: Sys_Portal, Opening_Width_m, Opening_Height_m) ===")
out("Cens Sys_Portal: " + ", ".join(f"{k} {v}" for k, v in
                                    S.Sys_Portal.value_counts(dropna=False).items()))
out(f"EA amb portal present: {pres.Code.nunique()} "
    f"({int((pres.Sys_Portal == 'Present complete').sum())} complets, "
    f"{int((pres.Sys_Portal == 'Present partial').sum())} parcials); "
    f"portals: {len(por)}")
m = por.dropna(subset=["Amplaria_m", "Alcaria_m"])
mb = m[m.Font == FONT_BD]
out(f"Llum mesurada a la BD en {len(mb)} portals de {mb.Code.nunique()} EA; "
    f"amb el segon portal d'EA02, n = {len(m)}")
out("Sense dada: " + ", ".join(
    f"{r.Code} ({r.Estat_portal.split()[1]})" for r in
    por[por.Font == "sense dada a la BD"].itertuples()))
for lab, col in [("Amplaria (m)", "Amplaria_m"), ("Alcaria (m)", "Alcaria_m"),
                 ("Ratio A/H", "Ratio_A_H"), ("Llum A x H (m2)", "Llum_m2")]:
    for nom, d in [("BD", mb), ("BD + EA02 N1.2", m)]:
        out(f"  {lab}, {nom} (n = {len(d)}): min {d[col].min():.2f}, "
            f"mediana {d[col].median():.2f}, max {d[col].max():.2f}")
out("Parella de muntants: " + ", ".join(f"{k} {v}" for k, v in
                                        por.Jamb_Fabric.value_counts().items()))
out("Posicio del portal: " + ", ".join(f"{k} {v}" for k, v in
                                       por.Posicio.value_counts(dropna=False).items()))

with open("output/logs/resultats_metriques_cap4.txt", "w", encoding="utf-8") as f:
    f.write("\n".join(log) + "\n")
