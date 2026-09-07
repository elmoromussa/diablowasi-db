#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
01_build_matriu_AX.py - Construccio de la matriu d'elements A-X (corpus DW v26)

Replica fidelment la logica de QRY_13_AX_Pattern_Export de la BD:
cada element es anul.lat (NULL) quan el seu sistema contenidor (Sys_*) es
'Not applicable' o 'Not observable' (gating), Chamber_Roof s'anul.la quan la
coberta es 'Natural bedrock'.

Entrada:  data/T_STRUCTURES.csv, data/L_SECTORS.csv, data/L_TYPOLOGY.csv
          (generats per 00_extract_data.sh)
Sortida:  data/matriu_AX_cinc_valors.csv (106 EA x 21 elements, valors del
          domini 0/1/2/3/9 despres del gating, mes Code/Sector/Typology/
          Record_Class)

Executar des d'analysis/:  python3 scripts/01_build_matriu_AX.py
"""
import pandas as pd

S = pd.read_csv("data/T_STRUCTURES.csv")
sec = pd.read_csv("data/L_SECTORS.csv")
typ = pd.read_csv("data/L_TYPOLOGY.csv")

S = S.merge(sec[["ID", "Sector_Name"]].rename(columns={"ID": "ID_Sector"}),
            on="ID_Sector", how="left")
S = S.merge(typ[["ID", "Name", "Record_Class"]]
            .rename(columns={"ID": "ID_Typology", "Name": "Typology"}),
            on="ID_Typology", how="left")

# Camp d'element -> (sistema que el condiciona, lletra del vocabulari)
GATE = {
    "Embedded_Base_Beams": ("Sys_Base", "A"),
    "Base_Level": ("Sys_Base", "B"),
    "Decorative_Socle": ("Sys_Base", "C"),
    "Tie_Walls": (None, "D"),
    "Timber_Brackets": ("Sys_Platform", "E"),
    "Transverse_Beams": ("Sys_Platform", "F"),
    "Corbelled_Courses": ("Sys_Platform", "G"),
    "Interbody_Cornice": ("Sys_Interface", "I"),
    "Corner_Quoins": ("Sys_Chamber", "J"),
    "Structural_Pilasters": ("Sys_Chamber", "K"),
    "Facade_Flank": ("Sys_Chamber", "L"),
    "Relief_Frieze": ("Sys_Chamber", "M"),
    "Sill": ("Sys_Portal", "N"),
    "Jambs": ("Sys_Portal", "O"),
    "Lintel": ("Sys_Portal", "Q"),
    "Upper_Crown": (None, "R"),
    "Eave_Beam": ("Sys_Eave", "S"),
    "Eave_Surface": ("Sys_Eave", "T"),
    "Return_Wall": ("Sys_Chamber", "V"),
    "Chamber_Roof": ("Sys_Chamber", "X"),
    "Platform_Surface": ("Sys_Platform", "Z"),
}

M = pd.DataFrame({
    "Code": S["Code"], "Sector": S["Sector_Name"],
    "Typology": S["Typology"], "Record_Class": S["Record_Class"],
})
for camp, (gate, lletra) in GATE.items():
    v = S[camp].copy()
    if gate:
        v = v.mask(S[gate].isin(["Not applicable", "Not observable"]))
    if lletra == "X":
        v = v.mask(S["Chamber_Roof_Type"] == "Natural bedrock")
    M[lletra] = v

lletres = sorted(l for _, l in GATE.values())
M = M[["Code", "Sector", "Typology", "Record_Class"] + lletres]
M.to_csv("data/matriu_AX_cinc_valors.csv", index=False)
print("data/matriu_AX_cinc_valors.csv:", M.shape[0], "files x",
      len(lletres), "elements")
