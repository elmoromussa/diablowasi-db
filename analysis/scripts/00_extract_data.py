#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
00_extract_data.py - Alternativa en Python a 00_extract_data.sh

Extrau les mateixes taules de la BD publicada (db/chachapoya_v26.accdb) cap
a analysis/data/ amb el paquet access-parser (pip install access-parser), que
llig el fitxer .accdb directament: no cal Access, ni el controlador ODBC/ACE,
ni mdbtools. Per als camps que consumeixen els scripts (numerics i de text)
el resultat es equivalent al de mdb-export. Executar des d'analysis/.
Les extraccions T_*.csv i L_*.csv no es versionen (vegeu data/.gitignore).
"""
import sys
import pandas as pd
from access_parser import AccessParser

DB = sys.argv[1] if len(sys.argv) > 1 else "../db/chachapoya_v26.accdb"
TAULES = ["T_STRUCTURES", "T_CONNECTIONS", "T_DECORATIONS", "T_LOST_ELEMENTS",
          "L_SECTORS", "L_SITES", "L_TYPOLOGY", "L_ELEMENTS", "L_STRUCT_BODY",
          "L_DEC_TYPE", "L_SUPPORT", "L_STATUS"]

db = AccessParser(DB)
for t in TAULES:
    df = pd.DataFrame(db.parse_table(t))
    df.to_csv(f"data/{t}.csv", index=False, encoding="utf-8")
    print(f"  data/{t}.csv ({len(df)} files)")
print("Extraccio completada.")
