#!/usr/bin/env bash
# Extrau les taules necessaries de la BD publicada (db/chachapoya_v26.accdb)
# cap a analysis/data/ mitjancant mdbtools. Executar des d'analysis/.
# Nota: la BD publica te les coordenades redactades; les extraccions locals
# T_*.csv i L_*.csv no es versionen (vegeu data/.gitignore).
set -e
DB="${1:-../db/chachapoya_v26.accdb}"
for t in T_STRUCTURES T_CONNECTIONS T_DECORATIONS T_LOST_ELEMENTS \
         L_SECTORS L_SITES L_TYPOLOGY L_ELEMENTS L_STRUCT_BODY L_DEC_TYPE \
         L_SUPPORT L_STATUS; do
  mdb-export "$DB" "$t" > "data/$t.csv"
  echo "  data/$t.csv"
done
echo "Extraccio completada."
