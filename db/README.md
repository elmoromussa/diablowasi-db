# Base de dades

`chachapoya_v26.accdb` és la base de dades Access tancada per a l'anàlisi, amb les dades: 106 elements arqueològics, 91 arestes de connexió, 12 paternitats, 7 datacions i 40 evidències de pèrdua, amb la bateria de validació `QRY_16_Validation_Check` a zero. Inclou les coordenades de les estructures (WGS84 i UTM 18S) i l'azimut de façana importats des de Metashape.

És l'única còpia Access del repositori. Es va construir seguint la cadena que descriu `docs/deltas/DELTA_v25g_v26.md`: pedaç de dades sobre la còpia v25, `BuildDB()` sobre una base en blanc, migració de les dades i `BuildForm()`. Les bases de fites anteriors es regeneren amb el codi de cada etiqueta.

No executes mai `BuildDB()` sobre aquesta base: esborra les taules de dades. Per a treballar-hi, fes-ne una còpia.
