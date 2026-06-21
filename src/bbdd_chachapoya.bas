Option Compare Database
Option Explicit

' =============================================================
'  BASE DE DADES ARQUEOLOGICA - LA PETACA & DIABLO WASI
'  Autor: Esteve Ribera Torre  |  TFM Arqueologia UA 2024-25
'  Versio sense continuacions de linia (& _)
' =============================================================
'  INSTRUCCIONS:
'  1. Alt+F11  -> Inserir -> Modul
'  2. Enganxar tot el codi
'  3. Cursor dins Sub RunAll() -> F5
' =============================================================

Sub RunAll()
    Dim db As DAO.Database
    Set db = CurrentDb()
    CreateAllTables db
    PopulateAllLookups db
    CreateAllRelationships db
    CreateAllQueries db
    db.TableDefs.Refresh
    db.QueryDefs.Refresh
    Set db = Nothing
    Dim msg As String
    msg = "Base de dades creada correctament!" & vbCrLf & vbCrLf
    msg = msg & "  13 taules" & vbCrLf
    msg = msg & "  Lookups poblats" & vbCrLf
    msg = msg & "  13 relacions" & vbCrLf
    msg = msg & "  10 consultes" & vbCrLf & vbCrLf
    msg = msg & "Comprova el Panell de navegacio."
    MsgBox msg, vbInformation, "Fet!"
End Sub

' ─────────────────────────────────────────────────────────────
Function TableExists(db As DAO.Database, n As String) As Boolean
    Dim t As DAO.TableDef
    For Each t In db.TableDefs
        If t.Name = n Then TableExists = True: Exit Function
    Next t
End Function

Function QueryExists(db As DAO.Database, n As String) As Boolean
    Dim q As DAO.QueryDef
    For Each q In db.QueryDefs
        If q.Name = n Then QueryExists = True: Exit Function
    Next q
End Function

Sub X(db As DAO.Database, sql As String)
    On Error GoTo Err_X
    db.Execute sql, dbFailOnError
    Exit Sub
Err_X:
    Debug.Print "AVIS: " & Err.Description
End Sub

' ─────────────────────────────────────────────────────────────
'  1. TAULES
' ─────────────────────────────────────────────────────────────
Sub CreateAllTables(db As DAO.Database)
    Dim sql As String

    If Not TableExists(db, "L_SITIOS") Then
        db.Execute "CREATE TABLE L_SITIOS (ID COUNTER CONSTRAINT PK_SIT PRIMARY KEY, Nom_Sitio TEXT(50) NOT NULL, Descripcio MEMO)", dbFailOnError
    End If

    If Not TableExists(db, "L_SECTORES") Then
        db.Execute "CREATE TABLE L_SECTORES (ID COUNTER CONSTRAINT PK_SEC PRIMARY KEY, ID_Sitio LONG NOT NULL, Nom_Sector TEXT(50) NOT NULL, Descripcio TEXT(255))", dbFailOnError
    End If

    If Not TableExists(db, "L_TIPOLOGIA") Then
        db.Execute "CREATE TABLE L_TIPOLOGIA (ID COUNTER CONSTRAINT PK_TIP PRIMARY KEY, Nom TEXT(60) NOT NULL, Descripcio MEMO)", dbFailOnError
    End If

    If Not TableExists(db, "L_SUPORT") Then
        db.Execute "CREATE TABLE L_SUPORT (ID COUNTER CONSTRAINT PK_SUP PRIMARY KEY, Nom TEXT(50) NOT NULL)", dbFailOnError
    End If

    If Not TableExists(db, "L_ESTAT") Then
        db.Execute "CREATE TABLE L_ESTAT (ID COUNTER CONSTRAINT PK_EST PRIMARY KEY, Nom TEXT(30) NOT NULL)", dbFailOnError
    End If

    If Not TableExists(db, "L_METODE_VOLUM") Then
        db.Execute "CREATE TABLE L_METODE_VOLUM (ID COUNTER CONSTRAINT PK_MV PRIMARY KEY, Nom TEXT(50) NOT NULL, Descripcio TEXT(255))", dbFailOnError
    End If

    If Not TableExists(db, "L_COORD_METODE") Then
        db.Execute "CREATE TABLE L_COORD_METODE (ID COUNTER CONSTRAINT PK_CM PRIMARY KEY, Nom TEXT(50) NOT NULL, Descripcio TEXT(255))", dbFailOnError
    End If

    If Not TableExists(db, "L_TIPUS_CONJUNT") Then
        db.Execute "CREATE TABLE L_TIPUS_CONJUNT (ID COUNTER CONSTRAINT PK_TC PRIMARY KEY, Nom TEXT(50) NOT NULL, Descripcio MEMO)", dbFailOnError
    End If

    If Not TableExists(db, "L_CAMPANYA") Then
        db.Execute "CREATE TABLE L_CAMPANYA (ID COUNTER CONSTRAINT PK_CAM PRIMARY KEY, Codi TEXT(4) NOT NULL, Nom_Camp TEXT(80), Descripcio MEMO)", dbFailOnError
    End If

    If Not TableExists(db, "T_CONJUNTS") Then
        db.Execute "CREATE TABLE T_CONJUNTS (ID COUNTER CONSTRAINT PK_CNJ PRIMARY KEY, Codi_Conjunt TEXT(20) NOT NULL, ID_Sector LONG NOT NULL, ID_Tipus_Conjunt LONG NOT NULL, N_Membres INTEGER, Notes MEMO)", dbFailOnError
    End If

    ' T_ESTRUCTURES: construim el SQL per parts per evitar limit de linies
    If Not TableExists(db, "T_ESTRUCTURES") Then
        sql = "CREATE TABLE T_ESTRUCTURES ("
        sql = sql & "ID COUNTER CONSTRAINT PK_EST PRIMARY KEY,"
        sql = sql & "Codi TEXT(20) NOT NULL,"
        sql = sql & "ID_Sector LONG NOT NULL,"
        sql = sql & "ID_Tipologia LONG,"
        sql = sql & "ID_Suport LONG,"
        sql = sql & "ID_Estructura_Parent LONG,"
        sql = sql & "ID_Conjunt LONG,"
        sql = sql & "N_Pisos INTEGER,"
        sql = sql & "Planta TEXT(20),"
        sql = sql & "N_Murs_Construits INTEGER,"
        sql = sql & "Largo_m SINGLE,"
        sql = sql & "Ancho_m SINGLE,"
        sql = sql & "Alto_m SINGLE,"
        sql = sql & "Altura_Aprox_m SINGLE,"
        sql = sql & "Orientacio_Vano TEXT(5),"
        sql = sql & "Dintel TEXT(20),"
        sql = sql & "Techo_Natural YESNO,"
        sql = sql & "Contraforts YESNO,"
        sql = sql & "Cornisa_Entre_Pisos YESNO,"
        sql = sql & "Estacas_Fusta YESNO,"
        sql = sql & "Arrebossat YESNO,"
        sql = sql & "Color_Arrebossat TEXT(20),"
        sql = sql & "Pintura_Sobre_Roca YESNO,"
        sql = sql & "Color_Pintura_Roca TEXT(20),"
        sql = sql & "Dec_Nicho_Quadrat YESNO,"
        sql = sql & "Dec_Relieve_T YESNO,"
        sql = sql & "Dec_Relieve_T_Inv YESNO,"
        sql = sql & "Dec_Relieve_L YESNO,"
        sql = sql & "Dec_Relieve_L_Inv YESNO,"
        sql = sql & "Dec_Zigzag YESNO,"
        sql = sql & "Dec_Escalonat YESNO,"
        sql = sql & "Dec_Fris_Greca YESNO,"
        sql = sql & "Pintura_Rupestre YESNO,"
        sql = sql & "PR_Antropomorfa YESNO,"
        sql = sql & "PR_Zoomorfa YESNO,"
        sql = sql & "PR_Geometrica YESNO,"
        sql = sql & "PR_Abstracta YESNO,"
        sql = sql & "PR_Escena_Decap YESNO,"
        sql = sql & "ID_Estat LONG,"
        sql = sql & "Saqueig YESNO,"
        sql = sql & "Incendi YESNO,"
        sql = sql & "Activitat_Animal YESNO,"
        sql = sql & "Evidencia_Acces_Modern YESNO,"
        sql = sql & "Restes_Humanes YESNO,"
        sql = sql & "NMI INTEGER,"
        sql = sql & "Connexio_Anatomica YESNO,"
        sql = sql & "Momificacio YESNO,"
        sql = sql & "Fardells_Funeraris YESNO,"
        sql = sql & "Restes_Disperses YESNO,"
        sql = sql & "Posicio_Flexionada YESNO,"
        sql = sql & "Cremacio_Ossos YESNO,"
        sql = sql & "Mat_Textils YESNO,"
        sql = sql & "Mat_Fusta_Cultural YESNO,"
        sql = sql & "Mat_Fibra_Vegetal YESNO,"
        sql = sql & "Mat_Ceramica YESNO,"
        sql = sql & "Mat_Fauna YESNO,"
        sql = sql & "Mat_Banya_Cervol YESNO,"
        sql = sql & "Mat_Altres YESNO,"
        sql = sql & "C14 YESNO,"
        sql = sql & "Crono_Segle_Ini INTEGER,"
        sql = sql & "Crono_Segle_Fi INTEGER,"
        sql = sql & "Volum_Interior_m3 SINGLE,"
        sql = sql & "Volum_Total_m3 SINGLE,"
        sql = sql & "ID_Metode_Volum LONG,"
        sql = sql & "Volum_Notes TEXT(200),"
        sql = sql & "Coord_Lat_WGS84 DOUBLE,"
        sql = sql & "Coord_Lon_WGS84 DOUBLE,"
        sql = sql & "Coord_E_UTM DOUBLE,"
        sql = sql & "Coord_N_UTM DOUBLE,"
        sql = sql & "Altitud_msnm SINGLE,"
        sql = sql & "Coord_Precisio_m SINGLE,"
        sql = sql & "ID_Coord_Metode LONG,"
        sql = sql & "URL_Pano TEXT(255),"
        sql = sql & "URL_Pano_2 TEXT(255),"
        sql = sql & "URL_Giga TEXT(255),"
        sql = sql & "URL_3D TEXT(255),"
        sql = sql & "Documentat_ChaXR YESNO,"
        sql = sql & "ID_Campanya LONG,"
        sql = sql & "Notes MEMO)"
        db.Execute sql, dbFailOnError
        Debug.Print "[OK] T_ESTRUCTURES (65 camps)"
    End If

    If Not TableExists(db, "T_DATACIONS") Then
        db.Execute "CREATE TABLE T_DATACIONS (ID COUNTER CONSTRAINT PK_DAT PRIMARY KEY, ID_Estructura LONG NOT NULL, Mostra_Tipus TEXT(30), Data_BP LONG, Sigma1_Ini INTEGER, Sigma1_Fi INTEGER, Sigma2_Ini INTEGER, Sigma2_Fi INTEGER, Ref_Laboratori TEXT(50), Ref_Bibliografica MEMO)", dbFailOnError
    End If

    If Not TableExists(db, "T_INDIVIDUS") Then
        db.Execute "CREATE TABLE T_INDIVIDUS (ID COUNTER CONSTRAINT PK_IND PRIMARY KEY, ID_Estructura LONG NOT NULL, Num_Individu INTEGER, Edat_Categoria TEXT(20), Sexe_Categoria TEXT(20), Estat_Preservacio TEXT(20), Notes MEMO)", dbFailOnError
    End If

    Debug.Print "→ 13 taules creades"
End Sub

' ─────────────────────────────────────────────────────────────
'  2. POBLAR LOOKUPS
' ─────────────────────────────────────────────────────────────
Sub PopulateAllLookups(db As DAO.Database)
    Dim i As Integer

    X db, "DELETE FROM L_SECTORES"
    X db, "DELETE FROM L_SITIOS"
    X db, "DELETE FROM L_TIPOLOGIA"
    X db, "DELETE FROM L_SUPORT"
    X db, "DELETE FROM L_ESTAT"
    X db, "DELETE FROM L_METODE_VOLUM"
    X db, "DELETE FROM L_COORD_METODE"
    X db, "DELETE FROM L_TIPUS_CONJUNT"
    X db, "DELETE FROM L_CAMPANYA"

    ' L_SITIOS
    db.Execute "INSERT INTO L_SITIOS (Nom_Sitio,Descripcio) VALUES ('La Petaca','>12.000 m2, 4 sectors, s.X-XVI. Lat -6.8311 / Lon -77.8084')", dbFailOnError
    db.Execute "INSERT INTO L_SITIOS (Nom_Sitio,Descripcio) VALUES ('Diablo Wasi','6 sectors, predomini cavitats. Lat -6.8475 / Lon -77.8154')", dbFailOnError

    ' L_SECTORES
    Dim s(11, 2) As String
    s(0, 0) = "1": s(0, 1) = "LP - General":  s(0, 2) = "La Petaca - context general"
    s(1, 0) = "1": s(1, 1) = "LP - Nord":     s(1, 2) = "La Petaca - Sector Nord"
    s(2, 0) = "1": s(2, 1) = "LP - Central":  s(2, 2) = "La Petaca - Sector Central"
    s(3, 0) = "1": s(3, 1) = "LP - Superior": s(3, 2) = "La Petaca - Sector Superior"
    s(4, 0) = "1": s(4, 1) = "LP - Sud":      s(4, 2) = "La Petaca - Sector Sud (>=96 estructures)"
    s(5, 0) = "2": s(5, 1) = "DW - General":  s(5, 2) = "Diablo Wasi - context general"
    s(6, 0) = "2": s(6, 1) = "DW - Sector 1": s(6, 2) = "Diablo Wasi - Sector 1 (~40 contextos)"
    s(7, 0) = "2": s(7, 1) = "DW - Sector 2": s(7, 2) = "Diablo Wasi - Sector 2 (1 cambra + cova)"
    s(8, 0) = "2": s(8, 1) = "DW - Sector 3": s(8, 2) = "Diablo Wasi - Sector 3 (cova subterranea)"
    s(9, 0) = "2": s(9, 1) = "DW - Sector 4": s(9, 2) = "Diablo Wasi - Sector 4 (~9 contextos)"
    s(10, 0) = "2": s(10, 1) = "DW - Sector 5": s(10, 2) = "Diablo Wasi - Sector 5"
    s(11, 0) = "2": s(11, 1) = "DW - Sector 6": s(11, 2) = "Diablo Wasi - Sector 6"
    For i = 0 To 11
        db.Execute "INSERT INTO L_SECTORES (ID_Sitio,Nom_Sector,Descripcio) VALUES (" & s(i, 0) & ",'" & s(i, 1) & "','" & s(i, 2) & "')", dbFailOnError
    Next i

    ' L_TIPOLOGIA
    Dim t(9, 1) As String
    t(0, 0) = "EA-MAU Mausoleu/Chullpa":   t(0, 1) = "Estructura construida (3+ murs + sostre) sobre repisa. 1-3 pisos. Predominant a LP."
    t(1, 0) = "EA-CAM Cambra funeraria":   t(1, 1) = "Cavitat natural tancada per 1 facana construida. Predominant a DW."
    t(2, 0) = "EA-PLA-R Plataforma repisa":t(2, 1) = "Plataforma sobre repisa natural. Funcio: transit o base per a mausoleus."
    t(3, 0) = "EA-PLA-V Plataforma volada":t(3, 1) = "Plataforma artificial sobre fustes i lloses, sense repisa natural."
    t(4, 0) = "NIX Nixol natural":         t(4, 1) = "Petita cavitat natural (<1m2). Ossari o enterrament secundari."
    t(5, 0) = "CAV Caverna/Cova":          t(5, 1) = "Gran cavitat natural (>1m2) amb us funerari o ritual documentat."
    t(6, 0) = "PR Pintura rupestre":       t(6, 1) = "Motiu pictoric sobre roca, documentat de forma independent."
    t(7, 0) = "MEN Mensula aillada":       t(7, 1) = "Element estructural aillat. Evidencia de circulacio perduda."
    t(8, 0) = "MIX Mixt":                 t(8, 1) = "Combinacio de dues o mes categories anteriors."
    t(9, 0) = "ND No determinat":         t(9, 1) = "Informacio insuficient per classificar."
    For i = 0 To 9
        db.Execute "INSERT INTO L_TIPOLOGIA (Nom,Descripcio) VALUES ('" & t(i, 0) & "','" & t(i, 1) & "')", dbFailOnError
    Next i

    ' L_SUPORT
    Dim sup(8) As String
    sup(0) = "Repisa natural amplia (>2m)"
    sup(1) = "Repisa natural estreta (<2m)"
    sup(2) = "Repisa artificial"
    sup(3) = "Cavitat gran (>10m2)"
    sup(4) = "Cavitat mitjana (1-10m2)"
    sup(5) = "Nixol natural (<1m2)"
    sup(6) = "Grieta"
    sup(7) = "Combinat"
    sup(8) = "ND"
    For i = 0 To 8
        db.Execute "INSERT INTO L_SUPORT (Nom) VALUES ('" & sup(i) & "')", dbFailOnError
    Next i

    ' L_ESTAT
    Dim e(4) As String
    e(0) = "Bo": e(1) = "Regular": e(2) = "Pre-col.lapse": e(3) = "Col.lapsat": e(4) = "ND"
    For i = 0 To 4
        db.Execute "INSERT INTO L_ESTAT (Nom) VALUES ('" & e(i) & "')", dbFailOnError
    Next i

    ' L_METODE_VOLUM
    db.Execute "INSERT INTO L_METODE_VOLUM (Nom,Descripcio) VALUES ('Calcul LxAxH','Calcul geometric per a cambres rectangulars.')", dbFailOnError
    db.Execute "INSERT INTO L_METODE_VOLUM (Nom,Descripcio) VALUES ('Model fotogrametric','Volum del model 3D (MeshLab, CloudCompare).')", dbFailOnError
    db.Execute "INSERT INTO L_METODE_VOLUM (Nom,Descripcio) VALUES ('Estimacio','Estimacio visual o mesures parcials.')", dbFailOnError
    db.Execute "INSERT INTO L_METODE_VOLUM (Nom,Descripcio) VALUES ('ND','No determinat.')", dbFailOnError

    ' L_COORD_METODE
    db.Execute "INSERT INTO L_COORD_METODE (Nom,Descripcio) VALUES ('Drone RTK','Precisio ~2-5 cm.')", dbFailOnError
    db.Execute "INSERT INTO L_COORD_METODE (Nom,Descripcio) VALUES ('GPS diferencial','Precisio ~5-20 cm.')", dbFailOnError
    db.Execute "INSERT INTO L_COORD_METODE (Nom,Descripcio) VALUES ('Fotogrametria','Model Metashape georeferenciat.')", dbFailOnError
    db.Execute "INSERT INTO L_COORD_METODE (Nom,Descripcio) VALUES ('GPS mobil','Precisio 3-10 m. Nomes orientacio.')", dbFailOnError
    db.Execute "INSERT INTO L_COORD_METODE (Nom,Descripcio) VALUES ('Estimacio','Estimacio sobre ortofoto. >10 m.')", dbFailOnError
    db.Execute "INSERT INTO L_COORD_METODE (Nom,Descripcio) VALUES ('ND','No determinat.')", dbFailOnError

    ' L_TIPUS_CONJUNT
    db.Execute "INSERT INTO L_TIPUS_CONJUNT (Nom,Descripcio) VALUES ('Alineament vertical','Estructures en una mateixa vertical (grieta o estrats successius).')", dbFailOnError
    db.Execute "INSERT INTO L_TIPUS_CONJUNT (Nom,Descripcio) VALUES ('Conjunt de repisa','Multiples EAs en una mateixa repisa horitzontal.')", dbFailOnError
    db.Execute "INSERT INTO L_TIPUS_CONJUNT (Nom,Descripcio) VALUES ('Plataforma amb mensules','Repisa volada + mensules que la composen.')", dbFailOnError
    db.Execute "INSERT INTO L_TIPUS_CONJUNT (Nom,Descripcio) VALUES ('Conjunt de cova','Caverna + estructures construides al seu interior.')", dbFailOnError
    db.Execute "INSERT INTO L_TIPUS_CONJUNT (Nom,Descripcio) VALUES ('Xarxa de circulacio','MENs i EA-PLAs que reconstrueixen circulacio aeria perduda.')", dbFailOnError
    db.Execute "INSERT INTO L_TIPUS_CONJUNT (Nom,Descripcio) VALUES ('Conjunt pintural','PR + EA adjacent visualment o funcionalment vinculada.')", dbFailOnError
    db.Execute "INSERT INTO L_TIPUS_CONJUNT (Nom,Descripcio) VALUES ('Grup funcional','Qualsevol altra agrupacio intra-sector.')", dbFailOnError

    ' L_CAMPANYA
    db.Execute "INSERT INTO L_CAMPANYA (Codi,Nom_Camp,Descripcio) VALUES ('2013','PALP I','Prospeccio inicial La Petaca. Dr. J. Marla Toyne (UCF).')", dbFailOnError
    db.Execute "INSERT INTO L_CAMPANYA (Codi,Nom_Camp,Descripcio) VALUES ('2016','PALP II','Ampliacio a Diablo Wasi. Primera doc. sistematica DW.')", dbFailOnError
    db.Execute "INSERT INTO L_CAMPANYA (Codi,Nom_Camp,Descripcio) VALUES ('2021','La Petaca Project','Doc. no invasiva integral. Fotogrametria+360+gigafotos.')", dbFailOnError
    db.Execute "INSERT INTO L_CAMPANYA (Codi,Nom_Camp,Descripcio) VALUES ('2023','PALP IV','Excavacio + reconstruccions 3D detallades.')", dbFailOnError

    Debug.Print "→ Lookups poblats"
End Sub

' ─────────────────────────────────────────────────────────────
'  3. RELACIONS
' ─────────────────────────────────────────────────────────────
Sub CreateAllRelationships(db As DAO.Database)
    Dim rn(12) As String
    rn(0) = "REL_SIT_SEC": rn(1) = "REL_SEC_EST": rn(2) = "REL_SEC_CNJ"
    rn(3) = "REL_TIP_EST": rn(4) = "REL_SUP_EST": rn(5) = "REL_EST_EST2"
    rn(6) = "REL_STA_EST": rn(7) = "REL_MV_EST":  rn(8) = "REL_CM_EST"
    rn(9) = "REL_TC_CNJ":  rn(10) = "REL_CNJ_EST": rn(11) = "REL_EST_DAT"
    rn(12) = "REL_EST_IND"
    Dim n As Integer
    For n = 0 To 12
        On Error Resume Next
        db.Relations.Delete rn(n)
        On Error GoTo 0
    Next n

    MkRel db, rn(0),  "L_SITIOS",       "ID", "L_SECTORES",    "ID_Sitio",             True,  False
    MkRel db, rn(1),  "L_SECTORES",     "ID", "T_ESTRUCTURES", "ID_Sector",            True,  False
    MkRel db, rn(2),  "L_SECTORES",     "ID", "T_CONJUNTS",    "ID_Sector",            True,  False
    MkRel db, rn(3),  "L_TIPOLOGIA",    "ID", "T_ESTRUCTURES", "ID_Tipologia",         False, False
    MkRel db, rn(4),  "L_SUPORT",       "ID", "T_ESTRUCTURES", "ID_Suport",            False, False
    MkRel db, rn(5),  "T_ESTRUCTURES",  "ID", "T_ESTRUCTURES", "ID_Estructura_Parent", False, True
    MkRel db, rn(6),  "L_ESTAT",        "ID", "T_ESTRUCTURES", "ID_Estat",             False, False
    MkRel db, rn(7),  "L_METODE_VOLUM", "ID", "T_ESTRUCTURES", "ID_Metode_Volum",      False, False
    MkRel db, rn(8),  "L_COORD_METODE", "ID", "T_ESTRUCTURES", "ID_Coord_Metode",      False, False
    MkRel db, rn(9),  "L_TIPUS_CONJUNT","ID", "T_CONJUNTS",    "ID_Tipus_Conjunt",     False, False
    MkRel db, rn(10), "T_CONJUNTS",     "ID", "T_ESTRUCTURES", "ID_Conjunt",           False, False
    MkRel db, rn(11), "T_ESTRUCTURES",  "ID", "T_DATACIONS",   "ID_Estructura",        True,  False
    MkRel db, rn(12), "T_ESTRUCTURES",  "ID", "T_INDIVIDUS",   "ID_Estructura",        True,  False

    db.Relations.Refresh
    Debug.Print "→ 13 relacions OK"
End Sub

Sub MkRel(db As DAO.Database, nm As String, pTbl As String, pFld As String, cTbl As String, cFld As String, del As Boolean, noInt As Boolean)
    Dim rel As DAO.Relation
    Dim fld As DAO.Field
    Dim fl As Long
    On Error GoTo ErrR
    If noInt Then
        fl = 2
    Else
        fl = dbRelationUpdateCascade
        If del Then fl = fl Or dbRelationDeleteCascade
    End If
    Set rel = db.CreateRelation(nm, pTbl, cTbl, fl)
    Set fld = rel.CreateField(pFld)
    fld.ForeignName = cFld
    rel.Fields.Append fld
    db.Relations.Append rel
    Exit Sub
ErrR:
    Debug.Print "  Avis relacio " & nm & ": " & Err.Description
End Sub

' ─────────────────────────────────────────────────────────────
'  4. CONSULTES
' ─────────────────────────────────────────────────────────────
Sub CreateAllQueries(db As DAO.Database)
    Dim qn(9) As String
    qn(0) = "QRY_01_Tipologia_x_Jaciment"
    qn(1) = "QRY_02_Decoracio_x_Jaciment"
    qn(2) = "QRY_03_Conservacio_x_Sector"
    qn(3) = "QRY_04_Estructures_C14"
    qn(4) = "QRY_05_Export_RSPSS"
    qn(5) = "QRY_06_Volumetria_x_Tipologia"
    qn(6) = "QRY_07_Export_QGIS"
    qn(7) = "QRY_08_Fills_dun_Pare"
    qn(8) = "QRY_09_Membres_Conjunt"
    qn(9) = "QRY_10_Cobertura_ChaXR"
    Dim i As Integer
    For i = 0 To 9
        If QueryExists(db, qn(i)) Then db.QueryDefs.Delete qn(i)
    Next i

    Dim q As String

    ' QRY_01
    q = "SELECT S.Nom_Sitio, T.Nom AS Tipologia, COUNT(E.ID) AS N "
    q = q & "FROM ((T_ESTRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORES AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITIOS AS S ON SC.ID_Sitio=S.ID) "
    q = q & "INNER JOIN L_TIPOLOGIA AS T ON E.ID_Tipologia=T.ID "
    q = q & "GROUP BY S.Nom_Sitio, T.Nom ORDER BY S.Nom_Sitio, T.Nom;"
    db.CreateQueryDef qn(0), q

    ' QRY_02
    q = "SELECT S.Nom_Sitio, "
    q = q & "SUM(IIF(E.Dec_Nicho_Quadrat=True,1,0)) AS N_Nicho, "
    q = q & "SUM(IIF(E.Dec_Relieve_T=True,1,0)) AS N_T, "
    q = q & "SUM(IIF(E.Dec_Relieve_L=True,1,0)) AS N_L, "
    q = q & "SUM(IIF(E.Dec_Zigzag=True,1,0)) AS N_Zigzag, "
    q = q & "SUM(IIF(E.Dec_Escalonat=True,1,0)) AS N_Escalonat, "
    q = q & "SUM(IIF(E.Pintura_Rupestre=True,1,0)) AS N_PR, "
    q = q & "COUNT(E.ID) AS Total "
    q = q & "FROM (T_ESTRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORES AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITIOS AS S ON SC.ID_Sitio=S.ID "
    q = q & "GROUP BY S.Nom_Sitio;"
    db.CreateQueryDef qn(1), q

    ' QRY_03
    q = "SELECT S.Nom_Sitio, SC.Nom_Sector, ES.Nom AS Estat, COUNT(*) AS N "
    q = q & "FROM (((T_ESTRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORES AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITIOS AS S ON SC.ID_Sitio=S.ID) "
    q = q & "INNER JOIN L_ESTAT AS ES ON E.ID_Estat=ES.ID) "
    q = q & "GROUP BY S.Nom_Sitio, SC.Nom_Sector, ES.Nom "
    q = q & "ORDER BY S.Nom_Sitio, SC.Nom_Sector;"
    db.CreateQueryDef qn(2), q

    ' QRY_04
    q = "SELECT E.Codi, S.Nom_Sitio, SC.Nom_Sector, T.Nom AS Tipologia, "
    q = q & "E.Crono_Segle_Ini, E.Crono_Segle_Fi "
    q = q & "FROM (((T_ESTRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORES AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITIOS AS S ON SC.ID_Sitio=S.ID) "
    q = q & "INNER JOIN L_TIPOLOGIA AS T ON E.ID_Tipologia=T.ID) "
    q = q & "WHERE E.C14=True ORDER BY E.Crono_Segle_Ini;"
    db.CreateQueryDef qn(3), q

    ' QRY_05 export pla
    q = "SELECT E.ID, E.Codi, S.Nom_Sitio AS Jaciment, SC.Nom_Sector AS Sector, "
    q = q & "T.Nom AS Tipologia, SU.Nom AS Suport, ES.Nom AS Estat, "
    q = q & "E.N_Pisos, E.Planta, E.N_Murs_Construits, E.Orientacio_Vano, "
    q = q & "E.Techo_Natural, E.Contraforts, E.Arrebossat, E.Pintura_Sobre_Roca, "
    q = q & "E.Dec_Nicho_Quadrat, E.Dec_Relieve_T, E.Dec_Relieve_T_Inv, "
    q = q & "E.Dec_Relieve_L, E.Dec_Relieve_L_Inv, E.Dec_Zigzag, "
    q = q & "E.Dec_Escalonat, E.Dec_Fris_Greca, E.Pintura_Rupestre, "
    q = q & "E.Saqueig, E.Incendi, E.Activitat_Animal, E.Evidencia_Acces_Modern, "
    q = q & "E.Restes_Humanes, E.NMI, E.Momificacio, E.Fardells_Funeraris, "
    q = q & "E.Cremacio_Ossos, E.Mat_Textils, E.Mat_Ceramica, E.Mat_Banya_Cervol, "
    q = q & "E.C14, E.Crono_Segle_Ini, E.Crono_Segle_Fi, "
    q = q & "E.Volum_Interior_m3, E.Volum_Total_m3, "
    q = q & "E.Coord_Lat_WGS84, E.Coord_Lon_WGS84, "
    q = q & "E.Coord_E_UTM, E.Coord_N_UTM, E.Altitud_msnm, "
    q = q & "E.Documentat_ChaXR "
    q = q & "FROM ((((T_ESTRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORES AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITIOS AS S ON SC.ID_Sitio=S.ID) "
    q = q & "INNER JOIN L_TIPOLOGIA AS T ON E.ID_Tipologia=T.ID) "
    q = q & "INNER JOIN L_SUPORT AS SU ON E.ID_Suport=SU.ID) "
    q = q & "INNER JOIN L_ESTAT AS ES ON E.ID_Estat=ES.ID;"
    db.CreateQueryDef qn(4), q

    ' QRY_06
    q = "SELECT S.Nom_Sitio, T.Nom AS Tipologia, "
    q = q & "COUNT(E.ID) AS N_Total, COUNT(E.Volum_Interior_m3) AS N_Volum, "
    q = q & "AVG(E.Volum_Interior_m3) AS Volum_Mitja, "
    q = q & "MIN(E.Volum_Interior_m3) AS Volum_Min, "
    q = q & "MAX(E.Volum_Interior_m3) AS Volum_Max "
    q = q & "FROM ((T_ESTRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORES AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITIOS AS S ON SC.ID_Sitio=S.ID) "
    q = q & "INNER JOIN L_TIPOLOGIA AS T ON E.ID_Tipologia=T.ID "
    q = q & "WHERE E.Volum_Interior_m3 IS NOT NULL "
    q = q & "GROUP BY S.Nom_Sitio, T.Nom ORDER BY S.Nom_Sitio, Volum_Mitja DESC;"
    db.CreateQueryDef qn(5), q

    ' QRY_07 QGIS
    q = "SELECT E.ID, E.Codi, S.Nom_Sitio, SC.Nom_Sector, T.Nom AS Tipologia, "
    q = q & "E.Coord_Lat_WGS84, E.Coord_Lon_WGS84, "
    q = q & "E.Coord_E_UTM, E.Coord_N_UTM, E.Altitud_msnm, "
    q = q & "E.Altura_Aprox_m, E.Coord_Precisio_m, ES.Nom AS Estat, "
    q = q & "E.N_Pisos, E.Volum_Interior_m3, "
    q = q & "E.Crono_Segle_Ini, E.Crono_Segle_Fi, "
    q = q & "E.Saqueig, E.Restes_Humanes, E.NMI, "
    q = q & "E.Documentat_ChaXR, E.URL_3D "
    q = q & "FROM (((T_ESTRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORES AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITIOS AS S ON SC.ID_Sitio=S.ID) "
    q = q & "INNER JOIN L_TIPOLOGIA AS T ON E.ID_Tipologia=T.ID) "
    q = q & "INNER JOIN L_ESTAT AS ES ON E.ID_Estat=ES.ID "
    q = q & "WHERE E.Coord_Lat_WGS84 IS NOT NULL ORDER BY S.Nom_Sitio, E.Codi;"
    db.CreateQueryDef qn(6), q

    ' QRY_08 fills d'un pare
    q = "SELECT E_f.Codi AS Codi_Fill, T.Nom AS Tipologia, E_f.Restes_Humanes, E_f.NMI "
    q = q & "FROM T_ESTRUCTURES AS E_f "
    q = q & "INNER JOIN L_TIPOLOGIA AS T ON E_f.ID_Tipologia=T.ID "
    q = q & "WHERE E_f.ID_Estructura_Parent=[ID del pare?] "
    q = q & "ORDER BY T.Nom, E_f.Codi;"
    db.CreateQueryDef qn(7), q

    ' QRY_09 membres conjunt
    q = "SELECT C.Codi_Conjunt, TC.Nom AS Tipus, E.Codi, T.Nom AS Tipologia, "
    q = q & "E.Coord_E_UTM, E.Coord_N_UTM, E.Altitud_msnm "
    q = q & "FROM (((T_ESTRUCTURES AS E "
    q = q & "INNER JOIN T_CONJUNTS AS C ON E.ID_Conjunt=C.ID) "
    q = q & "INNER JOIN L_TIPOLOGIA AS T ON E.ID_Tipologia=T.ID) "
    q = q & "INNER JOIN L_TIPUS_CONJUNT AS TC ON C.ID_Tipus_Conjunt=TC.ID) "
    q = q & "WHERE C.Codi_Conjunt=[Codi del conjunt?] "
    q = q & "ORDER BY E.Altitud_msnm DESC;"
    db.CreateQueryDef qn(8), q

    ' QRY_10 cobertura ChaXR
    q = "SELECT S.Nom_Sitio, CA.Codi AS Campanya, "
    q = q & "COUNT(E.ID) AS Total, "
    q = q & "SUM(IIF(E.Documentat_ChaXR=True,1,0)) AS Publicat_ChaXR, "
    q = q & "SUM(IIF(E.URL_3D IS NOT NULL,1,0)) AS Amb_Model_3D "
    q = q & "FROM ((T_ESTRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORES AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITIOS AS S ON SC.ID_Sitio=S.ID) "
    q = q & "LEFT JOIN L_CAMPANYA AS CA ON E.ID_Campanya=CA.ID "
    q = q & "GROUP BY S.Nom_Sitio, CA.Codi ORDER BY S.Nom_Sitio, CA.Codi;"
    db.CreateQueryDef qn(9), q

    Debug.Print "→ 10 consultes OK"
End Sub

' ─────────────────────────────────────────────────────────────
'  OPCIONAL: configurar ComboBox per als camps FK
'  Executar DESPRES de RunAll()
' ─────────────────────────────────────────────────────────────
Sub SetupCombos()
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim prp As DAO.Property
    Set db = CurrentDb()
    Set tdf = db.TableDefs("T_ESTRUCTURES")

    Dim c(7, 2) As String
    c(0, 0) = "ID_Sector":       c(0, 1) = "SELECT ID, Nom_Sector FROM L_SECTORES ORDER BY ID_Sitio, Nom_Sector":      c(0, 2) = "2"
    c(1, 0) = "ID_Tipologia":    c(1, 1) = "SELECT ID, Nom FROM L_TIPOLOGIA ORDER BY Nom":                              c(1, 2) = "2"
    c(2, 0) = "ID_Suport":       c(2, 1) = "SELECT ID, Nom FROM L_SUPORT ORDER BY ID":                                  c(2, 2) = "2"
    c(3, 0) = "ID_Estat":        c(3, 1) = "SELECT ID, Nom FROM L_ESTAT ORDER BY ID":                                   c(3, 2) = "2"
    c(4, 0) = "ID_Metode_Volum": c(4, 1) = "SELECT ID, Nom FROM L_METODE_VOLUM ORDER BY ID":                            c(4, 2) = "2"
    c(5, 0) = "ID_Coord_Metode": c(5, 1) = "SELECT ID, Nom FROM L_COORD_METODE ORDER BY ID":                            c(5, 2) = "2"
    c(6, 0) = "ID_Campanya":     c(6, 1) = "SELECT ID, Codi, Nom_Camp FROM L_CAMPANYA ORDER BY Codi":                   c(6, 2) = "3"
    c(7, 0) = "ID_Conjunt":      c(7, 1) = "SELECT ID, Codi_Conjunt FROM T_CONJUNTS ORDER BY Codi_Conjunt":             c(7, 2) = "2"

    Dim i As Integer
    For i = 0 To 7
        Set fld = tdf.Fields(c(i, 0))
        On Error Resume Next
        fld.Properties("DisplayControl") = 111
        If Err.Number <> 0 Then
            Set prp = fld.CreateProperty("DisplayControl", DB_INTEGER, 111)
            fld.Properties.Append prp
            Err.Clear
        End If
        fld.Properties("RowSourceType").Value = "Table/Query"
        fld.Properties("RowSource").Value = c(i, 1)
        fld.Properties("BoundColumn").Value = 1
        fld.Properties("ColumnCount").Value = CInt(c(i, 2))
        fld.Properties("ColumnWidths").Value = "0cm;5cm"
        fld.Properties("LimitToList").Value = True
        On Error GoTo 0
        Debug.Print "[OK] Combo: " & c(i, 0)
    Next i

    tdf.Fields.Refresh
    Set db = Nothing
    MsgBox "ComboBox configurats!", vbInformation
End Sub
