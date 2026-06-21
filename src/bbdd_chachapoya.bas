Attribute VB_Name = "Chachapoya_DB"
Option Compare Database
Option Explicit

' =============================================================
'  BASE DE DADES ARQUEOLOGICA - LA PETACA & DIABLO WASI
'  Autor: Esteve Ribera Torró  |  TFM Arqueologia UA 2024-25
' =============================================================
'  INSTRUCCIONS:
'  1. Obrir Access -> Alt+F11 (Editor VBA)
'  2. Inserir -> Modul
'  3. Enganxar tot aquest codi
'  4. Posar el cursor dins Sub RunAll() i prémer F5
'     (o clic en el botó "Executar" ▶)
' =============================================================

' ─────────────────────────────────────────────────────────────
'  PUNT D'ENTRADA PRINCIPAL
' ─────────────────────────────────────────────────────────────
Sub RunAll()
    Dim db As DAO.Database
    Set db = CurrentDb()

    Debug.Print "=== INICI CREACIÓ DB CHACHAPOYA ==="

    CreateAllTables db
    PopulateAllLookups db
    CreateAllRelationships db
    CreateAllQueries db

    db.TableDefs.Refresh
    db.QueryDefs.Refresh
    Set db = Nothing

    MsgBox "Base de dades creada correctament!" & vbCrLf & vbCrLf & _
           "  13 taules creades" & vbCrLf & _
           "  Taules lookup poblades" & vbCrLf & _
           "  13 relacions establertes" & vbCrLf & _
           "  10 consultes SQL creades" & vbCrLf & vbCrLf & _
           "Comprova el Panell de navegació." & vbCrLf & _
           "Pas següent: configura els ComboBox (veure instruccions).", _
           vbInformation, "Completat!"
End Sub

' ─────────────────────────────────────────────────────────────
'  HELPERS
' ─────────────────────────────────────────────────────────────
Function TableExists(db As DAO.Database, tName As String) As Boolean
    Dim t As DAO.TableDef
    For Each t In db.TableDefs
        If t.Name = tName Then TableExists = True: Exit Function
    Next t
End Function

Function QueryExists(db As DAO.Database, qName As String) As Boolean
    Dim q As DAO.QueryDef
    For Each q In db.QueryDefs
        If q.Name = qName Then QueryExists = True: Exit Function
    Next q
End Function

Sub SafeExec(db As DAO.Database, sql As String)
    On Error GoTo Err_SafeExec
    db.Execute sql, dbFailOnError
    Exit Sub
Err_SafeExec:
    Debug.Print "  AVIS: " & Err.Description & " | SQL: " & Left(sql, 60)
End Sub

' ─────────────────────────────────────────────────────────────
'  1. CREAR TAULES
' ─────────────────────────────────────────────────────────────
Sub CreateAllTables(db As DAO.Database)
    Debug.Print "Creant taules..."

    ' ── LOOKUPS ──────────────────────────────────────────────
    If Not TableExists(db, "L_SITIOS") Then
        db.Execute "CREATE TABLE L_SITIOS (" & _
            "ID COUNTER CONSTRAINT PK_SITIOS PRIMARY KEY, " & _
            "Nom_Sitio TEXT(50) NOT NULL, " & _
            "Descripcio MEMO)", dbFailOnError
        Debug.Print "  [OK] L_SITIOS"
    End If

    If Not TableExists(db, "L_SECTORES") Then
        db.Execute "CREATE TABLE L_SECTORES (" & _
            "ID COUNTER CONSTRAINT PK_SECTORES PRIMARY KEY, " & _
            "ID_Sitio LONG NOT NULL, " & _
            "Nom_Sector TEXT(50) NOT NULL, " & _
            "Descripcio TEXT(255))", dbFailOnError
        Debug.Print "  [OK] L_SECTORES"
    End If

    If Not TableExists(db, "L_TIPOLOGIA") Then
        db.Execute "CREATE TABLE L_TIPOLOGIA (" & _
            "ID COUNTER CONSTRAINT PK_TIPOLOGIA PRIMARY KEY, " & _
            "Nom TEXT(60) NOT NULL, " & _
            "Descripcio MEMO)", dbFailOnError
        Debug.Print "  [OK] L_TIPOLOGIA"
    End If

    If Not TableExists(db, "L_SUPORT") Then
        db.Execute "CREATE TABLE L_SUPORT (" & _
            "ID COUNTER CONSTRAINT PK_SUPORT PRIMARY KEY, " & _
            "Nom TEXT(50) NOT NULL)", dbFailOnError
        Debug.Print "  [OK] L_SUPORT"
    End If

    If Not TableExists(db, "L_ESTAT") Then
        db.Execute "CREATE TABLE L_ESTAT (" & _
            "ID COUNTER CONSTRAINT PK_ESTAT PRIMARY KEY, " & _
            "Nom TEXT(30) NOT NULL)", dbFailOnError
        Debug.Print "  [OK] L_ESTAT"
    End If

    If Not TableExists(db, "L_METODE_VOLUM") Then
        db.Execute "CREATE TABLE L_METODE_VOLUM (" & _
            "ID COUNTER CONSTRAINT PK_MET_VOL PRIMARY KEY, " & _
            "Nom TEXT(50) NOT NULL, " & _
            "Descripcio TEXT(255))", dbFailOnError
        Debug.Print "  [OK] L_METODE_VOLUM"
    End If

    If Not TableExists(db, "L_COORD_METODE") Then
        db.Execute "CREATE TABLE L_COORD_METODE (" & _
            "ID COUNTER CONSTRAINT PK_COORD_MET PRIMARY KEY, " & _
            "Nom TEXT(50) NOT NULL, " & _
            "Descripcio TEXT(255))", dbFailOnError
        Debug.Print "  [OK] L_COORD_METODE"
    End If

    If Not TableExists(db, "L_TIPUS_CONJUNT") Then
        db.Execute "CREATE TABLE L_TIPUS_CONJUNT (" & _
            "ID COUNTER CONSTRAINT PK_TIPUS_CNJ PRIMARY KEY, " & _
            "Nom TEXT(50) NOT NULL, " & _
            "Descripcio MEMO)", dbFailOnError
        Debug.Print "  [OK] L_TIPUS_CONJUNT"
    End If

    If Not TableExists(db, "L_CAMPANYA") Then
        db.Execute "CREATE TABLE L_CAMPANYA (" & _
            "ID COUNTER CONSTRAINT PK_CAMPANYA PRIMARY KEY, " & _
            "Codi TEXT(4) NOT NULL, " & _
            "Nom_Camp TEXT(80), " & _
            "Descripcio MEMO)", dbFailOnError
        Debug.Print "  [OK] L_CAMPANYA"
    End If

    ' ── TAULES SECUNDÀRIES ───────────────────────────────────
    If Not TableExists(db, "T_CONJUNTS") Then
        db.Execute "CREATE TABLE T_CONJUNTS (" & _
            "ID COUNTER CONSTRAINT PK_CONJUNTS PRIMARY KEY, " & _
            "Codi_Conjunt TEXT(20) NOT NULL, " & _
            "ID_Sector LONG NOT NULL, " & _
            "ID_Tipus_Conjunt LONG NOT NULL, " & _
            "N_Membres INTEGER, " & _
            "Notes MEMO)", dbFailOnError
        Debug.Print "  [OK] T_CONJUNTS"
    End If

    ' ── TAULA PRINCIPAL ──────────────────────────────────────
    If Not TableExists(db, "T_ESTRUCTURES") Then
        db.Execute _
            "CREATE TABLE T_ESTRUCTURES (" & _
            "ID COUNTER CONSTRAINT PK_ESTRUCTURES PRIMARY KEY, " & _
            "Codi TEXT(20) NOT NULL, " & _
            "ID_Sector LONG NOT NULL, " & _
            "ID_Tipologia LONG, " & _
            "ID_Suport LONG, " & _
            "ID_Estructura_Parent LONG, " & _
            "ID_Conjunt LONG, " & _
            "N_Pisos INTEGER, " & _
            "Planta TEXT(20), " & _
            "N_Murs_Construits INTEGER, " & _
            "Largo_m SINGLE, " & _
            "Ancho_m SINGLE, " & _
            "Alto_m SINGLE, " & _
            "Altura_Aprox_m SINGLE, " & _
            "Orientacio_Vano TEXT(5), " & _
            "Dintel TEXT(20), " & _
            "Techo_Natural YESNO, " & _
            "Contraforts YESNO, " & _
            "Cornisa_Entre_Pisos YESNO, " & _
            "Estacas_Fusta YESNO, " & _
            "Arrebossat YESNO, " & _
            "Color_Arrebossat TEXT(20), " & _
            "Pintura_Sobre_Roca YESNO, " & _
            "Color_Pintura_Roca TEXT(20), " & _
            "Dec_Nicho_Quadrat YESNO, " & _
            "Dec_Relieve_T YESNO, " & _
            "Dec_Relieve_T_Inv YESNO, " & _
            "Dec_Relieve_L YESNO, " & _
            "Dec_Relieve_L_Inv YESNO, " & _
            "Dec_Zigzag YESNO, " & _
            "Dec_Escalonat YESNO, " & _
            "Dec_Fris_Greca YESNO, " & _
            "Pintura_Rupestre YESNO, " & _
            "PR_Antropomorfa YESNO, " & _
            "PR_Zoomorfa YESNO, " & _
            "PR_Geometrica YESNO, " & _
            "PR_Abstracta YESNO, " & _
            "PR_Escena_Decap YESNO, " & _
            "ID_Estat LONG, " & _
            "Saqueig YESNO, " & _
            "Incendi YESNO, " & _
            "Activitat_Animal YESNO, " & _
            "Evidencia_Acces_Modern YESNO, " & _
            "Restes_Humanes YESNO, " & _
            "NMI INTEGER, " & _
            "Connexio_Anatomica YESNO, " & _
            "Momificacio YESNO, " & _
            "Fardells_Funeraris YESNO, " & _
            "Restes_Disperses YESNO, " & _
            "Posicio_Flexionada YESNO, " & _
            "Cremacio_Ossos YESNO, " & _
            "Mat_Textils YESNO, " & _
            "Mat_Fusta_Cultural YESNO, " & _
            "Mat_Fibra_Vegetal YESNO, " & _
            "Mat_Ceramica YESNO, " & _
            "Mat_Fauna YESNO, " & _
            "Mat_Banya_Cervol YESNO, " & _
            "Mat_Altres YESNO, " & _
            "C14 YESNO, " & _
            "Crono_Segle_Ini INTEGER, " & _
            "Crono_Segle_Fi INTEGER, " & _
            "Volum_Interior_m3 SINGLE, " & _
            "Volum_Total_m3 SINGLE, " & _
            "ID_Metode_Volum LONG, " & _
            "Volum_Notes TEXT(200), " & _
            "Coord_Lat_WGS84 DOUBLE, " & _
            "Coord_Lon_WGS84 DOUBLE, " & _
            "Coord_E_UTM DOUBLE, " & _
            "Coord_N_UTM DOUBLE, " & _
            "Altitud_msnm SINGLE, " & _
            "Coord_Precisio_m SINGLE, " & _
            "ID_Coord_Metode LONG, " & _
            "URL_Pano TEXT(255), " & _
            "URL_Pano_2 TEXT(255), " & _
            "URL_Giga TEXT(255), " & _
            "URL_3D TEXT(255), " & _
            "Documentat_ChaXR YESNO, " & _
            "ID_Campanya LONG, " & _
            "Notes MEMO)", dbFailOnError
        Debug.Print "  [OK] T_ESTRUCTURES (65 camps)"
    End If

    If Not TableExists(db, "T_DATACIONS") Then
        db.Execute "CREATE TABLE T_DATACIONS (" & _
            "ID COUNTER CONSTRAINT PK_DATACIONS PRIMARY KEY, " & _
            "ID_Estructura LONG NOT NULL, " & _
            "Mostra_Tipus TEXT(30), " & _
            "Data_BP LONG, " & _
            "Sigma1_Ini INTEGER, " & _
            "Sigma1_Fi INTEGER, " & _
            "Sigma2_Ini INTEGER, " & _
            "Sigma2_Fi INTEGER, " & _
            "Ref_Laboratori TEXT(50), " & _
            "Ref_Bibliografica MEMO)", dbFailOnError
        Debug.Print "  [OK] T_DATACIONS"
    End If

    If Not TableExists(db, "T_INDIVIDUS") Then
        db.Execute "CREATE TABLE T_INDIVIDUS (" & _
            "ID COUNTER CONSTRAINT PK_INDIVIDUS PRIMARY KEY, " & _
            "ID_Estructura LONG NOT NULL, " & _
            "Num_Individu INTEGER, " & _
            "Edat_Categoria TEXT(20), " & _
            "Sexe_Categoria TEXT(20), " & _
            "Estat_Preservacio TEXT(20), " & _
            "Notes MEMO)", dbFailOnError
        Debug.Print "  [OK] T_INDIVIDUS"
    End If

    Debug.Print "  → 13 taules OK"
End Sub

' ─────────────────────────────────────────────────────────────
'  2. POBLAR TAULES LOOKUP
' ─────────────────────────────────────────────────────────────
Sub PopulateAllLookups(db As DAO.Database)
    Debug.Print "Poblant lookups..."
    Dim i As Integer

    ' Netejar (ordre invers de dependència)
    SafeExec db, "DELETE FROM L_SECTORES"
    SafeExec db, "DELETE FROM L_SITIOS"
    SafeExec db, "DELETE FROM L_TIPOLOGIA"
    SafeExec db, "DELETE FROM L_SUPORT"
    SafeExec db, "DELETE FROM L_ESTAT"
    SafeExec db, "DELETE FROM L_METODE_VOLUM"
    SafeExec db, "DELETE FROM L_COORD_METODE"
    SafeExec db, "DELETE FROM L_TIPUS_CONJUNT"
    SafeExec db, "DELETE FROM L_CAMPANYA"

    ' ── L_SITIOS ─────────────────────────────────────────────
    db.Execute "INSERT INTO L_SITIOS (Nom_Sitio,Descripcio) VALUES " & _
        "('La Petaca','>12.000 m2 roca exposada, 4 sectors, s.X-XVI. Lat -6.831 / Lon -77.808')", dbFailOnError
    db.Execute "INSERT INTO L_SITIOS (Nom_Sitio,Descripcio) VALUES " & _
        "('Diablo Wasi','6 sectors, predomini cavitats funeraries. Lat -6.848 / Lon -77.815')", dbFailOnError

    ' ── L_SECTORES (ID_Sitio 1=LP, 2=DW) ────────────────────
    Dim secLP(4, 1) As String
    secLP(0, 0) = "LP - General":  secLP(0, 1) = "La Petaca - context general"
    secLP(1, 0) = "LP - Nord":     secLP(1, 1) = "La Petaca - Sector Nord (mida menor)"
    secLP(2, 0) = "LP - Central":  secLP(2, 1) = "La Petaca - Sector Central"
    secLP(3, 0) = "LP - Superior": secLP(3, 1) = "La Petaca - Sector Superior"
    secLP(4, 0) = "LP - Sud":      secLP(4, 1) = "La Petaca - Sector Sud (>=96 estructures)"
    For i = 0 To 4
        db.Execute "INSERT INTO L_SECTORES (ID_Sitio,Nom_Sector,Descripcio) VALUES (1,'" & _
            secLP(i, 0) & "','" & secLP(i, 1) & "')", dbFailOnError
    Next i

    Dim secDW(6, 1) As String
    secDW(0, 0) = "DW - General":  secDW(0, 1) = "Diablo Wasi - context general"
    secDW(1, 0) = "DW - Sector 1": secDW(1, 1) = "Diablo Wasi - Sector 1 (~40 contextos funeraris)"
    secDW(2, 0) = "DW - Sector 2": secDW(2, 1) = "Diablo Wasi - Sector 2 (1 cambra + cova basal)"
    secDW(3, 0) = "DW - Sector 3": secDW(3, 1) = "Diablo Wasi - Sector 3 (cova subterranea)"
    secDW(4, 0) = "DW - Sector 4": secDW(4, 1) = "Diablo Wasi - Sector 4 (~9 contextos funeraris)"
    secDW(5, 0) = "DW - Sector 5": secDW(5, 1) = "Diablo Wasi - Sector 5"
    secDW(6, 0) = "DW - Sector 6": secDW(6, 1) = "Diablo Wasi - Sector 6"
    For i = 0 To 6
        db.Execute "INSERT INTO L_SECTORES (ID_Sitio,Nom_Sector,Descripcio) VALUES (2,'" & _
            secDW(i, 0) & "','" & secDW(i, 1) & "')", dbFailOnError
    Next i

    ' ── L_TIPOLOGIA ──────────────────────────────────────────
    Dim tips(9, 1) As String
    tips(0, 0) = "EA-MAU Mausoleu/Chullpa":    tips(0, 1) = "Estructura construida (3+ murs + sostre artificial) sobre repisa. 1-3 pisos. Predominant a La Petaca."
    tips(1, 0) = "EA-CAM Cambra funeraria":     tips(1, 1) = "Cavitat natural tancada per 1 facana construida. Predominant a Diablo Wasi."
    tips(2, 0) = "EA-PLA-R Plataforma repisa":  tips(2, 1) = "Plataforma constructiva sobre repisa natural. Funcio: transit o base per a mausoleus."
    tips(3, 0) = "EA-PLA-V Plataforma volada":  tips(3, 1) = "Plataforma artificial sobre fustes i lloses, sense repisa natural de suport."
    tips(4, 0) = "NIX Nixol natural":           tips(4, 1) = "Petita cavitat natural (<1m2). Funcio: ossari o enterrament secundari."
    tips(5, 0) = "CAV Caverna/Cova":            tips(5, 1) = "Gran cavitat natural (>1m2) amb us funerari o ritual documentat."
    tips(6, 0) = "PR Pintura rupestre":          tips(6, 1) = "Motiu pictoric sobre roca, documentat de forma independent."
    tips(7, 0) = "MEN Mensula aillada":          tips(7, 1) = "Element estructural aillat sense estructura conservada. Evidencia de circulacio perduda."
    tips(8, 0) = "MIX Mixt":                    tips(8, 1) = "Combinacio de dues o mes categories anteriors."
    tips(9, 0) = "ND No determinat":            tips(9, 1) = "Informacio insuficient per classificar."
    For i = 0 To 9
        db.Execute "INSERT INTO L_TIPOLOGIA (Nom,Descripcio) VALUES ('" & _
            tips(i, 0) & "','" & tips(i, 1) & "')", dbFailOnError
    Next i

    ' ── L_SUPORT ─────────────────────────────────────────────
    Dim sups(8) As String
    sups(0) = "Repisa natural amplia (>2m)"
    sups(1) = "Repisa natural estreta (<2m)"
    sups(2) = "Repisa artificial"
    sups(3) = "Cavitat gran (>10m2)"
    sups(4) = "Cavitat mitjana (1-10m2)"
    sups(5) = "Nixol natural (<1m2)"
    sups(6) = "Grieta"
    sups(7) = "Combinat"
    sups(8) = "ND"
    For i = 0 To 8
        db.Execute "INSERT INTO L_SUPORT (Nom) VALUES ('" & sups(i) & "')", dbFailOnError
    Next i

    ' ── L_ESTAT ──────────────────────────────────────────────
    Dim estats(4) As String
    estats(0) = "Bo"
    estats(1) = "Regular"
    estats(2) = "Pre-col.lapse"
    estats(3) = "Col.lapsat"
    estats(4) = "ND"
    For i = 0 To 4
        db.Execute "INSERT INTO L_ESTAT (Nom) VALUES ('" & estats(i) & "')", dbFailOnError
    Next i

    ' ── L_METODE_VOLUM ───────────────────────────────────────
    db.Execute "INSERT INTO L_METODE_VOLUM (Nom,Descripcio) VALUES ('Calcul LxAxH','Calcul geometric per a cambres rectangulars.')", dbFailOnError
    db.Execute "INSERT INTO L_METODE_VOLUM (Nom,Descripcio) VALUES ('Model fotogrametric','Volum extret del model 3D (MeshLab, CloudCompare).')", dbFailOnError
    db.Execute "INSERT INTO L_METODE_VOLUM (Nom,Descripcio) VALUES ('Estimacio','Estimacio visual o mesures parcials.')", dbFailOnError
    db.Execute "INSERT INTO L_METODE_VOLUM (Nom,Descripcio) VALUES ('ND','No determinat.')", dbFailOnError

    ' ── L_COORD_METODE ───────────────────────────────────────
    db.Execute "INSERT INTO L_COORD_METODE (Nom,Descripcio) VALUES ('Drone RTK','Drone amb GPS RTK. Precisio ~2-5 cm.')", dbFailOnError
    db.Execute "INSERT INTO L_COORD_METODE (Nom,Descripcio) VALUES ('GPS diferencial','GNSS doble frequencia. Precisio ~5-20 cm.')", dbFailOnError
    db.Execute "INSERT INTO L_COORD_METODE (Nom,Descripcio) VALUES ('Fotogrametria','Coordenades del model Metashape georeferenciat.')", dbFailOnError
    db.Execute "INSERT INTO L_COORD_METODE (Nom,Descripcio) VALUES ('GPS mobil','GPS de mobil/tauleta. Precisio 3-10 m. Nomes orientacio.')", dbFailOnError
    db.Execute "INSERT INTO L_COORD_METODE (Nom,Descripcio) VALUES ('Estimacio','Estimacio sobre cartografia o ortofoto. >10 m.')", dbFailOnError
    db.Execute "INSERT INTO L_COORD_METODE (Nom,Descripcio) VALUES ('ND','No determinat.')", dbFailOnError

    ' ── L_TIPUS_CONJUNT ──────────────────────────────────────
    db.Execute "INSERT INTO L_TIPUS_CONJUNT (Nom,Descripcio) VALUES ('Alineament vertical','Estructures en una mateixa vertical de faralló (grieta o estrats successius).')", dbFailOnError
    db.Execute "INSERT INTO L_TIPUS_CONJUNT (Nom,Descripcio) VALUES ('Conjunt de repisa','Multiples EAs compartint una mateixa repisa horitzontal.')", dbFailOnError
    db.Execute "INSERT INTO L_TIPUS_CONJUNT (Nom,Descripcio) VALUES ('Plataforma amb mensules','Repisa volada + mensules que la composen o flanquegen.')", dbFailOnError
    db.Execute "INSERT INTO L_TIPUS_CONJUNT (Nom,Descripcio) VALUES ('Conjunt de cova','Caverna + estructures construides al seu interior.')", dbFailOnError
    db.Execute "INSERT INTO L_TIPUS_CONJUNT (Nom,Descripcio) VALUES ('Xarxa de circulacio','MENs i EA-PLAs que reconstrueixen circulacio aeria perduda.')", dbFailOnError
    db.Execute "INSERT INTO L_TIPUS_CONJUNT (Nom,Descripcio) VALUES ('Conjunt pintural','PR + EA adjacent visualment o funcionalment vinculada.')", dbFailOnError
    db.Execute "INSERT INTO L_TIPUS_CONJUNT (Nom,Descripcio) VALUES ('Grup funcional','Qualsevol altra agrupacio intra-sector amb coherencia funcional.')", dbFailOnError

    ' ── L_CAMPANYA ───────────────────────────────────────────
    db.Execute "INSERT INTO L_CAMPANYA (Codi,Nom_Camp,Descripcio) VALUES ('2013','PALP I','Primera campanya. Prospeccio i doc. inicial de La Petaca. Dr. J. Marla Toyne (UCF).')", dbFailOnError
    db.Execute "INSERT INTO L_CAMPANYA (Codi,Nom_Camp,Descripcio) VALUES ('2016','PALP II','Ampliacio a Diablo Wasi. Primera doc. sistematica de DW. Introduccio fotogrametria.')", dbFailOnError
    db.Execute "INSERT INTO L_CAMPANYA (Codi,Nom_Camp,Descripcio) VALUES ('2021','La Petaca Project','Doc. no invasiva integral (360, gigafotos, fotogrametria). Panograma Labs / UCF. DAFO.')", dbFailOnError
    db.Execute "INSERT INTO L_CAMPANYA (Codi,Nom_Camp,Descripcio) VALUES ('2023','PALP IV','Campanya d''excavacio arqueologica i reconstruccions 3D detallades de contextos clau.')", dbFailOnError

    Debug.Print "  → Lookups poblats OK"
End Sub

' ─────────────────────────────────────────────────────────────
'  3. RELACIONS
' ─────────────────────────────────────────────────────────────
Sub CreateAllRelationships(db As DAO.Database)
    Debug.Print "Establint relacions..."

    ' Elimina relacions existents per poder re-executar
    Dim relsToDelete(12) As String
    relsToDelete(0)  = "REL_SITIOS_SECTORES"
    relsToDelete(1)  = "REL_SECTORES_ESTRUCT"
    relsToDelete(2)  = "REL_SECTORES_CONJUNTS"
    relsToDelete(3)  = "REL_TIPOLOGIA_ESTRUCT"
    relsToDelete(4)  = "REL_SUPORT_ESTRUCT"
    relsToDelete(5)  = "REL_ESTAT_ESTRUCT"
    relsToDelete(6)  = "REL_METVOL_ESTRUCT"
    relsToDelete(7)  = "REL_COORDMET_ESTRUCT"
    relsToDelete(8)  = "REL_TIPUSCONJ_CONJ"
    relsToDelete(9)  = "REL_CONJUNTS_ESTRUCT"
    relsToDelete(10) = "REL_ESTRUCT_SELF"
    relsToDelete(11) = "REL_ESTRUCT_DATACIONS"
    relsToDelete(12) = "REL_ESTRUCT_INDIVIDUS"

    Dim n As Integer
    For n = 0 To 12
        On Error Resume Next
        db.Relations.Delete relsToDelete(n)
        On Error GoTo 0
    Next n

    ' Clau: parentTable.parentField → childTable.childField
    '       cascade=True activa DELETE CASCADE
    '       noIntegrity=True per a la relacio autoreferenciant
    MakeRelation db, "REL_SITIOS_SECTORES",    "L_SITIOS",        "ID", "L_SECTORES",     "ID_Sitio",             True, False
    MakeRelation db, "REL_SECTORES_ESTRUCT",   "L_SECTORES",      "ID", "T_ESTRUCTURES",  "ID_Sector",            True, False
    MakeRelation db, "REL_SECTORES_CONJUNTS",  "L_SECTORES",      "ID", "T_CONJUNTS",     "ID_Sector",            True, False
    MakeRelation db, "REL_TIPOLOGIA_ESTRUCT",  "L_TIPOLOGIA",     "ID", "T_ESTRUCTURES",  "ID_Tipologia",         False, False
    MakeRelation db, "REL_SUPORT_ESTRUCT",     "L_SUPORT",        "ID", "T_ESTRUCTURES",  "ID_Suport",            False, False
    MakeRelation db, "REL_ESTAT_ESTRUCT",      "L_ESTAT",         "ID", "T_ESTRUCTURES",  "ID_Estat",             False, False
    MakeRelation db, "REL_METVOL_ESTRUCT",     "L_METODE_VOLUM",  "ID", "T_ESTRUCTURES",  "ID_Metode_Volum",      False, False
    MakeRelation db, "REL_COORDMET_ESTRUCT",   "L_COORD_METODE",  "ID", "T_ESTRUCTURES",  "ID_Coord_Metode",      False, False
    MakeRelation db, "REL_TIPUSCONJ_CONJ",     "L_TIPUS_CONJUNT", "ID", "T_CONJUNTS",     "ID_Tipus_Conjunt",     False, False
    MakeRelation db, "REL_CONJUNTS_ESTRUCT",   "T_CONJUNTS",      "ID", "T_ESTRUCTURES",  "ID_Conjunt",           False, False
    MakeRelation db, "REL_ESTRUCT_SELF",       "T_ESTRUCTURES",   "ID", "T_ESTRUCTURES",  "ID_Estructura_Parent", False, True
    MakeRelation db, "REL_ESTRUCT_DATACIONS",  "T_ESTRUCTURES",   "ID", "T_DATACIONS",    "ID_Estructura",        True, False
    MakeRelation db, "REL_ESTRUCT_INDIVIDUS",  "T_ESTRUCTURES",   "ID", "T_INDIVIDUS",    "ID_Estructura",        True, False

    db.Relations.Refresh
    Debug.Print "  → 13 relacions OK"
End Sub

Sub MakeRelation(db As DAO.Database, relName As String, _
                 parentTbl As String, parentFld As String, _
                 childTbl As String, childFld As String, _
                 deleteCascade As Boolean, noIntegrity As Boolean)
    Dim rel As DAO.Relation
    Dim fld As DAO.Field
    Dim flags As Long

    On Error GoTo ErrRel
    If noIntegrity Then
        flags = dbRelationDontEnforceIntegrity
    Else
        flags = dbRelationUpdateCascade
        If deleteCascade Then flags = flags Or dbRelationDeleteCascade
    End If

    Set rel = db.CreateRelation(relName, parentTbl, childTbl, flags)
    Set fld = rel.CreateField(parentFld)
    fld.ForeignName = childFld
    rel.Fields.Append fld
    db.Relations.Append rel
    Debug.Print "  [OK] " & relName
    Exit Sub
ErrRel:
    Debug.Print "  [AVIS] " & relName & ": " & Err.Description
End Sub

' ─────────────────────────────────────────────────────────────
'  4. CONSULTES SQL
' ─────────────────────────────────────────────────────────────
Sub CreateAllQueries(db As DAO.Database)
    Debug.Print "Creant consultes..."

    Dim qs(9) As String  ' noms
    qs(0) = "QRY_01_Tipologia_x_Jaciment"
    qs(1) = "QRY_02_Decoracio_x_Jaciment"
    qs(2) = "QRY_03_Conservacio_x_Sector"
    qs(3) = "QRY_04_Estructures_C14"
    qs(4) = "QRY_05_Export_RSPSS"
    qs(5) = "QRY_06_Volumetria_x_Tipologia"
    qs(6) = "QRY_07_Export_QGIS"
    qs(7) = "QRY_08_Fills_dun_Pare"
    qs(8) = "QRY_09_Membres_Conjunt"
    qs(9) = "QRY_10_Cobertura_ChaXR"

    Dim i As Integer
    For i = 0 To 9
        If QueryExists(db, qs(i)) Then db.QueryDefs.Delete qs(i)
    Next i

    ' QRY_01 — Tipologia per jaciment
    db.CreateQueryDef qs(0), _
        "SELECT S.Nom_Sitio, T.Nom AS Tipologia, COUNT(E.ID) AS N " & _
        "FROM ((T_ESTRUCTURES AS E " & _
        "INNER JOIN L_SECTORES AS SC ON E.ID_Sector=SC.ID) " & _
        "INNER JOIN L_SITIOS AS S ON SC.ID_Sitio=S.ID) " & _
        "INNER JOIN L_TIPOLOGIA AS T ON E.ID_Tipologia=T.ID " & _
        "GROUP BY S.Nom_Sitio, T.Nom " & _
        "ORDER BY S.Nom_Sitio, T.Nom;"

    ' QRY_02 — Decoració per jaciment (per a χ²)
    db.CreateQueryDef qs(1), _
        "SELECT S.Nom_Sitio, " & _
        "SUM(IIF(E.Dec_Nicho_Quadrat=True,1,0)) AS N_Nicho, " & _
        "SUM(IIF(E.Dec_Relieve_T=True,1,0)) AS N_T, " & _
        "SUM(IIF(E.Dec_Relieve_L=True,1,0)) AS N_L, " & _
        "SUM(IIF(E.Dec_Zigzag=True,1,0)) AS N_Zigzag, " & _
        "SUM(IIF(E.Dec_Escalonat=True,1,0)) AS N_Escalonat, " & _
        "SUM(IIF(E.Pintura_Rupestre=True,1,0)) AS N_PR, " & _
        "COUNT(E.ID) AS Total " & _
        "FROM (T_ESTRUCTURES AS E " & _
        "INNER JOIN L_SECTORES AS SC ON E.ID_Sector=SC.ID) " & _
        "INNER JOIN L_SITIOS AS S ON SC.ID_Sitio=S.ID " & _
        "GROUP BY S.Nom_Sitio;"

    ' QRY_03 — Conservació per sector
    db.CreateQueryDef qs(2), _
        "SELECT S.Nom_Sitio, SC.Nom_Sector, ES.Nom AS Estat, COUNT(*) AS N " & _
        "FROM (((T_ESTRUCTURES AS E " & _
        "INNER JOIN L_SECTORES AS SC ON E.ID_Sector=SC.ID) " & _
        "INNER JOIN L_SITIOS AS S ON SC.ID_Sitio=S.ID) " & _
        "INNER JOIN L_ESTAT AS ES ON E.ID_Estat=ES.ID) " & _
        "GROUP BY S.Nom_Sitio, SC.Nom_Sector, ES.Nom " & _
        "ORDER BY S.Nom_Sitio, SC.Nom_Sector;"

    ' QRY_04 — Estructures amb C14
    db.CreateQueryDef qs(3), _
        "SELECT E.Codi, S.Nom_Sitio, SC.Nom_Sector, T.Nom AS Tipologia, " & _
        "E.Crono_Segle_Ini, E.Crono_Segle_Fi, " & _
        "E.Crono_Segle_Fi - E.Crono_Segle_Ini AS Rang_Segles " & _
        "FROM (((T_ESTRUCTURES AS E " & _
        "INNER JOIN L_SECTORES AS SC ON E.ID_Sector=SC.ID) " & _
        "INNER JOIN L_SITIOS AS S ON SC.ID_Sitio=S.ID) " & _
        "INNER JOIN L_TIPOLOGIA AS T ON E.ID_Tipologia=T.ID) " & _
        "WHERE E.C14=True " & _
        "ORDER BY E.Crono_Segle_Ini;"

    ' QRY_05 — Exportació plana per a R/SPSS
    db.CreateQueryDef qs(4), _
        "SELECT E.ID, E.Codi, S.Nom_Sitio AS Jaciment, SC.Nom_Sector AS Sector, " & _
        "T.Nom AS Tipologia, SU.Nom AS Suport, ES.Nom AS Estat_Conservacio, " & _
        "E.N_Pisos, E.Planta, E.N_Murs_Construits, E.Orientacio_Vano, E.Dintel, " & _
        "E.Techo_Natural, E.Contraforts, E.Arrebossat, E.Pintura_Sobre_Roca, " & _
        "E.Dec_Nicho_Quadrat, E.Dec_Relieve_T, E.Dec_Relieve_T_Inv, " & _
        "E.Dec_Relieve_L, E.Dec_Relieve_L_Inv, E.Dec_Zigzag, E.Dec_Escalonat, " & _
        "E.Dec_Fris_Greca, E.Pintura_Rupestre, " & _
        "E.Saqueig, E.Incendi, E.Activitat_Animal, E.Evidencia_Acces_Modern, " & _
        "E.Restes_Humanes, E.NMI, E.Momificacio, E.Fardells_Funeraris, E.Cremacio_Ossos, " & _
        "E.Mat_Textils, E.Mat_Ceramica, E.Mat_Banya_Cervol, " & _
        "E.C14, E.Crono_Segle_Ini, E.Crono_Segle_Fi, " & _
        "E.Volum_Interior_m3, E.Volum_Total_m3, " & _
        "E.Coord_Lat_WGS84, E.Coord_Lon_WGS84, E.Coord_E_UTM, E.Coord_N_UTM, E.Altitud_msnm, " & _
        "E.Documentat_ChaXR " & _
        "FROM ((((T_ESTRUCTURES AS E " & _
        "INNER JOIN L_SECTORES AS SC ON E.ID_Sector=SC.ID) " & _
        "INNER JOIN L_SITIOS AS S ON SC.ID_Sitio=S.ID) " & _
        "INNER JOIN L_TIPOLOGIA AS T ON E.ID_Tipologia=T.ID) " & _
        "INNER JOIN L_SUPORT AS SU ON E.ID_Suport=SU.ID) " & _
        "INNER JOIN L_ESTAT AS ES ON E.ID_Estat=ES.ID;"

    ' QRY_06 — Volumetria per tipologia i jaciment
    db.CreateQueryDef qs(5), _
        "SELECT S.Nom_Sitio, T.Nom AS Tipologia, " & _
        "COUNT(E.ID) AS N_Total, COUNT(E.Volum_Interior_m3) AS N_amb_Volum, " & _
        "AVG(E.Volum_Interior_m3) AS Volum_Mitja_m3, " & _
        "MIN(E.Volum_Interior_m3) AS Volum_Min_m3, " & _
        "MAX(E.Volum_Interior_m3) AS Volum_Max_m3 " & _
        "FROM ((T_ESTRUCTURES AS E " & _
        "INNER JOIN L_SECTORES AS SC ON E.ID_Sector=SC.ID) " & _
        "INNER JOIN L_SITIOS AS S ON SC.ID_Sitio=S.ID) " & _
        "INNER JOIN L_TIPOLOGIA AS T ON E.ID_Tipologia=T.ID " & _
        "WHERE E.Volum_Interior_m3 IS NOT NULL " & _
        "GROUP BY S.Nom_Sitio, T.Nom " & _
        "ORDER BY S.Nom_Sitio, Volum_Mitja_m3 DESC;"

    ' QRY_07 — Exportació per a QGIS
    db.CreateQueryDef qs(6), _
        "SELECT E.ID, E.Codi, S.Nom_Sitio, SC.Nom_Sector, T.Nom AS Tipologia, " & _
        "E.Coord_Lat_WGS84, E.Coord_Lon_WGS84, E.Coord_E_UTM, E.Coord_N_UTM, " & _
        "E.Altitud_msnm, E.Altura_Aprox_m, E.Coord_Precisio_m, " & _
        "ES.Nom AS Estat_Conservacio, E.N_Pisos, E.Volum_Interior_m3, " & _
        "E.Crono_Segle_Ini, E.Crono_Segle_Fi, " & _
        "E.Dec_Nicho_Quadrat, E.Dec_Relieve_T, E.Pintura_Rupestre, " & _
        "E.Saqueig, E.Restes_Humanes, E.NMI, E.Documentat_ChaXR, E.URL_3D " & _
        "FROM (((T_ESTRUCTURES AS E " & _
        "INNER JOIN L_SECTORES AS SC ON E.ID_Sector=SC.ID) " & _
        "INNER JOIN L_SITIOS AS S ON SC.ID_Sitio=S.ID) " & _
        "INNER JOIN L_TIPOLOGIA AS T ON E.ID_Tipologia=T.ID) " & _
        "INNER JOIN L_ESTAT AS ES ON E.ID_Estat=ES.ID " & _
        "WHERE E.Coord_Lat_WGS84 IS NOT NULL " & _
        "ORDER BY S.Nom_Sitio, E.Codi;"

    ' QRY_08 — Fills d'un element pare (contenció)
    db.CreateQueryDef qs(7), _
        "SELECT E_fill.Codi AS Codi_Fill, T.Nom AS Tipologia_Fill, " & _
        "E_fill.Restes_Humanes, E_fill.NMI, E_fill.ID_Conjunt " & _
        "FROM T_ESTRUCTURES AS E_fill " & _
        "INNER JOIN L_TIPOLOGIA AS T ON E_fill.ID_Tipologia=T.ID " & _
        "WHERE E_fill.ID_Estructura_Parent = [Introdueix l''ID del pare] " & _
        "ORDER BY T.Nom, E_fill.Codi;"

    ' QRY_09 — Membres d'un conjunt funcional
    db.CreateQueryDef qs(8), _
        "SELECT C.Codi_Conjunt, TC.Nom AS Tipus_Conjunt, " & _
        "E.Codi, T.Nom AS Tipologia, " & _
        "E.Coord_E_UTM, E.Coord_N_UTM, E.Altitud_msnm, E.ID_Estructura_Parent " & _
        "FROM (((T_ESTRUCTURES AS E " & _
        "INNER JOIN T_CONJUNTS AS C ON E.ID_Conjunt=C.ID) " & _
        "INNER JOIN L_TIPOLOGIA AS T ON E.ID_Tipologia=T.ID) " & _
        "INNER JOIN L_TIPUS_CONJUNT AS TC ON C.ID_Tipus_Conjunt=TC.ID) " & _
        "WHERE C.Codi_Conjunt = [Introdueix el codi del conjunt] " & _
        "ORDER BY E.Altitud_msnm DESC;"

    ' QRY_10 — Cobertura Chacha XR vs total
    db.CreateQueryDef qs(9), _
        "SELECT S.Nom_Sitio, CA.Codi AS Campanya, " & _
        "COUNT(E.ID) AS Total, " & _
        "SUM(IIF(E.Documentat_ChaXR=True,1,0)) AS Publicat_ChaXR, " & _
        "SUM(IIF(E.URL_3D IS NOT NULL,1,0)) AS Amb_Model_3D " & _
        "FROM ((T_ESTRUCTURES AS E " & _
        "INNER JOIN L_SECTORES AS SC ON E.ID_Sector=SC.ID) " & _
        "INNER JOIN L_SITIOS AS S ON SC.ID_Sitio=S.ID) " & _
        "LEFT JOIN L_CAMPANYA AS CA ON E.ID_Campanya=CA.ID " & _
        "GROUP BY S.Nom_Sitio, CA.Codi " & _
        "ORDER BY S.Nom_Sitio, CA.Codi;"

    Debug.Print "  → 10 consultes OK"
End Sub

' ─────────────────────────────────────────────────────────────
'  PAS FINAL: COMBOS (executar manualment si cal)
'  Configura els ComboBox per als camps FK de T_ESTRUCTURES
' ─────────────────────────────────────────────────────────────
Sub SetupCombos()
    ' Executa aquesta sub després de RunAll()
    ' Configura els desplegables per als camps FK principals

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Set db = CurrentDb()
    Set tdf = db.TableDefs("T_ESTRUCTURES")

    Dim combos(7, 3) As String  ' field, source, col1w, col2w
    combos(0, 0) = "ID_Sector":       combos(0, 1) = "SELECT ID, Nom_Sector FROM L_SECTORES ORDER BY ID_Sitio, Nom_Sector"
    combos(1, 0) = "ID_Tipologia":    combos(1, 1) = "SELECT ID, Nom FROM L_TIPOLOGIA ORDER BY Nom"
    combos(2, 0) = "ID_Suport":       combos(2, 1) = "SELECT ID, Nom FROM L_SUPORT ORDER BY ID"
    combos(3, 0) = "ID_Estat":        combos(3, 1) = "SELECT ID, Nom FROM L_ESTAT ORDER BY ID"
    combos(4, 0) = "ID_Metode_Volum": combos(4, 1) = "SELECT ID, Nom FROM L_METODE_VOLUM ORDER BY ID"
    combos(5, 0) = "ID_Coord_Metode": combos(5, 1) = "SELECT ID, Nom FROM L_COORD_METODE ORDER BY ID"
    combos(6, 0) = "ID_Campanya":     combos(6, 1) = "SELECT ID, Codi, Nom_Camp FROM L_CAMPANYA ORDER BY Codi"
    combos(7, 0) = "ID_Conjunt":      combos(7, 1) = "SELECT ID, Codi_Conjunt FROM T_CONJUNTS ORDER BY Codi_Conjunt"

    Dim i As Integer
    For i = 0 To 7
        Set fld = tdf.Fields(combos(i, 0))
        On Error Resume Next
        fld.Properties("DisplayControl") = 111  ' acComboBox
        If Err.Number <> 0 Then
            Dim prp As DAO.Property
            Set prp = fld.CreateProperty("DisplayControl", DB_INTEGER, 111)
            fld.Properties.Append prp
            Err.Clear
        End If
        On Error GoTo 0

        Dim setProp As Boolean
        setProp = True
        fld.Properties("RowSourceType").Value = "Table/Query"
        fld.Properties("RowSource").Value = combos(i, 1)
        fld.Properties("BoundColumn").Value = 1
        fld.Properties("ColumnCount").Value = 2
        fld.Properties("ColumnWidths").Value = "0cm;4cm"
        fld.Properties("LimitToList").Value = True
        Debug.Print "  [OK] ComboBox: " & combos(i, 0)
    Next i

    tdf.Fields.Refresh
    Set db = Nothing
    MsgBox "ComboBox configurats correctament!", vbInformation
End Sub
