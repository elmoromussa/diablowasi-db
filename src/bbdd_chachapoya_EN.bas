Option Compare Database
Option Explicit

' ================================================================
'  CHACHAPOYA ARCHAEOLOGICAL DATABASE  -  English version
'  La Petaca & Diablo Wasi (Leymebamba, Peru)
'  Author: Esteve Ribera Torro  |  TFM Arqueologia UA 2024-25
'
'  INSTRUCTIONS:
'  1. Open a NEW blank Access database
'  2. Alt+F11 -> Insert -> Module
'  3. Paste this code
'  4. Cursor inside Sub RunAll() -> F5
' ================================================================

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
    msg = "Database created successfully!" & vbCrLf & vbCrLf
    msg = msg & "  13 tables" & vbCrLf
    msg = msg & "  All lookups populated" & vbCrLf
    msg = msg & "  13 relationships" & vbCrLf
    msg = msg & "  10 queries" & vbCrLf & vbCrLf
    msg = msg & "Check the Navigation Panel."
    MsgBox msg, vbInformation, "Done!"
End Sub

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
    Debug.Print "Warning: " & Err.Description
End Sub

' -------------------------------------------------------------
'  1. CREATE TABLES
' -------------------------------------------------------------
Sub CreateAllTables(db As DAO.Database)
    Dim sql As String

    ' -- LOOKUP TABLES ----------------------------------------

    If Not TableExists(db, "L_SITES") Then
        db.Execute "CREATE TABLE L_SITES (ID COUNTER CONSTRAINT PK_SIT PRIMARY KEY, Site_Name TEXT(50) NOT NULL, Description MEMO)", dbFailOnError
    End If

    If Not TableExists(db, "L_SECTORS") Then
        db.Execute "CREATE TABLE L_SECTORS (ID COUNTER CONSTRAINT PK_SEC PRIMARY KEY, ID_Site LONG NOT NULL, Sector_Name TEXT(50) NOT NULL, Description TEXT(255))", dbFailOnError
    End If

    If Not TableExists(db, "L_TYPOLOGY") Then
        db.Execute "CREATE TABLE L_TYPOLOGY (ID COUNTER CONSTRAINT PK_TYP PRIMARY KEY, Name TEXT(60) NOT NULL, Description MEMO)", dbFailOnError
    End If

    If Not TableExists(db, "L_SUPPORT") Then
        db.Execute "CREATE TABLE L_SUPPORT (ID COUNTER CONSTRAINT PK_SUP PRIMARY KEY, Name TEXT(60) NOT NULL)", dbFailOnError
    End If

    If Not TableExists(db, "L_STATUS") Then
        db.Execute "CREATE TABLE L_STATUS (ID COUNTER CONSTRAINT PK_STA PRIMARY KEY, Name TEXT(30) NOT NULL)", dbFailOnError
    End If

    If Not TableExists(db, "L_VOL_METHOD") Then
        db.Execute "CREATE TABLE L_VOL_METHOD (ID COUNTER CONSTRAINT PK_VM PRIMARY KEY, Name TEXT(50) NOT NULL, Description TEXT(255))", dbFailOnError
    End If

    If Not TableExists(db, "L_COORD_METHOD") Then
        db.Execute "CREATE TABLE L_COORD_METHOD (ID COUNTER CONSTRAINT PK_CM PRIMARY KEY, Name TEXT(50) NOT NULL, Description TEXT(255))", dbFailOnError
    End If

    If Not TableExists(db, "L_GROUP_TYPE") Then
        db.Execute "CREATE TABLE L_GROUP_TYPE (ID COUNTER CONSTRAINT PK_GT PRIMARY KEY, Name TEXT(50) NOT NULL, Description MEMO)", dbFailOnError
    End If

    If Not TableExists(db, "L_CAMPAIGN") Then
        db.Execute "CREATE TABLE L_CAMPAIGN (ID COUNTER CONSTRAINT PK_CAM PRIMARY KEY, Code TEXT(4) NOT NULL, Campaign_Name TEXT(80), Description MEMO)", dbFailOnError
    End If

    ' -- SECONDARY TABLES -------------------------------------

    If Not TableExists(db, "T_GROUPS") Then
        db.Execute "CREATE TABLE T_GROUPS (ID COUNTER CONSTRAINT PK_GRP PRIMARY KEY, Group_Code TEXT(20) NOT NULL, ID_Sector LONG NOT NULL, ID_Group_Type LONG NOT NULL, N_Members INTEGER, Notes MEMO)", dbFailOnError
    End If

    ' -- MAIN TABLE -------------------------------------------

    If Not TableExists(db, "T_STRUCTURES") Then
        sql = "CREATE TABLE T_STRUCTURES ("
        sql = sql & "ID COUNTER CONSTRAINT PK_STR PRIMARY KEY,"
        sql = sql & "Code TEXT(20) NOT NULL,"
        sql = sql & "ID_Sector LONG NOT NULL,"
        sql = sql & "ID_Typology LONG,"
        sql = sql & "ID_Support LONG,"
        sql = sql & "ID_Parent LONG,"
        sql = sql & "ID_Group LONG,"
        sql = sql & "N_Floors INTEGER,"
        sql = sql & "Floor_Plan TEXT(20),"
        sql = sql & "N_Built_Walls INTEGER,"
        sql = sql & "Length_m SINGLE,"
        sql = sql & "Width_m SINGLE,"
        sql = sql & "Height_m SINGLE,"
        sql = sql & "Approx_Height_m SINGLE,"
        sql = sql & "Access_Orientation TEXT(5),"
        sql = sql & "Lintel TEXT(20),"
        sql = sql & "Natural_Roof YESNO,"
        sql = sql & "Buttresses YESNO,"
        sql = sql & "Interlevel_Cornice YESNO,"
        sql = sql & "Wooden_Stakes YESNO,"
        sql = sql & "Plastered YESNO,"
        sql = sql & "Plaster_Color TEXT(20),"
        sql = sql & "Rock_Painting YESNO,"
        sql = sql & "Rock_Paint_Color TEXT(20),"
        sql = sql & "Dec_Square_Niche YESNO,"
        sql = sql & "Dec_Relief_T YESNO,"
        sql = sql & "Dec_Relief_T_Inv YESNO,"
        sql = sql & "Dec_Relief_L YESNO,"
        sql = sql & "Dec_Relief_L_Inv YESNO,"
        sql = sql & "Dec_Zigzag YESNO,"
        sql = sql & "Dec_Stepped YESNO,"
        sql = sql & "Dec_Frieze YESNO,"
        sql = sql & "Rock_Art YESNO,"
        sql = sql & "RA_Anthropomorphic YESNO,"
        sql = sql & "RA_Zoomorphic YESNO,"
        sql = sql & "RA_Geometric YESNO,"
        sql = sql & "RA_Abstract YESNO,"
        sql = sql & "RA_Decap_Scene YESNO,"
        sql = sql & "ID_Status LONG,"
        sql = sql & "Looting YESNO,"
        sql = sql & "Fire_Damage YESNO,"
        sql = sql & "Animal_Activity YESNO,"
        sql = sql & "Modern_Access YESNO,"
        sql = sql & "Human_Remains YESNO,"
        sql = sql & "MNI INTEGER,"
        sql = sql & "Anatomical_Connection YESNO,"
        sql = sql & "Mummification YESNO,"
        sql = sql & "Funerary_Bundles YESNO,"
        sql = sql & "Dispersed_Remains YESNO,"
        sql = sql & "Flexed_Position YESNO,"
        sql = sql & "Bone_Burning YESNO,"
        sql = sql & "Mat_Textiles YESNO,"
        sql = sql & "Mat_Wood YESNO,"
        sql = sql & "Mat_VegFiber YESNO,"
        sql = sql & "Mat_Ceramics YESNO,"
        sql = sql & "Mat_Fauna YESNO,"
        sql = sql & "Mat_DeerAntler YESNO,"
        sql = sql & "Mat_Other YESNO,"
        sql = sql & "C14 YESNO,"
        sql = sql & "Chrono_Start_Cent INTEGER,"
        sql = sql & "Chrono_End_Cent INTEGER,"
        sql = sql & "Interior_Vol_m3 SINGLE,"
        sql = sql & "Total_Vol_m3 SINGLE,"
        sql = sql & "ID_Vol_Method LONG,"
        sql = sql & "Vol_Notes TEXT(200),"
        sql = sql & "Coord_Lat_WGS84 DOUBLE,"
        sql = sql & "Coord_Lon_WGS84 DOUBLE,"
        sql = sql & "Coord_E_UTM DOUBLE,"
        sql = sql & "Coord_N_UTM DOUBLE,"
        sql = sql & "Altitude_masl SINGLE,"
        sql = sql & "Coord_Precision_m SINGLE,"
        sql = sql & "ID_Coord_Method LONG,"
        sql = sql & "URL_Pano TEXT(255),"
        sql = sql & "URL_Pano_2 TEXT(255),"
        sql = sql & "URL_Giga TEXT(255),"
        sql = sql & "URL_3D TEXT(255),"
        sql = sql & "ChaXR_Documented YESNO,"
        sql = sql & "ID_Campaign LONG,"
        sql = sql & "Notes MEMO)"
        db.Execute sql, dbFailOnError
        Debug.Print "[OK] T_STRUCTURES (65 fields)"
    End If

    If Not TableExists(db, "T_DATING") Then
        db.Execute "CREATE TABLE T_DATING (ID COUNTER CONSTRAINT PK_DAT PRIMARY KEY, ID_Structure LONG NOT NULL, Sample_Type TEXT(30), Date_BP LONG, Sigma1_Start INTEGER, Sigma1_End INTEGER, Sigma2_Start INTEGER, Sigma2_End INTEGER, Lab_Reference TEXT(50), Bibliog_Reference MEMO)", dbFailOnError
    End If

    If Not TableExists(db, "T_INDIVIDUALS") Then
        db.Execute "CREATE TABLE T_INDIVIDUALS (ID COUNTER CONSTRAINT PK_IND PRIMARY KEY, ID_Structure LONG NOT NULL, Individual_No INTEGER, Age_Category TEXT(20), Sex_Category TEXT(20), Preservation TEXT(20), Notes MEMO)", dbFailOnError
    End If

    Debug.Print "-> 13 tables OK"
End Sub

' -------------------------------------------------------------
'  2. POPULATE LOOKUPS (English)
' -------------------------------------------------------------
Sub PopulateAllLookups(db As DAO.Database)
    Dim i As Integer

    X db, "DELETE FROM L_SECTORS"
    X db, "DELETE FROM L_SITES"
    X db, "DELETE FROM L_TYPOLOGY"
    X db, "DELETE FROM L_SUPPORT"
    X db, "DELETE FROM L_STATUS"
    X db, "DELETE FROM L_VOL_METHOD"
    X db, "DELETE FROM L_COORD_METHOD"
    X db, "DELETE FROM L_GROUP_TYPE"
    X db, "DELETE FROM L_CAMPAIGN"

    ' L_SITES
    db.Execute "INSERT INTO L_SITES (Site_Name,Description) VALUES ('La Petaca','>12,000 m2 exposed rock, 4 sectors, 10th-16th c. WGS84: Lat -6.8311 / Lon -77.8084')", dbFailOnError
    db.Execute "INSERT INTO L_SITES (Site_Name,Description) VALUES ('Diablo Wasi','6 sectors, predominance of funerary chambers in cavities. WGS84: Lat -6.8475 / Lon -77.8154')", dbFailOnError

    ' L_SECTORS
    Dim s(11, 2) As String
    s(0, 0) = "1": s(0, 1) = "LP - General":  s(0, 2) = "La Petaca - general context"
    s(1, 0) = "1": s(1, 1) = "LP - North":    s(1, 2) = "La Petaca - North Sector"
    s(2, 0) = "1": s(2, 1) = "LP - Central":  s(2, 2) = "La Petaca - Central Sector"
    s(3, 0) = "1": s(3, 1) = "LP - Upper":    s(3, 2) = "La Petaca - Upper Sector"
    s(4, 0) = "1": s(4, 1) = "LP - South":    s(4, 2) = "La Petaca - South Sector (>=96 structures)"
    s(5, 0) = "2": s(5, 1) = "DW - General":  s(5, 2) = "Diablo Wasi - general context"
    s(6, 0) = "2": s(6, 1) = "DW - Sector 1": s(6, 2) = "Diablo Wasi - Sector 1 (~40 funerary contexts)"
    s(7, 0) = "2": s(7, 1) = "DW - Sector 2": s(7, 2) = "Diablo Wasi - Sector 2 (1 chamber + basal cave)"
    s(8, 0) = "2": s(8, 1) = "DW - Sector 3": s(8, 2) = "Diablo Wasi - Sector 3 (underground cave)"
    s(9, 0) = "2": s(9, 1) = "DW - Sector 4": s(9, 2) = "Diablo Wasi - Sector 4 (~9 funerary contexts)"
    s(10,0) = "2": s(10,1) = "DW - Sector 5": s(10,2) = "Diablo Wasi - Sector 5"
    s(11,0) = "2": s(11,1) = "DW - Sector 6": s(11,2) = "Diablo Wasi - Sector 6"
    For i = 0 To 11
        db.Execute "INSERT INTO L_SECTORS (ID_Site,Sector_Name,Description) VALUES (" & s(i,0) & ",'" & s(i,1) & "','" & s(i,2) & "')", dbFailOnError
    Next i

    ' L_TYPOLOGY
    Dim t(9, 1) As String
    t(0,0) = "EA-MAU Mausoleum/Chullpa":   t(0,1) = "Built structure (3+ walls + artificial roof) on ledge. 1-3 storeys. Predominant at La Petaca."
    t(1,0) = "EA-CAM Funerary Chamber":    t(1,1) = "Natural cavity closed by 1 built facade. Predominant at Diablo Wasi."
    t(2,0) = "EA-PLA-R Ledge Platform":    t(2,1) = "Constructive platform on natural ledge. Function: transit or mausoleum base."
    t(3,0) = "EA-PLA-V Aerial Platform":   t(3,1) = "Artificial platform on wooden beams and slabs, without natural ledge support."
    t(4,0) = "NIX Natural Niche":          t(4,1) = "Small natural cavity (<1m2). Function: ossuary or secondary burial."
    t(5,0) = "CAV Cave/Cavern":            t(5,1) = "Large natural cavity (>1m2) with documented funerary or ritual use."
    t(6,0) = "PR Rock Art":               t(6,1) = "Pictorial motif on rock, independently documented."
    t(7,0) = "MEN Isolated Bracket":       t(7,1) = "Isolated structural element without preserved structure. Evidence of lost circulation network."
    t(8,0) = "MIX Mixed":                 t(8,1) = "Combination of two or more previous categories."
    t(9,0) = "ND Undetermined":           t(9,1) = "Insufficient information to classify."
    For i = 0 To 9
        db.Execute "INSERT INTO L_TYPOLOGY (Name,Description) VALUES ('" & t(i,0) & "','" & t(i,1) & "')", dbFailOnError
    Next i

    ' L_SUPPORT
    Dim sup(8) As String
    sup(0) = "Wide natural ledge (>2m)"
    sup(1) = "Narrow natural ledge (<2m)"
    sup(2) = "Artificial ledge"
    sup(3) = "Large cavity (>10m2)"
    sup(4) = "Medium cavity (1-10m2)"
    sup(5) = "Natural niche (<1m2)"
    sup(6) = "Fissure/Crack"
    sup(7) = "Combined"
    sup(8) = "ND"
    For i = 0 To 8
        db.Execute "INSERT INTO L_SUPPORT (Name) VALUES ('" & sup(i) & "')", dbFailOnError
    Next i

    ' L_STATUS
    Dim st(4) As String
    st(0) = "Good": st(1) = "Fair": st(2) = "Pre-collapse": st(3) = "Collapsed": st(4) = "ND"
    For i = 0 To 4
        db.Execute "INSERT INTO L_STATUS (Name) VALUES ('" & st(i) & "')", dbFailOnError
    Next i

    ' L_VOL_METHOD
    db.Execute "INSERT INTO L_VOL_METHOD (Name,Description) VALUES ('L x W x H Calculation','Geometric calculation for regular rectangular chambers.')", dbFailOnError
    db.Execute "INSERT INTO L_VOL_METHOD (Name,Description) VALUES ('Photogrammetric Model','Volume extracted from 3D model (MeshLab, CloudCompare).')", dbFailOnError
    db.Execute "INSERT INTO L_VOL_METHOD (Name,Description) VALUES ('Estimation','Visual estimation or partial measurements.')", dbFailOnError
    db.Execute "INSERT INTO L_VOL_METHOD (Name,Description) VALUES ('ND','Undetermined.')", dbFailOnError

    ' L_COORD_METHOD
    db.Execute "INSERT INTO L_COORD_METHOD (Name,Description) VALUES ('Drone RTK','RTK GPS drone. Precision ~2-5 cm.')", dbFailOnError
    db.Execute "INSERT INTO L_COORD_METHOD (Name,Description) VALUES ('Differential GPS','Dual-frequency GNSS. Precision ~5-20 cm.')", dbFailOnError
    db.Execute "INSERT INTO L_COORD_METHOD (Name,Description) VALUES ('Photogrammetry','Coordinates from georeferenced Metashape model.')", dbFailOnError
    db.Execute "INSERT INTO L_COORD_METHOD (Name,Description) VALUES ('Mobile GPS','Mobile/tablet GPS. Precision 3-10 m. Orientation only.')", dbFailOnError
    db.Execute "INSERT INTO L_COORD_METHOD (Name,Description) VALUES ('Estimation','Estimated from cartography or orthophoto. >10 m.')", dbFailOnError
    db.Execute "INSERT INTO L_COORD_METHOD (Name,Description) VALUES ('ND','Undetermined.')", dbFailOnError

    ' L_GROUP_TYPE
    db.Execute "INSERT INTO L_GROUP_TYPE (Name,Description) VALUES ('Vertical alignment','Structures on the same vertical of the cliff (fissure or successive strata).')", dbFailOnError
    db.Execute "INSERT INTO L_GROUP_TYPE (Name,Description) VALUES ('Ledge cluster','Multiple structures sharing the same horizontal ledge.')", dbFailOnError
    db.Execute "INSERT INTO L_GROUP_TYPE (Name,Description) VALUES ('Platform with brackets','Aerial or artificial platform + brackets composing or flanking it.')", dbFailOnError
    db.Execute "INSERT INTO L_GROUP_TYPE (Name,Description) VALUES ('Cave cluster','Cave + structures built inside it.')", dbFailOnError
    db.Execute "INSERT INTO L_GROUP_TYPE (Name,Description) VALUES ('Circulation network','MENs and EA-PLAs reconstructing a lost aerial circulation route.')", dbFailOnError
    db.Execute "INSERT INTO L_GROUP_TYPE (Name,Description) VALUES ('Rock art cluster','Rock art + adjacent structure visually or functionally linked.')", dbFailOnError
    db.Execute "INSERT INTO L_GROUP_TYPE (Name,Description) VALUES ('Functional group','Any other intra-sector grouping with functional coherence.')", dbFailOnError

    ' L_CAMPAIGN
    db.Execute "INSERT INTO L_CAMPAIGN (Code,Campaign_Name,Description) VALUES ('2013','PALP I','First campaign. Prospection and initial documentation of La Petaca. Dr. J. Marla Toyne (UCF).')", dbFailOnError
    db.Execute "INSERT INTO L_CAMPAIGN (Code,Campaign_Name,Description) VALUES ('2016','PALP II','Extension to Diablo Wasi. First systematic documentation of DW.')", dbFailOnError
    db.Execute "INSERT INTO L_CAMPAIGN (Code,Campaign_Name,Description) VALUES ('2021','La Petaca Project','Non-invasive integral documentation. Photogrammetry, 360, gigaphotos. Panograma Labs/UCF.')", dbFailOnError
    db.Execute "INSERT INTO L_CAMPAIGN (Code,Campaign_Name,Description) VALUES ('2023','PALP IV','Archaeological excavation campaign and detailed 3D reconstructions.')", dbFailOnError

    Debug.Print "-> Lookups populated (English)"
End Sub

' -------------------------------------------------------------
'  3. RELATIONSHIPS
' -------------------------------------------------------------
Sub CreateAllRelationships(db As DAO.Database)
    Dim rn(12) As String
    rn(0)  = "REL_SITES_SEC":    rn(1)  = "REL_SEC_STR":    rn(2)  = "REL_SEC_GRP"
    rn(3)  = "REL_TYP_STR":     rn(4)  = "REL_SUP_STR":    rn(5)  = "REL_STR_SELF"
    rn(6)  = "REL_STA_STR":     rn(7)  = "REL_VM_STR":     rn(8)  = "REL_CM_STR"
    rn(9)  = "REL_GT_GRP":      rn(10) = "REL_GRP_STR":    rn(11) = "REL_STR_DAT"
    rn(12) = "REL_STR_IND"
    Dim n As Integer
    For n = 0 To 12
        On Error Resume Next: db.Relations.Delete rn(n): On Error GoTo 0
    Next n

    MkRel db, rn(0),  "L_SITES",      "ID", "L_SECTORS",    "ID_Site",        True,  False
    MkRel db, rn(1),  "L_SECTORS",    "ID", "T_STRUCTURES", "ID_Sector",      True,  False
    MkRel db, rn(2),  "L_SECTORS",    "ID", "T_GROUPS",     "ID_Sector",      True,  False
    MkRel db, rn(3),  "L_TYPOLOGY",   "ID", "T_STRUCTURES", "ID_Typology",    False, False
    MkRel db, rn(4),  "L_SUPPORT",    "ID", "T_STRUCTURES", "ID_Support",     False, False
    MkRel db, rn(5),  "T_STRUCTURES", "ID", "T_STRUCTURES", "ID_Parent",      False, True
    MkRel db, rn(6),  "L_STATUS",     "ID", "T_STRUCTURES", "ID_Status",      False, False
    MkRel db, rn(7),  "L_VOL_METHOD", "ID", "T_STRUCTURES", "ID_Vol_Method",  False, False
    MkRel db, rn(8),  "L_COORD_METHOD","ID","T_STRUCTURES", "ID_Coord_Method",False, False
    MkRel db, rn(9),  "L_GROUP_TYPE", "ID", "T_GROUPS",     "ID_Group_Type",  False, False
    MkRel db, rn(10), "T_GROUPS",     "ID", "T_STRUCTURES", "ID_Group",       False, False
    MkRel db, rn(11), "T_STRUCTURES", "ID", "T_DATING",     "ID_Structure",   True,  False
    MkRel db, rn(12), "T_STRUCTURES", "ID", "T_INDIVIDUALS","ID_Structure",   True,  False

    db.Relations.Refresh
    Debug.Print "-> 13 relationships OK"
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
    Debug.Print "  Warning " & nm & ": " & Err.Description
End Sub

' -------------------------------------------------------------
'  4. QUERIES (English names + English field/table names)
' -------------------------------------------------------------
Sub CreateAllQueries(db As DAO.Database)
    Dim qn(9) As String
    qn(0) = "QRY_01_Typology_by_Site"
    qn(1) = "QRY_02_Decoration_by_Site"
    qn(2) = "QRY_03_Conservation_by_Sector"
    qn(3) = "QRY_04_C14_Structures"
    qn(4) = "QRY_05_Export_RStats"
    qn(5) = "QRY_06_Volumetry_by_Typology"
    qn(6) = "QRY_07_Export_QGIS"
    qn(7) = "QRY_08_Children_of_Parent"
    qn(8) = "QRY_09_Group_Members"
    qn(9) = "QRY_10_ChaXR_Coverage"
    Dim i As Integer
    For i = 0 To 9
        If QueryExists(db, qn(i)) Then db.QueryDefs.Delete qn(i)
    Next i

    Dim q As String

    ' QRY_01
    q = "SELECT S.Site_Name, T.Name AS Typology, COUNT(E.ID) AS N "
    q = q & "FROM ((T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID) "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID "
    q = q & "GROUP BY S.Site_Name, T.Name ORDER BY S.Site_Name, T.Name;"
    db.CreateQueryDef qn(0), q

    ' QRY_02
    q = "SELECT S.Site_Name, "
    q = q & "SUM(IIF(E.Dec_Square_Niche=True,1,0)) AS N_Niche, "
    q = q & "SUM(IIF(E.Dec_Relief_T=True,1,0)) AS N_T, "
    q = q & "SUM(IIF(E.Dec_Relief_L=True,1,0)) AS N_L, "
    q = q & "SUM(IIF(E.Dec_Zigzag=True,1,0)) AS N_Zigzag, "
    q = q & "SUM(IIF(E.Dec_Stepped=True,1,0)) AS N_Stepped, "
    q = q & "SUM(IIF(E.Rock_Art=True,1,0)) AS N_RockArt, "
    q = q & "COUNT(E.ID) AS Total "
    q = q & "FROM (T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID "
    q = q & "GROUP BY S.Site_Name;"
    db.CreateQueryDef qn(1), q

    ' QRY_03
    q = "SELECT S.Site_Name, SC.Sector_Name, ST.Name AS Status, COUNT(*) AS N "
    q = q & "FROM (((T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID) "
    q = q & "INNER JOIN L_STATUS AS ST ON E.ID_Status=ST.ID) "
    q = q & "GROUP BY S.Site_Name, SC.Sector_Name, ST.Name "
    q = q & "ORDER BY S.Site_Name, SC.Sector_Name;"
    db.CreateQueryDef qn(2), q

    ' QRY_04
    q = "SELECT E.Code, S.Site_Name, SC.Sector_Name, T.Name AS Typology, "
    q = q & "E.Chrono_Start_Cent, E.Chrono_End_Cent "
    q = q & "FROM (((T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID) "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID) "
    q = q & "WHERE E.C14=True ORDER BY E.Chrono_Start_Cent;"
    db.CreateQueryDef qn(3), q

    ' QRY_05 flat export
    q = "SELECT E.ID, E.Code, S.Site_Name AS Site, SC.Sector_Name AS Sector, "
    q = q & "T.Name AS Typology, SU.Name AS Support, ST.Name AS Status, "
    q = q & "E.N_Floors, E.Floor_Plan, E.N_Built_Walls, E.Access_Orientation, "
    q = q & "E.Natural_Roof, E.Buttresses, E.Plastered, E.Rock_Painting, "
    q = q & "E.Dec_Square_Niche, E.Dec_Relief_T, E.Dec_Relief_T_Inv, "
    q = q & "E.Dec_Relief_L, E.Dec_Relief_L_Inv, E.Dec_Zigzag, "
    q = q & "E.Dec_Stepped, E.Dec_Frieze, E.Rock_Art, "
    q = q & "E.Looting, E.Fire_Damage, E.Animal_Activity, E.Modern_Access, "
    q = q & "E.Human_Remains, E.MNI, E.Mummification, E.Funerary_Bundles, "
    q = q & "E.Bone_Burning, E.Mat_Textiles, E.Mat_Ceramics, E.Mat_DeerAntler, "
    q = q & "E.C14, E.Chrono_Start_Cent, E.Chrono_End_Cent, "
    q = q & "E.Interior_Vol_m3, E.Total_Vol_m3, "
    q = q & "E.Coord_Lat_WGS84, E.Coord_Lon_WGS84, "
    q = q & "E.Coord_E_UTM, E.Coord_N_UTM, E.Altitude_masl, "
    q = q & "E.ChaXR_Documented "
    q = q & "FROM ((((T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID) "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID) "
    q = q & "INNER JOIN L_SUPPORT AS SU ON E.ID_Support=SU.ID) "
    q = q & "INNER JOIN L_STATUS AS ST ON E.ID_Status=ST.ID;"
    db.CreateQueryDef qn(4), q

    ' QRY_06
    q = "SELECT S.Site_Name, T.Name AS Typology, "
    q = q & "COUNT(E.ID) AS N_Total, COUNT(E.Interior_Vol_m3) AS N_with_Vol, "
    q = q & "AVG(E.Interior_Vol_m3) AS Mean_Vol, "
    q = q & "MIN(E.Interior_Vol_m3) AS Min_Vol, "
    q = q & "MAX(E.Interior_Vol_m3) AS Max_Vol "
    q = q & "FROM ((T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID) "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID "
    q = q & "WHERE E.Interior_Vol_m3 IS NOT NULL "
    q = q & "GROUP BY S.Site_Name, T.Name ORDER BY S.Site_Name, Mean_Vol DESC;"
    db.CreateQueryDef qn(5), q

    ' QRY_07 QGIS export
    q = "SELECT E.ID, E.Code, S.Site_Name, SC.Sector_Name, T.Name AS Typology, "
    q = q & "E.Coord_Lat_WGS84, E.Coord_Lon_WGS84, "
    q = q & "E.Coord_E_UTM, E.Coord_N_UTM, E.Altitude_masl, "
    q = q & "E.Approx_Height_m, E.Coord_Precision_m, ST.Name AS Status, "
    q = q & "E.N_Floors, E.Interior_Vol_m3, "
    q = q & "E.Chrono_Start_Cent, E.Chrono_End_Cent, "
    q = q & "E.Looting, E.Human_Remains, E.MNI, "
    q = q & "E.ChaXR_Documented, E.URL_3D "
    q = q & "FROM (((T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID) "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID) "
    q = q & "INNER JOIN L_STATUS AS ST ON E.ID_Status=ST.ID "
    q = q & "WHERE E.Coord_Lat_WGS84 IS NOT NULL ORDER BY S.Site_Name, E.Code;"
    db.CreateQueryDef qn(6), q

    ' QRY_08
    q = "SELECT E_c.Code AS Child_Code, T.Name AS Typology, "
    q = q & "E_c.Human_Remains, E_c.MNI "
    q = q & "FROM T_STRUCTURES AS E_c "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E_c.ID_Typology=T.ID "
    q = q & "WHERE E_c.ID_Parent=[Parent ID?] "
    q = q & "ORDER BY T.Name, E_c.Code;"
    db.CreateQueryDef qn(7), q

    ' QRY_09
    q = "SELECT G.Group_Code, GT.Name AS Group_Type, E.Code, T.Name AS Typology, "
    q = q & "E.Coord_E_UTM, E.Coord_N_UTM, E.Altitude_masl "
    q = q & "FROM (((T_STRUCTURES AS E "
    q = q & "INNER JOIN T_GROUPS AS G ON E.ID_Group=G.ID) "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID) "
    q = q & "INNER JOIN L_GROUP_TYPE AS GT ON G.ID_Group_Type=GT.ID) "
    q = q & "WHERE G.Group_Code=[Group code?] "
    q = q & "ORDER BY E.Altitude_masl DESC;"
    db.CreateQueryDef qn(8), q

    ' QRY_10
    q = "SELECT S.Site_Name, CA.Code AS Campaign, "
    q = q & "COUNT(E.ID) AS Total, "
    q = q & "SUM(IIF(E.ChaXR_Documented=True,1,0)) AS Published_ChaXR, "
    q = q & "SUM(IIF(E.URL_3D IS NOT NULL,1,0)) AS With_3D_Model "
    q = q & "FROM ((T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID) "
    q = q & "LEFT JOIN L_CAMPAIGN AS CA ON E.ID_Campaign=CA.ID "
    q = q & "GROUP BY S.Site_Name, CA.Code ORDER BY S.Site_Name, CA.Code;"
    db.CreateQueryDef qn(9), q

    Debug.Print "-> 10 queries OK (English)"
End Sub
