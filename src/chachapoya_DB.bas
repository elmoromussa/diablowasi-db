Option Compare Database
Option Explicit

' ================================================================
'  CHACHAPOYA ARCHAEOLOGICAL DATABASE - COMPLETE BUILD SCRIPT v4
'  La Petaca & Diablo Wasi (Leymebamba, Amazonas, Peru)
'  Author: Esteve Ribera Torro | TFM Arqueologia UA
'  Version: v4 - full consolidation after form review (Aug 2026)
'
'  Changes vs v3:
'   - A-X presence fields converted YESNO -> BYTE with domain
'     0 = Absent (constructive decision)
'     1 = Present
'     9 = ND / not observable (collapse or documentation bias)
'     Default value 9 set via DAO (SetByteDefaults). This separates
'     real absence from taphonomic / observational bias, which is
'     essential for co-occurrence, cluster and Moran I analysis.
'   - New A-X fields: Lateral_Wall_Faces (L), Relief_Frieze (M).
'     Dec_Frieze removed from decoration booleans (now element M).
'     Q (lintel) derived from Lintel material field in QRY_13.
'   - Renamed: Interlevel_Cornice -> Interbody_Cornice (element I)
'              Cornice_Material  -> Interbody_Cornice_Material
'              Approx_Height_m   -> Height_Above_Base_m
'     Terminology: "nivell" reserved for N0/N1/Sup (constructive
'     levels); "cos" (body) for superposed storeys.
'   - New mortar block (H04, cf. Toyne & Anzellini 2017):
'     Mortar_Present (BYTE), Mortar_Type, Chinking_Stones (BYTE),
'     Mortar_Notes
'   - New access opening metrics: Opening_Width_cm, Opening_Height_cm
'   - New Dim_Method field (method used for L/W/H measurements)
'   - NEW TABLE T_CONNECTIONS: physical links between structures
'     (shared ledge, continuous platform, shared wall/beam) ->
'     adjacency matrix for aerial circulation network analysis (OE3)
'   - 2 new relationships (REL_STR_CONA, REL_STR_CONB) -> 20 total
'   - 1 new query QRY_14_Connections_Edges -> 14 total
'   - QRY_02 updated (frieze from Relief_Frieze)
'   - QRY_05 updated (renames + mortar + opening metrics)
'   - QRY_07 updated (Height_Above_Base_m)
'   - QRY_13 rewritten: full 24-column A-X matrix in letter order
'     (AX_A..AX_X aliases, AX_Q derived from Lintel)
'   - T_STRUCTURES: 124 fields
'
'  Run Sub BuildDB() on a NEW BLANK ACCESS DATABASE
'  Then run chachapoya_Form_v3_val.bas -> Sub BuildForm()
'  (chachapoya_patch_metric.bas is OBSOLETE: absorbed by Form v3)
' ================================================================

Sub BuildDB()
    Dim db As DAO.Database
    Set db = CurrentDb()
    CreateAllTables db
    SetByteDefaults db
    PopulateAllLookups db
    CreateAllRelationships db
    CreateAllQueries db
    db.TableDefs.Refresh
    db.QueryDefs.Refresh
    Set db = Nothing
    Dim msg As String
    msg = "DATABASE v4 BUILT SUCCESSFULLY!" & vbCrLf & vbCrLf
    msg = msg & "  19 tables (T_CONNECTIONS is new)" & vbCrLf
    msg = msg & "  All lookups populated" & vbCrLf
    msg = msg & "  20 relationships" & vbCrLf
    msg = msg & "  14 queries" & vbCrLf
    msg = msg & "  T_STRUCTURES: 124 fields" & vbCrLf & vbCrLf
    msg = msg & "Key changes (v4):" & vbCrLf
    msg = msg & "  A-X fields now BYTE: 0=Absent / 1=Present / 9=ND" & vbCrLf
    msg = msg & "  (default 9: absence must be positively recorded)" & vbCrLf
    msg = msg & "  New elements L and M; Q derived from Lintel" & vbCrLf
    msg = msg & "  Interbody cornice terminology (cos vs nivell)" & vbCrLf
    msg = msg & "  Mortar block + opening metrics" & vbCrLf
    msg = msg & "  T_CONNECTIONS + QRY_14 (network analysis, OE3)" & vbCrLf & vbCrLf
    msg = msg & "Next: run chachapoya_Form_v3_val.bas -> BuildForm()"
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

Private Sub X(db As DAO.Database, sql As String)
    On Error GoTo Err_X
    db.Execute sql, dbFailOnError
    Exit Sub
Err_X: Debug.Print "Warning: " & Err.Description
End Sub

' ================================================================
'  1. CREATE ALL TABLES
' ================================================================
Private Sub CreateAllTables(db As DAO.Database)
    Dim sql As String

    ' -- LOOKUP TABLES (no dependencies) --
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
    If Not TableExists(db, "L_MATERIAL_STATUS") Then
        db.Execute "CREATE TABLE L_MATERIAL_STATUS (ID COUNTER CONSTRAINT PK_MS PRIMARY KEY, Name TEXT(50) NOT NULL, Description TEXT(255))", dbFailOnError
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

    ' Decoration position lookup
    On Error Resume Next
    db.Execute "DROP TABLE L_STRUCT_BODY", dbFailOnError
    On Error GoTo 0
    db.Execute "CREATE TABLE L_STRUCT_BODY (ID COUNTER CONSTRAINT PK_SB PRIMARY KEY, Code TEXT(10) NOT NULL, Name TEXT(60) NOT NULL, Body_Level INTEGER, Description TEXT(255))", dbFailOnError

    ' Decoration type lookup
    On Error Resume Next
    db.Execute "DROP TABLE L_DEC_TYPE", dbFailOnError
    On Error GoTo 0
    db.Execute "CREATE TABLE L_DEC_TYPE (ID COUNTER CONSTRAINT PK_DT PRIMARY KEY, Name TEXT(60) NOT NULL, Description TEXT(255))", dbFailOnError

    ' -- SECONDARY TABLES --
    If Not TableExists(db, "T_GROUPS") Then
        db.Execute "CREATE TABLE T_GROUPS (ID COUNTER CONSTRAINT PK_GRP PRIMARY KEY, Group_Code TEXT(20) NOT NULL, ID_Sector LONG NOT NULL, ID_Group_Type LONG NOT NULL, N_Members INTEGER, Notes MEMO)", dbFailOnError
    End If

    ' -- MAIN TABLE: T_STRUCTURES (124 fields) --
    If Not TableExists(db, "T_STRUCTURES") Then
        sql = "CREATE TABLE T_STRUCTURES ("
        ' --- 1. Identification (7) ---
        sql = sql & "ID COUNTER CONSTRAINT PK_STR PRIMARY KEY,"
        sql = sql & "Code TEXT(20) NOT NULL,"
        sql = sql & "ID_Sector LONG NOT NULL,"
        sql = sql & "ID_Typology LONG,"
        sql = sql & "ID_Support LONG,"
        sql = sql & "ID_Parent LONG,"
        sql = sql & "ID_Group LONG,"
        ' --- 2. Morphology & dimensions (15) ---
        sql = sql & "N_Floors INTEGER,"
        sql = sql & "Floor_Plan TEXT(20),"
        sql = sql & "N_Built_Walls INTEGER,"
        sql = sql & "Length_m SINGLE,"
        sql = sql & "Width_m SINGLE,"
        sql = sql & "Height_m SINGLE,"
        sql = sql & "Height_Above_Base_m SINGLE,"
        sql = sql & "Dim_Method TEXT(30),"
        sql = sql & "Opening_Width_cm SINGLE,"
        sql = sql & "Opening_Height_cm SINGLE,"
        sql = sql & "Access_Orientation TEXT(5),"
        sql = sql & "Lintel TEXT(20),"
        sql = sql & "Natural_Roof YESNO,"
        sql = sql & "Buttresses YESNO,"
        sql = sql & "Wooden_Stakes YESNO,"
        ' --- 2b. Support geology detail (4) | H02: geology as determinant ---
        sql = sql & "Support_Width_cm SINGLE,"
        sql = sql & "Support_Depth_cm SINGLE,"
        sql = sql & "Support_Morphology TEXT(30),"
        sql = sql & "Support_Modified YESNO,"
        ' --- 2c. Masonry & mortar (6) | H01/H04: cf. Toyne & Anzellini 2017 ---
        sql = sql & "Masonry_Quality TEXT(20),"
        sql = sql & "Masonry_Type TEXT(30),"
        sql = sql & "Mortar_Present BYTE,"
        sql = sql & "Mortar_Type TEXT(30),"
        sql = sql & "Chinking_Stones BYTE,"
        sql = sql & "Mortar_Notes TEXT(150),"
        ' --- 3. A-X systems - Level N0 (10) | elements A-H ---
        '     BYTE fields: 0=Absent / 1=Present / 9=ND (default 9 via DAO)
        sql = sql & "Embedded_Base_Beams BYTE,"
        sql = sql & "Base_Level BYTE,"
        sql = sql & "Decorative_Socle BYTE,"
        sql = sql & "Tie_Walls BYTE,"
        sql = sql & "Timber_Brackets BYTE,"
        sql = sql & "Timber_Bracket_Count INTEGER,"
        sql = sql & "Transverse_Beams BYTE,"
        sql = sql & "Corbelled_Courses BYTE,"
        sql = sql & "Corbel_Material TEXT(20),"
        sql = sql & "Corbelled_Platform BYTE,"
        ' --- 3b. N0/N1 interface (2) | element I ---
        sql = sql & "Interbody_Cornice BYTE,"
        sql = sql & "Interbody_Cornice_Material TEXT(20),"
        ' --- 4. A-X systems - Level N1 (12) | elements J-P, R, V, W ---
        sql = sql & "Corner_Quoins BYTE,"
        sql = sql & "Structural_Pilasters BYTE,"
        sql = sql & "Lateral_Wall_Faces BYTE,"
        sql = sql & "Relief_Frieze BYTE,"
        sql = sql & "Sill BYTE,"
        sql = sql & "Jambs BYTE,"
        sql = sql & "Access_Opening BYTE,"
        sql = sql & "Recessed_Portal BYTE,"
        sql = sql & "Upper_Crown BYTE,"
        sql = sql & "Lateral_Walls BYTE,"
        sql = sql & "Rear_Wall BYTE,"
        sql = sql & "Rear_Wall_Built BYTE,"
        ' --- 5. A-X systems - Upper zone (4) | elements S, T, U, X ---
        sql = sql & "Eave_Beam BYTE,"
        sql = sql & "Eave_Surface BYTE,"
        sql = sql & "Eave BYTE,"
        sql = sql & "Chamber_Roof BYTE,"
        ' --- 6. Finishes (4) ---
        sql = sql & "Plastered YESNO,"
        sql = sql & "Plaster_Color TEXT(20),"
        sql = sql & "Rock_Painting YESNO,"
        sql = sql & "Rock_Paint_Color TEXT(20),"
        ' --- 6b. Landscape & orientation (2) | observational; QGIS later ---
        sql = sql & "Facade_Orientation TEXT(5),"
        sql = sql & "Visibility_Valley TEXT(10),"
        ' --- 7. Decoration summary booleans (13, detail in T_DECORATIONS) ---
        '     Dec_Frieze removed in v4: frieze is now element M (Relief_Frieze)
        sql = sql & "Dec_Square_Niche YESNO,"
        sql = sql & "Dec_Relief_T YESNO,"
        sql = sql & "Dec_Relief_T_Inv YESNO,"
        sql = sql & "Dec_Relief_L YESNO,"
        sql = sql & "Dec_Relief_L_Inv YESNO,"
        sql = sql & "Dec_Zigzag YESNO,"
        sql = sql & "Dec_Stepped YESNO,"
        sql = sql & "Rock_Art YESNO,"
        sql = sql & "RA_Anthropomorphic YESNO,"
        sql = sql & "RA_Zoomorphic YESNO,"
        sql = sql & "RA_Geometric YESNO,"
        sql = sql & "RA_Abstract YESNO,"
        sql = sql & "RA_Decap_Scene YESNO,"
        ' --- 8. Conservation (6) ---
        sql = sql & "ID_Arch_Status LONG,"
        sql = sql & "ID_Material_Status LONG,"
        sql = sql & "Looting YESNO,"
        sql = sql & "Fire_Damage YESNO,"
        sql = sql & "Animal_Activity YESNO,"
        sql = sql & "Modern_Access YESNO,"
        ' --- 9. Bioarchaeology (8) ---
        sql = sql & "Human_Remains YESNO,"
        sql = sql & "MNI INTEGER,"
        sql = sql & "Anatomical_Connection YESNO,"
        sql = sql & "Mummification YESNO,"
        sql = sql & "Funerary_Bundles YESNO,"
        sql = sql & "Dispersed_Remains YESNO,"
        sql = sql & "Flexed_Position YESNO,"
        sql = sql & "Bone_Burning YESNO,"
        ' --- 10. Cultural materials (7) ---
        sql = sql & "Mat_Textiles YESNO,"
        sql = sql & "Mat_Wood YESNO,"
        sql = sql & "Mat_VegFiber YESNO,"
        sql = sql & "Mat_Ceramics YESNO,"
        sql = sql & "Mat_Fauna YESNO,"
        sql = sql & "Mat_DeerAntler YESNO,"
        sql = sql & "Mat_Other YESNO,"
        ' --- 11. Chronology (3) ---
        sql = sql & "C14 YESNO,"
        sql = sql & "Chrono_Start_Cent INTEGER,"
        sql = sql & "Chrono_End_Cent INTEGER,"
        ' --- 11b. Constructive phases (2) | H03 / H04 ---
        sql = sql & "Construction_Phases INTEGER,"
        sql = sql & "Phase_Evidence TEXT(30),"
        ' --- 12. Volumetry & area (5) ---
        sql = sql & "Interior_Area_m2 SINGLE,"
        sql = sql & "Interior_Vol_m3 SINGLE,"
        sql = sql & "Total_Vol_m3 SINGLE,"
        sql = sql & "ID_Vol_Method LONG,"
        sql = sql & "Vol_Notes TEXT(200),"
        ' --- 13. Spatial coordinates (7) ---
        sql = sql & "Coord_Lat_WGS84 DOUBLE,"
        sql = sql & "Coord_Lon_WGS84 DOUBLE,"
        sql = sql & "Coord_E_UTM DOUBLE,"
        sql = sql & "Coord_N_UTM DOUBLE,"
        sql = sql & "Altitude_masl SINGLE,"
        sql = sql & "Coord_Precision_m SINGLE,"
        sql = sql & "ID_Coord_Method LONG,"
        ' --- 14. Digital documentation (7) ---
        sql = sql & "URL_Pano TEXT(255),"
        sql = sql & "URL_Pano_2 TEXT(255),"
        sql = sql & "URL_Giga TEXT(255),"
        sql = sql & "URL_3D TEXT(255),"
        sql = sql & "ChaXR_Documented YESNO,"
        sql = sql & "ID_Campaign LONG,"
        sql = sql & "Notes MEMO)"
        db.Execute sql, dbFailOnError
        Debug.Print "[OK] T_STRUCTURES (124 fields)"
    End If

    ' -- LINKED TABLES --
    If Not TableExists(db, "T_DATING") Then
        db.Execute "CREATE TABLE T_DATING (ID COUNTER CONSTRAINT PK_DAT PRIMARY KEY, ID_Structure LONG NOT NULL, Sample_Type TEXT(30), Date_BP LONG, Sigma1_Start INTEGER, Sigma1_End INTEGER, Sigma2_Start INTEGER, Sigma2_End INTEGER, Lab_Reference TEXT(50), Bibliog_Reference MEMO)", dbFailOnError
    End If
    If Not TableExists(db, "T_INDIVIDUALS") Then
        db.Execute "CREATE TABLE T_INDIVIDUALS (ID COUNTER CONSTRAINT PK_IND PRIMARY KEY, ID_Structure LONG NOT NULL, Individual_No INTEGER, Age_Category TEXT(20), Sex_Category TEXT(20), Preservation TEXT(20), Notes MEMO)", dbFailOnError
    End If

    ' Decoration detail
    On Error Resume Next
    db.Execute "DROP TABLE T_DECORATIONS", dbFailOnError
    On Error GoTo 0
    db.Execute "CREATE TABLE T_DECORATIONS (ID COUNTER CONSTRAINT PK_TDEC PRIMARY KEY, ID_Structure LONG NOT NULL, ID_Struct_Body LONG, ID_Dec_Type LONG, Body_No INTEGER, Color TEXT(20), Notes TEXT(255))", dbFailOnError

    ' Flexible feature recording (future-proof)
    On Error Resume Next
    db.Execute "DROP TABLE T_ARCH_FEATURES", dbFailOnError
    On Error GoTo 0
    sql = "CREATE TABLE T_ARCH_FEATURES ("
    sql = sql & "ID COUNTER CONSTRAINT PK_AF PRIMARY KEY,"
    sql = sql & "ID_Structure LONG NOT NULL,"
    sql = sql & "Feature_Code TEXT(40),"
    sql = sql & "Present YESNO,"
    sql = sql & "Feature_Count INTEGER,"
    sql = sql & "Material TEXT(20),"
    sql = sql & "Notes TEXT(255))"
    db.Execute sql, dbFailOnError

    ' NEW v4: physical connections between structures (OE3 network)
    On Error Resume Next
    db.Execute "DROP TABLE T_CONNECTIONS", dbFailOnError
    On Error GoTo 0
    sql = "CREATE TABLE T_CONNECTIONS ("
    sql = sql & "ID COUNTER CONSTRAINT PK_CON PRIMARY KEY,"
    sql = sql & "ID_Struct_A LONG NOT NULL,"
    sql = sql & "ID_Struct_B LONG NOT NULL,"
    sql = sql & "Connection_Type TEXT(30),"
    sql = sql & "Confidence TEXT(10),"
    sql = sql & "Notes TEXT(150))"
    db.Execute sql, dbFailOnError

    Debug.Print "-> 19 tables OK"
End Sub

' ================================================================
'  1b. DEFAULT VALUE 9 (ND) FOR ALL BYTE 0/1/9 FIELDS (via DAO)
'      Rationale: a new record must NOT silently claim absence.
'      Absence (0) has to be positively recorded by the researcher.
' ================================================================
Private Sub SetByteDefaults(db As DAO.Database)
    Dim fn(26) As String
    fn(0) = "Embedded_Base_Beams":  fn(1) = "Base_Level"
    fn(2) = "Decorative_Socle":     fn(3) = "Tie_Walls"
    fn(4) = "Timber_Brackets":      fn(5) = "Transverse_Beams"
    fn(6) = "Corbelled_Courses":    fn(7) = "Corbelled_Platform"
    fn(8) = "Interbody_Cornice":    fn(9) = "Corner_Quoins"
    fn(10) = "Structural_Pilasters": fn(11) = "Lateral_Wall_Faces"
    fn(12) = "Relief_Frieze":       fn(13) = "Sill"
    fn(14) = "Jambs":               fn(15) = "Access_Opening"
    fn(16) = "Recessed_Portal":     fn(17) = "Upper_Crown"
    fn(18) = "Lateral_Walls":       fn(19) = "Rear_Wall"
    fn(20) = "Rear_Wall_Built":     fn(21) = "Eave_Beam"
    fn(22) = "Eave_Surface":        fn(23) = "Eave"
    fn(24) = "Chamber_Roof":        fn(25) = "Mortar_Present"
    fn(26) = "Chinking_Stones"
    Dim i As Integer
    On Error Resume Next
    For i = 0 To 26
        db.TableDefs("T_STRUCTURES").Fields(fn(i)).DefaultValue = "9"
    Next i
    On Error GoTo 0
    Debug.Print "-> BYTE defaults (9 = ND) set on 27 fields"
End Sub

' ================================================================
'  2. POPULATE ALL LOOKUPS
' ================================================================
Private Sub PopulateAllLookups(db As DAO.Database)
    Dim i As Integer

    X db, "DELETE FROM L_SECTORS"
    X db, "DELETE FROM L_SITES"
    X db, "DELETE FROM L_TYPOLOGY"
    X db, "DELETE FROM L_SUPPORT"
    X db, "DELETE FROM L_STATUS"
    X db, "DELETE FROM L_MATERIAL_STATUS"
    X db, "DELETE FROM L_VOL_METHOD"
    X db, "DELETE FROM L_COORD_METHOD"
    X db, "DELETE FROM L_GROUP_TYPE"
    X db, "DELETE FROM L_CAMPAIGN"
    X db, "DELETE FROM L_STRUCT_BODY"
    X db, "DELETE FROM L_DEC_TYPE"

    ' L_SITES
    db.Execute "INSERT INTO L_SITES (Site_Name,Description) VALUES ('La Petaca','>12,000 m2 exposed rock, 4 sectors, 10th-16th c. WGS84: Lat -6.8311 / Lon -77.8084')", dbFailOnError
    db.Execute "INSERT INTO L_SITES (Site_Name,Description) VALUES ('Diablo Wasi','6 sectors, predominance of funerary chambers. WGS84: Lat -6.8475 / Lon -77.8154')", dbFailOnError

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
    s(10, 0) = "2": s(10, 1) = "DW - Sector 5": s(10, 2) = "Diablo Wasi - Sector 5"
    s(11, 0) = "2": s(11, 1) = "DW - Sector 6": s(11, 2) = "Diablo Wasi - Sector 6"
    For i = 0 To 11
        db.Execute "INSERT INTO L_SECTORS (ID_Site,Sector_Name,Description) VALUES (" & s(i, 0) & ",'" & s(i, 1) & "','" & s(i, 2) & "')", dbFailOnError
    Next i

    ' L_TYPOLOGY
    Dim t(9, 1) As String
    t(0, 0) = "EA-MAU Mausoleum/Chullpa":  t(0, 1) = "Built structure (3+ walls + artificial roof) on ledge. 1-3 storeys. Predominant at La Petaca."
    t(1, 0) = "EA-CAM Funerary Chamber":   t(1, 1) = "Natural cavity closed by 1 built facade. Predominant at Diablo Wasi."
    t(2, 0) = "EA-PLA-R Ledge Platform":   t(2, 1) = "Constructive platform on natural ledge. Function: transit or mausoleum base."
    t(3, 0) = "EA-PLA-V Aerial Platform":  t(3, 1) = "Artificial platform on wooden beams and slabs, without natural ledge support."
    t(4, 0) = "NIX Natural Niche":         t(4, 1) = "Small natural cavity (<1m2). Function: ossuary or secondary burial."
    t(5, 0) = "CAV Cave/Cavern":           t(5, 1) = "Large natural cavity (>1m2) with documented funerary or ritual use."
    t(6, 0) = "PR Rock Art":               t(6, 1) = "Pictorial motif on rock, independently documented."
    t(7, 0) = "MEN Isolated Bracket":      t(7, 1) = "Isolated structural element. Evidence of lost aerial circulation network."
    t(8, 0) = "MIX Mixed":                 t(8, 1) = "Combination of two or more previous categories."
    t(9, 0) = "ND Undetermined":           t(9, 1) = "Insufficient information to classify."
    For i = 0 To 9
        db.Execute "INSERT INTO L_TYPOLOGY (Name,Description) VALUES ('" & t(i, 0) & "','" & t(i, 1) & "')", dbFailOnError
    Next i

    ' L_SUPPORT
    Dim sup(8) As String
    sup(0) = "Wide natural ledge (>2m)": sup(1) = "Narrow natural ledge (<2m)": sup(2) = "Artificial ledge"
    sup(3) = "Large cavity (>10m2)": sup(4) = "Medium cavity (1-10m2)": sup(5) = "Natural niche (<1m2)"
    sup(6) = "Fissure/Crack": sup(7) = "Combined": sup(8) = "ND"
    For i = 0 To 8
        db.Execute "INSERT INTO L_SUPPORT (Name) VALUES ('" & sup(i) & "')", dbFailOnError
    Next i

    ' L_STATUS
    Dim st(4) As String
    st(0) = "Good": st(1) = "Fair": st(2) = "Pre-collapse": st(3) = "Collapsed": st(4) = "ND"
    For i = 0 To 4
        db.Execute "INSERT INTO L_STATUS (Name) VALUES ('" & st(i) & "')", dbFailOnError
    Next i

    ' L_MATERIAL_STATUS
    db.Execute "INSERT INTO L_MATERIAL_STATUS (Name,Description) VALUES ('Good','Material remains well preserved and identifiable.')", dbFailOnError
    db.Execute "INSERT INTO L_MATERIAL_STATUS (Name,Description) VALUES ('Fair','Partially preserved: some elements present and identifiable.')", dbFailOnError
    db.Execute "INSERT INTO L_MATERIAL_STATUS (Name,Description) VALUES ('Poor','Fragmentary: barely identifiable remains, heavily degraded or scattered.')", dbFailOnError
    db.Execute "INSERT INTO L_MATERIAL_STATUS (Name,Description) VALUES ('Absent','No movable remains documented (cause in Looting/Animal_Activity fields).')", dbFailOnError
    db.Execute "INSERT INTO L_MATERIAL_STATUS (Name,Description) VALUES ('ND','Not determined: not assessed or structure not accessible.')", dbFailOnError

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
    db.Execute "INSERT INTO L_GROUP_TYPE (Name,Description) VALUES ('Platform with brackets','Aerial platform + brackets composing or flanking it.')", dbFailOnError
    db.Execute "INSERT INTO L_GROUP_TYPE (Name,Description) VALUES ('Cave cluster','Cave + structures built inside it.')", dbFailOnError
    db.Execute "INSERT INTO L_GROUP_TYPE (Name,Description) VALUES ('Circulation network','MENs and EA-PLAs reconstructing a lost aerial circulation route.')", dbFailOnError
    db.Execute "INSERT INTO L_GROUP_TYPE (Name,Description) VALUES ('Rock art cluster','Rock art + adjacent structure visually or functionally linked.')", dbFailOnError
    db.Execute "INSERT INTO L_GROUP_TYPE (Name,Description) VALUES ('Functional group','Any other intra-sector grouping with functional coherence.')", dbFailOnError

    ' L_CAMPAIGN
    db.Execute "INSERT INTO L_CAMPAIGN (Code,Campaign_Name,Description) VALUES ('2013','PALP I','First campaign. Prospection and initial documentation of La Petaca. Dr. J. Marla Toyne (UCF).')", dbFailOnError
    db.Execute "INSERT INTO L_CAMPAIGN (Code,Campaign_Name,Description) VALUES ('2016','PALP II','Extension to Diablo Wasi. First systematic documentation of DW.')", dbFailOnError
    db.Execute "INSERT INTO L_CAMPAIGN (Code,Campaign_Name,Description) VALUES ('2021','La Petaca Project','Non-invasive integral documentation. Photogrammetry, 360, gigaphotos. Panograma Labs/UCF.')", dbFailOnError
    db.Execute "INSERT INTO L_CAMPAIGN (Code,Campaign_Name,Description) VALUES ('2023','PALP IV','Archaeological excavation campaign and detailed 3D reconstructions.')", dbFailOnError

    ' L_STRUCT_BODY (v4: interbody terminology - "cos" for storeys)
    Dim sb(7, 3) As String
    sb(0, 0) = "N0-SOC": sb(0, 1) = "Socle (N0)":          sb(0, 2) = "0":  sb(0, 3) = "Decorative socle - level 0 base element."
    sb(1, 0) = "N1-SPA": sb(1, 1) = "Spandrel (N1)":       sb(1, 2) = "1":  sb(1, 3) = "Lateral wall face of main body (outside portal frame)."
    sb(2, 0) = "N1-JAM": sb(2, 1) = "Jamb (N1)":           sb(2, 2) = "1":  sb(2, 3) = "Portal jamb - main body. Vertical frame element of access opening."
    sb(3, 0) = "N1-OVL": sb(3, 1) = "Over-lintel (N1)":    sb(3, 2) = "1":  sb(3, 3) = "Zone above lintel - decorative frieze area."
    sb(4, 0) = "N1-COR": sb(4, 1) = "Interbody cornice":   sb(4, 2) = "1":  sb(4, 3) = "Interbody cornice zone between superposed constructive bodies."
    sb(5, 0) = "N2-SPA": sb(5, 1) = "Spandrel (N2)":       sb(5, 2) = "2":  sb(5, 3) = "Lateral wall face - upper body (2-storey structures)."
    sb(6, 0) = "N2-JAM": sb(6, 1) = "Jamb (N2)":           sb(6, 2) = "2":  sb(6, 3) = "Portal jamb - upper body."
    sb(7, 0) = "ND":     sb(7, 1) = "Not determined":      sb(7, 2) = "-1": sb(7, 3) = "Position not determined."
    For i = 0 To 7
        db.Execute "INSERT INTO L_STRUCT_BODY (Code,Name,Body_Level,Description) VALUES ('" & sb(i, 0) & "','" & sb(i, 1) & "'," & sb(i, 2) & ",'" & sb(i, 3) & "')", dbFailOnError
    Next i

    ' L_DEC_TYPE
    Dim dt(10, 1) As String
    dt(0, 0) = "T-shaped niche":      dt(0, 1) = "Niche or bas-relief in T form. Vertical + horizontal element."
    dt(1, 0) = "T-shaped niche inv.": dt(1, 1) = "Inverted T niche/relief."
    dt(2, 0) = "L-shaped niche":      dt(2, 1) = "Niche or bas-relief in L form."
    dt(3, 0) = "L-shaped niche inv.": dt(3, 1) = "Inverted L niche/relief."
    dt(4, 0) = "Zigzag":              dt(4, 1) = "Zigzag or chevron motif."
    dt(5, 0) = "Stepped motif":       dt(5, 1) = "Stepped/staircase motif. Rare at DW and LP."
    dt(6, 0) = "Frieze / Greca":      dt(6, 1) = "Fretwork or repeating greca frieze. Common at La Petaca."
    dt(7, 0) = "Triangular motif":    dt(7, 1) = "Painted triangular/chevron pattern. Documented at DW (over-lintel zone)."
    dt(8, 0) = "Painted band":        dt(8, 1) = "Horizontal painted band (red/white). Interbody cornice zone."
    dt(9, 0) = "Square niche":        dt(9, 1) = "Square niches in series (hornacinas cuadradas)."
    dt(10, 0) = "ND":                 dt(10, 1) = "Decoration type not determined."
    For i = 0 To 10
        db.Execute "INSERT INTO L_DEC_TYPE (Name,Description) VALUES ('" & dt(i, 0) & "','" & dt(i, 1) & "')", dbFailOnError
    Next i

    Debug.Print "-> All lookups populated"
End Sub

' ================================================================
'  3. RELATIONSHIPS (20 total: 18 from v3 + 2 for T_CONNECTIONS)
' ================================================================
Private Sub CreateAllRelationships(db As DAO.Database)
    Dim rn(19) As String
    rn(0) = "REL_SIT_SEC":    rn(1) = "REL_SEC_STR":    rn(2) = "REL_SEC_GRP"
    rn(3) = "REL_TYP_STR":    rn(4) = "REL_SUP_STR":    rn(5) = "REL_STR_SELF"
    rn(6) = "REL_STA_STR":    rn(7) = "REL_MATSTA_STR": rn(8) = "REL_VM_STR"
    rn(9) = "REL_CM_STR":     rn(10) = "REL_GT_GRP":    rn(11) = "REL_GRP_STR"
    rn(12) = "REL_STR_DAT":   rn(13) = "REL_STR_IND":   rn(14) = "REL_STR_DEC"
    rn(15) = "REL_SB_DEC":    rn(16) = "REL_DT_DEC":    rn(17) = "REL_STR_AFEAT"
    rn(18) = "REL_STR_CONA":  rn(19) = "REL_STR_CONB"
    Dim n As Integer
    For n = 0 To 19
        On Error Resume Next: db.Relations.Delete rn(n): On Error GoTo 0
    Next n

    MkRel db, rn(0), "L_SITES", "ID", "L_SECTORS", "ID_Site", True, False
    MkRel db, rn(1), "L_SECTORS", "ID", "T_STRUCTURES", "ID_Sector", True, False
    MkRel db, rn(2), "L_SECTORS", "ID", "T_GROUPS", "ID_Sector", True, False
    MkRel db, rn(3), "L_TYPOLOGY", "ID", "T_STRUCTURES", "ID_Typology", False, False
    MkRel db, rn(4), "L_SUPPORT", "ID", "T_STRUCTURES", "ID_Support", False, False
    MkRel db, rn(5), "T_STRUCTURES", "ID", "T_STRUCTURES", "ID_Parent", False, True
    MkRel db, rn(6), "L_STATUS", "ID", "T_STRUCTURES", "ID_Arch_Status", False, False
    MkRel db, rn(7), "L_MATERIAL_STATUS", "ID", "T_STRUCTURES", "ID_Material_Status", False, False
    MkRel db, rn(8), "L_VOL_METHOD", "ID", "T_STRUCTURES", "ID_Vol_Method", False, False
    MkRel db, rn(9), "L_COORD_METHOD", "ID", "T_STRUCTURES", "ID_Coord_Method", False, False
    MkRel db, rn(10), "L_GROUP_TYPE", "ID", "T_GROUPS", "ID_Group_Type", False, False
    MkRel db, rn(11), "T_GROUPS", "ID", "T_STRUCTURES", "ID_Group", False, False
    MkRel db, rn(12), "T_STRUCTURES", "ID", "T_DATING", "ID_Structure", True, False
    MkRel db, rn(13), "T_STRUCTURES", "ID", "T_INDIVIDUALS", "ID_Structure", True, False
    MkRel db, rn(14), "T_STRUCTURES", "ID", "T_DECORATIONS", "ID_Structure", True, False
    MkRel db, rn(15), "L_STRUCT_BODY", "ID", "T_DECORATIONS", "ID_Struct_Body", False, False
    MkRel db, rn(16), "L_DEC_TYPE", "ID", "T_DECORATIONS", "ID_Dec_Type", False, False
    MkRel db, rn(17), "T_STRUCTURES", "ID", "T_ARCH_FEATURES", "ID_Structure", True, False
    ' v4: dual reference to T_STRUCTURES -> no enforced integrity (value 2),
    ' same treatment as the self-referencing REL_STR_SELF
    MkRel db, rn(18), "T_STRUCTURES", "ID", "T_CONNECTIONS", "ID_Struct_A", False, True
    MkRel db, rn(19), "T_STRUCTURES", "ID", "T_CONNECTIONS", "ID_Struct_B", False, True

    db.Relations.Refresh
    Debug.Print "-> 20 relationships OK"
End Sub

Private Sub MkRel(db As DAO.Database, nm As String, pT As String, pF As String, cT As String, cF As String, del As Boolean, noInt As Boolean)
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
    Set rel = db.CreateRelation(nm, pT, cT, fl)
    Set fld = rel.CreateField(pF)
    fld.ForeignName = cF
    rel.Fields.Append fld
    db.Relations.Append rel
    Exit Sub
ErrR: Debug.Print "  Warning " & nm & ": " & Err.Description
End Sub

' ================================================================
'  4. QUERIES (14 total)
'     QRY_01, QRY_03, QRY_04, QRY_06, QRY_08-QRY_12: unchanged
'     QRY_02: frieze count now from Relief_Frieze (element M)
'     QRY_05: renames + mortar + opening metrics
'     QRY_07: Height_Above_Base_m
'     QRY_13: rewritten - full 24-column A-X matrix (letter order)
'     QRY_14: NEW - connection edge list for network analysis (OE3)
' ================================================================
Private Sub CreateAllQueries(db As DAO.Database)
    Dim qn(13) As String
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
    qn(10) = "QRY_11_Masonry_by_Site"
    qn(11) = "QRY_12_Geology_Construction"
    qn(12) = "QRY_13_AX_Pattern_Export"
    qn(13) = "QRY_14_Connections_Edges"
    Dim i As Integer
    For i = 0 To 13
        If QueryExists(db, qn(i)) Then db.QueryDefs.Delete qn(i)
    Next i

    Dim q As String

    ' QRY_01 - Typology distribution by site
    q = "SELECT S.Site_Name, T.Name AS Typology, COUNT(E.ID) AS N "
    q = q & "FROM ((T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID) "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID "
    q = q & "GROUP BY S.Site_Name, T.Name ORDER BY S.Site_Name, T.Name;"
    db.CreateQueryDef qn(0), q

    ' QRY_02 - Decoration presence by site (chi-squared source)
    ' v4: N_Frieze from Relief_Frieze (BYTE, element M; counts value 1 only)
    q = "SELECT S.Site_Name, "
    q = q & "SUM(IIF(E.Dec_Square_Niche=True,1,0)) AS N_Niche, "
    q = q & "SUM(IIF(E.Dec_Relief_T=True,1,0)) AS N_T, "
    q = q & "SUM(IIF(E.Dec_Relief_L=True,1,0)) AS N_L, "
    q = q & "SUM(IIF(E.Dec_Zigzag=True,1,0)) AS N_Zigzag, "
    q = q & "SUM(IIF(E.Dec_Stepped=True,1,0)) AS N_Stepped, "
    q = q & "SUM(IIF(E.Relief_Frieze=1,1,0)) AS N_Frieze, "
    q = q & "SUM(IIF(E.Rock_Art=True,1,0)) AS N_RockArt, "
    q = q & "COUNT(E.ID) AS Total "
    q = q & "FROM (T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID "
    q = q & "GROUP BY S.Site_Name;"
    db.CreateQueryDef qn(1), q

    ' QRY_03 - Conservation status by sector
    q = "SELECT S.Site_Name, SC.Sector_Name, "
    q = q & "AS1.Name AS Arch_Status, MS.Name AS Material_Status, "
    q = q & "COUNT(*) AS N "
    q = q & "FROM ((((T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID) "
    q = q & "LEFT JOIN L_STATUS AS AS1 ON E.ID_Arch_Status=AS1.ID) "
    q = q & "LEFT JOIN L_MATERIAL_STATUS AS MS ON E.ID_Material_Status=MS.ID) "
    q = q & "GROUP BY S.Site_Name, SC.Sector_Name, AS1.Name, MS.Name "
    q = q & "ORDER BY S.Site_Name, SC.Sector_Name;"
    db.CreateQueryDef qn(2), q

    ' QRY_04 - Structures with C14 dating
    q = "SELECT E.Code, S.Site_Name, SC.Sector_Name, T.Name AS Typology, "
    q = q & "E.Chrono_Start_Cent, E.Chrono_End_Cent "
    q = q & "FROM (((T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID) "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID) "
    q = q & "WHERE E.C14=True ORDER BY E.Chrono_Start_Cent;"
    db.CreateQueryDef qn(3), q

    ' QRY_05 - Full flat export for R/SPSS
    ' v4: Relief_Frieze replaces Dec_Frieze; + mortar & opening metrics
    q = "SELECT E.ID, E.Code, S.Site_Name AS Site, SC.Sector_Name AS Sector, "
    q = q & "T.Name AS Typology, SU.Name AS Support, "
    q = q & "AS1.Name AS Arch_Status, MS.Name AS Material_Status, "
    q = q & "E.N_Floors, E.Floor_Plan, E.N_Built_Walls, E.Access_Orientation, "
    q = q & "E.Natural_Roof, E.Buttresses, E.Plastered, E.Rock_Painting, "
    q = q & "E.Base_Level, E.Corbelled_Platform, E.Timber_Brackets, "
    q = q & "E.Transverse_Beams, E.Corbelled_Courses, E.Embedded_Base_Beams, "
    q = q & "E.Tie_Walls, E.Access_Opening, E.Sill, E.Recessed_Portal, "
    q = q & "E.Structural_Pilasters, E.Eave, E.Upper_Crown, E.Corner_Quoins, "
    q = q & "E.Lateral_Wall_Faces, E.Relief_Frieze, E.Interbody_Cornice, "
    q = q & "E.Dec_Square_Niche, E.Dec_Relief_T, E.Dec_Relief_T_Inv, "
    q = q & "E.Dec_Relief_L, E.Dec_Relief_L_Inv, E.Dec_Zigzag, "
    q = q & "E.Dec_Stepped, E.Rock_Art, "
    q = q & "E.Looting, E.Fire_Damage, E.Animal_Activity, E.Modern_Access, "
    q = q & "E.Human_Remains, E.MNI, E.Mummification, E.Funerary_Bundles, "
    q = q & "E.Bone_Burning, E.Mat_Textiles, E.Mat_Ceramics, E.Mat_DeerAntler, "
    q = q & "E.C14, E.Chrono_Start_Cent, E.Chrono_End_Cent, "
    q = q & "E.Interior_Area_m2, E.Interior_Vol_m3, E.Total_Vol_m3, "
    q = q & "E.Coord_Lat_WGS84, E.Coord_Lon_WGS84, "
    q = q & "E.Coord_E_UTM, E.Coord_N_UTM, E.Altitude_masl, "
    q = q & "E.ChaXR_Documented, "
    q = q & "E.Support_Width_cm, E.Support_Depth_cm, "
    q = q & "E.Support_Morphology, E.Support_Modified, "
    q = q & "E.Masonry_Quality, E.Masonry_Type, "
    q = q & "E.Mortar_Present, E.Mortar_Type, E.Chinking_Stones, "
    q = q & "E.Opening_Width_cm, E.Opening_Height_cm, "
    q = q & "E.Height_Above_Base_m, "
    q = q & "E.Facade_Orientation, E.Visibility_Valley, "
    q = q & "E.Construction_Phases, E.Phase_Evidence "
    q = q & "FROM (((((T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID) "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID) "
    q = q & "INNER JOIN L_SUPPORT AS SU ON E.ID_Support=SU.ID) "
    q = q & "LEFT JOIN L_STATUS AS AS1 ON E.ID_Arch_Status=AS1.ID) "
    q = q & "LEFT JOIN L_MATERIAL_STATUS AS MS ON E.ID_Material_Status=MS.ID;"
    db.CreateQueryDef qn(4), q

    ' QRY_06 - Volumetry by typology
    q = "SELECT S.Site_Name, T.Name AS Typology, "
    q = q & "COUNT(E.ID) AS N_Total, COUNT(E.Interior_Vol_m3) AS N_with_Vol, "
    q = q & "AVG(E.Interior_Area_m2) AS Mean_Area, "
    q = q & "AVG(E.Interior_Vol_m3) AS Mean_Vol, "
    q = q & "MIN(E.Interior_Vol_m3) AS Min_Vol, MAX(E.Interior_Vol_m3) AS Max_Vol "
    q = q & "FROM ((T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID) "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID "
    q = q & "WHERE E.Interior_Vol_m3 IS NOT NULL "
    q = q & "GROUP BY S.Site_Name, T.Name ORDER BY S.Site_Name, Mean_Vol DESC;"
    db.CreateQueryDef qn(5), q

    ' QRY_07 - Spatial export for QGIS (v4: Height_Above_Base_m)
    q = "SELECT E.ID, E.Code, S.Site_Name, SC.Sector_Name, T.Name AS Typology, "
    q = q & "E.Coord_Lat_WGS84, E.Coord_Lon_WGS84, "
    q = q & "E.Coord_E_UTM, E.Coord_N_UTM, E.Altitude_masl, "
    q = q & "E.Height_Above_Base_m, E.Coord_Precision_m, "
    q = q & "AS1.Name AS Arch_Status, MS.Name AS Material_Status, "
    q = q & "E.N_Floors, E.Interior_Area_m2, E.Interior_Vol_m3, "
    q = q & "E.Chrono_Start_Cent, E.Chrono_End_Cent, "
    q = q & "E.Looting, E.Human_Remains, E.MNI, "
    q = q & "E.ChaXR_Documented, E.URL_3D, "
    q = q & "E.Facade_Orientation, E.Visibility_Valley, E.Masonry_Quality "
    q = q & "FROM ((((T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID) "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID) "
    q = q & "LEFT JOIN L_STATUS AS AS1 ON E.ID_Arch_Status=AS1.ID) "
    q = q & "LEFT JOIN L_MATERIAL_STATUS AS MS ON E.ID_Material_Status=MS.ID "
    q = q & "WHERE E.Coord_Lat_WGS84 IS NOT NULL ORDER BY S.Site_Name, E.Code;"
    db.CreateQueryDef qn(6), q

    ' QRY_08 - Children of a parent structure
    q = "SELECT E_c.Code AS Child_Code, T.Name AS Typology, "
    q = q & "E_c.Human_Remains, E_c.MNI "
    q = q & "FROM T_STRUCTURES AS E_c "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E_c.ID_Typology=T.ID "
    q = q & "WHERE E_c.ID_Parent=[Parent ID?] "
    q = q & "ORDER BY T.Name, E_c.Code;"
    db.CreateQueryDef qn(7), q

    ' QRY_09 - Members of a functional group
    q = "SELECT G.Group_Code, GT.Name AS Group_Type, E.Code, T.Name AS Typology, "
    q = q & "E.Coord_E_UTM, E.Coord_N_UTM, E.Altitude_masl "
    q = q & "FROM (((T_STRUCTURES AS E "
    q = q & "INNER JOIN T_GROUPS AS G ON E.ID_Group=G.ID) "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID) "
    q = q & "INNER JOIN L_GROUP_TYPE AS GT ON G.ID_Group_Type=GT.ID) "
    q = q & "WHERE G.Group_Code=[Group code?] "
    q = q & "ORDER BY E.Altitude_masl DESC;"
    db.CreateQueryDef qn(8), q

    ' QRY_10 - ChaXR documentation coverage
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

    ' QRY_11 - Masonry quality by site and typology (H01/H04)
    q = "SELECT S.Site_Name, T.Name AS Typology, "
    q = q & "E.Masonry_Quality, E.Masonry_Type, COUNT(E.ID) AS N "
    q = q & "FROM ((T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID) "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID "
    q = q & "WHERE E.Masonry_Quality IS NOT NULL "
    q = q & "GROUP BY S.Site_Name, T.Name, E.Masonry_Quality, E.Masonry_Type "
    q = q & "ORDER BY S.Site_Name, T.Name, E.Masonry_Quality;"
    db.CreateQueryDef qn(10), q

    ' QRY_12 - Geology-construction relationship (H02)
    q = "SELECT S.Site_Name, T.Name AS Typology, "
    q = q & "SU.Name AS Support_Type, E.Support_Morphology, "
    q = q & "E.Support_Modified, COUNT(E.ID) AS N, "
    q = q & "AVG(E.Support_Width_cm) AS Avg_Width_cm, "
    q = q & "AVG(E.Support_Depth_cm) AS Avg_Depth_cm "
    q = q & "FROM (((T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID) "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID) "
    q = q & "LEFT JOIN L_SUPPORT AS SU ON E.ID_Support=SU.ID "
    q = q & "GROUP BY S.Site_Name, T.Name, SU.Name, E.Support_Morphology, E.Support_Modified "
    q = q & "ORDER BY S.Site_Name, T.Name;"
    db.CreateQueryDef qn(11), q

    ' QRY_13 - A-X pattern export for R (v4: full 24-column matrix)
    ' Letter-ordered vector AX_A..AX_X with 0/1/9 values.
    ' AX_Q derived from Lintel material: Absent->0, ND/Null->9, else->1.
    ' In R: filter or weight cells with value 9 (not observable) before
    ' computing phi/Jaccard, hierarchical cluster, Moran I or CA.
    ' LEFT JOIN on L_TYPOLOGY: includes structures without typology.
    q = "SELECT E.ID, E.Code, S.Site_Name AS Site, SC.Sector_Name AS Sector, "
    q = q & "T.Name AS Typology, "
    q = q & "E.Coord_E_UTM, E.Coord_N_UTM, E.Altitude_masl, E.ID_Group, "
    q = q & "E.Embedded_Base_Beams AS AX_A, "
    q = q & "E.Base_Level AS AX_B, "
    q = q & "E.Decorative_Socle AS AX_C, "
    q = q & "E.Tie_Walls AS AX_D, "
    q = q & "E.Timber_Brackets AS AX_E, "
    q = q & "E.Transverse_Beams AS AX_F, "
    q = q & "E.Corbelled_Courses AS AX_G, "
    q = q & "E.Corbelled_Platform AS AX_H, "
    q = q & "E.Interbody_Cornice AS AX_I, "
    q = q & "E.Corner_Quoins AS AX_J, "
    q = q & "E.Structural_Pilasters AS AX_K, "
    q = q & "E.Lateral_Wall_Faces AS AX_L, "
    q = q & "E.Relief_Frieze AS AX_M, "
    q = q & "E.Sill AS AX_N, "
    q = q & "E.Jambs AS AX_O, "
    q = q & "E.Access_Opening AS AX_P, "
    q = q & "IIF(E.Lintel='Absent',0,IIF(E.Lintel Is Null Or E.Lintel='ND',9,1)) AS AX_Q, "
    q = q & "E.Upper_Crown AS AX_R, "
    q = q & "E.Eave_Beam AS AX_S, "
    q = q & "E.Eave_Surface AS AX_T, "
    q = q & "E.Eave AS AX_U, "
    q = q & "E.Lateral_Walls AS AX_V, "
    q = q & "E.Rear_Wall AS AX_W, "
    q = q & "E.Chamber_Roof AS AX_X, "
    q = q & "E.Timber_Bracket_Count, E.Corbel_Material, "
    q = q & "E.Interbody_Cornice_Material, E.Lintel, "
    q = q & "E.Rear_Wall_Built, E.Recessed_Portal, "
    q = q & "E.Masonry_Quality, E.Masonry_Type, "
    q = q & "E.Mortar_Present, E.Chinking_Stones, "
    q = q & "E.Support_Morphology, E.Support_Modified "
    q = q & "FROM ((T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID) "
    q = q & "LEFT JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID "
    q = q & "ORDER BY S.Site_Name, SC.Sector_Name, E.Code;"
    db.CreateQueryDef qn(12), q

    ' QRY_14 - NEW v4: connection edge list (OE3 aerial network)
    ' Exports the adjacency list with coordinates of both endpoints,
    ' ready for igraph (R) or line generation in QGIS.
    q = "SELECT C.ID, A.Code AS Code_A, B.Code AS Code_B, "
    q = q & "C.Connection_Type, C.Confidence, "
    q = q & "A.Coord_E_UTM AS E_UTM_A, A.Coord_N_UTM AS N_UTM_A, "
    q = q & "A.Altitude_masl AS Alt_A, "
    q = q & "B.Coord_E_UTM AS E_UTM_B, B.Coord_N_UTM AS N_UTM_B, "
    q = q & "B.Altitude_masl AS Alt_B, "
    q = q & "C.Notes "
    q = q & "FROM (T_CONNECTIONS AS C "
    q = q & "INNER JOIN T_STRUCTURES AS A ON C.ID_Struct_A=A.ID) "
    q = q & "INNER JOIN T_STRUCTURES AS B ON C.ID_Struct_B=B.ID "
    q = q & "ORDER BY C.Connection_Type, A.Code;"
    db.CreateQueryDef qn(13), q

    Debug.Print "-> 14 queries OK"
End Sub
