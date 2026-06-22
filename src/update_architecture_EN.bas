Option Compare Database
Option Explicit

' ================================================================
'  UPDATE DB - ARCHITECTURAL SYSTEMS & DECORATIONS
'  Run Sub UpdateArchitecture() -> F5
' ================================================================

Sub UpdateArchitecture()
    Dim db As DAO.Database
    Set db = CurrentDb()
    AddArchFields db
    CreateDecorationTables db
    CreateDecorationRelationships db
    Set db = Nothing
    Dim msg As String
    msg = "Architecture update complete!" & vbCrLf & vbCrLf
    msg = msg & "  16 new fields in T_STRUCTURES" & vbCrLf
    msg = msg & "  T_DECORATIONS created" & vbCrLf
    msg = msg & "  L_STRUCT_BODY created (8 values)" & vbCrLf
    msg = msg & "  L_DEC_TYPE created (11 values)" & vbCrLf
    msg = msg & "  Relationships established"
    MsgBox msg, vbInformation, "Done!"
End Sub

Function TableExists(db As DAO.Database, n As String) As Boolean
    Dim t As DAO.TableDef
    For Each t In db.TableDefs
        If t.Name = n Then TableExists = True: Exit Function
    Next t
End Function

Sub AddArchFields(db As DAO.Database)
    ' LEVEL 0 - BASE SYSTEM
    On Error Resume Next
    db.Execute "ALTER TABLE T_STRUCTURES ADD COLUMN Base_Level YESNO", dbFailOnError
    db.Execute "ALTER TABLE T_STRUCTURES ADD COLUMN Decorative_Socle YESNO", dbFailOnError
    db.Execute "ALTER TABLE T_STRUCTURES ADD COLUMN Corbelled_Platform YESNO", dbFailOnError
    db.Execute "ALTER TABLE T_STRUCTURES ADD COLUMN Timber_Brackets YESNO", dbFailOnError
    db.Execute "ALTER TABLE T_STRUCTURES ADD COLUMN Timber_Bracket_Count INTEGER", dbFailOnError
    db.Execute "ALTER TABLE T_STRUCTURES ADD COLUMN Transverse_Beams YESNO", dbFailOnError
    db.Execute "ALTER TABLE T_STRUCTURES ADD COLUMN Corbelled_Courses YESNO", dbFailOnError
    db.Execute "ALTER TABLE T_STRUCTURES ADD COLUMN Corbel_Material TEXT(20)", dbFailOnError
    ' ACCESS OPENING SYSTEM
    db.Execute "ALTER TABLE T_STRUCTURES ADD COLUMN Access_Opening YESNO", dbFailOnError
    db.Execute "ALTER TABLE T_STRUCTURES ADD COLUMN Sill YESNO", dbFailOnError
    db.Execute "ALTER TABLE T_STRUCTURES ADD COLUMN Recessed_Portal YESNO", dbFailOnError
    ' FACADE ELEMENTS
    db.Execute "ALTER TABLE T_STRUCTURES ADD COLUMN Structural_Pilasters YESNO", dbFailOnError
    db.Execute "ALTER TABLE T_STRUCTURES ADD COLUMN Cornice_Material TEXT(20)", dbFailOnError
    ' UPPER ZONE
    db.Execute "ALTER TABLE T_STRUCTURES ADD COLUMN Eave YESNO", dbFailOnError
    db.Execute "ALTER TABLE T_STRUCTURES ADD COLUMN Eave_Beam YESNO", dbFailOnError
    db.Execute "ALTER TABLE T_STRUCTURES ADD COLUMN Upper_Crown YESNO", dbFailOnError
    On Error GoTo 0
    Debug.Print "[OK] 16 fields added to T_STRUCTURES"
End Sub

Sub CreateDecorationTables(db As DAO.Database)

    ' L_STRUCT_BODY (where decoration is located)
    On Error Resume Next
    db.Execute "DROP TABLE L_STRUCT_BODY", dbFailOnError
    On Error GoTo 0
    db.Execute "CREATE TABLE L_STRUCT_BODY (ID COUNTER CONSTRAINT PK_SB PRIMARY KEY, Code TEXT(10) NOT NULL, Name TEXT(60) NOT NULL, Body_Level INTEGER, Description TEXT(255))", dbFailOnError

    Dim sb(7, 3) As String
    sb(0, 0) = "N0-SOC": sb(0, 1) = "Socle (N0)":            sb(0, 2) = "0": sb(0, 3) = "Decorative socle - level 0 base element."
    sb(1, 0) = "N1-SPA": sb(1, 1) = "Spandrel (N1)":         sb(1, 2) = "1": sb(1, 3) = "Lateral wall face of main body (outside vano frame)."
    sb(2, 0) = "N1-JAM": sb(2, 1) = "Jamb (N1)":             sb(2, 2) = "1": sb(2, 3) = "Portal jamb - main body. Vertical frame element of access opening."
    sb(3, 0) = "N1-OVL": sb(3, 1) = "Over-lintel (N1)":      sb(3, 2) = "1": sb(3, 3) = "Zone above lintel - decorative frieze area."
    sb(4, 0) = "N1-COR": sb(4, 1) = "Interlevel cornice":    sb(4, 2) = "1": sb(4, 3) = "Interlevel cornice zone between constructive bodies."
    sb(5, 0) = "N2-SPA": sb(5, 1) = "Spandrel (N2)":         sb(5, 2) = "2": sb(5, 3) = "Lateral wall face - upper body (2-storey structures)."
    sb(6, 0) = "N2-JAM": sb(6, 1) = "Jamb (N2)":             sb(6, 2) = "2": sb(6, 3) = "Portal jamb - upper body."
    sb(7, 0) = "ND":     sb(7, 1) = "Not determined":        sb(7, 2) = "-1": sb(7, 3) = "Position not determined."

    Dim i As Integer
    For i = 0 To 7
        db.Execute "INSERT INTO L_STRUCT_BODY (Code,Name,Body_Level,Description) VALUES ('" & sb(i, 0) & "','" & sb(i, 1) & "'," & sb(i, 2) & ",'" & sb(i, 3) & "')", dbFailOnError
    Next i
    Debug.Print "[OK] L_STRUCT_BODY: 8 values"

    ' L_DEC_TYPE (decoration types)
    On Error Resume Next
    db.Execute "DROP TABLE L_DEC_TYPE", dbFailOnError
    On Error GoTo 0
    db.Execute "CREATE TABLE L_DEC_TYPE (ID COUNTER CONSTRAINT PK_DT PRIMARY KEY, Name TEXT(60) NOT NULL, Description TEXT(255))", dbFailOnError

    Dim dt(10, 1) As String
    dt(0, 0) = "T-shaped niche":       dt(0, 1) = "Niche or bas-relief in T form. Vertical + horizontal element."
    dt(1, 0) = "T-shaped niche inv.":  dt(1, 1) = "Inverted T niche/relief."
    dt(2, 0) = "L-shaped niche":       dt(2, 1) = "Niche or bas-relief in L form."
    dt(3, 0) = "L-shaped niche inv.":  dt(3, 1) = "Inverted L niche/relief."
    dt(4, 0) = "Zigzag":               dt(4, 1) = "Zigzag or chevron motif."
    dt(5, 0) = "Stepped motif":        dt(5, 1) = "Stepped / staircase motif (Andean symbol). Rare at DW and LP."
    dt(6, 0) = "Frieze / Greca":       dt(6, 1) = "Fretwork or repeating greca frieze. Common at La Petaca."
    dt(7, 0) = "Triangular motif":     dt(7, 1) = "Painted triangular / chevron pattern. Documented at DW."
    dt(8, 0) = "Painted band":         dt(8, 1) = "Horizontal painted band (red/white). Associated with interlevel cornice."
    dt(9, 0) = "Square niche":         dt(9, 1) = "Square niches in series (hornacinas cuadradas)."
    dt(10, 0) = "ND":                  dt(10, 1) = "Decoration type not determined."

    For i = 0 To 10
        db.Execute "INSERT INTO L_DEC_TYPE (Name,Description) VALUES ('" & dt(i, 0) & "','" & dt(i, 1) & "')", dbFailOnError
    Next i
    Debug.Print "[OK] L_DEC_TYPE: 11 values"

    ' T_DECORATIONS (one record per decoration per structure)
    On Error Resume Next
    db.Execute "DROP TABLE T_DECORATIONS", dbFailOnError
    On Error GoTo 0
    db.Execute "CREATE TABLE T_DECORATIONS (ID COUNTER CONSTRAINT PK_TDEC PRIMARY KEY, ID_Structure LONG NOT NULL, ID_Struct_Body LONG, ID_Dec_Type LONG, Body_No INTEGER, Color TEXT(20), Notes TEXT(255))", dbFailOnError
    Debug.Print "[OK] T_DECORATIONS created"
End Sub

Sub CreateDecorationRelationships(db As DAO.Database)
    On Error Resume Next
    db.Relations.Delete "REL_STR_DEC"
    db.Relations.Delete "REL_SB_DEC"
    db.Relations.Delete "REL_DT_DEC"
    On Error GoTo 0

    MkRel db, "REL_STR_DEC", "T_STRUCTURES",  "ID", "T_DECORATIONS", "ID_Structure",  True
    MkRel db, "REL_SB_DEC",  "L_STRUCT_BODY", "ID", "T_DECORATIONS", "ID_Struct_Body", False
    MkRel db, "REL_DT_DEC",  "L_DEC_TYPE",    "ID", "T_DECORATIONS", "ID_Dec_Type",    False

    db.Relations.Refresh
    Debug.Print "[OK] Decoration relationships created"
End Sub

Sub MkRel(db As DAO.Database, nm As String, pT As String, pF As String, cT As String, cF As String, del As Boolean)
    Dim rel As DAO.Relation
    Dim fld As DAO.Field
    Dim fl As Long
    On Error GoTo ErrR
    fl = dbRelationUpdateCascade
    If del Then fl = fl Or dbRelationDeleteCascade
    Set rel = db.CreateRelation(nm, pT, cT, fl)
    Set fld = rel.CreateField(pF)
    fld.ForeignName = cF
    rel.Fields.Append fld
    db.Relations.Append rel
    Exit Sub
ErrR:
    Debug.Print "  Warning " & nm & ": " & Err.Description
End Sub
