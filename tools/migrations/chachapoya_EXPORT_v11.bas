Option Compare Database
Option Explicit

' ================================================================
'  CHACHAPOYA DB - STEP 1 OF 2: EXPORT FROM THE v11 SOURCE
'  Esteve Ribera Torro | TFM Arqueologia UA
'
'  RUN THIS ON v11_migrated.accdb (the SOURCE). Import a new blank
'  module, paste this, and run Sub ExportV11().
'
'  It writes five CSV files to <db folder>\transfer\ and touches
'  nothing in the database. Inspect and edit the files, then run
'  chachapoya_IMPORT_v15.bas on the v15 database.
'
'  WHAT IT WRITES
'    EXP_STRUCTURES.csv   35 rows, one per structure. Column names
'                         are v15 FIELD NAMES; lookups are exported
'                         as NAMES, never as IDs, because the IDs
'                         of L_STRUCT_BODY, L_SUPPORT, L_TYPOLOGY,
'                         L_MATERIAL_STATUS and L_GROUP_TYPE have
'                         MOVED between v11 and v15. Columns ending
'                         in _v11 are context for the reviewer and
'                         are ignored by the import.
'    EXP_DECORATIONS.csv  29 rows. Parent by Code, position by the
'                         L_STRUCT_BODY Code, type by name.
'    EXP_ROCKART_PROPOSED.csv  proposed ROC rows from the RA_*
'                         booleans. Position left BLANK on purpose:
'                         it needs the association test (delta 7.3).
'    EXP_LOST_PROPOSED.csv     proposed T_LOST_ELEMENTS rows from
'                         Lost_Body_Evidence. Element left BLANK.
'    EXP_DROPPED.csv      every value NOT transferred, one row per
'                         value, with the rule that dropped it. This
'                         is the reference sheet for the rule 17
'                         worklist.
'
'  TRANSFER RULES IMPLEMENTED (decisions D1-D8)
'    D1  the 20 A-X element fields: 0 -> 0, 9 -> NULL, 1 -> NULL.
'        v11 meant "present"; v15 value 1 means "present COMPLETE"
'        and 1/2/3 is a gradient nobody has assessed yet.
'    D2  value 9 in every observational field -> NULL. Those are the
'        v10 default, not judgements (delta 7bis finding a).
'    D3  Doc_Basis: only Direct access survives, same label and same
'        criterion in both scales. The rest -> NULL.
'    D4  Corbelled_Platform / Access_Opening / Eave / Rear_Wall:
'        the fields are gone in v15 and the 0s and 9s were already
'        converted into Sys_* by the previous migration. The 1s that
'        were left behind -> NULL, value kept as context.
'    D5  Pigment_Color "White, Red" / "Red, White" -> "Both" (still
'        in the v15 domain: 7.6 retired Both from T_DECORATIONS,
'        not from here). Plaster_Color "White, Red" -> NULL, there
'        is no secondary colour field for plaster.
'    D6  RockArt_Present from Rock_Art (same domain, same meaning),
'        plus proposed ROC rows from the RA_* booleans.
'    D7  Dec_Present = 1 where T_DECORATIONS rows exist, NULL
'        elsewhere. No row cannot tell 0 from 9.
'    D8  Lost_Body_Evidence -> proposed T_LOST_ELEMENTS rows,
'        scope Body, evidence type resolved by name.
'    B   ID_Support / ID_Support_Secondary = Fissure/Crack -> NULL
'        (the value was split into Vertical cleft and Bedding-plane
'        recess). ID_Support_Secondary = ND -> NULL. Absent in
'        ID_Material_Status -> Cultural_Materials_Present = 0.
'        _cm -> _m divides by 100. Pigment_Extent "Arch. elements"
'        -> "Architectural elements". Phase_Evidence "Adjacent" is
'        not in the v15 domain -> NULL.
'
'  Every field of the source is accounted for: transferred,
'  converted, dropped with a rule, or empty in all 35 records.
' ================================================================

' Leave empty to use <current database folder>\transfer\
Private Const EXPORT_DIR As String = ""

' Semicolon: opens straight into Excel under a Spanish locale.
Private Const SEP As String = ";"

' Provenance note written into the Notes of the proposed rows.
' Blank it if you would rather they arrive clean.
Private Const NOTE_RA As String = "Proposta automatica des dels booleans RA_* v11. Completar posicio."
Private Const NOTE_LOST As String = "Proposta automatica des de Lost_Body_Evidence v11. Completar element."

Private gDrop As String
Private gLog As String
Private gDropN As Long


' ================================================================
'  MAIN
' ================================================================
Public Sub ExportV11()
    Dim db As DAO.Database
    Dim d As String
    Dim n1 As Long
    Dim n2 As Long
    Dim n3 As Long
    Dim n4 As Long

    gDrop = ""
    gLog = ""
    gDropN = 0

    Set db = CurrentDb
    d = TargetDir()
    If Len(d) = 0 Then Exit Sub

    Say "CHACHAPOYA EXPORT v11 -> v15"
    Say "Source: " & CurrentProject.Name
    Say "Folder: " & d
    Say ""

    If Not CheckSource(db) Then
        MsgBox "The source does not look like the v11 schema. See the Immediate window.", vbCritical
        Exit Sub
    End If

    n1 = ExportStructures(db, d)
    n2 = ExportDecorations(db, d)
    n3 = ExportRockArt(db, d)
    n4 = ExportLost(db, d)

    WriteDropFile d
    CheckCoherence db

    Say ""
    Say "SUMMARY"
    Say "  EXP_STRUCTURES.csv .......... " & n1 & " rows"
    Say "  EXP_DECORATIONS.csv ......... " & n2 & " rows"
    Say "  EXP_ROCKART_PROPOSED.csv .... " & n3 & " rows"
    Say "  EXP_LOST_PROPOSED.csv ....... " & n4 & " rows"
    Say "  EXP_DROPPED.csv ............. " & gDropN & " values not transferred"
    Say ""
    Say "Next: open the files, review, then run ImportV15 on the v15 database."

    WriteLogFile d

    Dim m As String
    m = "Export finished." & vbCrLf & vbCrLf
    m = m & "Structures: " & n1 & vbCrLf
    m = m & "Decorations: " & n2 & vbCrLf
    m = m & "Proposed rock art rows: " & n3 & vbCrLf
    m = m & "Proposed lost element rows: " & n4 & vbCrLf
    m = m & "Values dropped: " & gDropN & vbCrLf & vbCrLf
    m = m & "Folder: " & d & vbCrLf
    m = m & "Detail in EXPORT_LOG.txt and in the Immediate window."
    MsgBox m, vbInformation
End Sub


' ================================================================
'  1. STRUCTURES
' ================================================================
Private Function ExportStructures(db As DAO.Database, ByVal d As String) As Long
    Dim rs As DAO.Recordset
    Dim rsD As DAO.Recordset
    Dim ff As Integer
    Dim hdr As String
    Dim ln As String
    Dim i As Long
    Dim n As Long
    Dim cd As String
    Dim v As Variant
    Dim s As String
    Dim el() As String
    Dim th() As String
    Dim dr() As String
    Dim sys() As String

    FillElementFields el
    FillThreeFields th
    FillDirectFields dr
    FillRetiredSystemFields sys

    ' ---------- header ----------
    hdr = ""
    Ap hdr, "Code"
    Ap hdr, "Sector_Name"
    Ap hdr, "Typology_Name"
    Ap hdr, "Support_Name"
    Ap hdr, "Support_Secondary_Name"
    Ap hdr, "Parent_Code"
    Ap hdr, "Group_Code"
    Ap hdr, "Arch_Status_Name"
    Ap hdr, "Material_Status_Name"
    Ap hdr, "Cultural_Materials_Present"
    Ap hdr, "Coord_Method_Name"
    Ap hdr, "Campaign_Code"
    Ap hdr, "Doc_Basis"
    Ap hdr, "Phase_Evidence"
    Ap hdr, "Pigment_Color"
    Ap hdr, "Plaster_Color"
    Ap hdr, "Pigment_Extent"
    Ap hdr, "Opening_Width_m"
    Ap hdr, "Opening_Height_m"
    Ap hdr, "Support_Width_m"
    Ap hdr, "Support_Depth_m"
    Ap hdr, "Dec_Present"
    Ap hdr, "RockArt_Present"
    For i = 0 To UBound(dr)
        Ap hdr, dr(i)
    Next i
    For i = 0 To UBound(th)
        Ap hdr, th(i)
    Next i
    For i = 0 To UBound(el)
        Ap hdr, el(i)
    Next i
    ' context columns, ignored by the import
    Ap hdr, "ID_v11"
    For i = 0 To UBound(el)
        Ap hdr, el(i) & "_v11"
    Next i
    For i = 0 To UBound(sys)
        Ap hdr, sys(i) & "_v11"
    Next i
    Ap hdr, "Doc_Basis_v11"
    Ap hdr, "Lost_Body_Evidence_v11"

    ff = FreeFile
    Open d & "EXP_STRUCTURES.csv" For Output As #ff
    Print #ff, Mid(hdr, 2)

    Set rs = db.OpenRecordset("SELECT * FROM T_STRUCTURES ORDER BY Code", dbOpenSnapshot)
    Do While Not rs.EOF
        cd = Nz(rs!Code, "")
        ln = ""

        Ap ln, Q(cd)
        Ap ln, Q(Nz(NameById(db, "L_SECTORS", "Sector_Name", rs!ID_Sector), ""))
        Ap ln, Q(Nz(NameById(db, "L_TYPOLOGY", "Name", rs!ID_Typology), ""))

        ' --- support: Fissure/Crack was split in v15, ND as secondary means nothing
        s = Nz(NameById(db, "L_SUPPORT", "Name", rs!ID_Support), "")
        If s = "Fissure/Crack" Then
            Drop cd, "ID_Support", s, "C: value split into Vertical cleft / Bedding-plane recess"
            s = ""
        End If
        Ap ln, Q(s)

        s = Nz(NameById(db, "L_SUPPORT", "Name", rs!ID_Support_Secondary), "")
        If s = "Fissure/Crack" Then
            Drop cd, "ID_Support_Secondary", s, "C: value split into Vertical cleft / Bedding-plane recess"
            s = ""
        ElseIf s = "ND" Then
            Drop cd, "ID_Support_Secondary", s, "C: ND as a secondary support means nothing (delta 7bis d)"
            s = ""
        End If
        Ap ln, Q(s)

        Ap ln, Q("")
        Ap ln, Q("")
        Ap ln, Q(Nz(NameById(db, "L_STATUS", "Name", rs!ID_Arch_Status), ""))

        ' --- material status split into presence + condition (delta 12.3)
        s = Nz(NameById(db, "L_MATERIAL_STATUS", "Name", rs!ID_Material_Status), "")
        If s = "Absent" Then
            Ap ln, Q("")
            Ap ln, "0"
            Drop cd, "ID_Material_Status", s, "B: Absent -> Cultural_Materials_Present = 0, status NULL"
        ElseIf s = "Good" Or s = "Fair" Or s = "Poor" Then
            Ap ln, Q(s)
            Ap ln, "1"
        ElseIf s = "ND" Then
            Ap ln, Q(s)
            Ap ln, ""
        Else
            Ap ln, Q("")
            Ap ln, ""
        End If

        Ap ln, Q(Nz(NameById(db, "L_COORD_METHOD", "Name", rs!ID_Coord_Method), ""))
        Ap ln, Q(Nz(NameById(db, "L_CAMPAIGN", "Code", rs!ID_Campaign), ""))

        ' --- Doc_Basis: only Direct access maps (D3)
        s = Nz(rs!Doc_Basis, "")
        If Len(s) > 0 And s <> "Direct access" Then
            Drop cd, "Doc_Basis", s, "C: the v11 scale does not map onto the ordinal one (12.2)"
            Ap ln, Q("")
        Else
            Ap ln, Q(s)
        End If

        ' --- Phase_Evidence: Adjacent is not in the v15 domain
        s = Nz(rs!Phase_Evidence, "")
        If s = "Adjacent" Then
            Drop cd, "Phase_Evidence", s, "C: not in the v15 domain - abutment belongs in T_CONNECTIONS"
            s = ""
        End If
        Ap ln, Q(s)

        ' --- colours (D5)
        s = Nz(rs!Pigment_Color, "")
        If s = "White, Red" Or s = "Red, White" Then
            Drop cd, "Pigment_Color", s, "B: converted to Both"
            s = "Both"
        End If
        Ap ln, Q(s)

        s = Nz(rs!Plaster_Color, "")
        If InStr(s, ",") > 0 Then
            Drop cd, "Plaster_Color", s, "C: no combined value and no secondary colour field for plaster"
            s = ""
        End If
        Ap ln, Q(s)

        ' --- Pigment_Extent: stray value normalised (delta 7bis b)
        s = Nz(rs!Pigment_Extent, "")
        If s = "Arch. elements" Then
            Drop cd, "Pigment_Extent", s, "B: normalised to Architectural elements"
            s = "Architectural elements"
        End If
        Ap ln, Q(s)

        ' --- cm to m
        Ap ln, Cm2M(rs!Opening_Width_cm)
        Ap ln, Cm2M(rs!Opening_Height_cm)
        Ap ln, Cm2M(rs!Support_Width_cm)
        Ap ln, Cm2M(rs!Support_Depth_cm)

        ' --- Dec_Present derived from the existence of rows (D7)
        Set rsD = db.OpenRecordset("SELECT COUNT(*) AS N FROM T_DECORATIONS WHERE ID_Structure=" & rs!ID, dbOpenSnapshot)
        If rsD!N > 0 Then
            Ap ln, "1"
        Else
            Ap ln, ""
        End If
        rsD.Close

        ' --- RockArt_Present from Rock_Art, three-value rule (D6)
        Ap ln, ThreeVal(rs!Rock_Art, cd, "Rock_Art")

        ' --- plain copies
        For i = 0 To UBound(dr)
            Ap ln, Fmt(rs.Fields(dr(i)).Value)
        Next i

        ' --- three-value observational fields (D2)
        For i = 0 To UBound(th)
            Ap ln, ThreeVal(rs.Fields(th(i)).Value, cd, th(i))
        Next i

        ' --- the 20 A-X elements (D1 + D2)
        For i = 0 To UBound(el)
            Ap ln, ElemVal(rs.Fields(el(i)).Value, cd, el(i))
        Next i

        ' --- context columns
        Ap ln, Fmt(rs!ID)
        For i = 0 To UBound(el)
            Ap ln, Fmt(rs.Fields(el(i)).Value)
        Next i
        For i = 0 To UBound(sys)
            v = rs.Fields(sys(i)).Value
            If Not IsNull(v) Then
                If v = 1 Then
                    Drop cd, sys(i), "1", "D4: field retired in v15, the 0s and 9s were already converted into Sys_*"
                End If
            End If
            Ap ln, Fmt(v)
        Next i
        Ap ln, Q(Nz(rs!Doc_Basis, ""))
        s = Nz(rs!Lost_Body_Evidence, "")
        If Len(s) > 0 Then
            Drop cd, "Lost_Body_Evidence", s, "C: field retired - see EXP_LOST_PROPOSED.csv"
        End If
        Ap ln, Q(s)

        ' --- decoration booleans: superseded by T_DECORATIONS, logged only
        DropDecBooleans rs, cd

        Print #ff, Mid(ln, 2)
        n = n + 1
        rs.MoveNext
    Loop
    rs.Close
    Close #ff

    Say "EXP_STRUCTURES.csv written: " & n & " rows"
    ExportStructures = n
End Function


' ================================================================
'  2. DECORATIONS
' ================================================================
Private Function ExportDecorations(db As DAO.Database, ByVal d As String) As Long
    Dim rs As DAO.Recordset
    Dim ff As Integer
    Dim hd As String
    Dim sql As String
    Dim ln As String
    Dim n As Long

    sql = "SELECT E.Code AS Parent_Code, B.Code AS Body_Code, "
    sql = sql & "T.Name AS Dec_Type_Name, D.Body_No, D.Color, "
    sql = sql & "D.Substrate, D.Notes, D.ID AS ID_v11 "
    sql = sql & "FROM ((T_DECORATIONS AS D "
    sql = sql & "INNER JOIN T_STRUCTURES AS E ON D.ID_Structure=E.ID) "
    sql = sql & "LEFT JOIN L_STRUCT_BODY AS B ON D.ID_Struct_Body=B.ID) "
    sql = sql & "LEFT JOIN L_DEC_TYPE AS T ON D.ID_Dec_Type=T.ID "
    sql = sql & "ORDER BY E.Code, D.ID"

    ff = FreeFile
    Open d & "EXP_DECORATIONS.csv" For Output As #ff
    hd = "Parent_Code" & SEP & "Body_Code" & SEP & "Dec_Type_Name" & SEP & "Body_No"
    hd = hd & SEP & "Color" & SEP & "Color_Secondary" & SEP & "Substrate" & SEP & "Notes" & SEP & "ID_v11"
    Print #ff, hd

    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    Do While Not rs.EOF
        ln = ""
        Ap ln, Q(Nz(rs!Parent_Code, ""))
        Ap ln, Q(Nz(rs!Body_Code, ""))
        Ap ln, Q(Nz(rs!Dec_Type_Name, ""))
        Ap ln, Fmt(rs!Body_No)
        Ap ln, Q(Nz(rs!Color, ""))
        Ap ln, Q("")
        Ap ln, Q(Nz(rs!Substrate, ""))
        Ap ln, Q(Nz(rs!Notes, ""))
        Ap ln, Fmt(rs!ID_v11)
        Print #ff, Mid(ln, 2)
        n = n + 1
        rs.MoveNext
    Loop
    rs.Close
    Close #ff

    Say "EXP_DECORATIONS.csv written: " & n & " rows"
    Say "  positions exported as L_STRUCT_BODY Code, never as ID: the IDs moved"
    ExportDecorations = n
End Function


' ================================================================
'  3. PROPOSED ROCK ART ROWS
' ================================================================
Private Function ExportRockArt(db As DAO.Database, ByVal d As String) As Long
    Dim rs As DAO.Recordset
    Dim ff As Integer
    Dim hd As String
    Dim n As Long
    Dim cd As String

    ff = FreeFile
    Open d & "EXP_ROCKART_PROPOSED.csv" For Output As #ff
    hd = "Parent_Code" & SEP & "Body_Code" & SEP & "Dec_Type_Name" & SEP & "Body_No"
    hd = hd & SEP & "Color" & SEP & "Color_Secondary" & SEP & "Substrate" & SEP & "Notes"
    Print #ff, hd

    Set rs = db.OpenRecordset("SELECT * FROM T_STRUCTURES ORDER BY Code", dbOpenSnapshot)
    Do While Not rs.EOF
        cd = Nz(rs!Code, "")
        If Nz(rs!RA_Anthropomorphic, 0) = 1 Then
            RaRow ff, cd, "RA Anthropomorphic"
            n = n + 1
        End If
        If Nz(rs!RA_Zoomorphic, 0) = 1 Then
            RaRow ff, cd, "RA Zoomorphic"
            n = n + 1
        End If
        If Nz(rs!RA_Geometric, 0) = 1 Then
            RaRow ff, cd, "RA Geometric"
            n = n + 1
        End If
        If Nz(rs!RA_Abstract, 0) = 1 Then
            RaRow ff, cd, "RA Abstract"
            n = n + 1
        End If
        rs.MoveNext
    Loop
    rs.Close
    Close #ff

    Say "EXP_ROCKART_PROPOSED.csv written: " & n & " rows"
    Say "  Body_Code left BLANK on purpose: the position needs the association test (7.3)"
    ExportRockArt = n
End Function

Private Sub RaRow(ByVal ff As Integer, ByVal cd As String, ByVal tp As String)
    Dim ln As String
    ln = ""
    Ap ln, Q(cd)
    Ap ln, Q("")
    Ap ln, Q(tp)
    Ap ln, ""
    Ap ln, Q("")
    Ap ln, Q("")
    Ap ln, Q("Bedrock")
    Ap ln, Q(NOTE_RA)
    Print #ff, Mid(ln, 2)
End Sub


' ================================================================
'  4. PROPOSED LOST ELEMENT ROWS
' ================================================================
Private Function ExportLost(db As DAO.Database, ByVal d As String) As Long
    Dim rs As DAO.Recordset
    Dim ff As Integer
    Dim hd As String
    Dim n As Long
    Dim skipped As Long
    Dim s As String
    Dim orig As String
    Dim ln As String

    ff = FreeFile
    Open d & "EXP_LOST_PROPOSED.csv" For Output As #ff
    hd = "Parent_Code" & SEP & "Element_Code" & SEP & "Evidence_Name"
    hd = hd & SEP & "Evidence_Scope" & SEP & "Position_Code" & SEP & "Notes"
    Print #ff, hd

    Set rs = db.OpenRecordset("SELECT Code, Lost_Body_Evidence FROM T_STRUCTURES ORDER BY Code", dbOpenSnapshot)
    Do While Not rs.EOF
        s = Nz(rs!Lost_Body_Evidence, "")
        orig = s
        If Len(s) > 0 Then
            If s = "Detached debris" Then s = "Detached fragment in situ"
            ' Only real evidence types produce a row. ND, None and any
            ' other stray value say there is nothing to record, and
            ' ID_Evidence_Type is NOT NULL: a row without it cannot
            ' exist. They stay in EXP_DROPPED.csv, which already has
            ' them, and nowhere else.
            If ExistsName(db, "L_LOST_EVIDENCE", "Name", s) Then
                ln = ""
                Ap ln, Q(Nz(rs!Code, ""))
                Ap ln, Q("")
                Ap ln, Q(s)
                Ap ln, Q("Body")
                Ap ln, Q("")
                Ap ln, Q(NOTE_LOST)
                Print #ff, Mid(ln, 2)
                n = n + 1
            Else
                Say "  no lost row for " & Nz(rs!Code, "") & ": '" & orig & "' is not an evidence type"
                skipped = skipped + 1
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
    Close #ff

    Say "EXP_LOST_PROPOSED.csv written: " & n & " rows"
    Say "  Element_Code left BLANK: the old field said there was evidence, not of what"
    If skipped > 0 Then Say "  " & skipped & " value(s) were not evidence types and produced no row"
    ExportLost = n
End Function


' ================================================================
'  5. COHERENCE WARNINGS
' ================================================================
Private Sub CheckCoherence(db As DAO.Database)
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim n As Long

    Say ""
    Say "COHERENCE WARNINGS (nothing is corrected, only reported)"

    sql = "SELECT E.Code FROM T_STRUCTURES AS E WHERE "
    sql = sql & "(E.Dec_Square_Niche=1 OR E.Dec_Relief_T=1 OR E.Dec_Relief_T_Inv=1 "
    sql = sql & "OR E.Dec_Relief_L=1 OR E.Dec_Relief_L_Inv=1 OR E.Dec_Zigzag=1 "
    sql = sql & "OR E.Dec_Stepped=1) AND E.ID NOT IN (SELECT ID_Structure FROM T_DECORATIONS)"
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    Do While Not rs.EOF
        Say "  " & rs!Code & ": decoration booleans set but no T_DECORATIONS row. Dec_Present exported as NULL."
        n = n + 1
        rs.MoveNext
    Loop
    rs.Close

    sql = "SELECT E.Code FROM T_STRUCTURES AS E WHERE "
    sql = sql & "(E.RA_Anthropomorphic=1 OR E.RA_Zoomorphic=1 OR E.RA_Geometric=1 "
    sql = sql & "OR E.RA_Abstract=1) AND (E.Rock_Art<>1 OR E.Rock_Art Is Null)"
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    Do While Not rs.EOF
        Say "  " & rs!Code & ": RA_* set but Rock_Art is not 1. A proposed ROC row exists while"
        Say "      RockArt_Present will be NULL - rule 29 will fire until you resolve it."
        n = n + 1
        rs.MoveNext
    Loop
    rs.Close

    Say "  RockArt_Present=1 with no ROC row yet: expected. Rule 28 fires until you"
    Say "      import EXP_ROCKART_PROPOSED.csv with the positions filled in."
    If n = 0 Then Say "  (no other warnings)"
End Sub


' ================================================================
'  FIELD LISTS - the single source of truth for the rules
' ================================================================

' The 20 A-X elements. v11: 0/1/9. v15: 0/1/2/3/9, where 1 means
' "present COMPLETE". Rule D1 lives here.
Private Sub FillElementFields(f() As String)
    ReDim f(0)
    AddF f, "Embedded_Base_Beams"
    AddF f, "Base_Level"
    AddF f, "Decorative_Socle"
    AddF f, "Tie_Walls"
    AddF f, "Timber_Brackets"
    AddF f, "Transverse_Beams"
    AddF f, "Corbelled_Courses"
    AddF f, "Interbody_Cornice"
    AddF f, "Corner_Quoins"
    AddF f, "Structural_Pilasters"
    AddF f, "Facade_Flank"
    AddF f, "Relief_Frieze"
    AddF f, "Sill"
    AddF f, "Jambs"
    AddF f, "Lintel"
    AddF f, "Upper_Crown"
    AddF f, "Return_Wall"
    AddF f, "Eave_Beam"
    AddF f, "Eave_Surface"
    AddF f, "Chamber_Roof"
End Sub

' Three-value fields, same name and same domain on both sides.
' Only rule D2 applies: the 9 goes to NULL.
Private Sub FillThreeFields(f() As String)
    ReDim f(0)
    AddF f, "Support_Modified"
    AddF f, "Mortar_Present"
    AddF f, "Chinking_Stones"
    AddF f, "Plaster_Present"
    AddF f, "Pigment_Present"
    AddF f, "Recessed_Frame"
    AddF f, "Looting"
    AddF f, "Fire_Damage"
    AddF f, "Animal_Activity"
    AddF f, "Modern_Access"
    AddF f, "Human_Remains"
    AddF f, "Anatomical_Connection"
    AddF f, "Mummification"
    AddF f, "Funerary_Bundles"
    AddF f, "Dispersed_Remains"
    AddF f, "Flexed_Position"
    AddF f, "Bone_Burning"
    AddF f, "Mat_Textiles"
    AddF f, "Mat_Wood"
    AddF f, "Mat_VegFiber"
    AddF f, "Mat_Ceramics"
    AddF f, "Mat_Fauna"
    AddF f, "Mat_DeerAntler"
    AddF f, "Mat_Other"
End Sub

' Same name, same domain, no rule: straight copy.
Private Sub FillDirectFields(f() As String)
    ReDim f(0)
    AddF f, "N_Basal_Bodies"
    AddF f, "N_Chamber_Bodies"
    AddF f, "N_Built_Walls"
    AddF f, "Floor_Plan"
    AddF f, "Length_m"
    AddF f, "Width_m"
    AddF f, "Height_m"
    AddF f, "Height_Above_Base_m"
    AddF f, "Dim_Method"
    AddF f, "Masonry_Quality"
    AddF f, "Masonry_Type"
    AddF f, "Mortar_Type"
    AddF f, "Mortar_Notes"
    AddF f, "Timber_Bracket_Count"
    AddF f, "Timber_Bracket_Role"
    AddF f, "Platform_Surface_Material"
    AddF f, "Interbody_Cornice_Material"
    AddF f, "Lintel_Material"
    AddF f, "Chamber_Roof_Type"
    AddF f, "Rear_Closure_Type"
    AddF f, "Plaster_Extent"
    AddF f, "Pigment_Substrate"
    AddF f, "Facade_Orientation"
    AddF f, "Visibility_Valley"
    AddF f, "Sys_Base"
    AddF f, "Sys_Platform"
    AddF f, "Sys_Portal"
    AddF f, "Sys_Eave"
    AddF f, "Sys_Chamber"
    AddF f, "Platform_Function"
    AddF f, "MNI"
    AddF f, "C14"
    AddF f, "ChaXR_Documented"
    AddF f, "Chrono_Start_Cent"
    AddF f, "Chrono_End_Cent"
    AddF f, "Construction_Phases"
    AddF f, "Coord_Lat_WGS84"
    AddF f, "Coord_Lon_WGS84"
    AddF f, "Coord_E_UTM"
    AddF f, "Coord_N_UTM"
    AddF f, "Altitude_masl"
    AddF f, "Coord_Precision_m"
    AddF f, "URL_Pano"
    AddF f, "URL_Pano_2"
    AddF f, "URL_Giga"
    AddF f, "URL_3D"
    AddF f, "Facade_Observability"
    AddF f, "Interior_Observability"
    AddF f, "Notes"
End Sub

' Retired in v15: their 0s and 9s were already folded into Sys_*.
Private Sub FillRetiredSystemFields(f() As String)
    ReDim f(0)
    AddF f, "Corbelled_Platform"
    AddF f, "Access_Opening"
    AddF f, "Eave"
    AddF f, "Rear_Wall"
End Sub


' ================================================================
'  VALUE RULES
' ================================================================

' D1 + D2 on the A-X elements.
Private Function ElemVal(ByVal v As Variant, ByVal cd As String, ByVal fld As String) As String
    If IsNull(v) Then
        ElemVal = ""
    ElseIf v = 0 Then
        ElemVal = "0"
    ElseIf v = 1 Then
        Drop cd, fld, "1", "D1: v11 present vs v15 present COMPLETE - resolve to 1, 2 or 3"
        ElemVal = ""
    ElseIf v = 9 Then
        Drop cd, fld, "9", "D2: 9 inherited from the v10 default, not a judgement"
        ElemVal = ""
    Else
        Drop cd, fld, CStr(v), "unexpected value in a three-value field"
        ElemVal = ""
    End If
End Function

' D2 on the three-value fields.
Private Function ThreeVal(ByVal v As Variant, ByVal cd As String, ByVal fld As String) As String
    If IsNull(v) Then
        ThreeVal = ""
    ElseIf v = 0 Then
        ThreeVal = "0"
    ElseIf v = 1 Then
        ThreeVal = "1"
    ElseIf v = 9 Then
        Drop cd, fld, "9", "D2: 9 inherited from the v10 default, not a judgement"
        ThreeVal = ""
    Else
        Drop cd, fld, CStr(v), "unexpected value in a three-value field"
        ThreeVal = ""
    End If
End Function

Private Function Cm2M(ByVal v As Variant) As String
    If IsNull(v) Then
        Cm2M = ""
    Else
        Cm2M = Replace(Trim(Str(CDbl(v) / 100)), ",", ".")
    End If
End Function

Private Sub DropDecBooleans(rs As DAO.Recordset, ByVal cd As String)
    Dim f() As String
    Dim i As Long
    Dim v As Variant

    ReDim f(0)
    AddF f, "Dec_Square_Niche"
    AddF f, "Dec_Relief_T"
    AddF f, "Dec_Relief_T_Inv"
    AddF f, "Dec_Relief_L"
    AddF f, "Dec_Relief_L_Inv"
    AddF f, "Dec_Zigzag"
    AddF f, "Dec_Stepped"
    AddF f, "RA_Anthropomorphic"
    AddF f, "RA_Zoomorphic"
    AddF f, "RA_Geometric"
    AddF f, "RA_Abstract"

    For i = 0 To UBound(f)
        v = rs.Fields(f(i)).Value
        If Not IsNull(v) Then
            If v = 1 Then
                Drop cd, f(i), "1", "C: boolean retired - the detail lives in T_DECORATIONS"
            End If
        End If
    Next i
End Sub


' ================================================================
'  UTILITIES
' ================================================================
Private Sub AddF(f() As String, ByVal s As String)
    Dim n As Long
    n = UBound(f)
    If n = 0 And Len(f(0)) = 0 Then
        f(0) = s
    Else
        ReDim Preserve f(n + 1)
        f(n + 1) = s
    End If
End Sub

Private Sub Ap(ByRef s As String, ByVal v As String)
    s = s & SEP & v
End Sub

' Text always quoted, embedded quotes doubled, newlines flattened.
Private Function Q(ByVal s As String) As String
    s = Replace(s, Chr(34), Chr(34) & Chr(34))
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    Q = Chr(34) & s & Chr(34)
End Function

Private Function Fmt(ByVal v As Variant) As String
    If IsNull(v) Then
        Fmt = ""
    ElseIf VarType(v) = vbBoolean Then
        If v Then Fmt = "1" Else Fmt = "0"
    ElseIf IsNumeric(v) Then
        Fmt = Replace(Trim(Str(v)), ",", ".")
    Else
        Fmt = Q(CStr(v))
    End If
End Function

Private Function ExistsName(db As DAO.Database, ByVal tbl As String, ByVal fld As String, ByVal v As String) As Boolean
    Dim rs As DAO.Recordset
    If Len(v) = 0 Then
        ExistsName = False
        Exit Function
    End If
    Set rs = db.OpenRecordset("SELECT ID FROM " & tbl & " WHERE " & fld & "='" & Replace(v, "'", "''") & "'", dbOpenSnapshot)
    ExistsName = Not rs.EOF
    rs.Close
End Function

Private Function NameById(db As DAO.Database, ByVal tbl As String, ByVal fld As String, ByVal id As Variant) As Variant
    Dim rs As DAO.Recordset
    If IsNull(id) Then
        NameById = Null
        Exit Function
    End If
    Set rs = db.OpenRecordset("SELECT " & fld & " AS V FROM " & tbl & " WHERE ID=" & CLng(id), dbOpenSnapshot)
    If rs.EOF Then
        NameById = Null
    Else
        NameById = rs!V
    End If
    rs.Close
End Function

Private Sub Drop(ByVal cd As String, ByVal fld As String, ByVal val As String, ByVal rule As String)
    Dim ln As String
    ln = ""
    Ap ln, Q(cd)
    Ap ln, Q(fld)
    Ap ln, Q(val)
    Ap ln, Q(rule)
    gDrop = gDrop & Mid(ln, 2) & vbCrLf
    gDropN = gDropN + 1
End Sub

Private Sub WriteDropFile(ByVal d As String)
    Dim ff As Integer
    ff = FreeFile
    Open d & "EXP_DROPPED.csv" For Output As #ff
    Print #ff, "Code" & SEP & "Field" & SEP & "Value_v11" & SEP & "Rule"
    If Len(gDrop) > 0 Then Print #ff, Left(gDrop, Len(gDrop) - 2)
    Close #ff
    Say "EXP_DROPPED.csv written: " & gDropN & " values"
End Sub

Private Function TargetDir() As String
    Dim d As String
    If Len(EXPORT_DIR) > 0 Then
        d = EXPORT_DIR
    Else
        d = CurrentProject.Path & "\transfer\"
    End If
    If Right(d, 1) <> "\" Then d = d & "\"
    If Len(Dir(d, vbDirectory)) = 0 Then
        On Error Resume Next
        MkDir Left(d, Len(d) - 1)
        On Error GoTo 0
    End If
    If Len(Dir(d, vbDirectory)) = 0 Then
        MsgBox "Cannot create the folder: " & d, vbCritical
        TargetDir = ""
    Else
        TargetDir = d
    End If
End Function

Private Function CheckSource(db As DAO.Database) As Boolean
    Dim need() As String
    Dim i As Long
    Dim ok As Boolean

    ReDim need(0)
    AddF need, "Lost_Body_Evidence"
    AddF need, "Corbelled_Platform"
    AddF need, "Rock_Art"
    AddF need, "Opening_Width_cm"
    AddF need, "ID_Material_Status"

    ok = True
    For i = 0 To UBound(need)
        If Not HasField(db, "T_STRUCTURES", need(i)) Then
            Say "MISSING in the source: T_STRUCTURES." & need(i)
            ok = False
        End If
    Next i
    CheckSource = ok
End Function

Private Function HasField(db As DAO.Database, ByVal tbl As String, ByVal fld As String) As Boolean
    Dim f As DAO.Field
    On Error GoTo NoField
    Set f = db.TableDefs(tbl).Fields(fld)
    HasField = True
    Exit Function
NoField:
    HasField = False
End Function

Private Sub Say(ByVal s As String)
    Debug.Print s
    gLog = gLog & s & vbCrLf
End Sub

Private Sub WriteLogFile(ByVal d As String)
    Dim ff As Integer
    ff = FreeFile
    Open d & "EXPORT_LOG.txt" For Output As #ff
    Print #ff, "CHACHAPOYA EXPORT LOG - " & Now()
    Print #ff, gLog
    Close #ff
End Sub
