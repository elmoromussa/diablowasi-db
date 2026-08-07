Option Compare Database
Option Explicit

' ================================================================
'  CHACHAPOYA - EXPORT FROM v15 TO INSPECTABLE CSV
'  Author: Esteve Ribera Torro | TFM Arqueologia UA
'  Spec: DELTA_v15_v16.md, section 10
'
'  RUN Sub ExportAll() ON THE v15 SOURCE DATABASE.
'  Then build a blank v16 with chachapoya_DB_v16.bas -> BuildDB()
'  and run chachapoya_IMPORT_v16.bas -> Sub ImportAll() there.
'
'  WHY TWO STEPS WITH A FILE IN BETWEEN. A direct database-to-
'  database copy is shorter to write and impossible to audit. The
'  CSV is the point: it can be opened, read, sorted and CORRECTED
'  before anything is written to the new database, which is the
'  only moment at which a human can still catch a bad mapping
'  cheaply. Every transfer this project has done has turned up
'  something at that step.
'
'  NO IDs CROSS THE BOUNDARY. Every foreign key is exported as the
'  HUMAN-READABLE VALUE it points at - a code or a name - and
'  resolved again on the far side. This is not tidiness: in the
'  v11->v15 transfer the L_STRUCT_BODY ids had been reorganised
'  between versions, so numerically identical keys pointed at
'  DIFFERENT positions and 19 of 29 decoration rows would have
'  been silently mislabelled. A wrong id raises no error and shows
'  no symptom; it just quietly changes what the data says.
'
'  THE CONSERVATIVE NULL RULE AND THE OLD_ COLUMNS. Fields whose
'  CRITERION changed in v16 must arrive NULL, because a value
'  carried across unreviewed would assert a judgement nobody made
'  under the new criterion. But blanking them in silence would
'  throw away the old reading, which is the best starting point
'  for the review. So each affected field is exported TWICE: the
'  live column, left blank, and an OLD_ column carrying what v15
'  said. The importer ignores every OLD_ column.
'
'  THAT MAKES THE CSV THE REVIEW WORKSPACE. Open it, look at the
'  photograph, and if the old value still holds under the new
'  criterion, copy it from OLD_Sill back into Sill. What you
'  copy is a judgement; what you leave blank stays honestly empty
'  and QRY_19_V16_Review will list it.
'
'  FIELDS ARE READ FROM THE TABLEDEF, NOT HARD-CODED. A hard-coded
'  list of 128 field names is a list that will be wrong after the
'  next schema change, and wrong silently. Reading them from the
'  table means a field dropped in v16 simply arrives as a column
'  the importer does not recognise, and says so.
' ================================================================

' Output folder. MUST EXIST and end with a path separator.
Private Const OUT_DIR As String = "C:\chachapoya_transfer\"

' Fields excluded from the dynamic sweep of T_STRUCTURES: the
' primary key, the code (written first) and every foreign key
' (written as a resolved name instead).
Private Const SKIP_LIST As String = "|ID|Code|ID_Sector|ID_Typology|ID_Support|ID_Support_Secondary|ID_Parent|ID_Group|ID_Arch_Status|ID_Material_Status|ID_Vol_Method|ID_Coord_Method|ID_Campaign|"

Sub ExportAll()
    Dim db As DAO.Database
    Set db = CurrentDb()

    If Dir(OUT_DIR, vbDirectory) = "" Then
        MsgBox "The folder " & OUT_DIR & " does not exist. Create it first, or edit OUT_DIR at the top of this module.", vbCritical, "Export aborted"
        Exit Sub
    End If

    Debug.Print "=== EXPORT v15 -> CSV ==="
    Debug.Print "Folder: " & OUT_DIR

    Dim nG As Long
    Dim nS As Long
    Dim nD As Long
    Dim nDat As Long
    Dim nInd As Long
    Dim nF As Long
    Dim nC As Long
    Dim nL As Long

    nG = ExportGroups(db)
    nS = ExportStructures(db)
    nD = ExportDecorations(db)
    nDat = ExportDating(db)
    nInd = ExportIndividuals(db)
    nF = ExportFeatures(db)
    nC = ExportConnections(db)
    nL = ExportLost(db)

    Set db = Nothing

    Dim msg As String
    msg = "EXPORT COMPLETE" & vbCrLf & vbCrLf
    msg = msg & "  T_GROUPS:         " & nG & vbCrLf
    msg = msg & "  T_STRUCTURES:     " & nS & vbCrLf
    msg = msg & "  T_DECORATIONS:    " & nD & vbCrLf
    msg = msg & "  T_DATING:         " & nDat & vbCrLf
    msg = msg & "  T_INDIVIDUALS:    " & nInd & vbCrLf
    msg = msg & "  T_ARCH_FEATURES:  " & nF & vbCrLf
    msg = msg & "  T_CONNECTIONS:    " & nC & vbCrLf
    msg = msg & "  T_LOST_ELEMENTS:  " & nL & vbCrLf & vbCrLf
    msg = msg & "Written to " & OUT_DIR & vbCrLf & vbCrLf
    msg = msg & "BEFORE IMPORTING, open T_STRUCTURES.csv." & vbCrLf
    msg = msg & "Sill, Jambs, Lintel are blank on purpose: the" & vbCrLf
    msg = msg & "differentiation criterion changed, so the old" & vbCrLf
    msg = msg & "values are not judgements under the new one." & vbCrLf
    msg = msg & "They are kept in the OLD_ columns. Copy back" & vbCrLf
    msg = msg & "only what you have actually re-checked." & vbCrLf & vbCrLf
    msg = msg & "Chamber_Roof is blank only where the roof was" & vbCrLf
    msg = msg & "natural bedrock or mixed: those need the" & vbCrLf
    msg = msg & "contact test. Built roofs came across intact."
    MsgBox msg, vbInformation, "Export v15"
    Debug.Print "=== EXPORT FINISHED ==="
End Sub

' ================================================================
'  CSV WRITING
'  Quote everything and double any internal quote. Quoting only
'  what "needs" it is how a Notes field containing a comma
'  silently shifts every column to its right.
' ================================================================
Private Function Q(v As Variant) As String
    Dim t As String
    If IsNull(v) Then
        Q = """"""
        Exit Function
    End If
    t = CStr(v)
    t = Replace(t, """", """""")
    Q = """" & t & """"
End Function

' Booleans stored as YESNO come out as 0 and -1, which is not what
' a person reading the file expects and not what the importer
' wants back. Normalise to 1 and 0 in both directions.
Private Function QB(v As Variant) As String
    If IsNull(v) Then
        QB = """"""
    ElseIf v Then
        QB = """1"""
    Else
        QB = """0"""
    End If
End Function

Private Function OpenOut(fname As String) As Integer
    Dim fn As Integer
    fn = FreeFile
    Open OUT_DIR & fname For Output As #fn
    OpenOut = fn
End Function

' ================================================================
'  FOREIGN KEY RESOLUTION
'  One helper per shape. All return "" for a null or unresolved
'  key, and an unresolved key is REPORTED: a foreign key pointing
'  at a row that no longer exists is a fact about the source
'  database, not a nuisance to swallow.
' ================================================================
Private Function NameOf(db As DAO.Database, tbl As String, col As String, id As Variant) As String
    If IsNull(id) Then Exit Function
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT " & col & " AS V FROM " & tbl & " WHERE ID=" & CLng(id), dbOpenSnapshot)
    If Not rs.EOF Then
        If Not IsNull(rs!V) Then NameOf = CStr(rs!V)
    Else
        Debug.Print "  *** dangling key: " & tbl & ".ID=" & id & " does not exist"
    End If
    rs.Close
    Set rs = Nothing
End Function

' Sector names are not unique across sites, so the site travels
' with them. Site and sector are exported as separate columns.
Private Function SiteOfSector(db As DAO.Database, id As Variant) As String
    If IsNull(id) Then Exit Function
    Dim rs As DAO.Recordset
    Dim q2 As String
    q2 = "SELECT S.Site_Name AS V FROM L_SECTORS AS SC "
    q2 = q2 & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID "
    q2 = q2 & "WHERE SC.ID=" & CLng(id)
    Set rs = db.OpenRecordset(q2, dbOpenSnapshot)
    If Not rs.EOF Then
        If Not IsNull(rs!V) Then SiteOfSector = CStr(rs!V)
    End If
    rs.Close
    Set rs = Nothing
End Function

Private Function CodeOfStructure(db As DAO.Database, id As Variant) As String
    CodeOfStructure = NameOf(db, "T_STRUCTURES", "Code", id)
End Function

Private Function IsIn(list As String, n As String) As Boolean
    IsIn = (InStr(1, list, "|" & n & "|", vbTextCompare) > 0)
End Function

' ================================================================
'  T_GROUPS - exported first because T_STRUCTURES points at it
' ================================================================
Private Function ExportGroups(db As DAO.Database) As Long
    Dim fn As Integer
    Dim rs As DAO.Recordset
    Dim n As Long
    fn = OpenOut("T_GROUPS.csv")
    Print #fn, "Group_Code,Site,Sector,Group_Type,N_Members,Notes"
    Set rs = db.OpenRecordset("SELECT * FROM T_GROUPS ORDER BY Group_Code", dbOpenSnapshot)
    Do While Not rs.EOF
        Dim ln As String
        ln = Q(rs!Group_Code)
        ln = ln & "," & Q(SiteOfSector(db, rs!ID_Sector))
        ln = ln & "," & Q(NameOf(db, "L_SECTORS", "Sector_Name", rs!ID_Sector))
        ln = ln & "," & Q(NameOf(db, "L_GROUP_TYPE", "Name", rs!ID_Group_Type))
        ln = ln & "," & Q(rs!N_Members)
        ln = ln & "," & Q(rs!Notes)
        Print #fn, ln
        n = n + 1
        rs.MoveNext
    Loop
    rs.Close
    Close #fn
    Debug.Print "-> T_GROUPS.csv: " & n & " rows"
    ExportGroups = n
End Function

' ================================================================
'  T_STRUCTURES
'  Header: Code, the twelve resolved keys, then every remaining
'  field in TableDef order, then the OLD_ columns.
' ================================================================
Private Function ExportStructures(db As DAO.Database) As Long
    Dim fn As Integer
    Dim rs As DAO.Recordset
    Dim td As DAO.TableDef
    Dim fld As DAO.Field
    Dim n As Long
    Dim nBlanked As Long
    Dim nRoof As Long

    Set td = db.TableDefs("T_STRUCTURES")
    fn = OpenOut("T_STRUCTURES.csv")

    Dim hdr As String
    hdr = "Code,Site,Sector,Typology,Support,Support_Secondary,Parent_Code,Group_Code,Arch_Status,Material_Status,Vol_Method,Coord_Method,Campaign"
    For Each fld In td.Fields
        If Not IsIn(SKIP_LIST, fld.Name) Then hdr = hdr & "," & fld.Name
    Next fld
    ' Informational columns. The importer skips anything starting
    ' with OLD_, so these can be edited or deleted freely.
    hdr = hdr & ",OLD_Sill,OLD_Jambs,OLD_Lintel,OLD_Chamber_Roof,OLD_Chamber_Roof_Type,OLD_Interbody_Cornice_Material"
    Print #fn, hdr

    Set rs = db.OpenRecordset("SELECT * FROM T_STRUCTURES ORDER BY Code", dbOpenSnapshot)
    Do While Not rs.EOF
        Dim ln As String
        ln = Q(rs!Code)
        ln = ln & "," & Q(SiteOfSector(db, rs!ID_Sector))
        ln = ln & "," & Q(NameOf(db, "L_SECTORS", "Sector_Name", rs!ID_Sector))
        ln = ln & "," & Q(NameOf(db, "L_TYPOLOGY", "Name", rs!ID_Typology))
        ln = ln & "," & Q(NameOf(db, "L_SUPPORT", "Name", rs!ID_Support))
        ln = ln & "," & Q(NameOf(db, "L_SUPPORT", "Name", rs!ID_Support_Secondary))
        ln = ln & "," & Q(CodeOfStructure(db, rs!ID_Parent))
        ln = ln & "," & Q(NameOf(db, "T_GROUPS", "Group_Code", rs!ID_Group))
        ln = ln & "," & Q(NameOf(db, "L_STATUS", "Name", rs!ID_Arch_Status))
        ln = ln & "," & Q(NameOf(db, "L_MATERIAL_STATUS", "Name", rs!ID_Material_Status))
        ln = ln & "," & Q(NameOf(db, "L_VOL_METHOD", "Name", rs!ID_Vol_Method))
        ln = ln & "," & Q(NameOf(db, "L_COORD_METHOD", "Name", rs!ID_Coord_Method))
        ln = ln & "," & Q(NameOf(db, "L_CAMPAIGN", "Code", rs!ID_Campaign))

        Dim rt As String
        rt = ""
        If Not IsNull(rs!Chamber_Roof_Type) Then rt = CStr(rs!Chamber_Roof_Type)

        For Each fld In td.Fields
            If Not IsIn(SKIP_LIST, fld.Name) Then
                If fld.Name = "Sill" Or fld.Name = "Jambs" Or fld.Name = "Lintel" Then
                    ' Always blank: the differentiation criterion
                    ' (delta 5.4) changes what counts as present, so
                    ' every one of these needs re-judging.
                    ln = ln & "," & """"""
                    If Not IsNull(rs.Fields(fld.Name).Value) Then nBlanked = nBlanked + 1
                ElseIf fld.Name = "Chamber_Roof" Or fld.Name = "Chamber_Roof_Type" Then
                    ' Blank ONLY where the roof was natural rock or
                    ' mixed: those need the contact test (delta 5.2).
                    ' A built roof was built either way and needs no
                    ' second judgement, so it travels intact - the
                    ' conservative rule applies where the criterion
                    ' moved, not everywhere as a reflex.
                    If rt = "Natural bedrock" Or rt = "Mixed" Then
                        ln = ln & "," & """"""
                        If fld.Name = "Chamber_Roof" Then nRoof = nRoof + 1
                    Else
                        ln = ln & "," & Q(rs.Fields(fld.Name).Value)
                    End If
                ElseIf fld.Type = dbBoolean Then
                    ln = ln & "," & QB(rs.Fields(fld.Name).Value)
                Else
                    ln = ln & "," & Q(rs.Fields(fld.Name).Value)
                End If
            End If
        Next fld

        ln = ln & "," & Q(rs!Sill)
        ln = ln & "," & Q(rs!Jambs)
        ln = ln & "," & Q(rs!Lintel)
        ln = ln & "," & Q(rs!Chamber_Roof)
        ln = ln & "," & Q(rs!Chamber_Roof_Type)
        ln = ln & "," & Q(rs!Interbody_Cornice_Material)

        Print #fn, ln
        n = n + 1
        rs.MoveNext
    Loop
    rs.Close
    Close #fn
    Debug.Print "-> T_STRUCTURES.csv: " & n & " rows"
    Debug.Print "   portal values blanked for review: " & nBlanked
    Debug.Print "   chamber roofs blanked (natural or mixed): " & nRoof
    ExportStructures = n
End Function

' ================================================================
'  T_DECORATIONS
'  The position goes out as a CODE, not an id and not a name: the
'  codes are stable across versions and the names are exactly
'  what v16 changed.
' ================================================================
Private Function ExportDecorations(db As DAO.Database) As Long
    Dim fn As Integer
    Dim rs As DAO.Recordset
    Dim n As Long
    Dim nSoc As Long
    fn = OpenOut("T_DECORATIONS.csv")
    Print #fn, "Structure_Code,Position_Code,Dec_Type,Body_No,Color,Color_Secondary,Substrate,Notes"
    Dim q2 As String
    q2 = "SELECT D.*, E.Code AS SCode, B.Code AS BCode, DT.Name AS DName "
    q2 = q2 & "FROM ((T_DECORATIONS AS D "
    q2 = q2 & "INNER JOIN T_STRUCTURES AS E ON D.ID_Structure=E.ID) "
    q2 = q2 & "LEFT JOIN L_STRUCT_BODY AS B ON D.ID_Struct_Body=B.ID) "
    q2 = q2 & "LEFT JOIN L_DEC_TYPE AS DT ON D.ID_Dec_Type=DT.ID "
    q2 = q2 & "ORDER BY E.Code"
    Set rs = db.OpenRecordset(q2, dbOpenSnapshot)
    Do While Not rs.EOF
        Dim ln As String
        ln = Q(rs!SCode)
        ln = ln & "," & Q(rs!BCode)
        ln = ln & "," & Q(rs!DName)
        ln = ln & "," & Q(rs!Body_No)
        ln = ln & "," & Q(rs!Color)
        ln = ln & "," & Q(rs!Color_Secondary)
        ln = ln & "," & Q(rs!Substrate)
        ln = ln & "," & Q(rs!Notes)
        Print #fn, ln
        If Nz(rs!BCode, "") = "SOC" Then nSoc = nSoc + 1
        n = n + 1
        rs.MoveNext
    Loop
    rs.Close
    Close #fn
    Debug.Print "-> T_DECORATIONS.csv: " & n & " rows"
    Debug.Print "   rows on SOC (will be remapped to BAS on import): " & nSoc
    ExportDecorations = n
End Function

Private Function ExportDating(db As DAO.Database) As Long
    Dim fn As Integer
    Dim rs As DAO.Recordset
    Dim n As Long
    fn = OpenOut("T_DATING.csv")
    Print #fn, "Structure_Code,Sample_Type,Date_BP,Sigma1_Start,Sigma1_End,Sigma2_Start,Sigma2_End,Lab_Reference,Bibliog_Reference"
    Dim q2 As String
    q2 = "SELECT D.*, E.Code AS SCode FROM T_DATING AS D "
    q2 = q2 & "INNER JOIN T_STRUCTURES AS E ON D.ID_Structure=E.ID ORDER BY E.Code"
    Set rs = db.OpenRecordset(q2, dbOpenSnapshot)
    Do While Not rs.EOF
        Dim ln As String
        ln = Q(rs!SCode)
        ln = ln & "," & Q(rs!Sample_Type)
        ln = ln & "," & Q(rs!Date_BP)
        ln = ln & "," & Q(rs!Sigma1_Start)
        ln = ln & "," & Q(rs!Sigma1_End)
        ln = ln & "," & Q(rs!Sigma2_Start)
        ln = ln & "," & Q(rs!Sigma2_End)
        ln = ln & "," & Q(rs!Lab_Reference)
        ln = ln & "," & Q(rs!Bibliog_Reference)
        Print #fn, ln
        n = n + 1
        rs.MoveNext
    Loop
    rs.Close
    Close #fn
    Debug.Print "-> T_DATING.csv: " & n & " rows"
    ExportDating = n
End Function

Private Function ExportIndividuals(db As DAO.Database) As Long
    Dim fn As Integer
    Dim rs As DAO.Recordset
    Dim n As Long
    fn = OpenOut("T_INDIVIDUALS.csv")
    Print #fn, "Structure_Code,Individual_No,Age_Category,Sex_Category,Preservation,Notes"
    Dim q2 As String
    q2 = "SELECT I.*, E.Code AS SCode FROM T_INDIVIDUALS AS I "
    q2 = q2 & "INNER JOIN T_STRUCTURES AS E ON I.ID_Structure=E.ID ORDER BY E.Code"
    Set rs = db.OpenRecordset(q2, dbOpenSnapshot)
    Do While Not rs.EOF
        Dim ln As String
        ln = Q(rs!SCode)
        ln = ln & "," & Q(rs!Individual_No)
        ln = ln & "," & Q(rs!Age_Category)
        ln = ln & "," & Q(rs!Sex_Category)
        ln = ln & "," & Q(rs!Preservation)
        ln = ln & "," & Q(rs!Notes)
        Print #fn, ln
        n = n + 1
        rs.MoveNext
    Loop
    rs.Close
    Close #fn
    Debug.Print "-> T_INDIVIDUALS.csv: " & n & " rows"
    ExportIndividuals = n
End Function

Private Function ExportFeatures(db As DAO.Database) As Long
    Dim fn As Integer
    Dim rs As DAO.Recordset
    Dim n As Long
    fn = OpenOut("T_ARCH_FEATURES.csv")
    Print #fn, "Structure_Code,Feature_Code,Present,Feature_Count,Material,Notes"
    Dim q2 As String
    q2 = "SELECT F.*, E.Code AS SCode FROM T_ARCH_FEATURES AS F "
    q2 = q2 & "INNER JOIN T_STRUCTURES AS E ON F.ID_Structure=E.ID ORDER BY E.Code"
    Set rs = db.OpenRecordset(q2, dbOpenSnapshot)
    Do While Not rs.EOF
        Dim ln As String
        ln = Q(rs!SCode)
        ln = ln & "," & Q(rs!Feature_Code)
        ln = ln & "," & QB(rs!Present)
        ln = ln & "," & Q(rs!Feature_Count)
        ln = ln & "," & Q(rs!Material)
        ln = ln & "," & Q(rs!Notes)
        Print #fn, ln
        n = n + 1
        rs.MoveNext
    Loop
    rs.Close
    Close #fn
    Debug.Print "-> T_ARCH_FEATURES.csv: " & n & " rows"
    ExportFeatures = n
End Function

Private Function ExportConnections(db As DAO.Database) As Long
    Dim fn As Integer
    Dim rs As DAO.Recordset
    Dim n As Long
    fn = OpenOut("T_CONNECTIONS.csv")
    Print #fn, "Code_A,Code_B,Connection_Type,Chrono_Relation,Confidence,Notes"
    Set rs = db.OpenRecordset("SELECT * FROM T_CONNECTIONS", dbOpenSnapshot)
    Do While Not rs.EOF
        Dim ln As String
        ln = Q(CodeOfStructure(db, rs!ID_Struct_A))
        ln = ln & "," & Q(CodeOfStructure(db, rs!ID_Struct_B))
        ln = ln & "," & Q(rs!Connection_Type)
        ln = ln & "," & Q(rs!Chrono_Relation)
        ln = ln & "," & Q(rs!Confidence)
        ln = ln & "," & Q(rs!Notes)
        Print #fn, ln
        n = n + 1
        rs.MoveNext
    Loop
    rs.Close
    Close #fn
    Debug.Print "-> T_CONNECTIONS.csv: " & n & " rows"
    ExportConnections = n
End Function

' ================================================================
'  T_LOST_ELEMENTS
'  ID_Position points at L_STRUCT_BODY, so it carries the same SOC
'  problem as the decoration rows and is exported as a code too.
' ================================================================
Private Function ExportLost(db As DAO.Database) As Long
    Dim fn As Integer
    Dim rs As DAO.Recordset
    Dim n As Long
    fn = OpenOut("T_LOST_ELEMENTS.csv")
    Print #fn, "Structure_Code,Element_Code,Evidence_Type,Evidence_Scope,Position_Code,Notes"
    Dim q2 As String
    q2 = "SELECT L.*, E.Code AS SCode, EV.Name AS EvName, B.Code AS BCode "
    q2 = q2 & "FROM ((T_LOST_ELEMENTS AS L "
    q2 = q2 & "INNER JOIN T_STRUCTURES AS E ON L.ID_Structure=E.ID) "
    q2 = q2 & "LEFT JOIN L_LOST_EVIDENCE AS EV ON L.ID_Evidence_Type=EV.ID) "
    q2 = q2 & "LEFT JOIN L_STRUCT_BODY AS B ON L.ID_Position=B.ID "
    q2 = q2 & "ORDER BY E.Code"
    Set rs = db.OpenRecordset(q2, dbOpenSnapshot)
    Do While Not rs.EOF
        Dim ln As String
        ln = Q(rs!SCode)
        ln = ln & "," & Q(rs!Element_Code)
        ln = ln & "," & Q(rs!EvName)
        ln = ln & "," & Q(rs!Evidence_Scope)
        ln = ln & "," & Q(rs!BCode)
        ln = ln & "," & Q(rs!Notes)
        Print #fn, ln
        n = n + 1
        rs.MoveNext
    Loop
    rs.Close
    Close #fn
    Debug.Print "-> T_LOST_ELEMENTS.csv: " & n & " rows"
    ExportLost = n
End Function
