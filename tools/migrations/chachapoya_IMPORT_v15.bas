Option Compare Database
Option Explicit

' ================================================================
'  CHACHAPOYA DB - STEP 2 OF 2: IMPORT INTO THE v15 DATABASE
'  Esteve Ribera Torro | TFM Arqueologia UA
'
'  RUN THIS ON THE NEW DATABASE, the one built by BuildDB() and
'  BuildForm(). Import a new blank module, paste this, set
'  IMPORT_DIR to the transfer folder and run Sub ImportV15().
'
'  It reads the CSV files written by chachapoya_EXPORT_v11.bas,
'  after you have reviewed and edited them.
'
'  THREE THINGS THIS SCRIPT DOES THAT A NAIVE IMPORT WOULD NOT
'
'  1. IT RESOLVES EVERY FOREIGN KEY BY NAME, NEVER BY ID. The IDs
'     of L_STRUCT_BODY, L_SUPPORT, L_TYPOLOGY, L_MATERIAL_STATUS
'     and L_GROUP_TYPE have moved between v11 and v15. Copying the
'     numbers would not raise a single error and would silently
'     turn 9 over-lintel motifs into jambs and 8 facade flanks into
'     interbody cornices.
'
'  2. IT EXPLICITLY WRITES NULL INTO EVERY FIELD THE CSV DOES NOT
'     CARRY. BuildDB sets DAO defaults: 0 on the 20 element fields,
'     Absent on the five Sys_*, Undetermined on Platform_Function.
'     A plain INSERT that omits a column FIRES THE DEFAULT, so an
'     untouched field would arrive as a verified absence. That is
'     precisely the false zero the whole domain exists to prevent,
'     and it would also hide the record from rule 17.
'
'  3. IT RUNS INSIDE A TRANSACTION. Any unresolved name or unknown
'     parent code rolls the whole import back, so the database is
'     never left half migrated.
'
'  Columns whose name ends in _v11, and ID_v11, are IGNORED: they
'  are the reviewer context written by the export. You can delete
'  them from the CSV or leave them, it makes no difference.
'
'  A column whose name matches no field is reported and skipped,
'  so you can add your own working columns to the sheet safely.
' ================================================================

' Leave empty to use <current database folder>\transfer\
Private Const IMPORT_DIR As String = ""

Private Const SEP As String = ";"

Private gCodes As Collection
Private gLog As String
Private gErr As Long
Private gPCode() As String
Private gPParent() As String
Private gPN As Long


' ================================================================
'  MAIN
' ================================================================
Public Sub ImportV15()
    Dim db As DAO.Database
    Dim ws As DAO.Workspace
    Dim d As String
    Dim n1 As Long
    Dim n2 As Long
    Dim n3 As Long
    Dim n4 As Long
    Dim inTrans As Boolean

    Set db = CurrentDb
    Set ws = DBEngine.Workspaces(0)
    Set gCodes = New Collection
    gLog = ""
    gErr = 0
    gPN = 0
    ReDim gPCode(0)
    ReDim gPParent(0)

    d = TargetDir()
    If Len(d) = 0 Then Exit Sub

    Say "CHACHAPOYA IMPORT v11 -> v15"
    Say "Target: " & CurrentProject.Name
    Say "Folder: " & d
    Say ""

    If Not CheckTarget(db) Then
        MsgBox "This does not look like the v15 schema. See the Immediate window.", vbCritical
        Exit Sub
    End If

    If db.OpenRecordset("SELECT COUNT(*) AS N FROM T_STRUCTURES", dbOpenSnapshot)!N > 0 Then
        MsgBox "T_STRUCTURES is not empty. Rebuild the database with BuildDB before importing.", vbCritical
        Exit Sub
    End If

    On Error GoTo Fail
    If Not PreflightWidths(db, d & "EXP_STRUCTURES.csv") Then
        MsgBox "Some values do not fit their field. Nothing was written - see the Immediate window.", vbCritical
        WriteLogFile d
        Exit Sub
    End If

    ws.BeginTrans
    inTrans = True

    n1 = ImportStructures(db, d & "EXP_STRUCTURES.csv")
    If gErr = 0 Then ResolveParents db
    If gErr = 0 Then n2 = ImportDecRows(db, d & "EXP_DECORATIONS.csv", "decorations")
    If gErr = 0 Then n3 = ImportDecRows(db, d & "EXP_ROCKART_PROPOSED.csv", "rock art")
    If gErr = 0 Then n4 = ImportLost(db, d & "EXP_LOST_PROPOSED.csv")

    If gErr > 0 Then
        ws.Rollback
        inTrans = False
        Say ""
        Say "ROLLED BACK. " & gErr & " error(s). Nothing was written."
        WriteLogFile d
        MsgBox "Import ABORTED: " & gErr & " error(s)." & vbCrLf & "Nothing was written. See IMPORT_LOG.txt.", vbCritical
        Exit Sub
    End If

    ws.CommitTrans
    inTrans = False

    Say ""
    Say "SUMMARY"
    Say "  T_STRUCTURES ......... " & n1 & " rows"
    Say "  T_DECORATIONS ........ " & n2 & " rows (architectural decoration)"
    Say "  T_DECORATIONS ........ " & n3 & " rows (proposed rock art)"
    Say "  T_LOST_ELEMENTS ...... " & n4 & " rows (proposed)"
    Say ""
    Say "Now open QRY_16 and work through rule 17: every NULL is a field"
    Say "that has not been assessed yet. EXP_DROPPED.csv tells you what"
    Say "the v11 value was for each one."

    WriteLogFile d

    Dim m As String
    m = "Import finished." & vbCrLf & vbCrLf
    m = m & "Structures: " & n1 & vbCrLf
    m = m & "Decorations: " & n2 & vbCrLf
    m = m & "Rock art rows: " & n3 & vbCrLf
    m = m & "Lost element rows: " & n4 & vbCrLf & vbCrLf
    m = m & "Detail in IMPORT_LOG.txt."
    MsgBox m, vbInformation
    Exit Sub

Fail:
    Say ""
    Say "RUNTIME ERROR " & Err.Number & ": " & Err.Description
    If inTrans Then
        ws.Rollback
        Say "ROLLED BACK. Nothing was written."
    End If
    WriteLogFile d
    MsgBox "Import aborted by an error: " & Err.Description & vbCrLf & "Nothing was written. See IMPORT_LOG.txt.", vbCritical
End Sub


' ================================================================
'  0. SCHEMA PATCH - run this once before ImportV15
'
'  Pigment_Extent is TEXT(20) in the build script but its own domain
'  contains "Architectural elements", which is 22 characters. The
'  field cannot store a value the form offers, so the first record
'  that carries it fails with error 3163 - whether it arrives from
'  this import or from someone picking it in the form. Widened to
'  TEXT(30). Idempotent: safe to run twice.
'
'  Fix it in chachapoya_DB_v15.bas too, line 517, so a rebuild does
'  not bring the problem back:
'      sql = sql & "Pigment_Extent TEXT(30),"
' ================================================================
Public Sub PatchV15Schema()
    Dim db As DAO.Database
    Dim sz As Long

    Set db = CurrentDb
    sz = db.TableDefs("T_STRUCTURES").Fields("Pigment_Extent").Size
    If sz >= 30 Then
        MsgBox "Already patched: Pigment_Extent is TEXT(" & sz & ").", vbInformation
        Exit Sub
    End If

    db.Execute "ALTER TABLE T_STRUCTURES ALTER COLUMN Pigment_Extent TEXT(30)", dbFailOnError
    db.TableDefs.Refresh
    sz = db.TableDefs("T_STRUCTURES").Fields("Pigment_Extent").Size
    Debug.Print "Pigment_Extent widened to TEXT(" & sz & ")"
    MsgBox "Pigment_Extent widened to TEXT(" & sz & "). Now run ImportV15.", vbInformation
End Sub


' ================================================================
'  1. STRUCTURES
' ================================================================
Private Function ImportStructures(db As DAO.Database, ByVal path As String) As Long
    Dim rs As DAO.Recordset
    Dim td As DAO.TableDef
    Dim f As DAO.Field
    Dim ff As Integer
    Dim ln As String
    Dim hdr() As String
    Dim val() As String
    Dim done As String
    Dim i As Long
    Dim n As Long
    Dim cd As String
    Dim nm As String
    Dim sv As String
    Dim id As Variant
    Dim newId As Long
    Dim e0 As Long
    Dim skipped As String

    If Len(Dir(path)) = 0 Then
        Say "MISSING: " & path
        gErr = gErr + 1
        Exit Function
    End If

    Set td = db.TableDefs("T_STRUCTURES")
    Set rs = db.OpenRecordset("T_STRUCTURES", dbOpenDynaset)

    ff = FreeFile
    Open path For Input As #ff
    Line Input #ff, ln
    hdr = SplitCsv(ln)

    Do While Not EOF(ff)
        Line Input #ff, ln
        If Len(Trim(ln)) > 0 Then
            val = SplitCsv(ln)
            done = "|"
            cd = ColVal(hdr, val, "Code")
            If Len(cd) = 0 Then
                Say "ERROR: a row has no Code"
                gErr = gErr + 1
            Else
                e0 = gErr
                rs.AddNew
                For i = 0 To UBound(hdr)
                    nm = Trim(hdr(i))
                    If i <= UBound(val) Then sv = Trim(val(i)) Else sv = ""

                    If Len(nm) = 0 Then
                        ' nothing
                    ElseIf nm = "ID_v11" Or Right(nm, 4) = "_v11" Then
                        ' reviewer context, ignored on purpose
                    ElseIf nm = "Parent_Code" Then
                        If Len(sv) > 0 Then Defer cd, sv
                    ElseIf IsLookupCol(nm) Then
                        If Len(sv) > 0 Then
                            id = LookupId(db, nm, sv)
                            If IsNull(id) Then
                                Say "ERROR: " & cd & " - " & nm & " = '" & sv & "' does not exist in the v15 lookup"
                                gErr = gErr + 1
                            Else
                                rs.Fields(LookupField(nm)).Value = id
                            End If
                        Else
                            rs.Fields(LookupField(nm)).Value = Null
                        End If
                        done = done & LookupField(nm) & "|"
                    ElseIf HasField(db, "T_STRUCTURES", nm) Then
                        On Error Resume Next
                        rs.Fields(nm).Value = ToVal(sv, td.Fields(nm))
                        If Err.Number <> 0 Then
                            Say "ERROR: " & cd & " - field " & nm & " rejected '" & sv & "': " & Err.Description
                            gErr = gErr + 1
                            Err.Clear
                        End If
                        On Error GoTo 0
                        done = done & nm & "|"
                    Else
                        If InStr(skipped, "|" & nm & "|") = 0 Then
                            skipped = skipped & "|" & nm & "|"
                            Say "  column ignored (no such field): " & nm
                        End If
                    End If
                Next i

                ' A row with an unresolved name is abandoned rather than
                ' half written: the transaction rolls everything back at
                ' the end anyway, but this keeps the log readable.
                If gErr > e0 Then
                    rs.CancelUpdate
                    GoTo NextRow
                End If

                ' Everything the sheet did not carry is written as NULL,
                ' never left to the DAO default. See the header, point 2.
                For Each f In td.Fields
                    If f.Name <> "ID" Then
                        If InStr(done, "|" & f.Name & "|") = 0 Then
                            If f.Type = dbBoolean Then
                                rs.Fields(f.Name).Value = False
                            Else
                                rs.Fields(f.Name).Value = Null
                            End If
                        End If
                    End If
                Next f

                rs.Update
                rs.Bookmark = rs.LastModified
                newId = rs!id
                On Error Resume Next
                gCodes.Add newId, cd
                If Err.Number <> 0 Then
                    Say "ERROR: duplicate Code in the sheet: " & cd
                    gErr = gErr + 1
                    Err.Clear
                End If
                On Error GoTo 0
                n = n + 1
            End If
        End If
NextRow:
    Loop
    Close #ff
    rs.Close

    Say "T_STRUCTURES: " & n & " rows inserted"
    ImportStructures = n
End Function

' Parent_Code needs a second pass: a parent may appear later in the sheet.
Private Sub ResolveParents(db As DAO.Database)
    Dim i As Long
    Dim a As Variant
    Dim b As Variant
    Dim sql As String
    Dim n As Long

    For i = 0 To gPN - 1
        a = CodeId(gPCode(i))
        b = CodeId(gPParent(i))
        If IsNull(b) Then
            Say "ERROR: " & gPCode(i) & " - Parent_Code '" & gPParent(i) & "' is not in the sheet"
            gErr = gErr + 1
        Else
            sql = "UPDATE T_STRUCTURES SET ID_Parent=" & CLng(b) & " WHERE ID=" & CLng(a)
            db.Execute sql, dbFailOnError
            n = n + 1
        End If
    Next i
    If n > 0 Then Say "  parent links resolved: " & n
End Sub


' ================================================================
'  2. DECORATION ROWS (real ones and proposed rock art alike)
' ================================================================
Private Function ImportDecRows(db As DAO.Database, ByVal path As String, ByVal label As String) As Long
    Dim rs As DAO.Recordset
    Dim ff As Integer
    Dim ln As String
    Dim hdr() As String
    Dim val() As String
    Dim n As Long
    Dim pc As String
    Dim bc As String
    Dim dt As String
    Dim s As String
    Dim id As Variant
    Dim pid As Variant
    Dim bid As Variant
    Dim tid As Variant
    Dim okRow As Boolean

    If Len(Dir(path)) = 0 Then
        Say "T_DECORATIONS (" & label & "): file not present, skipped"
        Exit Function
    End If

    Set rs = db.OpenRecordset("T_DECORATIONS", dbOpenDynaset)
    ff = FreeFile
    Open path For Input As #ff
    Line Input #ff, ln
    hdr = SplitCsv(ln)

    Do While Not EOF(ff)
        Line Input #ff, ln
        If Len(Trim(ln)) > 0 Then
            val = SplitCsv(ln)
            pc = ColVal(hdr, val, "Parent_Code")
            bc = ColVal(hdr, val, "Body_Code")
            dt = ColVal(hdr, val, "Dec_Type_Name")

            ' Everything is resolved BEFORE AddNew: a row that cannot be
            ' completed must never be started, or a NOT NULL field blows
            ' up at Update with a runtime error instead of being reported.
            id = CodeId(pc)
            pid = Null
            bid = Null
            tid = Null
            okRow = True

            If IsNull(id) Then
                Say "ERROR: " & label & " - unknown Parent_Code '" & pc & "'"
                gErr = gErr + 1
                okRow = False
            Else
                pid = id
            End If

            If Len(bc) > 0 Then
                bid = IdWhere(db, "L_STRUCT_BODY", "Code", bc)
                If IsNull(bid) Then
                    Say "ERROR: " & label & " - unknown L_STRUCT_BODY code '" & bc & "' (" & pc & ")"
                    gErr = gErr + 1
                    okRow = False
                End If
            End If

            If Len(dt) > 0 Then
                tid = IdWhere(db, "L_DEC_TYPE", "Name", dt)
                If IsNull(tid) Then
                    Say "ERROR: " & label & " - unknown decoration type '" & dt & "' (" & pc & ")"
                    gErr = gErr + 1
                    okRow = False
                End If
            End If

            If okRow Then
                rs.AddNew
                rs!ID_Structure = CLng(pid)
                If IsNull(bid) Then rs!ID_Struct_Body = Null Else rs!ID_Struct_Body = CLng(bid)
                If IsNull(tid) Then rs!ID_Dec_Type = Null Else rs!ID_Dec_Type = CLng(tid)

                s = ColVal(hdr, val, "Body_No")
                If Len(s) = 0 Then rs!Body_No = Null Else rs!Body_No = CLng(NumStr(s))
                rs!Color = NzTxt(ColVal(hdr, val, "Color"))
                rs!Color_Secondary = NzTxt(ColVal(hdr, val, "Color_Secondary"))
                rs!Substrate = NzTxt(ColVal(hdr, val, "Substrate"))
                rs!Notes = NzTxt(ColVal(hdr, val, "Notes"))
                rs.Update
                n = n + 1
            End If
        End If
    Loop
    Close #ff
    rs.Close

    Say "T_DECORATIONS (" & label & "): " & n & " rows inserted"
    ImportDecRows = n
End Function


' ================================================================
'  3. LOST ELEMENTS
' ================================================================
Private Function ImportLost(db As DAO.Database, ByVal path As String) As Long
    Dim rs As DAO.Recordset
    Dim ff As Integer
    Dim ln As String
    Dim hdr() As String
    Dim val() As String
    Dim n As Long
    Dim pc As String
    Dim s As String
    Dim id As Variant
    Dim pid As Variant
    Dim eid As Variant
    Dim posid As Variant
    Dim okRow As Boolean

    If Len(Dir(path)) = 0 Then
        Say "T_LOST_ELEMENTS: file not present, skipped"
        Exit Function
    End If

    Set rs = db.OpenRecordset("T_LOST_ELEMENTS", dbOpenDynaset)
    ff = FreeFile
    Open path For Input As #ff
    Line Input #ff, ln
    hdr = SplitCsv(ln)

    Do While Not EOF(ff)
        Line Input #ff, ln
        If Len(Trim(ln)) > 0 Then
            val = SplitCsv(ln)
            ' ID_Evidence_Type is NOT NULL, so the row is fully resolved
            ' before AddNew. Starting a row that cannot be finished turns
            ' a reportable problem into a runtime error at Update.
            pc = ColVal(hdr, val, "Parent_Code")
            pid = CodeId(pc)
            eid = Null
            posid = Null
            okRow = True

            If IsNull(pid) Then
                Say "ERROR: lost elements - unknown Parent_Code '" & pc & "'"
                gErr = gErr + 1
                okRow = False
            End If

            s = ColVal(hdr, val, "Evidence_Name")
            If Len(s) = 0 Then
                Say "SKIPPED: lost elements - no evidence type for " & pc & ", the field is mandatory"
                okRow = False
            Else
                eid = IdWhere(db, "L_LOST_EVIDENCE", "Name", s)
                If IsNull(eid) Then
                    Say "ERROR: lost elements - '" & s & "' is not in L_LOST_EVIDENCE (" & pc & ")"
                    gErr = gErr + 1
                    okRow = False
                End If
            End If

            s = ColVal(hdr, val, "Position_Code")
            If Len(s) > 0 Then
                posid = IdWhere(db, "L_STRUCT_BODY", "Code", s)
                If IsNull(posid) Then
                    Say "ERROR: lost elements - unknown position code '" & s & "' (" & pc & ")"
                    gErr = gErr + 1
                    okRow = False
                End If
            End If

            If okRow Then
                rs.AddNew
                rs!ID_Structure = CLng(pid)
                rs!ID_Evidence_Type = CLng(eid)
                ' Element_Code has AllowZeroLength = False: blank must be NULL
                rs!Element_Code = NzTxt(ColVal(hdr, val, "Element_Code"))
                rs!Evidence_Scope = NzTxt(ColVal(hdr, val, "Evidence_Scope"))
                If IsNull(posid) Then rs!ID_Position = Null Else rs!ID_Position = CLng(posid)
                rs!Notes = NzTxt(ColVal(hdr, val, "Notes"))
                rs.Update
                n = n + 1
            End If
        End If
    Loop
    Close #ff
    rs.Close

    Say "T_LOST_ELEMENTS: " & n & " rows inserted"
    ImportLost = n
End Function


' ================================================================
'  LOOKUP COLUMN MAP
' ================================================================
Private Function IsLookupCol(ByVal nm As String) As Boolean
    IsLookupCol = (Len(LookupField(nm)) > 0)
End Function

Private Function LookupField(ByVal nm As String) As String
    Select Case nm
        Case "Sector_Name": LookupField = "ID_Sector"
        Case "Typology_Name": LookupField = "ID_Typology"
        Case "Support_Name": LookupField = "ID_Support"
        Case "Support_Secondary_Name": LookupField = "ID_Support_Secondary"
        Case "Arch_Status_Name": LookupField = "ID_Arch_Status"
        Case "Material_Status_Name": LookupField = "ID_Material_Status"
        Case "Vol_Method_Name": LookupField = "ID_Vol_Method"
        Case "Coord_Method_Name": LookupField = "ID_Coord_Method"
        Case "Campaign_Code": LookupField = "ID_Campaign"
        Case "Group_Code": LookupField = "ID_Group"
        Case Else: LookupField = ""
    End Select
End Function

Private Function LookupId(db As DAO.Database, ByVal nm As String, ByVal v As String) As Variant
    Select Case nm
        Case "Sector_Name": LookupId = IdWhere(db, "L_SECTORS", "Sector_Name", v)
        Case "Typology_Name": LookupId = IdWhere(db, "L_TYPOLOGY", "Name", v)
        Case "Support_Name": LookupId = IdWhere(db, "L_SUPPORT", "Name", v)
        Case "Support_Secondary_Name": LookupId = IdWhere(db, "L_SUPPORT", "Name", v)
        Case "Arch_Status_Name": LookupId = IdWhere(db, "L_STATUS", "Name", v)
        Case "Material_Status_Name": LookupId = IdWhere(db, "L_MATERIAL_STATUS", "Name", v)
        Case "Vol_Method_Name": LookupId = IdWhere(db, "L_VOL_METHOD", "Name", v)
        Case "Coord_Method_Name": LookupId = IdWhere(db, "L_COORD_METHOD", "Name", v)
        Case "Campaign_Code": LookupId = IdWhere(db, "L_CAMPAIGN", "Code", v)
        Case "Group_Code": LookupId = IdWhere(db, "T_GROUPS", "Group_Code", v)
        Case Else: LookupId = Null
    End Select
End Function


' ================================================================
'  PREFLIGHT: every text value against the width of its field
'
'  Error 3163 says only "the field is too small" and names neither
'  the field nor the value, which is useless on a 130 column sheet.
'  This reports every offender in one pass, before anything is
'  written.
' ================================================================
Private Function PreflightWidths(db As DAO.Database, ByVal path As String) As Boolean
    Dim td As DAO.TableDef
    Dim ff As Integer
    Dim ln As String
    Dim hdr() As String
    Dim val() As String
    Dim i As Long
    Dim nm As String
    Dim sv As String
    Dim cd As String
    Dim bad As Long

    If Len(Dir(path)) = 0 Then
        Say "MISSING: " & path
        PreflightWidths = False
        Exit Function
    End If

    Set td = db.TableDefs("T_STRUCTURES")
    ff = FreeFile
    Open path For Input As #ff
    Line Input #ff, ln
    hdr = SplitCsv(ln)

    Do While Not EOF(ff)
        Line Input #ff, ln
        If Len(Trim(ln)) > 0 Then
            val = SplitCsv(ln)
            cd = ColVal(hdr, val, "Code")
            For i = 0 To UBound(hdr)
                nm = Trim(hdr(i))
                If i <= UBound(val) Then sv = Trim(val(i)) Else sv = ""
                If Len(sv) > 0 And nm <> "ID_v11" And Right(nm, 4) <> "_v11" Then
                    If Not IsLookupCol(nm) And nm <> "Parent_Code" Then
                        If HasField(db, "T_STRUCTURES", nm) Then
                            If td.Fields(nm).Type = dbText Then
                                If Len(sv) > td.Fields(nm).Size Then
                                    Say "TOO LONG: " & cd & " - " & nm & " needs " & Len(sv) & " chars, the field is TEXT(" & td.Fields(nm).Size & "): '" & sv & "'"
                                    bad = bad + 1
                                End If
                            End If
                        End If
                    End If
                End If
            Next i
        End If
    Loop
    Close #ff

    If bad > 0 Then
        Say ""
        Say bad & " value(s) do not fit. Widen the field (see PatchV15Schema) or shorten the value."
        PreflightWidths = False
    Else
        Say "Preflight: every text value fits its field"
        PreflightWidths = True
    End If
End Function


' ================================================================
'  UTILITIES
' ================================================================
Private Function IdWhere(db As DAO.Database, ByVal tbl As String, ByVal fld As String, ByVal v As String) As Variant
    Dim rs As DAO.Recordset
    Dim sql As String
    If Len(v) = 0 Then
        IdWhere = Null
        Exit Function
    End If
    sql = "SELECT ID FROM " & tbl & " WHERE " & fld & "='" & Replace(v, "'", "''") & "'"
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    If rs.EOF Then IdWhere = Null Else IdWhere = rs!id
    rs.Close
End Function

Private Function CodeId(ByVal cd As String) As Variant
    On Error GoTo NotThere
    CodeId = gCodes.Item(cd)
    Exit Function
NotThere:
    CodeId = Null
End Function

Private Sub Defer(ByVal cd As String, ByVal parent As String)
    ReDim Preserve gPCode(gPN)
    ReDim Preserve gPParent(gPN)
    gPCode(gPN) = cd
    gPParent(gPN) = parent
    gPN = gPN + 1
End Sub

Private Function ColVal(hdr() As String, val() As String, ByVal nm As String) As String
    Dim i As Long
    For i = 0 To UBound(hdr)
        If Trim(hdr(i)) = nm Then
            If i <= UBound(val) Then
                ColVal = Trim(val(i))
            Else
                ColVal = ""
            End If
            Exit Function
        End If
    Next i
    ColVal = ""
End Function

Private Function NzTxt(ByVal s As String) As Variant
    If Len(Trim(s)) = 0 Then NzTxt = Null Else NzTxt = s
End Function

Private Function ToVal(ByVal s As String, f As DAO.Field) As Variant
    s = Trim(s)
    If Len(s) = 0 Then
        If f.Type = dbBoolean Then ToVal = False Else ToVal = Null
        Exit Function
    End If
    Select Case f.Type
        Case dbBoolean
            ToVal = (s = "1" Or s = "-1" Or LCase(s) = "true" Or LCase(s) = "yes" Or LCase(s) = "si")
        Case dbByte, dbInteger, dbLong
            ToVal = CLng(NumStr(s))
        Case dbSingle, dbDouble, dbCurrency, dbDecimal
            ToVal = CDbl(NumStr(s))
        Case Else
            ToVal = s
    End Select
End Function

' Accepts both decimal separators, whatever Excel wrote back.
Private Function NumStr(ByVal s As String) As String
    Dim ds As String
    ds = Mid(CStr(1.5), 2, 1)
    s = Replace(s, ".", ds)
    s = Replace(s, ",", ds)
    NumStr = s
End Function

' CSV splitter that respects quoted fields and doubled quotes.
Private Function SplitCsv(ByVal ln As String) As String()
    Dim out() As String
    Dim n As Long
    Dim i As Long
    Dim c As String
    Dim cur As String
    Dim inQ As Boolean

    ReDim out(0)
    n = 0
    inQ = False
    For i = 1 To Len(ln)
        c = Mid(ln, i, 1)
        If inQ Then
            If c = Chr(34) Then
                If i < Len(ln) And Mid(ln, i + 1, 1) = Chr(34) Then
                    cur = cur & Chr(34)
                    i = i + 1
                Else
                    inQ = False
                End If
            Else
                cur = cur & c
            End If
        Else
            If c = Chr(34) Then
                inQ = True
            ElseIf c = SEP Then
                ReDim Preserve out(n)
                out(n) = cur
                n = n + 1
                cur = ""
            Else
                cur = cur & c
            End If
        End If
    Next i
    ReDim Preserve out(n)
    out(n) = cur
    SplitCsv = out
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

Private Function CheckTarget(db As DAO.Database) As Boolean
    Dim ok As Boolean
    ok = True
    If Not HasField(db, "T_STRUCTURES", "Dec_Present") Then ok = False
    If Not HasField(db, "T_STRUCTURES", "RockArt_Present") Then ok = False
    If Not HasField(db, "T_STRUCTURES", "Cultural_Materials_Present") Then ok = False
    If Not HasField(db, "T_STRUCTURES", "Opening_Width_m") Then ok = False
    If Not HasField(db, "T_STRUCTURES", "Area_m2") Then ok = False
    If Not HasField(db, "T_DECORATIONS", "Color_Secondary") Then ok = False
    If Not ok Then Say "The target is missing v15 fields. Run BuildDB first."
    CheckTarget = ok
End Function

Private Function TargetDir() As String
    Dim d As String
    If Len(IMPORT_DIR) > 0 Then
        d = IMPORT_DIR
    Else
        d = CurrentProject.Path & "\transfer\"
    End If
    If Right(d, 1) <> "\" Then d = d & "\"
    If Len(Dir(d, vbDirectory)) = 0 Then
        MsgBox "Folder not found: " & d & vbCrLf & "Set IMPORT_DIR at the top of the module.", vbCritical
        TargetDir = ""
    Else
        TargetDir = d
    End If
End Function

Private Sub Say(ByVal s As String)
    Debug.Print s
    gLog = gLog & s & vbCrLf
End Sub

Private Sub WriteLogFile(ByVal d As String)
    Dim ff As Integer
    ff = FreeFile
    Open d & "IMPORT_LOG.txt" For Output As #ff
    Print #ff, "CHACHAPOYA IMPORT LOG - " & Now()
    Print #ff, gLog
    Close #ff
End Sub
