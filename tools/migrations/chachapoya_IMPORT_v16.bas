Option Compare Database
Option Explicit

' ================================================================
'  CHACHAPOYA - IMPORT CSV INTO A FRESHLY BUILT v16
'  Author: Esteve Ribera Torro | TFM Arqueologia UA
'  Spec: DELTA_v15_v16.md, section 10
'
'  RUN Sub ImportAll() ON AN EMPTY v16 DATABASE, i.e. one on which
'  chachapoya_DB_v16.bas -> BuildDB() has just run and into which
'  nothing has been typed. The lookups must be populated (BuildDB
'  does that); the data tables must be empty. ImportAll refuses to
'  run if T_STRUCTURES already has rows: a half-finished import
'  run twice is far harder to diagnose than one that refused.
'
'  EVERY FOREIGN KEY IS RESOLVED BY NAME OR CODE, never carried as
'  a number. See the header of the export script for why: ids
'  reorganised between versions produce wrong references that
'  raise no error and show no symptom.
'
'  ALL KEYS ARE RESOLVED BEFORE rs.AddNew. This is not style. In
'  the v11->v15 transfer, resolving a key after opening the new
'  record turned a clean "lookup value not found" into a crash at
'  rs.Update with the record half-written, which says nothing
'  about what was actually wrong. Resolve first, report by row,
'  then write.
'
'  UNKNOWN COLUMNS ARE SKIPPED AND NAMED. A column in the CSV with
'  no matching field in v16 is not an error to abort on - it is
'  exactly what a withdrawn field looks like from here
'  (Interbody_Cornice_Material). It is reported once, by name, so
'  that a column skipped by ACCIDENT is visible too.
'
'  BLANK MEANS NULL, ALWAYS. Never the empty string: an empty
'  string satisfies no foreign key, fails rule 17 silently, and
'  is invisible in the datasheet. This is the same reason
'  AllowZeroLength is False on T_LOST_ELEMENTS.Element_Code.
' ================================================================

Private Const IN_DIR As String = "C:\chachapoya_transfer\"

' Columns the importer deliberately ignores.
Private Const IGNORE_PREFIX As String = "OLD_"

Private mCols() As String
Private mRows() As String
Private mNCol As Long
Private mNRow As Long

Sub ImportAll()
    Dim db As DAO.Database
    Set db = CurrentDb()

    If DCount("*", "T_STRUCTURES") > 0 Then
        MsgBox "T_STRUCTURES is not empty. Run this only on a freshly built v16." & vbCrLf & vbCrLf & "Build a blank database with chachapoya_DB_v16.bas -> BuildDB() and import into that.", vbCritical, "Import aborted"
        Exit Sub
    End If

    If Dir(IN_DIR, vbDirectory) = "" Then
        MsgBox "The folder " & IN_DIR & " does not exist. Edit IN_DIR at the top of this module.", vbCritical, "Import aborted"
        Exit Sub
    End If

    Debug.Print "=== IMPORT CSV -> v16 ==="

    Dim nG As Long
    Dim nS As Long
    Dim nP As Long
    Dim nD As Long
    Dim nDat As Long
    Dim nInd As Long
    Dim nF As Long
    Dim nC As Long
    Dim nL As Long

    nG = ImportGroups(db)
    nS = ImportStructures(db)
    nP = LinkParents(db)
    nD = ImportDecorations(db)
    nDat = ImportDating(db)
    nInd = ImportIndividuals(db)
    nF = ImportFeatures(db)
    nC = ImportConnections(db)
    nL = ImportLost(db)

    Set db = Nothing

    Dim msg As String
    msg = "IMPORT COMPLETE" & vbCrLf & vbCrLf
    msg = msg & "  T_GROUPS:         " & nG & vbCrLf
    msg = msg & "  T_STRUCTURES:     " & nS & vbCrLf
    msg = msg & "  parent links set: " & nP & vbCrLf
    msg = msg & "  T_DECORATIONS:    " & nD & vbCrLf
    msg = msg & "  T_DATING:         " & nDat & vbCrLf
    msg = msg & "  T_INDIVIDUALS:    " & nInd & vbCrLf
    msg = msg & "  T_ARCH_FEATURES:  " & nF & vbCrLf
    msg = msg & "  T_CONNECTIONS:    " & nC & vbCrLf
    msg = msg & "  T_LOST_ELEMENTS:  " & nL & vbCrLf & vbCrLf
    msg = msg & "CHECK THE IMMEDIATE WINDOW (Ctrl+G) before" & vbCrLf
    msg = msg & "doing anything else. Rows that could not be" & vbCrLf
    msg = msg & "written are listed there by source code, and" & vbCrLf
    msg = msg & "so are any CSV columns that were skipped." & vbCrLf & vbCrLf
    msg = msg & "THEN, IN ORDER:" & vbCrLf
    msg = msg & "  1. QRY_16_Validation_Check (39 rules)" & vbCrLf
    msg = msg & "  2. QRY_19_V16_Review (the manual worklist)" & vbCrLf
    msg = msg & "  3. Re-run the A-X matrix and compare with" & vbCrLf
    msg = msg & "     the previous results: the NA filter on" & vbCrLf
    msg = msg & "     natural roofs changes them."
    MsgBox msg, vbInformation, "Import v16"
    Debug.Print "=== IMPORT FINISHED ==="
End Sub

' ================================================================
'  CSV PARSER
'  Reads the whole file and walks it character by character. A
'  split on commas would be shorter and would break on the first
'  Notes field containing one; a split on line breaks would break
'  on the first memo containing one. Neither failure announces
'  itself - the columns just shift and the data becomes quietly
'  wrong, which is the failure mode this whole two-step transfer
'  exists to avoid.
' ================================================================
Private Function LoadCsv(fname As String) As Boolean
    Dim fn As Integer
    Dim buf As String
    Dim p As Long
    Dim ch As String
    Dim inQ As Boolean

    If Dir(IN_DIR & fname) = "" Then
        Debug.Print "  *** file not found: " & fname
        LoadCsv = False
        Exit Function
    End If

    ' Binary, not Input. Opening for Input reads in text mode, which
    ' can stop early on a stray end-of-file character inside a memo;
    ' binary gives back exactly the bytes on disk.
    fn = FreeFile
    Open IN_DIR & fname For Binary As #fn
    If LOF(fn) > 0 Then
        buf = Space$(LOF(fn))
        Get #fn, , buf
    End If
    Close #fn

    ' Fields are separated by Chr(1) and rows by Chr(2): both are
    ' control characters that cannot occur in the source data, so
    ' the split that follows cannot be confused by content.
    Dim outp As String
    inQ = False
    For p = 1 To Len(buf)
        ch = Mid$(buf, p, 1)
        If inQ Then
            If ch = """" Then
                If p < Len(buf) Then
                    If Mid$(buf, p + 1, 1) = """" Then
                        outp = outp & """"
                        p = p + 1
                    Else
                        inQ = False
                    End If
                Else
                    inQ = False
                End If
            Else
                outp = outp & ch
            End If
        Else
            If ch = """" Then
                inQ = True
            ElseIf ch = "," Then
                outp = outp & Chr(1)
            ElseIf ch = vbCr Then
                ' ignored: the line feed closes the row
            ElseIf ch = vbLf Then
                outp = outp & Chr(2)
            Else
                outp = outp & ch
            End If
        End If
    Next p

    Dim allRows() As String
    allRows = Split(outp, Chr(2))

    ' Header
    mCols = Split(allRows(0), Chr(1))
    mNCol = UBound(mCols) + 1
    Dim i As Long
    For i = 0 To UBound(mCols)
        mCols(i) = Trim$(mCols(i))
    Next i

    ' Body, skipping trailing blank rows
    Dim keep As Long
    keep = 0
    ReDim mRows(UBound(allRows))
    For i = 1 To UBound(allRows)
        If Len(Trim$(Replace(allRows(i), Chr(1), ""))) > 0 Then
            mRows(keep) = allRows(i)
            keep = keep + 1
        End If
    Next i
    mNRow = keep
    LoadCsv = True
End Function

' Value of a named column in a parsed row. Returns "" when the
' column does not exist, which the callers treat as absent.
Private Function CV(row As String, colName As String) As String
    Dim i As Long
    Dim parts() As String
    parts = Split(row, Chr(1))
    For i = 0 To UBound(mCols)
        If StrComp(mCols(i), colName, vbTextCompare) = 0 Then
            If i <= UBound(parts) Then CV = parts(i)
            Exit Function
        End If
    Next i
End Function

' Blank -> Null, never the empty string.
Private Function NV(s As String) As Variant
    If Len(Trim$(s)) = 0 Then
        NV = Null
    Else
        NV = s
    End If
End Function

Private Function ColIndex(colName As String) As Long
    Dim i As Long
    ColIndex = -1
    For i = 0 To UBound(mCols)
        If StrComp(mCols(i), colName, vbTextCompare) = 0 Then
            ColIndex = i
            Exit Function
        End If
    Next i
End Function

' ================================================================
'  KEY RESOLUTION
'  Returns 0 when unresolved. Callers decide whether 0 is fatal
'  for that row - a null typology is legitimate, a null sector is
'  not, and the schema already says which is which.
' ================================================================
Private Function IdOf(db As DAO.Database, tbl As String, col As String, v As String) As Long
    If Len(Trim$(v)) = 0 Then Exit Function
    Dim rs As DAO.Recordset
    Dim q2 As String
    q2 = "SELECT ID FROM " & tbl & " WHERE " & col & "='" & Replace(v, "'", "''") & "'"
    Set rs = db.OpenRecordset(q2, dbOpenSnapshot)
    If Not rs.EOF Then IdOf = rs!ID
    rs.Close
    Set rs = Nothing
End Function

Private Function SectorId(db As DAO.Database, site As String, sector As String) As Long
    If Len(Trim$(sector)) = 0 Then Exit Function
    Dim rs As DAO.Recordset
    Dim q2 As String
    q2 = "SELECT SC.ID FROM L_SECTORS AS SC "
    q2 = q2 & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID "
    q2 = q2 & "WHERE SC.Sector_Name='" & Replace(sector, "'", "''") & "'"
    If Len(Trim$(site)) > 0 Then
        q2 = q2 & " AND S.Site_Name='" & Replace(site, "'", "''") & "'"
    End If
    Set rs = db.OpenRecordset(q2, dbOpenSnapshot)
    If Not rs.EOF Then SectorId = rs!ID
    rs.Close
    Set rs = Nothing
End Function

' L_STRUCT_BODY by code, with the v16 remapping applied.
' SOC is withdrawn (delta 2.1) and its rows belong on BAS: the
' socle was never a position distinct from the basal body, only a
' treatment of it. Done here rather than in the export so that
' the CSV still shows what the source said.
Private Function BodyId(db As DAO.Database, code As String, ByRef remapped As Long) As Long
    Dim c As String
    c = Trim$(code)
    If Len(c) = 0 Then Exit Function
    If StrComp(c, "SOC", vbTextCompare) = 0 Then
        c = "BAS"
        remapped = remapped + 1
    End If
    BodyId = IdOf(db, "L_STRUCT_BODY", "Code", c)
End Function

Private Function StructId(db As DAO.Database, code As String) As Long
    StructId = IdOf(db, "T_STRUCTURES", "Code", code)
End Function

' ================================================================
'  T_GROUPS - first, because T_STRUCTURES points at it
' ================================================================
Private Function ImportGroups(db As DAO.Database) As Long
    If Not LoadCsv("T_GROUPS.csv") Then Exit Function
    Dim rs As DAO.Recordset
    Dim i As Long
    Dim n As Long
    Dim bad As Long
    Set rs = db.OpenRecordset("T_GROUPS", dbOpenDynaset)
    For i = 0 To mNRow - 1
        Dim r As String
        r = mRows(i)
        Dim gc As String
        Dim sid As Long
        Dim gt As Long
        gc = CV(r, "Group_Code")
        sid = SectorId(db, CV(r, "Site"), CV(r, "Sector"))
        gt = IdOf(db, "L_GROUP_TYPE", "Name", CV(r, "Group_Type"))
        ' Both are NOT NULL in the schema: resolve and check BEFORE
        ' AddNew, or the failure surfaces at Update instead of here.
        If sid = 0 Or gt = 0 Then
            Debug.Print "  *** T_GROUPS row " & (i + 2) & " (" & gc & "): unresolved sector or group type - skipped"
            bad = bad + 1
        Else
            rs.AddNew
            rs!Group_Code = gc
            rs!ID_Sector = sid
            rs!ID_Group_Type = gt
            rs!N_Members = NV(CV(r, "N_Members"))
            rs!Notes = NV(CV(r, "Notes"))
            rs.Update
            n = n + 1
        End If
    Next i
    rs.Close
    Debug.Print "-> T_GROUPS: " & n & " written, " & bad & " skipped"
    ImportGroups = n
End Function

' ================================================================
'  T_STRUCTURES
'  Two passes. The self-reference cannot be resolved on the way
'  in, because the parent may not exist yet; LinkParents does it
'  afterwards, when every code is present.
' ================================================================
Private Function ImportStructures(db As DAO.Database) As Long
    If Not LoadCsv("T_STRUCTURES.csv") Then Exit Function

    Dim td As DAO.TableDef
    Set td = db.TableDefs("T_STRUCTURES")

    ' Report the columns this run will not use, once, by name.
    Dim j As Long
    Dim skipped As String
    For j = 0 To UBound(mCols)
        If Left$(mCols(j), Len(IGNORE_PREFIX)) <> IGNORE_PREFIX Then
            If Not IsResolvedCol(mCols(j)) Then
                If Not FieldExists(td, mCols(j)) Then
                    skipped = skipped & mCols(j) & " "
                End If
            End If
        End If
    Next j
    If Len(skipped) > 0 Then
        Debug.Print "  [info] CSV columns with no field in v16, ignored: " & skipped
    End If

    Dim rs As DAO.Recordset
    Dim i As Long
    Dim n As Long
    Dim bad As Long
    Set rs = db.OpenRecordset("T_STRUCTURES", dbOpenDynaset)

    For i = 0 To mNRow - 1
        Dim r As String
        r = mRows(i)
        Dim cd As String
        cd = CV(r, "Code")

        ' Resolve everything first (see header).
        Dim sid As Long
        Dim tid As Long
        Dim su1 As Long
        Dim su2 As Long
        Dim gid As Long
        Dim ast As Long
        Dim mst As Long
        Dim vmt As Long
        Dim cmt As Long
        Dim cmp As Long
        sid = SectorId(db, CV(r, "Site"), CV(r, "Sector"))
        tid = IdOf(db, "L_TYPOLOGY", "Name", CV(r, "Typology"))
        su1 = IdOf(db, "L_SUPPORT", "Name", CV(r, "Support"))
        su2 = IdOf(db, "L_SUPPORT", "Name", CV(r, "Support_Secondary"))
        gid = IdOf(db, "T_GROUPS", "Group_Code", CV(r, "Group_Code"))
        ast = IdOf(db, "L_STATUS", "Name", CV(r, "Arch_Status"))
        mst = IdOf(db, "L_MATERIAL_STATUS", "Name", CV(r, "Material_Status"))
        vmt = IdOf(db, "L_VOL_METHOD", "Name", CV(r, "Vol_Method"))
        cmt = IdOf(db, "L_COORD_METHOD", "Name", CV(r, "Coord_Method"))
        cmp = IdOf(db, "L_CAMPAIGN", "Code", CV(r, "Campaign"))

        If Len(Trim$(cd)) = 0 Then
            Debug.Print "  *** T_STRUCTURES row " & (i + 2) & ": no Code - skipped"
            bad = bad + 1
        ElseIf sid = 0 Then
            Debug.Print "  *** T_STRUCTURES " & cd & ": sector not resolved (" & CV(r, "Site") & " / " & CV(r, "Sector") & ") - skipped"
            bad = bad + 1
        Else
            ' A named lookup value that does not resolve is worth
            ' saying out loud even when the field is nullable: it
            ' means a renamed value, not an empty one.
            WarnUnresolved cd, "Typology", CV(r, "Typology"), tid
            WarnUnresolved cd, "Support", CV(r, "Support"), su1
            WarnUnresolved cd, "Support_Secondary", CV(r, "Support_Secondary"), su2
            WarnUnresolved cd, "Group_Code", CV(r, "Group_Code"), gid
            WarnUnresolved cd, "Arch_Status", CV(r, "Arch_Status"), ast
            WarnUnresolved cd, "Material_Status", CV(r, "Material_Status"), mst
            WarnUnresolved cd, "Vol_Method", CV(r, "Vol_Method"), vmt
            WarnUnresolved cd, "Coord_Method", CV(r, "Coord_Method"), cmt
            WarnUnresolved cd, "Campaign", CV(r, "Campaign"), cmp

            rs.AddNew
            rs!Code = cd
            rs!ID_Sector = sid
            If tid > 0 Then rs!ID_Typology = tid
            If su1 > 0 Then rs!ID_Support = su1
            If su2 > 0 Then rs!ID_Support_Secondary = su2
            If gid > 0 Then rs!ID_Group = gid
            If ast > 0 Then rs!ID_Arch_Status = ast
            If mst > 0 Then rs!ID_Material_Status = mst
            If vmt > 0 Then rs!ID_Vol_Method = vmt
            If cmt > 0 Then rs!ID_Coord_Method = cmt
            If cmp > 0 Then rs!ID_Campaign = cmp

            Dim fld As DAO.Field
            For Each fld In td.Fields
                If fld.Name <> "ID" And fld.Name <> "Code" Then
                    If Left$(fld.Name, 3) <> "ID_" Then
                        If ColIndex(fld.Name) >= 0 Then
                            AssignField rs, fld, CV(r, fld.Name)
                        End If
                    End If
                End If
            Next fld

            rs.Update
            n = n + 1
        End If
    Next i
    rs.Close
    Debug.Print "-> T_STRUCTURES: " & n & " written, " & bad & " skipped"
    ImportStructures = n
End Function

' Writes one value with the type the field expects. Blank always
' becomes Null: a zero written where the CSV was blank would be a
' verified absence the observer never asserted, which is exactly
' the bias the whole default policy exists to prevent.
Private Sub AssignField(rs As DAO.Recordset, fld As DAO.Field, v As String)
    On Error GoTo Err_AF
    If Len(Trim$(v)) = 0 Then
        rs.Fields(fld.Name).Value = Null
        Exit Sub
    End If
    Select Case fld.Type
        Case dbBoolean
            rs.Fields(fld.Name).Value = (v = "1" Or StrComp(v, "True", vbTextCompare) = 0 Or v = "-1")
        Case dbByte, dbInteger, dbLong
            rs.Fields(fld.Name).Value = CLng(v)
        Case dbSingle, dbDouble, dbCurrency
            rs.Fields(fld.Name).Value = CDbl(Replace(v, ",", "."))
        Case dbDate
            rs.Fields(fld.Name).Value = CDate(v)
        Case Else
            rs.Fields(fld.Name).Value = v
    End Select
    Exit Sub
Err_AF:
    Debug.Print "  *** could not assign " & fld.Name & " = '" & v & "': " & Err.Description
    Err.Clear
End Sub

Private Function FieldExists(td As DAO.TableDef, n As String) As Boolean
    Dim f As DAO.Field
    For Each f In td.Fields
        If StrComp(f.Name, n, vbTextCompare) = 0 Then
            FieldExists = True
            Exit Function
        End If
    Next f
End Function

' The thirteen columns that carry resolved keys rather than field
' values. Without this the skip report would flag them all.
Private Function IsResolvedCol(n As String) As Boolean
    Const L As String = "|Code|Site|Sector|Typology|Support|Support_Secondary|Parent_Code|Group_Code|Arch_Status|Material_Status|Vol_Method|Coord_Method|Campaign|"
    IsResolvedCol = (InStr(1, L, "|" & n & "|", vbTextCompare) > 0)
End Function

Private Sub WarnUnresolved(code As String, colName As String, v As String, id As Long)
    If Len(Trim$(v)) > 0 And id = 0 Then
        Debug.Print "  [warn] " & code & ": " & colName & " '" & v & "' did not resolve - left null"
    End If
End Sub

' ================================================================
'  SECOND PASS: the self-referencing parent link
' ================================================================
Private Function LinkParents(db As DAO.Database) As Long
    If Not LoadCsv("T_STRUCTURES.csv") Then Exit Function
    Dim i As Long
    Dim n As Long
    For i = 0 To mNRow - 1
        Dim cd As String
        Dim pc As String
        cd = CV(mRows(i), "Code")
        pc = CV(mRows(i), "Parent_Code")
        If Len(Trim$(pc)) > 0 Then
            Dim pid As Long
            Dim cid As Long
            pid = StructId(db, pc)
            cid = StructId(db, cd)
            If pid = 0 Or cid = 0 Then
                Debug.Print "  *** parent link " & cd & " -> " & pc & ": one of the two codes does not exist"
            Else
                db.Execute "UPDATE T_STRUCTURES SET ID_Parent=" & pid & " WHERE ID=" & cid, dbFailOnError
                n = n + 1
            End If
        End If
    Next i
    Debug.Print "-> parent links: " & n
    LinkParents = n
End Function

' ================================================================
'  T_DECORATIONS
' ================================================================
Private Function ImportDecorations(db As DAO.Database) As Long
    If Not LoadCsv("T_DECORATIONS.csv") Then Exit Function
    Dim rs As DAO.Recordset
    Dim i As Long
    Dim n As Long
    Dim bad As Long
    Dim remap As Long
    Set rs = db.OpenRecordset("T_DECORATIONS", dbOpenDynaset)
    For i = 0 To mNRow - 1
        Dim r As String
        r = mRows(i)
        Dim sid As Long
        Dim bid As Long
        Dim did As Long
        Dim sc As String
        sc = CV(r, "Structure_Code")
        sid = StructId(db, sc)
        bid = BodyId(db, CV(r, "Position_Code"), remap)
        did = IdOf(db, "L_DEC_TYPE", "Name", CV(r, "Dec_Type"))
        If sid = 0 Then
            Debug.Print "  *** T_DECORATIONS row " & (i + 2) & ": structure '" & sc & "' not found - skipped"
            bad = bad + 1
        Else
            WarnUnresolved sc, "Position_Code", CV(r, "Position_Code"), bid
            WarnUnresolved sc, "Dec_Type", CV(r, "Dec_Type"), did
            rs.AddNew
            rs!ID_Structure = sid
            If bid > 0 Then rs!ID_Struct_Body = bid
            If did > 0 Then rs!ID_Dec_Type = did
            rs!Body_No = NV(CV(r, "Body_No"))
            ' Position_Relative is new in v16 and has no source
            ' column: it stays Null, which is correct - nobody has
            ' judged it yet.
            rs!Color = NV(CV(r, "Color"))
            rs!Color_Secondary = NV(CV(r, "Color_Secondary"))
            rs!Substrate = NV(CV(r, "Substrate"))
            rs!Notes = NV(CV(r, "Notes"))
            rs.Update
            n = n + 1
        End If
    Next i
    rs.Close
    Debug.Print "-> T_DECORATIONS: " & n & " written, " & bad & " skipped"
    Debug.Print "   SOC rows remapped to BAS: " & remap
    ImportDecorations = n
End Function

Private Function ImportDating(db As DAO.Database) As Long
    If Not LoadCsv("T_DATING.csv") Then Exit Function
    Dim rs As DAO.Recordset
    Dim i As Long
    Dim n As Long
    Dim bad As Long
    Set rs = db.OpenRecordset("T_DATING", dbOpenDynaset)
    For i = 0 To mNRow - 1
        Dim r As String
        r = mRows(i)
        Dim sid As Long
        sid = StructId(db, CV(r, "Structure_Code"))
        If sid = 0 Then
            Debug.Print "  *** T_DATING row " & (i + 2) & ": structure not found - skipped"
            bad = bad + 1
        Else
            rs.AddNew
            rs!ID_Structure = sid
            rs!Sample_Type = NV(CV(r, "Sample_Type"))
            rs!Date_BP = NV(CV(r, "Date_BP"))
            rs!Sigma1_Start = NV(CV(r, "Sigma1_Start"))
            rs!Sigma1_End = NV(CV(r, "Sigma1_End"))
            rs!Sigma2_Start = NV(CV(r, "Sigma2_Start"))
            rs!Sigma2_End = NV(CV(r, "Sigma2_End"))
            rs!Lab_Reference = NV(CV(r, "Lab_Reference"))
            rs!Bibliog_Reference = NV(CV(r, "Bibliog_Reference"))
            rs.Update
            n = n + 1
        End If
    Next i
    rs.Close
    Debug.Print "-> T_DATING: " & n & " written, " & bad & " skipped"
    ImportDating = n
End Function

Private Function ImportIndividuals(db As DAO.Database) As Long
    If Not LoadCsv("T_INDIVIDUALS.csv") Then Exit Function
    Dim rs As DAO.Recordset
    Dim i As Long
    Dim n As Long
    Dim bad As Long
    Set rs = db.OpenRecordset("T_INDIVIDUALS", dbOpenDynaset)
    For i = 0 To mNRow - 1
        Dim r As String
        r = mRows(i)
        Dim sid As Long
        sid = StructId(db, CV(r, "Structure_Code"))
        If sid = 0 Then
            Debug.Print "  *** T_INDIVIDUALS row " & (i + 2) & ": structure not found - skipped"
            bad = bad + 1
        Else
            rs.AddNew
            rs!ID_Structure = sid
            rs!Individual_No = NV(CV(r, "Individual_No"))
            rs!Age_Category = NV(CV(r, "Age_Category"))
            rs!Sex_Category = NV(CV(r, "Sex_Category"))
            rs!Preservation = NV(CV(r, "Preservation"))
            rs!Notes = NV(CV(r, "Notes"))
            rs.Update
            n = n + 1
        End If
    Next i
    rs.Close
    Debug.Print "-> T_INDIVIDUALS: " & n & " written, " & bad & " skipped"
    ImportIndividuals = n
End Function

Private Function ImportFeatures(db As DAO.Database) As Long
    If Not LoadCsv("T_ARCH_FEATURES.csv") Then Exit Function
    Dim rs As DAO.Recordset
    Dim i As Long
    Dim n As Long
    Dim bad As Long
    Set rs = db.OpenRecordset("T_ARCH_FEATURES", dbOpenDynaset)
    For i = 0 To mNRow - 1
        Dim r As String
        r = mRows(i)
        Dim sid As Long
        sid = StructId(db, CV(r, "Structure_Code"))
        If sid = 0 Then
            Debug.Print "  *** T_ARCH_FEATURES row " & (i + 2) & ": structure not found - skipped"
            bad = bad + 1
        Else
            rs.AddNew
            rs!ID_Structure = sid
            rs!Feature_Code = NV(CV(r, "Feature_Code"))
            rs!Present = (CV(r, "Present") = "1")
            rs!Feature_Count = NV(CV(r, "Feature_Count"))
            rs!Material = NV(CV(r, "Material"))
            rs!Notes = NV(CV(r, "Notes"))
            rs.Update
            n = n + 1
        End If
    Next i
    rs.Close
    Debug.Print "-> T_ARCH_FEATURES: " & n & " written, " & bad & " skipped"
    ImportFeatures = n
End Function

Private Function ImportConnections(db As DAO.Database) As Long
    If Not LoadCsv("T_CONNECTIONS.csv") Then Exit Function
    Dim rs As DAO.Recordset
    Dim i As Long
    Dim n As Long
    Dim bad As Long
    Set rs = db.OpenRecordset("T_CONNECTIONS", dbOpenDynaset)
    For i = 0 To mNRow - 1
        Dim r As String
        r = mRows(i)
        Dim a As Long
        Dim b As Long
        a = StructId(db, CV(r, "Code_A"))
        b = StructId(db, CV(r, "Code_B"))
        If a = 0 Or b = 0 Then
            Debug.Print "  *** T_CONNECTIONS row " & (i + 2) & ": " & CV(r, "Code_A") & " / " & CV(r, "Code_B") & " - one end not found, skipped"
            bad = bad + 1
        Else
            rs.AddNew
            rs!ID_Struct_A = a
            rs!ID_Struct_B = b
            rs!Connection_Type = NV(CV(r, "Connection_Type"))
            rs!Chrono_Relation = NV(CV(r, "Chrono_Relation"))
            rs!Confidence = NV(CV(r, "Confidence"))
            rs!Notes = NV(CV(r, "Notes"))
            rs.Update
            n = n + 1
        End If
    Next i
    rs.Close
    Debug.Print "-> T_CONNECTIONS: " & n & " written, " & bad & " skipped"
    ImportConnections = n
End Function

' ================================================================
'  T_LOST_ELEMENTS
'  ID_Evidence_Type is NOT NULL, so an unresolved evidence type
'  must stop the row here rather than at rs.Update. This is the
'  exact failure the v11->v15 transfer hit.
' ================================================================
Private Function ImportLost(db As DAO.Database) As Long
    If Not LoadCsv("T_LOST_ELEMENTS.csv") Then Exit Function
    Dim rs As DAO.Recordset
    Dim i As Long
    Dim n As Long
    Dim bad As Long
    Dim remap As Long
    Set rs = db.OpenRecordset("T_LOST_ELEMENTS", dbOpenDynaset)
    For i = 0 To mNRow - 1
        Dim r As String
        r = mRows(i)
        Dim sid As Long
        Dim ev As Long
        Dim pid As Long
        Dim ec As String
        Dim sc As String
        sc = CV(r, "Structure_Code")
        sid = StructId(db, sc)
        ev = IdOf(db, "L_LOST_EVIDENCE", "Name", CV(r, "Evidence_Type"))
        pid = BodyId(db, CV(r, "Position_Code"), remap)
        ec = Trim$(CV(r, "Element_Code"))
        If sid = 0 Then
            Debug.Print "  *** T_LOST_ELEMENTS row " & (i + 2) & ": structure '" & sc & "' not found - skipped"
            bad = bad + 1
        ElseIf ev = 0 Then
            Debug.Print "  *** T_LOST_ELEMENTS row " & (i + 2) & " (" & sc & "): evidence type '" & CV(r, "Evidence_Type") & "' not resolved and the field is NOT NULL - skipped"
            bad = bad + 1
        Else
            rs.AddNew
            rs!ID_Structure = sid
            rs!ID_Evidence_Type = ev
            ' AllowZeroLength is False on this field: an empty
            ' string would be rejected, so blank must be Null.
            If Len(ec) > 0 Then rs!Element_Code = ec
            rs!Evidence_Scope = NV(CV(r, "Evidence_Scope"))
            If pid > 0 Then rs!ID_Position = pid
            rs!Notes = NV(CV(r, "Notes"))
            rs.Update
            n = n + 1
        End If
    Next i
    rs.Close
    Debug.Print "-> T_LOST_ELEMENTS: " & n & " written, " & bad & " skipped"
    Debug.Print "   SOC positions remapped to BAS: " & remap
    ImportLost = n
End Function
