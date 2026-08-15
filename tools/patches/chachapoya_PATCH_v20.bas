Option Compare Database
Option Explicit

' ================================================================
'  CHACHAPOYA PATCH v19a -> v20 (in place, data preserved)
'  Author: Esteve Ribera Torro | TFM Arqueologia UA
'  Spec: DELTA_v19_v20.md
'
'  WHAT THIS DOES, in order:
'    0. Preflight: v19 marker (Jamb_Fabric_Reveal) and, above
'       all, DUPLICATE CODES. Identity blocks, it does not warn:
'       if two structures share a code the patch aborts listing
'       them, because step 6 (the unique index) would fail and
'       every cross-reference is ambiguous until they are fixed.
'    1. DROP Interbody_Cornice_Format (delta bloc 2). Recorded
'       loss: DW-S01-EA15 and DW-S01-EA36 carried 'Tabular
'       blocks'. Rule 50 retires with the field.
'    2. DROP T_ARCH_FEATURES.Present (bloc I): a row exists
'       because something was observed - the checkbox duplicated
'       the row's own existence and its default 0 made 11 of 16
'       rows lie.
'    3. ADD Upper_Crown_Format (bloc C): the form of the crown -
'       closing panel, projecting course, flush course, ND.
'       Default NULL: a judgement field, not a qualifier.
'    4. ADD L_SECTORS.Sector_Code + population BY NAME (bloc B):
'       the prefix every structure code must carry. Rule 57
'       watches the concordance the EA46 case broke silently.
'    5. UNIQUE INDEX on T_STRUCTURES.Code (bloc B): from now on
'       Access rejects a duplicate the moment it is typed.
'    6. Lookup updates BY NAME, never by ID: L_ELEMENTS row D
'       (Transverse wall or pier), L_TYPOLOGY MEN (examples with
'       final terms), L_DEC_TYPE outline bands (verdict-free
'       descriptions; the two geometric rows of the corpus are
'       retyped by the rename itself, rows point by ID).
'    7. Data cleanups: Phase_Evidence back to NULL where a single
'       phase left it without a subject (bloc G, 3 rows);
'       Interior_Observability back to NULL on rock art panels
'       (bloc H, 4 rows, values noted in the delta).
'
'  AFTER RUNNING THIS, IN ORDER:
'    1. chachapoya_DB_v20.bas -> RebuildQueriesV20()
'    2. chachapoya_Form_v20_val.bas -> BuildForm()
'    3. Open QRY_25_V20_Review: the worklist of this migration.
'
'  Every step is guarded, so re-running the patch is harmless.
' ================================================================

Sub PatchV20()
    Dim db As DAO.Database
    Set db = CurrentDb()

    Debug.Print "=== PATCH v19a -> v20 ==="

    ' Preflight 1: schema generation.
    If Not TableExists(db, "T_STRUCTURES") Then
        MsgBox "T_STRUCTURES does not exist. This is not a Chachapoya database.", vbCritical, "Patch aborted"
        Exit Sub
    End If
    If Not ColExists(db, "T_STRUCTURES", "Jamb_Fabric_Reveal") Then
        MsgBox "Jamb_Fabric_Reveal is missing. This database is older than v19 - apply PATCH v19 first.", vbCritical, "Patch aborted"
        Exit Sub
    End If

    ' Preflight 2: duplicate codes ABORT the patch.
    Dim rs As DAO.Recordset
    Dim dupList As String
    Set rs = db.OpenRecordset("SELECT Code, Count(*) AS N FROM T_STRUCTURES GROUP BY Code HAVING Count(*) > 1")
    Do While Not rs.EOF
        dupList = dupList & "  " & rs!Code & " (x" & rs!N & ")" & vbCrLf
        rs.MoveNext
    Loop
    rs.Close
    If Len(dupList) > 0 Then
        MsgBox "DUPLICATE CODES FOUND - fix them first, then re-run:" & vbCrLf & vbCrLf & dupList, vbCritical, "Patch aborted"
        Exit Sub
    End If

    Dim nDrop As Long
    Dim nAdd As Long

    ' ---------------------------------------------------------------
    ' 1-2. Column removals (bloc 2, bloc I)
    ' ---------------------------------------------------------------
    nDrop = nDrop + DropCol(db, "T_STRUCTURES", "Interbody_Cornice_Format")
    nDrop = nDrop + DropCol(db, "T_ARCH_FEATURES", "Present")

    ' ---------------------------------------------------------------
    ' 3. Upper_Crown_Format (bloc C)
    ' ---------------------------------------------------------------
    nAdd = nAdd + AddCol(db, "T_STRUCTURES", "Upper_Crown_Format", "TEXT(25)")

    ' ---------------------------------------------------------------
    ' 4. Sector_Code (bloc B) - population BY NAME
    ' ---------------------------------------------------------------
    nAdd = nAdd + AddCol(db, "L_SECTORS", "Sector_Code", "TEXT(10)")
    db.TableDefs.Refresh
    SetSector db, "LP - General", "LP-G"
    SetSector db, "LP - North", "LP-N"
    SetSector db, "LP - Central", "LP-C"
    SetSector db, "LP - Upper", "LP-U"
    SetSector db, "LP - South", "LP-S"
    SetSector db, "DW - General", "DW-G"
    SetSector db, "DW - Sector 1", "DW-S01"
    SetSector db, "DW - Sector 2", "DW-S02"
    SetSector db, "DW - Sector 3", "DW-S03"
    SetSector db, "DW - Sector 4", "DW-S04"
    SetSector db, "DW - Sector 5", "DW-S05"
    SetSector db, "DW - Sector 6", "DW-S06"

    ' ---------------------------------------------------------------
    ' 5. Unique index on Code (bloc B)
    ' ---------------------------------------------------------------
    If Not IndexExists(db, "T_STRUCTURES", "idx_Code_Unique") Then
        db.Execute "CREATE UNIQUE INDEX idx_Code_Unique ON T_STRUCTURES (Code)", dbFailOnError
        Debug.Print "-> unique index on Code created"
    Else
        Debug.Print "-> unique index already present"
    End If

    ' ---------------------------------------------------------------
    ' 6. Lookup updates BY NAME (blocs 1, F)
    ' ---------------------------------------------------------------
    db.Execute "UPDATE L_ELEMENTS SET Name_EN='Transverse wall or pier', Description='Transverse wall or pier anchoring the fabric to the rock; from a long anchoring wall to a compact pier - length does not change the letter. May anchor at any height of the fabric. Does not enclose any interior and does not count as a body. Against E: the corbel HANGS (embedded piece projecting), the pier RISES (fabric bearing downward). Against K: the pilaster lives IN the facade wall plane, the pier against the rock, outside any wall.' WHERE Code='D'", dbFailOnError
    Debug.Print "-> L_ELEMENTS: D is now Transverse wall or pier"

    db.Execute "UPDATE L_TYPOLOGY SET Description='Record whose evidence reduces to one or a few A-Z elements without a classifiable structure (bracket, transverse wall or pier, pilaster). Interpretation (circulation, earlier structure, support) goes to T_ARCH_FEATURES / Notes, never to the typology. The historical name comes from the first documented case, the bracket.' WHERE Name='MEN Isolated Structural Element'", dbFailOnError
    Debug.Print "-> L_TYPOLOGY: MEN examples updated"

    db.Execute "UPDATE L_DEC_TYPE SET Name='Outline band, defined', Name_VAL='Banda de contorn definida', Description='Pigment band following an outline, with a clean stroke and clear edges. The structural verdict lives in the record typology and attestation, never here. Topology (U, C, O) derives from the four span fields: QRY_27.' WHERE Name='RA U-shape geometric'", dbFailOnError
    db.Execute "UPDATE L_DEC_TYPE SET Name='Outline band, amorphous', Name_VAL='Banda de contorn amorfa', Description='Pigment band following an outline, with a diffuse stroke or ill-defined edges. The structural verdict lives in the record typology and attestation, never here. Topology (U, C, O) derives from the four span fields: QRY_27.' WHERE Name='RA U-shape organic'", dbFailOnError
    Debug.Print "-> L_DEC_TYPE: outline bands renamed, descriptions decoupled from the verdict"

    ' ---------------------------------------------------------------
    ' 7. Data cleanups (blocs G, H)
    ' ---------------------------------------------------------------
    Dim nPh As Long
    Dim nIo As Long
    nPh = DCount("*", "T_STRUCTURES", "Construction_Phases=1 AND Phase_Evidence Is Not Null")
    db.Execute "UPDATE T_STRUCTURES SET Phase_Evidence=Null WHERE Construction_Phases=1", dbFailOnError
    nIo = DCount("*", "T_STRUCTURES", "Interior_Observability Is Not Null AND ID_Typology IN (SELECT ID FROM L_TYPOLOGY WHERE Record_Class='Rock art panel')")
    db.Execute "UPDATE T_STRUCTURES SET Interior_Observability=Null WHERE ID_Typology IN (SELECT ID FROM L_TYPOLOGY WHERE Record_Class='Rock art panel')", dbFailOnError
    Debug.Print "-> cleanups: " & nPh & " phase-evidence and " & nIo & " interior-observability value(s) back to NULL"

    Set db = Nothing

    Dim msg As String
    msg = "PATCH v20 APPLIED" & vbCrLf & vbCrLf
    msg = msg & "  Columns dropped: " & nDrop & " of 2" & vbCrLf
    msg = msg & "  Columns added: " & nAdd & " of 2" & vbCrLf
    msg = msg & "  Cleanups: " & nPh & " + " & nIo & " values to NULL" & vbCrLf & vbCrLf
    msg = msg & "NEXT, IN THIS ORDER:" & vbCrLf & vbCrLf
    msg = msg & "1. chachapoya_DB_v20.bas -> RebuildQueriesV20()" & vbCrLf
    msg = msg & "   (NEVER BuildDB() on this database)" & vbCrLf & vbCrLf
    msg = msg & "2. chachapoya_Form_v20_val.bas -> BuildForm()" & vbCrLf & vbCrLf
    msg = msg & "3. Open QRY_25_V20_Review: the worklist of" & vbCrLf
    msg = msg & "   this migration." & vbCrLf
    MsgBox msg, vbInformation, "Patch v20"
    Debug.Print "=== PATCH FINISHED ==="
End Sub


Private Function AddCol(db As DAO.Database, tbl As String, col As String, typ As String) As Long
    If ColExists(db, tbl, col) Then
        Debug.Print "-> " & col & " already present"
        Exit Function
    End If
    db.Execute "ALTER TABLE " & tbl & " ADD COLUMN " & col & " " & typ, dbFailOnError
    db.TableDefs.Refresh
    AddCol = 1
    Debug.Print "-> " & col & " added to " & tbl
End Function

Private Function DropCol(db As DAO.Database, tbl As String, col As String) As Long
    If Not ColExists(db, tbl, col) Then
        Debug.Print "-> " & col & " already gone"
        Exit Function
    End If
    db.Execute "ALTER TABLE " & tbl & " DROP COLUMN " & col, dbFailOnError
    db.TableDefs.Refresh
    DropCol = 1
    Debug.Print "-> " & col & " dropped from " & tbl
End Function

Private Sub SetSector(db As DAO.Database, sName As String, sCode As String)
    db.Execute "UPDATE L_SECTORS SET Sector_Code='" & sCode & "' WHERE Sector_Name='" & sName & "'", dbFailOnError
End Sub

Private Function TableExists(db As DAO.Database, n As String) As Boolean
    Dim t As DAO.TableDef
    For Each t In db.TableDefs
        If t.Name = n Then TableExists = True: Exit Function
    Next t
End Function

Private Function ColExists(db As DAO.Database, tbl As String, col As String) As Boolean
    Dim f As DAO.Field
    On Error Resume Next
    For Each f In db.TableDefs(tbl).Fields
        If f.Name = col Then ColExists = True: Exit Function
    Next f
    On Error GoTo 0
End Function

Private Function IndexExists(db As DAO.Database, tbl As String, idx As String) As Boolean
    Dim i As DAO.Index
    On Error Resume Next
    For Each i In db.TableDefs(tbl).Indexes
        If i.Name = idx Then IndexExists = True: Exit Function
    Next i
    On Error GoTo 0
End Function
