Option Compare Database
Option Explicit

' ================================================================
'  CHACHAPOYA PATCH v24 -> v25 (in place, data preserved)
'  Author: Esteve Ribera Torro | TFM Arqueologia UA
'  Spec: DELTA_v24_v25.md
'
'  WHAT THIS DOES:
'    0. Preflight: v24 marker (Facade_Azimuth_Deg).
'    1. ADD Jamb_Fabric (TEXT 25, NULL) to T_STRUCTURES.
'       Composite jamb fabric: Monolithic slab / Composite
'       slab-masonry / Fabric as jamb. Gated on the form by
'       Jambs (house pattern: NULL keeps it editable).
'    2. ADD Niche_Partition (TEXT 15, NULL) to T_STRUCTURES.
'       Horizontal slab splitting a natural niche in two
'       sub-niches. NIX only (form gating); CAV excluded by
'       decision (delta, bloc B).
'    3. ADD Term_Lit (TEXT 30) + Lit_Source (TEXT 40) to
'       L_ELEMENTS and fill the three confirmed concordances:
'       B=platform-base, I=cornice, M=frieze (Guengerich 2014).
'    4. CREATE T_METRIC_REVIEW: the reconciliation table of the
'       metrics<->DB coherence pipeline. Populated ONLY by
'       CheckMetrics (chachapoya_CheckMetrics_v25.bas); decisions
'       are signed by hand; ApplyMetricReview executes them.
'
'  AFTER RUNNING THIS, IN ORDER:
'    1. chachapoya_DB_v25.bas -> RebuildQueriesV25()
'       (NEVER BuildDB() on this database)
'    2. chachapoya_Form_v25_val.bas -> BuildForm()
'       (with F_STRUCTURES CLOSED first)
'    3. chachapoya_Worklist_v25.bas -> BuildWorklist()
'    4. chachapoya_CheckMetrics_v25.bas -> CheckMetrics
'       with the CURRENT clean metric CSVs (the calibration
'       certificate: zero invalid, zero legacy, bodies squared).
'    5. Sign decisions in T_METRIC_REVIEW (or via the worklist),
'       then ApplyMetricReview.
'
'  Every step is guarded, so re-running the patch is harmless.
' ================================================================

Sub PatchV25()
    Dim db As DAO.Database
    Set db = CurrentDb()

    Debug.Print "=== PATCH v24 -> v25 ==="

    If Not TableExists(db, "T_STRUCTURES") Then
        MsgBox "T_STRUCTURES does not exist. This is not a Chachapoya database.", vbCritical, "Patch aborted"
        Exit Sub
    End If
    If Not ColExists(db, "T_STRUCTURES", "Facade_Azimuth_Deg") Then
        MsgBox "Facade_Azimuth_Deg is missing. This database is older than v24 - apply PATCH v24 first.", vbCritical, "Patch aborted"
        Exit Sub
    End If
    If Not TableExists(db, "L_ELEMENTS") Then
        MsgBox "L_ELEMENTS is missing. This database is older than v11.", vbCritical, "Patch aborted"
        Exit Sub
    End If

    Dim nAdd As Long
    nAdd = nAdd + AddCol(db, "T_STRUCTURES", "Jamb_Fabric", "TEXT(25)")
    nAdd = nAdd + AddCol(db, "T_STRUCTURES", "Niche_Partition", "TEXT(15)")
    nAdd = nAdd + AddCol(db, "L_ELEMENTS", "Term_Lit", "TEXT(30)")
    nAdd = nAdd + AddCol(db, "L_ELEMENTS", "Lit_Source", "TEXT(40)")

    ' The three confirmed concordances. Guarded: only fills NULL,
    ' so a hand-refined value survives a re-run.
    Dim nCon As Long
    nCon = nCon + SetLit(db, "B", "platform-base", "Guengerich 2014")
    nCon = nCon + SetLit(db, "I", "cornice", "Guengerich 2014")
    nCon = nCon + SetLit(db, "M", "frieze", "Guengerich 2014")

    ' The review table. FK inline: JET accepts the constraint in
    ' the CREATE and the relation shows up in the relationships
    ' window without touching the MkRel machinery.
    Dim nTbl As Long
    If Not TableExists(db, "T_METRIC_REVIEW") Then
        Dim sql As String
        sql = "CREATE TABLE T_METRIC_REVIEW ("
        sql = sql & "ID COUNTER CONSTRAINT PK_MREV PRIMARY KEY,"
        sql = sql & "ID_Structure LONG NOT NULL,"
        sql = sql & "Check_Code TEXT(4) NOT NULL,"
        sql = sql & "Field_Name TEXT(40),"
        sql = sql & "DB_Value TEXT(60),"
        sql = sql & "Proposed_Value TEXT(60),"
        sql = sql & "Evidence MEMO,"
        sql = sql & "CSV_Batch TEXT(80),"
        sql = sql & "Detected_On DATETIME,"
        sql = sql & "Decision TEXT(12),"
        sql = sql & "Decided_On DATETIME,"
        sql = sql & "Applied_On DATETIME,"
        sql = sql & "Notes MEMO,"
        sql = sql & "CONSTRAINT FK_MREV_STRUCT FOREIGN KEY (ID_Structure) REFERENCES T_STRUCTURES (ID))"
        db.Execute sql, dbFailOnError
        db.TableDefs.Refresh
        nTbl = 1
        Debug.Print "-> T_METRIC_REVIEW created"
        SetCaps db
    Else
        Debug.Print "-> T_METRIC_REVIEW already present"
    End If

    Set db = Nothing

    Dim msg As String
    msg = "PATCH v25 APPLIED" & vbCrLf & vbCrLf
    msg = msg & "  Columns added: " & nAdd & vbCrLf
    msg = msg & "  Concordances filled: " & nCon & vbCrLf
    msg = msg & "  Tables created: " & nTbl & vbCrLf & vbCrLf
    msg = msg & "NEXT, IN THIS ORDER:" & vbCrLf & vbCrLf
    msg = msg & "1. chachapoya_DB_v25.bas -> RebuildQueriesV25()" & vbCrLf
    msg = msg & "   (NEVER BuildDB() on this database)" & vbCrLf & vbCrLf
    msg = msg & "2. chachapoya_Form_v25_val.bas -> BuildForm()" & vbCrLf
    msg = msg & "   (amb F_STRUCTURES TANCAT abans)" & vbCrLf & vbCrLf
    msg = msg & "3. chachapoya_Worklist_v25.bas -> BuildWorklist()" & vbCrLf & vbCrLf
    msg = msg & "4. CheckMetrics amb els CSV metrics NETS vigents" & vbCrLf
    msg = msg & "   (el certificat de calibratge)." & vbCrLf & vbCrLf
    msg = msg & "5. Firmar decisions a T_METRIC_REVIEW i executar" & vbCrLf
    msg = msg & "   ApplyMetricReview." & vbCrLf
    MsgBox msg, vbInformation, "Patch v25"
    Debug.Print "=== PATCH FINISHED ==="
End Sub


Private Function SetLit(db As DAO.Database, elemCode As String, term As String, src As String) As Long
    Dim sql As String
    sql = "UPDATE L_ELEMENTS SET Term_Lit='" & term & "', Lit_Source='" & src & "' "
    sql = sql & "WHERE Code='" & elemCode & "' AND Term_Lit IS NULL"
    db.Execute sql, dbFailOnError
    SetLit = db.RecordsAffected
    If SetLit > 0 Then Debug.Print "-> L_ELEMENTS " & elemCode & ": Term_Lit='" & term & "' (" & src & ")"
End Function

Private Sub SetCaps(db As DAO.Database)
    ' Datasheet captions via DAO, the house pattern.
    Cap db, "T_METRIC_REVIEW", "ID_Structure", "Estructura (ID)"
    Cap db, "T_METRIC_REVIEW", "Check_Code", "Creuament"
    Cap db, "T_METRIC_REVIEW", "Field_Name", "Camp"
    Cap db, "T_METRIC_REVIEW", "DB_Value", "Valor BD"
    Cap db, "T_METRIC_REVIEW", "Proposed_Value", "Proposta (CSV)"
    Cap db, "T_METRIC_REVIEW", "Evidence", "Evidencia"
    Cap db, "T_METRIC_REVIEW", "CSV_Batch", "Lot CSV"
    Cap db, "T_METRIC_REVIEW", "Detected_On", "Detectat"
    Cap db, "T_METRIC_REVIEW", "Decision", "Decisio"
    Cap db, "T_METRIC_REVIEW", "Decided_On", "Decidit"
    Cap db, "T_METRIC_REVIEW", "Applied_On", "Aplicat"
    Cap db, "T_METRIC_REVIEW", "Notes", "Notes"
End Sub

Private Sub Cap(db As DAO.Database, tbl As String, fld As String, capText As String)
    Dim f As DAO.Field
    Dim p As DAO.Property
    On Error Resume Next
    Set f = db.TableDefs(tbl).Fields(fld)
    f.Properties("Caption") = capText
    If Err.Number <> 0 Then
        Err.Clear
        Set p = f.CreateProperty("Caption", dbText, capText)
        f.Properties.Append p
    End If
    On Error GoTo 0
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
End Function
