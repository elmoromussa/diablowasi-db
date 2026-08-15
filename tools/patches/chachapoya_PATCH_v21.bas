Option Compare Database
Option Explicit

' ================================================================
'  CHACHAPOYA PATCH v20 -> v21 (in place, data preserved)
'  Author: Esteve Ribera Torro | TFM Arqueologia UA
'  Spec: DELTA_v20_v21.md
'
'  WHAT THIS DOES, in order:
'    0. Preflight: v20 marker (Upper_Crown_Format). No duplicate
'       check this time: the v20 unique index already makes a
'       duplicate code impossible to enter.
'    1. ADD Masonry_Present (bloc 2): the fabric declaration.
'       0/1/9, NULL default - a judgement field, three epistemic
'       states preserved.
'    2. DERIVABLE FILL: every record already carrying ANY masonry
'       detail (quality, format, working, bond, mortar, chinking,
'       fabric divergence) gets Masonry_Present = 1 - the detail
'       could not have been recorded on no fabric. Records with
'       no detail at all stay NULL: that judgement goes to the
'       QRY_29 worklist (the 0 is what the EA55 corbel has been
'       waiting for; the 9 is a legitimate answer).
'    3. REPORT ONLY, no UPDATE: the Z-only platforms the retired
'       terrace criterion produced (Sys_Platform Present* with
'       E, F and G all confirmed 0). Reverting is a row-by-row
'       judgement - a slab resting on piers would be caught by a
'       blind UPDATE - so the codes are printed here and listed
'       in QRY_29; rule 62 keeps watching afterwards.
'
'  NOTHING IS DROPPED in this patch. The terrace criterion
'  revision (bloc 1) changes queries, the form and the manual,
'  never the schema.
'
'  AFTER RUNNING THIS, IN ORDER:
'    1. chachapoya_DB_v21.bas -> RebuildQueriesV21()
'       (NEVER BuildDB() on this database)
'    2. chachapoya_Form_v21_val.bas -> BuildForm()
'    3. chachapoya_Worklist_v21.bas -> BuildWorklist()
'    4. Open QRY_29_V21_Review: the worklist of this migration.
'
'  Every step is guarded, so re-running the patch is harmless.
' ================================================================

Sub PatchV21()
    Dim db As DAO.Database
    Set db = CurrentDb()

    Debug.Print "=== PATCH v20 -> v21 ==="

    ' Preflight: schema generation.
    If Not TableExists(db, "T_STRUCTURES") Then
        MsgBox "T_STRUCTURES does not exist. This is not a Chachapoya database.", vbCritical, "Patch aborted"
        Exit Sub
    End If
    If Not ColExists(db, "T_STRUCTURES", "Upper_Crown_Format") Then
        MsgBox "Upper_Crown_Format is missing. This database is older than v20 - apply PATCH v20 first.", vbCritical, "Patch aborted"
        Exit Sub
    End If

    Dim nAdd As Long

    ' ---------------------------------------------------------------
    ' 1. Masonry_Present (bloc 2)
    ' ---------------------------------------------------------------
    nAdd = nAdd + AddCol(db, "T_STRUCTURES", "Masonry_Present", "BYTE")

    ' ---------------------------------------------------------------
    ' 2. Derivable fill: detail implies fabric (bloc 2)
    ' ---------------------------------------------------------------
    Dim nOne As Long
    Dim sql As String
    sql = "UPDATE T_STRUCTURES SET Masonry_Present=1 "
    sql = sql & "WHERE Masonry_Present Is Null "
    sql = sql & "AND (Masonry_Quality Is Not Null Or Masonry_Type Is Not Null "
    sql = sql & "Or Stone_Format Is Not Null Or Stone_Format_Secondary Is Not Null "
    sql = sql & "Or Stone_Working Is Not Null Or Mortar_Present Is Not Null "
    sql = sql & "Or Mortar_Type Is Not Null Or Chinking_Stones Is Not Null "
    sql = sql & "Or Fabric Is Not Null)"
    db.Execute sql, dbFailOnError
    nOne = db.RecordsAffected
    Debug.Print "-> Masonry_Present = 1 derived on " & nOne & " record(s) carrying masonry detail"

    Dim nNul As Long
    nNul = DCount("*", "T_STRUCTURES", "Masonry_Present Is Null")
    Debug.Print "-> " & nNul & " record(s) left NULL: the judgement goes to QRY_29"

    ' ---------------------------------------------------------------
    ' 3. Report the Z-only platforms (bloc 1) - NO update
    ' ---------------------------------------------------------------
    Dim rs As DAO.Recordset
    Dim zList As String
    Dim nZ As Long
    Set rs = db.OpenRecordset("SELECT Code FROM T_STRUCTURES WHERE Sys_Platform Like 'Present*' AND Timber_Brackets=0 AND Transverse_Beams=0 AND Corbelled_Courses=0 ORDER BY Code")
    Do While Not rs.EOF
        zList = zList & "  " & rs!Code & vbCrLf
        nZ = nZ + 1
        rs.MoveNext
    Loop
    rs.Close
    If nZ > 0 Then
        Debug.Print "-> " & nZ & " Z-only platform(s) to revert BY HAND (old terrace criterion):"
        Debug.Print zList
    Else
        Debug.Print "-> no Z-only platforms found: the terrace revision costs no data pass"
    End If

    Set db = Nothing

    Dim msg As String
    msg = "PATCH v21 APPLIED" & vbCrLf & vbCrLf
    msg = msg & "  Columns added: " & nAdd & " of 1" & vbCrLf
    msg = msg & "  Masonry_Present derived to 1: " & nOne & vbCrLf
    msg = msg & "  Masonry judgements pending: " & nNul & vbCrLf
    msg = msg & "  Z-only platforms to revert by hand: " & nZ & vbCrLf & vbCrLf
    msg = msg & "NEXT, IN THIS ORDER:" & vbCrLf & vbCrLf
    msg = msg & "1. chachapoya_DB_v21.bas -> RebuildQueriesV21()" & vbCrLf
    msg = msg & "   (NEVER BuildDB() on this database)" & vbCrLf & vbCrLf
    msg = msg & "2. chachapoya_Form_v21_val.bas -> BuildForm()" & vbCrLf & vbCrLf
    msg = msg & "3. chachapoya_Worklist_v21.bas -> BuildWorklist()" & vbCrLf & vbCrLf
    msg = msg & "4. Open QRY_29_V21_Review: the worklist of" & vbCrLf
    msg = msg & "   this migration." & vbCrLf
    MsgBox msg, vbInformation, "Patch v21"
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
