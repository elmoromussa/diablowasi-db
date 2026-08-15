Option Compare Database
Option Explicit

' ================================================================
'  CHACHAPOYA PATCH v21 -> v22 (in place, data preserved)
'  Author: Esteve Ribera Torro | TFM Arqueologia UA
'  Spec: DELTA_v21_v22.md
'
'  WHAT THIS DOES, in order:
'    0. Preflight: v21 marker (Masonry_Present).
'    1. ADD Metrics_Available (bloc B): the metric-availability
'       gatekeeper on 9.Metr. 0/1/9, NULL default. OUT of the
'       rule-17 arrays on purpose: availability is not decidable
'       record-by-record until the extraction workflow runs.
'    2. DERIVABLE FILL: every record already carrying ANY
'       dimensional or volumetric value gets Metrics_Available=1
'       (2 records expected on the 2026-08-15 corpus). The rest
'       stay NULL - and deliberately do NOT go to any worklist.
'    3. PANEL MASONRY (bloc A): the four rock art panels get the
'       DERIVABLE Masonry_Present = 0 - a pure panel has no
'       fabric by definition (if it had, it would be
'       reclassified). Only NULLs are touched: a declared value
'       is never overwritten. This closes 4 of the 27 masonry
'       judgements of QRY_29; the naturals become answerable on
'       the form itself (2.Arq opens for them in the v22 form).
'    4. LEGACY QUERY CLEANUP (bloc E): the v21 rename orphaned
'       QRY_16g_Rules_57_60. Deleted here too, so the patch
'       leaves a clean house even before RebuildQueriesV22.
'
'  AFTER RUNNING THIS, IN ORDER:
'    1. chachapoya_DB_v22.bas -> RebuildQueriesV22()
'       (NEVER BuildDB() on this database)
'    2. chachapoya_Form_v22_val.bas -> BuildForm()
'    3. chachapoya_Worklist_v22.bas -> BuildWorklist()
'    4. The masonry judgements of the naturals live on in
'       QRY_29 (font 'Migracio v21') - now answerable from
'       2.Arq itself.
'
'  Every step is guarded, so re-running the patch is harmless.
' ================================================================

Sub PatchV22()
    Dim db As DAO.Database
    Set db = CurrentDb()

    Debug.Print "=== PATCH v21 -> v22 ==="

    ' Preflight: schema generation.
    If Not TableExists(db, "T_STRUCTURES") Then
        MsgBox "T_STRUCTURES does not exist. This is not a Chachapoya database.", vbCritical, "Patch aborted"
        Exit Sub
    End If
    If Not ColExists(db, "T_STRUCTURES", "Masonry_Present") Then
        MsgBox "Masonry_Present is missing. This database is older than v21 - apply PATCH v21 first.", vbCritical, "Patch aborted"
        Exit Sub
    End If

    Dim nAdd As Long

    ' ---------------------------------------------------------------
    ' 1. Metrics_Available (bloc B)
    ' ---------------------------------------------------------------
    nAdd = nAdd + AddCol(db, "T_STRUCTURES", "Metrics_Available", "BYTE")

    ' ---------------------------------------------------------------
    ' 2. Derivable fill: a value implies availability (bloc B)
    ' ---------------------------------------------------------------
    Dim nOne As Long
    Dim sql As String
    sql = "UPDATE T_STRUCTURES SET Metrics_Available=1 "
    sql = sql & "WHERE Metrics_Available Is Null "
    sql = sql & "AND (Length_m Is Not Null Or Width_m Is Not Null "
    sql = sql & "Or Height_m Is Not Null Or Height_Above_Base_m Is Not Null "
    sql = sql & "Or Dim_Method Is Not Null Or Opening_Width_m Is Not Null "
    sql = sql & "Or Opening_Height_m Is Not Null Or Support_Width_m Is Not Null "
    sql = sql & "Or Support_Depth_m Is Not Null Or Area_m2 Is Not Null "
    sql = sql & "Or Volume_m3 Is Not Null Or ID_Vol_Method Is Not Null "
    sql = sql & "Or Vol_Notes Is Not Null)"
    db.Execute sql, dbFailOnError
    nOne = db.RecordsAffected
    Debug.Print "-> Metrics_Available = 1 derived on " & nOne & " record(s) carrying metric data"
    Debug.Print "-> the NULLs stay NULL and go to NO worklist: the availability"
    Debug.Print "   judgement waits for the extraction workflow (delta, bloc B)"

    ' ---------------------------------------------------------------
    ' 3. Panel masonry: the derivable 0 (bloc A)
    ' ---------------------------------------------------------------
    Dim nPan As Long
    sql = "UPDATE T_STRUCTURES SET Masonry_Present=0 "
    sql = sql & "WHERE Masonry_Present Is Null "
    sql = sql & "AND ID_Typology IN (SELECT ID FROM L_TYPOLOGY WHERE Record_Class='Rock art panel')"
    db.Execute sql, dbFailOnError
    nPan = db.RecordsAffected
    Debug.Print "-> Masonry_Present = 0 derived on " & nPan & " rock art panel(s)"

    ' ---------------------------------------------------------------
    ' 4. Legacy query cleanup (bloc E)
    ' ---------------------------------------------------------------
    Dim nLeg As Long
    nLeg = nLeg + DropQuery(db, "QRY_16g_Rules_57_60")
    Debug.Print "-> legacy queries removed: " & nLeg

    Set db = Nothing

    Dim msg As String
    msg = "PATCH v22 APPLIED" & vbCrLf & vbCrLf
    msg = msg & "  Columns added: " & nAdd & " of 1" & vbCrLf
    msg = msg & "  Metrics_Available derived to 1: " & nOne & vbCrLf
    msg = msg & "  Panel Masonry_Present derived to 0: " & nPan & vbCrLf
    msg = msg & "  Legacy queries removed: " & nLeg & vbCrLf & vbCrLf
    msg = msg & "NEXT, IN THIS ORDER:" & vbCrLf & vbCrLf
    msg = msg & "1. chachapoya_DB_v22.bas -> RebuildQueriesV22()" & vbCrLf
    msg = msg & "   (NEVER BuildDB() on this database)" & vbCrLf & vbCrLf
    msg = msg & "2. chachapoya_Form_v22_val.bas -> BuildForm()" & vbCrLf & vbCrLf
    msg = msg & "3. chachapoya_Worklist_v22.bas -> BuildWorklist()" & vbCrLf & vbCrLf
    msg = msg & "4. Els judicis de maconeria dels naturals" & vbCrLf
    msg = msg & "   continuen a QRY_29 - ara responibles des" & vbCrLf
    msg = msg & "   de 2.Arq mateix." & vbCrLf
    MsgBox msg, vbInformation, "Patch v22"
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

Private Function DropQuery(db As DAO.Database, n As String) As Long
    Dim q As DAO.QueryDef
    For Each q In db.QueryDefs
        If q.Name = n Then
            db.QueryDefs.Delete n
            DropQuery = 1
            Exit Function
        End If
    Next q
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
