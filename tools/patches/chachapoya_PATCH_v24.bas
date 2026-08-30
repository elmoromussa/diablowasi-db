Option Compare Database
Option Explicit

' ================================================================
'  CHACHAPOYA PATCH v23 -> v24 (in place, data preserved)
'  Author: Esteve Ribera Torro | TFM Arqueologia UA
'  Spec: DELTA_v23_v24.md
'
'  WHAT THIS DOES:
'    0. Preflight: v23 markers (EA_Number, ID_Subsector).
'    1. ADD Facade_Azimuth_Deg (INTEGER, NULL) to T_STRUCTURES.
'       Values stay NULL: the field is INSTRUMENTAL and is
'       written only by the georef importer
'       (chachapoya_Import_Georef_v24.bas), never by hand.
'       Rule 68 (rebuilt next) watches its concordance with
'       the observational Facade_Orientation.
'
'  AFTER RUNNING THIS, IN ORDER:
'    1. chachapoya_DB_v24.bas -> RebuildQueriesV24()
'       (NEVER BuildDB() on this database)
'    2. chachapoya_Form_v24_val.bas -> BuildForm()
'       (with F_STRUCTURES CLOSED first)
'    3. chachapoya_Worklist_v24.bas -> BuildWorklist()
'       (unchanged in v24; rerun only to be safe)
'    4. chachapoya_Import_Georef_v24.bas -> ImportGeoref
'       with the CSV from export_georef_orientacions.py -
'       REGENERATED after the clean metrics pass, because the
'       bbox is the shared instrument and only a corpus with
'       zero PREFIX_DUBTOS certifies its calibration.
'
'  Every step is guarded, so re-running the patch is harmless.
' ================================================================

Sub PatchV24()
    Dim db As DAO.Database
    Set db = CurrentDb()

    Debug.Print "=== PATCH v23 -> v24 ==="

    If Not TableExists(db, "T_STRUCTURES") Then
        MsgBox "T_STRUCTURES does not exist. This is not a Chachapoya database.", vbCritical, "Patch aborted"
        Exit Sub
    End If
    If Not ColExists(db, "T_STRUCTURES", "EA_Number") Then
        MsgBox "EA_Number is missing. This database is older than v23 - apply PATCH v23 first.", vbCritical, "Patch aborted"
        Exit Sub
    End If
    If Not ColExists(db, "T_STRUCTURES", "ID_Subsector") Then
        MsgBox "ID_Subsector is missing. This database is older than v23 - apply PATCH v23 first.", vbCritical, "Patch aborted"
        Exit Sub
    End If

    Dim nAdd As Long
    nAdd = AddCol(db, "T_STRUCTURES", "Facade_Azimuth_Deg", "INTEGER")

    Set db = Nothing

    Dim msg As String
    msg = "PATCH v24 APPLIED" & vbCrLf & vbCrLf
    msg = msg & "  Columns added: " & nAdd & vbCrLf & vbCrLf
    msg = msg & "NEXT, IN THIS ORDER:" & vbCrLf & vbCrLf
    msg = msg & "1. chachapoya_DB_v24.bas -> RebuildQueriesV24()" & vbCrLf
    msg = msg & "   (NEVER BuildDB() on this database)" & vbCrLf & vbCrLf
    msg = msg & "2. chachapoya_Form_v24_val.bas -> BuildForm()" & vbCrLf
    msg = msg & "   (amb F_STRUCTURES TANCAT abans)" & vbCrLf & vbCrLf
    msg = msg & "3. chachapoya_Worklist_v24.bas -> BuildWorklist()" & vbCrLf & vbCrLf
    msg = msg & "4. Import georef: ImportGeoref amb el CSV" & vbCrLf
    msg = msg & "   REGENERAT despres de la passada neta de" & vbCrLf
    msg = msg & "   l'extractor de metriques (bbox certificats)." & vbCrLf
    MsgBox msg, vbInformation, "Patch v24"
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
