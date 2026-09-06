Option Compare Database
Option Explicit

' ================================================================
'  CHACHAPOYA MIGRATOR v25h(patched) -> v26
'  Run Sub MigrateFromV25h() INSIDE the freshly built v26
'  database (BuildDB() from chachapoya_DB_v26.bas already run,
'  lookups populated, data tables EMPTY). Edit SRC_PATH first.
'
'  DESIGN
'  1. Lookup alignment BY NAME: for every L_ table, the pairs
'     (ID, Name) of source and target must match one to one.
'     This is the foreign-keys-by-name guard operationalised:
'     if any lookup drifted, the migration refuses to start.
'  2. Field lists are built at runtime FROM THE TARGET schema,
'     so the retired Jamb_Fabric_Reveal column simply does not
'     travel - nothing is hard-coded, nothing else is dropped.
'  3. IDs are carried explicitly (autonumber insert with the
'     column listed), so every FK survives untouched.
'  4. Import order respects dependencies; verification recounts
'     every table against the source and spot-checks the five
'     v26 patch decisions on arrival.
'
'  House rules: ASCII, CRLF, no line continuations, DAO,
'  Private helpers, Val() at text-to-number frontiers.
' ================================================================

' The source path is PICKED at run time through the standard
' file dialog (same pattern as ImportGeoref) - nothing to edit.
Private SRC_PATH As String

Private fails As String

Public Sub MigrateFromV25h()
    Dim db As DAO.Database
    Set db = CurrentDb()
    fails = ""

    ' ---------- PHASE -1: PICK THE SOURCE FILE ----------
    SRC_PATH = PickSourceFile()
    If Len(SRC_PATH) = 0 Then
        MsgBox "No source file selected - nothing done.", vbInformation, "MigrateFromV25h"
        Set db = Nothing
        Exit Sub
    End If
    If SRC_PATH = db.Name Then
        MsgBox "The selected file IS the current database - pick the PATCHED v25h copy instead.", vbCritical, "MigrateFromV25h"
        Set db = Nothing
        Exit Sub
    End If
    Debug.Print "Source: " & SRC_PATH

    ' ---------- PHASE 0: IS THIS THE BUILT v26? ----------
    If Not TableExists(db, "T_STRUCTURES") Or Not TableExists(db, "L_SITES") Then
        MsgBox "This database has no v26 schema yet." & vbCrLf & "Run BuildDB() from chachapoya_DB_v26.bas FIRST, on this same (blank) database, then run the migrator again.", vbCritical, "MigrateFromV25h"
        Set db = Nothing
        Exit Sub
    End If

    ' ---------- PHASE 1: DATA TABLES EMPTY? ----------
    ' Before any lookup sync (which deletes lookup rows), the
    ' target data tables must hold nothing that points at them.
    Dim dt As Variant
    dt = Array("T_GROUPS", "T_STRUCTURES", "T_DECORATIONS", "T_ARCH_FEATURES", "T_CONNECTIONS", "T_DATING", "T_INDIVIDUALS", "T_LOST_ELEMENTS", "T_METRIC_REVIEW")
    Dim i As Integer
    For i = LBound(dt) To UBound(dt)
        If TblCount(db, CStr(dt(i))) > 0 Then
            fails = fails & "  " & dt(i) & " is not empty" & vbCrLf
        End If
    Next i
    If Len(fails) > 0 Then
        MsgBox "MIGRATION ABORTED - target not blank:" & vbCrLf & fails, vbCritical, "MigrateFromV25h"
        Set db = Nothing
        Exit Sub
    End If

    ' ---------- PHASE 2: LOOKUP ALIGNMENT, WITH SYNC ----------
    ' DOCTRINE (v26): live lookups are DATA with a history - the
    ' build script bootstraps a fresh vocabulary, but on migration
    ' the SOURCE is the authority. Autonumber holes (L_DEC_TYPE:
    ' 13, 14, 16 gone) are scars of the patch series and are NEVER
    ' renumbered: the data rows point at the surviving ids, and a
    ' sequential rebuild would silently remap them. A misaligned
    ' lookup is therefore REPOPULATED from the source (explicit
    ' ids), re-checked, and reported.
    Dim lk As Variant
    lk = Array("L_SITES;Site_Name", "L_SECTORS;Sector_Name", "L_SUBSECTORS;Name", "L_TYPOLOGY;Name", "L_SUPPORT;Name", "L_STATUS;Name", "L_MATERIAL_STATUS;Name", "L_VOL_METHOD;Name", "L_COORD_METHOD;Name", "L_GROUP_TYPE;Name", "L_CAMPAIGN;Code", "L_STRUCT_BODY;Code", "L_DEC_TYPE;Name", "L_ELEMENTS;Code", "L_LOST_EVIDENCE;Name")
    Dim parts() As String
    Dim synced As String
    synced = ""
    For i = LBound(lk) To UBound(lk)
        parts = Split(CStr(lk(i)), ";")
        If Not LookupAligned(db, parts(0), parts(1)) Then
            SyncLookup db, parts(0)
            If LookupAligned(db, parts(0), parts(1)) Then
                synced = synced & "  " & parts(0) & " synchronised from source" & vbCrLf
            Else
                fails = fails & "  " & parts(0) & ": still misaligned after sync" & vbCrLf
            End If
        End If
    Next i
    If Len(fails) > 0 Then
        MsgBox "MIGRATION ABORTED - lookups not aligned, nothing imported." & vbCrLf & vbCrLf & fails, vbCritical, "MigrateFromV25h"
        Set db = Nothing
        Exit Sub
    End If
    Debug.Print "MigrateFromV25h: all lookups aligned." & IIf(Len(synced) > 0, vbCrLf & synced, "")

    ' ---------- PHASE 3: IMPORT (dependency order) ----------
    For i = LBound(dt) To UBound(dt)
        ImportTable db, CStr(dt(i))
    Next i

    ' ---------- PHASE 4: VERIFICATION ----------
    Dim msg As String
    msg = "MIGRATION REPORT" & vbCrLf
    If Len(synced) > 0 Then
        msg = msg & "Lookups synchronised from source (id holes preserved):" & vbCrLf & synced & vbCrLf
    End If
    For i = LBound(dt) To UBound(dt)
        Dim nT As Long
        Dim nS As Long
        nT = TblCount(db, CStr(dt(i)))
        nS = SrcCount(db, CStr(dt(i)))
        msg = msg & "  " & dt(i) & ": " & nT & " / " & nS & IIf(nT = nS, " OK", " MISMATCH") & vbCrLf
    Next i

    ' Spot checks: the five v26 decisions must have arrived
    msg = msg & vbCrLf & "Spot checks (v26 decisions):" & vbCrLf
    msg = msg & Spot(db, "EA20 base NA", "SELECT Count(*) AS N FROM T_STRUCTURES WHERE Code='DW-S04-EA20' AND Sys_Base='Not applicable' AND N_Chamber_Bodies=0")
    msg = msg & Spot(db, "EA29 swapped", "SELECT Count(*) AS N FROM T_STRUCTURES WHERE Code='DW-S01-EA29' AND N_Basal_Bodies=1 AND N_Chamber_Bodies=0")
    msg = msg & Spot(db, "EA35 NULLs", "SELECT Count(*) AS N FROM T_STRUCTURES WHERE Code='DW-S01-EA35' AND N_Basal_Bodies Is Null AND N_Chamber_Bodies Is Null")
    msg = msg & Spot(db, "EA11 typology", "SELECT Count(*) AS N FROM T_STRUCTURES AS E INNER JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID WHERE E.Code='DW-S01-EA11' AND T.Name='Unclassifiable'")
    msg = msg & Spot(db, "11 counters", "SELECT Count(*) AS N FROM T_STRUCTURES WHERE Sys_Chamber='Absent' AND N_Chamber_Bodies Is Null")
    msg = msg & "  (11 counters check expects 0 remaining NULLs)" & vbCrLf
    msg = msg & vbCrLf & "Next: run the QRY_16_Validation_Check battery - expected at zero."
    Debug.Print msg
    MsgBox msg, vbInformation, "MigrateFromV25h"
    Set db = Nothing
End Sub

' ---------------- Private helpers ----------------

Private Function PickSourceFile() As String
    ' msoFileDialogFilePicker = 3; late-bound, no extra reference.
    Dim fd As Object
    Set fd = Application.FileDialog(3)
    fd.Title = "Tria la copia APANYADA de DB_v25h (la del PatchV26 en verd)"
    fd.AllowMultiSelect = False
    fd.Filters.Clear
    fd.Filters.Add "Access", "*.accdb"
    If fd.Show Then
        PickSourceFile = fd.SelectedItems(1)
    Else
        PickSourceFile = ""
    End If
End Function

Private Function TableExists(db As DAO.Database, tbl As String) As Boolean
    Dim td As DAO.TableDef
    TableExists = False
    For Each td In db.TableDefs
        If td.Name = tbl Then
            TableExists = True
            Exit Function
        End If
    Next td
End Function

Private Function TblCount(db As DAO.Database, tbl As String) As Long
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT Count(*) AS N FROM " & tbl, dbOpenSnapshot)
    TblCount = Val(rs!N & "")
    rs.Close
End Function

Private Function SrcCount(db As DAO.Database, tbl As String) As Long
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT Count(*) AS N FROM [" & tbl & "] IN '" & SRC_PATH & "'", dbOpenSnapshot)
    SrcCount = Val(rs!N & "")
    rs.Close
End Function

Private Function LookupAligned(db As DAO.Database, tbl As String, keyCol As String) As Boolean
    Dim nT As Long
    Dim nS As Long
    Dim nM As Long
    Dim rs As DAO.Recordset
    nT = TblCount(db, tbl)
    nS = SrcCount(db, tbl)
    ' JET: a mixed local/external join cannot use the IN clause
    ' (IN rebinds the WHOLE from-list); the per-table external
    ' form [path].[table] is the correct syntax here.
    Dim q As String
    q = "SELECT Count(*) AS N FROM " & tbl & " AS T "
    q = q & "INNER JOIN [" & SRC_PATH & "].[" & tbl & "] AS S "
    q = q & "ON T.ID=S.ID AND T.[" & keyCol & "]=S.[" & keyCol & "]"
    Set rs = db.OpenRecordset(q, dbOpenSnapshot)
    nM = Val(rs!N & "")
    rs.Close
    LookupAligned = (nT = nS And nM = nT)
    If Not LookupAligned Then
        Debug.Print "  " & tbl & ": target " & nT & ", source " & nS & ", matched " & nM
    End If
End Function

Private Sub SyncLookup(db As DAO.Database, tbl As String)
    ' Repopulate the target lookup from the source, EXPLICIT ids
    ' included (autonumber accepts them on INSERT...SELECT). Safe
    ' here by construction: phase 1 proved the data tables empty.
    Dim td As DAO.TableDef
    Dim fld As DAO.Field
    Dim lst As String
    Set td = db.TableDefs(tbl)
    lst = ""
    For Each fld In td.Fields
        If Len(lst) > 0 Then
            lst = lst & ", "
        End If
        lst = lst & "[" & fld.Name & "]"
    Next fld
    db.Execute "DELETE FROM " & tbl, dbFailOnError
    Dim q As String
    q = "INSERT INTO " & tbl & " (" & lst & ") "
    q = q & "SELECT " & lst & " FROM [" & tbl & "] IN '" & SRC_PATH & "'"
    db.Execute q, dbFailOnError
    Debug.Print "  -> " & tbl & " repopulated from source (" & db.RecordsAffected & " rows)"
End Sub

Private Sub ImportTable(db As DAO.Database, tbl As String)
    Dim td As DAO.TableDef
    Dim fld As DAO.Field
    Dim lst As String
    Set td = db.TableDefs(tbl)
    lst = ""
    For Each fld In td.Fields
        If Len(lst) > 0 Then
            lst = lst & ", "
        End If
        lst = lst & "[" & fld.Name & "]"
    Next fld
    Dim q As String
    q = "INSERT INTO " & tbl & " (" & lst & ") "
    q = q & "SELECT " & lst & " FROM [" & tbl & "] IN '" & SRC_PATH & "'"
    db.Execute q, dbFailOnError
    Debug.Print "  -> " & tbl & ": " & db.RecordsAffected & " rows imported"
End Sub

Private Function Spot(db As DAO.Database, label As String, sql As String) As String
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    Dim n As Long
    n = Val(rs!N & "")
    rs.Close
    Spot = "  " & label & ": " & n & vbCrLf
End Function
