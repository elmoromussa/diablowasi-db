Option Compare Database
Option Explicit

' ================================================================
'  CHACHAPOYA PATCH v25h -> v26 (DATA ONLY)
'  Decisions: DELTA_v25g_v26.md (sessio 2026-09-06, tot firmat)
'  Run Sub PatchV26() on a COPY of DB_v25h.accdb.
'
'  SCOPE: 24 value writes in 15 rows of T_STRUCTURES plus 5 note
'  redactions in T_LOST_ELEMENTS. NO schema change here: the
'  retirement of Jamb_Fabric_Reveal happens in the clean v26
'  rebuild (chachapoya_DB_v26.bas) - the migrator simply does
'  not carry the column. This keeps the patched copy readable
'  by every v25-era query until the moment of migration.
'
'  METHOD: every write has its precondition verified against
'  the live row BEFORE anything is written. One failed check
'  aborts the whole patch with a report and ZERO writes (all
'  writes run inside a single transaction).
'
'  House rules: ASCII only, CRLF, no line continuations, DAO,
'  Private helpers, Val() at every text-to-number frontier,
'  foreign keys resolved BY NAME, never by literal ID.
'  Note: script-written notes are ASCII (no accents); polish
'  diacritics later through the form if wanted.
' ================================================================

Private Const N_STRUCT As Long = 106
Private Const N_CONN As Long = 91
Private Const N_DATING As Long = 7
Private Const N_LOST As Long = 40

Private fails As String
Private nChecks As Long

Public Sub PatchV26()
    Dim db As DAO.Database
    Dim ws As DAO.Workspace
    Set db = CurrentDb()
    Set ws = DBEngine.Workspaces(0)

    fails = ""
    nChecks = 0

    ' ---------- PHASE 1: PRECONDITIONS (read only) ----------
    ChkCount db, "T_STRUCTURES", N_STRUCT
    ChkCount db, "T_CONNECTIONS", N_CONN
    ChkCount db, "T_DATING", N_DATING
    ChkCount db, "T_LOST_ELEMENTS", N_LOST

    ' A. The eleven NULL chamber counters under an Absent system
    Dim aCodes(10) As String
    Dim i As Integer
    aCodes(0) = "DW-S01-EA37": aCodes(1) = "DW-S01-EA60"
    aCodes(2) = "DW-S01-EA61": aCodes(3) = "DW-S01-EA71"
    aCodes(4) = "DW-S01-EA76": aCodes(5) = "DW-S01-EA77"
    aCodes(6) = "DW-S01-EA78": aCodes(7) = "DW-S01-EA79"
    aCodes(8) = "DW-S05-EA02": aCodes(9) = "DW-S06-EA01"
    aCodes(10) = "DW-S06-EA04"
    For i = 0 To 10
        ChkRow db, aCodes(i), "Sys_Chamber='Absent' AND N_Chamber_Bodies Is Null", "A: counter NULL under Absent"
    Next i

    ' B. EA29: the body sits in the wrong column
    ChkRow db, "DW-S01-EA29", "Sys_Chamber='Absent' AND N_Chamber_Bodies=1 AND N_Basal_Bodies=0", "B: EA29 swapped counters"

    ' C. EA20: aerial platform, NA doctrine
    ChkRow db, "DW-S04-EA20", "Sys_Base='Absent' AND Sys_Chamber='Not applicable' AND N_Chamber_Bodies=1", "C: EA20 state"

    ' D. EA35: pigment perimeter, counters must open to NULL
    ChkRow db, "DW-S01-EA35", "Sys_Chamber='Not observable' AND N_Chamber_Bodies=0 AND N_Basal_Bodies=0", "D: EA35 state"

    ' E. EA11: typology regressed to MAU, systems still pending
    ChkRow db, "DW-S01-EA11", "Sys_Base Is Null AND Sys_Chamber Is Null AND Sys_Platform Is Null AND Sys_Eave Is Null", "E: EA11 systems NULL"
    ChkTypoByName db, "DW-S01-EA11", "EA-MAU Mausoleum/Chullpa"
    ChkLookupName db, "L_TYPOLOGY", "Unclassifiable"

    ' F. The five v11 placeholders, addressed by Code + evidence NAME
    ChkLost db, "DW-S01-EA09", "Corbels into void"
    ChkLost db, "DW-S01-EA21", "Corbels into void"
    ChkLost db, "DW-S04-EA20", "Corbels into void"
    ChkLost db, "DW-S01-EA10", "Pigment on bedrock"
    ChkLost db, "DW-S01-EA12", "Truncated walls"

    If Len(fails) > 0 Then
        MsgBox "PATCH ABORTED - no writes performed." & vbCrLf & vbCrLf & "Failed preconditions:" & vbCrLf & fails, vbCritical, "PatchV26"
        Set db = Nothing
        Exit Sub
    End If
    Debug.Print "PatchV26: " & nChecks & " preconditions verified, all clean. Writing."

    ' ---------- PHASE 2: WRITES (one transaction) ----------
    Dim nW As Long
    nW = 0
    ws.BeginTrans

    ' A. Eleven derivable zeros (same doctrine as N_Built_Walls)
    For i = 0 To 10
        nW = nW + Wr(db, "UPDATE T_STRUCTURES SET N_Chamber_Bodies=0 WHERE Code='" & aCodes(i) & "'")
    Next i

    ' B. EA29
    nW = nW + Wr(db, "UPDATE T_STRUCTURES SET N_Chamber_Bodies=0, N_Basal_Bodies=1 WHERE Code='DW-S01-EA29'")

    ' C. EA20 (counter + type-level NA on the base)
    nW = nW + Wr(db, "UPDATE T_STRUCTURES SET N_Chamber_Bodies=0, Sys_Base='Not applicable' WHERE Code='DW-S04-EA20'")

    ' D. EA35 (assertions withdrawn to NULL)
    nW = nW + Wr(db, "UPDATE T_STRUCTURES SET N_Chamber_Bodies=Null, N_Basal_Bodies=Null WHERE Code='DW-S01-EA35'")

    ' E. EA11 (typology resolved BY NAME in VBA, id then embedded)
    Dim unclId As Long
    unclId = LookupIdByName(db, "L_TYPOLOGY", "Unclassifiable")
    Dim q As String
    q = "UPDATE T_STRUCTURES SET "
    q = q & "ID_Typology=" & unclId & ", "
    q = q & "Sys_Base='Not observable', Sys_Chamber='Not observable', "
    q = q & "Sys_Platform='Not observable', Sys_Eave='Not observable', "
    q = q & "N_Chamber_Bodies=Null, N_Basal_Bodies=Null "
    q = q & "WHERE Code='DW-S01-EA11'"
    nW = nW + Wr(db, q)

    ' F. The five notes (ASCII; evidence addressed by name)
    nW = nW + WrLost(db, "DW-S01-EA09", "Corbels into void", "Mensules romanents al buit; l'element sostingut (tauler de plataforma) perdut.")
    nW = nW + WrLost(db, "DW-S01-EA21", "Corbels into void", "Mensules romanents al buit; l'element sostingut (tauler de plataforma) perdut.")
    nW = nW + WrLost(db, "DW-S04-EA20", "Corbels into void", "Tauler de lloses parcialment perdut; les mensules romanen al buit.")
    nW = nW + WrLost(db, "DW-S01-EA10", "Pigment on bedrock", "Pigment sobre penya nua; la fabrica que el portava, desapareguda.")
    nW = nW + WrLost(db, "DW-S01-EA12", "Truncated walls", "Murs truncats en pla net; filades o cossos superiors perduts.")

    ws.CommitTrans

    ' ---------- PHASE 3: REPORT ----------
    Dim msg As String
    msg = "PATCH v26 applied." & vbCrLf
    msg = msg & "Rows touched by UPDATE statements: " & nW & " (expected 20)." & vbCrLf
    msg = msg & "Recount T_STRUCTURES: " & TblCount(db, "T_STRUCTURES") & " (expected " & N_STRUCT & ")" & vbCrLf
    msg = msg & "Recount T_CONNECTIONS: " & TblCount(db, "T_CONNECTIONS") & " (expected " & N_CONN & ")" & vbCrLf
    msg = msg & "Recount T_DATING: " & TblCount(db, "T_DATING") & " (expected " & N_DATING & ")" & vbCrLf
    msg = msg & "Recount T_LOST_ELEMENTS: " & TblCount(db, "T_LOST_ELEMENTS") & " (expected " & N_LOST & ")" & vbCrLf
    msg = msg & vbCrLf & "Next: BuildDB() from chachapoya_DB_v26.bas on a BLANK database, then MigrateFromV25h() pointing at THIS patched copy."
    Debug.Print msg
    MsgBox msg, vbInformation, "PatchV26"
    Set db = Nothing
End Sub

' ---------------- Private helpers ----------------

Private Function LookupIdByName(db As DAO.Database, tbl As String, nm As String) As Long
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT ID FROM " & tbl & " WHERE Name='" & nm & "'", dbOpenSnapshot)
    LookupIdByName = Val(rs!ID & "")
    rs.Close
End Function

Private Function TblCount(db As DAO.Database, tbl As String) As Long
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT Count(*) AS N FROM " & tbl, dbOpenSnapshot)
    TblCount = Val(Nz(rs!N, 0) & "")
    rs.Close
End Function

Private Sub ChkCount(db As DAO.Database, tbl As String, expected As Long)
    nChecks = nChecks + 1
    Dim n As Long
    n = TblCount(db, tbl)
    If n <> expected Then
        fails = fails & "  " & tbl & ": " & n & " rows, expected " & expected & vbCrLf
    End If
End Sub

Private Sub ChkRow(db As DAO.Database, code As String, cond As String, label As String)
    nChecks = nChecks + 1
    Dim rs As DAO.Recordset
    Dim q As String
    q = "SELECT Count(*) AS N FROM T_STRUCTURES WHERE Code='" & code & "' AND (" & cond & ")"
    Set rs = db.OpenRecordset(q, dbOpenSnapshot)
    If Val(Nz(rs!N, 0) & "") <> 1 Then
        fails = fails & "  " & code & " [" & label & "]: current state does not match" & vbCrLf
    End If
    rs.Close
End Sub

Private Sub ChkTypoByName(db As DAO.Database, code As String, typName As String)
    nChecks = nChecks + 1
    Dim rs As DAO.Recordset
    Dim q As String
    q = "SELECT Count(*) AS N FROM T_STRUCTURES AS E "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID "
    q = q & "WHERE E.Code='" & code & "' AND T.Name='" & typName & "'"
    Set rs = db.OpenRecordset(q, dbOpenSnapshot)
    If Val(Nz(rs!N, 0) & "") <> 1 Then
        fails = fails & "  " & code & ": typology is not '" & typName & "' as expected" & vbCrLf
    End If
    rs.Close
End Sub

Private Sub ChkLookupName(db As DAO.Database, tbl As String, nm As String)
    nChecks = nChecks + 1
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT Count(*) AS N FROM " & tbl & " WHERE Name='" & nm & "'", dbOpenSnapshot)
    If Val(Nz(rs!N, 0) & "") <> 1 Then
        fails = fails & "  " & tbl & ": name '" & nm & "' not found exactly once" & vbCrLf
    End If
    rs.Close
End Sub

Private Sub ChkLost(db As DAO.Database, code As String, evName As String)
    nChecks = nChecks + 1
    Dim rs As DAO.Recordset
    Dim q As String
    q = "SELECT Count(*) AS N FROM (T_LOST_ELEMENTS AS L "
    q = q & "INNER JOIN T_STRUCTURES AS E ON L.ID_Structure=E.ID) "
    q = q & "INNER JOIN L_LOST_EVIDENCE AS V ON L.ID_Evidence_Type=V.ID "
    q = q & "WHERE E.Code='" & code & "' AND V.Name='" & evName & "' "
    q = q & "AND L.Notes Like '*Completar element*'"
    Set rs = db.OpenRecordset(q, dbOpenSnapshot)
    If Val(Nz(rs!N, 0) & "") <> 1 Then
        fails = fails & "  " & code & " / " & evName & ": placeholder row not found exactly once" & vbCrLf
    End If
    rs.Close
End Sub

Private Function Wr(db As DAO.Database, sql As String) As Long
    db.Execute sql, dbFailOnError
    Wr = db.RecordsAffected
    Debug.Print "  -> " & db.RecordsAffected & " row: " & Left(sql, 90)
End Function

Private Function Esc(s As String) As String
    ' JET string literal: single quotes must be doubled.
    Esc = Replace(s, "'", "''")
End Function

Private Function WrLost(db As DAO.Database, code As String, evName As String, txt As String) As Long
    Dim q As String
    q = "UPDATE (T_LOST_ELEMENTS AS L "
    q = q & "INNER JOIN T_STRUCTURES AS E ON L.ID_Structure=E.ID) "
    q = q & "INNER JOIN L_LOST_EVIDENCE AS V ON L.ID_Evidence_Type=V.ID "
    q = q & "SET L.Notes='" & Esc(txt) & "' "
    q = q & "WHERE E.Code='" & code & "' AND V.Name='" & evName & "' "
    q = q & "AND L.Notes Like '*Completar element*'"
    db.Execute q, dbFailOnError
    WrLost = db.RecordsAffected
    Debug.Print "  -> nota " & code & " / " & evName & ": " & db.RecordsAffected & " row"
End Function
