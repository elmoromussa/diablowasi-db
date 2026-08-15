Option Compare Database
Option Explicit

' ================================================================
'  CHACHAPOYA PATCH v22 -> v23 (in place, data preserved)
'  Author: Esteve Ribera Torro | TFM Arqueologia UA
'  Spec: DELTA_v22_v23.md
'
'  WHAT THIS DOES, in order:
'    0. Preflight: v22 marker (Metrics_Available).
'    1. ADD EA_Number (bloc 5) and derive it from the code for
'       every record - explicit reversal of the v20 never-store
'       decision; rule 67 guards the concordance from here on.
'    2. CREATE L_SUBSECTORS + populate + ADD ID_Subsector +
'       relation (bloc 6). Values stay NULL: the subsector is
'       assigned record by record when known.
'    3. CHAMBER WALLS migrations (bloc 4), all derivable:
'       - chamber Absent or Not applicable with the count still
'         NULL -> 0 (the derivable padding; ~24 rows expected);
'       - chamber Not observable or unassessed with the count at
'         0 -> NULL (you cannot count the walls of a chamber you
'         cannot see; ~2 rows expected).
'       The 11 rows with walls counted and no chamber are NOT
'       touched: rule 66 lists them - that review is yours.
'
'  AFTER RUNNING THIS, IN ORDER:
'    1. chachapoya_DB_v23.bas -> RebuildQueriesV23()
'       (NEVER BuildDB() on this database)
'    2. chachapoya_Form_v23_val.bas -> BuildForm()
'       (with F_STRUCTURES CLOSED first)
'    3. chachapoya_Worklist_v23.bas -> BuildWorklist()
'    4. QRY_16 (Bateria): rule 66 lists the wall counts to
'       review row by row; rule 55 revised shrinks the old
'       fabric-reveal residue on QRY_24 to the 8 true rows.
'
'  Every step is guarded, so re-running the patch is harmless.
' ================================================================

Sub PatchV23()
    Dim db As DAO.Database
    Set db = CurrentDb()

    Debug.Print "=== PATCH v22 -> v23 ==="

    If Not TableExists(db, "T_STRUCTURES") Then
        MsgBox "T_STRUCTURES does not exist. This is not a Chachapoya database.", vbCritical, "Patch aborted"
        Exit Sub
    End If
    If Not ColExists(db, "T_STRUCTURES", "Metrics_Available") Then
        MsgBox "Metrics_Available is missing. This database is older than v22 - apply PATCH v22 first.", vbCritical, "Patch aborted"
        Exit Sub
    End If

    Dim nAdd As Long

    ' ---------------------------------------------------------------
    ' 1. EA_Number: add and derive (bloc 5)
    ' ---------------------------------------------------------------
    nAdd = nAdd + AddCol(db, "T_STRUCTURES", "EA_Number", "INTEGER")

    Dim nEA As Long
    db.Execute "UPDATE T_STRUCTURES SET EA_Number = Val(Mid(Code, InStr(Code,'EA')+2)) WHERE InStr(Code,'EA')>0 AND (EA_Number Is Null Or EA_Number<>Val(Mid(Code, InStr(Code,'EA')+2)))", dbFailOnError
    nEA = db.RecordsAffected
    Debug.Print "-> EA_Number derived on " & nEA & " record(s)"

    ' ---------------------------------------------------------------
    ' 2. Subsector: lookup, column, relation (bloc 6)
    ' ---------------------------------------------------------------
    Dim nSub As Long
    If Not TableExists(db, "L_SUBSECTORS") Then
        db.Execute "CREATE TABLE L_SUBSECTORS (ID COUNTER CONSTRAINT PK_SSEC PRIMARY KEY, Name TEXT(30) NOT NULL, Name_VAL TEXT(30))", dbFailOnError
        db.Execute "INSERT INTO L_SUBSECTORS (Name,Name_VAL) VALUES ('Upper','Superior (sup.)')", dbFailOnError
        db.Execute "INSERT INTO L_SUBSECTORS (Name,Name_VAL) VALUES ('Lower','Inferior (inf.)')", dbFailOnError
        db.Execute "INSERT INTO L_SUBSECTORS (Name,Name_VAL) VALUES ('North','Nord (N)')", dbFailOnError
        db.Execute "INSERT INTO L_SUBSECTORS (Name,Name_VAL) VALUES ('Central','Central (C)')", dbFailOnError
        db.Execute "INSERT INTO L_SUBSECTORS (Name,Name_VAL) VALUES ('South','Sud (S)')", dbFailOnError
        nSub = 1
        Debug.Print "-> L_SUBSECTORS created and populated (5 values)"
    Else
        Debug.Print "-> L_SUBSECTORS already present"
    End If
    nAdd = nAdd + AddCol(db, "T_STRUCTURES", "ID_Subsector", "LONG")

    On Error Resume Next
    db.Relations.Delete "REL_SSEC_STR"
    On Error GoTo 0
    Dim rel As DAO.Relation
    Set rel = db.CreateRelation("REL_SSEC_STR", "L_SUBSECTORS", "T_STRUCTURES", 0)
    Dim fld As DAO.Field
    Set fld = rel.CreateField("ID")
    fld.ForeignName = "ID_Subsector"
    rel.Fields.Append fld
    db.Relations.Append rel
    Debug.Print "-> relation REL_SSEC_STR created"

    ' ---------------------------------------------------------------
    ' 3. Chamber walls migrations (bloc 4)
    ' ---------------------------------------------------------------
    Dim nZero As Long, nNul As Long
    db.Execute "UPDATE T_STRUCTURES SET N_Built_Walls=0 WHERE N_Built_Walls Is Null AND Sys_Chamber IN ('Absent','Not applicable')", dbFailOnError
    nZero = db.RecordsAffected
    Debug.Print "-> derivable 0 filled on " & nZero & " record(s) with no chamber"

    db.Execute "UPDATE T_STRUCTURES SET N_Built_Walls=Null WHERE N_Built_Walls=0 AND (Sys_Chamber Is Null Or Sys_Chamber='Not observable')", dbFailOnError
    nNul = db.RecordsAffected
    Debug.Print "-> 0 returned to NULL on " & nNul & " record(s) with the chamber unknowable"

    Dim rs As DAO.Recordset
    Dim wList As String
    Dim nW As Long
    Set rs = db.OpenRecordset("SELECT Code FROM T_STRUCTURES WHERE N_Built_Walls>0 AND (Sys_Chamber Is Null Or Sys_Chamber IN ('Absent','Not applicable','Not observable')) ORDER BY Code")
    Do While Not rs.EOF
        wList = wList & "  " & rs!Code & vbCrLf
        nW = nW + 1
        rs.MoveNext
    Loop
    rs.Close
    If nW > 0 Then
        Debug.Print "-> " & nW & " wall count(s) WITHOUT a chamber, to review BY HAND (rule 66):"
        Debug.Print wList
    End If

    Set db = Nothing

    Dim msg As String
    msg = "PATCH v23 APPLIED" & vbCrLf & vbCrLf
    msg = msg & "  Columns added: " & nAdd & vbCrLf
    msg = msg & "  EA_Number derived: " & nEA & vbCrLf
    msg = msg & "  Chamber walls -> derivable 0: " & nZero & vbCrLf
    msg = msg & "  Chamber walls -> NULL (unknowable): " & nNul & vbCrLf
    msg = msg & "  Wall counts to review by hand (rule 66): " & nW & vbCrLf & vbCrLf
    msg = msg & "NEXT, IN THIS ORDER:" & vbCrLf & vbCrLf
    msg = msg & "1. chachapoya_DB_v23.bas -> RebuildQueriesV23()" & vbCrLf
    msg = msg & "   (NEVER BuildDB() on this database)" & vbCrLf & vbCrLf
    msg = msg & "2. chachapoya_Form_v23_val.bas -> BuildForm()" & vbCrLf
    msg = msg & "   (amb F_STRUCTURES TANCAT abans)" & vbCrLf & vbCrLf
    msg = msg & "3. chachapoya_Worklist_v23.bas -> BuildWorklist()" & vbCrLf & vbCrLf
    msg = msg & "4. QRY_16: la regla 66 llista els recomptes de" & vbCrLf
    msg = msg & "   murs a revisar fila a fila." & vbCrLf
    MsgBox msg, vbInformation, "Patch v23"
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
