Option Compare Database
Option Explicit

' ================================================================
'  CHACHAPOYA PATCH v17a -> v18 (in place, data preserved)
'  Author: Esteve Ribera Torro | TFM Arqueologia UA
'  Spec: DELTA_v17_v18.md
'
'  WHAT THIS DOES, in order:
'    1. T_STRUCTURES gains Platform_Surface (Z), 
'       Interbody_Cornice_Format and Sill_Coincides_Cornice,
'       with their padding defaults.
'    2. Data padding: Sill_Coincides_Cornice = 0 everywhere
'       (the coincidence window is listed in QRY_23 for manual
'       review); Platform_Surface = 0 under a closed platform
'       system, 9 under Not observable, NULL where the system
'       is open - those rows are the observation worklist.
'    3. Stone_Format / _Secondary: 'Tabular blocks' becomes
'       'Regular tabular blocks' (delta B1). The stored value
'       changes ONCE, here, so v17 exports and v18 exports
'       never mix the two spellings.
'    4. L_ELEMENTS gains Z (Platform surface); the description
'       of D records its v18 ungating. Y stays reserved.
'    5. L_STRUCT_BODY: RTW label becomes 'Mur de retorn (V)'
'       (delta A6). BY CODE, never by ID.
'    6. T_CONNECTIONS gains ID_Earlier; the positional values
'       'A is earlier' / 'B is earlier' (and their v16
'       spellings) migrate to Sequential + the named FK; the
'       REL_STR_CONE relationship is created.
'    7. T_LOST_ELEMENTS: the lone ID_Position value travels to
'       Notes with a constant formula, the REL_SB_LOST
'       relationship is deleted, and the column is dropped.
'
'  WHAT THIS DOES NOT DO: queries and the form. BuildDB() must
'  NEVER run on this database - it drops T_DECORATIONS,
'  T_LOST_ELEMENTS, T_CONNECTIONS and T_ARCH_FEATURES - so the
'  v18 script exposes RebuildQueriesV18(), which recreates the
'  35 queries and touches no table.
'
'  AFTER RUNNING THIS, IN ORDER:
'    1. Import chachapoya_DB_v18.bas -> run RebuildQueriesV18()
'    2. Import chachapoya_Form_v18_val.bas -> run BuildForm()
'    3. Open QRY_23_V18_Review: it is the worklist of this
'       migration (Z to observe, D paddings to confirm,
'       cornice-as-sill candidates, scope incoherences).
'
'  Every step is guarded, so re-running the patch is harmless.
' ================================================================

Sub PatchV18()
    Dim db As DAO.Database
    Set db = CurrentDb()

    Debug.Print "=== PATCH v17a -> v18 ==="

    ' Preflight: this must be a v17a Chachapoya database.
    If Not TableExists(db, "T_STRUCTURES") Then
        MsgBox "T_STRUCTURES does not exist. This is not a Chachapoya database.", vbCritical, "Patch aborted"
        Exit Sub
    End If
    If Not ColExists(db, "T_STRUCTURES", "Doc_Notes") Then
        MsgBox "Doc_Notes is missing. This database is older than v17a - apply PATCH v17a first.", vbCritical, "Patch aborted"
        Exit Sub
    End If

    Dim nCols As Long
    Dim nPad As Long
    Dim nOpen As Long
    Dim nStone As Long
    Dim nZ As Long
    Dim nConn As Long
    Dim nPos As Long

    ' ---------------------------------------------------------------
    ' 1. New T_STRUCTURES columns + padding defaults
    ' ---------------------------------------------------------------
    nCols = nCols + AddCol(db, "T_STRUCTURES", "Platform_Surface", "BYTE")
    nCols = nCols + AddCol(db, "T_STRUCTURES", "Interbody_Cornice_Format", "TEXT(20)")
    nCols = nCols + AddCol(db, "T_STRUCTURES", "Sill_Coincides_Cornice", "BYTE")
    db.TableDefs.Refresh
    SetDefault db, "T_STRUCTURES", "Platform_Surface", "0"
    SetDefault db, "T_STRUCTURES", "Sill_Coincides_Cornice", "0"

    ' ---------------------------------------------------------------
    ' 2. Data padding on the existing rows
    ' ---------------------------------------------------------------
    ' The coincidence field takes the padding 0 EVERYWHERE: the
    ' rows inside the window (Sill=0 with a real cornice) keep
    ' that 0 too, but QRY_23 lists them, because on them the 0 is
    ' a migration default and not yet a judgement.
    db.Execute "UPDATE T_STRUCTURES SET Sill_Coincides_Cornice=0 WHERE Sill_Coincides_Cornice Is Null", dbFailOnError

    ' Z: padding under a closed system, 9 under Not observable,
    ' NULL where the system is open - the conservative principle:
    ' a value carried across unobserved would assert a judgement
    ' nobody made.
    db.Execute "UPDATE T_STRUCTURES SET Platform_Surface=0 WHERE Sys_Platform IN ('Absent','Not applicable') AND Platform_Surface Is Null", dbFailOnError
    db.Execute "UPDATE T_STRUCTURES SET Platform_Surface=9 WHERE Sys_Platform='Not observable' AND Platform_Surface Is Null", dbFailOnError
    nPad = DCount("*", "T_STRUCTURES", "Platform_Surface Is Not Null")
    nOpen = DCount("*", "T_STRUCTURES", "Platform_Surface Is Null")
    Debug.Print "-> Platform_Surface: " & nPad & " padded, " & nOpen & " left NULL for observation (QRY_23)"

    ' ---------------------------------------------------------------
    ' 3. Stone_Format rename (delta B1)
    ' ---------------------------------------------------------------
    nStone = DCount("*", "T_STRUCTURES", "Stone_Format='Tabular blocks'")
    nStone = nStone + DCount("*", "T_STRUCTURES", "Stone_Format_Secondary='Tabular blocks'")
    db.Execute "UPDATE T_STRUCTURES SET Stone_Format='Regular tabular blocks' WHERE Stone_Format='Tabular blocks'", dbFailOnError
    db.Execute "UPDATE T_STRUCTURES SET Stone_Format_Secondary='Regular tabular blocks' WHERE Stone_Format_Secondary='Tabular blocks'", dbFailOnError
    Debug.Print "-> Stone format: " & nStone & " value(s) renamed to Regular tabular blocks"

    ' ---------------------------------------------------------------
    ' 4. L_ELEMENTS: Z enters, D gets its v18 description
    ' ---------------------------------------------------------------
    If DCount("*", "L_ELEMENTS", "Code='Z'") = 0 Then
        db.Execute "INSERT INTO L_ELEMENTS (Code, Name_EN, Name_VAL, Level_Type, Sys_Group, Is_System, Field_Name, Description) VALUES ('Z','Platform surface','Superficie de plataforma','N0','Platform',False,'Platform_Surface','Finished walking surface of the corbelled platform; component of the platform system (H). Material in Platform_Surface_Material.')", dbFailOnError
        nZ = 1
        Debug.Print "-> L_ELEMENTS: Z inserted (Y stays reserved for the banqueta)"
    Else
        Debug.Print "-> L_ELEMENTS: Z already present"
    End If
    db.Execute "UPDATE L_ELEMENTS SET Description='Transverse tie wall anchoring the fabric to the rock, at any level. Ungated from Sys_Base in v18 (delta C2): permanently active, like I and R.' WHERE Code='D'", dbFailOnError

    ' ---------------------------------------------------------------
    ' 5. L_STRUCT_BODY: one term everywhere (delta A6). BY CODE.
    ' ---------------------------------------------------------------
    db.Execute "UPDATE L_STRUCT_BODY SET Name_VAL='Mur de retorn (V)' WHERE Code='RTW'", dbFailOnError
    Debug.Print "-> L_STRUCT_BODY: RTW label is now Mur de retorn (V)"

    ' ---------------------------------------------------------------
    ' 6. T_CONNECTIONS: the direction leaves the pair (delta D1)
    ' ---------------------------------------------------------------
    nCols = nCols + AddCol(db, "T_CONNECTIONS", "ID_Earlier", "LONG")
    db.TableDefs.Refresh
    ' Positional values migrate to the named FK. The v16 corpus
    ' spelt them differently, so both spellings are covered; on
    ' the current data every row is Undetermined and these are
    ' no-ops, kept because other copies of the file exist.
    nConn = DCount("*", "T_CONNECTIONS", "Chrono_Relation IN ('A is earlier','A earlier than B')")
    db.Execute "UPDATE T_CONNECTIONS SET ID_Earlier=ID_Struct_A, Chrono_Relation='Sequential' WHERE Chrono_Relation IN ('A is earlier','A earlier than B')", dbFailOnError
    nConn = nConn + DCount("*", "T_CONNECTIONS", "Chrono_Relation IN ('B is earlier','B earlier than A')")
    db.Execute "UPDATE T_CONNECTIONS SET ID_Earlier=ID_Struct_B, Chrono_Relation='Sequential' WHERE Chrono_Relation IN ('B is earlier','B earlier than A')", dbFailOnError
    Debug.Print "-> T_CONNECTIONS: " & nConn & " directed edge(s) migrated to Sequential + ID_Earlier"

    ' Third edge into T_STRUCTURES: no enforced integrity (flag 2),
    ' same treatment as A and B.
    On Error Resume Next: db.Relations.Delete "REL_STR_CONE": On Error GoTo 0
    MkRelNoInt db, "REL_STR_CONE", "T_STRUCTURES", "ID", "T_CONNECTIONS", "ID_Earlier"

    ' ---------------------------------------------------------------
    ' 7. T_LOST_ELEMENTS: ID_Position out (delta D2)
    ' ---------------------------------------------------------------
    If ColExists(db, "T_LOST_ELEMENTS", "ID_Position") Then
        nPos = SavePositionToNotes(db)
        On Error Resume Next: db.Relations.Delete "REL_SB_LOST": On Error GoTo 0
        db.Execute "ALTER TABLE T_LOST_ELEMENTS DROP COLUMN ID_Position", dbFailOnError
        db.TableDefs.Refresh
        Debug.Print "-> ID_Position: " & nPos & " value(s) copied into Notes, column dropped"
    Else
        Debug.Print "-> ID_Position already gone"
    End If

    Set db = Nothing

    Dim msg As String
    msg = "PATCH v18 APPLIED" & vbCrLf & vbCrLf
    msg = msg & "  New columns added: " & nCols & " of 4" & vbCrLf
    msg = msg & "  Platform_Surface: " & nOpen & " row(s) await observation" & vbCrLf
    msg = msg & "  Stone format values renamed: " & nStone & vbCrLf
    If nZ = 1 Then msg = msg & "  Z inserted into L_ELEMENTS" & vbCrLf
    msg = msg & "  Directed connections migrated: " & nConn & vbCrLf
    msg = msg & "  ID_Position values saved to Notes: " & nPos & vbCrLf & vbCrLf
    msg = msg & "NEXT, IN THIS ORDER:" & vbCrLf & vbCrLf
    msg = msg & "1. Import chachapoya_DB_v18.bas and run" & vbCrLf
    msg = msg & "   RebuildQueriesV18() - queries only, no" & vbCrLf
    msg = msg & "   table is touched. NEVER run BuildDB() on" & vbCrLf
    msg = msg & "   this database." & vbCrLf & vbCrLf
    msg = msg & "2. Import chachapoya_Form_v18_val.bas and" & vbCrLf
    msg = msg & "   run BuildForm()." & vbCrLf & vbCrLf
    msg = msg & "3. Open QRY_23_V18_Review: it is the" & vbCrLf
    msg = msg & "   worklist of this migration." & vbCrLf
    MsgBox msg, vbInformation, "Patch v18"
    Debug.Print "=== PATCH FINISHED ==="
End Sub


' ================================================================
'  Guarded ALTER TABLE ADD COLUMN: 1 if added, 0 if present.
' ================================================================
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

' Sets a DefaultValue and reports on failure - a defaults routine
' that lies is worse than one that does nothing.
Private Sub SetDefault(db As DAO.Database, tbl As String, fld As String, val As String)
    On Error GoTo Err_SD
    db.TableDefs(tbl).Fields(fld).DefaultValue = val
    Debug.Print "-> default " & val & " on " & fld
    Exit Sub
Err_SD:
    Debug.Print "  *** default NOT set on " & tbl & "." & fld & ": " & Err.Description
End Sub

' ================================================================
'  ID_Position -> Notes, with a constant formula so the saved
'  values are findable later. Appended, never overwriting. BY
'  NAME through the lookup, never by raw id.
' ================================================================
Private Function SavePositionToNotes(db As DAO.Database) As Long
    Dim rs As DAO.Recordset
    Dim k As Long
    Dim nm As Variant
    Set rs = db.OpenRecordset("SELECT ID, Notes, ID_Position FROM T_LOST_ELEMENTS WHERE ID_Position Is Not Null", dbOpenDynaset)
    Do While Not rs.EOF
        nm = DLookup("Name", "L_STRUCT_BODY", "ID=" & rs!ID_Position)
        If IsNull(nm) Then nm = "id " & rs!ID_Position
        Dim t As String
        t = "position: " & CStr(nm)
        rs.Edit
        If IsNull(rs!Notes) Then
            rs!Notes = t
        Else
            rs!Notes = rs!Notes & " | " & t
        End If
        rs.Update
        k = k + 1
        rs.MoveNext
    Loop
    rs.Close
    SavePositionToNotes = k
End Function

' Relation with no enforced integrity: numeric flag 2, the same
'  treatment REL_STR_SELF and the A / B edges already use.
Private Sub MkRelNoInt(db As DAO.Database, nm As String, pT As String, pF As String, cT As String, cF As String)
    Dim rel As DAO.Relation
    Dim fld As DAO.Field
    On Error GoTo ErrR
    Set rel = db.CreateRelation(nm, pT, cT, 2)
    Set fld = rel.CreateField(pF)
    fld.ForeignName = cF
    rel.Fields.Append fld
    db.Relations.Append rel
    Debug.Print "-> relationship " & nm & " created"
    Exit Sub
ErrR:
    Debug.Print "  Warning " & nm & ": " & Err.Description
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
