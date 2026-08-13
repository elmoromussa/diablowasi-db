Option Compare Database
Option Explicit

' ================================================================
'  CHACHAPOYA PATCH v18 -> v19 (in place, data preserved)
'  Author: Esteve Ribera Torro | TFM Arqueologia UA
'  Spec: DELTA_v18_v19.md
'
'  WHAT THIS DOES, in order:
'    1. T_STRUCTURES gains Jamb_Fabric_Reveal (delta A1), with
'       its padding default of 0.
'    2. Data padding: Jamb_Fabric_Reveal = 0 everywhere. The
'       rows inside the window (Jambs at 0 or 2) keep that 0
'       too, but QRY_24 lists them, because on them the 0 is a
'       migration default and not yet a judgement. The formula
'       'jamb: masonry reveal, dressed' in Systems_Notes is the
'       entry clue (QRY_18_Notes_Review finds it).
'    3. Tie_Walls correction (delta C6): the v18 patch left as
'       assertions the zeros that were padding under the v17
'       gating. They go back to NULL - a judgement nobody made
'       must not read as one - and QRY_24 lists them for the
'       judgement to be actually made once.
'    4. L_TYPOLOGY: MEN promoted (delta C1). Stored Name,
'       Name_VAL and Description change ONCE, here. No rule or
'       query filters by MEN and the corpus holds no MEN
'       record, so this is a type-B1 update.
'    5. L_ELEMENTS, row D (delta C2): Name_EN 'Tie walls'
'       becomes 'Transverse wall' - 'tie' ASSERTS the bracing
'       function, and the function is precisely what we do not
'       know. The description states the anchoring freedom with
'       HEIGHT, never with LEVEL ('level' is a reserved word:
'       N0/N1 are defined by bodies, and a transverse wall is
'       the body of nothing). The FIELD NAME Tie_Walls is NOT
'       touched (field rename = broken queries, v11 lesson).
'
'  WHAT THIS DOES NOT DO: queries and the form. BuildDB() must
'  NEVER run on this database - it drops T_DECORATIONS,
'  T_LOST_ELEMENTS, T_CONNECTIONS and T_ARCH_FEATURES - so the
'  v19 script exposes RebuildQueriesV19(), which recreates the
'  36 queries and touches no table.
'
'  AFTER RUNNING THIS, IN ORDER:
'    1. Import chachapoya_DB_v19.bas -> run RebuildQueriesV19()
'    2. Import chachapoya_Form_v19_val.bas -> run BuildForm()
'    3. Open QRY_24_V19_Review: it is the worklist of this
'       migration (reveal candidates to confirm, D paddings
'       returned to NULL for judgement, bracket roles for the
'       three-question tree).
'
'  Every step is guarded, so re-running the patch is harmless.
' ================================================================

Sub PatchV19()
    Dim db As DAO.Database
    Set db = CurrentDb()

    Debug.Print "=== PATCH v18 -> v19 ==="

    ' Preflight: this must be a v18 Chachapoya database.
    If Not TableExists(db, "T_STRUCTURES") Then
        MsgBox "T_STRUCTURES does not exist. This is not a Chachapoya database.", vbCritical, "Patch aborted"
        Exit Sub
    End If
    If Not ColExists(db, "T_STRUCTURES", "Platform_Surface") Then
        MsgBox "Platform_Surface is missing. This database is older than v18 - apply PATCH v18 first.", vbCritical, "Patch aborted"
        Exit Sub
    End If

    Dim nCols As Long
    Dim nWin As Long
    Dim nTie As Long
    Dim nMen As Long

    ' ---------------------------------------------------------------
    ' 1. New T_STRUCTURES column + padding default (delta A1)
    ' ---------------------------------------------------------------
    nCols = AddCol(db, "T_STRUCTURES", "Jamb_Fabric_Reveal", "BYTE")
    db.TableDefs.Refresh
    SetDefault db, "T_STRUCTURES", "Jamb_Fabric_Reveal", "0"

    ' ---------------------------------------------------------------
    ' 2. Data padding on the existing rows
    ' ---------------------------------------------------------------
    ' Padding 0 EVERYWHERE, exactly as Sill_Coincides_Cornice was
    ' handled in v18: the window rows keep the 0 but QRY_24 lists
    ' them, because on them it is a migration default and not yet
    ' a judgement.
    db.Execute "UPDATE T_STRUCTURES SET Jamb_Fabric_Reveal=0 WHERE Jamb_Fabric_Reveal Is Null", dbFailOnError
    nWin = DCount("*", "T_STRUCTURES", "Jambs IN (0,2)")
    Debug.Print "-> Jamb_Fabric_Reveal padded; " & nWin & " row(s) inside the window await confirmation (QRY_24)"

    ' ---------------------------------------------------------------
    ' 3. Tie_Walls correction (delta C6)
    ' ---------------------------------------------------------------
    ' Under the v17 gating these zeros were padding; the v18 patch
    ' left them in place and the ungated field now reads them as
    ' assertions. Back to NULL: they join the rule-17 worklist for
    ' the real judgement. The zeros under Sys_Base = Present are
    ' NOT touched - the field was open there in v17, so they are
    ' presumed judgements, and the delimitation-test review (delta,
    ' outside section) confirms them by hand.
    nTie = DCount("*", "T_STRUCTURES", "Tie_Walls=0 AND Sys_Base IN ('Absent','Not applicable','Not observable')")
    db.Execute "UPDATE T_STRUCTURES SET Tie_Walls=Null WHERE Tie_Walls=0 AND Sys_Base IN ('Absent','Not applicable','Not observable')", dbFailOnError
    Debug.Print "-> Tie_Walls: " & nTie & " padding zero(s) returned to NULL for judgement (QRY_24)"

    ' ---------------------------------------------------------------
    ' 4. L_TYPOLOGY: MEN promoted (delta C1). BY NAME, never by ID.
    ' ---------------------------------------------------------------
    nMen = DCount("*", "L_TYPOLOGY", "Name='MEN Isolated Bracket'")
    If nMen = 1 Then
        db.Execute "UPDATE L_TYPOLOGY SET Name='MEN Isolated Structural Element', Name_VAL='MEN Element estructural aillat', Description='Record whose evidence reduces to one or a few A-Z elements without a classifiable structure (bracket, transverse wall, pilaster). Interpretation (circulation, earlier structure, support) goes to T_ARCH_FEATURES / Notes, never to the typology. The historical name comes from the first documented case, the bracket.' WHERE Name='MEN Isolated Bracket'", dbFailOnError
        Debug.Print "-> L_TYPOLOGY: MEN promoted to Isolated Structural Element"
    Else
        Debug.Print "-> L_TYPOLOGY: MEN already promoted (or absent)"
    End If

    ' ---------------------------------------------------------------
    ' 5. L_ELEMENTS, row D (delta C2). BY CODE, never by ID.
    ' ---------------------------------------------------------------
    db.Execute "UPDATE L_ELEMENTS SET Name_EN='Transverse wall', Description='Transverse wall anchoring the fabric to the rock; may anchor at any height of the fabric. Does not enclose any interior and does not count as a body. If it encloses a chamber it is a return wall (V); ungated since v18, permanently active like I and R.' WHERE Code='D'", dbFailOnError
    Debug.Print "-> L_ELEMENTS: D is now Transverse wall (field name Tie_Walls untouched)"

    Set db = Nothing

    Dim msg As String
    msg = "PATCH v19 APPLIED" & vbCrLf & vbCrLf
    msg = msg & "  New columns added: " & nCols & " of 1" & vbCrLf
    msg = msg & "  Reveal window rows to confirm: " & nWin & vbCrLf
    msg = msg & "  Tie_Walls paddings returned to NULL: " & nTie & vbCrLf
    msg = msg & "  MEN typology promoted: " & nMen & vbCrLf & vbCrLf
    msg = msg & "NEXT, IN THIS ORDER:" & vbCrLf & vbCrLf
    msg = msg & "1. Import chachapoya_DB_v19.bas and run" & vbCrLf
    msg = msg & "   RebuildQueriesV19() - queries only, no" & vbCrLf
    msg = msg & "   table is touched. NEVER run BuildDB() on" & vbCrLf
    msg = msg & "   this database." & vbCrLf & vbCrLf
    msg = msg & "2. Import chachapoya_Form_v19_val.bas and" & vbCrLf
    msg = msg & "   run BuildForm()." & vbCrLf & vbCrLf
    msg = msg & "3. Open QRY_24_V19_Review: it is the" & vbCrLf
    msg = msg & "   worklist of this migration." & vbCrLf
    MsgBox msg, vbInformation, "Patch v19"
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
