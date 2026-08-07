Option Compare Database
Option Explicit

' ================================================================
'  CHACHAPOYA - PATCH v17 -> v17a
'  Run Sub PatchV17a() ON AN EXISTING v17 DATABASE THAT ALREADY
'  HAS RECORDS. If nothing has been entered yet, rebuild from
'  scratch with the corrected chachapoya_DB_v17.bas instead: a
'  rebuild is always cleaner than a patch.
'
'  WHAT IT DOES, and why each piece is here.
'
'  1. NOTES ON EVERY TAB. Five tabs had no notes field of their
'     own - identification, decoration, chronology, metrics and
'     extras - so anything that did not fit an existing field on
'     those tabs had nowhere to go. QRY_18_Notes_Review exists to
'     turn repeated free text into new fields, and it can only do
'     that on tabs where free text is possible at all: the v17
'     field Outline_Geometry was born exactly that way, from a
'     word typed by hand into a Notes box.
'     The general field Notes is renamed Doc_Notes, so that all
'     seven now follow one pattern instead of six plus an
'     exception. RENAMING PRESERVES THE DATA - it is a TableDef
'     rename, not a drop and recreate.
'
'  2. THE ROCK ART VOCABULARY, REBUILT ON ONE AXIS. Entering data
'     showed the list was confusing, and the reason was
'     structural rather than a matter of wording: it mixed axes.
'     'Painted band' named a SHAPE while 'Perimeter band' named a
'     POSITION, so a perimeter band was also a painted band.
'     'Abstract' and 'Amorphous stain' both named DEGREE OF
'     LEGIBILITY and neither said where the boundary was. A list
'     whose entries answer different questions cannot be applied
'     consistently, however carefully each entry is defined.
'     The new list answers ONE question - what motif is it - and
'     the two U values absorb what Outline_Geometry was recording
'     separately.
'
'  3. OUTLINE_GEOMETRY IS DROPPED. It only ever discriminated
'     within the perimeter band, so as a general field it was a
'     column that stayed empty on most rows and asked a question
'     that had no meaning on them. Folded into the type, the same
'     information costs nothing and cannot be left inconsistent
'     with it.
'     BEFORE THE COLUMN GOES, ITS VALUES ARE COPIED INTO Notes on
'     each affected row. Dropping a column silently discards
'     judgements someone made; and here those judgements are
'     precisely what is needed to choose between the two U values.
'
'  4. WHAT THIS PATCH REFUSES TO DECIDE. 'RA Perimeter band' rows
'     are left WITHOUT A TYPE and listed. Geometric or organic is
'     a reading of the photograph, and a reading made by a patch
'     script is a reading nobody made. Everything else remaps
'     mechanically because the meaning is unchanged.
'
'  5. RULE 41 IS RETIRED and THE OTHERS ARE NOT RENUMBERED. The
'     numbers are cited across four documents; renumbering to
'     close a gap would invalidate every one of those citations
'     to save nothing. 45 active rules, numbered up to 46.
'
'  AFTER RUNNING THIS: re-run chachapoya_Form_v17_val.bas ->
'  BuildForm(), which rebuilds the form against the new schema.
' ================================================================

Sub PatchV17a()
    Dim db As DAO.Database
    Set db = CurrentDb()

    Debug.Print "=== PATCH v17 -> v17a ==="

    If Not TableExists(db, "T_DECORATIONS") Then
        MsgBox "T_DECORATIONS does not exist. This is not a Chachapoya database.", vbCritical, "Patch aborted"
        Exit Sub
    End If
    If Not ColExists(db, "T_DECORATIONS", "Span_Left") Then
        MsgBox "Span_Left is missing. This database is older than v17 - build v17 first.", vbCritical, "Patch aborted"
        Exit Sub
    End If

    Dim nNotes As Long
    Dim nRenamed As Long
    Dim nSaved As Long
    Dim nRemap As Long
    Dim nPerim As Long
    Dim nGeoDropped As Long

    ' ---------------------------------------------------------------
    ' 1. Notes fields
    ' ---------------------------------------------------------------
    nNotes = nNotes + AddMemo(db, "T_STRUCTURES", "Id_Notes")
    nNotes = nNotes + AddMemo(db, "T_STRUCTURES", "Dec_Notes")
    nNotes = nNotes + AddMemo(db, "T_STRUCTURES", "Chrono_Notes")
    nNotes = nNotes + AddMemo(db, "T_STRUCTURES", "Metric_Notes")
    nNotes = nNotes + AddMemo(db, "T_STRUCTURES", "Extra_Notes")

    ' A TableDef rename keeps the stored values. Guarded both ways so
    ' that re-running the patch is harmless.
    If ColExists(db, "T_STRUCTURES", "Notes") And Not ColExists(db, "T_STRUCTURES", "Doc_Notes") Then
        db.TableDefs("T_STRUCTURES").Fields("Notes").Name = "Doc_Notes"
        db.TableDefs.Refresh
        nRenamed = 1
        Debug.Print "-> Notes renamed to Doc_Notes (values preserved)"
    Else
        Debug.Print "-> Doc_Notes already present or Notes already gone"
    End If

    ' ---------------------------------------------------------------
    ' 2. Outline_Geometry: save into Notes, then drop
    ' ---------------------------------------------------------------
    If ColExists(db, "T_DECORATIONS", "Outline_Geometry") Then
        nSaved = SaveOutlineToNotes(db)
        db.Execute "ALTER TABLE T_DECORATIONS DROP COLUMN Outline_Geometry", dbFailOnError
        db.TableDefs.Refresh
        nGeoDropped = 1
        Debug.Print "-> Outline_Geometry: " & nSaved & " value(s) copied into Notes, column dropped"
    Else
        Debug.Print "-> Outline_Geometry already gone"
    End If

    ' ---------------------------------------------------------------
    ' 3. L_DEC_TYPE: remap the rows, then rebuild the rock art half
    ' ---------------------------------------------------------------
    ' ORDER MATTERS. The rows are repointed BEFORE any lookup entry
    ' is deleted, because a decoration row pointing at a deleted id
    ' is not detectably wrong afterwards - it just reads as blank.
    nRemap = nRemap + Repoint(db, "RA Abstract", "RA Geometric motif")
    nRemap = nRemap + Repoint(db, "RA Geometric", "RA Geometric motif")

    ' Perimeter band rows are NOT repointed: geometric or organic is
    ' a judgement. They are cleared and reported instead.
    nPerim = ClearType(db, "RA Perimeter band")

    RebuildRockArtTypes db

    ' ---------------------------------------------------------------
    ' 4. Rule 41 out of the battery
    ' ---------------------------------------------------------------
    Dim ruleFixed As Boolean
    ruleFixed = RetireRule41(db)

    Dim n As Long
    n = DCount("*", "L_DEC_TYPE")

    Set db = Nothing

    Dim msg As String
    msg = "PATCH v17a APPLIED" & vbCrLf & vbCrLf
    msg = msg & "  Notes fields added: " & nNotes & " of 5" & vbCrLf
    If nRenamed = 1 Then msg = msg & "  Notes renamed to Doc_Notes" & vbCrLf
    If nGeoDropped = 1 Then
        msg = msg & "  Outline_Geometry dropped; " & nSaved & " value(s)" & vbCrLf
        msg = msg & "  copied into the Notes of their row first" & vbCrLf
    End If
    msg = msg & "  Decoration rows remapped: " & nRemap & vbCrLf
    msg = msg & "  L_DEC_TYPE entries now: " & n & vbCrLf
    If ruleFixed Then msg = msg & "  Rule 41 retired (45 active, numbered to 46)" & vbCrLf
    msg = msg & vbCrLf
    msg = msg & "NEXT, IN THIS ORDER:" & vbCrLf & vbCrLf
    msg = msg & "1. Run chachapoya_Form_v17_val.bas -> BuildForm()." & vbCrLf
    msg = msg & "   The form is rebuilt against the new schema." & vbCrLf & vbCrLf
    If nPerim > 0 Then
        msg = msg & "2. " & nPerim & " row(s) were left WITHOUT A TYPE:" & vbCrLf
        msg = msg & "   they were Perimeter band, and choosing between" & vbCrLf
        msg = msg & "   U-shape geometric and U-shape organic is a" & vbCrLf
        msg = msg & "   reading of the photograph, not something a" & vbCrLf
        msg = msg & "   script should decide. Their old geometry, if" & vbCrLf
        msg = msg & "   any was recorded, is now in their Notes." & vbCrLf
        msg = msg & "   Open QRY_22_V17a_Review for the list." & vbCrLf
    Else
        msg = msg & "2. No perimeter band row needed manual review." & vbCrLf
    End If
    MsgBox msg, vbInformation, "Patch v17a"
    Debug.Print "=== PATCH FINISHED ==="
End Sub

' ================================================================
'  MEMO rather than TEXT, like the six notes fields that already
'  exist: a note that has to fit in 255 characters stops being a
'  note and starts being a code the observer invents.
' ================================================================
Private Function AddMemo(db As DAO.Database, tbl As String, col As String) As Long
    If ColExists(db, tbl, col) Then
        Debug.Print "-> " & col & " already present"
        Exit Function
    End If
    db.Execute "ALTER TABLE " & tbl & " ADD COLUMN " & col & " MEMO", dbFailOnError
    db.TableDefs.Refresh
    AddMemo = 1
    Debug.Print "-> " & col & " added"
End Function

' ================================================================
'  Outline_Geometry -> Notes, with a constant formula so that the
'  saved values are findable later. Appended, never overwriting an
'  existing note.
' ================================================================
Private Function SaveOutlineToNotes(db As DAO.Database) As Long
    Dim rs As DAO.Recordset
    Dim k As Long
    Set rs = db.OpenRecordset("SELECT ID, Notes, Outline_Geometry FROM T_DECORATIONS WHERE Outline_Geometry Is Not Null", dbOpenDynaset)
    Do While Not rs.EOF
        Dim t As String
        t = "outline: " & CStr(rs!Outline_Geometry)
        rs.Edit
        If IsNull(rs!Notes) Then
            rs!Notes = t
        Else
            rs!Notes = Left$(rs!Notes & " | " & t, 255)
        End If
        rs.Update
        k = k + 1
        rs.MoveNext
    Loop
    rs.Close
    SaveOutlineToNotes = k
End Function

' ================================================================
'  Repoint every decoration row from one type name to another.
'  BY NAME, never by id: ids are reorganised between versions and
'  a wrong id raises no error and shows no symptom.
' ================================================================
Private Function Repoint(db As DAO.Database, fromName As String, toName As String) As Long
    Dim idFrom As Long
    Dim idTo As Long
    idFrom = IdOfType(db, fromName)
    idTo = IdOfType(db, toName)
    If idFrom = 0 Then Exit Function
    If idTo = 0 Then
        ' The target does not exist yet - create it now, because
        ' repointing must happen before the old entries are deleted.
        db.Execute "INSERT INTO L_DEC_TYPE (Name, Name_VAL, Description) VALUES ('" & Esc(toName) & "','" & Esc(toName) & "','Created by patch v17a')", dbFailOnError
        idTo = IdOfType(db, toName)
        If idTo = 0 Then
            Debug.Print "  *** could not create " & toName & " - rows left untouched"
            Exit Function
        End If
    End If
    Dim k As Long
    k = DCount("*", "T_DECORATIONS", "ID_Dec_Type=" & idFrom)
    If k > 0 Then
        db.Execute "UPDATE T_DECORATIONS SET ID_Dec_Type=" & idTo & " WHERE ID_Dec_Type=" & idFrom, dbFailOnError
        Debug.Print "-> " & k & " row(s) repointed: " & fromName & " -> " & toName
    End If
    Repoint = k
End Function

' Clear the type on rows that need a human decision, leaving the
' row itself intact. A blank type is visible and fixable; a wrong
' one is neither.
Private Function ClearType(db As DAO.Database, nm As String) As Long
    Dim idFrom As Long
    idFrom = IdOfType(db, nm)
    If idFrom = 0 Then Exit Function
    Dim k As Long
    k = DCount("*", "T_DECORATIONS", "ID_Dec_Type=" & idFrom)
    If k > 0 Then
        db.Execute "UPDATE T_DECORATIONS SET ID_Dec_Type=Null WHERE ID_Dec_Type=" & idFrom, dbFailOnError
        Debug.Print "-> " & k & " row(s) left WITHOUT a type (were " & nm & "): manual review"
    End If
    ClearType = k
End Function

' ================================================================
'  Rebuild the rock art half of L_DEC_TYPE.
'  ONE AXIS: what motif is it. The v16 list asked three different
'  questions at once, which is why it could not be applied
'  consistently however carefully each entry was worded.
'  The architectural half is left exactly as it is.
' ================================================================
Private Sub RebuildRockArtTypes(db As DAO.Database)
    ' Delete only the RA entries nothing points at any more.
    Dim rs As DAO.Recordset
    Dim q As String
    q = "SELECT ID, Name FROM L_DEC_TYPE WHERE Name Like 'RA *'"
    Set rs = db.OpenRecordset(q, dbOpenSnapshot)
    Do While Not rs.EOF
        If DCount("*", "T_DECORATIONS", "ID_Dec_Type=" & rs!ID) = 0 Then
            db.Execute "DELETE FROM L_DEC_TYPE WHERE ID=" & rs!ID, dbFailOnError
        Else
            Debug.Print "  [kept] " & rs!Name & " still has rows pointing at it"
        End If
        rs.MoveNext
    Loop
    rs.Close

    ' Anthropomorphic keeps the note about the decapitation scene:
    ' the v17 delta retired that entry from the architectural list
    ' precisely because it belongs here.
    AddType db, "RA Anthropomorphic", "Antropomorf", "Rock art: anthropomorphic figure. Includes the decapitation scene: record the reading in Notes."
    AddType db, "RA Zoomorphic", "Zoomorf", "Rock art: zoomorphic figure."
    ' The two U values absorb what Outline_Geometry recorded apart.
    ' The distinction is not decorative: an ORTHOGONAL outline
    ' asserts a rectangular BUILT referent and is the evidence that
    ' a vanished structure stood there (schema 8bis.9); an organic
    ' one asserts nothing and may follow a natural recess.
    AddType db, "RA U-shape geometric", "Forma U geometrica", "Rock art: inverted-U band with straight runs and angular corners. Asserts a rectangular built referent: cross-check with T_LOST_ELEMENTS."
    AddType db, "RA U-shape organic", "Forma U organica", "Rock art: inverted-U band with curved or ill-defined outline. Asserts no built referent; may follow a natural recess."
    AddType db, "RA Geometric motif", "Motiu geometric", "Rock art: lines, bands or figures with recognisable regular organisation."
    ' Amorphous stain and Pigment traces are the pair that caused
    ' the confusion with the old Abstract. The boundary is whether
    ' the pigment has a readable EDGE, not how meaningful it looks.
    AddType db, "RA Amorphous stain", "Taca amorfa", "Rock art: defined pigment surface with a recognisable edge but no identifiable motif."
    AddType db, "RA Pigment traces", "Traces de pigment", "Rock art: scattered or degraded remains, too poor to say whether they formed a motif."
End Sub

Private Sub AddType(db As DAO.Database, nm As String, va As String, ds As String)
    If IdOfType(db, nm) > 0 Then
        db.Execute "UPDATE L_DEC_TYPE SET Name_VAL='" & Esc(va) & "', Description='" & Esc(ds) & "' WHERE Name='" & Esc(nm) & "'", dbFailOnError
        Debug.Print "-> " & nm & " updated"
        Exit Sub
    End If
    db.Execute "INSERT INTO L_DEC_TYPE (Name, Name_VAL, Description) VALUES ('" & Esc(nm) & "','" & Esc(va) & "','" & Esc(ds) & "')", dbFailOnError
    Debug.Print "-> " & nm & " added"
End Sub

' ================================================================
'  Rule 41 checked Outline_Geometry on a non-ROC row. With the
'  column gone the branch would fail the whole battery union, so
'  QRY_16e is rebuilt without it. THE OTHERS KEEP THEIR NUMBERS:
'  they are cited across four documents, and closing a gap in the
'  sequence would invalidate every citation to gain nothing.
'  QRY_22_V17a_Review is created at the same time - the rows this
'  patch deliberately refused to decide have to be findable.
' ================================================================
Private Function RetireRule41(db As DAO.Database) As Boolean
    On Error Resume Next
    db.QueryDefs.Delete "QRY_16_Validation_Check"
    db.QueryDefs.Delete "QRY_16e_Rules_40_46"
    db.QueryDefs.Delete "QRY_22_V17a_Review"
    Err.Clear
    On Error GoTo 0

    Dim q As String
    q = R40a()
    q = q & " UNION ALL " & R40b()
    q = q & " UNION ALL " & R42()
    q = q & " UNION ALL " & R43()
    q = q & " UNION ALL " & R44()
    q = q & " UNION ALL " & R45()
    q = q & " UNION ALL " & R46()
    q = q & ";"
    MkQ db, "QRY_16e_Rules_40_46", q

    q = "SELECT * FROM QRY_16a_Rules_1_11 "
    q = q & "UNION ALL SELECT * FROM QRY_16b_Rules_12_21 "
    q = q & "UNION ALL SELECT * FROM QRY_16c_Rules_22_31 "
    q = q & "UNION ALL SELECT * FROM QRY_16d_Rules_32_39 "
    q = q & "UNION ALL SELECT * FROM QRY_16e_Rules_40_46 "
    q = q & "ORDER BY Rule_No, Structure;"
    MkQ db, "QRY_16_Validation_Check", q

    q = "SELECT E.Code AS Structure, D.ID AS Row_ID, "
    q = q & "'Rock art row with no type' AS Review_Item, "
    q = q & "'Was Perimeter band: choose U-shape geometric or U-shape organic. The old outline value, if recorded, is in Notes' AS Reason, "
    q = q & "D.Span_Left, D.Span_Above, D.Span_Right, D.Span_Below, D.Notes "
    q = q & "FROM T_DECORATIONS AS D "
    q = q & "INNER JOIN T_STRUCTURES AS E ON D.ID_Structure=E.ID "
    q = q & "WHERE D.ID_Dec_Type Is Null "
    q = q & "ORDER BY E.Code, D.ID;"
    MkQ db, "QRY_22_V17a_Review", q

    RetireRule41 = True
End Function

Private Function R40a() As String
    Dim q As String
    q = "SELECT E.Code AS Structure, 40 AS Rule_No, "
    q = q & "'R40: rock art record carrying a structure-outline position' AS Rule_Violated, "
    q = q & "'A panel has no outline to follow: use Panel, or reclassify the record as a structure' AS Action "
    q = q & "FROM (((T_DECORATIONS AS D "
    q = q & "INNER JOIN T_STRUCTURES AS E ON D.ID_Structure=E.ID) "
    q = q & "INNER JOIN L_STRUCT_BODY AS B ON D.ID_Struct_Body=B.ID) "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID) "
    q = q & "WHERE T.Record_Class='Rock art panel' "
    q = q & "AND B.Code IN ('ROC-OVL','ROC-PER','ROC')"
    R40a = q
End Function

Private Function R40b() As String
    Dim q As String
    q = "SELECT E.Code AS Structure, 40 AS Rule_No, "
    q = q & "'R40: structure record carrying the Panel position' AS Rule_Violated, "
    q = q & "'Panel belongs to isolated PR records: give the row a contact position' AS Action "
    q = q & "FROM (((T_DECORATIONS AS D "
    q = q & "INNER JOIN T_STRUCTURES AS E ON D.ID_Structure=E.ID) "
    q = q & "INNER JOIN L_STRUCT_BODY AS B ON D.ID_Struct_Body=B.ID) "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID) "
    q = q & "WHERE T.Record_Class<>'Rock art panel' "
    q = q & "AND B.Code='ROC-PAN'"
    R40b = q
End Function

Private Function R42() As String
    Dim q As String
    q = "SELECT E.Code AS Structure, 42 AS Rule_No, "
    q = q & "'R42: band motif on the facade but Relief_Frieze is 0' AS Rule_Violated, "
    q = q & "'Raise M, or reconsider whether the motif is a frieze band' AS Action "
    q = q & "FROM (((T_DECORATIONS AS D "
    q = q & "INNER JOIN T_STRUCTURES AS E ON D.ID_Structure=E.ID) "
    q = q & "INNER JOIN L_STRUCT_BODY AS B ON D.ID_Struct_Body=B.ID) "
    q = q & "INNER JOIN L_DEC_TYPE AS DT ON D.ID_Dec_Type=DT.ID) "
    q = q & "WHERE B.Code IN ('OVL','FFL') AND E.Relief_Frieze=0 "
    q = q & "AND DT.Name IN ('Zigzag','Stepped motif','Square niche',"
    q = q & "'T-shaped niche','T-shaped niche inv.',"
    q = q & "'L-shaped niche','L-shaped niche inv.')"
    R42 = q
End Function

Private Function R43() As String
    Dim q As String
    q = "SELECT E.Code AS Structure, 43 AS Rule_No, "
    q = q & "'R43: materials declared present but all seven types are 0' AS Rule_Violated, "
    q = q & "'Record at least one type, set the unidentified ones to 9, or lower the aggregate' AS Action "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Cultural_Materials_Present=1 "
    q = q & "AND E.Mat_Textiles=0 AND E.Mat_Wood=0 AND E.Mat_VegFiber=0 "
    q = q & "AND E.Mat_Ceramics=0 AND E.Mat_Fauna=0 AND E.Mat_DeerAntler=0 "
    q = q & "AND E.Mat_Other=0"
    R43 = q
End Function

Private Function R44() As String
    Dim q As String
    q = "SELECT E.Code AS Structure, 44 AS Rule_No, "
    q = q & "'R44: condition of movable remains recorded where none are present' AS Rule_Violated, "
    q = q & "'Legacy of the v16 layout, where the gate and the field sat on different tabs' AS Action "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Cultural_Materials_Present=0 "
    q = q & "AND E.ID_Material_Status Is Not Null"
    R44 = q
End Function

Private Function R45() As String
    Dim q As String
    q = "SELECT E.Code AS Structure, 45 AS Rule_No, "
    q = q & "'R45: fabric declared multiple BY BODIES on a single-bodied structure' AS Rule_Violated, "
    q = q & "'Either the body count is wrong, or the plurality is within the one body' AS Action "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Fabric='Between bodies only' "
    q = q & "AND (Nz(E.N_Basal_Bodies,0)+Nz(E.N_Chamber_Bodies,0))<2"
    R45 = q
End Function

Private Function R46() As String
    Dim q As String
    q = "SELECT E.Code AS Structure, 46 AS Rule_No, "
    q = q & "'R46: two or more construction phases claimed with no phase evidence' AS Rule_Violated, "
    q = q & "'A phase is an interpretation: record what sustains it, or lower the count' AS Action "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Construction_Phases>=2 "
    q = q & "AND (E.Phase_Evidence Is Null Or E.Phase_Evidence='ND')"
    R46 = q
End Function

' ================================================================
'  HELPERS
'  MkQ prints on failure. The v16 bug that hid a missing battery
'  union for a whole version was a creation failure degraded to a
'  message nobody read: an error deliberately made non-fatal must
'  still leave a trace where someone looks.
' ================================================================
Private Sub MkQ(db As DAO.Database, nm As String, sql As String)
    On Error Resume Next
    db.QueryDefs.Delete nm
    Err.Clear
    Dim qd As DAO.QueryDef
    Set qd = db.CreateQueryDef(nm, sql)
    If Err.Number <> 0 Then
        Debug.Print "  *** query " & nm & " NOT created: " & Err.Description
        Err.Clear
    Else
        Debug.Print "-> query " & nm & " created"
    End If
    On Error GoTo 0
End Sub

Private Function IdOfType(db As DAO.Database, nm As String) As Long
    Dim v As Variant
    v = DLookup("ID", "L_DEC_TYPE", "Name='" & Esc(nm) & "'")
    If Not IsNull(v) Then IdOfType = CLng(v)
End Function

Private Function Esc(s As String) As String
    Esc = Replace(s, "'", "''")
End Function

Private Function TableExists(db As DAO.Database, n As String) As Boolean
    Dim td As DAO.TableDef
    For Each td In db.TableDefs
        If StrComp(td.Name, n, vbTextCompare) = 0 Then
            TableExists = True
            Exit Function
        End If
    Next td
End Function

Private Function ColExists(db As DAO.Database, tbl As String, col As String) As Boolean
    Dim f As DAO.Field
    For Each f In db.TableDefs(tbl).Fields
        If StrComp(f.Name, col, vbTextCompare) = 0 Then
            ColExists = True
            Exit Function
        End If
    Next f
End Function
