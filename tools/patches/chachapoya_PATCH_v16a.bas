Option Compare Database
Option Explicit

' ================================================================
'  CHACHAPOYA - PATCH v16 -> v16a
'  Run Sub PatchV16a() ON AN EXISTING v16 DATABASE that already has
'  records in it. If you have not entered anything yet, rebuild
'  from scratch with the corrected chachapoya_DB_v16.bas instead -
'  a rebuild is always cleaner than a patch.
'
'  WHAT WAS WRONG. L_LOST_EVIDENCE was created without a Name_VAL
'  column, while the lost-elements subform selects exactly that
'  column for its combo. The RowSource query failed, so the
'  EVIDENCE TYPE DROPDOWN CAME UP EMPTY - and with no evidence type
'  there is no way to open the row that value 3 requires, which
'  means "attested lost" could not be recorded at all.
'
'  WHY IT WENT UNNOTICED. The nine VL calls that write the
'  Valencian labels tolerate errors by design, so nine failed
'  updates in a row produced no message. Same failure family as the
'  QRY_16c naming bug: an error deliberately made non-fatal, with
'  nowhere for anyone to see it. VL now prints on failure.
'
'  This patch touches ONE lookup table. It adds no field to any
'  data table and changes no stored value.
' ================================================================

Sub PatchV16a()
    Dim db As DAO.Database
    Set db = CurrentDb()

    Debug.Print "=== PATCH v16 -> v16a ==="

    If Not TableExists(db, "L_LOST_EVIDENCE") Then
        MsgBox "L_LOST_EVIDENCE does not exist. This is not a v16 database.", vbCritical, "Patch aborted"
        Exit Sub
    End If

    Dim added As Boolean
    If ColExists(db, "L_LOST_EVIDENCE", "Name_VAL") Then
        Debug.Print "-> Name_VAL already present, no schema change needed"
    Else
        db.Execute "ALTER TABLE L_LOST_EVIDENCE ADD COLUMN Name_VAL TEXT(60)", dbFailOnError
        db.TableDefs.Refresh
        added = True
        Debug.Print "-> Name_VAL added to L_LOST_EVIDENCE"
    End If

    ' Labels are matched on the English Name, which is the stored
    ' value and has not changed. Re-running this is harmless.
    VL2 db, "Negative socket / impression", "Encaix negatiu / empremta"
    VL2 db, "Beam hole", "Forat de biga"
    VL2 db, "Break scar", "Cicatriu de despreniment"
    VL2 db, "Detached fragment in situ", "Fragment despres in situ"
    VL2 db, "Mortar imprint", "Empremta de morter"
    VL2 db, "Corbels into void", "Mensules al buit"
    VL2 db, "Pigment on bedrock", "Pigment sobre penya"
    VL2 db, "Truncated walls", "Murs truncats"
    VL2 db, "Other (see Notes)", "Altres (veure notes)"

    ' Anything left untranslated falls back to the English term, so
    ' a missing entry shows the term instead of a blank row.
    db.Execute "UPDATE L_LOST_EVIDENCE SET Name_VAL=Name WHERE Name_VAL Is Null", dbFailOnError

    Dim n As Long
    Dim blank As Long
    n = DCount("*", "L_LOST_EVIDENCE")
    blank = DCount("*", "L_LOST_EVIDENCE", "Name_VAL Is Null")

    Set db = Nothing

    Dim msg As String
    msg = "PATCH v16a APPLIED" & vbCrLf & vbCrLf
    If added Then
        msg = msg & "  Name_VAL added to L_LOST_EVIDENCE" & vbCrLf
    Else
        msg = msg & "  Name_VAL was already there" & vbCrLf
    End If
    msg = msg & "  Rows in the lookup: " & n & " (expected 9)" & vbCrLf
    msg = msg & "  Rows still without a label: " & blank & vbCrLf & vbCrLf
    msg = msg & "Now REOPEN F_STRUCTURES. Tab 12.Extra, lost" & vbCrLf
    msg = msg & "elements: the evidence type dropdown should" & vbCrLf
    msg = msg & "list nine values in Valencian." & vbCrLf & vbCrLf
    msg = msg & "If any element was left at a value other than 3" & vbCrLf
    msg = msg & "because the row could not be opened, this is the" & vbCrLf
    msg = msg & "moment to revisit it: rule 2 of QRY_16 lists the" & vbCrLf
    msg = msg & "3s that have no evidence row, but not the 3s that" & vbCrLf
    msg = msg & "were never entered."
    MsgBox msg, vbInformation, "Patch v16a"
    Debug.Print "=== PATCH FINISHED: " & n & " rows, " & blank & " unlabelled ==="
End Sub

Private Sub VL2(db As DAO.Database, en As String, va As String)
    On Error Resume Next
    db.Execute "UPDATE L_LOST_EVIDENCE SET Name_VAL='" & Replace(va, "'", "''") & "' WHERE Name='" & Replace(en, "'", "''") & "'", dbFailOnError
    If Err.Number <> 0 Then
        Debug.Print "  *** label failed for " & en & ": " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Sub

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
