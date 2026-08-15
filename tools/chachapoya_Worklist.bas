Option Compare Database
Option Explicit

' ================================================================
'  CHACHAPOYA WORKLIST NAVIGATOR (add-on for v23)
'  Author: Esteve Ribera Torro | TFM Arqueologia UA
'
'  WHAT THIS IS
'  One query and one form that turn the review machinery into a
'  clickable to-do list:
'
'    QRY_28_Worklist - the UNION of the four live sources:
'      the v19 review, the v20 review, the v21 review and
'      the rule battery.
'      IT IS THE TO-DO LIST ITSELF: no stored table, no done
'      checkbox - when you fix a record its rows simply vanish
'      on the next requery, so the list can never disagree with
'      the database (the usual duplication principle).
'
'    F_WORKLIST - a read-only datasheet over that query.
'      DOUBLE-CLICK any row -> F_STRUCTURES opens filtered on
'      that structure AND jumps to the tab the task lives on
'      (crown format -> 11.Sist, outline bands -> 4.Dec,
'      retypes -> 12.Extra, MEN pass and masonry judgement
'      -> 2.Arq, mass-top platforms -> 11.Sist, and so on).
'      The REFRESCA button (or F5) requeries: fixed rows
'      disappear, the count in the caption updates.
'
'  HOW TO USE
'    1. Import this module into the v23 database.
'    2. Run BuildWorklist() once. Re-running it is harmless.
'    3. Open F_WORKLIST and work top-down. Suggested order:
'       'Migracio v21' first (the mass-top reverts unlock
'       rule 62; the masonry judgements unlock rule 17),
'       then what remains of 'Migracio v19/v20', then leave
'       'Bateria' empty as the final check. KNOWN RESIDUE
'       (delta v21, bloc 6): the v19 fabric-reveal and the
'       v18 cornice-as-sill branches keep listing CONFIRMED
'       zeros - once the pass is documented, those rows are
'       history, not work.
'
'  Requires the v23 package applied (QRY_24, QRY_25, QRY_29
'  and QRY_16_Validation_Check must exist). v23 restricts
'  the fabric-reveal branch of QRY_24 to portals Present*:
'  the 53 inapplicable rows drop off the list and the
'  bloc-6 'known residue' shrinks to the 8 true candidates
'  (plus the cornice-as-sill zeros, unchanged). v22 adds no new
'  review query: the metric-availability judgement stays off
'  the worklist on purpose (delta v22, bloc B), the panel
'  masonry 0s are derived by the patch, and the naturals'
'  masonry rows of QRY_29 are now answerable on 2.Arq. Same VBA-project-access
'  requirement as BuildForm for the double-click code.
' ================================================================

Private mCode As String

Sub BuildWorklist()
    Dim db As DAO.Database
    Set db = CurrentDb()

    ' ------------------------------------------------------------
    ' 1. The union query. Column names are harmonised: Font /
    '    Structure / Tasca / Detall. The battery rows carry the
    '    rule number inside Tasca so nothing is lost.
    ' ------------------------------------------------------------
    Dim q As String
    q = "SELECT 'Migracio v19' AS Font, Structure, Review_Item AS Tasca, Reason AS Detall FROM QRY_24_V19_Review "
    q = q & "UNION ALL SELECT 'Migracio v20', Structure, Review_Item, Reason FROM QRY_25_V20_Review "
    q = q & "UNION ALL SELECT 'Migracio v21', Structure, Review_Item, Reason FROM QRY_29_V21_Review "
    q = q & "UNION ALL SELECT 'Bateria', Structure, 'Regla ' & Rule_No & ': ' & Rule_Violated, Action FROM QRY_16_Validation_Check "
    q = q & "ORDER BY Font, Structure;"

    On Error Resume Next
    db.QueryDefs.Delete "QRY_28_Worklist"
    On Error GoTo 0
    db.CreateQueryDef "QRY_28_Worklist", q
    Debug.Print "[OK] QRY_28_Worklist"

    ' ------------------------------------------------------------
    ' 2. The form.
    ' ------------------------------------------------------------
    Const WFRM = "F_WORKLIST"
    On Error Resume Next: DoCmd.DeleteObject acForm, WFRM: On Error GoTo 0

    Dim f As Form: Set f = CreateForm()
    Dim tmp As String: tmp = f.Name
    f.RecordSource = "QRY_28_Worklist"
    f.DefaultView = 2
    f.AllowEdits = False: f.AllowAdditions = False: f.AllowDeletions = False
    f.Caption = "Llista de treball - doble clic obri el registre"

    Dim T As Long: T = 50
    Dim c As Control, lb As Control

    Set c = CreateControl(tmp, acTextBox, acDetail, "", "", 40, T, 1500, 315)
    c.ControlSource = "Font"
    On Error Resume Next: c.Name = "Font_Col": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, c.Name, "", 40, T, 1500, 260)
    lb.Caption = "Font"

    Set c = CreateControl(tmp, acTextBox, acDetail, "", "", 1600, T, 1500, 315)
    c.ControlSource = "Structure"
    On Error Resume Next: c.Name = "Structure": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, c.Name, "", 1600, T, 1500, 260)
    lb.Caption = "Estructura"

    Set c = CreateControl(tmp, acTextBox, acDetail, "", "", 3160, T, 4200, 315)
    c.ControlSource = "Tasca"
    On Error Resume Next: c.Name = "Tasca": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, c.Name, "", 3160, T, 4200, 260)
    lb.Caption = "Tasca"

    Set c = CreateControl(tmp, acTextBox, acDetail, "", "", 7420, T, 5400, 315)
    c.ControlSource = "Detall"
    On Error Resume Next: c.Name = "Detall": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, c.Name, "", 7420, T, 5400, 260)
    lb.Caption = "Que fer"

    ' The refresh button lives in the header (visible in form
    ' view; in datasheet view use F5 / Maj+F9, the caption of the
    ' form reminds nobody - the manual note does).
    Dim bR As Control
    Set bR = CreateControl(tmp, acCommandButton, acDetail, "", "", 12900, T, 1400, 380)
    bR.Name = "cmdRefresh"
    bR.Caption = "Refresca"
    bR.OnClick = "[Event Procedure]"

    ' ------------------------------------------------------------
    ' 3. The module: double-click navigation + refresh. The tab
    '    jump reads the Tasca text - keyword map, with 11.Sist as
    '    the sensible fallback (most battery rules live there).
    ' ------------------------------------------------------------
    mCode = ""
    LG "Option Compare Database"
    LG "Option Explicit"
    LG ""
    LG "Private Sub GoToRecord()"
    LG "    If IsNull(Me!Structure) Then Exit Sub"
    LG "    DoCmd.OpenForm ""F_STRUCTURES"", acNormal, , ""Code='"" & Me!Structure & ""'"""
    LG "    Dim tg As String"
    LG "    tg = TabFor(Nz(Me!Tasca, """"))"
    LG "    On Error Resume Next"
    LG "    Forms(""F_STRUCTURES"")!tabMain.Pages(tg).SetFocus"
    LG "    On Error GoTo 0"
    LG "End Sub"
    LG ""
    LG "Private Function TabFor(t As String) As String"
    LG "    ' Keyword -> tab. Catalan and English stems, because"
    LG "    ' the three sources mix languages in Tasca."
    LG "    TabFor = ""pgSys"""
    LG "    If InStr(t, ""Outline"") > 0 Or InStr(t, ""banda"") > 0 Then TabFor = ""pgDec"""
    LG "    If InStr(t, ""12.Extra"") > 0 Or InStr(t, ""retype"") > 0 Then TabFor = ""pgExtra"""
    LG "    If InStr(t, ""2.Arq"") > 0 Or InStr(t, ""MEN"") > 0 Then TabFor = ""pgArq"""
    LG "    If InStr(t, ""Masonry"") > 0 Or InStr(t, ""maconeria"") > 0 Then TabFor = ""pgArq"""
    LG "    If InStr(t, ""metric"") > 0 Or InStr(t, ""Metrics"") > 0 Then TabFor = ""pgMetr"""
    LG "    If InStr(t, ""C14"") > 0 Or InStr(t, ""centur"") > 0 Or InStr(t, ""segle"") > 0 Then TabFor = ""pgCron"""
    LG "    If InStr(t, ""wall"") > 0 Or InStr(t, ""murs"") > 0 Then TabFor = ""pgArq"""
    LG "    If InStr(t, ""EA_Number"") > 0 Or InStr(t, ""EA number"") > 0 Or InStr(t, ""Subsector"") > 0 Then TabFor = ""pgId"""
    LG "    If InStr(t, ""portal-adjacent"") > 0 Or InStr(t, ""phase"") > 0 Or InStr(t, ""fase"") > 0 Then TabFor = ""pgArq"""
    LG "    If InStr(t, ""sector"") > 0 Then TabFor = ""pgId"""
    LG "End Function"
    LG ""
    LG "Private Sub Structure_DblClick(Cancel As Integer)"
    LG "    GoToRecord"
    LG "End Sub"
    LG ""
    LG "Private Sub Tasca_DblClick(Cancel As Integer)"
    LG "    GoToRecord"
    LG "End Sub"
    LG ""
    LG "Private Sub Detall_DblClick(Cancel As Integer)"
    LG "    GoToRecord"
    LG "End Sub"
    LG ""
    LG "Private Sub Font_Col_DblClick(Cancel As Integer)"
    LG "    GoToRecord"
    LG "End Sub"
    LG ""
    LG "Private Sub cmdRefresh_Click()"
    LG "    Me.Requery"
    LG "End Sub"
    ReplaceModule tmp

    DoCmd.Save acForm, tmp
    DoCmd.Close acForm, tmp
    DoCmd.Rename WFRM, acForm, tmp
    Debug.Print "[OK] F_WORKLIST"

    Set db = Nothing
    MsgBox "LLISTA DE TREBALL CREADA" & vbCrLf & vbCrLf & "Obri F_WORKLIST i treballa de dalt a baix:" & vbCrLf & vbCrLf & "  1. Filtra Font = 'Migracio v21' (reverts + judicis)" & vbCrLf & "  2. Despres el que quede de v19/v20" & vbCrLf & "  3. 'Bateria' buida = corpus coherent" & vbCrLf & "  (Residu conegut: els 0 confirmats de brancal" & vbCrLf & "  i cornisa-llindar; delta v21, bloc 6)" & vbCrLf & vbCrLf & "DOBLE CLIC en qualsevol fila obri el registre" & vbCrLf & "a la pestanya que toca. El boto REFRESCA (o F5)" & vbCrLf & "fa desapareixer les files ja resoltes.", vbInformation, "F_WORKLIST"
End Sub

Private Sub LG(s As String)
    mCode = mCode & s & vbCrLf
End Sub

Private Sub ReplaceModule(frmName As String)
    Dim f As Form
    Set f = Forms(frmName)
    f.HasModule = True
    With f.Module
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        .InsertLines 1, mCode
    End With
End Sub
