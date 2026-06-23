Attribute VB_Name = "addExtraFields"
Option Compare Database
Option Explicit

' ================================================================
'  ADD EXTRA FIELDS + T_ARCH_FEATURES + FORM TAB 11
'  Run Sub AddExtraFields() -> F5
' ================================================================

Const MT  As Long = 550
Const RG  As Long = 420
Const c1  As Long = 120
Const c2  As Long = 4880
Const LW  As Long = 1900
Const CW  As Long = 2500
Const CH  As Long = 315
Const lh  As Long = 270

Sub addExtraFields()
    Dim db As DAO.Database
    Set db = CurrentDb()
    AddFields db
    CreateArchFeaturesTable db
    Set db = Nothing
    UpdateForm
    Dim msg As String
    msg = "Update complete!" & vbCrLf & vbCrLf
    msg = msg & "  Embedded_Base_Beams added" & vbCrLf
    msg = msg & "  Tie_Walls added" & vbCrLf
    msg = msg & "  Corner_Quoins added" & vbCrLf
    msg = msg & "  T_ARCH_FEATURES created" & vbCrLf
    msg = msg & "  Tab 11 added to F_STRUCTURES"
    MsgBox msg, vbInformation, "Done!"
End Sub

' -- 1. NEW BOOLEAN FIELDS ----------------------------------------
Sub AddFields(db As DAO.Database)
    On Error Resume Next
    db.Execute "ALTER TABLE T_STRUCTURES ADD COLUMN Embedded_Base_Beams YESNO", dbFailOnError
    db.Execute "ALTER TABLE T_STRUCTURES ADD COLUMN Tie_Walls YESNO", dbFailOnError
    db.Execute "ALTER TABLE T_STRUCTURES ADD COLUMN Corner_Quoins YESNO", dbFailOnError
    On Error GoTo 0
    Debug.Print "[OK] 3 fields added"
End Sub

' -- 2. T_ARCH_FEATURES (flexible, future-proof) ------------------
Sub CreateArchFeaturesTable(db As DAO.Database)
    On Error Resume Next
    db.Relations.Delete "REL_STR_AFEAT"
    db.Execute "DROP TABLE T_ARCH_FEATURES", dbFailOnError
    On Error GoTo 0

    Dim sql As String
    sql = "CREATE TABLE T_ARCH_FEATURES ("
    sql = sql & "ID COUNTER CONSTRAINT PK_AF PRIMARY KEY,"
    sql = sql & "ID_Structure LONG NOT NULL,"
    sql = sql & "Feature_Code TEXT(40),"
    sql = sql & "Present YESNO,"
    sql = sql & "Feature_Count INTEGER,"
    sql = sql & "Material TEXT(20),"
    sql = sql & "Notes TEXT(255))"
    db.Execute sql, dbFailOnError

    ' Relationship
    Dim rel As DAO.Relation
    Dim fld As DAO.Field
    On Error GoTo ErrRel
    Set rel = db.CreateRelation("REL_STR_AFEAT", "T_STRUCTURES", "T_ARCH_FEATURES", dbRelationUpdateCascade Or dbRelationDeleteCascade)
    Set fld = rel.CreateField("ID")
    fld.ForeignName = "ID_Structure"
    rel.Fields.Append fld
    db.Relations.Append rel
    db.Relations.Refresh
    Debug.Print "[OK] T_ARCH_FEATURES created"
    Exit Sub
ErrRel:
    Debug.Print "  Warning REL: " & Err.Description
End Sub

' -- 3. FORM: ADD TAB 11 + CONTROLS + SUBFORM ---------------------
Sub UpdateForm()
    Const frm = "F_STRUCTURES"
    DoCmd.OpenForm frm, acDesign

    ' New page
    Dim pg As Control
    Set pg = CreateControl(frm, acPage, acDetail, "tabMain")
    pg.Name = "pgExtra"
    pg.Caption = "11.Extra"

    ' Section: confirmed elements
    SH frm, "pgExtra", "CONFIRMED CONSTRUCTIVE ELEMENTS", 0

    PCB frm, "pgExtra", "Embedded base beams:", "Embedded_Base_Beams", 1, 1
    PCB frm, "pgExtra", "Tie walls (perp. to cliff):", "Tie_Walls", 1, 2
    PCB frm, "pgExtra", "Corner quoins:", "Corner_Quoins", 2, 1

    ' Section: flexible features
    SH frm, "pgExtra", "CUSTOM FEATURES (T_ARCH_FEATURES) - add any element not covered above", 4

    ' Create and embed subform
    CreateFeaturesSubform
    Dim sf As Control
    Set sf = CreateControl(frm, acSubform, acDetail, "pgExtra", "", c1, MT + 5 * RG, 9000, 2600)
    sf.SourceObject = "F_ARCH_FEATURES"
    sf.LinkMasterFields = "ID"
    sf.LinkChildFields = "ID_Structure"
    On Error Resume Next: sf.Name = "sf_arch_features": On Error GoTo 0

    DoCmd.Save acForm, frm
    DoCmd.Close acForm, frm
    Debug.Print "[OK] Tab 11 added"
End Sub

Sub CreateFeaturesSubform()
    Const SFRM = "F_ARCH_FEATURES"
    On Error Resume Next
    DoCmd.DeleteObject acForm, SFRM
    On Error GoTo 0

    Dim f As Form
    Set f = CreateForm()
    Dim tmp As String
    tmp = f.Name

    f.RecordSource = "T_ARCH_FEATURES"
    f.DefaultView = 2    ' Continuous
    f.ScrollBars = 2
    f.NavigationButtons = False
    f.RecordSelectors = True
    f.Width = 9400
    f.Section(acDetail).Height = 400

    Dim L As Long: Dim T As Long
    T = 50

    ' Feature_Code
    L = 40
    Dim l1 As Control
    Set l1 = CreateControl(tmp, acLabel, acDetail, "", "", L, T + 15, 1000, 260)
    l1.Caption = "Feature:"
    l1.BackStyle = 0: l1.BorderStyle = 0
    Dim c1 As Control
    Set c1 = CreateControl(tmp, acComboBox, acDetail, "", "", L + 1060, T, 2400, 315)
    c1.ControlSource = "Feature_Code"
    c1.RowSourceType = "Value List"
    c1.RowSource = """Embedded base beams"";""Tie wall"";""Corner quoins"";""Unusual masonry bond"";""Pigment trace"";""Iron staining"";""Other"""
    c1.LimitToList = False
    On Error Resume Next: c1.Name = "Feature_Code": On Error GoTo 0

    ' Present checkbox
    L = 3560
    Dim l2 As Control
    Set l2 = CreateControl(tmp, acLabel, acDetail, "", "", L, T + 15, 800, 260)
    l2.Caption = "Present:"
    l2.BackStyle = 0: l2.BorderStyle = 0
    Dim c2 As Control
    Set c2 = CreateControl(tmp, acCheckBox, acDetail, "", "", L + 860, T + 20, 300, 300)
    c2.ControlSource = "Present"
    On Error Resume Next: c2.Name = "Present": On Error GoTo 0

    ' Count
    L = 4560
    Dim l3 As Control
    Set l3 = CreateControl(tmp, acLabel, acDetail, "", "", L, T + 15, 600, 260)
    l3.Caption = "Count:"
    l3.BackStyle = 0: l3.BorderStyle = 0
    Dim c3 As Control
    Set c3 = CreateControl(tmp, acTextBox, acDetail, "", "", L + 660, T, 700, 315)
    c3.ControlSource = "Feature_Count"
    On Error Resume Next: c3.Name = "Feature_Count": On Error GoTo 0

    ' Material
    L = 6120
    Dim l4 As Control
    Set l4 = CreateControl(tmp, acLabel, acDetail, "", "", L, T + 15, 700, 260)
    l4.Caption = "Material:"
    l4.BackStyle = 0: l4.BorderStyle = 0
    Dim c4 As Control
    Set c4 = CreateControl(tmp, acComboBox, acDetail, "", "", L + 760, T, 1200, 315)
    c4.ControlSource = "Material"
    c4.RowSourceType = "Value List"
    c4.RowSource = "Stone;Timber;Mixed;ND"
    c4.LimitToList = False
    On Error Resume Next: c4.Name = "Material": On Error GoTo 0

    ' Notes
    L = 8080
    Dim l5 As Control
    Set l5 = CreateControl(tmp, acLabel, acDetail, "", "", L, T + 15, 600, 260)
    l5.Caption = "Notes:"
    l5.BackStyle = 0: l5.BorderStyle = 0
    Dim c5 As Control
    Set c5 = CreateControl(tmp, acTextBox, acDetail, "", "", L + 660, T, 1500, 315)
    c5.ControlSource = "Notes"
    On Error Resume Next: c5.Name = "Notes": On Error GoTo 0

    DoCmd.Save acForm, tmp
    DoCmd.Close acForm, tmp
    DoCmd.Rename SFRM, acForm, tmp
    Debug.Print "[OK] F_ARCH_FEATURES subform created"
End Sub

' -- HELPERS ------------------------------------------------------
Sub SH(frm As String, pg As String, txt As String, row As Integer)
    Dim T As Long: T = MT + row * RG - 16
    Dim lh As Control
    Set lh = CreateControl(frm, acLabel, acDetail, pg, "", c1, T, 9000, 260)
    lh.Caption = "  " & UCase(txt)
    lh.BackStyle = 1
    lh.BackColor = RGB(214, 228, 247)
    lh.BorderStyle = 0
    lh.ForeColor = RGB(26, 60, 107)
    lh.FontBold = True
    lh.FontSize = 8
End Sub

Sub PCB(frm As String, pg As String, lbl As String, src As String, row As Integer, col As Integer)
    Dim L As Long: Dim T As Long
    If col = 1 Then L = c1 Else L = c2
    T = MT + row * RG
    Dim lc As Control
    Set lc = CreateControl(frm, acLabel, acDetail, pg, "", L, T + 22, LW, lh)
    lc.Caption = lbl: lc.BackStyle = 0: lc.BorderStyle = 0
    lc.TextAlign = 3: lc.ForeColor = RGB(55, 55, 80)
    Dim cc As Control
    Set cc = CreateControl(frm, acCheckBox, acDetail, pg, "", L + LW + 80, T + 20, 300, CH)
    cc.ControlSource = src
    On Error Resume Next: cc.Name = src: On Error GoTo 0
End Sub
