Option Compare Database
Option Explicit

' ================================================================
'  ADD SYSTEMS TAB + DECORATION SUBFORM TO F_STRUCTURES
'  Run Sub AddSystemsTab() -> F5
'  Then run Sub AddDecSubform() -> F5
' ================================================================

' Layout constants (same as main form)
Const MT  As Long = 550
Const RG  As Long = 420
Const C1  As Long = 120
Const C2  As Long = 4880
Const LW  As Long = 1900
Const CW  As Long = 2500
Const CH  As Long = 315
Const LH  As Long = 270

' -- ENTRY POINT 1: Systems tab -----------------------------------
Sub AddSystemsTab()
    Const FRM = "F_STRUCTURES"
    DoCmd.OpenForm FRM, acDesign

    ' Add page 10 to the existing tab control
    Dim pg As Control
    Set pg = CreateControl(FRM, acPage, acDetail, "tabMain")
    pg.Name = "pgSys"
    pg.Caption = "10.Systems"

    ' -- SECTION: LEVEL 0 - BASE ---------------------------------
    SH FRM, "pgSys", "LEVEL 0 - BASE SYSTEM", 0

    PCB FRM, "pgSys", "Base level (I):",           "Base_Level",        1, 1
    PCB FRM, "pgSys", "Decorative socle (J):",     "Decorative_Socle",  1, 2

    PCB FRM, "pgSys", "Corbelled access platform (M):", "Corbelled_Platform", 2, 1

    SH  FRM, "pgSys", "  Platform system components (if M = Yes)", 3
    PCB FRM, "pgSys", "Timber corbels (K):",        "Timber_Brackets",   4, 1
    PCT FRM, "pgSys", "Corbel count:",               "Timber_Bracket_Count", 4, 2
    PCB FRM, "pgSys", "Transverse beams (L):",      "Transverse_Beams",  5, 1
    PCB FRM, "pgSys", "Corbelled courses (N):",     "Corbelled_Courses", 5, 2
    PCV FRM, "pgSys", "Corbel material:",            "Corbel_Material",   6, 1, "Timber;Stone;Mixed;ND"

    ' -- SECTION: ACCESS OPENING ----------------------------------
    SH  FRM, "pgSys", "ACCESS OPENING SYSTEM", 7
    PCB FRM, "pgSys", "Access opening (O):",        "Access_Opening",    8, 1
    PCB FRM, "pgSys", "Sill / threshold (P):",      "Sill",              8, 2
    PCB FRM, "pgSys", "Recessed portal:",           "Recessed_Portal",   9, 1

    ' -- SECTION: FACADE & UPPER ZONE ----------------------------
    SH  FRM, "pgSys", "FACADE & UPPER ZONE", 10
    PCB FRM, "pgSys", "Structural pilasters (C):",  "Structural_Pilasters", 11, 1
    PCV FRM, "pgSys", "Interlevel cornice material:","Cornice_Material",  11, 2, "Stone slabs;Wooden beams;Mixed;ND"
    PCB FRM, "pgSys", "Eave / roof overhang (G):",  "Eave",              12, 1
    PCB FRM, "pgSys", "Eave-supporting beam (F):",  "Eave_Beam",         12, 2
    PCB FRM, "pgSys", "Upper crown / coping (H):",  "Upper_Crown",       13, 1

    DoCmd.Save acForm, FRM
    DoCmd.Close acForm, FRM
    MsgBox "Tab '10.Systems' added to F_STRUCTURES!", vbInformation, "Done!"
End Sub

' -- ENTRY POINT 2: Decoration subform ---------------------------
Sub AddDecSubform()
    Const FRM = "F_STRUCTURES"

    ' Create F_DECORATIONS as a subform
    CreateDecForm

    ' Open F_STRUCTURES in design and add the subform to pgDec
    DoCmd.OpenForm FRM, acDesign

    Dim L As Long: Dim T As Long
    L = C1
    T = MT + 9 * RG     ' place below existing decoration checkboxes

    Dim sf As Control
    Set sf = CreateControl(FRM, acSubform, acDetail, "pgDec", "", L, T, 8800, 1800)
    sf.SourceObject = "F_DECORATIONS"
    sf.LinkMasterFields = "ID"
    sf.LinkChildFields = "ID_Structure"

    ' Section header above subform
    Dim lh As Control
    Set lh = CreateControl(FRM, acLabel, acDetail, "pgDec", "", C1, T - 340, 8800, 260)
    lh.Caption = "  DETAILED DECORATION RECORDS (T_DECORATIONS)"
    lh.BackStyle = 1
    lh.BackColor = RGB(214, 228, 247)
    lh.BorderStyle = 0
    lh.ForeColor = RGB(26, 60, 107)
    lh.FontBold = True
    lh.FontSize = 8

    DoCmd.Save acForm, FRM
    DoCmd.Close acForm, FRM
    MsgBox "Decoration subform added to tab 4 (Dec.)!", vbInformation, "Done!"
End Sub

Sub CreateDecForm()
    Const SFRM = "F_DECORATIONS"
    On Error Resume Next
    DoCmd.DeleteObject acForm, SFRM
    On Error GoTo 0

    Dim f As Form
    Set f = CreateForm()
    Dim tmpName As String
    tmpName = f.Name
    f.RecordSource = "T_DECORATIONS"
    f.DefaultView = 2     ' Continuous forms
    f.ScrollBars = 2      ' Vertical only
    f.NavigationButtons = False
    f.RecordSelectors = False
    f.Width = 9000
    f.Section(acDetail).Height = 380

    ' Fields (compact horizontal layout)
    Dim L As Long: L = 40
    Dim T As Long: T = 40

    ' Struct Body combo
    Dim lbl As Control
    Set lbl = CreateControl(tmpName, acLabel, acDetail, "", "", L, T + 15, 1400, 260)
    lbl.Caption = "Location:"
    lbl.BackStyle = 0: lbl.BorderStyle = 0

    Dim c1 As Control
    Set c1 = CreateControl(tmpName, acComboBox, acDetail, "", "", L + 1460, T, 2200, 315)
    c1.ControlSource = "ID_Struct_Body"
    c1.RowSourceType = "Table/Query"
    c1.RowSource = "SELECT ID, Name FROM L_STRUCT_BODY ORDER BY Level, Name"
    c1.BoundColumn = 1: c1.ColumnCount = 2
    c1.ColumnWidths = "0cm;4cm": c1.LimitToList = True
    On Error Resume Next: c1.Name = "ID_Struct_Body": On Error GoTo 0

    ' Dec Type combo
    L = 3720
    Set lbl = CreateControl(tmpName, acLabel, acDetail, "", "", L, T + 15, 1200, 260)
    lbl.Caption = "Dec. type:"
    lbl.BackStyle = 0: lbl.BorderStyle = 0

    Dim c2 As Control
    Set c2 = CreateControl(tmpName, acComboBox, acDetail, "", "", L + 1260, T, 2200, 315)
    c2.ControlSource = "ID_Dec_Type"
    c2.RowSourceType = "Table/Query"
    c2.RowSource = "SELECT ID, Name FROM L_DEC_TYPE ORDER BY Name"
    c2.BoundColumn = 1: c2.ColumnCount = 2
    c2.ColumnWidths = "0cm;4cm": c2.LimitToList = True
    On Error Resume Next: c2.Name = "ID_Dec_Type": On Error GoTo 0

    ' Body No text
    L = 7200
    Set lbl = CreateControl(tmpName, acLabel, acDetail, "", "", L, T + 15, 600, 260)
    lbl.Caption = "Body:"
    lbl.BackStyle = 0: lbl.BorderStyle = 0

    Dim t1 As Control
    Set t1 = CreateControl(tmpName, acTextBox, acDetail, "", "", L + 660, T, 600, 315)
    t1.ControlSource = "Body_No"
    On Error Resume Next: t1.Name = "Body_No": On Error GoTo 0

    ' Color combo
    L = 8560
    Set lbl = CreateControl(tmpName, acLabel, acDetail, "", "", L, T + 15, 600, 260)
    lbl.Caption = "Color:"
    lbl.BackStyle = 0: lbl.BorderStyle = 0

    Dim c3 As Control
    Set c3 = CreateControl(tmpName, acComboBox, acDetail, "", "", L + 620, T, 1400, 315)
    c3.ControlSource = "Color"
    c3.RowSourceType = "Value List"
    c3.RowSource = "Red;White;Both;None;ND"
    c3.LimitToList = False
    On Error Resume Next: c3.Name = "Color": On Error GoTo 0

    DoCmd.Save acForm, SFRM
    DoCmd.Close acForm, SFRM
    DoCmd.Rename SFRM, acForm, tmpName
    Debug.Print "[OK] F_DECORATIONS created"
End Sub

' -- HELPERS ------------------------------------------------------

Sub SH(frm As String, pg As String, txt As String, row As Integer)
    Dim T As Long: T = MT + row * RG - 16
    Dim lh As Control
    Set lh = CreateControl(frm, acLabel, acDetail, pg, "", C1, T, 9000, 260)
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
    If col = 1 Then L = C1 Else L = C2
    T = MT + row * RG
    Dim lc As Control
    Set lc = CreateControl(frm, acLabel, acDetail, pg, "", L, T + 22, LW, LH)
    lc.Caption = lbl: lc.BackStyle = 0: lc.BorderStyle = 0
    lc.TextAlign = 3: lc.ForeColor = RGB(55, 55, 80)
    Dim cc As Control
    Set cc = CreateControl(frm, acCheckBox, acDetail, pg, "", L + LW + 80, T + 20, 300, CH)
    cc.ControlSource = src
    On Error Resume Next: cc.Name = src: On Error GoTo 0
End Sub

Sub PCT(frm As String, pg As String, lbl As String, src As String, row As Integer, col As Integer)
    Dim L As Long: Dim T As Long
    If col = 1 Then L = C1 Else L = C2
    T = MT + row * RG
    Dim lc As Control
    Set lc = CreateControl(frm, acLabel, acDetail, pg, "", L, T + 22, LW, LH)
    lc.Caption = lbl: lc.BackStyle = 0: lc.BorderStyle = 0
    lc.TextAlign = 3: lc.ForeColor = RGB(55, 55, 80)
    Dim cc As Control
    Set cc = CreateControl(frm, acTextBox, acDetail, pg, "", L + LW + 80, T, CW, CH)
    cc.ControlSource = src
    On Error Resume Next: cc.Name = src: On Error GoTo 0
End Sub

Sub PCV(frm As String, pg As String, lbl As String, src As String, row As Integer, col As Integer, vals As String)
    ' Value List ComboBox
    Dim L As Long: Dim T As Long
    If col = 1 Then L = C1 Else L = C2
    T = MT + row * RG
    Dim lc As Control
    Set lc = CreateControl(frm, acLabel, acDetail, pg, "", L, T + 22, LW, LH)
    lc.Caption = lbl: lc.BackStyle = 0: lc.BorderStyle = 0
    lc.TextAlign = 3: lc.ForeColor = RGB(55, 55, 80)
    Dim cc As Control
    Set cc = CreateControl(frm, acComboBox, acDetail, pg, "", L + LW + 80, T, CW, CH)
    cc.ControlSource = src
    cc.RowSourceType = "Value List"
    cc.RowSource = vals
    cc.LimitToList = False
    On Error Resume Next: cc.Name = src: On Error GoTo 0
End Sub
