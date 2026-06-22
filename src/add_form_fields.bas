Option Compare Database
Option Explicit

' ================================================================
'  ADD NEW FIELDS TO F_STRUCTURES FORM
'  Run Sub AddMissingFields() -> F5
' ================================================================

Sub AddMissingFields()
    Const FRM = "F_STRUCTURES"
    Const MT  = 550: Const RG  = 420
    Const C1  = 120: Const C2  = 4880
    Const LW  = 1900: Const CW = 2500
    Const CH  = 315: Const LH  = 270

    DoCmd.OpenForm FRM, acDesign

    ' -- STATUS TAB (pgEst): ID_Material_Status at row 6 col 1 --
    Dim L As Long: Dim T As Long
    L = C1: T = MT + 6 * RG

    Dim lbl As Control
    Set lbl = CreateControl(FRM, acLabel, acDetail, "pgEst", "", L, T + 22, LW, LH)
    lbl.Caption = "Material Status:"
    lbl.BackStyle = 0: lbl.BorderStyle = 0
    lbl.TextAlign = 3: lbl.ForeColor = RGB(55, 55, 80)

    Dim cmb As Control
    Set cmb = CreateControl(FRM, acComboBox, acDetail, "pgEst", "", L + LW + 80, T, CW, CH)
    cmb.ControlSource = "ID_Material_Status"
    cmb.RowSourceType = "Table/Query"
    cmb.RowSource = "SELECT ID, Name FROM L_MATERIAL_STATUS ORDER BY ID"
    cmb.BoundColumn = 1
    cmb.ColumnCount = 2
    cmb.ColumnWidths = "0cm;4cm"
    cmb.LimitToList = True
    On Error Resume Next: cmb.Name = "ID_Material_Status": On Error GoTo 0

    ' -- CHRONOLOGY TAB (pgCron): Interior_Area_m2 at row 9 col 1 -
    L = C1: T = MT + 9 * RG

    Set lbl = CreateControl(FRM, acLabel, acDetail, "pgCron", "", L, T + 22, LW, LH)
    lbl.Caption = "Interior area (m2):"
    lbl.BackStyle = 0: lbl.BorderStyle = 0
    lbl.TextAlign = 3: lbl.ForeColor = RGB(55, 55, 80)

    Dim txt As Control
    Set txt = CreateControl(FRM, acTextBox, acDetail, "pgCron", "", L + LW + 80, T, CW, CH)
    txt.ControlSource = "Interior_Area_m2"
    On Error Resume Next: txt.Name = "Interior_Area_m2": On Error GoTo 0

    ' -- Also update ID_Arch_Status label if needed ----------------
    Dim ctrl As Control
    For Each ctrl In Forms(FRM).Controls
        If ctrl.ControlType = acComboBox Then
            If ctrl.ControlSource = "ID_Arch_Status" Then
                On Error Resume Next
                ctrl.RowSourceType = "Table/Query"
                ctrl.RowSource = "SELECT ID, Name FROM L_STATUS ORDER BY ID"
                ctrl.BoundColumn = 1
                ctrl.ColumnCount = 2
                ctrl.ColumnWidths = "0cm;4cm"
                ctrl.LimitToList = True
                On Error GoTo 0
            End If
        End If
        ' Update label next to ID_Arch_Status combo
        If ctrl.ControlType = acLabel Then
            If ctrl.Caption = "Arch. Status:" Or ctrl.Caption = "General status:" Then
                ctrl.Caption = "Arch. Status:"
            End If
        End If
    Next ctrl

    DoCmd.Save acForm, FRM
    DoCmd.Close acForm, FRM

    Dim msg As String
    msg = "Done! Fields added to F_STRUCTURES:" & vbCrLf & vbCrLf
    msg = msg & "  - Tab 5 (Status): ID_Material_Status combo" & vbCrLf
    msg = msg & "  - Tab 8 (Chronology): Interior_Area_m2 field"
    MsgBox msg, vbInformation, "Done!"
End Sub
