Option Compare Database
Option Explicit

' ================================================================
'  FIX COMBO BOXES - F_STRUCTURES
'  Run Sub FixAllCombos() -> F5
'  Searches by ControlSource (not by name) - works always
' ================================================================

Sub FixAllCombos()
    Const FRM = "F_STRUCTURES"

    ' Map: ControlSource -> RowSource SQL, ColumnCount
    Dim cs(8)  As String   ' ControlSource field name
    Dim rs(8)  As String   ' RowSource SQL
    Dim cc(8)  As Integer  ' ColumnCount
    Dim cw(8)  As String   ' ColumnWidths

    cs(0) = "ID_Sector"
    rs(0) = "SELECT ID, Sector_Name FROM L_SECTORS ORDER BY ID_Site, Sector_Name"
    cc(0) = 2: cw(0) = "0cm;5cm"

    cs(1) = "ID_Typology"
    rs(1) = "SELECT ID, Name FROM L_TYPOLOGY ORDER BY Name"
    cc(1) = 2: cw(1) = "0cm;6cm"

    cs(2) = "ID_Support"
    rs(2) = "SELECT ID, Name FROM L_SUPPORT ORDER BY ID"
    cc(2) = 2: cw(2) = "0cm;5cm"

    cs(3) = "ID_Status"
    rs(3) = "SELECT ID, Name FROM L_STATUS ORDER BY ID"
    cc(3) = 2: cw(3) = "0cm;4cm"

    cs(4) = "ID_Vol_Method"
    rs(4) = "SELECT ID, Name FROM L_VOL_METHOD ORDER BY ID"
    cc(4) = 2: cw(4) = "0cm;5cm"

    cs(5) = "ID_Coord_Method"
    rs(5) = "SELECT ID, Name FROM L_COORD_METHOD ORDER BY ID"
    cc(5) = 2: cw(5) = "0cm;5cm"

    cs(6) = "ID_Campaign"
    rs(6) = "SELECT ID, Code, Campaign_Name FROM L_CAMPAIGN ORDER BY Code"
    cc(6) = 3: cw(6) = "0cm;1.5cm;5cm"

    cs(7) = "ID_Group"
    rs(7) = "SELECT ID, Group_Code FROM T_GROUPS ORDER BY Group_Code"
    cc(7) = 2: cw(7) = "0cm;4cm"

    cs(8) = "ID_Parent"
    rs(8) = "SELECT ID, Code FROM T_STRUCTURES ORDER BY Code"
    cc(8) = 2: cw(8) = "0cm;4cm"

    ' Open form in design view
    DoCmd.OpenForm FRM, acDesign
    Dim f As Form
    Set f = Forms(FRM)

    Dim ctrl  As Control
    Dim fixed As Integer
    fixed = 0

    ' Iterate ALL controls (includes controls inside tab pages)
    For Each ctrl In f.Controls
        If ctrl.ControlType = acComboBox Then
            Dim i As Integer
            For i = 0 To 8
                If ctrl.ControlSource = cs(i) Then
                    ctrl.RowSourceType = "Table/Query"
                    ctrl.RowSource     = rs(i)
                    ctrl.BoundColumn   = 1
                    ctrl.ColumnCount   = cc(i)
                    ctrl.ColumnWidths  = cw(i)
                    ctrl.LimitToList   = True
                    fixed = fixed + 1
                    Debug.Print "  [OK] " & cs(i)
                End If
            Next i
        End If
    Next ctrl

    DoCmd.Save acForm, FRM
    DoCmd.Close acForm, FRM

    Dim msg As String
    msg = fixed & " combo boxes configured correctly!"
    msg = msg & vbCrLf & "Open F_STRUCTURES to verify."
    MsgBox msg, vbInformation, "Done!"
End Sub
