Option Compare Database
Option Explicit

' ================================================================
'  CHACHAPOYA FORM BUILD SCRIPT v11 (VALENCIAN) - F_STRUCTURES
'  Author: Esteve Ribera Torro | TFM Arqueologia UA
'  Spec: DELTA_v10_v11.md (rev. 5)
'
'  PRINCIPLE: labels (UI) in Valencian | stored values in English
'
'  IMPORTANT: run AFTER the schema exists, i.e. after
'  chachapoya_DB_v11.bas -> BuildDB() on a blank database, or after
'  migrate_v10_to_v11.bas -> MigrateV10toV11() on the v10 one.
'
'  CHANGES FROM v10
'
'  1. EVERY CONTROLLED COMBO IS NOW TWO COLUMNS (section 0). Column 1
'     holds the English value that gets stored and is hidden; column 2
'     shows the Valencian label. In a one-column value list the stored
'     value IS the label, which made "relabel without changing stored
'     values" (1.4) literally impossible - the ND case would have
'     produced a corpus mixing 'ND' and 'Type undetermined'.
'
'  2. PC9 SPLITS INTO PC5 AND PC9. PC5 serves the 20 element fields
'     with the five-value domain; PC9 keeps three values for the
'     fields of 1.6, with 9 relabelled from ND to "No observable".
'     That relabel matters: 9 never meant "not observed", it means
'     the position cannot be examined.
'
'  3. TAB 11.Sist REORGANISED AROUND THE SYSTEMS (4.5). The five
'     Sys_* combos sit at the top and their component groups follow.
'     Lintel, Lintel_Material and Chamber_Roof_Type move here from
'     2.Arq, into the portal and chamber groups respectively (12bis).
'
'  4. TAB 4.Dec IS NOW ONLY THE T_DECORATIONS SUBFORM (7.4). The
'     quick boolean grid is gone with the fields behind it.
'
'  5. GATING, IN TWO LEVELS (4.4, 4.5, 4.6). Record_Class decides
'     which tabs are active; each Sys_* decides whether its component
'     group is. ID_Material_Status and Human_Remains do the same for
'     7.Mat and 6.Bio without needing a new field.
'
'  6. NEW SUBFORM F_CONNECTIONS (9.2), so recording the relation
'     between the five a/b pairs is the path of least resistance.
'
'  NOTE ON THE GATING CODE: levels 1 and 2 need real VBA behind the
'  form, which this script injects with Form.Module.AddFromString.
'  That requires "Trust access to the VBA project object model" in
'  the Access Trust Centre. If it is off, everything else still
'  builds and the form works - only the automatic enable/disable is
'  missing, and the script says so instead of failing silently.
' ================================================================

' Layout constants
Const MT  As Long = 550
Const RG  As Long = 420
Const C1  As Long = 120
Const C2  As Long = 5600
Const LW  As Long = 2600
Const CW  As Long = 2200
Const CH  As Long = 315
Const LH  As Long = 270
Const FW  As Long = 13200

' Two-column domains. Stored value first, Valencian label second.
' The stored side is English and never changes; only the label does.
Const DOM5 As String = "0;Absent;1;Present complet;2;Present parcial;3;Desaparegut;9;No observable"
Const DOM3 As String = "0;Absent;1;Present;9;No observable"
Const DOMSYS As String = "Present complete;Present complet;Present partial;Present parcial;Attested lost;Desaparegut;Absent;Absent;Not applicable;No aplicable;Not observable;No observable"
Const DOMGRP As String = "Present;Present;Absent;Absent;Not applicable;No aplicable;Not observable;No observable"

Sub BuildForm()
    CreateSubForms
    CreateMainForm
    SetFieldCaptionsVal
    Dim msg As String
    msg = "F_STRUCTURES v11 (val.) creada amb 12 pestanyes!" & vbCrLf & vbCrLf
    msg = msg & "  20 camps d'element amb domini de 5 valors" & vbCrLf
    msg = msg & "  24 camps observacionals amb 0/1/9" & vbCrLf
    msg = msg & "  Tots els combos de domini son de dues columnes:" & vbCrLf
    msg = msg & "  el valor guardat es en angles, l'etiqueta en valencia" & vbCrLf & vbCrLf
    msg = msg & "  11.Sist: els 5 sistemes manen sobre els seus grups" & vbCrLf
    msg = msg & "  4.Dec: nomes el subformulari T_DECORATIONS" & vbCrLf
    msg = msg & "  12.Extra: connexions i elements personalitzats" & vbCrLf & vbCrLf
    msg = msg & "El valor per defecte es 0 (Absent)." & vbCrLf
    msg = msg & "El 9 vol dir que la posicio NO es examinable," & vbCrLf
    msg = msg & "no que no s'haja mirat: s'ha de marcar a consciencia." & vbCrLf & vbCrLf
    msg = msg & "Executa QRY_16_Validation_Check periodicament."
    MsgBox msg, vbInformation, "Fet!"
End Sub

' ================================================================
'  COLUMN CAPTIONS FOR DATASHEET VIEW
'  The attached labels on the subform controls are the main
'  mechanism; the DAO captions back them up for when a table is
'  opened directly, outside the form.
' ================================================================
Private Sub SetFieldCaptionsVal()
    Dim db As DAO.Database
    Set db = CurrentDb()
    SetCap db, "T_DECORATIONS", "ID_Struct_Body", "Posicio"
    SetCap db, "T_DECORATIONS", "ID_Dec_Type", "Tipus dec."
    SetCap db, "T_DECORATIONS", "Body_No", "Num. cos"
    SetCap db, "T_DECORATIONS", "Color", "Color"
    SetCap db, "T_DECORATIONS", "Substrate", "Substrat"
    SetCap db, "T_DECORATIONS", "Notes", "Notes"
    SetCap db, "T_ARCH_FEATURES", "Feature_Code", "Element"
    SetCap db, "T_ARCH_FEATURES", "Present", "Present"
    SetCap db, "T_ARCH_FEATURES", "Feature_Count", "Nombre"
    SetCap db, "T_ARCH_FEATURES", "Material", "Material"
    SetCap db, "T_ARCH_FEATURES", "Notes", "Notes"
    SetCap db, "T_CONNECTIONS", "ID_Struct_A", "Estructura A"
    SetCap db, "T_CONNECTIONS", "ID_Struct_B", "Estructura B"
    SetCap db, "T_CONNECTIONS", "Connection_Type", "Tipus connexio"
    SetCap db, "T_CONNECTIONS", "Chrono_Relation", "Relacio cronologica"
    SetCap db, "T_CONNECTIONS", "Confidence", "Confianca"
    SetCap db, "T_CONNECTIONS", "Notes", "Notes"
    SetCap db, "T_LOST_ELEMENTS", "Element_Code", "Element"
    SetCap db, "T_LOST_ELEMENTS", "ID_Evidence_Type", "Tipus evidencia"
    SetCap db, "T_LOST_ELEMENTS", "Evidence_Scope", "Abast"
    SetCap db, "T_LOST_ELEMENTS", "ID_Position", "Posicio"
    SetCap db, "T_LOST_ELEMENTS", "Notes", "Notes"
    db.TableDefs.Refresh
    Set db = Nothing
    Debug.Print "[OK] Captions DAO establertes"
End Sub

Private Sub SetCap(db As DAO.Database, tbl As String, fld As String, cap As String)
    On Error GoTo TryCreate
    db.TableDefs(tbl).Fields(fld).Properties("Caption") = cap
    Exit Sub
TryCreate:
    On Error Resume Next
    Dim p As DAO.Property
    Set p = db.TableDefs(tbl).Fields(fld).CreateProperty("Caption", dbText, cap)
    db.TableDefs(tbl).Fields(fld).Properties.Append p
    On Error GoTo 0
End Sub

' ================================================================
'  HELPERS
' ================================================================
Private Sub SH(frm As String, pg As String, txt As String, row As Integer)
    Dim T As Long: T = MT + row * RG - 16
    Dim lh As Control
    Set lh = CreateControl(frm, acLabel, acDetail, pg, "", C1, T, FW - 200, 260)
    lh.Caption = "  " & UCase(txt)
    lh.BackStyle = 1: lh.BackColor = RGB(214, 228, 247)
    lh.BorderStyle = 0: lh.ForeColor = RGB(26, 60, 107)
    lh.FontBold = True: lh.FontSize = 8
End Sub

Private Sub PCB(frm As String, pg As String, lbl As String, src As String, row As Integer, col As Integer)
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

Private Sub PCT(frm As String, pg As String, lbl As String, src As String, row As Integer, col As Integer)
    AddCtrl frm, pg, lbl, src, acTextBox, row, col, CW
End Sub

Private Sub PCC(frm As String, pg As String, lbl As String, src As String, row As Integer, col As Integer)
    AddCtrl frm, pg, lbl, src, acComboBox, row, col, CW
End Sub

' Two-column value list. vals is "stored;label;stored;label;..."
' The stored column is hidden, so relabelling never touches the data.
Private Sub PCV(frm As String, pg As String, lbl As String, src As String, row As Integer, col As Integer, vals As String)
    AddCtrl frm, pg, lbl, src, acComboBox, row, col, CW
    TwoColumn frm, src, vals, "0cm;4cm"
End Sub

' Element fields: the five-value domain of 1.1.
Private Sub PC5(frm As String, pg As String, lbl As String, src As String, row As Integer, col As Integer)
    AddCtrl frm, pg, lbl, src, acComboBox, row, col, 2000
    TwoColumn frm, src, DOM5, "0cm;3.5cm"
End Sub

' Observational fields that keep three values (1.6). Only the label
' of 9 changes: ND becomes "No observable".
Private Sub PC9(frm As String, pg As String, lbl As String, src As String, row As Integer, col As Integer)
    AddCtrl frm, pg, lbl, src, acComboBox, row, col, 2000
    TwoColumn frm, src, DOM3, "0cm;3.5cm"
End Sub

' The three element-systems: six values (4.1).
Private Sub PCS(frm As String, pg As String, lbl As String, src As String, row As Integer, col As Integer)
    AddCtrl frm, pg, lbl, src, acComboBox, row, col, 2400
    TwoColumn frm, src, DOMSYS, "0cm;4cm"
End Sub

' The two grouping fields: four values (4.2).
Private Sub PCG(frm As String, pg As String, lbl As String, src As String, row As Integer, col As Integer)
    AddCtrl frm, pg, lbl, src, acComboBox, row, col, 2400
    TwoColumn frm, src, DOMGRP, "0cm;4cm"
End Sub

' LimitToList is ALWAYS True here, and not by choice: once the bound
' column is hidden (0cm) Access refuses to set it to False at all,
' raising error 7773. It has to, because with the stored value out of
' sight there would be no way to tell what typing free text should
' store. So every two-column domain combo is a closed list, which is
' what a controlled domain wants anyway - the escape hatch for genuine
' exceptions is an explicit "Other (see Notes)" entry plus the Notes
' field, the pattern the delta itself uses for Connection_Type.
Private Sub TwoColumn(frm As String, src As String, vals As String, widths As String)
    Dim ctrl As Control
    On Error Resume Next
    Set ctrl = Forms(frm).Controls(src)
    If Not ctrl Is Nothing Then
        ctrl.RowSourceType = "Value List"
        ctrl.RowSource = vals
        ctrl.ColumnCount = 2
        ctrl.BoundColumn = 1
        ctrl.ColumnWidths = widths
        ctrl.LimitToList = True
    End If
    On Error GoTo 0
End Sub

Private Sub PCM(frm As String, pg As String, lbl As String, src As String, row As Integer, col As Integer)
    Dim L As Long: Dim T As Long
    If col = 1 Then L = C1 Else L = C2
    T = MT + row * RG
    Dim lc As Control
    Set lc = CreateControl(frm, acLabel, acDetail, pg, "", L, T + 22, LW, LH)
    lc.Caption = lbl: lc.BackStyle = 0: lc.BorderStyle = 0
    lc.TextAlign = 3: lc.ForeColor = RGB(55, 55, 80)
    Dim cc As Control
    Set cc = CreateControl(frm, acTextBox, acDetail, pg, "", L + LW + 80, T, CW * 2 + 100, CH * 4)
    cc.ControlSource = src: cc.ScrollBars = 2
    On Error Resume Next: cc.Name = src: On Error GoTo 0
End Sub

Private Sub AddCtrl(frm As String, pg As String, lbl As String, src As String, cType As Integer, row As Integer, col As Integer, w As Long)
    Dim L As Long: Dim T As Long
    If col = 1 Then L = C1 Else L = C2
    T = MT + row * RG
    Dim lc As Control
    Set lc = CreateControl(frm, acLabel, acDetail, pg, "", L, T + 22, LW, LH)
    lc.Caption = lbl: lc.BackStyle = 0: lc.BorderStyle = 0
    lc.TextAlign = 3: lc.ForeColor = RGB(55, 55, 80)
    Dim cc As Control
    Set cc = CreateControl(frm, cType, acDetail, pg, "", L + LW + 80, T, w, CH)
    cc.ControlSource = src
    On Error Resume Next: cc.Name = src: On Error GoTo 0
End Sub

' ================================================================
'  SUBFORMS
'  Pattern: the control first (named after the field), then the
'  ATTACHED label (parent = control name). In datasheet view Access
'  shows the attached label's caption as the column header; without
'  one it showed the English field name.
' ================================================================
Private Sub CreateSubForms()
    CreateDecSubform
    CreateFeatSubform
    CreateConnSubform
End Sub

Private Sub CreateDecSubform()
    Const SFRM = "F_DECORATIONS"
    On Error Resume Next: DoCmd.DeleteObject acForm, SFRM: On Error GoTo 0
    Dim f As Form: Set f = CreateForm()
    Dim tmp As String: tmp = f.Name
    f.RecordSource = "T_DECORATIONS"
    f.DefaultView = 2: f.ScrollBars = 2
    f.NavigationButtons = False: f.Width = 15600
    f.Section(acDetail).Height = 400
    Dim T As Long: T = 50: Dim L As Long
    Dim lb As Control

    L = 40
    Dim c1 As Control: Set c1 = CreateControl(tmp, acComboBox, acDetail, "", "", L + 1060, T, 2200, 315)
    c1.ControlSource = "ID_Struct_Body"
    c1.RowSourceType = "Table/Query"
    c1.RowSource = "SELECT ID, Name FROM L_STRUCT_BODY ORDER BY Level_Type, Name"
    c1.BoundColumn = 1: c1.ColumnCount = 2: c1.ColumnWidths = "0cm;4cm": c1.LimitToList = True
    On Error Resume Next: c1.Name = "ID_Struct_Body": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "ID_Struct_Body", "", L, T + 15, 1000, 260)
    lb.Caption = "Posicio"

    L = 3400
    Dim c2 As Control: Set c2 = CreateControl(tmp, acComboBox, acDetail, "", "", L + 1060, T, 2000, 315)
    c2.ControlSource = "ID_Dec_Type"
    c2.RowSourceType = "Table/Query"
    c2.RowSource = "SELECT ID, Name FROM L_DEC_TYPE ORDER BY Name"
    c2.BoundColumn = 1: c2.ColumnCount = 2: c2.ColumnWidths = "0cm;4cm": c2.LimitToList = True
    On Error Resume Next: c2.Name = "ID_Dec_Type": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "ID_Dec_Type", "", L, T + 15, 1000, 260)
    lb.Caption = "Tipus dec."

    L = 6600
    Dim c3 As Control: Set c3 = CreateControl(tmp, acTextBox, acDetail, "", "", L + 660, T, 500, 315)
    c3.ControlSource = "Body_No"
    On Error Resume Next: c3.Name = "Body_No": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "Body_No", "", L, T + 15, 600, 260)
    lb.Caption = "Num. cos"

    L = 7900
    Dim c4 As Control: Set c4 = CreateControl(tmp, acComboBox, acDetail, "", "", L + 660, T, 1300, 315)
    c4.ControlSource = "Color": c4.RowSourceType = "Value List"
    c4.RowSource = "Red;Roig;White;Blanc;Both;Ambdos;Ochre;Ocre;None;Cap;ND;Indeterminat"
    c4.ColumnCount = 2: c4.BoundColumn = 1: c4.ColumnWidths = "0cm;3cm": c4.LimitToList = True
    On Error Resume Next: c4.Name = "Color": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "Color", "", L, T + 15, 600, 260)
    lb.Caption = "Color"

    L = 10000
    Dim c5 As Control: Set c5 = CreateControl(tmp, acComboBox, acDetail, "", "", L + 900, T, 1600, 315)
    c5.ControlSource = "Substrate": c5.RowSourceType = "Value List"
    c5.RowSource = "Plaster;Revoc;Masonry stone;Pedra de parament;Bedrock;Penya;ND;Indeterminat"
    c5.ColumnCount = 2: c5.BoundColumn = 1: c5.ColumnWidths = "0cm;4cm": c5.LimitToList = True
    On Error Resume Next: c5.Name = "Substrate": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "Substrate", "", L, T + 15, 840, 260)
    lb.Caption = "Substrat"

    L = 12700
    Dim c6 As Control: Set c6 = CreateControl(tmp, acTextBox, acDetail, "", "", L + 660, T, 2200, 315)
    c6.ControlSource = "Notes"
    On Error Resume Next: c6.Name = "Notes": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "Notes", "", L, T + 15, 600, 260)
    lb.Caption = "Notes"

    DoCmd.Save acForm, tmp: DoCmd.Close acForm, tmp
    DoCmd.Rename SFRM, acForm, tmp
    Debug.Print "[OK] F_DECORATIONS"
End Sub

Private Sub CreateFeatSubform()
    Const SFRM = "F_ARCH_FEATURES"
    On Error Resume Next: DoCmd.DeleteObject acForm, SFRM: On Error GoTo 0
    Dim f As Form: Set f = CreateForm()
    Dim tmp As String: tmp = f.Name
    f.RecordSource = "T_ARCH_FEATURES"
    f.DefaultView = 2: f.ScrollBars = 2
    f.NavigationButtons = False: f.Width = 11800
    f.Section(acDetail).Height = 400
    Dim T As Long: T = 50: Dim L As Long
    Dim lb As Control

    L = 40
    Dim c1 As Control: Set c1 = CreateControl(tmp, acComboBox, acDetail, "", "", L + 1060, T, 2400, 315)
    ' T_ARCH_FEATURES exists to record what the fixed vocabulary does
    ' not anticipate, so the open-ended case matters here more than
    ' anywhere: "Other (see Notes)" plus the Notes column is the escape
    ' hatch, since a hidden bound column cannot accept free text.
    c1.ControlSource = "Feature_Code": c1.RowSourceType = "Value List"
    c1.RowSource = "Pigment trace;Traca de pigment;Unusual bond;Aparell anomal;Textile fixation;Fixacio textil;Wooden peg;Clavilla de fusta;Other (see Notes);Altres (veure notes)"
    c1.ColumnCount = 2: c1.BoundColumn = 1: c1.ColumnWidths = "0cm;5cm": c1.LimitToList = True
    On Error Resume Next: c1.Name = "Feature_Code": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "Feature_Code", "", L, T + 15, 1000, 260)
    lb.Caption = "Element"

    L = 3560
    Dim c2 As Control: Set c2 = CreateControl(tmp, acCheckBox, acDetail, "", "", L + 760, T + 20, 300, 300)
    c2.ControlSource = "Present"
    On Error Resume Next: c2.Name = "Present": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "Present", "", L, T + 15, 700, 260)
    lb.Caption = "Present"

    L = 4700
    Dim c3 As Control: Set c3 = CreateControl(tmp, acTextBox, acDetail, "", "", L + 660, T, 700, 315)
    c3.ControlSource = "Feature_Count"
    On Error Resume Next: c3.Name = "Feature_Count": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "Feature_Count", "", L, T + 15, 600, 260)
    lb.Caption = "Nombre"

    L = 6200
    Dim c4 As Control: Set c4 = CreateControl(tmp, acComboBox, acDetail, "", "", L + 760, T, 1200, 315)
    c4.ControlSource = "Material": c4.RowSourceType = "Value List"
    c4.RowSource = "Stone;Pedra;Timber;Fusta;Mixed;Mixt;ND;Indeterminat"
    c4.ColumnCount = 2: c4.BoundColumn = 1: c4.ColumnWidths = "0cm;3cm": c4.LimitToList = True
    On Error Resume Next: c4.Name = "Material": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "Material", "", L, T + 15, 700, 260)
    lb.Caption = "Material"

    L = 8300
    Dim c5 As Control: Set c5 = CreateControl(tmp, acTextBox, acDetail, "", "", L + 660, T, 2600, 315)
    c5.ControlSource = "Notes"
    On Error Resume Next: c5.Name = "Notes": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "Notes", "", L, T + 15, 600, 260)
    lb.Caption = "Notes"

    DoCmd.Save acForm, tmp: DoCmd.Close acForm, tmp
    DoCmd.Rename SFRM, acForm, tmp
    Debug.Print "[OK] F_ARCH_FEATURES"
End Sub

' NEW v11 (9.2, 9.1bis). The closed Connection_Type list is what makes
' QRY_16 rule 20 computable at all: with free text there was no way to
' tell which connection types can carry a direction.
Private Sub CreateConnSubform()
    Const SFRM = "F_CONNECTIONS"
    On Error Resume Next: DoCmd.DeleteObject acForm, SFRM: On Error GoTo 0
    Dim f As Form: Set f = CreateForm()
    Dim tmp As String: tmp = f.Name
    f.RecordSource = "T_CONNECTIONS"
    f.DefaultView = 2: f.ScrollBars = 2
    f.NavigationButtons = False: f.Width = 15000
    f.Section(acDetail).Height = 400
    Dim T As Long: T = 50: Dim L As Long
    Dim lb As Control

    L = 40
    Dim c1 As Control: Set c1 = CreateControl(tmp, acComboBox, acDetail, "", "", L + 1200, T, 2000, 315)
    c1.ControlSource = "ID_Struct_A"
    c1.RowSourceType = "Table/Query"
    c1.RowSource = "SELECT ID, Code FROM T_STRUCTURES ORDER BY Code"
    c1.BoundColumn = 1: c1.ColumnCount = 2: c1.ColumnWidths = "0cm;4cm": c1.LimitToList = True
    On Error Resume Next: c1.Name = "ID_Struct_A": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "ID_Struct_A", "", L, T + 15, 1140, 260)
    lb.Caption = "Estructura A"

    L = 3300
    Dim c2 As Control: Set c2 = CreateControl(tmp, acComboBox, acDetail, "", "", L + 1200, T, 2000, 315)
    c2.ControlSource = "ID_Struct_B"
    c2.RowSourceType = "Table/Query"
    c2.RowSource = "SELECT ID, Code FROM T_STRUCTURES ORDER BY Code"
    c2.BoundColumn = 1: c2.ColumnCount = 2: c2.ColumnWidths = "0cm;4cm": c2.LimitToList = True
    On Error Resume Next: c2.Name = "ID_Struct_B": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "ID_Struct_B", "", L, T + 15, 1140, 260)
    lb.Caption = "Estructura B"

    L = 6600
    Dim c3 As Control: Set c3 = CreateControl(tmp, acComboBox, acDetail, "", "", L + 1200, T, 2400, 315)
    c3.ControlSource = "Connection_Type": c3.RowSourceType = "Value List"
    c3.RowSource = "Abutted vertical joint;Junta vertical adossada;Superposition;Superposicio;Bonded joint;Junta travada;Shared support;Suport compartit;Aerial connection;Connexio aeria;Other (see Notes);Altres (veure notes)"
    c3.ColumnCount = 2: c3.BoundColumn = 1: c3.ColumnWidths = "0cm;5cm": c3.LimitToList = True
    On Error Resume Next: c3.Name = "Connection_Type": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "Connection_Type", "", L, T + 15, 1140, 260)
    lb.Caption = "Tipus connexio"

    ' Direction is read from the joint: the structure showing the
    ' untoothed joint against the other's wall face is the later one.
    ' The order of A and B must NOT be swapped after entry (9.1).
    L = 10300
    Dim c4 As Control: Set c4 = CreateControl(tmp, acComboBox, acDetail, "", "", L + 1300, T, 2200, 315)
    c4.ControlSource = "Chrono_Relation": c4.RowSourceType = "Value List"
    c4.RowSource = "A earlier than B;A anterior a B;B earlier than A;B anterior a A;Contemporary;Contemporanis;Undetermined;Indeterminat"
    c4.ColumnCount = 2: c4.BoundColumn = 1: c4.ColumnWidths = "0cm;4.5cm": c4.LimitToList = True
    On Error Resume Next: c4.Name = "Chrono_Relation": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "Chrono_Relation", "", L, T + 15, 1240, 260)
    lb.Caption = "Cronologia"

    L = 14000
    Dim c5 As Control: Set c5 = CreateControl(tmp, acComboBox, acDetail, "", "", L + 900, T, 1200, 315)
    c5.ControlSource = "Confidence": c5.RowSourceType = "Value List"
    c5.RowSource = "High;Alta;Medium;Mitjana;Low;Baixa"
    c5.ColumnCount = 2: c5.BoundColumn = 1: c5.ColumnWidths = "0cm;3cm": c5.LimitToList = True
    On Error Resume Next: c5.Name = "Confidence": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "Confidence", "", L, T + 15, 840, 260)
    lb.Caption = "Confianca"

    L = 16200
    Dim c6 As Control: Set c6 = CreateControl(tmp, acTextBox, acDetail, "", "", L + 660, T, 2200, 315)
    c6.ControlSource = "Notes"
    On Error Resume Next: c6.Name = "Notes": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "Notes", "", L, T + 15, 600, 260)
    lb.Caption = "Notes"

    DoCmd.Save acForm, tmp: DoCmd.Close acForm, tmp
    DoCmd.Rename SFRM, acForm, tmp
    Debug.Print "[OK] F_CONNECTIONS"
End Sub

' ================================================================
'  MAIN FORM F_STRUCTURES
' ================================================================
Private Sub CreateMainForm()
    Const FRM = "F_STRUCTURES"
    On Error Resume Next: DoCmd.DeleteObject acForm, FRM: On Error GoTo 0

    Dim f As Form: Set f = CreateForm()
    Dim tmp As String: tmp = f.Name

    f.RecordSource = "T_STRUCTURES"
    f.DefaultView = 0: f.ScrollBars = 3
    f.NavigationButtons = True
    f.Caption = "Registre Estructura v11 - La Petaca i Diablo Wasi (PALP)"
    f.Width = FW

    f.Section(acDetail).Height = 9400
    f.Section(acDetail).BackColor = RGB(249, 249, 248)

    Dim h As Control
    Set h = CreateControl(tmp, acLabel, acDetail, "", "", 120, 80, 7000, 480)
    h.Caption = "REGISTRE D'ESTRUCTURA v11  -  La Petaca i Diablo Wasi (PALP)"
    h.FontSize = 13: h.FontBold = True
    h.ForeColor = RGB(26, 60, 107): h.BackStyle = 0: h.BorderStyle = 0

    Dim tc As Control
    Set tc = CreateControl(tmp, acTabCtl, acDetail, "", "", 60, 620, 13080, 8700)
    tc.Name = "tabMain"

    tc.Pages(0).Name = "pgId":  tc.Pages(0).Caption = "1.Id."
    tc.Pages(1).Name = "pgArq": tc.Pages(1).Caption = "2.Arq."
    Dim pg As Control
    Set pg = CreateControl(tmp, acPage, acDetail, "tabMain"): pg.Name = "pgAcab":  pg.Caption = "3.Acab."
    Set pg = CreateControl(tmp, acPage, acDetail, "tabMain"): pg.Name = "pgDec":   pg.Caption = "4.Dec."
    Set pg = CreateControl(tmp, acPage, acDetail, "tabMain"): pg.Name = "pgEst":   pg.Caption = "5.Estat"
    Set pg = CreateControl(tmp, acPage, acDetail, "tabMain"): pg.Name = "pgBio":   pg.Caption = "6.Bio."
    Set pg = CreateControl(tmp, acPage, acDetail, "tabMain"): pg.Name = "pgMat":   pg.Caption = "7.Mat."
    Set pg = CreateControl(tmp, acPage, acDetail, "tabMain"): pg.Name = "pgCron":  pg.Caption = "8.Cron."
    Set pg = CreateControl(tmp, acPage, acDetail, "tabMain"): pg.Name = "pgMetr":  pg.Caption = "9.Metr."
    Set pg = CreateControl(tmp, acPage, acDetail, "tabMain"): pg.Name = "pgDoc":   pg.Caption = "10.Doc."
    Set pg = CreateControl(tmp, acPage, acDetail, "tabMain"): pg.Name = "pgSys":   pg.Caption = "11.Sist."
    Set pg = CreateControl(tmp, acPage, acDetail, "tabMain"): pg.Name = "pgExtra": pg.Caption = "12.Extra"

    FillId tmp: FillArq tmp: FillAcab tmp: FillDec tmp: FillEst tmp
    FillBio tmp: FillMat tmp: FillCron tmp: FillMetr tmp: FillDoc tmp
    FillSys tmp: FillExtra tmp

    ' Tab 4.Dec is now nothing but this subform (7.4): the boolean
    ' grid went away with the fields behind it.
    Dim sf1 As Control
    Set sf1 = CreateControl(tmp, acSubform, acDetail, "pgDec", "", C1, MT + 1 * RG + 340, 12000, 5200)
    sf1.SourceObject = "F_DECORATIONS"
    sf1.LinkMasterFields = "ID": sf1.LinkChildFields = "ID_Structure"

    Dim sf2 As Control
    Set sf2 = CreateControl(tmp, acSubform, acDetail, "pgExtra", "", C1, MT + 1 * RG + 340, 12000, 2200)
    sf2.SourceObject = "F_ARCH_FEATURES"
    sf2.LinkMasterFields = "ID": sf2.LinkChildFields = "ID_Structure"

    ' Connections hang off structure A, so the pair is recorded from
    ' the record you are already looking at (9.2).
    Dim sf3 As Control
    Set sf3 = CreateControl(tmp, acSubform, acDetail, "pgExtra", "", C1, MT + 8 * RG + 340, 12000, 2200)
    sf3.SourceObject = "F_CONNECTIONS"
    sf3.LinkMasterFields = "ID": sf3.LinkChildFields = "ID_Struct_A"

    DoCmd.Save acForm, tmp
    DoCmd.Close acForm, tmp
    DoCmd.Rename FRM, acForm, tmp
    Debug.Print "[OK] F_STRUCTURES v11 (valencia) creada"

    ConfigureAllCombos FRM
    InjectGating FRM
End Sub

' ================================================================
'  TAB CONTENTS
' ================================================================

' TAB 1 - IDENTIFICATION AND SUPPORT DETAIL
Private Sub FillId(f As String)
    SH f, "pgId", "Identificacio i localitzacio", 0
    PCT f, "pgId", "Codi:",             "Code",                 1, 1
    PCC f, "pgId", "Sector:",           "ID_Sector",            2, 1
    PCC f, "pgId", "Tipologia:",        "ID_Typology",          3, 1
    PCC f, "pgId", "Suport geom.:",     "ID_Support",           4, 1
    PCC f, "pgId", "Suport secundari:", "ID_Support_Secondary", 1, 2
    PCC f, "pgId", "Element pare:",     "ID_Parent",            2, 2
    PCC f, "pgId", "Grup funcional:",   "ID_Group",             3, 2
    SH  f, "pgId", "Detall suport geologic (H02)", 5
    PC9 f, "pgId", "Geol. modificada:", "Support_Modified", 6, 1
End Sub

' TAB 2 - MORPHOLOGY, MASONRY AND PHASES
' v11: Lost_Body_Evidence is gone (its information is now a
' T_LOST_ELEMENTS row scoped Body), and the lintel and chamber-roof
' type fields moved to 11.Sist to sit inside their systems (12bis).
Private Sub FillArq(f As String)
    SH f, "pgArq", "Morfologia general", 0
    PCT f, "pgArq", "Cossos basals (N0):",     "N_Basal_Bodies",      1, 1
    PCT f, "pgArq", "Cossos cambra (N1):",     "N_Chamber_Bodies",    2, 1
    PCV f, "pgArq", "Planta:",                 "Floor_Plan",          3, 1, "Rectangular;Rectangular;Sub-rectangular;Sub-rectangular;Square;Quadrada;Circular;Circular;Sub-circular;Sub-circular;Trapezoidal;Trapezoidal;Irregular;Irregular;ND;Indeterminada"
    PCT f, "pgArq", "Murs construits:",        "N_Built_Walls",       4, 1
    PCT f, "pgArq", "Cota sobre la base (m):", "Height_Above_Base_m", 5, 1
    SH  f, "pgArq", "Facana i paisatge (observacional)", 6
    PCV f, "pgArq", "Orientacio facana:", "Facade_Orientation", 7, 1, "N;N;NE;NE;E;E;SE;SE;S;S;SW;SO;W;O;NW;NO;ND;Indeterminada"
    PCV f, "pgArq", "Visibilitat vall:",  "Visibility_Valley",  7, 2, "High;Alta;Medium;Mitjana;Low;Baixa;ND;Indeterminada"
    SH  f, "pgArq", "Maconeria i morter (T&A 2017 / H01, H04)", 8
    PCV f, "pgArq", "Qualitat maconeria:", "Masonry_Quality", 9, 1, "Good;Bona;Moderate;Moderada;Poor;Pobra;ND;Tipus indeterminat"
    PCV f, "pgArq", "Tipus aparell:",      "Masonry_Type",    9, 2, "Well-coursed;Filades regulars;Irregular-coursed;Filades irregulars;Uncoursed;Sense filades;Mixed;Mixt;ND;Tipus indeterminat"
    PC9 f, "pgArq", "Morter present:",     "Mortar_Present",  10, 1
    PCV f, "pgArq", "Tipus morter:",       "Mortar_Type",     10, 2, "Mud;Fang;Mud with gravel;Fang amb grava;Mud with organics;Fang amb organics;None dry-laid;Cap, en sec;ND;Tipus indeterminat"
    PC9 f, "pgArq", "Ripio / falques:",    "Chinking_Stones", 11, 1
    PCT f, "pgArq", "Notes morter:",       "Mortar_Notes",    11, 2
    SH  f, "pgArq", "Fases constructives (H03/H04)", 12
    PCT f, "pgArq", "Num. fases:",     "Construction_Phases", 13, 1
    PCV f, "pgArq", "Evidencia fase:", "Phase_Evidence",      13, 2, "C14;C14;Stratigraphy;Estratigrafia;Superposition;Superposicio;Mortar;Morter;ND;Indeterminada"
End Sub

' TAB 3 - SURFACE TREATMENTS
' Active for every record class (4.4): a rock art panel is DEFINED by
' its pigment, and a structural trace can keep pigment on the corbel.
Private Sub FillAcab(f As String)
    SH f, "pgAcab", "Revoc (lluit)", 0
    PC9 f, "pgAcab", "Revoc present:",  "Plaster_Present", 1, 1
    PCV f, "pgAcab", "Color revoc:",    "Plaster_Color",   2, 1, "White;Blanc;Cream;Crema;Red;Roig;Ochre;Ocre;Grey;Gris;ND;Tipus indeterminat"
    PCV f, "pgAcab", "Extensio revoc:", "Plaster_Extent",  3, 1, "Full facade;Facana sencera;Partial;Parcial;Traces only;Nomes traces;ND;Tipus indeterminat"
    SH f, "pgAcab", "Pigment aplicat", 4
    PC9 f, "pgAcab", "Pigment present:",  "Pigment_Present",   5, 1
    PCV f, "pgAcab", "Substrat pigment:", "Pigment_Substrate", 6, 1, "Plaster;Revoc;Masonry stone;Pedra de parament;Bedrock;Penya;Mixed;Mixt;ND;Tipus indeterminat"
    PCV f, "pgAcab", "Color pigment:",    "Pigment_Color",     5, 2, "Red;Roig;White;Blanc;Both;Ambdos;Ochre;Ocre;ND;Tipus indeterminat"
    ' v11: Perimeter/threshold added (5.8). Perimeter pigment appears in
    ' two structurally distinct contexts (EA11 razed, EA09 natural
    ' cavity); combined with substrate = Bedrock it isolates a practice
    ' of marking the threshold independently of the support type.
    PCV f, "pgAcab", "Extensio pigment:", "Pigment_Extent",    6, 2, "Whole facade;Facana sencera;Architectural elements;Elements arquitectonics;Decorative motifs;Motius decoratius;Perimeter/threshold;Perimetral / llindar;Traces;Traces;ND;Tipus indeterminat"
End Sub

' TAB 4 - DECORATION: the subform only (7.4)
Private Sub FillDec(f As String)
    SH f, "pgDec", "Registres de decoracio (T_DECORATIONS)", 0
End Sub

' TAB 5 - CONSERVATION AND OBSERVABILITY
Private Sub FillEst(f As String)
    SH f, "pgEst", "Estat de conservacio i alteracions", 0
    PCC f, "pgEst", "Estat estructura:",      "ID_Arch_Status",     1, 1
    PCC f, "pgEst", "Estat vestigis mobles:", "ID_Material_Status", 2, 1
    PC9 f, "pgEst", "Saquejada:",             "Looting",            3, 1
    PC9 f, "pgEst", "Dany per foc:",          "Fire_Damage",        4, 1
    PC9 f, "pgEst", "Activitat animal:",      "Animal_Activity",    3, 2
    PC9 f, "pgEst", "Acces modern:",          "Modern_Access",      4, 2
    SH  f, "pgEst", "Base documental i observabilitat (OE1)", 6
    PCV f, "pgEst", "Base documental:",  "Doc_Basis",              7, 1, "Direct access;Acces directe;Close-range photogrammetry;Fotogrametria proxima;Distant photogrammetry;Fotogrametria distant;Ground photography;Fotografia de terra;Published source;Font publicada;ND;Indeterminada"
    PCV f, "pgEst", "Observ. facana:",   "Facade_Observability",   8, 1, "Complete;Completa;Partial;Parcial;Poor;Deficient;ND;Indeterminada"
    PCV f, "pgEst", "Observ. interior:", "Interior_Observability", 8, 2, "Complete;Completa;Partial;Parcial;None;Nul-la;ND;Indeterminada"
End Sub

' TAB 6 - BIOARCHAEOLOGY. Human_Remains gates the rest (4.6).
Private Sub FillBio(f As String)
    SH f, "pgBio", "Context bioarqueologic", 0
    PC9 f, "pgBio", "Restes humanes:",     "Human_Remains",         1, 1
    PCT f, "pgBio", "MNI:",                "MNI",                   2, 1
    PC9 f, "pgBio", "Connexio anatomica:", "Anatomical_Connection", 3, 1
    PC9 f, "pgBio", "Mumificacio:",        "Mummification",         4, 1
    PC9 f, "pgBio", "Farcells funeraris:", "Funerary_Bundles",      5, 1
    PC9 f, "pgBio", "Restes disperses:",   "Dispersed_Remains",     1, 2
    PC9 f, "pgBio", "Posicio flexada:",    "Flexed_Position",       2, 2
    PC9 f, "pgBio", "Os cremat:",          "Bone_Burning",          3, 2
End Sub

' TAB 7 - CULTURAL MATERIALS. ID_Material_Status is already the gate,
' so no new field was needed (4.6).
Private Sub FillMat(f As String)
    SH f, "pgMat", "Materials culturals", 0
    PC9 f, "pgMat", "Textils:",          "Mat_Textiles",   1, 1
    PC9 f, "pgMat", "Fusta cultural:",   "Mat_Wood",       2, 1
    PC9 f, "pgMat", "Fibra vegetal:",    "Mat_VegFiber",   3, 1
    PC9 f, "pgMat", "Ceramica:",         "Mat_Ceramics",   4, 1
    PC9 f, "pgMat", "Fauna:",            "Mat_Fauna",      1, 2
    PC9 f, "pgMat", "Banya de cervol:",  "Mat_DeerAntler", 2, 2
    PC9 f, "pgMat", "Altres materials:", "Mat_Other",      3, 2
End Sub

' TAB 8 - CHRONOLOGY
Private Sub FillCron(f As String)
    SH  f, "pgCron", "Cronologia", 0
    PCB f, "pgCron", "Datacio C14:",      "C14",               1, 1
    PCT f, "pgCron", "Inici (segle dC):", "Chrono_Start_Cent", 2, 1
    PCT f, "pgCron", "Fi (segle dC):",    "Chrono_End_Cent",   3, 1
    PCC f, "pgCron", "Campanya:",         "ID_Campaign",       4, 1
End Sub

' TAB 9 - METRICS. Opening_Width_cm / Opening_Height_cm are NOT gated
' by Sys_Portal: the metric pass is a separate exercise (4.5).
Private Sub FillMetr(f As String)
    SH  f, "pgMetr", "Dimensions de l'estructura", 0
    PCT f, "pgMetr", "Longitud (m):", "Length_m",   1, 1
    PCT f, "pgMetr", "Amplada (m):",  "Width_m",    1, 2
    PCT f, "pgMetr", "Alcada (m):",   "Height_m",   2, 1
    PCV f, "pgMetr", "Metode dim.:",  "Dim_Method", 2, 2, "Photogrammetric model;Model fotogrametric;Tape measure;Cinta metrica;Laser;Laser;Estimation;Estimacio;perimeter pigment outline;Perimetre de pigment;ND;Indeterminat"
    SH  f, "pgMetr", "Dimensions de l'obertura d'acces (P)", 3
    PCT f, "pgMetr", "Amplada obertura (cm):", "Opening_Width_cm",  4, 1
    PCT f, "pgMetr", "Alcada obertura (cm):",  "Opening_Height_cm", 4, 2
    SH  f, "pgMetr", "Metrica del suport geologic (H02)", 5
    PCT f, "pgMetr", "Amplada suport (cm):", "Support_Width_cm", 6, 1
    PCT f, "pgMetr", "Profunditat (cm):",    "Support_Depth_cm", 6, 2
    SH  f, "pgMetr", "Volumetria i area", 7
    PCT f, "pgMetr", "Area interior (m2):", "Interior_Area_m2", 8, 1
    PCC f, "pgMetr", "Metode calc.:",       "ID_Vol_Method",    8, 2
    PCT f, "pgMetr", "Vol. interior (m3):", "Interior_Vol_m3",  9, 1
    PCT f, "pgMetr", "Notes vol.:",         "Vol_Notes",        9, 2
    PCT f, "pgMetr", "Vol. total (m3):",    "Total_Vol_m3",     10, 1
    SH  f, "pgMetr", "Coordenades espacials (Metashape / GPS)", 11
    PCT f, "pgMetr", "Lat WGS84:",      "Coord_Lat_WGS84",   12, 1
    PCT f, "pgMetr", "E UTM (m):",      "Coord_E_UTM",       12, 2
    PCT f, "pgMetr", "Lon WGS84:",      "Coord_Lon_WGS84",   13, 1
    PCT f, "pgMetr", "N UTM (m):",      "Coord_N_UTM",       13, 2
    PCT f, "pgMetr", "Altitud (msnm):", "Altitude_masl",     14, 1
    PCT f, "pgMetr", "Precisio (m):",   "Coord_Precision_m", 14, 2
    PCC f, "pgMetr", "Metode coord.:",  "ID_Coord_Method",   15, 1
End Sub

' TAB 10 - DIGITAL DOCUMENTATION
Private Sub FillDoc(f As String)
    SH  f, "pgDoc", "Documentacio digital", 0
    PCT f, "pgDoc", "URL Panorama 360:", "URL_Pano",         1, 1
    PCT f, "pgDoc", "URL Panorama 2:",   "URL_Pano_2",       2, 1
    PCT f, "pgDoc", "URL Gigafoto:",     "URL_Giga",         3, 1
    PCT f, "pgDoc", "URL Model 3D:",     "URL_3D",           4, 1
    PCB f, "pgDoc", "Publicat ChaXR:",   "ChaXR_Documented", 5, 1
    PCM f, "pgDoc", "Notes:",            "Notes",            6, 1
End Sub

' ================================================================
'  TAB 11 - CONSTRUCTIVE SYSTEMS
'  Reorganised around the systems (4.5): the five Sys_* combos come
'  first and each component group follows the system that governs it,
'  because that is the order the decisions are actually made in - you
'  decide there is a platform before you count its corbels.
'  I (Interbody_Cornice) and R (Upper_Crown) belong to no system and
'  stay permanently active, deliberately.
' ================================================================
Private Sub FillSys(f As String)
    SH  f, "pgSys", "Sistemes constructius - resolre primer", 0
    PCS f, "pgSys", "Sistema plataforma (H):", "Sys_Platform", 1, 1
    PCS f, "pgSys", "Sistema portal (P):",     "Sys_Portal",   1, 2
    PCS f, "pgSys", "Sistema rafec (U):",      "Sys_Eave",     2, 1
    PCG f, "pgSys", "Conjunt basal (A-D):",    "Sys_Base",     2, 2
    PCG f, "pgSys", "Conjunt cambra:",         "Sys_Chamber",  3, 1

    SH  f, "pgSys", "Conjunt basal: A B C D", 4
    PC5 f, "pgSys", "Jaceres basals (A):",    "Embedded_Base_Beams", 5, 1
    PC5 f, "pgSys", "Basament (B):",          "Base_Level",          5, 2
    PC5 f, "pgSys", "Socol decoratiu (C):",   "Decorative_Socle",    6, 1
    PC5 f, "pgSys", "Muret transversal (D):", "Tie_Walls",           6, 2

    SH  f, "pgSys", "Sistema plataforma: E + F + G", 7
    PC5 f, "pgSys", "Mensules fusta (E):",      "Timber_Brackets",      8, 1
    PCT f, "pgSys", "Num. mensules:",           "Timber_Bracket_Count", 8, 2
    PC5 f, "pgSys", "Bigues transversals (F):", "Transverse_Beams",     9, 1
    PC5 f, "pgSys", "Filades en voladis (G):",  "Corbelled_Courses",    9, 2
    PCV f, "pgSys", "Mat. superficie:",         "Platform_Surface_Material", 10, 1, "Timber;Fusta;Stone;Pedra;Mixed;Mixt;ND;Tipus indeterminat"
    PCV f, "pgSys", "Rol mensules (E):",        "Timber_Bracket_Role",  10, 2, "Platform support;Suport de plataforma;Isolated;Aillada;Both;Ambdos;ND;Tipus indeterminat"
    ' Form separated from function (5.3): OE3 exists to determine what
    ' the platform was FOR, so the morphology field must not presume it.
    PCV f, "pgSys", "Funcio plataforma:",       "Platform_Function",    11, 1, "Access;Acces;Circulation;Circulacio;Construction;Bastida constructiva;Support;Base de suport;Multiple;Multiple;Undetermined;Indeterminada"

    SH  f, "pgSys", "Interficie N0/N1 - sense gating", 12
    PC5 f, "pgSys", "Cornisa intercos (I):", "Interbody_Cornice",          13, 1
    PCV f, "pgSys", "Mat. cornisa (I):",     "Interbody_Cornice_Material", 13, 2, "Stone slabs;Lloses de pedra;Wooden beams;Bigues de fusta;Mixed;Mixt;ND;Tipus indeterminat"
    PC5 f, "pgSys", "Coronament (R):",       "Upper_Crown",                14, 1

    SH  f, "pgSys", "Sistema portal: N + O + Q", 15
    PC5 f, "pgSys", "Llindar (N):",     "Sill",            16, 1
    PC5 f, "pgSys", "Brancals (O):",    "Jambs",           16, 2
    PC5 f, "pgSys", "Dintell (Q):",     "Lintel",          17, 1
    PCV f, "pgSys", "Mat. dintell:",    "Lintel_Material", 17, 2, "Stone;Pedra;Wood;Fusta;Mixed;Mixt;ND;Tipus indeterminat"
    PC9 f, "pgSys", "Marc reculat:",    "Recessed_Frame",  18, 1

    SH  f, "pgSys", "Conjunt cambra: J K L M V X", 19
    PC5 f, "pgSys", "Cantoneres (J):",         "Corner_Quoins",        20, 1
    PC5 f, "pgSys", "Pilastres estruct. (K):", "Structural_Pilasters", 20, 2
    PC5 f, "pgSys", "Ala de facana (L):",      "Facade_Flank",         21, 1
    PC5 f, "pgSys", "Fris en relleu (M):",     "Relief_Frieze",        21, 2
    PC5 f, "pgSys", "Mur de retorn (V):",      "Return_Wall",          22, 1
    PC5 f, "pgSys", "Coberta cambra (X):",     "Chamber_Roof",         22, 2
    PCV f, "pgSys", "Tipus coberta (X):",      "Chamber_Roof_Type",    23, 1, "Natural bedrock;Penya natural;Built masonry;Obra de maconeria;Built timber and slabs;Fusta i lloses;Mixed;Mixt;ND;Tipus indeterminat"
    PCV f, "pgSys", "Tancament posterior:",    "Rear_Closure_Type",    23, 2, "Natural bedrock;Penya natural;Built masonry;Obra de maconeria;Mixed;Mixt;ND;Tipus indeterminat"

    SH  f, "pgSys", "Sistema rafec: S + T", 24
    PC5 f, "pgSys", "Biga suport rafec (S):", "Eave_Beam",    25, 1
    PC5 f, "pgSys", "Superficie rafec (T):",  "Eave_Surface", 25, 2
End Sub

' TAB 12 - CONNECTIONS AND CUSTOM FEATURES
Private Sub FillExtra(f As String)
    SH f, "pgExtra", "Elements personalitzats (T_ARCH_FEATURES)", 0
    SH f, "pgExtra", "Connexions amb altres estructures (T_CONNECTIONS)", 7
End Sub

' ================================================================
'  Table/Query COMBO CONFIGURATION
'  Matched by ControlSource, never by control name: the name
'  assignment in AddCtrl is best-effort, so a failed rename would
'  silently leave a combo unconfigured if we keyed on the name.
' ================================================================
Private Sub ConfigureAllCombos(frmName As String)
    DoCmd.OpenForm frmName, acDesign
    Dim f As Form: Set f = Forms(frmName)

    Dim cs(12) As String
    Dim rs(12) As String
    Dim cc(12) As Integer
    Dim cw(12) As String

    cs(0) = "ID_Sector":          rs(0) = "SELECT ID, Sector_Name FROM L_SECTORS ORDER BY ID_Site, Sector_Name": cc(0) = 2: cw(0) = "0cm;5cm"
    ' Typology carries Record_Class, which drives the tab gating
    cs(1) = "ID_Typology":        rs(1) = "SELECT ID, Name FROM L_TYPOLOGY ORDER BY Name":                       cc(1) = 2: cw(1) = "0cm;6cm"
    cs(2) = "ID_Support":         rs(2) = "SELECT ID, Name FROM L_SUPPORT ORDER BY ID":                          cc(2) = 2: cw(2) = "0cm;5cm"
    cs(3) = "ID_Arch_Status":     rs(3) = "SELECT ID, Name FROM L_STATUS ORDER BY ID":                           cc(3) = 2: cw(3) = "0cm;4cm"
    cs(4) = "ID_Material_Status": rs(4) = "SELECT ID, Name FROM L_MATERIAL_STATUS ORDER BY ID":                  cc(4) = 2: cw(4) = "0cm;5cm"
    cs(5) = "ID_Vol_Method":      rs(5) = "SELECT ID, Name FROM L_VOL_METHOD ORDER BY ID":                       cc(5) = 2: cw(5) = "0cm;5cm"
    cs(6) = "ID_Coord_Method":    rs(6) = "SELECT ID, Name FROM L_COORD_METHOD ORDER BY ID":                     cc(6) = 2: cw(6) = "0cm;5cm"
    cs(7) = "ID_Campaign":        rs(7) = "SELECT ID, Code, Campaign_Name FROM L_CAMPAIGN ORDER BY Code":        cc(7) = 3: cw(7) = "0cm;1.5cm;5cm"
    cs(8) = "ID_Group":           rs(8) = "SELECT ID, Group_Code FROM T_GROUPS ORDER BY Group_Code":             cc(8) = 2: cw(8) = "0cm;4cm"
    cs(9) = "ID_Parent":          rs(9) = "SELECT ID, Code FROM T_STRUCTURES ORDER BY Code":                     cc(9) = 2: cw(9) = "0cm;4cm"
    cs(10) = "ID_Struct_Body":    rs(10) = "SELECT ID, Name FROM L_STRUCT_BODY ORDER BY Level_Type, Name":       cc(10) = 2: cw(10) = "0cm;5cm"
    cs(11) = "ID_Dec_Type":       rs(11) = "SELECT ID, Name FROM L_DEC_TYPE ORDER BY Name":                      cc(11) = 2: cw(11) = "0cm;5cm"
    cs(12) = "ID_Support_Secondary": rs(12) = "SELECT ID, Name FROM L_SUPPORT ORDER BY ID":                      cc(12) = 2: cw(12) = "0cm;5cm"

    Dim ctrl As Control
    Dim i As Integer
    For Each ctrl In f.Controls
        If ctrl.ControlType = acComboBox Then
            For i = 0 To 12
                If ctrl.ControlSource = cs(i) Then
                    ctrl.RowSourceType = "Table/Query"
                    ctrl.RowSource = rs(i)
                    ctrl.BoundColumn = 1
                    ctrl.ColumnCount = cc(i)
                    ctrl.ColumnWidths = cw(i)
                    ctrl.LimitToList = True
                End If
            Next i
        End If
    Next ctrl

    DoCmd.Save acForm, frmName
    DoCmd.Close acForm, frmName
    Debug.Print "[OK] Tots els combos configurats"
End Sub

' ================================================================
'  GATING (4.4, 4.5, 4.6)
'
'  Level 1: Record_Class decides which tabs are active. It lives on
'  L_TYPOLOGY, not on T_STRUCTURES, so it is derived at runtime from
'  ID_Typology - one class per typology, never per record (6.1).
'
'  Level 2: each Sys_* decides whether its component group is
'  editable; ID_Material_Status and Human_Remains do the same for
'  7.Mat and 6.Bio.
'
'  Controls are located by ControlSource rather than by name, for the
'  same reason ConfigureAllCombos does: the name assignment is
'  best-effort and a silent miss would disable the wrong control.
'
'  This is the only part of the build that writes VBA into the form,
'  which Access allows only when "Trust access to the VBA project
'  object model" is enabled. If it is off, the form still works and
'  every field stays editable - only the automatic enabling is lost.
' ================================================================
Private Sub InjectGating(frmName As String)
    On Error GoTo Err_IG

    DoCmd.OpenForm frmName, acDesign
    Dim f As Form: Set f = Forms(frmName)
    f.HasModule = True
    f.Module.AddFromString GatingCode()

    ' Bind the events the generated code answers to.
    f.OnCurrent = "[Event Procedure]"
    SetAfterUpdate f, "ID_Typology"
    SetAfterUpdate f, "ID_Material_Status"
    SetAfterUpdate f, "Human_Remains"
    SetAfterUpdate f, "Sys_Platform"
    SetAfterUpdate f, "Sys_Portal"
    SetAfterUpdate f, "Sys_Eave"
    SetAfterUpdate f, "Sys_Base"
    SetAfterUpdate f, "Sys_Chamber"

    DoCmd.Save acForm, frmName
    DoCmd.Close acForm, frmName
    Debug.Print "[OK] Gating injectat (nivells 1 i 2)"
    Exit Sub

Err_IG:
    Debug.Print "[AVIS] No s'ha pogut injectar el gating: " & Err.Description
    Debug.Print "[AVIS] Activa 'Confiar en l'acces al model d'objectes de projectes VBA'"
    Debug.Print "[AVIS] al Centre de confianca d'Access i torna a executar BuildForm."
    Debug.Print "[AVIS] El formulari funciona igualment; nomes falta l'activacio automatica."
    On Error Resume Next
    DoCmd.Save acForm, frmName
    DoCmd.Close acForm, frmName
    On Error GoTo 0
End Sub

Private Sub SetAfterUpdate(f As Form, src As String)
    Dim c As Control
    On Error Resume Next
    For Each c In f.Controls
        If c.ControlType = acComboBox Or c.ControlType = acTextBox Then
            If c.ControlSource = src Then c.AfterUpdate = "[Event Procedure]"
        End If
    Next c
    On Error GoTo 0
End Sub

' NO Option STATEMENTS HERE. Setting HasModule = True already gives
' the form a module carrying "Option Compare Database", and adding a
' second one is a compile error ("duplicate Option statement").
' Option Explicit is skipped for the same reason - AddFromString
' appends, so it could not be placed legally anyway. Everything the
' generated code uses is explicitly declared, so nothing is lost.
Private Function GatingCode() As String
    Dim c As String
    c = "' Generated by chachapoya_Form_v11_val.bas. Do not edit by" & vbCrLf
    c = c & "' hand: re-running BuildForm replaces this module." & vbCrLf & vbCrLf

    c = c & "Private Sub Form_Current()" & vbCrLf
    c = c & "    ApplyGating" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf

    c = c & "Private Sub ID_Typology_AfterUpdate()" & vbCrLf & "    ApplyGating" & vbCrLf & "End Sub" & vbCrLf & vbCrLf
    c = c & "Private Sub ID_Material_Status_AfterUpdate()" & vbCrLf & "    ApplyGating" & vbCrLf & "End Sub" & vbCrLf & vbCrLf
    c = c & "Private Sub Human_Remains_AfterUpdate()" & vbCrLf & "    ApplyGating" & vbCrLf & "End Sub" & vbCrLf & vbCrLf
    c = c & "Private Sub Sys_Platform_AfterUpdate()" & vbCrLf & "    ApplyGating" & vbCrLf & "End Sub" & vbCrLf & vbCrLf
    c = c & "Private Sub Sys_Portal_AfterUpdate()" & vbCrLf & "    ApplyGating" & vbCrLf & "End Sub" & vbCrLf & vbCrLf
    c = c & "Private Sub Sys_Eave_AfterUpdate()" & vbCrLf & "    ApplyGating" & vbCrLf & "End Sub" & vbCrLf & vbCrLf
    c = c & "Private Sub Sys_Base_AfterUpdate()" & vbCrLf & "    ApplyGating" & vbCrLf & "End Sub" & vbCrLf & vbCrLf
    c = c & "Private Sub Sys_Chamber_AfterUpdate()" & vbCrLf & "    ApplyGating" & vbCrLf & "End Sub" & vbCrLf & vbCrLf

    ' Enables every control bound to a given field.
    c = c & "Private Sub EnSrc(src As String, en As Boolean)" & vbCrLf
    c = c & "    Dim ct As Control" & vbCrLf
    c = c & "    On Error Resume Next" & vbCrLf
    c = c & "    For Each ct In Me.Controls" & vbCrLf
    c = c & "        If ct.ControlType = acComboBox Or ct.ControlType = acTextBox Or ct.ControlType = acCheckBox Then" & vbCrLf
    c = c & "            If ct.ControlSource = src Then ct.Enabled = en" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    Next ct" & vbCrLf
    c = c & "End Sub" & vbCrLf & vbCrLf

    ' A system opens its group only when it is Present* or Attested
    ' lost. Under Absent / Not applicable / Not observable the
    ' components hold the padding 0 of section 1.2 and must not be
    ' edited - that is what keeps rule 4 satisfiable.
    c = c & "Private Function SysOpen(v As Variant) As Boolean" & vbCrLf
    c = c & "    If IsNull(v) Then" & vbCrLf
    c = c & "        SysOpen = True" & vbCrLf
    c = c & "    Else" & vbCrLf
    c = c & "        SysOpen = (Left(v, 7) = ""Present"") Or (v = ""Attested lost"")" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Function" & vbCrLf & vbCrLf

    c = c & "Private Sub ApplyGating()" & vbCrLf
    c = c & "    On Error Resume Next" & vbCrLf
    c = c & "    Dim rc As Variant" & vbCrLf
    c = c & "    rc = Null" & vbCrLf
    c = c & "    If Not IsNull(Me!ID_Typology) Then" & vbCrLf
    c = c & "        rc = DLookup(""Record_Class"", ""L_TYPOLOGY"", ""ID="" & Me!ID_Typology)" & vbCrLf
    c = c & "    End If" & vbCrLf & vbCrLf

    ' Level 1. An unclassified record keeps everything open: closing
    ' tabs on a record whose class is unknown would hide data the
    ' researcher may need to enter in order to classify it.
    c = c & "    Dim isBuilt As Boolean, isNat As Boolean, isTrace As Boolean, isArt As Boolean" & vbCrLf
    c = c & "    isBuilt = (rc = ""Built funerary structure"")" & vbCrLf
    c = c & "    isNat = (rc = ""Natural funerary context"")" & vbCrLf
    c = c & "    isTrace = (rc = ""Structural trace"")" & vbCrLf
    c = c & "    isArt = (rc = ""Rock art panel"")" & vbCrLf & vbCrLf
    c = c & "    Me!tabMain.Pages(""pgArq"").Enabled = Not (isNat Or isTrace Or isArt)" & vbCrLf
    c = c & "    Me!tabMain.Pages(""pgBio"").Enabled = Not (isTrace Or isArt)" & vbCrLf
    c = c & "    Me!tabMain.Pages(""pgMat"").Enabled = Not (isTrace Or isArt)" & vbCrLf
    c = c & "    Me!tabMain.Pages(""pgSys"").Enabled = Not isArt" & vbCrLf & vbCrLf

    ' Level 2: systems over their component groups (table 4.5).
    c = c & "    Dim b As Boolean" & vbCrLf
    c = c & "    b = SysOpen(Me!Sys_Base)" & vbCrLf
    c = c & "    EnSrc ""Embedded_Base_Beams"", b" & vbCrLf
    c = c & "    EnSrc ""Base_Level"", b" & vbCrLf
    c = c & "    EnSrc ""Decorative_Socle"", b" & vbCrLf
    c = c & "    EnSrc ""Tie_Walls"", b" & vbCrLf & vbCrLf
    c = c & "    b = SysOpen(Me!Sys_Platform)" & vbCrLf
    c = c & "    EnSrc ""Timber_Brackets"", b" & vbCrLf
    c = c & "    EnSrc ""Timber_Bracket_Count"", b" & vbCrLf
    c = c & "    EnSrc ""Timber_Bracket_Role"", b" & vbCrLf
    c = c & "    EnSrc ""Transverse_Beams"", b" & vbCrLf
    c = c & "    EnSrc ""Corbelled_Courses"", b" & vbCrLf
    c = c & "    EnSrc ""Platform_Surface_Material"", b" & vbCrLf
    c = c & "    EnSrc ""Platform_Function"", b" & vbCrLf & vbCrLf
    c = c & "    b = SysOpen(Me!Sys_Portal)" & vbCrLf
    c = c & "    EnSrc ""Sill"", b" & vbCrLf
    c = c & "    EnSrc ""Jambs"", b" & vbCrLf
    c = c & "    EnSrc ""Lintel"", b" & vbCrLf
    c = c & "    EnSrc ""Lintel_Material"", b" & vbCrLf
    c = c & "    EnSrc ""Recessed_Frame"", b" & vbCrLf & vbCrLf
    c = c & "    b = SysOpen(Me!Sys_Eave)" & vbCrLf
    c = c & "    EnSrc ""Eave_Beam"", b" & vbCrLf
    c = c & "    EnSrc ""Eave_Surface"", b" & vbCrLf & vbCrLf
    c = c & "    b = SysOpen(Me!Sys_Chamber)" & vbCrLf
    c = c & "    EnSrc ""Corner_Quoins"", b" & vbCrLf
    c = c & "    EnSrc ""Structural_Pilasters"", b" & vbCrLf
    c = c & "    EnSrc ""Facade_Flank"", b" & vbCrLf
    c = c & "    EnSrc ""Relief_Frieze"", b" & vbCrLf
    c = c & "    EnSrc ""Return_Wall"", b" & vbCrLf
    c = c & "    EnSrc ""Chamber_Roof"", b" & vbCrLf
    c = c & "    EnSrc ""Chamber_Roof_Type"", b" & vbCrLf
    c = c & "    EnSrc ""Rear_Closure_Type"", b" & vbCrLf & vbCrLf

    ' Level 2 outside 11.Sist (4.6). ID_Material_Status is already the
    ' gate for 7.Mat, and Human_Remains for the rest of 6.Bio, so
    ' neither needed a field of its own.
    c = c & "    Dim ms As Variant" & vbCrLf
    c = c & "    ms = Null" & vbCrLf
    c = c & "    If Not IsNull(Me!ID_Material_Status) Then" & vbCrLf
    c = c & "        ms = DLookup(""Name"", ""L_MATERIAL_STATUS"", ""ID="" & Me!ID_Material_Status)" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "    b = Not (ms = ""Absent"" Or ms = ""ND"")" & vbCrLf
    c = c & "    EnSrc ""Mat_Textiles"", b" & vbCrLf
    c = c & "    EnSrc ""Mat_Wood"", b" & vbCrLf
    c = c & "    EnSrc ""Mat_VegFiber"", b" & vbCrLf
    c = c & "    EnSrc ""Mat_Ceramics"", b" & vbCrLf
    c = c & "    EnSrc ""Mat_Fauna"", b" & vbCrLf
    c = c & "    EnSrc ""Mat_DeerAntler"", b" & vbCrLf
    c = c & "    EnSrc ""Mat_Other"", b" & vbCrLf & vbCrLf
    c = c & "    b = IsNull(Me!Human_Remains) Or Me!Human_Remains = 1" & vbCrLf
    c = c & "    EnSrc ""MNI"", b" & vbCrLf
    c = c & "    EnSrc ""Anatomical_Connection"", b" & vbCrLf
    c = c & "    EnSrc ""Mummification"", b" & vbCrLf
    c = c & "    EnSrc ""Funerary_Bundles"", b" & vbCrLf
    c = c & "    EnSrc ""Dispersed_Remains"", b" & vbCrLf
    c = c & "    EnSrc ""Flexed_Position"", b" & vbCrLf
    c = c & "    EnSrc ""Bone_Burning"", b" & vbCrLf
    c = c & "End Sub" & vbCrLf

    GatingCode = c
End Function
