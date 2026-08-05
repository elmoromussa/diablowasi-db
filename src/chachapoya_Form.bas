Option Compare Database
Option Explicit

' ================================================================
'  CHACHAPOYA FORM BUILD SCRIPT v12 (VALENCIAN) - F_STRUCTURES
'  Author: Esteve Ribera Torro | TFM Arqueologia UA
'  Spec: DELTA_v10_v11.md (rev. 5) + v12 addendum (gating estes)
'
'  PRINCIPLE: labels (UI) in Valencian | stored values in English
'
'  IMPORTANT: run AFTER the schema exists, i.e. after
'  chachapoya_DB_v12.bas -> BuildDB() on a blank database. To
'  upgrade an existing v11 database WITH DATA use
'  upgrade_form_v12.bas -> UpgradeV12() instead: it adds the v12
'  schema pieces and injects this same gating without a rebuild.
'
'  CHANGES FROM v11 (v12)
'
'  A. GATING NIVELL 2 NOU: Plaster_Present, Pigment_Present,
'     Mortar_Present i Dec_Present manen sobre els seus detalls;
'     Dec_Present governa el subformulari de decoracio.
'  B. GATING NIVELL 3: E -> nombre i rol de mensules (R8/R9); rol
'     Isolated bloqueja el material de plataforma (R7); Q -> material
'     del dintell (R10); X -> tipus de coberta (R14); I -> material
'     de cornisa. Les regles passen de detectades a impossibles.
'  C. QUICK-FILL AMB CONFIRMACIO: tancar una porta (sistema, revoc,
'     pigment, vestigis, restes humanes) ofereix escriure el 0 de
'     farciment / el 9 als camps depenents i buidar-ne el detall.
'     Mai s'escriu sense un Si explicit, i mai en navegar.
'  D. F_LOST_ELEMENTS: subformulari nou a 12.Extra amb validacio
'     R21 en viu, i recordatori emergent en marcar un 3 (R2).
'  E. F_CONNECTIONS: validacio R20 en viu (BeforeUpdate) i
'     autoassignacio Contemporary per a junta travada.
'  F. Avis roig a 11.Sist quan l'estat es Collapsed (R6), i el
'     substrat de pigment amaga Plaster quan el revoc es 0 (R13).
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

' Acumulador per a generar els moduls de formulari (gating i
' validacions de subformulari) de manera llegible.
Private mCode As String

Sub BuildForm()
    CreateSubForms
    CreateMainForm
    SetFieldCaptionsVal
    Dim msg As String
    msg = "F_STRUCTURES v12 (val.) creada amb 12 pestanyes!" & vbCrLf & vbCrLf
    msg = msg & "  20 camps d'element amb domini de 5 valors" & vbCrLf
    msg = msg & "  24 camps observacionals amb 0/1/9" & vbCrLf
    msg = msg & "  Tots els combos de domini son de dues columnes:" & vbCrLf
    msg = msg & "  el valor guardat es en angles, l'etiqueta en valencia" & vbCrLf & vbCrLf
    msg = msg & "  11.Sist: els 5 sistemes manen sobre els seus grups" & vbCrLf
    msg = msg & "  4.Dec: nomes el subformulari T_DECORATIONS" & vbCrLf
    msg = msg & "  12.Extra: connexions, elements personalitzats" & vbCrLf
    msg = msg & "  i evidencia dels elements desapareguts (regla 2)" & vbCrLf & vbCrLf
    msg = msg & "  Gating v12: 3 nivells + emplenat rapid confirmat" & vbCrLf
    msg = msg & "  4.Dec: Dec_Present (0/1/9) governa el subformulari" & vbCrLf & vbCrLf
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
    SetCap db, "T_STRUCTURES", "Dec_Present", "Decoracio present"
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
    CreateLostSubform
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

' NOU v12. La regla 2 exigeix una fila d'evidencia per a cada element
' codificat 3 (desaparegut), pero la v11 no donava cap via al
' formulari per a registrar-la: calia obrir la taula a ma. El
' recordatori emergent del gating (en marcar un 3) apunta aci.
Private Sub CreateLostSubform()
    Const SFRM = "F_LOST_ELEMENTS"
    On Error Resume Next: DoCmd.DeleteObject acForm, SFRM: On Error GoTo 0
    Dim f As Form: Set f = CreateForm()
    Dim tmp As String: tmp = f.Name
    f.RecordSource = "T_LOST_ELEMENTS"
    f.DefaultView = 2: f.ScrollBars = 2
    f.NavigationButtons = False: f.Width = 15600
    f.Section(acDetail).Height = 400
    Dim T As Long: T = 50: Dim L As Long
    Dim lb As Control

    L = 40
    Dim c1 As Control: Set c1 = CreateControl(tmp, acComboBox, acDetail, "", "", L + 900, T, 2400, 315)
    c1.ControlSource = "Element_Code"
    c1.RowSourceType = "Table/Query"
    ' Nomes elements reals: un sistema desaparegut es marca al seu
    ' Sys_* (Attested lost) i la fila d'evidencia apunta al component.
    c1.RowSource = "SELECT Code, Name_VAL FROM L_ELEMENTS WHERE Is_System=False ORDER BY ID"
    c1.BoundColumn = 1: c1.ColumnCount = 2: c1.ColumnWidths = "0.8cm;4cm": c1.LimitToList = True
    On Error Resume Next: c1.Name = "Element_Code": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "Element_Code", "", L, T + 15, 840, 260)
    lb.Caption = "Element"

    L = 3600
    Dim c2 As Control: Set c2 = CreateControl(tmp, acComboBox, acDetail, "", "", L + 1400, T, 2600, 315)
    c2.ControlSource = "ID_Evidence_Type"
    c2.RowSourceType = "Table/Query"
    c2.RowSource = "SELECT ID, Name FROM L_LOST_EVIDENCE ORDER BY ID"
    c2.BoundColumn = 1: c2.ColumnCount = 2: c2.ColumnWidths = "0cm;5.5cm": c2.LimitToList = True
    On Error Resume Next: c2.Name = "ID_Evidence_Type": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "ID_Evidence_Type", "", L, T + 15, 1340, 260)
    lb.Caption = "Tipus evidencia"

    L = 7800
    Dim c3 As Control: Set c3 = CreateControl(tmp, acComboBox, acDetail, "", "", L + 700, T, 1900, 315)
    c3.ControlSource = "Evidence_Scope"
    c3.RowSourceType = "Value List"
    c3.RowSource = "Element;Element;Body;Cos;Whole structure;Estructura sencera"
    c3.BoundColumn = 1: c3.ColumnCount = 2: c3.ColumnWidths = "0cm;4cm": c3.LimitToList = True
    On Error Resume Next: c3.Name = "Evidence_Scope": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "Evidence_Scope", "", L, T + 15, 640, 260)
    lb.Caption = "Abast"

    L = 10600
    Dim c4 As Control: Set c4 = CreateControl(tmp, acComboBox, acDetail, "", "", L + 800, T, 1800, 315)
    c4.ControlSource = "ID_Position"
    c4.RowSourceType = "Table/Query"
    c4.RowSource = "SELECT ID, Name FROM L_STRUCT_BODY ORDER BY Level_Type, Name"
    c4.BoundColumn = 1: c4.ColumnCount = 2: c4.ColumnWidths = "0cm;4.5cm": c4.LimitToList = True
    On Error Resume Next: c4.Name = "ID_Position": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "ID_Position", "", L, T + 15, 740, 260)
    lb.Caption = "Posicio"

    L = 13400
    Dim c5 As Control: Set c5 = CreateControl(tmp, acTextBox, acDetail, "", "", L + 660, T, 1500, 315)
    c5.ControlSource = "Notes"
    On Error Resume Next: c5.Name = "Notes": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "Notes", "", L, T + 15, 600, 260)
    lb.Caption = "Notes"

    DoCmd.Save acForm, tmp: DoCmd.Close acForm, tmp
    DoCmd.Rename SFRM, acForm, tmp
    Debug.Print "[OK] F_LOST_ELEMENTS"
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

    ' v12: mes alt per al tercer subformulari de 12.Extra
    f.Section(acDetail).Height = 11000
    f.Section(acDetail).BackColor = RGB(249, 249, 248)

    Dim h As Control
    Set h = CreateControl(tmp, acLabel, acDetail, "", "", 120, 80, 7000, 480)
    h.Caption = "REGISTRE D'ESTRUCTURA v11  -  La Petaca i Diablo Wasi (PALP)"
    h.FontSize = 13: h.FontBold = True
    h.ForeColor = RGB(26, 60, 107): h.BackStyle = 0: h.BorderStyle = 0

    Dim tc As Control
    Set tc = CreateControl(tmp, acTabCtl, acDetail, "", "", 60, 620, 13080, 10300)
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

    ' NOU v12: l'evidencia dels elements desapareguts (regla 2), al
    ' costat de les connexions perque totes dues son taules filles.
    Dim sf4 As Control
    Set sf4 = CreateControl(tmp, acSubform, acDetail, "pgExtra", "", C1, MT + 16 * RG + 140, 12000, 2400)
    On Error Resume Next: sf4.Name = "sfLost": On Error GoTo 0
    sf4.SourceObject = "F_LOST_ELEMENTS"
    sf4.LinkMasterFields = "ID": sf4.LinkChildFields = "ID_Structure"

    DoCmd.Save acForm, tmp
    DoCmd.Close acForm, tmp
    DoCmd.Rename FRM, acForm, tmp
    Debug.Print "[OK] F_STRUCTURES v11 (valencia) creada"

    ConfigureAllCombos FRM
    InjectGating FRM
    InjectSubformValidations
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
    ' v12: el judici agregat 0/1/9 que la retirada dels booleans havia
    ' deixat orfe. 0 o 9 desactiven el subformulari; les regles 22-23
    ' vigilen la coherencia amb les files de T_DECORATIONS.
    PC9 f, "pgDec", "Decoracio present:", "Dec_Present", 1, 1
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

    ' v12: avis de la regla 6, visible nomes quan ID_Arch_Status es
    ' Collapsed (ho commuta el gating).
    Dim lb As Control
    Set lb = CreateControl(f, acLabel, acDetail, "pgSys", "", C2, MT + 3 * RG, 7200, 280)
    lb.Name = "lblColl"
    lb.Caption = "AVIS: estructura colapsada - el 0 (absent) no es verificable. Useu 3 (amb evidencia) o 9."
    lb.ForeColor = RGB(180, 30, 30)
    lb.FontBold = True
    lb.BackStyle = 0
    lb.BorderStyle = 0
    lb.Visible = False
End Sub

' TAB 12 - CONNECTIONS AND CUSTOM FEATURES
Private Sub FillExtra(f As String)
    SH f, "pgExtra", "Elements personalitzats (T_ARCH_FEATURES)", 0
    SH f, "pgExtra", "Connexions amb altres estructures (T_CONNECTIONS)", 7
    SH f, "pgExtra", "Elements desapareguts - evidencia (T_LOST_ELEMENTS, regla 2)", 15
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
'  GATING v12 (tres nivells + quick-fill)
'
'  Nivell 1: Record_Class (derivat d'ID_Typology) mana sobre les
'  pestanyes. Nivell 2: cada porta (Sys_*, ID_Material_Status,
'  Human_Remains, Plaster/Pigment/Mortar/Dec_Present) mana sobre el
'  seu grup. Nivell 3: el detall d'un element nomes s'edita si
'  l'element en te (E -> nombre/rol; Q -> material; X -> tipus;
'  I -> material; rol Isolated bloqueja el material de plataforma).
'
'  QUICK-FILL: en tancar una porta s'ofereix escriure el 0 de
'  farciment (seccio 1.2) o el 9 als camps depenents, amb confirmacio
'  previa. Aixi les regles 4, 10, 19, 24 i 25 se satisfan en entrada.
'  L'assignacio per codi no dispara AfterUpdate: no hi ha recursio, i
'  Form_Current nomes activa/desactiva, mai escriu (analeg regla B).
'
'  El formulari PREVE; QRY_16 DETECTA. La bateria segueix sent
'  necessaria: taules obertes a ma, importacions o el Centre de
'  confianca desactivat esquiven el gating.
'
'  Requereix "Confiar en l'acces al model d'objectes de projectes
'  VBA". Si esta desactivat, el formulari es construeix igualment i
'  tot queda editable; nomes falta l'automatisme, i el script avisa.
' ================================================================
Private Sub InjectGating(frmName As String)
    On Error GoTo Err_IG

    DoCmd.OpenForm frmName, acDesign
    Dim f As Form: Set f = Forms(frmName)

    BuildGatingV12
    ReplaceModule frmName

    ' Vincula els esdeveniments que el codi generat respon.
    f.OnCurrent = "[Event Procedure]"
    SetAfterUpdate f, "ID_Typology"
    SetAfterUpdate f, "ID_Arch_Status"
    SetAfterUpdate f, "ID_Material_Status"
    SetAfterUpdate f, "Human_Remains"
    SetAfterUpdate f, "Sys_Base"
    SetAfterUpdate f, "Sys_Platform"
    SetAfterUpdate f, "Sys_Portal"
    SetAfterUpdate f, "Sys_Eave"
    SetAfterUpdate f, "Sys_Chamber"
    SetAfterUpdate f, "Plaster_Present"
    SetAfterUpdate f, "Pigment_Present"
    SetAfterUpdate f, "Mortar_Present"
    SetAfterUpdate f, "Dec_Present"
    SetAfterUpdate f, "Timber_Bracket_Role"

    Dim el(19) As String
    FillElems el
    Dim i As Integer
    For i = 0 To 19
        SetAfterUpdate f, el(i)
    Next i

    DoCmd.Save acForm, frmName
    DoCmd.Close acForm, frmName
    Debug.Print "[OK] Gating v12 injectat (nivells 1, 2 i 3 + quick-fill)"
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

' Validacions en viu dels subformularis. Als fulls de dades, Enabled
' afecta la columna sencera de totes les files, aixi que la logica
' condicional per fila ha de ser validacio BeforeUpdate, no gating.
Private Sub InjectSubformValidations()
    On Error GoTo Err_IS

    ' R21: amb abast Element cal el codi de l'element.
    DoCmd.OpenForm "F_LOST_ELEMENTS", acDesign
    mCode = ""
    LG "' Generat per chachapoya_Form_v12_val.bas. No editar a ma."
    LG "Option Compare Database"
    LG ""
    LG "Private Sub Form_BeforeUpdate(Cancel As Integer)"
    LG "    If Nz(Me!Evidence_Scope, """") = ""Element"" And IsNull(Me!Element_Code) Then"
    LG "        MsgBox ""Amb abast Element cal indicar el codi de l'element (regla 21)."", vbExclamation"
    LG "        Cancel = True"
    LG "    End If"
    LG "End Sub"
    ReplaceModule "F_LOST_ELEMENTS"
    Forms("F_LOST_ELEMENTS").BeforeUpdate = "[Event Procedure]"
    DoCmd.Save acForm, "F_LOST_ELEMENTS"
    DoCmd.Close acForm, "F_LOST_ELEMENTS"

    ' R20: la direccio cronologica nomes es llegeix d'una junta
    ' vertical adossada o d'una superposicio; una junta travada
    ' implica contemporaneitat i s'autoassigna si el camp esta buit.
    DoCmd.OpenForm "F_CONNECTIONS", acDesign
    Dim fc As Form: Set fc = Forms("F_CONNECTIONS")
    mCode = ""
    LG "' Generat per chachapoya_Form_v12_val.bas. No editar a ma."
    LG "Option Compare Database"
    LG ""
    LG "Private Sub Form_BeforeUpdate(Cancel As Integer)"
    LG "    Dim cr As String"
    LG "    cr = Nz(Me!Chrono_Relation, """")"
    LG "    If cr = ""A earlier than B"" Or cr = ""B earlier than A"" Then"
    LG "        Dim ct As String"
    LG "        ct = Nz(Me!Connection_Type, """")"
    LG "        If ct <> ""Abutted vertical joint"" And ct <> ""Superposition"" Then"
    LG "            MsgBox ""Una relacio direccional nomes es llegeix d'una junta vertical adossada o d'una superposicio (regla 20). Corregiu el tipus de connexio o marqueu Undetermined."", vbExclamation"
    LG "            Cancel = True"
    LG "        End If"
    LG "    End If"
    LG "End Sub"
    LG ""
    LG "Private Sub Connection_Type_AfterUpdate()"
    LG "    If Nz(Me!Connection_Type, """") = ""Bonded joint"" Then"
    LG "        If IsNull(Me!Chrono_Relation) Then Me!Chrono_Relation = ""Contemporary"""
    LG "    End If"
    LG "End Sub"
    ReplaceModule "F_CONNECTIONS"
    fc.BeforeUpdate = "[Event Procedure]"
    SetAfterUpdate fc, "Connection_Type"
    DoCmd.Save acForm, "F_CONNECTIONS"
    DoCmd.Close acForm, "F_CONNECTIONS"

    Debug.Print "[OK] Validacions R20 i R21 injectades als subformularis"
    Exit Sub

Err_IS:
    Debug.Print "[AVIS] No s'han pogut injectar les validacions de subformulari: " & Err.Description
    On Error Resume Next
    DoCmd.Save acForm, "F_LOST_ELEMENTS"
    DoCmd.Close acForm, "F_LOST_ELEMENTS"
    DoCmd.Save acForm, "F_CONNECTIONS"
    DoCmd.Close acForm, "F_CONNECTIONS"
    On Error GoTo 0
End Sub

Private Sub SetAfterUpdate(f As Form, src As String)
    Dim c As Control
    On Error Resume Next
    For Each c In f.Controls
        If c.ControlType = acComboBox Or c.ControlType = acTextBox Or c.ControlType = acCheckBox Then
            If c.ControlSource = src Then c.AfterUpdate = "[Event Procedure]"
        End If
    Next c
    On Error GoTo 0
End Sub

' Afig una linia al codi de modul en construccio.
Private Sub LG(s As String)
    mCode = mCode & s & vbCrLf
End Sub

' Substitueix el modul sencer d'un formulari (obert en disseny) pel
' codi acumulat a mCode. DeleteLines + InsertLines permet incloure les
' sentencies Option al codi generat, cosa que AddFromString no permet
' (el modul nou ja porta Option Compare i es duplicaria).
Private Sub ReplaceModule(frmName As String)
    Dim f As Form
    Set f = Forms(frmName)
    f.HasModule = True
    With f.Module
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        .InsertLines 1, mCode
    End With
End Sub

' Els 20 camps d'element del vocabulari A-X (sense H, P, U, W).
Private Sub FillElems(el() As String)
    el(0) = "Embedded_Base_Beams"
    el(1) = "Base_Level"
    el(2) = "Decorative_Socle"
    el(3) = "Tie_Walls"
    el(4) = "Timber_Brackets"
    el(5) = "Transverse_Beams"
    el(6) = "Corbelled_Courses"
    el(7) = "Interbody_Cornice"
    el(8) = "Corner_Quoins"
    el(9) = "Structural_Pilasters"
    el(10) = "Facade_Flank"
    el(11) = "Relief_Frieze"
    el(12) = "Sill"
    el(13) = "Jambs"
    el(14) = "Lintel"
    el(15) = "Upper_Crown"
    el(16) = "Eave_Beam"
    el(17) = "Eave_Surface"
    el(18) = "Return_Wall"
    el(19) = "Chamber_Roof"
End Sub

Private Sub BuildGatingV12()
    mCode = ""
    LG "' ============================================================"
    LG "' MODUL F_STRUCTURES v12 - generat per upgrade_form_v12.bas."
    LG "' No editar a ma: tornar a executar UpgradeV12 el substitueix."
    LG "' ============================================================"
    LG "Option Compare Database"
    LG ""
    LG "Private Sub Form_Current()"
    LG "    ApplyGating"
    LG "End Sub"
    LG ""
    LG "Private Sub ID_Typology_AfterUpdate()"
    LG "    ApplyGating"
    LG "End Sub"
    LG ""
    LG "Private Sub ID_Arch_Status_AfterUpdate()"
    LG "    ApplyGating"
    LG "End Sub"
    LG ""
    LG "Private Sub Timber_Bracket_Role_AfterUpdate()"
    LG "    ApplyGating"
    LG "End Sub"
    LG ""
    LG "Private Sub ID_Material_Status_AfterUpdate()"
    LG "    QuickFillMaterials"
    LG "    ApplyGating"
    LG "End Sub"
    LG ""
    LG "Private Sub Human_Remains_AfterUpdate()"
    LG "    QuickFillBio"
    LG "    ApplyGating"
    LG "End Sub"
    LG ""
    LG "Private Sub Sys_Base_AfterUpdate()"
    LG "    QuickFillSys ""Sys_Base"""
    LG "    ApplyGating"
    LG "End Sub"
    LG ""
    LG "Private Sub Sys_Platform_AfterUpdate()"
    LG "    QuickFillSys ""Sys_Platform"""
    LG "    ApplyGating"
    LG "End Sub"
    LG ""
    LG "Private Sub Sys_Portal_AfterUpdate()"
    LG "    QuickFillSys ""Sys_Portal"""
    LG "    ApplyGating"
    LG "End Sub"
    LG ""
    LG "Private Sub Sys_Eave_AfterUpdate()"
    LG "    QuickFillSys ""Sys_Eave"""
    LG "    ApplyGating"
    LG "End Sub"
    LG ""
    LG "Private Sub Sys_Chamber_AfterUpdate()"
    LG "    QuickFillSys ""Sys_Chamber"""
    LG "    ApplyGating"
    LG "End Sub"
    LG ""
    LG "Private Sub Plaster_Present_AfterUpdate()"
    LG "    If Nz(Me!Plaster_Present, 1) <> 1 Then FillGroup """", 0, ""Plaster_Color,Plaster_Extent"""
    LG "    ApplyGating"
    LG "End Sub"
    LG ""
    LG "Private Sub Pigment_Present_AfterUpdate()"
    LG "    If Nz(Me!Pigment_Present, 1) <> 1 Then FillGroup """", 0, ""Pigment_Substrate,Pigment_Color,Pigment_Extent"""
    LG "    ApplyGating"
    LG "End Sub"
    LG ""
    LG "Private Sub Mortar_Present_AfterUpdate()"
    LG "    Dim v As Variant"
    LG "    v = Me!Mortar_Present"
    LG "    If Not IsNull(v) Then"
    LG "        If v = 0 Then"
    LG "            If Nz(Me!Mortar_Type, """") <> ""None dry-laid"" Then Me!Mortar_Type = ""None dry-laid"""
    LG "        ElseIf v = 9 Then"
    LG "            If Not IsNull(Me!Mortar_Type) Then Me!Mortar_Type = Null"
    LG "        End If"
    LG "    End If"
    LG "    ApplyGating"
    LG "End Sub"
    LG ""
    LG "Private Sub Dec_Present_AfterUpdate()"
    LG "    If Nz(Me!Dec_Present, 1) <> 1 Then"
    LG "        If Not IsNull(Me!ID) Then"
    LG "            If DCount(""*"", ""T_DECORATIONS"", ""ID_Structure="" & Me!ID) > 0 Then"
    LG "                MsgBox ""Hi ha registres de decoracio per a esta estructura: la regla 23 els marcara mentre Dec_Present no siga 1. No s'esborra res automaticament."", vbExclamation"
    LG "            End If"
    LG "        End If"
    LG "    End If"
    LG "    ApplyGating"
    LG "End Sub"
    LG ""

    ' Els 20 stubs d'element: recordatori del 3 + refresc del gating.
    Dim el(19) As String
    FillElems el
    Dim i As Integer
    For i = 0 To 19
        LG "Private Sub " & el(i) & "_AfterUpdate()"
        LG "    ElemUpd """ & el(i) & """"
        LG "End Sub"
        LG ""
    Next i

    LG "Private Sub ElemUpd(fld As String)"
    LG "    On Error Resume Next"
    LG "    If Nz(Me(fld), 0) = 3 Then"
    LG "        MsgBox ""Heu marcat 3 (desaparegut). Registreu l'evidencia a 12.Extra (elements desapareguts): la regla 2 exigeix una fila per a cada 3."", vbInformation"
    LG "    End If"
    LG "    On Error GoTo 0"
    LG "    ApplyGating"
    LG "End Sub"
    LG ""
    LG "' Activa o desactiva tot control lligat a un camp donat."
    LG "Private Sub EnSrc(src As String, en As Boolean)"
    LG "    Dim ct As Control"
    LG "    On Error Resume Next"
    LG "    For Each ct In Me.Controls"
    LG "        If ct.ControlType = acComboBox Or ct.ControlType = acTextBox Or ct.ControlType = acCheckBox Then"
    LG "            If ct.ControlSource = src Then ct.Enabled = en"
    LG "        End If"
    LG "    Next ct"
    LG "    On Error GoTo 0"
    LG "End Sub"
    LG ""
    LG "' Un sistema obri el seu grup nomes en Present* o Attested lost."
    LG "Private Function SysOpen(v As Variant) As Boolean"
    LG "    If IsNull(v) Then"
    LG "        SysOpen = True"
    LG "    Else"
    LG "        SysOpen = (Left(v, 7) = ""Present"") Or (v = ""Attested lost"")"
    LG "    End If"
    LG "End Function"
    LG ""
    LG "' L'element te (o va tindre) entitat: 1, 2 o 3. NULL mante el"
    LG "' detall editable, perque el judici encara no s'ha fet."
    LG "Private Function ElemHas(fld As String) As Boolean"
    LG "    Dim v As Variant"
    LG "    v = Null"
    LG "    On Error Resume Next"
    LG "    v = Me(fld)"
    LG "    On Error GoTo 0"
    LG "    If IsNull(v) Then"
    LG "        ElemHas = True"
    LG "    Else"
    LG "        ElemHas = (v = 1 Or v = 2 Or v = 3)"
    LG "    End If"
    LG "End Function"
    LG ""
    LG "' QUICK-FILL. En tancar un sistema, ofereix el 0 de farciment"
    LG "' (Absent / No aplicable) o el 9 (No observable) als components,"
    LG "' i buida els camps de detall. Mai s'escriu sense confirmacio."
    LG "Private Sub QuickFillSys(sysField As String)"
    LG "    Dim v As Variant"
    LG "    v = Me(sysField)"
    LG "    If IsNull(v) Then Exit Sub"
    LG "    Dim target As Integer"
    LG "    If v = ""Absent"" Or v = ""Not applicable"" Then"
    LG "        target = 0"
    LG "    ElseIf v = ""Not observable"" Then"
    LG "        target = 9"
    LG "    Else"
    LG "        Exit Sub"
    LG "    End If"
    LG "    Dim comps As String"
    LG "    Dim clears As String"
    LG "    comps = """""
    LG "    clears = """""
    LG "    Select Case sysField"
    LG "        Case ""Sys_Base"""
    LG "            comps = ""Embedded_Base_Beams,Base_Level,Decorative_Socle,Tie_Walls"""
    LG "        Case ""Sys_Platform"""
    LG "            comps = ""Timber_Brackets,Transverse_Beams,Corbelled_Courses"""
    LG "            clears = ""Timber_Bracket_Count,Timber_Bracket_Role,Platform_Surface_Material,Platform_Function"""
    LG "        Case ""Sys_Portal"""
    LG "            comps = ""Sill,Jambs,Lintel,Recessed_Frame"""
    LG "            clears = ""Lintel_Material"""
    LG "        Case ""Sys_Eave"""
    LG "            comps = ""Eave_Beam,Eave_Surface"""
    LG "        Case ""Sys_Chamber"""
    LG "            comps = ""Corner_Quoins,Structural_Pilasters,Facade_Flank,Relief_Frieze,Return_Wall,Chamber_Roof"""
    LG "            clears = ""Chamber_Roof_Type,Rear_Closure_Type"""
    LG "    End Select"
    LG "    FillGroup comps, target, clears"
    LG "End Sub"
    LG ""
    LG "Private Sub QuickFillMaterials()"
    LG "    Dim ms As Variant"
    LG "    ms = Null"
    LG "    If Not IsNull(Me!ID_Material_Status) Then"
    LG "        ms = DLookup(""Name"", ""L_MATERIAL_STATUS"", ""ID="" & Me!ID_Material_Status)"
    LG "    End If"
    LG "    Dim g As String"
    LG "    g = ""Mat_Textiles,Mat_Wood,Mat_VegFiber,Mat_Ceramics,Mat_Fauna,Mat_DeerAntler,Mat_Other"""
    LG "    If Nz(ms, """") = ""Absent"" Then"
    LG "        FillGroup g, 0, """""
    LG "    ElseIf Nz(ms, """") = ""ND"" Then"
    LG "        FillGroup g, 9, """""
    LG "    End If"
    LG "End Sub"
    LG ""
    LG "Private Sub QuickFillBio()"
    LG "    Dim v As Variant"
    LG "    v = Me!Human_Remains"
    LG "    If IsNull(v) Then Exit Sub"
    LG "    Dim g As String"
    LG "    g = ""Anatomical_Connection,Mummification,Funerary_Bundles,Dispersed_Remains,Flexed_Position,Bone_Burning"""
    LG "    If v = 0 Then"
    LG "        FillGroup g, 0, ""MNI"""
    LG "    ElseIf v = 9 Then"
    LG "        FillGroup g, 9, ""MNI"""
    LG "    End If"
    LG "End Sub"
    LG ""
    LG "' Escriu target als camps de comps i Null als de clears, amb"
    LG "' confirmacio previa si algun valor canviaria. L'assignacio per"
    LG "' codi no dispara AfterUpdate, aixi que no hi ha recursio."
    LG "Private Sub FillGroup(comps As String, target As Integer, clears As String)"
    LG "    Dim a() As String"
    LG "    Dim c() As String"
    LG "    Dim i As Integer"
    LG "    Dim n As Integer"
    LG "    n = 0"
    LG "    a = Split(comps, "","")"
    LG "    c = Split(clears, "","")"
    LG "    On Error Resume Next"
    LG "    For i = 0 To UBound(a)"
    LG "        If Nz(Me(a(i)), -1) <> target Then n = n + 1"
    LG "    Next i"
    LG "    For i = 0 To UBound(c)"
    LG "        If Not IsNull(Me(c(i))) Then n = n + 1"
    LG "    Next i"
    LG "    On Error GoTo 0"
    LG "    If n = 0 Then Exit Sub"
    LG "    Dim msg As String"
    LG "    msg = ""Voleu emplenar automaticament "" & n & "" camp(s) del grup amb el valor coherent ("" & target & "" o buit)?"""
    LG "    If MsgBox(msg, vbYesNo + vbQuestion, ""Emplenat rapid"") <> vbYes Then Exit Sub"
    LG "    On Error Resume Next"
    LG "    For i = 0 To UBound(a)"
    LG "        Me(a(i)) = target"
    LG "    Next i"
    LG "    For i = 0 To UBound(c)"
    LG "        Me(c(i)) = Null"
    LG "    Next i"
    LG "    On Error GoTo 0"
    LG "End Sub"
    LG ""
    LG "Private Sub ApplyGating()"
    LG "    On Error Resume Next"
    LG "    Dim rc As Variant"
    LG "    rc = Null"
    LG "    If Not IsNull(Me!ID_Typology) Then"
    LG "        rc = DLookup(""Record_Class"", ""L_TYPOLOGY"", ""ID="" & Me!ID_Typology)"
    LG "    End If"
    LG ""
    LG "    ' Nivell 1: la classe de registre mana sobre les pestanyes."
    LG "    ' Un registre sense classificar ho mante tot obert."
    LG "    Dim isNat As Boolean, isTrace As Boolean, isArt As Boolean"
    LG "    isNat = (Nz(rc, """") = ""Natural funerary context"")"
    LG "    isTrace = (Nz(rc, """") = ""Structural trace"")"
    LG "    isArt = (Nz(rc, """") = ""Rock art panel"")"
    LG "    Me!tabMain.Pages(""pgArq"").Enabled = Not (isNat Or isTrace Or isArt)"
    LG "    Me!tabMain.Pages(""pgBio"").Enabled = Not (isTrace Or isArt)"
    LG "    Me!tabMain.Pages(""pgMat"").Enabled = Not (isTrace Or isArt)"
    LG "    Me!tabMain.Pages(""pgSys"").Enabled = Not isArt"
    LG ""
    LG "    ' Nivell 2: cada sistema mana sobre el seu grup (taula 4.5)."
    LG "    Dim b As Boolean"
    LG "    b = SysOpen(Me!Sys_Base)"
    LG "    EnSrc ""Embedded_Base_Beams"", b"
    LG "    EnSrc ""Base_Level"", b"
    LG "    EnSrc ""Decorative_Socle"", b"
    LG "    EnSrc ""Tie_Walls"", b"
    LG ""
    LG "    b = SysOpen(Me!Sys_Platform)"
    LG "    EnSrc ""Timber_Brackets"", b"
    LG "    EnSrc ""Transverse_Beams"", b"
    LG "    EnSrc ""Corbelled_Courses"", b"
    LG "    ' Nivell 3: el detall de les mensules nomes si E en te (R8, R9)"
    LG "    Dim eb As Boolean"
    LG "    eb = b And ElemHas(""Timber_Brackets"")"
    LG "    EnSrc ""Timber_Bracket_Count"", eb"
    LG "    EnSrc ""Timber_Bracket_Role"", eb"
    LG "    ' Una mensula aillada no suporta cap plataforma (R7)"
    LG "    EnSrc ""Platform_Surface_Material"", b And (Nz(Me!Timber_Bracket_Role, """") <> ""Isolated"")"
    LG "    EnSrc ""Platform_Function"", b"
    LG ""
    LG "    b = SysOpen(Me!Sys_Portal)"
    LG "    EnSrc ""Sill"", b"
    LG "    EnSrc ""Jambs"", b"
    LG "    EnSrc ""Lintel"", b"
    LG "    ' Un dintell absent no te material (R10)"
    LG "    EnSrc ""Lintel_Material"", b And ElemHas(""Lintel"")"
    LG "    EnSrc ""Recessed_Frame"", b"
    LG ""
    LG "    b = SysOpen(Me!Sys_Eave)"
    LG "    EnSrc ""Eave_Beam"", b"
    LG "    EnSrc ""Eave_Surface"", b"
    LG ""
    LG "    b = SysOpen(Me!Sys_Chamber)"
    LG "    EnSrc ""Corner_Quoins"", b"
    LG "    EnSrc ""Structural_Pilasters"", b"
    LG "    EnSrc ""Facade_Flank"", b"
    LG "    EnSrc ""Relief_Frieze"", b"
    LG "    EnSrc ""Return_Wall"", b"
    LG "    EnSrc ""Chamber_Roof"", b"
    LG "    ' El tipus de coberta nomes si X en te (R14)"
    LG "    EnSrc ""Chamber_Roof_Type"", b And ElemHas(""Chamber_Roof"")"
    LG "    EnSrc ""Rear_Closure_Type"", b"
    LG ""
    LG "    ' I i R no pertanyen a cap sistema; el material de la cornisa"
    LG "    ' nomes si I en te."
    LG "    EnSrc ""Interbody_Cornice_Material"", ElemHas(""Interbody_Cornice"")"
    LG ""
    LG "    ' Acabats: presencia mana sobre el detall (R24, R25)."
    LG "    Dim pOn As Boolean"
    LG "    pOn = IsNull(Me!Plaster_Present) Or Nz(Me!Plaster_Present, 1) = 1"
    LG "    EnSrc ""Plaster_Color"", pOn"
    LG "    EnSrc ""Plaster_Extent"", pOn"
    LG "    Dim gOn As Boolean"
    LG "    gOn = IsNull(Me!Pigment_Present) Or Nz(Me!Pigment_Present, 1) = 1"
    LG "    EnSrc ""Pigment_Substrate"", gOn"
    LG "    EnSrc ""Pigment_Color"", gOn"
    LG "    EnSrc ""Pigment_Extent"", gOn"
    LG ""
    LG "    ' Prevencio de R13: sense revoc verificat (0), el substrat"
    LG "    ' Plaster desapareix de la llista."
    LG "    Dim rsrc As String"
    LG "    rsrc = ""Masonry stone;Pedra de parament;Bedrock;Penya;Mixed;Mixt;ND;Tipus indeterminat"""
    LG "    If IsNull(Me!Plaster_Present) Or Nz(Me!Plaster_Present, 1) <> 0 Then"
    LG "        rsrc = ""Plaster;Revoc;"" & rsrc"
    LG "    End If"
    LG "    Dim ct2 As Control"
    LG "    For Each ct2 In Me.Controls"
    LG "        If ct2.ControlType = acComboBox Then"
    LG "            If ct2.ControlSource = ""Pigment_Substrate"" Then"
    LG "                If ct2.RowSource <> rsrc Then ct2.RowSource = rsrc"
    LG "            End If"
    LG "        End If"
    LG "    Next ct2"
    LG ""
    LG "    ' Morter: el tipus nomes si hi ha morter o encara no s'ha dit."
    LG "    EnSrc ""Mortar_Type"", IsNull(Me!Mortar_Present) Or Nz(Me!Mortar_Present, 1) = 1"
    LG ""
    LG "    ' Vestigis mobles: la porta es ID_Material_Status (4.6)."
    LG "    Dim ms As Variant"
    LG "    ms = Null"
    LG "    If Not IsNull(Me!ID_Material_Status) Then"
    LG "        ms = DLookup(""Name"", ""L_MATERIAL_STATUS"", ""ID="" & Me!ID_Material_Status)"
    LG "    End If"
    LG "    b = Not (Nz(ms, """") = ""Absent"" Or Nz(ms, """") = ""ND"")"
    LG "    EnSrc ""Mat_Textiles"", b"
    LG "    EnSrc ""Mat_Wood"", b"
    LG "    EnSrc ""Mat_VegFiber"", b"
    LG "    EnSrc ""Mat_Ceramics"", b"
    LG "    EnSrc ""Mat_Fauna"", b"
    LG "    EnSrc ""Mat_DeerAntler"", b"
    LG "    EnSrc ""Mat_Other"", b"
    LG ""
    LG "    ' Bioarqueologia: Human_Remains mana sobre el detall (4.6)."
    LG "    b = IsNull(Me!Human_Remains) Or Nz(Me!Human_Remains, 1) = 1"
    LG "    EnSrc ""MNI"", b"
    LG "    EnSrc ""Anatomical_Connection"", b"
    LG "    EnSrc ""Mummification"", b"
    LG "    EnSrc ""Funerary_Bundles"", b"
    LG "    EnSrc ""Dispersed_Remains"", b"
    LG "    EnSrc ""Flexed_Position"", b"
    LG "    EnSrc ""Bone_Burning"", b"
    LG ""
    LG "    ' Decoracio: Dec_Present mana sobre el subformulari."
    LG "    Dim dOn As Boolean"
    LG "    dOn = IsNull(Me!Dec_Present) Or Nz(Me!Dec_Present, 1) = 1"
    LG "    Dim ct3 As Control"
    LG "    For Each ct3 In Me.Controls"
    LG "        If ct3.ControlType = acSubform Then"
    LG "            If ct3.SourceObject = ""F_DECORATIONS"" Then ct3.Enabled = dOn"
    LG "        End If"
    LG "    Next ct3"
    LG ""
    LG "    ' Avis R6: en una estructura colapsada el 0 no es verificable."
    LG "    Dim st As Variant"
    LG "    st = Null"
    LG "    If Not IsNull(Me!ID_Arch_Status) Then"
    LG "        st = DLookup(""Name"", ""L_STATUS"", ""ID="" & Me!ID_Arch_Status)"
    LG "    End If"
    LG "    Me!lblColl.Visible = (Nz(st, """") = ""Collapsed"")"
    LG "End Sub"
End Sub
