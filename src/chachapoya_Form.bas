Option Compare Database
Option Explicit

' ================================================================
'  CHACHAPOYA FORM BUILD SCRIPT v3 (VALENCIAN) - F_STRUCTURES
'  Versio consolidada del formulari d'entrada de dades.
'
'  PRINCIPI: etiquetes (UI) en valencia | valors emmagatzemats en angles
'
'  Canvis respecte a chachapoya_Form_v2_val.bas:
'   - 12 pestanyes (absorbeix chachapoya_patch_metric.bas, ara OBSOLET):
'     1.Id 2.Arq 3.Acab 4.Dec 5.Estat 6.Bio 7.Mat 8.Cron
'     9.Metr 10.Doc 11.Sist 12.Extra
'   - Pestanya 11.Sist: TOTS els 24 elements A-X amb les lletres
'     CORRECTES del vocabulari normalitzat (abans hi havia un sistema
'     de lletres obsolet). Sistemes: E+F+G->H, N+O+Q->P, S+T->U.
'   - Camps A-X com a combos 0/1/9 (helper PC9):
'     0=Absent / 1=Present / 9=ND (per defecte 9, fixat a la BD)
'   - Etiquetes mes amples (LW 1900->2600) i columna 2 desplacada
'     (C2 4880->5600): elimina tots els textos tallats
'   - Pestanya 4.Dec: camps d'art rupestre reubicats SOTA la seua
'     capcalera i subformulari abaixat (abans tapava Rock_Art i
'     RA_Anthropomorphic, que eren inaccessibles)
'   - Subformularis: etiquetes ADJUNTES als controls (pare = control),
'     de manera que la vista full de dades mostra les capcaleres en
'     valencia (les captions DAO per si soles no ho resolien)
'   - Nous camps v4: morter (4), dimensions obertura (2), Dim_Method,
'     Height_Above_Base_m, cornisa intercos (terminologia cos/nivell)
'   - Combos de domini tancat per a Floor_Plan, Lintel (Q),
'     Access_Orientation, Dim_Method, Mortar_Type
'   - Etiquetes d'estat: "Estat estructura" / "Estat vestigis mobles"
'     (abans (I)/(M), ambigu amb les lletres del vocabulari A-X)
'
'  IMPORTANT: executar DESPRES de chachapoya_DB_v4.bas -> BuildDB()
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

Sub BuildForm()
    CreateSubForms
    CreateMainForm
    SetFieldCaptionsVal
    Dim msg As String
    msg = "F_STRUCTURES v3 (val.) creada amb 12 pestanyes!" & vbCrLf & vbCrLf
    msg = msg & "  Vocabulari A-X amb lletres correctes (24 elements)" & vbCrLf
    msg = msg & "  Camps A-X: combos 0=Absent / 1=Present / 9=ND" & vbCrLf
    msg = msg & "  Subformulari F_DECORATIONS (pestanya 4)" & vbCrLf
    msg = msg & "  Subformulari F_ARCH_FEATURES (pestanya 12)" & vbCrLf
    msg = msg & "  Etiquetes senceres (sense retalls)" & vbCrLf
    msg = msg & "  Capcaleres de subformularis en valencia" & vbCrLf & vbCrLf
    msg = msg & "chachapoya_patch_metric.bas ja NO s'ha d'executar."
    MsgBox msg, vbInformation, "Fet!"
End Sub

' ================================================================
'  CAPCALERES DE COLUMNA EN VISTA FULL DE DADES
'  Les etiquetes adjuntes als controls dels subformularis son la
'  solucio principal; les captions DAO es mantenen com a reforc per
'  a la vista de taula directa (T_DECORATIONS, T_ARCH_FEATURES,
'  T_CONNECTIONS obertes fora del formulari).
' ================================================================
Private Sub SetFieldCaptionsVal()
    Dim db As DAO.Database
    Set db = CurrentDb()
    SetCap db, "T_DECORATIONS", "ID_Struct_Body", "Posicio"
    SetCap db, "T_DECORATIONS", "ID_Dec_Type", "Tipus dec."
    SetCap db, "T_DECORATIONS", "Body_No", "Cos"
    SetCap db, "T_DECORATIONS", "Color", "Color"
    SetCap db, "T_ARCH_FEATURES", "Feature_Code", "Element"
    SetCap db, "T_ARCH_FEATURES", "Present", "Present"
    SetCap db, "T_ARCH_FEATURES", "Feature_Count", "Nombre"
    SetCap db, "T_ARCH_FEATURES", "Material", "Material"
    SetCap db, "T_ARCH_FEATURES", "Notes", "Notes"
    SetCap db, "T_CONNECTIONS", "ID_Struct_A", "Estructura A"
    SetCap db, "T_CONNECTIONS", "ID_Struct_B", "Estructura B"
    SetCap db, "T_CONNECTIONS", "Connection_Type", "Tipus connexio"
    SetCap db, "T_CONNECTIONS", "Confidence", "Confianca"
    SetCap db, "T_CONNECTIONS", "Notes", "Notes"
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

Private Sub PCV(frm As String, pg As String, lbl As String, src As String, row As Integer, col As Integer, vals As String)
    AddCtrl frm, pg, lbl, src, acComboBox, row, col, CW
    Dim ctrl As Control
    On Error Resume Next
    Set ctrl = Forms(frm).Controls(src)
    If Not ctrl Is Nothing Then
        ctrl.RowSourceType = "Value List"
        ctrl.RowSource = vals
    End If
    On Error GoTo 0
End Sub

' NOU v3: combo per a camps BYTE 0/1/9 del vocabulari A-X
' Columna 1 (amagada) = valor emmagatzemat; columna 2 = etiqueta
Private Sub PC9(frm As String, pg As String, lbl As String, src As String, row As Integer, col As Integer)
    AddCtrl frm, pg, lbl, src, acComboBox, row, col, 1500
    Dim ctrl As Control
    On Error Resume Next
    Set ctrl = Forms(frm).Controls(src)
    If Not ctrl Is Nothing Then
        ctrl.RowSourceType = "Value List"
        ctrl.RowSource = "0;Absent;1;Present;9;ND"
        ctrl.ColumnCount = 2
        ctrl.BoundColumn = 1
        ctrl.ColumnWidths = "0cm;2cm"
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
'  SUBFORMULARIS
'  Patro v3: primer el control (amb nom = camp), despres l'etiqueta
'  ADJUNTA (pare = nom del control). En vista full de dades, Access
'  mostra com a capcalera de columna la llegenda de l'etiqueta
'  adjunta; sense etiqueta adjunta mostrava el nom del camp angles.
' ================================================================
Private Sub CreateSubForms()
    CreateDecSubform
    CreateFeatSubform
End Sub

Private Sub CreateDecSubform()
    Const SFRM = "F_DECORATIONS"
    On Error Resume Next: DoCmd.DeleteObject acForm, SFRM: On Error GoTo 0
    Dim f As Form: Set f = CreateForm()
    Dim tmp As String: tmp = f.Name
    f.RecordSource = "T_DECORATIONS"
    f.DefaultView = 2: f.ScrollBars = 2
    f.NavigationButtons = False: f.Width = 10600
    f.Section(acDetail).Height = 400
    Dim T As Long: T = 50: Dim L As Long
    Dim lb As Control

    L = 40
    Dim c1 As Control: Set c1 = CreateControl(tmp, acComboBox, acDetail, "", "", L + 1060, T, 2400, 315)
    c1.ControlSource = "ID_Struct_Body"
    c1.RowSourceType = "Table/Query"
    c1.RowSource = "SELECT ID, Name FROM L_STRUCT_BODY ORDER BY Body_Level, Name"
    c1.BoundColumn = 1: c1.ColumnCount = 2: c1.ColumnWidths = "0cm;4cm": c1.LimitToList = True
    On Error Resume Next: c1.Name = "ID_Struct_Body": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "ID_Struct_Body", "", L, T + 15, 1000, 260)
    lb.Caption = "Posicio"

    L = 3600
    Dim c2 As Control: Set c2 = CreateControl(tmp, acComboBox, acDetail, "", "", L + 1060, T, 2200, 315)
    c2.ControlSource = "ID_Dec_Type"
    c2.RowSourceType = "Table/Query"
    c2.RowSource = "SELECT ID, Name FROM L_DEC_TYPE ORDER BY Name"
    c2.BoundColumn = 1: c2.ColumnCount = 2: c2.ColumnWidths = "0cm;4cm": c2.LimitToList = True
    On Error Resume Next: c2.Name = "ID_Dec_Type": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "ID_Dec_Type", "", L, T + 15, 1000, 260)
    lb.Caption = "Tipus dec."

    L = 7000
    Dim c3 As Control: Set c3 = CreateControl(tmp, acTextBox, acDetail, "", "", L + 660, T, 600, 315)
    c3.ControlSource = "Body_No"
    On Error Resume Next: c3.Name = "Body_No": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "Body_No", "", L, T + 15, 600, 260)
    lb.Caption = "Cos"

    L = 8400
    Dim c4 As Control: Set c4 = CreateControl(tmp, acComboBox, acDetail, "", "", L + 660, T, 1400, 315)
    c4.ControlSource = "Color": c4.RowSourceType = "Value List"
    c4.RowSource = "Red;White;Both;None;ND": c4.LimitToList = False
    On Error Resume Next: c4.Name = "Color": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "Color", "", L, T + 15, 600, 260)
    lb.Caption = "Color"

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
    c1.ControlSource = "Feature_Code": c1.RowSourceType = "Value List"
    c1.RowSource = """Pigment trace"";""Unusual bond"";""Textile fixation"";""Wooden peg"";""Other"""
    c1.LimitToList = False
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
    c4.RowSource = "Stone;Timber;Mixed;ND": c4.LimitToList = False
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

' ================================================================
'  FORMULARI PRINCIPAL F_STRUCTURES
' ================================================================
Private Sub CreateMainForm()
    Const FRM = "F_STRUCTURES"
    On Error Resume Next: DoCmd.DeleteObject acForm, FRM: On Error GoTo 0

    Dim f As Form: Set f = CreateForm()
    Dim tmp As String: tmp = f.Name

    f.RecordSource = "T_STRUCTURES"
    f.DefaultView = 0: f.ScrollBars = 3
    f.NavigationButtons = True
    f.Caption = "Registre Estructura v3 - La Petaca i Diablo Wasi (PALP)"
    f.Width = FW

    f.Section(acDetail).Height = 8900
    f.Section(acDetail).BackColor = RGB(249, 249, 248)

    Dim h As Control
    Set h = CreateControl(tmp, acLabel, acDetail, "", "", 120, 80, 7000, 480)
    h.Caption = "REGISTRE D'ESTRUCTURA v3  -  La Petaca i Diablo Wasi (PALP)"
    h.FontSize = 13: h.FontBold = True
    h.ForeColor = RGB(26, 60, 107): h.BackStyle = 0: h.BorderStyle = 0

    ' Control de pestanyes
    Dim tc As Control
    Set tc = CreateControl(tmp, acTabCtl, acDetail, "", "", 60, 620, 13080, 8200)
    tc.Name = "tabMain"

    ' 12 pestanyes
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

    ' Omplir cada pestanya
    FillId tmp: FillArq tmp: FillAcab tmp: FillDec tmp: FillEst tmp
    FillBio tmp: FillMat tmp: FillCron tmp: FillMetr tmp: FillDoc tmp
    FillSys tmp: FillExtra tmp

    ' Subformulari decoracions (pestanya 4) - v3: abaixat perque no
    ' tape els camps d'art rupestre (abans a la fila del subformulari)
    Dim sf1 As Control
    Set sf1 = CreateControl(tmp, acSubform, acDetail, "pgDec", "", C1, MT + 12 * RG + 340, 10800, 1700)
    sf1.SourceObject = "F_DECORATIONS"
    sf1.LinkMasterFields = "ID": sf1.LinkChildFields = "ID_Structure"

    Dim lhDec As Control
    Set lhDec = CreateControl(tmp, acLabel, acDetail, "pgDec", "", C1, MT + 12 * RG, 10800, 260)
    lhDec.Caption = "  REGISTRES DE DECORACIO (T_DECORATIONS)"
    lhDec.BackStyle = 1: lhDec.BackColor = RGB(214, 228, 247)
    lhDec.BorderStyle = 0: lhDec.ForeColor = RGB(26, 60, 107)
    lhDec.FontBold = True: lhDec.FontSize = 8

    ' Subformulari elements arquitectonics (pestanya 12)
    Dim sf2 As Control
    Set sf2 = CreateControl(tmp, acSubform, acDetail, "pgExtra", "", C1, MT + 1 * RG + 340, 12000, 2600)
    sf2.SourceObject = "F_ARCH_FEATURES"
    sf2.LinkMasterFields = "ID": sf2.LinkChildFields = "ID_Structure"

    DoCmd.Save acForm, tmp
    DoCmd.Close acForm, tmp
    DoCmd.Rename FRM, acForm, tmp
    Debug.Print "[OK] F_STRUCTURES v3 (valencia) creada"

    ConfigureAllCombos FRM
End Sub

' ================================================================
'  CONTINGUT DE LES PESTANYES
' ================================================================

' PESTANYA 1 - IDENTIFICACIO I DETALL DE SUPORT
' -----------------------------------------------
Private Sub FillId(f As String)
    SH f, "pgId", "Identificacio i localitzacio", 0
    PCT f, "pgId", "Codi:",           "Code",        1, 1
    PCC f, "pgId", "Sector:",         "ID_Sector",   2, 1
    PCC f, "pgId", "Tipologia:",      "ID_Typology", 3, 1
    PCC f, "pgId", "Suport geom.:",   "ID_Support",  4, 1
    PCC f, "pgId", "Element pare:",   "ID_Parent",   1, 2
    PCC f, "pgId", "Grup funcional:", "ID_Group",    2, 2
    SH  f, "pgId", "Detall suport geologic (H02)", 5
    PCV f, "pgId", "Morfologia suport:", "Support_Morphology", 6, 1, "Flat ledge;Concave ledge;Fissure;Small cavity;Medium cavity;Large cavity;Vertical no support;ND"
    PCB f, "pgId", "Geol. modificada:",  "Support_Modified",   6, 2
End Sub

' PESTANYA 2 - MORFOLOGIA, MACONERIA I FASES
' --------------------------------------------
Private Sub FillArq(f As String)
    SH f, "pgArq", "Morfologia general", 0
    PCT f, "pgArq", "Num. pisos (cossos):",   "N_Floors",            1, 1
    PCV f, "pgArq", "Planta:",                "Floor_Plan",          2, 1, "Rectangular;Sub-rectangular;Square;Circular;Sub-circular;Trapezoidal;Irregular;ND"
    PCT f, "pgArq", "Murs construits:",       "N_Built_Walls",       3, 1
    PCT f, "pgArq", "Cota sobre la base (m):", "Height_Above_Base_m", 4, 1
    PCV f, "pgArq", "Orient. acces:",         "Access_Orientation",  1, 2, "N;NE;E;SE;S;SW;W;NW;ND"
    PCV f, "pgArq", "Material dintell (Q):",  "Lintel",              2, 2, "Stone;Wood;Mixed;Absent;ND"
    PCB f, "pgArq", "Coberta natural:",       "Natural_Roof",        3, 2
    PCB f, "pgArq", "Contraforts:",           "Buttresses",          4, 2
    PCB f, "pgArq", "Estaques fusta:",        "Wooden_Stakes",       5, 2
    SH  f, "pgArq", "Facana i paisatge (observacional)", 6
    PCV f, "pgArq", "Orientacio facana:", "Facade_Orientation", 7, 1, "N;NE;E;SE;S;SW;W;NW;ND"
    PCV f, "pgArq", "Visibilitat vall:",  "Visibility_Valley",  7, 2, "High;Medium;Low;ND"
    SH  f, "pgArq", "Maconeria i morter (T&A 2017 / H01, H04)", 8
    PCV f, "pgArq", "Qualitat maconeria:", "Masonry_Quality", 9, 1, "Good;Moderate;Poor;ND"
    PCV f, "pgArq", "Tipus aparell:",      "Masonry_Type",    9, 2, "Well-coursed;Irregular-coursed;Uncoursed;Mixed;ND"
    PC9 f, "pgArq", "Morter present:",     "Mortar_Present",  10, 1
    PCV f, "pgArq", "Tipus morter:",       "Mortar_Type",     10, 2, "Mud;Mud with gravel;Mud with organics;None dry-laid;ND"
    PC9 f, "pgArq", "Ripio / falques:",    "Chinking_Stones", 11, 1
    PCT f, "pgArq", "Notes morter:",       "Mortar_Notes",    11, 2
    SH  f, "pgArq", "Fases constructives (H03/H04)", 12
    PCT f, "pgArq", "Num. fases:",     "Construction_Phases", 13, 1
    PCV f, "pgArq", "Evidencia fase:", "Phase_Evidence",      13, 2, "C14;Stratigraphy;Superposition;Mortar;ND"
End Sub

' PESTANYA 3 - ACABATS SUPERFICIALS
' -----------------------------------
Private Sub FillAcab(f As String)
    SH f, "pgAcab", "Acabats superficials", 0
    PCB f, "pgAcab", "Revocada:",         "Plastered",        1, 1
    PCT f, "pgAcab", "Color revoc:",      "Plaster_Color",    2, 1
    PCB f, "pgAcab", "Pintura rupestre:", "Rock_Painting",    3, 1
    PCT f, "pgAcab", "Color pintura:",    "Rock_Paint_Color", 4, 1
End Sub

' PESTANYA 4 - DECORACIO
' v3: art rupestre reubicat sota la seua capcalera; subformulari
' abaixat (fila 12). Dec_Frieze eliminat: el fris es l'element M
' (Relief_Frieze) a la pestanya 11.Sist.
' -----------------------
Private Sub FillDec(f As String)
    SH f, "pgDec", "Decoracio en baix relleu", 0
    PCB f, "pgDec", "Ninxol quadrat:",  "Dec_Square_Niche", 1, 1
    PCB f, "pgDec", "Relleu T:",        "Dec_Relief_T",     2, 1
    PCB f, "pgDec", "Relleu T inv.:",   "Dec_Relief_T_Inv", 3, 1
    PCB f, "pgDec", "Relleu L:",        "Dec_Relief_L",     4, 1
    PCB f, "pgDec", "Relleu L inv.:",   "Dec_Relief_L_Inv", 1, 2
    PCB f, "pgDec", "Zigzag:",          "Dec_Zigzag",       2, 2
    PCB f, "pgDec", "Motiu escalonat:", "Dec_Stepped",      3, 2
    SH f, "pgDec", "Art rupestre associat", 5
    PCB f, "pgDec", "Art rupestre:",  "Rock_Art",           6, 1
    PCB f, "pgDec", "Antropomorf:",   "RA_Anthropomorphic", 7, 1
    PCB f, "pgDec", "Zoomorf:",       "RA_Zoomorphic",      8, 1
    PCB f, "pgDec", "Geometric:",     "RA_Geometric",       6, 2
    PCB f, "pgDec", "Abstract:",      "RA_Abstract",        7, 2
    PCB f, "pgDec", "Escena decap.:", "RA_Decap_Scene",     8, 2
End Sub

' PESTANYA 5 - ESTAT DE CONSERVACIO
' v3: etiquetes sense (I)/(M) per a evitar col.lisio amb les lletres
' del vocabulari A-X
' -----------------------------------
Private Sub FillEst(f As String)
    SH f, "pgEst", "Estat de conservacio i alteracions", 0
    PCC f, "pgEst", "Estat estructura:",       "ID_Arch_Status",     1, 1
    PCC f, "pgEst", "Estat vestigis mobles:",  "ID_Material_Status", 2, 1
    PCB f, "pgEst", "Saquejada:",              "Looting",            3, 1
    PCB f, "pgEst", "Dany per foc:",           "Fire_Damage",        4, 1
    PCB f, "pgEst", "Activitat animal:",       "Animal_Activity",    5, 1
    PCB f, "pgEst", "Acces modern:",           "Modern_Access",      6, 1
End Sub

' PESTANYA 6 - BIOARQUEOLOGIA
' ----------------------------
Private Sub FillBio(f As String)
    SH f, "pgBio", "Context bioarqueologic", 0
    PCB f, "pgBio", "Restes humanes:",      "Human_Remains",         1, 1
    PCT f, "pgBio", "MNI:",                 "MNI",                   2, 1
    PCB f, "pgBio", "Connexio anatomica:",  "Anatomical_Connection", 3, 1
    PCB f, "pgBio", "Mumificacio:",         "Mummification",         4, 1
    PCB f, "pgBio", "Farcells funeraris:",  "Funerary_Bundles",      5, 1
    PCB f, "pgBio", "Restes disperses:",    "Dispersed_Remains",     1, 2
    PCB f, "pgBio", "Posicio flexada:",     "Flexed_Position",       2, 2
    PCB f, "pgBio", "Os cremat:",           "Bone_Burning",          3, 2
End Sub

' PESTANYA 7 - MATERIALS CULTURALS
' ----------------------------------
Private Sub FillMat(f As String)
    SH f, "pgMat", "Materials culturals", 0
    PCB f, "pgMat", "Textils:",          "Mat_Textiles",   1, 1
    PCB f, "pgMat", "Fusta cultural:",   "Mat_Wood",       2, 1
    PCB f, "pgMat", "Fibra vegetal:",    "Mat_VegFiber",   3, 1
    PCB f, "pgMat", "Ceramica:",         "Mat_Ceramics",   4, 1
    PCB f, "pgMat", "Fauna:",            "Mat_Fauna",      1, 2
    PCB f, "pgMat", "Banya de cervol:",  "Mat_DeerAntler", 2, 2
    PCB f, "pgMat", "Altres materials:", "Mat_Other",      3, 2
End Sub

' PESTANYA 8 - CRONOLOGIA
' -------------------------
Private Sub FillCron(f As String)
    SH  f, "pgCron", "Cronologia", 0
    PCB f, "pgCron", "Datacio C14:",      "C14",               1, 1
    PCT f, "pgCron", "Inici (segle dC):", "Chrono_Start_Cent", 2, 1
    PCT f, "pgCron", "Fi (segle dC):",    "Chrono_End_Cent",   3, 1
    PCC f, "pgCron", "Campanya:",         "ID_Campaign",       4, 1
End Sub

' PESTANYA 9 - METRICA (absorbeix el patch metric)
' --------------------------------------------------
Private Sub FillMetr(f As String)
    SH  f, "pgMetr", "Dimensions de l'estructura", 0
    PCT f, "pgMetr", "Longitud (m):", "Length_m",   1, 1
    PCT f, "pgMetr", "Amplada (m):",  "Width_m",    1, 2
    PCT f, "pgMetr", "Alcada (m):",   "Height_m",   2, 1
    PCV f, "pgMetr", "Metode dim.:",  "Dim_Method", 2, 2, "Photogrammetric model;Tape measure;Laser;Estimation;ND"
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

' PESTANYA 10 - DOCUMENTACIO DIGITAL
' ------------------------------------
Private Sub FillDoc(f As String)
    SH  f, "pgDoc", "Documentacio digital", 0
    PCT f, "pgDoc", "URL Panorama 360:", "URL_Pano",         1, 1
    PCT f, "pgDoc", "URL Panorama 2:",   "URL_Pano_2",       2, 1
    PCT f, "pgDoc", "URL Gigafoto:",     "URL_Giga",         3, 1
    PCT f, "pgDoc", "URL Model 3D:",     "URL_3D",           4, 1
    PCB f, "pgDoc", "Publicat ChaXR:",   "ChaXR_Documented", 5, 1
    PCM f, "pgDoc", "Notes:",            "Notes",            6, 1
End Sub

' PESTANYA 11 - SISTEMES CONSTRUCTIUS A-X (24 elements)
' Lletres corregides segons el vocabulari normalitzat.
' Q (dintell) es registra com a material a la pestanya 2
' i es deriva a QRY_13.
' -------------------------------------------------------
Private Sub FillSys(f As String)
    SH  f, "pgSys", "Nivell 0 - elements basals (A-D)", 0
    PC9 f, "pgSys", "Jaceres basals (A):",    "Embedded_Base_Beams", 1, 1
    PC9 f, "pgSys", "Basament (B):",          "Base_Level",          1, 2
    PC9 f, "pgSys", "Socol decoratiu (C):",   "Decorative_Socle",    2, 1
    PC9 f, "pgSys", "Muret transversal (D):", "Tie_Walls",           2, 2
    SH  f, "pgSys", "Sistema plataforma d'acces: E+F+G -> H", 3
    PC9 f, "pgSys", "Mensules fusta (E):",      "Timber_Brackets",      4, 1
    PCT f, "pgSys", "Num. mensules:",           "Timber_Bracket_Count", 4, 2
    PC9 f, "pgSys", "Bigues transversals (F):", "Transverse_Beams",     5, 1
    PC9 f, "pgSys", "Filades en voladis (G):",  "Corbelled_Courses",    5, 2
    PCV f, "pgSys", "Material E/G:",            "Corbel_Material",      6, 1, "Timber;Stone;Mixed;ND"
    PC9 f, "pgSys", "Plataforma acces (H):",    "Corbelled_Platform",   6, 2
    SH  f, "pgSys", "Interficie N0/N1", 7
    PC9 f, "pgSys", "Cornisa intercos (I):", "Interbody_Cornice",          8, 1
    PCV f, "pgSys", "Mat. cornisa (I):",     "Interbody_Cornice_Material", 8, 2, "Stone slabs;Wooden beams;Mixed;ND"
    SH  f, "pgSys", "Nivell 1 - alcat i sistema obertura: N+O+Q -> P", 9
    PC9 f, "pgSys", "Cantoneres (J):",          "Corner_Quoins",        10, 1
    PC9 f, "pgSys", "Pilastres estruct. (K):",  "Structural_Pilasters", 10, 2
    PC9 f, "pgSys", "Paraments laterals (L):",  "Lateral_Wall_Faces",   11, 1
    PC9 f, "pgSys", "Fris en relleu (M):",      "Relief_Frieze",        11, 2
    PC9 f, "pgSys", "Llindar (N):",             "Sill",                 12, 1
    PC9 f, "pgSys", "Brancals (O):",            "Jambs",                12, 2
    PC9 f, "pgSys", "Obertura d'acces (P):",    "Access_Opening",       13, 1
    PC9 f, "pgSys", "Portal enfonsat:",         "Recessed_Portal",      13, 2
    PC9 f, "pgSys", "Murs laterals (V):",       "Lateral_Walls",        14, 1
    PC9 f, "pgSys", "Mur posterior (W):",       "Rear_Wall",            14, 2
    PC9 f, "pgSys", "Mur post. construit:",     "Rear_Wall_Built",      15, 1
    PC9 f, "pgSys", "Coronament (R):",          "Upper_Crown",          15, 2
    SH  f, "pgSys", "Zona superior: S+T -> U", 16
    PC9 f, "pgSys", "Biga suport rafec (S):",  "Eave_Beam",    17, 1
    PC9 f, "pgSys", "Superficie rafec (T):",   "Eave_Surface", 17, 2
    PC9 f, "pgSys", "Rafec-voladis (U):",      "Eave",         18, 1
    PC9 f, "pgSys", "Coberta cambra (X):",     "Chamber_Roof", 18, 2
End Sub

' PESTANYA 12 - ELEMENTS PERSONALITZATS
' Els antics "elements addicionals" (A, D, J, V, W, X) ara son a la
' pestanya 11 amb la resta del vocabulari A-X.
' ---------------------------------------
Private Sub FillExtra(f As String)
    SH f, "pgExtra", "Elements personalitzats (T_ARCH_FEATURES)", 0
End Sub

' ================================================================
'  CONFIGURACIO DE COMBOS Table/Query
' ================================================================
Private Sub ConfigureAllCombos(frmName As String)
    DoCmd.OpenForm frmName, acDesign
    Dim f As Form: Set f = Forms(frmName)

    Dim cs(11) As String
    Dim rs(11) As String
    Dim cc(11) As Integer
    Dim cw(11) As String

    cs(0) = "ID_Sector":          rs(0) = "SELECT ID, Sector_Name FROM L_SECTORS ORDER BY ID_Site, Sector_Name": cc(0) = 2: cw(0) = "0cm;5cm"
    cs(1) = "ID_Typology":        rs(1) = "SELECT ID, Name FROM L_TYPOLOGY ORDER BY Name":                       cc(1) = 2: cw(1) = "0cm;6cm"
    cs(2) = "ID_Support":         rs(2) = "SELECT ID, Name FROM L_SUPPORT ORDER BY ID":                          cc(2) = 2: cw(2) = "0cm;5cm"
    cs(3) = "ID_Arch_Status":     rs(3) = "SELECT ID, Name FROM L_STATUS ORDER BY ID":                           cc(3) = 2: cw(3) = "0cm;4cm"
    cs(4) = "ID_Material_Status": rs(4) = "SELECT ID, Name FROM L_MATERIAL_STATUS ORDER BY ID":                  cc(4) = 2: cw(4) = "0cm;5cm"
    cs(5) = "ID_Vol_Method":      rs(5) = "SELECT ID, Name FROM L_VOL_METHOD ORDER BY ID":                       cc(5) = 2: cw(5) = "0cm;5cm"
    cs(6) = "ID_Coord_Method":    rs(6) = "SELECT ID, Name FROM L_COORD_METHOD ORDER BY ID":                     cc(6) = 2: cw(6) = "0cm;5cm"
    cs(7) = "ID_Campaign":        rs(7) = "SELECT ID, Code, Campaign_Name FROM L_CAMPAIGN ORDER BY Code":        cc(7) = 3: cw(7) = "0cm;1.5cm;5cm"
    cs(8) = "ID_Group":           rs(8) = "SELECT ID, Group_Code FROM T_GROUPS ORDER BY Group_Code":             cc(8) = 2: cw(8) = "0cm;4cm"
    cs(9) = "ID_Parent":          rs(9) = "SELECT ID, Code FROM T_STRUCTURES ORDER BY Code":                     cc(9) = 2: cw(9) = "0cm;4cm"
    cs(10) = "ID_Struct_Body":    rs(10) = "SELECT ID, Name FROM L_STRUCT_BODY ORDER BY Body_Level, Name":       cc(10) = 2: cw(10) = "0cm;5cm"
    cs(11) = "ID_Dec_Type":       rs(11) = "SELECT ID, Name FROM L_DEC_TYPE ORDER BY Name":                      cc(11) = 2: cw(11) = "0cm;5cm"

    Dim ctrl As Control
    Dim i As Integer
    For Each ctrl In f.Controls
        If ctrl.ControlType = acComboBox Then
            For i = 0 To 11
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
