Option Compare Database
Option Explicit

' ================================================================
'  CHACHAPOYA FORM BUILD SCRIPT v2 (VALENCIAN) - F_STRUCTURES
'  Versio en valencia del formulari d'entrada de dades.
'
'  Diferencies respecte a chachapoya_02_Form_v2.bas (angles):
'   - Totes les etiquetes de pestanyes, seccions i camps en valencia
'   - Llistes de valors (PCV) dels camps de text lliure en valencia:
'       Masonry_Quality: Bona/Moderada/Deficient/ND
'       Masonry_Type: Regular/Irregular/Sense cursar/Mixta/ND
'       Support_Morphology: Replana plana/Replana concava/...
'       Visibility_Valley: Alta/Mitja/Baixa/ND
'       Phase_Evidence: C14/Estratigrafia/Superposicio/Morter/ND
'       Corbel_Material: Fusta/Pedra/Mixt/ND
'       Cornice_Material: Lloses pedra/Bigues fusta/Mixt/ND
'   - Els ControlSource (noms de camp de la BD) resten en angles
'   - Els combos Table/Query mostren les taules lookup (en angles)
'   - Diacritics omesos en el codi per compatibilitat ASCII/cp1252
'
'  IMPORTANT: executar DESPRES de chachapoya_01_DB_v3.bas -> BuildDB()
' ================================================================

' Layout constants
Const MT  As Long = 550
Const RG  As Long = 420
Const C1  As Long = 120
Const C2  As Long = 4880
Const LW  As Long = 1900
Const CW  As Long = 2500
Const CH  As Long = 315
Const LH  As Long = 270
Const FW  As Long = 13200

Sub BuildForm()
    CreateSubForms
    CreateMainForm
    SetFieldCaptionsVal
    Dim msg As String
    msg = "F_STRUCTURES v2 (val.) creada amb 11 pestanyes!" & vbCrLf & vbCrLf
    msg = msg & "  Subformulari F_DECORATIONS (pestanya 4)" & vbCrLf
    msg = msg & "  Subformulari F_ARCH_FEATURES (pestanya 11)" & vbCrLf
    msg = msg & "  Tots els combos configurats" & vbCrLf
    msg = msg & "  Capcaleres de columna en valencia (T_DECORATIONS i T_ARCH_FEATURES)" & vbCrLf & vbCrLf
    msg = msg & "Camps nous en v2:" & vbCrLf
    msg = msg & "  Pestanya 1: Detall suport geologic (4 camps)" & vbCrLf
    msg = msg & "  Pestanya 2: Facana i paisatge (2 camps)" & vbCrLf
    msg = msg & "  Pestanya 2: Qualitat mamposteria (2 camps)" & vbCrLf
    msg = msg & "  Pestanya 2: Fases constructives (2 camps)"
    MsgBox msg, vbInformation, "Fet!"
End Sub

' ================================================================
'  CAPCALERES DE COLUMNA EN VISTA FULL DE DADES (DATASHEET VIEW)
'
'  En vista full de dades, Access mostra com a capcalera de columna
'  la propietat Caption del camp de la TableDef corresponent.
'  Si aquesta propietat no esta definida, mostra el nom del camp.
'  SetFieldCaptionsVal l'estableix via DAO per a T_DECORATIONS i
'  T_ARCH_FEATURES de manera que les columnes apareguen en valencia.
'  Nota: NO afecta els noms de camp (sempre en angles al schema).
' ================================================================
Private Sub SetFieldCaptionsVal()
    Dim db As DAO.Database
    Set db = CurrentDb()
    SetCap db, "T_DECORATIONS",  "ID_Struct_Body",  "Posicio"
    SetCap db, "T_DECORATIONS",  "ID_Dec_Type",     "Tipus dec."
    SetCap db, "T_DECORATIONS",  "Body_No",         "Cos"
    SetCap db, "T_DECORATIONS",  "Color",           "Color"
    SetCap db, "T_ARCH_FEATURES","Feature_Code",    "Element"
    SetCap db, "T_ARCH_FEATURES","Present",         "Present"
    SetCap db, "T_ARCH_FEATURES","Feature_Count",   "Nombre"
    SetCap db, "T_ARCH_FEATURES","Material",        "Material"
    SetCap db, "T_ARCH_FEATURES","Notes",           "Notes"
    db.TableDefs.Refresh
    Set db = Nothing
    Debug.Print "[OK] Capcaleres de columna establertes en valencia"
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
'  HELPERS (identics a la versio en angles)
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
    Dim L As Long: Dim T As Long
    If col = 1 Then L = C1 Else L = C2
    T = MT + row * RG
    Dim ctrl As Control
    On Error Resume Next
    Set ctrl = Forms(frm).Controls(src)
    If Not ctrl Is Nothing Then
        ctrl.RowSourceType = "Value List"
        ctrl.RowSource = vals
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
    f.NavigationButtons = False: f.Width = 9400
    f.Section(acDetail).Height = 400
    Dim T As Long: T = 50: Dim L As Long

    L = 40
    Dim lb1 As Control: Set lb1 = CreateControl(tmp, acLabel, acDetail, "", "", L, T+15, 1000, 260)
    lb1.Caption = "Posicio:": lb1.BackStyle = 0: lb1.BorderStyle = 0
    Dim c1 As Control: Set c1 = CreateControl(tmp, acComboBox, acDetail, "", "", L+1060, T, 2400, 315)
    c1.ControlSource = "ID_Struct_Body"
    c1.RowSourceType = "Table/Query"
    c1.RowSource = "SELECT ID, Name FROM L_STRUCT_BODY ORDER BY Body_Level, Name"
    c1.BoundColumn = 1: c1.ColumnCount = 2: c1.ColumnWidths = "0cm;4cm": c1.LimitToList = True
    On Error Resume Next: c1.Name = "ID_Struct_Body": On Error GoTo 0

    L = 3600
    Dim lb2 As Control: Set lb2 = CreateControl(tmp, acLabel, acDetail, "", "", L, T+15, 1000, 260)
    lb2.Caption = "Tipus dec.:": lb2.BackStyle = 0: lb2.BorderStyle = 0
    Dim c2 As Control: Set c2 = CreateControl(tmp, acComboBox, acDetail, "", "", L+1060, T, 2200, 315)
    c2.ControlSource = "ID_Dec_Type"
    c2.RowSourceType = "Table/Query"
    c2.RowSource = "SELECT ID, Name FROM L_DEC_TYPE ORDER BY Name"
    c2.BoundColumn = 1: c2.ColumnCount = 2: c2.ColumnWidths = "0cm;4cm": c2.LimitToList = True
    On Error Resume Next: c2.Name = "ID_Dec_Type": On Error GoTo 0

    L = 7000
    Dim lb3 As Control: Set lb3 = CreateControl(tmp, acLabel, acDetail, "", "", L, T+15, 600, 260)
    lb3.Caption = "Cos:": lb3.BackStyle = 0: lb3.BorderStyle = 0
    Dim c3 As Control: Set c3 = CreateControl(tmp, acTextBox, acDetail, "", "", L+660, T, 600, 315)
    c3.ControlSource = "Body_No"
    On Error Resume Next: c3.Name = "Body_No": On Error GoTo 0

    L = 8000
    Dim lb4 As Control: Set lb4 = CreateControl(tmp, acLabel, acDetail, "", "", L, T+15, 600, 260)
    lb4.Caption = "Color:": lb4.BackStyle = 0: lb4.BorderStyle = 0
    Dim c4 As Control: Set c4 = CreateControl(tmp, acComboBox, acDetail, "", "", L+660, T, 1400, 315)
    c4.ControlSource = "Color": c4.RowSourceType = "Value List"
    c4.RowSource = "Roig;Blanc;Ambdos;Cap;ND": c4.LimitToList = False
    On Error Resume Next: c4.Name = "Color": On Error GoTo 0

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
    f.NavigationButtons = False: f.Width = 9400
    f.Section(acDetail).Height = 400
    Dim T As Long: T = 50: Dim L As Long

    L = 40
    Dim lb1 As Control: Set lb1 = CreateControl(tmp, acLabel, acDetail, "", "", L, T+15, 1000, 260)
    lb1.Caption = "Element:": lb1.BackStyle = 0: lb1.BorderStyle = 0
    Dim c1 As Control: Set c1 = CreateControl(tmp, acComboBox, acDetail, "", "", L+1060, T, 2400, 315)
    c1.ControlSource = "Feature_Code": c1.RowSourceType = "Value List"
    c1.RowSource = """Jaceres basals"";""Muret transversal"";""Cantoneres"";""Trac pigment"";""Lligada inusual"";""Altre"""
    c1.LimitToList = False
    On Error Resume Next: c1.Name = "Feature_Code": On Error GoTo 0

    L = 3560
    Dim lb2 As Control: Set lb2 = CreateControl(tmp, acLabel, acDetail, "", "", L, T+15, 700, 260)
    lb2.Caption = "Present:": lb2.BackStyle = 0: lb2.BorderStyle = 0
    Dim c2 As Control: Set c2 = CreateControl(tmp, acCheckBox, acDetail, "", "", L+760, T+20, 300, 300)
    c2.ControlSource = "Present"
    On Error Resume Next: c2.Name = "Present": On Error GoTo 0

    L = 4700
    Dim lb3 As Control: Set lb3 = CreateControl(tmp, acLabel, acDetail, "", "", L, T+15, 600, 260)
    lb3.Caption = "Nombre:": lb3.BackStyle = 0: lb3.BorderStyle = 0
    Dim c3 As Control: Set c3 = CreateControl(tmp, acTextBox, acDetail, "", "", L+660, T, 700, 315)
    c3.ControlSource = "Feature_Count"
    On Error Resume Next: c3.Name = "Feature_Count": On Error GoTo 0

    L = 6200
    Dim lb4 As Control: Set lb4 = CreateControl(tmp, acLabel, acDetail, "", "", L, T+15, 700, 260)
    lb4.Caption = "Material:": lb4.BackStyle = 0: lb4.BorderStyle = 0
    Dim c4 As Control: Set c4 = CreateControl(tmp, acComboBox, acDetail, "", "", L+760, T, 1200, 315)
    c4.ControlSource = "Material": c4.RowSourceType = "Value List"
    c4.RowSource = "Pedra;Fusta;Mixt;ND": c4.LimitToList = False
    On Error Resume Next: c4.Name = "Material": On Error GoTo 0

    L = 8200
    Dim lb5 As Control: Set lb5 = CreateControl(tmp, acLabel, acDetail, "", "", L, T+15, 600, 260)
    lb5.Caption = "Notes:": lb5.BackStyle = 0: lb5.BorderStyle = 0
    Dim c5 As Control: Set c5 = CreateControl(tmp, acTextBox, acDetail, "", "", L+660, T, 1500, 315)
    c5.ControlSource = "Notes"
    On Error Resume Next: c5.Name = "Notes": On Error GoTo 0

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
    f.Caption = "Registre Estructura v2 - La Petaca i Diablo Wasi (PALP)"
    f.Width = FW

    f.Section(acDetail).Height = 7200
    f.Section(acDetail).BackColor = RGB(249, 249, 248)

    Dim h As Control
    Set h = CreateControl(tmp, acLabel, acDetail, "", "", 120, 80, 7000, 480)
    h.Caption = "REGISTRE D'ESTRUCTURA v2  -  La Petaca i Diablo Wasi (PALP)"
    h.FontSize = 13: h.FontBold = True
    h.ForeColor = RGB(26, 60, 107): h.BackStyle = 0: h.BorderStyle = 0

    ' Control de pestanyes
    Dim tc As Control
    Set tc = CreateControl(tmp, acTabCtl, acDetail, "", "", 60, 620, 13080, 6520)
    tc.Name = "tabMain"

    ' 11 pestanyes
    tc.Pages(0).Name = "pgId":   tc.Pages(0).Caption = "1.Id."
    tc.Pages(1).Name = "pgArq":  tc.Pages(1).Caption = "2.Arq."
    Dim pg As Control
    Set pg = CreateControl(tmp, acPage, acDetail, "tabMain"): pg.Name = "pgAcab": pg.Caption = "3.Acab."
    Set pg = CreateControl(tmp, acPage, acDetail, "tabMain"): pg.Name = "pgDec":  pg.Caption = "4.Dec."
    Set pg = CreateControl(tmp, acPage, acDetail, "tabMain"): pg.Name = "pgEst":  pg.Caption = "5.Estat"
    Set pg = CreateControl(tmp, acPage, acDetail, "tabMain"): pg.Name = "pgBio":  pg.Caption = "6.Bio."
    Set pg = CreateControl(tmp, acPage, acDetail, "tabMain"): pg.Name = "pgMat":  pg.Caption = "7.Mat."
    Set pg = CreateControl(tmp, acPage, acDetail, "tabMain"): pg.Name = "pgCron": pg.Caption = "8.Cron."
    Set pg = CreateControl(tmp, acPage, acDetail, "tabMain"): pg.Name = "pgSIG":  pg.Caption = "9.SIG"
    Set pg = CreateControl(tmp, acPage, acDetail, "tabMain"): pg.Name = "pgSys":  pg.Caption = "10.Sist."
    Set pg = CreateControl(tmp, acPage, acDetail, "tabMain"): pg.Name = "pgExtra":pg.Caption = "11.Extra"

    ' Omplir cada pestanya
    FillId   tmp: FillArq  tmp: FillAcab tmp: FillDec  tmp: FillEst  tmp
    FillBio  tmp: FillMat  tmp: FillCron tmp: FillSIG  tmp
    FillSys  tmp: FillExtra tmp

    ' Subformulari decoracions (pestanya 4)
    Dim sf1 As Control
    Set sf1 = CreateControl(tmp, acSubform, acDetail, "pgDec", "", C1, MT + 10 * RG, 9000, 1800)
    sf1.SourceObject = "F_DECORATIONS"
    sf1.LinkMasterFields = "ID": sf1.LinkChildFields = "ID_Structure"

    Dim lhDec As Control
    Set lhDec = CreateControl(tmp, acLabel, acDetail, "pgDec", "", C1, MT + 10 * RG - 300, 9000, 260)
    lhDec.Caption = "  REGISTRES DE DECORACIO (T_DECORATIONS)"
    lhDec.BackStyle = 1: lhDec.BackColor = RGB(214, 228, 247)
    lhDec.BorderStyle = 0: lhDec.ForeColor = RGB(26, 60, 107)
    lhDec.FontBold = True: lhDec.FontSize = 8

    ' Subformulari elements arquitectonics (pestanya 11)
    Dim sf2 As Control
    Set sf2 = CreateControl(tmp, acSubform, acDetail, "pgExtra", "", C1, MT + 6 * RG, 9000, 2600)
    sf2.SourceObject = "F_ARCH_FEATURES"
    sf2.LinkMasterFields = "ID": sf2.LinkChildFields = "ID_Structure"

    DoCmd.Save acForm, tmp
    DoCmd.Close acForm, tmp
    DoCmd.Rename FRM, acForm, tmp
    Debug.Print "[OK] F_STRUCTURES (valencia) creada"

    ConfigureAllCombos FRM
End Sub

' ================================================================
'  CONTINGUT DE LES PESTANYES
' ================================================================

' PESTANYA 1 - IDENTIFICACIO I DETALL DE SUPORT
' -----------------------------------------------
Private Sub FillId(f As String)
    SH f, "pgId", "Identificacio i localitzacio", 0
    PCT f, "pgId", "Codi:",              "Code",        1, 1
    PCC f, "pgId", "Sector:",            "ID_Sector",   2, 1
    PCC f, "pgId", "Tipologia:",         "ID_Typology", 3, 1
    PCC f, "pgId", "Suport geom.:",      "ID_Support",  4, 1
    PCC f, "pgId", "Element pare:",      "ID_Parent",   1, 2
    PCC f, "pgId", "Grup funcional:",    "ID_Group",    2, 2
    SH  f, "pgId", "Detall suport geologic (H02)", 5
    PCV f, "pgId", "Morfologia suport:", "Support_Morphology", 6, 1, "Replana plana;Replana concava;Fissura;Cavitat petita;Cavitat mitja;Cavitat gran;Vertical sense suport;ND"
    PCB f, "pgId", "Geol. modificada:", "Support_Modified",   6, 2
    PCT f, "pgId", "Amplada suport (cm):", "Support_Width_cm",  7, 1
    PCT f, "pgId", "Profunditat (cm):",    "Support_Depth_cm",  7, 2
End Sub

' PESTANYA 2 - DIMENSIONS, MORFOLOGIA I CAMPS ANALITICS
' -------------------------------------------------------
Private Sub FillArq(f As String)
    SH f, "pgArq", "Dimensions i morfologia", 0
    PCT f, "pgArq", "Num. pisos:",          "N_Floors",           1, 1
    PCT f, "pgArq", "Planta:",              "Floor_Plan",         2, 1
    PCT f, "pgArq", "Murs construits:",     "N_Built_Walls",      3, 1
    PCT f, "pgArq", "Longitud (m):",        "Length_m",           4, 1
    PCT f, "pgArq", "Amplada (m):",         "Width_m",            5, 1
    PCT f, "pgArq", "Alcada (m):",          "Height_m",           6, 1
    PCT f, "pgArq", "Alcada al penya (m):", "Approx_Height_m",    7, 1
    PCT f, "pgArq", "Orient. acces:",       "Access_Orientation",  1, 2
    PCT f, "pgArq", "Material dintell:",    "Lintel",             2, 2
    PCB f, "pgArq", "Coberta natural:",     "Natural_Roof",       3, 2
    PCB f, "pgArq", "Contraforts:",         "Buttresses",         4, 2
    PCB f, "pgArq", "Cornisa internivell:", "Interlevel_Cornice",  5, 2
    PCB f, "pgArq", "Estaques fusta:",      "Wooden_Stakes",      6, 2
    SH  f, "pgArq", "Facana i paisatge (observacional)", 8
    PCV f, "pgArq", "Orientacio facana:",   "Facade_Orientation",  9, 1, "N;NE;E;SE;S;SW;W;NW;ND"
    PCV f, "pgArq", "Visibilitat vall:",    "Visibility_Valley",   9, 2, "Alta;Mitja;Baixa;ND"
    SH  f, "pgArq", "Qualitat mamposteria (T&A 2017)", 10
    PCV f, "pgArq", "Qualitat mamp.:",      "Masonry_Quality",    11, 1, "Bona;Moderada;Deficient;ND"
    PCV f, "pgArq", "Tipus mamp.:",         "Masonry_Type",       11, 2, "Regular;Irregular;Sense cursar;Mixta;ND"
    SH  f, "pgArq", "Fases constructives (H03/H04)", 12
    PCT f, "pgArq", "Num. fases:",          "Construction_Phases", 13, 1
    PCV f, "pgArq", "Evidencia fase:",      "Phase_Evidence",      13, 2, "C14;Estratigrafia;Superposicio;Morter;ND"
End Sub

' PESTANYA 3 - ACABATS SUPERFICIALS
' -----------------------------------
Private Sub FillAcab(f As String)
    SH f, "pgAcab", "Acabats superficials", 0
    PCB f, "pgAcab", "Revocada:",        "Plastered",        1, 1
    PCT f, "pgAcab", "Color revoc:",     "Plaster_Color",    2, 1
    PCB f, "pgAcab", "Pintura rupestre:","Rock_Painting",    3, 1
    PCT f, "pgAcab", "Color pintura:",   "Rock_Paint_Color", 4, 1
End Sub

' PESTANYA 4 - DECORACIO
' -----------------------
Private Sub FillDec(f As String)
    SH f, "pgDec", "Decoracio en baix relleu", 0
    PCB f, "pgDec", "Ninxol quadrat:",  "Dec_Square_Niche",  1, 1
    PCB f, "pgDec", "Relleu T:",        "Dec_Relief_T",      2, 1
    PCB f, "pgDec", "Relleu T inv.:",   "Dec_Relief_T_Inv",  3, 1
    PCB f, "pgDec", "Relleu L:",        "Dec_Relief_L",      4, 1
    PCB f, "pgDec", "Relleu L inv.:",   "Dec_Relief_L_Inv",  5, 1
    PCB f, "pgDec", "Zigzag:",          "Dec_Zigzag",        6, 1
    PCB f, "pgDec", "Motiu escalonat:", "Dec_Stepped",       7, 1
    PCB f, "pgDec", "Greca/fris:",      "Dec_Frieze",        8, 1
    SH f, "pgDec", "Art rupestre associat", 9
    PCB f, "pgDec", "Art rupestre:",    "Rock_Art",           10, 1
    PCB f, "pgDec", "Antropomorf:",     "RA_Anthropomorphic", 11, 1
    PCB f, "pgDec", "Zoomorf:",         "RA_Zoomorphic",      1, 2
    PCB f, "pgDec", "Geometric:",       "RA_Geometric",       2, 2
    PCB f, "pgDec", "Abstract:",        "RA_Abstract",        3, 2
    PCB f, "pgDec", "Escena decap.:",   "RA_Decap_Scene",     4, 2
End Sub

' PESTANYA 5 - ESTAT DE CONSERVACIO
' -----------------------------------
Private Sub FillEst(f As String)
    SH f, "pgEst", "Estat de conservacio i alteracions", 0
    PCC f, "pgEst", "Estat arq. (I):",     "ID_Arch_Status",     1, 1
    PCC f, "pgEst", "Estat material (M):", "ID_Material_Status", 2, 1
    PCB f, "pgEst", "Saquejada:",          "Looting",            3, 1
    PCB f, "pgEst", "Dany per foc:",       "Fire_Damage",        4, 1
    PCB f, "pgEst", "Activitat animal:",   "Animal_Activity",    5, 1
    PCB f, "pgEst", "Acces modern:",       "Modern_Access",      6, 1
End Sub

' PESTANYA 6 - BIOARQUEOLOGIA
' ----------------------------
Private Sub FillBio(f As String)
    SH f, "pgBio", "Context bioarqueologic", 0
    PCB f, "pgBio", "Restes humanes:",     "Human_Remains",        1, 1
    PCT f, "pgBio", "MNI:",               "MNI",                  2, 1
    PCB f, "pgBio", "Connexio anatomica:","Anatomical_Connection", 3, 1
    PCB f, "pgBio", "Mumificacio:",        "Mummification",        4, 1
    PCB f, "pgBio", "Farcells funeraris:", "Funerary_Bundles",     5, 1
    PCB f, "pgBio", "Restes disperses:",  "Dispersed_Remains",    1, 2
    PCB f, "pgBio", "Posicio flexada:",   "Flexed_Position",      2, 2
    PCB f, "pgBio", "Os cremat:",         "Bone_Burning",         3, 2
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

' PESTANYA 8 - CRONOLOGIA I VOLUMETRIA
' --------------------------------------
Private Sub FillCron(f As String)
    SH  f, "pgCron", "Cronologia", 0
    PCB f, "pgCron", "Datacio C14:",        "C14",               1, 1
    PCT f, "pgCron", "Inici (segle dC):",   "Chrono_Start_Cent", 2, 1
    PCT f, "pgCron", "Fi (segle dC):",      "Chrono_End_Cent",   3, 1
    PCC f, "pgCron", "Campanya:",           "ID_Campaign",       4, 1
    SH  f, "pgCron", "Volumetria i area", 5
    PCT f, "pgCron", "Area interior (m2):", "Interior_Area_m2",  6, 1
    PCT f, "pgCron", "Vol. interior (m3):", "Interior_Vol_m3",   7, 1
    PCT f, "pgCron", "Vol. total (m3):",    "Total_Vol_m3",      8, 1
    PCC f, "pgCron", "Metode calc.:",       "ID_Vol_Method",     9, 1
    PCT f, "pgCron", "Notes vol.:",         "Vol_Notes",         6, 2
End Sub

' PESTANYA 9 - SIG I DOCUMENTACIO DIGITAL
' -----------------------------------------
Private Sub FillSIG(f As String)
    SH  f, "pgSIG", "Coordenades espacials (SIG)", 0
    PCT f, "pgSIG", "Lat WGS84:",       "Coord_Lat_WGS84",   1, 1
    PCT f, "pgSIG", "Lon WGS84:",       "Coord_Lon_WGS84",   2, 1
    PCT f, "pgSIG", "E UTM (m):",       "Coord_E_UTM",       3, 1
    PCT f, "pgSIG", "N UTM (m):",       "Coord_N_UTM",       4, 1
    PCT f, "pgSIG", "Altitud (msnm):",  "Altitude_masl",     5, 1
    PCT f, "pgSIG", "Precisio (m):",    "Coord_Precision_m", 6, 1
    PCC f, "pgSIG", "Metode coord.:",   "ID_Coord_Method",   7, 1
    SH  f, "pgSIG", "Documentacio digital", 0
    PCT f, "pgSIG", "URL Panorama 360:","URL_Pano",           1, 2
    PCT f, "pgSIG", "URL Panorama 2:",  "URL_Pano_2",         2, 2
    PCT f, "pgSIG", "URL Gigafoto:",    "URL_Giga",           3, 2
    PCT f, "pgSIG", "URL Model 3D:",    "URL_3D",             4, 2
    PCB f, "pgSIG", "Publicat ChaXR:", "ChaXR_Documented",  5, 2
    PCM f, "pgSIG", "Notes:",          "Notes",             9, 1
End Sub

' PESTANYA 10 - SISTEMES CONSTRUCTIUS A-X
' -----------------------------------------
Private Sub FillSys(f As String)
    SH  f, "pgSys", "Nivell 0 - Sistema base (I)", 0
    PCB f, "pgSys", "Basament (I):",              "Base_Level",          1, 1
    PCB f, "pgSys", "Socol decoratiu (J):",       "Decorative_Socle",    1, 2
    PCB f, "pgSys", "Plataforma en voladis (M):", "Corbelled_Platform",  2, 1
    SH  f, "pgSys", "  Sistema plataforma: K+L+N -> M", 3
    PCB f, "pgSys", "Mensules fusta (K):",        "Timber_Brackets",     4, 1
    PCT f, "pgSys", "Num. mensules:",              "Timber_Bracket_Count",4, 2
    PCB f, "pgSys", "Bigues transversals (L):",   "Transverse_Beams",    5, 1
    PCB f, "pgSys", "Filades en voladis (N):",    "Corbelled_Courses",   5, 2
    PCV f, "pgSys", "Material corbel:",           "Corbel_Material",     6, 1, "Fusta;Pedra;Mixt;ND"
    SH  f, "pgSys", "Sistema obertura: A+B+P -> O", 7
    PCB f, "pgSys", "Obertura d'acces (O):",      "Access_Opening",      8, 1
    PCB f, "pgSys", "Llindar (P):",               "Sill",                8, 2
    PCB f, "pgSys", "Portal enfonsat:",           "Recessed_Portal",     9, 1
    PCB f, "pgSys", "Brancals (O):",              "Jambs",               9, 2
    SH  f, "pgSys", "Facana i zona superior", 10
    PCB f, "pgSys", "Pilastres estructurals (C):","Structural_Pilasters",11, 1
    PCV f, "pgSys", "Mat. cornisa internivell:",  "Cornice_Material",    11, 2, "Lloses pedra;Bigues fusta;Mixt;ND"
    PCB f, "pgSys", "Rafec/voladis (G):",         "Eave",                12, 1
    PCB f, "pgSys", "Biga suport rafec (F):",     "Eave_Beam",           12, 2
    PCB f, "pgSys", "Coronament (H):",            "Upper_Crown",         13, 1
    PCB f, "pgSys", "Superficie rafec (T):",      "Eave_Surface",        13, 2
End Sub

' PESTANYA 11 - ELEMENTS ADDICIONALS
' ------------------------------------
Private Sub FillExtra(f As String)
    SH  f, "pgExtra", "Elements addicionals confirmats", 0
    PCB f, "pgExtra", "Jaceres basals empotrades:", "Embedded_Base_Beams", 1, 1
    PCB f, "pgExtra", "Murets transversals:",       "Tie_Walls",           1, 2
    PCB f, "pgExtra", "Cantoneres:",                "Corner_Quoins",       2, 1
    PCB f, "pgExtra", "Paraments laterals (V):",    "Lateral_Walls",       2, 2
    PCB f, "pgExtra", "Mur posterior (W):",         "Rear_Wall",           3, 1
    PCB f, "pgExtra", "Mur post. - construit:",     "Rear_Wall_Built",     3, 2
    PCB f, "pgExtra", "Coberta cambra (X):",        "Chamber_Roof",        4, 1
    SH  f, "pgExtra", "Elements personalitzats (T_ARCH_FEATURES)", 5
End Sub

' ================================================================
'  CONFIGURACIO DE COMBOS Table/Query (identic a la versio angles)
' ================================================================
Private Sub ConfigureAllCombos(frmName As String)
    DoCmd.OpenForm frmName, acDesign
    Dim f As Form: Set f = Forms(frmName)

    Dim cs(11) As String
    Dim rs(11) As String
    Dim cc(11) As Integer
    Dim cw(11) As String

    cs(0)="ID_Sector":         rs(0)="SELECT ID, Sector_Name FROM L_SECTORS ORDER BY ID_Site, Sector_Name":    cc(0)=2: cw(0)="0cm;5cm"
    cs(1)="ID_Typology":       rs(1)="SELECT ID, Name FROM L_TYPOLOGY ORDER BY Name":                          cc(1)=2: cw(1)="0cm;6cm"
    cs(2)="ID_Support":        rs(2)="SELECT ID, Name FROM L_SUPPORT ORDER BY ID":                             cc(2)=2: cw(2)="0cm;5cm"
    cs(3)="ID_Arch_Status":    rs(3)="SELECT ID, Name FROM L_STATUS ORDER BY ID":                              cc(3)=2: cw(3)="0cm;4cm"
    cs(4)="ID_Material_Status":rs(4)="SELECT ID, Name FROM L_MATERIAL_STATUS ORDER BY ID":                     cc(4)=2: cw(4)="0cm;5cm"
    cs(5)="ID_Vol_Method":     rs(5)="SELECT ID, Name FROM L_VOL_METHOD ORDER BY ID":                          cc(5)=2: cw(5)="0cm;5cm"
    cs(6)="ID_Coord_Method":   rs(6)="SELECT ID, Name FROM L_COORD_METHOD ORDER BY ID":                        cc(6)=2: cw(6)="0cm;5cm"
    cs(7)="ID_Campaign":       rs(7)="SELECT ID, Code, Campaign_Name FROM L_CAMPAIGN ORDER BY Code":           cc(7)=3: cw(7)="0cm;1.5cm;5cm"
    cs(8)="ID_Group":          rs(8)="SELECT ID, Group_Code FROM T_GROUPS ORDER BY Group_Code":                cc(8)=2: cw(8)="0cm;4cm"
    cs(9)="ID_Parent":         rs(9)="SELECT ID, Code FROM T_STRUCTURES ORDER BY Code":                        cc(9)=2: cw(9)="0cm;4cm"
    cs(10)="ID_Struct_Body":   rs(10)="SELECT ID, Name FROM L_STRUCT_BODY ORDER BY Body_Level, Name":          cc(10)=2:cw(10)="0cm;5cm"
    cs(11)="ID_Dec_Type":      rs(11)="SELECT ID, Name FROM L_DEC_TYPE ORDER BY Name":                         cc(11)=2:cw(11)="0cm;5cm"

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
