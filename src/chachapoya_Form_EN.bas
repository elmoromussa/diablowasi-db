Option Compare Database
Option Explicit

' ================================================================
'  CHACHAPOYA FORM BUILD SCRIPT v2 - F_STRUCTURES (11 tabs)
'  Run Sub BuildForm() AFTER chachapoya_01_DB_v2.bas -> BuildDB()
'
'  Changes vs v1:
'   Tab 1 (pgId):  + "Geological support detail" section
'                    Support_Morphology, Support_Modified,
'                    Support_Width_cm, Support_Depth_cm
'   Tab 2 (pgArq): + "Facade & landscape" section
'                    Facade_Orientation, Visibility_Valley
'                  + "Masonry quality" section
'                    Masonry_Quality, Masonry_Type
'                  + "Constructive phases" section
'                    Construction_Phases, Phase_Evidence
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
    Dim msg As String
    msg = "F_STRUCTURES v2 created with 11 tabs!" & vbCrLf & vbCrLf
    msg = msg & "  F_DECORATIONS subform (tab 4)" & vbCrLf
    msg = msg & "  F_ARCH_FEATURES subform (tab 11)" & vbCrLf
    msg = msg & "  All combos configured" & vbCrLf & vbCrLf
    msg = msg & "New fields in v2:" & vbCrLf
    msg = msg & "  Tab 1: Geological support detail (4 fields)" & vbCrLf
    msg = msg & "  Tab 2: Facade & landscape (2 fields)" & vbCrLf
    msg = msg & "  Tab 2: Masonry quality (2 fields)" & vbCrLf
    msg = msg & "  Tab 2: Constructive phases (2 fields)"
    MsgBox msg, vbInformation, "Done!"
End Sub

' ================================================================
'  HELPER FUNCTIONS
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
'  CREATE SUBFORMS
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
    lb1.Caption = "Location:": lb1.BackStyle = 0: lb1.BorderStyle = 0
    Dim c1 As Control: Set c1 = CreateControl(tmp, acComboBox, acDetail, "", "", L+1060, T, 2400, 315)
    c1.ControlSource = "ID_Struct_Body"
    c1.RowSourceType = "Table/Query"
    c1.RowSource = "SELECT ID, Name FROM L_STRUCT_BODY ORDER BY Body_Level, Name"
    c1.BoundColumn = 1: c1.ColumnCount = 2: c1.ColumnWidths = "0cm;4cm": c1.LimitToList = True
    On Error Resume Next: c1.Name = "ID_Struct_Body": On Error GoTo 0

    L = 3600
    Dim lb2 As Control: Set lb2 = CreateControl(tmp, acLabel, acDetail, "", "", L, T+15, 1000, 260)
    lb2.Caption = "Dec. type:": lb2.BackStyle = 0: lb2.BorderStyle = 0
    Dim c2 As Control: Set c2 = CreateControl(tmp, acComboBox, acDetail, "", "", L+1060, T, 2200, 315)
    c2.ControlSource = "ID_Dec_Type"
    c2.RowSourceType = "Table/Query"
    c2.RowSource = "SELECT ID, Name FROM L_DEC_Type ORDER BY Name"
    c2.BoundColumn = 1: c2.ColumnCount = 2: c2.ColumnWidths = "0cm;4cm": c2.LimitToList = True
    On Error Resume Next: c2.Name = "ID_Dec_Type": On Error GoTo 0

    L = 7000
    Dim lb3 As Control: Set lb3 = CreateControl(tmp, acLabel, acDetail, "", "", L, T+15, 600, 260)
    lb3.Caption = "Body:": lb3.BackStyle = 0: lb3.BorderStyle = 0
    Dim c3 As Control: Set c3 = CreateControl(tmp, acTextBox, acDetail, "", "", L+660, T, 600, 315)
    c3.ControlSource = "Body_No"
    On Error Resume Next: c3.Name = "Body_No": On Error GoTo 0

    L = 8000
    Dim lb4 As Control: Set lb4 = CreateControl(tmp, acLabel, acDetail, "", "", L, T+15, 600, 260)
    lb4.Caption = "Color:": lb4.BackStyle = 0: lb4.BorderStyle = 0
    Dim c4 As Control: Set c4 = CreateControl(tmp, acComboBox, acDetail, "", "", L+660, T, 1400, 315)
    c4.ControlSource = "Color": c4.RowSourceType = "Value List"
    c4.RowSource = "Red;White;Both;None;ND": c4.LimitToList = False
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
    lb1.Caption = "Feature:": lb1.BackStyle = 0: lb1.BorderStyle = 0
    Dim c1 As Control: Set c1 = CreateControl(tmp, acComboBox, acDetail, "", "", L+1060, T, 2400, 315)
    c1.ControlSource = "Feature_Code": c1.RowSourceType = "Value List"
    c1.RowSource = """Embedded base beams"";""Tie wall"";""Corner quoins"";""Pigment trace"";""Unusual bond"";""Other"""
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
    lb3.Caption = "Count:": lb3.BackStyle = 0: lb3.BorderStyle = 0
    Dim c3 As Control: Set c3 = CreateControl(tmp, acTextBox, acDetail, "", "", L+660, T, 700, 315)
    c3.ControlSource = "Feature_Count"
    On Error Resume Next: c3.Name = "Feature_Count": On Error GoTo 0

    L = 6200
    Dim lb4 As Control: Set lb4 = CreateControl(tmp, acLabel, acDetail, "", "", L, T+15, 700, 260)
    lb4.Caption = "Material:": lb4.BackStyle = 0: lb4.BorderStyle = 0
    Dim c4 As Control: Set c4 = CreateControl(tmp, acComboBox, acDetail, "", "", L+760, T, 1200, 315)
    c4.ControlSource = "Material": c4.RowSourceType = "Value List"
    c4.RowSource = "Stone;Timber;Mixed;ND": c4.LimitToList = False
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
'  CREATE MAIN FORM F_STRUCTURES
' ================================================================
Private Sub CreateMainForm()
    Const FRM = "F_STRUCTURES"
    On Error Resume Next: DoCmd.DeleteObject acForm, FRM: On Error GoTo 0

    Dim f As Form: Set f = CreateForm()
    Dim tmp As String: tmp = f.Name

    f.RecordSource = "T_STRUCTURES"
    f.DefaultView = 0: f.ScrollBars = 3
    f.NavigationButtons = True
    f.Caption = "Structure Record v2 - La Petaca & Diablo Wasi (PALP)"
    f.Width = FW

    f.Section(acDetail).Height = 7200
    f.Section(acDetail).BackColor = RGB(249, 249, 248)

    Dim h As Control
    Set h = CreateControl(tmp, acLabel, acDetail, "", "", 120, 80, 7000, 480)
    h.Caption = "STRUCTURE RECORD v2  -  La Petaca & Diablo Wasi (PALP)"
    h.FontSize = 13: h.FontBold = True
    h.ForeColor = RGB(26, 60, 107): h.BackStyle = 0: h.BorderStyle = 0

    ' Tab control
    Dim tc As Control
    Set tc = CreateControl(tmp, acTabCtl, acDetail, "", "", 60, 620, 13080, 6520)
    tc.Name = "tabMain"

    ' Pages (2 default + 9 new = 11 total)
    tc.Pages(0).Name = "pgId":   tc.Pages(0).Caption = "1.Ident."
    tc.Pages(1).Name = "pgArq":  tc.Pages(1).Caption = "2.Arq."
    Dim pg As Control
    Set pg = CreateControl(tmp, acPage, acDetail, "tabMain"): pg.Name = "pgAcab": pg.Caption = "3.Acab."
    Set pg = CreateControl(tmp, acPage, acDetail, "tabMain"): pg.Name = "pgDec":  pg.Caption = "4.Dec."
    Set pg = CreateControl(tmp, acPage, acDetail, "tabMain"): pg.Name = "pgEst":  pg.Caption = "5.Estat"
    Set pg = CreateControl(tmp, acPage, acDetail, "tabMain"): pg.Name = "pgBio":  pg.Caption = "6.Bio."
    Set pg = CreateControl(tmp, acPage, acDetail, "tabMain"): pg.Name = "pgMat":  pg.Caption = "7.Mat."
    Set pg = CreateControl(tmp, acPage, acDetail, "tabMain"): pg.Name = "pgCron": pg.Caption = "8.Cron."
    Set pg = CreateControl(tmp, acPage, acDetail, "tabMain"): pg.Name = "pgSIG":  pg.Caption = "9.SIG"
    Set pg = CreateControl(tmp, acPage, acDetail, "tabMain"): pg.Name = "pgSys":  pg.Caption = "10.Systems"
    Set pg = CreateControl(tmp, acPage, acDetail, "tabMain"): pg.Name = "pgExtra":pg.Caption = "11.Extra"

    ' Fill each tab
    FillId   tmp: FillArq  tmp: FillAcab tmp: FillDec  tmp: FillEst  tmp
    FillBio  tmp: FillMat  tmp: FillCron tmp: FillSIG  tmp
    FillSys  tmp: FillExtra tmp

    ' Subform: Decorations (tab 4)
    Dim sf1 As Control
    Set sf1 = CreateControl(tmp, acSubform, acDetail, "pgDec", "", C1, MT + 10 * RG, 9000, 1800)
    sf1.SourceObject = "F_DECORATIONS"
    sf1.LinkMasterFields = "ID": sf1.LinkChildFields = "ID_Structure"

    Dim lhDec As Control
    Set lhDec = CreateControl(tmp, acLabel, acDetail, "pgDec", "", C1, MT + 10 * RG - 300, 9000, 260)
    lhDec.Caption = "  DECORATION DETAIL RECORDS (T_DECORATIONS)"
    lhDec.BackStyle = 1: lhDec.BackColor = RGB(214, 228, 247)
    lhDec.BorderStyle = 0: lhDec.ForeColor = RGB(26, 60, 107)
    lhDec.FontBold = True: lhDec.FontSize = 8

    ' Subform: Arch Features (tab 11)
    Dim sf2 As Control
    Set sf2 = CreateControl(tmp, acSubform, acDetail, "pgExtra", "", C1, MT + 6 * RG, 9000, 2600)
    sf2.SourceObject = "F_ARCH_FEATURES"
    sf2.LinkMasterFields = "ID": sf2.LinkChildFields = "ID_Structure"

    DoCmd.Save acForm, tmp
    DoCmd.Close acForm, tmp
    DoCmd.Rename FRM, acForm, tmp
    Debug.Print "[OK] F_STRUCTURES v2 created"

    ConfigureAllCombos FRM
End Sub

' ================================================================
'  TAB CONTENT
' ================================================================

' TAB 1 - IDENTIFICATION & SUPPORT DETAIL (v2: +4 fields)
' ----------------------------------------------------------------
Private Sub FillId(f As String)
    SH f, "pgId", "Identification & location", 0
    PCT f, "pgId", "Code:",              "Code",        1, 1
    PCC f, "pgId", "Sector:",            "ID_Sector",   2, 1
    PCC f, "pgId", "Typology:",          "ID_Typology", 3, 1
    PCC f, "pgId", "Geom. support:",     "ID_Support",  4, 1
    PCC f, "pgId", "Parent element:",    "ID_Parent",   1, 2
    PCC f, "pgId", "Functional group:",  "ID_Group",    2, 2
    ' v2: Geological support detail (H02)
    SH  f, "pgId", "Geological support detail (H02)", 5
    PCV f, "pgId", "Support morphology:", "Support_Morphology", 6, 1, "Flat ledge;Concave ledge;Fissure;Small cavity;Medium cavity;Large cavity;Vertical no support;ND"
    PCB f, "pgId", "Geol. modified:",     "Support_Modified",   6, 2
    PCT f, "pgId", "Support width (cm):", "Support_Width_cm",   7, 1
    PCT f, "pgId", "Support depth (cm):", "Support_Depth_cm",   7, 2
End Sub

' TAB 2 - DIMENSIONS, MORPHOLOGY & NEW ANALYTICAL FIELDS (v2: +6 fields)
' ----------------------------------------------------------------
Private Sub FillArq(f As String)
    SH f, "pgArq", "Dimensions & morphology", 0
    PCT f, "pgArq", "No. floors:",        "N_Floors",          1, 1
    PCT f, "pgArq", "Floor plan:",        "Floor_Plan",        2, 1
    PCT f, "pgArq", "No. built walls:",   "N_Built_Walls",     3, 1
    PCT f, "pgArq", "Length (m):",        "Length_m",          4, 1
    PCT f, "pgArq", "Width (m):",         "Width_m",           5, 1
    PCT f, "pgArq", "Height (m):",        "Height_m",          6, 1
    PCT f, "pgArq", "Cliff height (m):",  "Approx_Height_m",   7, 1
    PCT f, "pgArq", "Access orientation:","Access_Orientation", 1, 2
    PCT f, "pgArq", "Lintel material:",   "Lintel",            2, 2
    PCB f, "pgArq", "Natural roof:",      "Natural_Roof",      3, 2
    PCB f, "pgArq", "Buttresses:",        "Buttresses",        4, 2
    PCB f, "pgArq", "Interlevel cornice:","Interlevel_Cornice", 5, 2
    PCB f, "pgArq", "Wooden stakes:",     "Wooden_Stakes",     6, 2
    ' v2: Facade & landscape (observational - QGIS analysis later)
    SH  f, "pgArq", "Facade & landscape (observational)", 8
    PCV f, "pgArq", "Facade orientation:", "Facade_Orientation", 9, 1, "N;NE;E;SE;S;SW;W;NW;ND"
    PCV f, "pgArq", "Valley visibility:",  "Visibility_Valley",  9, 2, "High;Medium;Low;ND"
    ' v2: Masonry quality (H01/H04 - cf. Toyne & Anzellini 2017)
    SH  f, "pgArq", "Masonry quality (Toyne & Anzellini 2017)", 10
    PCV f, "pgArq", "Masonry quality:",  "Masonry_Quality", 11, 1, "Good;Moderate;Poor;ND"
    PCV f, "pgArq", "Masonry type:",     "Masonry_Type",    11, 2, "Well-coursed;Irregular-coursed;Uncoursed;Mixed;ND"
    ' v2: Constructive phases (H03: accumulative saturation / H04: operative sequence)
    SH  f, "pgArq", "Constructive phases (H03 / H04)", 12
    PCT f, "pgArq", "N. phases:",      "Construction_Phases", 13, 1
    PCV f, "pgArq", "Phase evidence:", "Phase_Evidence",       13, 2, "C14;Stratigraphy;Superposition;Mortar;ND"
End Sub

' TAB 3 - FINISHES (unchanged)
' ----------------------------------------------------------------
Private Sub FillAcab(f As String)
    SH f, "pgAcab", "Surface finishes", 0
    PCB f, "pgAcab", "Plastered:",        "Plastered",        1, 1
    PCT f, "pgAcab", "Plaster color:",    "Plaster_Color",    2, 1
    PCB f, "pgAcab", "Rock painting:",    "Rock_Painting",    3, 1
    PCT f, "pgAcab", "Rock paint color:", "Rock_Paint_Color", 4, 1
End Sub

' TAB 4 - DECORATION (unchanged)
' ----------------------------------------------------------------
Private Sub FillDec(f As String)
    SH f, "pgDec", "Bas-relief decoration", 0
    PCB f, "pgDec", "Square niche:",    "Dec_Square_Niche",  1, 1
    PCB f, "pgDec", "Relief T:",        "Dec_Relief_T",      2, 1
    PCB f, "pgDec", "Relief T inv.:",   "Dec_Relief_T_Inv",  3, 1
    PCB f, "pgDec", "Relief L:",        "Dec_Relief_L",      4, 1
    PCB f, "pgDec", "Relief L inv.:",   "Dec_Relief_L_Inv",  5, 1
    PCB f, "pgDec", "Zigzag:",          "Dec_Zigzag",        6, 1
    PCB f, "pgDec", "Stepped motif:",   "Dec_Stepped",       7, 1
    PCB f, "pgDec", "Frieze:",          "Dec_Frieze",        8, 1
    SH f, "pgDec", "Associated rock art", 9
    PCB f, "pgDec", "Rock art:",        "Rock_Art",           10, 1
    PCB f, "pgDec", "Anthropomorphic:", "RA_Anthropomorphic", 11, 1
    PCB f, "pgDec", "Zoomorphic:",      "RA_Zoomorphic",      1, 2
    PCB f, "pgDec", "Geometric:",       "RA_Geometric",       2, 2
    PCB f, "pgDec", "Abstract:",        "RA_Abstract",        3, 2
    PCB f, "pgDec", "Decap. scene:",    "RA_Decap_Scene",     4, 2
End Sub

' TAB 5 - CONSERVATION (unchanged)
' ----------------------------------------------------------------
Private Sub FillEst(f As String)
    SH f, "pgEst", "Conservation status & alterations", 0
    PCC f, "pgEst", "Arch. status (I):",    "ID_Arch_Status",     1, 1
    PCC f, "pgEst", "Material status (M):", "ID_Material_Status", 2, 1
    PCB f, "pgEst", "Looting:",             "Looting",            3, 1
    PCB f, "pgEst", "Fire damage:",         "Fire_Damage",        4, 1
    PCB f, "pgEst", "Animal activity:",     "Animal_Activity",    5, 1
    PCB f, "pgEst", "Modern access:",       "Modern_Access",      6, 1
End Sub

' TAB 6 - BIOARCHAEOLOGY (unchanged)
' ----------------------------------------------------------------
Private Sub FillBio(f As String)
    SH f, "pgBio", "Bioarchaeological context", 0
    PCB f, "pgBio", "Human remains:",       "Human_Remains",      1, 1
    PCT f, "pgBio", "MNI:",                 "MNI",                2, 1
    PCB f, "pgBio", "Anatomical connect.:", "Anatomical_Connection",3, 1
    PCB f, "pgBio", "Mummification:",       "Mummification",      4, 1
    PCB f, "pgBio", "Funerary bundles:",    "Funerary_Bundles",   5, 1
    PCB f, "pgBio", "Dispersed remains:",   "Dispersed_Remains",  1, 2
    PCB f, "pgBio", "Flexed position:",     "Flexed_Position",    2, 2
    PCB f, "pgBio", "Bone burning:",        "Bone_Burning",       3, 2
End Sub

' TAB 7 - CULTURAL MATERIALS (unchanged)
' ----------------------------------------------------------------
Private Sub FillMat(f As String)
    SH f, "pgMat", "Cultural materials", 0
    PCB f, "pgMat", "Textiles:",       "Mat_Textiles",   1, 1
    PCB f, "pgMat", "Cultural wood:",  "Mat_Wood",       2, 1
    PCB f, "pgMat", "Vegetal fiber:",  "Mat_VegFiber",   3, 1
    PCB f, "pgMat", "Ceramics:",       "Mat_Ceramics",   4, 1
    PCB f, "pgMat", "Fauna:",          "Mat_Fauna",      1, 2
    PCB f, "pgMat", "Deer antler:",    "Mat_DeerAntler", 2, 2
    PCB f, "pgMat", "Other materials:","Mat_Other",      3, 2
End Sub

' TAB 8 - CHRONOLOGY & VOLUMETRY (unchanged)
' ----------------------------------------------------------------
Private Sub FillCron(f As String)
    SH  f, "pgCron", "Chronology", 0
    PCB f, "pgCron", "C14 dating:",         "C14",              1, 1
    PCT f, "pgCron", "Start (century CE):", "Chrono_Start_Cent",2, 1
    PCT f, "pgCron", "End (century CE):",   "Chrono_End_Cent",  3, 1
    PCC f, "pgCron", "Campaign:",           "ID_Campaign",      4, 1
    SH  f, "pgCron", "Volumetry & area", 5
    PCT f, "pgCron", "Interior area (m2):", "Interior_Area_m2", 6, 1
    PCT f, "pgCron", "Interior vol. (m3):", "Interior_Vol_m3",  7, 1
    PCT f, "pgCron", "Total vol. (m3):",    "Total_Vol_m3",     8, 1
    PCC f, "pgCron", "Calc. method:",       "ID_Vol_Method",    9, 1
    PCT f, "pgCron", "Volume notes:",       "Vol_Notes",        6, 2
End Sub

' TAB 9 - GIS & DIGITAL DOCUMENTATION (unchanged)
' ----------------------------------------------------------------
Private Sub FillSIG(f As String)
    SH  f, "pgSIG", "Spatial coordinates (GIS)", 0
    PCT f, "pgSIG", "Lat WGS84:",        "Coord_Lat_WGS84",   1, 1
    PCT f, "pgSIG", "Lon WGS84:",        "Coord_Lon_WGS84",   2, 1
    PCT f, "pgSIG", "E UTM (m):",        "Coord_E_UTM",       3, 1
    PCT f, "pgSIG", "N UTM (m):",        "Coord_N_UTM",       4, 1
    PCT f, "pgSIG", "Altitude (masl):",  "Altitude_masl",     5, 1
    PCT f, "pgSIG", "Precision (m):",    "Coord_Precision_m", 6, 1
    PCC f, "pgSIG", "Coord. method:",    "ID_Coord_Method",   7, 1
    SH  f, "pgSIG", "Digital documentation", 0
    PCT f, "pgSIG", "URL Panorama 360:", "URL_Pano",           1, 2
    PCT f, "pgSIG", "URL Panorama 2:",   "URL_Pano_2",         2, 2
    PCT f, "pgSIG", "URL Gigaphoto:",    "URL_Giga",           3, 2
    PCT f, "pgSIG", "URL 3D Model:",     "URL_3D",             4, 2
    PCB f, "pgSIG", "ChaXR Documented:","ChaXR_Documented",   5, 2
    PCM f, "pgSIG", "Notes:",           "Notes",              9, 1
End Sub

' TAB 10 - ARCH. SYSTEMS (unchanged)
' ----------------------------------------------------------------
Private Sub FillSys(f As String)
    SH  f, "pgSys", "Level 0 - Base system (I)", 0
    PCB f, "pgSys", "Base level (I):",          "Base_Level",          1, 1
    PCB f, "pgSys", "Decorative socle (J):",    "Decorative_Socle",    1, 2
    PCB f, "pgSys", "Corbelled platform (M):",  "Corbelled_Platform",  2, 1
    SH  f, "pgSys", "  Platform system: K+L+N -> M", 3
    PCB f, "pgSys", "Timber corbels (K):",      "Timber_Brackets",     4, 1
    PCT f, "pgSys", "Corbel count:",             "Timber_Bracket_Count",4, 2
    PCB f, "pgSys", "Transverse beams (L):",    "Transverse_Beams",    5, 1
    PCB f, "pgSys", "Corbelled courses (N):",   "Corbelled_Courses",   5, 2
    PCV f, "pgSys", "Corbel material:",         "Corbel_Material",     6, 1, "Timber;Stone;Mixed;ND"
    SH  f, "pgSys", "Access opening system: A+B+P -> O", 7
    PCB f, "pgSys", "Access opening (O):",      "Access_Opening",      8, 1
    PCB f, "pgSys", "Sill/threshold (P):",      "Sill",                8, 2
    PCB f, "pgSys", "Recessed portal:",         "Recessed_Portal",     9, 1
    PCB f, "pgSys", "Jambs (O):",               "Jambs",               9, 2
    SH  f, "pgSys", "Facade & upper zone", 10
    PCB f, "pgSys", "Structural pilasters (C):","Structural_Pilasters",11, 1
    PCV f, "pgSys", "Interlevel cornice mat.:","Cornice_Material",     11, 2, "Stone slabs;Wooden beams;Mixed;ND"
    PCB f, "pgSys", "Eave / overhang (G):",     "Eave",                12, 1
    PCB f, "pgSys", "Eave-supporting beam (F):", "Eave_Beam",          12, 2
    PCB f, "pgSys", "Upper crown/coping (H):",  "Upper_Crown",         13, 1
    PCB f, "pgSys", "Eave surface (T):",         "Eave_Surface",        13, 2
End Sub

' TAB 11 - EXTRA ELEMENTS (unchanged)
' ----------------------------------------------------------------
Private Sub FillExtra(f As String)
    SH  f, "pgExtra", "Additional confirmed elements", 0
    PCB f, "pgExtra", "Embedded base beams:", "Embedded_Base_Beams", 1, 1
    PCB f, "pgExtra", "Tie walls (perp.):",   "Tie_Walls",           1, 2
    PCB f, "pgExtra", "Corner quoins:",        "Corner_Quoins",       2, 1
    PCB f, "pgExtra", "Lateral walls (V):",    "Lateral_Walls",       2, 2
    PCB f, "pgExtra", "Rear wall (W):",         "Rear_Wall",           3, 1
    PCB f, "pgExtra", "Rear wall - built:",     "Rear_Wall_Built",     3, 2
    PCB f, "pgExtra", "Chamber roof (X):",      "Chamber_Roof",        4, 1
    SH  f, "pgExtra", "Custom features (T_ARCH_FEATURES) - add any unlisted element", 5
End Sub

' ================================================================
'  CONFIGURE ALL COMBOS (Table/Query type; value lists set inline)
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
    Debug.Print "[OK] All combos configured"
End Sub
