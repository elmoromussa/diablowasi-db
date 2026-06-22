Attribute VB_Name = "generate_mainform"
Option Compare Database
Option Explicit

' ================================================================
'  CREAR FORMULARI PRINCIPAL F_ESTRUCTURES
'  Executa Sub CreateMainForm() -> F5
' ================================================================

' Constantes de disseny (twips: 1440 = 1 polzada)
Const FW  As Long = 13200  ' amplada del formulari
Const LW  As Long = 1900  ' amplada etiqueta
Const CW  As Long = 2500  ' amplada control
Const CH  As Long = 315   ' alcada control
Const lh  As Long = 270   ' alcada etiqueta
Const MT  As Long = 550   ' marge superior pagina
Const RG  As Long = 420   ' gap entre files
Const c1  As Long = 120   ' columna 1 esquerra
Const c2  As Long = 4880  ' columna 2 esquerra

' ?? ENTRADA PRINCIPAL ???????????????????????????????????????????
Sub CreateMainForm()
    Const frm = "F_STRUCTURES"

    On Error Resume Next
    DoCmd.DeleteObject acForm, frm
    On Error GoTo 0

    ' Crea formulari
    Dim f As Form
    Set f = CreateForm()
    Dim tmpName As String
    tmpName = f.Name  ' nom temporal (ex: "Form1")
    f.RecordSource = "T_STRUCTURES"
    f.DefaultView = 0
    f.ScrollBars = 3
    f.NavigationButtons = True
    f.Caption = "Funerary Structure Record"
    f.Width = FW

    ' --- DETALL (sense capcalera, tot en detail) ---
    f.Section(acDetail).Height = 7000
    f.Section(acDetail).BackColor = RGB(249, 249, 248)

    ' Titol en la seccio detail
    Dim h As Control
    Set h = CreateControl(tmpName, acLabel, acDetail, "", "", 120, 80, 7000, 480)
    h.Caption = "STRUCTURE RECORD  -  La Petaca & Diablo Wasi (PALP)"
    h.FontSize = 13
    h.FontBold = True
    h.ForeColor = RGB(26, 60, 107)
    h.BackStyle = 0
    h.BorderStyle = 0

    ' Pestanya principal (comenca mes avall per deixar lloc al titol)
    Dim tc As Control
    Set tc = CreateControl(tmpName, acTabCtl, acDetail, "", "", 60, 620, 13080, 6260)
    tc.Name = "tabMain"

    ' Pagines (les dues primeres ja existeixen)
    tc.Pages(0).Name = "pgId":   tc.Pages(0).Caption = "1.Ident."
    tc.Pages(1).Name = "pgArq":  tc.Pages(1).Caption = "2.Arq."

    Dim pg As Control
    Set pg = CreateControl(tmpName, acPage, acDetail, "tabMain")
    pg.Name = "pgAcab": pg.Caption = "3.Acab."
    Set pg = CreateControl(tmpName, acPage, acDetail, "tabMain")
    pg.Name = "pgDec":  pg.Caption = "4.Dec."
    Set pg = CreateControl(tmpName, acPage, acDetail, "tabMain")
    pg.Name = "pgEst":  pg.Caption = "5.Estat"
    Set pg = CreateControl(tmpName, acPage, acDetail, "tabMain")
    pg.Name = "pgBio":  pg.Caption = "6.Bio."
    Set pg = CreateControl(tmpName, acPage, acDetail, "tabMain")
    pg.Name = "pgMat":  pg.Caption = "7.Mat."
    Set pg = CreateControl(tmpName, acPage, acDetail, "tabMain")
    pg.Name = "pgCron": pg.Caption = "8.Cron."
    Set pg = CreateControl(tmpName, acPage, acDetail, "tabMain")
    pg.Name = "pgSIG":  pg.Caption = "9.SIG"

    ' Omple cada pestanya
    FillId tmpName
    FillArq tmpName
    FillAcab tmpName
    FillDec tmpName
    FillEst tmpName
    FillBio tmpName
    FillMat tmpName
    FillCron tmpName
    FillSIG tmpName

    ' Guarda amb nom temporal i reanomena a F_ESTRUCTURES
    DoCmd.Save acForm, tmpName
    DoCmd.Close acForm, tmpName
    DoCmd.Rename frm, acForm, tmpName

    ' Configura combos en segon pas
    SetCombos frm

    Dim msg2 As String
    msg2 = "Formulari F_ESTRUCTURES creat correctament!"
    msg2 = msg2 & vbCrLf & "Obre-el des del Panell de navegacio."
    MsgBox msg2, vbInformation, "Fet!"
End Sub

' ?? HELPERS ?????????????????????????????????????????????????????

' Etiqueta de seccio (barra blava de titol)
Sub SH(frm As String, pg As String, txt As String, row As Integer)
    Dim T As Long: T = MT + row * RG - 16
    Dim lh As Control
    Set lh = CreateControl(frm, acLabel, acDetail, pg, "", c1, T, FW - 200, 260)
    lh.Caption = "  " & UCase(txt)
    lh.BackStyle = 1
    lh.BackColor = RGB(214, 228, 247)
    lh.BorderStyle = 0
    lh.ForeColor = RGB(26, 60, 107)
    lh.FontBold = True
    lh.FontSize = 8
End Sub

' Control de text
Sub PCT(frm As String, pg As String, lbl As String, src As String, row As Integer, col As Integer)
    AddCtrl frm, pg, lbl, src, acTextBox, row, col, CW
End Sub

' ComboBox
Sub PCC(frm As String, pg As String, lbl As String, src As String, row As Integer, col As Integer)
    AddCtrl frm, pg, lbl, src, acComboBox, row, col, CW
End Sub

' CheckBox
Sub PCB(frm As String, pg As String, lbl As String, src As String, row As Integer, col As Integer)
    Dim L As Long: Dim T As Long
    If col = 1 Then L = c1 Else L = c2
    T = MT + row * RG

    Dim lc As Control
    Set lc = CreateControl(frm, acLabel, acDetail, pg, "", L, T + 22, LW, lh)
    lc.Caption = lbl
    lc.BackStyle = 0: lc.BorderStyle = 0
    lc.TextAlign = 3
    lc.ForeColor = RGB(55, 55, 80)

    Dim cc As Control
    Set cc = CreateControl(frm, acCheckBox, acDetail, pg, "", L + LW + 80, T + 20, 300, CH)
    cc.ControlSource = src
    On Error Resume Next: cc.Name = src: On Error GoTo 0
End Sub

' Camp memo (notes)
Sub PCM(frm As String, pg As String, lbl As String, src As String, row As Integer, col As Integer)
    Dim L As Long: Dim T As Long
    If col = 1 Then L = c1 Else L = c2
    T = MT + row * RG

    Dim lc As Control
    Set lc = CreateControl(frm, acLabel, acDetail, pg, "", L, T + 22, LW, lh)
    lc.Caption = lbl
    lc.BackStyle = 0: lc.BorderStyle = 0
    lc.TextAlign = 3: lc.ForeColor = RGB(55, 55, 80)

    Dim cc As Control
    Set cc = CreateControl(frm, acTextBox, acDetail, pg, "", L + LW + 80, T, CW * 2 + 100, CH * 4)
    cc.ControlSource = src
    cc.ScrollBars = 2
    On Error Resume Next: cc.Name = src: On Error GoTo 0
End Sub

Sub AddCtrl(frm As String, pg As String, lbl As String, src As String, cType As Integer, row As Integer, col As Integer, w As Long)
    Dim L As Long: Dim T As Long
    If col = 1 Then L = c1 Else L = c2
    T = MT + row * RG

    Dim lc As Control
    Set lc = CreateControl(frm, acLabel, acDetail, pg, "", L, T + 22, LW, lh)
    lc.Caption = lbl
    lc.BackStyle = 0: lc.BorderStyle = 0
    lc.TextAlign = 3: lc.ForeColor = RGB(55, 55, 80)

    Dim cc As Control
    Set cc = CreateControl(frm, cType, acDetail, pg, "", L + LW + 80, T, w, CH)
    cc.ControlSource = src
    On Error Resume Next: cc.Name = src: On Error GoTo 0
End Sub

' ?? PESTANYA 1: IDENTIFICACIO ???????????????????????????????????
Sub FillId(f As String)
    SH f, "pgId", "Identificacio i localitzacio", 0
    PCT f, "pgId", "Code:", "Code", 1, 1
    PCC f, "pgId", "Sector:", "ID_Sector", 2, 1
    PCC f, "pgId", "Typology:", "ID_Typology", 3, 1
    PCC f, "pgId", "Geom. Support:", "ID_Support", 4, 1
    PCC f, "pgId", "Parent element (inside):", "ID_Parent", 1, 2
    PCC f, "pgId", "Functional group:", "ID_Group", 2, 2
End Sub

' ?? PESTANYA 2: ARQUITECTURA ????????????????????????????????????
Sub FillArq(f As String)
    SH f, "pgArq", "Dimensions i morfologia", 0
    PCT f, "pgArq", "No. floors:", "N_Floors", 1, 1
    PCT f, "pgArq", "Floor plan:", "Floor_Plan", 2, 1
    PCT f, "pgArq", "No. built walls:", "N_Built_Walls", 3, 1
    PCT f, "pgArq", "Length (m):", "Length_m", 4, 1
    PCT f, "pgArq", "Width (m):", "Width_m", 5, 1
    PCT f, "pgArq", "Height (m):", "Height_m", 6, 1
    PCT f, "pgArq", "Approx. cliff height (m):", "Approx_Height_m", 7, 1
    PCT f, "pgArq", "Access orientation:", "Access_Orientation", 1, 2
    PCT f, "pgArq", "Lintel type:", "Lintel", 2, 2
    PCB f, "pgArq", "Natural roof:", "Natural_Roof", 3, 2
    PCB f, "pgArq", "Buttresses:", "Buttresses", 4, 2
    PCB f, "pgArq", "Interlevel cornice:", "Interlevel_Cornice", 5, 2
    PCB f, "pgArq", "Wooden stakes:", "Wooden_Stakes", 6, 2
End Sub

' ?? PESTANYA 3: ACABATS ?????????????????????????????????????????
Sub FillAcab(f As String)
    SH f, "pgAcab", "Acabats superficials", 0
    PCB f, "pgAcab", "Plastered:", "Plastered", 1, 1
    PCT f, "pgAcab", "Plaster color:", "Plaster_Color", 2, 1
    PCB f, "pgAcab", "Rock painting:", "Rock_Painting", 3, 1
    PCT f, "pgAcab", "Rock paint color:", "Rock_Paint_Color", 4, 1
End Sub

' ?? PESTANYA 4: DECORACIO ???????????????????????????????????????
Sub FillDec(f As String)
    SH f, "pgDec", "Baix-relleu", 0
    PCB f, "pgDec", "Ninxol quadrat:", "Dec_Square_Niche", 1, 1
    PCB f, "pgDec", "Relief T:", "Dec_Relief_T", 2, 1
    PCB f, "pgDec", "Relief T inv.:", "Dec_Relief_T_Inv", 3, 1
    PCB f, "pgDec", "Relief L:", "Dec_Relief_L", 4, 1
    PCB f, "pgDec", "Relief L inv.:", "Dec_Relief_L_Inv", 5, 1
    PCB f, "pgDec", "Zigzag:", "Dec_Zigzag", 6, 1
    PCB f, "pgDec", "Stepped motif:", "Dec_Stepped", 7, 1
    PCB f, "pgDec", "Frieze:", "Dec_Frieze", 8, 1
    SH f, "pgDec", "Pintura rupestre associada", 9
    PCB f, "pgDec", "Rock art:", "Rock_Art", 10, 1
    PCB f, "pgDec", "Anthropomorphic:", "RA_Anthropomorphic", 11, 1
    PCB f, "pgDec", "Zoomorphic:", "RA_Zoomorphic", 1, 2
    PCB f, "pgDec", "Geometric:", "RA_Geometric", 2, 2
    PCB f, "pgDec", "Abstract:", "RA_Abstract", 3, 2
    PCB f, "pgDec", "Decapitation scene:", "RA_Decap_Scene", 4, 2
End Sub

' ?? PESTANYA 5: ESTAT DE CONSERVACIO ???????????????????????????
Sub FillEst(f As String)
    SH f, "pgEst", "Estat de conservacio i alteracions", 0
    PCC f, "pgEst", "General status:", "ID_Status", 1, 1
    PCB f, "pgEst", "Looting:", "Looting", 2, 1
    PCB f, "pgEst", "Fire damage:", "Fire_Damage", 3, 1
    PCB f, "pgEst", "Animal activity:", "Animal_Activity", 4, 1
    PCB f, "pgEst", "Modern access:", "Modern_Access", 5, 1
End Sub

' ?? PESTANYA 6: BIOARQUEOLOGIA ??????????????????????????????????
Sub FillBio(f As String)
    SH f, "pgBio", "Context bioarqueologic", 0
    PCB f, "pgBio", "Human remains:", "Human_Remains", 1, 1
    PCT f, "pgBio", "MNI:", "MNI", 2, 1
    PCB f, "pgBio", "Anatomical connection:", "Anatomical_Connection", 3, 1
    PCB f, "pgBio", "Mummification:", "Mummification", 4, 1
    PCB f, "pgBio", "Funerary bundles:", "Funerary_Bundles", 5, 1
    PCB f, "pgBio", "Dispersed remains:", "Dispersed_Remains", 1, 2
    PCB f, "pgBio", "Flexed position:", "Flexed_Position", 2, 2
    PCB f, "pgBio", "Cremacio d'ossos:", "Bone_Burning", 3, 2
End Sub

' ?? PESTANYA 7: MATERIALS ???????????????????????????????????????
Sub FillMat(f As String)
    SH f, "pgMat", "Materials culturals associats", 0
    PCB f, "pgMat", "Textiles:", "Mat_Textiles", 1, 1
    PCB f, "pgMat", "Cultural wood:", "Mat_Wood", 2, 1
    PCB f, "pgMat", "Vegetal fiber:", "Mat_VegFiber", 3, 1
    PCB f, "pgMat", "Ceramics:", "Mat_Ceramics", 4, 1
    PCB f, "pgMat", "Fauna:", "Mat_Fauna", 1, 2
    PCB f, "pgMat", "Deer antler:", "Mat_DeerAntler", 2, 2
    PCB f, "pgMat", "Other materials:", "Mat_Other", 3, 2
End Sub

' ?? PESTANYA 8: CRONOLOGIA I VOLUMETRIA ????????????????????????
Sub FillCron(f As String)
    SH f, "pgCron", "Cronologia", 0
    PCB f, "pgCron", "C14 dating:", "C14", 1, 1
    PCT f, "pgCron", "Start (century CE):", "Chrono_Start_Cent", 2, 1
    PCT f, "pgCron", "End (century CE):", "Chrono_End_Cent", 3, 1
    PCC f, "pgCron", "Campaign:", "ID_Campaign", 4, 1
    SH f, "pgCron", "Volumetria", 5
    PCT f, "pgCron", "Interior volume (m3):", "Interior_Vol_m3", 6, 1
    PCT f, "pgCron", "Total volume (m3):", "Total_Vol_m3", 7, 1
    PCC f, "pgCron", "Calc. method:", "ID_Vol_Method", 8, 1
    PCT f, "pgCron", "Volume notes:", "Vol_Notes", 6, 2
End Sub

' ?? PESTANYA 9: SIG I DOCUMENTACIO ?????????????????????????????
Sub FillSIG(f As String)
    SH f, "pgSIG", "Coordenades espacials", 0
    PCT f, "pgSIG", "Lat WGS84:", "Coord_Lat_WGS84", 1, 1
    PCT f, "pgSIG", "Lon WGS84:", "Coord_Lon_WGS84", 2, 1
    PCT f, "pgSIG", "E UTM (m):", "Coord_E_UTM", 3, 1
    PCT f, "pgSIG", "N UTM (m):", "Coord_N_UTM", 4, 1
    PCT f, "pgSIG", "Altitude (masl):", "Altitude_masl", 5, 1
    PCT f, "pgSIG", "Precision (m):", "Coord_Precision_m", 6, 1
    PCC f, "pgSIG", "Coord. method:", "ID_Coord_Method", 7, 1
    SH f, "pgSIG", "Documentacio digital", 0
    PCT f, "pgSIG", "URL Panorama 360:", "URL_Pano", 1, 2
    PCT f, "pgSIG", "URL Panorama 2:", "URL_Pano_2", 2, 2
    PCT f, "pgSIG", "URL Gigaphoto:", "URL_Giga", 3, 2
    PCT f, "pgSIG", "URL 3D Model:", "URL_3D", 4, 2
    PCB f, "pgSIG", "ChaXR Documented:", "ChaXR_Documented", 5, 2
    PCM f, "pgSIG", "Notes:", "Notes", 9, 1
End Sub

' ?? CONFIGURAR COMBOS ???????????????????????????????????????????
Sub SetCombos(frmName As String)
    DoCmd.OpenForm frmName, acDesign
    Dim f As Form
    Set f = Forms(frmName)

    Dim c(8, 2) As String
    c(0, 0) = "ID_Sector":        c(0, 1) = "SELECT ID, Sector_Name FROM L_SECTORS ORDER BY ID_Sitio, Nom_Sector": c(0, 2) = "2"
    c(1, 0) = "ID_Typology":     c(1, 1) = "SELECT ID, Name FROM L_TYPOLOGY ORDER BY Nom":                        c(1, 2) = "2"
    c(2, 0) = "ID_Support":        c(2, 1) = "SELECT ID, Name FROM L_SUPPORT ORDER BY ID":                            c(2, 2) = "2"
    c(3, 0) = "ID_Status":         c(3, 1) = "SELECT ID, Name FROM L_STATUS ORDER BY ID":                             c(3, 2) = "2"
    c(4, 0) = "ID_Vol_Method":  c(4, 1) = "SELECT ID, Name FROM L_VOL_METHOD ORDER BY ID":                      c(4, 2) = "2"
    c(5, 0) = "ID_Coord_Method":  c(5, 1) = "SELECT ID, Name FROM L_COORD_METHOD ORDER BY ID":                      c(5, 2) = "2"
    c(6, 0) = "ID_Campaign":      c(6, 1) = "SELECT ID, Code, Campaign_Name FROM L_CAMPAIGN ORDER BY Codi":             c(6, 2) = "3"
    c(7, 0) = "ID_Group":       c(7, 1) = "SELECT ID, Group_Code FROM T_GROUPS ORDER BY Codi_Conjunt":       c(7, 2) = "2"
    c(8, 0) = "ID_Parent": c(8, 1) = "SELECT ID, Code FROM T_STRUCTURES ORDER BY Code":               c(8, 2) = "2"

    Dim i As Integer
    For i = 0 To 8
        On Error Resume Next
        Dim ctrl As Control
        Set ctrl = f.Controls(c(i, 0))
        If Not ctrl Is Nothing Then
            ctrl.RowSourceType = "Table/Query"
            ctrl.RowSource = c(i, 1)
            ctrl.BoundColumn = 1
            ctrl.ColumnCount = CInt(c(i, 2))
            ctrl.ColumnWidths = "0cm;5cm"
            ctrl.LimitToList = True
        End If
        Set ctrl = Nothing
        On Error GoTo 0
    Next i

    DoCmd.Save acForm, frmName
    DoCmd.Close acForm, frmName
End Sub
