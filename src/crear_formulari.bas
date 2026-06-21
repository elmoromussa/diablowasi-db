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
Const LH  As Long = 270   ' alcada etiqueta
Const MT  As Long = 550   ' marge superior pagina
Const RG  As Long = 420   ' gap entre files
Const C1  As Long = 120   ' columna 1 esquerra
Const C2  As Long = 4880  ' columna 2 esquerra

' ?? ENTRADA PRINCIPAL ???????????????????????????????????????????
Sub CreateMainForm()
    Const FRM = "F_ESTRUCTURES"

    On Error Resume Next
    DoCmd.DeleteObject acForm, FRM
    On Error GoTo 0

    ' Crea formulari
    Dim f As Form
    Set f = CreateForm()
    Dim tmpName As String
    tmpName = f.Name  ' nom temporal (ex: "Form1")
    f.RecordSource = "T_ESTRUCTURES"
    f.DefaultView = 0
    f.ScrollBars = 3
    f.NavigationButtons = True
    f.Caption = "Fitxa d'Estructura Funeraria"
    f.Width = FW

    ' --- DETALL (sense capcalera, tot en detail) ---
    f.Section(acDetail).Height = 7000
    f.Section(acDetail).BackColor = RGB(249, 249, 248)

    ' Titol en la seccio detail
    Dim h As Control
    Set h = CreateControl(tmpName, acLabel, acDetail, "", "", 120, 80, 7000, 480)
    h.Caption = "FITXA D ESTRUCTURA  La Petaca & Diablo Wasi (PALP)"
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
    FillId   tmpName
    FillArq  tmpName
    FillAcab tmpName
    FillDec  tmpName
    FillEst  tmpName
    FillBio  tmpName
    FillMat  tmpName
    FillCron tmpName
    FillSIG  tmpName

    ' Guarda amb nom temporal i reanomena a F_ESTRUCTURES
    DoCmd.Save acForm, tmpName
    DoCmd.Close acForm, tmpName
    DoCmd.Rename FRM, acForm, tmpName

    ' Configura combos en segon pas
    SetCombos FRM

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
    Set lh = CreateControl(frm, acLabel, acDetail, pg, "", C1, T, FW - 200, 260)
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
    If col = 1 Then L = C1 Else L = C2
    T = MT + row * RG

    Dim lc As Control
    Set lc = CreateControl(frm, acLabel, acDetail, pg, "", L, T + 22, LW, LH)
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
    If col = 1 Then L = C1 Else L = C2
    T = MT + row * RG

    Dim lc As Control
    Set lc = CreateControl(frm, acLabel, acDetail, pg, "", L, T + 22, LW, LH)
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
    If col = 1 Then L = C1 Else L = C2
    T = MT + row * RG

    Dim lc As Control
    Set lc = CreateControl(frm, acLabel, acDetail, pg, "", L, T + 22, LW, LH)
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
    PCT f, "pgId", "Codi:",              "Codi",              1, 1
    PCC f, "pgId", "Sector:",            "ID_Sector",         2, 1
    PCC f, "pgId", "Tipologia:",         "ID_Tipologia",      3, 1
    PCC f, "pgId", "Suport geomorf.:",   "ID_Suport",         4, 1
    PCC f, "pgId", "Element pare (dins de):", "ID_Estructura_Parent", 1, 2
    PCC f, "pgId", "Conjunt funcional:", "ID_Conjunt",        2, 2
End Sub

' ?? PESTANYA 2: ARQUITECTURA ????????????????????????????????????
Sub FillArq(f As String)
    SH f, "pgArq", "Dimensions i morfologia", 0
    PCT f, "pgArq", "Nre. pisos:",         "N_Pisos",           1, 1
    PCT f, "pgArq", "Planta:",             "Planta",            2, 1
    PCT f, "pgArq", "Nre. murs construits:", "N_Murs_Construits", 3, 1
    PCT f, "pgArq", "Llarg (m):",          "Largo_m",           4, 1
    PCT f, "pgArq", "Ample (m):",          "Ancho_m",           5, 1
    PCT f, "pgArq", "Alcada estructura (m):", "Alto_m",         6, 1
    PCT f, "pgArq", "Alt. sobre sol (m):", "Altura_Aprox_m",   7, 1
    PCT f, "pgArq", "Orientacio vano:",    "Orientacio_Vano",   1, 2
    PCT f, "pgArq", "Tipus dintell:",      "Dintel",            2, 2
    PCB f, "pgArq", "Techo natural:",      "Techo_Natural",     3, 2
    PCB f, "pgArq", "Contraforts:",        "Contraforts",       4, 2
    PCB f, "pgArq", "Cornisa entre pisos:","Cornisa_Entre_Pisos", 5, 2
    PCB f, "pgArq", "Estacas de fusta:",   "Estacas_Fusta",    6, 2
End Sub

' ?? PESTANYA 3: ACABATS ?????????????????????????????????????????
Sub FillAcab(f As String)
    SH f, "pgAcab", "Acabats superficials", 0
    PCB f, "pgAcab", "Arrebossat:",           "Arrebossat",         1, 1
    PCT f, "pgAcab", "Color arrebossat:",     "Color_Arrebossat",   2, 1
    PCB f, "pgAcab", "Pintura sobre roca:",   "Pintura_Sobre_Roca", 3, 1
    PCT f, "pgAcab", "Color pintura roca:",   "Color_Pintura_Roca", 4, 1
End Sub

' ?? PESTANYA 4: DECORACIO ???????????????????????????????????????
Sub FillDec(f As String)
    SH f, "pgDec", "Baix-relleu", 0
    PCB f, "pgDec", "Ninxol quadrat:",    "Dec_Nicho_Quadrat",  1, 1
    PCB f, "pgDec", "Relleu T:",          "Dec_Relieve_T",      2, 1
    PCB f, "pgDec", "Relleu T inv.:",     "Dec_Relieve_T_Inv",  3, 1
    PCB f, "pgDec", "Relleu L:",          "Dec_Relieve_L",      4, 1
    PCB f, "pgDec", "Relleu L inv.:",     "Dec_Relieve_L_Inv",  5, 1
    PCB f, "pgDec", "Zigzag:",            "Dec_Zigzag",         6, 1
    PCB f, "pgDec", "Escalonat:",         "Dec_Escalonat",      7, 1
    PCB f, "pgDec", "Fris / greca:",      "Dec_Fris_Greca",     8, 1
    SH  f, "pgDec", "Pintura rupestre associada", 9
    PCB f, "pgDec", "Pintura rupestre:",  "Pintura_Rupestre",   10, 1
    PCB f, "pgDec", "Antropomorfa:",      "PR_Antropomorfa",    11, 1
    PCB f, "pgDec", "Zoomorfa:",          "PR_Zoomorfa",        1,  2
    PCB f, "pgDec", "Geometrica:",        "PR_Geometrica",      2,  2
    PCB f, "pgDec", "Abstracta:",         "PR_Abstracta",       3,  2
    PCB f, "pgDec", "Escena decapitacio:","PR_Escena_Decap",    4,  2
End Sub

' ?? PESTANYA 5: ESTAT DE CONSERVACIO ???????????????????????????
Sub FillEst(f As String)
    SH f, "pgEst", "Estat de conservacio i alteracions", 0
    PCC f, "pgEst", "Estat general:",       "ID_Estat",               1, 1
    PCB f, "pgEst", "Saqueig:",             "Saqueig",                2, 1
    PCB f, "pgEst", "Incendi:",             "Incendi",                3, 1
    PCB f, "pgEst", "Activitat animal:",    "Activitat_Animal",       4, 1
    PCB f, "pgEst", "Acces modern:",        "Evidencia_Acces_Modern", 5, 1
End Sub

' ?? PESTANYA 6: BIOARQUEOLOGIA ??????????????????????????????????
Sub FillBio(f As String)
    SH f, "pgBio", "Context bioarqueologic", 0
    PCB f, "pgBio", "Restes humanes:",        "Restes_Humanes",      1, 1
    PCT f, "pgBio", "NMI:",                   "NMI",                 2, 1
    PCB f, "pgBio", "Connexio anatomica:",    "Connexio_Anatomica",  3, 1
    PCB f, "pgBio", "Momificacio:",           "Momificacio",         4, 1
    PCB f, "pgBio", "Fardells funeraris:",    "Fardells_Funeraris",  5, 1
    PCB f, "pgBio", "Restes disperses:",      "Restes_Disperses",   1, 2
    PCB f, "pgBio", "Posicio flexionada:",    "Posicio_Flexionada",  2, 2
    PCB f, "pgBio", "Cremacio d'ossos:",      "Cremacio_Ossos",      3, 2
End Sub

' ?? PESTANYA 7: MATERIALS ???????????????????????????????????????
Sub FillMat(f As String)
    SH f, "pgMat", "Materials culturals associats", 0
    PCB f, "pgMat", "Textils:",          "Mat_Textils",       1, 1
    PCB f, "pgMat", "Fusta cultural:",   "Mat_Fusta_Cultural",2, 1
    PCB f, "pgMat", "Fibra vegetal:",    "Mat_Fibra_Vegetal", 3, 1
    PCB f, "pgMat", "Ceramica:",         "Mat_Ceramica",      4, 1
    PCB f, "pgMat", "Fauna:",            "Mat_Fauna",         1, 2
    PCB f, "pgMat", "Banya de cervol:",  "Mat_Banya_Cervol",  2, 2
    PCB f, "pgMat", "Altres materials:", "Mat_Altres",        3, 2
End Sub

' ?? PESTANYA 8: CRONOLOGIA I VOLUMETRIA ????????????????????????
Sub FillCron(f As String)
    SH  f, "pgCron", "Cronologia", 0
    PCB f, "pgCron", "Datacio C14:",          "C14",             1, 1
    PCT f, "pgCron", "Inici (segle d.n.e.):", "Crono_Segle_Ini", 2, 1
    PCT f, "pgCron", "Fi (segle d.n.e.):",    "Crono_Segle_Fi",  3, 1
    PCC f, "pgCron", "Campanya:",             "ID_Campanya",     4, 1
    SH  f, "pgCron", "Volumetria", 5
    PCT f, "pgCron", "Volum interior (m3):",  "Volum_Interior_m3", 6, 1
    PCT f, "pgCron", "Volum total (m3):",     "Volum_Total_m3",    7, 1
    PCC f, "pgCron", "Metode calcul:",        "ID_Metode_Volum",   8, 1
    PCT f, "pgCron", "Notes volum:",          "Volum_Notes",       6, 2
End Sub

' ?? PESTANYA 9: SIG I DOCUMENTACIO ?????????????????????????????
Sub FillSIG(f As String)
    SH  f, "pgSIG", "Coordenades espacials", 0
    PCT f, "pgSIG", "Lat WGS84:",        "Coord_Lat_WGS84",  1, 1
    PCT f, "pgSIG", "Lon WGS84:",        "Coord_Lon_WGS84",  2, 1
    PCT f, "pgSIG", "E UTM (m):",        "Coord_E_UTM",      3, 1
    PCT f, "pgSIG", "N UTM (m):",        "Coord_N_UTM",      4, 1
    PCT f, "pgSIG", "Altitud (msnm):",   "Altitud_msnm",     5, 1
    PCT f, "pgSIG", "Precisio (m):",     "Coord_Precisio_m", 6, 1
    PCC f, "pgSIG", "Metode coords:",    "ID_Coord_Metode",  7, 1
    SH  f, "pgSIG", "Documentacio digital", 0
    PCT f, "pgSIG", "URL Panorama 360:",  "URL_Pano",         1, 2
    PCT f, "pgSIG", "URL Panorama 2:",    "URL_Pano_2",       2, 2
    PCT f, "pgSIG", "URL Gigafoto:",      "URL_Giga",         3, 2
    PCT f, "pgSIG", "URL Model 3D:",      "URL_3D",           4, 2
    PCB f, "pgSIG", "Documentat ChaXR:", "Documentat_ChaXR", 5, 2
    PCM f, "pgSIG", "Notes:",            "Notes",             9, 1
End Sub

' ?? CONFIGURAR COMBOS ???????????????????????????????????????????
Sub SetCombos(frmName As String)
    DoCmd.OpenForm frmName, acDesign
    Dim f As Form
    Set f = Forms(frmName)

    Dim c(8, 2) As String
    c(0, 0) = "ID_Sector":        c(0, 1) = "SELECT ID, Nom_Sector FROM L_SECTORES ORDER BY ID_Sitio, Nom_Sector": c(0, 2) = "2"
    c(1, 0) = "ID_Tipologia":     c(1, 1) = "SELECT ID, Nom FROM L_TIPOLOGIA ORDER BY Nom":                        c(1, 2) = "2"
    c(2, 0) = "ID_Suport":        c(2, 1) = "SELECT ID, Nom FROM L_SUPORT ORDER BY ID":                            c(2, 2) = "2"
    c(3, 0) = "ID_Estat":         c(3, 1) = "SELECT ID, Nom FROM L_ESTAT ORDER BY ID":                             c(3, 2) = "2"
    c(4, 0) = "ID_Metode_Volum":  c(4, 1) = "SELECT ID, Nom FROM L_METODE_VOLUM ORDER BY ID":                      c(4, 2) = "2"
    c(5, 0) = "ID_Coord_Metode":  c(5, 1) = "SELECT ID, Nom FROM L_COORD_METODE ORDER BY ID":                      c(5, 2) = "2"
    c(6, 0) = "ID_Campanya":      c(6, 1) = "SELECT ID, Codi, Nom_Camp FROM L_CAMPANYA ORDER BY Codi":             c(6, 2) = "3"
    c(7, 0) = "ID_Conjunt":       c(7, 1) = "SELECT ID, Codi_Conjunt FROM T_CONJUNTS ORDER BY Codi_Conjunt":       c(7, 2) = "2"
    c(8, 0) = "ID_Estructura_Parent": c(8, 1) = "SELECT ID, Codi FROM T_ESTRUCTURES ORDER BY Codi":               c(8, 2) = "2"

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
