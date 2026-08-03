Option Compare Database
Option Explicit

' ================================================================
'  CHACHAPOYA FORM BUILD SCRIPT v10 (VALENCIAN) - F_STRUCTURES
'  Versio consolidada del formulari d'entrada de dades.
'
'  PRINCIPI: etiquetes (UI) en valencia | valors emmagatzemats en angles
'
'  Canvis respecte a chachapoya_Form_v3_val.bas:
'   - Domini 0/1/9 estes a TOTS els camps observacionals (63 combos PC9):
'     morfologia, decoracio, art rupestre, alteracions, bioarqueologia i
'     materials culturals, a mes del vocabulari A-X ja convertit en v4.
'     Nomes C14 i ChaXR_Documented queden com a casella (metadades del
'     corpus: el seu FALSE no es mai ambigu).
'   - Pestanya 3.Acab reconstruida: operacio separada del substrat
'     (revoc / pigment + Pigment_Substrate). Set camps en lloc de quatre.
'   - Pestanya 5.Estat: nou bloc d'observabilitat (Doc_Basis,
'     Facade_Observability, Interior_Observability) que fa interpretables
'     els valors 9 i permet defensar la mostra (OE1, QRY_15).
'   - Pestanya 1.Id: nou combo "Suport secundari" (ID_Support_Secondary).
'     La llista L_SUPPORT es mante intacta; els suports compostos es
'     registren com a parella ordenada (dominant + secundari) en lloc del
'     valor opac "Combined".
'   - Renoms: N_Bodies, Rear_Wall_Type (llista), Platform_Surface_Material,
'     L_STRUCT_BODY.Body_No.
'   - v10: "Contraforts" i "Estaques fusta" ELIMINATS del formulari: eren
'     un segon nom per a la pilastra (K) i per a la mensula (E). Les bandes
'     verticals roges que emmarquen la facana estan integrades al pla del
'     parament i pugen tota l'alcada: aixo es una pilastra, no un contrafort.
'     Els elements de fusta que sobreixen son horitzontals i encastats: son
'     mensules. Mantindre els dos parells nomes podia generar registres
'     incoherents (el mateix element codificat en un camp, en l'altre o en
'     tots dos).
'   - v10: nou camp "Rol mensules" (Timber_Bracket_Role): suport de
'     plataforma / aillada / ambdos. Es el que buscava l'antic camp
'     d'estaques: no un element distint, sino un ROL distint. Les mensules
'     que no sostenen plataforma son la contrapartida interna de la
'     tipologia MEN (mensula aillada = xarxa de circulacio perduda).
'   - v10: "Escena decap." eliminat de 4.Dec; ara es una fila de
'     T_DECORATIONS amb el tipus "Decapitation scene".
'   - v10: T_DECORATIONS admet tractament cromatic per element gracies al
'     tipus "Plain colour field" i a l'ampliacio de L_STRUCT_BODY (11
'     posicions amb els noms del vocabulari A-X): aixi es registra que
'     "les pilastres son roges" i "el llenc de facana blanc".
'   - v9: N_Bodies desdoblat en N_Basal_Bodies + N_Chamber_Bodies. Els
'     nivells N0/N1/Sup son CATEGORIES FUNCIONALS, no un sistema de
'     numeracio: no s'amplien amb N2, N3... La repeticio va als comptadors.
'     CRITERI OPERATIU: un cos es N1 si te (o tenia) obertura d'acces;
'     si no en te, es N0, per fina que siga la seua fabrica.
'   - v9: Lost_Body_Evidence - evidencia de cos perdut. Les bandes de
'     pigment sobre la roca per damunt del cos conservat son el fantasma
'     d'un cos desaparegut. Fixa quan els elements A-X valen 9 i no 0.
'   - v9: L_STRUCT_BODY nomes descriu POSICIO dins d'un cos; quin cos ho
'     diu T_DECORATIONS.Body_No. Escala a qualsevol nombre de cossos.
'   - v8: Support_Morphology eliminat (redundant amb la parella
'     ID_Support + ID_Support_Secondary, que ja descriu el suport).
'   - v8: Access_Orientation eliminat i fos en Facade_Orientation. En una
'     estructura de penya-segat son la mateixa variable: una sola entrada.
'   - v8: Natural_Roof eliminat i absorbit per Chamber_Roof_Type. L'element X
'     conserva la presencia (0/1/9) per a la matriu A-X; el tipus registra si
'     la cambra es tanca amb la roca natural o amb obra. Mateix patro que
'     W + Rear_Wall_Type.
'   - Subformulari F_DECORATIONS: nova columna "Substrat" per als casos
'     mixtos (fris sobre revoc + brancals sobre pedra en la mateixa
'     estructura).
'
'  IMPORTANT: executar DESPRES de chachapoya_DB_v10.bas -> BuildDB()
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
    msg = "F_STRUCTURES v10 (val.) creada amb 12 pestanyes!" & vbCrLf & vbCrLf
    msg = msg & "  59 camps amb domini 0/1/9 (Absent/Present/ND)" & vbCrLf
    msg = msg & "  Nomes C14 i ChaXR_Documented son caselles" & vbCrLf
    msg = msg & "  Acabats: revoc i pigment separats + substrat" & vbCrLf
    msg = msg & "  Bloc d'observabilitat a 5.Estat" & vbCrLf
    msg = msg & "  Suport secundari a 1.Id" & vbCrLf
    msg = msg & "  Subformularis F_DECORATIONS i F_ARCH_FEATURES" & vbCrLf & vbCrLf
    msg = msg & "Recorda: el valor per defecte es 9 (ND)." & vbCrLf
    msg = msg & "L'absencia (0) s'ha de marcar activament." & vbCrLf & vbCrLf
    msg = msg & "Un cos es N1 nomes si te (o tenia) obertura." & vbCrLf
    msg = msg & "Executa QRY_16_Validation_Check periodicament."
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
    c4.RowSource = "Red;White;Both;Ochre;None;ND": c4.LimitToList = False
    On Error Resume Next: c4.Name = "Color": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "Color", "", L, T + 15, 600, 260)
    lb.Caption = "Color"

    ' v7: substrat per posicio - resol els casos mixtos (fris sobre revoc
    ' i brancals sobre pedra en la mateixa estructura)
    L = 10000
    Dim c5 As Control: Set c5 = CreateControl(tmp, acComboBox, acDetail, "", "", L + 900, T, 1600, 315)
    c5.ControlSource = "Substrate": c5.RowSourceType = "Value List"
    c5.RowSource = "Plaster;Masonry stone;Bedrock;ND": c5.LimitToList = False
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
    f.Caption = "Registre Estructura v10 - La Petaca i Diablo Wasi (PALP)"
    f.Width = FW

    f.Section(acDetail).Height = 9400
    f.Section(acDetail).BackColor = RGB(249, 249, 248)

    Dim h As Control
    Set h = CreateControl(tmp, acLabel, acDetail, "", "", 120, 80, 7000, 480)
    h.Caption = "REGISTRE D'ESTRUCTURA v10  -  La Petaca i Diablo Wasi (PALP)"
    h.FontSize = 13: h.FontBold = True
    h.ForeColor = RGB(26, 60, 107): h.BackStyle = 0: h.BorderStyle = 0

    ' Control de pestanyes
    Dim tc As Control
    Set tc = CreateControl(tmp, acTabCtl, acDetail, "", "", 60, 620, 13080, 8700)
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
    Set sf1 = CreateControl(tmp, acSubform, acDetail, "pgDec", "", C1, MT + 10 * RG + 340, 12000, 1700)
    sf1.SourceObject = "F_DECORATIONS"
    sf1.LinkMasterFields = "ID": sf1.LinkChildFields = "ID_Structure"

    Dim lhDec As Control
    Set lhDec = CreateControl(tmp, acLabel, acDetail, "pgDec", "", C1, MT + 10 * RG, 12000, 260)
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
    Debug.Print "[OK] F_STRUCTURES v10 (valencia) creada"

    ConfigureAllCombos FRM
End Sub

' ================================================================
'  CONTINGUT DE LES PESTANYES
' ================================================================

' PESTANYA 1 - IDENTIFICACIO I DETALL DE SUPORT
' v7: suport secundari (parella ordenada en lloc del valor "Combined")
' v8: Support_Morphology eliminat - la parella ID_Support +
'     ID_Support_Secondary ja descriu el suport sense risc de contradiccio.
' -----------------------------------------------
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

' PESTANYA 2 - MORFOLOGIA, MACONERIA I FASES
' v9: comptadors de cossos separats per nivell + evidencia de cos perdut
' --------------------------------------------
Private Sub FillArq(f As String)
    SH f, "pgArq", "Morfologia general", 0
    PCT f, "pgArq", "Cossos basals (N0):",    "N_Basal_Bodies",      1, 1
    PCT f, "pgArq", "Cossos cambra (N1):",    "N_Chamber_Bodies",    2, 1
    PCV f, "pgArq", "Evidencia cos perdut:",  "Lost_Body_Evidence",  3, 1, "Pigment on bedrock;Truncated walls;Empty beam sockets;Corbels into void;Detached debris;None;ND"
    PCV f, "pgArq", "Planta:",                "Floor_Plan",          4, 1, "Rectangular;Sub-rectangular;Square;Circular;Sub-circular;Trapezoidal;Irregular;ND"
    PCT f, "pgArq", "Murs construits:",       "N_Built_Walls",       5, 1
    PCT f, "pgArq", "Cota sobre la base (m):", "Height_Above_Base_m", 6, 1
    PCV f, "pgArq", "Material dintell (Q):", "Lintel",           1, 2, "Stone;Wood;Mixed;Absent;ND"
    PCV f, "pgArq", "Tipus coberta (X):",    "Chamber_Roof_Type", 2, 2, "Natural bedrock;Built masonry;Built timber and slabs;Mixed;ND"
    SH  f, "pgArq", "Facana i paisatge (observacional)", 7
    PCV f, "pgArq", "Orientacio facana:", "Facade_Orientation", 8, 1, "N;NE;E;SE;S;SW;W;NW;ND"
    PCV f, "pgArq", "Visibilitat vall:",  "Visibility_Valley",  8, 2, "High;Medium;Low;ND"
    SH  f, "pgArq", "Maconeria i morter (T&A 2017 / H01, H04)", 9
    PCV f, "pgArq", "Qualitat maconeria:", "Masonry_Quality", 10, 1, "Good;Moderate;Poor;ND"
    PCV f, "pgArq", "Tipus aparell:",      "Masonry_Type",    10, 2, "Well-coursed;Irregular-coursed;Uncoursed;Mixed;ND"
    PC9 f, "pgArq", "Morter present:",     "Mortar_Present",  11, 1
    PCV f, "pgArq", "Tipus morter:",       "Mortar_Type",     11, 2, "Mud;Mud with gravel;Mud with organics;None dry-laid;ND"
    PC9 f, "pgArq", "Ripio / falques:",    "Chinking_Stones", 12, 1
    PCT f, "pgArq", "Notes morter:",       "Mortar_Notes",    12, 2
    SH  f, "pgArq", "Fases constructives (H03/H04)", 13
    PCT f, "pgArq", "Num. fases:",     "Construction_Phases", 14, 1
    PCV f, "pgArq", "Evidencia fase:", "Phase_Evidence",      14, 2, "C14;Stratigraphy;Superposition;Mortar;ND"
End Sub

' PESTANYA 3 - TRACTAMENTS SUPERFICIALS
' v7: operacio separada del substrat. Revocar i pintar son dues
' operacions distintes de la cadena; Pigment_Substrate registra si el
' pigment es va aplicar sobre revoc preparat o directament sobre la
' pedra del parament (H01, H04) i controla la preservacio diferencial.
' ---------------------------------------
Private Sub FillAcab(f As String)
    SH f, "pgAcab", "Revoc (lluit)", 0
    PC9 f, "pgAcab", "Revoc present:", "Plaster_Present", 1, 1
    PCV f, "pgAcab", "Color revoc:",   "Plaster_Color",   2, 1, "White;Cream;Red;Ochre;Grey;ND"
    PCV f, "pgAcab", "Extensio revoc:", "Plaster_Extent", 3, 1, "Full facade;Partial;Traces only;ND"
    SH f, "pgAcab", "Pigment aplicat", 4
    PC9 f, "pgAcab", "Pigment present:", "Pigment_Present",   5, 1
    PCV f, "pgAcab", "Substrat pigment:", "Pigment_Substrate", 6, 1, "Plaster;Masonry stone;Bedrock;Mixed;ND"
    PCV f, "pgAcab", "Color pigment:",   "Pigment_Color",     5, 2, "Red;White;Both;Ochre;ND"
    PCV f, "pgAcab", "Extensio pigment:", "Pigment_Extent",   6, 2, "Whole facade;Architectural elements;Decorative motifs;Traces;ND"
End Sub

' PESTANYA 4 - DECORACIO
' v7: booleans de decoracio i art rupestre amb domini 0/1/9. Aquests
' camps alimenten QRY_02 (khi-quadrat LP vs DW): sense el 9 la prova
' mesuraria preservacio diferencial de facanes i es llegiria com a
' practica decorativa diferencial.
' -----------------------
Private Sub FillDec(f As String)
    SH f, "pgDec", "Decoracio en baix relleu", 0
    PC9 f, "pgDec", "Ninxol quadrat:",  "Dec_Square_Niche", 1, 1
    PC9 f, "pgDec", "Relleu T:",        "Dec_Relief_T",     2, 1
    PC9 f, "pgDec", "Relleu T inv.:",   "Dec_Relief_T_Inv", 3, 1
    PC9 f, "pgDec", "Relleu L:",        "Dec_Relief_L",     4, 1
    PC9 f, "pgDec", "Relleu L inv.:",   "Dec_Relief_L_Inv", 1, 2
    PC9 f, "pgDec", "Zigzag:",          "Dec_Zigzag",       2, 2
    PC9 f, "pgDec", "Motiu escalonat:", "Dec_Stepped",      3, 2
    SH f, "pgDec", "Art rupestre associat (sobre el penyal)", 5
    PC9 f, "pgDec", "Art rupestre:",  "Rock_Art",           6, 1
    PC9 f, "pgDec", "Antropomorf:",   "RA_Anthropomorphic", 7, 1
    PC9 f, "pgDec", "Zoomorf:",       "RA_Zoomorphic",      8, 1
    PC9 f, "pgDec", "Geometric:",     "RA_Geometric",       6, 2
    PC9 f, "pgDec", "Abstract:",      "RA_Abstract",        7, 2
End Sub

' PESTANYA 5 - ESTAT DE CONSERVACIO I OBSERVABILITAT
' v7: alteracions amb domini 0/1/9 + bloc d'observabilitat, que fa
' interpretables els valors 9 de tot el registre (OE1, QRY_15).
' ----------------------------------------------------
Private Sub FillEst(f As String)
    SH f, "pgEst", "Estat de conservacio i alteracions", 0
    PCC f, "pgEst", "Estat estructura:",       "ID_Arch_Status",     1, 1
    PCC f, "pgEst", "Estat vestigis mobles:",  "ID_Material_Status", 2, 1
    PC9 f, "pgEst", "Saquejada:",              "Looting",            3, 1
    PC9 f, "pgEst", "Dany per foc:",           "Fire_Damage",        4, 1
    PC9 f, "pgEst", "Activitat animal:",       "Animal_Activity",    3, 2
    PC9 f, "pgEst", "Acces modern:",           "Modern_Access",      4, 2
    SH  f, "pgEst", "Base documental i observabilitat (OE1)", 6
    PCV f, "pgEst", "Base documental:",   "Doc_Basis",              7, 1, "Direct access;Close-range photogrammetry;Distant photogrammetry;Ground photography;Published source;ND"
    PCV f, "pgEst", "Observ. facana:",    "Facade_Observability",   8, 1, "Complete;Partial;Poor;ND"
    PCV f, "pgEst", "Observ. interior:",  "Interior_Observability", 8, 2, "Complete;Partial;None;ND"
End Sub

' PESTANYA 6 - BIOARQUEOLOGIA
' v7: domini 0/1/9. Els interiors s'observen sovint per una obertura o
' des d'un dron: es pot veure un farcell sense poder determinar flexio.
' Un FALSE aci hauria estat massivament fals.
' ----------------------------
Private Sub FillBio(f As String)
    SH f, "pgBio", "Context bioarqueologic", 0
    PC9 f, "pgBio", "Restes humanes:",      "Human_Remains",         1, 1
    PCT f, "pgBio", "MNI:",                 "MNI",                   2, 1
    PC9 f, "pgBio", "Connexio anatomica:",  "Anatomical_Connection", 3, 1
    PC9 f, "pgBio", "Mumificacio:",         "Mummification",         4, 1
    PC9 f, "pgBio", "Farcells funeraris:",  "Funerary_Bundles",      5, 1
    PC9 f, "pgBio", "Restes disperses:",    "Dispersed_Remains",     1, 2
    PC9 f, "pgBio", "Posicio flexada:",     "Flexed_Position",       2, 2
    PC9 f, "pgBio", "Os cremat:",           "Bone_Burning",          3, 2
End Sub

' PESTANYA 7 - MATERIALS CULTURALS
' v7: domini 0/1/9 (requereixen visio interior, com la bioarqueologia)
' ----------------------------------
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
    PCV f, "pgSys", "Mat. superficie H:",       "Platform_Surface_Material", 6, 1, "Timber;Stone;Mixed;ND"
    PC9 f, "pgSys", "Plataforma acces (H):",    "Corbelled_Platform",   6, 2
    PCV f, "pgSys", "Rol mensules (E):",        "Timber_Bracket_Role",  7, 1, "Platform support;Isolated;Both;ND"
    SH  f, "pgSys", "Interficie N0/N1", 8
    PC9 f, "pgSys", "Cornisa intercos (I):", "Interbody_Cornice",          9, 1
    PCV f, "pgSys", "Mat. cornisa (I):",     "Interbody_Cornice_Material", 9, 2, "Stone slabs;Wooden beams;Mixed;ND"
    SH  f, "pgSys", "Nivell 1 - alcat i sistema obertura: N+O+Q -> P", 10
    PC9 f, "pgSys", "Cantoneres (J):",          "Corner_Quoins",        11, 1
    PC9 f, "pgSys", "Pilastres estruct. (K):",  "Structural_Pilasters", 11, 2
    PC9 f, "pgSys", "Paraments laterals (L):",  "Lateral_Wall_Faces",   12, 1
    PC9 f, "pgSys", "Fris en relleu (M):",      "Relief_Frieze",        12, 2
    PC9 f, "pgSys", "Llindar (N):",             "Sill",                 13, 1
    PC9 f, "pgSys", "Brancals (O):",            "Jambs",                13, 2
    PC9 f, "pgSys", "Obertura d'acces (P):",    "Access_Opening",       14, 1
    PC9 f, "pgSys", "Portal enfonsat:",         "Recessed_Portal",      14, 2
    PC9 f, "pgSys", "Murs laterals (V):",       "Lateral_Walls",        15, 1
    PC9 f, "pgSys", "Mur posterior (W):",       "Rear_Wall",            15, 2
    PCV f, "pgSys", "Tipus mur post.:",         "Rear_Wall_Type",       16, 1, "Built masonry;Natural bedrock;Mixed;ND"
    PC9 f, "pgSys", "Coronament (R):",          "Upper_Crown",          16, 2
    SH  f, "pgSys", "Zona superior: S+T -> U", 17
    PC9 f, "pgSys", "Biga suport rafec (S):",  "Eave_Beam",    18, 1
    PC9 f, "pgSys", "Superficie rafec (T):",   "Eave_Surface", 18, 2
    PC9 f, "pgSys", "Rafec-voladis (U):",      "Eave",         19, 1
    PC9 f, "pgSys", "Coberta cambra (X):",     "Chamber_Roof", 19, 2
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

    Dim cs(12) As String
    Dim rs(12) As String
    Dim cc(12) As Integer
    Dim cw(12) As String

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
    cs(10) = "ID_Struct_Body":    rs(10) = "SELECT ID, Name FROM L_STRUCT_BODY ORDER BY Level_Type, Name":       cc(10) = 2: cw(10) = "0cm;5cm"
    cs(12) = "ID_Support_Secondary": rs(12) = "SELECT ID, Name FROM L_SUPPORT ORDER BY ID":            cc(12) = 2: cw(12) = "0cm;5cm"
    cs(11) = "ID_Dec_Type":       rs(11) = "SELECT ID, Name FROM L_DEC_TYPE ORDER BY Name":                      cc(11) = 2: cw(11) = "0cm;5cm"

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
