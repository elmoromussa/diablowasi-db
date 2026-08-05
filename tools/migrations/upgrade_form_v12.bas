Option Compare Database
Option Explicit

' ================================================================
'  CHACHAPOYA - UPGRADE FORM v12 (GATING ESTES + QUICK-FILL)
'  La Petaca & Diablo Wasi (Leymebamba, Amazonas, Peru)
'  Author: Esteve Ribera Torro | TFM Arqueologia UA
'
'  EXECUTAR SOBRE LA BD v11 (migrada o construida de zero), amb el
'  formulari F_STRUCTURES v11 ja creat per chachapoya_Form_v11_val.bas.
'  Es idempotent: es pot tornar a executar sense duplicar res.
'
'  REQUEREIX: "Confiar en l'acces al model d'objectes de projectes
'  VBA" activat al Centre de confianca d'Access (com el gating v11).
'
'  QUE FA (sis blocs):
'   U1. Camp nou T_STRUCTURES.Dec_Present (BYTE 0/1/9, defecte 0).
'       Restaura la distincio absencia verificada / no observable per
'       a la decoracio, que la v11 havia perdut en eliminar els
'       booleans: sense este camp, "cap fila a T_DECORATIONS" no
'       distingia 0 de 9. Els 35 registres existents queden NULL
'       (regla B) i es resolen a ma.
'   U2. QRY_16c_Rules_22_26 (regles 22-26) i reconstruccio de
'       QRY_16_Validation_Check com a unio a+b+c.
'   U3. Subformulari nou F_LOST_ELEMENTS amb validacio R21 en viu.
'       La regla 2 exigeix evidencia per a cada 3 (desaparegut), pero
'       la v11 no donava cap via al formulari per a registrar-la.
'   U4. Validacio R20 en viu a F_CONNECTIONS (BeforeUpdate) i
'       autoassignacio Contemporary per a junta travada.
'   U5. Afegits a F_STRUCTURES: combo Dec_Present a 4.Dec, avis
'       d'estructura colapsada a 11.Sist (R6), subformulari
'       d'elements desapareguts a 12.Extra.
'   U6. Modul de gating v12 (substitueix el v11 sencer):
'       - Nivell 1 (Record_Class -> pestanyes) i nivell 2 (Sys_* ->
'         grups; ID_Material_Status -> 7.Mat; Human_Remains -> 6.Bio)
'         identics a la v11.
'       - Nivell 2 nou: Plaster_Present -> color/extensio;
'         Pigment_Present -> substrat/color/extensio; Mortar_Present
'         -> tipus; Dec_Present -> subformulari de decoracio.
'       - Nivell 3 nou (detall depenent d'element): E -> nombre i rol
'         de mensules; rol Isolated -> bloqueja material de
'         plataforma (R7); Q -> material del dintell (R10); X ->
'         tipus de coberta (R14); I -> material de cornisa.
'       - QUICK-FILL amb confirmacio: en tancar una porta (sistema a
'         Absent / No aplicable / No observable, revoc o pigment a
'         0/9, vestigis Absent/ND, restes humanes 0/9) s'ofereix
'         emplenar els camps depenents amb el valor coherent (el 0 de
'         farciment de la seccio 1.2, el 9, o buit). Aixi les regles
'         4, 10, 19, 24 i 25 queden satisfetes en el moment d'entrada
'         i no com a troballa posterior.
'       - Recordatori en marcar un 3: cal una fila d'evidencia (R2).
'       - El combo Pigment_Substrate amaga l'opcio Plaster quan
'         Plaster_Present = 0 (prevencio de R13).
'
'  RESTRICCIONS RESPECTADES: cap continuacio de linia, nomes ASCII,
'  sql = sql & "..." en linies separades, un unic Sub public.
' ================================================================

Private mCode As String

Sub UpgradeV12()
    Dim db As DAO.Database
    Set db = CurrentDb()

    Debug.Print "================================================================"
    Debug.Print "UPGRADE FORM v12 : INICI (" & Now & ")"
    Debug.Print "================================================================"

    U1_SchemaDecPresent db
    U2_ValidationRules22to26 db
    U3_LostSubform
    U4_ConnValidation
    U5_FormAdditions
    U6_InjectGatingV12

    Debug.Print "================================================================"
    Debug.Print "UPGRADE FORM v12 : COMPLET"
    Debug.Print "================================================================"

    Dim msg As String
    msg = "Upgrade v12 aplicat!" & vbCrLf & vbCrLf
    msg = msg & "  Dec_Present afegit (els registres existents queden buits:" & vbCrLf
    msg = msg & "  resoleu-los a 0/1/9; les regles 22-23 els vigilen)." & vbCrLf
    msg = msg & "  Gating estes: revoc, pigment, morter, decoracio," & vbCrLf
    msg = msg & "  i detalls d'element (E, Q, X, I, rol de mensules)." & vbCrLf
    msg = msg & "  Quick-fill amb confirmacio en tancar seccions." & vbCrLf
    msg = msg & "  Subformulari d'elements desapareguts a 12.Extra." & vbCrLf
    msg = msg & "  QRY_16 ara porta 26 regles." & vbCrLf & vbCrLf
    msg = msg & "El gating preveu; QRY_16 continua detectant: executeu-la" & vbCrLf
    msg = msg & "periodicament igualment (taules obertes a ma, importacions" & vbCrLf
    msg = msg & "o gating desactivat poden esquivar el formulari)."
    MsgBox msg, vbInformation, "Fet!"

    Set db = Nothing
End Sub

' ================================================================
'  HELPERS COMPARTITS
' ================================================================
Private Function FldExists(db As DAO.Database, tbl As String, fld As String) As Boolean
    Dim f As DAO.Field
    On Error Resume Next
    For Each f In db.TableDefs(tbl).Fields
        If f.Name = fld Then FldExists = True: Exit Function
    Next f
    On Error GoTo 0
End Function

Private Function QryExists(db As DAO.Database, n As String) As Boolean
    Dim q As DAO.QueryDef
    For Each q In db.QueryDefs
        If q.Name = n Then QryExists = True: Exit Function
    Next q
End Function

Private Sub DropQry(db As DAO.Database, n As String)
    If QryExists(db, n) Then db.QueryDefs.Delete n
End Sub

Private Function CtlBySrc(f As Form, src As String) As Control
    Dim c As Control
    For Each c In f.Controls
        If c.ControlType = acComboBox Or c.ControlType = acTextBox Or c.ControlType = acCheckBox Then
            If c.ControlSource = src Then
                Set CtlBySrc = c
                Exit Function
            End If
        End If
    Next c
End Function

Private Function CtlByName(f As Form, n As String) As Boolean
    Dim c As Control
    For Each c In f.Controls
        If c.Name = n Then CtlByName = True: Exit Function
    Next c
End Function

Private Sub SetAU(f As Form, src As String)
    Dim c As Control
    On Error Resume Next
    For Each c In f.Controls
        If c.ControlType = acComboBox Or c.ControlType = acTextBox Or c.ControlType = acCheckBox Then
            If c.ControlSource = src Then c.AfterUpdate = "[Event Procedure]"
        End If
    Next c
    On Error GoTo 0
End Sub

' Acumulador per a generar codi de modul llegiblement.
Private Sub L(s As String)
    mCode = mCode & s & vbCrLf
End Sub

' Substitueix el modul sencer d'un formulari pel codi acumulat.
Private Sub ReplaceModule(frmName As String)
    Dim f As Form
    Set f = Forms(frmName)
    f.HasModule = True
    With f.Module
        If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
        .InsertLines 1, mCode
    End With
End Sub

' ================================================================
'  U1 - Dec_Present (BYTE 0/1/9, defecte 0 nomes per a files noves)
' ================================================================
Private Sub U1_SchemaDecPresent(db As DAO.Database)
    Debug.Print "---- [U1] START ----"
    If FldExists(db, "T_STRUCTURES", "Dec_Present") Then
        Debug.Print "[U1] Dec_Present ja existeix - saltat"
    Else
        db.Execute "ALTER TABLE T_STRUCTURES ADD COLUMN Dec_Present BYTE", dbFailOnError
        db.TableDefs.Refresh
        Debug.Print "[U1] Dec_Present afegit (els 35 registres existents queden NULL, regla B)"
    End If
    On Error Resume Next
    db.TableDefs("T_STRUCTURES").Fields("Dec_Present").DefaultValue = "0"
    On Error GoTo 0
    Dim p As DAO.Property
    On Error Resume Next
    db.TableDefs("T_STRUCTURES").Fields("Dec_Present").Properties("Caption") = "Decoracio present"
    If Err.Number <> 0 Then
        Err.Clear
        Set p = db.TableDefs("T_STRUCTURES").Fields("Dec_Present").CreateProperty("Caption", dbText, "Decoracio present")
        db.TableDefs("T_STRUCTURES").Fields("Dec_Present").Properties.Append p
    End If
    On Error GoTo 0
    Debug.Print "[U1] DefaultValue 0 i caption establits"
    Debug.Print "---- [U1] END ----"
End Sub

' ================================================================
'  U2 - Regles 22-26 i reconstruccio de la bateria
'  R22/R23: coherencia Dec_Present <-> files de T_DECORATIONS.
'  R24/R25: detall de revoc/pigment sota presencia 0 o 9.
'  R26: tipus de morter incompatible amb Mortar_Present = 0.
' ================================================================
Private Sub U2_ValidationRules22to26(db As DAO.Database)
    Debug.Print "---- [U2] START ----"

    DropQry db, "QRY_16c_Rules_22_26"

    Dim q As String
    q = "SELECT E.Code AS Structure, 22 AS Rule_No, "
    q = q & "'R22: decoration declared present but T_DECORATIONS has no row' AS Rule_Violated, "
    q = q & "'Add the decoration rows, or set Dec_Present to 0 / 9' AS Action "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Dec_Present=1 "
    q = q & "AND E.ID NOT IN (SELECT D.ID_Structure FROM T_DECORATIONS AS D) "
    q = q & "UNION ALL SELECT E.Code, 23, "
    q = q & "'R23: decoration rows exist but Dec_Present is not 1', "
    q = q & "'Set Dec_Present to 1, or review the T_DECORATIONS rows' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE (E.Dec_Present<>1 OR E.Dec_Present Is Null) "
    q = q & "AND E.ID IN (SELECT D.ID_Structure FROM T_DECORATIONS AS D) "
    q = q & "UNION ALL SELECT E.Code, 24, "
    q = q & "'R24: plaster detail recorded although Plaster_Present is 0 or 9', "
    q = q & "'Clear Plaster_Color / Plaster_Extent, or correct Plaster_Present' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Plaster_Present IN (0,9) "
    q = q & "AND (E.Plaster_Color Is Not Null OR E.Plaster_Extent Is Not Null) "
    q = q & "UNION ALL SELECT E.Code, 25, "
    q = q & "'R25: pigment detail recorded although Pigment_Present is 0 or 9', "
    q = q & "'Clear the pigment detail fields, or correct Pigment_Present' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Pigment_Present IN (0,9) "
    q = q & "AND (E.Pigment_Substrate Is Not Null OR E.Pigment_Color Is Not Null OR E.Pigment_Extent Is Not Null) "
    q = q & "UNION ALL SELECT E.Code, 26, "
    q = q & "'R26: Mortar_Present is 0 but Mortar_Type is not None dry-laid', "
    q = q & "'Set Mortar_Type to None dry-laid or clear it, or correct Mortar_Present' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Mortar_Present=0 "
    q = q & "AND E.Mortar_Type Is Not Null AND E.Mortar_Type<>'None dry-laid';"

    db.CreateQueryDef "QRY_16c_Rules_22_26", q
    Debug.Print "[U2] QRY_16c_Rules_22_26 creada"

    Dim hasB As Boolean
    hasB = QryExists(db, "QRY_16b_Rules_12_21")
    If Not QryExists(db, "QRY_16a_Rules_1_11") Then
        Debug.Print "[U2] AVIS: QRY_16a no trobada - bateria NO reconstruida (executeu abans la migracio)"
        Debug.Print "---- [U2] END ----"
        Exit Sub
    End If

    DropQry db, "QRY_16_Validation_Check"
    Dim c As String
    c = "SELECT * FROM QRY_16a_Rules_1_11"
    If hasB Then c = c & " UNION ALL SELECT * FROM QRY_16b_Rules_12_21"
    c = c & " UNION ALL SELECT * FROM QRY_16c_Rules_22_26"
    c = c & " ORDER BY Rule_No, Structure;"
    db.CreateQueryDef "QRY_16_Validation_Check", c
    db.QueryDefs.Refresh
    Debug.Print "[U2] QRY_16_Validation_Check reconstruida amb " & IIf(hasB, "26", "16") & " regles"
    Debug.Print "---- [U2] END ----"
End Sub

' ================================================================
'  U3 - F_LOST_ELEMENTS (nou subformulari, full de dades)
'  Amb validacio R21 en viu al seu propi modul.
' ================================================================
Private Sub U3_LostSubform()
    Debug.Print "---- [U3] START ----"
    Const SFRM = "F_LOST_ELEMENTS"
    On Error Resume Next
    DoCmd.DeleteObject acForm, SFRM
    On Error GoTo 0

    Dim f As Form
    Set f = CreateForm()
    Dim tmp As String
    tmp = f.Name
    f.RecordSource = "T_LOST_ELEMENTS"
    f.DefaultView = 2
    f.ScrollBars = 2
    f.NavigationButtons = False
    f.Width = 15600
    f.Section(acDetail).Height = 400

    Dim T As Long
    T = 50
    Dim Lp As Long
    Dim lb As Control

    Lp = 40
    Dim c1 As Control
    Set c1 = CreateControl(tmp, acComboBox, acDetail, "", "", Lp + 900, T, 2400, 315)
    c1.ControlSource = "Element_Code"
    c1.RowSourceType = "Table/Query"
    c1.RowSource = "SELECT Code, Name_VAL FROM L_ELEMENTS WHERE Is_System=False ORDER BY ID"
    c1.BoundColumn = 1
    c1.ColumnCount = 2
    c1.ColumnWidths = "0.8cm;4cm"
    c1.LimitToList = True
    On Error Resume Next
    c1.Name = "Element_Code"
    On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "Element_Code", "", Lp, T + 15, 840, 260)
    lb.Caption = "Element"

    Lp = 3600
    Dim c2 As Control
    Set c2 = CreateControl(tmp, acComboBox, acDetail, "", "", Lp + 1400, T, 2600, 315)
    c2.ControlSource = "ID_Evidence_Type"
    c2.RowSourceType = "Table/Query"
    c2.RowSource = "SELECT ID, Name FROM L_LOST_EVIDENCE ORDER BY ID"
    c2.BoundColumn = 1
    c2.ColumnCount = 2
    c2.ColumnWidths = "0cm;5.5cm"
    c2.LimitToList = True
    On Error Resume Next
    c2.Name = "ID_Evidence_Type"
    On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "ID_Evidence_Type", "", Lp, T + 15, 1340, 260)
    lb.Caption = "Tipus evidencia"

    Lp = 7800
    Dim c3 As Control
    Set c3 = CreateControl(tmp, acComboBox, acDetail, "", "", Lp + 700, T, 1900, 315)
    c3.ControlSource = "Evidence_Scope"
    c3.RowSourceType = "Value List"
    c3.RowSource = "Element;Element;Body;Cos;Whole structure;Estructura sencera"
    c3.BoundColumn = 1
    c3.ColumnCount = 2
    c3.ColumnWidths = "0cm;4cm"
    c3.LimitToList = True
    On Error Resume Next
    c3.Name = "Evidence_Scope"
    On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "Evidence_Scope", "", Lp, T + 15, 640, 260)
    lb.Caption = "Abast"

    Lp = 10600
    Dim c4 As Control
    Set c4 = CreateControl(tmp, acComboBox, acDetail, "", "", Lp + 800, T, 1800, 315)
    c4.ControlSource = "ID_Position"
    c4.RowSourceType = "Table/Query"
    c4.RowSource = "SELECT ID, Name FROM L_STRUCT_BODY ORDER BY Level_Type, Name"
    c4.BoundColumn = 1
    c4.ColumnCount = 2
    c4.ColumnWidths = "0cm;4.5cm"
    c4.LimitToList = True
    On Error Resume Next
    c4.Name = "ID_Position"
    On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "ID_Position", "", Lp, T + 15, 740, 260)
    lb.Caption = "Posicio"

    Lp = 13400
    Dim c5 As Control
    Set c5 = CreateControl(tmp, acTextBox, acDetail, "", "", Lp + 660, T, 1500, 315)
    c5.ControlSource = "Notes"
    On Error Resume Next
    c5.Name = "Notes"
    On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "Notes", "", Lp, T + 15, 600, 260)
    lb.Caption = "Notes"

    DoCmd.Save acForm, tmp
    DoCmd.Close acForm, tmp
    DoCmd.Rename SFRM, acForm, tmp
    Debug.Print "[U3] F_LOST_ELEMENTS creat"

    ' Modul amb la validacio R21 en viu.
    On Error GoTo Err_U3
    DoCmd.OpenForm SFRM, acDesign
    mCode = ""
    L "' Generat per upgrade_form_v12.bas. No editar a ma."
    L "Option Compare Database"
    L ""
    L "Private Sub Form_BeforeUpdate(Cancel As Integer)"
    L "    If Nz(Me!Evidence_Scope, """") = ""Element"" And IsNull(Me!Element_Code) Then"
    L "        MsgBox ""Amb abast Element cal indicar el codi de l'element (regla 21)."", vbExclamation"
    L "        Cancel = True"
    L "    End If"
    L "End Sub"
    ReplaceModule SFRM
    Forms(SFRM).BeforeUpdate = "[Event Procedure]"
    DoCmd.Save acForm, SFRM
    DoCmd.Close acForm, SFRM
    Debug.Print "[U3] Validacio R21 injectada"
    Debug.Print "---- [U3] END ----"
    Exit Sub
Err_U3:
    Debug.Print "[U3] AVIS: no s'ha pogut injectar la validacio: " & Err.Description
    Debug.Print "[U3] Activeu l'acces al model d'objectes VBA al Centre de confianca"
    On Error Resume Next
    DoCmd.Save acForm, SFRM
    DoCmd.Close acForm, SFRM
    On Error GoTo 0
End Sub

' ================================================================
'  U4 - Validacio R20 en viu a F_CONNECTIONS
'  Una direccio cronologica nomes es llegible d'una junta vertical
'  adossada o d'una superposicio; una junta travada implica
'  contemporaneitat i s'autoassigna si el camp esta buit.
' ================================================================
Private Sub U4_ConnValidation()
    Debug.Print "---- [U4] START ----"
    Const SFRM = "F_CONNECTIONS"
    On Error GoTo Err_U4
    DoCmd.OpenForm SFRM, acDesign
    Dim f As Form
    Set f = Forms(SFRM)

    mCode = ""
    L "' Generat per upgrade_form_v12.bas. No editar a ma."
    L "Option Compare Database"
    L ""
    L "Private Sub Form_BeforeUpdate(Cancel As Integer)"
    L "    Dim cr As String"
    L "    cr = Nz(Me!Chrono_Relation, """")"
    L "    If cr = ""A earlier than B"" Or cr = ""B earlier than A"" Then"
    L "        Dim ct As String"
    L "        ct = Nz(Me!Connection_Type, """")"
    L "        If ct <> ""Abutted vertical joint"" And ct <> ""Superposition"" Then"
    L "            MsgBox ""Una relacio direccional nomes es llegeix d'una junta vertical adossada o d'una superposicio (regla 20). Corregiu el tipus de connexio o marqueu Undetermined."", vbExclamation"
    L "            Cancel = True"
    L "        End If"
    L "    End If"
    L "End Sub"
    L ""
    L "Private Sub Connection_Type_AfterUpdate()"
    L "    If Nz(Me!Connection_Type, """") = ""Bonded joint"" Then"
    L "        If IsNull(Me!Chrono_Relation) Then Me!Chrono_Relation = ""Contemporary"""
    L "    End If"
    L "End Sub"
    ReplaceModule SFRM
    f.BeforeUpdate = "[Event Procedure]"
    SetAU f, "Connection_Type"
    DoCmd.Save acForm, SFRM
    DoCmd.Close acForm, SFRM
    Debug.Print "[U4] Validacio R20 i autoassignacio Contemporary injectades"
    Debug.Print "---- [U4] END ----"
    Exit Sub
Err_U4:
    Debug.Print "[U4] AVIS: no s'ha pogut injectar: " & Err.Description
    On Error Resume Next
    DoCmd.Save acForm, SFRM
    DoCmd.Close acForm, SFRM
    On Error GoTo 0
End Sub

' ================================================================
'  U5 - Afegits a F_STRUCTURES (disseny)
'  1. Redimensiona la seccio i el tab per a fer lloc a 12.Extra.
'  2. Combo Dec_Present a 4.Dec (domini 0/1/9 de dues columnes).
'  3. Etiqueta lblColl a 11.Sist (avis R6, oculta per defecte).
'  4. Capcalera + subformulari sfLost a 12.Extra.
' ================================================================
Private Sub U5_FormAdditions()
    Debug.Print "---- [U5] START ----"
    Const FRM = "F_STRUCTURES"
    Const MT As Long = 550
    Const RG As Long = 420
    Const C1 As Long = 120
    Const C2 As Long = 5600
    Const LW As Long = 2600
    Const FW As Long = 13200

    DoCmd.OpenForm FRM, acDesign
    Dim f As Form
    Set f = Forms(FRM)

    ' 1. Espai vertical per al tercer subformulari de 12.Extra.
    If f.Section(acDetail).Height < 11000 Then f.Section(acDetail).Height = 11000
    On Error Resume Next
    If f!tabMain.Height < 10300 Then f!tabMain.Height = 10300
    On Error GoTo 0
    Debug.Print "[U5] Seccio i tab redimensionats"

    ' 2. Dec_Present a 4.Dec (si no hi es ja).
    Dim T As Long
    Dim lb As Control
    Dim cc As Control
    If CtlBySrc(f, "Dec_Present") Is Nothing Then
        T = MT + 1 * RG
        Set lb = CreateControl(FRM, acLabel, acDetail, "pgDec", "", C1, T + 22, LW, 270)
        lb.Caption = "Decoracio present:"
        lb.BackStyle = 0
        lb.BorderStyle = 0
        lb.TextAlign = 3
        lb.ForeColor = RGB(55, 55, 80)
        Set cc = CreateControl(FRM, acComboBox, acDetail, "pgDec", "", C1 + LW + 80, T, 2000, 315)
        cc.ControlSource = "Dec_Present"
        cc.RowSourceType = "Value List"
        cc.RowSource = "0;Absent;1;Present;9;No observable"
        cc.ColumnCount = 2
        cc.BoundColumn = 1
        cc.ColumnWidths = "0cm;3.5cm"
        cc.LimitToList = True
        On Error Resume Next
        cc.Name = "Dec_Present"
        On Error GoTo 0
        Debug.Print "[U5] Combo Dec_Present afegit a 4.Dec"
    Else
        Debug.Print "[U5] Dec_Present ja te control - saltat"
    End If

    ' 3. Avis d'estructura colapsada (R6) a 11.Sist.
    If Not CtlByName(f, "lblColl") Then
        Set lb = CreateControl(FRM, acLabel, acDetail, "pgSys", "", C2, MT + 3 * RG, 7200, 280)
        lb.Name = "lblColl"
        lb.Caption = "AVIS: estructura colapsada - el 0 (absent) no es verificable. Useu 3 (amb evidencia) o 9."
        lb.ForeColor = RGB(180, 30, 30)
        lb.FontBold = True
        lb.BackStyle = 0
        lb.BorderStyle = 0
        lb.Visible = False
        Debug.Print "[U5] lblColl afegida a 11.Sist"
    Else
        Debug.Print "[U5] lblColl ja existeix - saltada"
    End If

    ' 4. Elements desapareguts a 12.Extra.
    If Not CtlByName(f, "sfLost") Then
        Set lb = CreateControl(FRM, acLabel, acDetail, "pgExtra", "", C1, MT + 15 * RG - 16, FW - 200, 260)
        lb.Caption = "  ELEMENTS DESAPAREGUTS - EVIDENCIA (T_LOST_ELEMENTS, REGLA 2)"
        lb.BackStyle = 1
        lb.BackColor = RGB(214, 228, 247)
        lb.BorderStyle = 0
        lb.ForeColor = RGB(26, 60, 107)
        lb.FontBold = True
        lb.FontSize = 8
        Dim sf As Control
        Set sf = CreateControl(FRM, acSubform, acDetail, "pgExtra", "", C1, MT + 16 * RG + 140, 12000, 2400)
        sf.Name = "sfLost"
        sf.SourceObject = "F_LOST_ELEMENTS"
        sf.LinkMasterFields = "ID"
        sf.LinkChildFields = "ID_Structure"
        Debug.Print "[U5] Subformulari sfLost afegit a 12.Extra"
    Else
        Debug.Print "[U5] sfLost ja existeix - saltat"
    End If

    DoCmd.Save acForm, FRM
    DoCmd.Close acForm, FRM
    Debug.Print "---- [U5] END ----"
End Sub

' ================================================================
'  U6 - Modul de gating v12 (substitucio completa del modul v11)
' ================================================================
Private Sub U6_InjectGatingV12()
    Debug.Print "---- [U6] START ----"
    Const FRM = "F_STRUCTURES"
    On Error GoTo Err_U6

    DoCmd.OpenForm FRM, acDesign
    Dim f As Form
    Set f = Forms(FRM)

    BuildGatingV12
    ReplaceModule FRM

    ' Vincula els esdeveniments que el codi generat respon.
    f.OnCurrent = "[Event Procedure]"
    SetAU f, "ID_Typology"
    SetAU f, "ID_Arch_Status"
    SetAU f, "ID_Material_Status"
    SetAU f, "Human_Remains"
    SetAU f, "Sys_Base"
    SetAU f, "Sys_Platform"
    SetAU f, "Sys_Portal"
    SetAU f, "Sys_Eave"
    SetAU f, "Sys_Chamber"
    SetAU f, "Plaster_Present"
    SetAU f, "Pigment_Present"
    SetAU f, "Mortar_Present"
    SetAU f, "Dec_Present"
    SetAU f, "Timber_Bracket_Role"

    Dim el(19) As String
    FillElems el
    Dim i As Integer
    For i = 0 To 19
        SetAU f, el(i)
    Next i

    DoCmd.Save acForm, FRM
    DoCmd.Close acForm, FRM
    Debug.Print "[U6] Gating v12 injectat (nivells 1, 2 i 3 + quick-fill)"
    Debug.Print "---- [U6] END ----"
    Exit Sub

Err_U6:
    Debug.Print "[U6] AVIS: no s'ha pogut injectar el gating: " & Err.Description
    Debug.Print "[U6] Activeu 'Confiar en l'acces al model d'objectes de projectes VBA'"
    Debug.Print "[U6] al Centre de confianca d'Access i torneu a executar UpgradeV12."
    On Error Resume Next
    DoCmd.Save acForm, FRM
    DoCmd.Close acForm, FRM
    On Error GoTo 0
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
    L "' ============================================================"
    L "' MODUL F_STRUCTURES v12 - generat per upgrade_form_v12.bas."
    L "' No editar a ma: tornar a executar UpgradeV12 el substitueix."
    L "' ============================================================"
    L "Option Compare Database"
    L ""
    L "Private Sub Form_Current()"
    L "    ApplyGating"
    L "End Sub"
    L ""
    L "Private Sub ID_Typology_AfterUpdate()"
    L "    ApplyGating"
    L "End Sub"
    L ""
    L "Private Sub ID_Arch_Status_AfterUpdate()"
    L "    ApplyGating"
    L "End Sub"
    L ""
    L "Private Sub Timber_Bracket_Role_AfterUpdate()"
    L "    ApplyGating"
    L "End Sub"
    L ""
    L "Private Sub ID_Material_Status_AfterUpdate()"
    L "    QuickFillMaterials"
    L "    ApplyGating"
    L "End Sub"
    L ""
    L "Private Sub Human_Remains_AfterUpdate()"
    L "    QuickFillBio"
    L "    ApplyGating"
    L "End Sub"
    L ""
    L "Private Sub Sys_Base_AfterUpdate()"
    L "    QuickFillSys ""Sys_Base"""
    L "    ApplyGating"
    L "End Sub"
    L ""
    L "Private Sub Sys_Platform_AfterUpdate()"
    L "    QuickFillSys ""Sys_Platform"""
    L "    ApplyGating"
    L "End Sub"
    L ""
    L "Private Sub Sys_Portal_AfterUpdate()"
    L "    QuickFillSys ""Sys_Portal"""
    L "    ApplyGating"
    L "End Sub"
    L ""
    L "Private Sub Sys_Eave_AfterUpdate()"
    L "    QuickFillSys ""Sys_Eave"""
    L "    ApplyGating"
    L "End Sub"
    L ""
    L "Private Sub Sys_Chamber_AfterUpdate()"
    L "    QuickFillSys ""Sys_Chamber"""
    L "    ApplyGating"
    L "End Sub"
    L ""
    L "Private Sub Plaster_Present_AfterUpdate()"
    L "    If Nz(Me!Plaster_Present, 1) <> 1 Then FillGroup """", 0, ""Plaster_Color,Plaster_Extent"""
    L "    ApplyGating"
    L "End Sub"
    L ""
    L "Private Sub Pigment_Present_AfterUpdate()"
    L "    If Nz(Me!Pigment_Present, 1) <> 1 Then FillGroup """", 0, ""Pigment_Substrate,Pigment_Color,Pigment_Extent"""
    L "    ApplyGating"
    L "End Sub"
    L ""
    L "Private Sub Mortar_Present_AfterUpdate()"
    L "    Dim v As Variant"
    L "    v = Me!Mortar_Present"
    L "    If Not IsNull(v) Then"
    L "        If v = 0 Then"
    L "            If Nz(Me!Mortar_Type, """") <> ""None dry-laid"" Then Me!Mortar_Type = ""None dry-laid"""
    L "        ElseIf v = 9 Then"
    L "            If Not IsNull(Me!Mortar_Type) Then Me!Mortar_Type = Null"
    L "        End If"
    L "    End If"
    L "    ApplyGating"
    L "End Sub"
    L ""
    L "Private Sub Dec_Present_AfterUpdate()"
    L "    If Nz(Me!Dec_Present, 1) <> 1 Then"
    L "        If Not IsNull(Me!ID) Then"
    L "            If DCount(""*"", ""T_DECORATIONS"", ""ID_Structure="" & Me!ID) > 0 Then"
    L "                MsgBox ""Hi ha registres de decoracio per a esta estructura: la regla 23 els marcara mentre Dec_Present no siga 1. No s'esborra res automaticament."", vbExclamation"
    L "            End If"
    L "        End If"
    L "    End If"
    L "    ApplyGating"
    L "End Sub"
    L ""

    ' Els 20 stubs d'element: recordatori del 3 + refresc del gating.
    Dim el(19) As String
    FillElems el
    Dim i As Integer
    For i = 0 To 19
        L "Private Sub " & el(i) & "_AfterUpdate()"
        L "    ElemUpd """ & el(i) & """"
        L "End Sub"
        L ""
    Next i

    L "Private Sub ElemUpd(fld As String)"
    L "    On Error Resume Next"
    L "    If Nz(Me(fld), 0) = 3 Then"
    L "        MsgBox ""Heu marcat 3 (desaparegut). Registreu l'evidencia a 12.Extra (elements desapareguts): la regla 2 exigeix una fila per a cada 3."", vbInformation"
    L "    End If"
    L "    On Error GoTo 0"
    L "    ApplyGating"
    L "End Sub"
    L ""
    L "' Activa o desactiva tot control lligat a un camp donat."
    L "Private Sub EnSrc(src As String, en As Boolean)"
    L "    Dim ct As Control"
    L "    On Error Resume Next"
    L "    For Each ct In Me.Controls"
    L "        If ct.ControlType = acComboBox Or ct.ControlType = acTextBox Or ct.ControlType = acCheckBox Then"
    L "            If ct.ControlSource = src Then ct.Enabled = en"
    L "        End If"
    L "    Next ct"
    L "    On Error GoTo 0"
    L "End Sub"
    L ""
    L "' Un sistema obri el seu grup nomes en Present* o Attested lost."
    L "Private Function SysOpen(v As Variant) As Boolean"
    L "    If IsNull(v) Then"
    L "        SysOpen = True"
    L "    Else"
    L "        SysOpen = (Left(v, 7) = ""Present"") Or (v = ""Attested lost"")"
    L "    End If"
    L "End Function"
    L ""
    L "' L'element te (o va tindre) entitat: 1, 2 o 3. NULL mante el"
    L "' detall editable, perque el judici encara no s'ha fet."
    L "Private Function ElemHas(fld As String) As Boolean"
    L "    Dim v As Variant"
    L "    v = Null"
    L "    On Error Resume Next"
    L "    v = Me(fld)"
    L "    On Error GoTo 0"
    L "    If IsNull(v) Then"
    L "        ElemHas = True"
    L "    Else"
    L "        ElemHas = (v = 1 Or v = 2 Or v = 3)"
    L "    End If"
    L "End Function"
    L ""
    L "' QUICK-FILL. En tancar un sistema, ofereix el 0 de farciment"
    L "' (Absent / No aplicable) o el 9 (No observable) als components,"
    L "' i buida els camps de detall. Mai s'escriu sense confirmacio."
    L "Private Sub QuickFillSys(sysField As String)"
    L "    Dim v As Variant"
    L "    v = Me(sysField)"
    L "    If IsNull(v) Then Exit Sub"
    L "    Dim target As Integer"
    L "    If v = ""Absent"" Or v = ""Not applicable"" Then"
    L "        target = 0"
    L "    ElseIf v = ""Not observable"" Then"
    L "        target = 9"
    L "    Else"
    L "        Exit Sub"
    L "    End If"
    L "    Dim comps As String"
    L "    Dim clears As String"
    L "    comps = """""
    L "    clears = """""
    L "    Select Case sysField"
    L "        Case ""Sys_Base"""
    L "            comps = ""Embedded_Base_Beams,Base_Level,Decorative_Socle,Tie_Walls"""
    L "        Case ""Sys_Platform"""
    L "            comps = ""Timber_Brackets,Transverse_Beams,Corbelled_Courses"""
    L "            clears = ""Timber_Bracket_Count,Timber_Bracket_Role,Platform_Surface_Material,Platform_Function"""
    L "        Case ""Sys_Portal"""
    L "            comps = ""Sill,Jambs,Lintel,Recessed_Frame"""
    L "            clears = ""Lintel_Material"""
    L "        Case ""Sys_Eave"""
    L "            comps = ""Eave_Beam,Eave_Surface"""
    L "        Case ""Sys_Chamber"""
    L "            comps = ""Corner_Quoins,Structural_Pilasters,Facade_Flank,Relief_Frieze,Return_Wall,Chamber_Roof"""
    L "            clears = ""Chamber_Roof_Type,Rear_Closure_Type"""
    L "    End Select"
    L "    FillGroup comps, target, clears"
    L "End Sub"
    L ""
    L "Private Sub QuickFillMaterials()"
    L "    Dim ms As Variant"
    L "    ms = Null"
    L "    If Not IsNull(Me!ID_Material_Status) Then"
    L "        ms = DLookup(""Name"", ""L_MATERIAL_STATUS"", ""ID="" & Me!ID_Material_Status)"
    L "    End If"
    L "    Dim g As String"
    L "    g = ""Mat_Textiles,Mat_Wood,Mat_VegFiber,Mat_Ceramics,Mat_Fauna,Mat_DeerAntler,Mat_Other"""
    L "    If Nz(ms, """") = ""Absent"" Then"
    L "        FillGroup g, 0, """""
    L "    ElseIf Nz(ms, """") = ""ND"" Then"
    L "        FillGroup g, 9, """""
    L "    End If"
    L "End Sub"
    L ""
    L "Private Sub QuickFillBio()"
    L "    Dim v As Variant"
    L "    v = Me!Human_Remains"
    L "    If IsNull(v) Then Exit Sub"
    L "    Dim g As String"
    L "    g = ""Anatomical_Connection,Mummification,Funerary_Bundles,Dispersed_Remains,Flexed_Position,Bone_Burning"""
    L "    If v = 0 Then"
    L "        FillGroup g, 0, ""MNI"""
    L "    ElseIf v = 9 Then"
    L "        FillGroup g, 9, ""MNI"""
    L "    End If"
    L "End Sub"
    L ""
    L "' Escriu target als camps de comps i Null als de clears, amb"
    L "' confirmacio previa si algun valor canviaria. L'assignacio per"
    L "' codi no dispara AfterUpdate, aixi que no hi ha recursio."
    L "Private Sub FillGroup(comps As String, target As Integer, clears As String)"
    L "    Dim a() As String"
    L "    Dim c() As String"
    L "    Dim i As Integer"
    L "    Dim n As Integer"
    L "    n = 0"
    L "    a = Split(comps, "","")"
    L "    c = Split(clears, "","")"
    L "    On Error Resume Next"
    L "    For i = 0 To UBound(a)"
    L "        If Nz(Me(a(i)), -1) <> target Then n = n + 1"
    L "    Next i"
    L "    For i = 0 To UBound(c)"
    L "        If Not IsNull(Me(c(i))) Then n = n + 1"
    L "    Next i"
    L "    On Error GoTo 0"
    L "    If n = 0 Then Exit Sub"
    L "    Dim msg As String"
    L "    msg = ""Voleu emplenar automaticament "" & n & "" camp(s) del grup amb el valor coherent ("" & target & "" o buit)?"""
    L "    If MsgBox(msg, vbYesNo + vbQuestion, ""Emplenat rapid"") <> vbYes Then Exit Sub"
    L "    On Error Resume Next"
    L "    For i = 0 To UBound(a)"
    L "        Me(a(i)) = target"
    L "    Next i"
    L "    For i = 0 To UBound(c)"
    L "        Me(c(i)) = Null"
    L "    Next i"
    L "    On Error GoTo 0"
    L "End Sub"
    L ""
    L "Private Sub ApplyGating()"
    L "    On Error Resume Next"
    L "    Dim rc As Variant"
    L "    rc = Null"
    L "    If Not IsNull(Me!ID_Typology) Then"
    L "        rc = DLookup(""Record_Class"", ""L_TYPOLOGY"", ""ID="" & Me!ID_Typology)"
    L "    End If"
    L ""
    L "    ' Nivell 1: la classe de registre mana sobre les pestanyes."
    L "    ' Un registre sense classificar ho mante tot obert."
    L "    Dim isNat As Boolean, isTrace As Boolean, isArt As Boolean"
    L "    isNat = (Nz(rc, """") = ""Natural funerary context"")"
    L "    isTrace = (Nz(rc, """") = ""Structural trace"")"
    L "    isArt = (Nz(rc, """") = ""Rock art panel"")"
    L "    Me!tabMain.Pages(""pgArq"").Enabled = Not (isNat Or isTrace Or isArt)"
    L "    Me!tabMain.Pages(""pgBio"").Enabled = Not (isTrace Or isArt)"
    L "    Me!tabMain.Pages(""pgMat"").Enabled = Not (isTrace Or isArt)"
    L "    Me!tabMain.Pages(""pgSys"").Enabled = Not isArt"
    L ""
    L "    ' Nivell 2: cada sistema mana sobre el seu grup (taula 4.5)."
    L "    Dim b As Boolean"
    L "    b = SysOpen(Me!Sys_Base)"
    L "    EnSrc ""Embedded_Base_Beams"", b"
    L "    EnSrc ""Base_Level"", b"
    L "    EnSrc ""Decorative_Socle"", b"
    L "    EnSrc ""Tie_Walls"", b"
    L ""
    L "    b = SysOpen(Me!Sys_Platform)"
    L "    EnSrc ""Timber_Brackets"", b"
    L "    EnSrc ""Transverse_Beams"", b"
    L "    EnSrc ""Corbelled_Courses"", b"
    L "    ' Nivell 3: el detall de les mensules nomes si E en te (R8, R9)"
    L "    Dim eb As Boolean"
    L "    eb = b And ElemHas(""Timber_Brackets"")"
    L "    EnSrc ""Timber_Bracket_Count"", eb"
    L "    EnSrc ""Timber_Bracket_Role"", eb"
    L "    ' Una mensula aillada no suporta cap plataforma (R7)"
    L "    EnSrc ""Platform_Surface_Material"", b And (Nz(Me!Timber_Bracket_Role, """") <> ""Isolated"")"
    L "    EnSrc ""Platform_Function"", b"
    L ""
    L "    b = SysOpen(Me!Sys_Portal)"
    L "    EnSrc ""Sill"", b"
    L "    EnSrc ""Jambs"", b"
    L "    EnSrc ""Lintel"", b"
    L "    ' Un dintell absent no te material (R10)"
    L "    EnSrc ""Lintel_Material"", b And ElemHas(""Lintel"")"
    L "    EnSrc ""Recessed_Frame"", b"
    L ""
    L "    b = SysOpen(Me!Sys_Eave)"
    L "    EnSrc ""Eave_Beam"", b"
    L "    EnSrc ""Eave_Surface"", b"
    L ""
    L "    b = SysOpen(Me!Sys_Chamber)"
    L "    EnSrc ""Corner_Quoins"", b"
    L "    EnSrc ""Structural_Pilasters"", b"
    L "    EnSrc ""Facade_Flank"", b"
    L "    EnSrc ""Relief_Frieze"", b"
    L "    EnSrc ""Return_Wall"", b"
    L "    EnSrc ""Chamber_Roof"", b"
    L "    ' El tipus de coberta nomes si X en te (R14)"
    L "    EnSrc ""Chamber_Roof_Type"", b And ElemHas(""Chamber_Roof"")"
    L "    EnSrc ""Rear_Closure_Type"", b"
    L ""
    L "    ' I i R no pertanyen a cap sistema; el material de la cornisa"
    L "    ' nomes si I en te."
    L "    EnSrc ""Interbody_Cornice_Material"", ElemHas(""Interbody_Cornice"")"
    L ""
    L "    ' Acabats: presencia mana sobre el detall (R24, R25)."
    L "    Dim pOn As Boolean"
    L "    pOn = IsNull(Me!Plaster_Present) Or Nz(Me!Plaster_Present, 1) = 1"
    L "    EnSrc ""Plaster_Color"", pOn"
    L "    EnSrc ""Plaster_Extent"", pOn"
    L "    Dim gOn As Boolean"
    L "    gOn = IsNull(Me!Pigment_Present) Or Nz(Me!Pigment_Present, 1) = 1"
    L "    EnSrc ""Pigment_Substrate"", gOn"
    L "    EnSrc ""Pigment_Color"", gOn"
    L "    EnSrc ""Pigment_Extent"", gOn"
    L ""
    L "    ' Prevencio de R13: sense revoc verificat (0), el substrat"
    L "    ' Plaster desapareix de la llista."
    L "    Dim rsrc As String"
    L "    rsrc = ""Masonry stone;Pedra de parament;Bedrock;Penya;Mixed;Mixt;ND;Tipus indeterminat"""
    L "    If IsNull(Me!Plaster_Present) Or Nz(Me!Plaster_Present, 1) <> 0 Then"
    L "        rsrc = ""Plaster;Revoc;"" & rsrc"
    L "    End If"
    L "    Dim ct2 As Control"
    L "    For Each ct2 In Me.Controls"
    L "        If ct2.ControlType = acComboBox Then"
    L "            If ct2.ControlSource = ""Pigment_Substrate"" Then"
    L "                If ct2.RowSource <> rsrc Then ct2.RowSource = rsrc"
    L "            End If"
    L "        End If"
    L "    Next ct2"
    L ""
    L "    ' Morter: el tipus nomes si hi ha morter o encara no s'ha dit."
    L "    EnSrc ""Mortar_Type"", IsNull(Me!Mortar_Present) Or Nz(Me!Mortar_Present, 1) = 1"
    L ""
    L "    ' Vestigis mobles: la porta es ID_Material_Status (4.6)."
    L "    Dim ms As Variant"
    L "    ms = Null"
    L "    If Not IsNull(Me!ID_Material_Status) Then"
    L "        ms = DLookup(""Name"", ""L_MATERIAL_STATUS"", ""ID="" & Me!ID_Material_Status)"
    L "    End If"
    L "    b = Not (Nz(ms, """") = ""Absent"" Or Nz(ms, """") = ""ND"")"
    L "    EnSrc ""Mat_Textiles"", b"
    L "    EnSrc ""Mat_Wood"", b"
    L "    EnSrc ""Mat_VegFiber"", b"
    L "    EnSrc ""Mat_Ceramics"", b"
    L "    EnSrc ""Mat_Fauna"", b"
    L "    EnSrc ""Mat_DeerAntler"", b"
    L "    EnSrc ""Mat_Other"", b"
    L ""
    L "    ' Bioarqueologia: Human_Remains mana sobre el detall (4.6)."
    L "    b = IsNull(Me!Human_Remains) Or Nz(Me!Human_Remains, 1) = 1"
    L "    EnSrc ""MNI"", b"
    L "    EnSrc ""Anatomical_Connection"", b"
    L "    EnSrc ""Mummification"", b"
    L "    EnSrc ""Funerary_Bundles"", b"
    L "    EnSrc ""Dispersed_Remains"", b"
    L "    EnSrc ""Flexed_Position"", b"
    L "    EnSrc ""Bone_Burning"", b"
    L ""
    L "    ' Decoracio: Dec_Present mana sobre el subformulari."
    L "    Dim dOn As Boolean"
    L "    dOn = IsNull(Me!Dec_Present) Or Nz(Me!Dec_Present, 1) = 1"
    L "    Dim ct3 As Control"
    L "    For Each ct3 In Me.Controls"
    L "        If ct3.ControlType = acSubform Then"
    L "            If ct3.SourceObject = ""F_DECORATIONS"" Then ct3.Enabled = dOn"
    L "        End If"
    L "    Next ct3"
    L ""
    L "    ' Avis R6: en una estructura colapsada el 0 no es verificable."
    L "    Dim st As Variant"
    L "    st = Null"
    L "    If Not IsNull(Me!ID_Arch_Status) Then"
    L "        st = DLookup(""Name"", ""L_STATUS"", ""ID="" & Me!ID_Arch_Status)"
    L "    End If"
    L "    Me!lblColl.Visible = (Nz(st, """") = ""Collapsed"")"
    L "End Sub"
End Sub
