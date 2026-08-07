Option Compare Database
Option Explicit

' ================================================================
'  CHACHAPOYA FORM BUILD SCRIPT v16 (VALENCIAN) - F_STRUCTURES
'  Author: Esteve Ribera Torro | TFM Arqueologia UA
'  Spec: DELTA_v15_v16.md
'
'  PRINCIPLE: labels (UI) in Valencian | stored values in English
'
'  IMPORTANT: run AFTER the schema exists, i.e. after
'  chachapoya_DB_v16.bas -> BuildDB() on a blank database.
'  There is no in-place upgrade: v16 is built from scratch and
'  the data is transferred with EXPORT_v15 / IMPORT_v16.
'
'  CANVIS DE LA v15 A LA v16
'
'  Q. 2.Arq: FORMAT I TREBALL DE LA PEDRA (delta 1). Masonry_Type
'     deia com s'apila la pedra i Masonry_Quality un judici
'     d'execucio; cap dels dos deia QUINA PEDRA ES, que es el
'     primer gest de la cadena operativa. Dos combos ORTOGONALS,
'     perque el vocabulari inicial barrejava tres eixos
'     (morfologia natural, grau de treball, dimensio) i una llista
'     unica deixava sense valor possible un mur de lloses laminars
'     ben escairades. Van DAVANT de Masonry_Type: es tria la pedra
'     abans d'apilar-la, i eixe es l'ordre de la cadena operativa.
'     Stone_Format duu parella ordenada i per tant NO porta Mixed.
'
'  R. 2.Arq: PLA D'ACCES I ORIENTACIO DEL PORTAL (delta 3). La
'     facana es fixa com el PLA EXPOSAT. En una estructura real
'     l'obertura es al mur estret perpendicular al farallo i la
'     decoracio al mur llarg paral-lel: si la facana la definira
'     l'obertura, Orientacio facana deixaria d'apuntar a la vall i
'     Visibilitat vall mesuraria un pla que no es veu. Access_Plane
'     registra la divergencia, que es una VARIABLE i no una
'     anomalia, i declara a mes el pla de referencia sense el qual
'     esquerra i dreta no signifiquen res a 4.Dec.
'
'  S. 4.Dec: POSICIO RELATIVA als DOS subformularis (delta 4.1).
'     Body_No indexa el cos constructiu, que una pintura sobre la
'     penya no te - pero no se substitueix per esquerra/dreta,
'     perque son eixos distints i les files arquitectoniques
'     continuen necessitant el numero de cos. CONVENCIO: esquerra i
'     dreta DES DE L'OBSERVADOR SITUAT DAVANT DEL PLA, mai des de
'     l'estructura; el pla de referencia el declara Access_Plane.
'     A la meitat rupestre, Num. cos queda DESACTIVAT (regla 37).
'
'  T. 4.Dec: SUBSTRAT COM A VERIFICADOR a la meitat rupestre
'     (delta 4.6). Revoc i pedra de parament son impossibles en una
'     fila ROC per definicio del criteri d'associacio: una fila ROC
'     amb substrat de parament es una fila que havia d'anar a
'     l'altra meitat. La llista es redueix a Penya i Mixt (Mixt
'     nomes per a la posicio Solapada), i les regles 38a-38b
'     detecten el que hi entre pel costat. Una incoherencia abans
'     silenciosa passa a ser un error detectat.
'
'  U. 11.Sist: MAT. CORNISA ELIMINAT (delta 5.1). No perque de fet
'     siguen sempre de pedra sino perque no podrien no ser-ho: un
'     element horitzontal de fusta volat entre dos cossos es una
'     mensula (E) o una biga transversal (F), de manera que el
'     valor 'Bigues de fusta' descrivia una PLATAFORMA MAL
'     CLASSIFICADA. Cau amb ell un nivell 3 de gating.
'
'  V. 11.Sist: TERMINOLOGIA. 'Tancament posterior' passa a 'Fons
'     de cambra (W)': en arqueologia funeraria 'tancament'
'     s'enten com el segellat de la tomba, i el risc era registrar
'     la cosa equivocada sense adonar-se'n. La coberta de cambra
'     duu ara el criteri de contacte a l'etiqueta: SENSE CONTACTE,
'     NO HI HA SOSTRE - una visera que passa a metres per damunt
'     no tanca res, i aixo es ID_Support, no X.
'
'  W. RECORDATORI DE CRITERI (delta 5.4, sense canvi de codi). Un
'     element A-X es present quan hi ha un component FISICAMENT
'     DIFERENCIAT del parament contigu. Una obertura resolta
'     deixant un buit al mur es Llindar=0, Brancals=0, Dintell=0 -
'     i Sys_Portal present igualment, perque l'obertura hi es. El
'     gating no ho pot imposar (el sistema i els components diuen
'     coses distintes), de manera que viu a les etiquetes.
'
'  CHANGES FROM v13 (v14) - rev. 5 del delta
'
'  L. CINC VALORS NOMES PER ALS 20 ELEMENTS A-X. Morter, revoc,
'     pigment, decoracio i art rupestre tornen a PC9. El criteri de
'     la v13 ("capes aplicades") era mal formulat: el valor 3
'     exigeix que l'evidencia de la perdua siga de NATURALESA
'     DISTINTA de la cosa perduda, i l'unica evidencia de pigment
'     es pigment. A mes, Plaster_Extent i Pigment_Extent ja
'     registraven la conservacio parcial amb "Traces".
'  M. Cultural_Materials_Present: camp general de 7.Mat, analeg a
'     Human_Remains. ID_Material_Status perd el valor Absent.
'  N. Doc_Basis: escala ordinal de qualitat documental.
'  O. C14 governa el subformulari de datacions; els segles queden
'     sempre editables (la font habitual no es radiocarbonica).
'  P. Totes les dimensions lineals en METRES, amb 2 decimals.
'
'  CHANGES FROM v12 (v13)
'
'  G. GATING D'E CORREGIT (delta 1). E era tractat com un component
'     mes de Sys_Platform, de manera que amb el sistema absent el
'     camp es bloquejava i l'emplenat rapid oferia posar-lo a 0 -
'     activament fals en una cambra amb mensula lateral. E es
'     l'UNIC element amb existencia independent del seu sistema, i
'     la regla 1 de la bateria ja l'exemptava: el gating v12 no
'     recollia una exempcio que la validacio ja feia. Ara E queda
'     sempre editable, i Count i Role pengen d'E (nivell 3) en lloc
'     del sistema.
'  H. 4.Dec: DUES SUBSECCIONS sobre la MATEIXA taula (delta 7.4).
'     Decoracio arquitectonica (posicions no-ROC) i pintura
'     rupestre (posicions ROC), cada una amb els seus combos ja
'     filtrats. L'usuari no veu mai una llista barrejada i no es
'     mante res dues vegades: la separacio analitica la dona
'     Level_Type, no la separacio fisica de les dades.
'  I. RockArt_Present al costat de Dec_Present (delta 7.5), cada un
'     governant el seu subformulari.
'  J. CINC VALORS a les capes aplicades (delta 8.3): morter, revoc,
'     pigment, decoracio i art rupestre passen de PC9 a PC5.
'  K. Color_Secondary al subformulari de decoracio; 'Both' retirat
'     de la llista de colors (delta 7.6).
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
Const C2  As Long = 6200
Const LW  As Long = 3000
Const CW  As Long = 2600
' v17: LW 2600->3000, CW 2200->2600, C2 5600->6200. Les
' etiquetes llargues es tallaven i els combos no mostraven els
' valors sencers. Amb els nous valors la columna 1 ocupa
' 120-5800 i la 2 6200-11880, dins dels 13080 de la pestanya.
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
    msg = "F_STRUCTURES v17 (val.) creada amb 12 pestanyes!" & vbCrLf & vbCrLf
    msg = msg & "  20 camps d'element amb domini de 5 valors" & vbCrLf
    msg = msg & "  24 camps observacionals amb 0/1/9" & vbCrLf
    msg = msg & "  Tots els combos de domini son de dues columnes:" & vbCrLf
    msg = msg & "  el valor guardat es en angles, l'etiqueta en valencia" & vbCrLf & vbCrLf
    msg = msg & "  11.Sist: els 6 sistemes manen sobre els seus grups" & vbCrLf
    msg = msg & "  4.Dec: nomes el subformulari T_DECORATIONS" & vbCrLf
    msg = msg & "  12.Extra: connexions, elements personalitzats" & vbCrLf
    msg = msg & "  i evidencia dels elements desapareguts (regla 2)" & vbCrLf & vbCrLf
    msg = msg & "  Gating v12: 3 nivells + emplenat rapid confirmat" & vbCrLf
    msg = msg & "  v16: 2.Arq porta format i treball de la pedra," & vbCrLf
    msg = msg & "  pla d'acces i orientacio del portal" & vbCrLf
    msg = msg & "  v17: 4.Dec - fora la posicio relativa; quatre" & vbCrLf
    msg = msg & "  trams del contorn + geometria del trac a la" & vbCrLf
    msg = msg & "  meitat rupestre; Num. cos ocult; LEFT JOIN, que" & vbCrLf
    msg = msg & "  recupera les files sense posicio assignada" & vbCrLf
    msg = msg & "  v17: 2.Arq - Fabrica, posicio del portal, marc" & vbCrLf
    msg = msg & "  reculat mogut a morfologia, evidencia de fase" & vbCrLf
    msg = msg & "  ampliada a 9 valors" & vbCrLf
    msg = msg & "  v17: el gating nou tanca NOMES amb No aplicable," & vbCrLf
    msg = msg & "  mai amb Absent, que es el valor per defecte" & vbCrLf
    msg = msg & "  v17: 9.Metr - obertura sota Sys_Portal i" & vbCrLf
    msg = msg & "  volumetria sota Sys_Chamber" & vbCrLf
    msg = msg & "  v17: estat dels vestigis mobles passa a 7.Mat" & vbCrLf
    msg = msg & "  v17: Sys_Interface governa la cornisa intercos" & vbCrLf
    msg = msg & "  v17: 12.Extra - connexions entrants (lectura)," & vbCrLf
    msg = msg & "  amb la cronologia girada; parella normalitzada" & vbCrLf
    msg = msg & "  v16: Mat. cornisa eliminat (I es sempre pedra)" & vbCrLf
    msg = msg & "  4.Dec: dues subseccions (arquitectonica / rupestre)" & vbCrLf
    msg = msg & "  sobre la mateixa taula, amb combos filtrats" & vbCrLf & vbCrLf
    msg = msg & "Els quatre trams arriben BUITS, no a 0: buit vol" & vbCrLf
    msg = msg & "dir que encara no s han mirat." & vbCrLf & vbCrLf
    msg = msg & "El valor per defecte dels elements es 0 (Absent)." & vbCrLf
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
    SetCap db, "T_DECORATIONS", "Span_Left", "Tram esquerra"
    SetCap db, "T_DECORATIONS", "Span_Above", "Tram damunt"
    SetCap db, "T_DECORATIONS", "Span_Right", "Tram dreta"
    SetCap db, "T_DECORATIONS", "Span_Below", "Tram davall"
    SetCap db, "T_DECORATIONS", "Outline_Geometry", "Geometria del trac"
    SetCap db, "T_DECORATIONS", "Color", "Color"
    SetCap db, "T_DECORATIONS", "Substrate", "Substrat"
    SetCap db, "T_DECORATIONS", "Notes", "Notes"
    SetCap db, "T_STRUCTURES", "Dec_Present", "Decoracio present"
    SetCap db, "T_STRUCTURES", "RockArt_Present", "Art rupestre present"
    SetCap db, "T_DECORATIONS", "Color_Secondary", "Color secundari"
    SetCap db, "T_STRUCTURES", "Arch_Notes", "Notes arquitectura"
    SetCap db, "T_STRUCTURES", "Finish_Notes", "Notes acabats"
    SetCap db, "T_STRUCTURES", "Condition_Notes", "Notes conservacio"
    SetCap db, "T_STRUCTURES", "Bio_Notes", "Notes bioarqueologia"
    SetCap db, "T_STRUCTURES", "Materials_Notes", "Notes materials"
    SetCap db, "T_STRUCTURES", "Systems_Notes", "Notes sistemes"
    SetCap db, "T_STRUCTURES", "Stone_Format", "Format pedra"
    SetCap db, "T_STRUCTURES", "Stone_Format_Secondary", "Format pedra secundari"
    SetCap db, "T_STRUCTURES", "Stone_Working", "Treball pedra"
    SetCap db, "T_STRUCTURES", "Access_Plane", "Pla d'acces"
    SetCap db, "T_STRUCTURES", "Portal_Orientation", "Orientacio portal"
    SetCap db, "T_STRUCTURES", "Portal_Position", "Posicio portal"
    SetCap db, "T_STRUCTURES", "Fabric", "Fabrica"
    SetCap db, "T_STRUCTURES", "Sys_Interface", "Sistema interficie"
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
    Set lh = CreateControl(frm, acLabel, acDetail, pg, "", C1, T, FW - 400, 260)
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

' rev. 6: per-tab notes field. Wide, multi-line, spanning both
' columns. Placed on the tab it belongs to because a note gets
' written when the case turns up, and the case turns up while
' filling in that tab - not on a separate review tab, which most
' notes would never reach.
Private Sub PCN(frm As String, pg As String, lbl As String, src As String, row As Integer)
    AddCtrl frm, pg, lbl, src, acTextBox, row, 1, 8600
    Dim c As Control
    On Error Resume Next
    For Each c In Forms(frm).Controls
        If c.ControlType = acTextBox Then
            If c.ControlSource = src Then
                c.Height = 700
                c.ScrollBars = 2
                c.CanGrow = True
            End If
        End If
    Next c
    On Error GoTo 0
End Sub

' rev. 5: metric field in METRES, shown with two decimals. Without
' the explicit format Access would render 0.62 as 0.6 and the
' observer would think the entry had been rounded.
Private Sub PCD(frm As String, pg As String, lbl As String, src As String, row As Integer, col As Integer)
    AddCtrl frm, pg, lbl, src, acTextBox, row, col, CW
    Dim c As Control
    On Error Resume Next
    For Each c In Forms(frm).Controls
        If c.ControlType = acTextBox Then
            If c.ControlSource = src Then
                c.Format = "0.00"
                c.DecimalPlaces = 2
            End If
        End If
    Next c
    On Error GoTo 0
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
    ' v13 (delta 7.4): dos subformularis sobre LA MATEIXA taula.
    ' El filtre i els combos restringits fan que l'usuari no veja
    ' mai una llista barrejada; la separacio analitica la dona
    ' Level_Type, no la separacio fisica de les dades.
    CreateDecSubform "F_DECORATIONS", False
    CreateDecSubform "F_ROCKART", True
    CreateFeatSubform
    CreateConnSubform
    CreateConnInSubform
    CreateLostSubform
    CreateDatingSubform
End Sub

' isRock = False -> decoracio arquitectonica (posicions no-ROC)
' isRock = True  -> pintura rupestre associada (posicions ROC)
' El Filter deixa fora les files de l'altra meitat; els RowSource
' dels combos impedeixen crear-ne de mal classificades.
Private Sub CreateDecSubform(SFRM As String, isRock As Boolean)
    On Error Resume Next: DoCmd.DeleteObject acForm, SFRM: On Error GoTo 0
    Dim f As Form: Set f = CreateForm()
    Dim tmp As String: tmp = f.Name
    ' El RecordSource porta el filtre: cada subformulari nomes veu
    ' la seua meitat de la taula. ND queda a la banda arquitectonica
    ' perque una posicio indeterminada d'una decoracio de facana no
    ' es art rupestre.
    ' v17: LEFT JOIN, no INNER. Amb l'INNER JOIN de la v16, les
    ' files sense posicio assignada queien de LES DUES meitats i
    ' no es podien ni veure ni completar - 16 de 56 a la copia
    ' local, quinze d'elles propostes automatiques de la migracio
    ' v11 amb la nota 'completar posicio'. Ara van a la meitat
    ' arquitectonica, que es on viu tambe ND.
    Dim rsq As String
    rsq = "SELECT D.* FROM T_DECORATIONS AS D LEFT JOIN L_STRUCT_BODY AS B ON D.ID_Struct_Body=B.ID WHERE "
    If isRock Then
        rsq = rsq & "B.Level_Type='ROC'"
    Else
        rsq = rsq & "(B.Level_Type<>'ROC' OR B.Level_Type Is Null)"
    End If
    f.RecordSource = rsq
    f.DefaultView = 2: f.ScrollBars = 2
    f.NavigationButtons = False
    f.Section(acDetail).Height = 400
    Dim T As Long: T = 50: Dim L As Long
    Dim lb As Control

    ' Sort_Order (v13) restaura l'ordre constructiu bottom-to-top;
    ' abans l'ordenacio per Level_Type + Name deixava el rafec al
    ' final i cada nivell en ordre alfabetic.
    Dim posq As String
    posq = "SELECT ID, Name_VAL FROM L_STRUCT_BODY WHERE Level_Type"
    If isRock Then
        posq = posq & "='ROC'"
    Else
        posq = posq & "<>'ROC'"
    End If
    posq = posq & " ORDER BY Sort_Order"

    L = 40
    Dim c1 As Control: Set c1 = CreateControl(tmp, acComboBox, acDetail, "", "", L + 1060, T, 2200, 315)
    c1.ControlSource = "ID_Struct_Body"
    c1.RowSourceType = "Table/Query"
    c1.RowSource = posq
    c1.BoundColumn = 1: c1.ColumnCount = 2: c1.ColumnWidths = "0cm;4cm": c1.LimitToList = True
    On Error Resume Next: c1.Name = "ID_Struct_Body": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "ID_Struct_Body", "", L, T + 15, 1000, 260)
    lb.Caption = "Posicio"
    L = L + 3360

    Dim c2 As Control: Set c2 = CreateControl(tmp, acComboBox, acDetail, "", "", L + 1060, T, 2000, 315)
    c2.ControlSource = "ID_Dec_Type"
    c2.RowSourceType = "Table/Query"
    ' Els repertoris NO son compartits (comprovat contra el corpus).
    ' Els tipus rupestres duen prefix 'RA '; els generics que
    ' serveixen a totes dues bandes (color pla, banda pintada, ND)
    ' es repeteixen a proposit en les dues llistes.
    If isRock Then
        c2.RowSource = "SELECT ID, Name_VAL FROM L_DEC_TYPE WHERE Name Like 'RA *' OR Name='Plain colour field' OR Name='Painted band' OR Name='ND' ORDER BY Name"
    Else
        c2.RowSource = "SELECT ID, Name_VAL FROM L_DEC_TYPE WHERE Name Not Like 'RA *' ORDER BY Name"
    End If
    c2.BoundColumn = 1: c2.ColumnCount = 2: c2.ColumnWidths = "0cm;4cm": c2.LimitToList = True
    On Error Resume Next: c2.Name = "ID_Dec_Type": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "ID_Dec_Type", "", L, T + 15, 1000, 260)
    lb.Caption = "Tipus dec."
    L = L + 3200

    ' v17 (delta B2): a la meitat rupestre el control s'OCULTA, no
    ' es limita a quedar desactivat. Body_No indexa el cos
    ' constructiu i una pintura sobre la penya no esta en cap cos:
    ' ensenyar una casella morta a cada fila no informa de res i
    ' ocupa l'amplaria que ara necessiten els trams. Es crea
    ' igualment perque la taula es la mateixa, i la regla 37
    ' detecta el que hi entre per un altre costat.
    Dim c3 As Control: Set c3 = CreateControl(tmp, acTextBox, acDetail, "", "", L + 660, T, 500, 315)
    c3.ControlSource = "Body_No"
    On Error Resume Next: c3.Name = "Body_No": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "Body_No", "", L, T + 15, 600, 260)
    lb.Caption = "Num. cos"
    If isRock Then
        c3.Visible = False
        lb.Visible = False
    Else
        L = L + 1300
    End If

    ' v17 (delta A1): ELS QUATRE TRAMS DEL CONTORN, nomes a la
    ' meitat rupestre. Substitueixen Position_Relative, que servia
    ' les dues meitats amb referents distints i per tant no en
    ' servia cap: a la banda arquitectonica 'esquerra' designava
    ' QUINA INSTANCIA d'un parell portava el motiu, i a la
    ' rupestre ON ESTA LA PINTURA respecte del volum construit.
    ' Quatre camps i no un valor unic perque un valor unic NO POT
    ' DISTINGIR 'la banda no cobria la base' de 'el tram de base
    ' no es observable', i eixa es exactament la diferencia entre
    ' una U invertida i un anell tancat mal conservat.
    ' U invertida, envoltant i flanquejant NO s'emmagatzemen: es
    ' deriven a QRY_20, de manera que una U amb un 9 a la base hi
    ' apareix com a CANDIDATA i no com a afirmacio.
    ' CONVENCIO, corregida respecte de la v16: ESQUERRA I DRETA
    ' DES DE L'OBSERVADOR SITUAT DAVANT DEL PLA EXPOSAT, mai des
    ' del pla d'acces - la convencio v16 les invertia en cada
    ' estructura amb l'acces desviat de la facana.
    If isRock Then
        DecSpan tmp, "Span_Left", "Esquerra", L, T
        L = L + 1500
        DecSpan tmp, "Span_Above", "Damunt", L, T
        L = L + 1500
        DecSpan tmp, "Span_Right", "Dreta", L, T
        L = L + 1500
        DecSpan tmp, "Span_Below", "Davall", L, T
        L = L + 1500

        ' v17 (delta A2): geometria del trac. A TOTES les files
        ' rupestres i no filtrada per tipus de decoracio: les tres
        ' files del corpus que descriuen una U invertida porten
        ' TRES TIPUS DIFERENTS, de manera que qualsevol filtre per
        ' tipus les deixaria fora precisament a elles.
        ' Un contorn ortogonal AFIRMA UN REFERENT CONSTRUIT
        ' rectangular; un de corb no afirma res i pot estar
        ' resseguint un rebaix natural. Es el camp que decideix si
        ' un cas de pintura perimetral sense fabrica es un registre
        ' d'estructura o un d'art rupestre.
        Dim c3c As Control: Set c3c = CreateControl(tmp, acComboBox, acDetail, "", "", L + 900, T, 1500, 315)
        c3c.ControlSource = "Outline_Geometry": c3c.RowSourceType = "Value List"
        c3c.RowSource = "Orthogonal;Ortogonal;Curvilinear;Corb;Irregular;Irregular;ND;Indeterminada"
        c3c.ColumnCount = 2: c3c.BoundColumn = 1: c3c.ColumnWidths = "0cm;3.5cm": c3c.LimitToList = True
        On Error Resume Next: c3c.Name = "Outline_Geometry": On Error GoTo 0
        Set lb = CreateControl(tmp, acLabel, acDetail, "Outline_Geometry", "", L, T + 15, 840, 260)
        lb.Caption = "Geometria"
        L = L + 2500
    End If

    Dim c4 As Control: Set c4 = CreateControl(tmp, acComboBox, acDetail, "", "", L + 660, T, 1300, 315)
    ' v13: 'Both' retirat (delta 7.6). La parella ordenada Color +
    ' Color_Secondary el substitueix amb avantatge, perque diu quin
    ' domina. Cap fila del corpus v12 el portava: migracio nul-la.
    c4.ControlSource = "Color": c4.RowSourceType = "Value List"
    c4.RowSource = "Red;Roig;White;Blanc;Ochre;Ocre;None;Cap;ND;Indeterminat"
    c4.ColumnCount = 2: c4.BoundColumn = 1: c4.ColumnWidths = "0cm;3cm": c4.LimitToList = True
    On Error Resume Next: c4.Name = "Color": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "Color", "", L, T + 15, 600, 260)
    lb.Caption = "Color"
    L = L + 2100

    ' v13: color secundari per als casos bicroms amb els colors
    ' junts o contigus dins d'un mateix motiu (camp clar amb vora
    ' roja), on partir-ho en dues files inventaria dos motius on
    ' n'hi ha un i duplicaria la posicio.
    Dim c4b As Control: Set c4b = CreateControl(tmp, acComboBox, acDetail, "", "", L + 900, T, 1300, 315)
    c4b.ControlSource = "Color_Secondary": c4b.RowSourceType = "Value List"
    ' rev. 6: sense ND. Un color secundari indeterminat no diu res
    ' util; si es veu que hi ha un segon color pero no se'n
    ' distingeix el to, el registre honest es deixar-ho buit i
    ' explicar-ho a Notes. Aixi el camp nomes afirma "hi ha segon
    ' color i es aquest", i buit/ple equival a monocrom/policrom
    ' sense necessitat de cap camp derivat.
    c4b.RowSource = "Red;Roig;White;Blanc;Ochre;Ocre"
    c4b.ColumnCount = 2: c4b.BoundColumn = 1: c4b.ColumnWidths = "0cm;3cm": c4b.LimitToList = True
    On Error Resume Next: c4b.Name = "Color_Secondary": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "Color_Secondary", "", L, T + 15, 840, 260)
    lb.Caption = "Color sec."
    L = L + 2400

    Dim c5 As Control: Set c5 = CreateControl(tmp, acComboBox, acDetail, "", "", L + 900, T, 1600, 315)
    ' v16 (delta 4.6): a la meitat rupestre el substrat deixa de
    ' ser una tria i passa a ser un VERIFICADOR. Revoc i pedra de
    ' parament son impossibles en una fila ROC per definicio del
    ' criteri d'associacio: si el pigment toca la fabrica, allo
    ' es pigment de l'estructura i va a l'altra meitat (test 1).
    ' Nomes queden Penya i Mixt, i Mixt exclusivament per a la
    ' posicio Solapada, que per definicio abasta fabrica i roca.
    ' Les regles 38a i 38b detecten les files mal encaminades:
    ' una incoherencia abans silenciosa passa a ser un error
    ' detectat, de manera que restringir la llista no perd res.
    c5.ControlSource = "Substrate": c5.RowSourceType = "Value List"
    If isRock Then
        c5.RowSource = "Bedrock;Penya;Mixed;Mixt"
    Else
        c5.RowSource = "Plaster;Revoc;Masonry stone;Pedra de parament;Bedrock;Penya;ND;Indeterminat"
    End If
    c5.ColumnCount = 2: c5.BoundColumn = 1: c5.ColumnWidths = "0cm;4cm": c5.LimitToList = True
    On Error Resume Next: c5.Name = "Substrate": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "Substrate", "", L, T + 15, 840, 260)
    lb.Caption = "Substrat"
    L = L + 2700

    Dim c6 As Control: Set c6 = CreateControl(tmp, acTextBox, acDetail, "", "", L + 660, T, 1400, 315)
    c6.ControlSource = "Notes"
    On Error Resume Next: c6.Name = "Notes": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "Notes", "", L, T + 15, 600, 260)
    lb.Caption = "Notes"
    f.Width = L + 2200

    DoCmd.Save acForm, tmp: DoCmd.Close acForm, tmp
    DoCmd.Rename SFRM, acForm, tmp
    Debug.Print "[OK] " & SFRM
End Sub

' Un tram del contorn. Domini de tres valors, el mateix que la
' resta de camps observacionals: 0 absent, 1 present, 9 no
' observable. Sense valor per defecte - buit vol dir que encara
' no s'ha mirat, i eixe es precisament l'estat que cal poder
' distingir del 0.
Private Sub DecSpan(tmp As String, src As String, cap As String, L As Long, T As Long)
    Dim c As Control
    Set c = CreateControl(tmp, acComboBox, acDetail, "", "", L + 780, T, 700, 315)
    c.ControlSource = src: c.RowSourceType = "Value List"
    c.RowSource = DOM3
    c.ColumnCount = 2: c.BoundColumn = 1: c.ColumnWidths = "0cm;3cm": c.LimitToList = True
    On Error Resume Next: c.Name = src: On Error GoTo 0
    Dim lb As Control
    Set lb = CreateControl(tmp, acLabel, acDetail, src, "", L, T + 15, 720, 260)
    lb.Caption = cap
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
    ' rev. 6: Vertical association afegida. Registra una relacio de
    ' verticalitat observable entre estructures (proximitat, posicio
    ' relativa, morfologia del farallo) sense afirmar el mecanisme:
    ' la hipotesi de politja o de suport d'escala continua sent
    ' especulativa i va a T_ARCH_FEATURES, pero l'associacio SI que
    ' s'observa. Confidence en registra la seguretat. Material per a
    ' OE3.
    ' CRITERI quan concorren dues relacions (junta vertical entre
    ' fabriques sobre una base compartida): LA JUNTA MANA SOBRE EL
    ' SUPORT, perque la junta porta la direccio cronologica (regla
    ' 20) i el suport compartit no.
    ' v17: Horizontal association, simetrica de la vertical.
    ' Registra una relacio d'horitzontalitat OBSERVABLE entre
    ' dues estructures - alineacio sobre la mateixa cornisa
    ' natural o repisa, sense contacte fisic - sense afirmar-ne
    ' el mecanisme, igual que la vertical.
    c3.RowSource = "Abutted vertical joint;Junta vertical adossada;Superposition;Superposicio;Bonded joint;Junta travada;Shared support;Suport compartit;Vertical association;Associacio de verticalitat;Horizontal association;Associacio d'horitzontalitat;Aerial connection;Connexio aeria;Other (see Notes);Altres (veure notes)"
    c3.ColumnCount = 2: c3.BoundColumn = 1: c3.ColumnWidths = "0cm;5cm": c3.LimitToList = True
    On Error Resume Next: c3.Name = "Connection_Type": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "Connection_Type", "", L, T + 15, 1140, 260)
    lb.Caption = "Tipus connexio"

    ' Direction is read from the joint: the structure showing the
    ' untoothed joint against the other's wall face is the later one.
    ' v17 (delta D1): LA INSTRUCCIO 'NO INVERTIR L'ORDRE'
    ' DESAPAREIX. Ara el formulari NORMALITZA la parella a
    ' BeforeUpdate - l'ID menor sempre a A - i, quan les gira,
    ' GIRA TAMBE LA CRONOLOGIA, de manera que el sentit es
    ' conserva. Amb l'ordre normalitzat, l'index unic
    ' UQ_CONN_PAIR fa impossible entrar la mateixa connexio dues
    ' vegades: era el comportament per defecte de la v16, i les
    ' dues uniques files de la copia local eren la mateixa
    ' connexio entrada des de cada extrem.
    L = 10300
    Dim c4 As Control: Set c4 = CreateControl(tmp, acComboBox, acDetail, "", "", L + 1300, T, 2200, 315)
    c4.ControlSource = "Chrono_Relation": c4.RowSourceType = "Value List"
    c4.RowSource = "A is earlier;A es anterior;B is earlier;B es anterior;Contemporary;Contemporanis;Undetermined;Indeterminat"
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
'  v17 - F_CONN_IN: LES CONNEXIONS VISTES DES DE L'ALTRE EXTREM
'
'  Una connexio es una aresta entre dues estructures, pero es
'  guarda com a parella ordenada i en la v16 nomes es mostrava des
'  d'A. Des de B no es veia res, aixi que es tornava a entrar: la
'  duplicacio no era descuit sino el comportament per defecte, i
'  les dues uniques files de la copia local eren la mateixa
'  connexio entrada des de cada extrem.
'
'  NOMES LECTURA a proposit. Editar la mateixa fila des dels dos
'  costats reintroduiria per la porta del darrere el problema que
'  la normalitzacio resol. Per a modificar-la, s'obri el registre
'  on es va entrar - i el codi hi apareix.
'
'  LA CRONOLOGIA ES MOSTRA GIRADA. Es text calculat i no valor
'  emmagatzemat, de manera que no trenca la convencio d'idiomes:
'  el que es guarda continua sent angles a T_CONNECTIONS.
' ================================================================
Private Sub CreateConnInSubform()
    Const SFRM = "F_CONN_IN"
    On Error Resume Next: DoCmd.DeleteObject acForm, SFRM: On Error GoTo 0
    Dim f As Form: Set f = CreateForm()
    Dim tmp As String: tmp = f.Name
    Dim rsq As String
    rsq = "SELECT C.ID_Struct_B AS Link_ID, E.Code AS Other_Code, "
    rsq = rsq & "C.Connection_Type, C.Confidence, C.Notes, "
    rsq = rsq & "IIF(C.Chrono_Relation='A is earlier','Aquesta es posterior', "
    rsq = rsq & "IIF(C.Chrono_Relation='B is earlier','Aquesta es anterior', "
    rsq = rsq & "IIF(C.Chrono_Relation='Contemporary','Contemporanis','Indeterminat'))) AS Chrono_Seen "
    rsq = rsq & "FROM T_CONNECTIONS AS C "
    rsq = rsq & "INNER JOIN T_STRUCTURES AS E ON C.ID_Struct_A=E.ID"
    f.RecordSource = rsq
    f.DefaultView = 2: f.ScrollBars = 2
    f.NavigationButtons = False: f.Width = 12000
    f.AllowEdits = False: f.AllowAdditions = False: f.AllowDeletions = False
    f.Section(acDetail).Height = 400
    Dim T As Long: T = 50: Dim L As Long
    Dim lb As Control
    Dim c As Control

    L = 40
    Set c = CreateControl(tmp, acTextBox, acDetail, "", "", L + 1200, T, 2000, 315)
    c.ControlSource = "Other_Code"
    On Error Resume Next: c.Name = "Other_Code": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "Other_Code", "", L, T + 15, 1140, 260)
    lb.Caption = "Registrada des de"

    L = 3400
    Set c = CreateControl(tmp, acTextBox, acDetail, "", "", L + 1200, T, 2400, 315)
    c.ControlSource = "Connection_Type"
    On Error Resume Next: c.Name = "Connection_Type_In": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "Connection_Type_In", "", L, T + 15, 1140, 260)
    lb.Caption = "Tipus connexio"

    L = 7200
    Set c = CreateControl(tmp, acTextBox, acDetail, "", "", L + 1240, T, 2400, 315)
    c.ControlSource = "Chrono_Seen"
    On Error Resume Next: c.Name = "Chrono_Seen": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "Chrono_Seen", "", L, T + 15, 1180, 260)
    lb.Caption = "Cronologia"

    L = 11000
    Set c = CreateControl(tmp, acTextBox, acDetail, "", "", L + 900, T, 1200, 315)
    c.ControlSource = "Confidence"
    On Error Resume Next: c.Name = "Confidence_In": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "Confidence_In", "", L, T + 15, 840, 260)
    lb.Caption = "Confianca"

    L = 13200
    Set c = CreateControl(tmp, acTextBox, acDetail, "", "", L + 660, T, 2200, 315)
    c.ControlSource = "Notes"
    On Error Resume Next: c.Name = "Notes_In": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "Notes_In", "", L, T + 15, 600, 260)
    lb.Caption = "Notes"
    f.Width = 16200

    DoCmd.Save acForm, tmp: DoCmd.Close acForm, tmp
    DoCmd.Rename SFRM, acForm, tmp
    Debug.Print "[OK] F_CONN_IN"
End Sub

' NOU rev. 5. T_DATING existia com a taula des de la v1 pero no
' tenia cap subformulari: les datacions calia entrar-les obrint la
' taula a ma. Ara penja de 8.Cron i C14 el governa, que es el que
' l'usuari demanava - amb C14 = No, no te sentit poder escriure als
' camps de la mostra.
Private Sub CreateDatingSubform()
    Const SFRM = "F_DATING"
    On Error Resume Next: DoCmd.DeleteObject acForm, SFRM: On Error GoTo 0
    Dim f As Form: Set f = CreateForm()
    Dim tmp As String: tmp = f.Name
    f.RecordSource = "T_DATING"
    f.DefaultView = 2: f.ScrollBars = 2
    f.NavigationButtons = False: f.Width = 16000
    f.Section(acDetail).Height = 400
    Dim T As Long: T = 50: Dim L As Long
    Dim lb As Control

    L = 40
    Dim d1 As Control: Set d1 = CreateControl(tmp, acComboBox, acDetail, "", "", L + 800, T, 2000, 315)
    d1.ControlSource = "Sample_Type": d1.RowSourceType = "Value List"
    d1.RowSource = "Charcoal;Carbo;Bone;Os;Textile;Textil;Wood;Fusta;Vegetal fiber;Fibra vegetal;ND;Indeterminat"
    d1.ColumnCount = 2: d1.BoundColumn = 1: d1.ColumnWidths = "0cm;4cm": d1.LimitToList = True
    On Error Resume Next: d1.Name = "Sample_Type": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "Sample_Type", "", L, T + 15, 740, 260)
    lb.Caption = "Mostra"

    L = 3000
    Dim d2 As Control: Set d2 = CreateControl(tmp, acTextBox, acDetail, "", "", L + 700, T, 1200, 315)
    d2.ControlSource = "Date_BP"
    On Error Resume Next: d2.Name = "Date_BP": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "Date_BP", "", L, T + 15, 640, 260)
    lb.Caption = "Data BP"

    L = 5100
    Dim d3 As Control: Set d3 = CreateControl(tmp, acTextBox, acDetail, "", "", L + 800, T, 900, 315)
    d3.ControlSource = "Sigma1_Start"
    On Error Resume Next: d3.Name = "Sigma1_Start": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "Sigma1_Start", "", L, T + 15, 740, 260)
    lb.Caption = "1s inici"

    L = 6900
    Dim d4 As Control: Set d4 = CreateControl(tmp, acTextBox, acDetail, "", "", L + 660, T, 900, 315)
    d4.ControlSource = "Sigma1_End"
    On Error Resume Next: d4.Name = "Sigma1_End": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "Sigma1_End", "", L, T + 15, 600, 260)
    lb.Caption = "1s fi"

    L = 8560
    Dim d5 As Control: Set d5 = CreateControl(tmp, acTextBox, acDetail, "", "", L + 800, T, 900, 315)
    d5.ControlSource = "Sigma2_Start"
    On Error Resume Next: d5.Name = "Sigma2_Start": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "Sigma2_Start", "", L, T + 15, 740, 260)
    lb.Caption = "2s inici"

    L = 10360
    Dim d6 As Control: Set d6 = CreateControl(tmp, acTextBox, acDetail, "", "", L + 660, T, 900, 315)
    d6.ControlSource = "Sigma2_End"
    On Error Resume Next: d6.Name = "Sigma2_End": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "Sigma2_End", "", L, T + 15, 600, 260)
    lb.Caption = "2s fi"

    L = 12020
    Dim d7 As Control: Set d7 = CreateControl(tmp, acTextBox, acDetail, "", "", L + 700, T, 1500, 315)
    d7.ControlSource = "Lab_Reference"
    On Error Resume Next: d7.Name = "Lab_Reference": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "Lab_Reference", "", L, T + 15, 640, 260)
    lb.Caption = "Lab."

    L = 14320
    Dim d8 As Control: Set d8 = CreateControl(tmp, acTextBox, acDetail, "", "", L + 640, T, 1000, 315)
    d8.ControlSource = "Bibliog_Reference"
    On Error Resume Next: d8.Name = "Bibliog_Reference": On Error GoTo 0
    Set lb = CreateControl(tmp, acLabel, acDetail, "Bibliog_Reference", "", L, T + 15, 580, 260)
    lb.Caption = "Bibl."

    DoCmd.Save acForm, tmp: DoCmd.Close acForm, tmp
    DoCmd.Rename SFRM, acForm, tmp
    Debug.Print "[OK] F_DATING"
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
    c2.RowSource = "SELECT ID, Name_VAL FROM L_LOST_EVIDENCE ORDER BY ID"
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
    c4.RowSource = "SELECT ID, Name_VAL FROM L_STRUCT_BODY ORDER BY Sort_Order"
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

    ' v13: dues subseccions a 4.Dec, una per meitat de la taula.
    Dim sf1 As Control
    Set sf1 = CreateControl(tmp, acSubform, acDetail, "pgDec", "", C1, MT + 2 * RG + 200, 12400, 2700)
    sf1.SourceObject = "F_DECORATIONS"
    sf1.LinkMasterFields = "ID": sf1.LinkChildFields = "ID_Structure"

    ' rev. 5: subformulari de datacions, governat per C14.
    Dim sfDat As Control
    Set sfDat = CreateControl(tmp, acSubform, acDetail, "pgCron", "", C1, MT + 7 * RG + 200, 12400, 2400)
    On Error Resume Next: sfDat.Name = "sfDating": On Error GoTo 0
    sfDat.SourceObject = "F_DATING"
    sfDat.LinkMasterFields = "ID": sfDat.LinkChildFields = "ID_Structure"

    Dim sf1b As Control
    Set sf1b = CreateControl(tmp, acSubform, acDetail, "pgDec", "", C1, MT + 11 * RG + 200, 12400, 2700)
    On Error Resume Next: sf1b.Name = "sfRockArt": On Error GoTo 0
    sf1b.SourceObject = "F_ROCKART"
    sf1b.LinkMasterFields = "ID": sf1b.LinkChildFields = "ID_Structure"

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

    ' v17: connexions entrants, nomes lectura. La cronologia es
    ' mostra GIRADA des d'aquest costat: si des de l'altra
    ' estructura s'ha dit que ella es l'anterior, aci ha de
    ' llegir-se 'aquesta es posterior'. Sense la inversio la
    ' llista mentiria cada vegada que hi haja direccio.
    Dim sf3b As Control
    Set sf3b = CreateControl(tmp, acSubform, acDetail, "pgExtra", "", C1, MT + 15 * RG + 340, 12000, 1700)
    On Error Resume Next: sf3b.Name = "sfConnIn": On Error GoTo 0
    sf3b.SourceObject = "F_CONN_IN"
    sf3b.LinkMasterFields = "ID": sf3b.LinkChildFields = "Link_ID"

    ' NOU v12: l'evidencia dels elements desapareguts (regla 2), al
    ' costat de les connexions perque totes dues son taules filles.
    Dim sf4 As Control
    Set sf4 = CreateControl(tmp, acSubform, acDetail, "pgExtra", "", C1, MT + 21 * RG + 140, 12000, 2400)
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
    ' v17: traslladat aci des del bloc de fases, on no pintava
    ' res. Es un qualificador del PLA DE FACANA (rev. 7): el
    ' recul afecta el parament sencer i el portal hi queda
    ' inscrit, de manera que no es gateja per Sys_Portal.
    PC9 f, "pgArq", "Marc reculat facana:", "Recessed_Frame", 5, 2
    ' v16 (delta 3.2): FACANA = EL PLA EXPOSAT, el que mira a la
    ' vall. Una estructura del corpus te l'obertura al mur estret
    ' perpendicular al farallo i la decoracio al mur llarg
    ' paral-lel: fixar la facana per l'obertura deixaria aquest
    ' camp sense apuntar a la vall i degradaria tota H06 per salvar
    ' la definicio d'un element. El flanc de facana (L) deixa de
    ' dependre de l'obertura: on l'obertura es a la facana, L son
    ' els paraments que la flanquegen; on no, L es el parament
    ' corregut.
    SH  f, "pgArq", "Facana i paisatge (observacional)", 6
    PCV f, "pgArq", "Orientacio facana:", "Facade_Orientation", 7, 1, "N;N;NE;NE;E;E;SE;SE;S;S;SW;SO;W;O;NW;NO;ND;Indeterminada"
    PCV f, "pgArq", "Visibilitat vall:",  "Visibility_Valley",  7, 2, "High;Alta;Medium;Mitjana;Low;Baixa;ND;Indeterminada"
    ' Que l'acces no estiga al pla exposat NO es una anomalia a
    ' absorbir: diu que la circulacio mana sobre l'exhibicio -
    ' s'entra per on es camina, per la repisa, i s'exhibeix cap a
    ' on es mira.
    ' v17: LA SEGONA FAENA HA DESAPAREGUT. A la v16 aquest camp
    ' declarava tambe el pla de referencia de la posicio relativa
    ' de 4.Dec; eixe camp s'ha retirat, i els trams que el
    ' substitueixen es llegeixen des del pla EXPOSAT. La
    ' convencio v16 invertia esquerra i dreta precisament en les
    ' estructures amb l'acces desviat de la facana.
    ' v17: aquest camp i l'orientacio del portal es desactiven
    ' quan Sys_Chamber es declara Absent o No aplicable - MAI amb
    ' el camp buit, que vol dir 'encara no s'ha avaluat'.
    ' Orientacio de facana i visibilitat de la vall NO es
    ' desactiven: la facana es el PLA EXPOSAT i una massa basal
    ' en te i mira cap a algun lloc; tancar-les eliminaria de
    ' H06 el seu grup de control natural.
    PCV f, "pgArq", "Pla d'acces:", "Access_Plane", 8, 1, "Facade;Facana;Return wall;Mur lateral;Rear;Fons;ND;Indeterminat"
    ' NO gatejat darrere de Sys_Portal, pel mateix motiu que
    ' Recessed_Frame no ho esta: una obertura arrasada te
    ' orientacio coneguda. La DIVERGENCIA amb l'orientacio de
    ' facana es ara calculable, i eixa es la variable (regla 36).
    PCV f, "pgArq", "Orientacio portal:", "Portal_Orientation", 8, 2, "N;N;NE;NE;E;E;SE;SE;S;S;SW;SO;W;O;NW;NO;ND;Indeterminada"
    ' v17: punt obert heretat tancat. Descentrar un portal es una
    ' DECISIO DE PLANIFICACIO, no un accident de fabrica: toca
    ' H01 i H04. Mateixa convencio d'observador que els trams de
    ' 4.Dec: DES DE DAVANT DEL PLA EXPOSAT.
    PCV f, "pgArq", "Posicio portal:", "Portal_Position", 9, 1, "Centred;Centrat;Off-centre left;Descentrat a l'esquerra;Off-centre right;Descentrat a la dreta;NA;No aplicable;ND;Indeterminada"
    SH  f, "pgArq", "Maconeria i morter (T&A 2017 / H01, H04)", 10
    ' v16 (delta 1): QUINA PEDRA ES, davant de com s'apila. El
    ' test primari de format es FUNCIONAL - quantes peces fan una
    ' filada - perque es llig directament a la facana i no exigeix
    ' mesurar res: irregular (sense dues cares planes
    ' subparal-leles), tabular (una pedra, una filada), laminar
    ' (cal apilar-ne diverses per filada). Desempat quantitatiu
    ' per als dubtosos: gruix <= 1/4 de la dimensio major ->
    ' laminar. ABAST: EL PARAMENT, no els elements singulars
    ' (dintell, superficie de rafec, tancament), que ja tenen el
    ' seu camp de material. Per aixo 'lloses de gran format' no hi
    ' es: era una categoria de dimensio dins d'un eix de
    ' morfologia.
    PCV f, "pgArq", "Format pedra:",    "Stone_Format",           11, 1, "Irregular stones;Pedres irregulars;Tabular blocks;Blocs tabulars;Laminar slabs;Lloses laminars;ND;Indeterminat"
    ' Parella ordenada com ID_Support, i per aixo SENSE valor
    ' Mixed: el projecte ja ha retirat dues vegades el valor que
    ' la parella substitueix (Combined a L_SUPPORT en v8, Both al
    ' color en v13). Dominancia = major SUPERFICIE de parament, no
    ' major nombre de peces: amb lloses menudes i blocs grans,
    ' comptar peces inverteix el resultat. Sense ND: buit ja vol
    ' dir 'cap segon format' (regles 32-33).
    PCV f, "pgArq", "Format secundari:", "Stone_Format_Secondary", 11, 2, "Irregular stones;Pedres irregulars;Tabular blocks;Blocs tabulars;Laminar slabs;Lloses laminars"
    ' ORDINAL: Unworked < Semi-dressed < Dressed, utilitzable com a
    ' proxy d'inversio de treball. Mixed i ND son FORA D'ESCALA i
    ' cauen d'eixes analisis, com els Not applicable. Registra
    ' EVIDENCIA POSITIVA de treball (traces d'eina, aristes vives
    ' regulars, cares que trenquen el pla de fractura natural);
    ' sense eixes traces la cara plana s'atribueix a la fractura i
    ' el valor es Unworked, que aixi no afirma 'ningu no la va
    ' tocar' sino 'no hi ha evidencia que la tocaren'. ND queda per
    ' al dubte genuin: un gres estratificat fractura en cares
    ' indistingibles del carejat. Cap valor 'no observable': el 9
    ' existeix per a protegir el 0, i aci no hi ha 0 a protegir.
    ' Doc_Basis i Facade_Observability ja separen 'no s'hi veia' de
    ' 'no era decidible'.
    PCV f, "pgArq", "Treball pedra:",   "Stone_Working",          12, 1, "Unworked;Sense treballar;Semi-dressed;Semiescairada;Dressed;Escairada;Mixed;Mixt;ND;Indeterminat"
    PCV f, "pgArq", "Qualitat maconeria:", "Masonry_Quality", 12, 2, "Good;Bona;Moderate;Moderada;Poor;Pobra;ND;Tipus indeterminat"
    PCV f, "pgArq", "Tipus aparell:",      "Masonry_Type",    13, 1, "Well-coursed;Filades regulars;Irregular-coursed;Filades irregulars;Uncoursed;Sense filades;Mixed;Mixt;ND;Tipus indeterminat"
    ' rev. 5: el morter no es una capa que es perd per zones sino
    ' un ATRIBUT DE LA TECNICA de fabrica: un mur en sec no ho es a
    ' trossos. El que varia amb la conservacio es la visibilitat,
    ' que ja registren Facade_Observability i Mortar_Notes.
    PC9 f, "pgArq", "Morter present:",     "Mortar_Present",  13, 2
    PCV f, "pgArq", "Tipus morter:",       "Mortar_Type",     14, 1, "Mud;Fang;Mud with gravel;Fang amb grava;Mud with organics;Fang amb organics;None dry-laid;Cap, en sec;ND;Tipus indeterminat"
    PC9 f, "pgArq", "Ripio / falques:",    "Chinking_Stones", 14, 2
    PCT f, "pgArq", "Notes morter:",       "Mortar_Notes",    15, 1
    ' v17 (delta F2): EL COMPTADOR QUE SUBSTITUEIX T_BODIES.
    ' El 'Mixt' del tipus d'aparell i del treball de la pedra diu
    ' que el registre te mes d'un valor; aquest diu SI EIXA
    ' BARREJA COINCIDEIX AMB LA DIVISIO EN COSSOS.
    ' La distincio entre cossos / dins d'un cos es tot el sentit
    ' de l'operacio: si la divergencia resulta estar sobretot
    ' DINS d'un mateix cos, una taula de cossos no resoldria res
    ' - el mateix 'Mixt' reapareixeria un nivell mes avall.
    ' FRONTERA AMB LES FASES: la divergencia es una OBSERVACIO
    ' (la pedra canvia, i aixo es veu); la fase es una
    ' INTERPRETACIO (hi va haver dos moments, i aixo s'argumenta).
    ' Pot haver-hi canvi de fabrica sense cap fase afirmada: un
    ' canvi de proveiment, o dos paletes el mateix dia.
    ' Es desactiva amb un sol cos: no hi ha res a comparar.
    ' Els tres valors plurals son EXCLOENTS perque el criteri no
    ' es ON hi ha divergencia - un cos plural per dins tambe ho
    ' es respecte del vei, i aixi els valors se solapaven - sino
    ' SI LA DIVISIO COINCIDEIX AMB ELS COSSOS.
    SH  f, "pgArq", "Fabrica (F1)", 16
    PCV f, "pgArq", "Fabrica:", "Fabric", 17, 1, "Single;Unica;Between bodies only;Multiple per cossos;Within a body;Multiple dins d'un cos;Not observable;No observable"
    SH  f, "pgArq", "Fases constructives (H03/H04)", 18
    PCT f, "pgArq", "Num. fases:",     "Construction_Phases", 19, 1
    ' v17: la llista v16 no tenia valor per a les dues
    ' observacions mes frequents en camp - la junta vertical
    ' adossada i el canvi de fabrica -, de manera que es podia
    ' declarar una fase sense res registrable al darrere.
    PCV f, "pgArq", "Evidencia fase:", "Phase_Evidence", 19, 2, "Abutted vertical joint;Junta vertical adossada;Superposition;Superposicio;Fabric change;Canvi de fabrica;Mortar difference;Diferencia de morter;Blocked or altered opening;Obertura tapiada o modificada;Added mass or annex;Massa afegida o annex;Radiocarbon;C14;Stratigraphy;Estratigrafia;ND;Indeterminada"
    SH  f, "pgArq", "Observacions sobre morfologia, maconeria i fases", 20
    PCN f, "pgArq", "Notes:", "Arch_Notes", 21
End Sub

' TAB 3 - SURFACE TREATMENTS
' Active for every record class (4.4): a rock art panel is DEFINED by
' its pigment, and a structural trace can keep pigment on the corbel.
Private Sub FillAcab(f As String)
    ' rev. 5: revoc i pigment tornen a tres valors. El 2 duplicava
    ' el que Plaster_Extent / Pigment_Extent ja registren amb
    ' "Traces", i el 3 no es assertable: l'unica evidencia de
    ' pigment es pigment, aixi que 3 i 0 es confondrien.
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
    SH  f, "pgAcab", "Observacions sobre acabats", 8
    PCN f, "pgAcab", "Notes:", "Finish_Notes", 9
End Sub

' TAB 4 - DECORATION: the subform only (7.4)
Private Sub FillDec(f As String)
    ' v13: dues subseccions sobre la mateixa taula (delta 7.4).
    ' Cada judici agregat governa el seu subformulari; les regles
    ' 22-23 (no-ROC) i 28-29 (ROC) vigilen cada meitat per separat.
    ' rev. 5: judicis agregats, tres valors. Cada motiu ja te la
    ' seua fila amb el seu propi estat: un 2 al camp agregat
    ' duplicaria informacio que viu al detall.
    SH  f, "pgDec", "Decoracio arquitectonica (T_DECORATIONS)", 0
    PC9 f, "pgDec", "Decoracio present:", "Dec_Present", 1, 1
    SH  f, "pgDec", "Pintura rupestre associada (posicions ROC)", 9
    PC9 f, "pgDec", "Art rupestre present:", "RockArt_Present", 10, 1
End Sub

' TAB 5 - CONSERVATION AND OBSERVABILITY
Private Sub FillEst(f As String)
    SH f, "pgEst", "Estat de conservacio i alteracions", 0
    PCC f, "pgEst", "Estat estructura:",      "ID_Arch_Status",     1, 1
    ' v17 (delta D2): L'ESTAT DELS VESTIGIS MOBLES SE'N VA A
    ' 7.Mat. La porta (Cultural_Materials_Present) vivia a 7.Mat
    ' i el camp que governa vivia aci, de manera que es preguntava
    ' en quin estat estan ABANS de si n'hi ha - i la contradiccio
    ' era indetectable, perque el gating de nivell 2 opera dins
    ' d'una pestanya. 5.Estat queda per a l'ESTRUCTURA CONSTRUIDA;
    ' 7.Mat, per al seu contingut, amb la mateixa forma que 6.Bio.
    PC9 f, "pgEst", "Saquejada:",             "Looting",            2, 1
    PC9 f, "pgEst", "Dany per foc:",          "Fire_Damage",        3, 1
    PC9 f, "pgEst", "Activitat animal:",      "Animal_Activity",    2, 2
    PC9 f, "pgEst", "Acces modern:",          "Modern_Access",      3, 2
    SH  f, "pgEst", "Base documental i observabilitat (OE1)", 6
    ' rev. 5: escala ordinal de qualitat documental. La frontera
    ' entre els dos primers valors es FISICA i no metrica: amb
    ' cordes s'esta evidentment a menys de 5 m, aixi que el que
    ' distingeix l'acces directe es el CONTACTE. La tecnica de
    ' captura ja queda descrita per les URL de 10.Doc.
    PCV f, "pgEst", "Base documental:",  "Doc_Basis",              7, 1, "Direct access;Acces directe (contacte);Close-range (<5m);Proximitat (<5m);Medium-range (5-30m);Distancia mitjana (5-30m);Long-range (>30m);Llarga distancia (>30m);ND;Indeterminada"
    PCV f, "pgEst", "Observ. facana:",   "Facade_Observability",   8, 1, "Complete;Completa;Partial;Parcial;Poor;Deficient;ND;Indeterminada"
    PCV f, "pgEst", "Observ. interior:", "Interior_Observability", 8, 2, "Complete;Completa;Partial;Parcial;None;Nul-la;ND;Indeterminada"
    SH  f, "pgEst", "Observacions sobre conservacio", 10
    PCN f, "pgEst", "Notes:", "Condition_Notes", 11
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
    SH  f, "pgBio", "Observacions bioarqueologiques", 7
    PCN f, "pgBio", "Notes:", "Bio_Notes", 8
End Sub

' TAB 7 - CULTURAL MATERIALS. rev. 5: Cultural_Materials_Present es
' ara la porta, analeg exacte de Human_Remains a 6.Bio.
' ID_Material_Status barrejava dues variables en una escala
' (Good/Fair/Poor son graus de conservacio, Absent es una afirmacio
' de presencia) - el mateix error que Lintel tenia en v10.
Private Sub FillMat(f As String)
    SH f, "pgMat", "Materials culturals", 0
    PC9 f, "pgMat", "Materials culturals:", "Cultural_Materials_Present", 1, 1
    ' v17 (delta D2-D3): l'estat arriba de 5.Estat i queda davall
    ' de la seua porta. Les tres capes NO son redundants: la
    ' presencia es contestable des de trenta metres, el tipus
    ' exigeix identificar-los i l'estat exigeix veure'ls prou be
    ' per a jutjar-ho. Una estructura observada de lluny on es veu
    ' que hi ha material sense distingir-ne cap categoria val 1
    ' amb els set tipus a 9 - i si la presencia es deduira dels
    ' set, eixe registre diria 'no hi ha materials', que es fals.
    PCC f, "pgMat", "Estat vestigis mobles:", "ID_Material_Status", 1, 2
    SH f, "pgMat", "Detall per tipus de material", 2
    PC9 f, "pgMat", "Textils:",          "Mat_Textiles",   3, 1
    PC9 f, "pgMat", "Fusta cultural:",   "Mat_Wood",       4, 1
    PC9 f, "pgMat", "Fibra vegetal:",    "Mat_VegFiber",   5, 1
    PC9 f, "pgMat", "Ceramica:",         "Mat_Ceramics",   6, 1
    PC9 f, "pgMat", "Fauna:",            "Mat_Fauna",      3, 2
    PC9 f, "pgMat", "Banya de cervol:",  "Mat_DeerAntler", 4, 2
    PC9 f, "pgMat", "Altres materials:", "Mat_Other",      5, 2
    SH  f, "pgMat", "Observacions sobre materials culturals", 8
    PCN f, "pgMat", "Notes:", "Materials_Notes", 9
End Sub

' TAB 8 - CHRONOLOGY
Private Sub FillCron(f As String)
    SH  f, "pgCron", "Cronologia", 0
    PCB f, "pgCron", "Datacio C14:",      "C14",               1, 1
    PCT f, "pgCron", "Inici (segle dC):", "Chrono_Start_Cent", 2, 1
    PCT f, "pgCron", "Fi (segle dC):",    "Chrono_End_Cent",   3, 1
    PCC f, "pgCron", "Campanya:",         "ID_Campaign",       4, 1
    SH  f, "pgCron", "Datacions radiocarboniques (T_DATING)", 6
End Sub

' TAB 9 - METRICS. Opening_Width_m / Opening_Height_m are NOT gated
' by Sys_Portal: the metric pass is a separate exercise (4.5).
' rev. 5: TOTES les dimensions lineals en METRES. No es cosmetic -
' una consulta pot sumar o comparar camps sense conversions, i
' dividir per 100 a ma es on s'introdueixen errors silenciosos.
' Els controls es formaten amb 2 decimals (PCD) perque l'usuari
' veja el que ha escrit: 0,62 i no 0,6. La regla 31 marca qualsevol
' dimensio superior a 20 m, que es l'error d'entrar 62 pensant en
' centimetres.
Private Sub FillMetr(f As String)
    SH  f, "pgMetr", "Dimensions de l'estructura", 0
    PCD f, "pgMetr", "Longitud (m):", "Length_m",   1, 1
    PCD f, "pgMetr", "Amplada (m):",  "Width_m",    1, 2
    PCD f, "pgMetr", "Alcada (m):",   "Height_m",   2, 1
    PCV f, "pgMetr", "Metode dim.:",  "Dim_Method", 2, 2, "Photogrammetric model;Model fotogrametric;Tape measure;Cinta metrica;Laser;Laser;Estimation;Estimacio;perimeter pigment outline;Perimetre de pigment;ND;Indeterminat"
    SH  f, "pgMetr", "Dimensions de l'obertura d'acces (P)", 3
    PCD f, "pgMetr", "Amplada obertura (m):", "Opening_Width_m",  4, 1
    PCD f, "pgMetr", "Alcada obertura (m):",  "Opening_Height_m", 4, 2
    SH  f, "pgMetr", "Metrica del suport geologic (H02)", 5
    PCD f, "pgMetr", "Amplada suport (m):", "Support_Width_m", 6, 1
    PCD f, "pgMetr", "Profunditat (m):",    "Support_Depth_m", 6, 2
    ' rev. 7: UN volum i UNA superficie. Interior_Vol i Total_Vol no
    ' volien dir el mateix segons la tipologia - en un mausoleu la
    ' diferencia ES la fabrica, pero en una cambra dins d'una
    ' cavitat l'"interior" es espai natural que ningu no va excavar.
    ' El desglossament fi (per cos, superficie de plataforma, volum
    ' construit) va a Vol_Notes: son casos puntuals, i un camp
    ' quedaria buit en la majoria de registres.
    SH  f, "pgMetr", "Volumetria i superficie", 7
    PCD f, "pgMetr", "Superficie (m2):", "Area_m2",       8, 1
    PCC f, "pgMetr", "Metode calc.:",    "ID_Vol_Method", 8, 2
    PCD f, "pgMetr", "Volum (m3):",      "Volume_m3",     9, 1
    PCT f, "pgMetr", "Notes vol.:",      "Vol_Notes",     9, 2
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
'  v17: I (Interbody_Cornice) is now governed by Sys_Interface.
'  R (Upper_Crown) belongs to no system and stays permanently
'  active, deliberately: a one-bodied structure has no interbody
'  cornice, but it still has a top.
' ================================================================
Private Sub FillSys(f As String)
    SH  f, "pgSys", "Sistemes constructius - resolre primer", 0
    PCS f, "pgSys", "Sistema plataforma (H):", "Sys_Platform", 1, 1
    PCS f, "pgSys", "Sistema portal (P):",     "Sys_Portal",   1, 2
    PCS f, "pgSys", "Sistema rafec (U):",      "Sys_Eave",     2, 1
    PCG f, "pgSys", "Conjunt basal (A-D):",    "Sys_Base",     2, 2
    PCG f, "pgSys", "Conjunt cambra:",         "Sys_Chamber",  3, 1
    ' v17 (delta E3): SISE SISTEMA. Passa el test que va
    ' rebutjar un sistema de facana en v16 - una facana no pot
    ' ser absent mai, i un sistema que no pot ser absent no fa
    ' el que fan els sistemes -, perque una interficie SI que
    ' pot ser absent: una estructura d'un sol cos no en te.
    ' Governa NOMES la cornisa intercos (I). El coronament (R)
    ' es queda permanentment actiu: amb un sol cos no hi ha
    ' cornisa entre cossos, pero l'estructura te part de dalt
    ' igualment, i tancar-los junts tancaria una pregunta que
    ' si que te resposta.
    PCG f, "pgSys", "Sistema interficie:",     "Sys_Interface", 3, 2

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

    SH  f, "pgSys", "Interficie N0/N1 (I sota Sys_Interface; R sempre actiu)", 12
    ' v16 (delta 5.1): Mat. cornisa ELIMINAT. Una cornisa es
    ' SEMPRE de pedra, no per costum sino per definicio: un
    ' element horitzontal de fusta volat entre dos cossos es una
    ' mensula (E) o una biga transversal (F), i el que hi ha
    ' aleshores es plataforma, no cornisa. El valor 'Bigues de
    ' fusta' descrivia una plataforma mal classificada. T i U ja
    ' porten SEMPRE PEDRA a la definicio i no tenen camp de
    ' material propi: la cornisa era l'excepcio incoherent.
    ' La clausula val ara mes com a TEST D'IDENTIFICACIO de G
    ' contra I (regla 39).
    PC5 f, "pgSys", "Cornisa intercos (I):", "Interbody_Cornice",          13, 1
    PC5 f, "pgSys", "Coronament (R):",       "Upper_Crown",                14, 1

    ' v16 (delta 5.4, 5.5): CRITERI DE DIFERENCIACIO. Un element
    ' A-X es present quan hi ha un component FISICAMENT
    ' DIFERENCIAT del parament contigu - per format de pedra,
    ' dimensio, material o tractament. Una obertura resolta
    ' deixant simplement un buit al mur es N=0, O=0, Q=0 - i
    ' Sys_Portal present igualment, perque l'obertura hi es.
    ' No es cap contradiccio: el SISTEMA registra que existeix
    ' obertura, els COMPONENTS si cada posicio es va resoldre
    ' amb un element distint. Eixos zeros son un resultat de
    ' primera magnitud sobre la inversio de treball (H01, H04):
    ' amb el criteri contrari tots els portals serien iguals.
    ' Portal asimetric (un brancal diferenciat i l'altre no):
    ' valor 2 i nota amb formula constant a Notes sistemes,
    ' 'asymmetric: left only'. Menys de tres casos al corpus, de
    ' manera que no es crea camp; si el patro es repeteix, eixe
    ' sera el senyal que en cal un.
    SH  f, "pgSys", "Sistema portal: N + O + Q", 15
    PC5 f, "pgSys", "Llindar (N):",     "Sill",            16, 1
    PC5 f, "pgSys", "Brancals (O):",    "Jambs",           16, 2
    PC5 f, "pgSys", "Dintell (Q):",     "Lintel",          17, 1
    PCV f, "pgSys", "Mat. dintell:",    "Lintel_Material", 17, 2, "Stone;Pedra;Wood;Fusta;Mixed;Mixt;ND;Tipus indeterminat"

    SH  f, "pgSys", "Conjunt cambra: J K L M V X", 19
    PC5 f, "pgSys", "Cantoneres (J):",         "Corner_Quoins",        20, 1
    PC5 f, "pgSys", "Pilastres estruct. (K):", "Structural_Pilasters", 20, 2
    PC5 f, "pgSys", "Flanc de facana (L):",    "Facade_Flank",         21, 1
    PC5 f, "pgSys", "Fris en relleu (M):",     "Relief_Frieze",        21, 2
    PC5 f, "pgSys", "Mur lateral de cambra (V):", "Return_Wall",       22, 1
    ' v16 (delta 5.2): X registra el TANCAMENT EFECTIU de
    ' l'espai funerari, per obra o per roca EN CONTACTE amb la
    ' fabrica. Test binari, sense gradient: SENSE CONTACTE, NO
    ' HI HA SOSTRE - 20 cm de buit es buit igual que 3 m. Un
    ' mausoleu amb la visera separada es X = 0 i la visera va a
    ' ID_Support i a Notes, perque la roca no tanca res. A
    ' QRY_13, X exporta NA quan el tipus es penya natural: no es
    ' una absencia sino una solucio NO CONSTRUIDA, i comptar-la
    ' com a element feia que la matriu A-X comptara geologia com
    ' si fora construccio.
    PC5 f, "pgSys", "Coberta cambra (X):",     "Chamber_Roof",         22, 2
    PCV f, "pgSys", "Tipus coberta (X):",      "Chamber_Roof_Type",    23, 1, "Natural bedrock;Penya natural;Built masonry;Obra de maconeria;Built timber and slabs;Fusta i lloses;Mixed;Mixt;ND;Tipus indeterminat"
    ' v16 (delta 2.4): 'Tancament posterior' passa a 'Fons de
    ' cambra'. En arqueologia funeraria 'tancament' s'enten per
    ' defecte com el SEGELLAT de la tomba - la llosa que obtura
    ' l'obertura -, i el risc era registrar-hi la cosa
    ' equivocada sense adonar-se'n. Al corpus no s'ha documentat
    ' cap llosa ni muret de tancament d'obertura, i eixa
    ' absencia es un resultat sobre l'acces recurrent a la
    ' cambra, no un silenci.
    PCV f, "pgSys", "Fons de cambra (W):",     "Rear_Closure_Type",    23, 2, "Natural bedrock;Penya natural;Built masonry;Obra de maconeria;Mixed;Mixt;ND;Tipus indeterminat"

    SH  f, "pgSys", "Sistema rafec: S + T", 24
    PC5 f, "pgSys", "Biga suport rafec (S):", "Eave_Beam",    25, 1
    PC5 f, "pgSys", "Superficie rafec (T):",  "Eave_Surface", 25, 2

    SH  f, "pgSys", "Observacions sobre sistemes i elements", 26
    PCN f, "pgSys", "Notes:", "Systems_Notes", 27

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
    SH f, "pgExtra", "Connexions registrades des d'aquesta estructura", 7
    ' v17 (delta D1): la llista de connexions ENTRANTS. Sense
    ' ella, des de B no es veia res del que s'havia registrat des
    ' d'A i la duplicacio era el comportament per defecte.
    SH f, "pgExtra", "Connexions registrades des d'altres estructures (nomes lectura)", 14
    SH f, "pgExtra", "Elements desapareguts - evidencia (T_LOST_ELEMENTS, regla 2)", 20
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

    ' rev. 7: els combos de taula mostren Name_VAL i no Name. El
    ' formulari tenia dues menes de combo i nomes una estava
    ' localitzada: els de llista de valors (PC5/PC9/PCV) porten les
    ' dues columnes a la cadena del codi i es veien en valencia,
    ' mentre que els de taula llegien Name i es veien en angles.
    ' Mig formulari en cada idioma.
    ' El patro es el mateix que ja usaven les llistes de valors -
    ' columna emmagatzemada oculta, etiqueta visible - pero aci
    ' l'emmagatzemat es un ID numeric, cosa que ho fa encara mes
    ' segur: reanomenar una etiqueta no pot tocar cap valor.
    ' Name continua sent el terme de referencia: exportacions,
    ' publicacio i les regles de QRY_16 que hi busquen per nom.
    cs(0) = "ID_Sector":          rs(0) = "SELECT ID, Sector_Name FROM L_SECTORS ORDER BY ID_Site, Sector_Name": cc(0) = 2: cw(0) = "0cm;5cm"
    ' Typology carries Record_Class, which drives the tab gating
    cs(1) = "ID_Typology":        rs(1) = "SELECT ID, Name_VAL FROM L_TYPOLOGY ORDER BY Name":                       cc(1) = 2: cw(1) = "0cm;6cm"
    cs(2) = "ID_Support":         rs(2) = "SELECT ID, Name_VAL FROM L_SUPPORT ORDER BY ID":                          cc(2) = 2: cw(2) = "0cm;5cm"
    cs(3) = "ID_Arch_Status":     rs(3) = "SELECT ID, Name_VAL FROM L_STATUS ORDER BY ID":                           cc(3) = 2: cw(3) = "0cm;4cm"
    cs(4) = "ID_Material_Status": rs(4) = "SELECT ID, Name_VAL FROM L_MATERIAL_STATUS ORDER BY ID":                  cc(4) = 2: cw(4) = "0cm;5cm"
    cs(5) = "ID_Vol_Method":      rs(5) = "SELECT ID, Name_VAL FROM L_VOL_METHOD ORDER BY ID":                       cc(5) = 2: cw(5) = "0cm;5cm"
    cs(6) = "ID_Coord_Method":    rs(6) = "SELECT ID, Name_VAL FROM L_COORD_METHOD ORDER BY ID":                     cc(6) = 2: cw(6) = "0cm;5cm"
    cs(7) = "ID_Campaign":        rs(7) = "SELECT ID, Code, Campaign_Name FROM L_CAMPAIGN ORDER BY Code":        cc(7) = 3: cw(7) = "0cm;1.5cm;5cm"
    cs(8) = "ID_Group":           rs(8) = "SELECT ID, Group_Code FROM T_GROUPS ORDER BY Group_Code":             cc(8) = 2: cw(8) = "0cm;4cm"
    cs(9) = "ID_Parent":          rs(9) = "SELECT ID, Code FROM T_STRUCTURES ORDER BY Code":                     cc(9) = 2: cw(9) = "0cm;4cm"
    cs(10) = "ID_Struct_Body":    rs(10) = "SELECT ID, Name_VAL FROM L_STRUCT_BODY ORDER BY Sort_Order":       cc(10) = 2: cw(10) = "0cm;5cm"
    cs(11) = "ID_Dec_Type":       rs(11) = "SELECT ID, Name_VAL FROM L_DEC_TYPE ORDER BY Name":                      cc(11) = 2: cw(11) = "0cm;5cm"
    cs(12) = "ID_Support_Secondary": rs(12) = "SELECT ID, Name_VAL FROM L_SUPPORT ORDER BY ID":                      cc(12) = 2: cw(12) = "0cm;5cm"

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
    SetAfterUpdate f, "Cultural_Materials_Present"
    SetAfterUpdate f, "Human_Remains"
    SetAfterUpdate f, "C14"
    SetAfterUpdate f, "Sys_Base"
    SetAfterUpdate f, "Sys_Platform"
    SetAfterUpdate f, "Sys_Portal"
    SetAfterUpdate f, "Sys_Eave"
    SetAfterUpdate f, "Sys_Chamber"
    SetAfterUpdate f, "Sys_Interface"
    SetAfterUpdate f, "N_Basal_Bodies"
    SetAfterUpdate f, "N_Chamber_Bodies"
    SetAfterUpdate f, "Plaster_Present"
    SetAfterUpdate f, "Pigment_Present"
    SetAfterUpdate f, "Mortar_Present"
    SetAfterUpdate f, "Dec_Present"
    SetAfterUpdate f, "RockArt_Present"
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
    LG "    ' v17: NORMALITZACIO DE LA PARELLA. L'ID menor sempre"
    LG "    ' a A. En girar-los cal GIRAR TAMBE LA CRONOLOGIA, o el"
    LG "    ' sentit s'inverteix en silenci. Amb l'ordre normalitzat,"
    LG "    ' l'index unic UQ_CONN_PAIR fa impossible la duplicacio."
    LG "    If Not IsNull(Me!ID_Struct_A) And Not IsNull(Me!ID_Struct_B) Then"
    LG "        If Me!ID_Struct_A = Me!ID_Struct_B Then"
    LG "            MsgBox ""Una connexio necessita dues estructures distintes."", vbExclamation"
    LG "            Cancel = True"
    LG "            Exit Sub"
    LG "        End If"
    LG "        If Me!ID_Struct_A > Me!ID_Struct_B Then"
    LG "            Dim sw As Long"
    LG "            sw = Me!ID_Struct_A"
    LG "            Me!ID_Struct_A = Me!ID_Struct_B"
    LG "            Me!ID_Struct_B = sw"
    LG "            If Nz(Me!Chrono_Relation, """") = ""A is earlier"" Then"
    LG "                Me!Chrono_Relation = ""B is earlier"""
    LG "            ElseIf Nz(Me!Chrono_Relation, """") = ""B is earlier"" Then"
    LG "                Me!Chrono_Relation = ""A is earlier"""
    LG "            End If"
    LG "        End If"
    LG "    End If"
    LG "    Dim cr As String"
    LG "    cr = Nz(Me!Chrono_Relation, """")"
    LG "    If cr = ""A is earlier"" Or cr = ""B is earlier"" Then"
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
    LG "Private Sub Cultural_Materials_Present_AfterUpdate()"
    LG "    QuickFillMaterials"
    LG "    ApplyGating"
    LG "End Sub"
    LG ""
    LG "Private Sub C14_AfterUpdate()"
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
    LG "Private Sub Sys_Interface_AfterUpdate()"
    LG "    QuickFillSys ""Sys_Interface"""
    LG "    ApplyGating"
    LG "End Sub"
    LG ""
    LG "Private Sub N_Basal_Bodies_AfterUpdate()"
    LG "    ApplyGating"
    LG "End Sub"
    LG ""
    LG "Private Sub N_Chamber_Bodies_AfterUpdate()"
    LG "    ApplyGating"
    LG "End Sub"
    LG ""
    LG "Private Sub Plaster_Present_AfterUpdate()"
    LG "    If Me!Plaster_Present = 0 Or Me!Plaster_Present = 9 Then FillGroup """", 0, ""Plaster_Color,Plaster_Extent"""
    LG "    ApplyGating"
    LG "End Sub"
    LG ""
    LG "Private Sub Pigment_Present_AfterUpdate()"
    LG "    If Me!Pigment_Present = 0 Or Me!Pigment_Present = 9 Then FillGroup """", 0, ""Pigment_Substrate,Pigment_Color,Pigment_Extent"""
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
    LG "Private Sub RockArt_Present_AfterUpdate()"
    LG "    ApplyGating"
    LG "End Sub"
    LG ""
    LG "Private Sub Dec_Present_AfterUpdate()"
    LG "    If Me!Dec_Present = 0 Or Me!Dec_Present = 9 Then"
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
    LG "            ' E exclos a proposit (delta 1): tancar el sistema"
    LG "            ' no autoritza a negar una mensula observada."
    LG "            comps = ""Transverse_Beams,Corbelled_Courses"""
    LG "            clears = ""Timber_Bracket_Count,Timber_Bracket_Role,Platform_Surface_Material,Platform_Function"""
    LG "        Case ""Sys_Portal"""
    LG "            comps = ""Sill,Jambs,Lintel"""
    LG "            clears = ""Lintel_Material"""
    LG "        Case ""Sys_Eave"""
    LG "            comps = ""Eave_Beam,Eave_Surface"""
    LG "        Case ""Sys_Chamber"""
    LG "            comps = ""Corner_Quoins,Structural_Pilasters,Facade_Flank,Relief_Frieze,Return_Wall,Chamber_Roof"""
    LG "            clears = ""Chamber_Roof_Type,Rear_Closure_Type"""
    LG "        Case ""Sys_Interface"""
    LG "            ' v17: nomes I. R no penja de cap sistema."
    LG "            comps = ""Interbody_Cornice"""
    LG "    End Select"
    LG "    FillGroup comps, target, clears"
    LG "End Sub"
    LG ""
    LG "Private Sub QuickFillMaterials()"
    LG "    Dim v As Variant"
    LG "    v = Me!Cultural_Materials_Present"
    LG "    If IsNull(v) Then Exit Sub"
    LG "    Dim g As String"
    LG "    g = ""Mat_Textiles,Mat_Wood,Mat_VegFiber,Mat_Ceramics,Mat_Fauna,Mat_DeerAntler,Mat_Other"""
    LG "    If v = 0 Then"
    LG "        FillGroup g, 0, ""ID_Material_Status"""
    LG "    ElseIf v = 9 Then"
    LG "        FillGroup g, 9, ""ID_Material_Status"""
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
    LG "    ' E NO es gateja pel sistema (delta 1). Es l'unic element"
    LG "    ' del vocabulari amb existencia independent del seu"
    LG "    ' sistema: una mensula aillada no ha d'haver portat mai"
    LG "    ' cap plataforma. La regla 1 de la bateria ja l'exempta."
    LG "    EnSrc ""Timber_Brackets"", True"
    LG "    EnSrc ""Transverse_Beams"", b"
    LG "    EnSrc ""Corbelled_Courses"", b"
    LG "    ' Nivell 3: el detall de les mensules penja d'E, no del"
    LG "    ' sistema (R8, R9)"
    LG "    Dim eb As Boolean"
    LG "    eb = ElemHas(""Timber_Brackets"")"
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
    LG "    ' v17 (delta E3): la cornisa intercos passa a penjar de"
    LG "    ' Sys_Interface. El coronament (R) NO: amb un sol cos no"
    LG "    ' hi ha cornisa entre cossos, pero l'estructura te part"
    LG "    ' de dalt igualment. El material de la cornisa va"
    LG "    ' desapareixer en v16: una cornisa es sempre de pedra."
    LG "    EnSrc ""Interbody_Cornice"", SysOpen(Me!Sys_Interface)"
    LG ""
    LG "    ' v17 (delta E1): pla d'acces i orientacio del portal"
    LG "    ' NOMES amb cambra. Gatejats per Sys_Chamber i no per"
    LG "    ' Sys_Portal, perque una obertura arrasada dins d'una"
    LG "    ' cambra real conserva orientacio coneguda (decisio v16)."
    LG "    ' NOMES amb 'No aplicable', mai amb 'Absent': Absent es"
    LG "    ' el valor per defecte dels sistemes, de manera que no"
    LG "    ' pot valdre com a declaracio de l'usuari - bloquejaria"
    LG "    ' aquests camps en tots els registres nous abans que"
    LG "    ' ningu haja dit res. 'No aplicable' sempre es deliberat."
    LG "    ' Orientacio de facana i visibilitat de la vall queden"
    LG "    ' obertes: la facana es el pla EXPOSAT, i una massa basal"
    LG "    ' en te i mira cap a algun lloc."
    LG "    Dim chDecl As Boolean"
    LG "    chDecl = (Nz(Me!Sys_Chamber, """") = ""Not applicable"")"
    LG "    EnSrc ""Access_Plane"", Not chDecl"
    LG "    EnSrc ""Portal_Orientation"", Not chDecl"
    LG ""
    LG "    ' v17: la posicio del portal a la facana penja del"
    LG "    ' sistema portal, com la resta del seu grup."
    LG "    EnSrc ""Portal_Position"", Nz(Me!Sys_Portal, """") <> ""Not applicable"""
    LG ""
    LG "    ' v17: Fabric NO es gateja. Es va intentar tancar-lo"
    LG "    ' amb menys de dos cossos i era un error de fons: una"
    LG "    ' estructura d'UN SOL COS pot tindre dues fabriques"
    LG "    ' -el valor 'dins d'un cos'-, de manera que el bloqueig"
    LG "    ' tancava justament el cas que el camp havia de recollir."
    LG "    ' L'unica combinacio impossible es 'per cossos' amb un"
    LG "    ' sol cos, i eixa la vigila la regla 45."
    LG ""
    LG "    ' v17 (punt 8): 9.Metr deixa de suposar mausoleu. Els"
    LG "    ' dos blocs que no valen per a una cova, una terrassa o"
    LG "    ' un panell es tanquen pels SISTEMES, que es on ja viu la"
    LG "    ' informacio, i tambe nomes amb 'No aplicable'."
    LG "    Dim poNA As Boolean, chNA As Boolean"
    LG "    poNA = (Nz(Me!Sys_Portal, """") = ""Not applicable"")"
    LG "    chNA = (Nz(Me!Sys_Chamber, """") = ""Not applicable"")"
    LG "    EnSrc ""Opening_Width_m"", Not poNA"
    LG "    EnSrc ""Opening_Height_m"", Not poNA"
    LG "    EnSrc ""Area_m2"", Not chNA"
    LG "    EnSrc ""Volume_m3"", Not chNA"
    LG "    EnSrc ""ID_Vol_Method"", Not chNA"
    LG "    EnSrc ""Vol_Notes"", Not chNA"
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
    LG "    If IsNull(Me!Plaster_Present) Or Me!Plaster_Present <> 0 Then"
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
    LG "    ' rev. 5: la porta de 7.Mat es un camp propi, no el nom"
    LG "    ' d'un valor del lookup d'estat: mes directe, sense"
    LG "    ' DLookup, i tanca tambe l'estat de conservacio."
    LG "    b = IsNull(Me!Cultural_Materials_Present) Or Nz(Me!Cultural_Materials_Present, 1) = 1"
    LG "    EnSrc ""ID_Material_Status"", b"
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
    LG "    ' Datacions: C14 governa el subformulari de T_DATING."
    LG "    ' Els segles queden SEMPRE editables: la font habitual de"
    LG "    ' l'atribucio cronologica no es radiocarbonica sino"
    LG "    ' tipologica, i bloquejar-los impediria registrar-la en la"
    LG "    ' immensa majoria del corpus."
    LG "    Dim ct4 As Control"
    LG "    For Each ct4 In Me.Controls"
    LG "        If ct4.ControlType = acSubform Then"
    LG "            If ct4.SourceObject = ""F_DATING"" Then ct4.Enabled = Nz(Me!C14, False)"
    LG "        End If"
    LG "    Next ct4"
    LG ""
    LG "    ' Decoracio i art rupestre: cada judici agregat governa el"
    LG "    ' seu subformulari (delta 7.4, 7.5). Els valors 1 i 2 obrin;"
    LG "    ' 3 tambe, perque una decoracio desapareguda es registra"
    LG "    ' igualment amb la seua evidencia."
    LG "    Dim dOn As Boolean, rOn As Boolean"
    LG "    dOn = IsNull(Me!Dec_Present) Or Nz(Me!Dec_Present, 1) = 1"
    LG "    rOn = IsNull(Me!RockArt_Present) Or Nz(Me!RockArt_Present, 1) = 1"
    LG "    Dim ct3 As Control"
    LG "    For Each ct3 In Me.Controls"
    LG "        If ct3.ControlType = acSubform Then"
    LG "            If ct3.SourceObject = ""F_DECORATIONS"" Then ct3.Enabled = dOn"
    LG "            If ct3.SourceObject = ""F_ROCKART"" Then ct3.Enabled = rOn"
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
