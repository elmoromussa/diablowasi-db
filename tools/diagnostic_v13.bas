Option Compare Database
Option Explicit

' ================================================================
'  CHACHAPOYA - DIAGNOSTIC PREVI A LA v13
'  La Petaca & Diablo Wasi (Leymebamba, Amazonas, Peru)
'  Author: Esteve Ribera Torro | TFM Arqueologia UA
'  Spec: DELTA_v12_v13.md, seccio 11 (comprovacions pendents)
'
'  QUE FA: NOMES LLIG. Cap INSERT, cap UPDATE, cap DELETE, cap
'  canvi d'esquema. Es pot executar sobre la base de treball sense
'  copia previa i tantes vegades com calga.
'
'  QUE RESPON: les nou comprovacions que el delta deixa obertes i
'  que decideixen com s'implementa la v13.
'
'    D1  Registres amb ID_Support = Fissure/Crack  -> punt 2
'    D2  Diaclasi com a primaria vs secundaria     -> punt 2
'    D3  Parelles de formes de suport               -> punt 2
'    D4  N_Basal_Bodies > 0 amb Base_Level = 0      -> punt 4 (R27)
'    D5  Candidats a fitxa PR i art rupestre        -> punt 7.2 / 7.3
'    D6  Distribucio de Color a T_DECORATIONS       -> punt 7.6
'    D7  Distribucio de Chinking_Stones             -> punt 8.6
'    D8  Estat dels camps que canvien de defecte    -> punt 8.2
'    D9  Capes aplicades: candidats al valor 3      -> punt 8.3
'
'  COM S'EXECUTA: Alt+F11, modul nou, apegar, cursor dins de
'  RunDiagnostic i F5. El resultat ix per la finestra immediata
'  (Ctrl+G). Per a conservar-lo, copiar-lo i apegar-lo al delta.
' ================================================================

Sub RunDiagnostic()
    Dim db As DAO.Database
    Set db = CurrentDb()

    Debug.Print String(64, "=")
    Debug.Print "DIAGNOSTIC PREVI A LA v13  (" & Now & ")"
    Debug.Print "Registres a T_STRUCTURES: " & DCount("*", "T_STRUCTURES")
    Debug.Print String(64, "=")

    D1_SupportFissure db
    D2_CleftPrimaryVsSecondary db
    D3_SupportPairs db
    D4_BasalBodiesWithoutB db
    D5_RockArtCandidates db
    D6_ColorDistribution db
    D7_ChinkingVariance db
    D8_DefaultChangeFields db
    D9_LayerLossCandidates db

    Debug.Print String(64, "=")
    Debug.Print "DIAGNOSTIC COMPLET. Cap dada modificada."
    Debug.Print String(64, "=")

    Set db = Nothing
End Sub

' ================================================================
'  HELPERS
' ================================================================
Private Sub H(t As String)
    Debug.Print ""
    Debug.Print "-- " & t & " " & String(58 - Len(t), "-")
End Sub

Private Function FldExists(db As DAO.Database, tbl As String, fld As String) As Boolean
    Dim f As DAO.Field
    On Error Resume Next
    For Each f In db.TableDefs(tbl).Fields
        If f.Name = fld Then FldExists = True: Exit Function
    Next f
    On Error GoTo 0
End Function

' Imprimeix una consulta de dues columnes: etiqueta i recompte.
Private Sub PrintPairs(db As DAO.Database, sql As String, lbl As String, cnt As String)
    Dim rs As DAO.Recordset
    On Error GoTo Err_P
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    If rs.EOF Then
        Debug.Print "   (cap resultat)"
    Else
        Do While Not rs.EOF
            Debug.Print "   " & Nz(rs(lbl), "(buit)") & " : " & rs(cnt)
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
    Exit Sub
Err_P:
    Debug.Print "   ERROR: " & Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
End Sub

' Imprimeix la llista de codis d'una condicio sobre T_STRUCTURES.
Private Sub PrintCodes(db As DAO.Database, whereClause As String)
    Dim rs As DAO.Recordset
    Dim s As String
    Dim n As Long
    On Error GoTo Err_C
    Set rs = db.OpenRecordset("SELECT Code FROM T_STRUCTURES WHERE " & whereClause & " ORDER BY Code", dbOpenSnapshot)
    Do While Not rs.EOF
        s = s & rs!Code & "  "
        n = n + 1
        If n Mod 5 = 0 Then
            Debug.Print "   " & s
            s = ""
        End If
        rs.MoveNext
    Loop
    If Len(s) > 0 Then Debug.Print "   " & s
    If n = 0 Then Debug.Print "   (cap)"
    rs.Close
    Set rs = Nothing
    Exit Sub
Err_C:
    Debug.Print "   ERROR: " & Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
End Sub

' Recompte per valor d'un camp BYTE, incloent-hi els NULL.
Private Sub PrintByteDist(db As DAO.Database, fld As String)
    If Not FldExists(db, "T_STRUCTURES", fld) Then
        Debug.Print "   " & fld & ": CAMP INEXISTENT"
        Exit Sub
    End If
    Dim n0 As Long, n1 As Long, n2 As Long, n3 As Long, n9 As Long, nN As Long
    n0 = DCount("*", "T_STRUCTURES", "[" & fld & "]=0")
    n1 = DCount("*", "T_STRUCTURES", "[" & fld & "]=1")
    n2 = DCount("*", "T_STRUCTURES", "[" & fld & "]=2")
    n3 = DCount("*", "T_STRUCTURES", "[" & fld & "]=3")
    n9 = DCount("*", "T_STRUCTURES", "[" & fld & "]=9")
    nN = DCount("*", "T_STRUCTURES", "[" & fld & "] Is Null")
    Dim ln As String
    ln = "   " & Left(fld & String(28, " "), 28)
    ln = ln & " 0:" & n0 & "  1:" & n1 & "  2:" & n2 & "  3:" & n3
    ln = ln & "  9:" & n9 & "  NULL:" & nN
    Debug.Print ln
End Sub

' ================================================================
'  D1 - Registres amb ID_Support = Fissure/Crack   (delta punt 2)
'  Decideix si el desdoblament es un simple renomenament o exigeix
'  reassignacio registre a registre.
' ================================================================
Private Sub D1_SupportFissure(db As DAO.Database)
    H "D1. ID_Support = Fissure/Crack (punt 2)"

    Dim idF As Variant
    idF = DLookup("ID", "L_SUPPORT", "Name Like 'Fissure*'")
    If IsNull(idF) Then
        Debug.Print "   El valor Fissure/Crack no existeix a L_SUPPORT."
        Debug.Print "   -> Ja migrat, o la taula no es la esperada."
        Exit Sub
    End If

    Dim nP As Long, nS As Long
    nP = DCount("*", "T_STRUCTURES", "ID_Support=" & idF)
    nS = DCount("*", "T_STRUCTURES", "ID_Support_Secondary=" & idF)
    Debug.Print "   Com a suport primari:   " & nP
    Debug.Print "   Com a suport secundari: " & nS
    Debug.Print "   TOTAL a reclassificar:  " & (nP + nS)
    Debug.Print ""
    Debug.Print "   Codis (primari):"
    PrintCodes db, "ID_Support=" & idF
    Debug.Print "   Codis (secundari):"
    PrintCodes db, "ID_Support_Secondary=" & idF
    Debug.Print ""
    Debug.Print "   ACCIO: revisar cada codi al model 3D i anotar si es"
    Debug.Print "   VERTICAL (diaclasi) o HORITZONTAL (junta d'estratificacio)."
    Debug.Print "   Si tots son d'una mena -> renomenar la fila i afegir l'altra."
    Debug.Print "   Si n'hi ha de les dues -> cal reassignar-los un a un."
End Sub

' ================================================================
'  D2 - Frequencia de cada forma com a primaria i com a secundaria
'  (delta punt 2). La lectura interessant: quines formes no s'usen
'  mai com a suport principal.
' ================================================================
Private Sub D2_CleftPrimaryVsSecondary(db As DAO.Database)
    H "D2. Cada forma de suport: primaria vs secundaria (punt 2)"

    Dim rs As DAO.Recordset
    On Error GoTo Err_D2
    Set rs = db.OpenRecordset("SELECT ID, Name FROM L_SUPPORT ORDER BY ID", dbOpenSnapshot)
    Debug.Print "   " & Left("FORMA" & String(32, " "), 32) & " PRIM  SEC"
    Do While Not rs.EOF
        Dim nP As Long, nS As Long
        nP = DCount("*", "T_STRUCTURES", "ID_Support=" & rs!ID)
        nS = DCount("*", "T_STRUCTURES", "ID_Support_Secondary=" & rs!ID)
        Dim ln2 As String
        ln2 = "   " & Left(rs!Name & String(32, " "), 32) & " "
        ln2 = ln2 & Right("   " & nP, 4) & "  " & Right("   " & nS, 4)
        Debug.Print ln2
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    Debug.Print ""
    Debug.Print "   LECTURA: una forma amb PRIM=0 i SEC>0 no s'usa mai com a"
    Debug.Print "   suport principal, nomes com a estabilitzador. Es un resultat."
    Exit Sub
Err_D2:
    Debug.Print "   ERROR: " & Err.Description
End Sub

' ================================================================
'  D3 - Parelles de formes (delta punt 2)
'  Les combinacions poden ser, elles mateixes, tipus de solucio
'  constructiva: la classificacio que ix de H02.
' ================================================================
Private Sub D3_SupportPairs(db As DAO.Database)
    H "D3. Parelles primari + secundari (punt 2)"

    Dim q As String
    q = "SELECT SP.Name & '  +  ' & SS.Name AS Parella, COUNT(*) AS N "
    q = q & "FROM (T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SUPPORT AS SP ON E.ID_Support=SP.ID) "
    q = q & "INNER JOIN L_SUPPORT AS SS ON E.ID_Support_Secondary=SS.ID "
    q = q & "GROUP BY SP.Name & '  +  ' & SS.Name "
    q = q & "ORDER BY COUNT(*) DESC;"
    PrintPairs db, q, "Parella", "N"

    Debug.Print ""
    Debug.Print "   Estructures amb suport simple (sense secundari): " & DCount("*", "T_STRUCTURES", "ID_Support Is Not Null And ID_Support_Secondary Is Null")
    Debug.Print "   Estructures sense cap suport assignat: " & DCount("*", "T_STRUCTURES", "ID_Support Is Null")
End Sub

' ================================================================
'  D4 - N_Basal_Bodies > 0 amb Base_Level = 0   (delta punt 4)
'  Decideix si la regla R27 es util o generadora de falsos positius.
' ================================================================
Private Sub D4_BasalBodiesWithoutB(db As DAO.Database)
    H "D4. Cossos basals sense element B (punt 4, regla R27)"

    Dim n As Long
    n = DCount("*", "T_STRUCTURES", "N_Basal_Bodies>0 And Base_Level=0")
    Debug.Print "   Casos amb N_Basal_Bodies>0 i Base_Level=0: " & n
    If n > 0 Then PrintCodes db, "N_Basal_Bodies>0 And Base_Level=0"

    Debug.Print ""
    Debug.Print "   Context:"
    Debug.Print "   N_Basal_Bodies>0 (total): " & DCount("*", "T_STRUCTURES", "N_Basal_Bodies>0")
    Debug.Print "   Base_Level=1,2,3 (total): " & DCount("*", "T_STRUCTURES", "Base_Level In (1,2,3)")
    Debug.Print "   N_Basal_Bodies Is Null:   " & DCount("*", "T_STRUCTURES", "N_Basal_Bodies Is Null")
    Debug.Print ""
    Debug.Print "   LECTURA: si els casos son pocs i son de veres incoherents,"
    Debug.Print "   R27 es util. Si son molts i tots explicables (basaments no"
    Debug.Print "   registrats com a B diferenciat), R27 generaria soroll."
End Sub

' ================================================================
'  D5 - Art rupestre: que hi ha ja al corpus  (delta 7.2 i 7.3)
'  Determina quantes entrades ROC calen i sobre quins casos aplicar
'  els quatre tests d'associacio.
' ================================================================
Private Sub D5_RockArtCandidates(db As DAO.Database)
    H "D5. Art rupestre i pigment sobre penya (punts 7.2, 7.3)"

    ' Registres ja tipificats com a PR
    Dim q As String
    q = "SELECT T.Name AS Tipus, COUNT(*) AS N FROM T_STRUCTURES AS E "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID "
    q = q & "WHERE T.Record_Class='Rock art panel' GROUP BY T.Name;"
    Debug.Print "   Registres amb classe Rock art panel:"
    PrintPairs db, q, "Tipus", "N"

    Debug.Print ""
    Debug.Print "   Estructures amb Pigment_Substrate='Bedrock': " & DCount("*", "T_STRUCTURES", "Pigment_Substrate='Bedrock'")
    PrintCodes db, "Pigment_Substrate='Bedrock'"

    Debug.Print ""
    Debug.Print "   Per extensio del pigment:"
    q = "SELECT Nz(Pigment_Extent,'(buit)') AS Ext, COUNT(*) AS N "
    q = q & "FROM T_STRUCTURES WHERE Pigment_Present In (1,2,3) "
    q = q & "GROUP BY Nz(Pigment_Extent,'(buit)') ORDER BY COUNT(*) DESC;"
    PrintPairs db, q, "Ext", "N"

    Debug.Print ""
    Debug.Print "   Files de T_DECORATIONS amb Substrate='Bedrock': " & DCount("*", "T_DECORATIONS", "Substrate='Bedrock'")
    Debug.Print ""
    Debug.Print "   ACCIO: sobre els codis llistats, aplicar els quatre tests"
    Debug.Print "   d'associacio (delta 7.3) i anotar quin decideix cada cas."
    Debug.Print "   Si algun no cau clarament en cap test, el criteri es incomplet."
End Sub

' ================================================================
'  D6 - Distribucio de Color a T_DECORATIONS   (delta 7.6)
'  Decideix la migracio de 'Both' a la parella dominant/secundari.
' ================================================================
Private Sub D6_ColorDistribution(db As DAO.Database)
    H "D6. Color a T_DECORATIONS (punt 7.6)"

    Dim q As String
    q = "SELECT Nz(Color,'(buit)') AS Col, COUNT(*) AS N FROM T_DECORATIONS "
    q = q & "GROUP BY Nz(Color,'(buit)') ORDER BY COUNT(*) DESC;"
    PrintPairs db, q, "Col", "N"

    Debug.Print ""
    Debug.Print "   Files totals a T_DECORATIONS: " & DCount("*", "T_DECORATIONS")
    Debug.Print "   Files amb Color='Both': " & DCount("*", "T_DECORATIONS", "Color='Both'")
    Debug.Print ""
    Debug.Print "   ACCIO: per a cada fila amb 'Both', determinar al model quin"
    Debug.Print "   color domina. Si no es resoluble -> Color='ND' amb nota."
End Sub

' ================================================================
'  D7 - Variancia de Chinking_Stones   (delta 8.6)
'  Un camp quasi constant no discrimina res.
' ================================================================
Private Sub D7_ChinkingVariance(db As DAO.Database)
    H "D7. Variancia de Chinking_Stones (punt 8.6)"

    PrintByteDist db, "Chinking_Stones"

    Dim n0 As Long, n1 As Long, nTot As Long
    n0 = DCount("*", "T_STRUCTURES", "Chinking_Stones=0")
    n1 = DCount("*", "T_STRUCTURES", "Chinking_Stones=1")
    nTot = n0 + n1
    Debug.Print ""
    If nTot = 0 Then
        Debug.Print "   Cap observacio valida encara: comprovacio ajornada."
    Else
        Debug.Print "   Denominador valid (0+1): " & nTot
        Debug.Print "   Proporcio de presents: " & Format(n1 / nTot, "0.0%")
        If n1 / nTot > 0.9 Or n1 / nTot < 0.1 Then
            Debug.Print "   -> QUASI CONSTANT. El camp no discrimina: valorar retirar-lo"
            Debug.Print "      o substituir-lo per una variable mes fina (densitat,"
            Debug.Print "      regularitat, posicio del ripio)."
        Else
            Debug.Print "   -> Te variancia util. Mantindre'l i valorar promocio a 5 valors."
        End If
    End If

    Debug.Print ""
    Debug.Print "   Per comparacio, Mortar_Present:"
    PrintByteDist db, "Mortar_Present"
End Sub

' ================================================================
'  D8 - Estat actual dels camps que canvien de defecte  (delta 8.2)
'  Els 0 existents no es toquen mai (regla B); aixo mostra quants
'  n'hi ha i quants podrien ser 0 heretats per inercia.
' ================================================================
Private Sub D8_DefaultChangeFields(db As DAO.Database)
    H "D8. Camps que passen a defecte NULL (punt 8.2)"

    Dim f(8) As String
    f(0) = "Mortar_Present"
    f(1) = "Chinking_Stones"
    f(2) = "Pigment_Present"
    f(3) = "Plaster_Present"
    f(4) = "Dec_Present"
    f(5) = "Looting"
    f(6) = "Fire_Damage"
    f(7) = "Animal_Activity"
    f(8) = "Modern_Access"

    Dim i As Integer
    For i = 0 To 8
        PrintByteDist db, f(i)
    Next i

    Debug.Print ""
    Debug.Print "   LECTURA: el canvi de defecte nomes afecta registres FUTURS"
    Debug.Print "   (regla B). Els 0 que ja hi son continuen sent 0. Si una"
    Debug.Print "   columna te molts 0 i pocs 9, val la pena revisar si aquells"
    Debug.Print "   0 son absencies verificades o heretades del defecte v11."
End Sub

' ================================================================
'  D9 - Capes aplicades: hi ha casos que demanen el valor 3?
'  (delta 8.3). Si no n'hi ha cap, la promocio a cinc valors es
'  teoricament correcta pero practicament prematura.
' ================================================================
Private Sub D9_LayerLossCandidates(db As DAO.Database)
    H "D9. Capes aplicades: candidats al valor 3 (punt 8.3)"

    Debug.Print "   Files de T_LOST_ELEMENTS per tipus d'evidencia:"
    Dim q As String
    q = "SELECT EV.Name AS Evid, COUNT(*) AS N FROM T_LOST_ELEMENTS AS L "
    q = q & "INNER JOIN L_LOST_EVIDENCE AS EV ON L.ID_Evidence_Type=EV.ID "
    q = q & "GROUP BY EV.Name ORDER BY COUNT(*) DESC;"
    PrintPairs db, q, "Evid", "N"

    Debug.Print ""
    Dim idM As Long
    Dim idP As Long
    idM = Nz(DLookup("ID", "L_LOST_EVIDENCE", "Name='Mortar imprint'"), 0)
    idP = Nz(DLookup("ID", "L_LOST_EVIDENCE", "Name='Pigment on bedrock'"), 0)
    Debug.Print "   Files amb evidencia 'Mortar imprint': " & DCount("*", "T_LOST_ELEMENTS", "ID_Evidence_Type=" & idM)
    Debug.Print "   Files amb evidencia 'Pigment on bedrock': " & DCount("*", "T_LOST_ELEMENTS", "ID_Evidence_Type=" & idP)
    Debug.Print ""
    Debug.Print "   ACCIO: si hi ha zero files a tota T_LOST_ELEMENTS, aquesta"
    Debug.Print "   comprovacio no diu res encara. Cal mirar-ho a ma: hi ha"
    Debug.Print "   estructures amb empremta de morter sense morter, o amb"
    Debug.Print "   adherencies de revoc sense revoc? Si no n'hi ha cap, la"
    Debug.Print "   promocio de 8.3 es correcta pero pot esperar."
End Sub
