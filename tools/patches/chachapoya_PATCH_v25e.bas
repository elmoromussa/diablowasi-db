Option Compare Database
Option Explicit

' ================================================================
'  chachapoya_PATCH_v25e.bas  (2026-08-31)
'  La Petaca & Diablo Wasi (Leymebamba, Amazonas, Peru)
'  Author: Esteve Ribera Torro | TFM Arqueologia UA
'  Spec: DELTA_v25d_v25e.md
'
'  Paquet quirurgic v25e. Dos blocs publics, cadascun amb mode
'  informe (per defecte) i mode aplicacio (argument True):
'
'    RepairC5Locale [True]
'      Repara el bug de locale d'ApplyMetricReview v25: CDbl
'      llegia el punt decimal de Proposed_Value com a separador
'      de milers sota locale de coma (1.39 -> 139), i 125 valors
'      C5 van entrar a T_STRUCTURES multiplicats per 100. La
'      reparacio NO divideix per 100: RE-DERIVA cada valor des de
'      Proposed_Value de T_METRIC_REVIEW (files C5 aplicades amb
'      camp numeric), que conserva el valor bo exacte. Per fila:
'        - viu = proposta        -> ja correcte, salta (idempotent)
'        - viu = proposta x 100  -> repara (o llista, en informe)
'        - qualsevol altra cosa  -> [DIVERGENT], NO es toca
'      En aplicar, estampa a Notes de la fila TMR:
'      "Repaired v25e (locale CDbl x100)". Esperat sobre el corpus
'      v25d: 125 reparats, 28 omesos (Dim_Method, text), 0
'      divergents.
'
'    SurgeryMuntantsV25e [True]
'      Acta v25e de muntants: reclassifica les sis parelles
'      entrades duplicant l'unic membre observat, i passa el
'      Jamb_Fabric_Reveal de DW-S02-EA02 d'1 a 9 (parella ND: la
'      fabrica-fa-de-brancal no es va poder verificar). Dues
'      passades: la primera verifica TOTS els valors actuals i,
'      si cap no quadra, AVORTA sense escriure res; la segona
'      escriu i re-verifica.
'
'  Sequencia completa v25e (BD amb dades, estat v25d):
'    1. Importar/reemplacar moduls: aquest PATCH, DB v25e,
'       CheckMetrics v25e, Form v25e
'    2. RepairC5Locale            (informe)
'    3. RepairC5Locale True       (aplicacio)
'    4. SurgeryMuntantsV25e       (informe)
'    5. SurgeryMuntantsV25e True  (aplicacio)
'    6. RebuildQueriesV25         (modul DB v25e: R69 esmenada)
'    7. BuildForm                 (modul Form v25e: combo ampliat;
'       F_STRUCTURES tancat abans)
'    8. Obrir QRY_16_Validation_Check i confirmar el verd
'  MAI BuildDB() sobre esta base.
' ================================================================

Public Sub RepairC5Locale(Optional ByVal Apply As Boolean = False)
    Dim db As DAO.Database
    Set db = CurrentDb
    Dim sql As String
    sql = "SELECT ID, ID_Structure, Field_Name, Proposed_Value, Notes "
    sql = sql & "FROM T_METRIC_REVIEW "
    sql = sql & "WHERE Check_Code='C5' AND Applied_On Is Not Null"
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset(sql, dbOpenDynaset)
    Dim nOK As Long, nRep As Long, nDiv As Long, nSkip As Long
    Debug.Print "=== RepairC5Locale (" & IIf(Apply, "APLICACIO", "informe") & ") ==="
    Dim fld As String, pv As String, sid As Long
    Dim p As Double, cd As String
    Dim liveV As Variant
    Dim rs2 As DAO.Recordset
    Do While Not rs.EOF
        fld = Nz(rs!Field_Name, "")
        pv = Nz(rs!Proposed_Value, "")
        sid = Nz(rs!ID_Structure, 0)
        If fld = "" Or pv = "" Then
            nSkip = nSkip + 1
        ElseIf Not IsNumericFieldP(db, fld) Then
            nSkip = nSkip + 1
        Else
            p = Val(Replace(pv, ",", "."))
            cd = Nz(DLookup("Code", "T_STRUCTURES", "ID=" & sid), "?")
            liveV = DLookup(fld, "T_STRUCTURES", "ID=" & sid)
            If IsNull(liveV) Then
                Debug.Print "  [ANOMALIA] " & cd & "." & fld & ": camp NULL amb fila aplicada (no es toca)"
                nDiv = nDiv + 1
            ElseIf Abs(CDbl(liveV) - p) < 0.005 Then
                nOK = nOK + 1
            ElseIf Abs(CDbl(liveV) - p * 100#) < 0.5 Then
                If Apply Then
                    Set rs2 = db.OpenRecordset("SELECT " & fld & " FROM T_STRUCTURES WHERE ID=" & sid, dbOpenDynaset)
                    If Not rs2.EOF Then
                        rs2.Edit
                        rs2.Fields(fld).Value = p
                        rs2.Update
                    End If
                    rs2.Close
                    rs.Edit
                    rs!Notes = AppendTxt(Nz(rs!Notes, ""), "Repaired v25e (locale CDbl x100)")
                    rs.Update
                    Debug.Print "  [REPARAT] " & cd & "." & fld & ": " & liveV & " -> " & pv
                Else
                    Debug.Print "  [REPARARIA] " & cd & "." & fld & ": " & liveV & " -> " & pv
                End If
                nRep = nRep + 1
            Else
                Debug.Print "  [DIVERGENT] " & cd & "." & fld & ": viu=" & liveV & " proposta=" & pv & " (edicio manual posterior? NO es toca)"
                nDiv = nDiv + 1
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
    Debug.Print "--- correctes=" & nOK & "  " & IIf(Apply, "reparats=", "repararia=") & nRep & "  divergents=" & nDiv & "  omesos(no numerics)=" & nSkip
    If nDiv > 0 Then Debug.Print "ATENCIO: hi ha files divergents o anomales. Revisa-les a ma abans de continuar."
    If Not Apply Then Debug.Print "Mode informe: cap escriptura. Executa RepairC5Locale True per a aplicar."
End Sub

Public Sub SurgeryMuntantsV25e(Optional ByVal Apply As Boolean = False)
    Dim db As DAO.Database
    Set db = CurrentDb
    Dim codes(5) As String, oldV(5) As String, newV(5) As String
    codes(0) = "DW-S04-EA18": oldV(0) = "Slab-Slab": newV(0) = "Slab-NObs"
    codes(1) = "DW-S02-EA01": oldV(1) = "Composite-Composite": newV(1) = "Composite-NObs"
    codes(2) = "DW-S03-EA01": oldV(2) = "Composite-Composite": newV(2) = "Composite-NObs"
    codes(3) = "DW-S01-EA03": oldV(3) = "Composite-Composite": newV(3) = "Composite-NObs"
    codes(4) = "DW-S04-EA12": oldV(4) = "Composite-Composite": newV(4) = "ND"
    codes(5) = "DW-S01-EA22": oldV(5) = "Fabric-Fabric": newV(5) = "Fabric-NObs"
    Debug.Print "=== SurgeryMuntantsV25e (" & IIf(Apply, "APLICACIO", "informe") & ") ==="
    ' ---- Passada 1: verificacio total. Res no s'escriu si una
    ' ---- sola fila no quadra amb l'acta.
    Dim bad As Long, i As Integer
    Dim cur As Variant
    bad = 0
    For i = 0 To 5
        cur = DLookup("Jamb_Fabric", "T_STRUCTURES", "Code='" & codes(i) & "'")
        If Nz(cur, "(NULL)") <> oldV(i) Then
            Debug.Print "  [NO QUADRA] " & codes(i) & ": esperat '" & oldV(i) & "', trobat '" & Nz(cur, "(NULL)") & "'"
            bad = bad + 1
        Else
            Debug.Print "  [VERIFICAT] " & codes(i) & ": " & oldV(i) & " -> " & newV(i)
        End If
    Next i
    Dim jf As Variant, jfr As Variant
    jf = DLookup("Jamb_Fabric", "T_STRUCTURES", "Code='DW-S02-EA02'")
    jfr = DLookup("Jamb_Fabric_Reveal", "T_STRUCTURES", "Code='DW-S02-EA02'")
    If Nz(jf, "(NULL)") <> "ND" Or Nz(jfr, -1) <> 1 Then
        Debug.Print "  [NO QUADRA] DW-S02-EA02: esperat parella 'ND' i JFR=1, trobat '" & Nz(jf, "(NULL)") & "' i JFR=" & Nz(jfr, "(NULL)")
        bad = bad + 1
    Else
        Debug.Print "  [VERIFICAT] DW-S02-EA02: JFR 1 -> 9 (la parella ND es mante)"
    End If
    If bad > 0 Then
        Debug.Print "AVORTAT: " & bad & " verificacions fallides. Cap escriptura."
        Exit Sub
    End If
    If Not Apply Then
        Debug.Print "Verificacio en verd (7/7). Executa SurgeryMuntantsV25e True per a aplicar."
        Exit Sub
    End If
    ' ---- Passada 2: escriptura + nota d'acta + re-verificacio.
    For i = 0 To 5
        db.Execute "UPDATE T_STRUCTURES SET Jamb_Fabric='" & newV(i) & "' WHERE Code='" & codes(i) & "'", dbFailOnError
    Next i
    db.Execute "UPDATE T_STRUCTURES SET Jamb_Fabric_Reveal=9 WHERE Code='DW-S02-EA02'", dbFailOnError
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT Systems_Notes FROM T_STRUCTURES WHERE Code='DW-S02-EA02'", dbOpenDynaset)
    If Not rs.EOF Then
        rs.Edit
        rs!Systems_Notes = AppendTxt(Nz(rs!Systems_Notes, ""), "JFR 1->9 (v25e): parella de muntants ND, l'observacio positiva de fabrica-fa-de-brancal no es verificable amb la documentacio actual; pendent de revisio amb foto o en camp.")
        rs.Update
    End If
    rs.Close
    bad = 0
    For i = 0 To 5
        If Nz(DLookup("Jamb_Fabric", "T_STRUCTURES", "Code='" & codes(i) & "'"), "") <> newV(i) Then bad = bad + 1
    Next i
    If Nz(DLookup("Jamb_Fabric_Reveal", "T_STRUCTURES", "Code='DW-S02-EA02'"), -1) <> 9 Then bad = bad + 1
    If bad = 0 Then
        Debug.Print "APLICAT i re-verificat: 6 reclassificacions + JFR d'EA02 + nota d'acta."
        Debug.Print "Seguent pas: RebuildQueriesV25 (R69 esmenada) i reobrir QRY_16_Validation_Check."
    Else
        Debug.Print "ATENCIO: " & bad & " re-verificacions fallides. Revisa a ma."
    End If
End Sub

' ---------------- auxiliars privats ----------------

Private Function IsNumericFieldP(db As DAO.Database, fld As String) As Boolean
    On Error GoTo fail
    Dim t As Integer
    t = db.TableDefs("T_STRUCTURES").Fields(fld).Type
    IsNumericFieldP = (t = dbByte Or t = dbInteger Or t = dbLong Or t = dbSingle Or t = dbDouble Or t = dbCurrency Or t = dbDecimal)
    Exit Function
fail:
    IsNumericFieldP = False
End Function

Private Function AppendTxt(base As String, addTxt As String) As String
    If Trim(base) = "" Then
        AppendTxt = addTxt
    Else
        AppendTxt = base & " | " & addTxt
    End If
End Function
