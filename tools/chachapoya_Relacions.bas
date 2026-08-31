Option Compare Database
Option Explicit

' ================================================================
'  chachapoya_Relacions_v25f.bas  (2026-08-31)
'  La Petaca & Diablo Wasi (Leymebamba, Amazonas, Peru)
'  Author: Esteve Ribera Torro | TFM Arqueologia UA
'  Spec: DELTA_v25e_v25f.md
'
'  Entrada quirurgica del model de relacions (arestes + paternitat):
'
'    RelacionsV25f [True]
'      Mode informe (per defecte): verifica TOT (codis, parelles
'      duplicades, estat previ de cada fila a tocar) i llista el
'      pla sencer sense escriure res. Mode aplicacio (True): torna
'      a verificar, escriu, i re-verifica.
'
'  Contingut de l acta:
'    - 9 assignacions ID_Parent (diedre EA76: 5 filles; cova EA42:
'      +EA34; cova DW-S05-EA02: 1; cova DW-S06-EA04: 2)
'    - 1 correccio de tipologia: DW-S01-EA76 passa de 9
'      (Unclassifiable, reservat a estructures construides) a 6
'      (CAV context natural), doctrina de L_TYPOLOGY
'    - 4 esmenes de files existents de T_CONNECTIONS (ID 3, 4, 5,
'      6): tipus i cronologia firmada en sessio
'    - 72 arestes noves en cinc families del domini tancat v11:
'      Shared support (33), Abutted vertical joint (10),
'      Superposition (1), Associated natural context (1),
'      Vertical association (27)
'    - Deduplicacio aplicada: quan un parell apareix en mes d una
'      llista, guanya la relacio mes especifica (contacte >
'      proximitat > estrat compartit); 12 parells d estrat
'      subsumits per adossament (fet derivable, cap perdua)
'
'  Normalitzacio del parell: ID_Struct_A < ID_Struct_B (index
'  unic UQ_CONN_PAIR). Regla 52 (ID_Earlier dins del parell i
'  coherent amb Chrono_Relation) es respecta per construccio.
'
'  Sequencia: RelacionsV25f -> revisar informe -> RelacionsV25f
'  True -> reobrir QRY_14_Connections_Edges i QRY_16.
'  MAI BuildDB() sobre esta base.
' ================================================================

Public Sub RelacionsV25f(Optional ByVal Apply As Boolean = False)
    Dim db As DAO.Database
    Set db = CurrentDb
    Debug.Print "=== RelacionsV25f (" & IIf(Apply, "APLICACIO", "informe") & ") ==="
    Dim bad As Long
    bad = 0

    ' ---------- 1. PATERNITAT ----------
    Dim pc(8) As String, pp(8) As String
    pc(0) = "DW-S01-EA02": pp(0) = "DW-S01-EA76"
    pc(1) = "DW-S01-EA13": pp(1) = "DW-S01-EA76"
    pc(2) = "DW-S01-EA14": pp(2) = "DW-S01-EA76"
    pc(3) = "DW-S01-EA51": pp(3) = "DW-S01-EA76"
    pc(4) = "DW-S01-EA59": pp(4) = "DW-S01-EA76"
    pc(5) = "DW-S01-EA34": pp(5) = "DW-S01-EA42"
    pc(6) = "DW-S05-EA01": pp(6) = "DW-S05-EA02"
    pc(7) = "DW-S06-EA01": pp(7) = "DW-S06-EA04"
    pc(8) = "DW-S06-EA02": pp(8) = "DW-S06-EA04"
    Dim i As Integer
    Dim cid As Long, pid As Long
    Dim curp As Variant
    For i = 0 To 8
        cid = IdOf(pc(i)): pid = IdOf(pp(i))
        If cid = 0 Or pid = 0 Then
            Debug.Print "  [CODI INEXISTENT] " & pc(i) & " o " & pp(i)
            bad = bad + 1
        Else
            curp = DLookup("ID_Parent", "T_STRUCTURES", "ID=" & cid)
            If IsNull(curp) Then
                Debug.Print "  [PARE] " & pc(i) & " -> " & pp(i)
            ElseIf curp = pid Then
                Debug.Print "  [PARE JA ASSIGNAT] " & pc(i) & " -> " & pp(i) & " (idempotent)"
            Else
                Debug.Print "  [CONFLICTE PARE] " & pc(i) & " ja te un altre pare (ID=" & curp & ")"
                bad = bad + 1
            End If
        End If
    Next i

    ' ---------- 2. TIPOLOGIA DEL DIEDRE ----------
    Dim t76 As Variant
    t76 = DLookup("ID_Typology", "T_STRUCTURES", "Code='DW-S01-EA76'")
    If Nz(t76, 0) = 9 Then
        Debug.Print "  [TIPOLOGIA] DW-S01-EA76: 9 (Unclassifiable) -> 6 (CAV); el diedre es context natural"
    ElseIf Nz(t76, 0) = 6 Then
        Debug.Print "  [TIPOLOGIA JA CORRECTA] DW-S01-EA76 = 6 (idempotent)"
    Else
        Debug.Print "  [NO QUADRA] DW-S01-EA76: tipologia esperada 9, trobada " & Nz(t76, "(NULL)")
        bad = bad + 1
    End If

    ' ---------- 3. ESMENES DE FILES EXISTENTS ----------
    bad = bad + ChkRow(db, 3, "DW-S01-EA22", "DW-S01-EA71", "Superposition", "Abutted vertical joint + Sequential (EA71 anterior)")
    bad = bad + ChkRow(db, 4, "DW-S01-EA40", "DW-S01-EA72", "Abutted vertical joint", "+ Sequential (EA40 anterior)")
    bad = bad + ChkRow(db, 5, "DW-S04-EA02", "DW-S04-EA21", "Other (see Notes)", "Associated natural context + Sequential (EA21 anterior)")
    bad = bad + ChkRow(db, 6, "DW-S04-EA12", "DW-S04-EA20", "Superposition", "Abutted vertical joint")

    ' ---------- 4. ARESTES NOVES ----------
    Dim col As Collection
    Set col = New Collection
    ' --- Shared support (mateix estrat de suport, encadenades) ---
    AddE col, "DW-S01-EA06", "DW-S01-EA07", "Shared support", "Undetermined", "", ""
    AddE col, "DW-S01-EA08", "DW-S01-EA24", "Shared support", "Undetermined", "", ""
    AddE col, "DW-S01-EA23", "DW-S01-EA22", "Shared support", "Undetermined", "", ""
    AddE col, "DW-S01-EA17", "DW-S01-EA16", "Shared support", "Undetermined", "", ""
    AddE col, "DW-S01-EA16", "DW-S01-EA52", "Shared support", "Undetermined", "", ""
    AddE col, "DW-S01-EA52", "DW-S01-EA13", "Shared support", "Undetermined", "", ""
    AddE col, "DW-S01-EA13", "DW-S01-EA02", "Shared support", "Undetermined", "", ""
    AddE col, "DW-S01-EA02", "DW-S01-EA47", "Shared support", "Undetermined", "", ""
    AddE col, "DW-S01-EA48", "DW-S01-EA46", "Shared support", "Undetermined", "", ""
    AddE col, "DW-S01-EA02", "DW-S01-EA14", "Shared support", "Undetermined", "", ""
    AddE col, "DW-S01-EA59", "DW-S01-EA18", "Shared support", "Undetermined", "", ""
    AddE col, "DW-S01-EA46", "DW-S01-EA49", "Shared support", "Undetermined", "", ""
    AddE col, "DW-S01-EA49", "DW-S01-EA73", "Shared support", "Undetermined", "", ""
    AddE col, "DW-S01-EA19", "DW-S01-EA45", "Shared support", "Undetermined", "", ""
    AddE col, "DW-S01-EA44", "DW-S01-EA04", "Shared support", "Undetermined", "", ""
    AddE col, "DW-S01-EA53", "DW-S01-EA54", "Shared support", "Undetermined", "", ""
    AddE col, "DW-S01-EA54", "DW-S01-EA55", "Shared support", "Undetermined", "", ""
    AddE col, "DW-S01-EA55", "DW-S01-EA33", "Shared support", "Undetermined", "", ""
    AddE col, "DW-S01-EA33", "DW-S01-EA61", "Shared support", "Undetermined", "", ""
    AddE col, "DW-S01-EA28", "DW-S01-EA58", "Shared support", "Undetermined", "", ""
    AddE col, "DW-S01-EA58", "DW-S01-EA57", "Shared support", "Undetermined", "", ""
    AddE col, "DW-S01-EA57", "DW-S01-EA29", "Shared support", "Undetermined", "", ""
    AddE col, "DW-S01-EA29", "DW-S01-EA56", "Shared support", "Undetermined", "", ""
    AddE col, "DW-S01-EA56", "DW-S01-EA30", "Shared support", "Undetermined", "", ""
    AddE col, "DW-S01-EA30", "DW-S01-EA31", "Shared support", "Undetermined", "", ""
    AddE col, "DW-S01-EA31", "DW-S01-EA32", "Shared support", "Undetermined", "", ""
    AddE col, "DW-S01-EA41", "DW-S01-EA40", "Shared support", "Undetermined", "", ""
    AddE col, "DW-S01-EA36", "DW-S01-EA74", "Shared support", "Undetermined", "", ""
    AddE col, "DW-S04-EA02", "DW-S04-EA08", "Shared support", "Undetermined", "", ""
    AddE col, "DW-S04-EA03", "DW-S04-EA11", "Shared support", "Undetermined", "", ""
    AddE col, "DW-S04-EA11", "DW-S04-EA17", "Shared support", "Undetermined", "", ""
    AddE col, "DW-S04-EA09", "DW-S04-EA12", "Shared support", "Undetermined", "", ""
    AddE col, "DW-S04-EA20", "DW-S04-EA18", "Shared support", "Undetermined", "", ""
    ' --- Abutted vertical joint (adossades) ---
    AddE col, "DW-S01-EA17", "DW-S01-EA71", "Abutted vertical joint", "Undetermined", "", ""
    AddE col, "DW-S01-EA14", "DW-S01-EA51", "Abutted vertical joint", "Undetermined", "", ""
    AddE col, "DW-S01-EA51", "DW-S01-EA59", "Abutted vertical joint", "Undetermined", "", ""
    AddE col, "DW-S01-EA73", "DW-S01-EA19", "Abutted vertical joint", "Undetermined", "", ""
    AddE col, "DW-S01-EA48", "DW-S01-EA47", "Abutted vertical joint", "Undetermined", "", ""
    AddE col, "DW-S01-EA03", "DW-S01-EA04", "Abutted vertical joint", "Undetermined", "", "Llistada tambe com a proximitat vertical; prima el contacte (v25f)"
    AddE col, "DW-S01-EA04", "DW-S01-EA05", "Abutted vertical joint", "Undetermined", "", ""
    AddE col, "DW-S01-EA45", "DW-S01-EA44", "Abutted vertical joint", "Undetermined", "", ""
    AddE col, "DW-S02-EA01", "DW-S02-EA02", "Abutted vertical joint", "Undetermined", "", ""
    AddE col, "DW-S01-EA20", "DW-S01-EA75", "Abutted vertical joint", "Undetermined", "", "EA75 al fons de la cova (v25f)"
    ' --- Superposition (apilades) ---
    AddE col, "DW-S01-EA06", "DW-S01-EA62", "Superposition", "Sequential", "DW-S01-EA06", "EA62 descansa sobre EA06; EA06 anterior per estratigrafia (v25f)"
    ' --- Associated natural context ---
    AddE col, "DW-S03-EA01", "DW-S03-EA02", "Associated natural context", "Sequential", "DW-S03-EA02", "Mausoleu associat al nitxo natural EA02; el context natural precedeix (v25f)"
    ' --- Vertical association (proximitat vertical/diagonal, sense contacte) ---
    AddE col, "DW-S01-EA09", "DW-S01-EA10", "Vertical association", "Undetermined", "", ""
    AddE col, "DW-S01-EA18", "DW-S01-EA17", "Vertical association", "Undetermined", "", ""
    AddE col, "DW-S01-EA59", "DW-S01-EA52", "Vertical association", "Undetermined", "", ""
    AddE col, "DW-S01-EA52", "DW-S01-EA15", "Vertical association", "Undetermined", "", ""
    AddE col, "DW-S01-EA14", "DW-S01-EA13", "Vertical association", "Undetermined", "", ""
    AddE col, "DW-S01-EA03", "DW-S01-EA02", "Vertical association", "Undetermined", "", ""
    AddE col, "DW-S01-EA36", "DW-S01-EA35", "Vertical association", "Undetermined", "", ""
    AddE col, "DW-S01-EA35", "DW-S01-EA34", "Vertical association", "Undetermined", "", ""
    AddE col, "DW-S01-EA34", "DW-S01-EA33", "Vertical association", "Undetermined", "", ""
    AddE col, "DW-S01-EA72", "DW-S01-EA39", "Vertical association", "Undetermined", "", "Quasi al mateix nivell (dAlt 0.7 m); mes diagonal que vertical (QC v25f)"
    AddE col, "DW-S01-EA39", "DW-S01-EA38", "Vertical association", "Undetermined", "", ""
    AddE col, "DW-S02-EA01", "DW-S02-EA04", "Vertical association", "Undetermined", "", ""
    AddE col, "DW-S02-EA04", "DW-S02-EA06", "Vertical association", "Undetermined", "", ""
    AddE col, "DW-S04-EA15", "DW-S04-EA14", "Vertical association", "Undetermined", "", ""
    AddE col, "DW-S04-EA14", "DW-S04-EA20", "Vertical association", "Undetermined", "", ""
    AddE col, "DW-S04-EA12", "DW-S04-EA11", "Vertical association", "Undetermined", "", ""
    AddE col, "DW-S04-EA11", "DW-S04-EA10", "Vertical association", "Undetermined", "", ""
    AddE col, "DW-S04-EA18", "DW-S04-EA17", "Vertical association", "Undetermined", "", ""
    AddE col, "DW-S04-EA19", "DW-S04-EA09", "Vertical association", "Undetermined", "", ""
    AddE col, "DW-S04-EA06", "DW-S04-EA05", "Vertical association", "Undetermined", "", ""
    AddE col, "DW-S04-EA05", "DW-S04-EA04", "Vertical association", "Undetermined", "", ""
    AddE col, "DW-S04-EA04", "DW-S04-EA03", "Vertical association", "Undetermined", "", ""
    AddE col, "DW-S04-EA03", "DW-S04-EA02", "Vertical association", "Undetermined", "", ""
    AddE col, "DW-S04-EA08", "DW-S04-EA07", "Vertical association", "Undetermined", "", ""
    AddE col, "DW-S04-EA07", "DW-S04-EA01", "Vertical association", "Undetermined", "", ""
    AddE col, "DW-S01-EA60", "DW-S01-EA39", "Vertical association", "Undetermined", "", ""
    AddE col, "DW-S01-EA20", "DW-S01-EA22", "Vertical association", "Undetermined", "", ""

    Dim reg As Collection
    Set reg = New Collection
    Dim s As String, p() As String
    Dim ka As Long, kb As Long, tmpL As Long
    Dim nNew As Long
    nNew = 0
    For i = 1 To col.Count
        s = col(i)
        p = Split(s, "|")
        ka = IdOf(p(0)): kb = IdOf(p(1))
        If ka = 0 Or kb = 0 Then
            Debug.Print "  [CODI INEXISTENT] " & p(0) & " o " & p(1)
            bad = bad + 1
        ElseIf ka = kb Then
            Debug.Print "  [PARELL AMB SI MATEIX] " & p(0)
            bad = bad + 1
        Else
            If ka > kb Then
                tmpL = ka: ka = kb: kb = tmpL
            End If
            On Error Resume Next
            reg.Add "x", "K" & ka & "_" & kb
            If Err.Number <> 0 Then
                On Error GoTo 0
                Debug.Print "  [DUPLICAT AL LOT] " & p(0) & " <-> " & p(1)
                bad = bad + 1
            Else
                On Error GoTo 0
                If DCount("*", "T_CONNECTIONS", "(ID_Struct_A=" & ka & " AND ID_Struct_B=" & kb & ") OR (ID_Struct_A=" & kb & " AND ID_Struct_B=" & ka & ")") > 0 Then
                    Debug.Print "  [JA EXISTEIX A LA TAULA] " & p(0) & " <-> " & p(1)
                    bad = bad + 1
                Else
                    nNew = nNew + 1
                End If
            End If
        End If
    Next i
    Debug.Print "  [ARESTES NOVES VERIFICADES] " & nNew & " de " & col.Count

    If bad > 0 Then
        Debug.Print "AVORTAT: " & bad & " verificacions fallides. Cap escriptura."
        Exit Sub
    End If
    If Not Apply Then
        Debug.Print "Verificacio en verd. Executa RelacionsV25f True per a aplicar."
        Exit Sub
    End If

    ' ---------- 5. ESCRIPTURA ----------
    For i = 0 To 8
        cid = IdOf(pc(i)): pid = IdOf(pp(i))
        If IsNull(DLookup("ID_Parent", "T_STRUCTURES", "ID=" & cid)) Then
            db.Execute "UPDATE T_STRUCTURES SET ID_Parent=" & pid & " WHERE ID=" & cid, dbFailOnError
        End If
    Next i
    If Nz(DLookup("ID_Typology", "T_STRUCTURES", "Code='DW-S01-EA76'"), 0) = 9 Then
        db.Execute "UPDATE T_STRUCTURES SET ID_Typology=6 WHERE Code='DW-S01-EA76'", dbFailOnError
    End If
    UpdRow db, 3, "Abutted vertical joint", "Sequential", IdOf("DW-S01-EA71"), "EA22b descansa parcialment sobre EA71 (escalo de massa reduida); prima la junta adossada. EA71 anterior per estratigrafia (v25f)."
    UpdRow db, 4, "", "Sequential", IdOf("DW-S01-EA40"), "EA40 anterior a EA72 per estratigrafia (v25f)."
    UpdRow db, 5, "Associated natural context", "Sequential", IdOf("DW-S04-EA21"), "Plataforma apilada sobre el nitxo natural EA21; el context natural precedeix la construccio (v25f)."
    UpdRow db, 6, "Abutted vertical joint", "", 0, ""
    Dim nIns As Long
    nIns = 0
    For i = 1 To col.Count
        p = Split(col(i), "|")
        ka = IdOf(p(0)): kb = IdOf(p(1))
        If ka > kb Then
            tmpL = ka: ka = kb: kb = tmpL
        End If
        s = "INSERT INTO T_CONNECTIONS (ID_Struct_A, ID_Struct_B, Connection_Type, Chrono_Relation, Confidence, ID_Earlier, Notes) VALUES ("
        s = s & ka & ", " & kb & ", '" & p(2) & "', '" & p(3) & "', 'High', "
        If p(4) = "" Then
            s = s & "NULL"
        Else
            s = s & IdOf(p(4))
        End If
        If p(5) = "" Then
            s = s & ", NULL)"
        Else
            s = s & ", '" & p(5) & "')"
        End If
        db.Execute s, dbFailOnError
        nIns = nIns + 1
    Next i

    ' ---------- 6. RE-VERIFICACIO ----------
    Debug.Print "APLICAT: " & nIns & " arestes inserides. Total a T_CONNECTIONS: " & DCount("*", "T_CONNECTIONS")
    Debug.Print "ID_Parent poblats: " & DCount("*", "T_STRUCTURES", "ID_Parent Is Not Null")
    Debug.Print "Sequential amb ID_Earlier: " & DCount("*", "T_CONNECTIONS", "Chrono_Relation='Sequential' AND ID_Earlier Is Not Null")
    Debug.Print "Seguent pas: reobrir QRY_14_Connections_Edges i QRY_16_Validation_Check."
End Sub

' ---------------- auxiliars privats ----------------

Private Sub AddE(col As Collection, a As String, b As String, t As String, chr_ As String, earl As String, note As String)
    col.Add a & "|" & b & "|" & t & "|" & chr_ & "|" & earl & "|" & note
End Sub

Private Function IdOf(code As String) As Long
    IdOf = Nz(DLookup("ID", "T_STRUCTURES", "Code='" & code & "'"), 0)
End Function

Private Function ChkRow(db As DAO.Database, rid As Long, ca As String, cb As String, expType As String, planned As String) As Long
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT * FROM T_CONNECTIONS WHERE ID=" & rid, dbOpenDynaset)
    If rs.EOF Then
        Debug.Print "  [FILA INEXISTENT] T_CONNECTIONS ID " & rid
        ChkRow = 1
    Else
        Dim a As String, b As String
        a = Nz(DLookup("Code", "T_STRUCTURES", "ID=" & rs!ID_Struct_A), "?")
        b = Nz(DLookup("Code", "T_STRUCTURES", "ID=" & rs!ID_Struct_B), "?")
        If (a = ca And b = cb) Or (a = cb And b = ca) Then
            If Nz(rs!Connection_Type, "") = expType Then
                Debug.Print "  [ESMENA ID " & rid & "] " & a & " <-> " & b & ": " & planned
                ChkRow = 0
            Else
                Debug.Print "  [NO QUADRA ID " & rid & "] tipus esperat '" & expType & "', trobat '" & Nz(rs!Connection_Type, "(NULL)") & "'"
                ChkRow = 1
            End If
        Else
            Debug.Print "  [NO QUADRA ID " & rid & "] parell esperat " & ca & "/" & cb & ", trobat " & a & "/" & b
            ChkRow = 1
        End If
    End If
    rs.Close
End Function

Private Sub UpdRow(db As DAO.Database, rid As Long, newType As String, newChrono As String, earlId As Long, newNote As String)
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT * FROM T_CONNECTIONS WHERE ID=" & rid, dbOpenDynaset)
    If Not rs.EOF Then
        rs.Edit
        If newType <> "" Then rs!Connection_Type = newType
        If newChrono <> "" Then rs!Chrono_Relation = newChrono
        If earlId > 0 Then rs!ID_Earlier = earlId
        If newNote <> "" Then rs!Notes = newNote
        rs.Update
    End If
    rs.Close
End Sub
