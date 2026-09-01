Option Compare Database
Option Explicit

' ================================================================
'  chachapoya_Tancament_v25g.bas  (2026-09-01)
'  Tancament de la worklist i la bateria: cascada d'observacions
'  firmades en sessio (P0-P8) + decisions de regla (4.1-4.23) +
'  noves EA78/EA79 amb les seues arestes.
'  TancamentV25g [True]: informe per defecte; True per a aplicar.
'  La cascada NOMES escriu sobre camps NULL (mai trepitja res).
'  Les cirurgies puntuals verifiquen el valor actual abans.
'  Els sis valors [DEFAULT] triats per l'assistent es marquen a
'  l'informe: revisa'ls i discrepa abans del True si cal.
'  Despres: importar chachapoya_DB_v25g.bas (substituint el DB
'  v25e), executar RebuildQueriesV25 (R54 i R66 esmenades) i
'  reobrir QRY_16_Validation_Check.
' ================================================================

Private Const ELEMS As String = "Embedded_Base_Beams;Base_Level;Decorative_Socle;Tie_Walls;Timber_Brackets;Transverse_Beams;Corbelled_Courses;Interbody_Cornice;Corner_Quoins;Structural_Pilasters;Facade_Flank;Relief_Frieze;Sill;Jambs;Lintel;Upper_Crown;Eave_Beam;Eave_Surface;Return_Wall;Chamber_Roof;Platform_Surface"
Private Const BIOF As String = "Anatomical_Connection;Mummification;Funerary_Bundles;Dispersed_Remains;Flexed_Position;Bone_Burning"
Private Const MATF As String = "Mat_Textiles;Mat_Wood;Mat_VegFiber;Mat_Ceramics;Mat_Fauna;Mat_DeerAntler;Mat_Other"
Private Const DANF As String = "Looting;Fire_Damage;Animal_Activity;Modern_Access"
Private Planned As Collection

Private Const MASF As String = "Mortar_Present;Plaster_Present;Pigment_Present;Dec_Present;RockArt_Present;Chinking_Stones"

Public Sub TancamentV25g(Optional ByVal Apply As Boolean = False)
    Dim db As DAO.Database
    Set db = CurrentDb
    Set Planned = New Collection
    Debug.Print "=== TancamentV25g (" & IIf(Apply, "APLICACIO", "informe") & ") ==="
    Dim nW As Long
    nW = 0

    ' ---------- BLOC 1: cirurgies puntuals (verificades) ----------
    Dim bad As Long
    bad = 0
    bad = bad + Srg(db, Apply, nW, "DW-S01-EA07", "Facade_Flank", "1", "0", "4.1 B: els murs no son de cambra")
    bad = bad + SrgN(db, Apply, nW, "DW-S01-EA07", "N_Built_Walls", "0", "4.1 B")
    bad = bad + Srg(db, Apply, nW, "DW-S01-EA07", "Sys_Eave", "Absent", "Present complete", "4.9: cos basal coronat per superficie en voladis sobre biga [DEFAULT complete]")
    bad = bad + Srg(db, Apply, nW, "DW-S01-EA07", "Eave_Surface", "0", "1", "4.9: la superficie en voladis ES l'element T")
    bad = bad + Srg(db, Apply, nW, "DW-S01-EA16", "Sys_Chamber", "Absent", "Not applicable", "4.2: el ninxol natural fa d'espai; murs de tancament legitims (R66 esmenada)")
    bad = bad + Srg(db, Apply, nW, "DW-S01-EA16", "Base_Level", "0", "1", "4.17 B1: basament d'anivellacio de ninxol / suport de mensula")
    bad = bad + Srg(db, Apply, nW, "DW-S01-EA16", "Sys_Base", "Absent", "Present complete", "4.17 [DEFAULT complete]")
    bad = bad + Srg(db, Apply, nW, "DW-S01-EA71", "Facade_Flank", "2", "0", "4.3 B")
    bad = bad + SrgN(db, Apply, nW, "DW-S01-EA71", "N_Built_Walls", "0", "4.3 B")
    bad = bad + Srg(db, Apply, nW, "DW-S01-EA72", "Facade_Flank", "1", "0", "4.4 B")
    bad = bad + SrgN(db, Apply, nW, "DW-S01-EA72", "N_Built_Walls", "0", "4.4 B")
    bad = bad + Srg(db, Apply, nW, "DW-S01-EA72", "N_Chamber_Bodies", "1", "0", "4.4/4.6 B: sense cambra ni portal")
    bad = bad + Srg(db, Apply, nW, "DW-S01-EA72", "Sys_Platform", "Absent", "Present partial", "4.6 [DEFAULT]: bigues F conservades com a sol d'un nivell superior (porxo/circulacio)")
    bad = bad + Srg(db, Apply, nW, "DW-S04-EA01", "Facade_Flank", "2", "0", "4.5 B")
    bad = bad + SrgN(db, Apply, nW, "DW-S04-EA01", "N_Built_Walls", "0", "4.5 B")
    bad = bad + Srg(db, Apply, nW, "DW-S04-EA01", "N_Chamber_Bodies", "1", "0", "4.7 B")
    bad = bad + Srg(db, Apply, nW, "DW-S03-EA03", "Structural_Pilasters", "2", "0", "4.8: el pilaret exempt de la boca no es element K; passa a T_ARCH_FEATURES")
    bad = bad + Srg(db, Apply, nW, "DW-S01-EA18", "Sys_Interface", "Absent", "Present", "4.10: correccio firmada")
    bad = bad + Srg(db, Apply, nW, "DW-S01-EA62", "Sys_Portal", "Absent", "Attested lost", "4.11 A")
    bad = bad + Srg(db, Apply, nW, "DW-S01-EA74", "Sys_Portal", "Absent", "Attested lost", "4.12 A")
    bad = bad + Srg(db, Apply, nW, "DW-S02-EA04", "Sys_Portal", "Absent", "Attested lost", "4.13 A")
    bad = bad + Srg(db, Apply, nW, "DW-S01-EA73", "Upper_Crown", "0", "9", "4.14: base muraria apenes conservada")
    bad = bad + Srg(db, Apply, nW, "DW-S01-EA73", "Tie_Walls", "0", "9", "4.15")
    bad = bad + Srg(db, Apply, nW, "DW-S01-EA04", "N_Basal_Bodies", "1", "0", "4.16: N1 arranca del terreny; el recompte era l'error")
    bad = bad + Srg(db, Apply, nW, "DW-S01-EA06", "Upper_Crown_Format", "", "Projecting course", "4.18")
    bad = bad + Srg(db, Apply, nW, "DW-S01-EA03", "Jamb_Fabric_Reveal", "0", "9", "4.20 [DEFAULT]: muntant ocult no observable, JFR indeterminable")
    bad = bad + Srg(db, Apply, nW, "DW-S04-EA18", "Jamb_Fabric_Reveal", "0", "9", "4.21 [DEFAULT]: idem")
    If bad > 0 Then
        Debug.Print "AVORTAT: " & bad & " cirurgies no quadren amb l'estat actual. Cap escriptura."
        Exit Sub
    End If

    ' ---------- BLOC 2: cascada (nomes sobre NULL) ----------
    Dim i As Integer, c As Variant
    ' Naturals amb interior no observable -> 9
    For Each c In Split("DW-S01-EA43;DW-S01-EA44;DW-S01-EA49;DW-S03-EA02;DW-S04-EA21;DW-S01-EA75;DW-S05-EA02", ";")
        FillNulls db, Apply, nW, CStr(c), ELEMS & ";" & BIOF & ";" & MATF & ";" & DANF & ";" & MASF & ";Recessed_Frame;Human_Remains", "9", "P0 interior no observable"
    Next c
    ' Naturals observables -> 0 (el pare descriu el context, no els fills)
    For Each c In Split("DW-S01-EA42;DW-S01-EA76;DW-S06-EA04", ";")
        FillNulls db, Apply, nW, CStr(c), ELEMS & ";" & BIOF & ";" & MATF & ";" & DANF & ";" & MASF & ";Recessed_Frame;Human_Remains", "0", "P0 context observat sense contingut propi"
    Next c
    ' EA78/EA79 nous [DEFAULT perfil 0]
    For Each c In Split("DW-S01-EA78;DW-S01-EA79", ";")
        FillNulls db, Apply, nW, CStr(c), ELEMS & ";" & BIOF & ";" & MATF & ";" & DANF & ";" & MASF & ";Recessed_Frame;Human_Remains;Support_Modified", "0", "[DEFAULT] perfil de ninxol observat; digues si cal 9"
    Next c
    ' EA11: Mat a 0 (P2), resta del bloc a 9 (P5)
    FillNulls db, Apply, nW, "DW-S01-EA11", MATF, "0", "P2"
    FillNulls db, Apply, nW, "DW-S01-EA11", ELEMS & ";" & BIOF & ";" & MASF & ";Recessed_Frame;Human_Remains;Support_Modified", "9", "P5 colapsada, nomes pintura perimetral"
    ' EA77
    FillNulls db, Apply, nW, "DW-S01-EA77", "Masonry_Present", "1", "P8"
    FillNulls db, Apply, nW, "DW-S01-EA77", ELEMS, "9", "P8 cap element llegible"
    FillNulls db, Apply, nW, "DW-S01-EA77", MATF & ";" & DANF & ";Human_Remains;Recessed_Frame", "0", "P8"
    FillNulls db, Apply, nW, "DW-S01-EA77", "Mortar_Present;Chinking_Stones", "9", "[DEFAULT] deteriorament extrem: no avaluable"
    ' Restes humanes explicites
    For Each c In Split("DW-S01-EA05;DW-S01-EA32;DW-S01-EA33;DW-S01-EA37;DW-S01-EA55;DW-S01-EA60;DW-S01-EA61", ";")
        FillNulls db, Apply, nW, CStr(c), "Human_Remains", "0", "P1"
    Next c
    For Each c In Split("DW-S01-EA19;DW-S01-EA50;DW-S06-EA01", ";")
        FillNulls db, Apply, nW, CStr(c), "Human_Remains", "9", "P1 no observable"
    Next c
    ' Cascada bioarq: HR=0 -> 0 ; HR=9 -> 9 (tot el corpus, nomes NULL)
    CascadeBio db, Apply, nW
    ' Materials
    For Each c In Split("Mat_Textiles;Mat_Wood;Mat_VegFiber;Mat_Fauna", ";")
        FillNulls db, Apply, nW, "DW-S01-EA17", CStr(c), "1", "P2 presencia (matisacio a Materials_Notes)"
    Next c
    FillNulls db, Apply, nW, "DW-S01-EA17", "Mat_Ceramics;Mat_DeerAntler;Mat_Other", "0", "P2"
    For Each c In Split("DW-S01-EA07;DW-S01-EA08;DW-S01-EA50;DW-S04-EA01;DW-S06-EA01", ";")
        FillNulls db, Apply, nW, CStr(c), MATF, "9", "P2 no observables"
    Next c
    FillAllNull db, Apply, nW, MATF, "0", "P2 resta: absencia verificada"
    ' Danys
    For Each c In Split("DW-S01-EA23;DW-S01-EA26;DW-S01-EA52", ";")
        FillNulls db, Apply, nW, CStr(c), DANF, "9", "P3"
    Next c
    FillNulls db, Apply, nW, "DW-S01-EA73", "Modern_Access", "1", "P3"
    FillNulls db, Apply, nW, "DW-S04-EA19", "Modern_Access", "1", "P3"
    For Each c In Split("DW-S01-EA60;DW-S01-EA73;DW-S04-EA11;DW-S04-EA19", ";")
        FillNulls db, Apply, nW, CStr(c), DANF, "0", "P3"
    Next c
    ' Maconeria: declaracions v21 i valors A/B
    For Each c In Split("DW-S01-EA25;DW-S01-EA26;DW-S04-EA11;DW-S04-EA13;DW-S04-EA17", ";")
        FillNulls db, Apply, nW, CStr(c), "Masonry_Present", "0", "B fabrica=0"
    Next c
    FillNulls db, Apply, nW, "DW-S01-EA16", "Masonry_Present", "1", "B"
    FillNulls db, Apply, nW, "DW-S01-EA16", "Mortar_Present", "9", "B"
    FillNulls db, Apply, nW, "DW-S01-EA16", "Chinking_Stones", "1", "B"
    FillNulls db, Apply, nW, "DW-S01-EA12", "Mortar_Present", "1", "A"
    Dim ch As Variant
    For Each ch In Array("DW-S01-EA03|1", "DW-S01-EA09|0", "DW-S01-EA14|1", "DW-S01-EA21|0", "DW-S01-EA51|1", "DW-S01-EA53|0", "DW-S01-EA59|9", "DW-S03-EA03|1", "DW-S05-EA01|1", "DW-S06-EA02|1")
        FillNulls db, Apply, nW, Split(ch, "|")(0), "Chinking_Stones", Split(ch, "|")(1), "A falques"
    Next ch
    FillNulls db, Apply, nW, "DW-S01-EA10", "Dec_Present", "0", "P4"
    FillNulls db, Apply, nW, "DW-S01-EA39", "Dec_Present", "0", "P4"
    FillNulls db, Apply, nW, "DW-S01-EA71", "RockArt_Present", "0", "P4"
    FillNulls db, Apply, nW, "DW-S04-EA03", "RockArt_Present", "0", "P4"
    ' Cascada maconeria: Masonry_Present=0 -> Mortar/Chinking 0
    CascadeMas db, Apply, nW
    ' Marc rebaixat
    For Each c In Split("DW-S02-EA02;DW-S04-EA03;DW-S04-EA04;DW-S04-EA05;DW-S06-EA02", ";")
        FillNulls db, Apply, nW, CStr(c), "Recessed_Frame", "0", "P7"
    Next c
    FillNulls db, Apply, nW, "DW-S04-EA15", "Recessed_Frame", "9", "P7 marc no conservat"
    FillAllNull db, Apply, nW, "Recessed_Frame", "0", "P7 sense portal: inaplicable"
    ' Plataformes atestades per mensules (P6)
    For Each c In Split("DW-S01-EA38;DW-S01-EA39;DW-S01-EA40;DW-S04-EA03", ";")
        FillNulls db, Apply, nW, CStr(c), "Platform_Surface", "3", "P6 mensules atesten la superficie perduda"
        RaiseIfAbsent db, Apply, nW, CStr(c), "Sys_Platform", "Attested lost", "P6 [DEFAULT] coherencia amb Z=3"
    Next c
    FillNulls db, Apply, nW, "DW-S01-EA71", "Upper_Crown", "0", "P6"
    FillNulls db, Apply, nW, "DW-S01-EA72", "Upper_Crown", "0", "P6"
    ' Suport modificat: pendents a 9 + conversio 0->9 en construides
    FillAllNull db, Apply, nW, "Support_Modified", "9", "P-SuportMod: pendents a 9"
    For Each c In Split("DW-S01-EA37;DW-S06-EA01;DW-S01-EA14;DW-S01-EA51;DW-S01-EA73;DW-S01-EA77;DW-S01-EA74;DW-S02-EA04;DW-S01-EA19", ";")
        Flip db, Apply, nW, CStr(c), "Support_Modified", "0", "9", "P-SuportMod: el dubte sempre hi es en construides"
    Next c
    ' Residus: qualsevol NULL restant de la bateria queda llistat
    Debug.Print "  --- residus NULL restants (si cap, perfecte) ---"
    ReportNulls db

    If Not Apply Then
        Debug.Print "INFORME: " & nW & " escriptures planificades. Revisa els [DEFAULT] i executa TancamentV25g True."
        Exit Sub
    End If

    ' ---------- BLOC 3: ARCH_FEATURES + arestes noves + notes ----------
    Dim sid As Long
    sid = Nz(DLookup("ID", "T_STRUCTURES", "Code='DW-S03-EA03'"), 0)
    If sid > 0 And DCount("*", "T_ARCH_FEATURES", "ID_Structure=" & sid & " AND Feature_Code='Freestanding pillar'") = 0 Then
        db.Execute "INSERT INTO T_ARCH_FEATURES (ID_Structure, Feature_Code, Feature_Count, Material, Notes) VALUES (" & sid & ", 'Freestanding pillar', 1, 'Stone', 'Pilaret exempt de maconeria a la boca de la cova, amb petit basament; funcio indeterminada (v25g, decisio 4.8)')", dbFailOnError
        Debug.Print "  [ARCH_FEATURES] pilaret de DW-S03-EA03 registrat"
    End If
    AddANC db, "DW-S01-EA78", "DW-S01-EA33", "Ninxol natural adossat a EA33; el context natural precedeix (v25g)"
    AddANC db, "DW-S01-EA79", "DW-S01-EA07", "Ninxol natural damunt de la confluencia EA07-EA08 (v25g)"
    AddANC db, "DW-S01-EA79", "DW-S01-EA08", "Ninxol natural damunt de la confluencia EA07-EA08 (v25g)"
    AppendNote db, "DW-S01-EA72", "Systems_Notes", "Sense cambra tancada (v25g): mur de dos pisos a l'esquerra i facana adjacent a la dreta; hipotesis obertes: porxo, estructura de pre-assecat o circulacio cap a un nivell superior (EA38?)."
    AppendNote db, "DW-S01-EA11", "Condition_Notes", "Colapsada: nomes es conserva la pintura perimetral; probable esfondrament de la repisa de suport (v25g)."
    AppendNote db, "DW-S01-EA17", "Materials_Notes", "Presencia parcial/fragmentaria de textils, fusta, fibra vegetal i fauna (v25g)."
    AppendNote db, "DW-S01-EA07", "Systems_Notes", "Cos basal coronat per rafec: superficie en voladis sobre biga transversal; el modul S+T-U es independent de la cambra (v25g, decisio 4.9)."
    Debug.Print "APLICAT: " & nW & " escriptures. Seguent: importar DB v25g, RebuildQueriesV25, i reobrir QRY_16."
End Sub

' ---------------- auxiliars privats ----------------

Private Function Srg(db As DAO.Database, ByVal Apply As Boolean, ByRef nW As Long, ByVal code As String, ByVal fld As String, ByVal oldV As String, ByVal newV As String, ByVal why As String) As Long
    Dim cur As Variant
    cur = DLookup(fld, "T_STRUCTURES", "Code='" & code & "'")
    Dim curS As String
    curS = Trim(Nz(cur, ""))
    If curS = newV Then
        Debug.Print "  [JA FET] " & code & "." & fld & " = " & newV
        Srg = 0
    ElseIf curS = oldV Then
        Debug.Print "  [CIRURGIA] " & code & "." & fld & ": '" & oldV & "' -> '" & newV & "'  (" & why & ")"
        If Apply Then WriteFld db, code, fld, newV
        nW = nW + 1
        Srg = 0
    Else
        Debug.Print "  [NO QUADRA] " & code & "." & fld & ": esperat '" & oldV & "', trobat '" & curS & "'"
        Srg = 1
    End If
End Function

Private Function SrgN(db As DAO.Database, ByVal Apply As Boolean, ByRef nW As Long, ByVal code As String, ByVal fld As String, ByVal newV As String, ByVal why As String) As Long
    Dim cur As Variant
    cur = DLookup(fld, "T_STRUCTURES", "Code='" & code & "'")
    If Nz(cur, 0) = 0 Then
        Debug.Print "  [JA FET] " & code & "." & fld & " = 0"
    Else
        Debug.Print "  [CIRURGIA] " & code & "." & fld & ": " & cur & " -> 0  (" & why & ")"
        If Apply Then WriteFld db, code, fld, newV
        nW = nW + 1
    End If
    SrgN = 0
End Function

Private Sub WriteFld(db As DAO.Database, ByVal code As String, ByVal fld As String, ByVal v As String)
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT [" & fld & "] FROM T_STRUCTURES WHERE Code='" & code & "'", dbOpenDynaset)
    If Not rs.EOF Then
        rs.Edit
        If IsNumeric(v) And Not (fld Like "Sys_*" Or fld Like "*Format" Or fld = "Jamb_Fabric") Then
            rs.Fields(fld).Value = Val(v)
        Else
            rs.Fields(fld).Value = v
        End If
        rs.Update
    End If
    rs.Close
End Sub

Private Sub FillNulls(db As DAO.Database, ByVal Apply As Boolean, ByRef nW As Long, ByVal code As String, ByVal flds As String, ByVal v As String, ByVal why As String)
    Dim f As Variant
    For Each f In Split(flds, ";")
        If FieldExists(db, CStr(f)) Then
            Dim cur As Variant
            cur = DLookup("[" & f & "]", "T_STRUCTURES", "Code='" & code & "'")
            If IsNull(cur) Then
                On Error Resume Next
                Planned.Add "x", "K" & code & "." & f
                If Err.Number = 0 Then
                    On Error GoTo 0
                    Debug.Print "  [CASCADA] " & code & "." & f & " -> " & v & "  (" & why & ")"
                    If Apply Then WriteFld db, code, CStr(f), v
                    nW = nW + 1
                Else
                    On Error GoTo 0
                End If
            End If
        End If
    Next f
End Sub

Private Sub FillAllNull(db As DAO.Database, ByVal Apply As Boolean, ByRef nW As Long, ByVal flds As String, ByVal v As String, ByVal why As String)
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT Code FROM T_STRUCTURES", dbOpenSnapshot)
    Do While Not rs.EOF
        FillNulls db, Apply, nW, CStr(rs!code), flds, v, why
        rs.MoveNext
    Loop
    rs.Close
End Sub

Private Sub CascadeBio(db As DAO.Database, ByVal Apply As Boolean, ByRef nW As Long)
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT Code, Human_Remains FROM T_STRUCTURES WHERE Human_Remains=0 OR Human_Remains=9", dbOpenSnapshot)
    Do While Not rs.EOF
        FillNulls db, Apply, nW, CStr(rs!code), BIOF, CStr(rs!Human_Remains), "cascada Human_Remains=" & rs!Human_Remains
        rs.MoveNext
    Loop
    rs.Close
End Sub

Private Sub CascadeMas(db As DAO.Database, ByVal Apply As Boolean, ByRef nW As Long)
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT Code FROM T_STRUCTURES WHERE Masonry_Present=0", dbOpenSnapshot)
    Do While Not rs.EOF
        FillNulls db, Apply, nW, CStr(rs!code), "Mortar_Present;Chinking_Stones", "0", "cascada sense fabrica"
        rs.MoveNext
    Loop
    rs.Close
End Sub

Private Sub RaiseIfAbsent(db As DAO.Database, ByVal Apply As Boolean, ByRef nW As Long, ByVal code As String, ByVal fld As String, ByVal v As String, ByVal why As String)
    If Trim(Nz(DLookup(fld, "T_STRUCTURES", "Code='" & code & "'"), "")) = "Absent" Then
        Debug.Print "  [CASCADA] " & code & "." & fld & ": Absent -> " & v & "  (" & why & ")"
        If Apply Then WriteFld db, code, fld, v
        nW = nW + 1
    End If
End Sub

Private Sub Flip(db As DAO.Database, ByVal Apply As Boolean, ByRef nW As Long, ByVal code As String, ByVal fld As String, ByVal oldV As String, ByVal newV As String, ByVal why As String)
    Dim cur As Variant
    cur = DLookup(fld, "T_STRUCTURES", "Code='" & code & "'")
    If Not IsNull(cur) Then
        If CStr(cur) = oldV Then
            Debug.Print "  [CONVERSIO] " & code & "." & fld & ": " & oldV & " -> " & newV & "  (" & why & ")"
            If Apply Then WriteFld db, code, fld, newV
            nW = nW + 1
        End If
    End If
End Sub

Private Sub AddANC(db As DAO.Database, ByVal natC As String, ByVal strC As String, ByVal note As String)
    Dim a As Long, b As Long, t As Long
    a = Nz(DLookup("ID", "T_STRUCTURES", "Code='" & natC & "'"), 0)
    b = Nz(DLookup("ID", "T_STRUCTURES", "Code='" & strC & "'"), 0)
    If a = 0 Or b = 0 Then
        Debug.Print "  [ARESTA OMESA] codi inexistent: " & natC & " / " & strC
        Exit Sub
    End If
    Dim ea As Long
    ea = a
    If a > b Then
        t = a: a = b: b = t
    End If
    If DCount("*", "T_CONNECTIONS", "(ID_Struct_A=" & a & " AND ID_Struct_B=" & b & ") OR (ID_Struct_A=" & b & " AND ID_Struct_B=" & a & ")") = 0 Then
        db.Execute "INSERT INTO T_CONNECTIONS (ID_Struct_A, ID_Struct_B, Connection_Type, Chrono_Relation, Confidence, ID_Earlier, Notes) VALUES (" & a & ", " & b & ", 'Associated natural context', 'Sequential', 'High', " & ea & ", '" & note & "')", dbFailOnError
        Debug.Print "  [ARESTA] " & natC & " <-> " & strC & " (Associated natural context, earlier=" & natC & ")"
    Else
        Debug.Print "  [ARESTA JA EXISTENT] " & natC & " <-> " & strC
    End If
End Sub

Private Sub AppendNote(db As DAO.Database, ByVal code As String, ByVal fld As String, ByVal txt As String)
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT [" & fld & "] FROM T_STRUCTURES WHERE Code='" & code & "'", dbOpenDynaset)
    If Not rs.EOF Then
        Dim cur As String
        cur = Nz(rs.Fields(fld).Value, "")
        If InStr(cur, Left(txt, 40)) = 0 Then
            rs.Edit
            If Trim(cur) = "" Then
                rs.Fields(fld).Value = txt
            Else
                rs.Fields(fld).Value = cur & " | " & txt
            End If
            rs.Update
        End If
    End If
    rs.Close
End Sub

Private Function FieldExists(db As DAO.Database, ByVal fld As String) As Boolean
    On Error GoTo no
    Dim t As Integer
    t = db.TableDefs("T_STRUCTURES").Fields(fld).Type
    FieldExists = True
    Exit Function
no:
    FieldExists = False
End Function

Private Sub ReportNulls(db As DAO.Database)
    Dim allf As String
    allf = ELEMS & ";" & BIOF & ";" & MATF & ";" & DANF & ";" & MASF & ";Recessed_Frame;Human_Remains;Support_Modified"
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT * FROM T_STRUCTURES", dbOpenSnapshot)
    Dim f As Variant, n As Long
    Do While Not rs.EOF
        For Each f In Split(allf, ";")
            If FieldExists(db, CStr(f)) Then
                If IsNull(rs.Fields(CStr(f)).Value) Then
                    Debug.Print "    [RESIDU] " & rs!code & "." & f
                    n = n + 1
                End If
            End If
        Next f
        rs.MoveNext
    Loop
    rs.Close
    If n = 0 Then Debug.Print "    (cap residu previst despres de l'aplicacio)"
End Sub
