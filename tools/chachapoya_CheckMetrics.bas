Option Compare Database
Option Explicit

' ================================================================
'  chachapoya_CheckMetrics_v25.bas
'
'  METRICS <-> DATABASE COHERENCE PIPELINE (Fase 2)
'  Author: Esteve Ribera Torro | TFM Arqueologia UA
'  Spec: DELTA_v24_v25.md (bloc A)
'
'  Two public entry points (run from the Immediate window):
'
'    CheckMetrics ["...\metriques_shapes_audit_vN.csv"]
'      Reads the FOUR metric CSVs of one extractor pass (pick the
'      shapes_audit file; the siblings are derived by name) and
'      crosses them against T_STRUCTURES. Every discrepancy
'      becomes ONE ROW in T_METRIC_REVIEW with Decision='Pending'.
'      NOTHING IS WRITTEN to T_STRUCTURES here: the QC advises,
'      the researcher signs.
'
'      Idempotent: each run clears its own Pending rows and
'      repopulates; SIGNED rows (Accept / Keep DB / Investigate)
'      are never touched and never duplicated - a discrepancy
'      already signed with the same proposal is skipped.
'
'    ApplyMetricReview
'      Executes the rows signed 'Accept' (and only those), writes
'      Proposed_Value into Field_Name, stamps Applied_On and
'      prints the acta. 'Keep DB' and 'Investigate' are never
'      applied. Re-running is harmless: applied rows are skipped.
'
'  DIRECTIONAL PRIMACY (closed in conversation, delta bloc A):
'    - A shape is POSITIVE evidence: it can raise a presence or a
'      count, never lower one. Fewer shapes than the field says
'      is an acta note, not a discrepancy.
'    - Counts from shapes are a FLOOR (undrawn brackets exist).
'    - Metric values are instrumental: they fill NULL fields via
'      review; a future conflict against a hand value is a
'      Pending row like any other.
'    - No write without a signature.
'
'  ELEMENT MAP: token -> A-Z letter is fixed here (convention
'  v1.1, section 2.1); letter -> field is read live from
'  L_ELEMENTS.Field_Name, so a vocabulary delta never desyncs
'  this module. 'wall' splits by plane: f -> L (Facade_Flank),
'  r/l -> V (Return_Wall); b has no field in v24+ (absorbed by
'  Sys_Chamber) and is skipped.
'
'  CHECKS:
'    C1   Timber_Bracket_Count vs distinct 'bra' instances
'    C2   element presence fields vs existence of element shapes
'    SLOT 'slot' shapes vs Support_Modified
'    C3a  body-code grammar (X.Y with Y >= 1)
'    C3b  bodies cited by shapes vs N_Basal/N_Chamber_Bodies
'    C5   integration of NULL metric fields (Area_m2, opening
'         dims from -c/-r portals, Length/Width/Height_m from
'         walls, Support_Depth_m from 'led', Dim_Method)
'    C7   N_Built_Walls vs distinct wall planes drawn
'
'  VBA constraints: ASCII, CRLF, DAO, no line continuations,
'  Private helpers, FKs by name.
' ================================================================

Private Const DEC_PENDING As String = "Pending"
Private Const DEC_ACCEPT As String = "Accept"

Private mBatch As String

Public Sub CheckMetrics(Optional auditPath As String = "")
    Dim db As DAO.Database
    Set db = CurrentDb()

    If Not TableOK(db, "T_METRIC_REVIEW") Then
        Debug.Print "[Check] *** T_METRIC_REVIEW no existeix: aplica PATCH v25 primer"
        Exit Sub
    End If

    If auditPath = "" Then auditPath = PickFile()
    If auditPath = "" Then Exit Sub
    If InStr(auditPath, "shapes_audit") = 0 Then
        Debug.Print "[Check] *** tria el fitxer metriques_shapes_audit (els germans es deriven del nom)"
        Exit Sub
    End If
    Dim mursPath As String, estPath As String, portPath As String
    mursPath = Replace(auditPath, "shapes_audit", "murs")
    estPath = Replace(auditPath, "shapes_audit", "estructures")
    portPath = Replace(auditPath, "shapes_audit", "portals")
    If Len(Dir(auditPath)) = 0 Or Len(Dir(mursPath)) = 0 Or Len(Dir(estPath)) = 0 Then
        Debug.Print "[Check] *** falta algun dels CSV germans (audit/murs/estructures)"
        Exit Sub
    End If

    mBatch = Mid(auditPath, InStrRev(auditPath, "\") + 1)

    Debug.Print "================================================================"
    Debug.Print "[Check] lot: " & mBatch
    Debug.Print "[Check] mode: detectar i proposar; cap escriptura a T_STRUCTURES"

    ' ---- maps from the database ----
    Dim codeToID As New Collection
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT ID, Code FROM T_STRUCTURES", dbOpenSnapshot)
    Do While Not rs.EOF
        KAdd codeToID, CStr(rs!ID), Nz(rs!Code, "")
        rs.MoveNext
    Loop
    rs.Close

    Dim letterToField As New Collection
    Set rs = db.OpenRecordset("SELECT Code, Field_Name FROM L_ELEMENTS WHERE Field_Name Is Not Null AND Field_Name <> ''", dbOpenSnapshot)
    Do While Not rs.EOF
        KAdd letterToField, Nz(rs!Field_Name, ""), Nz(rs!Code, "")
        rs.MoveNext
    Loop
    rs.Close

    ' token -> letter (convention v1.1, 2.1). 'wall' handled apart.
    Dim tokLetter As New Collection
    KAdd tokLetter, "E", "bra"
    KAdd tokLetter, "F", "tbeam"
    KAdd tokLetter, "S", "ebeam"
    KAdd tokLetter, "A", "bbeam"
    KAdd tokLetter, "I", "corn"
    KAdd tokLetter, "R", "crown"
    KAdd tokLetter, "N", "sill"
    KAdd tokLetter, "O", "jamb"
    KAdd tokLetter, "Q", "lint"
    KAdd tokLetter, "K", "pil"
    KAdd tokLetter, "D", "tie"
    KAdd tokLetter, "X", "roof"
    KAdd tokLetter, "Z", "plat"
    KAdd tokLetter, "P", "port"
    KAdd tokLetter, "U", "eave"

    ' signatures of already-signed rows: never resurrect them
    Dim signedKeys As New Collection
    Set rs = db.OpenRecordset("SELECT ID_Structure, Check_Code, Field_Name, Proposed_Value FROM T_METRIC_REVIEW WHERE Decision Is Not Null AND Decision <> '" & DEC_PENDING & "'", dbOpenSnapshot)
    Do While Not rs.EOF
        KAdd signedKeys, "1", SigKey(Nz(rs!ID_Structure, 0), Nz(rs!Check_Code, ""), Nz(rs!Field_Name, ""), Nz(rs!Proposed_Value, ""))
        rs.MoveNext
    Loop
    rs.Close

    ' idempotency: this run owns the Pending set
    db.Execute "DELETE FROM T_METRIC_REVIEW WHERE Decision='" & DEC_PENDING & "'", dbFailOnError
    Debug.Print "[Check] files Pending anteriors esborrades (les firmades es conserven)"

    ' ---- load the audit CSV ----
    Dim hdr() As String, rows As Collection
    Set rows = LoadCsv(auditPath, hdr)
    If rows Is Nothing Then Exit Sub
    Dim cCode As Long, cLabel As Long, cValid As Long, cGeom As Long
    Dim cPlane As Long, cElem As Long, cBody As Long, cLetter As Long
    Dim cBody2 As Long, cStatus As Long, cAReal As Long, cWid As Long
    Dim cHei As Long, cMetric As Long, cMVal As Long
    cCode = HIdx(hdr, "Code"): cLabel = HIdx(hdr, "Label"): cValid = HIdx(hdr, "Valid")
    cGeom = HIdx(hdr, "Geometry"): cPlane = HIdx(hdr, "Plane"): cElem = HIdx(hdr, "Element")
    cBody = HIdx(hdr, "Body"): cLetter = HIdx(hdr, "Letter"): cBody2 = HIdx(hdr, "Body2")
    cStatus = HIdx(hdr, "Status"): cAReal = HIdx(hdr, "Area_Real_m2")
    cWid = HIdx(hdr, "Width_m"): cHei = HIdx(hdr, "Height_m")
    cMetric = HIdx(hdr, "Metric"): cMVal = HIdx(hdr, "Metric_Value_m")
    If cCode < 0 Or cLabel < 0 Or cElem < 0 Then
        Debug.Print "[Check] *** capcalera de l'audit no reconeguda"
        Exit Sub
    End If

    ' ---- load the murs CSV ----
    Dim mhdr() As String, mrows As Collection
    Set mrows = LoadCsv(mursPath, mhdr)
    Dim mCode As Long, mWall As Long, mBody As Long, mLL As Long, mHH As Long
    Dim mWC As Long, mHC As Long
    mCode = HIdx(mhdr, "Code"): mWall = HIdx(mhdr, "Wall"): mBody = HIdx(mhdr, "Body")
    mLL = HIdx(mhdr, "LL_m"): mHH = HIdx(mhdr, "HH_m")
    mWC = HIdx(mhdr, "Width_Cons_m"): mHC = HIdx(mhdr, "Height_Cons_m")

    ' ---- per-structure accumulators, keyed by Code ----
    Dim codes As New Collection
    Dim braInst As New Collection      ' code -> Collection of body|letter
    Dim elemShapes As New Collection   ' code|token -> evidence labels
    Dim elemBest As New Collection     ' code|token -> best status rank (3=c,2=cons,1=r)
    Dim bodiesB As New Collection      ' code -> Collection of basal bodies
    Dim bodiesC As New Collection      ' code -> Collection of chamber bodies
    Dim wallPlanes As New Collection   ' code -> Collection of planes
    Dim intArea As New Collection      ' code -> sum t-int Area_Real
    Dim portRow As New Collection      ' code -> "w|h|status" best f-port -c/-r
    Dim ledVals As New Collection      ' code -> Collection of led dd values

    Dim i As Long
    Dim r() As String
    Dim nShapes As Long
    For i = 1 To rows.Count
        r = rows(i)
        If LCase(F(r, cValid)) = "si" Then
            nShapes = nShapes + 1
            Dim cd As String, tok As String, pl As String, bd As String, bd2 As String
            Dim st As String, geo As String
            cd = F(r, cCode): tok = F(r, cElem): pl = F(r, cPlane)
            bd = F(r, cBody): bd2 = F(r, cBody2): st = F(r, cStatus): geo = F(r, cGeom)
            KAdd codes, "1", cd

            ' C1 material
            If tok = "bra" Then SubAdd braInst, cd, bd & "|" & F(r, cLetter)

            ' C3 material
            If bd <> "" Then AddBody bodiesB, bodiesC, cd, bd
            If bd2 <> "" Then AddBody bodiesB, bodiesC, cd, bd2

            ' C3a grammar
            If bd <> "" And Not BodyOK(bd) Then AddRow db, codeToID, signedKeys, cd, "C3a", "", "", "", "cos '" & bd & "' fora de gramatica (X.Y, Y>=1) a la shape '" & F(r, cLabel) & "'"
            If bd2 <> "" And Not BodyOK(bd2) Then AddRow db, codeToID, signedKeys, cd, "C3a", "", "", "", "cos '" & bd2 & "' fora de gramatica (X.Y, Y>=1) a la shape '" & F(r, cLabel) & "'"

            ' C2 material: mapped tokens
            If tok <> "" And HasK(tokLetter, tok) Then
                SubAdd elemShapes, cd & "|" & tok, F(r, cLabel)
                RankUp elemBest, cd & "|" & tok, StRank(st)
            End If
            ' wall split by plane; bare body polygons count as wall faces
            If geo = "polygon" And (tok = "wall" Or tok = "") And bd <> "" Then
                If pl = "f" Then
                    SubAdd elemShapes, cd & "|wallF", F(r, cLabel)
                    RankUp elemBest, cd & "|wallF", StRank(st)
                End If
                If pl = "r" Or pl = "l" Then
                    SubAdd elemShapes, cd & "|wallV", F(r, cLabel)
                    RankUp elemBest, cd & "|wallV", StRank(st)
                End If
                If pl = "f" Or pl = "r" Or pl = "l" Or pl = "b" Then SubAdd wallPlanes, cd, pl
            End If
            If tok = "wall" And geo = "line" Then
                If pl = "f" Then SubAdd wallPlanes, cd, pl
                If pl = "r" Or pl = "l" Then SubAdd wallPlanes, cd, pl
            End If

            ' SLOT
            If tok = "slot" Then
                SubAdd elemShapes, cd & "|slot", F(r, cLabel)
            End If

            ' C5 material
            If tok = "int" And geo = "polygon" And pl = "t" Then
                NumAdd intArea, cd, Val(F(r, cAReal))
            End If
            If tok = "port" And geo = "polygon" And pl = "f" Then
                If st = "-c" Or st = "-r" Then
                    If Not HasK(portRow, cd) Then KAdd portRow, F(r, cWid) & "|" & F(r, cHei) & "|" & st, cd
                End If
            End If
            If tok = "led" And F(r, cMetric) = "dd" Then
                SubAdd ledVals, cd, F(r, cMVal)
            End If
        End If
    Next i
    Debug.Print "[Check] shapes valides llegides: " & nShapes & " (" & codes.Count & " estructures)"

    ' ---- now walk the corpus and emit rows ----
    Dim nC1 As Long, nC2 As Long, nSlot As Long, nC3b As Long, nC5 As Long, nC7 As Long, nInfo As Long
    Dim k As Long
    For k = 1 To codes.Count
        cd = codes(k)
        If Not HasK(codeToID, cd) Then
            Debug.Print "  [SENSE BD] " & cd & " (el CSV cita una estructura que la BD no te)"
        Else
            Dim sid As Long
            sid = CLng(codeToID(cd))

            ' ---------- C1 ----------
            Dim nSh As Long, vBD As Variant
            nSh = SubCount(braInst, cd)
            vBD = DLookup("Timber_Bracket_Count", "T_STRUCTURES", "ID=" & sid)
            If nSh > 0 Then
                If IsNull(vBD) Then
                    nC1 = nC1 + AddRow(db, codeToID, signedKeys, cd, "C1", "Timber_Bracket_Count", "", CStr(nSh), SubList(braInst, cd, "instancies bra: "))
                ElseIf nSh > CLng(vBD) Then
                    nC1 = nC1 + AddRow(db, codeToID, signedKeys, cd, "C1", "Timber_Bracket_Count", CStr(vBD), CStr(nSh), SubList(braInst, cd, "instancies bra: "))
                ElseIf nSh < CLng(vBD) Then
                    nInfo = nInfo + 1
                    Debug.Print "  [nota C1] " & cd & ": BD=" & vBD & " shapes=" & nSh & " (mensules no dibuixades: legitim)"
                End If
            End If

            ' ---------- C2 + SLOT ----------
            nC2 = nC2 + C2One(db, codeToID, signedKeys, letterToField, cd, sid, "bra", elemShapes, elemBest)
            nC2 = nC2 + C2One(db, codeToID, signedKeys, letterToField, cd, sid, "tbeam", elemShapes, elemBest)
            nC2 = nC2 + C2One(db, codeToID, signedKeys, letterToField, cd, sid, "ebeam", elemShapes, elemBest)
            nC2 = nC2 + C2One(db, codeToID, signedKeys, letterToField, cd, sid, "bbeam", elemShapes, elemBest)
            nC2 = nC2 + C2One(db, codeToID, signedKeys, letterToField, cd, sid, "corn", elemShapes, elemBest)
            nC2 = nC2 + C2One(db, codeToID, signedKeys, letterToField, cd, sid, "crown", elemShapes, elemBest)
            nC2 = nC2 + C2One(db, codeToID, signedKeys, letterToField, cd, sid, "sill", elemShapes, elemBest)
            nC2 = nC2 + C2One(db, codeToID, signedKeys, letterToField, cd, sid, "jamb", elemShapes, elemBest)
            nC2 = nC2 + C2One(db, codeToID, signedKeys, letterToField, cd, sid, "lint", elemShapes, elemBest)
            nC2 = nC2 + C2One(db, codeToID, signedKeys, letterToField, cd, sid, "pil", elemShapes, elemBest)
            nC2 = nC2 + C2One(db, codeToID, signedKeys, letterToField, cd, sid, "tie", elemShapes, elemBest)
            nC2 = nC2 + C2One(db, codeToID, signedKeys, letterToField, cd, sid, "roof", elemShapes, elemBest)
            nC2 = nC2 + C2One(db, codeToID, signedKeys, letterToField, cd, sid, "plat", elemShapes, elemBest)
            nC2 = nC2 + C2One(db, codeToID, signedKeys, letterToField, cd, sid, "port", elemShapes, elemBest)
            nC2 = nC2 + C2One(db, codeToID, signedKeys, letterToField, cd, sid, "eave", elemShapes, elemBest)
            nC2 = nC2 + C2Wall(db, codeToID, signedKeys, cd, sid, "wallF", "Facade_Flank", elemShapes, elemBest)
            nC2 = nC2 + C2Wall(db, codeToID, signedKeys, cd, sid, "wallV", "Return_Wall", elemShapes, elemBest)

            If HasK(elemShapes, cd & "|slot") Then
                Dim vSM As Variant
                vSM = DLookup("Support_Modified", "T_STRUCTURES", "ID=" & sid)
                If Nz(vSM, 0) <> 1 Then
                    nSlot = nSlot + AddRow(db, codeToID, signedKeys, cd, "SLOT", "Support_Modified", CStr(Nz(vSM, "")), "1", SubList(elemShapes, cd & "|slot", "encaixos de mensula dibuixats: "))
                End If
            End If

            ' ---------- C3b ----------
            Dim nB As Long, nC As Long
            nB = SubCount(bodiesB, cd): nC = SubCount(bodiesC, cd)
            Dim vNB As Variant, vNC As Variant
            vNB = DLookup("N_Basal_Bodies", "T_STRUCTURES", "ID=" & sid)
            vNC = DLookup("N_Chamber_Bodies", "T_STRUCTURES", "ID=" & sid)
            If nB > Nz(vNB, 0) Then
                nC3b = nC3b + AddRow(db, codeToID, signedKeys, cd, "C3b", "N_Basal_Bodies", CStr(Nz(vNB, "")), CStr(nB), SubList(bodiesB, cd, "cossos basals citats: "))
            End If
            If nC > Nz(vNC, 0) Then
                nC3b = nC3b + AddRow(db, codeToID, signedKeys, cd, "C3b", "N_Chamber_Bodies", CStr(Nz(vNC, "")), CStr(nC), SubList(bodiesC, cd, "cossos de cambra citats: "))
            End If

            ' ---------- C5 (only NULL fields get a proposal) ----------
            If HasK(intArea, cd) Then
                nC5 = nC5 + PropNull(db, codeToID, signedKeys, cd, sid, "Area_m2", Fmt2(CDbl(intArea(cd))), "suma de t-int (Area_Real_m2)")
            End If
            If HasK(portRow, cd) Then
                Dim pw() As String
                pw = Split(portRow(cd), "|")
                If Val(pw(0)) > 0 Then nC5 = nC5 + PropNull(db, codeToID, signedKeys, cd, sid, "Opening_Width_m", Fmt2(Val(pw(0))), "f-port amb estatus " & pw(2))
                If Val(pw(1)) > 0 Then nC5 = nC5 + PropNull(db, codeToID, signedKeys, cd, sid, "Opening_Height_m", Fmt2(Val(pw(1))), "f-port amb estatus " & pw(2))
            End If
            If HasK(ledVals, cd) Then
                Dim lm As Double
                lm = SubMean(ledVals, cd)
                If lm > 0 Then nC5 = nC5 + PropNull(db, codeToID, signedKeys, cd, sid, "Support_Depth_m", Fmt2(lm), "mitjana dels dd de 'led'")
            End If
            nC5 = nC5 + C5Dims(db, codeToID, signedKeys, mrows, mCode, mWall, mBody, mLL, mHH, mWC, mHC, cd, sid)

            ' ---------- C7 ----------
            Dim nPl As Long, vNW As Variant
            nPl = SubCount(wallPlanes, cd)
            vNW = DLookup("N_Built_Walls", "T_STRUCTURES", "ID=" & sid)
            If nPl > 0 Then
                If nPl > Nz(vNW, 0) Then
                    nC7 = nC7 + AddRow(db, codeToID, signedKeys, cd, "C7", "N_Built_Walls", CStr(Nz(vNW, "")), CStr(nPl), SubList(wallPlanes, cd, "plans de mur dibuixats: "))
                ElseIf nPl < Nz(vNW, 0) Then
                    nInfo = nInfo + 1
                    Debug.Print "  [nota C7] " & cd & ": BD=" & vNW & " plans dibuixats=" & nPl & " (murs no documentables: legitim)"
                End If
            End If
        End If
    Next k

    Debug.Print "----------------------------------------------------------------"
    Debug.Print "[Check] C1 recomptes de mensules:  " & nC1
    Debug.Print "[Check] C2 presencies d'element:   " & nC2
    Debug.Print "[Check] SLOT suport modificat:     " & nSlot
    Debug.Print "[Check] C3b recomptes de cossos:   " & nC3b
    Debug.Print "[Check] C5 integracio (camps NULL):" & nC5
    Debug.Print "[Check] C7 recompte de murs:       " & nC7
    Debug.Print "[Check] notes informatives (acta): " & nInfo
    Debug.Print "[Check] Revisa T_METRIC_REVIEW (o la worklist), firma Decision"
    Debug.Print "[Check] (Accept / Keep DB / Investigate) i executa ApplyMetricReview."
    Debug.Print "================================================================"
    Set db = Nothing
End Sub

Public Sub ApplyMetricReview()
    Dim db As DAO.Database
    Set db = CurrentDb()
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT * FROM T_METRIC_REVIEW WHERE Decision='" & DEC_ACCEPT & "' AND Applied_On Is Null ORDER BY ID", dbOpenDynaset)
    Debug.Print "================================================================"
    Debug.Print "[Apply] executant les files firmades 'Accept'"
    Dim nApp As Long, nSkip As Long
    Do While Not rs.EOF
        Dim fld As String, pv As String, sid As Long
        fld = Nz(rs!Field_Name, ""): pv = Nz(rs!Proposed_Value, ""): sid = rs!ID_Structure
        If fld = "" Or pv = "" Then
            nSkip = nSkip + 1
            Debug.Print "  [OMESA] fila " & rs!ID & " (" & Nz(rs!Check_Code, "") & "): sense camp o sense proposta (Investigate-only)"
        Else
            Dim cd As String
            cd = Nz(DLookup("Code", "T_STRUCTURES", "ID=" & sid), "?")
            Dim oldV As Variant
            oldV = DLookup(fld, "T_STRUCTURES", "ID=" & sid)
            Dim rs2 As DAO.Recordset
            Set rs2 = db.OpenRecordset("SELECT " & fld & " FROM T_STRUCTURES WHERE ID=" & sid, dbOpenDynaset)
            If Not rs2.EOF Then
                rs2.Edit
                If IsNumericField(db, fld) Then
                    rs2.Fields(fld).Value = CDbl(Replace(pv, ",", "."))
                Else
                    rs2.Fields(fld).Value = pv
                End If
                rs2.Update
                rs.Edit
                rs!Applied_On = Now
                rs.Update
                nApp = nApp + 1
                Debug.Print "  [APLICAT] " & cd & "." & fld & ": " & Nz(oldV, "NULL") & " -> " & pv & "  (" & Nz(rs!Check_Code, "") & ")"
            End If
            rs2.Close
        End If
        rs.MoveNext
    Loop
    rs.Close
    Debug.Print "----------------------------------------------------------------"
    Debug.Print "[Apply] files aplicades: " & nApp
    If nSkip > 0 Then Debug.Print "[Apply] omeses (sense camp/proposta): " & nSkip
    Debug.Print "[Apply] Guarda aquest text com a acta de la passada."
    Debug.Print "================================================================"
    Set db = Nothing
End Sub

' ================================================================
'  C2 helpers
' ================================================================

Private Function C2One(db As DAO.Database, codeToID As Collection, signedKeys As Collection, letterToField As Collection, cd As String, sid As Long, tok As String, elemShapes As Collection, elemBest As Collection) As Long
    If Not HasK(elemShapes, cd & "|" & tok) Then Exit Function
    Dim letter As String, fld As String
    letter = TokLetterOf(tok)
    If letter = "" Then Exit Function
    If Not HasK(letterToField, letter) Then Exit Function
    fld = letterToField(letter)
    C2One = C2Field(db, codeToID, signedKeys, cd, sid, tok, fld, elemShapes, elemBest)
End Function

Private Function C2Wall(db As DAO.Database, codeToID As Collection, signedKeys As Collection, cd As String, sid As Long, key As String, fld As String, elemShapes As Collection, elemBest As Collection) As Long
    If Not HasK(elemShapes, cd & "|" & key) Then Exit Function
    C2Wall = C2Field(db, codeToID, signedKeys, cd, sid, key, fld, elemShapes, elemBest)
End Function

Private Function C2Field(db As DAO.Database, codeToID As Collection, signedKeys As Collection, cd As String, sid As Long, tok As String, fld As String, elemShapes As Collection, elemBest As Collection) As Long
    Dim v As Variant
    v = DLookup(fld, "T_STRUCTURES", "ID=" & sid)
    Dim rank As Long
    rank = CLng(elemBest(cd & "|" & tok))
    Dim isTxt As Boolean
    isTxt = Not IsNumericField(db, fld)
    Dim ok As Boolean, prop As String
    If isTxt Then
        ok = (Left(Nz(v, ""), 7) = "Present") Or (Nz(v, "") = "Attested lost")
        If rank = 3 Then prop = "Present complete" Else prop = "Present partial"
        If rank = 1 Then prop = "Attested lost"
    Else
        ok = (Nz(v, -1) = 1) Or (Nz(v, -1) = 2) Or (Nz(v, -1) = 3)
        If rank = 3 Then prop = "1" Else prop = "2"
        If rank = 1 Then prop = "3"
    End If
    If Not ok Then
        C2Field = AddRow(db, codeToID, signedKeys, cd, "C2", fld, CStr(Nz(v, "")), prop, SubList(elemShapes, cd & "|" & tok, "shapes de l'element: "))
    End If
End Function

Private Function TokLetterOf(tok As String) As String
    Select Case tok
        Case "bra": TokLetterOf = "E"
        Case "tbeam": TokLetterOf = "F"
        Case "ebeam": TokLetterOf = "S"
        Case "bbeam": TokLetterOf = "A"
        Case "corn": TokLetterOf = "I"
        Case "crown": TokLetterOf = "R"
        Case "sill": TokLetterOf = "N"
        Case "jamb": TokLetterOf = "O"
        Case "lint": TokLetterOf = "Q"
        Case "pil": TokLetterOf = "K"
        Case "tie": TokLetterOf = "D"
        Case "roof": TokLetterOf = "X"
        Case "plat": TokLetterOf = "Z"
        Case "port": TokLetterOf = "P"
        Case "eave": TokLetterOf = "U"
    End Select
End Function

' ================================================================
'  C5 dimension proposals from the murs CSV
'  Length_m  <- max over facade rows of (LL_m else Width_Cons_m)
'  Height_m  <- sum over facade BODIES of best height per body
'  Width_m   <- max over r/l rows of (LL_m else Width_Cons_m)
'  Plus Dim_Method='Photogrammetry' when any of the three lands.
'  Proposals only where the DB field is NULL: the R->BD
'  conversation of the v24 delta stays sovereign over semantics;
'  each of these rows is signed one by one.
' ================================================================

Private Function C5Dims(db As DAO.Database, codeToID As Collection, signedKeys As Collection, mrows As Collection, mCode As Long, mWall As Long, mBody As Long, mLL As Long, mHH As Long, mWC As Long, mHC As Long, cd As String, sid As Long) As Long
    If mrows Is Nothing Then Exit Function
    Dim i As Long, r() As String
    Dim lenMax As Double, widMax As Double
    Dim hBodies As New Collection
    For i = 1 To mrows.Count
        r = mrows(i)
        If F(r, mCode) = cd Then
            Dim w As String
            w = F(r, mWall)
            Dim d1 As Double
            If w = "f" Then
                d1 = Val(F(r, mLL))
                If d1 = 0 Then d1 = Val(F(r, mWC))
                If d1 > lenMax Then lenMax = d1
                Dim hb As Double
                hb = Val(F(r, mHH))
                If hb = 0 Then hb = Val(F(r, mHC))
                RankUpD hBodies, F(r, mBody), hb
            End If
            If w = "r" Or w = "l" Then
                d1 = Val(F(r, mLL))
                If d1 = 0 Then d1 = Val(F(r, mWC))
                If d1 > widMax Then widMax = d1
            End If
        End If
    Next i
    Dim hSum As Double, j As Long
    For j = 1 To hBodies.Count
        hSum = hSum + CDbl(hBodies(j))
    Next j
    Dim n As Long
    If lenMax > 0 Then n = n + PropNull(db, codeToID, signedKeys, cd, sid, "Length_m", Fmt2(lenMax), "extensio maxima del mur de facana (murs CSV)")
    If hSum > 0 Then n = n + PropNull(db, codeToID, signedKeys, cd, sid, "Height_m", Fmt2(hSum), "suma d'alcades per cos del mur de facana (murs CSV)")
    If widMax > 0 Then n = n + PropNull(db, codeToID, signedKeys, cd, sid, "Width_m", Fmt2(widMax), "extensio maxima dels murs r/l (murs CSV)")
    If n > 0 Then n = n + PropNull(db, codeToID, signedKeys, cd, sid, "Dim_Method", "Photogrammetry", "dimensions proposades des del corpus de shapes")
    C5Dims = n
End Function

' ================================================================
'  Row machinery
' ================================================================

Private Function PropNull(db As DAO.Database, codeToID As Collection, signedKeys As Collection, cd As String, sid As Long, fld As String, pv As String, ev As String) As Long
    Dim v As Variant
    v = DLookup(fld, "T_STRUCTURES", "ID=" & sid)
    If IsNull(v) Or Trim(Nz(v, "")) = "" Then
        PropNull = AddRow(db, codeToID, signedKeys, cd, "C5", fld, "", pv, ev)
    ElseIf CStr(v) <> pv Then
        Debug.Print "  [nota C5] " & cd & "." & fld & ": BD=" & v & " CSV=" & pv & " (camp ja poblat: conflicte a firmar si cal, no proposat d'ofici)"
    End If
End Function

Private Function AddRow(db As DAO.Database, codeToID As Collection, signedKeys As Collection, cd As String, chk As String, fld As String, dbv As String, pv As String, ev As String) As Long
    If Not HasK(codeToID, cd) Then Exit Function
    Dim sid As Long
    sid = CLng(codeToID(cd))
    If HasK(signedKeys, SigKey(sid, chk, fld, pv)) Then
        Debug.Print "  [ja firmada] " & cd & " " & chk & " " & fld & " -> " & pv
        Exit Function
    End If
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("T_METRIC_REVIEW", dbOpenDynaset)
    rs.AddNew
    rs!ID_Structure = sid
    rs!Check_Code = chk
    If fld <> "" Then rs!Field_Name = fld
    If dbv <> "" Then rs!DB_Value = dbv
    If pv <> "" Then rs!Proposed_Value = pv
    rs!Evidence = ev
    rs!CSV_Batch = mBatch
    rs!Detected_On = Now
    rs!Decision = DEC_PENDING
    rs.Update
    rs.Close
    AddRow = 1
    Debug.Print "  [" & chk & "] " & cd & IIf(fld <> "", "." & fld, "") & ": BD='" & dbv & "' -> proposta '" & pv & "'"
End Function

Private Function SigKey(sid As Long, chk As String, fld As String, pv As String) As String
    SigKey = sid & "#" & chk & "#" & fld & "#" & pv
End Function

' ================================================================
'  CSV loading (quote-aware, header-indexed)
' ================================================================

Private Function LoadCsv(path As String, ByRef hdr() As String) As Collection
    Dim fnum As Integer
    fnum = FreeFile
    On Error GoTo bad
    Open path For Input As #fnum
    On Error GoTo 0
    Dim s As String
    Line Input #fnum, s
    If Left(s, 3) = Chr(239) & Chr(187) & Chr(191) Then s = Mid(s, 4)
    hdr = SplitCsv(s)
    Dim out As New Collection
    Do While Not EOF(fnum)
        Line Input #fnum, s
        If Len(Trim(s)) > 0 Then out.Add SplitCsv(s)
    Loop
    Close #fnum
    Set LoadCsv = out
    Exit Function
bad:
    Debug.Print "[Check] *** no puc obrir: " & path
    Set LoadCsv = Nothing
End Function

Private Function SplitCsv(s As String) As String()
    Dim out() As String
    ReDim out(0 To 63)
    Dim n As Long, i As Long, ch As String, cur As String
    Dim inQ As Boolean
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If inQ Then
            If ch = Chr(34) Then
                If Mid(s, i + 1, 1) = Chr(34) Then
                    cur = cur & Chr(34)
                    i = i + 1
                Else
                    inQ = False
                End If
            Else
                cur = cur & ch
            End If
        Else
            If ch = Chr(34) Then
                inQ = True
            ElseIf ch = "," Then
                out(n) = cur
                n = n + 1
                cur = ""
            Else
                cur = cur & ch
            End If
        End If
    Next i
    out(n) = cur
    ReDim Preserve out(0 To n)
    SplitCsv = out
End Function

Private Function HIdx(hdr() As String, name As String) As Long
    Dim i As Long
    HIdx = -1
    For i = 0 To UBound(hdr)
        If Trim(hdr(i)) = name Then HIdx = i: Exit Function
    Next i
End Function

Private Function F(r() As String, idx As Long) As String
    If idx < 0 Or idx > UBound(r) Then Exit Function
    F = Trim(r(idx))
End Function

' ================================================================
'  Small collection helpers (keyed sets and sub-collections)
' ================================================================

Private Sub KAdd(col As Collection, v As String, k As String)
    On Error Resume Next
    col.Add v, k
    On Error GoTo 0
End Sub

Private Function HasK(col As Collection, k As String) As Boolean
    On Error Resume Next
    Dim v As Variant
    v = col(k)
    HasK = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Sub SubAdd(col As Collection, k As String, item As String)
    Dim sub1 As Collection
    If HasK(col, k) Then
        Set sub1 = col(k)
    Else
        Set sub1 = New Collection
        col.Add sub1, k
    End If
    On Error Resume Next
    sub1.Add item, item
    On Error GoTo 0
End Sub

Private Function SubCount(col As Collection, k As String) As Long
    If Not HasK(col, k) Then Exit Function
    Dim sub1 As Collection
    Set sub1 = col(k)
    SubCount = sub1.Count
End Function

Private Function SubList(col As Collection, k As String, pre As String) As String
    If Not HasK(col, k) Then Exit Function
    Dim sub1 As Collection
    Set sub1 = col(k)
    Dim i As Long, s As String
    For i = 1 To sub1.Count
        If i > 1 Then s = s & ", "
        If i > 8 Then s = s & "...": Exit For
        s = s & sub1(i)
    Next i
    SubList = pre & s
End Function

Private Function SubMean(col As Collection, k As String) As Double
    If Not HasK(col, k) Then Exit Function
    Dim sub1 As Collection
    Set sub1 = col(k)
    Dim i As Long, t As Double, n As Long
    For i = 1 To sub1.Count
        If Val(sub1(i)) > 0 Then
            t = t + Val(sub1(i))
            n = n + 1
        End If
    Next i
    If n > 0 Then SubMean = t / n
End Function

Private Sub NumAdd(col As Collection, k As String, v As Double)
    Dim cur As Double
    If HasK(col, k) Then
        cur = CDbl(col(k))
        col.Remove k
    End If
    col.Add CStr(cur + v), k
End Sub

Private Sub RankUp(col As Collection, k As String, rank As Long)
    Dim cur As Long
    If HasK(col, k) Then
        cur = CLng(col(k))
        If rank > cur Then
            col.Remove k
            col.Add CStr(rank), k
        End If
    Else
        col.Add CStr(rank), k
    End If
End Sub

Private Sub RankUpD(col As Collection, k As String, v As Double)
    Dim cur As Double
    If HasK(col, k) Then
        cur = CDbl(col(k))
        If v > cur Then
            col.Remove k
            col.Add CStr(v), k
        End If
    Else
        col.Add CStr(v), k
    End If
End Sub

Private Sub AddBody(bodiesB As Collection, bodiesC As Collection, cd As String, bd As String)
    If Left(bd, 2) = "0." Then
        SubAdd bodiesB, cd, bd
    Else
        SubAdd bodiesC, cd, bd
    End If
End Sub

Private Function BodyOK(bd As String) As Boolean
    Dim p As Long
    p = InStr(bd, ".")
    If p < 2 Then Exit Function
    If Not IsNumeric(Left(bd, p - 1)) Then Exit Function
    If Not IsNumeric(Mid(bd, p + 1)) Then Exit Function
    BodyOK = (CLng(Mid(bd, p + 1)) >= 1)
End Function

Private Function StRank(st As String) As Long
    ' 3 = -c (complete), 2 = cons (partial), 1 = -r (restituted)
    Select Case st
        Case "-c": StRank = 3
        Case "-r": StRank = 1
        Case Else: StRank = 2
    End Select
End Function

Private Function IsNumericField(db As DAO.Database, fld As String) As Boolean
    Dim t As Long
    On Error Resume Next
    t = db.TableDefs("T_STRUCTURES").Fields(fld).Type
    On Error GoTo 0
    IsNumericField = (t = dbByte) Or (t = dbInteger) Or (t = dbLong) Or (t = dbSingle) Or (t = dbDouble) Or (t = dbCurrency) Or (t = dbDecimal)
End Function

Private Function Fmt2(v As Double) As String
    Fmt2 = Format(v, "0.00")
    Fmt2 = Replace(Fmt2, ",", ".")
End Function

Private Function TableOK(db As DAO.Database, n As String) As Boolean
    Dim t As DAO.TableDef
    For Each t In db.TableDefs
        If t.Name = n Then TableOK = True: Exit Function
    Next t
End Function

Private Function PickFile() As String
    Dim fd As Object
    On Error Resume Next
    Set fd = Application.FileDialog(3)
    On Error GoTo 0
    If fd Is Nothing Then
        Debug.Print "[Check] passa la ruta del CSV audit com a argument: CheckMetrics ""C:\...\metriques_shapes_audit_vN.csv"""
        Exit Function
    End If
    fd.Title = "Tria metriques_shapes_audit_*.csv"
    fd.AllowMultiSelect = False
    If fd.Show = -1 Then PickFile = fd.SelectedItems(1)
End Function
