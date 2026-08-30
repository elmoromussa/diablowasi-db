Attribute VB_Name = "Import_Georef_v24"
Option Compare Database
Option Explicit

' ================================================================
'  chachapoya_Import_Georef_v24.bas
'
'  Importador del CSV de coordenades i orientacions generat per
'  export_georef_orientacions.py (Metashape 2.3.0). Una fila per
'  estructura; matching per Code (mai per ID).
'
'  QUE FA:
'   - Coordenades del punt de circulacio: Coord_Lat/Lon_WGS84,
'     Coord_E/N_UTM (EPSG:32718, ja convertides a l'exportacio des
'     del mateix punt geocentric - aci NO es converteix res),
'     Altitude_masl, Coord_Precision_m.
'   - ID_Coord_Method resolt PER NOM ('Photogrammetry'), mai per ID.
'   - Facade_Azimuth_Deg (camp v24): azimut enter 0-359. Si el camp
'     no existeix (patch v24 no aplicat), s'omet amb avis i la resta
'     de la importacio continua.
'   - Facade_Orientation NO ES TOCA MAI: es judici observacional.
'     L'importador nomes COMPARA azimut vs sector declarat i llista
'     les divergencies superiors a 67.5 graus (previsualitzacio de
'     la regla 68: el detector d'antipodes).
'
'  POLITICA D'ESCRIPTURA (principi conservador):
'   - Per defecte NOMES OMPLI CAMPS NULL. Un valor existent que
'     difereix del CSV es llista com a [CONFLICTE] i no es toca.
'   - ImportGeoref ruta, True -> mode sobreescriptura: els canvis
'     s'apliquen pero CADA UN es llista com a [CANVI]. Mai en silenci.
'
'  RECONCILIACIO BIDIRECCIONAL:
'   - [CSV sense BD]: codis del CSV sense registre a T_STRUCTURES.
'   - [BD sense CSV]: registres de T_STRUCTURES dels sectors coberts
'     pel CSV que no tenen fila al CSV.
'
'  US (finestra Immediate):
'    ImportGeoref                      -> tria el fitxer, mode ompli-NULL
'    ImportGeoref "C:\ruta\fitxer.csv" -> ruta directa
'    ImportGeoref "C:\ruta\f.csv", True -> mode sobreescriptura
'
'  El resultat s'escriu a la finestra Immediate (Ctrl+G).
' ================================================================

Private Const METHOD_NAME As String = "Photogrammetry"
Private Const TOL_DEG As Double = 0.00000005   ' mig compte del 7e decimal
Private Const TOL_M As Double = 0.005          ' mig cm (CSV a 2 decimals)
Private Const TOL_AZ As Double = 0.5           ' azimut enter
Private Const R68_THRESHOLD As Double = 67.5   ' mes d'un sector adjacent

Public Sub ImportGeoref(Optional csvPath As String = "", Optional overwriteExisting As Boolean = False)
    Dim db As DAO.Database
    Set db = CurrentDb

    If Len(csvPath) = 0 Then csvPath = PickFile()
    If Len(csvPath) = 0 Then
        Debug.Print "[Import] cancel.lat: cap fitxer CSV"
        Exit Sub
    End If
    If Len(Dir(csvPath)) = 0 Then
        Debug.Print "[Import] *** el fitxer no existeix: " & csvPath
        Exit Sub
    End If

    Debug.Print "================================================================"
    Debug.Print "[Import] " & csvPath
    If overwriteExisting Then
        Debug.Print "[Import] MODE SOBREESCRIPTURA: els valors existents es canviaran i es llistaran"
    Else
        Debug.Print "[Import] mode conservador: nomes s'omplin camps NULL"
    End If

    Dim hasAzim As Boolean
    hasAzim = FieldExists(db, "T_STRUCTURES", "Facade_Azimuth_Deg")
    If Not hasAzim Then
        Debug.Print "[Import] AVIS: Facade_Azimuth_Deg no existeix (patch v24 no aplicat); l'azimut s'ometra"
    End If

    Dim methodID As Variant
    methodID = LookupMethodID(db)
    If IsNull(methodID) Then
        Debug.Print "[Import] *** L_COORD_METHOD no conte '" & METHOD_NAME & "': avortat"
        Exit Sub
    End If

    Dim csvCodes As New Collection
    Dim prefixes As New Collection
    Dim nRows As Long, nMatch As Long, nNoMatch As Long, nBad As Long, nDup As Long
    Dim nFill As Long, nConf As Long, nOver As Long, nMeth As Long, nR68 As Long

    Dim f As Integer
    f = FreeFile
    Open csvPath For Input As #f

    Dim s As String
    Line Input #f, s
    Dim p As Long
    p = InStr(s, "Code,")
    If p = 0 Then
        Close #f
        Debug.Print "[Import] *** capcalera no reconeguda (falta 'Code,'): avortat"
        Exit Sub
    End If
    s = Mid(s, p)

    Do While Not EOF(f)
        Line Input #f, s
        If Len(Trim(s)) > 0 Then
            Dim a() As String
            a = Split(s, ",", 13)
            If UBound(a) < 12 Then
                nBad = nBad + 1
                Debug.Print "  [MALFORMADA] " & Left(s, 60)
            Else
                nRows = nRows + 1
                Dim code As String
                code = Trim(a(0))
                If HasKey(csvCodes, code) Then
                    nDup = nDup + 1
                    Debug.Print "  [DUPLICAT AL CSV] " & code & " (fila ignorada)"
                Else
                    csvCodes.Add code, code
                    p = InStr(code, "-EA")
                    If p > 0 Then AddKey prefixes, Left(code, p - 1)
                    ProcessRow db, a, code, overwriteExisting, hasAzim, methodID, nMatch, nNoMatch, nFill, nConf, nOver, nMeth, nR68
                End If
            End If
        End If
    Loop
    Close #f

    ' -- reconciliacio BD -> CSV: registres dels sectors coberts sense fila --
    Dim nDbMissing As Long
    Dim rsAll As DAO.Recordset
    Set rsAll = db.OpenRecordset("SELECT Code FROM T_STRUCTURES ORDER BY Code", dbOpenSnapshot)
    Do While Not rsAll.EOF
        Dim c As String
        c = Nz(rsAll!Code, "")
        p = InStr(c, "-EA")
        If p > 0 Then
            If HasKey(prefixes, Left(c, p - 1)) And Not HasKey(csvCodes, c) Then
                nDbMissing = nDbMissing + 1
                Debug.Print "  [BD sense CSV] " & c
            End If
        End If
        rsAll.MoveNext
    Loop
    rsAll.Close

    Debug.Print "----------------------------------------------------------------"
    Debug.Print "[Import] files llegides:        " & nRows & IIf(nBad > 0, "  (" & nBad & " malformades)", "")
    Debug.Print "[Import] casades amb la BD:     " & nMatch
    Debug.Print "[Import] CSV sense BD:          " & nNoMatch
    Debug.Print "[Import] BD sense CSV:          " & nDbMissing & "  (sectors coberts pel CSV)"
    Debug.Print "[Import] camps omplits (NULL):  " & nFill
    Debug.Print "[Import] conflictes no tocats:  " & nConf
    If overwriteExisting Then Debug.Print "[Import] valors sobreescrits:   " & nOver
    Debug.Print "[Import] metode assignat:       " & nMeth & "  ('" & METHOD_NAME & "')"
    Debug.Print "[Import] divergencies R68:      " & nR68 & "  (azimut vs Facade_Orientation > " & R68_THRESHOLD & " graus)"
    If nDup > 0 Then Debug.Print "[Import] duplicats al CSV:      " & nDup
    Debug.Print "================================================================"
End Sub

' ================================================================
'  Processament d'una fila
'  Index de columnes del CSV:
'   0 Code / 1 Lat / 2 Lon / 3 E / 4 N / 5 Alt / 6 Prec /
'   7 Azimut / 8 Sector8 / 9 Tilt / 10 OrientChunk / 11 MarkerChunk / 12 Notes
' ================================================================
Private Sub ProcessRow(db As DAO.Database, a() As String, code As String, overwrite As Boolean, hasAzim As Boolean, methodID As Variant, ByRef nMatch As Long, ByRef nNoMatch As Long, ByRef nFill As Long, ByRef nConf As Long, ByRef nOver As Long, ByRef nMeth As Long, ByRef nR68 As Long)
    Dim rs As DAO.Recordset
    Dim sql As String
    sql = "SELECT * FROM T_STRUCTURES "
    sql = sql & "WHERE Code='" & Replace(code, "'", "''") & "'"
    Set rs = db.OpenRecordset(sql, dbOpenDynaset)

    If rs.EOF Then
        nNoMatch = nNoMatch + 1
        Debug.Print "  [CSV sense BD] " & code
        rs.Close
        Exit Sub
    End If
    nMatch = nMatch + 1

    Dim wrote As Boolean
    wrote = False
    rs.Edit

    Dim hasCoords As Boolean
    hasCoords = Len(Trim(a(3))) > 0

    If hasCoords Then
        ApplyNum rs, "Coord_Lat_WGS84", a(1), TOL_DEG, overwrite, code, wrote, nFill, nConf, nOver
        ApplyNum rs, "Coord_Lon_WGS84", a(2), TOL_DEG, overwrite, code, wrote, nFill, nConf, nOver
        ApplyNum rs, "Coord_E_UTM", a(3), TOL_M, overwrite, code, wrote, nFill, nConf, nOver
        ApplyNum rs, "Coord_N_UTM", a(4), TOL_M, overwrite, code, wrote, nFill, nConf, nOver
        ApplyNum rs, "Altitude_masl", a(5), TOL_M, overwrite, code, wrote, nFill, nConf, nOver
        ApplyNum rs, "Coord_Precision_m", a(6), TOL_M, overwrite, code, wrote, nFill, nConf, nOver

        ' Metode de coordenades, resolt per nom
        If IsNull(rs!ID_Coord_Method) Then
            rs!ID_Coord_Method = methodID
            nMeth = nMeth + 1
            wrote = True
        ElseIf rs!ID_Coord_Method <> methodID Then
            If overwrite Then
                Debug.Print "  [CANVI] " & code & ".ID_Coord_Method: " & rs!ID_Coord_Method & " -> " & methodID & " (" & METHOD_NAME & ")"
                rs!ID_Coord_Method = methodID
                nOver = nOver + 1
                wrote = True
            Else
                Debug.Print "  [CONFLICTE] " & code & ".ID_Coord_Method: la BD te un altre metode (no tocat)"
                nConf = nConf + 1
            End If
        End If
    End If

    ' Azimut (camp v24)
    If hasAzim And Len(Trim(a(7))) > 0 Then
        ApplyNum rs, "Facade_Azimuth_Deg", a(7), TOL_AZ, overwrite, code, wrote, nFill, nConf, nOver
    End If

    ' Previsualitzacio R68: azimut calculat vs sector observat
    If Len(Trim(a(7))) > 0 And Not IsNull(rs!Facade_Orientation) Then
        Dim ctr As Double
        ctr = SectorCenter(Nz(rs!Facade_Orientation, ""))
        If ctr >= 0 Then
            Dim d As Double
            d = CircDiff(Val(a(7)), ctr)
            If d > R68_THRESHOLD Then
                nR68 = nR68 + 1
                Debug.Print "  [R68?] " & code & ": azimut " & Trim(a(7)) & " (" & Trim(a(8)) & ") vs Facade_Orientation='" & rs!Facade_Orientation & "' (dif. " & Format(d, "0") & " graus) - bbox girat o error d'entrada?"
            End If
        End If
    End If

    If wrote Then
        rs.Update
    Else
        rs.CancelUpdate
    End If
    rs.Close
End Sub

' ================================================================
'  Helpers
' ================================================================

' Escriu txt (numeric) a fld segons la politica: NULL s'ompli;
' valor diferent es conflicte (o canvi llistat en mode overwrite);
' valor igual (dins de tol) no es toca.
Private Sub ApplyNum(rs As DAO.Recordset, fld As String, txt As String, tol As Double, overwrite As Boolean, code As String, ByRef wrote As Boolean, ByRef nFill As Long, ByRef nConf As Long, ByRef nOver As Long)
    Dim t As String
    t = Trim(txt)
    If Len(t) = 0 Then Exit Sub
    Dim v As Double
    v = Val(t)
    If IsNull(rs.Fields(fld).Value) Then
        rs.Fields(fld).Value = v
        nFill = nFill + 1
        wrote = True
    ElseIf Abs(CDbl(rs.Fields(fld).Value) - v) > tol Then
        If overwrite Then
            Debug.Print "  [CANVI] " & code & "." & fld & ": " & rs.Fields(fld).Value & " -> " & v
            rs.Fields(fld).Value = v
            nOver = nOver + 1
            wrote = True
        Else
            Debug.Print "  [CONFLICTE] " & code & "." & fld & ": BD=" & rs.Fields(fld).Value & " CSV=" & v & " (no tocat)"
            nConf = nConf + 1
        End If
    End If
End Sub

Private Function LookupMethodID(db As DAO.Database) As Variant
    LookupMethodID = Null
    Dim rs As DAO.Recordset
    Dim sql As String
    sql = "SELECT ID FROM L_COORD_METHOD "
    sql = sql & "WHERE Name='" & METHOD_NAME & "'"
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    If Not rs.EOF Then LookupMethodID = rs!ID
    rs.Close
End Function

Private Function SectorCenter(sector As String) As Double
    Select Case Trim(sector)
        Case "N": SectorCenter = 0
        Case "NE": SectorCenter = 45
        Case "E": SectorCenter = 90
        Case "SE": SectorCenter = 135
        Case "S": SectorCenter = 180
        Case "SW": SectorCenter = 225
        Case "W": SectorCenter = 270
        Case "NW": SectorCenter = 315
        Case Else: SectorCenter = -1
    End Select
End Function

Private Function CircDiff(aDeg As Double, bDeg As Double) As Double
    Dim d As Double
    d = Abs(aDeg - bDeg)
    If d > 180 Then d = 360 - d
    CircDiff = d
End Function

Private Function FieldExists(db As DAO.Database, tbl As String, fld As String) As Boolean
    On Error Resume Next
    Dim s As String
    s = db.TableDefs(tbl).Fields(fld).Name
    FieldExists = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Function HasKey(col As Collection, key As String) As Boolean
    On Error Resume Next
    Dim v As Variant
    v = col(key)
    HasKey = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Sub AddKey(col As Collection, key As String)
    If Not HasKey(col, key) Then col.Add key, key
End Sub

Private Function PickFile() As String
    PickFile = ""
    On Error Resume Next
    Dim fd As Object
    Set fd = Application.FileDialog(3)
    If Err.Number = 0 And Not fd Is Nothing Then
        On Error GoTo 0
        fd.Title = "CSV de coordenades i orientacions (Metashape)"
        fd.AllowMultiSelect = False
        fd.Filters.Clear
        fd.Filters.Add "CSV", "*.csv"
        If fd.Show = -1 Then PickFile = fd.SelectedItems(1)
        Exit Function
    End If
    Err.Clear
    On Error GoTo 0
    PickFile = InputBox("Ruta completa del CSV:", "Import georef")
End Function
