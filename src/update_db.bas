Attribute VB_Name = "update_db"
Option Compare Database
Option Explicit

' ================================================================
'  UPDATE T_STRUCTURES - Add fields + rename ID_Status
'  Run Sub UpdateDB() -> F5
' ================================================================

Sub UpdateDB()
    Dim db As DAO.Database
    Set db = CurrentDb()

    RenameStatusField db
    AddNewFields db
    CreateMaterialStatus db
    UpdateRelationships db
    UpdateQueries db
    UpdateFormCombos
    db.TableDefs.Refresh
    Set db = Nothing

    Dim msg As String
    msg = "Database updated successfully!" & vbCrLf & vbCrLf
    msg = msg & "  ID_Status renamed to ID_Arch_Status" & vbCrLf
    msg = msg & "  Interior_Area_m2 added" & vbCrLf
    msg = msg & "  ID_Material_Status added" & vbCrLf
    msg = msg & "  L_MATERIAL_STATUS created and populated" & vbCrLf
    msg = msg & "  Queries updated" & vbCrLf
    msg = msg & "  Form combos updated"
    MsgBox msg, vbInformation, "Done!"
End Sub

' -- 1. RENAME ID_Status -> ID_Arch_Status --------------------
Sub RenameStatusField(db As DAO.Database)
    ' Delete relationship first (references field name)
    On Error Resume Next
    db.Relations.Delete "REL_STA_STR"
    On Error GoTo 0

    ' Rename field via DAO
    Dim tdf As DAO.TableDef
    Set tdf = db.TableDefs("T_STRUCTURES")
    On Error Resume Next
    tdf.Fields("ID_Status").Name = "ID_Arch_Status"
    tdf.Fields.Refresh
    On Error GoTo 0
    Debug.Print "[OK] ID_Status renamed to ID_Arch_Status"
End Sub

' -- 2. ADD NEW COLUMNS ---------------------------------------
Sub AddNewFields(db As DAO.Database)
    On Error Resume Next
    db.Execute "ALTER TABLE T_STRUCTURES ADD COLUMN Interior_Area_m2 SINGLE", dbFailOnError
    db.Execute "ALTER TABLE T_STRUCTURES ADD COLUMN ID_Material_Status LONG", dbFailOnError
    On Error GoTo 0
    Debug.Print "[OK] Interior_Area_m2 and ID_Material_Status added"
End Sub

' -- 3. CREATE L_MATERIAL_STATUS ------------------------------
Sub CreateMaterialStatus(db As DAO.Database)
    ' Create table if not exists
    If Not TableExists(db, "L_MATERIAL_STATUS") Then
        db.Execute "CREATE TABLE L_MATERIAL_STATUS (ID COUNTER CONSTRAINT PK_MS PRIMARY KEY, Name TEXT(50) NOT NULL, Description TEXT(255))", dbFailOnError
    End If

    ' Clear and populate
    db.Execute "DELETE FROM L_MATERIAL_STATUS", dbFailOnError

    Dim v(4, 1) As String
    v(0, 0) = "Good"
    v(0, 1) = "Material remains well preserved and identifiable (mummies, bundles, textiles in good condition)."
    v(1, 0) = "Fair"
    v(1, 1) = "Partially preserved: some elements present and identifiable, others missing or degraded."
    v(2, 0) = "Poor"
    v(2, 1) = "Fragmentary: barely identifiable remains, heavily degraded or scattered."
    v(3, 0) = "Absent"
    v(3, 1) = "No movable remains documented (cause recorded separately via Looting, Animal_Activity, etc.)."
    v(4, 0) = "ND"
    v(4, 1) = "Not determined: not assessed or structure not accessible during documentation."

    Dim i As Integer
    For i = 0 To 4
        db.Execute "INSERT INTO L_MATERIAL_STATUS (Name,Description) VALUES ('" & v(i, 0) & "','" & v(i, 1) & "')", dbFailOnError
    Next i
    Debug.Print "[OK] L_MATERIAL_STATUS created with 7 values"
End Sub

' -- 4. RECREATE RELATIONSHIPS --------------------------------
Sub UpdateRelationships(db As DAO.Database)
    ' Recreate architecture status relationship (renamed field)
    On Error Resume Next
    db.Relations.Delete "REL_STA_STR"
    On Error GoTo 0
    MkRel db, "REL_STA_STR", "L_STATUS", "ID", "T_STRUCTURES", "ID_Arch_Status", False, False

    ' Create material status relationship
    On Error Resume Next
    db.Relations.Delete "REL_MATSTA_STR"
    On Error GoTo 0
    MkRel db, "REL_MATSTA_STR", "L_MATERIAL_STATUS", "ID", "T_STRUCTURES", "ID_Material_Status", False, False

    db.Relations.Refresh
    Debug.Print "[OK] Relationships updated"
End Sub

Sub MkRel(db As DAO.Database, nm As String, pTbl As String, pFld As String, cTbl As String, cFld As String, del As Boolean, noInt As Boolean)
    Dim rel As DAO.Relation
    Dim fld As DAO.Field
    Dim fl As Long
    On Error GoTo ErrR
    fl = dbRelationUpdateCascade
    If del Then fl = fl Or dbRelationDeleteCascade
    If noInt Then fl = 2
    Set rel = db.CreateRelation(nm, pTbl, cTbl, fl)
    Set fld = rel.CreateField(pFld)
    fld.ForeignName = cFld
    rel.Fields.Append fld
    db.Relations.Append rel
    Exit Sub
ErrR:
    Debug.Print "  Warning " & nm & ": " & Err.Description
End Sub

Function TableExists(db As DAO.Database, n As String) As Boolean
    Dim T As DAO.TableDef
    For Each T In db.TableDefs
        If T.Name = n Then TableExists = True: Exit Function
    Next T
End Function

' -- 5. UPDATE QUERIES ----------------------------------------
Sub UpdateQueries(db As DAO.Database)
    ' QRY_03 - Conservation by Sector (uses ID_Arch_Status now)
    Dim q As String
    On Error Resume Next
    db.QueryDefs.Delete "QRY_03_Conservation_by_Sector"
    On Error GoTo 0
    q = "SELECT S.Site_Name, SC.Sector_Name, "
    q = q & "AS1.Name AS Arch_Status, MS.Name AS Material_Status, "
    q = q & "COUNT(*) AS N "
    q = q & "FROM ((((T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID) "
    q = q & "LEFT JOIN L_STATUS AS AS1 ON E.ID_Arch_Status=AS1.ID) "
    q = q & "LEFT JOIN L_MATERIAL_STATUS AS MS ON E.ID_Material_Status=MS.ID) "
    q = q & "GROUP BY S.Site_Name, SC.Sector_Name, AS1.Name, MS.Name "
    q = q & "ORDER BY S.Site_Name, SC.Sector_Name;"
    db.CreateQueryDef "QRY_03_Conservation_by_Sector", q

    ' QRY_05 - Full flat export (update status field references)
    On Error Resume Next
    db.QueryDefs.Delete "QRY_05_Export_RStats"
    On Error GoTo 0
    q = "SELECT E.ID, E.Code, S.Site_Name AS Site, SC.Sector_Name AS Sector, "
    q = q & "T.Name AS Typology, SU.Name AS Support, "
    q = q & "AS1.Name AS Arch_Status, MS.Name AS Material_Status, "
    q = q & "E.N_Floors, E.Floor_Plan, E.N_Built_Walls, E.Access_Orientation, "
    q = q & "E.Natural_Roof, E.Buttresses, E.Plastered, E.Rock_Painting, "
    q = q & "E.Dec_Square_Niche, E.Dec_Relief_T, E.Dec_Relief_T_Inv, "
    q = q & "E.Dec_Relief_L, E.Dec_Relief_L_Inv, E.Dec_Zigzag, "
    q = q & "E.Dec_Stepped, E.Dec_Frieze, E.Rock_Art, "
    q = q & "E.Looting, E.Fire_Damage, E.Animal_Activity, E.Modern_Access, "
    q = q & "E.Human_Remains, E.MNI, E.Mummification, E.Funerary_Bundles, "
    q = q & "E.Bone_Burning, E.Mat_Textiles, E.Mat_Ceramics, E.Mat_DeerAntler, "
    q = q & "E.C14, E.Chrono_Start_Cent, E.Chrono_End_Cent, "
    q = q & "E.Interior_Area_m2, E.Interior_Vol_m3, E.Total_Vol_m3, "
    q = q & "E.Coord_Lat_WGS84, E.Coord_Lon_WGS84, "
    q = q & "E.Coord_E_UTM, E.Coord_N_UTM, E.Altitude_masl, "
    q = q & "E.ChaXR_Documented "
    q = q & "FROM (((((T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID) "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID) "
    q = q & "INNER JOIN L_SUPPORT AS SU ON E.ID_Support=SU.ID) "
    q = q & "LEFT JOIN L_STATUS AS AS1 ON E.ID_Arch_Status=AS1.ID) "
    q = q & "LEFT JOIN L_MATERIAL_STATUS AS MS ON E.ID_Material_Status=MS.ID;"
    db.CreateQueryDef "QRY_05_Export_RStats", q

    ' QRY_07 - QGIS export (update status field)
    On Error Resume Next
    db.QueryDefs.Delete "QRY_07_Export_QGIS"
    On Error GoTo 0
    q = "SELECT E.ID, E.Code, S.Site_Name, SC.Sector_Name, T.Name AS Typology, "
    q = q & "E.Coord_Lat_WGS84, E.Coord_Lon_WGS84, "
    q = q & "E.Coord_E_UTM, E.Coord_N_UTM, E.Altitude_masl, "
    q = q & "E.Approx_Height_m, E.Coord_Precision_m, "
    q = q & "AS1.Name AS Arch_Status, MS.Name AS Material_Status, "
    q = q & "E.N_Floors, E.Interior_Area_m2, E.Interior_Vol_m3, "
    q = q & "E.Chrono_Start_Cent, E.Chrono_End_Cent, "
    q = q & "E.Looting, E.Human_Remains, E.MNI, "
    q = q & "E.ChaXR_Documented, E.URL_3D "
    q = q & "FROM ((((T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID) "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID) "
    q = q & "LEFT JOIN L_STATUS AS AS1 ON E.ID_Arch_Status=AS1.ID) "
    q = q & "LEFT JOIN L_MATERIAL_STATUS AS MS ON E.ID_Material_Status=MS.ID "
    q = q & "WHERE E.Coord_Lat_WGS84 IS NOT NULL ORDER BY S.Site_Name, E.Code;"
    db.CreateQueryDef "QRY_07_Export_QGIS", q

    Debug.Print "[OK] Queries updated"
End Sub

' -- 6. UPDATE FORM COMBOS ------------------------------------
Sub UpdateFormCombos()
    Const frm = "F_STRUCTURES"
    DoCmd.OpenForm frm, acDesign
    Dim f As Form
    Set f = Forms(frm)

    Dim ctrl As Control
    For Each ctrl In f.Controls
        If ctrl.ControlType = acComboBox Then
            ' Update old ID_Status -> ID_Arch_Status
            If ctrl.ControlSource = "ID_Status" Then
                ctrl.ControlSource = "ID_Arch_Status"
                Debug.Print "[OK] Form combo updated: ID_Status -> ID_Arch_Status"
            End If
            ' Configure ID_Arch_Status combo (might be named ID_Arch_Status already)
            If ctrl.ControlSource = "ID_Arch_Status" Then
                ctrl.RowSourceType = "Table/Query"
                ctrl.RowSource = "SELECT ID, Name FROM L_STATUS ORDER BY ID"
                ctrl.BoundColumn = 1
                ctrl.ColumnCount = 2
                ctrl.ColumnWidths = "0cm;4cm"
                ctrl.LimitToList = True
            End If
        End If
    Next ctrl

    DoCmd.Save acForm, frm
    DoCmd.Close acForm, frm
    Debug.Print "[OK] Form combo for ID_Arch_Status configured"
    Debug.Print "NOTE: Add ID_Material_Status and Interior_Area_m2 to the form manually (Field List panel)"
End Sub
