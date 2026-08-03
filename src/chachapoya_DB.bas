Option Compare Database
Option Explicit

' ================================================================
'  CHACHAPOYA ARCHAEOLOGICAL DATABASE - COMPLETE BUILD SCRIPT v10
'  La Petaca & Diablo Wasi (Leymebamba, Amazonas, Peru)
'  Author: Esteve Ribera Torro | TFM Arqueologia UA
'  Version: v10 - duplicate elements merged, vocabulary closed (Aug 2026)
'  File numbering aligned across DB / Form / schema / methodology.
'
'  Changes vs v4:
'   - 0/1/9 DOMAIN EXTENDED TO ALL OBSERVATIONAL FIELDS (63 BYTE fields).
'     Rationale: YESNO is the only Access type that cannot hold Null, so a
'     FALSE conflates "verified absent" with "could not be observed". Every
'     field whose FALSE is ambiguous is now BYTE 0/1/9, default 9.
'     Converted: morphology (Natural_Roof, Buttresses, Wooden_Stakes,
'     Support_Modified), decoration (7), rock art (6), alterations (4),
'     bioarchaeology (7), cultural materials (7), plus the A-X set from v4.
'     Bioarchaeology and materials matter most: cliff interiors are usually
'     observed through an opening, so a FALSE there is very often "not seen".
'     Decoration matters analytically: those fields feed QRY_02, the source
'     of the LP vs DW chi-squared. Without 0/1/9 that test would measure
'     differential facade preservation and read it as differential practice.
'     STILL YESNO (deliberately): C14 and ChaXR_Documented. These are corpus
'     metadata, not observations - their FALSE is never ambiguous.
'   - SURFACE TREATMENTS SPLIT (operation vs substrate):
'     Plastered/Rock_Painting -> Plaster_Present, Plaster_Color,
'     Plaster_Extent, Pigment_Present, Pigment_Substrate, Pigment_Color,
'     Pigment_Extent. Pigment_Substrate (Plaster / Masonry stone / Bedrock /
'     Mixed / ND) records whether paint was applied over a prepared render or
'     directly onto the masonry face - two different operative sequences.
'     T_DECORATIONS gains Substrate for per-position mixed cases.
'   - DOCUMENTATION CONTEXT: Doc_Basis, Facade_Observability,
'     Interior_Observability + QRY_15_Observability_Bias. These make the 9s
'     interpretable: the sample can be justified, not merely disclaimed (OE1).
'   - ID_Support_Secondary (FK -> L_SUPPORT): composite geological supports are
'     recorded as an ordered pair (dominant + secondary) instead of the opaque
'     'Combined' value, which stays in the lookup for non-decomposable cases.
'     L_SUPPORT itself is unchanged.
'   - Support_Morphology reduced to strictly morphological descriptors
'     (Flat / Concave / Convex / Stepped / Irregular / ND) to remove the
'     overlap with L_SUPPORT, which keeps the support CLASS.
'   - Rear_Wall_Built (BYTE) -> Rear_Wall_Type TEXT(20): under the 0/1/9
'     domain, 0 means "absent" everywhere else, but here it meant "natural
'     bedrock" - a value, not an absence. Now Built masonry / Natural bedrock /
'     Mixed / ND. Rear_Wall (W) still carries the A-X presence.
'   - Corbel_Material -> Platform_Surface_Material: E is timber and G is stone
'     by definition, so the old field was partly derivable and could contradict
'     E/G. It now records only the surface material of platform H.
'   - N_Floors -> N_Bodies and L_STRUCT_BODY.Body_Level -> Body_No, completing
'     the v4 terminology split ("nivell" = N0/N1/Sup; "cos" = superposed body).
'   - v8: Support_Morphology DROPPED - redundant with the ID_Support +
'     ID_Support_Secondary pair, and the overlap had already produced a
'     contradictory test record (Medium cavity + Fissure). QRY_12 reworked.
'   - v8: Access_Orientation DROPPED - on a cliff face it is the same variable
'     as Facade_Orientation. One field, one entry, no divergence to arbitrate.
'   - v8: Natural_Roof DROPPED, absorbed into Chamber_Roof_Type TEXT(25)
'     (Natural bedrock / Built masonry / Built timber and slabs / Mixed / ND).
'     Same pattern as W + Rear_Wall_Type: element X carries presence for the
'     A-X matrix, the type field carries the natural-vs-built distinction.
'     NOTE ON MATRIX SEMANTICS: AX_X now means "chamber closed above", which
'     may be a geological given rather than a constructive decision. QRY_13
'     exports Chamber_Roof_Type so a built-only variant of X can be derived
'     in R when the analysis concerns constructive choices specifically.
'   - v8: QRY_02 rebuilt - three counters per motif (Present / Absent / ND)
'     so the chi-squared denominator can be chosen explicitly. See the long
'     comment on that query.
'   - v8: QRY_16_Validation_Check - non-blocking coherence report, including
'     the rule "A-X must be 9, not 0, when ID_Arch_Status = Collapsed".
'   - v9: N_Bodies SPLIT into N_Basal_Bodies + N_Chamber_Bodies.
'     The single counter could not answer its own question: in a structure with
'     two basal masses under one chamber, was it 3 or 1? The two levels of the
'     vocabulary (N0 / N1) are FUNCTIONAL CATEGORIES, not a numbering scheme,
'     so they must not be extended with N2, N3... - "N2" would mean "second
'     basal mass" in one structure and "second chamber" in another, and the
'     A-X matrix would stop being comparable. Repetition therefore goes in the
'     counters, while the category set stays fixed at three.
'     OPERATIVE CRITERION for assigning a body to a level:
'       a body is N1 if it holds (or held) an access opening;
'       otherwise it is N0, however fine its masonry.
'   - v9: Lost_Body_Evidence TEXT(40) - Pigment on bedrock / Truncated walls /
'     Empty beam sockets / Corbels into void / Detached debris / None / ND.
'     Vertical bands of pigment applied directly to the bedrock ABOVE a
'     surviving body are the ghost of a body that is gone: the paint outlived
'     the masonry that carried it. Note this is already partly captured by
'     Pigment_Substrate = Bedrock. Two consequences: it serves H03 (a one-body
'     structure and a mutilated two-body structure are no longer conflated),
'     and it fixes when A-X elements must be 9 rather than 0 - the same logic
'     as the collapse rule, applied vertically. Enforced by QRY_16 rule 8.
'   - v9: L_STRUCT_BODY carries POSITION ONLY (Socle, Spandrel, Jamb,
'     Over-lintel, Cornice); the body index lives in T_DECORATIONS.Body_No.
'     The old N1-SPA / N2-SPA pairs described the same position on different
'     bodies and would have needed N3-SPA for a three-body structure. Level_Type
'     replaces the old Body_No column in the lookup. 8 entries -> 6.
'   - v9: QRY_16 gains two rules (8: lost body with upper elements coded 0;
'     9: chamber bodies counted without an access opening).
'   - v10: TWO DUPLICATED ELEMENT FIELDS MERGED AWAY.
'     Buttresses and Structural_Pilasters (K) denoted the SAME element under
'     two names: the red vertical bands flanking the facade are integrated in
'     the wall plane and run full height, which is a pilaster. Nothing at the
'     sites projects from the facade plane as a true buttress would.
'     Buttresses DROPPED; K carries it.
'     Wooden_Stakes and Timber_Brackets (E) likewise: the protruding horizontal
'     timbers are corbels, and "stake" was simply a second name for them.
'     Wooden_Stakes DROPPED; E carries it.
'     Keeping both pairs could only produce incoherent records, since the same
'     physical element could be coded in either field or in both.
'   - v10: Timber_Bracket_Role TEXT(25) - Platform support / Isolated /
'     Both / ND. This is what the old Wooden_Stakes field was reaching for:
'     not a different element, but a different ROLE. Corbels that do not
'     support a platform are the intra-structure counterpart of the MEN
'     typology (isolated bracket = evidence of a lost aerial circulation
'     network), so recording the role connects the two scales. QRY_16 rule 10.
'   - v10: RA_Decap_Scene DROPPED. A boolean true in one or two structures
'     contributes no variance to any test while occupying a slot in the vector.
'     Recorded instead as a T_DECORATIONS row (new L_DEC_TYPE value) with
'     RA_Anthropomorphic=1, which also gains position, colour and notes.
'   - v10: L_SUPPORT value 'Combined' DROPPED (8 values). Since v8 the ordered
'     pair ID_Support + ID_Support_Secondary carries composite supports; the
'     old value could now only be used badly, because choosing it as the
'     dominant class records no component at all. Three-component supports:
'     record the two dominant ones and note the third in Notes.
'   - v10: L_STRUCT_BODY expanded 6 -> 11 entries, renamed to the A-X
'     vocabulary. 'Spandrel' dropped - it was element L (lateral wall face)
'     under a different and strictly incorrect name. Every position that can
'     carry a surface treatment now has an entry.
'   - v10: L_DEC_TYPE 11 -> 13. 'Plain colour field' turns T_DECORATIONS into
'     the general per-position record of SURFACE TREATMENT, which is how
'     "pilasters red, wall face white" is recorded - the structure-level fields
'     on tab 3 remain the summary. 'Decapitation scene' absorbs the dropped
'     boolean.
'   - T_STRUCTURES: 129 fields | 21 relationships | 16 queries
'
'  Run Sub BuildDB() on a NEW BLANK ACCESS DATABASE
'  Then run chachapoya_Form_v10_val.bas -> Sub BuildForm()
'  (chachapoya_patch_metric.bas is OBSOLETE: absorbed by the form script)
' ================================================================

Sub BuildDB()
    Dim db As DAO.Database
    Set db = CurrentDb()
    CreateAllTables db
    SetByteDefaults db
    PopulateAllLookups db
    CreateAllRelationships db
    CreateAllQueries db
    db.TableDefs.Refresh
    db.QueryDefs.Refresh
    Set db = Nothing
    Dim msg As String
    msg = "DATABASE v10 BUILT SUCCESSFULLY!" & vbCrLf & vbCrLf
    msg = msg & "  19 tables | 21 relationships | 16 queries" & vbCrLf
    msg = msg & "  T_STRUCTURES: 129 fields" & vbCrLf
    msg = msg & "  59 BYTE fields, default 9 (ND)" & vbCrLf & vbCrLf
    msg = msg & "Key changes (v10):" & vbCrLf
    msg = msg & "  Buttresses dropped: same element as pilaster (K)" & vbCrLf
    msg = msg & "  Wooden_Stakes dropped: same element as corbel (E)" & vbCrLf
    msg = msg & "  Timber_Bracket_Role: platform support vs isolated" & vbCrLf
    msg = msg & "  RA_Decap_Scene -> L_DEC_TYPE row" & vbCrLf
    msg = msg & "  L_STRUCT_BODY 6 -> 11 positions (A-X names)" & vbCrLf
    msg = msg & "  L_DEC_TYPE + Plain colour field (per-element pigment)" & vbCrLf
    msg = msg & "  L_SUPPORT: 'Combined' dropped (use the pair)" & vbCrLf & vbCrLf
    msg = msg & "Next: run chachapoya_Form_v10_val.bas -> BuildForm()"
    MsgBox msg, vbInformation, "Done!"
End Sub

Function TableExists(db As DAO.Database, n As String) As Boolean
    Dim t As DAO.TableDef
    For Each t In db.TableDefs
        If t.Name = n Then TableExists = True: Exit Function
    Next t
End Function

Function QueryExists(db As DAO.Database, n As String) As Boolean
    Dim q As DAO.QueryDef
    For Each q In db.QueryDefs
        If q.Name = n Then QueryExists = True: Exit Function
    Next q
End Function

Private Sub X(db As DAO.Database, sql As String)
    On Error GoTo Err_X
    db.Execute sql, dbFailOnError
    Exit Sub
Err_X: Debug.Print "Warning: " & Err.Description
End Sub

' ================================================================
'  1. CREATE ALL TABLES
' ================================================================
Private Sub CreateAllTables(db As DAO.Database)
    Dim sql As String

    ' -- LOOKUP TABLES (no dependencies) --
    If Not TableExists(db, "L_SITES") Then
        db.Execute "CREATE TABLE L_SITES (ID COUNTER CONSTRAINT PK_SIT PRIMARY KEY, Site_Name TEXT(50) NOT NULL, Description MEMO)", dbFailOnError
    End If
    If Not TableExists(db, "L_SECTORS") Then
        db.Execute "CREATE TABLE L_SECTORS (ID COUNTER CONSTRAINT PK_SEC PRIMARY KEY, ID_Site LONG NOT NULL, Sector_Name TEXT(50) NOT NULL, Description TEXT(255))", dbFailOnError
    End If
    If Not TableExists(db, "L_TYPOLOGY") Then
        db.Execute "CREATE TABLE L_TYPOLOGY (ID COUNTER CONSTRAINT PK_TYP PRIMARY KEY, Name TEXT(60) NOT NULL, Description MEMO)", dbFailOnError
    End If
    If Not TableExists(db, "L_SUPPORT") Then
        db.Execute "CREATE TABLE L_SUPPORT (ID COUNTER CONSTRAINT PK_SUP PRIMARY KEY, Name TEXT(60) NOT NULL)", dbFailOnError
    End If
    If Not TableExists(db, "L_STATUS") Then
        db.Execute "CREATE TABLE L_STATUS (ID COUNTER CONSTRAINT PK_STA PRIMARY KEY, Name TEXT(30) NOT NULL)", dbFailOnError
    End If
    If Not TableExists(db, "L_MATERIAL_STATUS") Then
        db.Execute "CREATE TABLE L_MATERIAL_STATUS (ID COUNTER CONSTRAINT PK_MS PRIMARY KEY, Name TEXT(50) NOT NULL, Description TEXT(255))", dbFailOnError
    End If
    If Not TableExists(db, "L_VOL_METHOD") Then
        db.Execute "CREATE TABLE L_VOL_METHOD (ID COUNTER CONSTRAINT PK_VM PRIMARY KEY, Name TEXT(50) NOT NULL, Description TEXT(255))", dbFailOnError
    End If
    If Not TableExists(db, "L_COORD_METHOD") Then
        db.Execute "CREATE TABLE L_COORD_METHOD (ID COUNTER CONSTRAINT PK_CM PRIMARY KEY, Name TEXT(50) NOT NULL, Description TEXT(255))", dbFailOnError
    End If
    If Not TableExists(db, "L_GROUP_TYPE") Then
        db.Execute "CREATE TABLE L_GROUP_TYPE (ID COUNTER CONSTRAINT PK_GT PRIMARY KEY, Name TEXT(50) NOT NULL, Description MEMO)", dbFailOnError
    End If
    If Not TableExists(db, "L_CAMPAIGN") Then
        db.Execute "CREATE TABLE L_CAMPAIGN (ID COUNTER CONSTRAINT PK_CAM PRIMARY KEY, Code TEXT(4) NOT NULL, Campaign_Name TEXT(80), Description MEMO)", dbFailOnError
    End If

    ' Decoration position lookup
    On Error Resume Next
    db.Execute "DROP TABLE L_STRUCT_BODY", dbFailOnError
    On Error GoTo 0
    db.Execute "CREATE TABLE L_STRUCT_BODY (ID COUNTER CONSTRAINT PK_SB PRIMARY KEY, Code TEXT(10) NOT NULL, Name TEXT(60) NOT NULL, Level_Type TEXT(10), Description TEXT(255))", dbFailOnError

    ' Decoration type lookup
    On Error Resume Next
    db.Execute "DROP TABLE L_DEC_TYPE", dbFailOnError
    On Error GoTo 0
    db.Execute "CREATE TABLE L_DEC_TYPE (ID COUNTER CONSTRAINT PK_DT PRIMARY KEY, Name TEXT(60) NOT NULL, Description TEXT(255))", dbFailOnError

    ' -- SECONDARY TABLES --
    If Not TableExists(db, "T_GROUPS") Then
        db.Execute "CREATE TABLE T_GROUPS (ID COUNTER CONSTRAINT PK_GRP PRIMARY KEY, Group_Code TEXT(20) NOT NULL, ID_Sector LONG NOT NULL, ID_Group_Type LONG NOT NULL, N_Members INTEGER, Notes MEMO)", dbFailOnError
    End If

    ' -- MAIN TABLE: T_STRUCTURES (124 fields) --
    If Not TableExists(db, "T_STRUCTURES") Then
        sql = "CREATE TABLE T_STRUCTURES ("
        ' --- 1. Identification (8) ---
        sql = sql & "ID COUNTER CONSTRAINT PK_STR PRIMARY KEY,"
        sql = sql & "Code TEXT(20) NOT NULL,"
        sql = sql & "ID_Sector LONG NOT NULL,"
        sql = sql & "ID_Typology LONG,"
        sql = sql & "ID_Support LONG,"
        sql = sql & "ID_Support_Secondary LONG,"
        sql = sql & "ID_Parent LONG,"
        sql = sql & "ID_Group LONG,"
        ' --- 2. Morphology & dimensions (15) | v9: body counters + lost-body evidence ---
        sql = sql & "N_Basal_Bodies INTEGER,"
        sql = sql & "N_Chamber_Bodies INTEGER,"
        sql = sql & "Lost_Body_Evidence TEXT(40),"
        sql = sql & "Floor_Plan TEXT(20),"
        sql = sql & "N_Built_Walls INTEGER,"
        sql = sql & "Length_m SINGLE,"
        sql = sql & "Width_m SINGLE,"
        sql = sql & "Height_m SINGLE,"
        sql = sql & "Height_Above_Base_m SINGLE,"
        sql = sql & "Dim_Method TEXT(30),"
        sql = sql & "Opening_Width_cm SINGLE,"
        sql = sql & "Opening_Height_cm SINGLE,"
        sql = sql & "Lintel TEXT(20),"
        ' --- 2b. Support geology detail (3) | H02: geology as determinant ---
        sql = sql & "Support_Width_cm SINGLE,"
        sql = sql & "Support_Depth_cm SINGLE,"
        sql = sql & "Support_Modified BYTE,"
        ' --- 2c. Masonry & mortar (6) | H01/H04: cf. Toyne & Anzellini 2017 ---
        sql = sql & "Masonry_Quality TEXT(20),"
        sql = sql & "Masonry_Type TEXT(30),"
        sql = sql & "Mortar_Present BYTE,"
        sql = sql & "Mortar_Type TEXT(30),"
        sql = sql & "Chinking_Stones BYTE,"
        sql = sql & "Mortar_Notes TEXT(150),"
        ' --- 3. A-X systems - Level N0 (10) | elements A-H ---
        '     BYTE fields: 0=Absent / 1=Present / 9=ND (default 9 via DAO)
        sql = sql & "Embedded_Base_Beams BYTE,"
        sql = sql & "Base_Level BYTE,"
        sql = sql & "Decorative_Socle BYTE,"
        sql = sql & "Tie_Walls BYTE,"
        sql = sql & "Timber_Brackets BYTE,"
        sql = sql & "Timber_Bracket_Count INTEGER,"
        sql = sql & "Timber_Bracket_Role TEXT(25),"
        sql = sql & "Transverse_Beams BYTE,"
        sql = sql & "Corbelled_Courses BYTE,"
        sql = sql & "Platform_Surface_Material TEXT(20),"
        sql = sql & "Corbelled_Platform BYTE,"
        ' --- 3b. N0/N1 interface (2) | element I ---
        sql = sql & "Interbody_Cornice BYTE,"
        sql = sql & "Interbody_Cornice_Material TEXT(20),"
        ' --- 4. A-X systems - Level N1 (12) | elements J-P, R, V, W ---
        sql = sql & "Corner_Quoins BYTE,"
        sql = sql & "Structural_Pilasters BYTE,"
        sql = sql & "Lateral_Wall_Faces BYTE,"
        sql = sql & "Relief_Frieze BYTE,"
        sql = sql & "Sill BYTE,"
        sql = sql & "Jambs BYTE,"
        sql = sql & "Access_Opening BYTE,"
        sql = sql & "Recessed_Portal BYTE,"
        sql = sql & "Upper_Crown BYTE,"
        sql = sql & "Lateral_Walls BYTE,"
        sql = sql & "Rear_Wall BYTE,"
        sql = sql & "Rear_Wall_Type TEXT(20),"
        ' --- 5. A-X systems - Upper zone (4) | elements S, T, U, X ---
        sql = sql & "Eave_Beam BYTE,"
        sql = sql & "Eave_Surface BYTE,"
        sql = sql & "Eave BYTE,"
        sql = sql & "Chamber_Roof BYTE,"
        sql = sql & "Chamber_Roof_Type TEXT(25),"
        ' --- 6. Surface treatments (7) | v7: operation separated from substrate ---
        '     Plastering and pigment application are two distinct operations.
        '     Pigment_Substrate is the key variable: painting on plaster requires
        '     a preparatory operation; painting directly on masonry means the
        '     stone face WAS the intended finished surface (H01, H04).
        '     It also controls differential preservation: plaster spalls far more
        '     readily than pigment absorbed into porous sandstone.
        sql = sql & "Plaster_Present BYTE,"
        sql = sql & "Plaster_Color TEXT(20),"
        sql = sql & "Plaster_Extent TEXT(20),"
        sql = sql & "Pigment_Present BYTE,"
        sql = sql & "Pigment_Substrate TEXT(20),"
        sql = sql & "Pigment_Color TEXT(20),"
        sql = sql & "Pigment_Extent TEXT(20),"
        ' --- 6b. Landscape & orientation (2) | observational; QGIS later ---
        sql = sql & "Facade_Orientation TEXT(5),"
        sql = sql & "Visibility_Valley TEXT(10),"
        ' --- 7. Decoration summary, BYTE 0/1/9 (13, detail in T_DECORATIONS) ---
        '     Dec_Frieze removed in v4: frieze is now element M (Relief_Frieze)
        sql = sql & "Dec_Square_Niche BYTE,"
        sql = sql & "Dec_Relief_T BYTE,"
        sql = sql & "Dec_Relief_T_Inv BYTE,"
        sql = sql & "Dec_Relief_L BYTE,"
        sql = sql & "Dec_Relief_L_Inv BYTE,"
        sql = sql & "Dec_Zigzag BYTE,"
        sql = sql & "Dec_Stepped BYTE,"
        sql = sql & "Rock_Art BYTE,"
        sql = sql & "RA_Anthropomorphic BYTE,"
        sql = sql & "RA_Zoomorphic BYTE,"
        sql = sql & "RA_Geometric BYTE,"
        sql = sql & "RA_Abstract BYTE,"
        ' --- 8. Conservation (6) | alterations BYTE 0/1/9 ---
        sql = sql & "ID_Arch_Status LONG,"
        sql = sql & "ID_Material_Status LONG,"
        sql = sql & "Looting BYTE,"
        sql = sql & "Fire_Damage BYTE,"
        sql = sql & "Animal_Activity BYTE,"
        sql = sql & "Modern_Access BYTE,"
        ' --- 9. Bioarchaeology (8) | BYTE 0/1/9: interiors are rarely fully observed ---
        sql = sql & "Human_Remains BYTE,"
        sql = sql & "MNI INTEGER,"
        sql = sql & "Anatomical_Connection BYTE,"
        sql = sql & "Mummification BYTE,"
        sql = sql & "Funerary_Bundles BYTE,"
        sql = sql & "Dispersed_Remains BYTE,"
        sql = sql & "Flexed_Position BYTE,"
        sql = sql & "Bone_Burning BYTE,"
        ' --- 10. Cultural materials (7) | BYTE 0/1/9 ---
        sql = sql & "Mat_Textiles BYTE,"
        sql = sql & "Mat_Wood BYTE,"
        sql = sql & "Mat_VegFiber BYTE,"
        sql = sql & "Mat_Ceramics BYTE,"
        sql = sql & "Mat_Fauna BYTE,"
        sql = sql & "Mat_DeerAntler BYTE,"
        sql = sql & "Mat_Other BYTE,"
        ' --- 11. Chronology (3) ---
        sql = sql & "C14 YESNO,"
        sql = sql & "Chrono_Start_Cent INTEGER,"
        sql = sql & "Chrono_End_Cent INTEGER,"
        ' --- 11b. Constructive phases (2) | H03 / H04 ---
        sql = sql & "Construction_Phases INTEGER,"
        sql = sql & "Phase_Evidence TEXT(30),"
        ' --- 12. Volumetry & area (5) ---
        sql = sql & "Interior_Area_m2 SINGLE,"
        sql = sql & "Interior_Vol_m3 SINGLE,"
        sql = sql & "Total_Vol_m3 SINGLE,"
        sql = sql & "ID_Vol_Method LONG,"
        sql = sql & "Vol_Notes TEXT(200),"
        ' --- 13. Spatial coordinates (7) ---
        sql = sql & "Coord_Lat_WGS84 DOUBLE,"
        sql = sql & "Coord_Lon_WGS84 DOUBLE,"
        sql = sql & "Coord_E_UTM DOUBLE,"
        sql = sql & "Coord_N_UTM DOUBLE,"
        sql = sql & "Altitude_masl SINGLE,"
        sql = sql & "Coord_Precision_m SINGLE,"
        sql = sql & "ID_Coord_Method LONG,"
        ' --- 14. Digital documentation & observability (10) ---
        sql = sql & "URL_Pano TEXT(255),"
        sql = sql & "URL_Pano_2 TEXT(255),"
        sql = sql & "URL_Giga TEXT(255),"
        sql = sql & "URL_3D TEXT(255),"
        sql = sql & "ChaXR_Documented YESNO,"
        sql = sql & "ID_Campaign LONG,"
        sql = sql & "Doc_Basis TEXT(30),"
        sql = sql & "Facade_Observability TEXT(20),"
        sql = sql & "Interior_Observability TEXT(20),"
        sql = sql & "Notes MEMO)"
        db.Execute sql, dbFailOnError
        Debug.Print "[OK] T_STRUCTURES (129 fields)"
    End If

    ' -- LINKED TABLES --
    If Not TableExists(db, "T_DATING") Then
        db.Execute "CREATE TABLE T_DATING (ID COUNTER CONSTRAINT PK_DAT PRIMARY KEY, ID_Structure LONG NOT NULL, Sample_Type TEXT(30), Date_BP LONG, Sigma1_Start INTEGER, Sigma1_End INTEGER, Sigma2_Start INTEGER, Sigma2_End INTEGER, Lab_Reference TEXT(50), Bibliog_Reference MEMO)", dbFailOnError
    End If
    If Not TableExists(db, "T_INDIVIDUALS") Then
        db.Execute "CREATE TABLE T_INDIVIDUALS (ID COUNTER CONSTRAINT PK_IND PRIMARY KEY, ID_Structure LONG NOT NULL, Individual_No INTEGER, Age_Category TEXT(20), Sex_Category TEXT(20), Preservation TEXT(20), Notes MEMO)", dbFailOnError
    End If

    ' Decoration detail
    On Error Resume Next
    db.Execute "DROP TABLE T_DECORATIONS", dbFailOnError
    On Error GoTo 0
    db.Execute "CREATE TABLE T_DECORATIONS (ID COUNTER CONSTRAINT PK_TDEC PRIMARY KEY, ID_Structure LONG NOT NULL, ID_Struct_Body LONG, ID_Dec_Type LONG, Body_No INTEGER, Color TEXT(20), Substrate TEXT(20), Notes TEXT(255))", dbFailOnError

    ' Flexible feature recording (future-proof)
    On Error Resume Next
    db.Execute "DROP TABLE T_ARCH_FEATURES", dbFailOnError
    On Error GoTo 0
    sql = "CREATE TABLE T_ARCH_FEATURES ("
    sql = sql & "ID COUNTER CONSTRAINT PK_AF PRIMARY KEY,"
    sql = sql & "ID_Structure LONG NOT NULL,"
    sql = sql & "Feature_Code TEXT(40),"
    sql = sql & "Present YESNO,"
    sql = sql & "Feature_Count INTEGER,"
    sql = sql & "Material TEXT(20),"
    sql = sql & "Notes TEXT(255))"
    db.Execute sql, dbFailOnError

    ' NEW v4: physical connections between structures (OE3 network)
    On Error Resume Next
    db.Execute "DROP TABLE T_CONNECTIONS", dbFailOnError
    On Error GoTo 0
    sql = "CREATE TABLE T_CONNECTIONS ("
    sql = sql & "ID COUNTER CONSTRAINT PK_CON PRIMARY KEY,"
    sql = sql & "ID_Struct_A LONG NOT NULL,"
    sql = sql & "ID_Struct_B LONG NOT NULL,"
    sql = sql & "Connection_Type TEXT(30),"
    sql = sql & "Confidence TEXT(10),"
    sql = sql & "Notes TEXT(150))"
    db.Execute sql, dbFailOnError

    Debug.Print "-> 19 tables OK"
End Sub

' ================================================================
'  1b. DEFAULT VALUE 9 (ND) FOR ALL BYTE 0/1/9 FIELDS (via DAO)
'      Rationale: a new record must NOT silently claim absence.
'      Absence (0) has to be positively recorded by the researcher.
' ================================================================
Private Sub SetByteDefaults(db As DAO.Database)
    Dim fn(58) As String
    fn(0) = "Embedded_Base_Beams"
    fn(1) = "Base_Level"
    fn(2) = "Decorative_Socle"
    fn(3) = "Tie_Walls"
    fn(4) = "Timber_Brackets"
    fn(5) = "Transverse_Beams"
    fn(6) = "Corbelled_Courses"
    fn(7) = "Corbelled_Platform"
    fn(8) = "Interbody_Cornice"
    fn(9) = "Corner_Quoins"
    fn(10) = "Structural_Pilasters"
    fn(11) = "Lateral_Wall_Faces"
    fn(12) = "Relief_Frieze"
    fn(13) = "Sill"
    fn(14) = "Jambs"
    fn(15) = "Access_Opening"
    fn(16) = "Upper_Crown"
    fn(17) = "Lateral_Walls"
    fn(18) = "Rear_Wall"
    fn(19) = "Eave_Beam"
    fn(20) = "Eave_Surface"
    fn(21) = "Eave"
    fn(22) = "Chamber_Roof"
    fn(23) = "Recessed_Portal"
    fn(24) = "Mortar_Present"
    fn(25) = "Chinking_Stones"
    fn(26) = "Support_Modified"
    fn(27) = "Plaster_Present"
    fn(28) = "Pigment_Present"
    fn(29) = "Dec_Square_Niche"
    fn(30) = "Dec_Relief_T"
    fn(31) = "Dec_Relief_T_Inv"
    fn(32) = "Dec_Relief_L"
    fn(33) = "Dec_Relief_L_Inv"
    fn(34) = "Dec_Zigzag"
    fn(35) = "Dec_Stepped"
    fn(36) = "Rock_Art"
    fn(37) = "RA_Anthropomorphic"
    fn(38) = "RA_Zoomorphic"
    fn(39) = "RA_Geometric"
    fn(40) = "RA_Abstract"
    fn(41) = "Looting"
    fn(42) = "Fire_Damage"
    fn(43) = "Animal_Activity"
    fn(44) = "Modern_Access"
    fn(45) = "Human_Remains"
    fn(46) = "Anatomical_Connection"
    fn(47) = "Mummification"
    fn(48) = "Funerary_Bundles"
    fn(49) = "Dispersed_Remains"
    fn(50) = "Flexed_Position"
    fn(51) = "Bone_Burning"
    fn(52) = "Mat_Textiles"
    fn(53) = "Mat_Wood"
    fn(54) = "Mat_VegFiber"
    fn(55) = "Mat_Ceramics"
    fn(56) = "Mat_Fauna"
    fn(57) = "Mat_DeerAntler"
    fn(58) = "Mat_Other"
    Dim i As Integer
    On Error Resume Next
    For i = 0 To 58
        db.TableDefs("T_STRUCTURES").Fields(fn(i)).DefaultValue = "9"
    Next i
    On Error GoTo 0
    Debug.Print "-> BYTE defaults (9 = ND) set on 59 fields"
End Sub

' ================================================================
'  2. POPULATE ALL LOOKUPS
' ================================================================
Private Sub PopulateAllLookups(db As DAO.Database)
    Dim i As Integer

    X db, "DELETE FROM L_SECTORS"
    X db, "DELETE FROM L_SITES"
    X db, "DELETE FROM L_TYPOLOGY"
    X db, "DELETE FROM L_SUPPORT"
    X db, "DELETE FROM L_STATUS"
    X db, "DELETE FROM L_MATERIAL_STATUS"
    X db, "DELETE FROM L_VOL_METHOD"
    X db, "DELETE FROM L_COORD_METHOD"
    X db, "DELETE FROM L_GROUP_TYPE"
    X db, "DELETE FROM L_CAMPAIGN"
    X db, "DELETE FROM L_STRUCT_BODY"
    X db, "DELETE FROM L_DEC_TYPE"

    ' L_SITES
    db.Execute "INSERT INTO L_SITES (Site_Name,Description) VALUES ('La Petaca','>12,000 m2 exposed rock, 4 sectors, 10th-16th c. WGS84: Lat -6.8311 / Lon -77.8084')", dbFailOnError
    db.Execute "INSERT INTO L_SITES (Site_Name,Description) VALUES ('Diablo Wasi','6 sectors, predominance of funerary chambers. WGS84: Lat -6.8475 / Lon -77.8154')", dbFailOnError

    ' L_SECTORS
    Dim s(11, 2) As String
    s(0, 0) = "1": s(0, 1) = "LP - General":  s(0, 2) = "La Petaca - general context"
    s(1, 0) = "1": s(1, 1) = "LP - North":    s(1, 2) = "La Petaca - North Sector"
    s(2, 0) = "1": s(2, 1) = "LP - Central":  s(2, 2) = "La Petaca - Central Sector"
    s(3, 0) = "1": s(3, 1) = "LP - Upper":    s(3, 2) = "La Petaca - Upper Sector"
    s(4, 0) = "1": s(4, 1) = "LP - South":    s(4, 2) = "La Petaca - South Sector (>=96 structures)"
    s(5, 0) = "2": s(5, 1) = "DW - General":  s(5, 2) = "Diablo Wasi - general context"
    s(6, 0) = "2": s(6, 1) = "DW - Sector 1": s(6, 2) = "Diablo Wasi - Sector 1 (~40 funerary contexts)"
    s(7, 0) = "2": s(7, 1) = "DW - Sector 2": s(7, 2) = "Diablo Wasi - Sector 2 (1 chamber + basal cave)"
    s(8, 0) = "2": s(8, 1) = "DW - Sector 3": s(8, 2) = "Diablo Wasi - Sector 3 (underground cave)"
    s(9, 0) = "2": s(9, 1) = "DW - Sector 4": s(9, 2) = "Diablo Wasi - Sector 4 (~9 funerary contexts)"
    s(10, 0) = "2": s(10, 1) = "DW - Sector 5": s(10, 2) = "Diablo Wasi - Sector 5"
    s(11, 0) = "2": s(11, 1) = "DW - Sector 6": s(11, 2) = "Diablo Wasi - Sector 6"
    For i = 0 To 11
        db.Execute "INSERT INTO L_SECTORS (ID_Site,Sector_Name,Description) VALUES (" & s(i, 0) & ",'" & s(i, 1) & "','" & s(i, 2) & "')", dbFailOnError
    Next i

    ' L_TYPOLOGY
    Dim t(9, 1) As String
    t(0, 0) = "EA-MAU Mausoleum/Chullpa":  t(0, 1) = "Built structure (3+ walls + artificial roof) on ledge. 1-3 storeys. Predominant at La Petaca."
    t(1, 0) = "EA-CAM Funerary Chamber":   t(1, 1) = "Natural cavity closed by 1 built facade. Predominant at Diablo Wasi."
    t(2, 0) = "EA-PLA-R Ledge Platform":   t(2, 1) = "Constructive platform on natural ledge. Function: transit or mausoleum base."
    t(3, 0) = "EA-PLA-V Aerial Platform":  t(3, 1) = "Artificial platform on wooden beams and slabs, without natural ledge support."
    t(4, 0) = "NIX Natural Niche":         t(4, 1) = "Small natural cavity (<1m2). Function: ossuary or secondary burial."
    t(5, 0) = "CAV Cave/Cavern":           t(5, 1) = "Large natural cavity (>1m2) with documented funerary or ritual use."
    t(6, 0) = "PR Rock Art":               t(6, 1) = "Pictorial motif on rock, independently documented."
    t(7, 0) = "MEN Isolated Bracket":      t(7, 1) = "Isolated structural element. Evidence of lost aerial circulation network."
    t(8, 0) = "MIX Mixed":                 t(8, 1) = "Combination of two or more previous categories."
    t(9, 0) = "ND Undetermined":           t(9, 1) = "Insufficient information to classify."
    For i = 0 To 9
        db.Execute "INSERT INTO L_TYPOLOGY (Name,Description) VALUES ('" & t(i, 0) & "','" & t(i, 1) & "')", dbFailOnError
    Next i

    ' L_SUPPORT
    Dim sup(7) As String
    sup(0) = "Wide natural ledge (>2m)": sup(1) = "Narrow natural ledge (<2m)": sup(2) = "Artificial ledge"
    sup(3) = "Large cavity (>10m2)": sup(4) = "Medium cavity (1-10m2)": sup(5) = "Natural niche (<1m2)"
    sup(6) = "Fissure/Crack": sup(7) = "ND"
    For i = 0 To 7
        db.Execute "INSERT INTO L_SUPPORT (Name) VALUES ('" & sup(i) & "')", dbFailOnError
    Next i

    ' L_STATUS
    Dim st(4) As String
    st(0) = "Good": st(1) = "Fair": st(2) = "Pre-collapse": st(3) = "Collapsed": st(4) = "ND"
    For i = 0 To 4
        db.Execute "INSERT INTO L_STATUS (Name) VALUES ('" & st(i) & "')", dbFailOnError
    Next i

    ' L_MATERIAL_STATUS
    db.Execute "INSERT INTO L_MATERIAL_STATUS (Name,Description) VALUES ('Good','Material remains well preserved and identifiable.')", dbFailOnError
    db.Execute "INSERT INTO L_MATERIAL_STATUS (Name,Description) VALUES ('Fair','Partially preserved: some elements present and identifiable.')", dbFailOnError
    db.Execute "INSERT INTO L_MATERIAL_STATUS (Name,Description) VALUES ('Poor','Fragmentary: barely identifiable remains, heavily degraded or scattered.')", dbFailOnError
    db.Execute "INSERT INTO L_MATERIAL_STATUS (Name,Description) VALUES ('Absent','No movable remains documented (cause in Looting/Animal_Activity fields).')", dbFailOnError
    db.Execute "INSERT INTO L_MATERIAL_STATUS (Name,Description) VALUES ('ND','Not determined: not assessed or structure not accessible.')", dbFailOnError

    ' L_VOL_METHOD
    db.Execute "INSERT INTO L_VOL_METHOD (Name,Description) VALUES ('L x W x H Calculation','Geometric calculation for regular rectangular chambers.')", dbFailOnError
    db.Execute "INSERT INTO L_VOL_METHOD (Name,Description) VALUES ('Photogrammetric Model','Volume extracted from 3D model (MeshLab, CloudCompare).')", dbFailOnError
    db.Execute "INSERT INTO L_VOL_METHOD (Name,Description) VALUES ('Estimation','Visual estimation or partial measurements.')", dbFailOnError
    db.Execute "INSERT INTO L_VOL_METHOD (Name,Description) VALUES ('ND','Undetermined.')", dbFailOnError

    ' L_COORD_METHOD
    db.Execute "INSERT INTO L_COORD_METHOD (Name,Description) VALUES ('Drone RTK','RTK GPS drone. Precision ~2-5 cm.')", dbFailOnError
    db.Execute "INSERT INTO L_COORD_METHOD (Name,Description) VALUES ('Differential GPS','Dual-frequency GNSS. Precision ~5-20 cm.')", dbFailOnError
    db.Execute "INSERT INTO L_COORD_METHOD (Name,Description) VALUES ('Photogrammetry','Coordinates from georeferenced Metashape model.')", dbFailOnError
    db.Execute "INSERT INTO L_COORD_METHOD (Name,Description) VALUES ('Mobile GPS','Mobile/tablet GPS. Precision 3-10 m. Orientation only.')", dbFailOnError
    db.Execute "INSERT INTO L_COORD_METHOD (Name,Description) VALUES ('Estimation','Estimated from cartography or orthophoto. >10 m.')", dbFailOnError
    db.Execute "INSERT INTO L_COORD_METHOD (Name,Description) VALUES ('ND','Undetermined.')", dbFailOnError

    ' L_GROUP_TYPE
    db.Execute "INSERT INTO L_GROUP_TYPE (Name,Description) VALUES ('Vertical alignment','Structures on the same vertical of the cliff (fissure or successive strata).')", dbFailOnError
    db.Execute "INSERT INTO L_GROUP_TYPE (Name,Description) VALUES ('Ledge cluster','Multiple structures sharing the same horizontal ledge.')", dbFailOnError
    db.Execute "INSERT INTO L_GROUP_TYPE (Name,Description) VALUES ('Platform with brackets','Aerial platform + brackets composing or flanking it.')", dbFailOnError
    db.Execute "INSERT INTO L_GROUP_TYPE (Name,Description) VALUES ('Cave cluster','Cave + structures built inside it.')", dbFailOnError
    db.Execute "INSERT INTO L_GROUP_TYPE (Name,Description) VALUES ('Circulation network','MENs and EA-PLAs reconstructing a lost aerial circulation route.')", dbFailOnError
    db.Execute "INSERT INTO L_GROUP_TYPE (Name,Description) VALUES ('Rock art cluster','Rock art + adjacent structure visually or functionally linked.')", dbFailOnError
    db.Execute "INSERT INTO L_GROUP_TYPE (Name,Description) VALUES ('Functional group','Any other intra-sector grouping with functional coherence.')", dbFailOnError

    ' L_CAMPAIGN
    db.Execute "INSERT INTO L_CAMPAIGN (Code,Campaign_Name,Description) VALUES ('2013','PALP I','First campaign. Prospection and initial documentation of La Petaca. Dr. J. Marla Toyne (UCF).')", dbFailOnError
    db.Execute "INSERT INTO L_CAMPAIGN (Code,Campaign_Name,Description) VALUES ('2016','PALP II','Extension to Diablo Wasi. First systematic documentation of DW.')", dbFailOnError
    db.Execute "INSERT INTO L_CAMPAIGN (Code,Campaign_Name,Description) VALUES ('2021','La Petaca Project','Non-invasive integral documentation. Photogrammetry, 360, gigaphotos. Panograma Labs/UCF.')", dbFailOnError
    db.Execute "INSERT INTO L_CAMPAIGN (Code,Campaign_Name,Description) VALUES ('2023','PALP IV','Archaeological excavation campaign and detailed 3D reconstructions.')", dbFailOnError

    ' L_STRUCT_BODY - POSITION WITHIN A BODY ONLY (which body: T_DECORATIONS.Body_No)
    ' v10: expanded from 6 to 11 entries and renamed to match the A-X vocabulary.
    ' "Spandrel" was dropped: it denoted the SAME thing as element L (lateral wall
    ' face) under a different - and, strictly, wrong - name (a spandrel is properly
    ' the area beside an arch). Every position that can carry a surface treatment
    ' now has an entry, so T_DECORATIONS can record "pilasters red, wall face
    ' white" as two rows. Element letters given for cross-reference.
    Dim sb(10, 3) As String
    sb(0, 0) = "BAS":  sb(0, 1) = "Base level (B)":        sb(0, 2) = "N0":  sb(0, 3) = "Basal mass treated as a distinct constructive element."
    sb(1, 0) = "SOC":  sb(1, 1) = "Socle (C)":             sb(1, 2) = "N0":  sb(1, 3) = "Decorative socle - treatment of the basal mass."
    sb(2, 0) = "LWF":  sb(2, 1) = "Lateral wall face (L)": sb(2, 2) = "N1":  sb(2, 3) = "Wall face outside the portal frame. Replaces the old Spandrel entry."
    sb(3, 0) = "PIL":  sb(3, 1) = "Pilaster (K)":          sb(3, 2) = "N1":  sb(3, 3) = "Structural pilaster framing the facade, full height."
    sb(4, 0) = "QUO":  sb(4, 1) = "Corner quoin (J)":      sb(4, 2) = "N1":  sb(4, 3) = "Larger stones set vertically at the angles."
    sb(5, 0) = "JAM":  sb(5, 1) = "Jamb (O)":              sb(5, 2) = "N1":  sb(5, 3) = "Portal jamb - vertical frame element of the access opening."
    sb(6, 0) = "OVL":  sb(6, 1) = "Over-lintel":           sb(6, 2) = "N1":  sb(6, 3) = "Zone above the lintel - decorative frieze area (element M)."
    sb(7, 0) = "COR":  sb(7, 1) = "Interbody cornice (I)": sb(7, 2) = "N1":  sb(7, 3) = "Cornice zone between superposed constructive bodies."
    sb(8, 0) = "CRO":  sb(8, 1) = "Upper crown (R)":       sb(8, 2) = "N1":  sb(8, 3) = "Upper crown or coping of the body."
    sb(9, 0) = "EAV":  sb(9, 1) = "Eave (U)":              sb(9, 2) = "SUP": sb(9, 3) = "Eave / roof overhang above the facade."
    sb(10, 0) = "ND":  sb(10, 1) = "Not determined":       sb(10, 2) = "ND":  sb(10, 3) = "Position not determined."
    For i = 0 To 10
        db.Execute "INSERT INTO L_STRUCT_BODY (Code,Name,Level_Type,Description) VALUES ('" & sb(i, 0) & "','" & sb(i, 1) & "','" & sb(i, 2) & "','" & sb(i, 3) & "')", dbFailOnError
    Next i

    ' L_DEC_TYPE
    ' v10: two values added.
    ' "Plain colour field" makes T_DECORATIONS the general per-position record of
    ' SURFACE TREATMENT, not only of motifs: it is how "pilasters red, wall face
    ' white" gets recorded. "Decapitation scene" replaces the former
    ' RA_Decap_Scene boolean, which would have been 1 in one or two structures
    ' and therefore contributed no variance to any test while occupying a slot
    ' in the vector; as a T_DECORATIONS row it gains position, colour and notes.
    Dim dt(12, 1) As String
    dt(0, 0) = "T-shaped niche":        dt(0, 1) = "Niche or bas-relief in T form. Vertical + horizontal element."
    dt(1, 0) = "T-shaped niche inv.":   dt(1, 1) = "Inverted T niche/relief."
    dt(2, 0) = "L-shaped niche":        dt(2, 1) = "Niche or bas-relief in L form."
    dt(3, 0) = "L-shaped niche inv.":   dt(3, 1) = "Inverted L niche/relief."
    dt(4, 0) = "Zigzag":                dt(4, 1) = "Zigzag or chevron motif."
    dt(5, 0) = "Stepped motif":         dt(5, 1) = "Stepped/staircase motif. Rare at DW and LP."
    dt(6, 0) = "Frieze / Greca":        dt(6, 1) = "Fretwork or repeating greca frieze. Common at La Petaca."
    dt(7, 0) = "Triangular motif":      dt(7, 1) = "Painted triangular/chevron pattern. Documented at DW (over-lintel zone)."
    dt(8, 0) = "Painted band":          dt(8, 1) = "Horizontal painted band (red/white). Interbody cornice zone."
    dt(9, 0) = "Square niche":          dt(9, 1) = "Square niches in series (hornacinas cuadradas)."
    dt(10, 0) = "Plain colour field":    dt(10, 1) = "NEW v10. Flat chromatic application with no motif: a whole element painted one colour."
    dt(11, 0) = "Decapitation scene":    dt(11, 1) = "NEW v10. Decapitation scene. Replaces the RA_Decap_Scene boolean; set RA_Anthropomorphic=1 as well."
    dt(12, 0) = "ND":                    dt(12, 1) = "Decoration type not determined."
    For i = 0 To 12
        db.Execute "INSERT INTO L_DEC_TYPE (Name,Description) VALUES ('" & dt(i, 0) & "','" & dt(i, 1) & "')", dbFailOnError
    Next i

    Debug.Print "-> All lookups populated"
End Sub

' ================================================================
'  3. RELATIONSHIPS (20 total: 18 from v3 + 2 for T_CONNECTIONS)
' ================================================================
Private Sub CreateAllRelationships(db As DAO.Database)
    Dim rn(20) As String
    rn(0) = "REL_SIT_SEC":    rn(1) = "REL_SEC_STR":    rn(2) = "REL_SEC_GRP"
    rn(3) = "REL_TYP_STR":    rn(4) = "REL_SUP_STR":    rn(5) = "REL_STR_SELF"
    rn(6) = "REL_STA_STR":    rn(7) = "REL_MATSTA_STR": rn(8) = "REL_VM_STR"
    rn(9) = "REL_CM_STR":     rn(10) = "REL_GT_GRP":    rn(11) = "REL_GRP_STR"
    rn(12) = "REL_STR_DAT":   rn(13) = "REL_STR_IND":   rn(14) = "REL_STR_DEC"
    rn(15) = "REL_SB_DEC":    rn(16) = "REL_DT_DEC":    rn(17) = "REL_STR_AFEAT"
    rn(18) = "REL_STR_CONA":  rn(19) = "REL_STR_CONB"
    rn(20) = "REL_SUP2_STR"
    Dim n As Integer
    For n = 0 To 20
        On Error Resume Next: db.Relations.Delete rn(n): On Error GoTo 0
    Next n

    MkRel db, rn(0), "L_SITES", "ID", "L_SECTORS", "ID_Site", True, False
    MkRel db, rn(1), "L_SECTORS", "ID", "T_STRUCTURES", "ID_Sector", True, False
    MkRel db, rn(2), "L_SECTORS", "ID", "T_GROUPS", "ID_Sector", True, False
    MkRel db, rn(3), "L_TYPOLOGY", "ID", "T_STRUCTURES", "ID_Typology", False, False
    MkRel db, rn(4), "L_SUPPORT", "ID", "T_STRUCTURES", "ID_Support", False, False
    MkRel db, rn(5), "T_STRUCTURES", "ID", "T_STRUCTURES", "ID_Parent", False, True
    MkRel db, rn(6), "L_STATUS", "ID", "T_STRUCTURES", "ID_Arch_Status", False, False
    MkRel db, rn(7), "L_MATERIAL_STATUS", "ID", "T_STRUCTURES", "ID_Material_Status", False, False
    MkRel db, rn(8), "L_VOL_METHOD", "ID", "T_STRUCTURES", "ID_Vol_Method", False, False
    MkRel db, rn(9), "L_COORD_METHOD", "ID", "T_STRUCTURES", "ID_Coord_Method", False, False
    MkRel db, rn(10), "L_GROUP_TYPE", "ID", "T_GROUPS", "ID_Group_Type", False, False
    MkRel db, rn(11), "T_GROUPS", "ID", "T_STRUCTURES", "ID_Group", False, False
    MkRel db, rn(12), "T_STRUCTURES", "ID", "T_DATING", "ID_Structure", True, False
    MkRel db, rn(13), "T_STRUCTURES", "ID", "T_INDIVIDUALS", "ID_Structure", True, False
    MkRel db, rn(14), "T_STRUCTURES", "ID", "T_DECORATIONS", "ID_Structure", True, False
    MkRel db, rn(15), "L_STRUCT_BODY", "ID", "T_DECORATIONS", "ID_Struct_Body", False, False
    MkRel db, rn(16), "L_DEC_TYPE", "ID", "T_DECORATIONS", "ID_Dec_Type", False, False
    MkRel db, rn(17), "T_STRUCTURES", "ID", "T_ARCH_FEATURES", "ID_Structure", True, False
    ' v4: dual reference to T_STRUCTURES -> no enforced integrity (value 2),
    ' same treatment as the self-referencing REL_STR_SELF
    MkRel db, rn(18), "T_STRUCTURES", "ID", "T_CONNECTIONS", "ID_Struct_A", False, True
    MkRel db, rn(19), "T_STRUCTURES", "ID", "T_CONNECTIONS", "ID_Struct_B", False, True
    ' v7: secondary support type (composite geological supports)
    MkRel db, rn(20), "L_SUPPORT", "ID", "T_STRUCTURES", "ID_Support_Secondary", False, False

    db.Relations.Refresh
    Debug.Print "-> 21 relationships OK"
End Sub

Private Sub MkRel(db As DAO.Database, nm As String, pT As String, pF As String, cT As String, cF As String, del As Boolean, noInt As Boolean)
    Dim rel As DAO.Relation
    Dim fld As DAO.Field
    Dim fl As Long
    On Error GoTo ErrR
    If noInt Then
        fl = 2
    Else
        fl = dbRelationUpdateCascade
        If del Then fl = fl Or dbRelationDeleteCascade
    End If
    Set rel = db.CreateRelation(nm, pT, cT, fl)
    Set fld = rel.CreateField(pF)
    fld.ForeignName = cF
    rel.Fields.Append fld
    db.Relations.Append rel
    Exit Sub
ErrR: Debug.Print "  Warning " & nm & ": " & Err.Description
End Sub

' ================================================================
'  4. QUERIES (14 total)
'     QRY_01, QRY_03, QRY_04, QRY_06, QRY_08-QRY_12: unchanged
'     QRY_02: frieze count now from Relief_Frieze (element M)
'     QRY_05: renames + mortar + opening metrics
'     QRY_07: Height_Above_Base_m
'     QRY_13: rewritten - full 24-column A-X matrix (letter order)
'     QRY_14: NEW - connection edge list for network analysis (OE3)
' ================================================================
Private Sub CreateAllQueries(db As DAO.Database)
    Dim qn(15) As String
    qn(0) = "QRY_01_Typology_by_Site"
    qn(1) = "QRY_02_Decoration_by_Site"
    qn(2) = "QRY_03_Conservation_by_Sector"
    qn(3) = "QRY_04_C14_Structures"
    qn(4) = "QRY_05_Export_RStats"
    qn(5) = "QRY_06_Volumetry_by_Typology"
    qn(6) = "QRY_07_Export_QGIS"
    qn(7) = "QRY_08_Children_of_Parent"
    qn(8) = "QRY_09_Group_Members"
    qn(9) = "QRY_10_ChaXR_Coverage"
    qn(10) = "QRY_11_Masonry_by_Site"
    qn(11) = "QRY_12_Geology_Construction"
    qn(12) = "QRY_13_AX_Pattern_Export"
    qn(13) = "QRY_14_Connections_Edges"
    qn(14) = "QRY_15_Observability_Bias"
    qn(15) = "QRY_16_Validation_Check"
    Dim i As Integer
    For i = 0 To 15
        If QueryExists(db, qn(i)) Then db.QueryDefs.Delete qn(i)
    Next i

    Dim q As String

    ' QRY_01 - Typology distribution by site
    q = "SELECT S.Site_Name, T.Name AS Typology, COUNT(E.ID) AS N "
    q = q & "FROM ((T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID) "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID "
    q = q & "GROUP BY S.Site_Name, T.Name ORDER BY S.Site_Name, T.Name;"
    db.CreateQueryDef qn(0), q

    ' QRY_02 - Decoration presence by site (chi-squared source)
    ' *** CRITICAL - READ BEFORE RUNNING ANY CHI-SQUARED TEST ***
    ' The Dec_* and Rock_Art fields are BYTE 0/1/9. The N_* counters below
    ' count value 1 only, so a 9 (not observable) is NOT counted as present.
    ' But "Total" counts EVERY structure, including those coded 9. Feeding
    ' N_x and Total straight into a chi-squared therefore inflates the
    ' denominator with structures that were never assessed, and understates
    ' the frequency of the motif. Because facade preservation differs between
    ' La Petaca and Diablo Wasi, that inflation is NOT equal across sites: the
    ' test would then measure differential preservation and the result would
    ' read as differential decorative practice - exactly the artefact the
    ' 0/1/9 domain exists to prevent.
    ' Three columns are provided per motif so the correct denominator can be
    ' chosen explicitly:
    '   N_x        = structures where the motif is PRESENT      (value 1)
    '   N_x_Absent = structures where it is VERIFIED ABSENT      (value 0)
    '   N_x_ND     = structures NOT OBSERVABLE                   (value 9)
    ' Valid denominator for the test = N_x + N_x_Absent (never Total).
    ' Total is retained only to report coverage: N_x_ND / Total is the share
    ' of the corpus excluded, and must be stated when reporting the test.
    ' v4: N_Frieze from Relief_Frieze (BYTE, element M; counts value 1 only)
    q = "SELECT S.Site_Name, "
    q = q & "SUM(IIF(E.Dec_Square_Niche=1,1,0)) AS N_Niche, "
    q = q & "SUM(IIF(E.Dec_Square_Niche=0,1,0)) AS N_Niche_Absent, "
    q = q & "SUM(IIF(E.Dec_Square_Niche=9,1,0)) AS N_Niche_ND, "
    q = q & "SUM(IIF(E.Dec_Relief_T=1,1,0)) AS N_T, "
    q = q & "SUM(IIF(E.Dec_Relief_T=0,1,0)) AS N_T_Absent, "
    q = q & "SUM(IIF(E.Dec_Relief_T=9,1,0)) AS N_T_ND, "
    q = q & "SUM(IIF(E.Dec_Relief_T_Inv=1,1,0)) AS N_T_Inv, "
    q = q & "SUM(IIF(E.Dec_Relief_T_Inv=0,1,0)) AS N_T_Inv_Absent, "
    q = q & "SUM(IIF(E.Dec_Relief_T_Inv=9,1,0)) AS N_T_Inv_ND, "
    q = q & "SUM(IIF(E.Dec_Relief_L=1,1,0)) AS N_L, "
    q = q & "SUM(IIF(E.Dec_Relief_L=0,1,0)) AS N_L_Absent, "
    q = q & "SUM(IIF(E.Dec_Relief_L=9,1,0)) AS N_L_ND, "
    q = q & "SUM(IIF(E.Dec_Relief_L_Inv=1,1,0)) AS N_L_Inv, "
    q = q & "SUM(IIF(E.Dec_Relief_L_Inv=0,1,0)) AS N_L_Inv_Absent, "
    q = q & "SUM(IIF(E.Dec_Relief_L_Inv=9,1,0)) AS N_L_Inv_ND, "
    q = q & "SUM(IIF(E.Dec_Zigzag=1,1,0)) AS N_Zigzag, "
    q = q & "SUM(IIF(E.Dec_Zigzag=0,1,0)) AS N_Zigzag_Absent, "
    q = q & "SUM(IIF(E.Dec_Zigzag=9,1,0)) AS N_Zigzag_ND, "
    q = q & "SUM(IIF(E.Dec_Stepped=1,1,0)) AS N_Stepped, "
    q = q & "SUM(IIF(E.Dec_Stepped=0,1,0)) AS N_Stepped_Absent, "
    q = q & "SUM(IIF(E.Dec_Stepped=9,1,0)) AS N_Stepped_ND, "
    q = q & "SUM(IIF(E.Relief_Frieze=1,1,0)) AS N_Frieze, "
    q = q & "SUM(IIF(E.Relief_Frieze=0,1,0)) AS N_Frieze_Absent, "
    q = q & "SUM(IIF(E.Relief_Frieze=9,1,0)) AS N_Frieze_ND, "
    q = q & "SUM(IIF(E.Rock_Art=1,1,0)) AS N_RockArt, "
    q = q & "SUM(IIF(E.Rock_Art=0,1,0)) AS N_RockArt_Absent, "
    q = q & "SUM(IIF(E.Rock_Art=9,1,0)) AS N_RockArt_ND, "
    q = q & "COUNT(E.ID) AS Total "
    q = q & "FROM (T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID "
    q = q & "GROUP BY S.Site_Name;"
    db.CreateQueryDef qn(1), q

    ' QRY_03 - Conservation status by sector
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
    db.CreateQueryDef qn(2), q

    ' QRY_04 - Structures with C14 dating
    q = "SELECT E.Code, S.Site_Name, SC.Sector_Name, T.Name AS Typology, "
    q = q & "E.Chrono_Start_Cent, E.Chrono_End_Cent "
    q = q & "FROM (((T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID) "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID) "
    q = q & "WHERE E.C14=True ORDER BY E.Chrono_Start_Cent;"
    db.CreateQueryDef qn(3), q

    ' QRY_05 - Full flat export for R/SPSS
    ' v4: Relief_Frieze replaces Dec_Frieze; + mortar & opening metrics
    q = "SELECT E.ID, E.Code, S.Site_Name AS Site, SC.Sector_Name AS Sector, "
    q = q & "T.Name AS Typology, SU.Name AS Support, "
    q = q & "AS1.Name AS Arch_Status, MS.Name AS Material_Status, "
    q = q & "E.N_Basal_Bodies, E.N_Chamber_Bodies, E.Lost_Body_Evidence, "
    q = q & "E.Floor_Plan, E.N_Built_Walls, "
    q = q & "E.Chamber_Roof_Type, "
    q = q & "E.Plaster_Present, E.Plaster_Extent, "
    q = q & "E.Pigment_Present, E.Pigment_Substrate, E.Pigment_Extent, "
    q = q & "E.Base_Level, E.Corbelled_Platform, E.Timber_Brackets, "
    q = q & "E.Transverse_Beams, E.Corbelled_Courses, E.Embedded_Base_Beams, "
    q = q & "E.Tie_Walls, E.Access_Opening, E.Sill, E.Recessed_Portal, "
    q = q & "E.Structural_Pilasters, E.Eave, E.Upper_Crown, E.Corner_Quoins, "
    q = q & "E.Lateral_Wall_Faces, E.Relief_Frieze, E.Interbody_Cornice, "
    q = q & "E.Dec_Square_Niche, E.Dec_Relief_T, E.Dec_Relief_T_Inv, "
    q = q & "E.Dec_Relief_L, E.Dec_Relief_L_Inv, E.Dec_Zigzag, "
    q = q & "E.Dec_Stepped, E.Rock_Art, "
    q = q & "E.Looting, E.Fire_Damage, E.Animal_Activity, E.Modern_Access, "
    q = q & "E.Human_Remains, E.MNI, E.Mummification, E.Funerary_Bundles, "
    q = q & "E.Bone_Burning, E.Mat_Textiles, E.Mat_Ceramics, E.Mat_DeerAntler, "
    q = q & "E.C14, E.Chrono_Start_Cent, E.Chrono_End_Cent, "
    q = q & "E.Interior_Area_m2, E.Interior_Vol_m3, E.Total_Vol_m3, "
    q = q & "E.Coord_Lat_WGS84, E.Coord_Lon_WGS84, "
    q = q & "E.Coord_E_UTM, E.Coord_N_UTM, E.Altitude_masl, "
    q = q & "E.ChaXR_Documented, "
    q = q & "E.Support_Width_cm, E.Support_Depth_cm, E.Support_Modified, "
    q = q & "E.Masonry_Quality, E.Masonry_Type, "
    q = q & "E.Mortar_Present, E.Mortar_Type, E.Chinking_Stones, "
    q = q & "E.Opening_Width_cm, E.Opening_Height_cm, "
    q = q & "E.Height_Above_Base_m, "
    q = q & "E.Facade_Orientation, E.Visibility_Valley, "
    q = q & "E.Construction_Phases, E.Phase_Evidence, "
    q = q & "E.Doc_Basis, E.Facade_Observability, E.Interior_Observability "
    q = q & "FROM (((((T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID) "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID) "
    q = q & "INNER JOIN L_SUPPORT AS SU ON E.ID_Support=SU.ID) "
    q = q & "LEFT JOIN L_STATUS AS AS1 ON E.ID_Arch_Status=AS1.ID) "
    q = q & "LEFT JOIN L_MATERIAL_STATUS AS MS ON E.ID_Material_Status=MS.ID;"
    db.CreateQueryDef qn(4), q

    ' QRY_06 - Volumetry by typology
    q = "SELECT S.Site_Name, T.Name AS Typology, "
    q = q & "COUNT(E.ID) AS N_Total, COUNT(E.Interior_Vol_m3) AS N_with_Vol, "
    q = q & "AVG(E.Interior_Area_m2) AS Mean_Area, "
    q = q & "AVG(E.Interior_Vol_m3) AS Mean_Vol, "
    q = q & "MIN(E.Interior_Vol_m3) AS Min_Vol, MAX(E.Interior_Vol_m3) AS Max_Vol "
    q = q & "FROM ((T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID) "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID "
    q = q & "WHERE E.Interior_Vol_m3 IS NOT NULL "
    q = q & "GROUP BY S.Site_Name, T.Name ORDER BY S.Site_Name, Mean_Vol DESC;"
    db.CreateQueryDef qn(5), q

    ' QRY_07 - Spatial export for QGIS (v4: Height_Above_Base_m)
    q = "SELECT E.ID, E.Code, S.Site_Name, SC.Sector_Name, T.Name AS Typology, "
    q = q & "E.Coord_Lat_WGS84, E.Coord_Lon_WGS84, "
    q = q & "E.Coord_E_UTM, E.Coord_N_UTM, E.Altitude_masl, "
    q = q & "E.Height_Above_Base_m, E.Coord_Precision_m, "
    q = q & "AS1.Name AS Arch_Status, MS.Name AS Material_Status, "
    q = q & "E.N_Basal_Bodies, E.N_Chamber_Bodies, E.Lost_Body_Evidence, "
    q = q & "E.Interior_Area_m2, E.Interior_Vol_m3, "
    q = q & "E.Chrono_Start_Cent, E.Chrono_End_Cent, "
    q = q & "E.Looting, E.Human_Remains, E.MNI, "
    q = q & "E.ChaXR_Documented, E.URL_3D, "
    q = q & "E.Facade_Orientation, E.Visibility_Valley, E.Masonry_Quality, "
    q = q & "E.Doc_Basis, E.Facade_Observability "
    q = q & "FROM ((((T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID) "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID) "
    q = q & "LEFT JOIN L_STATUS AS AS1 ON E.ID_Arch_Status=AS1.ID) "
    q = q & "LEFT JOIN L_MATERIAL_STATUS AS MS ON E.ID_Material_Status=MS.ID "
    q = q & "WHERE E.Coord_Lat_WGS84 IS NOT NULL ORDER BY S.Site_Name, E.Code;"
    db.CreateQueryDef qn(6), q

    ' QRY_08 - Children of a parent structure
    q = "SELECT E_c.Code AS Child_Code, T.Name AS Typology, "
    q = q & "E_c.Human_Remains, E_c.MNI "
    q = q & "FROM T_STRUCTURES AS E_c "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E_c.ID_Typology=T.ID "
    q = q & "WHERE E_c.ID_Parent=[Parent ID?] "
    q = q & "ORDER BY T.Name, E_c.Code;"
    db.CreateQueryDef qn(7), q

    ' QRY_09 - Members of a functional group
    q = "SELECT G.Group_Code, GT.Name AS Group_Type, E.Code, T.Name AS Typology, "
    q = q & "E.Coord_E_UTM, E.Coord_N_UTM, E.Altitude_masl "
    q = q & "FROM (((T_STRUCTURES AS E "
    q = q & "INNER JOIN T_GROUPS AS G ON E.ID_Group=G.ID) "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID) "
    q = q & "INNER JOIN L_GROUP_TYPE AS GT ON G.ID_Group_Type=GT.ID) "
    q = q & "WHERE G.Group_Code=[Group code?] "
    q = q & "ORDER BY E.Altitude_masl DESC;"
    db.CreateQueryDef qn(8), q

    ' QRY_10 - ChaXR documentation coverage
    q = "SELECT S.Site_Name, CA.Code AS Campaign, "
    q = q & "COUNT(E.ID) AS Total, "
    q = q & "SUM(IIF(E.ChaXR_Documented=True,1,0)) AS Published_ChaXR, "
    q = q & "SUM(IIF(E.URL_3D IS NOT NULL,1,0)) AS With_3D_Model "
    q = q & "FROM ((T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID) "
    q = q & "LEFT JOIN L_CAMPAIGN AS CA ON E.ID_Campaign=CA.ID "
    q = q & "GROUP BY S.Site_Name, CA.Code ORDER BY S.Site_Name, CA.Code;"
    db.CreateQueryDef qn(9), q

    ' QRY_11 - Masonry quality by site and typology (H01/H04)
    q = "SELECT S.Site_Name, T.Name AS Typology, "
    q = q & "E.Masonry_Quality, E.Masonry_Type, COUNT(E.ID) AS N "
    q = q & "FROM ((T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID) "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID "
    q = q & "WHERE E.Masonry_Quality IS NOT NULL "
    q = q & "GROUP BY S.Site_Name, T.Name, E.Masonry_Quality, E.Masonry_Type "
    q = q & "ORDER BY S.Site_Name, T.Name, E.Masonry_Quality;"
    db.CreateQueryDef qn(10), q

    ' QRY_12 - Geology-construction relationship (H02)
    ' v8: Support_Morphology was dropped as redundant; the geological support
    ' is now described by the ordered pair ID_Support (dominant class) +
    ' ID_Support_Secondary (second component of composite supports), which
    ' carries the same information without the risk of the two fields
    ' contradicting each other.
    q = "SELECT S.Site_Name, T.Name AS Typology, "
    q = q & "SU.Name AS Support_Primary, SU2.Name AS Support_Secondary, "
    q = q & "E.Support_Modified, COUNT(E.ID) AS N, "
    q = q & "AVG(E.Support_Width_cm) AS Avg_Width_cm, "
    q = q & "AVG(E.Support_Depth_cm) AS Avg_Depth_cm "
    q = q & "FROM ((((T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID) "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID) "
    q = q & "LEFT JOIN L_SUPPORT AS SU ON E.ID_Support=SU.ID) "
    q = q & "LEFT JOIN L_SUPPORT AS SU2 ON E.ID_Support_Secondary=SU2.ID "
    q = q & "GROUP BY S.Site_Name, T.Name, SU.Name, SU2.Name, E.Support_Modified "
    q = q & "ORDER BY S.Site_Name, T.Name;"
    db.CreateQueryDef qn(11), q

    ' QRY_13 - A-X pattern export for R (v4: full 24-column matrix)
    ' Letter-ordered vector AX_A..AX_X with 0/1/9 values.
    ' AX_Q derived from Lintel material: Absent->0, ND/Null->9, else->1.
    ' In R: filter or weight cells with value 9 (not observable) before
    ' computing phi/Jaccard, hierarchical cluster, Moran I or CA.
    ' LEFT JOIN on L_TYPOLOGY: includes structures without typology.
    q = "SELECT E.ID, E.Code, S.Site_Name AS Site, SC.Sector_Name AS Sector, "
    q = q & "T.Name AS Typology, "
    q = q & "E.Coord_E_UTM, E.Coord_N_UTM, E.Altitude_masl, E.ID_Group, "
    q = q & "E.N_Basal_Bodies, E.N_Chamber_Bodies, E.Lost_Body_Evidence, "
    q = q & "E.Embedded_Base_Beams AS AX_A, "
    q = q & "E.Base_Level AS AX_B, "
    q = q & "E.Decorative_Socle AS AX_C, "
    q = q & "E.Tie_Walls AS AX_D, "
    q = q & "E.Timber_Brackets AS AX_E, "
    q = q & "E.Transverse_Beams AS AX_F, "
    q = q & "E.Corbelled_Courses AS AX_G, "
    q = q & "E.Corbelled_Platform AS AX_H, "
    q = q & "E.Interbody_Cornice AS AX_I, "
    q = q & "E.Corner_Quoins AS AX_J, "
    q = q & "E.Structural_Pilasters AS AX_K, "
    q = q & "E.Lateral_Wall_Faces AS AX_L, "
    q = q & "E.Relief_Frieze AS AX_M, "
    q = q & "E.Sill AS AX_N, "
    q = q & "E.Jambs AS AX_O, "
    q = q & "E.Access_Opening AS AX_P, "
    q = q & "IIF(E.Lintel='Absent',0,IIF(E.Lintel Is Null Or E.Lintel='ND',9,1)) AS AX_Q, "
    q = q & "E.Upper_Crown AS AX_R, "
    q = q & "E.Eave_Beam AS AX_S, "
    q = q & "E.Eave_Surface AS AX_T, "
    q = q & "E.Eave AS AX_U, "
    q = q & "E.Lateral_Walls AS AX_V, "
    q = q & "E.Rear_Wall AS AX_W, "
    q = q & "E.Chamber_Roof AS AX_X, "
    q = q & "E.Timber_Bracket_Count, E.Timber_Bracket_Role, "
    q = q & "E.Platform_Surface_Material, "
    q = q & "E.Interbody_Cornice_Material, E.Lintel, "
    q = q & "E.Rear_Wall_Type, E.Recessed_Portal, "
    q = q & "E.Masonry_Quality, E.Masonry_Type, "
    q = q & "E.Mortar_Present, E.Chinking_Stones, "
    q = q & "E.Plaster_Present, E.Pigment_Present, E.Pigment_Substrate, "
    q = q & "E.Support_Modified, E.Chamber_Roof_Type, "
    q = q & "E.Doc_Basis, E.Facade_Observability, E.Interior_Observability "
    q = q & "FROM ((T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID) "
    q = q & "LEFT JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID "
    q = q & "ORDER BY S.Site_Name, SC.Sector_Name, E.Code;"
    db.CreateQueryDef qn(12), q

    ' QRY_14 - NEW v4: connection edge list (OE3 aerial network)
    ' Exports the adjacency list with coordinates of both endpoints,
    ' ready for igraph (R) or line generation in QGIS.
    q = "SELECT C.ID, A.Code AS Code_A, B.Code AS Code_B, "
    q = q & "C.Connection_Type, C.Confidence, "
    q = q & "A.Coord_E_UTM AS E_UTM_A, A.Coord_N_UTM AS N_UTM_A, "
    q = q & "A.Altitude_masl AS Alt_A, "
    q = q & "B.Coord_E_UTM AS E_UTM_B, B.Coord_N_UTM AS N_UTM_B, "
    q = q & "B.Altitude_masl AS Alt_B, "
    q = q & "C.Notes "
    q = q & "FROM (T_CONNECTIONS AS C "
    q = q & "INNER JOIN T_STRUCTURES AS A ON C.ID_Struct_A=A.ID) "
    q = q & "INNER JOIN T_STRUCTURES AS B ON C.ID_Struct_B=B.ID "
    q = q & "ORDER BY C.Connection_Type, A.Code;"
    db.CreateQueryDef qn(13), q


    ' QRY_15 - NEW v7: documentation and observability bias report (OE1)
    ' The 0/1/9 domain records WHERE the gaps are; this query quantifies them,
    ' so the sample can be defended rather than merely disclaimed. Use it to
    ' justify analytical subsets in R (e.g. restrict to Facade_Observability
    ' = 'Complete' before computing A-X co-occurrence).
    q = "SELECT S.Site_Name, SC.Sector_Name, "
    q = q & "E.Doc_Basis, E.Facade_Observability, E.Interior_Observability, "
    q = q & "COUNT(E.ID) AS N "
    q = q & "FROM (T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID "
    q = q & "GROUP BY S.Site_Name, SC.Sector_Name, E.Doc_Basis, "
    q = q & "E.Facade_Observability, E.Interior_Observability "
    q = q & "ORDER BY S.Site_Name, SC.Sector_Name;"
    db.CreateQueryDef qn(14), q


    ' QRY_16 - NEW v8: data validation battery (non-blocking report)
    ' Lists records that violate a coherence rule. Deliberately a REPORT and
    ' not a table-level constraint: a hard rule would block the researcher from
    ' recording a genuinely observed absence in a partially collapsed structure,
    ' and an Access table validation rule would have to hard-code the numeric ID
    ' of the 'Collapsed' status, which is brittle. Run it periodically during
    ' data entry; an empty result set means the corpus is coherent.
    q = "SELECT E.Code AS Structure, "
    q = q & "'A-X coded 0 (absent) although the structure is Collapsed' AS Rule_Violated, "
    q = q & "'Absence cannot be verified on a collapsed structure: use 9 (ND)' AS Action "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "INNER JOIN L_STATUS AS ST ON E.ID_Arch_Status=ST.ID "
    q = q & "WHERE ST.Name='Collapsed' AND ("
    q = q & "E.Embedded_Base_Beams=0 Or E.Base_Level=0 Or E.Decorative_Socle=0 Or E.Tie_Walls=0 Or E.Timber_Brackets=0 Or E.Transverse_Beams=0 Or "
    q = q & "E.Corbelled_Courses=0 Or E.Corbelled_Platform=0 Or E.Interbody_Cornice=0 Or E.Corner_Quoins=0 Or E.Structural_Pilasters=0 Or E.Lateral_Wall_Faces=0 Or "
    q = q & "E.Relief_Frieze=0 Or E.Sill=0 Or E.Jambs=0 Or E.Access_Opening=0 Or E.Upper_Crown=0 Or E.Lateral_Walls=0 Or E.Rear_Wall=0 Or E.Eave_Beam=0 Or "
    q = q & "E.Eave_Surface=0 Or E.Eave=0 Or E.Chamber_Roof=0"
    q = q & ") "
    q = q & "UNION ALL "
    q = q & "SELECT E.Code, "
    q = q & "'Pigment present but substrate not recorded', "
    q = q & "'Set Pigment_Substrate: the plaster / masonry distinction is the point' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Pigment_Present=1 AND (E.Pigment_Substrate Is Null Or E.Pigment_Substrate='ND') "
    q = q & "UNION ALL "
    q = q & "SELECT E.Code, "
    q = q & "'Pigment on plaster but plaster recorded as absent', "
    q = q & "'Reconcile Plaster_Present with Pigment_Substrate' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Pigment_Substrate='Plaster' AND E.Plaster_Present=0 "
    q = q & "UNION ALL "
    q = q & "SELECT E.Code, "
    q = q & "'Chamber roof present but type not recorded', "
    q = q & "'Set Chamber_Roof_Type: natural bedrock vs built is the decision' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Chamber_Roof=1 AND (E.Chamber_Roof_Type Is Null Or E.Chamber_Roof_Type='ND') "
    q = q & "UNION ALL "
    q = q & "SELECT E.Code, "
    q = q & "'Rear wall present but type not recorded', "
    q = q & "'Set Rear_Wall_Type: built masonry vs natural bedrock' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Rear_Wall=1 AND (E.Rear_Wall_Type Is Null Or E.Rear_Wall_Type='ND') "
    q = q & "UNION ALL "
    q = q & "SELECT E.Code, "
    q = q & "'System H present but no component (E, F or G) recorded', "
    q = q & "'E+F+G -> H: a platform needs at least one support element' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Corbelled_Platform=1 AND E.Timber_Brackets=0 "
    q = q & "AND E.Transverse_Beams=0 AND E.Corbelled_Courses=0 "
    q = q & "UNION ALL "
    q = q & "SELECT E.Code, "
    q = q & "'System P present but no component (N, O or Q) recorded', "
    q = q & "'N+O+Q -> P: an opening needs sill, jambs or lintel' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Access_Opening=1 AND E.Sill=0 AND E.Jambs=0 AND E.Lintel='Absent' "
    q = q & "UNION ALL "
    q = q & "SELECT E.Code, "
    q = q & "'Lost body inferred but upper-zone elements coded 0 (absent)', "
    q = q & "'Elements of a vanished body are not observable: use 9 (ND)' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Lost_Body_Evidence Is Not Null "
    q = q & "AND E.Lost_Body_Evidence<>'None' AND E.Lost_Body_Evidence<>'ND' "
    q = q & "AND (E.Upper_Crown=0 Or E.Eave=0 Or E.Eave_Beam=0 "
    q = q & "Or E.Eave_Surface=0 Or E.Chamber_Roof=0) "
    q = q & "UNION ALL "
    q = q & "SELECT E.Code, "
    q = q & "'Chamber bodies recorded but no access opening', "
    q = q & "'A body counts as N1 only if it has (or had) an access opening' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.N_Chamber_Bodies>0 AND E.Access_Opening=0 "
    q = q & "UNION ALL "
    q = q & "SELECT E.Code, "
    q = q & "'Timber corbels present but their role is not recorded', "
    q = q & "'Set Timber_Bracket_Role: platform support or isolated corbel' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Timber_Brackets=1 "
    q = q & "AND (E.Timber_Bracket_Role Is Null Or E.Timber_Bracket_Role='ND') "
    q = q & "ORDER BY Rule_Violated, Structure;"
    db.CreateQueryDef qn(15), q

    Debug.Print "-> 16 queries OK"
End Sub
