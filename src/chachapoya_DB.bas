Option Compare Database
Option Explicit

' ================================================================
'  CHACHAPOYA ARCHAEOLOGICAL DATABASE - COMPLETE BUILD SCRIPT v14
'  La Petaca & Diablo Wasi (Leymebamba, Amazonas, Peru)
'  Author: Esteve Ribera Torro | TFM Arqueologia UA
'  Spec: DELTA_v12_v13.md (rev. 5)
'
'  Run Sub BuildDB() on a NEW BLANK ACCESS DATABASE.
'  Then run chachapoya_Form_v14_val.bas -> Sub BuildForm()
'
'  NO MIGRATION PATH IN v13. The 35 v12 records are re-entered by
'  hand into a database built by this script. That decision removes
'  the whole migration apparatus (rule B, FK-safe UPDATE+INSERT,
'  reassignment of the six Fissure/Crack records, normalisation of
'  the stray 'Arch. elements' value) and, more importantly, removes
'  the 9s inherited from the v10 default: under v13 every stored
'  value is a deliberate judgement, which is the precondition for
'  publishing any percentage.
'
'  WHAT CHANGED FROM v10
'
'  1. FIVE-VALUE ELEMENT DOMAIN (1.1). The 20 A-X element fields go
'     from 0/1/9 to 0/1/2/3/9: Absent / Present complete / Present
'     partial / Attested lost / Not observable. The gradient is what
'     lets "what was built" (1+2+3), "what survives" (1+2) and
'     "intact" (1) be counted separately. Default is now 0, not 9:
'     under the new domain a new record starts from "examined,
'     nothing there" and 9 is claimed explicitly.
'     The other 24 observational fields keep three values (1.6):
'     "present partial" and "attested lost" are morphological
'     categories of built elements, and forcing them onto processes
'     (Looting) or portable remains (Mat_Textiles) means nothing.
'
'  2. H, P AND U WERE NEVER ELEMENTS, THEY WERE SYSTEMS (4.1).
'     Corbelled_Platform, Access_Opening and Eave are replaced by
'     Sys_Platform (E+F+G), Sys_Portal (N+O+Q) and Sys_Eave (S+T),
'     TEXT(20) over a six-value domain that adds Not applicable to
'     the gradient. Sys_Base and Sys_Chamber join them as grouping
'     fields for gating. The distinction between "not applicable"
'     and "verified absent" is the same logic as 0 vs 9 raised to
'     the level of the system - without it the published
'     percentages are wrong (4.3).
'
'  3. LINTEL SPLIT (5.4). The v10 field was a material lookup that
'     conflated presence with type. Now Lintel BYTE (element Q) plus
'     Lintel_Material TEXT(20), the pattern Chamber_Roof /
'     Chamber_Roof_Type already used.
'
'  4. DECORATION BOOLEANS DROPPED (7). The 12 Dec_*/RA_*/Rock_Art
'     flags duplicated T_DECORATIONS, which also records position,
'     type, colour and substrate - and with only 35 records the two
'     already disagreed in 6 cases. QRY_02 is rebuilt on
'     T_DECORATIONS. Relief_Frieze stays: it is an architectural
'     element, not a decoration (7.3).
'
'  5. T_LOST_ELEMENTS + L_LOST_EVIDENCE (2). Value 3 is the only one
'     in the domain that is an inference rather than an observation;
'     without a record of the physical evidence it is not
'     replicable. It is also the evidential basis of OE3.
'
'  6. L_ELEMENTS (3). The A-X vocabulary existed only as field
'     names: not referenceable, not exportable. Now a lookup, with a
'     unique index on Code so T_LOST_ELEMENTS can point at it.
'
'  7. RECORD_CLASS ON L_TYPOLOGY (6.1). Of the 35 current records, 5
'     are not structures; any percentage over 35 is wrong. MIX is
'     gone (every EA-CAM is mixed by definition, so it distinguished
'     nothing) and ND is split into Unclassifiable (a result) and
'     Not yet classified (a working status).
'
'  8. RENAMES: Lateral_Walls -> Return_Wall and Lateral_Wall_Faces ->
'     Facade_Flank (both v10 names said "lateral" about different
'     things), Recessed_Portal -> Recessed_Frame, Rear_Wall_Type ->
'     Rear_Closure_Type. Rear_Wall itself is dropped, redundant with
'     Sys_Chamber (5.5).
'
'  WHAT CHANGED FROM v11 (v12)
'
'  9. DEC_PRESENT (v12). Dropping the decoration booleans left the
'     ABSENCE of decoration unrecordable: no rows in T_DECORATIONS
'     could mean verified absence (0) or not observable (9), which
'     breaks the observability axiom the whole domain is built on.
'     Dec_Present BYTE 0/1/9 restores the aggregate judgement; the
'     detail still lives ONLY in T_DECORATIONS. Rules 22-23 keep
'     the two in agreement.
'
'  10. RULES 22-26 (v12). QRY_16c adds: Dec_Present coherence (22,
'     23), plaster/pigment detail under presence 0 or 9 (24, 25),
'     and Mortar_Type incompatible with Mortar_Present = 0 (26).
'     They are the corpus-level counterpart of the v12 form gating:
'     the form PREVENTS at entry, the battery DETECTS what got in
'     around the form (raw table edits, imports, gating disabled).
'
'  WHAT CHANGED FROM v12 (v13) - see DELTA_v12_v13.md
'
'  11. L_SUPPORT: FISSURE/CRACK SPLIT (delta 2). The single value
'      fused two geological phenomena with OPPOSITE structural
'      logics: a vertical cleft CONTAINS (the rock acts as
'      formwork, confinement) while a bedding-plane recess ANCHORS
'      (a continuous slot for embedding A or E, working in shear
'      and moment). Sharing one cell was an accident of vocabulary,
'      not an affinity. The mechanical mode is now written into
'      Description rather than held in a separate field: it is
'      deducible from the form's name, so a Support_Mode column
'      would be derivable - not information, but an opportunity for
'      contradiction.
'
'  12. DEFAULT NULL WHERE NO GATING NEEDS A ZERO (delta 8.2). The
'      v11 default 0 existed because rule 4 REQUIRES the padding
'      zero on components of a non-applicable system. Fields
'      outside any system inherited it for uniformity, with no
'      reason of their own. A false 0 ASSERTS TOO MUCH: it inflates
'      the denominator with unverified absences, and since that
'      bias correlates with observation quality - which differs
'      between LP and DW - it reintroduces exactly the artefact the
'      whole domain exists to prevent. NULL, not 9: 9 is an
'      assertion ("this position cannot be examined"), so using it
'      as a default would make an untouched field indistinguishable
'      from a deliberate one and would disable rule 17.
'
'  13. FIVE-VALUE DOMAIN EXTENDED TO APPLIED LAYERS (delta 8.3).
'      Mortar_Present, Plaster_Present, Pigment_Present,
'      Dec_Present and RockArt_Present move to 0/1/2/3/9. NOT a
'      blanket promotion: these are the only three-value fields
'      that record a LAYER APPLIED TO THE FABRIC, so they have
'      physical integrity, degrade gradually and leave a trace when
'      lost. L_LOST_EVIDENCE already carried 'Mortar imprint' - the
'      evidence catalogue anticipated a case the field domain could
'      not express. Processes (Looting), portable remains (Mat_*)
'      and qualifiers keep three values: "partially present" and
'      "attested lost" mean nothing there.
'      GENERAL RULE: five values for fabric or layers on fabric;
'      three for processes, portable remains and qualifiers.
'
'  14. ROCK ART INTEGRATED INTO T_DECORATIONS (delta 7). Associated
'      rock art becomes rows of the SAME table as architectural
'      decoration - no second table, which would duplicate the
'      structure, the relationships and the validation rules, and
'      force a UNION whenever the whole set is wanted. The only
'      real obstacle was position: L_STRUCT_BODY held architectural
'      positions only, so a painting on bedrock landed on 'ND',
'      which means "position not determined", not "position not
'      architectural". Three ROC entries fix it, and Level_Type
'      becomes the analytical discriminator: separating the two
'      sets is WHERE Level_Type = 'ROC' or its complement.
'
'  15. Color_Secondary on T_DECORATIONS, 'Both' retired (delta
'      7.6); RockArt_Present beside Dec_Present (7.5); Sort_Order
'      on L_STRUCT_BODY so the list follows the bottom-to-top
'      vocabulary instead of alphabetical order within each level;
'      QRY_17_RockArt_All with a Context column (7.7); rules 27-30
'      and rules 22-23 restricted to non-ROC rows.
'
'  22 tables | T_STRUCTURES: 121 fields | 25 relationships | 27 queries
' ================================================================

Sub BuildDB()
    Dim db As DAO.Database
    Set db = CurrentDb()
    CreateAllTables db
    SetByteDefaults db
    SetTextDefaults db
    PopulateAllLookups db
    CreateAllRelationships db
    CreateAllQueries db
    db.TableDefs.Refresh
    db.QueryDefs.Refresh
    Set db = Nothing

    Dim msg As String
    msg = "DATABASE v14 BUILT SUCCESSFULLY!" & vbCrLf & vbCrLf
    msg = msg & "  22 tables | 25 relationships | 28 queries" & vbCrLf
    msg = msg & "  T_STRUCTURES: 128 fields" & vbCrLf
    msg = msg & "  46 observational BYTE fields" & vbCrLf & vbCrLf & vbCrLf
    msg = msg & "Key changes (v11):" & vbCrLf
    msg = msg & "  20 element fields: five-value domain 0/1/2/3/9" & vbCrLf
    msg = msg & "  H, P, U are now Sys_Platform / Sys_Portal / Sys_Eave" & vbCrLf
    msg = msg & "  Sys_Base and Sys_Chamber added for gating" & vbCrLf
    msg = msg & "  Lintel split into Lintel + Lintel_Material" & vbCrLf
    msg = msg & "  Decoration booleans dropped: use T_DECORATIONS" & vbCrLf
    msg = msg & "  T_LOST_ELEMENTS records the evidence behind value 3" & vbCrLf
    msg = msg & "  L_ELEMENTS: the A-X vocabulary as a real lookup" & vbCrLf
    msg = msg & "  Record_Class on L_TYPOLOGY: filter every analysis" & vbCrLf
    msg = msg & "  v12: Dec_Present + validation rules 22-26" & vbCrLf
    msg = msg & "  v13: Fissure/Crack split into cleft + bedding plane" & vbCrLf
    msg = msg & "  v13: applied layers now 0/1/2/3/9 (mortar, plaster," & vbCrLf
    msg = msg & "       pigment, decoration, rock art)" & vbCrLf
    msg = msg & "  v13: rock art as ROC rows of T_DECORATIONS" & vbCrLf
    msg = msg & "  v13: rules 27-30 | QRY_17_RockArt_All" & vbCrLf
    msg = msg & "  v14: five-value domain = the 20 A-X elements ONLY" & vbCrLf
    msg = msg & "  v14: Cultural_Materials_Present | Doc_Basis ordinal" & vbCrLf
    msg = msg & "  v14: all linear fields in METRES | Ground support" & vbCrLf
    msg = msg & "  v14: C14 gates the dating subform | rule 31" & vbCrLf & vbCrLf
    msg = msg & "TWO DEFAULTS, DELIBERATELY:" & vbCrLf
    msg = msg & "  0 on element fields governed by a Sys_* (rule 4" & vbCrLf
    msg = msg & "    needs the padding zero)" & vbCrLf
    msg = msg & "  NULL everywhere else: empty = not yet assessed," & vbCrLf
    msg = msg & "    9 = assessed and not examinable, 0 = assessed" & vbCrLf
    msg = msg & "    and absent. Rule 17 lists what is still NULL." & vbCrLf & vbCrLf
    msg = msg & "Next: run chachapoya_Form_v13_val.bas -> BuildForm()"
    MsgBox msg, vbInformation, "Done!"
End Sub

' ================================================================
'  SHARED HELPERS
' ================================================================
Private Function TableExists(db As DAO.Database, n As String) As Boolean
    Dim t As DAO.TableDef
    For Each t In db.TableDefs
        If t.Name = n Then TableExists = True: Exit Function
    Next t
End Function

Private Function QueryExists(db As DAO.Database, n As String) As Boolean
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

' The 20 element fields with the system that gates each one
' (1.5 and 4.5). Single source of truth for the defaults, the
' QRY_13 NA gating and the QRY_16 rules, so they cannot disagree.
' 0 = letter, 1 = field, 2 = gating system ("" = none: I and R).
Private Sub FillElementMap(m() As String)
    ReDim m(19, 2)
    m(0, 0) = "A":  m(0, 1) = "Embedded_Base_Beams":  m(0, 2) = "Sys_Base"
    m(1, 0) = "B":  m(1, 1) = "Base_Level":           m(1, 2) = "Sys_Base"
    m(2, 0) = "C":  m(2, 1) = "Decorative_Socle":     m(2, 2) = "Sys_Base"
    m(3, 0) = "D":  m(3, 1) = "Tie_Walls":            m(3, 2) = "Sys_Base"
    m(4, 0) = "E":  m(4, 1) = "Timber_Brackets":      m(4, 2) = "Sys_Platform"
    m(5, 0) = "F":  m(5, 1) = "Transverse_Beams":     m(5, 2) = "Sys_Platform"
    m(6, 0) = "G":  m(6, 1) = "Corbelled_Courses":    m(6, 2) = "Sys_Platform"
    m(7, 0) = "I":  m(7, 1) = "Interbody_Cornice":    m(7, 2) = ""
    m(8, 0) = "J":  m(8, 1) = "Corner_Quoins":        m(8, 2) = "Sys_Chamber"
    m(9, 0) = "K":  m(9, 1) = "Structural_Pilasters": m(9, 2) = "Sys_Chamber"
    m(10, 0) = "L": m(10, 1) = "Facade_Flank":        m(10, 2) = "Sys_Chamber"
    m(11, 0) = "M": m(11, 1) = "Relief_Frieze":       m(11, 2) = "Sys_Chamber"
    m(12, 0) = "N": m(12, 1) = "Sill":                m(12, 2) = "Sys_Portal"
    m(13, 0) = "O": m(13, 1) = "Jambs":               m(13, 2) = "Sys_Portal"
    m(14, 0) = "Q": m(14, 1) = "Lintel":              m(14, 2) = "Sys_Portal"
    m(15, 0) = "R": m(15, 1) = "Upper_Crown":         m(15, 2) = ""
    m(16, 0) = "S": m(16, 1) = "Eave_Beam":           m(16, 2) = "Sys_Eave"
    m(17, 0) = "T": m(17, 1) = "Eave_Surface":        m(17, 2) = "Sys_Eave"
    m(18, 0) = "V": m(18, 1) = "Return_Wall":         m(18, 2) = "Sys_Chamber"
    m(19, 0) = "X": m(19, 1) = "Chamber_Roof":        m(19, 2) = "Sys_Chamber"
End Sub

' The 20 three-value fields (0/1/9). v13 moved five of the former
' 25 to the five-value domain (delta 8.3, FillLayerFields below):
' these are processes, portable remains and qualifiers, which have
' no physical integrity of their own, so "partially present" and
' "attested lost" would mean nothing on them.
' Extending this array automatically extends the BYTE defaults,
' QRY_16s_Observational_Nulls and rule 17: one list, one truth.
Private Sub FillThreeValueFields(f() As String)
    ReDim f(19)
    f(0) = "Recessed_Frame"
    f(1) = "Looting"
    f(2) = "Fire_Damage"
    f(3) = "Animal_Activity"
    f(4) = "Modern_Access"
    f(5) = "Human_Remains"
    f(6) = "Anatomical_Connection"
    f(7) = "Mummification"
    f(8) = "Funerary_Bundles"
    f(9) = "Dispersed_Remains"
    f(10) = "Flexed_Position"
    f(11) = "Bone_Burning"
    f(12) = "Mat_Textiles"
    f(13) = "Mat_Wood"
    f(14) = "Mat_VegFiber"
    f(15) = "Mat_Ceramics"
    f(16) = "Mat_Fauna"
    f(17) = "Mat_DeerAntler"
    f(18) = "Mat_Other"
    f(19) = "Support_Modified"
End Sub

' The remaining three-value fields: applied layers, aggregate
' judgements and qualifiers. Kept apart from FillThreeValueFields
' only so the arrays stay readable; both behave identically.
'
' WHY NO FIVE-VALUE PROMOTION HERE (rev. 5 corrects the v13 spec).
' The v13 criterion said "layers applied to the fabric", but what
' actually licenses value 3 is something else:
'   value 3 needs the evidence of loss to be of a DIFFERENT NATURE
'   from the thing lost - an empty socket is not a corbel, a mortar
'   imprint is not mortar;
'   value 2 needs the original extent to be inferable - on a wall
'   you can see how far it reached, on a painted surface you cannot.
' Pigment fails both: THE ONLY EVIDENCE OF PIGMENT IS PIGMENT, so 3
' and 0 necessarily collapse, and "present complete" is never
' assertable. Plaster likewise, and Plaster_Extent / Pigment_Extent
' ALREADY carry "Traces only" / "Traces" - value 2 duplicated them.
' Mortar is not a layer lost in patches but AN ATTRIBUTE OF THE
' MASONRY TECHNIQUE: a dry-laid wall is not dry-laid in places.
' What conservation changes is visibility, and that is what
' Facade_Observability and Mortar_Notes record - a Mortar_Present=2
' would describe the state of the 3D model, not the structure.
' Dec_Present and RockArt_Present are AGGREGATE judgements: every
' motif already has its own row with its own state.
'
' GENERAL RULE (rev. 5): the distinction is DISCRETE vs CONTINUOUS.
'   Five values: the 20 A-X elements, and only them.
'   Three values: everything else - technique attributes,
'   continuous layers, processes, portable remains, aggregates.
Private Sub FillLayerFields(f() As String)
    ReDim f(7)
    f(0) = "Mortar_Present"
    f(1) = "Plaster_Present"
    f(2) = "Pigment_Present"
    f(3) = "Dec_Present"
    f(4) = "RockArt_Present"
    f(5) = "Chinking_Stones"
    f(6) = "Recessed_Frame"
    f(7) = "Cultural_Materials_Present"
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
    ' v11: Record_Class drives the tab gating and every analytical filter (6.1)
    If Not TableExists(db, "L_TYPOLOGY") Then
        db.Execute "CREATE TABLE L_TYPOLOGY (ID COUNTER CONSTRAINT PK_TYP PRIMARY KEY, Name TEXT(60) NOT NULL, Name_VAL TEXT(60), Record_Class TEXT(30), Description MEMO)", dbFailOnError
    End If
    If Not TableExists(db, "L_SUPPORT") Then
        db.Execute "CREATE TABLE L_SUPPORT (ID COUNTER CONSTRAINT PK_SUP PRIMARY KEY, Name TEXT(60) NOT NULL, Name_VAL TEXT(60), Support_Mode TEXT(20), Description TEXT(255))", dbFailOnError
    End If
    If Not TableExists(db, "L_STATUS") Then
        db.Execute "CREATE TABLE L_STATUS (ID COUNTER CONSTRAINT PK_STA PRIMARY KEY, Name TEXT(30) NOT NULL, Name_VAL TEXT(30))", dbFailOnError
    End If
    If Not TableExists(db, "L_MATERIAL_STATUS") Then
        db.Execute "CREATE TABLE L_MATERIAL_STATUS (ID COUNTER CONSTRAINT PK_MS PRIMARY KEY, Name TEXT(50) NOT NULL, Name_VAL TEXT(50), Description TEXT(255))", dbFailOnError
    End If
    If Not TableExists(db, "L_VOL_METHOD") Then
        db.Execute "CREATE TABLE L_VOL_METHOD (ID COUNTER CONSTRAINT PK_VM PRIMARY KEY, Name TEXT(50) NOT NULL, Name_VAL TEXT(50), Description TEXT(255))", dbFailOnError
    End If
    If Not TableExists(db, "L_COORD_METHOD") Then
        db.Execute "CREATE TABLE L_COORD_METHOD (ID COUNTER CONSTRAINT PK_CM PRIMARY KEY, Name TEXT(50) NOT NULL, Name_VAL TEXT(50), Description TEXT(255))", dbFailOnError
    End If
    If Not TableExists(db, "L_GROUP_TYPE") Then
        db.Execute "CREATE TABLE L_GROUP_TYPE (ID COUNTER CONSTRAINT PK_GT PRIMARY KEY, Name TEXT(50) NOT NULL, Name_VAL TEXT(50), Description MEMO)", dbFailOnError
    End If
    If Not TableExists(db, "L_CAMPAIGN") Then
        db.Execute "CREATE TABLE L_CAMPAIGN (ID COUNTER CONSTRAINT PK_CAM PRIMARY KEY, Code TEXT(4) NOT NULL, Campaign_Name TEXT(80), Description MEMO)", dbFailOnError
    End If

    ' NEW v11: the A-X vocabulary as a real entity (3). UQ_ELEM_CODE is
    ' not decoration: JET only accepts a relationship against a primary
    ' key or a unique index, and T_LOST_ELEMENTS.Element_Code needs one.
    If Not TableExists(db, "L_ELEMENTS") Then
        sql = "CREATE TABLE L_ELEMENTS ("
        sql = sql & "ID COUNTER CONSTRAINT PK_ELEM PRIMARY KEY,"
        sql = sql & "Code TEXT(2) NOT NULL,"
        sql = sql & "Name_EN TEXT(60),"
        sql = sql & "Name_VAL TEXT(60),"
        sql = sql & "Level_Type TEXT(10),"
        sql = sql & "Sys_Group TEXT(20),"
        sql = sql & "Is_System YESNO,"
        sql = sql & "Field_Name TEXT(40),"
        sql = sql & "Description MEMO,"
        sql = sql & "CONSTRAINT UQ_ELEM_CODE UNIQUE (Code))"
        db.Execute sql, dbFailOnError
    End If

    ' NEW v11: the evidence vocabulary behind value 3 (2)
    If Not TableExists(db, "L_LOST_EVIDENCE") Then
        db.Execute "CREATE TABLE L_LOST_EVIDENCE (ID COUNTER CONSTRAINT PK_LEV PRIMARY KEY, Name TEXT(60) NOT NULL, Description TEXT(150))", dbFailOnError
    End If

    ' Decoration position lookup
    On Error Resume Next
    db.Execute "DROP TABLE L_STRUCT_BODY", dbFailOnError
    On Error GoTo 0
    db.Execute "CREATE TABLE L_STRUCT_BODY (ID COUNTER CONSTRAINT PK_SB PRIMARY KEY, Code TEXT(10) NOT NULL, Name TEXT(60) NOT NULL, Name_VAL TEXT(60), Level_Type TEXT(10), Sort_Order INTEGER, Description TEXT(255))", dbFailOnError

    ' Decoration type lookup
    On Error Resume Next
    db.Execute "DROP TABLE L_DEC_TYPE", dbFailOnError
    On Error GoTo 0
    db.Execute "CREATE TABLE L_DEC_TYPE (ID COUNTER CONSTRAINT PK_DT PRIMARY KEY, Name TEXT(60) NOT NULL, Name_VAL TEXT(60), Description TEXT(255))", dbFailOnError

    ' -- SECONDARY TABLES --
    If Not TableExists(db, "T_GROUPS") Then
        db.Execute "CREATE TABLE T_GROUPS (ID COUNTER CONSTRAINT PK_GRP PRIMARY KEY, Group_Code TEXT(20) NOT NULL, ID_Sector LONG NOT NULL, ID_Group_Type LONG NOT NULL, N_Members INTEGER, Notes MEMO)", dbFailOnError
    End If

    ' -- MAIN TABLE: T_STRUCTURES (119 fields) --
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
        ' --- 2. Morphology & dimensions (11) ---
        '     Lost_Body_Evidence is gone: lost bodies are recorded as
        '     T_LOST_ELEMENTS rows scoped Body, which also say WHAT the
        '     evidence was rather than only that there was some (2).
        sql = sql & "N_Basal_Bodies INTEGER,"
        sql = sql & "N_Chamber_Bodies INTEGER,"
        sql = sql & "Floor_Plan TEXT(20),"
        sql = sql & "N_Built_Walls INTEGER,"
        sql = sql & "Length_m SINGLE,"
        sql = sql & "Width_m SINGLE,"
        sql = sql & "Height_m SINGLE,"
        sql = sql & "Height_Above_Base_m SINGLE,"
        sql = sql & "Dim_Method TEXT(30),"
        sql = sql & "Opening_Width_m SINGLE,"
        sql = sql & "Opening_Height_m SINGLE,"
        ' --- 2b. Support geology detail (3) | H02 ---
        sql = sql & "Support_Width_m SINGLE,"
        sql = sql & "Support_Depth_m SINGLE,"
        sql = sql & "Support_Modified BYTE,"
        ' --- 2c. Masonry & mortar (6) | H01/H04 ---
        sql = sql & "Masonry_Quality TEXT(20),"
        sql = sql & "Masonry_Type TEXT(30),"
        sql = sql & "Mortar_Present BYTE,"
        sql = sql & "Mortar_Type TEXT(30),"
        sql = sql & "Chinking_Stones BYTE,"
        sql = sql & "Mortar_Notes TEXT(150),"
        ' --- 3. Constructive systems (6) | NEW v11, section 4 ---
        '     Six-value domain for the three element-systems:
        '     Present complete / Present partial / Attested lost /
        '     Absent / Not applicable / Not observable.
        '     Four values for the two grouping fields:
        '     Present / Absent / Not applicable / Not observable.
        '     Not applicable exports as NA and enters no denominator;
        '     Absent exports as 0, because an absence is a result.
        sql = sql & "Sys_Base TEXT(20),"
        sql = sql & "Sys_Platform TEXT(20),"
        sql = sql & "Sys_Portal TEXT(20),"
        sql = sql & "Sys_Eave TEXT(20),"
        sql = sql & "Sys_Chamber TEXT(20),"
        ' Form separated from function: naming the platform "access"
        ' in the data would presume the conclusion OE3 exists to reach
        sql = sql & "Platform_Function TEXT(20),"
        ' --- 3b. Level N0 elements (10) | A-G ---
        '     BYTE five-value: 0=Absent 1=Complete 2=Partial
        '     3=Attested lost 9=Not observable (default 0 via DAO)
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
        ' --- 3c. N0/N1 interface (2) | element I, gated by nothing ---
        sql = sql & "Interbody_Cornice BYTE,"
        sql = sql & "Interbody_Cornice_Material TEXT(20),"
        ' --- 4. Level N1 elements (12) | J-Q, R, V ---
        sql = sql & "Corner_Quoins BYTE,"
        sql = sql & "Structural_Pilasters BYTE,"
        sql = sql & "Facade_Flank BYTE,"
        sql = sql & "Relief_Frieze BYTE,"
        sql = sql & "Sill BYTE,"
        sql = sql & "Jambs BYTE,"
        sql = sql & "Lintel BYTE,"
        sql = sql & "Lintel_Material TEXT(20),"
        ' rev. 7: qualifier of the FACADE PLANE, not of the portal.
        ' The recess affects the whole wall face and the portal is
        ' inscribed in it, so gating it behind Sys_Portal blocked it
        ' precisely where there is a recess but no portal. It sits
        ' with Masonry_Type and Facade_Orientation on tab 2.Arq now.
        ' NOT promoted to a Sys_Facade: the five systems are either
        ' combinations of identifiable components (E+F+G, N+O+Q,
        ' S+T) or grouping fields for the gating. A facade is not a
        ' combination of anything - it is the plane on which
        ' everything else happens, and it could never be Absent. A
        ' system that cannot be absent does not do what systems do.
        sql = sql & "Recessed_Frame BYTE,"
        sql = sql & "Upper_Crown BYTE,"
        sql = sql & "Return_Wall BYTE,"
        ' Rear_Wall (W) is gone, redundant with Sys_Chamber; the type
        ' field survives because a chamber that uses the bedrock as its
        ' rear closure was not built like one that raises a wall (5.5)
        sql = sql & "Rear_Closure_Type TEXT(20),"
        ' --- 5. Upper zone (4) | S, T, X ---
        sql = sql & "Eave_Beam BYTE,"
        sql = sql & "Eave_Surface BYTE,"
        sql = sql & "Chamber_Roof BYTE,"
        sql = sql & "Chamber_Roof_Type TEXT(25),"
        ' --- 6. Surface treatments (7) ---
        sql = sql & "Plaster_Present BYTE,"
        sql = sql & "Plaster_Color TEXT(20),"
        sql = sql & "Plaster_Extent TEXT(20),"
        sql = sql & "Pigment_Present BYTE,"
        sql = sql & "Pigment_Substrate TEXT(20),"
        sql = sql & "Pigment_Color TEXT(20),"
        sql = sql & "Pigment_Extent TEXT(30),"
        ' --- 6b. Landscape & orientation (2) ---
        sql = sql & "Facade_Orientation TEXT(5),"
        sql = sql & "Visibility_Valley TEXT(10),"
        ' --- 7. Decoration (1): Dec_Present carries the aggregate
        '     0/1/9 judgement (v12); the DETAIL lives only in
        '     T_DECORATIONS (7.2). Rules 22-23 keep them coherent. ---
        sql = sql & "Dec_Present BYTE,"
        ' v13: rock art associated with the structure. Sibling of
        ' Dec_Present, not a substitute: with the two subforms of
        ' tab 4.Dec separated, one aggregate judgement each is what
        ' the user actually sees. Rules 28-29 keep each in step with
        ' its own half of T_DECORATIONS (ROC vs non-ROC rows).
        sql = sql & "RockArt_Present BYTE,"
        ' --- 8. Conservation (6) ---
        sql = sql & "ID_Arch_Status LONG,"
        ' rev. 5: ID_Material_Status mixed two variables in one
        ' scale - Good/Fair/Poor are degrees of preservation, but
        ' Absent is a statement of presence. Same error Lintel had
        ' in v10 (a material lookup conflating presence with type),
        ' and same fix. NOTE the asymmetry with ID_Arch_Status: a
        ' structure always exists, it IS the record. Portable
        ' remains may not, so they need a presence field that the
        ' architecture does not. Boundary: remains reduced to dust
        ' or unidentifiable fragments are Present=1 + Status=Poor;
        ' Present=0 is reserved for an empty chamber.
        sql = sql & "Cultural_Materials_Present BYTE,"
        sql = sql & "ID_Material_Status LONG,"
        sql = sql & "Looting BYTE,"
        sql = sql & "Fire_Damage BYTE,"
        sql = sql & "Animal_Activity BYTE,"
        sql = sql & "Modern_Access BYTE,"
        ' --- 9. Bioarchaeology (8) ---
        sql = sql & "Human_Remains BYTE,"
        sql = sql & "MNI INTEGER,"
        sql = sql & "Anatomical_Connection BYTE,"
        sql = sql & "Mummification BYTE,"
        sql = sql & "Funerary_Bundles BYTE,"
        sql = sql & "Dispersed_Remains BYTE,"
        sql = sql & "Flexed_Position BYTE,"
        sql = sql & "Bone_Burning BYTE,"
        ' --- 10. Cultural materials (7) ---
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
        ' rev. 7: ONE volume, ONE area. Interior_Vol / Total_Vol did
        ' not mean the same thing across typologies, which is worse
        ' than not having them: in a mausoleum the difference between
        ' the two IS the built fabric, but in a chamber inside a
        ' cavity the "interior" is natural space nobody excavated and
        ' the "total" includes bedrock, so the subtraction means
        ' nothing comparable. Averaging one column across both would
        ' have produced a meaningless figure.
        ' Volume_m3 is the FUNERARY SPACE - comparable across every
        ' typology and the one that relates to MNI (H01). Where a
        ' finer breakdown is possible (per body, platform surface,
        ' built volume) it goes in Vol_Notes: those are occasional
        ' cases, and a field would be empty in most records. If they
        ' turn out to be frequent, THAT is when to formalise them -
        ' same criterion as the pulley hypothesis.
        sql = sql & "Area_m2 SINGLE,"
        sql = sql & "Volume_m3 SINGLE,"
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
        ' rev. 5: Doc_Basis becomes an ORDINAL SCALE of documentary
        ' quality: Direct access / Close-range (<5m) / Medium-range
        ' (5-30m) / Long-range (>30m) / ND. The old list mixed
        ' access mode with capture technique; the technique is
        ' already described by the URLs on tab 10.Doc, so what this
        ' field adds is HOW CLOSE. The boundary between the first
        ' two values is PHYSICAL, not metric: on ropes you are
        ' obviously within 5 m, so what distinguishes Direct access
        ' is CONTACT - the thing that actually changes what can be
        ' recorded. Thresholds live in the labels so the criterion
        ' applies without consulting a manual.
        sql = sql & "Doc_Basis TEXT(30),"
        sql = sql & "Facade_Observability TEXT(20),"
        sql = sql & "Interior_Observability TEXT(20),"
        ' rev. 6: per-tab notes. A tab earns a notes field when it
        ' holds INTERPRETIVE JUDGEMENTS the domain cannot express -
        ' not every tab does: 9.Metr already has Vol_Notes and
        ' Dim_Method, masonry has Mortar_Notes, and the child tables
        ' carry their own Notes.
        ' They live on the tab they belong to, NOT on a single
        ' review tab: a note gets written when the case turns up,
        ' and the case turns up while filling in that tab. A
        ' read-only overview is QRY_18_Notes_Review instead.
        ' CAVEAT worth keeping in mind: notes are the black hole of
        ' the record. Anything written there stops being analysable,
        ' and the temptation to write instead of deciding the field
        ' value is strong. If the same kind of observation keeps
        ' recurring in the notes, that is the signal that a FIELD is
        ' needed - same criterion applied to the pulley hypothesis
        ' and to chromatic relation.
        sql = sql & "Arch_Notes MEMO,"
        sql = sql & "Finish_Notes MEMO,"
        sql = sql & "Condition_Notes MEMO,"
        sql = sql & "Bio_Notes MEMO,"
        sql = sql & "Materials_Notes MEMO,"
        sql = sql & "Systems_Notes MEMO,"
        sql = sql & "Notes MEMO)"
        db.Execute sql, dbFailOnError
        Debug.Print "[OK] T_STRUCTURES (119 fields)"
    End If

    ' -- LINKED TABLES --
    If Not TableExists(db, "T_DATING") Then
        db.Execute "CREATE TABLE T_DATING (ID COUNTER CONSTRAINT PK_DAT PRIMARY KEY, ID_Structure LONG NOT NULL, Sample_Type TEXT(30), Date_BP LONG, Sigma1_Start INTEGER, Sigma1_End INTEGER, Sigma2_Start INTEGER, Sigma2_End INTEGER, Lab_Reference TEXT(50), Bibliog_Reference MEMO)", dbFailOnError
    End If
    If Not TableExists(db, "T_INDIVIDUALS") Then
        db.Execute "CREATE TABLE T_INDIVIDUALS (ID COUNTER CONSTRAINT PK_IND PRIMARY KEY, ID_Structure LONG NOT NULL, Individual_No INTEGER, Age_Category TEXT(20), Sex_Category TEXT(20), Preservation TEXT(20), Notes MEMO)", dbFailOnError
    End If

    ' Decoration detail - now the ONLY record of decoration (7.4)
    On Error Resume Next
    db.Execute "DROP TABLE T_DECORATIONS", dbFailOnError
    On Error GoTo 0
        ' v13: Color_Secondary (delta 7.6). The documented bichrome cases
    ' have the two colours TOGETHER OR ADJACENT within a single motif
    ' (a pale field with a red border), so splitting them into two
    ' rows would invent two motifs where there is one and duplicate
    ' the position. Ordered pair instead - the same pattern as
    ' ID_Support + ID_Support_Secondary, and for the same reason: it
    ' says which one dominates. The old 'Both' value is retired
    ' (0 rows carried it in the v12 corpus, so nothing is lost).
    db.Execute "CREATE TABLE T_DECORATIONS (ID COUNTER CONSTRAINT PK_TDEC PRIMARY KEY, ID_Structure LONG NOT NULL, ID_Struct_Body LONG, ID_Dec_Type LONG, Body_No INTEGER, Color TEXT(20), Color_Secondary TEXT(20), Substrate TEXT(20), Notes TEXT(255))", dbFailOnError

    ' NEW v11: the evidence behind every value 3 (2). Element_Code is
    ' OPTIONAL on purpose: a vanished body or a razed structure has no
    ' single element code, and rev. 3's NOT NULL rejected exactly the
    ' rows the Body / Whole structure scopes exist to hold. The
    ' "mandatory if and only if scope = Element" rule lives in QRY_16
    ' rule 21 and in form validation, not in the DDL.
    On Error Resume Next
    db.Execute "DROP TABLE T_LOST_ELEMENTS", dbFailOnError
    On Error GoTo 0
    sql = "CREATE TABLE T_LOST_ELEMENTS ("
    sql = sql & "ID COUNTER CONSTRAINT PK_LOST PRIMARY KEY,"
    sql = sql & "ID_Structure LONG NOT NULL,"
    sql = sql & "Element_Code TEXT(2),"
    sql = sql & "ID_Evidence_Type LONG NOT NULL,"
    sql = sql & "Evidence_Scope TEXT(20),"
    sql = sql & "ID_Position LONG,"
    sql = sql & "Notes MEMO)"
    db.Execute sql, dbFailOnError
    ' Blank must mean NULL, never "": an empty string satisfies neither
    ' the foreign key nor rule 21, and is invisible in the datasheet.
    ' Refresh first and report on failure, for the same reason as the
    ' defaults: TableDefs is stale right after a CREATE, and a silent
    ' miss here would only surface as a broken foreign key much later.
    db.TableDefs.Refresh
    On Error GoTo Err_AZL
    db.TableDefs("T_LOST_ELEMENTS").Fields("Element_Code").AllowZeroLength = False
    GoTo Done_AZL
Err_AZL:
    Debug.Print "  *** AllowZeroLength NOT set on T_LOST_ELEMENTS.Element_Code: " & Err.Description
Done_AZL:
    On Error GoTo 0

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

    ' Physical connections between structures (OE3 network).
    ' v11: Chrono_Relation turns the edge list from an undirected graph
    ' into a potentially directed one, which is what H03 needs (9.1).
    On Error Resume Next
    db.Execute "DROP TABLE T_CONNECTIONS", dbFailOnError
    On Error GoTo 0
    sql = "CREATE TABLE T_CONNECTIONS ("
    sql = sql & "ID COUNTER CONSTRAINT PK_CON PRIMARY KEY,"
    sql = sql & "ID_Struct_A LONG NOT NULL,"
    sql = sql & "ID_Struct_B LONG NOT NULL,"
    sql = sql & "Connection_Type TEXT(30),"
    sql = sql & "Chrono_Relation TEXT(20),"
    sql = sql & "Confidence TEXT(10),"
    sql = sql & "Notes TEXT(150))"
    db.Execute sql, dbFailOnError

    Debug.Print "-> 22 tables OK"
End Sub

' ================================================================
'  1b. DEFAULT 0 (ABSENT) ON EVERY OBSERVATIONAL BYTE FIELD
'
'  Set via DAO, never in the DDL: JET only honours the DEFAULT clause
'  in DDL executed through ADO in ANSI-92 mode, and this whole script
'  uses db.Execute (DAO). Section 0 of the delta.
'
'  v10 defaulted to 9 so that a new record would not silently claim
'  absence. v11 inverts it: under the five-value domain 0 means "the
'  position was examined and there was nothing", which is the normal
'  finding, while 9 means the position cannot be examined at all -
'  a claim the researcher should have to make deliberately.
' ================================================================
Private Sub SetByteDefaults(db As DAO.Database)
    ' MUST refresh first. TableDefs was read before CreateAllTables
    ' ran, so without this every lookup of T_STRUCTURES raises "item
    ' not found" - and with a blanket On Error Resume Next the whole
    ' loop would fail silently while still reporting success. That is
    ' exactly the bug this script inherited from v10, where the
    ' defaults were most likely never set either.
    db.TableDefs.Refresh

    ' v13 (delta 8.2): TWO DEFAULTS, and the split is the point.
    '
    ' DEFAULT 0 on the 20 A-X element fields ONLY. They need it
    ' because rule 4 REQUIRES the padding zero on components of a
    ' non-applicable system, and the form's quick-fill writes it.
    '
    ' DEFAULT NULL everywhere else. Those fields belong to no
    ' system and have no padding to satisfy: they inherited the
    ' v11 default for uniformity, with no reason of their own.
    ' A false 0 asserts too much - it inflates the denominator with
    ' unverified absences, and the bias correlates with observation
    ' quality, which differs between LP and DW.
    '
    ' NULL AND NOT 9, deliberately. 9 is not an empty cell, it is
    ' an assertion: "this position cannot be examined". As a
    ' default it would make an untouched field indistinguishable
    ' from a deliberate one and would disable rule 17, whose whole
    ' job is listing what has not been assessed yet. The v12 corpus
    ' showed exactly this: 20 nines on Fire_Damage, inherited from
    ' the v10 default, none of them a real judgement.
    '
    ' Three states are preserved: NULL = not yet assessed,
    ' 9 = assessed and not examinable, 0 = assessed and absent.
    db.TableDefs.Refresh

    Dim m() As String
    Dim f() As String
    Dim g() As String
    FillElementMap m
    FillThreeValueFields f
    FillLayerFields g

    Dim i As Integer
    Dim ok As Integer
    Dim bad As Integer
    For i = 0 To 19
        If SetDef(db, "T_STRUCTURES", m(i, 1), "0") Then ok = ok + 1 Else bad = bad + 1
    Next i
    Debug.Print "-> Default 0 set on " & ok & " of 20 element fields | failures: " & bad

    ' The rest: DefaultValue cleared, so a new record starts empty.
    Dim nN As Integer
    For i = 0 To 19
        If ClearDef(db, "T_STRUCTURES", f(i)) Then nN = nN + 1
    Next i
    For i = 0 To 7
        If ClearDef(db, "T_STRUCTURES", g(i)) Then nN = nN + 1
    Next i
    Debug.Print "-> Default cleared (NULL) on " & nN & " of 28 non-gated observational fields"
    Debug.Print "-> Five-value domain: the 20 A-X elements, and only them"
End Sub

' Empties DefaultValue so a new record starts NULL on this field.
Private Function ClearDef(db As DAO.Database, tbl As String, fld As String) As Boolean
    On Error GoTo Err_CD
    db.TableDefs(tbl).Fields(fld).DefaultValue = ""
    ClearDef = True
    Exit Function
Err_CD:
    Debug.Print "   [default] ERROR clearing " & fld & ": " & Err.Description
End Function

' ================================================================
'  1c. TEXT DEFAULTS (1.7, 5.3, 9.1)
'  Absent on the five systems, coherent with the 0 of the elements;
'  Undetermined where the value is a judgement not yet made.
' ================================================================
Private Sub SetTextDefaults(db As DAO.Database)
    db.TableDefs.Refresh

    Dim ok As Integer
    Dim bad As Integer
    Dim q As String
    q = Chr(34)
    If SetDef(db, "T_STRUCTURES", "Sys_Platform", q & "Absent" & q) Then ok = ok + 1 Else bad = bad + 1
    If SetDef(db, "T_STRUCTURES", "Sys_Portal", q & "Absent" & q) Then ok = ok + 1 Else bad = bad + 1
    If SetDef(db, "T_STRUCTURES", "Sys_Eave", q & "Absent" & q) Then ok = ok + 1 Else bad = bad + 1
    If SetDef(db, "T_STRUCTURES", "Sys_Base", q & "Absent" & q) Then ok = ok + 1 Else bad = bad + 1
    If SetDef(db, "T_STRUCTURES", "Sys_Chamber", q & "Absent" & q) Then ok = ok + 1 Else bad = bad + 1
    If SetDef(db, "T_STRUCTURES", "Platform_Function", q & "Undetermined" & q) Then ok = ok + 1 Else bad = bad + 1
    If SetDef(db, "T_CONNECTIONS", "Chrono_Relation", q & "Undetermined" & q) Then ok = ok + 1 Else bad = bad + 1
    Debug.Print "-> TEXT defaults set on " & ok & " of 7 fields | failures: " & bad
End Sub

' Sets one DefaultValue and says whether it worked. Never report a
' count that was incremented regardless of the outcome: a defaults
' routine that lies is worse than one that does nothing, because the
' records it silently fails to protect look fine until analysis.
Private Function SetDef(db As DAO.Database, tbl As String, fld As String, val As String) As Boolean
    On Error GoTo Err_SD
    db.TableDefs(tbl).Fields(fld).DefaultValue = val
    SetDef = True
    Exit Function
Err_SD:
    Debug.Print "  *** default NOT set on " & tbl & "." & fld & ": " & Err.Description
End Function

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
    X db, "DELETE FROM L_ELEMENTS"
    X db, "DELETE FROM L_LOST_EVIDENCE"

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

    ' L_TYPOLOGY (6.1, 6.2, 6.3)
    ' MIX is gone: every EA-CAM is mixed by definition (natural cavity
    ' plus built facade), so the category distinguished nothing, and
    ' ID_Support plus the Sys_* fields already carry the combination.
    ' ND is split: Unclassifiable is a RESULT (evidence insufficient to
    ' classify a built structure, as at EA11), Not yet classified is a
    ' working status excluded from every analytical query.
    Dim t(9, 2) As String
    t(0, 0) = "EA-MAU Mausoleum/Chullpa": t(0, 1) = "Built funerary structure": t(0, 2) = "Built structure (3+ walls + artificial roof) on ledge. 1-3 storeys. Predominant at La Petaca."
    t(1, 0) = "EA-CAM Funerary Chamber":  t(1, 1) = "Built funerary structure": t(1, 2) = "Natural cavity closed by 1 built facade. Predominant at Diablo Wasi."
    t(2, 0) = "EA-PLA-R Ledge Platform":  t(2, 1) = "Built funerary structure": t(2, 2) = "Constructive platform on natural ledge. Function: transit or mausoleum base."
    t(3, 0) = "EA-PLA-V Aerial Platform": t(3, 1) = "Built funerary structure": t(3, 2) = "Artificial platform on wooden beams and slabs, without natural ledge support."
    t(4, 0) = "NIX Natural Niche":        t(4, 1) = "Natural funerary context": t(4, 2) = "Small natural cavity (<1m2). Function: ossuary or secondary burial."
    t(5, 0) = "CAV Cave/Cavern":          t(5, 1) = "Natural funerary context": t(5, 2) = "Large natural cavity (>1m2) with documented funerary or ritual use."
    t(6, 0) = "PR Rock Art":              t(6, 1) = "Rock art panel":           t(6, 2) = "Pictorial motif on rock, independently documented."
    t(7, 0) = "MEN Isolated Bracket":     t(7, 1) = "Structural trace":         t(7, 2) = "Isolated structural element. Evidence of lost aerial circulation network."
    t(8, 0) = "Unclassifiable":           t(8, 1) = "Built funerary structure": t(8, 2) = "Insufficient evidence to classify a BUILT structure (e.g. EA11: pigment perimeter with no surviving construction). Natural contexts are always classifiable as NIX or CAV."
    t(9, 0) = "Not yet classified":       t(9, 1) = "Pending classification":   t(9, 2) = "Working status: pending manual classification review. Excluded from every analytical query."
    For i = 0 To 9
        db.Execute "INSERT INTO L_TYPOLOGY (Name,Record_Class,Description) VALUES ('" & t(i, 0) & "','" & t(i, 1) & "','" & t(i, 2) & "')", dbFailOnError
    Next i

    ' L_SUPPORT - v13: Fissure/Crack split in two (delta 2).
    ' Support_Mode is written here rather than held as a field on
    ' T_STRUCTURES: the mechanical mode follows deterministically
    ' from the form's name, so a per-record column would be
    ' derivable - not information, but an opportunity for
    ' contradiction, and near-constant per form on 35 records.
    ' Writing it down is what makes the deduction legitimate: it
    ' fixes the criterion for the observer and leaves the derivation
    ' available to R as an aggregate variable.
    ' MICRO-LEDGE VS BEDDING-PLANE RECESS ARE THE SAME PHENOMENON
    ' SEEN FROM OPPOSITE SIDES: when a soft stratum erodes, the bed
    ' below is left projecting (micro-ledge: you rest on it) and a
    ' slot is left (recess: you embed into it). Which one it is
    ' depends solely on what the builder did, not on the rock -
    ' hence the criterion in each Description.
    Dim sup(10, 2) As String
    sup(0, 0) = "Wide natural ledge (>2m)":   sup(0, 1) = "Gravitational rest": sup(0, 2) = "Ledge over 2 m wide. The structure rests on it in compression."
    sup(1, 0) = "Narrow natural ledge (<2m)": sup(1, 1) = "Gravitational rest": sup(1, 2) = "Ledge under 2 m wide. The structure rests on it in compression."
    sup(2, 0) = "Artificial ledge":           sup(2, 1) = "Gravitational rest": sup(2, 2) = "Ledge built or enlarged by cutting. Cross-check with Support_Modified."
    sup(3, 0) = "Large cavity (>10m2)":       sup(3, 1) = "Gravitational rest": sup(3, 2) = "Cavity over 10 m2. The chamber rests on the cavity floor."
    sup(4, 0) = "Medium cavity (1-10m2)":     sup(4, 1) = "Gravitational rest": sup(4, 2) = "Cavity of 1-10 m2. The chamber rests on the cavity floor."
    sup(5, 0) = "Natural niche (<1m2)":       sup(5, 1) = "Gravitational rest": sup(5, 2) = "Cavity under 1 m2."
    sup(6, 0) = "Micro-ledge (<50cm)":        sup(6, 1) = "Gravitational rest": sup(6, 2) = "Projecting bed under 50 cm. Criterion: the structure RESTS on it. If it EMBEDS into the slot instead, record a bedding-plane recess."
    sup(7, 0) = "Vertical cleft":             sup(7, 1) = "Confinement":        sup(7, 2) = "Vertical joint / diaclase crossing the beds. Supplies two lateral rock faces, constrains the plan and is filled to generate the basal level: the rock acts as formwork."
    sup(8, 0) = "Bedding-plane recess":       sup(8, 1) = "Embedment":          sup(8, 2) = "Horizontal recess left by differential erosion of a soft stratum between competent beds. Supplies a continuous slot for embedding base beams (A) or corbels (E), working in shear and moment rather than compression."
    ' rev. 5: Ground makes explicit an assumption the other ten
    ' forms all share silently - that the structure sits ELEVATED
    ' OVER A VOID. Ground is its negation, and the operative test is
    ' negative and checkable: THERE IS NO VOID BELOW.
    ' It is also the only form on SEDIMENT rather than rock, which
    ' is not a nuance: on rock the structure does not settle and
    ' needs no footing; on sediment it does. That probably explains
    ' the presence or absence of embedded base beams (A) and tie
    ' walls (D). Entry guidance, deliberately NOT gated in case a
    ' case contradicts it: Height_Above_Base_m should be 0 or near
    ' it, and Sys_Platform should normally be Not applicable rather
    ' than Absent - corbels, beams, courses and eave all exist to
    ' solve problems of verticality.
    sup(9, 0) = "Ground":                     sup(9, 1) = "Gravitational rest": sup(9, 2) = "Ground surface at the cliff base or foot of the slope. The structure sits on soil, not elevated over a void. The only form in the catalogue with no vertical component: check before assigning it."
    sup(10, 0) = "ND":                        sup(10, 1) = "ND":                sup(10, 2) = "Support form not determined."
    For i = 0 To 10
        db.Execute "INSERT INTO L_SUPPORT (Name,Support_Mode,Description) VALUES ('" & sup(i, 0) & "','" & sup(i, 1) & "','" & Replace(sup(i, 2), "'", "''") & "')", dbFailOnError
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
    db.Execute "INSERT INTO L_MATERIAL_STATUS (Name,Description) VALUES ('Poor','Fragmentary: barely identifiable remains, heavily degraded or scattered. Use with Cultural_Materials_Present=1; absence is recorded there, not here.')", dbFailOnError
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
    ' rev. 6: the Xa/Xb convention lived ONLY in the code string, as
    ' a textual habit. Nothing made the relation queryable: counting
    ' supra-structural units meant text-matching on Code, which is
    ' fragile. This type makes the unit countable.
    ' NOTE the division of labour with T_CONNECTIONS: a connection is
    ' BINARY and describes the joint (abutted, bonded, superposition,
    ' and it carries the chronological direction); a group is N-ARY
    ' and describes the unit. Three bodies Xa/Xb/Xc give three
    ' connections but ONE group.
    db.Execute "INSERT INTO L_GROUP_TYPE (Name,Description) VALUES ('Supra-structural unit','Two or more coherent constructive units on a shared base, recorded separately (codes Xa, Xb...) but forming one built whole. Criterion: two units with attributes of their own, not necessarily two chambers - a mausoleum plus a corbelled platform qualifies.')", dbFailOnError
    db.Execute "INSERT INTO L_GROUP_TYPE (Name,Description) VALUES ('Functional group','Any other intra-sector grouping with functional coherence.')", dbFailOnError

    ' L_CAMPAIGN
    db.Execute "INSERT INTO L_CAMPAIGN (Code,Campaign_Name,Description) VALUES ('2013','PALP I','First campaign. Prospection and initial documentation of La Petaca. Dr. J. Marla Toyne (UCF).')", dbFailOnError
    db.Execute "INSERT INTO L_CAMPAIGN (Code,Campaign_Name,Description) VALUES ('2016','PALP II','Extension to Diablo Wasi. First systematic documentation of DW.')", dbFailOnError
    db.Execute "INSERT INTO L_CAMPAIGN (Code,Campaign_Name,Description) VALUES ('2021','La Petaca Project','Non-invasive integral documentation. Photogrammetry, 360, gigaphotos. Panograma Labs/UCF.')", dbFailOnError
    db.Execute "INSERT INTO L_CAMPAIGN (Code,Campaign_Name,Description) VALUES ('2023','PALP IV','Archaeological excavation campaign and detailed 3D reconstructions.')", dbFailOnError

    ' L_STRUCT_BODY - position within a body only.
    ' v11: LWF "Lateral wall face (L)" becomes FFL "Facade flank (L)",
    ' following the rename of element L (5.1).
    ' v13 (delta 7.2): three ROC entries, and Sort_Order.
    '
    ' WHY ROC AT ALL: T_DECORATIONS is the single record of surface
    ' treatment, and rock art associated with a structure belongs in
    ' it - a second table would duplicate the structure, the
    ' relationships and the validation rules, and force a UNION
    ' whenever the whole set is wanted. The only real obstacle was
    ' position: with architectural positions only, a painting on
    ' bedrock landed on 'ND', which means "position not determined",
    ' not "position not architectural". Two very different things
    ' collapsed into one value - the same class of error the 0/1/9
    ' domain exists to prevent.
    '
    ' LEVEL_TYPE = 'ROC' IS THE ANALYTICAL DISCRIMINATOR: separating
    ' architectural decoration from rock art is WHERE Level_Type =
    ' 'ROC' or its complement. No new column on T_DECORATIONS.
    '
    ' ROC-PAN exists for rows hanging off an ISOLATED rock art
    ' record (typology PR), which is itself a T_STRUCTURES row and
    ' therefore also describes its motifs through T_DECORATIONS.
    ' Without it those rows would fall back to 'ND' - the very
    ' problem being fixed.
    '
    ' SORT_ORDER: the list was ordered by Level_Type then Name, so
    ' Eave (SUP) fell last and each level ran alphabetically rather
    ' than constructively. The explicit order restores the
    ' bottom-to-top logic of the A-X vocabulary.
    Dim sb(13, 4) As String
    sb(0, 0) = "BAS":     sb(0, 1) = "Base level (B)":            sb(0, 2) = "N0":  sb(0, 3) = "10": sb(0, 4) = "Basal mass treated as a distinct constructive element."
    sb(1, 0) = "SOC":     sb(1, 1) = "Socle (C)":                 sb(1, 2) = "N0":  sb(1, 3) = "20": sb(1, 4) = "Decorative socle - treatment of the basal mass."
    sb(2, 0) = "COR":     sb(2, 1) = "Interbody cornice (I)":     sb(2, 2) = "N1":  sb(2, 3) = "30": sb(2, 4) = "Cornice zone between superposed constructive bodies."
    sb(3, 0) = "QUO":     sb(3, 1) = "Corner quoin (J)":          sb(3, 2) = "N1":  sb(3, 3) = "40": sb(3, 4) = "Larger stones set vertically at the angles."
    sb(4, 0) = "PIL":     sb(4, 1) = "Pilaster (K)":              sb(4, 2) = "N1":  sb(4, 3) = "50": sb(4, 4) = "Structural pilaster framing the facade, full height."
    sb(5, 0) = "FFL":     sb(5, 1) = "Facade flank (L)":          sb(5, 2) = "N1":  sb(5, 3) = "60": sb(5, 4) = "Wall face in the facade plane, flanking the opening."
    sb(6, 0) = "JAM":     sb(6, 1) = "Jamb (O)":                  sb(6, 2) = "N1":  sb(6, 3) = "70": sb(6, 4) = "Portal jamb - vertical frame element of the access opening."
    sb(7, 0) = "OVL":     sb(7, 1) = "Over-lintel":               sb(7, 2) = "N1":  sb(7, 3) = "80": sb(7, 4) = "Zone above the lintel - decorative frieze area (element M)."
    sb(8, 0) = "CRO":     sb(8, 1) = "Upper crown (R)":           sb(8, 2) = "N1":  sb(8, 3) = "90": sb(8, 4) = "Upper crown or coping of the body."
    sb(9, 0) = "EAV":     sb(9, 1) = "Eave (U)":                  sb(9, 2) = "SUP": sb(9, 3) = "100": sb(9, 4) = "Eave / roof overhang above the facade."
    sb(10, 0) = "ROC":    sb(10, 1) = "Adjacent bedrock":         sb(10, 2) = "ROC": sb(10, 3) = "200": sb(10, 4) = "Rock face adjacent to the structure, outside the built fabric. Use when tests 2 or 3 of the association criterion apply (delta 7.3)."
    sb(11, 0) = "ROC-PER": sb(11, 1) = "Opening perimeter (rock)": sb(11, 2) = "ROC": sb(11, 3) = "210": sb(11, 4) = "Marks framing a natural opening, threshold or cleft mouth. Test 3 of the association criterion."
    sb(12, 0) = "ROC-PAN": sb(12, 1) = "Panel (no architectural ref.)": sb(12, 2) = "ROC": sb(12, 3) = "220": sb(12, 4) = "For rows hanging off an isolated rock art record (typology PR), which has no architectural position to refer to."
    sb(13, 0) = "ND":     sb(13, 1) = "Not determined":           sb(13, 2) = "ND": sb(13, 3) = "999": sb(13, 4) = "Position not determined. NOT to be used for non-architectural positions: those are the ROC entries."
    For i = 0 To 13
        db.Execute "INSERT INTO L_STRUCT_BODY (Code,Name,Level_Type,Sort_Order,Description) VALUES ('" & sb(i, 0) & "','" & sb(i, 1) & "','" & sb(i, 2) & "'," & sb(i, 3) & ",'" & Replace(sb(i, 4), "'", "''") & "')", dbFailOnError
    Next i

    ' L_DEC_TYPE
    Dim dt(18, 1) As String
    dt(0, 0) = "T-shaped niche":      dt(0, 1) = "Niche or bas-relief in T form. Vertical + horizontal element."
    dt(1, 0) = "T-shaped niche inv.": dt(1, 1) = "Inverted T niche/relief."
    dt(2, 0) = "L-shaped niche":      dt(2, 1) = "Niche or bas-relief in L form."
    dt(3, 0) = "L-shaped niche inv.": dt(3, 1) = "Inverted L niche/relief."
    dt(4, 0) = "Zigzag":              dt(4, 1) = "Zigzag or chevron motif."
    dt(5, 0) = "Stepped motif":       dt(5, 1) = "Stepped/staircase motif. Rare at DW and LP."
    dt(6, 0) = "Frieze / Greca":      dt(6, 1) = "Fretwork or repeating greca frieze. Common at La Petaca."
    dt(7, 0) = "Triangular motif":    dt(7, 1) = "Painted triangular/chevron pattern. Documented at DW (over-lintel zone)."
    dt(8, 0) = "Painted band":        dt(8, 1) = "Horizontal painted band (red/white). Interbody cornice zone."
    dt(9, 0) = "Square niche":        dt(9, 1) = "Square niches in series (hornacinas cuadradas)."
    dt(10, 0) = "Plain colour field": dt(10, 1) = "Flat chromatic application with no motif: a whole element painted one colour."
    dt(11, 0) = "Decapitation scene": dt(11, 1) = "Decapitation scene. Record the anthropomorphic reading in Notes."
    dt(12, 0) = "ND":                 dt(12, 1) = "Decoration type not determined."
    ' v13: rock art types. The rupestrian repertoire is NOT shared
    ' with the architectural one - that was checked against the
    ' corpus - so these are additional entries, not a merge. Some
    ' generic types (Plain colour field, Painted band, ND) do serve
    ' both, which is why Level_Type and not ID_Dec_Type is the
    ' discriminator between the two sets.
    dt(13, 0) = "RA Anthropomorphic": dt(13, 1) = "Rock art: anthropomorphic figure."
    dt(14, 0) = "RA Zoomorphic":      dt(14, 1) = "Rock art: zoomorphic figure."
    dt(15, 0) = "RA Geometric":       dt(15, 1) = "Rock art: geometric motif."
    dt(16, 0) = "RA Abstract":        dt(16, 1) = "Rock art: abstract or non-figurative motif."
    dt(17, 0) = "RA Amorphous stain": dt(17, 1) = "Rock art: amorphous colour stain with no discernible motif."
    dt(18, 0) = "RA Perimeter band":  dt(18, 1) = "Rock art: band or marks framing an opening, threshold or structure outline. Cross-check with T_LOST_ELEMENTS when the fabric is gone."
    For i = 0 To 18
        db.Execute "INSERT INTO L_DEC_TYPE (Name,Description) VALUES ('" & dt(i, 0) & "','" & dt(i, 1) & "')", dbFailOnError
    Next i

    PopulateElements db
    PopulateLostEvidence db
    PopulateValencianLabels db

    Debug.Print "-> All lookups populated"
End Sub

' NEW v11 (3): the 24 elements A-X, bottom to top. H, P and U carry
' Is_System = True and point at their Sys_* field; W stays as a
' vocabulary entry with no field of its own, since Sys_Chamber
' absorbed it (5.5).

' ================================================================
'  Name_VAL (rev. 7) - UI LABELS ONLY
'
'  WHY THIS EXISTS. The form had two kinds of combo and only one of
'  them was localised: value-list combos (PC5, PC9, PCV) carry both
'  columns in the code string ("Red;Roig;White;Blanc;..."), so they
'  showed Valencian, while table-driven combos read Name and showed
'  English. Half the form in each language.
'
'  The fix is the SAME PATTERN the value lists already used - stored
'  column hidden, label column visible - only here the stored column
'  is a numeric ID, which makes it even safer: renaming a label
'  cannot touch a single stored value.
'
'  Name stays as the reference term: exports, publication and the
'  QRY_16 rules that match on it (Record_Class via typology,
'  L_STATUS for the collapse warning) all keep reading Name.
'  Name_VAL is display only.
'
'  Populated by UPDATE ... WHERE Name = ... rather than by rewriting
'  every INSERT: the translations stay in one readable block, and
'  adding Name_ES later is a copy of this Sub, not a rewrite of the
'  whole populate section.
' ================================================================
Private Sub PopulateValencianLabels(db As DAO.Database)
    VL db, "L_TYPOLOGY", "EA-MAU Mausoleum/Chullpa", "EA-MAU Mausoleu/Chullpa"
    VL db, "L_TYPOLOGY", "EA-CAM Funerary Chamber", "EA-CAM Cambra funeraria"
    VL db, "L_TYPOLOGY", "EA-PLA-R Ledge Platform", "EA-PLA-R Plataforma en repisa"
    VL db, "L_TYPOLOGY", "EA-PLA-V Aerial Platform", "EA-PLA-V Plataforma aeria"
    VL db, "L_TYPOLOGY", "NIX Natural Niche", "NIX Ninxol natural"
    VL db, "L_TYPOLOGY", "CAV Cave/Cavern", "CAV Cova/Cavitat"
    VL db, "L_TYPOLOGY", "PR Rock Art", "PR Art rupestre"
    VL db, "L_TYPOLOGY", "MEN Isolated Bracket", "MEN Mensula aillada"
    VL db, "L_TYPOLOGY", "Unclassifiable", "No classificable"
    VL db, "L_TYPOLOGY", "Not yet classified", "Pendent de classificar"

    VL db, "L_STATUS", "Good", "Bo"
    VL db, "L_STATUS", "Fair", "Regular"
    VL db, "L_STATUS", "Pre-collapse", "Pre-colapse"
    VL db, "L_STATUS", "Collapsed", "Colapsat"
    VL db, "L_STATUS", "ND", "Indeterminat"

    VL db, "L_MATERIAL_STATUS", "Good", "Bo"
    VL db, "L_MATERIAL_STATUS", "Fair", "Regular"
    VL db, "L_MATERIAL_STATUS", "Poor", "Deficient"
    VL db, "L_MATERIAL_STATUS", "ND", "Indeterminat"

    VL db, "L_SUPPORT", "Wide natural ledge (>2m)", "Repisa natural ampla (>2m)"
    VL db, "L_SUPPORT", "Narrow natural ledge (<2m)", "Repisa natural estreta (<2m)"
    VL db, "L_SUPPORT", "Artificial ledge", "Repisa artificial"
    VL db, "L_SUPPORT", "Large cavity (>10m2)", "Cavitat gran (>10m2)"
    VL db, "L_SUPPORT", "Medium cavity (1-10m2)", "Cavitat mitjana (1-10m2)"
    VL db, "L_SUPPORT", "Natural niche (<1m2)", "Ninxol natural (<1m2)"
    VL db, "L_SUPPORT", "Micro-ledge (<50cm)", "Micro-repisa (<50cm)"
    VL db, "L_SUPPORT", "Vertical cleft", "Diaclasi (escletxa vertical)"
    VL db, "L_SUPPORT", "Bedding-plane recess", "Junta d'estratificacio"
    VL db, "L_SUPPORT", "Ground", "Terreny (base del cingle)"
    VL db, "L_SUPPORT", "ND", "Indeterminat"

    VL db, "L_DEC_TYPE", "T-shaped niche", "Ninxol en T"
    VL db, "L_DEC_TYPE", "T-shaped niche inv.", "Ninxol en T invertida"
    VL db, "L_DEC_TYPE", "L-shaped niche", "Ninxol en L"
    VL db, "L_DEC_TYPE", "L-shaped niche inv.", "Ninxol en L invertida"
    VL db, "L_DEC_TYPE", "Zigzag", "Ziga-zaga"
    VL db, "L_DEC_TYPE", "Stepped motif", "Motiu escalonat"
    VL db, "L_DEC_TYPE", "Frieze / Greca", "Fris / Greca"
    VL db, "L_DEC_TYPE", "Triangular motif", "Motiu triangular"
    VL db, "L_DEC_TYPE", "Painted band", "Banda pintada"
    VL db, "L_DEC_TYPE", "Square niche", "Ninxol quadrat"
    VL db, "L_DEC_TYPE", "Plain colour field", "Camp de color pla"
    VL db, "L_DEC_TYPE", "Decapitation scene", "Escena de decapitacio"
    VL db, "L_DEC_TYPE", "ND", "Indeterminat"
    VL db, "L_DEC_TYPE", "RA Anthropomorphic", "AR Antropomorf"
    VL db, "L_DEC_TYPE", "RA Zoomorphic", "AR Zoomorf"
    VL db, "L_DEC_TYPE", "RA Geometric", "AR Geometric"
    VL db, "L_DEC_TYPE", "RA Abstract", "AR Abstracte"
    VL db, "L_DEC_TYPE", "RA Amorphous stain", "AR Taca amorfa"
    VL db, "L_DEC_TYPE", "RA Perimeter band", "AR Banda perimetral"

    VL db, "L_STRUCT_BODY", "Base level (B)", "Basament (B)"
    VL db, "L_STRUCT_BODY", "Socle (C)", "Socol (C)"
    VL db, "L_STRUCT_BODY", "Interbody cornice (I)", "Cornisa intercos (I)"
    VL db, "L_STRUCT_BODY", "Corner quoin (J)", "Cantonera (J)"
    VL db, "L_STRUCT_BODY", "Pilaster (K)", "Pilastra (K)"
    VL db, "L_STRUCT_BODY", "Facade flank (L)", "Ala de facana (L)"
    VL db, "L_STRUCT_BODY", "Jamb (O)", "Brancal (O)"
    VL db, "L_STRUCT_BODY", "Over-lintel", "Sobre-dintell"
    VL db, "L_STRUCT_BODY", "Upper crown (R)", "Coronament (R)"
    VL db, "L_STRUCT_BODY", "Eave (U)", "Rafec (U)"
    VL db, "L_STRUCT_BODY", "Adjacent bedrock", "Penya adjacent"
    VL db, "L_STRUCT_BODY", "Opening perimeter (rock)", "Perimetre d'obertura (penya)"
    VL db, "L_STRUCT_BODY", "Panel (no architectural ref.)", "Panell (sense ref. arquitectonica)"
    VL db, "L_STRUCT_BODY", "Not determined", "No determinada"

    VL db, "L_GROUP_TYPE", "Vertical alignment", "Alineament vertical"
    VL db, "L_GROUP_TYPE", "Ledge cluster", "Conjunt de repisa"
    VL db, "L_GROUP_TYPE", "Platform with brackets", "Plataforma amb mensules"
    VL db, "L_GROUP_TYPE", "Cave cluster", "Conjunt de cova"
    VL db, "L_GROUP_TYPE", "Circulation network", "Xarxa de circulacio"
    VL db, "L_GROUP_TYPE", "Rock art cluster", "Conjunt d'art rupestre"
    VL db, "L_GROUP_TYPE", "Supra-structural unit", "Unitat supraestructural"
    VL db, "L_GROUP_TYPE", "Functional group", "Grup funcional"

    VL db, "L_LOST_EVIDENCE", "Negative socket / impression", "Encaix negatiu / empremta"
    VL db, "L_LOST_EVIDENCE", "Beam hole", "Forat de biga"
    VL db, "L_LOST_EVIDENCE", "Break scar", "Cicatriu de despreniment"
    VL db, "L_LOST_EVIDENCE", "Detached fragment in situ", "Fragment despres in situ"
    VL db, "L_LOST_EVIDENCE", "Mortar imprint", "Empremta de morter"
    VL db, "L_LOST_EVIDENCE", "Corbels into void", "Mensules al buit"
    VL db, "L_LOST_EVIDENCE", "Pigment on bedrock", "Pigment sobre penya"
    VL db, "L_LOST_EVIDENCE", "Truncated walls", "Murs truncats"
    VL db, "L_LOST_EVIDENCE", "Other (see Notes)", "Altres (veure notes)"

    ' Any row left untranslated falls back to the English term, so a
    ' missing entry degrades gracefully instead of showing a blank.
    Dim t(8) As String
    t(0) = "L_TYPOLOGY": t(1) = "L_STATUS": t(2) = "L_MATERIAL_STATUS"
    t(3) = "L_SUPPORT": t(4) = "L_DEC_TYPE": t(5) = "L_STRUCT_BODY"
    t(6) = "L_GROUP_TYPE": t(7) = "L_LOST_EVIDENCE": t(8) = "L_VOL_METHOD"
    Dim i As Integer
    For i = 0 To 8
        On Error Resume Next
        db.Execute "UPDATE " & t(i) & " SET Name_VAL=Name WHERE Name_VAL Is Null", dbFailOnError
        On Error GoTo 0
    Next i
    On Error Resume Next
    db.Execute "UPDATE L_COORD_METHOD SET Name_VAL=Name WHERE Name_VAL Is Null", dbFailOnError
    On Error GoTo 0

    Debug.Print "-> Valencian UI labels populated (Name_VAL)"
End Sub

Private Sub VL(db As DAO.Database, tbl As String, en As String, va As String)
    On Error Resume Next
    db.Execute "UPDATE " & tbl & " SET Name_VAL='" & Replace(va, "'", "''") & "' WHERE Name='" & Replace(en, "'", "''") & "'", dbFailOnError
    On Error GoTo 0
End Sub

Private Sub PopulateElements(db As DAO.Database)
    Dim el(23, 7) As String
    el(0, 0) = "A":  el(0, 1) = "Embedded base beams":         el(0, 2) = "Jaceres basals":         el(0, 3) = "N0":    el(0, 4) = "":         el(0, 5) = "False": el(0, 6) = "Embedded_Base_Beams":  el(0, 7) = "Timber beams embedded in the basal masonry."
    el(1, 0) = "B":  el(1, 1) = "Base level":                  el(1, 2) = "Basament":               el(1, 3) = "N0":    el(1, 4) = "":         el(1, 5) = "False": el(1, 6) = "Base_Level":           el(1, 7) = "Constructed basal level supporting the structure."
    el(2, 0) = "C":  el(2, 1) = "Decorative socle":            el(2, 2) = "Socol decoratiu":        el(2, 3) = "N0":    el(2, 4) = "":         el(2, 5) = "False": el(2, 6) = "Decorative_Socle":     el(2, 7) = "Decorative treatment of the basal mass."
    el(3, 0) = "D":  el(3, 1) = "Tie walls":                   el(3, 2) = "Muret transversal":      el(3, 3) = "N0":    el(3, 4) = "":         el(3, 5) = "False": el(3, 6) = "Tie_Walls":            el(3, 7) = "Transverse tie wall at base level."
    el(4, 0) = "E":  el(4, 1) = "Timber brackets (corbels)":   el(4, 2) = "Mensules de fusta":      el(4, 3) = "N0":    el(4, 4) = "Platform": el(4, 5) = "False": el(4, 6) = "Timber_Brackets":      el(4, 7) = "Protruding horizontal timber corbels; component of the platform system (H)."
    el(5, 0) = "F":  el(5, 1) = "Transverse beams":            el(5, 2) = "Bigues transversals":    el(5, 3) = "N0":    el(5, 4) = "Platform": el(5, 5) = "False": el(5, 6) = "Transverse_Beams":     el(5, 7) = "Spanning beams; component of the platform system (H)."
    el(6, 0) = "G":  el(6, 1) = "Corbelled courses":           el(6, 2) = "Filades en voladis":     el(6, 3) = "N0":    el(6, 4) = "Platform": el(6, 5) = "False": el(6, 6) = "Corbelled_Courses":    el(6, 7) = "Masonry courses projecting in corbel; component of the platform system (H)."
    el(7, 0) = "H":  el(7, 1) = "Corbelled platform (system)": el(7, 2) = "Plataforma en voladis":  el(7, 3) = "N0":    el(7, 4) = "Platform": el(7, 5) = "True":  el(7, 6) = "Sys_Platform":         el(7, 7) = "Composite system: Timber_Brackets + Transverse_Beams + Corbelled_Courses."
    el(8, 0) = "I":  el(8, 1) = "Interbody cornice":           el(8, 2) = "Cornisa intercos":       el(8, 3) = "N0-N1": el(8, 4) = "":         el(8, 5) = "False": el(8, 6) = "Interbody_Cornice":    el(8, 7) = "Cornice zone between superposed bodies. Not gated by any system (4.5)."
    el(9, 0) = "J":  el(9, 1) = "Corner quoins":               el(9, 2) = "Cantoneres":             el(9, 3) = "N1":    el(9, 4) = "":         el(9, 5) = "False": el(9, 6) = "Corner_Quoins":        el(9, 7) = "Larger stones set vertically at the corners; component of Sys_Chamber."
    el(10, 0) = "K": el(10, 1) = "Structural pilasters":       el(10, 2) = "Pilastres estructurals": el(10, 3) = "N1":  el(10, 4) = "":        el(10, 5) = "False": el(10, 6) = "Structural_Pilasters": el(10, 7) = "Full-height pilaster integrated in the wall plane; component of Sys_Chamber."
    el(11, 0) = "L": el(11, 1) = "Facade flank":               el(11, 2) = "Ala de facana":         el(11, 3) = "N1":   el(11, 4) = "":        el(11, 5) = "False": el(11, 6) = "Facade_Flank":         el(11, 7) = "Wall face in the facade plane, flanking the opening; component of Sys_Chamber."
    el(12, 0) = "M": el(12, 1) = "Relief frieze":              el(12, 2) = "Fris en relleu":        el(12, 3) = "N1":   el(12, 4) = "":        el(12, 5) = "False": el(12, 6) = "Relief_Frieze":        el(12, 7) = "Architectural element, not a decoration (7.3); component of Sys_Chamber."
    el(13, 0) = "N": el(13, 1) = "Sill":                       el(13, 2) = "Llindar":               el(13, 3) = "N1":   el(13, 4) = "Portal": el(13, 5) = "False": el(13, 6) = "Sill":                 el(13, 7) = "Lower frame element of the access opening; component of the portal system (P)."
    el(14, 0) = "O": el(14, 1) = "Jambs":                      el(14, 2) = "Brancals":              el(14, 3) = "N1":   el(14, 4) = "Portal": el(14, 5) = "False": el(14, 6) = "Jambs":                el(14, 7) = "Vertical frame elements of the access opening; component of the portal system (P)."
    el(15, 0) = "P": el(15, 1) = "Access portal (system)":     el(15, 2) = "Sistema portal":        el(15, 3) = "N1":   el(15, 4) = "Portal": el(15, 5) = "True":  el(15, 6) = "Sys_Portal":           el(15, 7) = "Composite system: Sill + Jambs + Lintel."
    el(16, 0) = "Q": el(16, 1) = "Lintel":                     el(16, 2) = "Dintell":               el(16, 3) = "N1":   el(16, 4) = "Portal": el(16, 5) = "False": el(16, 6) = "Lintel":               el(16, 7) = "Upper frame element of the access opening; component of the portal system (P)."
    el(17, 0) = "R": el(17, 1) = "Upper crown":                el(17, 2) = "Coronament":            el(17, 3) = "N1":   el(17, 4) = "":       el(17, 5) = "False": el(17, 6) = "Upper_Crown":          el(17, 7) = "Upper crown or coping of the body. Not gated by any system (4.5)."
    el(18, 0) = "S": el(18, 1) = "Eave beam":                  el(18, 2) = "Biga de suport rafec":  el(18, 3) = "SUP":  el(18, 4) = "Eave":   el(18, 5) = "False": el(18, 6) = "Eave_Beam":            el(18, 7) = "Supporting beam of the eave; component of the eave system (U)."
    el(19, 0) = "T": el(19, 1) = "Eave surface":               el(19, 2) = "Superficie de rafec":   el(19, 3) = "SUP":  el(19, 4) = "Eave":   el(19, 5) = "False": el(19, 6) = "Eave_Surface":         el(19, 7) = "Finished surface of the eave; component of the eave system (U)."
    el(20, 0) = "U": el(20, 1) = "Eave (system)":              el(20, 2) = "Rafec en voladis":      el(20, 3) = "SUP":  el(20, 4) = "Eave":   el(20, 5) = "True":  el(20, 6) = "Sys_Eave":             el(20, 7) = "Composite system: Eave_Beam + Eave_Surface."
    el(21, 0) = "V": el(21, 1) = "Return wall":                el(21, 2) = "Mur de retorn":         el(21, 3) = "N1":   el(21, 4) = "":       el(21, 5) = "False": el(21, 6) = "Return_Wall":          el(21, 7) = "Wall perpendicular to the facade plane, returning toward the cliff; component of Sys_Chamber."
    el(22, 0) = "W": el(22, 1) = "Rear wall":                  el(22, 2) = "Mur posterior":         el(22, 3) = "N1":   el(22, 4) = "":       el(22, 5) = "False": el(22, 6) = "":                     el(22, 7) = "Vocabulary entry retained for reference; no field of its own in v11 (absorbed by Sys_Chamber, 5.5)."
    el(23, 0) = "X": el(23, 1) = "Chamber roof":               el(23, 2) = "Coberta de cambra":     el(23, 3) = "N1":   el(23, 4) = "":       el(23, 5) = "False": el(23, 6) = "Chamber_Roof":         el(23, 7) = "Element closing the chamber above; component of Sys_Chamber."

    Dim sql As String
    Dim i As Integer
    For i = 0 To 23
        sql = "INSERT INTO L_ELEMENTS (Code, Name_EN, Name_VAL, Level_Type, Sys_Group, Is_System, Field_Name, Description) VALUES ("
        sql = sql & "'" & el(i, 0) & "',"
        sql = sql & "'" & el(i, 1) & "',"
        sql = sql & "'" & el(i, 2) & "',"
        sql = sql & "'" & el(i, 3) & "',"
        sql = sql & "'" & el(i, 4) & "',"
        sql = sql & el(i, 5) & ","
        sql = sql & "'" & el(i, 6) & "',"
        sql = sql & "'" & el(i, 7) & "')"
        db.Execute sql, dbFailOnError
    Next i
    Debug.Print "-> L_ELEMENTS: 24 rows (A-X)"
End Sub

' NEW v11 (2): the closed list of evidence types. A value of 3 is
' admitted for any element if and only if one of these can be pointed
' at; if none applies, the correct value is 0 or 9. The lookup acts
' as a methodological filter.
Private Sub PopulateLostEvidence(db As DAO.Database)
    Dim ev(8, 1) As String
    ev(0, 0) = "Negative socket / impression": ev(0, 1) = "Empty socket or impression left in the wall face"
    ev(1, 0) = "Beam hole":                    ev(1, 1) = "Through-hole for a spanning beam"
    ev(2, 0) = "Break scar":                   ev(2, 1) = "Detachment scar on masonry or bedrock"
    ev(3, 0) = "Detached fragment in situ":    ev(3, 1) = "Identifiable fallen fragment at the foot of the structure"
    ev(4, 0) = "Mortar imprint":               ev(4, 1) = "Mortar imprint without the element it once bonded"
    ev(5, 0) = "Corbels into void":            ev(5, 1) = "Corbels that no longer support anything"
    ev(6, 0) = "Pigment on bedrock":           ev(6, 1) = "Pigment on bedrock now bare of the masonry that carried it"
    ev(7, 0) = "Truncated walls":              ev(7, 1) = "Walls truncated on a clean plane"
    ev(8, 0) = "Other (see Notes)":            ev(8, 1) = "See the Notes field for detail"

    Dim i As Integer
    For i = 0 To 8
        db.Execute "INSERT INTO L_LOST_EVIDENCE (Name,Description) VALUES ('" & ev(i, 0) & "','" & ev(i, 1) & "')", dbFailOnError
    Next i
    Debug.Print "-> L_LOST_EVIDENCE: 9 rows"
End Sub

' ================================================================
'  3. RELATIONSHIPS (25: 21 from v10 + 4 for T_LOST_ELEMENTS)
' ================================================================
Private Sub CreateAllRelationships(db As DAO.Database)
    Dim rn(24) As String
    rn(0) = "REL_SIT_SEC":    rn(1) = "REL_SEC_STR":    rn(2) = "REL_SEC_GRP"
    rn(3) = "REL_TYP_STR":    rn(4) = "REL_SUP_STR":    rn(5) = "REL_STR_SELF"
    rn(6) = "REL_STA_STR":    rn(7) = "REL_MATSTA_STR": rn(8) = "REL_VM_STR"
    rn(9) = "REL_CM_STR":     rn(10) = "REL_GT_GRP":    rn(11) = "REL_GRP_STR"
    rn(12) = "REL_STR_DAT":   rn(13) = "REL_STR_IND":   rn(14) = "REL_STR_DEC"
    rn(15) = "REL_SB_DEC":    rn(16) = "REL_DT_DEC":    rn(17) = "REL_STR_AFEAT"
    rn(18) = "REL_STR_CONA":  rn(19) = "REL_STR_CONB":  rn(20) = "REL_SUP2_STR"
    rn(21) = "REL_STR_LOST":  rn(22) = "REL_ELEM_LOST": rn(23) = "REL_LEV_LOST"
    rn(24) = "REL_SB_LOST"
    Dim n As Integer
    For n = 0 To 24
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
    ' Dual reference to T_STRUCTURES -> no enforced integrity (flag 2),
    ' same treatment as the self-referencing REL_STR_SELF
    MkRel db, rn(18), "T_STRUCTURES", "ID", "T_CONNECTIONS", "ID_Struct_A", False, True
    MkRel db, rn(19), "T_STRUCTURES", "ID", "T_CONNECTIONS", "ID_Struct_B", False, True
    MkRel db, rn(20), "L_SUPPORT", "ID", "T_STRUCTURES", "ID_Support_Secondary", False, False
    ' NEW v11. REL_ELEM_LOST is the one that needs UQ_ELEM_CODE.
    MkRel db, rn(21), "T_STRUCTURES", "ID", "T_LOST_ELEMENTS", "ID_Structure", True, False
    MkRel db, rn(22), "L_ELEMENTS", "Code", "T_LOST_ELEMENTS", "Element_Code", False, False
    MkRel db, rn(23), "L_LOST_EVIDENCE", "ID", "T_LOST_ELEMENTS", "ID_Evidence_Type", False, False
    MkRel db, rn(24), "L_STRUCT_BODY", "ID", "T_LOST_ELEMENTS", "ID_Position", False, False

    db.Relations.Refresh
    Debug.Print "-> 25 relationships OK"
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
'  4. QUERIES (25 objects)
'     Created in dependency order: sources before dependants.
' ================================================================
Private Sub CreateAllQueries(db As DAO.Database)
    Dim qn(27) As String
    qn(0) = "QRY_01_Typology_by_Site"
    qn(1) = "QRY_02s_Decoration_Typed"
    qn(2) = "QRY_02a_Decoration_Flags"
    qn(3) = "QRY_02_Decoration_by_Site"
    qn(4) = "QRY_03_Conservation_by_Sector"
    qn(5) = "QRY_04_C14_Structures"
    qn(6) = "QRY_05_Export_RStats"
    qn(7) = "QRY_06_Volumetry_by_Typology"
    qn(8) = "QRY_07_Export_QGIS"
    qn(9) = "QRY_08_Children_of_Parent"
    qn(10) = "QRY_09_Group_Members"
    qn(11) = "QRY_10_ChaXR_Coverage"
    qn(12) = "QRY_11_Masonry_by_Site"
    qn(13) = "QRY_12_Geology_Construction"
    qn(14) = "QRY_13_AX_Pattern_Export"
    qn(15) = "QRY_14_Connections_Edges"
    qn(16) = "QRY_15_Observability_Bias"
    qn(17) = "QRY_16s_Element_Values"
    qn(18) = "QRY_16s_Lost_Cover"
    qn(19) = "QRY_16s_Damage_Count"
    qn(20) = "QRY_16s_Observational_Nulls"
    qn(21) = "QRY_16s_Null_Count"
    qn(22) = "QRY_16a_Rules_1_11"
    qn(23) = "QRY_16b_Rules_12_21"
    qn(24) = "QRY_16c_Rules_22_31"
    qn(25) = "QRY_17_RockArt_All"
    qn(26) = "QRY_18_Notes_Review"
    qn(27) = "QRY_16_Validation_Check"
    Dim i As Integer
    ' Reverse order so dependants go before their sources
    For i = 27 To 0 Step -1
        If QueryExists(db, qn(i)) Then db.QueryDefs.Delete qn(i)
    Next i

    BuildBasicQueries db
    BuildDecorationQueries db
    BuildExportQueries db
    BuildAXExport db
    BuildRockArtQuery db
    BuildNotesQuery db
    BuildValidationHelpers db
    BuildValidationBattery db

    Debug.Print "-> 28 queries OK"
End Sub

Private Sub MkQuery(db As DAO.Database, qn As String, sql As String)
    On Error GoTo Err_MQ
    db.CreateQueryDef qn, sql
    Exit Sub
Err_MQ: Debug.Print "  *** FAILED to create " & qn & ": " & Err.Description
End Sub

Private Sub BuildBasicQueries(db As DAO.Database)
    Dim q As String

    ' QRY_01 - Typology distribution by site
    q = "SELECT S.Site_Name, T.Name AS Typology, T.Record_Class, COUNT(E.ID) AS N "
    q = q & "FROM ((T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID) "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID "
    q = q & "GROUP BY S.Site_Name, T.Name, T.Record_Class "
    q = q & "ORDER BY S.Site_Name, T.Name;"
    MkQuery db, "QRY_01_Typology_by_Site", q

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
    MkQuery db, "QRY_03_Conservation_by_Sector", q

    ' QRY_04 - Structures with C14 dating
    q = "SELECT E.Code, S.Site_Name, SC.Sector_Name, T.Name AS Typology, "
    q = q & "E.Chrono_Start_Cent, E.Chrono_End_Cent "
    q = q & "FROM (((T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID) "
    q = q & "LEFT JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID) "
    q = q & "WHERE E.C14=True ORDER BY E.Chrono_Start_Cent;"
    MkQuery db, "QRY_04_C14_Structures", q

    ' QRY_06 - Volumetry by typology
    q = "SELECT S.Site_Name, T.Name AS Typology, "
    q = q & "COUNT(E.ID) AS N_Total, COUNT(E.Volume_m3) AS N_with_Vol, "
    q = q & "AVG(E.Area_m2) AS Mean_Area, "
    q = q & "AVG(E.Volume_m3) AS Mean_Vol, "
    q = q & "MIN(E.Volume_m3) AS Min_Vol, MAX(E.Volume_m3) AS Max_Vol "
    q = q & "FROM ((T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID) "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID "
    q = q & "WHERE E.Volume_m3 IS NOT NULL "
    q = q & "GROUP BY S.Site_Name, T.Name ORDER BY S.Site_Name, Mean_Vol DESC;"
    MkQuery db, "QRY_06_Volumetry_by_Typology", q

    ' QRY_08 - Children of a parent structure
    q = "SELECT E_c.Code AS Child_Code, T.Name AS Typology, "
    q = q & "E_c.Human_Remains, E_c.MNI "
    q = q & "FROM T_STRUCTURES AS E_c "
    q = q & "LEFT JOIN L_TYPOLOGY AS T ON E_c.ID_Typology=T.ID "
    q = q & "WHERE E_c.ID_Parent=[Parent ID?] "
    q = q & "ORDER BY T.Name, E_c.Code;"
    MkQuery db, "QRY_08_Children_of_Parent", q

    ' QRY_09 - Members of a functional group
    q = "SELECT G.Group_Code, GT.Name AS Group_Type, E.Code, T.Name AS Typology, "
    q = q & "E.Coord_E_UTM, E.Coord_N_UTM, E.Altitude_masl "
    q = q & "FROM (((T_STRUCTURES AS E "
    q = q & "INNER JOIN T_GROUPS AS G ON E.ID_Group=G.ID) "
    q = q & "LEFT JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID) "
    q = q & "INNER JOIN L_GROUP_TYPE AS GT ON G.ID_Group_Type=GT.ID) "
    q = q & "WHERE G.Group_Code=[Group code?] "
    q = q & "ORDER BY E.Altitude_masl DESC;"
    MkQuery db, "QRY_09_Group_Members", q

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
    MkQuery db, "QRY_10_ChaXR_Coverage", q

    ' QRY_11 - Masonry quality by site and typology (H01/H04)
    q = "SELECT S.Site_Name, T.Name AS Typology, "
    q = q & "E.Masonry_Quality, E.Masonry_Type, COUNT(E.ID) AS N "
    q = q & "FROM ((T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID) "
    q = q & "LEFT JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID "
    q = q & "WHERE E.Masonry_Quality IS NOT NULL "
    q = q & "GROUP BY S.Site_Name, T.Name, E.Masonry_Quality, E.Masonry_Type "
    q = q & "ORDER BY S.Site_Name, T.Name, E.Masonry_Quality;"
    MkQuery db, "QRY_11_Masonry_by_Site", q

    ' QRY_12 - Geology-construction relationship (H02)
    q = "SELECT S.Site_Name, T.Name AS Typology, "
    q = q & "SU.Name AS Support_Primary, SU2.Name AS Support_Secondary, "
    q = q & "E.Support_Modified, COUNT(E.ID) AS N, "
    q = q & "AVG(E.Support_Width_m) AS Avg_Width_m, "
    q = q & "AVG(E.Support_Depth_m) AS Avg_Depth_m "
    q = q & "FROM ((((T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID) "
    q = q & "LEFT JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID) "
    q = q & "LEFT JOIN L_SUPPORT AS SU ON E.ID_Support=SU.ID) "
    q = q & "LEFT JOIN L_SUPPORT AS SU2 ON E.ID_Support_Secondary=SU2.ID "
    q = q & "GROUP BY S.Site_Name, T.Name, SU.Name, SU2.Name, E.Support_Modified "
    q = q & "ORDER BY S.Site_Name, T.Name;"
    MkQuery db, "QRY_12_Geology_Construction", q

    ' QRY_15 - Documentation and observability bias report (OE1)
    ' The 0/9 distinction records WHERE the gaps are; this quantifies
    ' them, so the sample can be defended rather than merely disclaimed.
    q = "SELECT S.Site_Name, SC.Sector_Name, "
    q = q & "E.Doc_Basis, E.Facade_Observability, E.Interior_Observability, "
    q = q & "COUNT(E.ID) AS N "
    q = q & "FROM (T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID "
    q = q & "GROUP BY S.Site_Name, SC.Sector_Name, E.Doc_Basis, "
    q = q & "E.Facade_Observability, E.Interior_Observability "
    q = q & "ORDER BY S.Site_Name, SC.Sector_Name;"
    MkQuery db, "QRY_15_Observability_Bias", q
End Sub

' ================================================================
'  QRY_02 REBUILT ON T_DECORATIONS (7.6)
'
'  The v10 version was built entirely on the Dec_* booleans that v11
'  drops, and it is the source of the LP vs DW chi-squared.
'  T_DECORATIONS records only PRESENCES, so the triad that lets the
'  denominator be chosen has to be reconstructed:
'      Present         a row of that motif exists
'      Verified absent no row AND Facade_Observability = 'Complete'
'      Not evaluable   no row AND observability is anything else
'  Valid denominator = Present + Verified absent. NEVER the total:
'  facade preservation differs between the two sites, so an inflated
'  denominator would not inflate equally, and the test would measure
'  differential preservation while reading as differential practice.
'
'  THREE OBJECTS, FOR TWO REASONS.
'  (a) A structure can carry the same motif in two positions, so a
'      direct LEFT JOIN would duplicate its row and SUM(IIF()) would
'      count it twice. The GROUP BY in the flag query collapses that
'      before any counting: MAX(IIF(...)) reads 1 for one match or
'      five.
'  (b) JET refuses a LEFT JOIN whose right operand is a parenthesised
'      INNER JOIN (error 3296), but accepts one against a SAVED
'      QUERY - hence QRY_02s. Matching motifs by NAME rather than by
'      hard-coded L_DEC_TYPE autonumbers is deliberate (section 12).
' ================================================================
Private Sub BuildDecorationQueries(db As DAO.Database)
    Dim q As String

    q = "SELECT D.ID_Structure, DT.Name AS Dec_Name "
    q = q & "FROM T_DECORATIONS AS D "
    q = q & "INNER JOIN L_DEC_TYPE AS DT ON D.ID_Dec_Type=DT.ID;"
    MkQuery db, "QRY_02s_Decoration_Typed", q

    q = "SELECT E.ID, E.Code, S.Site_Name AS Site, SC.Sector_Name AS Sector, "
    q = q & "T.Record_Class, E.Facade_Observability, "
    q = q & "IIF(E.Facade_Observability='Complete',1,0) AS Obs_Complete, "
    q = q & "MAX(IIF(DX.Dec_Name='Square niche',1,0)) AS Has_Niche, "
    q = q & "MAX(IIF(DX.Dec_Name='T-shaped niche',1,0)) AS Has_T, "
    q = q & "MAX(IIF(DX.Dec_Name='T-shaped niche inv.',1,0)) AS Has_T_Inv, "
    q = q & "MAX(IIF(DX.Dec_Name='L-shaped niche',1,0)) AS Has_L, "
    q = q & "MAX(IIF(DX.Dec_Name='L-shaped niche inv.',1,0)) AS Has_L_Inv, "
    q = q & "MAX(IIF(DX.Dec_Name='Zigzag',1,0)) AS Has_Zigzag, "
    q = q & "MAX(IIF(DX.Dec_Name='Stepped motif',1,0)) AS Has_Stepped, "
    q = q & "MAX(IIF(DX.Dec_Name='Frieze / Greca',1,0)) AS Has_Frieze, "
    q = q & "IIF(E.Pigment_Present=1 AND E.Pigment_Substrate='Bedrock',1,0) AS Has_RockArt "
    q = q & "FROM ((((T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID) "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID) "
    q = q & "LEFT JOIN QRY_02s_Decoration_Typed AS DX ON DX.ID_Structure=E.ID) "
    q = q & "WHERE T.Record_Class<>'Pending classification' "
    q = q & "GROUP BY E.ID, E.Code, S.Site_Name, SC.Sector_Name, T.Record_Class, "
    q = q & "E.Facade_Observability, E.Pigment_Present, E.Pigment_Substrate "
    q = q & "ORDER BY S.Site_Name, E.Code;"
    MkQuery db, "QRY_02a_Decoration_Flags", q

    Dim m(8, 1) As String
    m(0, 0) = "Has_Niche":    m(0, 1) = "Niche"
    m(1, 0) = "Has_T":        m(1, 1) = "T"
    m(2, 0) = "Has_T_Inv":    m(2, 1) = "T_Inv"
    m(3, 0) = "Has_L":        m(3, 1) = "L"
    m(4, 0) = "Has_L_Inv":    m(4, 1) = "L_Inv"
    m(5, 0) = "Has_Zigzag":   m(5, 1) = "Zigzag"
    m(6, 0) = "Has_Stepped":  m(6, 1) = "Stepped"
    m(7, 0) = "Has_Frieze":   m(7, 1) = "Frieze"
    m(8, 0) = "Has_RockArt":  m(8, 1) = "RockArt"

    Dim i As Integer
    q = "SELECT F.Site, "
    For i = 0 To 8
        q = q & "SUM(IIF(F." & m(i, 0) & "=1,1,0)) AS N_" & m(i, 1) & ", "
        q = q & "SUM(IIF(F." & m(i, 0) & "=0 AND F.Obs_Complete=1,1,0)) AS N_" & m(i, 1) & "_Absent, "
        q = q & "SUM(IIF(F." & m(i, 0) & "=0 AND F.Obs_Complete=0,1,0)) AS N_" & m(i, 1) & "_ND, "
    Next i
    q = q & "COUNT(F.ID) AS Total "
    q = q & "FROM QRY_02a_Decoration_Flags AS F "
    q = q & "GROUP BY F.Site ORDER BY F.Site;"
    MkQuery db, "QRY_02_Decoration_by_Site", q
End Sub

Private Sub BuildExportQueries(db As DAO.Database)
    Dim q As String

    ' QRY_05 - Flat export for R. LEFT JOINs throughout: an export that
    ' silently drops records with no typology or no support recorded is
    ' a bug, not a filter.
    q = "SELECT E.ID, E.Code, S.Site_Name AS Site, SC.Sector_Name AS Sector, "
    q = q & "T.Name AS Typology, T.Record_Class, SU.Name AS Support, "
    q = q & "AS1.Name AS Arch_Status, MS.Name AS Material_Status, "
    q = q & "E.N_Basal_Bodies, E.N_Chamber_Bodies, "
    q = q & "E.Floor_Plan, E.N_Built_Walls, "
    q = q & "E.Sys_Platform, E.Sys_Portal, E.Sys_Eave, E.Sys_Base, E.Sys_Chamber, "
    q = q & "E.Platform_Function, "
    q = q & "E.Chamber_Roof_Type, E.Rear_Closure_Type, "
    q = q & "E.Plaster_Present, E.Plaster_Extent, "
    q = q & "E.Pigment_Present, E.Pigment_Substrate, E.Pigment_Extent, "
    q = q & "E.Embedded_Base_Beams, E.Base_Level, E.Decorative_Socle, E.Tie_Walls, "
    q = q & "E.Timber_Brackets, E.Transverse_Beams, E.Corbelled_Courses, "
    q = q & "E.Interbody_Cornice, E.Corner_Quoins, E.Structural_Pilasters, "
    q = q & "E.Facade_Flank, E.Relief_Frieze, E.Sill, E.Jambs, "
    q = q & "E.Lintel, E.Lintel_Material, E.Recessed_Frame, "
    q = q & "E.Upper_Crown, E.Eave_Beam, E.Eave_Surface, "
    q = q & "E.Return_Wall, E.Chamber_Roof, "
    q = q & "E.Looting, E.Fire_Damage, E.Animal_Activity, E.Modern_Access, "
    q = q & "E.Human_Remains, E.MNI, E.Mummification, E.Funerary_Bundles, "
    q = q & "E.Bone_Burning, E.Mat_Textiles, E.Mat_Ceramics, E.Mat_DeerAntler, "
    q = q & "E.C14, E.Chrono_Start_Cent, E.Chrono_End_Cent, "
    q = q & "E.Area_m2, E.Volume_m3, "
    q = q & "E.Coord_Lat_WGS84, E.Coord_Lon_WGS84, "
    q = q & "E.Coord_E_UTM, E.Coord_N_UTM, E.Altitude_masl, "
    q = q & "E.ChaXR_Documented, "
    q = q & "E.Support_Width_m, E.Support_Depth_m, E.Support_Modified, "
    q = q & "E.Masonry_Quality, E.Masonry_Type, "
    q = q & "E.Mortar_Present, E.Mortar_Type, E.Chinking_Stones, "
    q = q & "E.Opening_Width_m, E.Opening_Height_m, "
    q = q & "E.Height_Above_Base_m, "
    q = q & "E.Facade_Orientation, E.Visibility_Valley, "
    q = q & "E.Construction_Phases, E.Phase_Evidence, "
    q = q & "E.Timber_Bracket_Count, E.Timber_Bracket_Role, "
    q = q & "E.Platform_Surface_Material, E.Interbody_Cornice_Material, "
    q = q & "E.Doc_Basis, E.Facade_Observability, E.Interior_Observability "
    q = q & "FROM ((((((T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID) "
    q = q & "LEFT JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID) "
    q = q & "LEFT JOIN L_SUPPORT AS SU ON E.ID_Support=SU.ID) "
    q = q & "LEFT JOIN L_STATUS AS AS1 ON E.ID_Arch_Status=AS1.ID) "
    q = q & "LEFT JOIN L_MATERIAL_STATUS AS MS ON E.ID_Material_Status=MS.ID) "
    q = q & "ORDER BY S.Site_Name, SC.Sector_Name, E.Code;"
    MkQuery db, "QRY_05_Export_RStats", q

    ' QRY_07 - Spatial export. Lost_Body_Evidence is replaced by a
    ' per-structure count from T_LOST_ELEMENTS (12bis).
    q = "SELECT E.ID, E.Code, S.Site_Name, SC.Sector_Name, "
    q = q & "T.Name AS Typology, T.Record_Class, "
    q = q & "E.Coord_Lat_WGS84, E.Coord_Lon_WGS84, "
    q = q & "E.Coord_E_UTM, E.Coord_N_UTM, E.Altitude_masl, "
    q = q & "E.Height_Above_Base_m, E.Coord_Precision_m, "
    q = q & "AS1.Name AS Arch_Status, MS.Name AS Material_Status, "
    q = q & "E.N_Basal_Bodies, E.N_Chamber_Bodies, "
    q = q & "(SELECT COUNT(*) FROM T_LOST_ELEMENTS AS LE WHERE LE.ID_Structure=E.ID) AS N_Lost_Elements, "
    q = q & "E.Sys_Platform, E.Sys_Portal, E.Sys_Eave, "
    q = q & "E.Area_m2, E.Volume_m3, "
    q = q & "E.Chrono_Start_Cent, E.Chrono_End_Cent, "
    q = q & "E.Looting, E.Human_Remains, E.MNI, "
    q = q & "E.ChaXR_Documented, E.URL_3D, "
    q = q & "E.Facade_Orientation, E.Visibility_Valley, E.Masonry_Quality, "
    q = q & "E.Doc_Basis, E.Facade_Observability "
    q = q & "FROM (((((T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID) "
    q = q & "LEFT JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID) "
    q = q & "LEFT JOIN L_STATUS AS AS1 ON E.ID_Arch_Status=AS1.ID) "
    q = q & "LEFT JOIN L_MATERIAL_STATUS AS MS ON E.ID_Material_Status=MS.ID) "
    q = q & "WHERE E.Coord_Lat_WGS84 IS NOT NULL "
    q = q & "ORDER BY S.Site_Name, E.Code;"
    MkQuery db, "QRY_07_Export_QGIS", q

    ' QRY_14 - Connection edge list. Chrono_Relation makes the graph
    ' potentially directed, which is what H03 needs (9.3).
    q = "SELECT C.ID, A.Code AS Code_A, B.Code AS Code_B, "
    q = q & "C.Connection_Type, C.Chrono_Relation, C.Confidence, "
    q = q & "A.Coord_E_UTM AS E_UTM_A, A.Coord_N_UTM AS N_UTM_A, "
    q = q & "A.Altitude_masl AS Alt_A, "
    q = q & "B.Coord_E_UTM AS E_UTM_B, B.Coord_N_UTM AS N_UTM_B, "
    q = q & "B.Altitude_masl AS Alt_B, "
    q = q & "C.Notes "
    q = q & "FROM ((T_CONNECTIONS AS C "
    q = q & "INNER JOIN T_STRUCTURES AS A ON C.ID_Struct_A=A.ID) "
    q = q & "INNER JOIN T_STRUCTURES AS B ON C.ID_Struct_B=B.ID) "
    q = q & "ORDER BY C.Connection_Type, A.Code;"
    MkQuery db, "QRY_14_Connections_Edges", q
End Sub

' ================================================================
'  QRY_13 - THE A-X MATRIX, WITH NA SEMANTICS (12bis, 4.3, 4.7)
'
'  20 element columns, not 24: H, P and U are systems now and get
'  SYS_ columns; W is gone entirely, its information carried by
'  Rear_Closure_Type. AX_Q reads the Lintel BYTE directly.
'
'  THE NA RULE IS THE POINT. When a system is Not applicable or Not
'  observable the query exports NULL - NA in R - for the system AND
'  for every component of it, because those components hold the
'  technical 0 of section 1.2, a padding value that asserts nothing.
'  Exporting it raw would put unverified absences into denominators
'  and the published percentages would be wrong. Absent, by contrast,
'  exports real values: a mausoleum with no eave is a result.
' ================================================================
Private Sub BuildAXExport(db As DAO.Database)
    Dim ax() As String
    FillElementMap ax

    Dim cm(8, 1) As String
    cm(0, 0) = "Timber_Bracket_Count":       cm(0, 1) = "Sys_Platform"
    cm(1, 0) = "Timber_Bracket_Role":        cm(1, 1) = "Sys_Platform"
    cm(2, 0) = "Platform_Surface_Material":  cm(2, 1) = "Sys_Platform"
    cm(3, 0) = "Platform_Function":          cm(3, 1) = "Sys_Platform"
    cm(4, 0) = "Lintel_Material":            cm(4, 1) = "Sys_Portal"
    cm(5, 0) = "Recessed_Frame":             cm(5, 1) = "Sys_Portal"
    cm(6, 0) = "Chamber_Roof_Type":          cm(6, 1) = "Sys_Chamber"
    cm(7, 0) = "Rear_Closure_Type":          cm(7, 1) = "Sys_Chamber"
    cm(8, 0) = "Interbody_Cornice_Material": cm(8, 1) = ""

    Dim sy(4, 1) As String
    sy(0, 0) = "Sys_Platform": sy(0, 1) = "SYS_H"
    sy(1, 0) = "Sys_Portal":   sy(1, 1) = "SYS_P"
    sy(2, 0) = "Sys_Eave":     sy(2, 1) = "SYS_U"
    sy(3, 0) = "Sys_Base":     sy(3, 1) = "SYS_BASE"
    sy(4, 0) = "Sys_Chamber":  sy(4, 1) = "SYS_CHAMBER"

    Dim q As String
    Dim i As Integer

    q = "SELECT E.ID, E.Code, S.Site_Name AS Site, SC.Sector_Name AS Sector, "
    q = q & "T.Name AS Typology, T.Record_Class, "
    q = q & "E.Coord_E_UTM, E.Coord_N_UTM, E.Altitude_masl, E.ID_Group, "
    q = q & "E.N_Basal_Bodies, E.N_Chamber_Bodies, "
    For i = 0 To 19
        q = q & Gated(ax(i, 1), ax(i, 2), "AX_" & ax(i, 0))
    Next i
    For i = 0 To 4
        q = q & Gated(sy(i, 0), sy(i, 0), sy(i, 1))
    Next i
    For i = 0 To 8
        q = q & Gated(cm(i, 0), cm(i, 1), cm(i, 0))
    Next i
    q = q & "E.Masonry_Quality, E.Masonry_Type, "
    q = q & "E.Mortar_Present, E.Chinking_Stones, "
    q = q & "E.Plaster_Present, E.Pigment_Present, E.Pigment_Substrate, "
    q = q & "E.Support_Modified, "
    q = q & "E.Doc_Basis, E.Facade_Observability, E.Interior_Observability "
    q = q & "FROM (((T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID) "
    q = q & "LEFT JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID) "
    q = q & "ORDER BY S.Site_Name, SC.Sector_Name, E.Code;"
    MkQuery db, "QRY_13_AX_Pattern_Export", q
End Sub

' One SELECT column, wrapped in the NA gate when it belongs to a
' system. gate = "" means never gated.
Private Function Gated(fld As String, gate As String, outName As String) As String
    If Len(gate) = 0 Then
        Gated = "E." & fld & " AS " & outName & ", "
    Else
        Gated = "IIF(E." & gate & "='Not applicable' Or E." & gate & "='Not observable',Null,E." & fld & ") AS " & outName & ", "
    End If
End Function

' ================================================================
'  QRY_16 HELPERS
'
'  Rules 1, 2, 4 and 6 each iterate over all 20 element fields.
'  Written directly that would be ~80 near-identical UNION branches.
'  QRY_16s_Element_Values pivots the columns into long format once,
'  and each rule becomes a single branch that also names the
'  offending field in its message.
'
'  QRY_16s_Null_Count deliberately does NOT group the unpivot query:
'  JET inlines a saved query wherever it is referenced, so building
'  on the 44-branch QRY_16s_Observational_Nulls dragged those
'  branches into the battery and blew the "query too complex"
'  ceiling. A plain sum of 44 IIFs on one table costs nothing. The
'  unpivot survives as the standalone review checklist.
' ================================================================
' ================================================================
'  QRY_17_RockArt_All (v13, delta 7.7)
'
'  The two halves of the rock art record live in different places -
'  associated paintings as ROC rows hanging off a structure,
'  isolated ones as ROC rows hanging off a PR record - but BOTH ARE
'  ROWS OF T_DECORATIONS, because a PR record is itself a
'  T_STRUCTURES row. So the union needs no UNION at all: one query
'  over T_DECORATIONS, joined to L_TYPOLOGY to find out whether the
'  parent is a panel or a structure.
'
'  Context is a VARIABLE, not administrative noise: it lets you ask
'  whether motifs, colours or position relative to the threshold
'  differ depending on whether there is a structure there.
'
'  KNOWN LIMITATION, accepted: the coordinates of an associated
'  painting are those of its parent structure, not of the painting.
'  Irrelevant at cliff scale (sector distribution, density,
'  visibility). If fine position is ever needed, the cheap fix is
'  optional coordinates on T_DECORATIONS - not done now.
' ================================================================
' QRY_18 (rev. 6): every notes field of every record, side by side.
' Serves the review pass, and - more usefully - makes it possible to
' SPOT PATTERNS: when the same observation keeps recurring in free
' text, that is the signal that it should be a field.
Private Sub BuildNotesQuery(db As DAO.Database)
    Dim q As String
    q = "SELECT E.Code, S.Site_Name, SC.Sector_Name, "
    q = q & "E.Arch_Notes, E.Finish_Notes, E.Condition_Notes, "
    q = q & "E.Bio_Notes, E.Materials_Notes, E.Systems_Notes, "
    q = q & "E.Mortar_Notes, E.Vol_Notes, E.Notes "
    q = q & "FROM (T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID "
    q = q & "WHERE E.Arch_Notes Is Not Null OR E.Finish_Notes Is Not Null "
    q = q & "OR E.Condition_Notes Is Not Null OR E.Bio_Notes Is Not Null "
    q = q & "OR E.Materials_Notes Is Not Null OR E.Systems_Notes Is Not Null "
    q = q & "OR E.Mortar_Notes Is Not Null OR E.Vol_Notes Is Not Null "
    q = q & "OR E.Notes Is Not Null "
    q = q & "ORDER BY S.Site_Name, SC.Sector_Name, E.Code;"
    MkQuery db, "QRY_18_Notes_Review", q
End Sub

Private Sub BuildRockArtQuery(db As DAO.Database)
    Dim q As String
    q = "SELECT E.Code AS Parent_Code, "
    q = q & "IIF(T.Record_Class='Rock art panel','Isolated','Associated') AS Context, "
    q = q & "S.Site_Name, SC.Sector_Name, "
    q = q & "B.Code AS Position_Code, B.Name AS Position, "
    q = q & "DT.Name AS Dec_Type, D.Color, D.Color_Secondary, D.Substrate, "
    q = q & "D.Body_No, E.Facade_Orientation, E.Visibility_Valley, "
    q = q & "E.Coord_E_UTM, E.Coord_N_UTM, E.Altitude_masl, "
    q = q & "E.Facade_Observability, D.Notes "
    q = q & "FROM (((((T_DECORATIONS AS D "
    q = q & "INNER JOIN T_STRUCTURES AS E ON D.ID_Structure=E.ID) "
    q = q & "INNER JOIN L_STRUCT_BODY AS B ON D.ID_Struct_Body=B.ID) "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID) "
    q = q & "LEFT JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID) "
    q = q & "LEFT JOIN L_DEC_TYPE AS DT ON D.ID_Dec_Type=DT.ID "
    q = q & "WHERE B.Level_Type='ROC' "
    q = q & "ORDER BY S.Site_Name, SC.Sector_Name, E.Code;"
    MkQuery db, "QRY_17_RockArt_All", q
End Sub

Private Sub BuildValidationHelpers(db As DAO.Database)
    Dim m() As String
    Dim f() As String
    Dim g() As String
    FillElementMap m
    FillThreeValueFields f
    FillLayerFields g

    Dim q As String
    Dim i As Integer

    ' Long format: one row per structure per element
    q = ""
    For i = 0 To 19
        If i > 0 Then q = q & " UNION ALL "
        q = q & "SELECT E.ID AS ID_Structure, E.Code AS Code, "
        q = q & "'" & m(i, 0) & "' AS Element_Code, "
        q = q & "'" & m(i, 1) & "' AS Field_Name, "
        q = q & "E." & m(i, 1) & " AS Element_Value, "
        q = q & "'" & m(i, 2) & "' AS Sys_Field, "
        If Len(m(i, 2)) = 0 Then
            q = q & "Null AS Sys_Value "
        Else
            q = q & "E." & m(i, 2) & " AS Sys_Value "
        End If
        q = q & "FROM T_STRUCTURES AS E"
    Next i
    q = q & ";"
    MkQuery db, "QRY_16s_Element_Values", q

    ' Every (structure, element) pair that has justifying evidence. A
    ' Body or Whole structure row covers that structure's whole
    ' vocabulary at once (2), which the cross join expresses.
    q = "SELECT L.ID_Structure, L.Element_Code "
    q = q & "FROM T_LOST_ELEMENTS AS L "
    q = q & "WHERE L.Evidence_Scope='Element' AND L.Element_Code Is Not Null "
    q = q & "UNION "
    q = q & "SELECT L.ID_Structure, M.Code "
    q = q & "FROM T_LOST_ELEMENTS AS L, L_ELEMENTS AS M "
    q = q & "WHERE L.Evidence_Scope='Body' OR L.Evidence_Scope='Whole structure';"
    MkQuery db, "QRY_16s_Lost_Cover", q

    q = "SELECT V.ID_Structure, "
    q = q & "SUM(IIF(V.Element_Value=2 Or V.Element_Value=3,1,0)) AS N_Damaged "
    q = q & "FROM QRY_16s_Element_Values AS V "
    q = q & "GROUP BY V.ID_Structure;"
    MkQuery db, "QRY_16s_Damage_Count", q

    ' The manual-review checklist: every observational field still NULL
    q = ""
    For i = 0 To 19
        If i > 0 Then q = q & " UNION ALL "
        q = q & "SELECT E.ID AS ID_Structure, E.Code AS Code, '" & m(i, 1) & "' AS Field_Name "
        q = q & "FROM T_STRUCTURES AS E WHERE E." & m(i, 1) & " Is Null"
    Next i
    For i = 0 To 19
        q = q & " UNION ALL "
        q = q & "SELECT E.ID, E.Code, '" & f(i) & "' "
        q = q & "FROM T_STRUCTURES AS E WHERE E." & f(i) & " Is Null"
    Next i
    For i = 0 To 7
        q = q & " UNION ALL "
        q = q & "SELECT E.ID, E.Code, '" & g(i) & "' "
        q = q & "FROM T_STRUCTURES AS E WHERE E." & g(i) & " Is Null"
    Next i
    q = q & ";"
    MkQuery db, "QRY_16s_Observational_Nulls", q

    q = "SELECT E.ID AS ID_Structure, E.Code AS Code, "
    For i = 0 To 19
        If i > 0 Then q = q & "+"
        q = q & "IIF(E." & m(i, 1) & " Is Null,1,0)"
    Next i
    For i = 0 To 19
        q = q & "+IIF(E." & f(i) & " Is Null,1,0)"
    Next i
    For i = 0 To 7
        q = q & "+IIF(E." & g(i) & " Is Null,1,0)"
    Next i
    q = q & " AS N_Null FROM T_STRUCTURES AS E;"
    MkQuery db, "QRY_16s_Null_Count", q
End Sub

' ================================================================
'  QRY_16 - VALIDATION BATTERY, 31 RULES (section 12 + v12 + v13)
'
'  Non-blocking by design: a hard constraint would stop the
'  researcher recording a genuinely observed absence in a partly
'  collapsed structure. An empty result means the corpus is coherent.
'
'  THE TWO KINDS OF ZERO (rev. 5). Rules 4 and 6 demanded opposite
'  things until they were separated:
'      0 of assertion - the field's system is Present*/Attested lost,
'          or it has no system (I, R). "Examined, nothing there", so
'          on a collapse it is unverifiable and rule 6 flags it.
'      0 of padding   - the system is Not applicable / Not observable
'          / Absent. Asserts nothing, only avoids a NULL, and rule 4
'          actively requires it.
'  Without that split, a collapsed cave with Sys_Portal = Not
'  applicable would fire rule 6 on the very Sill/Jambs/Lintel zeros
'  rule 4 demands.
'
'  STORED IN THREE PARTS. All 26 rules parse and run on their own,
'  but a single UNION of the 31 resulting branches trips JET's
'  "query too complex" ceiling. QRY_16_Validation_Check unions the
'  parts, which are also separately inspectable during the review.
' ================================================================
Private Sub BuildValidationBattery(db As DAO.Database)
    Dim q As String

    q = R01()
    q = q & " UNION ALL " & R02Elements()
    q = q & " UNION ALL " & R02Systems()
    q = q & " UNION ALL " & R03()
    q = q & " UNION ALL " & R04()
    q = q & " UNION ALL " & R05()
    q = q & " UNION ALL " & R06()
    q = q & " UNION ALL " & R07()
    q = q & " UNION ALL " & R08()
    q = q & " UNION ALL " & R09()
    q = q & " UNION ALL " & R10()
    q = q & " UNION ALL " & R11()
    q = q & ";"
    MkQuery db, "QRY_16a_Rules_1_11", q

    q = R12()
    q = q & " UNION ALL " & R13()
    q = q & " UNION ALL " & R14()
    q = q & " UNION ALL " & R15()
    q = q & " UNION ALL " & R16()
    q = q & " UNION ALL " & R17()
    q = q & " UNION ALL " & R18()
    q = q & " UNION ALL " & R19()
    q = q & " UNION ALL " & R20()
    q = q & " UNION ALL " & R21()
    q = q & ";"
    MkQuery db, "QRY_16b_Rules_12_21", q

    q = R22()
    q = q & " UNION ALL " & R23()
    q = q & " UNION ALL " & R24()
    q = q & " UNION ALL " & R25()
    q = q & " UNION ALL " & R26()
    q = q & " UNION ALL " & R27()
    q = q & " UNION ALL " & R28()
    q = q & " UNION ALL " & R29()
    q = q & " UNION ALL " & R30()
    q = q & " UNION ALL " & R31()
    q = q & ";"
    MkQuery db, "QRY_16c_Rules_22_30", q

    q = "SELECT * FROM QRY_16a_Rules_1_11 "
    q = q & "UNION ALL SELECT * FROM QRY_16b_Rules_12_21 "
    q = q & "UNION ALL SELECT * FROM QRY_16c_Rules_22_31 "
    q = q & "ORDER BY Rule_No, Structure;"
    MkQuery db, "QRY_16_Validation_Check", q
End Sub

' R1 CARRIES TWO EXEMPTIONS THE DELTA DID NOT ANTICIPATE, added after
' the rule fired 13 times on the real corpus, all of them
' archaeologically sound.
'
' (a) NOT OBSERVABLE IS EXEMPT. Seeing a component while being unable
'     to resolve the system it might belong to is not a contradiction,
'     it is partial observability - the norm on a cliff face. The same
'     epistemic move as the REV5 acotacio of rule 6.
'
' (b) TIMBER_BRACKETS IS EXEMPT. Element E is the one item in the
'     vocabulary with an existence independent of its system: an
'     isolated corbel need not have carried a platform, and a corbel
'     on a funerary chamber whose platform cannot be affirmed does not
'     stop the structure being a funerary chamber. That is why
'     Timber_Bracket_Role exists, and why the MEN typology exists.
'     Table 4.5 lists E under Sys_Platform for GATING, which is a
'     different question from whether E implies H.
'
' F and G stay inside the rule: they are far harder to read as
' anything but platform components.
Private Function R01() As String
    Dim q As String
    q = "SELECT V.Code AS Structure, 1 AS Rule_No, "
    q = q & "'R1: ' & V.Field_Name & ' is present (' & V.Element_Value & ') but ' & V.Sys_Field & ' = ' & V.Sys_Value AS Rule_Violated, "
    q = q & "'Raise the system, or correct the component' AS Action "
    q = q & "FROM QRY_16s_Element_Values AS V "
    q = q & "WHERE V.Element_Value IN (1,2,3) AND V.Sys_Field<>'' "
    q = q & "AND V.Sys_Value Is Not Null "
    q = q & "AND V.Sys_Value Not Like 'Present*' AND V.Sys_Value<>'Attested lost' "
    q = q & "AND V.Sys_Value<>'Not observable' "
    q = q & "AND V.Field_Name<>'Timber_Brackets'"
    R01 = q
End Function

' LEFT JOIN, not a correlated NOT EXISTS: JET resolves outer
' references unreliably inside UNION branches.
Private Function R02Elements() As String
    Dim q As String
    q = "SELECT V.Code, 2, "
    q = q & "'R2: ' & V.Field_Name & ' coded 3 (attested lost) with no evidence row', "
    q = q & "'Add a T_LOST_ELEMENTS row, or use 0 / 9 instead' "
    q = q & "FROM QRY_16s_Element_Values AS V "
    q = q & "LEFT JOIN QRY_16s_Lost_Cover AS LC "
    q = q & "ON (V.ID_Structure=LC.ID_Structure) AND (V.Element_Code=LC.Element_Code) "
    q = q & "WHERE V.Element_Value=3 AND LC.ID_Structure Is Null"
    R02Elements = q
End Function

Private Function R02Systems() As String
    Dim s(2, 1) As String
    s(0, 0) = "Sys_Platform": s(0, 1) = "H"
    s(1, 0) = "Sys_Portal":   s(1, 1) = "P"
    s(2, 0) = "Sys_Eave":     s(2, 1) = "U"

    Dim q As String
    Dim i As Integer
    For i = 0 To 2
        If i > 0 Then q = q & " UNION ALL "
        q = q & "SELECT E.Code, 2, "
        q = q & "'R2: " & s(i, 0) & " is Attested lost with no evidence row', "
        q = q & "'Add a T_LOST_ELEMENTS row, or use Absent / Not observable' "
        q = q & "FROM T_STRUCTURES AS E "
        q = q & "WHERE E." & s(i, 0) & "='Attested lost' "
        q = q & "AND E.ID NOT IN (SELECT LC.ID_Structure FROM QRY_16s_Lost_Cover AS LC "
        q = q & "WHERE LC.Element_Code='" & s(i, 1) & "')"
    Next i
    R02Systems = q
End Function

Private Function R03() As String
    Dim s(2, 2) As String
    s(0, 0) = "Sys_Platform": s(0, 1) = "E.Timber_Brackets=0 AND E.Transverse_Beams=0 AND E.Corbelled_Courses=0": s(0, 2) = "E, F and G"
    s(1, 0) = "Sys_Portal":   s(1, 1) = "E.Sill=0 AND E.Jambs=0 AND E.Lintel=0":                                  s(1, 2) = "N, O and Q"
    s(2, 0) = "Sys_Eave":     s(2, 1) = "E.Eave_Beam=0 AND E.Eave_Surface=0":                                     s(2, 2) = "S and T"

    Dim q As String
    Dim i As Integer
    For i = 0 To 2
        If i > 0 Then q = q & " UNION ALL "
        q = q & "SELECT E.Code, 3, "
        q = q & "'R3: " & s(i, 0) & " is present but " & s(i, 2) & " are all 0', "
        q = q & "'A present system needs at least one component at 1, 2 or 3' "
        q = q & "FROM T_STRUCTURES AS E "
        q = q & "WHERE E." & s(i, 0) & " Like 'Present*' AND " & s(i, 1)
    Next i
    R03 = q
End Function

Private Function R04() As String
    Dim q As String
    q = "SELECT V.Code, 4, "
    q = q & "'R4: ' & V.Field_Name & ' = ' & V.Element_Value & ' although ' & V.Sys_Field & ' is Not applicable', "
    q = q & "'Components of a non-applicable system are recorded as 0' "
    q = q & "FROM QRY_16s_Element_Values AS V "
    q = q & "WHERE V.Sys_Value='Not applicable' AND V.Element_Value<>0"
    R04 = q
End Function

' 30% of the 20 element fields is 6, so the test is "more than 6".
Private Function R05() As String
    Dim q As String
    q = "SELECT E.Code, 5, "
    q = q & "'R5: recorded Good but more than 30% of elements are partial or lost', "
    q = q & "'Reconcile ID_Arch_Status with the element values' "
    q = q & "FROM ((T_STRUCTURES AS E "
    q = q & "INNER JOIN L_STATUS AS ST ON E.ID_Arch_Status=ST.ID) "
    q = q & "INNER JOIN QRY_16s_Damage_Count AS DC ON E.ID=DC.ID_Structure) "
    q = q & "WHERE ST.Name='Good' AND DC.N_Damaged>6"
    R05 = q
End Function

Private Function R06() As String
    Dim q As String
    q = "SELECT V.Code, 6, "
    q = q & "'R6: ' & V.Field_Name & ' coded 0 (absent) on a Collapsed structure', "
    q = q & "'Absence is not verifiable on a collapse: use 3 with evidence, or 9' "
    q = q & "FROM ((QRY_16s_Element_Values AS V "
    q = q & "INNER JOIN T_STRUCTURES AS E ON V.ID_Structure=E.ID) "
    q = q & "INNER JOIN L_STATUS AS ST ON E.ID_Arch_Status=ST.ID) "
    q = q & "WHERE ST.Name='Collapsed' AND V.Element_Value=0 "
    q = q & "AND (V.Sys_Field='' OR V.Sys_Value Like 'Present*' OR V.Sys_Value='Attested lost')"
    R06 = q
End Function

Private Function R07() As String
    Dim q As String
    q = "SELECT E.Code, 7, "
    q = q & "'R7: Platform_Surface_Material set although the corbels are Isolated', "
    q = q & "'Clear the material, or correct Timber_Bracket_Role' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Platform_Surface_Material Is Not Null "
    q = q & "AND E.Timber_Bracket_Role='Isolated'"
    R07 = q
End Function

Private Function R08() As String
    Dim q As String
    q = "SELECT E.Code, 8, "
    q = q & "'R8: corbels present but Timber_Bracket_Count is empty', "
    q = q & "'Count the corbels: the number carries the platform argument' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Timber_Brackets IN (1,2) AND E.Timber_Bracket_Count Is Null"
    R08 = q
End Function

Private Function R09() As String
    Dim q As String
    q = "SELECT E.Code, 9, "
    q = q & "'R9: corbels present but Timber_Bracket_Role is not recorded', "
    q = q & "'Set the role: platform support or isolated corbel' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Timber_Brackets IN (1,2) "
    q = q & "AND (E.Timber_Bracket_Role Is Null Or E.Timber_Bracket_Role='ND')"
    R09 = q
End Function

Private Function R10() As String
    Dim q As String
    q = "SELECT E.Code, 10, "
    q = q & "'R10: Lintel is 0 (absent) but Lintel_Material is set', "
    q = q & "'Clear Lintel_Material, or correct the Lintel value' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Lintel=0 AND E.Lintel_Material Is Not Null"
    R10 = q
End Function

Private Function R11() As String
    Dim q As String
    q = "SELECT E.Code, 11, "
    q = q & "'R11: chamber bodies counted but Sys_Portal is Absent', "
    q = q & "'A body is N1 only if it has (or had) an access opening' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.N_Chamber_Bodies>0 AND E.Sys_Portal='Absent'"
    R11 = q
End Function

' First branch of QRY_16b: carries the explicit aliases that name
' that half's columns when it is opened on its own.
Private Function R12() As String
    Dim q As String
    q = "SELECT E.Code AS Structure, 12 AS Rule_No, "
    q = q & "'R12: pigment present but the substrate is not recorded' AS Rule_Violated, "
    q = q & "'Set Pigment_Substrate: plaster vs masonry vs bedrock is the point' AS Action "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Pigment_Present=1 "
    q = q & "AND (E.Pigment_Substrate Is Null Or E.Pigment_Substrate='ND')"
    R12 = q
End Function

Private Function R13() As String
    Dim q As String
    q = "SELECT E.Code, 13, "
    q = q & "'R13: pigment recorded on plaster but plaster is absent', "
    q = q & "'Reconcile Plaster_Present with Pigment_Substrate' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Pigment_Substrate='Plaster' AND E.Plaster_Present=0"
    R13 = q
End Function

Private Function R14() As String
    Dim q As String
    q = "SELECT E.Code, 14, "
    q = q & "'R14: chamber roof present but its type is not recorded', "
    q = q & "'Set Chamber_Roof_Type: natural bedrock vs built is the decision' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Chamber_Roof IN (1,2) "
    q = q & "AND (E.Chamber_Roof_Type Is Null Or E.Chamber_Roof_Type='ND')"
    R14 = q
End Function

Private Function R15() As String
    Dim q As String
    q = "SELECT E.Code, 15, "
    q = q & "'R15: classified as a structural trace but human remains are recorded', "
    q = q & "'Reclassify the record, or correct Human_Remains' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID "
    q = q & "WHERE T.Record_Class='Structural trace' AND E.Human_Remains=1"
    R15 = q
End Function

' "Filled" means real content: a bio or materials field at 1, or an
' MNI. The padding zeros the gating leaves behind do not count, which
' is what stops this rule contradicting rule 17.
Private Function R16() As String
    Dim q As String
    q = "SELECT E.Code, 16, "
    q = q & "'R16: bio or materials content on a record whose class closes those tabs', "
    q = q & "'Reclassify the record, or clear the content' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID "
    q = q & "WHERE T.Record_Class IN ('Structural trace','Rock art panel') "
    q = q & "AND (E.MNI Is Not Null "
    q = q & "OR E.Human_Remains=1 OR E.Anatomical_Connection=1 OR E.Mummification=1 "
    q = q & "OR E.Funerary_Bundles=1 OR E.Dispersed_Remains=1 OR E.Flexed_Position=1 "
    q = q & "OR E.Bone_Burning=1 "
    q = q & "OR E.Mat_Textiles=1 OR E.Mat_Wood=1 OR E.Mat_VegFiber=1 "
    q = q & "OR E.Mat_Ceramics=1 OR E.Mat_Fauna=1 OR E.Mat_DeerAntler=1 "
    q = q & "OR E.Mat_Other=1)"
    R16 = q
End Function

' Not a complaint: the manual-review worklist of section 10. One row
' per structure with a count; the field by field detail is in
' QRY_16s_Observational_Nulls.
Private Function R17() As String
    Dim q As String
    q = "SELECT NC.Code, 17, "
    q = q & "'R17: ' & NC.N_Null & ' observational field(s) still NULL (not yet assessed)', "
    q = q & "'Resolve each to 0, 1, 2, 3 or 9 - see QRY_16s_Observational_Nulls' "
    q = q & "FROM QRY_16s_Null_Count AS NC "
    q = q & "WHERE NC.N_Null>0"
    R17 = q
End Function

Private Function R18() As String
    Dim q As String
    q = "SELECT E.Code, 18, "
    q = q & "'R18: human remains recorded but MNI is empty', "
    q = q & "'Set the minimum number of individuals' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Human_Remains=1 AND E.MNI Is Null"
    R18 = q
End Function

Private Function R19() As String
    Dim q As String
    q = "SELECT E.Code, 19, "
    q = q & "'R19: cultural materials declared absent or not observable but a Mat_ field is not 0/9', "
    q = q & "'Set the Mat_ fields to match, or correct Cultural_Materials_Present' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Cultural_Materials_Present=0 "
    q = q & "AND (E.Mat_Textiles<>0 OR E.Mat_Wood<>0 OR E.Mat_VegFiber<>0 "
    q = q & "OR E.Mat_Ceramics<>0 OR E.Mat_Fauna<>0 OR E.Mat_DeerAntler<>0 OR E.Mat_Other<>0)"
    R19 = q
End Function

' Direction is read from the joint: only an abutted vertical joint or
' a superposition can carry it (9.1bis). A bonded joint implies
' Contemporary; the rest imply nothing.
Private Function R20() As String
    Dim q As String
    q = "SELECT A.Code, 20, "
    q = q & "'R20: directional chronology on a connection type that implies no abutment', "
    q = q & "'Use Abutted vertical joint / Superposition, or set Undetermined' "
    q = q & "FROM T_CONNECTIONS AS C "
    q = q & "INNER JOIN T_STRUCTURES AS A ON C.ID_Struct_A=A.ID "
    q = q & "WHERE C.Chrono_Relation IN ('A earlier than B','B earlier than A') "
    q = q & "AND (C.Connection_Type Is Null "
    q = q & "OR (C.Connection_Type<>'Abutted vertical joint' "
    q = q & "AND C.Connection_Type<>'Superposition'))"
    R20 = q
End Function

Private Function R21() As String
    Dim q As String
    q = "SELECT E.Code, 21, "
    q = q & "'R21: lost-element row scoped Element but no Element_Code given', "
    q = q & "'Element_Code is required when the scope is Element' "
    q = q & "FROM T_LOST_ELEMENTS AS L "
    q = q & "INNER JOIN T_STRUCTURES AS E ON L.ID_Structure=E.ID "
    q = q & "WHERE L.Evidence_Scope='Element' AND L.Element_Code Is Null"
    R21 = q
End Function

' ================================================================
'  RULES 22-26 (v12) - the corpus-level counterpart of the form
'  gating. Non-blocking, like the rest of the battery.
' ================================================================

' R22/R23: Dec_Present and T_DECORATIONS must agree. Two directions,
' two rules, so the worklist names which side to fix.
Private Function R22() As String
    Dim q As String
    q = "SELECT E.Code, 22, "
    q = q & "'R22: architectural decoration declared present but no non-ROC row', "
    q = q & "'Add the decoration rows, or correct Dec_Present' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Dec_Present=1 "
    q = q & "AND E.ID NOT IN (SELECT D.ID_Structure FROM T_DECORATIONS AS D "
    q = q & "INNER JOIN L_STRUCT_BODY AS B ON D.ID_Struct_Body=B.ID "
    q = q & "WHERE B.Level_Type<>'ROC')"
    R22 = q
End Function

Private Function R23() As String
    Dim q As String
    q = "SELECT E.Code, 23, "
    q = q & "'R23: non-ROC decoration rows exist but Dec_Present is not 1', "
    q = q & "'Set Dec_Present, or move the rows to a ROC position' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE (E.Dec_Present<>1 OR E.Dec_Present Is Null) "
    q = q & "AND E.ID IN (SELECT D.ID_Structure FROM T_DECORATIONS AS D "
    q = q & "INNER JOIN L_STRUCT_BODY AS B ON D.ID_Struct_Body=B.ID "
    q = q & "WHERE B.Level_Type<>'ROC')"
    R23 = q
End Function

' R24/R25: under presence 0 or 9 the detail is INAPPLICABLE, not
' undetermined - the same logic M5 applied to Lintel_Material.
Private Function R24() As String
    Dim q As String
    q = "SELECT E.Code, 24, "
    q = q & "'R24: plaster detail recorded although Plaster_Present is 0 or 9', "
    q = q & "'Clear Plaster_Color / Plaster_Extent, or correct Plaster_Present' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Plaster_Present IN (0,9) "
    q = q & "AND (E.Plaster_Color Is Not Null OR E.Plaster_Extent Is Not Null)"
    R24 = q
End Function

Private Function R25() As String
    Dim q As String
    q = "SELECT E.Code, 25, "
    q = q & "'R25: pigment detail recorded although Pigment_Present is 0 or 9', "
    q = q & "'Clear the pigment detail fields, or correct Pigment_Present' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Pigment_Present IN (0,9) "
    q = q & "AND (E.Pigment_Substrate Is Not Null OR E.Pigment_Color Is Not Null OR E.Pigment_Extent Is Not Null)"
    R25 = q
End Function

' R26: no mortar means dry-laid; any other type contradicts it.
Private Function R26() As String
    Dim q As String
    q = "SELECT E.Code, 26, "
    q = q & "'R26: Mortar_Present is 0 but Mortar_Type is not None dry-laid', "
    q = q & "'Set Mortar_Type to None dry-laid or clear it, or correct Mortar_Present' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Mortar_Present=0 "
    q = q & "AND E.Mortar_Type Is Not Null AND E.Mortar_Type<>'None dry-laid'"
    R26 = q
End Function

' ================================================================
'  RULES 27-30 (v13) - see DELTA_v12_v13.md
' ================================================================

' R27 (delta 4): a basal body is a MASS OF MASONRY, not a platform.
' H is a system and an interface element: it closes N0 and
' preconditions N1, but it is not a superposed body. If it were
' counted as one, a mausoleum on a corbelled platform and one on a
' masonry podium would both read 2,1 - and those are precisely two
' constructive solutions that must stay distinguishable.
' Fired on 3 of 31 records in the v12 corpus: a low enough rate to
' be signal rather than noise, but check whether the record is
' simply half-entered before correcting it.
Private Function R27() As String
    Dim q As String
    q = "SELECT E.Code, 27, "
    q = q & "'R27: basal bodies counted but Base_Level (B) is 0', "
    q = q & "'A basal body is a masonry mass. If it is a platform, use Sys_Platform and set N_Basal_Bodies to 0' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.N_Basal_Bodies>0 AND E.Base_Level=0"
    R27 = q
End Function

' R28 / R29 (delta 7.5): the ROC counterpart of rules 22-23.
' RockArt_Present governs the ROC half of T_DECORATIONS exactly as
' Dec_Present governs the architectural half.
Private Function R28() As String
    Dim q As String
    q = "SELECT E.Code, 28, "
    q = q & "'R28: rock art declared present but no ROC row', "
    q = q & "'Add the rock art rows, or correct RockArt_Present' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.RockArt_Present=1 "
    q = q & "AND E.ID NOT IN (SELECT D.ID_Structure FROM T_DECORATIONS AS D "
    q = q & "INNER JOIN L_STRUCT_BODY AS B ON D.ID_Struct_Body=B.ID "
    q = q & "WHERE B.Level_Type='ROC')"
    R28 = q
End Function

Private Function R29() As String
    Dim q As String
    q = "SELECT E.Code, 29, "
    q = q & "'R29: ROC rows exist but RockArt_Present is not 1', "
    q = q & "'Set RockArt_Present, or move the rows to an architectural position' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE (E.RockArt_Present<>1 OR E.RockArt_Present Is Null) "
    q = q & "AND E.ID IN (SELECT D.ID_Structure FROM T_DECORATIONS AS D "
    q = q & "INNER JOIN L_STRUCT_BODY AS B ON D.ID_Struct_Body=B.ID "
    q = q & "WHERE B.Level_Type='ROC')"
    R29 = q
End Function

' R31 (rev. 5): implausible linear dimension. Cheap, and it
' catches precisely the error the cm -> m conversion introduces -
' entering 62 while thinking in centimetres would give a 62-metre
' ledge, and no other rule would notice.
Private Function R31() As String
    Dim q As String
    q = "SELECT E.Code, 31, "
    q = q & "'R31: implausible linear dimension (over 20 m) - check the unit', "
    q = q & "'All linear fields are in METRES since rev. 5' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Length_m>20 OR E.Width_m>20 OR E.Height_m>20 "
    q = q & "OR E.Support_Width_m>20 OR E.Support_Depth_m>20 "
    q = q & "OR E.Opening_Width_m>20 OR E.Opening_Height_m>20"
    R31 = q
End Function

' R30 (delta 7.9): NOT a prohibition - an alert. Pigment on bedrock
' following architectural elements is THE DIAGNOSTIC CASE of a lost
' body: vertical bands on the rock aligned with the jambs of the
' body below, where the architecture was there and the paint has
' outlived it. Blocking the combination would prevent recording the
' very evidence that justifies a value 3 with 'Pigment on bedrock'.
Private Function R30() As String
    Dim q As String
    q = "SELECT E.Code, 30, "
    q = q & "'R30: pigment on bedrock following architectural elements - possible lost body', "
    q = q & "'If a body has vanished, open a T_LOST_ELEMENTS row scoped Body. Otherwise ignore' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Pigment_Substrate='Bedrock' "
    q = q & "AND E.Pigment_Extent='Architectural elements' "
    q = q & "AND E.ID NOT IN (SELECT L.ID_Structure FROM T_LOST_ELEMENTS AS L)"
    R30 = q
End Function
