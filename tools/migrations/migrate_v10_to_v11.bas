Option Compare Database
Option Explicit

' ----------------------------------------------------------------
'  FINAL DELETIONS SWITCH (delta step 13.14)
'
'  Leave this False until the manual review is finished. Setting it
'  True still does not delete anything on its own: M11a and M11b each
'  re-check their own preconditions against the live data first and
'  refuse to run if the replacement information is not in place. This
'  flag only says "I intend to delete", the checks say "it is safe to".
'
'  With it False, every run still REPORTS how close each precondition
'  is, so it doubles as a progress meter for the review.
' ----------------------------------------------------------------
Private Const ALLOW_FINAL_DELETIONS As Boolean = False

' ================================================================
'  CHACHAPOYA ARCHAEOLOGICAL DATABASE - MIGRATION SCRIPT v10 -> v11
'  La Petaca & Diablo Wasi (Leymebamba, Amazonas, Peru)
'  Author: Esteve Ribera Torro | TFM Arqueologia UA
'  Spec: DELTA_v10_v11.md (rev. 5)
'
'  PURPOSE: transform the EXISTING v10 database (35 records in
'  T_STRUCTURES, 29 in T_DECORATIONS) into the v11 schema, keeping
'  every existing record intact. Delivered in independently testable
'  blocks, in the order fixed by DELTA_v10_v11.md section 13.
'
'  RULES THIS SCRIPT FOLLOWS THROUGHOUT (see plan):
'   A. Strict order: rename -> add -> populate (UPDATE) -> delete.
'      No field is ever removed before its replacement is populated
'      and verified.
'   B. Existing NULLs in observational fields are NEVER touched.
'      DEFAULT 0 (via DAO) only ever applies to NEW records.
'   C. Lintel TEXT -> BYTE: rename to Lintel_Material (DAO, keeps
'      data) -> new Lintel BYTE -> populate by UPDATE -> verify count
'      -> Lintel_Material is never deleted.
'   D. Corbelled_Platform / Access_Opening / Eave -> Sys_* : only the
'      already-recorded 0 and 9 values are migrated by script; 1 and
'      NULL are left NULL on the new field (pending manual review).
'      The 4 old fields are deleted only in the final block (13.14).
'   E. Debug.Print at the start and end of every block, with counts
'      of affected records, so a failure is always locatable.
'
'  STATUS OF THIS FILE (updated at every delivery):
'    [DONE]    M1  - New lookups (L_ELEMENTS, L_LOST_EVIDENCE),
'                     typology preconditions (EA09 MIX->CAV,
'                     Timber_Bracket_Role / Platform_Surface_Material
'                     fix), L_SUPPORT expansion, L_TYPOLOGY rebuild
'                     (Record_Class, ND split, MIX removal).
'    [DONE]    M3  - Domain defaults (SetByteDefaults v11): DEFAULT 0
'                     via DAO on the 19 five-value element fields plus
'                     the 24 relabelled three-value fields. No existing
'                     value is touched - only affects future new rows.
'    [DONE]    M4  - Field renames via DAO (Return_Wall, Facade_Flank,
'                     Recessed_Frame, Rear_Closure_Type) + the
'                     L_STRUCT_BODY LWF -> FFL entry.
'    [DONE]    M5  - Lintel TEXT -> BYTE split: rename to
'                     Lintel_Material, new Lintel BYTE, 21 values
'                     mapped per delta 5.4, ~14 NULLs preserved.
'    [DONE]    M6  - The five Sys_* fields + Platform_Function, and the
'                     scripted migration of the recorded 0/9 values of
'                     H, P and U. Never overwrites manual work.
'    [DONE]    M7  - T_LOST_ELEMENTS + its 4 relationships, and
'                     Chrono_Relation on T_CONNECTIONS. Structure
'                     only: both are populated by hand.
'    [DONE]    M9a - QRY_02 rebuilt on T_DECORATIONS (+ the new
'                     QRY_02a flag query), QRY_05 and QRY_07 updated.
'    [DONE]    M9b - QRY_13 rebuilt (20 AX + 5 SYS columns, NA
'                     semantics) and QRY_14 with Chrono_Relation.
'    [DONE]    M10a - QRY_16s_Element_Values (long-format helper) and
'                     QRY_16_Validation_Check with rules 1-11.
'    [DONE]    M10b - QRY_16 rules 12-21; QRY_16 now carries all 21.
'    [READY]   M11a - Deletes the 12 Dec_*/RA_*/Rock_Art booleans.
'    [READY]   M11b - Deletes Lost_Body_Evidence and the four absorbed
'                     H/P/U/W fields.
'                     Both are gated by ALLOW_FINAL_DELETIONS above AND
'                     by live precondition checks. Until then they only
'                     report how far the manual review has got.
'
'  Run on the EXISTING v10 database (NOT a blank one).
'  Run VerifyMigration() after each block to sanity-check progress.
' ================================================================

Sub MigrateV10toV11()
    Dim db As DAO.Database
    Set db = CurrentDb()

    Debug.Print "================================================================"
    Debug.Print "MIGRATE V10 -> V11 : RUN START (" & Now & ")"
    Debug.Print "================================================================"

    M1_NewLookupsAndTypology db
    M3_DomainAndDefaults db
    M4_FieldRenames db
    M5_LintelSplit db
    M6_SysFields db
    M7_LostElementsAndConnections db
    M9a_QueriesDecorationAndExports db
    M9b_QueriesAXMatrix db
    M10a_QRY16_Rules1to11 db
    M10b_QRY16_Rules12to21 db

    ' Step 13.14. These check their own preconditions and honour
    ' ALLOW_FINAL_DELETIONS; with the flag False they only report.
    M11a_DeleteDecorationBooleans db
    M11b_DeleteAbsorbedFields db
    ' M11a_DeleteDecorationBooleans db
    ' M11b_DeleteAbsorbedFields db

    Debug.Print "================================================================"
    Debug.Print "MIGRATE V10 -> V11 : RUN COMPLETE (blocks: M1, M3-M7, M9a, M9b, M10a, M10b, M11a, M11b)"
    Debug.Print "================================================================"

    Set db = Nothing
End Sub

' ================================================================
'  SHARED HELPERS
' ================================================================
Private Function TblExists(db As DAO.Database, n As String) As Boolean
    Dim t As DAO.TableDef
    For Each t In db.TableDefs
        If t.Name = n Then TblExists = True: Exit Function
    Next t
End Function

Private Function FldExists(db As DAO.Database, tbl As String, fld As String) As Boolean
    Dim f As DAO.Field
    If Not TblExists(db, tbl) Then Exit Function
    For Each f In db.TableDefs(tbl).Fields
        If f.Name = fld Then FldExists = True: Exit Function
    Next f
End Function

' The 20 element fields of the A-X vocabulary (delta 1.5) with the
' system that gates each one (delta 4.5). SINGLE SOURCE OF TRUTH:
' the QRY_13 NA gating and the QRY_16 validation rules both read it,
' so they can never end up disagreeing about which element belongs to
' which system - a disagreement that would make the validation battery
' contradict the export it is supposed to validate.
' Columns: 0 = letter, 1 = v11 field name, 2 = gating Sys_* ("" = none;
' elements I and R belong to no system).
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

' The 24 observational fields that keep the three-value 0/1/9 domain
' (delta 1.6). Together with the 20 of FillElementMap these are the 44
' fields where a NULL means "not yet assessed" and which rule 17 tracks.
' KEEP IN SYNC with the arrays in M3_DomainAndDefaults, which lists the
' same 24 but under whichever name they carry before the M4 renames.
Private Sub FillThreeValueFields(f() As String)
    ReDim f(23)
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
    f(20) = "Mortar_Present"
    f(21) = "Chinking_Stones"
    f(22) = "Plaster_Present"
    f(23) = "Pigment_Present"
End Sub

' ================================================================
'  BLOCK M1 - New lookups + typology preconditions + EA09 fixes
'  DELTA sections: 2 (T_LOST_ELEMENTS lookup), 3 (L_ELEMENTS),
'  5.7 (L_SUPPORT), 6.1/6.2/6.3 (L_TYPOLOGY), 11 (EA09).
'  DELTA section 13 steps covered: 1, 2, 3, 4 (L_SUPPORT part only).
' ================================================================
Private Sub M1_NewLookupsAndTypology(db As DAO.Database)
    Debug.Print "---- [M1] START ----"

    M1_CreateAndPopulate_L_ELEMENTS db
    M1_CreateAndPopulate_L_LOST_EVIDENCE db
    M1_ExpandLSupport db
    M1_ReclassifyEA09_MixToCav db
    M1_CorrectEA09_BracketAndPlatformMaterial db
    M1_RebuildLTypology db

    db.TableDefs.Refresh
    Debug.Print "---- [M1] END ----"
End Sub

' ----------------------------------------------------------------
'  M1a - L_ELEMENTS (24 rows, A..X vocabulary, bottom-to-top)
'  UQ_ELEM_CODE unique index is required before M7 creates the
'  T_LOST_ELEMENTS.Element_Code -> L_ELEMENTS.Code relationship.
' ----------------------------------------------------------------
Private Sub M1_CreateAndPopulate_L_ELEMENTS(db As DAO.Database)
    If TblExists(db, "L_ELEMENTS") Then
        Debug.Print "[M1] L_ELEMENTS already exists - skipped"
        Exit Sub
    End If

    Dim sql As String
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
    Debug.Print "[M1] L_ELEMENTS table created (with unique index UQ_ELEM_CODE on Code)"

    ' Sys_Group domain per DELTA section 3: Platform / Portal / Eave /
    ' empty. Sys_Base and Sys_Chamber are NOT in that domain (they are
    ' grouping fields without their own vocabulary letter, section 4.2),
    ' so their component elements (A-D, J/K/L/M/V/X) get an empty
    ' Sys_Group here, exactly as specified.
    ' Level_Type for Q/X/S/T/U is not given letter-by-letter in the
    ' delta; assigned here by analogy with the v10 code banner comments
    ' ("Level N0: elements A-H", "Level N1: elements J-P, R, V, W",
    ' "Upper zone: S, T, U, X"). Adjust later if the researcher's
    ' vocabulary disagrees - this table carries no computed logic yet.
    ' Columns: 0=Code 1=Name_EN 2=Name_VAL 3=Level_Type 4=Sys_Group
    '          5=Is_System 6=Field_Name 7=Description
    Dim el(23, 7) As String

    el(0, 0) = "A":  el(0, 1) = "Embedded base beams":         el(0, 2) = "Jaceres basals":         el(0, 3) = "N0":    el(0, 4) = "":         el(0, 5) = "False": el(0, 6) = "Embedded_Base_Beams":  el(0, 7) = "Timber beams embedded in the basal masonry."
    el(1, 0) = "B":  el(1, 1) = "Base level":                  el(1, 2) = "Basament":               el(1, 3) = "N0":    el(1, 4) = "":         el(1, 5) = "False": el(1, 6) = "Base_Level":           el(1, 7) = "Constructed basal level supporting the structure."
    el(2, 0) = "C":  el(2, 1) = "Decorative socle":            el(2, 2) = "Socol decoratiu":        el(2, 3) = "N0":    el(2, 4) = "":         el(2, 5) = "False": el(2, 6) = "Decorative_Socle":     el(2, 7) = "Decorative treatment of the basal mass."
    el(3, 0) = "D":  el(3, 1) = "Tie walls":                   el(3, 2) = "Muret transversal":      el(3, 3) = "N0":    el(3, 4) = "":         el(3, 5) = "False": el(3, 6) = "Tie_Walls":            el(3, 7) = "Transverse tie wall at base level."
    el(4, 0) = "E":  el(4, 1) = "Timber brackets (corbels)":   el(4, 2) = "Mensules de fusta":      el(4, 3) = "N0":    el(4, 4) = "Platform": el(4, 5) = "False": el(4, 6) = "Timber_Brackets":      el(4, 7) = "Protruding horizontal timber corbels; component of the platform system (H)."
    el(5, 0) = "F":  el(5, 1) = "Transverse beams":            el(5, 2) = "Bigues transversals":    el(5, 3) = "N0":    el(5, 4) = "Platform": el(5, 5) = "False": el(5, 6) = "Transverse_Beams":     el(5, 7) = "Spanning beams; component of the platform system (H)."
    el(6, 0) = "G":  el(6, 1) = "Corbelled courses":           el(6, 2) = "Filades en voladis":     el(6, 3) = "N0":    el(6, 4) = "Platform": el(6, 5) = "False": el(6, 6) = "Corbelled_Courses":    el(6, 7) = "Masonry courses projecting in corbel; component of the platform system (H)."
    el(7, 0) = "H":  el(7, 1) = "Corbelled platform (system)": el(7, 2) = "Plataforma en voladis":  el(7, 3) = "N0":    el(7, 4) = "Platform": el(7, 5) = "True":  el(7, 6) = "Sys_Platform":         el(7, 7) = "Composite system: Timber_Brackets + Transverse_Beams + Corbelled_Courses."
    el(8, 0) = "I":  el(8, 1) = "Interbody cornice":           el(8, 2) = "Cornisa intercos":       el(8, 3) = "N0-N1": el(8, 4) = "":         el(8, 5) = "False": el(8, 6) = "Interbody_Cornice":    el(8, 7) = "Cornice zone between superposed bodies. Not gated by any system (section 4.5)."
    el(9, 0) = "J":  el(9, 1) = "Corner quoins":               el(9, 2) = "Cantoneres":             el(9, 3) = "N1":    el(9, 4) = "":         el(9, 5) = "False": el(9, 6) = "Corner_Quoins":        el(9, 7) = "Larger stones set vertically at the corners; component of Sys_Chamber."
    el(10, 0) = "K": el(10, 1) = "Structural pilasters":       el(10, 2) = "Pilastres estructurals": el(10, 3) = "N1":  el(10, 4) = "":        el(10, 5) = "False": el(10, 6) = "Structural_Pilasters": el(10, 7) = "Full-height pilaster integrated in the wall plane; component of Sys_Chamber."
    el(11, 0) = "L": el(11, 1) = "Facade flank":               el(11, 2) = "Ala de facana":         el(11, 3) = "N1":   el(11, 4) = "":        el(11, 5) = "False": el(11, 6) = "Facade_Flank":         el(11, 7) = "Wall face in the facade plane, flanking the opening; component of Sys_Chamber."
    el(12, 0) = "M": el(12, 1) = "Relief frieze":              el(12, 2) = "Fris en relleu":        el(12, 3) = "N1":   el(12, 4) = "":        el(12, 5) = "False": el(12, 6) = "Relief_Frieze":        el(12, 7) = "Architectural relief frieze; component of Sys_Chamber."
    el(13, 0) = "N": el(13, 1) = "Sill":                       el(13, 2) = "Llindar":               el(13, 3) = "N1":   el(13, 4) = "Portal": el(13, 5) = "False": el(13, 6) = "Sill":                 el(13, 7) = "Lower frame element of the access opening; component of the portal system (P)."
    el(14, 0) = "O": el(14, 1) = "Jambs":                      el(14, 2) = "Brancals":              el(14, 3) = "N1":   el(14, 4) = "Portal": el(14, 5) = "False": el(14, 6) = "Jambs":                el(14, 7) = "Vertical frame elements of the access opening; component of the portal system (P)."
    el(15, 0) = "P": el(15, 1) = "Access portal (system)":     el(15, 2) = "Sistema portal":        el(15, 3) = "N1":   el(15, 4) = "Portal": el(15, 5) = "True":  el(15, 6) = "Sys_Portal":           el(15, 7) = "Composite system: Sill + Jambs + Lintel."
    el(16, 0) = "Q": el(16, 1) = "Lintel":                     el(16, 2) = "Dintell":               el(16, 3) = "N1":   el(16, 4) = "Portal": el(16, 5) = "False": el(16, 6) = "Lintel":               el(16, 7) = "Upper frame element of the access opening; component of the portal system (P)."
    el(17, 0) = "R": el(17, 1) = "Upper crown":                el(17, 2) = "Coronament":            el(17, 3) = "N1":   el(17, 4) = "":       el(17, 5) = "False": el(17, 6) = "Upper_Crown":          el(17, 7) = "Upper crown or coping of the body. Not gated by any system (section 4.5)."
    el(18, 0) = "S": el(18, 1) = "Eave beam":                  el(18, 2) = "Biga de suport rafec":  el(18, 3) = "SUP":  el(18, 4) = "Eave":   el(18, 5) = "False": el(18, 6) = "Eave_Beam":            el(18, 7) = "Supporting beam of the eave; component of the eave system (U)."
    el(19, 0) = "T": el(19, 1) = "Eave surface":               el(19, 2) = "Superficie de rafec":   el(19, 3) = "SUP":  el(19, 4) = "Eave":   el(19, 5) = "False": el(19, 6) = "Eave_Surface":         el(19, 7) = "Finished surface of the eave; component of the eave system (U)."
    el(20, 0) = "U": el(20, 1) = "Eave (system)":              el(20, 2) = "Rafec en voladis":      el(20, 3) = "SUP":  el(20, 4) = "Eave":   el(20, 5) = "True":  el(20, 6) = "Sys_Eave":             el(20, 7) = "Composite system: Eave_Beam + Eave_Surface."
    el(21, 0) = "V": el(21, 1) = "Return wall":                el(21, 2) = "Mur de retorn":         el(21, 3) = "N1":   el(21, 4) = "":       el(21, 5) = "False": el(21, 6) = "Return_Wall":          el(21, 7) = "Wall perpendicular to the facade plane, returning toward the cliff; component of Sys_Chamber."
    el(22, 0) = "W": el(22, 1) = "Rear wall":                  el(22, 2) = "Mur posterior":         el(22, 3) = "N1":   el(22, 4) = "":       el(22, 5) = "False": el(22, 6) = "":                     el(22, 7) = "Vocabulary entry retained for reference; no own field in v11 (absorbed by Sys_Chamber, delta 5.5)."
    el(23, 0) = "X": el(23, 1) = "Chamber roof":               el(23, 2) = "Coberta de cambra":     el(23, 3) = "N1":   el(23, 4) = "":       el(23, 5) = "False": el(23, 6) = "Chamber_Roof":         el(23, 7) = "Element closing the chamber above; component of Sys_Chamber."

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
    Debug.Print "[M1] L_ELEMENTS populated: " & DCount("*", "L_ELEMENTS") & " rows (expected 24)"
End Sub

' ----------------------------------------------------------------
'  M1b - L_LOST_EVIDENCE (9 rows, section 2)
' ----------------------------------------------------------------
Private Sub M1_CreateAndPopulate_L_LOST_EVIDENCE(db As DAO.Database)
    If TblExists(db, "L_LOST_EVIDENCE") Then
        Debug.Print "[M1] L_LOST_EVIDENCE already exists - skipped"
        Exit Sub
    End If

    Dim sql As String
    sql = "CREATE TABLE L_LOST_EVIDENCE ("
    sql = sql & "ID COUNTER CONSTRAINT PK_LEV PRIMARY KEY,"
    sql = sql & "Name TEXT(60) NOT NULL,"
    sql = sql & "Description TEXT(150))"
    db.Execute sql, dbFailOnError
    Debug.Print "[M1] L_LOST_EVIDENCE table created"

    Dim ev(8, 1) As String
    ev(0, 0) = "Negative socket / impression": ev(0, 1) = "Empty socket or impression left in the wall face"
    ev(1, 0) = "Beam hole":                    ev(1, 1) = "Through-hole for a spanning beam"
    ev(2, 0) = "Break scar":                   ev(2, 1) = "Detachment scar on masonry or bedrock"
    ev(3, 0) = "Detached fragment in situ":    ev(3, 1) = "Identifiable fallen fragment at the foot of the structure"
    ev(4, 0) = "Mortar imprint":                ev(4, 1) = "Mortar imprint without the element it once bonded"
    ev(5, 0) = "Corbels into void":              ev(5, 1) = "Corbels that no longer support anything"
    ev(6, 0) = "Pigment on bedrock":             ev(6, 1) = "Pigment on bedrock now bare of the masonry that carried it"
    ev(7, 0) = "Truncated walls":                ev(7, 1) = "Walls truncated on a clean plane"
    ev(8, 0) = "Other (see Notes)":              ev(8, 1) = "See the Notes field for detail"

    Dim i As Integer
    For i = 0 To 8
        sql = "INSERT INTO L_LOST_EVIDENCE (Name, Description) VALUES ("
        sql = sql & "'" & ev(i, 0) & "',"
        sql = sql & "'" & ev(i, 1) & "')"
        db.Execute sql, dbFailOnError
    Next i
    Debug.Print "[M1] L_LOST_EVIDENCE populated: 9 rows"
End Sub

' ----------------------------------------------------------------
'  M1c - L_SUPPORT expansion (section 5.7)
'  Appended at the end (new ID) - never renumber existing rows,
'  which are already referenced by ID_Support / ID_Support_Secondary.
' ----------------------------------------------------------------
Private Sub M1_ExpandLSupport(db As DAO.Database)
    If DCount("*", "L_SUPPORT", "Name='Micro-ledge (<50cm)'") > 0 Then
        Debug.Print "[M1] L_SUPPORT: Micro-ledge already present - skipped"
        Exit Sub
    End If
    db.Execute "INSERT INTO L_SUPPORT (Name) VALUES ('Micro-ledge (<50cm)')", dbFailOnError
    Debug.Print "[M1] L_SUPPORT: Micro-ledge (<50cm) added"
End Sub

' ----------------------------------------------------------------
'  M1d - EA09: reclassify MIX -> CAV (precondition for MIX removal,
'  section 6.3 and 11). Must run before the MIX DELETE below.
' ----------------------------------------------------------------
Private Sub M1_ReclassifyEA09_MixToCav(db As DAO.Database)
    Dim idMix As Variant
    Dim idCav As Variant
    idMix = DLookup("ID", "L_TYPOLOGY", "Name='MIX Mixed'")
    idCav = DLookup("ID", "L_TYPOLOGY", "Name='CAV Cave/Cavern'")

    If IsNull(idMix) Then
        Debug.Print "[M1] EA09 reclass: MIX typology not found (already migrated) - skipped"
        Exit Sub
    End If
    If IsNull(idCav) Then
        Debug.Print "[M1] EA09 reclass: ABORTED - CAV typology not found"
        Exit Sub
    End If

    db.Execute "UPDATE T_STRUCTURES SET ID_Typology=" & idCav & " WHERE Code='DW-S01-EA09' AND ID_Typology=" & idMix, dbFailOnError
    Debug.Print "[M1] EA09 reclass MIX->CAV: " & db.RecordsAffected & " record(s) affected"
End Sub

' ----------------------------------------------------------------
'  M1e - EA09 incoherence fix (section 11): Timber_Bracket_Role
'  Isolated -> Platform support; and EA10/EA21 (isolated corbels,
'  no real platform): clear Platform_Surface_Material.
'  NOTE: Sys_Platform / Platform_Function assignment for EA09 is
'  NOT done here - those fields do not exist until block M6, and the
'  complete/partial judgement is explicitly manual (step 9).
' ----------------------------------------------------------------
Private Sub M1_CorrectEA09_BracketAndPlatformMaterial(db As DAO.Database)
    db.Execute "UPDATE T_STRUCTURES SET Timber_Bracket_Role='Platform support' WHERE Code='DW-S01-EA09' AND Timber_Bracket_Role='Isolated'", dbFailOnError
    Debug.Print "[M1] EA09 Timber_Bracket_Role fix: " & db.RecordsAffected & " record(s) affected"

    db.Execute "UPDATE T_STRUCTURES SET Platform_Surface_Material=Null WHERE Code IN ('DW-S01-EA10','DW-S01-EA21') AND Platform_Surface_Material Is Not Null", dbFailOnError
    Debug.Print "[M1] EA10/EA21 Platform_Surface_Material cleared: " & db.RecordsAffected & " record(s) affected"
End Sub

' ----------------------------------------------------------------
'  M1f - L_TYPOLOGY rebuild (sections 6.1, 6.2, 6.3)
'  Order: add Record_Class column -> populate it for existing
'  categories -> split ND (UPDATE in place + INSERT Unclassifiable)
'  -> verify MIX unreferenced -> DELETE MIX (dbFailOnError direct,
'  never via a silent wrapper).
' ----------------------------------------------------------------
Private Sub M1_RebuildLTypology(db As DAO.Database)
    If Not FldExists(db, "L_TYPOLOGY", "Record_Class") Then
        db.Execute "ALTER TABLE L_TYPOLOGY ADD COLUMN Record_Class TEXT(30)", dbFailOnError
        Debug.Print "[M1] L_TYPOLOGY: Record_Class column added"
    Else
        Debug.Print "[M1] L_TYPOLOGY: Record_Class column already exists - skipped"
    End If

    db.Execute "UPDATE L_TYPOLOGY SET Record_Class='Built funerary structure' WHERE Name Like 'EA-MAU*' Or Name Like 'EA-CAM*' Or Name Like 'EA-PLA-R*' Or Name Like 'EA-PLA-V*'", dbFailOnError
    db.Execute "UPDATE L_TYPOLOGY SET Record_Class='Natural funerary context' WHERE Name Like 'NIX*' Or Name Like 'CAV*'", dbFailOnError
    db.Execute "UPDATE L_TYPOLOGY SET Record_Class='Structural trace' WHERE Name Like 'MEN*'", dbFailOnError
    db.Execute "UPDATE L_TYPOLOGY SET Record_Class='Rock art panel' WHERE Name Like 'PR*'", dbFailOnError
    Debug.Print "[M1] L_TYPOLOGY: Record_Class populated for MAU/CAM/PLA-R/PLA-V/NIX/CAV/MEN/PR"

    ' ND split: UPDATE in place (never DELETE+INSERT - REL_TYP_STR
    ' would block the DELETE while EA38/EA39/EA40b etc. reference ID 10)
    Dim idND As Variant
    idND = DLookup("ID", "L_TYPOLOGY", "Name='ND Undetermined'")
    If IsNull(idND) Then
        Debug.Print "[M1] L_TYPOLOGY: ND row already migrated - skipped"
    Else
        db.Execute "UPDATE L_TYPOLOGY SET Name='Not yet classified', Description='Working status: pending manual classification review.', Record_Class='Pending classification' WHERE ID=" & idND, dbFailOnError
        Debug.Print "[M1] L_TYPOLOGY: ND row (ID=" & idND & ") renamed to Not yet classified"
    End If

    If DCount("*", "L_TYPOLOGY", "Name='Unclassifiable'") > 0 Then
        Debug.Print "[M1] L_TYPOLOGY: Unclassifiable already present - skipped"
    Else
        db.Execute "INSERT INTO L_TYPOLOGY (Name, Description, Record_Class) VALUES ('Unclassifiable','Insufficient evidence to classify a BUILT structure (e.g. EA11: pigment perimeter with no surviving construction). Natural contexts are always classifiable as NIX or CAV.','Built funerary structure')", dbFailOnError
        Debug.Print "[M1] L_TYPOLOGY: Unclassifiable inserted"
    End If

    ' MIX removal - precondition: zero structures may still reference it
    Dim idMix As Variant
    idMix = DLookup("ID", "L_TYPOLOGY", "Name='MIX Mixed'")
    If IsNull(idMix) Then
        Debug.Print "[M1] L_TYPOLOGY: MIX already removed - skipped"
        Exit Sub
    End If

    Dim nMix As Long
    nMix = DCount("*", "T_STRUCTURES", "ID_Typology=" & idMix)
    Debug.Print "[M1] MIX precondition check: " & nMix & " structure(s) still reference MIX (must be 0)"

    If nMix <> 0 Then
        Debug.Print "[M1] ABORTED: MIX NOT deleted - reclassify the remaining structure(s) first"
    Else
        db.Execute "DELETE FROM L_TYPOLOGY WHERE ID=" & idMix, dbFailOnError
        Debug.Print "[M1] L_TYPOLOGY: MIX row deleted (ID=" & idMix & ")"
    End If
End Sub

' ================================================================
'  BLOCK M3 - Domain and defaults (SetByteDefaults v11)
'  DELTA sections: 1.1/1.5 (five-value domain, 19 fields here),
'  1.6 (three-value fields relabelled), 1.7 (mechanism).
'  DELTA section 13 step covered: 5.
'
'  WHAT THIS BLOCK DOES: sets Field.DefaultValue = "0" via DAO on
'  every observational BYTE field. v10 had default 9 (ND); v11 makes
'  0 (Absent) the default, because under the new domain a new record
'  must start from "examined, nothing there" and the researcher marks
'  9 explicitly when the position is not examinable.
'
'  WHAT THIS BLOCK DELIBERATELY DOES NOT DO (rule B): it does not
'  UPDATE a single stored value. The 35 existing records keep every
'  value they have, NULLs included. DefaultValue only ever applies to
'  rows inserted after this point. Converting the existing NULLs to 0
'  would assert hundreds of unverified absences (delta section 10).
'
'  NOTE ON FIELD NAMES: the four renames happen in M4, so this block
'  still uses the v10 names (Lateral_Wall_Faces, Lateral_Walls,
'  Recessed_Portal). DAO keeps DefaultValue across a later rename.
'  Lintel is absent here on purpose: it is still TEXT until M5.
' ================================================================
Private Sub M3_DomainAndDefaults(db As DAO.Database)
    Debug.Print "---- [M3] START ----"

    ' -- Group 1: the 19 five-value element fields (delta 1.5).
    '    Lintel is the 20th and joins the domain in M5 (delta 1.5 note).
    '    Names are v10 here; M4 renames three of them.
    Dim e(18) As String
    e(0) = "Embedded_Base_Beams"      ' A
    e(1) = "Base_Level"               ' B
    e(2) = "Decorative_Socle"         ' C
    e(3) = "Tie_Walls"                ' D
    e(4) = "Timber_Brackets"          ' E
    e(5) = "Transverse_Beams"         ' F
    e(6) = "Corbelled_Courses"        ' G
    e(7) = "Interbody_Cornice"        ' I
    e(8) = "Corner_Quoins"            ' J
    e(9) = "Structural_Pilasters"     ' K
    e(10) = "Lateral_Wall_Faces|Facade_Flank"   ' L, renamed in M4
    e(11) = "Relief_Frieze"           ' M
    e(12) = "Sill"                    ' N
    e(13) = "Jambs"                   ' O
    e(14) = "Upper_Crown"             ' R
    e(15) = "Eave_Beam"               ' S
    e(16) = "Eave_Surface"            ' T
    e(17) = "Lateral_Walls|Return_Wall"         ' V, renamed in M4
    e(18) = "Chamber_Roof"            ' X

    ' -- Group 2: the three-value 0/1/9 fields (delta 1.6). Same
    '    default 0; only the label of 9 changes (ND -> Not observable),
    '    which is a form-side concern handled in the form script.
    Dim o(23) As String
    o(0) = "Recessed_Portal|Recessed_Frame"     ' renamed in M4
    o(1) = "Looting"
    o(2) = "Fire_Damage"
    o(3) = "Animal_Activity"
    o(4) = "Modern_Access"
    o(5) = "Human_Remains"
    o(6) = "Anatomical_Connection"
    o(7) = "Mummification"
    o(8) = "Funerary_Bundles"
    o(9) = "Dispersed_Remains"
    o(10) = "Flexed_Position"
    o(11) = "Bone_Burning"
    o(12) = "Mat_Textiles"
    o(13) = "Mat_Wood"
    o(14) = "Mat_VegFiber"
    o(15) = "Mat_Ceramics"
    o(16) = "Mat_Fauna"
    o(17) = "Mat_DeerAntler"
    o(18) = "Mat_Other"
    o(19) = "Support_Modified"
    o(20) = "Mortar_Present"
    o(21) = "Chinking_Stones"
    o(22) = "Plaster_Present"
    o(23) = "Pigment_Present"

    Dim nOK As Integer
    Dim nFail As Integer
    Dim i As Integer

    For i = 0 To 18
        If M3_SetDefault(db, e(i), "0") Then nOK = nOK + 1 Else nFail = nFail + 1
    Next i
    Debug.Print "[M3] Five-value element fields: " & nOK & " default(s) set to 0 (expected 19)"

    Dim nOK2 As Integer
    nOK2 = 0
    For i = 0 To 23
        If M3_SetDefault(db, o(i), "0") Then nOK2 = nOK2 + 1 Else nFail = nFail + 1
    Next i
    Debug.Print "[M3] Three-value 0/1/9 fields: " & nOK2 & " default(s) set to 0 (expected 24)"

    Debug.Print "[M3] Total defaults set: " & (nOK + nOK2) & " (expected 43) | failures: " & nFail
    Debug.Print "[M3] Existing stored values and NULLs left untouched (rule B)"
    Debug.Print "[M3] Lintel excluded here on purpose - still TEXT until M5"
    Debug.Print "---- [M3] END ----"
End Sub

' Sets one DefaultValue and reports success, so a mistyped or missing
' field name surfaces in the log instead of being swallowed.
' fldSpec may list alternatives separated by "|": M3 runs before the
' M4 renames on a first pass but after them on every later one, so the
' three renamed fields have to be findable under either name.
Private Function M3_SetDefault(db As DAO.Database, fldSpec As String, val As String) As Boolean
    Dim cand() As String
    Dim fld As String
    Dim i As Integer
    cand = Split(fldSpec, "|")
    For i = 0 To UBound(cand)
        If FldExists(db, "T_STRUCTURES", cand(i)) Then
            fld = cand(i)
            Exit For
        End If
    Next i
    If Len(fld) = 0 Then
        Debug.Print "[M3] MISSING FIELD: " & fldSpec & " - default NOT set"
        Exit Function
    End If
    On Error GoTo Err_SD
    db.TableDefs("T_STRUCTURES").Fields(fld).DefaultValue = val
    M3_SetDefault = True
    Exit Function
Err_SD:
    Debug.Print "[M3] ERROR on " & fld & ": " & Err.Description
End Function

' ================================================================
'  BLOCK M4 - Field renames (DAO only)
'  DELTA sections: 5.1 (Return_Wall / Facade_Flank), 5.2
'  (Recessed_Frame), 5.5 (Rear_Closure_Type).
'  DELTA section 13 step covered: 6.
'
'  WHY DAO AND NOT SQL: JET DDL has no RENAME COLUMN. The only way
'  to rename a field while keeping its stored data is
'  TableDefs(t).Fields(old).Name = new. An ALTER TABLE ADD + UPDATE
'  + DROP round trip would work too, but it is three chances to lose
'  data where DAO offers zero.
'
'  WHY THE OLD NAMES WERE WRONG (5.1): both v10 names contained
'  "lateral" while denoting different things. V is the wall that
'  returns toward the cliff, perpendicular to the facade plane;
'  L is the wall face inside the facade plane. The data confirm they
'  are independent variables, so the names had to stop colliding.
'
'  SIDE EFFECT TO EXPECT: a DAO rename does NOT rewrite the SQL of
'  existing QueryDefs (Access Name AutoCorrect only tracks changes
'  made through the UI). QRY_05, QRY_13 and QRY_16 therefore refer to
'  fields that no longer exist until M9/M10 rebuild them. That is
'  expected and harmless as long as nobody opens them meanwhile.
' ================================================================
Private Sub M4_FieldRenames(db As DAO.Database)
    Debug.Print "---- [M4] START ----"

    Dim n As Integer
    n = 0
    If M4_RenameField(db, "T_STRUCTURES", "Lateral_Walls", "Return_Wall") Then n = n + 1
    If M4_RenameField(db, "T_STRUCTURES", "Lateral_Wall_Faces", "Facade_Flank") Then n = n + 1
    If M4_RenameField(db, "T_STRUCTURES", "Recessed_Portal", "Recessed_Frame") Then n = n + 1
    If M4_RenameField(db, "T_STRUCTURES", "Rear_Wall_Type", "Rear_Closure_Type") Then n = n + 1
    Debug.Print "[M4] T_STRUCTURES fields renamed this run: " & n & " (4 on a first run, 0 when re-run)"

    db.TableDefs.Refresh

    ' L_STRUCT_BODY: same element L, so the position entry follows the
    ' new name. Same ID (T_DECORATIONS rows keep pointing at it);
    ' only Code and Name change, per delta 5.1.
    If DCount("*", "L_STRUCT_BODY", "Code='LWF'") > 0 Then
        db.Execute "UPDATE L_STRUCT_BODY SET Code='FFL', Name='Facade flank (L)' WHERE Code='LWF'", dbFailOnError
        Debug.Print "[M4] L_STRUCT_BODY LWF -> FFL: " & db.RecordsAffected & " row(s) affected"
    Else
        Debug.Print "[M4] L_STRUCT_BODY: LWF entry already migrated - skipped"
    End If

    Debug.Print "[M4] NOTE: QRY_05 / QRY_13 / QRY_16 now reference dropped names."
    Debug.Print "[M4] They stay broken until M9/M10 rebuild them. Do not open them."
    Debug.Print "---- [M4] END ----"
End Sub

' Renames one field, idempotently: if the new name is already there the
' rename has been done, and if neither name is there something is wrong
' and it is reported rather than silently ignored.
Private Function M4_RenameField(db As DAO.Database, tbl As String, oldName As String, newName As String) As Boolean
    If FldExists(db, tbl, newName) Then
        Debug.Print "[M4] " & tbl & "." & newName & " already exists - skipped"
        Exit Function
    End If
    If Not FldExists(db, tbl, oldName) Then
        Debug.Print "[M4] ERROR: neither " & oldName & " nor " & newName & " found in " & tbl
        Exit Function
    End If
    On Error GoTo Err_RN
    db.TableDefs(tbl).Fields(oldName).Name = newName
    Debug.Print "[M4] " & tbl & ": " & oldName & " -> " & newName
    M4_RenameField = True
    Exit Function
Err_RN:
    Debug.Print "[M4] ERROR renaming " & oldName & " -> " & newName & ": " & Err.Description
End Function

' ================================================================
'  BLOCK M5 - Lintel: TEXT lookup -> BYTE element + material field
'  DELTA section 5.4. DELTA section 13 step covered: 7.
'
'  THE PROBLEM: v10 Lintel is a material lookup (Stone 10 / Absent 6 /
'  Wood 3 / ND 2) that conflates PRESENCE with TYPE. v11 splits it the
'  way Chamber_Roof / Chamber_Roof_Type already were:
'      Lintel           BYTE, five-value domain (element Q)
'      Lintel_Material  TEXT(20)  Stone / Wood / Mixed / ND
'
'  WHY A THREE-STEP DANCE: JET cannot convert TEXT to BYTE in place
'  while 'Stone' is sitting in the column. So: rename the old column
'  (DAO, data intact) -> add a brand-new BYTE column -> populate it
'  from the text -> keep both. Nothing is dropped: Lintel_Material is
'  a v11 field, not a casualty.
'
'  THE ONE DESTRUCTIVE EDIT IN THIS BLOCK: rows whose material reads
'  'Absent' or 'ND' get Lintel_Material cleared to NULL, because once
'  presence is 0 or 9 the "material" is not undetermined, it is
'  inapplicable (delta 5.4). No information is lost - 0 and 9 encode
'  those two states completely - but the text is gone for good, which
'  is why the counts are validated BEFORE the UPDATEs run and the
'  block aborts on any value the mapping does not cover.
'
'  NULLS: the ~14 rows with a NULL Lintel in v10 stay NULL. ALTER
'  TABLE ADD COLUMN does not backfill, and no UPDATE touches them
'  (rule B). They are resolved by the manual review of section 10.
' ================================================================
Private Sub M5_LintelSplit(db As DAO.Database)
    Debug.Print "---- [M5] START ----"

    ' Detected before anything moves: on a re-run both fields already
    ' exist, so the "before" counts describe the migrated state and
    ' must not be printed as if they were expectations.
    Dim alreadySplit As Boolean
    alreadySplit = FldExists(db, "T_STRUCTURES", "Lintel_Material") And FldExists(db, "T_STRUCTURES", "Lintel")

    ' -- Step 1: rename the v10 material column (DAO keeps the data).
    If FldExists(db, "T_STRUCTURES", "Lintel_Material") Then
        Debug.Print "[M5] Lintel_Material already exists - rename skipped"
    Else
        If Not FldExists(db, "T_STRUCTURES", "Lintel") Then
            Debug.Print "[M5] ABORTED: neither Lintel nor Lintel_Material found"
            Debug.Print "---- [M5] END ----"
            Exit Sub
        End If
        db.TableDefs("T_STRUCTURES").Fields("Lintel").Name = "Lintel_Material"
        db.TableDefs.Refresh
        Debug.Print "[M5] T_STRUCTURES: Lintel -> Lintel_Material (data preserved)"
    End If

    ' -- Step 2: measure BEFORE touching anything.
    Dim nStone As Long, nWood As Long, nMixed As Long
    Dim nAbsent As Long, nND As Long, nNull As Long
    Dim nTotal As Long, nOther As Long, nMapped As Long

    nStone = DCount("*", "T_STRUCTURES", "Lintel_Material='Stone'")
    nWood = DCount("*", "T_STRUCTURES", "Lintel_Material='Wood'")
    nMixed = DCount("*", "T_STRUCTURES", "Lintel_Material='Mixed'")
    nAbsent = DCount("*", "T_STRUCTURES", "Lintel_Material='Absent'")
    nND = DCount("*", "T_STRUCTURES", "Lintel_Material='ND'")
    nNull = DCount("*", "T_STRUCTURES", "Lintel_Material Is Null")
    nTotal = DCount("*", "T_STRUCTURES")
    nMapped = nStone + nWood + nMixed + nAbsent + nND
    nOther = nTotal - nMapped - nNull

    If alreadySplit Then
        Debug.Print "[M5] (already split - the counts below are the CURRENT state, not a baseline)"
        Debug.Print "[M5] CURRENT - Stone: " & nStone & " | Wood: " & nWood & " | Mixed: " & nMixed & " | NULL: " & nNull
    Else
        Debug.Print "[M5] BEFORE - Stone: " & nStone & " | Wood: " & nWood & " | Mixed: " & nMixed
        Debug.Print "[M5] BEFORE - Absent: " & nAbsent & " | ND: " & nND & " | NULL: " & nNull
        Debug.Print "[M5] BEFORE - mapped: " & nMapped & " (delta 5.4 expects 21) | unmapped: " & nOther
    End If

    ' Unmapped values would silently become NULL in the new BYTE field,
    ' which is real information loss. Stop instead.
    If nOther > 0 Then
        Debug.Print "[M5] ABORTED: " & nOther & " row(s) hold a Lintel_Material value"
        Debug.Print "[M5] outside Stone/Wood/Mixed/Absent/ND. Nothing was changed."
        Debug.Print "[M5] Inspect them and extend the mapping before re-running."
        Debug.Print "---- [M5] END ----"
        Exit Sub
    End If

    ' A drift from 21 is not dangerous (every value still maps), but it
    ' means the spec snapshot and the data have diverged - say so.
    ' Meaningless on a re-run, where Absent and ND are already cleared.
    If nMapped <> 21 And Not alreadySplit Then
        Debug.Print "[M5] WARNING: mapped count is " & nMapped & ", delta 5.4 documents 21."
        Debug.Print "[M5] All values are still covered by the mapping, so migration proceeds."
    End If

    ' -- Step 3: the new BYTE column. Born NULL on all 35 rows, which is
    '    exactly what delta 1.7 asks for; the default only serves new rows.
    If FldExists(db, "T_STRUCTURES", "Lintel") Then
        Debug.Print "[M5] Lintel BYTE column already exists - creation skipped"
    Else
        db.Execute "ALTER TABLE T_STRUCTURES ADD COLUMN Lintel BYTE", dbFailOnError
        db.TableDefs.Refresh
        Debug.Print "[M5] T_STRUCTURES: Lintel BYTE column created (all rows NULL)"
    End If

    ' Lintel is the 20th five-value element field; M3 could not set its
    ' default because it was still TEXT back then (delta 1.5 note).
    On Error Resume Next
    db.TableDefs("T_STRUCTURES").Fields("Lintel").DefaultValue = "0"
    On Error GoTo 0
    Debug.Print "[M5] Lintel DefaultValue set to 0 (new records only)"

    ' -- Step 4: populate, per the mapping table of delta 5.4.
    '    Mixed is not in that table (the corpus has none) but is mapped
    '    to 1 defensively: it is a present lintel of mixed material.
    db.Execute "UPDATE T_STRUCTURES SET Lintel=1 WHERE Lintel_Material IN ('Stone','Wood','Mixed')", dbFailOnError
    Debug.Print "[M5] Stone/Wood/Mixed -> Lintel=1 : " & db.RecordsAffected & " row(s)"

    db.Execute "UPDATE T_STRUCTURES SET Lintel=0, Lintel_Material=Null WHERE Lintel_Material='Absent'", dbFailOnError
    Debug.Print "[M5] Absent -> Lintel=0, material cleared : " & db.RecordsAffected & " row(s)"

    db.Execute "UPDATE T_STRUCTURES SET Lintel=9, Lintel_Material=Null WHERE Lintel_Material='ND'", dbFailOnError
    Debug.Print "[M5] ND -> Lintel=9, material cleared : " & db.RecordsAffected & " row(s)"

    ' -- Step 5: verify the transformation against the before-counts.
    Dim aOne As Long, aZero As Long, aNine As Long, aNull As Long, aMat As Long
    aOne = DCount("*", "T_STRUCTURES", "Lintel=1")
    aZero = DCount("*", "T_STRUCTURES", "Lintel=0")
    aNine = DCount("*", "T_STRUCTURES", "Lintel=9")
    aNull = DCount("*", "T_STRUCTURES", "Lintel Is Null")
    aMat = DCount("*", "T_STRUCTURES", "Lintel_Material Is Not Null")

    If alreadySplit Then
        Debug.Print "[M5] CURRENT - Lintel 1/0/9/NULL: " & aOne & "/" & aZero & "/" & aNine & "/" & aNull & " | material set on " & aMat
        Debug.Print "[M5] (re-run: nothing to migrate, " & (aOne + aZero + aNine) & " value(s) already in place)"
    Else
        Debug.Print "[M5] AFTER - Lintel=1: " & aOne & " (expected " & (nStone + nWood + nMixed) & ")"
        Debug.Print "[M5] AFTER - Lintel=0: " & aZero & " (expected " & nAbsent & ")"
        Debug.Print "[M5] AFTER - Lintel=9: " & aNine & " (expected " & nND & ")"
        Debug.Print "[M5] AFTER - Lintel NULL: " & aNull & " (expected " & nNull & ", rule B)"
        Debug.Print "[M5] AFTER - Lintel_Material not null: " & aMat & " (expected " & (nStone + nWood + nMixed) & ")"
        If aOne = nStone + nWood + nMixed And aZero = nAbsent And aNine = nND And aNull = nNull Then
            Debug.Print "[M5] MIGRATION VERIFIED: " & (aOne + aZero + aNine) & " value(s) migrated, " & aNull & " NULL(s) preserved"
        Else
            Debug.Print "[M5] *** MISMATCH - review the numbers above before continuing ***"
        End If
    End If

    Debug.Print "---- [M5] END ----"
End Sub

' ================================================================
'  BLOCK M6 - Constructive systems (Sys_*) + Platform_Function
'  DELTA sections: 4.1 (the three element-systems and the v10 value
'  mapping), 4.2 (the two grouping fields), 5.3 (Platform_Function),
'  1.7 (types and DAO defaults).
'  DELTA section 13 step covered: 8.
'
'  WHAT CHANGES CONCEPTUALLY: H, P and U were never really elements,
'  they were systems (H = E+F+G, P = N+O+Q, U = S+T). In v11 they stop
'  having a field of their own and BECOME the Sys_* fields, with a
'  six-value domain that keeps the conservation gradient instead of
'  flattening it into 0/1/9.
'
'  THE ONLY SCRIPTED PART (delta 4.1): the v10 values 0 and 9 are
'  already-made observations and translate deterministically.
'      0     -> 'Absent'
'      9     -> 'Not observable'
'      1     -> NULL   (needs the complete/partial/lost judgement)
'      NULL  -> NULL   (never converted, rule B)
'  Sys_Base and Sys_Chamber have no v10 counterpart at all (Rear_Wall
'  is not a substitute: Sys_Chamber aggregates six elements), so they
'  are created empty and filled entirely by hand.
'
'  RE-RUN SAFETY - IMPORTANT: every UPDATE carries "AND Sys_X Is Null".
'  Step 9 of the delta is a long manual assignment session, and this
'  script is re-run after it to deliver later blocks. Without that
'  guard a re-run would silently revert hand-made values back to the
'  scripted ones. With it, M6 only ever fills blanks.
'
'  NO DELETIONS HERE: Corbelled_Platform, Access_Opening and Eave stay
'  in the table, untouched, until M11b - and only once you confirm the
'  Sys_* review is finished (rule A).
' ================================================================
Private Sub M6_SysFields(db As DAO.Database)
    Debug.Print "---- [M6] START ----"

    ' -- Step 1: the six new fields. All 35 rows are born NULL;
    '    the DAO default only ever applies to records created later.
    Dim nNew As Integer
    nNew = 0
    If M6_AddTextField(db, "Sys_Platform", 20, "Absent") Then nNew = nNew + 1
    If M6_AddTextField(db, "Sys_Portal", 20, "Absent") Then nNew = nNew + 1
    If M6_AddTextField(db, "Sys_Eave", 20, "Absent") Then nNew = nNew + 1
    If M6_AddTextField(db, "Sys_Base", 20, "Absent") Then nNew = nNew + 1
    If M6_AddTextField(db, "Sys_Chamber", 20, "Absent") Then nNew = nNew + 1
    If M6_AddTextField(db, "Platform_Function", 20, "Undetermined") Then nNew = nNew + 1
    db.TableDefs.Refresh
    Debug.Print "[M6] New fields created this run: " & nNew & " (6 on a first run, 0 when re-run)"

    ' -- Step 2: migrate the recorded 0/9 values of H, P and U.
    M6_MigrateSystem db, "Corbelled_Platform", "Sys_Platform", "H"
    M6_MigrateSystem db, "Access_Opening", "Sys_Portal", "P"
    M6_MigrateSystem db, "Eave", "Sys_Eave", "U"

    ' -- Step 3: report what manual assignment is now unblocked.
    Debug.Print "[M6] --- pending manual assignment (delta step 9) ---"
    Debug.Print "[M6] Sys_Platform still NULL: " & DCount("*", "T_STRUCTURES", "Sys_Platform Is Null")
    Debug.Print "[M6] Sys_Portal still NULL: " & DCount("*", "T_STRUCTURES", "Sys_Portal Is Null")
    Debug.Print "[M6] Sys_Eave still NULL: " & DCount("*", "T_STRUCTURES", "Sys_Eave Is Null")
    Debug.Print "[M6] Sys_Base still NULL: " & DCount("*", "T_STRUCTURES", "Sys_Base Is Null") & " (no v10 source, fully manual)"
    Debug.Print "[M6] Sys_Chamber still NULL: " & DCount("*", "T_STRUCTURES", "Sys_Chamber Is Null") & " (no v10 source, fully manual)"
    Debug.Print "[M6] Platform_Function still NULL: " & DCount("*", "T_STRUCTURES", "Platform_Function Is Null") & " (manual)"
    Debug.Print "[M6] Old H/P/U fields kept intact until M11b (rule A)"
    Debug.Print "---- [M6] END ----"
End Sub

' Adds one TEXT field with a DAO default. Returns True only when it
' actually created it, so a re-run reports 0 instead of pretending.
Private Function M6_AddTextField(db As DAO.Database, fld As String, sz As Integer, defVal As String) As Boolean
    If FldExists(db, "T_STRUCTURES", fld) Then
        Debug.Print "[M6] " & fld & " already exists - skipped"
        Exit Function
    End If
    On Error GoTo Err_AF
    db.Execute "ALTER TABLE T_STRUCTURES ADD COLUMN " & fld & " TEXT(" & sz & ")", dbFailOnError
    db.TableDefs.Refresh
    db.TableDefs("T_STRUCTURES").Fields(fld).DefaultValue = Chr(34) & defVal & Chr(34)
    Debug.Print "[M6] " & fld & " TEXT(" & sz & ") created, default " & defVal & " (new records only)"
    M6_AddTextField = True
    Exit Function
Err_AF:
    Debug.Print "[M6] ERROR creating " & fld & ": " & Err.Description
End Function

' Translates the 0 and 9 of one v10 element field into its v11 system
' field, leaving 1 and NULL for the researcher. Only fills blanks.
Private Sub M6_MigrateSystem(db As DAO.Database, srcFld As String, sysFld As String, letter As String)
    If Not FldExists(db, "T_STRUCTURES", srcFld) Then
        Debug.Print "[M6] " & letter & ": source " & srcFld & " not found - already dropped? skipped"
        Exit Sub
    End If
    If Not FldExists(db, "T_STRUCTURES", sysFld) Then
        Debug.Print "[M6] " & letter & ": target " & sysFld & " missing - skipped"
        Exit Sub
    End If

    Dim c0 As Long, c9 As Long, c1 As Long, cN As Long
    c0 = DCount("*", "T_STRUCTURES", srcFld & "=0")
    c9 = DCount("*", "T_STRUCTURES", srcFld & "=9")
    c1 = DCount("*", "T_STRUCTURES", srcFld & "=1")
    cN = DCount("*", "T_STRUCTURES", srcFld & " Is Null")
    Debug.Print "[M6] " & letter & " (" & srcFld & ") source: 0=" & c0 & " 9=" & c9 & " 1=" & c1 & " NULL=" & cN

    db.Execute "UPDATE T_STRUCTURES SET " & sysFld & "='Absent' WHERE " & srcFld & "=0 AND " & sysFld & " Is Null", dbFailOnError
    Dim a0 As Long
    a0 = db.RecordsAffected
    db.Execute "UPDATE T_STRUCTURES SET " & sysFld & "='Not observable' WHERE " & srcFld & "=9 AND " & sysFld & " Is Null", dbFailOnError
    Dim a9 As Long
    a9 = db.RecordsAffected

    Debug.Print "[M6] " & letter & " -> " & sysFld & ": " & a0 & " set Absent, " & a9 & " set Not observable (" & (a0 + a9) & " total)"
    Debug.Print "[M6] " & letter & ": " & c1 & " value(s) 1 and " & cN & " NULL(s) left NULL for manual review"
End Sub

' ================================================================
'  BLOCK M7 - T_LOST_ELEMENTS + Chrono_Relation
'  DELTA sections: 2 (the table, its lookup and the scope rule),
'  9.1 (Chrono_Relation), 9.1bis (closed Connection_Type list).
'  DELTA section 13 steps covered: 10 and 11 (structure only).
'
'  WHY THE TABLE EXISTS: value 3 (Attested lost) is the only one in
'  the domain that is an inference rather than an observation. Without
'  a record of the physical evidence behind it, it is not replicable -
'  and it is the evidential basis of OE3, the reconstruction of the
'  lost aerial circulation infrastructure.
'
'  ELEMENT_CODE IS OPTIONAL, ON PURPOSE (delta 2, REV4): rows whose
'  scope is Body or Whole structure have no single element code - a
'  vanished body or a razed structure is not one element. Making the
'  column NOT NULL, as rev. 3 did, would have rejected exactly the
'  rows the migration of Lost_Body_Evidence and EA11 need to create.
'  The "mandatory if and only if scope = Element" rule lives in
'  QRY_16 rule 21 and in form validation, not in the DDL.
'
'  AllowZeroLength is forced off on Element_Code so that "no code"
'  can only ever be NULL. An empty string would satisfy neither the
'  foreign key nor rule 21, and would be invisible in the datasheet.
'
'  NO DATA IS INSERTED HERE. Migrating Lost_Body_Evidence (up to 25
'  values), the EA11 whole-structure row and the five a/b connection
'  pairs are all manual tasks per delta section 15 - the script only
'  builds the place to put them.
' ================================================================
Private Sub M7_LostElementsAndConnections(db As DAO.Database)
    Debug.Print "---- [M7] START ----"

    ' -- Step 1: the table.
    If TblExists(db, "T_LOST_ELEMENTS") Then
        Debug.Print "[M7] T_LOST_ELEMENTS already exists - creation skipped"
    Else
        Dim sql As String
        sql = "CREATE TABLE T_LOST_ELEMENTS ("
        sql = sql & "ID COUNTER CONSTRAINT PK_LOST PRIMARY KEY,"
        sql = sql & "ID_Structure LONG NOT NULL,"
        sql = sql & "Element_Code TEXT(2),"
        sql = sql & "ID_Evidence_Type LONG NOT NULL,"
        sql = sql & "Evidence_Scope TEXT(20),"
        sql = sql & "ID_Position LONG,"
        sql = sql & "Notes MEMO)"
        db.Execute sql, dbFailOnError
        db.TableDefs.Refresh
        Debug.Print "[M7] T_LOST_ELEMENTS created (7 fields)"

        On Error Resume Next
        db.TableDefs("T_LOST_ELEMENTS").Fields("Element_Code").AllowZeroLength = False
        On Error GoTo 0
        Debug.Print "[M7] Element_Code AllowZeroLength = False (blank means NULL, never '')"
    End If

    ' -- Step 2: the four relationships. Element_Code -> L_ELEMENTS.Code
    '    is the one that needs the UQ_ELEM_CODE unique index built in M1:
    '    JET only accepts a relationship against a primary key or a
    '    unique index, never against a plain column.
    Dim n As Integer
    n = 0
    If M7_MkRel(db, "REL_STR_LOST", "T_STRUCTURES", "ID", "T_LOST_ELEMENTS", "ID_Structure", True) Then n = n + 1
    If M7_MkRel(db, "REL_ELEM_LOST", "L_ELEMENTS", "Code", "T_LOST_ELEMENTS", "Element_Code", False) Then n = n + 1
    If M7_MkRel(db, "REL_LEV_LOST", "L_LOST_EVIDENCE", "ID", "T_LOST_ELEMENTS", "ID_Evidence_Type", False) Then n = n + 1
    If M7_MkRel(db, "REL_SB_LOST", "L_STRUCT_BODY", "ID", "T_LOST_ELEMENTS", "ID_Position", False) Then n = n + 1
    db.Relations.Refresh
    Debug.Print "[M7] Relationships created this run: " & n & " (4 on a first run, 0 when re-run)"

    ' -- Step 3: Chrono_Relation on T_CONNECTIONS (delta 9.1).
    '    Direction is read from the joint: the structure showing the
    '    untoothed joint against the other's wall face is the later one.
    If Not TblExists(db, "T_CONNECTIONS") Then
        Debug.Print "[M7] T_CONNECTIONS not found - Chrono_Relation skipped"
    ElseIf FldExists(db, "T_CONNECTIONS", "Chrono_Relation") Then
        Debug.Print "[M7] T_CONNECTIONS.Chrono_Relation already exists - skipped"
    Else
        db.Execute "ALTER TABLE T_CONNECTIONS ADD COLUMN Chrono_Relation TEXT(20)", dbFailOnError
        db.TableDefs.Refresh
        On Error Resume Next
        db.TableDefs("T_CONNECTIONS").Fields("Chrono_Relation").DefaultValue = Chr(34) & "Undetermined" & Chr(34)
        On Error GoTo 0
        Debug.Print "[M7] T_CONNECTIONS.Chrono_Relation TEXT(20) added, default Undetermined"
    End If

    Debug.Print "[M7] --- awaiting manual entry (delta section 15) ---"
    Debug.Print "[M7] T_LOST_ELEMENTS rows: " & DCount("*", "T_LOST_ELEMENTS") & " (Lost_Body_Evidence migration + EA11)"
    Debug.Print "[M7] T_CONNECTIONS rows: " & DCount("*", "T_CONNECTIONS") & " (the five a/b pairs)"
    Debug.Print "[M7] Lost_Body_Evidence values still to migrate: " & DCount("*", "T_STRUCTURES", "Lost_Body_Evidence Is Not Null And Lost_Body_Evidence<>'None' And Lost_Body_Evidence<>'ND'")
    Debug.Print "---- [M7] END ----"
End Sub

' Creates one relationship, reporting instead of failing silently.
' delCascade mirrors v10's MkRel: child rows follow the parent.
Private Function M7_MkRel(db As DAO.Database, nm As String, pT As String, pF As String, cT As String, cF As String, delCascade As Boolean) As Boolean
    Dim r As DAO.Relation
    For Each r In db.Relations
        If r.Name = nm Then
            Debug.Print "[M7] " & nm & " already exists - skipped"
            Exit Function
        End If
    Next r

    Dim rel As DAO.Relation
    Dim fld As DAO.Field
    Dim fl As Long
    On Error GoTo Err_R
    fl = dbRelationUpdateCascade
    If delCascade Then fl = fl Or dbRelationDeleteCascade
    Set rel = db.CreateRelation(nm, pT, cT, fl)
    Set fld = rel.CreateField(pF)
    fld.ForeignName = cF
    rel.Fields.Append fld
    db.Relations.Append rel
    Debug.Print "[M7] " & nm & ": " & pT & "." & pF & " -> " & cT & "." & cF
    M7_MkRel = True
    Exit Function
Err_R:
    Debug.Print "[M7] ERROR creating " & nm & ": " & Err.Description
End Function

' ================================================================
'  BLOCK M9a - QRY_02 (rebuilt), QRY_05, QRY_07
'  DELTA sections: 7.6 (QRY_02 reconstruction), 12bis (inventory of
'  affected queries). DELTA section 13 step covered: 13 (part).
'
'  QRY_02 IS THE DELICATE ONE. Its v10 version was built entirely on
'  the Dec_* booleans that section 7.2 deletes, and it is the source
'  of the LP vs DW chi-squared. T_DECORATIONS records only PRESENCES,
'  so the Present / Verified absent / Not evaluable triad that lets
'  the denominator be chosen has to be reconstructed:
'      Present        a row of that motif exists
'      Verified absent no row AND Facade_Observability = 'Complete'
'      Not evaluable   no row AND observability is anything else
'  Valid denominator = Present + Verified absent. Never the total.
'
'  WHY THREE QUERY OBJECTS. The delta sketches "LEFT JOIN +
'  SUM(IIF())", but a structure can carry the same motif in two
'  positions (a T-niche on the socle and another over the lintel are
'  two legitimate T_DECORATIONS rows). A direct LEFT JOIN duplicates
'  the structure row and SUM(IIF()) would then count that structure
'  twice, silently inflating the very statistic this query exists to
'  produce. So:
'      QRY_02s_Decoration_Typed   decorations with the type name
'      QRY_02a_Decoration_Flags   one row per structure, 0/1 per motif
'      QRY_02_Decoration_by_Site  aggregates those flags by site
'  The GROUP BY in the flag query collapses the duplicates before any
'  counting happens - MAX(IIF(...)) reads 1 whether there is one
'  matching row or five.
'
'  WHY THE EXTRA "s" QUERY (JET error 3296). The obvious one-query
'  form would be
'      LEFT JOIN (T_DECORATIONS AS D INNER JOIN L_DEC_TYPE AS DT ...)
'  but JET refuses a LEFT JOIN whose right operand is a parenthesised
'  INNER JOIN, and raises "JOIN expression not supported": the inner
'  join could drop rows, which makes the outer join ambiguous. Access
'  does accept a LEFT JOIN against a SAVED QUERY, so the type lookup
'  is resolved once in QRY_02s and joined from there. This keeps the
'  motif matching by NAME - hard-coding L_DEC_TYPE autonumber IDs
'  into query SQL is exactly what delta section 12 warns against.
'
'  ROCK ART: the Rock_Art / RA_* booleans are gone (7.2). Their
'  replacement is Pigment_Present = 1 with Pigment_Substrate =
'  'Bedrock', which is how rock art was always identifiable anyway.
'
'  RECORD_CLASS: mandatory filter per 6.1. Pending classification is a
'  working status, not a result, so it is excluded from the analysis.
' ================================================================
Private Sub M9a_QueriesDecorationAndExports(db As DAO.Database)
    Debug.Print "---- [M9a] START ----"

    ' Drop dependants before their sources.
    M9_DropQuery db, "QRY_02_Decoration_by_Site"
    M9_DropQuery db, "QRY_02a_Decoration_Flags"
    M9_DropQuery db, "QRY_02s_Decoration_Typed"
    M9_DropQuery db, "QRY_05_Export_RStats"
    M9_DropQuery db, "QRY_07_Export_QGIS"

    M9a_BuildDecorationTyped db
    M9a_BuildDecorationFlags db
    M9a_BuildDecorationBySite db
    M9a_BuildExportRStats db
    M9a_BuildExportQGIS db

    db.QueryDefs.Refresh
    Debug.Print "---- [M9a] END ----"
End Sub

Private Sub M9_DropQuery(db As DAO.Database, qn As String)
    Dim q As DAO.QueryDef
    For Each q In db.QueryDefs
        If q.Name = qn Then
            db.QueryDefs.Delete qn
            Exit Sub
        End If
    Next q
End Sub

' Creates one query, containing any failure. Without this a single bad
' statement aborts the whole block and leaves every later query
' dropped-but-not-recreated, which is how one error reads as four.
Private Sub M9_CreateQuery(db As DAO.Database, qn As String, sql As String, note As String)
    On Error GoTo Err_CQ
    db.CreateQueryDef qn, sql
    Debug.Print "[M9] " & qn & " created (" & note & ")"
    Exit Sub
Err_CQ:
    Debug.Print "[M9] *** FAILED to create " & qn & " *** " & Err.Description
End Sub

' Decorations with their type name resolved. Exists so that the flag
' query can LEFT JOIN a saved query instead of a parenthesised inner
' join, which JET rejects with error 3296.
Private Sub M9a_BuildDecorationTyped(db As DAO.Database)
    Dim q As String
    q = "SELECT D.ID_Structure, DT.Name AS Dec_Name "
    q = q & "FROM T_DECORATIONS AS D "
    q = q & "INNER JOIN L_DEC_TYPE AS DT ON D.ID_Dec_Type=DT.ID;"
    M9_CreateQuery db, "QRY_02s_Decoration_Typed", q, "type names resolved"
End Sub

' One row per structure, with a 0/1 flag per motif. The GROUP BY is
' what makes repeated motifs harmless.
Private Sub M9a_BuildDecorationFlags(db As DAO.Database)
    Dim q As String
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
    M9_CreateQuery db, "QRY_02a_Decoration_Flags", q, "one row per structure"
End Sub

' Three counters per motif, so the chi-squared denominator can be
' chosen explicitly. Built by loop: 9 motifs x 3 counters.
Private Sub M9a_BuildDecorationBySite(db As DAO.Database)
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

    Dim q As String
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
    M9_CreateQuery db, "QRY_02_Decoration_by_Site", q, "9 motifs x Present/Absent/ND"
End Sub

' Flat export for R. Dropped: Lost_Body_Evidence, the 8 Dec_*/Rock_Art
' booleans, and the four absorbed H/P/U/W fields. Added: the five
' Sys_*, Platform_Function, Lintel + Lintel_Material, and Record_Class.
' JOINs to L_TYPOLOGY and L_SUPPORT are LEFT, not INNER as in v10: an
' export that silently drops records with no typology or no support
' recorded is a bug, not a filter.
Private Sub M9a_BuildExportRStats(db As DAO.Database)
    Dim q As String
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
    M9_CreateQuery db, "QRY_05_Export_RStats", q, "v11 field names, Sys_* added"
End Sub

' Spatial export. Lost_Body_Evidence is replaced by a per-structure
' count from T_LOST_ELEMENTS, as delta 12bis suggests.
Private Sub M9a_BuildExportQGIS(db As DAO.Database)
    Dim q As String
    q = "SELECT E.ID, E.Code, S.Site_Name, SC.Sector_Name, "
    q = q & "T.Name AS Typology, T.Record_Class, "
    q = q & "E.Coord_Lat_WGS84, E.Coord_Lon_WGS84, "
    q = q & "E.Coord_E_UTM, E.Coord_N_UTM, E.Altitude_masl, "
    q = q & "E.Height_Above_Base_m, E.Coord_Precision_m, "
    q = q & "AS1.Name AS Arch_Status, MS.Name AS Material_Status, "
    q = q & "E.N_Basal_Bodies, E.N_Chamber_Bodies, "
    q = q & "(SELECT COUNT(*) FROM T_LOST_ELEMENTS AS LE WHERE LE.ID_Structure=E.ID) AS N_Lost_Elements, "
    q = q & "E.Sys_Platform, E.Sys_Portal, E.Sys_Eave, "
    q = q & "E.Interior_Area_m2, E.Interior_Vol_m3, "
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
    M9_CreateQuery db, "QRY_07_Export_QGIS", q, "lost-element count replaces Lost_Body_Evidence"
End Sub

' ================================================================
'  BLOCK M9b - QRY_13 (A-X matrix) and QRY_14 (connection edges)
'  DELTA sections: 12bis (QRY_13 v11 specification), 4.3 and 4.7
'  (analytical semantics of NA), 4.5 (which field belongs to which
'  system), 9.3 (QRY_14). DELTA section 13 step covered: 13 (part).
'
'  THE MATRIX SHRINKS FROM 24 TO 20 ELEMENT COLUMNS. H, P and U are
'  no longer elements but systems, and get SYS_ columns instead;
'  W disappears entirely (5.5), its information carried by
'  Rear_Closure_Type. AX_Q now reads the Lintel BYTE directly - the
'  v10 IIF over material text is gone with the split of M5.
'
'  THE NA RULE IS THE POINT OF THIS QUERY (4.3, 4.7). A system marked
'  Not applicable or Not observable must export NULL - which R reads
'  as NA - for the system column AND for every one of its component
'  fields, because those components hold the technical 0 of section
'  1.2, a padding value that asserts nothing. Exporting that 0 raw
'  would put unverified absences into the denominator, and the
'  published percentages would be wrong.
'      Absent, by contrast, exports its real values: a mausoleum
'      with no eave is a result, not a gap.
'  Elements I and R belong to no system (4.5) and are never gated.
'
'  Gating is applied to the type/material companions too, not just to
'  the A-X letters: Lintel_Material under Sys_Portal, Chamber_Roof_Type
'  and Rear_Closure_Type under Sys_Chamber, and so on per table 4.5.
' ================================================================
Private Sub M9b_QueriesAXMatrix(db As DAO.Database)
    Debug.Print "---- [M9b] START ----"

    M9_DropQuery db, "QRY_13_AX_Pattern_Export"
    M9_DropQuery db, "QRY_14_Connections_Edges"

    M9b_BuildAXExport db
    M9b_BuildConnectionEdges db

    db.QueryDefs.Refresh
    Debug.Print "---- [M9b] END ----"
End Sub

' Builds one SELECT column, wrapped in the NA gate when the field
' belongs to a system. gate = "" means the field is never gated.
Private Function M9_Gated(fld As String, gate As String, outName As String) As String
    If Len(gate) = 0 Then
        M9_Gated = "E." & fld & " AS " & outName & ", "
    Else
        M9_Gated = "IIF(E." & gate & "='Not applicable' Or E." & gate & "='Not observable',Null,E." & fld & ") AS " & outName & ", "
    End If
End Function

Private Sub M9b_BuildAXExport(db As DAO.Database)
    Dim ax() As String
    FillElementMap ax

    ' Type and material companions, gated by the same table 4.5.
    ' Interbody_Cornice_Material is ungated: element I belongs to no
    ' system, so its material is never padding (delta 4.5).
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

    ' The five system columns, which are themselves NA-gated.
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
        q = q & M9_Gated(ax(i, 1), ax(i, 2), "AX_" & ax(i, 0))
    Next i

    For i = 0 To 4
        q = q & M9_Gated(sy(i, 0), sy(i, 0), sy(i, 1))
    Next i

    For i = 0 To 8
        q = q & M9_Gated(cm(i, 0), cm(i, 1), cm(i, 0))
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

    M9_CreateQuery db, "QRY_13_AX_Pattern_Export", q, "20 AX + 5 SYS columns, NA semantics"
End Sub

' Edge list for igraph / QGIS. Chrono_Relation turns what was an
' undirected graph into a potentially directed one, which is what
' H03 needs (delta 9.3).
Private Sub M9b_BuildConnectionEdges(db As DAO.Database)
    Dim q As String
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
    M9_CreateQuery db, "QRY_14_Connections_Edges", q, "Chrono_Relation added, directed graph"
End Sub

' ================================================================
'  BLOCK M10a - QRY_16 validation battery, rules 1-11
'  DELTA section 12 (the 21 rules and the REV5 definition of the two
'  kinds of zero). DELTA section 13 step covered: 13 (part).
'
'  NON-BLOCKING BY DESIGN, as in v10: a hard table constraint would
'  stop the researcher recording a genuinely observed absence in a
'  partly collapsed structure. An empty result set means the corpus
'  is coherent; run it periodically during data entry.
'
'  THE HELPER QUERY IS THE WHOLE TRICK. Rules 1, 2, 4 and 6 each have
'  to iterate over all 20 element fields. Written directly that would
'  be roughly 80 UNION branches of near-identical SQL. Instead
'  QRY_16s_Element_Values pivots the 20 columns into long format once
'      ID_Structure | Code | Element_Code | Field_Name |
'      Element_Value | Sys_Field | Sys_Value
'  and every one of those rules becomes a single short branch that
'  also names the offending field in its message.
'
'  THE TWO KINDS OF ZERO (REV5). Rules 4 and 6 demanded opposite
'  things until rev. 5 separated them:
'      0 of assertion - the field's system IS Present*/Attested lost,
'                       or the field has no system (I and R). It means
'                       "the position was examined and nothing was
'                       there", so on a Collapsed structure it is
'                       unverifiable and rule 6 flags it.
'      0 of padding   - the system is Not applicable / Not observable /
'                       Absent. It asserts nothing, it only avoids a
'                       NULL, and rule 4 actively requires it.
'  Rule 6 therefore excludes padding zeros. Without that exclusion a
'  collapsed cave with Sys_Portal = Not applicable would fire rule 6
'  on the very Sill/Jambs/Lintel zeros that rule 4 demands.
'
'  RULE 6 STAYS QUIET FOR NOW, CORRECTLY. While the Sys_* fields are
'  still unassigned they are NULL, so no zero counts as an assertion
'  and the rule reports nothing. It becomes informative as you work
'  through the manual assignment - which is the right behaviour: the
'  rule cannot say anything useful before it knows which systems exist.
' ================================================================
Private Sub M10a_QRY16_Rules1to11(db As DAO.Database)
    Debug.Print "---- [M10a] START ----"

    M9_DropQuery db, "QRY_16_Validation_Check"
    M9_DropQuery db, "QRY_16s_Element_Values"
    M9_DropQuery db, "QRY_16s_Lost_Cover"
    M9_DropQuery db, "QRY_16s_Damage_Count"

    M10_BuildElementValues db
    M10_BuildLostCover db
    M10_BuildDamageCount db
    M10_BuildValidationCheck db, False
    M10_DiagnoseRules db

    db.QueryDefs.Refresh
    Debug.Print "---- [M10a] END ----"
End Sub

' Every (structure, element) pair that has justifying evidence. A row
' scoped Body or Whole structure covers the whole vocabulary of that
' structure at once (delta section 2), which is what the cross join
' with L_ELEMENTS expresses.
' Exists so rule 2 can be a LEFT JOIN instead of a correlated NOT
' EXISTS: JET resolves outer references unreliably inside UNION
' branches, which is what "Too few parameters" means when it happens.
Private Sub M10_BuildLostCover(db As DAO.Database)
    Dim q As String
    q = "SELECT L.ID_Structure, L.Element_Code "
    q = q & "FROM T_LOST_ELEMENTS AS L "
    q = q & "WHERE L.Evidence_Scope='Element' AND L.Element_Code Is Not Null "
    q = q & "UNION "
    q = q & "SELECT L.ID_Structure, M.Code "
    q = q & "FROM T_LOST_ELEMENTS AS L, L_ELEMENTS AS M "
    q = q & "WHERE L.Evidence_Scope='Body' OR L.Evidence_Scope='Whole structure';"
    M9_CreateQuery db, "QRY_16s_Lost_Cover", q, "structure/element pairs with evidence"
End Sub

' Partial-or-lost element count per structure, for rule 5. Same
' reason: a plain join beats a correlated scalar subquery here.
Private Sub M10_BuildDamageCount(db As DAO.Database)
    Dim q As String
    q = "SELECT V.ID_Structure, "
    q = q & "SUM(IIF(V.Element_Value=2 Or V.Element_Value=3,1,0)) AS N_Damaged "
    q = q & "FROM QRY_16s_Element_Values AS V "
    q = q & "GROUP BY V.ID_Structure;"
    M9_CreateQuery db, "QRY_16s_Damage_Count", q, "partial/lost count per structure"
End Sub

' Builds each rule as a standalone throwaway query and opens it, so a
' failure names the rule instead of just breaking the whole battery.
Private Sub M10_DiagnoseRules(db As DAO.Database)
    Debug.Print "[M10] --- per-rule diagnosis ---"
    M10_TestRule db, "R1", M10_Rule01()
    M10_TestRule db, "R2a", M10_Rule02_Elements()
    M10_TestRule db, "R2b", M10_Rule02_Systems()
    M10_TestRule db, "R3", M10_Rule03()
    M10_TestRule db, "R4", M10_Rule04()
    M10_TestRule db, "R5", M10_Rule05()
    M10_TestRule db, "R6", M10_Rule06()
    M10_TestRule db, "R7", M10_Rule07()
    M10_TestRule db, "R8", M10_Rule08()
    M10_TestRule db, "R9", M10_Rule09()
    M10_TestRule db, "R10", M10_Rule10()
    M10_TestRule db, "R11", M10_Rule11()
End Sub

Private Sub M10_TestRule(db As DAO.Database, label As String, sql As String)
    Dim tmpName As String
    Dim rs As DAO.Recordset
    Dim n As Long
    tmpName = "ZZZ_TEST_" & label
    On Error Resume Next
    db.QueryDefs.Delete tmpName
    On Error GoTo 0

    On Error GoTo Err_TR
    db.CreateQueryDef tmpName, sql & ";"
    Set rs = db.OpenRecordset(tmpName, dbOpenSnapshot)
    If Not rs.EOF Then
        rs.MoveLast
        n = rs.RecordCount
    End If
    rs.Close
    Set rs = Nothing
    Debug.Print "[M10] " & label & ": OK, " & n & " finding(s)"
    db.QueryDefs.Delete tmpName
    Exit Sub
Err_TR:
    Debug.Print "[M10] " & label & ": *** FAILS *** " & Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    db.QueryDefs.Delete tmpName
    On Error GoTo 0
End Sub

' Pivots the 20 element columns into long format, one row per
' structure per element, carrying the value of its gating system.
Private Sub M10_BuildElementValues(db As DAO.Database)
    Dim m() As String
    FillElementMap m

    Dim q As String
    Dim i As Integer
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
    M9_CreateQuery db, "QRY_16s_Element_Values", q, "20 elements in long format"
End Sub

' Assembles the battery. full = False stops after rule 11 (M10a);
' M10b calls it again with True to add rules 12-21.
'
' WHY THREE OBJECTS AND NOT ONE. All 21 rules parse and run happily on
' their own, but a single UNION of the 26 resulting branches - many of
' them carrying joins - trips JET's "query too complex" ceiling. The
' halves are well inside it, so the battery is stored as
'     QRY_16a_Rules_1_11   13 branches
'     QRY_16b_Rules_12_21  10 branches
'     QRY_16_Validation_Check   the two, unioned
' The name the delta specifies still holds the whole battery, and the
' halves are separately inspectable during the manual review.
Private Sub M10_BuildValidationCheck(db As DAO.Database, full As Boolean)
    M9_DropQuery db, "QRY_16_Validation_Check"
    M9_DropQuery db, "QRY_16a_Rules_1_11"
    If full Then M9_DropQuery db, "QRY_16b_Rules_12_21"

    M10_BuildHalfA db
    If full Then M10_BuildHalfB db

    Dim c As String
    c = "SELECT * FROM QRY_16a_Rules_1_11"
    If full Then c = c & " UNION ALL SELECT * FROM QRY_16b_Rules_12_21"
    c = c & " ORDER BY Rule_No, Structure;"
    M9_CreateQuery db, "QRY_16_Validation_Check", c, IIf(full, "21 rules", "rules 1-11 (12-21 pending M10b)")
End Sub

Private Sub M10_BuildHalfA(db As DAO.Database)
    Dim q As String

    q = M10_Rule01()
    q = q & " UNION ALL " & M10_Rule02_Elements()
    q = q & " UNION ALL " & M10_Rule02_Systems()
    q = q & " UNION ALL " & M10_Rule03()
    q = q & " UNION ALL " & M10_Rule04()
    q = q & " UNION ALL " & M10_Rule05()
    q = q & " UNION ALL " & M10_Rule06()
    q = q & " UNION ALL " & M10_Rule07()
    q = q & " UNION ALL " & M10_Rule08()
    q = q & " UNION ALL " & M10_Rule09()
    q = q & " UNION ALL " & M10_Rule10()
    q = q & " UNION ALL " & M10_Rule11()
    q = q & ";"
    M9_CreateQuery db, "QRY_16a_Rules_1_11", q, "13 branches"
End Sub

Private Sub M10_BuildHalfB(db As DAO.Database)
    Dim q As String

    q = M10_Rule12()
    q = q & " UNION ALL " & M10_Rule13()
    q = q & " UNION ALL " & M10_Rule14()
    q = q & " UNION ALL " & M10_Rule15()
    q = q & " UNION ALL " & M10_Rule16()
    q = q & " UNION ALL " & M10_Rule17()
    q = q & " UNION ALL " & M10_Rule18()
    q = q & " UNION ALL " & M10_Rule19()
    q = q & " UNION ALL " & M10_Rule20()
    q = q & " UNION ALL " & M10_Rule21()
    q = q & ";"
    M9_CreateQuery db, "QRY_16b_Rules_12_21", q, "10 branches"
End Sub

' R1: a component recorded present while its system says it is not.
' R1 CARRIES TWO EXEMPTIONS THE DELTA DID NOT ANTICIPATE. Applied
' after the rule fired 13 times on the real corpus, all of them
' archaeologically sound - a validation rule that cries wolf teaches
' the researcher to ignore it.
'
' (a) NOT OBSERVABLE IS EXEMPT. Seeing a component while being unable
'     to resolve the system it might belong to is not a contradiction,
'     it is partial observability, which on a cliff face is the norm.
'     'Not observable' says the system-level question cannot be
'     answered, never that the component was not seen. This is the
'     same epistemic move as the REV5 acotacio of rule 6.
'
' (b) TIMBER_BRACKETS IS EXEMPT. Element E is the one item in the
'     vocabulary with an existence independent of its system: an
'     isolated corbel need not have carried a platform - it could have
'     served as a pulley mount or anything else - and a corbel on a
'     funerary chamber whose platform cannot be affirmed does not stop
'     the structure being a funerary chamber. That independence is
'     precisely why v10 introduced Timber_Bracket_Role, and why the
'     MEN typology exists at all. Table 4.5 lists E under Sys_Platform
'     for gating purposes, which is a different question from whether
'     E implies H.
'
' F and G stay inside the rule: transverse beams and corbelled courses
' are far harder to read as anything but platform components, so a
' finding there is worth a look.
Private Function M10_Rule01() As String
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
    M10_Rule01 = q
End Function

' R2a: an element coded 3 with nothing in T_LOST_ELEMENTS to justify
' it. A Body or Whole structure row covers every element of that
' structure in one go (delta section 2).
Private Function M10_Rule02_Elements() As String
    Dim q As String
    q = "SELECT V.Code, 2, "
    q = q & "'R2: ' & V.Field_Name & ' coded 3 (attested lost) with no evidence row', "
    q = q & "'Add a T_LOST_ELEMENTS row, or use 0 / 9 instead' "
    q = q & "FROM QRY_16s_Element_Values AS V "
    q = q & "LEFT JOIN QRY_16s_Lost_Cover AS LC "
    q = q & "ON (V.ID_Structure=LC.ID_Structure) AND (V.Element_Code=LC.Element_Code) "
    q = q & "WHERE V.Element_Value=3 AND LC.ID_Structure Is Null"
    M10_Rule02_Elements = q
End Function

' R2b: the same rule for the three system fields, which also admit
' Attested lost (delta section 2, end).
Private Function M10_Rule02_Systems() As String
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
        ' NOT IN over a subquery with no outer reference: nothing for
        ' JET to correlate, and no derived table in a join to trip on.
        q = q & "FROM T_STRUCTURES AS E "
        q = q & "WHERE E." & s(i, 0) & "='Attested lost' "
        q = q & "AND E.ID NOT IN (SELECT LC.ID_Structure FROM QRY_16s_Lost_Cover AS LC "
        q = q & "WHERE LC.Element_Code='" & s(i, 1) & "')"
    Next i
    M10_Rule02_Systems = q
End Function

' R3: a system declared present whose components are all absent.
' This is the rule that catches DW-S04-EA09.
Private Function M10_Rule03() As String
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
    M10_Rule03 = q
End Function

' R4: under a Not applicable system the components must hold the
' padding 0 (delta 1.2). NULLs are rule 17's business, and JET
' excludes them here anyway since NULL<>0 is not True.
Private Function M10_Rule04() As String
    Dim q As String
    q = "SELECT V.Code, 4, "
    q = q & "'R4: ' & V.Field_Name & ' = ' & V.Element_Value & ' although ' & V.Sys_Field & ' is Not applicable', "
    q = q & "'Components of a non-applicable system are recorded as 0' "
    q = q & "FROM QRY_16s_Element_Values AS V "
    q = q & "WHERE V.Sys_Value='Not applicable' AND V.Element_Value<>0"
    M10_Rule04 = q
End Function

' R5: status Good with more than 30% of the 20 elements partial or
' lost. 30% of 20 is 6, so the test is "more than 6".
Private Function M10_Rule05() As String
    Dim q As String
    q = "SELECT E.Code, 5, "
    q = q & "'R5: recorded Good but more than 30% of elements are partial or lost', "
    q = q & "'Reconcile ID_Arch_Status with the element values' "
    q = q & "FROM ((T_STRUCTURES AS E "
    q = q & "INNER JOIN L_STATUS AS ST ON E.ID_Arch_Status=ST.ID) "
    q = q & "INNER JOIN QRY_16s_Damage_Count AS DC ON E.ID=DC.ID_Structure) "
    q = q & "WHERE ST.Name='Good' AND DC.N_Damaged>6"
    M10_Rule05 = q
End Function

' R6: on a collapsed structure an absence cannot be verified, so a 0
' OF ASSERTION is wrong - 3 (with evidence) or 9 are the legitimate
' values. Padding zeros are excluded, per the REV5 note.
Private Function M10_Rule06() As String
    Dim q As String
    q = "SELECT V.Code, 6, "
    q = q & "'R6: ' & V.Field_Name & ' coded 0 (absent) on a Collapsed structure', "
    q = q & "'Absence is not verifiable on a collapse: use 3 with evidence, or 9' "
    q = q & "FROM ((QRY_16s_Element_Values AS V "
    q = q & "INNER JOIN T_STRUCTURES AS E ON V.ID_Structure=E.ID) "
    q = q & "INNER JOIN L_STATUS AS ST ON E.ID_Arch_Status=ST.ID) "
    q = q & "WHERE ST.Name='Collapsed' AND V.Element_Value=0 "
    q = q & "AND (V.Sys_Field='' OR V.Sys_Value Like 'Present*' OR V.Sys_Value='Attested lost')"
    M10_Rule06 = q
End Function

' R7: an isolated corbel supports no platform, so a platform surface
' material contradicts it.
Private Function M10_Rule07() As String
    Dim q As String
    q = "SELECT E.Code, 7, "
    q = q & "'R7: Platform_Surface_Material set although the corbels are Isolated', "
    q = q & "'Clear the material, or correct Timber_Bracket_Role' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Platform_Surface_Material Is Not Null "
    q = q & "AND E.Timber_Bracket_Role='Isolated'"
    M10_Rule07 = q
End Function

Private Function M10_Rule08() As String
    Dim q As String
    q = "SELECT E.Code, 8, "
    q = q & "'R8: corbels present but Timber_Bracket_Count is empty', "
    q = q & "'Count the corbels: the number carries the platform argument' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Timber_Brackets IN (1,2) AND E.Timber_Bracket_Count Is Null"
    M10_Rule08 = q
End Function

' R9: the role is what separates a platform support from the
' intra-structure counterpart of the MEN typology.
Private Function M10_Rule09() As String
    Dim q As String
    q = "SELECT E.Code, 9, "
    q = q & "'R9: corbels present but Timber_Bracket_Role is not recorded', "
    q = q & "'Set the role: platform support or isolated corbel' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Timber_Brackets IN (1,2) "
    q = q & "AND (E.Timber_Bracket_Role Is Null Or E.Timber_Bracket_Role='ND')"
    M10_Rule09 = q
End Function

' R10: an absent lintel has no material - the counterpart of the
' cleanup M5 performed on the migrated values.
Private Function M10_Rule10() As String
    Dim q As String
    q = "SELECT E.Code, 10, "
    q = q & "'R10: Lintel is 0 (absent) but Lintel_Material is set', "
    q = q & "'Clear Lintel_Material, or correct the Lintel value' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Lintel=0 AND E.Lintel_Material Is Not Null"
    M10_Rule10 = q
End Function

' R11: the operative criterion for N1 - a body counts as a chamber
' body only if it has, or had, an access opening.
Private Function M10_Rule11() As String
    Dim q As String
    q = "SELECT E.Code, 11, "
    q = q & "'R11: chamber bodies counted but Sys_Portal is Absent', "
    q = q & "'A body is N1 only if it has (or had) an access opening' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.N_Chamber_Bodies>0 AND E.Sys_Portal='Absent'"
    M10_Rule11 = q
End Function

' ================================================================
'  BLOCK M10b - QRY_16 rules 12-21, completing the battery
'  DELTA section 12, blocks C (finishes and types), D (record class
'  and gating) and E (bio, materials and connections).
'
'  RULE 17 IS NOT A COMPLAINT, IT IS THE WORKLIST. Delta section 10
'  is explicit that the 35 existing records carry between 5 and 91
'  NULLs each and that resolving them is manual: converting them by
'  script would assert hundreds of unverified absences. Rule 17 is
'  how that backlog stays visible. It is reported ONE ROW PER
'  STRUCTURE with a count, so the battery stays readable - the field
'  by field detail is in QRY_16s_Observational_Nulls, which doubles
'  as the actual review checklist.
'
'  RULE 16 vs RULE 17. They look contradictory until you read the
'  operative definition of "filled" in section 12: a tab closed by
'  Record_Class gating counts as filled only if a bio/materials field
'  holds a 1, or MNI has a value. The padding zeros the gating leaves
'  behind are not content, so they do not trigger rule 16.
' ================================================================
Private Sub M10b_QRY16_Rules12to21(db As DAO.Database)
    Debug.Print "---- [M10b] START ----"

    M9_DropQuery db, "QRY_16_Validation_Check"
    M9_DropQuery db, "QRY_16s_Null_Count"
    M9_DropQuery db, "QRY_16s_Observational_Nulls"

    M10_BuildObservationalNulls db
    M10_BuildNullCount db
    M10_BuildValidationCheck db, True
    M10_DiagnoseRules2 db

    db.QueryDefs.Refresh
    Debug.Print "---- [M10b] END ----"
End Sub

' Every observational field that is still NULL, in long format: the
' 20 element fields plus the 24 three-value ones, 44 in all.
Private Sub M10_BuildObservationalNulls(db As DAO.Database)
    Dim m() As String
    Dim f() As String
    FillElementMap m
    FillThreeValueFields f

    Dim q As String
    Dim i As Integer
    Dim first As Boolean
    first = True

    For i = 0 To 19
        If Not first Then q = q & " UNION ALL "
        q = q & "SELECT E.ID AS ID_Structure, E.Code AS Code, '" & m(i, 1) & "' AS Field_Name "
        q = q & "FROM T_STRUCTURES AS E WHERE E." & m(i, 1) & " Is Null"
        first = False
    Next i
    For i = 0 To 23
        q = q & " UNION ALL "
        q = q & "SELECT E.ID, E.Code, '" & f(i) & "' "
        q = q & "FROM T_STRUCTURES AS E WHERE E." & f(i) & " Is Null"
    Next i
    q = q & ";"
    M9_CreateQuery db, "QRY_16s_Observational_Nulls", q, "44 observational fields, NULLs only"
End Sub

' Counts the pending NULLs per structure as a plain sum of 44 IIFs on
' a single table - deliberately NOT by grouping the unpivot query.
' JET inlines a saved query wherever it is referenced, so building
' this on the 44-branch QRY_16s_Observational_Nulls made rule 17 drag
' those 44 branches into the battery and blew the "query too complex"
' ceiling. The unpivot survives as the standalone review checklist,
' where it works fine; nothing in QRY_16 depends on it any more.
Private Sub M10_BuildNullCount(db As DAO.Database)
    Dim m() As String
    Dim f() As String
    FillElementMap m
    FillThreeValueFields f

    Dim q As String
    Dim i As Integer
    q = "SELECT E.ID AS ID_Structure, E.Code AS Code, "
    For i = 0 To 19
        If i > 0 Then q = q & "+"
        q = q & "IIF(E." & m(i, 1) & " Is Null,1,0)"
    Next i
    For i = 0 To 23
        q = q & "+IIF(E." & f(i) & " Is Null,1,0)"
    Next i
    q = q & " AS N_Null "
    q = q & "FROM T_STRUCTURES AS E;"
    M9_CreateQuery db, "QRY_16s_Null_Count", q, "pending NULLs per structure, no unpivot"
End Sub

Private Sub M10_DiagnoseRules2(db As DAO.Database)
    Debug.Print "[M10] --- per-rule diagnosis (12-21) ---"
    M10_TestRule db, "R12", M10_Rule12()
    M10_TestRule db, "R13", M10_Rule13()
    M10_TestRule db, "R14", M10_Rule14()
    M10_TestRule db, "R15", M10_Rule15()
    M10_TestRule db, "R16", M10_Rule16()
    M10_TestRule db, "R17", M10_Rule17()
    M10_TestRule db, "R18", M10_Rule18()
    M10_TestRule db, "R19", M10_Rule19()
    M10_TestRule db, "R20", M10_Rule20()
    M10_TestRule db, "R21", M10_Rule21()
End Sub

' R12: painting on a prepared render and painting straight onto the
' masonry are two different operative sequences - the substrate is
' the whole point of splitting the field (delta v10 rationale).
Private Function M10_Rule12() As String
    Dim q As String
    ' First branch of QRY_16b, so it carries the explicit aliases that
    ' name that half's columns when it is opened on its own.
    q = "SELECT E.Code AS Structure, 12 AS Rule_No, "
    q = q & "'R12: pigment present but the substrate is not recorded' AS Rule_Violated, "
    q = q & "'Set Pigment_Substrate: plaster vs masonry vs bedrock is the point' AS Action "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Pigment_Present=1 "
    q = q & "AND (E.Pigment_Substrate Is Null Or E.Pigment_Substrate='ND')"
    M10_Rule12 = q
End Function

Private Function M10_Rule13() As String
    Dim q As String
    q = "SELECT E.Code, 13, "
    q = q & "'R13: pigment recorded on plaster but plaster is absent', "
    q = q & "'Reconcile Plaster_Present with Pigment_Substrate' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Pigment_Substrate='Plaster' AND E.Plaster_Present=0"
    M10_Rule13 = q
End Function

' R14: element X carries presence; the type field carries the
' natural-bedrock vs built distinction, which is the constructive
' decision the analysis needs.
Private Function M10_Rule14() As String
    Dim q As String
    q = "SELECT E.Code, 14, "
    q = q & "'R14: chamber roof present but its type is not recorded', "
    q = q & "'Set Chamber_Roof_Type: natural bedrock vs built is the decision' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Chamber_Roof IN (1,2) "
    q = q & "AND (E.Chamber_Roof_Type Is Null Or E.Chamber_Roof_Type='ND')"
    M10_Rule14 = q
End Function

Private Function M10_Rule15() As String
    Dim q As String
    q = "SELECT E.Code, 15, "
    q = q & "'R15: classified as a structural trace but human remains are recorded', "
    q = q & "'Reclassify the record, or correct Human_Remains' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID "
    q = q & "WHERE T.Record_Class='Structural trace' AND E.Human_Remains=1"
    M10_Rule15 = q
End Function

' R16: a tab closed by Record_Class gating holds only padding zeros.
' Real content there means the classification or the data is wrong.
Private Function M10_Rule16() As String
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
    M10_Rule16 = q
End Function

' R17: the manual-review backlog of delta section 10. One row per
' structure; QRY_16s_Observational_Nulls has the field by field list.
Private Function M10_Rule17() As String
    Dim q As String
    q = "SELECT NC.Code, 17, "
    q = q & "'R17: ' & NC.N_Null & ' observational field(s) still NULL (not yet assessed)', "
    q = q & "'Resolve each to 0, 1, 2, 3 or 9 - see QRY_16s_Observational_Nulls' "
    q = q & "FROM QRY_16s_Null_Count AS NC "
    q = q & "WHERE NC.N_Null>0"
    M10_Rule17 = q
End Function

Private Function M10_Rule18() As String
    Dim q As String
    q = "SELECT E.Code, 18, "
    q = q & "'R18: human remains recorded but MNI is empty', "
    q = q & "'Set the minimum number of individuals' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Human_Remains=1 AND E.MNI Is Null"
    M10_Rule18 = q
End Function

Private Function M10_Rule19() As String
    Dim q As String
    q = "SELECT E.Code, 19, "
    q = q & "'R19: movable remains recorded Absent but a Mat_ field is not 0', "
    q = q & "'Reconcile ID_Material_Status with the Mat_ fields' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "INNER JOIN L_MATERIAL_STATUS AS MS ON E.ID_Material_Status=MS.ID "
    q = q & "WHERE MS.Name='Absent' "
    q = q & "AND (E.Mat_Textiles<>0 OR E.Mat_Wood<>0 OR E.Mat_VegFiber<>0 "
    q = q & "OR E.Mat_Ceramics<>0 OR E.Mat_Fauna<>0 OR E.Mat_DeerAntler<>0 "
    q = q & "OR E.Mat_Other<>0)"
    M10_Rule19 = q
End Function

' R20: direction is read from the joint. Only an abutted vertical
' joint or a superposition can carry it (delta 9.1bis); a bonded
' joint implies Contemporary, and the rest imply nothing.
Private Function M10_Rule20() As String
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
    M10_Rule20 = q
End Function

' R21: the counterpart of making Element_Code optional (delta 2).
Private Function M10_Rule21() As String
    Dim q As String
    q = "SELECT E.Code, 21, "
    q = q & "'R21: lost-element row scoped Element but no Element_Code given', "
    q = q & "'Element_Code is required when the scope is Element' "
    q = q & "FROM T_LOST_ELEMENTS AS L "
    q = q & "INNER JOIN T_STRUCTURES AS E ON L.ID_Structure=E.ID "
    q = q & "WHERE L.Evidence_Scope='Element' AND L.Element_Code Is Null"
    M10_Rule21 = q
End Function

' ================================================================
'  BLOCKS M11a / M11b - The final deletions (delta step 13.14)
'  DELTA sections: 7.2 and 7.5 (decoration booleans), 2 and 8
'  (Lost_Body_Evidence), 4.1 (H/P/U absorbed), 5.5 (W removed).
'
'  THIS IS THE ONLY IRREVERSIBLE PART OF THE MIGRATION. Rule A says
'  nothing is deleted before its replacement is populated AND
'  verified, so neither block trusts a human memory of "yes, I did
'  the review": each re-derives its precondition from the live data
'  every time it runs, and refuses if the replacement is not there.
'
'  WHY THE BOOLEANS GO (7.1): with only 35 records the Dec_* flags
'  and T_DECORATIONS already disagree in 6 cases - 17%. Two places
'  recording the same fact is one place too many, and T_DECORATIONS
'  is the one that also carries position, type, colour and substrate.
' ================================================================
Private Sub M11a_DeleteDecorationBooleans(db As DAO.Database)
    Debug.Print "---- [M11a] START ----"

    ' Motif booleans paired with the L_DEC_TYPE name that must back them.
    Dim d(6, 1) As String
    d(0, 0) = "Dec_Square_Niche": d(0, 1) = "Square niche"
    d(1, 0) = "Dec_Relief_T":     d(1, 1) = "T-shaped niche"
    d(2, 0) = "Dec_Relief_T_Inv": d(2, 1) = "T-shaped niche inv."
    d(3, 0) = "Dec_Relief_L":     d(3, 1) = "L-shaped niche"
    d(4, 0) = "Dec_Relief_L_Inv": d(4, 1) = "L-shaped niche inv."
    d(5, 0) = "Dec_Zigzag":       d(5, 1) = "Zigzag"
    d(6, 0) = "Dec_Stepped":      d(6, 1) = "Stepped motif"

    Dim unbacked As Long
    Dim i As Integer
    unbacked = 0

    For i = 0 To 6
        If FldExists(db, "T_STRUCTURES", d(i, 0)) Then
            Dim n As Long
            n = DCount("*", "T_STRUCTURES", d(i, 0) & "=1 AND ID NOT IN (SELECT ID_Structure FROM QRY_02s_Decoration_Typed WHERE Dec_Name='" & d(i, 1) & "')")
            If n > 0 Then Debug.Print "[M11a] " & d(i, 0) & "=1 with no '" & d(i, 1) & "' row: " & n
            unbacked = unbacked + n
        End If
    Next i

    ' Rock art is replaced by pigment on bedrock (7.2), so its evidence
    ' is Pigment_Present=1 with Pigment_Substrate='Bedrock'.
    If FldExists(db, "T_STRUCTURES", "Rock_Art") Then
        Dim nRA As Long
        nRA = DCount("*", "T_STRUCTURES", "Rock_Art=1 AND NOT (Pigment_Present=1 AND Pigment_Substrate='Bedrock')")
        If nRA > 0 Then Debug.Print "[M11a] Rock_Art=1 without pigment on bedrock: " & nRA
        unbacked = unbacked + nRA
    End If

    Debug.Print "[M11a] Precondition (delta 7.5): " & unbacked & " boolean(s) at 1 with no backing record (must be 0)"

    If unbacked > 0 Then
        Debug.Print "[M11a] NOT DELETING: migrate those to T_DECORATIONS first"
        Debug.Print "---- [M11a] END ----"
        Exit Sub
    End If
    If Not ALLOW_FINAL_DELETIONS Then
        Debug.Print "[M11a] Precondition met, but ALLOW_FINAL_DELETIONS is False - nothing deleted"
        Debug.Print "---- [M11a] END ----"
        Exit Sub
    End If

    Dim f(11) As String
    f(0) = "Dec_Square_Niche":   f(1) = "Dec_Relief_T":      f(2) = "Dec_Relief_T_Inv"
    f(3) = "Dec_Relief_L":       f(4) = "Dec_Relief_L_Inv":  f(5) = "Dec_Zigzag"
    f(6) = "Dec_Stepped":        f(7) = "Rock_Art":          f(8) = "RA_Anthropomorphic"
    f(9) = "RA_Zoomorphic":      f(10) = "RA_Geometric":     f(11) = "RA_Abstract"

    Dim nDel As Integer
    nDel = 0
    For i = 0 To 11
        If M11_DropField(db, f(i)) Then nDel = nDel + 1
    Next i
    db.TableDefs.Refresh
    Debug.Print "[M11a] Fields deleted this run: " & nDel & " (12 the first time)"
    Debug.Print "---- [M11a] END ----"
End Sub

Private Sub M11b_DeleteAbsorbedFields(db As DAO.Database)
    Debug.Print "---- [M11b] START ----"

    ' Precondition 1: every Sys_* resolved. Until then the H/P/U
    ' information still only exists in the old fields.
    Dim blanks As Long
    blanks = 0
    blanks = blanks + M11_BlankCount(db, "Sys_Platform")
    blanks = blanks + M11_BlankCount(db, "Sys_Portal")
    blanks = blanks + M11_BlankCount(db, "Sys_Eave")
    blanks = blanks + M11_BlankCount(db, "Sys_Base")
    blanks = blanks + M11_BlankCount(db, "Sys_Chamber")
    Debug.Print "[M11b] Precondition A: " & blanks & " unassigned Sys_* cell(s) (must be 0 of 175)"

    ' Precondition 2: the real Lost_Body_Evidence values have become
    ' T_LOST_ELEMENTS rows. None/ND never generate one (delta 2).
    Dim unmigrated As Long
    unmigrated = 0
    If FldExists(db, "T_STRUCTURES", "Lost_Body_Evidence") Then
        unmigrated = DCount("*", "T_STRUCTURES", "Lost_Body_Evidence Is Not Null AND Lost_Body_Evidence<>'None' AND Lost_Body_Evidence<>'ND' AND ID NOT IN (SELECT ID_Structure FROM T_LOST_ELEMENTS)")
    End If
    Debug.Print "[M11b] Precondition B: " & unmigrated & " Lost_Body_Evidence value(s) not yet in T_LOST_ELEMENTS (must be 0)"

    If blanks > 0 Or unmigrated > 0 Then
        Debug.Print "[M11b] NOT DELETING: finish the assignment and the migration first"
        Debug.Print "---- [M11b] END ----"
        Exit Sub
    End If
    If Not ALLOW_FINAL_DELETIONS Then
        Debug.Print "[M11b] Preconditions met, but ALLOW_FINAL_DELETIONS is False - nothing deleted"
        Debug.Print "---- [M11b] END ----"
        Exit Sub
    End If

    Dim f(4) As String
    f(0) = "Lost_Body_Evidence"
    f(1) = "Corbelled_Platform"
    f(2) = "Access_Opening"
    f(3) = "Eave"
    f(4) = "Rear_Wall"

    Dim nDel As Integer
    Dim i As Integer
    nDel = 0
    For i = 0 To 4
        If M11_DropField(db, f(i)) Then nDel = nDel + 1
    Next i
    db.TableDefs.Refresh
    Debug.Print "[M11b] Fields deleted this run: " & nDel & " (5 the first time)"
    Debug.Print "[M11b] Now run Compact and Repair from the Access UI to reclaim the space"
    Debug.Print "---- [M11b] END ----"
End Sub

Private Function M11_BlankCount(db As DAO.Database, fld As String) As Long
    If Not FldExists(db, "T_STRUCTURES", fld) Then Exit Function
    M11_BlankCount = DCount("*", "T_STRUCTURES", "[" & fld & "] Is Null")
End Function

Private Function M11_DropField(db As DAO.Database, fld As String) As Boolean
    If Not FldExists(db, "T_STRUCTURES", fld) Then
        Debug.Print "[M11] " & fld & " already gone - skipped"
        Exit Function
    End If
    On Error GoTo Err_DF
    db.Execute "ALTER TABLE T_STRUCTURES DROP COLUMN " & fld, dbFailOnError
    Debug.Print "[M11] " & fld & " DELETED"
    M11_DropField = True
    Exit Function
Err_DF:
    Debug.Print "[M11] ERROR deleting " & fld & ": " & Err.Description
End Function

' ================================================================
'  VERIFICATION (grows with each block; final version after M11b)
' ================================================================
Sub VerifyMigration()
    Dim db As DAO.Database
    Set db = CurrentDb()

    Debug.Print "================================================================"
    Debug.Print "VERIFY MIGRATION (" & Now & ")"
    Debug.Print "================================================================"

    ' --- Checks valid from M1 onward ---
    Debug.Print "[Verify] T_STRUCTURES record count: " & DCount("*", "T_STRUCTURES") & " (expected 35)"

    If TblExists(db, "L_ELEMENTS") Then
        Debug.Print "[Verify] L_ELEMENTS row count: " & DCount("*", "L_ELEMENTS") & " (expected 24)"
    Else
        Debug.Print "[Verify] L_ELEMENTS: TABLE MISSING"
    End If

    If TblExists(db, "L_LOST_EVIDENCE") Then
        Debug.Print "[Verify] L_LOST_EVIDENCE row count: " & DCount("*", "L_LOST_EVIDENCE") & " (expected 9)"
    Else
        Debug.Print "[Verify] L_LOST_EVIDENCE: TABLE MISSING"
    End If

    If FldExists(db, "L_TYPOLOGY", "Record_Class") Then
        Debug.Print "[Verify] L_TYPOLOGY rows with Record_Class NULL: " & DCount("*", "L_TYPOLOGY", "Record_Class Is Null") & " (expected 0)"
    Else
        Debug.Print "[Verify] L_TYPOLOGY.Record_Class: FIELD MISSING"
    End If

    Debug.Print "[Verify] L_TYPOLOGY rows named MIX Mixed: " & DCount("*", "L_TYPOLOGY", "Name='MIX Mixed'") & " (expected 0)"
    Debug.Print "[Verify] T_STRUCTURES rows with ID_Typology Null and Code='DW-S01-EA09': " & DCount("*", "T_STRUCTURES", "Code='DW-S01-EA09' And ID_Typology Is Null") & " (expected 0)"

    ' --- Checks valid from M3 onward ---
    ' Defaults must read "0". A field still showing "9" was missed by M3.
    Debug.Print "[Verify] --- M3 defaults (sample) ---"
    VerifyDefault db, "Embedded_Base_Beams", "0"
    VerifyDefault db, "Chamber_Roof", "0"
    VerifyDefault db, "Human_Remains", "0"
    VerifyDefault db, "Pigment_Present", "0"

    ' Rule B guard: no block may ever touch a stored value. The
    ' baselines below were measured right after M3 ran on the live
    ' database; they must stay identical until the manual review of
    ' section 10 begins. A drop means something converted NULL to 0.
    ' Skipped on an empty database: this same routine is useful for
    ' checking a fresh BuildDB, and three false INVESTIGATE lines
    ' every run would only teach you to ignore the real ones.
    Debug.Print "[Verify] --- Rule B: existing NULLs must be preserved ---"
    If DCount("*", "T_STRUCTURES") = 0 Then
        Debug.Print "[Verify] (empty database - baselines not applicable, skipped)"
    Else
        VerifyNullCount "Base_Level", 5
        VerifyNullCount "Human_Remains", 8
        VerifyNullCount "Chamber_Roof", 18
    End If

    ' --- Checks valid from M4 onward ---
    Debug.Print "[Verify] --- M4 renames ---"
    VerifyRenamed db, "Lateral_Walls", "Return_Wall"
    VerifyRenamed db, "Lateral_Wall_Faces", "Facade_Flank"
    VerifyRenamed db, "Recessed_Portal", "Recessed_Frame"
    VerifyRenamed db, "Rear_Wall_Type", "Rear_Closure_Type"
    Debug.Print "[Verify] L_STRUCT_BODY rows with Code='LWF': " & DCount("*", "L_STRUCT_BODY", "Code='LWF'") & " (expected 0)"
    Debug.Print "[Verify] L_STRUCT_BODY rows with Code='FFL': " & DCount("*", "L_STRUCT_BODY", "Code='FFL'") & " (expected 1)"

    ' --- Checks valid from M5 onward ---
    ' The delta's own acceptance test for the Lintel split (5.4 step 5).
    Debug.Print "[Verify] --- M5 Lintel split ---"
    If FldExists(db, "T_STRUCTURES", "Lintel") And FldExists(db, "T_STRUCTURES", "Lintel_Material") Then
        Dim v1 As Long, v0 As Long, v9 As Long, vN As Long
        v1 = DCount("*", "T_STRUCTURES", "Lintel=1")
        v0 = DCount("*", "T_STRUCTURES", "Lintel=0")
        v9 = DCount("*", "T_STRUCTURES", "Lintel=9")
        vN = DCount("*", "T_STRUCTURES", "Lintel Is Null")
        Debug.Print "[Verify] Lintel 1/0/9 = " & v1 & "/" & v0 & "/" & v9 & " | total migrated " & (v1 + v0 + v9) & " (expected 21)"
        Debug.Print "[Verify] Lintel NULL: " & vN & " (expected 14, resolved by manual review)"
        Debug.Print "[Verify] Lintel is BYTE: " & (db.TableDefs("T_STRUCTURES").Fields("Lintel").Type = dbByte)
        ' Delta rule 10 of QRY_16: an absent lintel cannot have a material.
        Debug.Print "[Verify] Lintel=0 with a material set: " & DCount("*", "T_STRUCTURES", "Lintel=0 And Lintel_Material Is Not Null") & " (expected 0)"
        Debug.Print "[Verify] Lintel=9 with a material set: " & DCount("*", "T_STRUCTURES", "Lintel=9 And Lintel_Material Is Not Null") & " (expected 0)"
    Else
        Debug.Print "[Verify] Lintel split: NOT DONE (one of the two fields is missing)"
    End If

    ' --- Checks valid from M6 onward ---
    Debug.Print "[Verify] --- M6 system fields ---"
    VerifyFieldPresent db, "Sys_Platform"
    VerifyFieldPresent db, "Sys_Portal"
    VerifyFieldPresent db, "Sys_Eave"
    VerifyFieldPresent db, "Sys_Base"
    VerifyFieldPresent db, "Sys_Chamber"
    VerifyFieldPresent db, "Platform_Function"

    ' Invariant of delta 4.1: a Sys_* may only be blank where the v10
    ' source held 1 (needs judgement) or was itself NULL. If more rows
    ' are blank than that, a scripted 0/9 failed to translate.
    VerifySystemMigration db, "Corbelled_Platform", "Sys_Platform"
    VerifySystemMigration db, "Access_Opening", "Sys_Portal"
    VerifySystemMigration db, "Eave", "Sys_Eave"

    ' --- Checks valid from M7 onward ---
    Debug.Print "[Verify] --- M7 lost elements and connections ---"
    If TblExists(db, "T_LOST_ELEMENTS") Then
        Debug.Print "[Verify] T_LOST_ELEMENTS: present | rows " & DCount("*", "T_LOST_ELEMENTS")
        VerifyRelation db, "REL_STR_LOST"
        VerifyRelation db, "REL_ELEM_LOST"
        VerifyRelation db, "REL_LEV_LOST"
        VerifyRelation db, "REL_SB_LOST"
        ' Rule 21 of QRY_16, checkable as soon as rows start appearing.
        Debug.Print "[Verify] Scope='Element' rows without Element_Code: " & DCount("*", "T_LOST_ELEMENTS", "Evidence_Scope='Element' And Element_Code Is Null") & " (expected 0)"
    Else
        Debug.Print "[Verify] T_LOST_ELEMENTS: TABLE MISSING"
    End If
    If FldExists(db, "T_CONNECTIONS", "Chrono_Relation") Then
        Debug.Print "[Verify] T_CONNECTIONS.Chrono_Relation: present | rows " & DCount("*", "T_CONNECTIONS")
    Else
        Debug.Print "[Verify] T_CONNECTIONS.Chrono_Relation: FIELD MISSING"
    End If

    ' --- Checks valid from M9a onward ---
    ' Existence proves nothing: a query can exist and still fail on
    ' open because it names a field that was renamed. Open each one.
    Debug.Print "[Verify] --- M9a queries (opened, not just listed) ---"
    VerifyQueryRuns db, "QRY_02a_Decoration_Flags"
    VerifyQueryRuns db, "QRY_02_Decoration_by_Site"
    VerifyQueryRuns db, "QRY_05_Export_RStats"
    VerifyQueryRuns db, "QRY_07_Export_QGIS"

    VerifyQueryRuns db, "QRY_13_AX_Pattern_Export"
    VerifyQueryRuns db, "QRY_14_Connections_Edges"

    ' QRY_13 must carry exactly 20 AX_ and 5 SYS_ columns (delta 12bis),
    ' and must no longer carry AX_H / AX_P / AX_U / AX_W.
    VerifyAXColumns db

    ' --- Checks valid from M10a onward ---
    Debug.Print "[Verify] --- M10 validation battery ---"
    VerifyQueryRuns db, "QRY_16s_Element_Values"
    VerifyQueryRuns db, "QRY_16s_Lost_Cover"
    VerifyQueryRuns db, "QRY_16s_Damage_Count"
    VerifyQueryRuns db, "QRY_16s_Observational_Nulls"
    VerifyQueryRuns db, "QRY_16a_Rules_1_11"
    VerifyQueryRuns db, "QRY_16b_Rules_12_21"
    VerifyQueryRuns db, "QRY_16_Validation_Check"
    ' 20 elements x 35 structures = 700 rows expected in the helper.
    ' QRY_16 rows are findings, not errors in the script: a non-zero
    ' count is the report doing its job.
    ReportValidationBreakdown db

    ' --- Checks valid from M11 onward ---
    ' The user's fourth acceptance criterion: no retired field may
    ' still be named by any query. Checked against every QueryDef,
    ' not just the ones this script wrote.
    Debug.Print "[Verify] --- M11 retired fields ---"
    VerifyRetiredFields db
    ' TODO M5: Lintel BYTE + Lintel_Material TEXT both exist; 21 values migrated; NULLs preserved
    ' TODO M6: Sys_* fields exist; H/P/U 0/9 migrated; 1/NULL left NULL
    ' TODO M7: T_LOST_ELEMENTS exists with its 4 relationships
    ' TODO M9/M10: no QueryDef.SQL references a field no longer in T_STRUCTURES
    ' TODO M11a/M11b: Dec_*/RA_*/Rock_Art/Lost_Body_Evidence/Corbelled_Platform/
    '                 Access_Opening/Eave/Rear_Wall no longer exist

    Debug.Print "================================================================"
    Debug.Print "VERIFY MIGRATION: DONE. Run Compact and Repair manually from the"
    Debug.Print "Access UI (Database Tools) once all blocks are complete - it"
    Debug.Print "cannot be invoked from code on the currently open connection."
    Debug.Print "================================================================"

    Set db = Nothing
End Sub

' Scans every saved query for references to a field that v11 retires.
' Runs before the deletions too, where it is exactly as useful: it
' says which queries would break the moment M11 fires.
Private Sub VerifyRetiredFields(db As DAO.Database)
    Dim gone(16) As String
    gone(0) = "Dec_Square_Niche":   gone(1) = "Dec_Relief_T":      gone(2) = "Dec_Relief_T_Inv"
    gone(3) = "Dec_Relief_L":       gone(4) = "Dec_Relief_L_Inv":  gone(5) = "Dec_Zigzag"
    gone(6) = "Dec_Stepped":        gone(7) = "Rock_Art":          gone(8) = "RA_Anthropomorphic"
    gone(9) = "RA_Zoomorphic":      gone(10) = "RA_Geometric":     gone(11) = "RA_Abstract"
    gone(12) = "Lost_Body_Evidence": gone(13) = "Corbelled_Platform"
    gone(14) = "Access_Opening":    gone(15) = "Eave":             gone(16) = "Rear_Wall"

    Dim q As DAO.QueryDef
    Dim i As Integer
    Dim hits As Long
    hits = 0
    For Each q In db.QueryDefs
        If Left(q.Name, 1) <> "~" Then
            For i = 0 To 16
                ' Word boundaries matter: "Eave" must not match
                ' "Eave_Beam", and "Rock_Art" must not match "Rock_Art"
                ' inside a longer name.
                If M11_NamesField(q.SQL, gone(i)) Then
                    Debug.Print "[Verify] " & q.Name & " still names " & gone(i)
                    hits = hits + 1
                End If
            Next i
        End If
    Next q
    If hits = 0 Then
        Debug.Print "[Verify] No query references a retired field (OK)"
    Else
        Debug.Print "[Verify] " & hits & " reference(s) to retired fields - these queries break when M11 runs"
    End If
End Sub

' True when sql names fld as a whole identifier, not as the prefix of
' a longer one. Eave vs Eave_Beam is the case that matters.
Private Function M11_NamesField(sql As String, fld As String) As Boolean
    Dim p As Long
    Dim after As String
    Dim before As String
    p = InStr(1, sql, fld, vbTextCompare)
    Do While p > 0
        before = ""
        If p > 1 Then before = Mid(sql, p - 1, 1)
        after = Mid(sql, p + Len(fld), 1)
        If before <> "_" And Not (before Like "[A-Za-z0-9]") Then
            If after <> "_" And Not (after Like "[A-Za-z0-9]") Then
                M11_NamesField = True
                Exit Function
            End If
        End If
        p = InStr(p + 1, sql, fld, vbTextCompare)
    Loop
End Function

' Findings per rule. These are not script errors - the battery is a
' report, and a non-empty result is it working. Printed per rule so
' the manual review has a worklist rather than a single total.
Private Sub ReportValidationBreakdown(db As DAO.Database)
    Dim rs As DAO.Recordset
    On Error GoTo Err_VB
    Set rs = db.OpenRecordset("SELECT Rule_No, COUNT(*) AS N FROM QRY_16_Validation_Check GROUP BY Rule_No ORDER BY Rule_No", dbOpenSnapshot)
    If rs.EOF Then
        Debug.Print "[Verify] QRY_16: no findings - the corpus is coherent on the active rules"
    Else
        Do While Not rs.EOF
            Debug.Print "[Verify] QRY_16 rule " & rs!Rule_No & ": " & rs!N & " finding(s)"
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
    Exit Sub
Err_VB:
    Debug.Print "[Verify] QRY_16 breakdown failed: " & Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
End Sub

' The A-X matrix went from 24 element columns to 20: H, P and U became
' systems and W was dropped altogether (delta 5.5, 12bis).
Private Sub VerifyAXColumns(db As DAO.Database)
    Dim rs As DAO.Recordset
    Dim f As DAO.Field
    Dim nAX As Integer
    Dim nSYS As Integer
    Dim gone As String
    On Error GoTo Err_AX
    Set rs = db.OpenRecordset("QRY_13_AX_Pattern_Export", dbOpenSnapshot)
    For Each f In rs.Fields
        If Left(f.Name, 3) = "AX_" Then nAX = nAX + 1
        If Left(f.Name, 4) = "SYS_" Then nSYS = nSYS + 1
        If f.Name = "AX_H" Or f.Name = "AX_P" Or f.Name = "AX_U" Or f.Name = "AX_W" Then
            gone = gone & f.Name & " "
        End If
    Next f
    rs.Close
    Set rs = Nothing
    Debug.Print "[Verify] QRY_13 element columns: " & nAX & " (expected 20) | system columns: " & nSYS & " (expected 5)"
    If Len(gone) > 0 Then
        Debug.Print "[Verify] QRY_13 still exports retired columns: " & gone & "- INVESTIGATE"
    Else
        Debug.Print "[Verify] QRY_13 no longer exports AX_H / AX_P / AX_U / AX_W (OK)"
    End If
    Exit Sub
Err_AX:
    Debug.Print "[Verify] QRY_13 column check failed: " & Err.Description
End Sub

' Actually opens the query. This is the check that catches a stale
' field name, a bad JOIN nesting or an unbalanced parenthesis - none
' of which show up merely by asking whether the QueryDef exists.
Private Sub VerifyQueryRuns(db As DAO.Database, qn As String)
    Dim rs As DAO.Recordset
    On Error GoTo Err_Q
    Set rs = db.OpenRecordset(qn, dbOpenSnapshot)
    Dim n As Long
    n = 0
    If Not rs.EOF Then
        rs.MoveLast
        n = rs.RecordCount
    End If
    Debug.Print "[Verify] " & qn & ": opens OK, " & n & " row(s), " & rs.Fields.Count & " column(s)"
    rs.Close
    Set rs = Nothing
    Exit Sub
Err_Q:
    Debug.Print "[Verify] " & qn & ": *** FAILS TO OPEN *** " & Err.Description
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    On Error GoTo 0
End Sub

Private Sub VerifyRelation(db As DAO.Database, nm As String)
    Dim r As DAO.Relation
    For Each r In db.Relations
        If r.Name = nm Then
            Debug.Print "[Verify] " & nm & ": present"
            Exit Sub
        End If
    Next r
    Debug.Print "[Verify] " & nm & ": RELATIONSHIP MISSING"
End Sub

Private Sub VerifyFieldPresent(db As DAO.Database, fld As String)
    If FldExists(db, "T_STRUCTURES", fld) Then
        Debug.Print "[Verify] " & fld & ": present | assigned " & DCount("*", "T_STRUCTURES", "[" & fld & "] Is Not Null") & " / 35"
    Else
        Debug.Print "[Verify] " & fld & ": FIELD MISSING"
    End If
End Sub

' A blank Sys_* is legitimate only where the v10 field held 1 or NULL.
Private Sub VerifySystemMigration(db As DAO.Database, srcFld As String, sysFld As String)
    If Not FldExists(db, "T_STRUCTURES", srcFld) Then
        Debug.Print "[Verify] " & sysFld & ": source " & srcFld & " already dropped - check skipped"
        Exit Sub
    End If
    Dim expectedBlank As Long
    Dim actualBlank As Long
    expectedBlank = DCount("*", "T_STRUCTURES", srcFld & "=1 Or " & srcFld & " Is Null")
    actualBlank = DCount("*", "T_STRUCTURES", "[" & sysFld & "] Is Null")
    If actualBlank <= expectedBlank Then
        Debug.Print "[Verify] " & sysFld & " blank: " & actualBlank & " (at most " & expectedBlank & " pending, OK)"
    Else
        Debug.Print "[Verify] " & sysFld & " blank: " & actualBlank & " (EXCEEDS the " & expectedBlank & " expected - a 0/9 failed to migrate)"
    End If
End Sub

Private Sub VerifyNullCount(fld As String, baseline As Long)
    Dim n As Long
    n = DCount("*", "T_STRUCTURES", "[" & fld & "] Is Null")
    If n = baseline Then
        Debug.Print "[Verify] " & fld & " NULLs: " & n & " (baseline " & baseline & ", OK)"
    Else
        Debug.Print "[Verify] " & fld & " NULLs: " & n & " (BASELINE WAS " & baseline & " - INVESTIGATE)"
    End If
End Sub

Private Sub VerifyRenamed(db As DAO.Database, oldName As String, newName As String)
    Dim hasOld As Boolean
    Dim hasNew As Boolean
    hasOld = FldExists(db, "T_STRUCTURES", oldName)
    hasNew = FldExists(db, "T_STRUCTURES", newName)
    If hasNew And Not hasOld Then
        Debug.Print "[Verify] " & oldName & " -> " & newName & " (OK)"
    ElseIf hasOld And Not hasNew Then
        Debug.Print "[Verify] " & newName & ": RENAME NOT DONE (" & oldName & " still present)"
    ElseIf hasOld And hasNew Then
        Debug.Print "[Verify] " & newName & ": BOTH NAMES PRESENT - INVESTIGATE"
    Else
        Debug.Print "[Verify] " & newName & ": NEITHER NAME PRESENT - INVESTIGATE"
    End If
End Sub

Private Sub VerifyDefault(db As DAO.Database, fld As String, expected As String)
    If Not FldExists(db, "T_STRUCTURES", fld) Then
        Debug.Print "[Verify] " & fld & ": FIELD MISSING"
        Exit Sub
    End If
    Dim actual As String
    actual = "" & db.TableDefs("T_STRUCTURES").Fields(fld).DefaultValue
    If actual = expected Then
        Debug.Print "[Verify] " & fld & " default = " & actual & " (OK)"
    Else
        Debug.Print "[Verify] " & fld & " default = " & actual & " (EXPECTED " & expected & ")"
    End If
End Sub
