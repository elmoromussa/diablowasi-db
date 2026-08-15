Option Compare Database
Option Explicit

' ================================================================
'  CHACHAPOYA ARCHAEOLOGICAL DATABASE - COMPLETE BUILD SCRIPT v20
'  La Petaca & Diablo Wasi (Leymebamba, Amazonas, Peru)
'  Author: Esteve Ribera Torro | TFM Arqueologia UA
'  Spec: DELTA_v19_v20.md
'
'  Run Sub BuildDB() on a NEW BLANK ACCESS DATABASE.
'  Then run chachapoya_Form_v20_val.bas -> Sub BuildForm()
'
'  WHAT CHANGED IN v16 (delta v15->v16)
'
'  ============ v17 (delta v16->v17) ============
'
'  V1. POSITION_RELATIVE REMOVED FROM T_DECORATIONS (A1). The
'      field served two halves of tab 4.Dec that DO NOT SHARE A
'      REFERENT. On architectural rows Left named WHICH INSTANCE
'      of a paired element bore the motif - an identificational
'      relation internal to the structure. On ROC rows it named
'      WHERE THE PAINTING LIES with respect to the built volume -
'      a topological relation between two separate objects. Three
'      consequences were live errors, not discomfort:
'      (a) Above/Below meant nothing on architectural rows: the
'          vertical axis is ID_Struct_Body x Body_No.
'      (b) Both meant two different things per half.
'      (c) The declared reference plane (Access_Plane) is the
'          WRONG one for rock art, which is seen from the valley,
'          not from the entrance: on a structure with
'          Access_Plane = Return wall the convention INVERTED
'          left and right.
'      Corpus evidence: 7 of 56 rows carried a value, ALL of them
'      ROC, none architectural. No replacement on the
'      architectural half - laterality goes to Notes until a real
'      case appears (same criterion that kept the ROC positions
'      to four: no categories without cases).
'
'  V2. FOUR SPAN FIELDS ON T_DECORATIONS (A1.3). Span_Left,
'      Span_Above, Span_Right, Span_Below, domain 0/1/9, NULL by
'      default. They record WHICH SEGMENTS OF THE OUTLINE - of a
'      structure PRESENT OR VANISHED - the motif occupies.
'      Not an enumerated list of coverages, for two reasons: four
'      segments give fifteen combinations and the corpus already
'      produced four distinct ones; and, decisively, a single
'      value CANNOT DISTINGUISH 'the band did not cover the base'
'      from 'the base segment is not observable', which is
'      exactly what separates an inverted U from a badly
'      preserved closed ring. Inverted U, Surrounding and
'      Flanking are DERIVED in QRY_20, never stored: the same
'      aggregate-judgement pattern as Dec_Present.
'      CONVENTION, corrected from v16: left and right AS SEEN BY
'      AN OBSERVER FACING THE EXPOSED PLANE, never Access_Plane.
'
'  V3. OUTLINE_GEOMETRY (A2). Orthogonal / Curvilinear /
'      Irregular / ND, on every ROC row and nowhere else. NOT
'      gated by decoration type: the three corpus rows
'      describing an inverted U carry THREE DIFFERENT TYPES (RA
'      Amorphous stain, RA Perimeter band, Painted band), so any
'      type filter would exclude precisely them.
'      This is not ornament. An orthogonal outline asserts a
'      RECTANGULAR BUILT REFERENT; a curved one asserts nothing
'      and may follow a natural recess. It is the field that
'      decides which half of the corpus a vanished-structure
'      record belongs to (see V5).
'
'  V4. L_SUPPORT: TWO ENTRIES, ONE NEW MODE (B1).
'      Rock dihedral - re-entrant angle between two rock planes,
'      natural and not cut. Geologically the negative of a wedge
'      block detached along two intersecting discontinuities.
'      Boundary with the cavities is GEOMETRIC, not metric: a
'      cavity is hollowed INTO the wall and has a mouth; a
'      dihedral is open in two directions. Field test: how many
'      walls does the rock save you?
'      Rock surface - support of rock art panels with no
'      associated structure. Support_Mode = Substrate, A NEW
'      MODE, because a surface carrying pigment is neither
'      gravitational rest nor confinement nor embedment. The own
'      mode keeps it out of the constructive-support counts.
'      Written criterion (guidance, NOT a validation rule): the
'      primary support answers VERTICAL forces, the secondary
'      answers HORIZONTAL ones.
'
'  V5. VANISHED STRUCTURE = STRUCTURE RECORD (C1). A structure
'      of which only the perimeter pigment survives is a
'      T_STRUCTURES record with typology Unclassifiable, the A-X
'      vocabulary at 9, a T_LOST_ELEMENTS row with scope Whole
'      structure and evidence Pigment on bedrock, and the
'      ROC-PER decoration row. Registered as rock art it would
'      drop out of the structure counts, and the per-sector
'      counts are what H02 and H03 measure: excluding the
'      vanished ones counts PRESERVATION and reads as
'      constructive density.
'      Boundary: orthogonal or aligned with architectural
'      elements -> structure. Curved or irregular -> rock art.
'      Rule 40 guards both directions.
'
'  V6. CONNECTIONS: THE ORDER STOPS CARRYING MEANING (D1). The
'      corpus held two rows and they were THE SAME CONNECTION,
'      entered from each end. The cause was structural, not
'      careless: F_CONNECTIONS linked on ID_Struct_A only, so
'      from B nothing was visible. Chrono_Relation stops saying
'      'A earlier than B' and NAMES WHICH STRUCTURE IS EARLIER;
'      with the order free of meaning it is normalised (lower ID
'      always in A) and a UNIQUE INDEX makes duplication
'      impossible instead of something to watch for. The v16
'      instruction 'never swap A and B after entry' disappears.
'      Horizontal association joins Vertical association as a
'      connection type.
'
'  V7. ID_MATERIAL_STATUS MOVES TO TAB 7 (D2). The gate
'      (Cultural_Materials_Present) lived on tab 7 and the thing
'      it gates on tab 5, so the condition of movable remains
'      was asked BEFORE whether there were any - and the
'      contradiction was undetectable, because level-2 gating
'      operates within a tab. Tab 5 becomes the built structure;
'      tab 7 becomes its contents, with the same shape as tab 6.
'      The three layers are NOT redundant: presence is
'      answerable from 30 m, type needs identification, and
'      condition needs to see them well. Rules 43-44.
'
'  V8. FRIEZE BOUNDARY (D4). Relief_Frieze (M) is the AGGREGATE
'      claim that there is a frieze; the rows say WHICH MOTIF.
'      Not alternatives - both are recorded. The socle/BAS
'      precedent does not transfer, because Chachapoya friezes
'      ARE the fabric, so 'fabric to the element, motif to the
'      row' would lose the motif repertoire that H01 and H04
'      compare. Rule 42 is unidirectional and checks OVL AND
'      FFL, since M is a treatment of the facade wall face and a
'      frieze may run without crossing the portal.
'      Three types RETIRED from L_DEC_TYPE, all unused in the
'      56 corpus rows: 'Frieze / Greca' (the only entry that
'      names no motif - it says what M already says, and ND
'      covers the unresolved case), 'Triangular motif' (never
'      observed as distinct from the zigzag) and 'Decapitation
'      scene' (a single case, and ROCK ART, not architectural
'      decoration: it belongs to RA Anthropomorphic with the
'      reading in Notes).
'      The remaining architectural repertoire is CONSTRUCTIVE:
'      zigzag, stepped, T, inverted T, L, inverted L and square
'      niches are all built into the fabric and all read as
'      frieze types, painted or not. Rule 42 covers all seven.
'
'  V9. GATING AND THE NON-APPLICABLE (E). Access_Plane and
'      Portal_Orientation are gated BY Sys_Chamber, not by
'      Sys_Portal: gating by the portal would break the v16
'      decision that a razed opening still has a known
'      orientation. Facade_Orientation and Visibility_Valley
'      STAY OPEN - v16 defined the facade as the EXPOSED PLANE,
'      deliberately unhooked from the opening, and a basal mass
'      has an exposed plane and faces somewhere; disabling it
'      would remove H06's natural control group.
'      General criterion: DERIVABLE -> gating; A JUDGEMENT ->
'      the Not applicable value.
'      Sys_Interface is added as a sixth system. It passes the
'      test that rejected a Sys_Facade in v16 - a facade can
'      never be Absent, an interface can. It governs ONLY
'      Interbody_Cornice (I): with one body there is no cornice
'      between bodies, but Upper_Crown (R) still exists, so R
'      stays permanently active.
'
'  V10. CARDINALITY: T_BODIES DEFERRED, WITH A COUNTER FITTED
'      (F). 24 of the 33 records with body counts have more than
'      one body, but Stone_Format and Stone_Working are filled
'      on only 4 of 36 and Mixed has never been used: deciding
'      now would structure data not yet collected. The decisive
'      argument is the UNIT OF ANALYSIS - two bodies of one
'      structure are NOT independent observations, so rows per
'      body would inflate the sample with repetitions and make
'      clustering reflect how many bodies a structure has.
'      THE STRUCTURE REMAINS THE UNIT OF ANALYSIS.
'      Fabric records the
'      phenomenon Mixed was destroying, and the second pass will
'      produce the deciding figure as a BY-PRODUCT. The
'      between-bodies / within-a-body distinction is what makes
'      the figure mean anything: if divergence is mostly WITHIN
'      a body, T_BODIES would solve nothing.
'      Boundary: fabric divergence is an OBSERVATION, a
'      construction phase is an INTERPRETATION. Rules 45-46.
'      Mixed is redefined: it says the record holds more than
'      one value; Fabric says whether that plurality
'      coincides with the division into bodies.
'
'  V11. Portal_Position (inherited open point): Centred /
'      Off-centre left / Off-centre right / NA / ND, gated by
'      Sys_Portal, same observer convention as V2. Off-centring
'      a portal is a planning decision, not an accident of
'      fabric: H01 and H04.
'
'  ============ v16 and earlier ============
'
'  1. STONE FORMAT AND WORKING (delta 1). Masonry_Type recorded HOW
'     the stone is coursed and Masonry_Quality a judgement of
'     execution; neither recorded WHICH STONE it is - the first
'     gesture of the operative sequence. Two ORTHOGONAL fields,
'     because the vocabulary first proposed mixed three axes
'     (natural morphology, degree of working, dimension) and a
'     single list left well-dressed laminar slabs with no possible
'     value. Stone_Format takes an ordered pair like ID_Support,
'     so there is no Mixed value; Stone_Working is ORDINAL
'     (Unworked < Semi-dressed < Dressed) with Mixed and ND
'     off-scale. Scope is the WALL FACE, not the singular elements
'     (lintel, eave surface, rear closure) which have their own
'     material fields.
'
'  2. ACCESS_PLANE AND PORTAL_ORIENTATION (delta 3). One real
'     structure has its opening on the narrow wall perpendicular
'     to the cliff and its decoration on the long wall parallel to
'     it, so 'facade' as exposed plane and 'facade' as compositional
'     plane came apart. FACADE IS FIXED AS THE EXPOSED PLANE:
'     pinning it to the opening would leave Facade_Orientation not
'     pointing at the valley and degrade the whole of H06 to save
'     the definition of one element. Access_Plane records the
'     divergence, which is a variable and not an anomaly:
'     circulation overriding display. Portal_Orientation is NOT a
'     duplicate - the plane is relational and intrinsic, the
'     orientation absolute and for QGIS; neither is derivable from
'     the other. Access_Plane also DECLARES THE REFERENCE PLANE
'     that makes left/right meaningful in T_DECORATIONS.
'
'  3. POSITION_RELATIVE ON T_DECORATIONS (delta 4.1). Body_No
'     indexes the constructive body, which a painting on bedrock
'     does not have. It is NOT replaced by left/right: those are
'     two different axes, and Body_No is still needed on the
'     architectural rows. The new field serves BOTH halves of tab
'     4.Dec and resolves the laterality of FFL, JAM, PRT and RTW
'     with one mechanism instead of duplicated lookup entries.
'
'  4. L_STRUCT_BODY (delta 2 and 4.4). SOC is withdrawn: 'a
'     decoration located on the decorative treatment of the basal
'     mass' is circular, and splitting it by motif type would make
'     the position field encode what ID_Dec_Type already says.
'     RTW, PRT, SIL and LIN are added: v15 had no sill and no
'     lintel, and Over-lintel is the frieze zone (M), not the
'     lintel - a painted portal had to be recorded against an
'     element that was not the painted one. PRT exists because a
'     continuous band framing the opening is ONE gesture, not
'     three. The ROC entries are redefined on PURELY GEOMETRIC
'     contact (overlapping / perimeter / proximate), which removed
'     a real overlap: v15 sent tests 2 AND 3 to ROC and test 3
'     also to ROC-PER, so framing marks matched two entries.
'
'  5. INTERBODY_CORNICE_MATERIAL WITHDRAWN (delta 5.1). Not
'     because cornices happen to be stone but because they COULD
'     NOT be otherwise: a horizontal timber projecting between two
'     bodies is a corbel (E) or a transverse beam (F), so the
'     value 'Wooden beams' described a misclassified platform. T
'     and U already carry ALWAYS STONE in the definition with no
'     material field; the cornice was the incoherent exception.
'     The clause now works as an IDENTIFICATION TEST for G vs I.
'
'  6. X IS A STATE, NOT AN ELEMENT (delta 5.2, 5.3). Defined as
'     'closed above by whatever solution', Chamber_Roof entered
'     the A-X matrix identically whether masonry was built or a
'     rock visor did the work. With many natural-bedrock roofs in
'     the corpus, Jaccard and correspondence analysis were
'     counting GEOLOGY AS CONSTRUCTION. QRY_13 now exports X as
'     NA when the type is Natural bedrock - NA and not 0, because
'     it is not an absence but an unbuilt solution. Definition
'     also tightened: NO CONTACT, NO ROOF. A visor passing metres
'     above the structure closes nothing; that is ID_Support.
'
'  7. DIFFERENTIATION CRITERION, GENERALISED (delta 5.4, 5.5). An
'     A-X element is present when there is a component PHYSICALLY
'     DIFFERENTIATED from the adjoining wall face. An opening left
'     as a plain gap is Sill=0, Jambs=0, Lintel=0 - and Sys_Portal
'     present all the same, because the opening exists. The system
'     records THAT THERE IS AN OPENING; N, O and Q record whether
'     each position was resolved with a distinct element. Those
'     zeros are a first-order result about labour investment
'     (H01, H04): under the opposite criterion every portal would
'     look alike and the variable would vanish.
'
'  8. RULES 32-39 AND A BUG FIX. QRY_16c was CREATED as
'     'QRY_16c_Rules_22_30' but referenced as '..._22_31' by both
'     the drop array and QRY_16_Validation_Check, so the union
'     failed to be created - and MkQuery degrades that failure to
'     a Debug.Print, which nobody reads. The battery had no entry
'     point. Name corrected, union now over four partials, and
'     BuildDB counts the queries it created against the number
'     expected: a silenced error must still leave a trace where
'     someone looks.
'
'  9. NOT IN v16, DELIBERATELY. Per-subbody description (N0, N1.1,
'     N1.2 with derived interlevels) is a question of CARDINALITY,
'     not vocabulary: it would need a T_BODIES table with the A-X
'     elements hanging off each body. Deferred to v17, and the
'     figure that decides it - how many records have bodies with
'     genuinely divergent attributes - does not exist yet.
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

' ================================================================
'  v18 (DELTA_v17_v18, all blocks closed 2026-08-11)
'
'  A1. Recessed_Frame gated by N_Chamber_Bodies=0 (derivable:
'      no chamber body, no facade, no frame). Rule 47.
'  A2. Portal_Orientation closed while Access_Plane='Facade'
'      (one datum, two fields); QRY_07 exports the effective
'      column so R never sees the duplicate.
'  A3. EA-TER closes facade and portal fields (derivable from
'      the typology). Rule 48.
'  A4. Sill_Coincides_Cornice: element sharing - the interbody
'      cornice serving as the sill. Rule 49.
'  A6. 'Mur de retorn' unified in every UI label; stored names
'      untouched.
'  A7. Rear of the chamber DEFINED AGAINST THE FACADE PLANE,
'      never against the access: an access on the return wall
'      must not rotate Rear_Closure_Type by 90 degrees.
'  B1. 'Tabular blocks' -> 'Regular tabular blocks' (stored
'      value; migration in PATCH v18).
'  B2. 'Large blocks' added to the stone format lists.
'  B3. Interbody_Cornice_Format: laminar against tabular stone
'      in the cornice, hanging from I. Rule 50.
'  C1. Platform_Surface (Z): the platform gets its surface,
'      exact parallel of T on the eave. Joins the element map
'      as the 21st five-value field, gated by Sys_Platform;
'      Platform_Surface_Material now hangs from it. Rule 51.
'      Y stays reserved for the abutted mass (banqueta).
'  C2. Tie_Walls (D) UNGATED from Sys_Base: it appears detached
'      from the basal mass - over corbelled platforms, as an
'      abutting support, or alone. Like I and R it
'      is permanently active; Sys_Base keeps A, B, C.
'  C3. Connection_Type + 'Associated natural context': the
'      terrace-plus-niche case as two records, one edge.
'  D1. Chrono_Relation -> Sequential / Contemporary /
'      Undetermined, and ID_Earlier NAMES the earlier structure
'      by FK. The direction leaves the A/B positions for good;
'      the incoming subform stops inverting. Rule 52.
'  D2. T_LOST_ELEMENTS.ID_Position removed (1 filled row in
'      27); Evidence_Scope kept under a hard coherence rule
'      with Element_Code (rule 53, mirror of 21).
'  E1. ID_Support filtered by typology: NIX autofills its 1:1
'      value, CAV keeps the medium/large choice, PR is limited
'      to the rock surface. Rule 54.
'  F.  Fixes found in review: QRY_18 still read E.Notes (gone
'      in v17a - the query raised a parameter prompt);
'      Sys_Interface never exported from QRY_13; QRY_13 gated
'      Recessed_Frame by Sys_Portal against the v17 decision.
'
'  WHAT CHANGED IN v19 (delta v18->v19)
'
'  A1. Jamb_Fabric_Reveal: second POSITION-RESOLUTION
'      qualifier - the lateral edge of the opening resolved
'      by a dressed terminal face of the fabric (masonry
'      reveal). The class is NOT element sharing: the cornice
'      case shares one existing piece across two positions,
'      this one has no element at all. What they share is the
'      question - with the element at 0, how was the position
'      resolved? Window Jambs IN (0, 2); padding 0 outside;
'      DOM3Q labels in the form, where the 9 reads
'      'Indeterminat': inside the window the not-examinable
'      cannot exist, so the 9 can only mean assessed and not
'      decidable. Rule 55.
'  A2. Rule 3, portal branch, exempt while the reveal is
'      declared (AND Jamb_Fabric_Reveal<>1): the all-zero
'      portal is the legitimate node-2 case of the manual.
'      The equivalent platform hole stays OPEN on purpose
'      (v18 delta A5): a known, accepted false positive.
'  B1. Rule 9 fires on Is Null ONLY: ND is a judgement, not
'      an empty cell. Rule 56 guards the coherence the old
'      condition pretended to (support role under a denied
'      platform; the reverse direction was already R7).
'  B2. Timber_Bracket_Role UI labels renamed around the one
'      question the field answers (the link with H); stored
'      values untouched.
'  C1. MEN promoted: 'MEN Isolated Structural Element',
'      description anchored to the A-Z vocabulary; no
'      subtypes (the element fields already say which).
'  C2. L_ELEMENTS row D: Name_EN 'Transverse wall' ('tie'
'      asserts the function, and the function is what we do
'      not know); anchoring freedom said with HEIGHT, never
'      LEVEL. The field name Tie_Walls is NOT touched.
'  C6. The v17 padding zeros of D go back to NULL (PATCH
'      v19): a judgement nobody made must not read as one.
'
'  WHAT CHANGED IN v20 (delta v19->v20) - 15 closed blocks
'
'  1.  D widened: transverse wall OR PIER (label already v19a);
'      delimitation tests against E (hangs/rises) and K (wall
'      plane / against the rock). The isolated-K exemption
'      candidate died: an isolated pier is D -> MEN, live today.
'  2.  Interbody_Cornice_Format DROPPED; rule 50 retired.
'  3.  Two-question criterion for N0-only tops (finish -> R;
'      use surface -> platform + Z); R generalised with its two
'      faces written; G = progression (>=2 courses); the
'      crown-by-default escape is outlawed in the manual.
'  4.  Natural niches with perimeter paint: NIX + ROC-PER rows;
'      glossary splits natural niche from decorative niche.
'  A.  Two families of zero: DOM5 labels say 'constatat' and
'      'atestat'; vestige tabs carry the epistemic banner; the
'      wood elements (A, E, F, S) stay in the strong family
'      because embedded wood leaves sockets.
'  B.  Code identity: UNIQUE INDEX; Sector_Code in L_SECTORS;
'      rule 57 (concordance); EA number as query expression;
'      QRY_26_Next_EA (max+1, gaps are history).
'  C.  Upper_Crown_Format: closing panel / projecting course /
'      flush course / ND. Rule 58. Which face it crowns is
'      derived, never declared.
'  D.  Rule 6 action names the third exit (confirm legibility);
'      no blanket 0->9 conversion - per-surface judgement
'      (model record: DW-S01-EA13).
'  E.  Portal-adjacent fields gated by the chamber body count
'      (rule 59); masonry deliberately NOT gated by walls (13
'      wall-less records carry real fabric).
'  F.  Outline bands: verdict-free types (defined/amorphous);
'      8bis.9 demotes geometry from criterion to clue; topology
'      derives from spans (QRY_27).
'  G.  No phase default; Phase_Evidence gated by the count;
'      rule 60.
'  H.  Rock-art class: position autofilled to Panel; interior
'      observability closed; Body_No hidden in datasheet too.
'  I.  T_ARCH_FEATURES.Present dropped; catalogue gains Socket
'      and Access bench (the counting instrument of element Y).
'  J.  Structural trace opens 2.Arq (a wall has fabric).
'  K.  No field colouring (it would duplicate the battery);
'      the Validate-this-record button reuses it instead.
'
'  WHAT CHANGED IN v21 (delta v20->v21) - 6 closed blocks
'
'  1.  TERRACE CRITERION REVISED (reopens v20 bloc 3 at the
'      designer's explicit request): Z is RESERVED for the
'      surface of the CANTILEVERED system (slabs on corbels
'      or on a transverse beam). The top of the basal mass
'      of a terrace is NEVER Z and does not open
'      Sys_Platform: a terrace already answers the use
'      question by its own typology. Only one real question
'      remains on an N0-only top: the constructed finish
'      (R). Rule 62 guards the boundary: a platform whose
'      three supports are all CONFIRMED absent is the top
'      of the mass, not a platform. Nines do not fire it:
'      the v18 C1 case (surface with unresolvable supports)
'      is preserved. The 'Terrace two questions' branch of
'      QRY_25 is retired; Z-only rows go to QRY_29 for a
'      row-by-row revert (no blind UPDATE).
'  2.  Masonry_Present: the fabric declaration. The class
'      cannot derive it (an isolated corbel MEN has no
'      fabric; a pier MEN does - DW-S01-EA55 is the field
'      counterexample), so under the v17 criterion the
'      judgement gets a field that gates the masonry block
'      (quality, formats, working, bond, mortar, chinking,
'      fabric divergence). 0/1/9, NULL default, joins the
'      layer family (rule 17 lists it). Rule 61 mirrors the
'      plaster/pigment presence rules. Plaster and pigment
'      are deliberately NOT chained to it. The v17 note
'      still holds: Fabric is not gated BY BODIES - its
'      gate is now the fabric declaration itself.
'  3.  Outline band combo FIX (v20 erratum): the v20
'      reasoning ('a pure panel outlines nothing') confused
'      the panel CLASS with the whole ROC half of the
'      subform filter. F_ROCKART serves ALL ROC rows, and
'      ROC-PER - exactly where outline bands live - could
'      not reach them. Form-side fix: rock combo gains the
'      two bands, architectural combo excludes them, and
'      the panel class restricts its list to RA* by the
'      bloc-H per-class pattern. No stored value changes.
'  4.  12.Extra catalogue FIX (v20 erratum): bloc I decided
'      Socket / negative interface and Access bench, but
'      the form value list was never updated - the two
'      entries were unreachable and the retype worklist
'      could not be worked. Form-side fix only.
'  5.  '(opc.)' suffix on the ALWAYS-OPTIONAL fields
'      (secondary support, secondary format, secondary
'      colour): their NULL is a legitimate final state in
'      EVERY record state, so a static label cue does not
'      lie - unlike the per-field colouring bloc K
'      rejected, whose subject was state-dependent.
'  6.  Worklist residue DOCUMENTED: the QRY_24 fabric-reveal
'      and QRY_23 cornice-as-sill branches fire on the 0
'      and a confirmed 0 is indistinguishable from the
'      migration 0. Once the pass is done, surviving rows
'      are expected residue, not pending work (the rule-6
'      deliberate-friction precedent). No query change.
'
'  WHAT CHANGED IN v22 (delta v21->v22) - 5 closed blocks
'
'  A.  MASONRY GATE, THE OTHER HALF (fixes the v21 friction
'      Esteve found on the populated DB: 23 of 27 pending
'      masonry judgements sat on classes whose 2.Arq tab
'      was CLOSED). Natural funerary contexts OPEN 2.Arq -
'      a cavity can hold real fabric and basal bodies (the
'      field photo: beams and a slab floor inside a niche)
'      - with Masonry_Present as the gatekeeper: 0 or 9
'      closes the whole block in seconds, 1 opens the
'      detail. Rock art panels stay closed and get the
'      DERIVABLE 0 (a pure panel has no fabric by
'      definition): one-time UPDATE in the patch, and the
'      typology AfterUpdate fills it for future panels -
'      never touching a declared value.
'  B.  Metrics_Available: the metric-availability
'      gatekeeper on 9.Metr (designer's request). 0/1/9,
'      NULL default; gates the dimensional and volumetric
'      fields. COORDINATES AND NOTES STAY OUT: a structure
'      with no measurable dimensions still has a position.
'      Deliberately OUTSIDE the layer array, unlike
'      Masonry_Present: availability is not decidable
'      record-by-record until the extraction workflow
'      runs, and rule 17 would flood the worklist with an
'      unanswerable question. Rule 63 guards the detail
'      side; the derivation in the patch covers the rest.
'  C.  C14-ONLY DATING (explicit reversal of the documented
'      decision that kept the century fields always
'      editable for typological attribution): the project
'      will only date by radiocarbon, so the century
'      fields gate behind the C14 flag. Rule 64 guards the
'      datasheet side. Zero data cost: no record carried
'      centuries or C14 at migration time.
'  D.  The structure code joins the header strip (outside
'      the tab control, visible from every tab); the
'      static title finally stops saying v11.
'  E.  LEGACY QUERY CLEANUP (fixes the v21 defect: the
'      rename of the seventh rules partial left
'      QRY_16g_Rules_57_60 behind as an inert orphan,
'      because the delete loop only knows current names).
'      Known legacy names are now deleted explicitly.
'
'  RebuildQueriesV22() below rebuilds QUERIES ONLY: it is the
'  step the PATCH v22 instructions call for on a populated
'  database, where BuildDB() must never run (it drops tables).
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
    ' v16: a silenced MkQuery error must still leave a trace
    ' where someone looks. See header note 8.
    ReportQueryCount db
    db.TableDefs.Refresh
    db.QueryDefs.Refresh
    Set db = Nothing

    Dim msg As String
    msg = "DATABASE v22 BUILT SUCCESSFULLY!" & vbCrLf & vbCrLf
    msg = msg & "  22 tables | 25 relationships | 40 queries" & vbCrLf
    msg = msg & "  T_STRUCTURES: 144 fields" & vbCrLf
    msg = msg & "  48 observational BYTE fields" & vbCrLf & vbCrLf & vbCrLf
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
    msg = msg & "  v14: C14 gates the dating subform | rule 31" & vbCrLf
    msg = msg & "  v16: Stone_Format + Stone_Working (masonry)" & vbCrLf
    msg = msg & "  v16: Access_Plane + Portal_Orientation" & vbCrLf
    msg = msg & "  v16: Position_Relative on T_DECORATIONS" & vbCrLf
    msg = msg & "  v16: L_STRUCT_BODY 18 entries (-SOC, +RTW PRT" & vbCrLf
    msg = msg & "       SIL LIN ROC-OVL); ROC redefined by contact" & vbCrLf
    msg = msg & "  v16: Interbody_Cornice_Material withdrawn" & vbCrLf
    msg = msg & "  v16: X exports NA when the roof is natural rock" & vbCrLf
    msg = msg & "  v16: rules 32-39 | QRY_16d | QRY_19 review list" & vbCrLf
    msg = msg & "  v16a: L_LOST_EVIDENCE.Name_VAL added - the evidence" & vbCrLf
    msg = msg & "        dropdown was empty and value 3 unrecordable" & vbCrLf
    msg = msg & "  v16: QRY_16c name fixed - the battery union was" & vbCrLf
    msg = msg & "       silently failing to be created in v15" & vbCrLf
    msg = msg & "  v17: Position_Relative removed; four Span_*" & vbCrLf
    msg = msg & "       Span fields on T_DECORATIONS" & vbCrLf
    msg = msg & "  v17: Sys_Interface (sixth system, governs I only)" & vbCrLf
    msg = msg & "  v17: Fabric (single / between bodies / within):" & vbCrLf
    msg = msg & "       T_BODIES deferred WITH A COUNTER FITTED" & vbCrLf
    msg = msg & "  v17: Portal_Position | Horizontal association" & vbCrLf
    msg = msg & "  v17: connection order normalised + unique index" & vbCrLf
    msg = msg & "  v17: L_SUPPORT +Rock dihedral +Rock surface" & vbCrLf
    msg = msg & "  v17: L_DEC_TYPE 16 entries: Frieze/Greca," & vbCrLf
    msg = msg & "       Triangular motif and Decapitation scene" & vbCrLf
    msg = msg & "       retired (all unused; the last is rock art)" & vbCrLf
    msg = msg & "  v17a: notes on every tab; Notes -> Doc_Notes" & vbCrLf
    msg = msg & "  v17a: rock art types rebuilt on one axis;" & vbCrLf
    msg = msg & "        Outline_Geometry withdrawn into the list" & vbCrLf
    msg = msg & "  v17: rules 40-46 less 41 | QRY_20 | QRY_21" & vbCrLf
    msg = msg & "  v18: Platform_Surface (Z) joins the map: 21" & vbCrLf
    msg = msg & "       five-value elements, gated by Sys_Platform" & vbCrLf
    msg = msg & "  v18: Tie_Walls (D) ungated from Sys_Base" & vbCrLf
    msg = msg & "  v18: Interbody_Cornice_Format under I;" & vbCrLf
    msg = msg & "       Sill_Coincides_Cornice (cornice as sill)" & vbCrLf
    msg = msg & "  v18: Recessed_Frame gated by the body count;" & vbCrLf
    msg = msg & "       EA-TER closes facade and portal fields" & vbCrLf
    msg = msg & "  v18: Chrono_Relation 3 values + ID_Earlier FK" & vbCrLf
    msg = msg & "  v18: T_LOST_ELEMENTS.ID_Position removed" & vbCrLf
    msg = msg & "  v18: Regular tabular / Large blocks; support" & vbCrLf
    msg = msg & "       filtered by typology (NIX 1:1 autofill)" & vbCrLf
    msg = msg & "  v18: rules 47-54 | QRY_16f | QRY_23 review" & vbCrLf
    msg = msg & "  v19: reveal (A1) | R3 exempt | R9 Is Null;" & vbCrLf
    msg = msg & "       rules 55-56 | MEN promoted | QRY_24" & vbCrLf
    msg = msg & "  v18 FIX: QRY_18 read E.Notes (gone in v17a);" & vbCrLf
    msg = msg & "       QRY_13 now exports Sys_Interface" & vbCrLf
    msg = msg & "  v20: D widened to pier | Upper_Crown_Format" & vbCrLf
    msg = msg & "       | outline bands verdict-free | rules" & vbCrLf
    msg = msg & "       57-60 | Code unique | QRY_25/26/27" & vbCrLf
    msg = msg & "  v21: terrace criterion revised (Z reserved" & vbCrLf
    msg = msg & "       for the cantilevered surface, rule 62)" & vbCrLf
    msg = msg & "  v21: Masonry_Present gates the masonry" & vbCrLf
    msg = msg & "       block (rule 61) | QRY_29 review" & vbCrLf
    msg = msg & "  v22: naturals open 2.Arq | panels derive" & vbCrLf
    msg = msg & "       Masonry 0 | Metrics_Available gate" & vbCrLf
    msg = msg & "       (rule 63) | C14-only dating (rule" & vbCrLf
    msg = msg & "       64) | header code | legacy cleanup" & vbCrLf & vbCrLf
    msg = msg & "TWO DEFAULTS, DELIBERATELY:" & vbCrLf
    msg = msg & "  0 on element fields governed by a Sys_* (rule 4" & vbCrLf
    msg = msg & "    needs the padding zero)" & vbCrLf
    msg = msg & "  NULL everywhere else: empty = not yet assessed," & vbCrLf
    msg = msg & "    9 = assessed and not examinable, 0 = assessed" & vbCrLf
    msg = msg & "    and absent. Rule 17 lists what is still NULL." & vbCrLf & vbCrLf
    msg = msg & "Next: run chachapoya_Form_v22_val.bas -> BuildForm()"
    MsgBox msg, vbInformation, "Done!"
End Sub

' Queries only, tables untouched: safe on a database that already
' holds records. This is step 2 of the PATCH v18 sequence.
Public Sub RebuildQueriesV22()
    Dim db As DAO.Database
    Set db = CurrentDb()
    CreateAllQueries db
    ReportQueryCount db
    db.QueryDefs.Refresh
    Set db = Nothing
    MsgBox "Queries rebuilt against the v22 schema." & vbCrLf & "Tables and data untouched.", vbInformation, "RebuildQueriesV22"
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
' v18 (delta C1): the map grows to 21 entries with Z, and (C2)
' D loses its gate. Extending this array automatically extends
' the defaults, the long-format unpivot, rules 1-2-4-6 and the
' null worklist: one list, one truth.
Private Sub FillElementMap(m() As String)
    ReDim m(20, 2)
    m(0, 0) = "A":  m(0, 1) = "Embedded_Base_Beams":  m(0, 2) = "Sys_Base"
    m(1, 0) = "B":  m(1, 1) = "Base_Level":           m(1, 2) = "Sys_Base"
    m(2, 0) = "C":  m(2, 1) = "Decorative_Socle":     m(2, 2) = "Sys_Base"
    ' v18 (delta C2): D ungated - anchors at any level, not only
    ' inside the basal mass. Empty gate, like R.
    m(3, 0) = "D":  m(3, 1) = "Tie_Walls":            m(3, 2) = ""
    m(4, 0) = "E":  m(4, 1) = "Timber_Brackets":      m(4, 2) = "Sys_Platform"
    m(5, 0) = "F":  m(5, 1) = "Transverse_Beams":     m(5, 2) = "Sys_Platform"
    m(6, 0) = "G":  m(6, 1) = "Corbelled_Courses":    m(6, 2) = "Sys_Platform"
    m(7, 0) = "I":  m(7, 1) = "Interbody_Cornice":    m(7, 2) = "Sys_Interface"
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
    m(20, 0) = "Z": m(20, 1) = "Platform_Surface":    m(20, 2) = "Sys_Platform"
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
    ReDim f(8)
    f(0) = "Mortar_Present"
    f(1) = "Plaster_Present"
    f(2) = "Pigment_Present"
    f(3) = "Dec_Present"
    f(4) = "RockArt_Present"
    f(5) = "Chinking_Stones"
    f(6) = "Recessed_Frame"
    f(7) = "Cultural_Materials_Present"
    ' v21 (bloc 2): the fabric declaration joins the family -
    ' a genuine per-record judgement, so rule 17 must list it.
    f(8) = "Masonry_Present"
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
        db.Execute "CREATE TABLE L_SECTORS (ID COUNTER CONSTRAINT PK_SEC PRIMARY KEY, ID_Site LONG NOT NULL, Sector_Name TEXT(50) NOT NULL, Sector_Code TEXT(10), Description TEXT(255))", dbFailOnError
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
        ' v16a BUG FIX: Name_VAL was missing here while the form's
        ' combo selects it, so the RowSource query failed and the
        ' evidence-type dropdown came up EMPTY - which makes value 3
        ' unrecordable, since every 3 needs its evidence row. The
        ' nine VL calls failed silently too (VL tolerates errors by
        ' design). Only this lookup was affected; the other ten were
        ' checked.
        db.Execute "CREATE TABLE L_LOST_EVIDENCE (ID COUNTER CONSTRAINT PK_LEV PRIMARY KEY, Name TEXT(60) NOT NULL, Name_VAL TEXT(60), Description TEXT(150))", dbFailOnError
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
        ' v22 (bloc B): THE METRIC-AVAILABILITY GATEKEEPER.
        ' 0/1/9, NULL default. Gates the dimensional and
        ' volumetric fields of 9.Metr (and the two opening
        ' dims, and the support metrics, and the base-height
        ' cota); coordinates and notes stay out - a structure
        ' with nothing measurable still has a position. OUT
        ' of the layer array on purpose (see SetByteDefaults).
        sql = sql & "Metrics_Available BYTE,"
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
        ' --- 2c. Masonry & mortar (10) | H01/H04 ---
        ' v21 (bloc 2): THE FABRIC DECLARATION. Presence of
        ' masonry is a judgement the class cannot derive - an
        ' isolated corbel MEN has no fabric, a pier MEN does
        ' (DW-S01-EA55) - so under the v17 criterion it gets a
        ' field, family of Plaster_Present / Pigment_Present.
        ' 0/1/9, NULL default (three epistemic states), joins
        ' the layer array so rule 17 lists it. Gates the whole
        ' block below, Fabric included; rule 61 guards the
        ' datasheet side. Plaster and pigment stay independent.
        sql = sql & "Masonry_Present BYTE,"
        sql = sql & "Masonry_Quality TEXT(20),"
        ' v16 (delta 1): WHICH STONE, as against Masonry_Type
        ' (how it is coursed). Two orthogonal fields because the
        ' single list mixed morphology, working and dimension.
        ' Format (v18, deltas B1-B2): Irregular stones /
        ' Regular tabular blocks / Large blocks / Laminar
        ' slabs / ND. 'Regular' entered the stored value when
        ' Large blocks arrived: a large block is ALSO tabular,
        ' so the bare term stopped naming one class. Large =
        ' one piece is one course on its own, the labour
        ' signal the functional test exists to read.
        ' Primary test is FUNCTIONAL - how many
        ' pieces make one course - because it reads straight off
        ' the facade and needs no measuring. Ordered pair, so
        ' there is NO Mixed value: the project already retired
        ' Combined (L_SUPPORT, v8) and Both (Color, v13) for the
        ' same reason. Dominance = greater wall AREA, not more
        ' pieces: small laminar slabs and large tabular blocks
        ' invert the answer if counted by piece.
        sql = sql & "Stone_Format TEXT(30),"
        sql = sql & "Stone_Format_Secondary TEXT(30),"
        ' ORDINAL: Unworked < Semi-dressed < Dressed, usable as a
        ' proxy for labour investment. Mixed and ND are OFF-SCALE
        ' and drop out of those analyses, like Not applicable.
        ' Records POSITIVE EVIDENCE of working (tool traces,
        ' regular arrises, faces breaking the natural fracture
        ' plane); with no such traces the flat face is credited
        ' to the fracture and the value is Unworked. ND is for
        ' the genuine doubt - a bedded sandstone fractures into
        ' faces indistinguishable from dressing. NO ninth value:
        ' 9 exists to protect the 0, and these fields have no 0
        ' to protect. Doc_Basis and Facade_Observability already
        ' separate 'could not see' from 'not decidable'.
        sql = sql & "Stone_Working TEXT(20),"
        sql = sql & "Masonry_Type TEXT(30),"
        sql = sql & "Mortar_Present BYTE,"
        sql = sql & "Mortar_Type TEXT(30),"
        sql = sql & "Chinking_Stones BYTE,"
        ' v17 (delta F2, revised): THE COUNTER FITTED IN PLACE
        ' OF T_BODIES. Mixed on Masonry_Type and Stone_Working
        ' says the record holds more than one value; this says
        ' whether that plurality FOLLOWS THE DIVISION INTO BODIES
        ' or cuts across it.
        ' Single / Between bodies only / Within a body /
        ' Not observable. The three are MUTUALLY EXCLUSIVE because
        ' the criterion is not WHERE the divergence sits - a body
        ' that is internally plural is also plural against its
        ' neighbour, so 'where' overlaps - but whether the
        ' division COINCIDES with the bodies.
        ' That distinction is the whole point: 'Between bodies
        ' only' is exactly the case a T_BODIES table would
        ' resolve, and 'Within a body' is the case it would NOT,
        ' because the same Mixed would reappear one level down.
        ' BOUNDARY WITH Construction_Phases: divergence is an
        ' OBSERVATION (the stone changes, and that is visible);
        ' a phase is an INTERPRETATION (there were two moments
        ' of building, and that is argued). There can be a
        ' change of fabric with no phase claimed - a change of
        ' supply, or two masons the same day.
        sql = sql & "Fabric TEXT(25),"
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
        ' v17 (delta E3): SIXTH SYSTEM. Passes the test that
        ' rejected a Sys_Facade in v16 - a facade can never be
        ' Absent, so a system was the wrong shape for it; an
        ' interface CAN be absent, because a single-bodied
        ' structure has none, and that absence is real and
        ' checkable. Governs ONLY Interbody_Cornice (I):
        ' Upper_Crown (R) stays permanently active, because a
        ' one-body structure still has a top.
        sql = sql & "Sys_Interface TEXT(20),"
        ' Form separated from function: naming the platform "access"
        ' in the data would presume the conclusion OE3 exists to reach
        sql = sql & "Platform_Function TEXT(20),"
        ' --- 3b. Level N0 elements (10) | A-G ---
        '     BYTE five-value: 0=Absent 1=Complete 2=Partial
        '     3=Attested lost 9=Not observable (default 0 via DAO)
        sql = sql & "Embedded_Base_Beams BYTE,"
        sql = sql & "Base_Level BYTE,"
        sql = sql & "Decorative_Socle BYTE,"
        ' v18 (delta C2): D UNGATED. It appears detached from the
        ' basal mass - over corbelled platforms, as an abutting
        ' support, or alone - so Sys_Base is no
        ' longer its gate: like I and R it is permanently active
        ' and its 0 is always an assertion, never padding.
        sql = sql & "Tie_Walls BYTE,"
        sql = sql & "Timber_Brackets BYTE,"
        sql = sql & "Timber_Bracket_Count INTEGER,"
        sql = sql & "Timber_Bracket_Role TEXT(25),"
        sql = sql & "Transverse_Beams BYTE,"
        sql = sql & "Corbelled_Courses BYTE,"
        ' v18 (delta C1): THE PLATFORM GETS ITS SURFACE. Z is the
        ' finished walking surface of the corbelled platform, the
        ' exact parallel of T on the eave, with the full five-value
        ' domain and gated by Sys_Platform like F and G. The
        ' material keeps its own field below and now hangs from Z:
        ' no surface, no material (rule 51).
        sql = sql & "Platform_Surface BYTE,"
        sql = sql & "Platform_Surface_Material TEXT(20),"
        ' --- 3c. N0/N1 interface (1) | element I, gated by nothing ---
        ' v16 (delta 5.1): Interbody_Cornice_Material is GONE. Not
        ' because cornices happen to be stone but because they
        ' could not be otherwise - a projecting timber between two
        ' bodies is E or F, so 'Wooden beams' described a
        ' misclassified platform, never a timber cornice. T and U
        ' already carry ALWAYS STONE in the definition with no
        ' material field of their own. The clause is now worth
        ' more as an identification test for G vs I (rule 39).
        sql = sql & "Interbody_Cornice BYTE,"
        ' v18 (delta B3): stone format of the cornice - laminar
        ' (thin slabs) against tabular (thick blocks). Follows the
        ' Lintel_Material pattern: only while I has entity (rule 50).
        ' --- 4. Level N1 elements (12) | J-Q, R, V ---
        sql = sql & "Corner_Quoins BYTE,"
        sql = sql & "Structural_Pilasters BYTE,"
        sql = sql & "Facade_Flank BYTE,"
        sql = sql & "Relief_Frieze BYTE,"
        sql = sql & "Sill BYTE,"
        sql = sql & "Jambs BYTE,"
        sql = sql & "Lintel BYTE,"
        sql = sql & "Lintel_Material TEXT(20),"
        ' v18 (delta A4): ELEMENT SHARING, the case the gradient
        ' cannot say. Where the interbody cornice itself serves as
        ' the sill, N is honestly 0 (no differentiated sill) and
        ' yet the position is resolved. 1 = the cornice does the
        ' job; 0 = it does not (and the padding 0 outside the
        ' window); 9 = cannot be seen. Meaningful only while
        ' Sill=0 and I is present: rule 49 watches the window and
        ' the form gates it.
        sql = sql & "Sill_Coincides_Cornice BYTE,"
        ' v19 (delta A1): SECOND POSITION-RESOLUTION QUALIFIER.
        ' The lateral edge of the opening resolved by the fabric
        ' itself - a finished, deliberate terminal face (masonry
        ' reveal), NO element at all: nothing is shared here,
        ' unlike the cornice-as-sill case above, where one real
        ' piece does two jobs. Meaningful only while Jambs is
        ' 0 or 2 (with O=2, one jamb one reveal - the case the
        ' manual declared undecidable at node 1); with O=1 there
        ' is no position left to resolve another way. Rule 55
        ' watches the window and the form gates it; padding 0
        ' outside, exactly like the cornice-as-sill field above.
        sql = sql & "Jamb_Fabric_Reveal BYTE,"
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
        ' v20 (bloc C): the FORM of the crown - closing panel
        ' (the cave-mouth wall above the eave, the founding
        ' case), projecting course (the terrace slabs), flush
        ' course, or ND. Default NULL (a judgement field);
        ' visible with R present or attested; rule 58 watches
        ' the two present values. Which face of R it crowns is
        ' DERIVED from the record (counters and systems), so
        ' that is deliberately NOT a value here.
        sql = sql & "Upper_Crown_Format TEXT(25),"
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
        ' --- 6b. Landscape & orientation (4) ---
        ' FACADE = THE EXPOSED PLANE (v16, delta 3.2). One real
        ' structure has its opening on the narrow wall
        ' perpendicular to the cliff and its decoration on the
        ' long parallel wall. Pinning the facade to the opening
        ' would leave this field not pointing at the valley and
        ' Visibility_Valley measuring a plane nobody sees: the
        ' whole of H06 degraded to save one element definition.
        sql = sql & "Facade_Orientation TEXT(5),"
        sql = sql & "Visibility_Valley TEXT(10),"
        ' Facade / Return wall / Rear / ND. The divergence is a
        ' VARIABLE, not an anomaly to absorb: it says circulation
        ' overrode display - entry where one can walk, display
        ' towards where one looks.
        ' v17: THE SECOND JOB IS GONE. In v16 this field also
        ' declared the reference plane for Position_Relative;
        ' that field is removed (A1), and the Span_* fields that
        ' replace it read from the EXPOSED plane, never from
        ' this one. Access_Plane keeps only its first job.
        ' v17 gating: this field and Portal_Orientation are
        ' disabled when Sys_Chamber is declared Absent or Not
        ' applicable - NEVER on NULL, which means 'not yet
        ' assessed'. Gated by the CHAMBER and not by the portal,
        ' so that a razed opening in a real chamber keeps both
        ' fields open, as v16 intended.
        sql = sql & "Access_Plane TEXT(20),"
        ' NOT a duplicate of Access_Plane: the plane is
        ' relational and intrinsic (comparable across different
        ' absolute orientations), the orientation is absolute and
        ' goes to QGIS against topography and circulation.
        ' Neither is derivable from the other without knowing the
        ' geometry of the case. NOT gated by Sys_Portal, for the
        ' same reason Recessed_Frame is not: a razed opening has
        ' a known orientation, and an Attested lost portal with a
        ' recorded orientation is precisely an informative case.
        ' The DIVERGENCE between the two fields is now
        ' computable, which is what makes it analysable.
        sql = sql & "Portal_Orientation TEXT(5),"
        ' v17: inherited open point closed. Centred /
        ' Off-centre left / Off-centre right / NA / ND, gated
        ' by Sys_Portal. Same observer convention as the Span_*
        ' fields: FROM IN FRONT OF THE EXPOSED PLANE.
        ' Off-centring a portal is a planning decision, not an
        ' accident of fabric: H01 and H04.
        sql = sql & "Portal_Position TEXT(20),"
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
        ' v17: vocabulary widened. The v16 list (C14 / Stratigraphy
        ' / Superposition / Mortar / ND) had no value for the two
        ' commonest field observations - an abutted vertical joint
        ' and a change of fabric - so a phase could be claimed
        ' with nothing recordable behind it. See rule 46.
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
        ' v17a: SET CANONIC DE NOTES, UNA PER PESTANYA. Cinc
        ' pestanyes no en tenien - identificacio, decoracio,
        ' cronologia, metrica i extra -, de manera que el que no
        ' cabia en un camp d'eixes pestanyes no tenia on anar.
        ' QRY_18_Notes_Review existeix per a convertir text lliure
        ' repetit en camps nous, i nomes ho pot fer alla on el
        ' text lliure es possible: Outline_Geometry va naixer aixi,
        ' d'una paraula escrita a ma en un camp de notes.
        ' Notes passa a dir-se Doc_Notes perque les set seguisquen
        ' un sol patro en lloc de sis mes una excepcio.
        sql = sql & "Id_Notes MEMO,"
        sql = sql & "Dec_Notes MEMO,"
        sql = sql & "Chrono_Notes MEMO,"
        sql = sql & "Metric_Notes MEMO,"
        sql = sql & "Extra_Notes MEMO,"
        sql = sql & "Doc_Notes MEMO)"
        db.Execute sql, dbFailOnError
        ' v20 (bloc B): identity blocks, it does not warn. A
        ' duplicate code is never a legitimate transitional
        ' state, so the engine enforces it, not the battery.
        db.Execute "CREATE UNIQUE INDEX idx_Code_Unique ON T_STRUCTURES (Code)", dbFailOnError
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
    ' v17 (delta A1): POSITION_RELATIVE IS GONE. It served two
    ' halves of the table that do not share a referent - see the
    ' V1 header note. The architectural half gets NO replacement
    ' (no laterality in 56 rows; QUO and PIL do not need it, which
    ' also disposes of the multiple-pilaster problem before it
    ' arises), and the ROC half gets four SEGMENT fields.
    '
    ' Span_Left / Span_Above / Span_Right / Span_Below, 0/1/9,
    ' NULL by default: WHICH SEGMENTS OF THE OUTLINE - of a
    ' structure PRESENT OR VANISHED - this motif occupies.
    ' Four fields and not one enumerated value because a single
    ' value cannot distinguish 'the band did not cover the base'
    ' from 'the base segment is not observable', and that is
    ' precisely what separates an inverted U from a badly
    ' preserved closed ring. The corpus has cases where the upper
    ' segment is masked by the rock overhang and by the body
    ' above: marking it Absent would assert something nobody
    ' could look at.
    ' Inverted U, Surrounding and Flanking are DERIVED in QRY_20,
    ' never stored - the aggregate-judgement pattern again, with
    ' the advantage that an inverted U with a 9 on the base shows
    ' up as a CANDIDATE for Surrounding instead of being asserted.
    ' CONVENTION, corrected from v16: LEFT AND RIGHT AS SEEN BY AN
    ' OBSERVER FACING THE EXPOSED PLANE, never Access_Plane - v16
    ' inverted them on every structure with a diverted entrance.
    ' Enabled on ROC-OVL, ROC-PER and ROC. NOT on ROC-PAN, which
    ' hangs off a PR record and by definition has no structure of
    ' reference, nor on the architectural half.
    '
    ' v17a: OUTLINE_GEOMETRY WITHDRAWN. It only ever
    ' discriminated within the perimeter band, so as a general
    ' field it sat empty on most rows and asked a question that
    ' had no meaning on them. The distinction it carried now
    ' lives in the type list, as 'Outline band, defined' against
    ' 'Outline band, amorphous' (v20 names). Same information, one field fewer,
    ' and it can no longer contradict the type.
    ' v20 (bloc F) REWRITES the claim: geometry is a CLUE, not
    ' the criterion - at La Petaca structures are outlined with
    ' C and even O bands. The structural verdict is the full
    ' attestation judgement at record level (typology +
    ' T_LOST_ELEMENTS), never a property of the band; the type
    ' says only the stroke, and the topology derives from the
    ' spans (QRY_27).
    sql = "CREATE TABLE T_DECORATIONS ("
    sql = sql & "ID COUNTER CONSTRAINT PK_TDEC PRIMARY KEY,"
    sql = sql & "ID_Structure LONG NOT NULL,"
    sql = sql & "ID_Struct_Body LONG,"
    sql = sql & "ID_Dec_Type LONG,"
    sql = sql & "Body_No INTEGER,"
    sql = sql & "Span_Left BYTE,"
    sql = sql & "Span_Above BYTE,"
    sql = sql & "Span_Right BYTE,"
    sql = sql & "Span_Below BYTE,"
    sql = sql & "Color TEXT(20),"
    sql = sql & "Color_Secondary TEXT(20),"
    sql = sql & "Substrate TEXT(20),"
    sql = sql & "Notes TEXT(255))"
    db.Execute sql, dbFailOnError

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
    ' v18 (delta D2): ID_Position REMOVED. One filled row in 27
    ' said everything: the column asked a question nobody was
    ' answering, because Element_Code already places an element
    ' and a Body / Whole structure row has no single position by
    ' definition. Its lone value travelled to Notes in the patch.
    ' Evidence_Scope STAYS, under rule 53 (mirror of 21): it is
    ' what tells a lost body from a lost element, and that
    ' distinction carries the R30 diagnosis of the perimetral Us.
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
    ' v20 (bloc I): Present dropped - a row exists because
    ' something was observed; the checkbox duplicated the
    ' row's own existence and its default 0 made rows lie.
    sql = sql & "Feature_Count INTEGER,"
    sql = sql & "Material TEXT(20),"
    sql = sql & "Notes TEXT(255))"
    db.Execute sql, dbFailOnError

    ' Physical connections between structures (OE3 network).
    ' v11: Chrono_Relation turns the edge list from an undirected graph
    ' into a potentially directed one, which is what H03 needs (9.1).
    ' v17 (delta D1): THE ORDER OF A AND B NO LONGER MEANS
    ' ANYTHING. In v16 the direction lived in the position -
    ' 'A earlier than B' - which is why the schema had to carry
    ' the instruction 'never swap A and B after entry', and why
    ' the pair could not be normalised. Chrono_Relation now
    ' NAMES WHICH STRUCTURE IS EARLIER (A is earlier / B is
    ' earlier / Contemporary / Undetermined, read against the
    ' normalised order), the form forces the LOWER ID into
    ' ID_Struct_A on BeforeUpdate, and the unique index below
    ' makes the duplicate pair IMPOSSIBLE rather than something
    ' to watch for.
    ' This was not hypothetical: the corpus held two rows and
    ' they were the same connection entered from each end. The
    ' cause was structural - F_CONNECTIONS linked on
    ' ID_Struct_A alone, so from B nothing was visible. The
    ' read-only incoming list in the form is the other half of
    ' the fix, and it must show the chronology INVERTED or it
    ' lies from B's side.
    On Error Resume Next
    db.Execute "DROP TABLE T_CONNECTIONS", dbFailOnError
    On Error GoTo 0
    sql = "CREATE TABLE T_CONNECTIONS ("
    sql = sql & "ID COUNTER CONSTRAINT PK_CON PRIMARY KEY,"
    sql = sql & "ID_Struct_A LONG NOT NULL,"
    sql = sql & "ID_Struct_B LONG NOT NULL,"
    ' v18 (delta D1): THE DIRECTION LEAVES THE PAIR FOR GOOD. The
    ' v17 fix normalised the order but kept the reading positional:
    ' 'A is earlier' still meant nothing without knowing which one
    ' A was, and the incoming subform had to INVERT it to not lie.
    ' Chrono_Relation now says only WHETHER there is a direction
    ' (Sequential / Contemporary / Undetermined) and ID_Earlier
    ' NAMES THE EARLIER STRUCTURE by foreign key, resolved by name
    ' in the form like every other FK. Swapping A and B no longer
    ' touches the chronology at all, and the incoming list shows
    ' the stored fact instead of a computed inversion. Rule 52
    ' keeps the triple coherent.
    sql = sql & "ID_Earlier LONG,"
    sql = sql & "Connection_Type TEXT(30),"
    sql = sql & "Chrono_Relation TEXT(20),"
    sql = sql & "Confidence TEXT(10),"
    sql = sql & "Notes TEXT(150))"
    db.Execute sql, dbFailOnError

    ' The pair, unique in the normalised order. Duplication
    ' stops being a rule in the battery and becomes impossible.
    On Error Resume Next
    db.Execute "CREATE UNIQUE INDEX UQ_CONN_PAIR ON T_CONNECTIONS (ID_Struct_A, ID_Struct_B)", dbFailOnError
    On Error GoTo 0

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
    For i = 0 To 20
        If SetDef(db, "T_STRUCTURES", m(i, 1), "0") Then ok = ok + 1 Else bad = bad + 1
    Next i
    Debug.Print "-> Default 0 set on " & ok & " of 21 element fields | failures: " & bad

    ' v18 (delta A4): Sill_Coincides_Cornice takes the PADDING 0,
    ' not NULL. Its applicability is derivable (Sill=0 with a real
    ' cornice), so the gating principle applies; kept OUT of the
    ' three-value array on purpose, because that array drives the
    ' NULL defaults and rule 17, and a padding field inside it
    ' would put every structure on the worklist.
    If SetDef(db, "T_STRUCTURES", "Sill_Coincides_Cornice", "0") Then
        Debug.Print "-> Default 0 (padding) on Sill_Coincides_Cornice"
    End If

    ' v19 (delta A1): same reasoning, same treatment - the
    ' reveal qualifier is derivable-inapplicable outside its
    ' window, so it takes the padding 0 and stays OUT of the
    ' three-value array. Note its 9 is NOT the usual 'not
    ' observable': the window (O at 0 or 2) already asserts
    ' the position was examined, so the only reading left is
    ' 'assessed, not decidable' - the same move R9 made for ND.
    If SetDef(db, "T_STRUCTURES", "Jamb_Fabric_Reveal", "0") Then
        Debug.Print "-> Default 0 (padding) on Jamb_Fabric_Reveal"
    End If

    ' v22 (bloc B): Metrics_Available is a judgement field
    ' with NULL default, but it stays OUT of the layer array
    ' ON PURPOSE, unlike Masonry_Present: metric availability
    ' is not decidable record-by-record until the extraction
    ' workflow runs, so rule 17 would put the whole corpus on
    ' the worklist for a question nobody can answer yet. No
    ' default is ever set on it, so it is born NULL; rule 63
    ' guards the detail side.

    ' The rest: DefaultValue cleared, so a new record starts empty.
    Dim nN As Integer
    For i = 0 To 19
        If ClearDef(db, "T_STRUCTURES", f(i)) Then nN = nN + 1
    Next i
    For i = 0 To 8
        If ClearDef(db, "T_STRUCTURES", g(i)) Then nN = nN + 1
    Next i
    Debug.Print "-> Default cleared (NULL) on " & nN & " of 29 non-gated observational fields"
    Debug.Print "-> Five-value domain: the 21 A-Z elements, and only them"
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
    If SetDef(db, "T_STRUCTURES", "Sys_Interface", q & "Absent" & q) Then ok = ok + 1 Else bad = bad + 1
    If SetDef(db, "T_STRUCTURES", "Platform_Function", q & "Undetermined" & q) Then ok = ok + 1 Else bad = bad + 1
    If SetDef(db, "T_CONNECTIONS", "Chrono_Relation", q & "Undetermined" & q) Then ok = ok + 1 Else bad = bad + 1
    Debug.Print "-> TEXT defaults set on " & ok & " of 8 fields | failures: " & bad
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
    ' v20 (bloc B): Sector_Code is the prefix every structure
    ' code must carry; rule 57 watches the concordance the
    ' EA46 case broke silently. LP keeps its directional
    ' sectors: future La Petaca codes read LP-N-EA01 etc.
    Dim s(11, 3) As String
    s(0, 0) = "1": s(0, 1) = "LP - General":  s(0, 2) = "La Petaca - general context":                 s(0, 3) = "LP-G"
    s(1, 0) = "1": s(1, 1) = "LP - North":    s(1, 2) = "La Petaca - North Sector":                    s(1, 3) = "LP-N"
    s(2, 0) = "1": s(2, 1) = "LP - Central":  s(2, 2) = "La Petaca - Central Sector":                  s(2, 3) = "LP-C"
    s(3, 0) = "1": s(3, 1) = "LP - Upper":    s(3, 2) = "La Petaca - Upper Sector":                    s(3, 3) = "LP-U"
    s(4, 0) = "1": s(4, 1) = "LP - South":    s(4, 2) = "La Petaca - South Sector (>=96 structures)":  s(4, 3) = "LP-S"
    s(5, 0) = "2": s(5, 1) = "DW - General":  s(5, 2) = "Diablo Wasi - general context":               s(5, 3) = "DW-G"
    s(6, 0) = "2": s(6, 1) = "DW - Sector 1": s(6, 2) = "Diablo Wasi - Sector 1 (~40 funerary contexts)": s(6, 3) = "DW-S01"
    s(7, 0) = "2": s(7, 1) = "DW - Sector 2": s(7, 2) = "Diablo Wasi - Sector 2 (1 chamber + basal cave)": s(7, 3) = "DW-S02"
    s(8, 0) = "2": s(8, 1) = "DW - Sector 3": s(8, 2) = "Diablo Wasi - Sector 3 (underground cave)":   s(8, 3) = "DW-S03"
    s(9, 0) = "2": s(9, 1) = "DW - Sector 4": s(9, 2) = "Diablo Wasi - Sector 4 (~9 funerary contexts)": s(9, 3) = "DW-S04"
    s(10, 0) = "2": s(10, 1) = "DW - Sector 5": s(10, 2) = "Diablo Wasi - Sector 5":                   s(10, 3) = "DW-S05"
    s(11, 0) = "2": s(11, 1) = "DW - Sector 6": s(11, 2) = "Diablo Wasi - Sector 6":                   s(11, 3) = "DW-S06"
    For i = 0 To 11
        db.Execute "INSERT INTO L_SECTORS (ID_Site,Sector_Name,Sector_Code,Description) VALUES (" & s(i, 0) & ",'" & s(i, 1) & "','" & s(i, 3) & "','" & s(i, 2) & "')", dbFailOnError
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
    ' v17: renamed from 'EA-PLA-R Ledge Platform'. PLATFORM is now
    ' reserved for the CORBELLED surface projecting over the void
    ' - element H and typology EA-PLA-V, which are the same thing
    ' at two scales, alone or inside a larger structure. A built
    ' mass levelling a natural ledge is a different object and
    ' was borrowing the word.
    ' 'Basal body' was rejected as the replacement: it is already
    ' taken by N_Basal_Bodies and by the BAS decoration position,
    ' so it would name a whole recording unit AND a component of
    ' one - the exact confusion this delta has been undoing.
    t(2, 0) = "EA-TER Ledge Terrace":     t(2, 1) = "Built funerary structure": t(2, 2) = "Built mass levelling a natural ledge. Function: transit or mausoleum base - the typology does not resolve which, Platform_Function does."
    t(3, 0) = "EA-PLA-V Aerial Platform": t(3, 1) = "Built funerary structure": t(3, 2) = "Artificial platform on wooden beams and slabs, without natural ledge support."
    t(4, 0) = "NIX Natural Niche":        t(4, 1) = "Natural funerary context": t(4, 2) = "Small natural cavity (<1m2). Function: ossuary or secondary burial."
    t(5, 0) = "CAV Cave/Cavern":          t(5, 1) = "Natural funerary context": t(5, 2) = "Large natural cavity (>1m2) with documented funerary or ritual use."
    t(6, 0) = "PR Rock Art":              t(6, 1) = "Rock art panel":           t(6, 2) = "Pictorial motif on rock, independently documented."
    t(7, 0) = "MEN Isolated Structural Element": t(7, 1) = "Structural trace": t(7, 2) = "Record whose evidence reduces to one or a few A-Z elements without a classifiable structure (bracket, transverse wall or pier, pilaster). Interpretation (circulation, earlier structure, support) goes to T_ARCH_FEATURES / Notes, never to the typology. The historical name comes from the first documented case, the bracket."
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
    Dim sup(12, 2) As String
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
    ' v17 (delta B1): the re-entrant angle between two rock
    ' planes with a roof over it. Geologically the NEGATIVE OF A
    ' WEDGE BLOCK detached along two intersecting
    ' discontinuities, the roof being the third fracture plane
    ' - but the name is taken from the descriptive tradition
    ' (dihedral) and not from the genesis, consistent with the
    ' purely geometric criterion adopted for the ROC positions.
    ' BOUNDARY WITH THE CAVITIES, geometric and not metric: a
    ' cavity is hollowed INTO the wall and is entered through a
    ' mouth; a dihedral has no mouth and is open in two
    ' directions. Field test: how many walls does the rock save
    ' you? The triangular roof is a usual description, NOT a
    ' condition - what counts is that the rock supplies two
    ' walls and a cover.
    sup(10, 0) = "Rock dihedral":            sup(10, 1) = "Confinement":        sup(10, 2) = "Re-entrant angle between two roughly orthogonal rock planes, usually roofed by a third. NATURAL, not cut. Not a cavity: a cavity is hollowed into the wall and has a mouth, a dihedral has none. Test: how many walls does the rock save?"
    ' v17 (delta B1b): support of rock art panels with NO
    ' associated structure. The mode is NEW because none of the
    ' three existing ones describes what it does: a surface
    ' carrying pigment is not gravitational rest, not
    ' confinement and not embedment. Its own mode keeps it out
    ' of every constructive-support count.
    sup(11, 0) = "Rock surface":             sup(11, 1) = "Substrate":          sup(11, 2) = "Exposed rock face carrying pigment, with no structure resting on or against it. For PR records only: on a structure record the rock is context, not support."
    sup(12, 0) = "ND":                        sup(12, 1) = "ND":                sup(12, 2) = "Support form not determined."
    ' WRITTEN CRITERION, guidance and not a validation rule: the
    ' PRIMARY support answers VERTICAL forces (it carries the
    ' weight), the SECONDARY answers HORIZONTAL ones (it
    ' stabilises). Deliberately unenforced: three of the four
    ' secondaries currently recorded are ledges or cavities,
    ' which carry weight, and they may well be right - a
    ' structure can rest on a ledge and on a cavity floor at
    ' once. A hard rule would flag them for nothing.
    For i = 0 To 12
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
    ' v16 (delta 2.1): SOC WITHDRAWN. C is defined as the
    ' DECORATIVE TREATMENT of the basal mass, so 'a decoration
    ' located on the socle' means 'a decoration located on the
    ' decorative treatment of the basal mass' - the position
    ' presupposes what the row is there to document. Keeping both
    ' entries and splitting them by motif type was considered and
    ' rejected: it would make the POSITION field encode the TYPE,
    ' which ID_Dec_Type already carries, and a painted relief
    ' frieze would have no possible value. Boundary now fixed: C is
    ' plastic treatment (projecting course, moulding, change of
    ' bond or stone format); a BAS row is an APPLIED motif.
    '
    ' v16 (delta 2.3): RTW, PRT, SIL, LIN added. v15 had no sill
    ' and no lintel, and Over-lintel is the frieze zone (M), not
    ' the lintel - so a painted portal had to be recorded against
    ' an element that was not the painted one. PRT exists because a
    ' continuous band framing the whole opening is ONE gesture, not
    ' three: the vocabulary already treats N+O+Q as a SYSTEM
    ' because the components act together. Use the elemental
    ' position when the motif BREAKS between elements or decorates
    ' one alone; PRT when the treatment runs continuously round the
    ' opening. IN DOUBT, ELEMENTAL: one can always aggregate
    ' afterwards, never disaggregate. RTW answers a documented
    ' case - decoration on the lateral wall and NOT on the facade,
    ' which is decoration not aimed at the viewer in the valley and
    ' bears directly on H06.
    '
    ' v16 (delta 2.2): terminology. 'Ala de facana' was a calque of
    ' flank and said nothing in the field; 'mur de retorn' names
    ' the geometric operation rather than the element, and 'lateral'
    ' was free again once L became 'flanc'. The English field names
    ' Facade_Flank and Return_Wall are UNCHANGED - they are the
    ' published ones - and 'in return' survives as the technical
    ' gloss in the Description column below.
    '
    ' v16 (delta 4.4): THE ROC ENTRIES ARE REDEFINED ON GEOMETRIC
    ' CONTACT. v15 sent 'tests 2 or 3' to ROC and test 3 also to
    ' ROC-PER, so any painting framing an opening matched BOTH
    ' entries and the choice fell to the observer. The new criterion
    ' is purely geometric - resolvable by looking, without
    ' consulting ID_Support - and the three are applied IN ORDER,
    ' most contact to least: overlapping, then perimeter, then
    ' proximate. ROC-OVL is new and is the analytically richest of
    ' the four: a motif crossing the joint between rock and masonry
    ' shows that whoever painted treated the two surfaces as one,
    ' and that the paint post-dates the fabric at that point
    ' (H04, H06). It carries Substrate = Mixed and the fabric side
    ' is recorded by Pigment_* on tab 3.Acab: ONE motif, ONE row.
    ' Test 1 restated accordingly - PARTIAL contact does not make a
    ' painting architectural; being WHOLLY on the fabric does.
    ' 'Proximate' needs an OUTER LIMIT or it swallows what should
    ' be a PR record: same accident of the cliff (ledge, cavity,
    ' cleft, recess) that houses the structure. Crossing a bench
    ' edge, a natural cornice or a void to reach it means PR.
    Dim sb(17, 4) As String
    sb(0, 0) = "BAS":     sb(0, 1) = "Basal body (B/C)":          sb(0, 2) = "N0":  sb(0, 3) = "10": sb(0, 4) = "Basal mass as a whole. Applied motif on the basal body; the plastic treatment itself is element C on tab 11.Sist."
    sb(1, 0) = "COR":     sb(1, 1) = "Interbody cornice (I)":     sb(1, 2) = "N1":  sb(1, 3) = "30": sb(1, 4) = "Cornice zone between superposed constructive bodies. Always stone: a projecting timber is E or F."
    sb(2, 0) = "QUO":     sb(2, 1) = "Corner quoin (J)":          sb(2, 2) = "N1":  sb(2, 3) = "40": sb(2, 4) = "Larger stones set vertically at the angles."
    sb(3, 0) = "PIL":     sb(3, 1) = "Pilaster (K)":              sb(3, 2) = "N1":  sb(3, 3) = "50": sb(3, 4) = "Structural pilaster framing the facade, full height."
    sb(4, 0) = "FFL":     sb(4, 1) = "Facade flank (L)":          sb(4, 2) = "N1":  sb(4, 3) = "60": sb(4, 4) = "Wall face within the facade plane. Where the opening is on the facade, the faces flanking it; where it is not, the running wall face. Does not depend on the opening."
    sb(5, 0) = "RTW":     sb(5, 1) = "Return wall (V)":           sb(5, 2) = "N1":  sb(5, 3) = "62": sb(5, 4) = "Chamber side wall, in return perpendicular to the facade plane, turning towards the cliff. Distinguished from L: L does not turn."
    sb(6, 0) = "PRT":     sb(6, 1) = "Portal system (P)":         sb(6, 2) = "N1":  sb(6, 3) = "64": sb(6, 4) = "Treatment running continuously round the opening, not attributable to any one component. Use an elemental position instead when the motif breaks between elements."
    sb(7, 0) = "SIL":     sb(7, 1) = "Sill (N)":                  sb(7, 2) = "N1":  sb(7, 3) = "65": sb(7, 4) = "Lower framing member of the access opening - the one stepped on."
    sb(8, 0) = "JAM":     sb(8, 1) = "Jamb (O)":                  sb(8, 2) = "N1":  sb(8, 3) = "70": sb(8, 4) = "Portal jamb - vertical frame element of the access opening."
    sb(9, 0) = "LIN":     sb(9, 1) = "Lintel (Q)":                sb(9, 2) = "N1":  sb(9, 3) = "75": sb(9, 4) = "Upper horizontal member of the access opening. NOT the same as Over-lintel, which is the zone above it."
    sb(10, 0) = "OVL":    sb(10, 1) = "Over-lintel (M)":          sb(10, 2) = "N1": sb(10, 3) = "80": sb(10, 4) = "Zone above the lintel - decorative frieze area (element M). Not the lintel itself: that is LIN."
    sb(11, 0) = "CRO":    sb(11, 1) = "Upper crown (R)":          sb(11, 2) = "N1": sb(11, 3) = "90": sb(11, 4) = "Upper crown or coping of the body."
    sb(12, 0) = "EAV":    sb(12, 1) = "Eave (U)":                 sb(12, 2) = "SUP": sb(12, 3) = "100": sb(12, 4) = "Eave / roof overhang above the facade."
    sb(13, 0) = "ROC-OVL": sb(13, 1) = "Overlapping":             sb(13, 2) = "ROC": sb(13, 3) = "200": sb(13, 4) = "One motif with part on the built fabric and part on the rock. Substrate = Mixed. Applied FIRST of the three."
    sb(14, 0) = "ROC-PER": sb(14, 1) = "Perimeter":               sb(14, 2) = "ROC": sb(14, 3) = "210": sb(14, 4) = "Follows the outline of the structure - PRESENT OR VANISHED - in tangential contact. Says outline, not contact, because contact is an observation of the present state: bands on bare rock aligned with the jambs of a body that is gone belong here (cross with T_LOST_ELEMENTS, rule 30)."
    sb(15, 0) = "ROC":    sb(15, 1) = "Proximate":                sb(15, 2) = "ROC": sb(15, 3) = "220": sb(15, 4) = "Touches the fabric nowhere, but lies within the same accident of the cliff - ledge, cavity, cleft or recess - that houses the structure. If a bench edge, natural cornice or void must be crossed to reach it, it is a PR record instead."
    sb(16, 0) = "ROC-PAN": sb(16, 1) = "Panel (no architectural ref.)": sb(16, 2) = "ROC": sb(16, 3) = "230": sb(16, 4) = "For rows hanging off an isolated rock art record (typology PR), which has no architectural position to refer to. NEVER chosen from a structure record."
    sb(17, 0) = "ND":     sb(17, 1) = "Not determined":           sb(17, 2) = "ND": sb(17, 3) = "999": sb(17, 4) = "Position not determined. NOT to be used for non-architectural positions: those are the ROC entries."
    For i = 0 To 17
        db.Execute "INSERT INTO L_STRUCT_BODY (Code,Name,Level_Type,Sort_Order,Description) VALUES ('" & sb(i, 0) & "','" & sb(i, 1) & "','" & sb(i, 2) & "'," & sb(i, 3) & ",'" & Replace(sb(i, 4), "'", "''") & "')", dbFailOnError
    Next i

    ' L_DEC_TYPE
    Dim dt(16, 1) As String
    dt(0, 0) = "T-shaped niche":      dt(0, 1) = "Niche or bas-relief in T form. Vertical + horizontal element."
    dt(1, 0) = "T-shaped niche inv.": dt(1, 1) = "Inverted T niche/relief."
    dt(2, 0) = "L-shaped niche":      dt(2, 1) = "Niche or bas-relief in L form."
    dt(3, 0) = "L-shaped niche inv.": dt(3, 1) = "Inverted L niche/relief."
    dt(4, 0) = "Zigzag":              dt(4, 1) = "Zigzag or chevron motif."
    dt(5, 0) = "Stepped motif":       dt(5, 1) = "Stepped/staircase motif. Rare at DW and LP."
    ' v17 (delta D4.2): 'Frieze / Greca' RETIRED. Zero uses in
    ' 56 rows, and it was the only entry in the list that names
    ' no motif - it said there is a frieze, which is exactly
    ' what Relief_Frieze (M) already says. The others (zigzag,
    ' stepped, triangular) do specify. ND covers the frieze
    ' whose motif could not be resolved, and keeping two values
    ' for one situation is what went wrong with the ROC
    ' positions in v15. If the greca as a specific Andean motif
    ' ever deserves an entry, it will be a NEW one with a motif
    ' definition, not this one.
    ' v17: 'Triangular motif' RETIRED - no entry in the corpus
    ' carries it, and it was never observed as a distinct motif
    ' from the zigzag.
    ' v17: 'Decapitation scene' RETIRED - a single documented
    ' case, and it is ROCK ART, not architectural decoration.
    ' It belongs to the rupestrian repertoire as RA
    ' Anthropomorphic with the reading in Notes, which is what
    ' its own description already instructed. A whole entry in
    ' the architectural list for one rock art motif inverted
    ' the proportion between vocabulary and corpus.
    dt(6, 0) = "Painted band":        dt(6, 1) = "Horizontal painted band (red/white). Interbody cornice zone."
    dt(7, 0) = "Square niche":        dt(7, 1) = "Square niches in series (hornacinas cuadradas). CONSTRUCTIVE motif: a frieze type."
    dt(8, 0) = "Plain colour field":  dt(8, 1) = "Flat chromatic application with no motif: a whole element painted one colour."
    dt(9, 0) = "ND":                  dt(9, 1) = "Decoration type not determined."
    ' v13: rock art types. The rupestrian repertoire is NOT shared
    ' with the architectural one - that was checked against the
    ' corpus - so these are additional entries, not a merge. Some
    ' generic types (Plain colour field, Painted band, ND) do serve
    ' both, which is why Level_Type and not ID_Dec_Type is the
    ' discriminator between the two sets.
    ' v17a: LA MEITAT RUPESTRE, REFETA SOBRE UN SOL EIX.
    ' La llista anterior barrejava tres preguntes: 'Painted band'
    ' nomenava una FORMA i 'Perimeter band' una POSICIO, de
    ' manera que una banda perimetral tambe era una banda
    ' pintada; i 'Abstract' i 'Amorphous stain' nomenaven totes
    ' dues un GRAU DE LLEGIBILITAT sense dir on era la frontera.
    ' Una llista les entrades de la qual responen preguntes
    ' distintes no es pot aplicar de manera consistent per molt
    ' be que es definisca cada entrada.
    ' La nova respon NOMES 'quin motiu es', i les dues formes
    ' en U absorbeixen el que Outline_Geometry registrava a
    ' banda. La distincio no es decorativa: un contorn
    ' ortogonal AFIRMA UN REFERENT CONSTRUIT rectangular i es
    ' l'evidencia d'estructura desapareguda (8bis.9); un
    ' d'organic no afirma res i pot resseguir un rebaix natural.
    dt(10, 0) = "RA Anthropomorphic": dt(10, 1) = "Rock art: anthropomorphic figure. Includes the decapitation scene: record the reading in Notes."
    dt(11, 0) = "RA Zoomorphic":      dt(11, 1) = "Rock art: zoomorphic figure."
        ' v20 (bloc F): the outline bands stop carrying the
    ' structural verdict - at La Petaca structures are outlined
    ' with C and even O bands, so no property of the band can
    ' decide it alone. The type says the STROKE; the topology
    ' (U, C, O) DERIVES from the four span fields (QRY_27); the
    ' verdict lives in the record typology and attestation.
    ' Losing the RA prefix moves them to the non-rock-art side
    ' of the type combo BY ITSELF, which is correct: a pure
    ' panel outlines nothing (a U as panel iconography is RA
    ' Geometric motif).
    dt(12, 0) = "Outline band, defined": dt(12, 1) = "Pigment band following an outline, with a clean stroke and clear edges. The structural verdict lives in the record typology and attestation, never here. Topology (U, C, O) derives from the four span fields: QRY_27."
    dt(13, 0) = "Outline band, amorphous": dt(13, 1) = "Pigment band following an outline, with a diffuse stroke or ill-defined edges. The structural verdict lives in the record typology and attestation, never here. Topology (U, C, O) derives from the four span fields: QRY_27."
    dt(14, 0) = "RA Geometric motif": dt(14, 1) = "Rock art: lines, bands or figures with recognisable regular organisation."
    ' La frontera entre les dues seguents es si el pigment te
    ' VORA LLEGIBLE, no si sembla significatiu: era la confusio
    ' que produia el parell Abstract / Amorphous.
    dt(15, 0) = "RA Amorphous stain": dt(15, 1) = "Rock art: defined pigment surface with a recognisable edge but no identifiable motif."
    dt(16, 0) = "RA Pigment traces":  dt(16, 1) = "Rock art: scattered or degraded remains, too poor to say whether they formed a motif."
    For i = 0 To 16
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
    VL db, "L_TYPOLOGY", "EA-TER Ledge Terrace", "EA-TER Terrassa en repisa"
    VL db, "L_TYPOLOGY", "EA-PLA-V Aerial Platform", "EA-PLA-V Plataforma aeria"
    VL db, "L_TYPOLOGY", "NIX Natural Niche", "NIX Ninxol natural"
    VL db, "L_TYPOLOGY", "CAV Cave/Cavern", "CAV Cova/Cavitat"
    VL db, "L_TYPOLOGY", "PR Rock Art", "PR Art rupestre"
    VL db, "L_TYPOLOGY", "MEN Isolated Structural Element", "MEN Element estructural aillat"
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
    VL db, "L_SUPPORT", "Rock dihedral", "Diedre rocos (angle amb sostre)"
    VL db, "L_SUPPORT", "Rock surface", "Superficie de roca"
    VL db, "L_SUPPORT", "ND", "Indeterminat"

    VL db, "L_DEC_TYPE", "T-shaped niche", "Ninxol en T"
    VL db, "L_DEC_TYPE", "T-shaped niche inv.", "Ninxol en T invertida"
    VL db, "L_DEC_TYPE", "L-shaped niche", "Ninxol en L"
    VL db, "L_DEC_TYPE", "L-shaped niche inv.", "Ninxol en L invertida"
    VL db, "L_DEC_TYPE", "Zigzag", "Ziga-zaga"
    VL db, "L_DEC_TYPE", "Stepped motif", "Motiu escalonat"
    VL db, "L_DEC_TYPE", "Painted band", "Banda pintada"
    VL db, "L_DEC_TYPE", "Square niche", "Ninxol quadrat"
    VL db, "L_DEC_TYPE", "Plain colour field", "Camp de color pla"
    VL db, "L_DEC_TYPE", "ND", "Indeterminat"
    VL db, "L_DEC_TYPE", "RA Anthropomorphic", "Antropomorf"
    VL db, "L_DEC_TYPE", "RA Zoomorphic", "Zoomorf"
    VL db, "L_DEC_TYPE", "Outline band, defined", "Banda de contorn definida"
    VL db, "L_DEC_TYPE", "Outline band, amorphous", "Banda de contorn amorfa"
    VL db, "L_DEC_TYPE", "RA Geometric motif", "Motiu geometric"
    VL db, "L_DEC_TYPE", "RA Amorphous stain", "Taca amorfa"
    VL db, "L_DEC_TYPE", "RA Pigment traces", "Traces de pigment"

    VL db, "L_STRUCT_BODY", "Basal body (B/C)", "Cos basal (B/C)"
    VL db, "L_STRUCT_BODY", "Interbody cornice (I)", "Cornisa intercos (I)"
    VL db, "L_STRUCT_BODY", "Corner quoin (J)", "Cantonera (J)"
    VL db, "L_STRUCT_BODY", "Pilaster (K)", "Pilastra (K)"
    VL db, "L_STRUCT_BODY", "Facade flank (L)", "Flanc de facana (L)"
    ' v18 (delta A6): one term everywhere - mur de retorn.
    VL db, "L_STRUCT_BODY", "Return wall (V)", "Mur de retorn (V)"
    VL db, "L_STRUCT_BODY", "Portal system (P)", "Sistema portal (P)"
    VL db, "L_STRUCT_BODY", "Sill (N)", "Llindar (N)"
    VL db, "L_STRUCT_BODY", "Jamb (O)", "Brancal (O)"
    ' Q is ALWAYS 'Dintell' in the interface and NEVER 'llinda':
    ' llinda and llindar are near-homographs meaning opposite ends
    ' of the same frame, and a direct source of recording error.
    VL db, "L_STRUCT_BODY", "Lintel (Q)", "Dintell (Q)"
    VL db, "L_STRUCT_BODY", "Over-lintel (M)", "Sobre-dintell (M)"
    VL db, "L_STRUCT_BODY", "Upper crown (R)", "Coronament (R)"
    VL db, "L_STRUCT_BODY", "Eave (U)", "Rafec (U)"
    VL db, "L_STRUCT_BODY", "Overlapping", "Solapada"
    VL db, "L_STRUCT_BODY", "Perimeter", "Perimetral"
    VL db, "L_STRUCT_BODY", "Proximate", "Proxima"
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

' v16a: tolerating the error is right - one bad label must not stop
' the build - but a tolerated error still has to leave a trace where
' someone looks. A missing Name_VAL column silently swallowed nine
' updates in a row and nobody knew until the dropdown came up empty.
Private Sub VL(db As DAO.Database, tbl As String, en As String, va As String)
    On Error Resume Next
    db.Execute "UPDATE " & tbl & " SET Name_VAL='" & Replace(va, "'", "''") & "' WHERE Name='" & Replace(en, "'", "''") & "'", dbFailOnError
    If Err.Number <> 0 Then
        Debug.Print "  *** VL failed on " & tbl & " / " & en & ": " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Private Sub PopulateElements(db As DAO.Database)
    Dim el(24, 7) As String
    el(0, 0) = "A":  el(0, 1) = "Embedded base beams":         el(0, 2) = "Jaceres basals":         el(0, 3) = "N0":    el(0, 4) = "":         el(0, 5) = "False": el(0, 6) = "Embedded_Base_Beams":  el(0, 7) = "Timber beams embedded in the basal masonry."
    el(1, 0) = "B":  el(1, 1) = "Base level":                  el(1, 2) = "Basament":               el(1, 3) = "N0":    el(1, 4) = "":         el(1, 5) = "False": el(1, 6) = "Base_Level":           el(1, 7) = "Constructed basal level supporting the structure."
    el(2, 0) = "C":  el(2, 1) = "Decorative socle":            el(2, 2) = "Socol decoratiu":        el(2, 3) = "N0":    el(2, 4) = "":         el(2, 5) = "False": el(2, 6) = "Decorative_Socle":     el(2, 7) = "Decorative treatment of the basal mass."
    el(3, 0) = "D":  el(3, 1) = "Transverse wall or pier":                   el(3, 2) = "Muret transversal":      el(3, 3) = "N0":    el(3, 4) = "":         el(3, 5) = "False": el(3, 6) = "Tie_Walls":            el(3, 7) = "Transverse wall or pier anchoring the fabric to the rock; from a long anchoring wall to a compact pier - length does not change the letter. May anchor at any height of the fabric. Does not enclose any interior and does not count as a body. Against E: the corbel HANGS, the pier RISES. Against K: the pilaster lives IN the facade wall plane, the pier against the rock. Ungated since v18, permanently active like I and R."
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
    ' v18 (delta C1): Z closes the platform the way T closes the
    ' eave. Y stays RESERVED for the abutted mass (banqueta),
    ' pending its three documented cases.
    el(24, 0) = "Z": el(24, 1) = "Platform surface":            el(24, 2) = "Superficie de plataforma": el(24, 3) = "N0": el(24, 4) = "Platform": el(24, 5) = "False": el(24, 6) = "Platform_Surface":  el(24, 7) = "Finished walking surface of the corbelled platform; component of the platform system (H). Material in Platform_Surface_Material."
    el(23, 0) = "X": el(23, 1) = "Chamber roof":               el(23, 2) = "Coberta de cambra":     el(23, 3) = "N1":   el(23, 4) = "":       el(23, 5) = "False": el(23, 6) = "Chamber_Roof":         el(23, 7) = "Element closing the chamber above; component of Sys_Chamber."

    Dim sql As String
    Dim i As Integer
    For i = 0 To 24
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
    Debug.Print "-> L_ELEMENTS: 25 rows (A-X + Z; Y reserved)"
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
    ' v18: REL_SB_LOST gone with ID_Position; the slot now holds
    ' the third edge of T_CONNECTIONS into T_STRUCTURES.
    rn(24) = "REL_STR_CONE"
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
    ' Third reference to T_STRUCTURES from the same table -> no
    ' enforced integrity (flag 2), same treatment as A and B.
    MkRel db, rn(24), "T_STRUCTURES", "ID", "T_CONNECTIONS", "ID_Earlier", False, True

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
    Dim qn(40) As String
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
    qn(27) = "QRY_16d_Rules_32_39"
    qn(28) = "QRY_19_V16_Review"
    qn(29) = "QRY_16e_Rules_40_46"
    qn(30) = "QRY_20_RockArt_Span"
    qn(31) = "QRY_21_V17_Review"
    qn(32) = "QRY_16_Validation_Check"
    ' v18: rules 47-54 (now 47-56) and the migration worklists.
    qn(33) = "QRY_16f_Rules_47_56"
    qn(34) = "QRY_23_V18_Review"
    ' v19: the migration worklist of the v18->v19 patch.
    qn(35) = "QRY_24_V19_Review"
    ' v20: the seventh rules partial and the three new tools.
    ' v21: rules 61-62 join the seventh partial.
    qn(36) = "QRY_16g_Rules_57_64"
    qn(37) = "QRY_25_V20_Review"
    qn(38) = "QRY_26_Next_EA"
    qn(39) = "QRY_27_Outline_Topology"
    ' v21: the migration worklist of the v20->v21 patch.
    ' QRY_28 is taken by the worklist navigator add-on.
    qn(40) = "QRY_29_V21_Review"
    Dim i As Integer
    ' Reverse order so dependants go before their sources
    For i = 40 To 0 Step -1
        If QueryExists(db, qn(i)) Then db.QueryDefs.Delete qn(i)
    Next i

    ' v22 (bloc E): LEGACY NAMES. The loop above only knows
    ' the current names, so a renamed partial survives as an
    ' inert orphan - the v21 rename left QRY_16g_Rules_57_60
    ' behind and it took a manual delete. Known legacy names
    ' are removed explicitly from now on.
    Dim legacy(1) As String
    legacy(0) = "QRY_16g_Rules_57_60"
    legacy(1) = "QRY_16g_Rules_57_62"
    For i = 0 To 1
        If QueryExists(db, legacy(i)) Then db.QueryDefs.Delete legacy(i)
    Next i

    BuildBasicQueries db
    BuildDecorationQueries db
    BuildExportQueries db
    BuildAXExport db
    BuildRockArtQuery db
    BuildNotesQuery db
    BuildValidationHelpers db
    BuildValidationBattery db
    BuildReviewQuery db
    BuildSpanQuery db
    BuildV17ReviewQuery db
    BuildV18ReviewQuery db
    BuildV19ReviewQuery db
    BuildV20ReviewQuery db
    BuildNextEAQuery db
    BuildOutlineTopologyQuery db
    BuildV21ReviewQuery db

    ' No hard-coded 'OK': ReportQueryCount (called from BuildDB)
    ' counts what actually exists against what was expected.
    Debug.Print "-> query build finished; see the count below"
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
    q = q & "IIF(E.Relief_Frieze IN (1,2,3),1,0) AS Has_Frieze, "
    q = q & "IIF(E.Pigment_Present=1 AND E.Pigment_Substrate='Bedrock',1,0) AS Has_RockArt "
    q = q & "FROM ((((T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID) "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID) "
    q = q & "LEFT JOIN QRY_02s_Decoration_Typed AS DX ON DX.ID_Structure=E.ID) "
    q = q & "WHERE T.Record_Class<>'Pending classification' "
    q = q & "GROUP BY E.ID, E.Code, S.Site_Name, SC.Sector_Name, T.Record_Class, "
    q = q & "E.Facade_Observability, E.Pigment_Present, E.Pigment_Substrate, E.Relief_Frieze "
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
    q = q & "E.Masonry_Present, E.Metrics_Available, E.Masonry_Quality, E.Masonry_Type, "
    q = q & "E.Stone_Format, E.Stone_Format_Secondary, E.Stone_Working, "
    q = q & "E.Mortar_Present, E.Mortar_Type, E.Chinking_Stones, "
    q = q & "E.Opening_Width_m, E.Opening_Height_m, "
    q = q & "E.Height_Above_Base_m, "
    q = q & "E.Facade_Orientation, E.Visibility_Valley, "
    q = q & "E.Access_Plane, E.Portal_Orientation, "
    q = q & "E.Construction_Phases, E.Phase_Evidence, "
    q = q & "E.Timber_Bracket_Count, E.Timber_Bracket_Role, "
    ' v18: the two new companion fields ride along.
    q = q & "E.Platform_Surface_Material, "
    q = q & "E.Sill_Coincides_Cornice, "
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
    q = "SELECT E.ID, E.Code, "
    ' v20 (bloc B): the EA number as a query expression - no
    ' stored field (it would duplicate the code), numeric
    ' sorting and QGIS labels come for free.
    q = q & "Val(Mid(E.Code, InStr(E.Code, '-EA') + 3)) AS EA_Num, "
    q = q & "S.Site_Name, SC.Sector_Name, SC.Sector_Code, "
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
    ' v16: Access_Plane and Portal_Orientation go to QGIS (H06).
    ' Their DIVERGENCE from Facade_Orientation is the analysable
    ' quantity: entry displaced from the exposed plane means
    ' circulation overrode display, and that can now be crossed
    ' with typology, sector, support form and ledge width.
    q = q & "E.Facade_Orientation, E.Visibility_Valley, E.Masonry_Quality, "
    q = q & "E.Access_Plane, E.Portal_Orientation, "
    ' v18 (delta A2): where the access is ON the facade the two
    ' orientations are one datum recorded once; the effective
    ' column resolves it so R never averages a duplicate against
    ' a blank.
    q = q & "IIF(E.Access_Plane='Facade',E.Facade_Orientation,E.Portal_Orientation) AS Portal_Orientation_Effective, "
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
    ' v18 (delta D1): the earlier structure by code, resolved
    ' through the FK - the graph is directed by name, not by
    ' the accident of the A/B positions.
    q = q & "EC.Code AS Code_Earlier, "
    q = q & "A.Coord_E_UTM AS E_UTM_A, A.Coord_N_UTM AS N_UTM_A, "
    q = q & "A.Altitude_masl AS Alt_A, "
    q = q & "B.Coord_E_UTM AS E_UTM_B, B.Coord_N_UTM AS N_UTM_B, "
    q = q & "B.Altitude_masl AS Alt_B, "
    q = q & "C.Notes "
    q = q & "FROM (((T_CONNECTIONS AS C "
    q = q & "INNER JOIN T_STRUCTURES AS A ON C.ID_Struct_A=A.ID) "
    q = q & "INNER JOIN T_STRUCTURES AS B ON C.ID_Struct_B=B.ID) "
    q = q & "LEFT JOIN T_STRUCTURES AS EC ON C.ID_Earlier=EC.ID) "
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

    Dim cm(6, 1) As String
    cm(0, 0) = "Timber_Bracket_Count":       cm(0, 1) = "Sys_Platform"
    cm(1, 0) = "Timber_Bracket_Role":        cm(1, 1) = "Sys_Platform"
    cm(2, 0) = "Platform_Surface_Material":  cm(2, 1) = "Sys_Platform"
    cm(3, 0) = "Platform_Function":          cm(3, 1) = "Sys_Platform"
    cm(4, 0) = "Lintel_Material":            cm(4, 1) = "Sys_Portal"
    ' v18: Recessed_Frame leaves this map - its gate is the body
    ' count, not a system (delta A1). v20: the cornice format
    ' leaves with its field (bloc 2).
    cm(5, 0) = "Chamber_Roof_Type":          cm(5, 1) = "Sys_Chamber"
    cm(6, 0) = "Rear_Closure_Type":          cm(6, 1) = "Sys_Chamber"

    Dim sy(5, 1) As String
    sy(0, 0) = "Sys_Platform": sy(0, 1) = "SYS_H"
    sy(1, 0) = "Sys_Portal":   sy(1, 1) = "SYS_P"
    sy(2, 0) = "Sys_Eave":     sy(2, 1) = "SYS_U"
    sy(3, 0) = "Sys_Base":     sy(3, 1) = "SYS_BASE"
    sy(4, 0) = "Sys_Chamber":  sy(4, 1) = "SYS_CHAMBER"
    ' v18 FIX: the sixth system never exported - I was gated by
    ' it in the matrix while its own state stayed invisible.
    sy(5, 0) = "Sys_Interface": sy(5, 1) = "SYS_INTERFACE"

    Dim q As String
    Dim i As Integer

    q = "SELECT E.ID, E.Code, S.Site_Name AS Site, SC.Sector_Name AS Sector, "
    q = q & "T.Name AS Typology, T.Record_Class, "
    q = q & "E.Coord_E_UTM, E.Coord_N_UTM, E.Altitude_masl, E.ID_Group, "
    q = q & "E.N_Basal_Bodies, E.N_Chamber_Bodies, "
    ' v16 (delta 5.2, 5.3): X IS A STATE, NOT AN ELEMENT. Defined
    ' as 'closed above by whatever solution', Chamber_Roof entered
    ' this matrix identically whether masonry was raised or a rock
    ' visor did the work. With many natural-bedrock roofs in the
    ' corpus, Jaccard and correspondence analysis were counting
    ' GEOLOGY AS CONSTRUCTION and inflating the similarity between
    ' structures that share no constructive decision at all - the
    ' same family of error as ID_Material_Status and as Lintel in
    ' v10, two variables inside one scale.
    '
    ' NA AND NOT 0: it is not an absence but an UNBUILT SOLUTION.
    ' A 0 would be as false as a 1. Not applicable enters no
    ' denominator, which is exactly the behaviour needed.
    '
    ' X = 0 STILL EXPORTS 0. Under the presence+type pattern the
    ' type is NULL when the presence is 0, so the override below
    ' can only fire on 1, 2 or 3 - a real absence of closure
    ' (including a visor passing metres above without contact:
    ' no contact, no roof) is a genuine result and stays a 0.
    For i = 0 To 20
        If ax(i, 0) = "X" Then
            q = q & GatedRoof()
        Else
            q = q & Gated(ax(i, 1), ax(i, 2), "AX_" & ax(i, 0))
        End If
    Next i
    For i = 0 To 5
        q = q & Gated(sy(i, 0), sy(i, 0), sy(i, 1))
    Next i
    For i = 0 To 6
        q = q & Gated(cm(i, 0), cm(i, 1), cm(i, 0))
    Next i
    ' v18 (delta A1): Recessed_Frame nullified where there is no
    ' chamber body - no facade, no frame - instead of by system.
    q = q & "IIF(E.N_Chamber_Bodies=0,Null,E.Recessed_Frame) AS Recessed_Frame, "
    q = q & "E.Masonry_Present, E.Metrics_Available, E.Masonry_Quality, E.Masonry_Type, "
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
' v16: X exports NA when the closure is natural rock (delta 5.3).
' Kept as its own function rather than a special case inside
' Gated, so that the general gating rule stays readable and this
' exception stays visible as an exception.
Private Function GatedRoof() As String
    Dim g As String
    g = "IIF(E.Sys_Chamber='Not applicable' Or E.Sys_Chamber='Not observable',Null,"
    g = g & "IIF(E.Chamber_Roof_Type='Natural bedrock',Null,E.Chamber_Roof)) AS AX_X, "
    GatedRoof = g
End Function

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
    ' v18 BUG FIX: this query still read E.Notes, renamed to
    ' Doc_Notes by v17a, so opening it raised a parameter
    ' prompt. The five v17a tab notes join the sweep too: this
    ' is the query that turns recurring free text into
    ' candidate fields, so it must see all of it.
    q = q & "E.Id_Notes, E.Arch_Notes, E.Finish_Notes, "
    q = q & "E.Dec_Notes, E.Condition_Notes, E.Bio_Notes, "
    q = q & "E.Materials_Notes, E.Chrono_Notes, E.Metric_Notes, "
    q = q & "E.Mortar_Notes, E.Vol_Notes, E.Systems_Notes, "
    q = q & "E.Extra_Notes, E.Doc_Notes "
    q = q & "FROM (T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID) "
    q = q & "INNER JOIN L_SITES AS S ON SC.ID_Site=S.ID "
    q = q & "WHERE E.Id_Notes Is Not Null OR E.Arch_Notes Is Not Null "
    q = q & "OR E.Finish_Notes Is Not Null OR E.Dec_Notes Is Not Null "
    q = q & "OR E.Condition_Notes Is Not Null OR E.Bio_Notes Is Not Null "
    q = q & "OR E.Materials_Notes Is Not Null OR E.Chrono_Notes Is Not Null "
    q = q & "OR E.Metric_Notes Is Not Null OR E.Mortar_Notes Is Not Null "
    q = q & "OR E.Vol_Notes Is Not Null OR E.Systems_Notes Is Not Null "
    q = q & "OR E.Extra_Notes Is Not Null OR E.Doc_Notes Is Not Null "
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
    q = q & "D.Span_Left, D.Span_Above, D.Span_Right, D.Span_Below, "
    q = q & "E.Facade_Orientation, E.Portal_Orientation, E.Access_Plane, "
    q = q & "E.Visibility_Valley, "
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
    For i = 0 To 20
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
    For i = 0 To 20
        If i > 0 Then q = q & " UNION ALL "
        q = q & "SELECT E.ID AS ID_Structure, E.Code AS Code, '" & m(i, 1) & "' AS Field_Name "
        q = q & "FROM T_STRUCTURES AS E WHERE E." & m(i, 1) & " Is Null"
    Next i
    For i = 0 To 19
        q = q & " UNION ALL "
        q = q & "SELECT E.ID, E.Code, '" & f(i) & "' "
        q = q & "FROM T_STRUCTURES AS E WHERE E." & f(i) & " Is Null"
    Next i
    For i = 0 To 8
        q = q & " UNION ALL "
        q = q & "SELECT E.ID, E.Code, '" & g(i) & "' "
        q = q & "FROM T_STRUCTURES AS E WHERE E." & g(i) & " Is Null"
    Next i
    q = q & ";"
    MkQuery db, "QRY_16s_Observational_Nulls", q

    q = "SELECT E.ID AS ID_Structure, E.Code AS Code, "
    For i = 0 To 20
        If i > 0 Then q = q & "+"
        q = q & "IIF(E." & m(i, 1) & " Is Null,1,0)"
    Next i
    For i = 0 To 19
        q = q & "+IIF(E." & f(i) & " Is Null,1,0)"
    Next i
    For i = 0 To 8
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
    ' v16 BUG FIX: v15 created this query as QRY_16c_Rules_22_30
    ' while both the drop array and the union below referred to
    ' QRY_16c_Rules_22_31. CreateQueryDef on the union therefore
    ' failed - and MkQuery degrades that failure to a Debug.Print,
    ' so BuildDB finished looking successful while THE BATTERY HAD
    ' NO ENTRY POINT. The three partials existed and opened fine;
    ' only the union was missing, which is why it went unnoticed.
    MkQuery db, "QRY_16c_Rules_22_31", q

    BuildValidationD db
    BuildValidationE db
    BuildValidationF db
    BuildValidationG db

    q = "SELECT * FROM QRY_16a_Rules_1_11 "
    q = q & "UNION ALL SELECT * FROM QRY_16b_Rules_12_21 "
    q = q & "UNION ALL SELECT * FROM QRY_16c_Rules_22_31 "
    q = q & "UNION ALL SELECT * FROM QRY_16d_Rules_32_39 "
    q = q & "UNION ALL SELECT * FROM QRY_16e_Rules_40_46 "
    q = q & "UNION ALL SELECT * FROM QRY_16f_Rules_47_56 "
    q = q & "UNION ALL SELECT * FROM QRY_16g_Rules_57_64 "
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
    ' v18 (delta C1): Z can carry the system on its own - a
    ' surface whose supports cannot be resolved is still a platform.
    s(0, 0) = "Sys_Platform": s(0, 1) = "E.Timber_Brackets=0 AND E.Transverse_Beams=0 AND E.Corbelled_Courses=0 AND E.Platform_Surface=0": s(0, 2) = "E, F, G and Z"
    ' v19 (delta A2): the portal branch is exempt when the
    ' reveal is declared - the all-zero portal is then the
    ' legitimate node-2 case of the manual (opening resolved
    ' in the fabric), not a half-entered record. Same route as
    ' the Timber_Brackets exemption on R1. The platform keeps
    ' its equivalent hole OPEN (delta v18 A5, EA-PLA-V with
    ' all components 0 is legal and R3 still fires): a known,
    ' accepted false positive until a qualifier of its own
    ' accumulates the three cases.
    s(1, 0) = "Sys_Portal":   s(1, 1) = "E.Sill=0 AND E.Jambs=0 AND E.Lintel=0 AND E.Jamb_Fabric_Reveal<>1":      s(1, 2) = "N, O and Q"
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
    q = q & "'Absence on a collapse demands a legible surface: use 3 with evidence, 9, or confirm surface legibility in Notes' "
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

' v19 (delta B1): R9 fires on Is Null ONLY. ND is a judgement
' ('assessed, not decidable' - the NULL/9/0 distinction carried
' into a TEXT field), not an empty cell; with the old condition
' the branch could never be emptied (7 of 35 rows) and the
' battery accumulated permanent noise.
Private Function R09() As String
    Dim q As String
    q = "SELECT E.Code, 9, "
    q = q & "'R9: corbels present but Timber_Bracket_Role is not recorded', "
    q = q & "'Run the three-question tree: support, no link, or ND' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Timber_Brackets IN (1,2) "
    q = q & "AND E.Timber_Bracket_Role Is Null"
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

' ================================================================
'  v16 - QUERY COUNT REPORT
'  MkQuery deliberately degrades a failure to a Debug.Print, which
'  is the right call: 30 queries must not be lost because one of
'  them is malformed. But in v15 that silence hid a real failure
'  (QRY_16_Validation_Check never existed) through several working
'  sessions. A silenced error still has to leave a trace where
'  someone looks, so BuildDB now compares actual against expected.
' ================================================================
Private Sub ReportQueryCount(db As DAO.Database)
    Dim n As Integer
    Dim qd As DAO.QueryDef
    db.QueryDefs.Refresh
    For Each qd In db.QueryDefs
        If Left(qd.Name, 4) = "QRY_" Then n = n + 1
    Next qd
    Debug.Print "-> queries created: " & n & " of 41 expected"
    If n < 35 Then
        Debug.Print "  *** SOME QUERIES FAILED. Scroll up for the"
        Debug.Print "  *** 'FAILED to create' lines naming them."
    End If
End Sub

' ================================================================
'  v16 - QRY_19_V16_REVIEW
'  The worklist of the manual revision the v15->v16 transfer
'  requires. Three of the four branches exist because fields
'  arrive NULL by the conservative principle - a value carried
'  across unreviewed would assert a judgement nobody made under
'  the new criterion - and the fourth because decoration rows
'  cannot be emptied without losing the whole record, so they are
'  LISTED instead.
' ================================================================
Private Sub BuildReviewQuery(db As DAO.Database)
    Dim q As String
    q = "SELECT E.Code AS Structure, 'Portal elements' AS Review_Item, "
    q = q & "'Re-judge N, O and Q under the differentiation criterion: a plain gap is 0' AS Reason "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Sys_Portal Like 'Present*' "
    q = q & "AND (E.Sill Is Null Or E.Jambs Is Null Or E.Lintel Is Null) "

    q = q & "UNION ALL SELECT E.Code, 'Chamber roof', "
    q = q & "'Re-judge X under the contact test: no contact, no roof' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Chamber_Roof Is Null "

    q = q & "UNION ALL SELECT E.Code, 'Stone format and working', "
    q = q & "'New v16 fields: enter from photograph or field sheet' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Stone_Format Is Null "

    q = q & "UNION ALL SELECT E.Code, 'Decoration position', "
    q = q & "'Row recorded as JAM or OVL may now belong to PRT, SIL or LIN' "
    q = q & "FROM ((T_DECORATIONS AS D "
    q = q & "INNER JOIN T_STRUCTURES AS E ON D.ID_Structure=E.ID) "
    q = q & "INNER JOIN L_STRUCT_BODY AS B ON D.ID_Struct_Body=B.ID) "
    q = q & "WHERE B.Code IN ('JAM','OVL') "
    q = q & "ORDER BY Structure, Review_Item;"
    MkQuery db, "QRY_19_V16_Review", q
End Sub

' ================================================================
'  v16 - VALIDATION BATTERY, FOURTH PARTIAL (RULES 32-39)
'  Nine branches for eight rule numbers: R38 splits, because the
'  substrate a ROC row must carry depends on which ROC position it
'  is. Same principle as the rest of the battery: a REPORT, never
'  a table constraint - a hard rule would block a genuine
'  observation in a partially collapsed structure.
' ================================================================
Private Sub BuildValidationD(db As DAO.Database)
    Dim q As String
    q = R32()
    q = q & " UNION ALL " & R33()
    q = q & " UNION ALL " & R34()
    q = q & " UNION ALL " & R35()
    q = q & " UNION ALL " & R36()
    q = q & " UNION ALL " & R37()
    q = q & " UNION ALL " & R38a()
    q = q & " UNION ALL " & R38b()
    q = q & " UNION ALL " & R39()
    q = q & ";"
    MkQuery db, "QRY_16d_Rules_32_39", q
End Sub

' R32-R33: the ordered pair, same two failure modes already
' guarded on ID_Support - an orphan secondary, and a pair that
' says the same thing twice.
Private Function R32() As String
    Dim q As String
    q = "SELECT E.Code AS Structure, 32 AS Rule_No, "
    q = q & "'R32: secondary stone format set with no dominant format' AS Rule_Violated, "
    q = q & "'Set Stone_Format, or clear the secondary' AS Action "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Stone_Format_Secondary Is Not Null "
    q = q & "AND (E.Stone_Format Is Null Or E.Stone_Format='ND')"
    R32 = q
End Function

Private Function R33() As String
    Dim q As String
    q = "SELECT E.Code, 33, "
    q = q & "'R33: primary and secondary stone format are the same value', "
    q = q & "'The ordered pair needs two distinct formats' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Stone_Format_Secondary=E.Stone_Format"
    R33 = q
End Function

' R34 is DELIBERATELY ONE-DIRECTIONAL. C and the BAS rows are
' independent: there can be a moulding with no motif and a motif
' with no moulding. Only one combination is incoherent - a RELIEF
' on the basal body IS a plastic treatment, so C cannot be 0.
Private Function R34() As String
    Dim q As String
    q = "SELECT E.Code, 34, "
    q = q & "'R34: relief motif on the basal body but Decorative_Socle is 0', "
    q = q & "'A relief on the basal body is itself a plastic treatment' "
    q = q & "FROM (((T_DECORATIONS AS D "
    q = q & "INNER JOIN T_STRUCTURES AS E ON D.ID_Structure=E.ID) "
    q = q & "INNER JOIN L_STRUCT_BODY AS B ON D.ID_Struct_Body=B.ID) "
    q = q & "INNER JOIN L_DEC_TYPE AS DT ON D.ID_Dec_Type=DT.ID) "
    q = q & "WHERE B.Code='BAS' AND E.Decorative_Socle=0 "
    q = q & "AND DT.Name IN ('Zigzag','Stepped motif',"
    q = q & "'Square niche','T-shaped niche',"
    q = q & "'T-shaped niche inv.','L-shaped niche','L-shaped niche inv.')"
    R34 = q
End Function

' R35: same presence-detail pattern as R10 and R13, applied to the
' two positions v16 adds.
Private Function R35() As String
    Dim q As String
    q = "SELECT E.Code, 35, "
    q = q & "'R35: decoration recorded on a portal element that is absent', "
    q = q & "'Reconcile the decoration position with Sill or Lintel' "
    q = q & "FROM ((T_DECORATIONS AS D "
    q = q & "INNER JOIN T_STRUCTURES AS E ON D.ID_Structure=E.ID) "
    q = q & "INNER JOIN L_STRUCT_BODY AS B ON D.ID_Struct_Body=B.ID) "
    q = q & "WHERE (B.Code='LIN' AND E.Lintel=0) "
    q = q & "OR (B.Code='SIL' AND E.Sill=0)"
    R35 = q
End Function

' R36: if the access IS on the exposed plane, the two orientations
' must agree. Fires only when both are recorded - a missing value
' is pending work and belongs to rule 17, not here.
Private Function R36() As String
    Dim q As String
    q = "SELECT E.Code, 36, "
    q = q & "'R36: access is on the facade but the two orientations differ', "
    q = q & "'One of the two is mistaken, or Access_Plane is not Facade' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Access_Plane='Facade' "
    q = q & "AND E.Portal_Orientation Is Not Null "
    q = q & "AND E.Facade_Orientation Is Not Null "
    q = q & "AND E.Portal_Orientation<>E.Facade_Orientation"
    R36 = q
End Function

Private Function R37() As String
    Dim q As String
    q = "SELECT E.Code, 37, "
    q = q & "'R37: rock art row carries a constructive body number', "
    q = q & "'A painting on bedrock is in no body: clear Body_No' "
    q = q & "FROM ((T_DECORATIONS AS D "
    q = q & "INNER JOIN T_STRUCTURES AS E ON D.ID_Structure=E.ID) "
    q = q & "INNER JOIN L_STRUCT_BODY AS B ON D.ID_Struct_Body=B.ID) "
    q = q & "WHERE B.Level_Type='ROC' AND D.Body_No Is Not Null"
    R37 = q
End Function

' R38a-R38b: Substrate stops being a free choice on the rock art
' half and becomes a VERIFIER. A ROC row with a masonry substrate
' is a row that should have gone to the other half under test 1;
' an overlapping row that is not Mixed contradicts its own
' position. Fire on a recorded mismatch only: NULL is pending
' work, which rule 17 already lists.
Private Function R38a() As String
    Dim q As String
    q = "SELECT E.Code, 38, "
    q = q & "'R38: rock art row whose substrate is not bedrock', "
    q = q & "'Re-check test 1: this is probably pigment of the structure' "
    q = q & "FROM ((T_DECORATIONS AS D "
    q = q & "INNER JOIN T_STRUCTURES AS E ON D.ID_Structure=E.ID) "
    q = q & "INNER JOIN L_STRUCT_BODY AS B ON D.ID_Struct_Body=B.ID) "
    q = q & "WHERE B.Code IN ('ROC','ROC-PER','ROC-PAN') "
    q = q & "AND D.Substrate Is Not Null AND D.Substrate<>'Bedrock'"
    R38a = q
End Function

Private Function R38b() As String
    Dim q As String
    q = "SELECT E.Code, 38, "
    q = q & "'R38: overlapping row whose substrate is not mixed', "
    q = q & "'An overlapping motif spans fabric and rock by definition' "
    q = q & "FROM ((T_DECORATIONS AS D "
    q = q & "INNER JOIN T_STRUCTURES AS E ON D.ID_Structure=E.ID) "
    q = q & "INNER JOIN L_STRUCT_BODY AS B ON D.ID_Struct_Body=B.ID) "
    q = q & "WHERE B.Code='ROC-OVL' "
    q = q & "AND D.Substrate Is Not Null AND D.Substrate<>'Mixed'"
    R38b = q
End Function

' R39 is test T2 of the G-vs-I criterion, automated. It catches
' exactly the case that worries the recorder most: a razed
' platform coded as a cornice. An interbody cornice needs two
' bodies to sit between - without a second body the projecting
' course is G, not I.
Private Function R39() As String
    Dim q As String
    q = "SELECT E.Code, 39, "
    q = q & "'R39: interbody cornice recorded with fewer than two bodies', "
    q = q & "'A cornice sits BETWEEN bodies: check whether this is corbelling (G)' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Interbody_Cornice IN (1,2,3) "
    q = q & "AND Nz(E.N_Basal_Bodies,0)+Nz(E.N_Chamber_Bodies,0)<2"
    R39 = q
End Function

' ================================================================
'  v17 - VALIDATION BATTERY, FIFTH PARTIAL (RULES 40-46)
'  Seven branches for six rule numbers: R40 splits, because the
'  mismatch between record class and decoration position can be
'  made in either direction. R41 was retired in v17a with the
'  column it watched; THE OTHERS KEEP THEIR NUMBERS, because
'  they are cited across four documents and closing the gap
'  would invalidate every citation to gain nothing.
'  A REPORT, never a table constraint.
' ================================================================
Private Sub BuildValidationE(db As DAO.Database)
    Dim q As String
    q = R40a()
    q = q & " UNION ALL " & R40b()
    q = q & " UNION ALL " & R42()
    q = q & " UNION ALL " & R43()
    q = q & " UNION ALL " & R44()
    q = q & " UNION ALL " & R45()
    q = q & " UNION ALL " & R46()
    q = q & ";"
    MkQuery db, "QRY_16e_Rules_40_46", q
End Sub

' R40 IS THE RULE THAT WOULD HAVE CAUGHT EA11. A record of class
' Rock art panel has no structure outline to follow, so the three
' contact positions are impossible on it; conversely, Panel is
' reserved for rows hanging off a PR record and must never appear
' on a structure. The corpus held exactly this: a PR record with a
' ROC-PER row, while the description of the Unclassifiable
' typology already named that same structure as its example.
Private Function R40a() As String
    Dim q As String
    q = "SELECT E.Code AS Structure, 40 AS Rule_No, "
    q = q & "'R40: rock art record carrying a structure-outline position' AS Rule_Violated, "
    q = q & "'A panel has no outline to follow: use Panel, or reclassify the record as a structure' AS Action "
    q = q & "FROM (((T_DECORATIONS AS D "
    q = q & "INNER JOIN T_STRUCTURES AS E ON D.ID_Structure=E.ID) "
    q = q & "INNER JOIN L_STRUCT_BODY AS B ON D.ID_Struct_Body=B.ID) "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID) "
    q = q & "WHERE T.Record_Class='Rock art panel' "
    q = q & "AND B.Code IN ('ROC-OVL','ROC-PER','ROC')"
    R40a = q
End Function

Private Function R40b() As String
    Dim q As String
    q = "SELECT E.Code AS Structure, 40 AS Rule_No, "
    q = q & "'R40: structure record carrying the Panel position' AS Rule_Violated, "
    q = q & "'Panel belongs to isolated PR records: give the row a contact position' AS Action "
    q = q & "FROM (((T_DECORATIONS AS D "
    q = q & "INNER JOIN T_STRUCTURES AS E ON D.ID_Structure=E.ID) "
    q = q & "INNER JOIN L_STRUCT_BODY AS B ON D.ID_Struct_Body=B.ID) "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID) "
    q = q & "WHERE T.Record_Class<>'Rock art panel' "
    q = q & "AND B.Code='ROC-PAN'"
    R40b = q
End Function

' R42 IS UNIDIRECTIONAL, exactly like R34. M and the rows are
' independent in one direction - there can be a frieze whose motif
' could not be resolved, and that is an honest record - but a
' CONSTRUCTIVE MOTIF on the facade with M = 0 is a contradiction.
'
' THE LIST IS THE WHOLE ARCHITECTURAL REPERTOIRE BAR TWO.
' Zigzag, stepped, T, inverted T, L, inverted L and square niches
' are all CONSTRUCTIVE motifs - built into the fabric by setting
' the stones in and out - and all of them are read as frieze
' types. Whether they are also painted is a separate question,
' which is why Color exists on the row.
' Painted band and Plain colour field stay OUT: they are applied
' to the surface, not modelled in it, and a painted band across
' the facade is not evidence of a relief frieze.
' Checks OVL AND FFL: M is a treatment of the facade wall face,
' and a frieze may run across it without crossing the portal.
' Calibration on the v16 corpus: 14 rows sit at OVL or FFL with
' a constructive motif; the rule fires on ONE of them, and ten
' more have M still NULL, which is rule 17's business, not this
' one's.
Private Function R42() As String
    Dim q As String
    q = "SELECT E.Code AS Structure, 42 AS Rule_No, "
    q = q & "'R42: band motif on the facade but Relief_Frieze is 0' AS Rule_Violated, "
    q = q & "'Raise M, or reconsider whether the motif is a frieze band' AS Action "
    q = q & "FROM (((T_DECORATIONS AS D "
    q = q & "INNER JOIN T_STRUCTURES AS E ON D.ID_Structure=E.ID) "
    q = q & "INNER JOIN L_STRUCT_BODY AS B ON D.ID_Struct_Body=B.ID) "
    q = q & "INNER JOIN L_DEC_TYPE AS DT ON D.ID_Dec_Type=DT.ID) "
    q = q & "WHERE B.Code IN ('OVL','FFL') AND E.Relief_Frieze=0 "
    q = q & "AND DT.Name IN ('Zigzag','Stepped motif','Square niche',"
    q = q & "'T-shaped niche','T-shaped niche inv.',"
    q = q & "'L-shaped niche','L-shaped niche inv.')"
    R42 = q
End Function

' R43-R44: the two directions R19 did not cover. The three layers
' of cultural material are NOT redundant - presence is answerable
' from 30 m, type needs identification, condition needs to see
' them well - but nothing was watching the aggregate against the
' detail in this direction, nor the condition against the gate.
Private Function R43() As String
    Dim q As String
    q = "SELECT E.Code AS Structure, 43 AS Rule_No, "
    q = q & "'R43: materials declared present but all seven types are 0' AS Rule_Violated, "
    q = q & "'Record at least one type, set the unidentified ones to 9, or lower the aggregate' AS Action "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Cultural_Materials_Present=1 "
    q = q & "AND E.Mat_Textiles=0 AND E.Mat_Wood=0 AND E.Mat_VegFiber=0 "
    q = q & "AND E.Mat_Ceramics=0 AND E.Mat_Fauna=0 AND E.Mat_DeerAntler=0 "
    q = q & "AND E.Mat_Other=0"
    R43 = q
End Function

Private Function R44() As String
    Dim q As String
    q = "SELECT E.Code AS Structure, 44 AS Rule_No, "
    q = q & "'R44: condition of movable remains recorded where none are present' AS Rule_Violated, "
    q = q & "'Legacy of the v16 layout, where the gate and the field sat on different tabs' AS Action "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Cultural_Materials_Present=0 "
    q = q & "AND E.ID_Material_Status Is Not Null"
    R44 = q
End Function

' R45 CHECKS ONE VALUE, NOT TWO. 'Between bodies only' needs two
' bodies to divide between - same shape as R39. But 'Within a
' body' is PERFECTLY POSSIBLE on a single-bodied structure, and
' is in fact the only value such a structure can take when its
' fabric is plural. An earlier draft flagged both and would have
' reported the very case the field exists to record.
Private Function R45() As String
    Dim q As String
    q = "SELECT E.Code AS Structure, 45 AS Rule_No, "
    q = q & "'R45: fabric declared multiple BY BODIES on a single-bodied structure' AS Rule_Violated, "
    q = q & "'Either the body count is wrong, or the plurality is within the one body' AS Action "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Fabric='Between bodies only' "
    q = q & "AND (Nz(E.N_Basal_Bodies,0)+Nz(E.N_Chamber_Bodies,0))<2"
    R45 = q
End Function

' R46 IS WHAT HOLDS UP THE DIVERGENCE / PHASE BOUNDARY. Divergence
' is an observation, a phase is an interpretation - and an
' interpretation with no evidence field filled is exactly what the
' boundary exists to prevent. The four records that declare two
' phases carried no evidence at all when v17 was specified.
Private Function R46() As String
    Dim q As String
    q = "SELECT E.Code AS Structure, 46 AS Rule_No, "
    q = q & "'R46: two or more construction phases claimed with no phase evidence' AS Rule_Violated, "
    q = q & "'A phase is an interpretation: record what sustains it, or lower the count' AS Action "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Construction_Phases>=2 "
    q = q & "AND (E.Phase_Evidence Is Null Or E.Phase_Evidence='ND')"
    R46 = q
End Function

' ================================================================
'  RULES 47-56 (v18 + v19). Every rule watches a window some other part
'  of the delta opened: the battery is the corpus-level mirror of
'  the form gating, as always.
' ================================================================
Private Sub BuildValidationF(db As DAO.Database)
    Dim q As String
    q = R47()
    q = q & " UNION ALL " & R48()
    q = q & " UNION ALL " & R49()
    q = q & " UNION ALL " & R51()
    q = q & " UNION ALL " & R52a()
    q = q & " UNION ALL " & R52b()
    q = q & " UNION ALL " & R53()
    q = q & " UNION ALL " & R54()
    q = q & " UNION ALL " & R55()
    q = q & " UNION ALL " & R56()
    q = q & ";"
    MkQuery db, "QRY_16f_Rules_47_56", q
End Sub

' v20: the seventh partial. Rules 57-60 get their own UNION
' rather than growing 16f past what JET reliably tolerates -
' an implementation decision the delta left open on purpose.
' v21: rules 61-62 join it (six branches keep it well under
' the 16f ceiling). v22: rules 63-64 (eight branches, still
' comfortable).
Private Sub BuildValidationG(db As DAO.Database)
    Dim q As String
    q = R57()
    q = q & " UNION ALL " & R58()
    q = q & " UNION ALL " & R59()
    q = q & " UNION ALL " & R60()
    q = q & " UNION ALL " & R61()
    q = q & " UNION ALL " & R62()
    q = q & " UNION ALL " & R63()
    q = q & " UNION ALL " & R64()
    q = q & ";"
    MkQuery db, "QRY_16g_Rules_57_64", q
End Sub

' R61 (v21, bloc 2): masonry detail carrying a value while the
' fabric declaration denies or cannot see any masonry. Mirror
' of the plaster/pigment presence rules (R24, R25). A NULL
' declaration is pending work and belongs to rule 17, not here.
Private Function R61() As String
    Dim q As String
    q = "SELECT E.Code, 61, "
    q = q & "'R61: masonry detail recorded but Masonry_Present is 0 or 9', "
    q = q & "'Clear the detail fields, or set Masonry_Present to 1' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Masonry_Present IN (0,9) "
    q = q & "AND (E.Masonry_Quality Is Not Null Or E.Masonry_Type Is Not Null "
    q = q & "Or E.Stone_Format Is Not Null Or E.Stone_Format_Secondary Is Not Null "
    q = q & "Or E.Stone_Working Is Not Null Or E.Mortar_Present Is Not Null "
    q = q & "Or E.Mortar_Type Is Not Null Or E.Chinking_Stones Is Not Null "
    q = q & "Or E.Fabric Is Not Null)"
    R61 = q
End Function

' R62 (v21, bloc 1): a platform whose three supports are all
' CONFIRMED absent is not a platform - it is the top of the
' mass, and Z is reserved for the surface of the cantilevered
' system. Nines do NOT fire it: a surface with unresolvable
' supports is still a platform (v18 C1, preserved intact).
Private Function R62() As String
    Dim q As String
    q = "SELECT E.Code, 62, "
    q = q & "'R62: platform system open but all three supports are confirmed absent', "
    q = q & "'The top of the mass is not a platform: set Sys_Platform to Absent (the finish goes to R), or revise E/F/G' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Sys_Platform Like 'Present*' "
    q = q & "AND E.Timber_Brackets=0 AND E.Transverse_Beams=0 "
    q = q & "AND E.Corbelled_Courses=0"
    R62 = q
End Function

' R63 (v22, bloc B): metric detail carrying a value while the
' availability declaration says there is none. Mirror of R61
' on the metric gatekeeper; NULL is legitimate here and is
' NOT rule-17 work (the field stays out of the arrays).
Private Function R63() As String
    Dim q As String
    q = "SELECT E.Code, 63, "
    q = q & "'R63: metric field carries a value but Metrics_Available is 0 or 9', "
    q = q & "'Clear the metric fields, or set Metrics_Available to 1' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Metrics_Available IN (0,9) "
    q = q & "AND (E.Length_m Is Not Null Or E.Width_m Is Not Null "
    q = q & "Or E.Height_m Is Not Null Or E.Height_Above_Base_m Is Not Null "
    q = q & "Or E.Dim_Method Is Not Null Or E.Opening_Width_m Is Not Null "
    q = q & "Or E.Opening_Height_m Is Not Null Or E.Support_Width_m Is Not Null "
    q = q & "Or E.Support_Depth_m Is Not Null Or E.Area_m2 Is Not Null "
    q = q & "Or E.Volume_m3 Is Not Null Or E.ID_Vol_Method Is Not Null "
    q = q & "Or E.Vol_Notes Is Not Null)"
    R63 = q
End Function

' R64 (v22, bloc C): centuries recorded with the C14 flag
' down. The project dates by radiocarbon only (explicit
' reversal of the always-editable-centuries decision), so a
' century without a C14 behind it has no source. Guards the
' datasheet side; the form gates the fields.
Private Function R64() As String
    Dim q As String
    q = "SELECT E.Code, 64, "
    q = q & "'R64: chronology centuries recorded but the C14 flag is down', "
    q = q & "'Centuries come from radiocarbon only: tick C14 (with its T_DATING row), or clear them' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.C14=0 "
    q = q & "AND (E.Chrono_Start_Cent Is Not Null Or E.Chrono_End_Cent Is Not Null)"
    R64 = q
End Function

' R57 (v20, bloc B): the concordance the EA46 case broke
' silently - a code claiming one sector while the combo points
' to another. Sectors without a code (none, by construction)
' are skipped.
Private Function R57() As String
    Dim q As String
    q = "SELECT E.Code, 57, "
    q = q & "'R57: structure code does not start with the code of its sector', "
    q = q & "'Fix the code, or fix the sector' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID "
    q = q & "WHERE SC.Sector_Code Is Not Null "
    q = q & "AND Left(E.Code, Len(SC.Sector_Code) + 1) <> SC.Sector_Code & '-'"
    R57 = q
End Function

' R58 (v20, bloc C): a present crown wants its form recorded.
' Worklist semantics - ND is a complete answer and the branch
' can be emptied (the R9 lesson, built in from birth). Attested
' crowns (3) are deliberately outside: the form MAY be
' attestable from fallen slabs, but it is not demandable.
Private Function R58() As String
    Dim q As String
    q = "SELECT E.Code, 58, "
    q = q & "'R58: upper crown present but its format is not recorded', "
    q = q & "'Closing panel, projecting course, flush course - or ND' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Upper_Crown IN (1,2) "
    q = q & "AND E.Upper_Crown_Format Is Null"
    R58 = q
End Function

' R59 (v20, bloc E): the portal-adjacent landscape fields have
' no subject without a chamber. The form gates them by the
' counter; this guards datasheet edits.
Private Function R59() As String
    Dim q As String
    q = "SELECT E.Code, 59, "
    q = q & "'R59: portal-adjacent field carries a value but the chamber body count is 0', "
    q = q & "'Without a chamber there is no access: clear the field, or fix the count' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.N_Chamber_Bodies=0 "
    q = q & "AND (E.Access_Plane Is Not Null Or E.Portal_Orientation Is Not Null Or E.Portal_Position Is Not Null Or (E.Recessed_Frame Is Not Null And E.Recessed_Frame<>0))"
    R59 = q
End Function

' R60 (v20, bloc G): two or more phases claimed with no
' evidence recorded. ND is legal; the branch can be emptied.
Private Function R60() As String
    Dim q As String
    q = "SELECT E.Code, 60, "
    q = q & "'R60: multiple construction phases claimed but no phase evidence recorded', "
    q = q & "'Record the evidence - ND is a complete answer' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Construction_Phases>=2 "
    q = q & "AND E.Phase_Evidence Is Null"
    R60 = q
End Function

' R55 (v19, delta A1) is the shape of 49 on the second
' position-resolution qualifier: a reveal claimed outside its
' window carries a padding value promoted to a claim.
Private Function R55() As String
    Dim q As String
    q = "SELECT E.Code, 55, "
    q = q & "'R55: fabric reveal claimed outside its window (Jambs at 0 or 2)', "
    q = q & "'Either O is 0 or 2, or the claim must come down' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Jamb_Fabric_Reveal=1 "
    q = q & "AND (E.Jambs Is Null Or E.Jambs NOT IN (0,2))"
    R55 = q
End Function

' R56 (v19, delta B3): if the corbels supported a platform, the
' system is present or attested - a support role under an
' Absent or Not applicable platform is a contradiction. The
' reverse direction is already R7.
Private Function R56() As String
    Dim q As String
    q = "SELECT E.Code, 56, "
    q = q & "'R56: bracket role claims platform support but Sys_Platform denies the platform', "
    q = q & "'Raise the system to present / attested, or correct the role' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Timber_Bracket_Role IN ('Platform support','Both') "
    q = q & "AND E.Sys_Platform IN ('Absent','Not applicable')"
    R56 = q
End Function

' R47 (delta A1): a recessed frame qualifies the facade plane, and
' with N_Chamber_Bodies=0 there is no facade to qualify. Derivable,
' so the form gates it and this rule catches table-level entry.
Private Function R47() As String
    Dim q As String
    q = "SELECT E.Code AS Structure, 47 AS Rule_No, "
    q = q & "'R47: recessed frame recorded with no chamber body (no facade)' AS Rule_Violated, "
    q = q & "'Raise N_Chamber_Bodies, or lower the frame to its padding 0' AS Action "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.N_Chamber_Bodies=0 AND E.Recessed_Frame=1"
    R47 = q
End Function

' R48 (delta A3): a ledge terrace has neither facade nor portal.
' Portal_Position tolerates its own NA value, which is padding.
Private Function R48() As String
    Dim q As String
    q = "SELECT E.Code, 48, "
    q = q & "'R48: EA-TER carrying facade or portal data', "
    q = q & "'A ledge terrace has neither facade nor portal: clear the fields, or reclassify the record' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID "
    q = q & "WHERE T.Name Like 'EA-TER*' "
    q = q & "AND (E.Facade_Orientation Is Not Null Or E.Access_Plane Is Not Null "
    q = q & "Or E.Portal_Orientation Is Not Null Or (E.Portal_Position Is Not Null And E.Portal_Position<>'NA') "
    q = q & "Or E.Recessed_Frame=1)"
    R48 = q
End Function

' R49 (delta A4): the coincidence claim only makes sense inside its
' window - no differentiated sill, and a cornice with entity.
Private Function R49() As String
    Dim q As String
    q = "SELECT E.Code, 49, "
    q = q & "'R49: cornice-as-sill claimed outside its window (Sill=0 with a real cornice)', "
    q = q & "'Either N is 0 and I has entity, or the claim must come down' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Sill_Coincides_Cornice=1 "
    q = q & "AND (E.Sill<>0 Or E.Sill Is Null "
    q = q & "Or E.Interbody_Cornice=0 Or E.Interbody_Cornice=9 Or E.Interbody_Cornice Is Null)"
    R49 = q
End Function

' R50 RETIRED in v20 (bloc 2) together with the field it
' watched, Interbody_Cornice_Format. Numbering is never
' reused, like 41.

' R51 (delta C1): same shape again - no surface, no material. The
' Isolated-corbel direction stays with R7.
Private Function R51() As String
    Dim q As String
    q = "SELECT E.Code, 51, "
    q = q & "'R51: platform surface material set although Z is absent', "
    q = q & "'Clear the material, or raise Platform_Surface' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Platform_Surface_Material Is Not Null "
    q = q & "AND E.Platform_Surface=0"
    R51 = q
End Function

' R52 (delta D1), two directions like R38 and R40: a Sequential
' edge must name one of its own two structures as the earlier one,
' and a non-sequential edge must name nobody.
Private Function R52a() As String
    Dim q As String
    q = "SELECT A.Code, 52, "
    q = q & "'R52: Sequential connection with no valid earlier structure named', "
    q = q & "'ID_Earlier must name one of the two structures of the pair' "
    q = q & "FROM T_CONNECTIONS AS C "
    q = q & "INNER JOIN T_STRUCTURES AS A ON C.ID_Struct_A=A.ID "
    q = q & "WHERE C.Chrono_Relation='Sequential' "
    q = q & "AND (C.ID_Earlier Is Null "
    q = q & "Or (C.ID_Earlier<>C.ID_Struct_A And C.ID_Earlier<>C.ID_Struct_B))"
    R52a = q
End Function

Private Function R52b() As String
    Dim q As String
    q = "SELECT A.Code, 52, "
    q = q & "'R52: non-sequential connection naming an earlier structure', "
    q = q & "'Only Sequential carries a direction: clear ID_Earlier, or raise the relation' "
    q = q & "FROM T_CONNECTIONS AS C "
    q = q & "INNER JOIN T_STRUCTURES AS A ON C.ID_Struct_A=A.ID "
    q = q & "WHERE C.Chrono_Relation<>'Sequential' "
    q = q & "AND C.ID_Earlier Is Not Null"
    R52b = q
End Function

' R53 (delta D2): the mirror of rule 21. A named element with a
' Body or Whole scope is incoherent in the other direction, and 16
' rows of the v17a corpus sat exactly there.
Private Function R53() As String
    Dim q As String
    q = "SELECT E.Code, 53, "
    q = q & "'R53: lost-element row names an element but its scope is not Element', "
    q = q & "'Set the scope to Element, or clear the code for a Body / Whole row' "
    q = q & "FROM T_LOST_ELEMENTS AS L "
    q = q & "INNER JOIN T_STRUCTURES AS E ON L.ID_Structure=E.ID "
    q = q & "WHERE L.Element_Code Is Not Null AND L.Evidence_Scope<>'Element'"
    R53 = q
End Function

' R54 (delta E1): natural typologies carry their own support form.
' ND stays legal everywhere: it is a doubt, not a contradiction.
Private Function R54() As String
    Dim q As String
    q = "SELECT E.Code, 54, "
    q = q & "'R54: ' & T.Name & ' recorded on support ' & SU.Name, "
    q = q & "'The typology implies its support form: reconcile the two fields' "
    q = q & "FROM ((T_STRUCTURES AS E "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID) "
    q = q & "INNER JOIN L_SUPPORT AS SU ON E.ID_Support=SU.ID) "
    q = q & "WHERE (T.Name Like 'NIX*' AND SU.Name<>'Natural niche (<1m2)' AND SU.Name<>'ND') "
    q = q & "OR (T.Name Like 'CAV*' AND SU.Name Not Like '*cavity*' AND SU.Name<>'ND') "
    q = q & "OR (T.Name Like 'PR *' AND SU.Name<>'Rock surface' AND SU.Name<>'ND')"
    R54 = q
End Function

' ================================================================
'  v17 - QRY_20_ROCKART_SPAN
'  The coverage labels DERIVED from the four segment fields. They
'  are computed and never stored, so that a band with a 9 on one
'  segment appears as a CANDIDATE rather than being asserted: the
'  distinction between an inverted U and a badly preserved closed
'  ring is the whole reason the segments are four fields and not
'  one enumerated value.
' ================================================================
Private Sub BuildSpanQuery(db As DAO.Database)
    Dim q As String
    q = "SELECT E.Code AS Structure, D.ID AS Row_ID, "
    q = q & "B.Name AS Position, DT.Name AS Dec_Type, "
    q = q & "D.Span_Left, D.Span_Above, D.Span_Right, D.Span_Below, "
    q = q & "IIF(D.Span_Left=1 AND D.Span_Above=1 AND D.Span_Right=1 AND D.Span_Below=1,'Surrounding',"
    q = q & "IIF(D.Span_Left=1 AND D.Span_Above=1 AND D.Span_Right=1,'Inverted U',"
    q = q & "IIF(D.Span_Left=1 AND D.Span_Right=1,'Flanking',"
    q = q & "IIF(D.Span_Left=1 OR D.Span_Above=1 OR D.Span_Right=1 OR D.Span_Below=1,'Partial',"
    q = q & "'Not recorded')))) AS Coverage, "
    q = q & "IIF(D.Span_Left=9 OR D.Span_Above=9 OR D.Span_Right=9 OR D.Span_Below=9,'Yes','') AS Under_Read, "
    q = q & "D.Color, D.Notes "
    q = q & "FROM (((T_DECORATIONS AS D "
    q = q & "INNER JOIN T_STRUCTURES AS E ON D.ID_Structure=E.ID) "
    q = q & "INNER JOIN L_STRUCT_BODY AS B ON D.ID_Struct_Body=B.ID) "
    q = q & "LEFT JOIN L_DEC_TYPE AS DT ON D.ID_Dec_Type=DT.ID) "
    q = q & "WHERE B.Level_Type='ROC' "
    q = q & "ORDER BY E.Code, D.ID;"
    MkQuery db, "QRY_20_RockArt_Span", q
End Sub

' ================================================================
'  v17 - QRY_21_V17_REVIEW
'  The worklist of the v16->v17 transfer. Every branch exists
'  because something arrives NULL by the conservative principle,
'  or because a row cannot be corrected automatically without
'  asserting a judgement nobody made.
' ================================================================
Private Sub BuildV17ReviewQuery(db As DAO.Database)
    Dim q As String
    ' The 7 rows that carried Position_Relative: the declared
    ' segment came across as 1, the other three are NULL by the
    ' conservative principle and have to be read off the notes.
    q = "SELECT E.Code AS Structure, 'Segment coverage' AS Review_Item, "
    q = q & "'Rock art row with no segment recorded: read the outline off the photograph' AS Reason "
    q = q & "FROM ((T_DECORATIONS AS D "
    q = q & "INNER JOIN T_STRUCTURES AS E ON D.ID_Structure=E.ID) "
    q = q & "INNER JOIN L_STRUCT_BODY AS B ON D.ID_Struct_Body=B.ID) "
    q = q & "WHERE B.Level_Type='ROC' AND B.Code<>'ROC-PAN' "
    q = q & "AND D.Span_Left Is Null AND D.Span_Above Is Null "
    q = q & "AND D.Span_Right Is Null AND D.Span_Below Is Null "
    ' The orphan rows: no position at all, therefore invisible in
    ' both halves of tab 4.Dec under the v16 INNER JOIN. 16 of 56
    ' in the corpus. Listed here so they cannot hide again.
    q = q & "UNION ALL SELECT E.Code, 'Decoration row with no position', "
    q = q & "'Invisible in both halves of tab 4.Dec until a position is assigned' "
    q = q & "FROM T_DECORATIONS AS D "
    q = q & "INNER JOIN T_STRUCTURES AS E ON D.ID_Structure=E.ID "
    q = q & "WHERE D.ID_Struct_Body Is Null "
    ' Connections: the chronology changed meaning, so every row
    ' that carries a direction has to be re-read once.
    q = q & "UNION ALL SELECT E.Code, 'Connection chronology', "
    q = q & "'Chrono_Relation now names which structure is earlier: re-read against the normalised order' "
    q = q & "FROM T_CONNECTIONS AS C "
    q = q & "INNER JOIN T_STRUCTURES AS E ON C.ID_Struct_A=E.ID "
    q = q & "WHERE C.Chrono_Relation Is Not Null "
    q = q & "AND C.Chrono_Relation<>'Undetermined' "
    ' The counter that replaces T_BODIES only counts if it is
    ' filled. Restricted to multi-bodied records: on a single body
    ' the field is disabled and there is nothing to compare.
    q = q & "UNION ALL SELECT E.Code, 'Fabric divergence', "
    q = q & "'New v17 field on a multi-bodied structure: this is the figure that decides T_BODIES' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Fabric Is Null "
    q = q & "AND (Nz(E.N_Basal_Bodies,0)+Nz(E.N_Chamber_Bodies,0))>=2 "
    ' Phase evidence: without it the divergence/phase boundary
    ' cannot be sustained. Duplicates R46 on purpose - the battery
    ' reports an error, this is a worklist.
    q = q & "UNION ALL SELECT E.Code, 'Phase evidence', "
    q = q & "'Two phases claimed with nothing recorded that sustains them' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Construction_Phases>=2 "
    q = q & "AND (E.Phase_Evidence Is Null Or E.Phase_Evidence='ND') "
    q = q & "ORDER BY Structure, Review_Item;"
    MkQuery db, "QRY_21_V17_Review", q
End Sub

' ================================================================
'  v18 - QRY_23_V18_REVIEW
'  The worklist of the v17a->v18 transfer. Every branch exists
'  because a padding value must become an observation, or because
'  a new field opens a question on rows recorded before it
'  existed. Empty on a fresh build, by construction.
' ================================================================
Private Sub BuildV18ReviewQuery(db As DAO.Database)
    Dim q As String
    ' Z arrives NULL wherever the platform system is open: the
    ' surface has to be read off the record before the A-Z matrix
    ' can use the row.
    q = "SELECT E.Code AS Structure, 'Platform surface (Z)' AS Review_Item, "
    q = q & "'New v18 element on an open platform system: observe and record it' AS Reason "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Platform_Surface Is Null "
    q = q & "AND (E.Sys_Platform Like 'Present*' Or E.Sys_Platform='Attested lost') "
    ' D ungated: under a closed Sys_Base its 0 was padding, and
    ' padding asserts nothing. Now that the field is permanently
    ' active every 0 is read as an observation, so these rows
    ' need the judgement actually made once.
    q = q & "UNION ALL SELECT E.Code, 'Tie walls (D) padding', "
    q = q & "'D was ungated in v18: this 0 was padding under a closed Sys_Base - confirm it as observed, or set 9' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Sys_Base IN ('Absent','Not applicable') "
    q = q & "AND E.Tie_Walls=0 "
    ' The coincidence window: candidates recorded before the
    ' field existed. The 0 they carry is the migration default,
    ' not a judgement.
    q = q & "UNION ALL SELECT E.Code, 'Cornice as sill', "
    q = q & "'Sill=0 with a real cornice: does the cornice serve as the sill? Confirm 0, or raise to 1 / 9' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Sill=0 AND E.Interbody_Cornice IN (1,2,3) "
    q = q & "AND E.Sill_Coincides_Cornice=0 "
    ' Scope incoherences: duplicates R53 on purpose - the battery
    ' reports an error, this is a worklist.
    q = q & "UNION ALL SELECT E.Code, 'Evidence scope', "
    q = q & "'Element_Code with a Body / Whole scope: set the scope to Element, or clear the code' "
    q = q & "FROM T_LOST_ELEMENTS AS L "
    q = q & "INNER JOIN T_STRUCTURES AS E ON L.ID_Structure=E.ID "
    q = q & "WHERE L.Element_Code Is Not Null AND L.Evidence_Scope<>'Element' "
    ' Directed edges that arrived from the positional convention:
    ' the migration names ID_Earlier automatically, but any row
    ' still Sequential without a name needs the joint re-read.
    q = q & "UNION ALL SELECT A.Code, 'Connection direction', "
    q = q & "'Sequential with no earlier structure named: re-read the joint and set the Anterior field' "
    q = q & "FROM T_CONNECTIONS AS C "
    q = q & "INNER JOIN T_STRUCTURES AS A ON C.ID_Struct_A=A.ID "
    q = q & "WHERE C.Chrono_Relation='Sequential' AND C.ID_Earlier Is Null "
    q = q & "ORDER BY Structure, Review_Item;"
    MkQuery db, "QRY_23_V18_Review", q
End Sub

' ================================================================
'  v19 - QRY_24_V19_REVIEW
'  The worklist of the v18->v19 migration. Three branches:
'  the reveal window carrying its migration 0, the D paddings
'  returned to NULL by the patch, and the bracket roles still
'  unrecorded that must go through the three-question tree.
'  Empty on a fresh build, by construction.
' ================================================================
Private Sub BuildV19ReviewQuery(db As DAO.Database)
    Dim q As String
    ' The reveal candidates: rows inside the window recorded
    ' before the field existed. Their 0 is the migration
    ' default, not yet a judgement. The constant formula
    ' 'jamb: masonry reveal, dressed' in Systems_Notes is the
    ' entry clue (QRY_18_Notes_Review retrieves it).
    q = "SELECT E.Code AS Structure, 'Fabric reveal (window)' AS Review_Item, "
    q = q & "'Jambs at 0 or 2: is the position resolved by a dressed terminal face? Confirm 0, or raise to 1 / 9' AS Reason "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Jambs IN (0,2) "
    q = q & "AND E.Jamb_Fabric_Reveal=0 "
    ' The D paddings the patch returned to NULL: the judgement
    ' the v17 gating skipped has to be actually made once.
    q = q & "UNION ALL SELECT E.Code, 'Transverse wall (D) judgement', "
    q = q & "'This NULL was a v17 padding zero: observe and record D - the 9 is a legitimate answer' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Tie_Walls Is Null "
    q = q & "AND E.Sys_Base IN ('Absent','Not applicable','Not observable') "
    ' The roles still NULL: duplicates R9 on purpose - the
    ' battery reports an error, this is a worklist.
    q = q & "UNION ALL SELECT E.Code, 'Bracket role (three-question tree)', "
    q = q & "'Corbels present with no role: attestation, no link, or ND - the ND is a complete answer' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Timber_Brackets IN (1,2) "
    q = q & "AND E.Timber_Bracket_Role Is Null "
    q = q & "ORDER BY Structure, Review_Item;"
    MkQuery db, "QRY_24_V19_Review", q
End Sub

' ================================================================
'  v21 - QRY_29_V21_REVIEW
'  The worklist of the v20->v21 migration. Two branches: the
'  Z-only platforms produced by the retired terrace criterion
'  (duplicates rule 62 on purpose - the battery reports an
'  error, this is a worklist), and the fabric declarations
'  the patch could not derive (duplicates rule 17 on purpose:
'  this judgement is the one the migration opens). QRY_28 is
'  the worklist navigator, hence the jump in numbering.
'  Empty on a fresh build, by construction.
' ================================================================
Private Sub BuildV21ReviewQuery(db As DAO.Database)
    Dim q As String
    ' The rows the old two-question criterion opened: revert
    ' by hand, never by blind UPDATE - a platform slab resting
    ' on piers would be caught in the net. Expected revert:
    ' Sys_Platform to Absent, Z to padding 0, surface material
    ' and platform function cleared; the finish, if any, is R.
    q = "SELECT E.Code AS Structure, 'Mass top, not platform' AS Review_Item, "
    q = q & "'Old terrace criterion: all three supports confirmed absent - revert Sys_Platform to Absent (the finish goes to R)' AS Reason "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Sys_Platform Like 'Present*' "
    q = q & "AND E.Timber_Brackets=0 AND E.Transverse_Beams=0 "
    q = q & "AND E.Corbelled_Courses=0 "
    ' The fabric declarations left NULL by the patch: every
    ' record with no masonry detail at all needs the judgement
    ' made once. The 9 is a legitimate answer; the 0 is the
    ' answer the EA55 corbel has been waiting for.
    q = q & "UNION ALL SELECT E.Code, 'Masonry judgement', "
    q = q & "'New v21 declaration: is there masonry fabric at all? 0, 1 or 9 - the patch derived the 1s' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Masonry_Present Is Null "
    q = q & "ORDER BY Review_Item, Structure;"
    MkQuery db, "QRY_29_V21_Review", q
End Sub

' ================================================================
'  v20 - QRY_25_V20_REVIEW
'  The worklist of the v19a->v20 migration. Five branches:
'  crown formats to record, amorphous outline bands to confirm
'  against their spans, 12.Extra rows to retype now that the
'  catalogue counts, the terrace two-question review, and the
'  MEN records whose 2.Arq the old gating kept shut.
'  Empty on a fresh build, by construction.
' ================================================================
Private Sub BuildV20ReviewQuery(db As DAO.Database)
    Dim q As String
    q = "SELECT E.Code AS Structure, 'Crown format' AS Review_Item, "
    q = q & "'R present: record the format - closing panel, projecting course, flush course, or ND' AS Reason "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "WHERE E.Upper_Crown IN (1,2) AND E.Upper_Crown_Format Is Null "
    ' The old organic value mixed curved with ill-defined:
    ' each row decides between defined and amorphous with its
    ' spans in front (QRY_27 derives the topology).
    q = q & "UNION ALL SELECT E.Code, 'Outline band stroke', "
    q = q & "'Was organic: confirm amorphous, or move to defined - the old value mixed curvature with definition' "
    q = q & "FROM (T_DECORATIONS AS D "
    q = q & "INNER JOIN T_STRUCTURES AS E ON D.ID_Structure=E.ID) "
    q = q & "INNER JOIN L_DEC_TYPE AS DT ON D.ID_Dec_Type=DT.ID "
    q = q & "WHERE DT.Name='Outline band, amorphous' "
    ' The catalogue candidates hiding in Other: sockets and
    ' benches, targeted by the note stems that identified them.
    q = q & "UNION ALL SELECT E.Code, '12.Extra retype', "
    q = q & "'Candidate socket or bench: retype from Other now that the catalogue counts' "
    q = q & "FROM T_ARCH_FEATURES AS AF "
    q = q & "INNER JOIN T_STRUCTURES AS E ON AF.ID_Structure=E.ID "
    q = q & "WHERE AF.Feature_Code='Other (see Notes)' "
    q = q & "AND (AF.Notes Like '*forat*' Or AF.Notes Like '*perfora*' Or AF.Notes Like '*interf*' Or AF.Notes Like '*banquet*' Or AF.Notes Like '*ancad*') "
    ' v21 (bloc 1): the 'Terrace two questions' branch is
    ' RETIRED with the criterion that motivated it. The top
    ' of a terrace is not Z and does not open Sys_Platform;
    ' the only question left on an N0-only top is the finish
    ' (R), and rule 62 plus QRY_29 handle the rows the old
    ' criterion produced.
    ' The MEN records the old class gating kept out of 2.Arq.
    ' v21 (bloc 2): the branch respects the fabric declaration
    ' - an isolated corbel with Masonry_Present=0 has NOTHING
    ' to fill on this tab and must not sit on the list forever.
    q = q & "UNION ALL SELECT E.Code, 'MEN 2.Arq pass', "
    q = q & "'The tab is open now: fabric declaration first; masonry detail only if there is fabric' "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID "
    q = q & "WHERE T.Record_Class='Structural trace' "
    q = q & "AND (E.Masonry_Present Is Null Or (E.Masonry_Present=1 AND E.Masonry_Quality Is Null)) "
    q = q & "ORDER BY Review_Item, Structure;"
    MkQuery db, "QRY_25_V20_Review", q
End Sub

' ================================================================
'  v20 - QRY_26_NEXT_EA (bloc B)
'  The max-plus-one criterion made visible at entry time: per
'  sector, the highest EA and the next free number. Gaps are
'  campaign history, never free space.
' ================================================================
Private Sub BuildNextEAQuery(db As DAO.Database)
    Dim q As String
    q = "SELECT SC.Sector_Code, Count(E.ID) AS N_Structures, "
    q = q & "Max(Val(Mid(E.Code, InStr(E.Code, '-EA') + 3))) AS Max_EA, "
    q = q & "Max(Val(Mid(E.Code, InStr(E.Code, '-EA') + 3))) + 1 AS Next_EA "
    q = q & "FROM T_STRUCTURES AS E "
    q = q & "INNER JOIN L_SECTORS AS SC ON E.ID_Sector=SC.ID "
    q = q & "GROUP BY SC.Sector_Code "
    q = q & "ORDER BY SC.Sector_Code;"
    MkQuery db, "QRY_26_Next_EA", q
End Sub

' ================================================================
'  v20 - QRY_27_OUTLINE_TOPOLOGY (bloc F)
'  The shape of every outline band DERIVED from its four span
'  fields - no constant formula, no new entry. 1,1,1,0 IS the
'  inverted U; all four is an O; anything partial reads as an
'  arc/C with the covered runs listed. The structure->U /
'  niche->C hypothesis becomes a crosstab of this against
'  typology and site.
' ================================================================
Private Sub BuildOutlineTopologyQuery(db As DAO.Database)
    Dim q As String
    q = "SELECT E.Code AS Structure, T.Name AS Typology, DT.Name AS Band_Type, "
    q = q & "D.Span_Left, D.Span_Above, D.Span_Right, D.Span_Below, "
    q = q & "IIf(Nz(D.Span_Left,0)=1 And Nz(D.Span_Above,0)=1 And Nz(D.Span_Right,0)=1 And Nz(D.Span_Below,0)=1, 'O (full circuit)', "
    q = q & "IIf(Nz(D.Span_Left,0)=1 And Nz(D.Span_Above,0)=1 And Nz(D.Span_Right,0)=1, 'U inverted', "
    q = q & "IIf(Nz(D.Span_Left,0)+Nz(D.Span_Above,0)+Nz(D.Span_Right,0)+Nz(D.Span_Below,0)>=2, 'Arc / C', 'Single run'))) AS Outline_Shape "
    q = q & "FROM ((T_DECORATIONS AS D "
    q = q & "INNER JOIN T_STRUCTURES AS E ON D.ID_Structure=E.ID) "
    q = q & "INNER JOIN L_TYPOLOGY AS T ON E.ID_Typology=T.ID) "
    q = q & "INNER JOIN L_DEC_TYPE AS DT ON D.ID_Dec_Type=DT.ID "
    q = q & "WHERE DT.Name Like 'Outline band*' "
    q = q & "ORDER BY E.Code;"
    MkQuery db, "QRY_27_Outline_Topology", q
End Sub

