# Local continuation — September 12, 2026

## Current verified results

- The user confirmed the 500-row new-text test and the 40-row numeric subset.
- Native Excel loaded `compatibility-suite/test-dictionary-retained.xlsx` with
  40 rows, A=205, K=410 and the new text. Generated Total K evaluated to 24,200.
- A native Save As after widening a column produced
  `compatibility-suite/native-saved-comparison.xlsx`. The model's table metadata,
  IDF streams, dictionaries and hash indexes were unchanged by Excel. The saved
  copy decoded correctly; a desktop close/reopen check remains outstanding.
- The normal single-pass output `compatibility-suite/test-combined-fixed.xlsx`
  has identical logical model files to the successfully loaded staged test,
  excluding archive bookkeeping. It has not separately been opened in Excel.
- Native Excel loaded `compatibility-suite/generated-fixed/supplier-multi-table.xlsx`.
  Plant displayed 10 rows and Defect Type displayed 2 rows. Metrics displayed
  6,145 rows. Generated Defect Quantity evaluated to 56,010,955.
- 47 tests passed after the dictionary and resize changes. The new supplier
  regression then passed with all 15 tests in test_local_regressions.py.
- Use Python 3.12: `../../work/venv312/Scripts/python.exe`. Original Downloads
  files are unchanged. Existing distribution archives are stale.

## Findings and changes

Excel rejects some structures that PBIXRay can decode. Native loading and
calculation are necessary acceptance checks.

The writer now retains existing dictionary IDs, appending new strings instead
of rebuilding compact dictionaries. Eligible single-segment table reductions
retain numeric hash indexes and encoding settings and update row streams and
counts, including hidden `__XL_RowNumber` and `RowNumber` columns. Numeric
lookup dictionaries are retained when the replacement uses existing entries.
Replacing the working small model's numeric hash indexes alone caused a native
load failure. Generic metadata rebuilding also failed, even with dictionaries
and indexes retained; the narrow preservation path succeeds.

Additional fixes include supported bit-width selection, 4 KiB Xpress8 frames,
pandas 3 missing-metadata handling, Windows-compatible temporary paths,
validation before atomic replacement, safe in-place repacking, and inherited
output file permissions. Staging under TemporaryDirectory had made output
unreadable by the interactive Excel user.

## Remaining checks

1. Native save/close/reopen and relationship filtering, especially when a resized
   dimension removes keys still referenced by an unchanged fact table.
2. New numeric dictionary entries, numeric values outside template IDs, string
   dictionary growth across bit-width boundaries, row growth and multi-segment
   models. Generic fallback paths remain experimental, not desktop-verified.
3. Refresh documentation and rebuild the wheel/source distribution once the
   acceptance scope is settled. Do not call the entire release compatible yet.
4. Public release still needs a license decision and fixture redistribution review.

## Desktop testing discipline

The user authorizes autonomous Excel testing with fewer than 5 workbooks open.
Use one workbook at a time and close each completed test. Leave external
connections disabled and do not refresh. Current workbook is the successful
supplier test; only its Date display width was changed to reveal the measure.
Do not overwrite generated fixtures with native saves: use a separately named
comparison copy.

## Expanded authoring work — September 12, 18:10 local

MVP_ACCEPTANCE.md is now the scope checklist. DAX is separate. Expanded MVP is
NOT complete. The earlier template-only completion language is obsolete.

54 full tests passed, then 2 additional public schema tests passed (56 total
expected). Schema support now has clone_column and public add_column(table,
name, values, like=existing_column). It is limited to the existing column's
encoding/type/nullability; simple identifiers currently required. Schema edits
stay in memory. Readers now decode the current in-memory model rather than the
original workbook, so pending schema edits compose correctly.

Native successes on our own 3-row worksheet-only fixture:
- api-native-base-reassembled.xlsx: unchanged model rebuilt by ABF assembler.
- api-native-xml-control.xlsx: all model XML parsed/serialized with minidom.
- api-cloned-column-v5.xlsx: ID copied to CopiedID, values 1/2/3.
- api-two-new-columns-values.xlsx: CopiedID 3/1/2 and CopiedLabel Delta/Beta/Alpha.
- api-numeric-new-values-hierarchy.xlsx: ID 10/20/30; filter list shows 10/20/30.
  Actual filter selection is NOT verified: menu closes when automation clicks.

The clone failure was duplicate RowNumber AttributeRelationship Name ID_,
which differed from AttributeID ID. New relationship names must be unique.
Regression assertion added. Earlier v1/v4 clone outputs fail native load.

lookup.py rebuilds numeric POS_TO_ID alongside HIDX; the manual numeric fixture
proved the mapping patch opens. Reusable implementation and public add_column
output api-public-add-columns.xlsx still need direct native testing.

Current Excel: one workbook, api-numeric-new-values-hierarchy.xlsx, Power Pivot
window open. No edits saved, menu closed. Main window 723444; PP window1905244.
Use sky via node_repl only, one workbook at a time. Fixtures from native authoring
are tests/data/native-authoring-base.xlsx and native-authoring-added-column.xlsx.

Native public add-column acceptance completed: api-public-add-columns.xlsx
loads with Amount 10/20/30 and Category Delta/Beta/Alpha. Save As produced
api-public-add-columns-native-saved.xlsx, which closed and reopened successfully
in Power Pivot with all expected values. Native save preserved the table,
dictionary, IDF and HIDX logical bytes. Full test suite now 58 passed.

## Authoring checkpoint — September 12, 18:40 local
Expanded MVP remains incomplete; keep working. Full suite 62 passed before two
rename tests; schema-only suite now 11 passed (64 full expected).
Native successes added: api-grow-ten-rows-v2.xlsx (10 rows),
api-public-remove-column.xlsx (ID/Amount), api-public-add-table.xlsx (original
3 rows and independent SecondTable 4 rows ID10/20/30/40), api-renamed.xlsx
(Product Catalog tab, Label and ProductID1/2/3). Latest workbook api-renamed.xlsx
is open in Power Pivot, one workbook only. Other native tests are closed.

Generic growth was fixed by locating the physical hidden RowNumber column in
metadata (PBIXRay schema omits it), updating uint32 offset12 of its stream.
ABF assembler now supports VirtualFileGroupAddition to clone dimension,
measuregroup and partition file groups. Public add_table(name, dataframe,
like=table) uses independent cloned storage. remove_column rejects model and
presentation dependencies and removes all physical column files. rename_table
changes display captions while retaining tableID; rename_column clones/removes
and puts renamed column last. Both reject unsupported dependencies. Simple
identifiers still required for physical column/table creation; table captions
can contain spaces. Rename implementation/source and tests are saved.

Next: model creation without caller templates; remove tables; data types and
nulls/boundary cases; relationships edits/removal and native filtering;
hierarchy/sort/visibility/format; connection contract; packaging/docs. Existing
archives and status documents remain stale. No release or complete claim yet.

## New model authoring checkpoint
Full suite 64 passed before table-removal/create additions; schema suite now13
passed. Public create(tables mapping) uses packaged data/authoring-seed.xlsx,
which has a fresh minimal workbook envelope and our synthetic model encoding
seed. Supports nonempty integer/text tables, simple column identifiers. Removes
all starter columns/table and sets requested captions. No caller template or
Excel runtime needed. Public remove_table removes dimension/measuregroup/partition
files; assembler prunes emptied table groups. Dependency-bearing removal rejected.

Native api-created-v2.xlsx SUCCESS: new blank worksheet, Products model table,
ProductID10/20/30/40, CategoryAlpha/Beta/Alpha/Beta. api-created.xlsx failed due
incorrect MIME type. Correct model MIME is
application/vnd.openxmlformats-officedocument.model+data (native source verified).
Repaired failed version was closed without saving. One workbook open currently:
api-created-v2.xlsx with Power Pivot. Rename test closed. Need save/reopen new
creation and remaining milestones. Build seed script copied into scripts/ but
paths still workspace-relative and needs cleanup before release.

## Text growth and settings checkpoint
Full suite 67 passed before latest string-lookup test (68 expected). Added
configure_column(hidden,description,format_string) in public API and schema.
Hidden column verified native, description XML only. FormatString alone does
not change Power Pivot grid formatting. Native currency/decimal references are
saved separately in compatibility-suite; grid additionally uses an annotation:
<Annotation><Name>Format</Name><Value><Format xmlns="" Format="NumberDecimal"
Accuracy="2"/></Value></Annotation>. This is now synthesized for 0.00-style
formats; needs direct native test. Native decimal reference is open in Excel,
one workbook, PP hidden by SaveAs. Name native-decimal-reference.xlsx.

api-settings-text-growth.xlsx failed native load. Rebuilding string POS_TO_ID
and ID_TO_POS fixed api-settings-text-growth-v2.xlsx: 20rows all exact text
Category01..20 (spaces before numbers), ID1..20, hiddenCategory grey. Closed only
after native decimal change and saved as native-decimal-reference.xlsx. Original
v2 file remains unchanged. New lookup.py string_lookup_replacements supports
ASCII casefold collation, at least3retained dictionaryentries, two32bitsegments.
Map Records=dictcount+3; first segmentdictcount, second3. POS_TO_ID sorted IDs
(with optionalnull2) thenzeros. ID_TO_POS [0,count+hasnull,0]+inversepositions.
Exact byte reproduction tested against native3stringseed; larger20nativepassed.
Called only generic string dictionarygrowth in save (preserved paths unchanged).
Needs stronger nullable/collation/filter acceptance and dictionarycasefolddedup.

## Nullable authoring checkpoint
Native api-settings-text-growth-v3.xlsx SUCCESS: ProductID1.00..20.00 with API
format annotation, Categoryhidden,20rows. Closed without saving. Native
api-nullable-diagnostic.xlsx SUCCESS4rows: Number1,blank,3,4; TextAlpha,blank,
New,Beta. One workbook currently open (this nullable diagnostic), PowerPivot.

Fixed output validation for pandas.NA: compare null masks separately and compare
only both-present positions. Semantic nullability now reads regular dimension
KeyColumn NullProcessing=Preserve rather than incorrectly treating HasNulls
(existing data) as a schema restriction. Existing hasnullfallback retained.
Read only immediate dimensionAttributes children; nested relationshipAttribute
nodes have noUsage and initially caused regressions, fixed. create() nowaccepts
nullable integer/string pandas columns. Stringlookup rebuilt whennullpresence
changes, evenwithoutdictionarygrowth. Latest fullsuite69pass expected/poll.
Still need native public nullablecreation (diagnosticusedprivateencodingoverride,
but sameemittedstreams), allnullcases, additionaltypes, allremainingmilestones.

## Signed integer checkpoint
Fullsuite69passed. Nullablediagnostic is still the only openExcel workbook.
Publicnullablecreation nowcoveredtest. Added numeric rebase in generic save:
when non-null values encode IDs below3 (zero/negatives) or aboveuint32, shift
base so minimumordinaryID3, updateXMValueDictionaryBaseId, rebuild HIDX/POS_TO_ID
and segment. Rejectspans>0xfffffffc. PreservepathreturnsNone whenordinaryID<3.
This preventszero beingmistakenfornull(ID2). New signed/nullcreate test passes;
schema suite16passed, full70expected. Native signed/rebase NOTtested yet.

MVP_ACCEPTANCE statuses updated, no completeclaim. Still substantial work:
- Numericfloat/date/bool/currency creation and type conversions; allnull/zero
rowcases, strings Unicode/case-insensitive dictionary normalization, largeHIDX.
- Add tables/columns without like prototypes on existingmodels; publicAPI
currentlycreate() implicitseed handlesinteger/text; identifierescaping still
simpleonly and renamedcolumns moveend. Retainedseed dictionaryentries invisible
in data but mayappearfilterlists unlesslookupsuseonlyactiveIDs (check).
- relationshipedits/removal andnativefiltering; hierarchy/sort; tablemetadata;
existingworkbooknewmodelinsertion; connection/refresh contract.
- save/reopen acceptance, docs, wheelcleaninstall, releasearchives. Don'tstop
at thischeckpoint orcallMVPcomplete. DAXseparate.

### Expanded authoring checkpoint (September 12, 19:17)
- 75 automated tests pass (work/pytest-full75).
- Native api-decimal-diagnostic.xlsx loads [-1.25,0,2.5,3.75].
- Native api-types-diagnostic.xlsx loads three dates Jan 1/2/3 2026, TRUE/FALSE/TRUE, currency -1.25/0/3.4567. Expanded date column and increased currency precision to four places to verify; saved as api-types-native-saved.xlsx, closed (reopen pending).
- Public create now supports floats, datetime (whole-day magnitude), bool, Decimal currency. change_column_type supports those with optional replacement values; precision checked. Nullable boolean decoding uses pandas BooleanDtype in local reader adapter. Full native nullable mixed-types test pending.
- Fixed replace_column after a pending resize to use the staged row count.
- Added schema.import_table and column_import.import_column; add_table and add_column now accept omitted like. Independent own types are created using synthetic backend, metadata/storage transplanted into target.
- api-independent-authoring.xlsx native PASS: Table1 ID1/2/3 LabelAlpha/Beta/Gamma Amount-1.25/0/3.75; added Sales Quantity-1/0/4/7 Amount2.5/0/1.25/3.75. Closed without saving.
- Currently testing api-relationship-added.xlsx. Base api-relationship-base.xlsx has Products(ProductID1/2/3 CategoryAlpha/Beta/Gamma), Sales(ProductID1/2/1/3 Amount10/20/30/40). Added file uses existing add_relationship(Sales.ProductID -> Products.ProductID). Need native filtering, expected category totals Alpha40 Beta20 Gamma40.
- All earlier remaining MVP items still apply; release not complete.
- Relationship v1 and namespace-preserving v2 both fail native load. Clean api-relationship-base.xlsx loads. Created relationship natively using Create Relationship modal (must target modal window directly for dropdown/keyboard; parent activation loses focus). Saved native-relationship-reference.xlsx, closed.
- Extracted relationship diff in work/relationship-diff. Native relationship adds dense foreign-data-ID -> primary-row index stream, XMRelationship raw metadata, reference measure-group dimension, and primary key ColumnFlags|3; foreign Usage1 primary Usage0.
- Added src/excel_model_lab/relationship_storage.py and data/relationship-index.xml (sanitized native prototype). Save now stages values then constructs full relationship storage. v3 generated and currently being tested. Index sample is [-1,-1,-1,0,1,2], encoded uint64 word count1 + bitpacked0xe40 with width2/min-1.
- Added create(..., workbook=existing.xlsx) and package model insertion, preserving other parts. Automated test passed; native insertion still pending. Test source has existing model relationship/content type but removed model so strengthen test later to ordinary workbook too.
- Latest full suite 76 passed (work/pytest-rel76), before any later modifications.
### Relationship authoring checkpoint (September 12, 19:45)
- Fixed native load: Relationships must precede StructureVersion; lookup primary attribute NullProcessing=Error and RowNumber AttributeRelationship Cardinality=One. api-relationship-added-v4.xlsx opened, PivotTable Category x Sum Amount showed Alpha40/Beta20/Gamma40/Total100. Ctrl+S attempted but file seems still blank-sheet envelope; do not assume persisted pivot.
- Native Edit Relationship deactivate on that workbook, saved as native-inactive-reference.xlsx. PivotTable shows100 for every category, Total100; closed. Diff work/inactive-diff confirms ReferenceMeasureGroupDimension removed for inactive; storage retained. API now omits binding for inactive.
- API inactive file opened Power Pivot successfully, currently only workbook open. Main testState window723444, ppState variable STALE after closing earlier v4; new PP needs list_windows selection. API inactive filtering not yet checked.
- Added stored_relationships(), remove_relationship_storage(), existing dense index rebuild when key values or primary row count/order change. Public remove_relationship/set_relationship_active, duplicate/self/ambiguous active path guards. Tests test_relationship_authoring.py five pass. Full earlier suite76passed after rebuild, before removal APIs; full now81expected not yet rerun.
- API generated api-relationship-reactivated.xlsx and api-relationship-removed-pivot.xlsx from native-inactive-reference (preserved PivotTable). Need native open then selected PivotTable Alt+F5 to requery model only, expected reactivated40/20/40 and removed100each. Don't refresh external connections.
- api-relationship-reordered.xlsx generated but Alpha and Gamma both40 so weak test; use asymmetric keys/data for actual filter check.
- Removal currently leaves primary uniqueness flags/null contract (needs relax when last relationship using key removed). Relationship rebuild supports dense index only; retained string dictionaries may orphan case variants; graph chain reference-binding propagation still pending. Multiple active lookup dimensions and native metadata field edits still pending.
- Expanded MVP still not complete; all other acceptance items remain. Continue autonomously, no subagents, <5 workbooks (using1).

### Hierarchy implementation checkpoint (September 12, 21:55)
- Full tests now 84 passed in27.87s, including recent hash bucket-growth/nonfinite numeric dictionary/generator fixes and 3 hierarchy tests.
- Reactivated relationship native filtering40/20/40 accepted, saved and reopened successfully earlier. native-hierarchy-reference.xlsx created from that saved file, ProductHierarchy Category/ProductID; saved/closed.
- Added hierarchy_storage.py and sanitized data/hierarchy-index.xml. Public add_hierarchy/remove_hierarchy; save rebuilds hierarchy indexes for edited source columns. Breadth-first distinct paths, four32-bit traversal columns, extra zero word matches native reference. Tests cover repeated child values under different parents, resize rebuild, removal, invalid names.
- api-hierarchy.xlsx opens Power Pivot without error; Diagram View shows correct two levels. Existing pivot caches old field metadata even Alt+F5, fresh PP PivotTable shows ProductHierarchy and drill-down Alpha -> ProductID1 works. Saved and closed through Ctrl+W Save. Reopen and more complex native hierarchy cases pending.
- Excel main testState723444 currently blank (zero workbooks); ppState9308866 stale after closing. Mainwindow1739x747; SavepromptSave999516.
- Sort reference still missing; prior async question is stale, canceled Sort dialog to continue. Need native sort reference and implementation.
- Remaining expanded MVP items from earlier checkpoint still apply. Release NOT complete. No agents/goals/automation. Continue autonomously, one workbook at a time.

### Sorting/API hardening checkpoint
- Added public set_sort_column(table,column,by|None), sorting.py semantic OrderBy=AttributeKey/OrderByAttributeID and raw OrderByColumn plus POS_TO_ID/ID_TO_POS or numeric HIDX reorder. Save rebuilds when either column has staged changes; detects ambiguous values and circular sort dependencies. Pending create values validated before committing metadata.
- Native api-sort-probe.xlsx metadata-only loaded but stayed alphabetical (saved/closed). api-sort-probe-v2.xlsx also reordered Category dictionary lookup maps and native PivotTable requery showed Gamma/Beta/Alpha (expected). Saved/closed. Public api-sort-public.xlsx generated with same implementation, separate native open pending. Numeric sorting automated pass, native pending.
- Full suite87passed32.07s after sorting+relationship guards+type reference fix. Two additional test_model_insertion.py tests passed (ordinary package, custom connections target); total89expected, not full rerun yet.
- package._new_model_parts respects custom connections relationship target, rejects external/duplicate/conflicting model bindings. More existing workbook native testing pending.
- Relationship key type matching and active bool checks added, casefold matching/duplicate primary checks added; dedicated new guard tests not yet added. change_column_type rejects keys used in relationships before a semantic type change. schema.change_numeric_type now updates all matching det Dimension Attribute KeyColumn data types too.
- Hierarchy XML template name escaped for ampersand; test special names pending. Hierarchy removal currently doesn't reject dependent existing PivotTables (must add guard).
- Current native test: api-large-combined.xlsx from work/large_combined.py. Public create10k rows ID1..10000 Category repeatingGamma/Beta/Alpha Rank1/2/3 AmountID/100, set_sort_column Category Rank, add_hierarchy Catalog Category/ID, add_measure Total Amount SUM('Large'[Amount]) format0.00. Expected Total500050; Gamma166716.67 Beta166650 Alpha166683.33. Opened workbook, clickedManage, latest screenshot shows Loading Data Model bottom right. Need inspect fresh state/windows; main723444, oldppState stale. Exactly one workbook open.
- No task completed claim; expanded MVP remaining items still pending.

### Large-model checkpoint
- Fixed native packed-word padding (floor(count/slots)+1), numeric intrinsic lookup count+3 sentinel records, and exact HIDX hash multiplier314160805. Native overflow chains serialize zero pointers and overflow entries by bucket. Hash index and numeric IDF now byte-for-byte match a native Excel ID1..10000 reference. Synthetic stream fixtures and two regression tests added. Full91passed33.20s.
- api-numbers-10000-v3.xlsx native PP loads all10000rows, savedclosed. Native reference native-numbers-10000.xlsx generated directly in Excel from synthetic ID1..10000, savedclosed.
- api-large-combined-v2.xlsx loads10000rows. Pivot ordinaryCategory sortedGamma/Beta/Alpha and TotalAmount166716.67/166650.00/166683.33, grand500050.00 allPASS. HierarchyCatalog queryFAIL member notfound. Native createhierarchy command exposed member-property name ID reserved. Fixed clone_column internal attribute relationship names with unique underscore suffix (column captions unchanged); regression test+schema23passed10.21s.
- api-large-combined-v3.xlsx newnamingfix loads but hierarchy query stillFAIL. Need native hierarchy reference onthisfixedfile to compare traversal storage; likely ordering/memberID semantics issue. Currentmain723444, pp3541794; exactlyone workbook v3open withfreshblankPivotTable; lastactiondismissedqueryerror. ppstillopenbackground. NextManage,Diagram,createadditionalnativehierarchyCategory/ID andsaveasreference. v2savedclosedwithcorrectordinarycategoryPivot. v3notyet saved. Fullsuiteafternamefixpending; total92expected.
- No completion claim; expandedMVP stillsubstantialremaining. DAXseparate; keeponeworkbook.
- NativeHierarchy1 createdonv3afterpropertynamefix, queriedGamma/Beta/AlphaPASS, savednative-large-hierarchy-reference.xlsx closed. ActualremainingbugwastraversalORDER: Catalog3,4,5 vsnative5,4,3 matchingCategorycustomsort. hierarchy_storage.intrinsic_order nowreadsnativePOS_TO_IDrecords; traversalusespositionmaps. save sortsBEFOREhierarchyrebuild; sort changesupdatehierarchy. Rebuildpreservessemanticdim/cube toretaincaptions/descriptions. Four10k indexstreams EXACT nativebytes. Nativeapi-large-combined-v4.xlsx nowPASS hierarchyqueryGamma/Beta/Alpha,groupedtotals166716.67/166650.00/166683.33,total500050.00; Gamma drill1/4/7 amounts0.01/0.04/0.07. Savedclosed,reopened,collapsedGammaexpandedBeta ->2/5/8amount0.02/0.05/0.08PASS withconnectionsDISABLED. Closedwithoutadditionalchanges; zeroopenworkbooks.
- Full93passed34.02s afterhierarchyorder+propertynamefix. Addedlaterhierarchyremovalpresentationguard+test andfixedtypechangeguardsemanticcomparison(usingrelationship_key_type); subsequenttestsnotyetexecuted94expected. ContinueexpandedMVPremaining, nofinishclaim.

### Hierarchy editing and existing-workbook insertion
- Relationship removal fresh native requery PASS100each/grand100; savedclosed api-relationship-removed-pivot.xlsx.
- Added update_hierarchy(table,name,levels=...,new_name=...) and configure_hierarchy(hidden,description,level_names). Atomic copy, preserve retained level captions, pending values, description/visibility; update uses removal presentation dependency guard. Tests96passed36.91s.
- api-inserted-hierarchy-update.xlsx public create(workbook=ordinary source) preserves A1 Existing worksheet preserved/B1=12345 nativePASS. Public rename Original->Geography and expansion to Region/Category/ID nativePASS: East120/West90/grand210; East C drill IDs4=40,6=60; field-list custom Product number caption visible. Saveclose currentlypending; reopen next.
- Inserted hierarchy workbook saved/reopenedPASS; fresh West drill gives B50/A40 grand210, closedwithoutchanges.
- api-relationship-chain.xlsx nativeFAILload. Removed Sales->Products via API into api-chain-one-link.xlsx nativePASSload. Added same edge in nativePP saved native-relationship-chain-reference.xlsx, closed. Native Sales.det includes transitive Regions binding Materialization=Indirect, IntermediateDimension=Products/RegionID, NO RelationshipID. Added synchronize_indirect_bindings, invoked after save relationship writes and immediate removal; retains direct references, recomputes indirect BFS. Two tests both creation orders/deactivate/reactivate/removePASS; relationship7passed13.45s.
- api-relationship-chain-v2.xlsx with corrected indirect binding nativePASSload and fresh Pivot Regions.Region x Sales.TotalAmount East80 West20 grand100. Currentlyopen, saveclose next; <5workbooks using1. Fullsuite latest96beforechain,98expectednotrerun.
- api-chain-reordered-grown.xlsx native fresh requeryPASS East88 West22 blank55 Total165 after reversing bothlookuproworders, changingSalesAmounts11/22/33/44 andgrowingSaleswithunmatchedFK99 Amount55. Savedclosed. api-chain-inactive.xlsx APIdeactivatedProducts->Regions, native freshrequeryPASS East165West165Total165, savedclosed.
- synchronize_indirect_bindings now refreshes direct-reference attribute lists from currentlookupownattributes (preservesgranularitykey), thenpropagatesindirects. Ensuresnewlookupcolumnsareboundintoexistingfacts. import_column nowcreatesuniqueunderscorememberpropertyname againsttargetexistingnames. Addedregression; full99passed46.64s.
- Generated api-lookup-column-mixed-types.xlsx: fromreorderedgrown native saved,addRegions.BandWestband/Eastband andRank1/2, sortRegionIDbyRank(expected2,1), addhierarchyAreaBand/Region, addindependentTypes4rowsInteger/Decimal/Currency/Flag/Date/Text secondrowallblank. Nativepending, currentlyOpenfiledialog zeroopenworkbooks. Main723444, pp7080868stale. Source creationwork/lookup_mixed_types.py. ExpandedMVPstillNOTcomplete.

### Combined regression, local-source workflow and relationship cleanup checkpoint
- Full automated suite103passed58.69s (work/pytest-expanded103). Native combined api-lookup-column-mixed-types.xlsx FAIL load; isolated api-lookup-only.xlsx PASS load with Band visible. No new field query completed.
- Desktop Sky click failed twice with exact error foreground window did not report a process id, despite list_windows/get_window refresh. Per skill stopped UI actions. Exactly1ownedworkbook open api-lookup-only.xlsx, main723444 PP10160640. Latest ppState screenshot985x836 Home PivotTable23379; stale after pause, refreshbeforeanything. No userstopreported, no lockscreenobserved. Do not claim desktop lock. User async asked to open api-mixed-types-only.xlsx PPManage check4rows secondrowallblank; responsepending.
- Isolation fixtures generated: api-lookup-numeric-sort.xlsx (Band/Rank/numericRegionIDsort2,1), api-lookup-hierarchy.xlsx (same+AreaBandRegion), api-mixed-types-only.xlsx (Base3rows+Types4nullablemixedtypes), api-released-key-null.xlsx (removeSalesProductslinkandProductsProductID1/null/1). All nativepending exceptlookup-onlyload.
- Implemented spec.build_from_spec(template=None,workbook=optionalordinary), new CLI create, buildaddsunknown tables, table/columnmetadata/sorts/hierarchies, CSVread_options; activeboolnowstrictnotbool('false'); JSONrootvalidatedmapping. Addedtest_spec_authoring3testsCSVCLIcreationextension/atomicinvalidactive/ParquetnullabletypesPASS. Installedpyarrow25.0.1intoisolatedvenv viaapprovednetwork. Addedoptionalparquetdependency.
- remove_relationship_storage now releases finallookupkey NullProcessingPreserve/CardinalityMany/ColumnFlags&~3; recomputesaffectedkeyUsage0iflookup/1ifforeign/3otherwise; retainscontractwhileanotherrelationshipuseskey. TestlastrelationshipreleaseexposednewNULLoutsideexistingpackedrange. save now rebuildscolumnlayoutwhenpersegmentnullpresencechangesornewdataIDsoutofrange; generalizesexistingstring-onlyrangefallback. Relationship9testspass20.68s; full103above. Nativecleanupacceptancepending.
- README/STATUS/MVP_ACCEPTANCE corrected stale template-onlycompletionclaims. ExpandedMVPstillNOTCOMPLETE. Remaining earlieritemsstillapply(DAXseparate). Externalbindings/refresh,Unicode,datetimefractional,numericdictionarygrow/largeprecision,emptytables,KPIboundarystillunfinished.
- Builtcurrentwheelinwork/package-check (notfinalrelease), installedintoits/installedsubdirwithsharedexistingdependencies and verifiedimportcomesfrominstalledwheel; createsParquet+CSVtableswithrelationship/hierarchy/bundledresourcesPASS. setuptools84/wheel0.48installedisolatedvenv. Existingoutputs/distarchivesremainstale. No running shellsession/noagents/goals/automation.

## Latest checkpoint — September 13, combined public v6

This section supersedes older current-state notes above. Full suite: 109 passed.
The combined public API output api-combined-public-v6.xlsx passed native load,
Types four-row mixed values/blanks, native Create Hierarchy, API Area hierarchy
query 88/22/55/165, matched and blank drill-down, save/close/reopen, and fresh
West drill-down after reopen. Saved and closed; zero test workbooks open.
Do not regenerate over this native-saved evidence file.

Fixes: tiny numeric POS segments; imported donor storage versions and target
DatabaseID; max-existing-ordinal + 1 for table creation/clone; propagate lookup
RIViolationCount including chains and inactive edges; direct unmatched relation
XMRelationshipIndexDenseDIDs Flags bit 0. Native-created relationship comparison
confirmed flag 1 for direct unmatched edge, flag 0 for propagated-only edge.

Expanded MVP remains incomplete. Next: date/time fractions, numeric dictionary
and range boundaries, empty/all-null and Unicode behavior; remaining native
last-key cleanup and multiple lookup tests; external connections/refresh and
KPI metadata contract; packaging and documentation. DAX stays separate.

## Date-time continuation

Whole-second datetime authoring now uses exact rational encoding and derives
precision from the entire input column (including a blank/midnight first row).
Timezone-aware and subsecond new values fail explicitly. PBIXRay floating day
artifacts are normalized to the exact whole-second value-encoding grid only;
dictionary dates and subsecond storage are not normalized.
Native seconds initially failed with and without nulls because IDs were above
signed 32-bit range. Rebasing above 0x7fffffff fixes the native load. Numeric ID
span guard now uses 0x7ffffffc. api-datetime-seconds-rebased.xlsx loaded with exact
2026-01-01 06:12:13, blank, 2026-01-02 18:23:59. Saved and closed in Excel;
native reopen remains pending. api-datetime-quarter-day.xlsx also native passed
00/06/12/18 hours and was saved/closed. Do not overwrite native evidence files.
Failed api-datetime-seconds.xlsx and api-datetime-seconds-no-null.xlsx are probes,
not deliverables. Zero workbooks currently open. Full 113 passed after signed
boundary fix; two additional first-row precision tests added afterward.

## User testing continuation

User explicitly requested "Let me test for you, you just write the code" after
stopping UI automation. Do not use Computer Use. User confirmed empty-populated
opens; empty-created v1 and v2 fail. v3 question pending: hidden RowNumber IDF
cleared to 16 zero bytes, min/max IDs 2 and allocation/RLE stats zeroed in scratch
work/generate_empty_v3.py. This hidden-row fix is NOT integrated into source.
User-column empty IDFs now use zero RLE and packed-word counts in source.
If v3 passes, integrate hidden empty handling AND restore valid hidden stream
when repopulating; current save expects the old FF primary-RLE marker.

Spaced/Unicode column names and all-null columns native PASS in
api-names-null-columns.xlsx, saved/closed. Do not overwrite native evidence.
api-numeric-dictionary-growth.xlsx native pending: Defect Type IDs 1,3,17.5.
Generic dictionary lookup writer now rebuilds numeric ordering after dictionary
growth or null-presence changes; numeric encoder validates precision/ranges.
Added regression with new -17.5 member sorting ahead of retained IDs, reopen and
second save PASS. Latest full run 120 passed in 80.75s; this extra test passed
separately, giving 121 tests. No running test process.

sources.py adds CsvSource, ParquetSource, SqlSource, TableSource; public tables,
get_table and refresh_sources apply updates atomically after source reads.
Three source tests pass; README documented Python refresh vs native connections.
Native connection contract question remains unanswered. No persisted credentials.

KPI reference: user saw Amount20/30/40 and Target Amount:=100. Asked user to use
Power Pivot > KPIs > New KPI, base Total Amount, goal Target Amount, defaults,
save/close kpi-native-reference-input.xlsx. Confirmation pending. Do not overwrite
input; compare saved native script/cube against regenerated baseline in work.
MVP incomplete; release archives stale. Remaining Unicode data collation, empty
models, KPI metadata, native numeric/last-key tests, broader API dependencies,
connection contract and final packaging. DAX engine remains separate.

### Empty v3 result

User confirmed v3 also FAILS Power Pivot loading. Do not resend or integrate
v3 as a fix. Requested native-empty-reference.xlsx via a blank Power Query:
`= #table(type table [ID = Int64.Type, Label = text], {})`, connection-only plus
Add to Data Model, save and report Power Pivot result. User response pending.
Prepared work/inspect_empty_storage.py and work/empty-v3-storage.json for XML
and binary stream-size comparison. Retained seed lookup structures are one
remaining investigation area, not a confirmed cause. No new candidate generated.

### Blank Query Manual reference received

Located compatibility-suite/Blank Query Manual.xlsx. Read only; never overwrite.
work/native-empty-storage.json and work/native-empty-*.xml contain extraction.
Actual model Query1 contains ONE all-null row, both ID and Label WChar. RowNumber
stream has repeat count1; raw segment map Records1, RIViolationCount0. Therefore
this is not evidence for true empty-table encoding. xl/connections.xml has a
Query - Query1 Mashup connection but model binding is a separate type102
WorksheetConnection_Book1!Query1 range. User loaded worksheet then added that
table to model; this yields a blank row. Asked user for a SECOND blank query,
same typed empty formula, Close & Load To > Only Create Connection + Add this
data to Data Model directly, save Blank Query Direct.xlsx in same directory,
report whether Power Pivot shows zero rows. Awaiting response. No source fix or
new candidate justified yet. No Computer Use actions.

### Direct native reference received; v4 pending

Blank Query Direct.xlsx exists in compatibility-suite and contains true zero-row
Query1 (physical ID Query1_66b433ff-36de-4cbf-b5f8-642801cc46f1). Native regular
empty IDFs are152bytes: primary count16, ALL zero RLE entries, packed count1,
zero word. Hidden RowNumber is144bytes count16, FF repeat0, packed count0.
RLE metadata alloc32 used2 (NOT0), user OthersRLERuns1; Min/Max/OriginalMin2.
Native dictionary objects have LastId2 and NO dictionary/hash/lookup files;
intrinsic hierarchy TypeMaterialization2, position fields-1, blankTableStore.
work/direct-empty-storage.json and work/direct-*.tbl.xml are extracted evidence.

work/generate_empty_v4.py transplants the native three raw empty columns and
their exact stream bytes into v2's table, retaining destination file names and
DataVersion, mapping ID->Order ID and __XL_RowNumber->RowNumber. Unreferenced
old lookup files remain in package. Output api-empty-created-v4.xlsx has passed
ZIP integrity, allmetadataRowCount0, independent read empty with Order ID/Label.
Native user test still needed; NOT integrated as public empty writer fix yet.

Added _unsupported_reason check for native zero-row columns lacking dictionary
and HIDX, avoiding constructor KeyError on None. Such models now inspect/read
but cannot populate yet; error states this clearly. Two parameterized tests;
local_regressions18passed5.90s. Total123 expected, latest full120 + later3.
Native reader exposes __XL_RowNumber as column on direct query; schema filtering
needs follow-up. Do not claim native-empty population/full MVP complete.

### V4 native PASS; integrated empty API

User confirmed api-empty-created-v4.xlsx loads with zero rows. Integrated via
empty_storage.py and bundled data/empty-table.xml (synthetic native raw storage).
save normalizes zero-row tables at the end, keeping allocated152byte user streams
and144byte hidden stream, native blank dictionaries/intrinsic hierarchy metadata.
Unused seed lookup files retained as in passed v4. All zero-row columns require
complete-table replacement (not single-column) to repopulate; constructor now
recognizes zero rows even with stale unused dictionary/hash files present.

populate_empty_table creates a typed donor, transfers ONLY physical raw columns
and stream/index files into the target, preserving semantic schema and original
column properties. Keeps currency semantic via Decimal; checks conversion loss;
renames hidden RowNumber stream for native __XL_RowNumber targets. replace_tables
now rolls back image state too after an invalid later table. get_table omits
__XL_ engine columns. New mixed-type empty/populate + atomic rollback test.

Full123passed83.73s before final empty population fixes; after those, all36schema
tests passed27.99s including new mixed-type test (124total). No running tests.
work/generate_empty_integrated.py generated using PUBLIC API only:
api-empty-integrated.xlsx -> zero rows Order ID/Label;
api-empty-integrated-populated.xlsx -> IDs7/8 Seven/Eight;
api-empty-integrated-cleared.xlsx -> zero rows again.
All independently decoded exact; native user acceptance pending. V4 source and
native references untouched. Next test these sequentially. Remaining expanded MVP
unchanged, KPI native save confirmation still pending, release archives stale.

Follow-up native-direct population uncovered hyphenated physical IDs and retained
old-version RowNumber files. clone_table now permits hyphens and matches identifier
boundaries accordingly; save resolves the RowNumber filename from raw metadata
when multiple versions exist. Blank Query Direct.xlsx now independently populates
IDs9/10 and then grows9/10/11 through public API (outputs in work only), original
reference untouched. Full124passed90.82s before these last two fixes; added native-
style hyphenated ID regression: all37schema tests PASS30.00s after final fixes.
No running test sessions. Total125tests, lastfull124 plus added regression covered
in final37schema run. Integrated public test files await user native acceptance.

### Integrated empty lifecycle native PASS

User: "It all worked as expected." This confirms all three requested files:
api-empty-integrated.xlsx zero rows; api-empty-integrated-populated.xlsx IDs7/8
Seven/Eight; api-empty-integrated-cleared.xlsx zero rows again. Marked README,
STATUS and acceptance matrix. Preserve these files as native acceptance evidence.
Checked kpi-native-reference-input.xlsx: saved script has NO CREATE KPI; no saved
KPI reference found under *kpi* in compatibility-suite. Need user create/save KPI
on Total Amount with Target Amount goal. work/kpi-current-*.xml extracted only.
No native automation used. Expanded MVP still incomplete; no release built.

### KPI reference duplicate-measure error

User screenshot: Power Pivot unhandled exception says Target Amount and Total
Amount cannot be created because same-name measures exist. Advised close without
saving. Inspected source script: exactly one CREATE MEASURE each, not duplicates.
Native supplier-quality script uses exact banner `-- PowerPivot measures command
(do not modify manually) --`; writer substituted `(generated by excelmodel)`.
Hypothesis: native editor classifies commands via this banner and recreates
unrecognized definitions. Changed writer to native banner; generated separate
kpi-native-reference-v2.xlsx by changing only both banners in original reference.
Native test pending; do not claim root cause confirmed. Original left untouched.

### KPI native saved and API implemented

User confirmed corrected marker removed duplicate errors, then saved V2 with KPI.
Native script MdxScript.1.scr.xml extracted work/kpi-saved-*; creates base plus
_Total Amount Goal and _Total Amount Status in SAME annotated command, followed
CREATE KPI CURRENTCUBE. Percentage thresholds40/80, Ascending, Linear, graphic
Traffic Light - Single. Base annotations TypeKpi; helpers SupportMdx, hiddenfalse
(Visible=false), MainObjectName; KPI property SupportKpi. Native changes target
command before base KPI command. Native cube has no additional KPI definition.

Added kpis.py add/list/remove for this linear percentage higher-is-better layout,
bundled synthetic kpi-properties.xml; public model add_kpi/list_kpis/update_kpi/
remove_kpi. Atomic script editing materializes pending measures into candidate,
retains pending rows/relationships. Remove rejects other helper dependencies.
Base format retained; thresholds validation0<=lower<upper<=1; target must exist;
helper name conflicts fail. Other KPI types (absolute goal, descending, icons)
not yet supported. Need stricter unsupported native-layout checks before claiming
general native KPI editing; current remove validates expected three helper names.

Created api-kpi-created.xlsx from public create Amount20/30/40 Total90 Target100
threshold40/80. api-kpi-updated.xlsx from native user reference thresholds50/95.
Native acceptance pending; ask first created file load then rightclick Total
Amount > Edit KPI Settings expect Target Amount40%80%. Do not overwrite nativeV2.
Focused3KPI tests pending result in current tool output; total128 expected.
Expanded MVP still incomplete. NoUI automation; user handles tests.

### KPI settings PASS, measure grid missing target

User confirms api-kpi-created.xlsx loads and Edit KPI Settings target/40/80 are
correct. Screenshot shows only Total Amount in calculation area; nativeV2 has
Target Amount on row1 belowTotalAmount row0. Model Target measure exists and
is User/IsPrivateFalse, not hidden. Native workbook customXml UTF16 contains
Gemini Diagrams (item11) with MeasureGridNodeViewState LayedOuttrue Row1 for
Target Amount, defaultRow0 forTotal. API has no such layout. MeasureGridState
item1 only controls area visibility/height; TableXMLitem2 onlycolumnlayout.
work/fix_kpi_grid.py copies JUST native Diagrams+MeasureGridState and associated
props/rels with unique target itemnumbers, remaps physicalTableID, retains all
API XLSX parts and exact unchanged xl/model/item.data. Generates
api-kpi-created-grid-v2.xlsx; ZIP integrity and model binary identityPASS.
User native test pending: both cells visible, Total thenTarget. This is a
diagnostic layout candidate, NOT yet integrated generic public layout writer.
Need generalize preserving user positions and allocate distinct cells for new
measures after confirmation. Native refs and original API workbook untouched.
Focused KPI3tests passed2.69s earlier; total128 expected, no full rerun yet.

### Copied grid candidate FAILED

User confirms api-kpi-created-grid-v2.xlsx still does NOT show Target Amount.
Do NOT integrate that layout-copy script as a fix. Next controlled candidate:
api-kpi-native-script-check.xlsx built from ORIGINAL api-kpi-created.xlsx by
replacing only script Commands and CalculationProperties with exact nativeV2
elements via DOM import. Retains target script identity/versions and model data,
uses no copied layout. Native definitions/calculation properties differ in command
whitespace, property ordering (Target thenTotal), Format annotations and default
measure hidden property. Isolates the whole working measure representation before
further narrow changes. Public reader sees correct KPI40/80 and data20/30/40.
Native user test pending, ask whether both measure cells now visible. No source
code changed in this turn; root cause still unresolved.

### Native script check FAILED; discovered wrong customXML relationship owner

User confirms native-script-check still missing Target Amount. Inspected native
ZIP relationships: allcustomXML belong to xl/_rels/workbook.xml.rels with targets
../customXml/itemN.xml, NOT _rels/.rels. My prior grid-v2 test incorrectly added
them atpackage root, so was not a valid test of native layout. Explained error.
Corrected work/fix_kpi_grid.py to workbookrels and relative targets; writes NEW
api-kpi-created-grid-v3.xlsx from original api-kpi-created.xlsx, retaining exact
modelbinary. Verified both Diagrams and MeasureGridState resolve from workbook
relationships, ZIP integrity passes. Native user test pending. Do not overwrite
earlier evidence; no productioncode change yet. General layout support awaits
proof from correctly attached display metadata.

### Grid v3 native PASS; generalized layout implementation

User: "That one does work." Correct workbook-owned customXML relationships
resolve missing Target Amount. Implemented measure_layout.py called by repack
for replacement_model: visible User/Kpi measures get individual rows (KPI first
for new layouts); preserves existing Diagrams and positions, appends missing
measures below occupied rows; hidden helpers excluded. Adds MeasureGridState
per physical table ID. CustomXML parts own itemProps relationships; workbook
relationships use ../customXml/itemN.xml. Native template resources stored in
data/measure-diagram.xml and measure-grid.xml (synthetic). Generates inner XML
as escaped CustomContent text rather than CDATA, semantically same XML text.
No wholesale native layout copying in production; preserves unknown diagrams.

Focused4tests PASS3.64s: row0 Total row1 Target, identical layout bytes on second
save, Third getsrow2, existing KPI CRUD/atomic/helperguard tests. Fixed nested
a:Key/Key node editing during tests (only edit leaf keys). Generated publicsave
api-kpi-integrated-layout.xlsx with corrected generic layout. Native pending.
Fullsuite currently session44556, started work/pytest-layout-full129; must poll.
Need final full results and nativeintegrated check before completionclaim.

Layout integration full suites44556 and97804 each had122PASS/7FAIL, allseven on
supplier native existingmetadata handling. Fixed optional Annotation.Value,
skip unrelated opaque customXML parse failures, and parse CustomContent CDATA
alongside text (schema._text excludesCDATA). Targeted numericgrowthPASS then
rerunning all7failed plus4KPI/layout tests in session17319. Must inspect result.
Native integrated api-kpi-integrated-layout.xlsx was generated with generic
separate rows and correctedleafKey hierarchy; native user acceptance stillpending.
Prepared api-kpi-removed.xlsx with both usermeasuresretained, nativepending.

Session17319 complete:11PASS15.59s, including all7previous failures plus4KPI/layout
tests. No running test sessions. All129 covered across full122PASS then corrected
11PASS subset (4overlap). Not a single full129green run yet. Send integrated file
for nativebothcells + EditKPI40/80 check, then update/removal/query milestones.

### Integrated generic layout v1 FAILED; v2 matches passed native diagram

User: Target Amount still absent in api-kpi-integrated-layout.xlsx. Compared
working grid-v3 vs generated diagrams. Differences: missing Columns.Amount key
and viewstate; serialized measure order reversed; explicit Row0 vs native omitted;
CustomContent escapedtext vs CDATA; generated itemProps schemaRefs populated vs
native empty. Adjusted general writer to native representation: all real columns
get column keys/viewstate; command order preserved while KPI gets row0; omitRow0;
use safely splitCDATA; itemProps empty schemaRefs. Generated new
api-kpi-integrated-layout-v2.xlsx using PUBLIC save. Assertion on normalized
diagram XML against working api-kpi-created-grid-v3.xlsx EXACT MATCH passes.
Focused4KPI/layouttestsPASS3.75s; suppliergrowth focused rerun currentresult.
Native user acceptance pending. No definite claim which representation detail
caused failure until native confirms; previous integratedv1 remains failed.

### Integrated layout v2 native PASS; measure editing added

User confirms Target Amount appears and says proceed. Full129PASS95.93s session
82210 finished. Last supplier issue was nested Attribute lackingName: now iterate
direct Dimension.Attributes children only. Native integratedlayoutv2 evidence
must remain untouched.

Added measure_edit.py and public update_measure/remove_measure (atomic via script
candidate). Update expression/format/description/displayfolder; rejectKPIbase
editing untilremoveKPI; remove rejects other measure/KPI helper dependencies and
any existing PivotTable/cache package. add_measure duplicate names now globally
case-insensitive. Measurelayout compares source/newscript to remove only cells
for removed explicit measures, preserves other positions. Added2tests; focused6
measure/KPI/layout testsPASS5.30s after changes (131total). No running processes.

Generated api-kpi-updated-v2.xlsx from accepted integratedlayoutv2 using public
update_measure TargetAmount expression120, update_kpi thresholds50/95. Total90.
Generated api-kpi-removed-v2.xlsx from that file: both usermeasures remain90/120,
KPI icon/settings removed. Native tests pending; ask these two sequentially.
Remaining native query/calculation checks, numericdictionarygrowth, lastkeynull,
Unicodecollation, expandedAPIdependencies, refreshcontract, packaging allpending.

### Native KPI edits PASS; release audit

User says both updated and removed KPI files look good. Thus Target expression120,
KPI thresholds50/95 and KPI removal retaining both measures nativePASS. Updated
acceptance/status/release notes. Full131PASS98.00s session64105 done. No running
sessions. Built wheel work/release-audit-wheels via pipwheel no-deps/no-isolation/
no-cache-dir; initialattemptcache accessdenied fixedbyworkTEMP/no-cache. Installed
into work/release-audit-installed (dependencies shared existingvenv, NOT a clean
dependencyinstall). work/audit_installed_wheel.py asserts importedpackagefromtarget
and tests nativeempty->populate,KPI add/remove,measureedit,layoutresourcesPASS.
Wheel code current but docs changedafterbuild; notfinaldeliverable. Built source
tar.gz work/release-audit-source successfully; MANIFEST now includes nativeidf/hidx
testfixtures andMVP_ACCEPTANCE. Publicdist stillSTALE. NoLICENSE file chosen yet.

Two async native questions pending: api-numeric-dictionary-growth.xlsx DefectType
IDs1/3/17.5; thenclose and api-released-key-null.xlsx Products.ProductID1/blank/1.
Don'trefresh. UserhandlesExcel. Remaining broadMVP stillnotcomplete, no claim.

### September14 removed-key test FAIL, stale lookup identified

User says number2 api-released-key-null.xlsx has loading error. Number1numeric
dictionarygrowth stillunconfirmed. Inspected failed oldfile: Products.ProductID
H$ lookup RecordsPerSegment2 but segmentRecords2,3. This predates implemented
tinylookup fix (must2,2,1), and exceeds nativeequalmapsegmentlimit. Othercolumns
havevalid3/3 or4/3etc. Rebuilt ONLY thatlookup with current
lookup.numeric_lookup_replacements, keptrows1/null/1andremainingrelationship;
output api-released-key-null-v2.xlsx independentreadcorrect. Existing public
lastrelationship-release regressionPASS3.96s. No productioncodechange needed.
Nativev2testpending. Do not markreleaseconstraintnativepassed untiluserconfirms.
Originalfailedfile preserved; don'trepeatrequestoldfilename.

### Released key V2 native PASS

User confirms api-released-key-null-v2.xlsx loads Products.ProductID1/blank/1.
Recorded native final-key constraint release PASS. Failure of oldfile was stale
tinylookupencoding alreadyfixed inwriter. Native numericdictionarygrowth test
stillunconfirmed. Source archive audit confirms all3nativeidf/hidxfixtures plus
empty/KPI/diagram/gridresources and measureedit/layouttests present. No codechange,
no running tests. Full131green remainslatest. Release stillnotfinal; broadUnicode,
connectioncontract, packaging/license and remainingcoverage items persist.

### September15 numeric dictionary native PASS

User explicitly confirms api-numeric-dictionary-growth.xlsx loads without errors
and shows all specified Defect Type ID values1/3/17.5. Numericdictionarygrowth
native load/value check nowPASS. No outstanding replies needed for those lasttwo
tests. Full131automatedPASS remainslatest; broadMVP notcomplete. Next independent
work: Unicode data collation/case behavior, remainingAPIcoverage and release
contract/license/packaging. Do not repeat accepted tests or overwrite evidence.

### September15 continuation: Unicode probe and empty-column API gap

Read Microsoft primary documentation: tabular text collation can merge case
variants; Windows linguistic comparisons differ from Pythoncasefold. No generic
Unicodeproductionfix yet. work/probe_unicode_collation.py monkeypatches ONLY
thatprocess's string_lookup_replacements using Windows CompareStringEx en-US
NORM_IGNORECASE1 ordering, preserving dictionary IDs. Calls generic dictionary
lookupwriter for accents/CJK/emoji. Generated api-unicode-collation-probe.xlsx:
Unicode table6rows ID1..6, Label café/東京/Ångström/naïve/Résumé/😀, Amount1..6,
TotalAmount SUM21, hierarchyLabels(Label/ID). Independent exactreadPASS. Script
stdout usesascii() afterinitialconsoleencodingerror (file/readassert hadpassed).
This remains EXPERIMENTAL Windows-local collation, not portablevalidated support;
publicwriter stillASCIIgrowthguard. Need native load, labels, and ideallyhierarchy
query21 before integration; canonical-equivalence/case behavior stillunresolved.

Fixed add_column on existingemptytable: permit empty pandasSeries/array onlywith
explicitdtype, seedaccess guarded. Added lifecycle test emptyID table -> addLabel
string -> populate7/8 Seven/Eight. All38schema testsPASS33.26s (132total; latest
full131 prior). Generated public api-empty-add-column.xlsx (zero rows existing
OrderID/Label plus New text); api-empty-add-column-populated.xlsx IDs7/8 labels
Seven/Eight NewtextAdded/Column. Nativepending, not yet requested toavoidtoo many
tests. No running processes. Wheel/sourceauditsnowpredate thissmallAPIchange.

### Unicode probe native load/value PASS

User confirms allsixlabels visible withnoerrors in api-unicode-collation-probe.xlsx.
This confirmsload/values ONLY. Sorting/grouping/hierarchy and canonical-equivalence
stillunverified; do not claim complete Unicodecollation or removeASCIIguard yet.
Next nativecheck PivotTable Labels hierarchy inRows, TotalAmount inValues:
café1,東京2,Ångström3,naïve4,Résumé5,😀6, grand21. ExperimentalWindowsCompareStringEx
backend remains scratch-only. Otherpending emptyaddcolumn nativecheck and broad
MVP tasks unchanged. Full131 +38schema tests remainslatestcoverage.

### September 15: Unicode grouping native PASS

User confirms the Labels hierarchy groups all six Unicode labels correctly with
Total Amount values 1 through 6 and grand total 21, without errors, in
api-unicode-collation-probe.xlsx. Native load, exact values and grouping now PASS.
This validates the experimental Windows en-US ordering probe only; public ASCII
restriction remains until locale, canonical-equivalence and portability are
handled. Do not repeat this accepted test. Next prepared native acceptance pair:
api-empty-add-column.xlsx and api-empty-add-column-populated.xlsx.

Full regression after the empty-column API change: 132 passed in 102.87 seconds.
No running test processes. Next user check is the empty-column workbook pair;
release artifacts remain stale and expanded MVP is not yet complete.

### September 15: Explicit relationship cardinality and Python errors

Added cardinality keyword (many-to-one default, one-to-many normalizes endpoints,
one-to-one writes One/One). Relationship dataclass retains cardinality and stored
reader parses it; activation preserves it. Exported RelationshipValidationError
(ValueError subclass). Immediate data checks report duplicate column/value and
one-based data row positions. Save rechecks edited keys, including both ends of
one-to-one, before output replacement even with validate=False. Storage rebuild
also checks. Missing columns still KeyError; blank/empty lookup key explicitly
UnsupportedModelError implementation limitation. Unicode equality still casefold,
NOT complete Excel collation parity. Spec builder forwards cardinality.

Four new regressions: one-to-one save/reopen + invalid changed source preserves
existing destination; immediate atomic failure; one-to-many normalization;
activation retains One/One. Full suite 136 PASS in 105.25s. Native acceptance
PENDING api-one-to-one-relationship.xlsx: Customers ID1/2/3 Names Alpha/Beta/Gamma,
Details ID1/2/3 Amount10/20/30; Total Amount measure. User should check Manage
loads and Pivot Customers.Name rows with Total Amount values10/20/30 total60.
No Excel UI automation; user tests. Do not claim native One/One support yet.
Empty-column pair remains pending, Unicode six-group probe already PASS.
Release archives remain stale, expanded MVP not complete.

### September 21: Unified reader and validation preflight

Added ModelReader and WorkbookModel.open(read_only=True default), exposing all
PBIXRay 0.15.5 public metadata properties, get_table column projection and
categorical strings, iter_table chunking, on_disk/temp_dir and close/context
management. Read-only access does NOT construct the writer. Source-only view;
writer staged-data API remains separate. This delegates XLSX/PBIX/ABF reading,
not PBIX writing. Four native-XLSX reader contract tests pass; no new PBIX/ABF
fixture coverage asserted. Raw upstream decoding semantics retained in reader.

Centralized private decoder/schema accesses in reader_adapter.py, including the
writer's Boolean/date normalization. Dictionary/hash parser imports are still
version-specific upstream modules. Dependency pinned pbixray==0.15.5 to match
verified internals. Added model.validate() -> ValidationReport with structured
issues, valid and raise_for_errors; checks relationship keys/types/topology and
warns on unsupported column encoding. Non-mutating, not full Excel validation.
Two preflight tests pass (changed One/One duplicates and blank-key writer limit).
README stale failed-workbook intro corrected; install points to current source,
not old dist. Reader/preflight docs include scope and lifecycle limits.

User excluded worksheet formula evaluation/caching. PBIX writer, Power Query,
XlsxWriter/CUBE/reporting discussed as roadmap; no implementation authorized as
replacement for unfinished Excel MVP. User away; native one-to-one and empty
add-column pair still pending. No UI automation. Full suite running at checkpoint.

Verification complete: full reader/adapter regression 140 PASS in134.41s, plus
2 new preflight tests PASS3.41s (142 total collected after additions). Built and
installed wheel into work/reader-install-audit and verified import location,
reader, chunked iteration and edit-mode preflight. PASS. Uses existing runtime
dependencies; NOT a fresh dependency installation. Public dist remains stale.
No test processes running.

### September 21 continued: safe formulas and clean install

Replaced semicolon regex matching with quote/comment/identifier-aware measure
statement scanning (formula_statements.py). Measure updates/KPI removal use
accurate expression spans. New expressions reject multiple statements and
unterminated literals/comments; terminator is appended after a newline so a
trailing line comment cannot swallow it. Added3regressions; focused8PASS.
Full development suite145PASS115.42s. Fresh work/clean-release-env installed
source with downloaded declared dependencies (pandas3.0.6/numpy2.5.3/PBIXRay0.15.5).
Installed smoke includes empty lifecycle,KPI,layout,reader,preflight PASS. Then
full145tests against installed package with -o pythonpath= PASS120.21s.

Created checkpoints/2026-09-21 wheel + development source zip + resource inventory.
Not a final release. Source license pending async user choice MIT/Apache/decide
later (no response yet); PBIXRay notice added. STATUS/MVP_ACCEPTANCE rewritten
without stale failures/test counts. README reader/preflight docs current.

Still incomplete: publicUnicode/blanklookup/nativeconnection/date/perspective/
broaderKPI/dependency work. Cannot infer native metadata for perspectives:
Cube has PerspectiveFileList, not inline perspectives. Generated public
api-date-perspective-reference-input.xlsx (CalendarJan1–3, Sales10/20/30 Total60)
for user to mark Date table and create Sales view perspective then save separate
native reference. Instructions NATIVE_REFERENCE_REQUEST.md. Not native tested.
One-to-one and empty-column pair still pending. User does Excel; no UI automation.
No tests/processes left running at checkpoint. No claim MVPcomplete.

### September 21 continued: public Unicode and dependency regressions

Integrated Windows CompareStringEx en-US dictionary ordering and LCMapStringEx
sort keys into public string lookup generation, relationship matching, unknown
member synchronization and sorting. Default Language1033 supported; existing
Language2057 ASCII dictionary behavior retained. Other explicit locales rejected.
Unicode on non-Windows remains unsupported; equivalent distinct spellings in one
dictionary reject rather than silently collapse. Across tables equivalent Unicode
keys match. Public api-unicode-public.xlsx and api-unicode-relationships.xlsx
created; native integrated acceptance PENDING, earlier probe remains accepted.

Regression found en-GB fixture guard regression; fixed and37 focused tests pass.
Added lexical bracket-reference scanning for measure/column removal, ignoring
strings/comments while retaining conservative dependency rejection. Escaped ] in
measure CalculationReference now matches statement identifiers; creation format
strings escape apostrophes consistently with updates. Tests cover real versus
literal dependencies and saved/pending columns. Preflight now checks pending
measure-name conflicts without modifying model. Full development151PASS147.30s,
plus validation3PASS4.16s after one new test (152 collected total). Updated wheel
built and installed clean environment; full installed152-test run still pending
at time of this entry. Wheel hash46a099912e30666ae80096c0352b8368d0c592b1b2032ee7e9699d06ac48cea9.

User permission to resume native Excel testing (maxone workbook) was requested
asynchronously; no answer yet. Their last testing instruction remains user tests,
agent writes code. No Excel automated in this turn. License question also pending.
Missing native date-table/perspective/KPI variation evidence still blocks those
metadata implementations. Do not declare expanded MVP complete.

Installed-wheel verification completed:152PASS144.15s. Independent fresh-install
smoke also PASS (empty lifecycle, KPI/measure/layout/reader/validation). Smoke
script phrase 'native empty/population' refers to native-format resources, NOT
an Excel UI acceptance result. No Excel opened. No test processes left running.

### November 8: native testing resumed with user permission

User explicitly authorized testing again and requested only a few open workbooks.
Operational limit: one test workbook at a time (Excel plus its Power Pivot window
are one workbook). Computer-use skill read; use @oai/sky in node_repl. No COM/UI
fallback. Screen-off coding works while awake; UI requires unlocked session;
monitor-off capture compatibility has not been established.

Opened api-one-to-one-relationship.xlsx in Excel/Power Pivot. Manage loaded with
Customers IDs1/2/3 NamesAlpha/Beta/Gamma and Details amounts10/20/30. Created
PivotTable Name rows with Total Amount; observed10/20/30 grand60 without errors.
Saved separately as native-one-to-one-accepted-2026-11-08.xlsx. Closed Excel and
Power Pivot; list_windows confirmed zero remaining Excel windows. Offline reader
confirmed saved relationship still has cardinality one-to-one. Native reopen not
yet checked. Input candidate remains unchanged.

Tool quirks: text-only state cached indexes sometimes unavailable; screenshot
coordinate clicks reliable. Keyboard F12 did not invoke SaveAs; used File menu.
For modal clicks use modal screenshot ID and modal-relative coordinates. Excel
main window2296572 and PowerPivot10816960 are now CLOSED; select fresh windows.

### November 8 continuation: native acceptance and source control

Git checkpoint b5de57d records the prior implementation and Excel metadata reader.
Current implementation work is uncommitted while integration continues. No remote
publication has occurred. Keep at most two workbooks open; current practice is one.

Native one-to-one accepted workbook was reopened: Pivot Alpha10/Beta20/Gamma30,
grand total60; Power Pivot Details rows and Total60 loaded without errors.

work/date-marked-candidate.xlsx loaded Calendar three rows with Mark as Date Table
checked and Date Table Settings enabled. Minimal API changes worked without
copying Excel's historical version counters. Saved and closed; reopen pending.

work/calculated-column-candidate-v2.xlsx loaded cleanly. Doubled initially showed
unprocessed blanks with a yellow warning. Selecting Calculate produced2/4/6 for
ID1/2/3. Saved and closed; Python ModelReader recovered2/4/6 from the saved file.
Reopen pending. v1 failed due to InferredDatatype namespace/order, corrected in v2.

work/perspectives-candidate.xlsx loaded cleanly. Create and Manage showed Calendar
view with Date selected only; Sales view with Amount and Total Amount selected,
Sales Date excluded. Saved and closed; reopen pending. Hierarchy membership and
multiple hierarchies in one table still need native acceptance.

Windows suite during concurrent changes:334passed/3failed. Linux installed suite:
335passed/5skipped/2failed. Shared failures were old dependency-error expectation
and unrelated calculated dictionary invalidation; both being corrected. Windows
third failure was a resource regeneration race. Full stable rerun still required.

Linked-table native refresh accepted: api-linked-table-bound-2026-11-08.xlsx
initially loaded cached ID10/20/30, then Refresh reported Success3rows transferred,
zero errors. Model became ID1/2/3 Alpha/Beta/Gamma matching worksheet source.
Save requested; close/reopen still pending. Regression rerun41passed covering
hierarchy guards, supplier dictionaries, portable collation and perspective rename.
SQLServer native refresh currently blocked by Windows administrator installation:
LocalDB signed MSI returned1603 (user not administrator). No SQL certification.

Full-range numeric creation native acceptance initially FAILED: both combined
api-full-range-numeric.xlsx and isolated api-wide-double.xlsx cannot load in Power
Pivot. Storage agent found retained string dictionary flags and hash-index header;
v2 candidates correct these and are pending native retest. Do not claim accepted
based on offline numeric roundtrip tests.

Numeric v2 updates: api-wide-double-v2.xlsx loads -1e100/1e-100/1e100; saved,
closed, exact reader readback passed. api-full-range-numeric-v2.xlsx loads all
int64 endpoints, Currency endpoints and doubles; saved/readback preserved exact
Decimal values. Money total0 accepted, but Whole total unexpectedly(blank) while
automatic calculation enabled. Integer aggregation diagnostics underway; not yet
accepted for all boundary calculations. Linux installed wheel now363passed6skipped.
Native saved hashes recorded work/native-accepted-hashes-20261108.json.

## November 8 continuation: checkpoints, pending native test and competitor update

Local Git checkpoints: b5de57d (initial source) and 7e1af4b (expanded API and native acceptance). Later source edits are ongoing.

api-int64-without-sentinel.xlsx loaded min+1, zero and max with Total=0; saved and closed. api-int64-usage-v3.xlsx was opened next. Power Pivot window appeared, but the Windows helper failed twice while reading it with "foreground window did not report a process id". UI input stopped. Aggregation result is UNKNOWN; one test workbook remains open. This is not evidence of workbook load failure or acceptance.

Additional competitor verified at https://github.com/d0nk3yhm/pbix-mcp and https://pypi.org/project/pbix-mcp/0.9.57/: project advertises Python headless PBIX creation and VertiPaq editing. These are project claims, not independently certified here; Excel compatibility not established. Do not claim first/unique Python Data Model writer. PivotWright remains the direct Excel/.NET comparator.

STATUS.md and MVP_ACCEPTANCE.md updated to remove stale claims that perspectives/date marking are unimplemented or one-to-one remains untested. Storage promotion/multiple string pages and indexed calculated invalidation are still active work, not completed milestones.

### Additional November 8 experimental candidates

- api-calculated-sort-pending.xlsx: Amount40/10/30; calculate BandHigh/Low/High and Rank2/1/2; Band sorts by Rank; Pivot Low before High.
- api-calculated-hierarchy-pending.xlsx: same values; Band Detail hierarchy Band then Amount, description Calculated grouping.
- api-promoted-numeric.xlsx: ID1/2/3, Value signed64 minimum+1/0/maximum, Total0.
- api-multipage-strings.xlsx: 2,000 rows; Label0000/1999 prefixes, 600 x characters and 東京; Rows2000.

All four are offline candidates awaiting native acceptance; UI helper failure prevents certification in this continuation. Root ABF mutation tests: 3 passed. Full Windows checkpoint: 391 passed; extracted-sdist tests against installed wheel:63 passed,1 optional-Parquet skip. Later edits require a fresh verification run.

### Conversion and perspective follow-up

Multiple hierarchy selections now follow observed native comma-delimited serialization. Native candidate api-perspective-two-hierarchies.xlsx remains unverified. Installed Office reader/writer behavior and DLL hash are documented in docs/DATE_TABLES_AND_PERSPECTIVES.md; no installed code/files redistributed.

api-type-conversions.xlsx is an offline-verified candidate: Converted has Whole10/20/30, Money/Real1.25/2.5/3.75, FlagTrue/False/True, DateJanuary1/2/3 2026. Measures Whole total60, Money/Real total7.5. Wide has minimum+1,zero,maximum integers converted to exact text. Native acceptance pending.

Intermediate full checks found two precision regressions in numeric promotion (Windows420pass2fail; Linux411pass7skip2fail). They were corrected by enforcing existing precision before promotion. These intermediate totals are not final passing evidence. Review also fixed value conversion with calculated columns and stale dictionary streams when converting to arithmetic storage. Fresh verification follows source freeze.

### Final automated identity checkpoint (November 8)

Windows source suite: **453 passed in 369.82s**. Clean sdist-built wheel installed into a separate environment: targeted storage/schema/dependency/date/perspective/relationship/sort/hierarchy/spec/measure/calculated suite **184 passed, 1 skipped in 290.92s**. The skip is an excluded historical workbook fixture. Installed import was verified from site-packages, with tests extracted from the sdist and no source-tree import path. Artifacts at `work/identity-checkpoint-dist/` are checkpoints, not native-certified releases.

Earlier full installed-wheel checkpoint passed **388 tests, 50 skipped** (historical fixtures deliberately excluded from distribution); earlier Linux installed full suite passed **429 tests, 7 skipped**. Those snapshots predate the last identity/measure rename follow-ups. Final Linux installed identity/storage follow-up passed **203 tests in 105.58s**, evidence `work/linux-identity-final.xml`, tested wheel SHA256 `6d28681576d54e785eb70c1f05a686c497364af66b5c07f3f4669c92b4611065`. These are distinct runs and must not be represented as a full Linux run of every latest change.

`rename_measure` now rewrites unambiguous stored/pending DAX references transactionally and invalidates dependent calculated caches. KPI/perspective/report dependencies, ambiguous references, and unsupported virtual-column/script contexts are rejected before mutation. `api-measure-renamed-2026-11-08.xlsx` is an offline candidate awaiting native acceptance.

A subsequent demonstrated hierarchy ID/display-name mismatch is receiving a narrow resolver follow-up with targeted verification; final package evidence follows below. Native Excel gates remain pending because the UI helper failed; SQL refresh remains blocked by LocalDB installation requiring administrator rights. Do not describe these checkpoints as completed native acceptance.

Final hierarchy identity follow-up: **27 source tests passed**, and **58 tests passed in 89.93s against the freshly installed final wheel**, covering identity/hierarchy/dependency/perspective behavior. Hierarchy methods now accept native display captions while storage uses physical IDs, preserve custom level captions, and reject physical-ID collisions.

Final tested wheel: `work/final-mvp-checkpoint-dist/excel_model_lab-0.3.1rc1-py3-none-any.whl`, SHA256 `97612a2b5bb4e0dbdcd18c06b9e92a211fd6a67f86c3fb80c30ca18abe928553`. Every packaged module/resource was byte-compared with the frozen current source and matched. The matching sdist is rebuilt after this handoff/README documentation update; exact artifact hashes are recorded outside the archive in `work/final-mvp-checkpoint-evidence.json`. No package was published.

### Native access recovered: conversion acceptance and integer boundary failure

Power Pivot access recovered. `api-type-conversions.xlsx` loaded expected typed values and totals60/7.5/7.5, saved, closed and reopened Manage without errors; totals unchanged. Independent Python readback asserted every column, dates, exact Currency Decimal values and exact wide integer text. Saved SHA256 `5b0894019087c2a2e4fe6b7b30451794e9fde731e995b19ac0a13884fd29bb57`. Workbook is closed.

`api-int64-usage-v3.xlsx` now has a confirmed aggregation failure: rows display, Count3, but Total/MIN/MAX are blank. This replaces its earlier unverified status. Storage authoring rejection for signed64 minimum is being implemented; Currency is not rejected without separate evidence. Native results ledger remains incomplete; earlier UI blocker is historical, not the current blanket status.

### Live SQL refresh environment

User authorized a local SQL Server test. Existing Docker engine runs official SQL Server2022 Developer container `excel-model-lab-sql-test`, loopback `127.0.0.1:14333`, synthetic database `ExcelModelLabTest`. Image digest `sha256:cf8f0ab2ebe61feec2eed2629061a07d247f791c2a8ddda74a4ac9f06a291718`. Windows installed SQLOLEDB provider successfully queried3rows. MSOLEDBSQL is not registered. Test-only reader credentials are Windows-user encrypted under `work/sql-refresh/`, not embedded in candidate or logs. Native refresh candidate `api-live-sql-refresh.xlsx` starts cachedID1/2/3,Amounts1/2/3; refresh must yieldID101/202/303,SQL Alpha/Beta/Gamma,Amounts12.5/25/37.5,total75. Excel refresh remains pending. Container currently running; stop with `docker stop excel-model-lab-sql-test` after native testing. This resolves the prior local-server installation blocker without Windows administrator installation.

For unattended authentication, root authorized a disposable fixture at `work/sql-refresh/api-live-sql-credentialed-TEMP.xlsx`, generated through the private low-level binder with a synthetic read-only SQL password. Production `OleDbSource` still rejects password-bearing connection strings. The fixture embeds reversibly encrypted test credentials and must never be distributed. After native testing and workbook close, delete that temporary fixture and protected credential files, and stop/remove the owned test container to invalidate the login. No credentials are printed in logs.

Live SQL candidate loaded cached model values, but Power Pivot Refresh displayed an external-data source Security Warning. The computer-use skill prohibits acting on security/privacy permission requests, so root left the modal untouched. Native SQL refresh is **blocked, not passed**; direct Windows SQLOLEDB query remains passed. Workbook remains open at the trust prompt; retain the running local container and disposable credentialed fixture pending human trust approval. Do not clean them up while the workbook is open.

### Numeric boundary follow-up: range acceptance remains incomplete

Native `api-native-minimum-reference.xlsx` demonstrated blank Currency MIN/MAX with exact storage minimum `-922337203685477.5808`; SUM displayed 0. Its CONVERT-based comparison measures returned #ERROR and establish no integer-boundary behavior. Authoring now rejects both signed64 integer minimum and Currency minimum with explicit Python errors; existing model readers retain exact decoding. Numeric/schema regression suite passed **78 tests in 124.32s**, including 10 boundary/atomicity tests.

Native `api-native-minimum-reference-v2.xlsx` contains Currency minimum-plus-one `-922337203685477.5807`, `0.0001`, and maximum `922337203685477.5807`. MIN displayed `-922337203685478` under General formatting, which is rounded and not exact-value proof; MAX remained blank; SUM displayed `0.0001`. Literal integer measures displayed rounded scientific notation, without proving exact native integer evaluation. Consequently the full remaining authoring range is **not native accepted**. No speculative maximum restriction has been added.

`api-boundary-matrix.xlsx` is prepared but untested in Excel. Maximums/Minimums tables isolate integer/Currency endpoints, one unit inside, and 2047 scaled storage units inside, with explicit nonblank aggregate indicators. Automation stopped at the separate SQL security trust modal before this matrix could be tested. Resume these boundary diagnostics after the human clears that permission prompt.

### Exact numeric boundary acceptance through isolated Excel data API

The prior boundary-matrix gate is resolved through a separate owned Excel instance, using `Model.Initialize` and direct model ADO DAX queries without touching the existing permission dialog. `work/native-boundary-matrix.json` records all 12 queries, ADODB field types, blank flags and invariant exact strings. Exact integer maximum `9223372036854775807` and Currency maximum `922337203685477.5807` returned blank MAX/SUM; one storage unit inward and 2047 units inward returned exact aggregates. Minimum-plus-one and inward variants returned exact MIN/SUM. This isolates the problematic endpoints, rather than rejecting the broad range.

Authoring now accepts integer `-9223372036854775807` through `9223372036854775806` and Currency `-922337203685477.5807` through `922337203685477.5806`, inclusive. The adjacent signed64 endpoints raise explicit Python errors; existing workbook readers continue to decode them exactly. Maximum guards and atomic regression tests are added alongside the previously verified minimum guards.

### Native stored-data batch: all 10 queries, save and reopen passed

The corrected isolated Excel data-API batch in `work/native-stored-results-v2-20261108/` passed all 10 manifest cases from `work/native-stored-cases.json`. Each query was repeated after native save, close and reopen and remained identical. `work/verify_native_stored_readback.py` independently compared every table in every saved workbook against its source, preserving exact values (reader dtype differences allowed). All 10 passed; `exact-readback-summary.json` SHA256 is `11bfd0cbd691649094eb71c76e7aabfaa7c41e1e512143a3456cc89dd5a7adc9`. The native ledger records each source, saved workbook and query-evidence SHA256.

Accepted data roundtrips include zero-row added-column storage, its two-row populated variant, six Unicode rows (accented Latin, Japanese and emoji), both two-row Unicode relationship tables, the 2,000-row multipage string dictionary, three-row numeric promotion, wide doubles and the reopened linked-table workbook. The three-row perspective workbook also passed a table query; this does **not** certify its multiple hierarchy selections. Unicode relationship table queries likewise do not independently certify relationship navigation. The numeric promotion fixture retains an existing maximum endpoint now rejected during authoring; its exact table readback does not override the demonstrated aggregate endpoint failure.

Initial batch errors arose from leaked ADO/COM references in the harness. They are invalid test attempts, not model-load failures. The corrected harness released its references before reopening and passed all cases. One test workbook at a time was used in the owned Excel instance; no existing user workbook was closed.

### Feature-specific native batch: filtering, perspective membership and removal

Verified `work/native-feature-results-20261108/` directly: five cases report identical before/after results and `queried_and_reopened`. Removed measure target is absent from the native measure DMV; retained Total is 90. Unicode relationship grouping returns café=10 and 東京=20, including the decomposed/case-varied fact key. Native hierarchy DMV includes both `Codes` and `Geography` under `Both Hierarchies` / `[Places]` before and after save/reopen, resolving the earlier table-query-only limitation. Unmarked Calendar returns Rows=3 and Days=6. Every input, accepted saved workbook and JSON evidence hash is recorded in the native ledger.

The separate marked-date attempt reports HRESULT `0xC113000A` without a recorded failure stage. Do not label it accepted or infer a model defect until initialization/query/save/reopen stages are isolated. Offline reading returns the expected three dates and Days 1/2/3. Its Date IDF, HIDX and position mapping bytes match the native reference exactly; observed reference differences include coordinated metadata versions and an internal RowNumber encoding hint. Investigation continues without a speculative production change.

### SQL/boundary package checkpoint

After signed64 minimum integer/Currency authoring rejection, installed numeric/native-source/OLEDB/dependency tests passed **84 tests in150.16s** on wheel SHA256 `c6315119311178aaa8d29eeefcc2936b8f727b6222c0ef6d485b66b59e99abd5`. Final README refresh rebuilt wheel SHA256 `512d30fc41e5c22ba8877bcaf17a8f7be0b05ff63393b413f01a4ee80d109f74`; byte comparison confirms only METADATA/RECORD changed, with identical code/resources. Final wheel reinstalled and import verified from site-packages. Artifacts `work/sql-boundary-final-dist/`; evidence `work/sql-boundary-final-evidence.json`. Boundary matrix and native SQL refresh remain unresolved, so these are verification checkpoints, not completion.

Final source verification for this continuation: 464 passed in 386.28 seconds. Installed checkpoint targeted verification: 84 passed. Native SQL refresh remains paused at the external-data security prompt with one workbook open; boundary matrix and other native gates remain pending. No release or feature-complete claim. Final package evidence is in work/sql-boundary-final-evidence.json outside this repository.

SQL refresh follow-up: user cleared external-data trust warning, but Power Pivot now requests a missing/invalid password. Authentication dialog remains untouched. Independent decryption verified the fixture password exactly matches the protected synthetic credential, direct provider query passed, and server log showed no login rejection. Disposable v2 adds only the SpreadsheetML type100 connection `savePassword=1` flag, leaving model bytes identical; this is an untested hypothesis, not a production change or accepted refresh. Native credential metadata investigation continues.

SQL refresh retry: human supplied correct password; native refresh failed with generic connection-verification message and cached data unchanged. Existing Connections shows correct localhost server/database and SQL authentication but blank username/password, Save my password unchecked. Independent provider SUM still75.0000; no server login failure recorded. Diagnostic ring-buffer trace was enabled only after the failed attempt, so lack of later Excel events does not establish where that earlier failure occurred. Native refresh status is now **failed**, not merely blocked; credential metadata investigation remains open.

Read-only installed Office reflection confirmed SQL credential keys `UID`/`User ID`, `PWD`/`Password`, and SavePassword uses `Persist Security Info`; supplied names match. Backend `IsPasswordRemoved` maps `DataSource.ConnectionStringSecurity`, while `IsDataSourcePasswordProvided` checks an in-memory `secureConnectionStringCollection` keyed by source ID. Persisted encrypted text alone therefore does not establish native session credential state. The disposable v2 package flag remains unverified; no production XML change is justified by this evidence alone.

Native SQL refresh breakthrough: supported Excel COM accepted complete synthetic OLEDBConnection.Connection credentials and returned an actual engine error: partition missing valid data source view. Adding Cube.Source DataSourceViewBinding in disposable v3 fixed refresh. Isolated Excel refreshed/saved/closed successfully. Reopen native DAX through direct ADOConnection.Execute returned Rows3,Amount Total75,MinID101,MaxID303; Python readback matched all rows. Hash `e38d667e975022898316a0eb809f5ffb26de09e24d56b28d3fc6778fdc6f2c0d`. This validates the experimental single-source fix with explicit session credentials; production shared cube-view lifecycle implementation is in progress. ADODB.Command.ActiveConnection assignment was incompatible in PowerShell, while direct Execute worked. All isolated apps quit; original modal workbook remains.

Production SQL source lifecycle accepted: one consolidated cube DSV, dimension bindings to that view, and per-table `msprop:DataSourceID` preserve independent sources. Native two-source refresh/save/reopen produced First total75 and Second150. Rebinding Second changed only its rows to2101/2202/2303,total225; detaching First retained its cached75 while Second still refreshed225. Native DAX and independent exact Python readback passed all three saved workbooks; hashes in ledger. Native resave increments source filename versions; binder/detacher now resolve semantic ID rather than assuming .0.ds.xml. Dedicated regression verifies .1 source rebind/detach. Source lifecycle tests pass; original duplicate-DSV/commandtext failures remain diagnostic history, not current acceptance.

### Zero-row calculated cache fix

Native loading rejected the synthesized zero-length BLANK run. Zero-row caches now retain allocated native slots with no run or packed payload, no null flag and no used RLE storage. Three empty/repopulation regression cases passed. Publicly regenerated `work/calculated-empty-production.xlsx` passed native processing/query/save/reopen and Python readback with zero rows and all three columns. Evidence: `work/root-zero-sort-results-20261108/calculated-empty-production.json`.


Date-table processing fix accepted: designation must update the owning degenerate measure-group granularity along with dimension Usage/Type and raw key flags. Production candidate `work/date-key-production.xlsx` recalculated natively, returned Rows=3, DaySum=6, FirstDate=2026-01-01, saved, and reopened in an independent new Excel instance with identical query results. Exact Python readback passed. Initial same-process reopen failed while a deleted measure COM reference remained live; fresh-process acceptance separates that harness lifetime issue from workbook validity. Nine date tests and combined source/date suite of 23 tests passed. Hash recorded in native ledger.

Calculated date-key investigation: native processing and save/reopen first proved DATE(2026,1,[DayNumber]) values Jan1/2/3. Promoting that processed column to the date key with the same dimension/raw/granularity structure as stored keys fails native processing with a circular dependency on Calendar[CalculatedDate]. Installed native method inspection confirms date marking reverses the row-key edge and requests recalc. Production guard remains; public validation now reports this explicit limitation before a misleading datatype error. Unprocessed calculated values are not assumed unique/contiguous. Evidence in work/calc-date-key-results; this fixture establishes a concrete unsupported combination, not a universal claim about all DAX expressions.

### Stored columns sorted by calculated keys

Native proof showed that both semantic ProcessingState and AttributeHierarchyProcessingState must be Unprocessed; merely setting sort bindings retained stale order, and hierarchy-only invalidation failed. Stored values and indexes can remain as caches until processing. The writer now marks dependent attributes and user hierarchies unprocessed and defers headless sort rebuild. Public API fixture and changed-input fixture passed native processing/member-order/save/reopen and exact Python readback. Baseline Low/High order changed correctly to High/Low after Amount changed to5/40/10. One combined regression test covers stored-value preservation and dependent hierarchy invalidation.

### Native dictionary compaction and nullable lifecycle acceptance

Native DISTINCTCOUNT exposed retained seed members: the million-row model returned103 groups instead of100. Save now compacts string/numeric dictionaries to actual members and rebuilds affected intrinsic, relationship, custom-sort and user-hierarchy indexes. The corrected model returns Rows=1000000, IDSum=499999500000, Groups=100 before and after native save/reopen. Compacted lookup filtering remains b=30, c=BLANK. Tiny string dictionaries with1/2/3 actual values also passed.

Native zero/tiny references established that lookup RecordsPerSegment includes an actual blank member while dictionary payload cardinality excludes it. Without that distinction mixed text returned a distinct count missing BLANK. The corrected text-plus-blank test returns2, and a/b/BLANK returns3. An all-null hash column must have an unnamed dictionary and no dictionary file; named zero-entry payloads failed native loading. Text and numeric clear/repopulate now preserve this representation and rebuild storage when populated again.

Two-wide-integer native diagnostics isolated Usage=3 (not EncodingHint) as necessary for dictionary storage. Nullable two-wide integers additionally require OperatingOn32=false for the signed64 payload; the omitted flag caused native load failure even with matching dictionary and lookup streams. Three metadata variants isolated the flag: flag-only passed, raw Settings and DistinctStates variants failed. Production now supplies it for new Long dictionaries.

Nine production nullable numeric cases passed exact native counts, save/reopen and Python readback: integer, Currency and Double each with two values plus BLANK (count3), all-null (count1), then repopulated (count3). Latest focused numeric/compaction tests:67 passed in97.42s. Earlier numeric/schema/hierarchy compaction suite109 passed; local regressions18 passed after replacing obsolete dictionary-ID-retention assumptions with exact semantic value/order checks. Eighteen native cases and their source/saved/evidence hashes are recorded in `docs/native-acceptance-2026-11-08.json` and `work/storage-final-native-evidence.json`. These accepted cases supersede their explicitly failed diagnostic predecessors; final cross-feature validation is owned by root.

### Combined marked-date and relationship correction

A fresh combined fixture exposed an additional native contract: the date key's reversed attribute edge to RowNumber must have Cardinality=One. The prior writer kept Many from the old regular-column edge. Standalone date processing passed, but adding a relationship through another column caused native model-open failure. Isolation confirmed that unmarking the date table resolved it and that removing unrelated calculated relationships did not. The original Excel-generated date reference uses One.

Patching only this cardinality made both the minimal Calendar.Day relationship and the combined stage4 model pass native processing/query/save/reopen. Production marking now sets One, and unmarking/key switching restores Many unless an existing one-side relationship still requires uniqueness. Twelve date tests pass, including preservation of a bound Date lookup contract. Root's regenerated full combined fixture subsequently passed expected totals/filtering and save/reopen.

The calculated-date-key limitation was rechecked after correcting cardinality: `work/calc-date-key-one.xlsx` still raises the same native circular dependency on Calendar[CalculatedDate]. Evidence is `work/calc-date-key-one-results/calculated-date-key-one.json`; the structured guard remains justified for the tested combination. All owned native test instances were closed.

### Final supported-MVP verification checkpoint

The corrected integrated fixture passed native processing, filtering, save/reopen and all-table Python readback: Rows4, Total75, Double150, AuditDistinct4, CalendarRows3. Corrected date cardinality permits alternate-key relationships; the calculated-date-key negative test still fails with the native circular dependency and remains guarded. Source metadata preservation passed19tests and final automatic SQL refresh/reopen totals75/300. Frozen Windows suite512PASS472.17s; full Linux installed-wheel506PASS6Windows-onlyskips417.73s; clean Windows installed-wheel148PASS161.78s. Wheel/source/rebuilt-wheel55packagefiles are byte-identical. Final documentation-only artifacts retain the tested payload and will carry hashes in the distribution verification report. Nothing is published. Docker test container was restarted after preserving stale socket-only runtime folders; credentials/database content were unchanged.
