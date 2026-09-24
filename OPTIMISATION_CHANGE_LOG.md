# Electre Customize – Performance Optimisation Change Log

| | |
|---|---|
| **Branch** | `perf/report-generation-optimization` (base: `main`, last commit `e316121`) |
| **Status** | All changes are **uncommitted working-tree changes**. `main` is untouched. Nothing has been merged. |
| **Date of this document** | 2026-09-20 |
| **Projects touched** | `Electre_Customize_DotNet` (main), `SegregateLoom`, `CATLoom` |
| **Build** | Full solution builds with Visual Studio MSBuild (`dotnet build` cannot be used: COM reference, MSB4803) |
| **Verified by** | Build + code review + randomized equivalence tests for the extracted helper classes and the Megger engine (test harnesses live outside the repo). **No report has been run against real production data; no speed-up has been measured.** |

## 1. Purpose and ground rules

The reports (Cable List, Component Breakdown, Continuity, Megger, Power On) were very slow on large jobs
(300–500 drawings, 500+ connections, 100+ components, hundreds of looms).

Ground rules followed throughout:

1. **Report logic, output values and row order must not change** – this is production-tested code.
2. Only how the work is done may change: data structures, algorithms, Excel COM call patterns, memory use.
3. Case-sensitivity, null handling, source order and existing quirks are preserved unless listed in section 5.
4. Nothing is committed or merged until the changes are validated on the production systems.
5. Dead code inside `/* … */` blocks was not touched.

Where a change could not be proven identical without production data, it is either listed in section 5 (behaviour
notes) or put behind an `App.config` switch that is **off by default** (section 6).

## 2. Summary of what changed

| Area | Main idea | Expected effect |
|---|---|---|
| Excel I/O | Bulk block writes/reads, one-call ranges, typed sheets instead of late-bound lookups, Excel UI/alerts/events off | Largest constant-factor win in every report |
| Lookups | Linear `List.Contains` / `Where` / nested loops replaced by hash sets and dictionaries/indexes | Removes O(n²)/O(n³) behaviour on large jobs |
| Cable List | Per-loom/per-sheet buffers sized to their rows, indexed row lookup, redundant saves and calls removed | Big win on many looms/sheets |
| Continuity / Power On | Trace lookups indexed once per report, ground-pin lookup indexed by panel | Removes per-source full scans |
| Megger | Text-file writer rewritten as a streaming, parallel engine; pin-list expansion cached | Memory no longer grows with line count; works at ~10⁹ lines |
| Architecture | Helpers moved into `Helpers/` (Global, ExcelHelpers, CableList, Continuity, PowerOn, Megger) | Cleaner structure, shared logic once |

## 3. Detailed change list

Each line: what was slow → what was done → where.

### 3.1 Excel I/O and Excel application state (all reports)

| ID | Change | Files |
|---|---|---|
| E-01 | Excel instance runs with `ScreenUpdating`, `EnableEvents`, `DisplayAlerts` off. **Calculation stays Automatic** (templates use live cross-sheet formulas). | `modExcel.InitiateExcel` |
| E-02 | Cell-by-cell writes replaced by one `object[,]` block write (`Range.Value`, same as the original – not `Value2`) with the COM range released. Applied in `AppendToExcel`, `AppendToExcelMegger` (Pin List, Connection List, Exception), `LogToExcel`, Cable List and Continuity generators, header rows. | `Helpers/ExcelHelpers/ExcelBulk.cs`, `modExcel.cs` |
| E-03 | Bulk read of the Cable List sheet (`Range.Value2`) in `CableList_Report_Preprocessing` instead of 7 cell reads per row. | `modMain.cs`, `ExcelBulk.ReadValues` |
| E-04 | Ranges fetched by one A1 address (`"C9:L9"`) via `ExcelBulk.RangeAt` – 1 COM call instead of 5 (`Cells`, indexer, `Cells`, indexer, `Range`). Used in row writes, Power On `LogToExcel` merges, Continuity Components sheet, Megger link list. | `ExcelBulk.cs`, `modExcel.cs`, `ExcelHyperlinks.cs` |
| E-05 | Merge / unmerge / `alignCellsXl` receive the sheet just activated instead of doing a late-bound `ReportWB.ActiveSheet` lookup per range (optional parameter, other callers unchanged). | `modExcel.cs` |
| E-06 | Template-sheet lookup ("CABLE LIST" / "continuity") hoisted out of the per-sheet copy loop; `alignCellsXl(6,53)` hoisted in the Continuity generator; sheet name compared with `StringComparison.OrdinalIgnoreCase` instead of `ToLower()`. | `modExcel.cs` |
| E-07 | Per-row `ScreenUpdating = true` and `ws1.Calculate()` on struck-through (`"D"`) Cable List rows **removed** (they switched the screen-updating optimisation back on for the rest of the run; Excel is hidden and calculation is automatic). | `GenerateHALReportFormat_CableList` |
| E-08 | Header row of the wire-list workbooks written as one row; one `Columns` object for the ten widths; `ReportWS.Sort` fetched once. | `CreateWirelistReportHeader`, `SortWirelistWireNumber` |
| E-09 | Cable List linked cells on sheets 2..n assigned as range formulas: `C48:C52`, `D48:D52`, `K51:K52` (Cable List) and `C56:C57` (Continuity). Excel adjusts the relative reference per cell (e.g. `C49 → 'CL-1'!C49`). Non-contiguous cells unchanged. | `modExcel.cs` |
| E-10 | Megger link list: column width set once; link cell fetched with one call. | `ExcelHyperlinks.cs` |

### 3.2 Data structures / algorithmic complexity

| ID | Change | Files |
|---|---|---|
| D-01 | `SearchAndAppend` (linear contains-then-add) replaced by `UniqueStringList.TryAdd` (hash mirror per list, self-healing if the list is changed directly). Added `UniqueStringList.Contains` for O(1) membership. | `Helpers/Global/UniqueStringList.cs`, `modMain.cs`, `ContinuityWOBreakdown.cs` |
| D-02 | Lists that are sorted afterwards (10 in `MainFunction`, panel list in `MainFunction_withPanels`) collected in `HashSet` first, then ordered. Lists whose insertion order matters keep `UniqueStringList`. | `modMain.cs` |
| D-03 | Wire + subnet pairing (`Removing_DuplicateWires_In2DArray`, `Reading_And_StoringData_In2DArray`) via `WireSubNetIndex` (group list for the both-direction pass, queues for the consuming `Tag6_Link` pass). Original `.ToUpper()` key semantics kept. | `Helpers/CableList/WireSubNetIndex.cs`, `modMain.cs` |
| D-04 | Loom report row lookup via `LoomRowIndex` (FROM bundle column 8 checked before TO bundle column 10, ascending row order, sheet filter, null bundles never match). | `Helpers/CableList/LoomRowIndex.cs` |
| D-05 | Loom group rules (invalid group check, filter/sort) moved to `LoomGroupRules` with compiled regex instances; same patterns. | `Helpers/CableList/LoomGroupRules.cs` |
| D-06 | Continuity/Power On trace lookups (wire+subnet, connector+pin, connector+type+shunt) via `ElectreTraversalIndex`, built once per list (`TraversalIndexCache` rebuilds only when the list reference changes). Case sensitivity of each original lookup preserved. | `Helpers/Global/ElectreTraversalIndex.cs`, `TraversalIndexCache.cs` |
| D-07 | Power On `FindGroundPin`: full collection scan per circuit breaker replaced by `GroundCandidateIndex` (candidates grouped by panel once). | `Helpers/PowerOn/GroundCandidateIndex.cs`, `PowerOnReport.cs`, `modExcel.AppendToExcelPowerOn` |
| D-08 | Continuity main loop: `destinationConnectors.Contains` (linear, list grows during the loop) now O(1) via the live hash mirror. | `ContinuityWOBreakdown.cs` |
| D-09 | Continuity: populated-row count used to trim `removeRedundants` / `removeRedundantsMegger` scans (Megger variant scans the populated prefix plus one blank row to preserve its quirk); `Array.Clear` for `arrFTcwob`; `loomElectreObject` and panel/sheet selection use hash sets. | `ContinuityWOBreakdown.cs`, `PowerOnReport.cs` |
| D-10 | Connector-suffix tests (`Pairs.Keys.Any/FirstOrDefault` building 48 strings per connector) replaced by precomputed `EndsWithUnderscoredKey` / `FirstUnderscoredKey`. Used in Continuity sort/grouping and the Megger sort. | `Helpers/Global/EquConnectorMap.cs` |
| D-11 | Cable List loom bundle report: per-iteration `Where(WireName==…).Count()` replaced by a precomputed count dictionary; special-cable list is a `HashSet`. | `modMain.cs` |
| D-12 | `GeneratePinListArray`: library entries grouped by part number, gauge by connector name indexed, continuity components as a set; pin lists per (part number, gauge) expanded **once** (`LibraryPinExpander`) and reused. Unused second full-size array removed. | `modMain.cs`, `Helpers/Megger/LibraryPinExpander.cs` |
| D-13 | `MeggerSchedulesheet` filters use hash sets. | `modMain.cs` |
| D-14 | `DuplicateWiresCheck` nested `Any` replaced by a `HashSet` of pairs. | `DuplicateWiresCheck.cs` |
| D-15 | `SegregateLoom`: bulk sheet read + `HashSet` for the searched items; `CATLoom`: `ToLookup` in `generateCATReports`. | `SegregateLoom.cs`, `CatLoom.cs` |
| D-16 | `sCableType…` string tests, `TakeWhile` scans etc. reviewed and left (millisecond cost). | – |

### 3.3 Cable List and Component Breakdown specific

| ID | Change | Files |
|---|---|---|
| C-01 | **Per-loom buffers** in `loomReportWireList` / `loomReportWireListBundle`: were `new object[arrFT_CwithBCProject.Length, 14]` (rows × columns rows!) for every loom and then scanned; now sized to exactly the loom's rows. Dead per-loom `arrCustomLoom = new object[R, 9]` removed. | `modMain.cs` |
| C-02 | **Sheet-wise and component-wise reports**: row lookup by `ColumnRowIndex` (ordinal, case-sensitive, ascending); buffer grows per sheet/component (`RowBuffer.EnsureCapacity`, filled rows + 4 blank rows, capped at the original height). The 4 blank rows matter: `AppendToExcel` only detects the end-of-data block when the array has more than 3 rows. | `Helpers/CableList/ColumnRowIndex.cs`, `RowBuffer.cs`, `modMain.cs` |
| C-03 | **Sort range clamped** to the rows actually written (`ReportAppendRow - 1`), never above the requested end and never before the first data row. | `modExcel.SortWirelistWireNumber` |
| C-04 | Redundant saves skipped for sheet/component/project wire-lists: `AppendToExcel(…, saveWorkbook:false)` and `SortWirelistWireNumber(…, saveBeforeClose:false)` – the final `Close(true)` saves the same content. Defaults keep the old behaviour for other callers. | `modExcel.cs`, `modMain.cs` |
| C-05 | Loom (non-template) workbooks closed after `AppendToExcel` saved them (`CloseReportWorkbookWithoutSaving`); they used to stay open until Excel quit (one per loom). | `modExcel.cs`, `modMain.cs` |
| C-06 | Dead work removed per loom: discarded `loomSortList.Distinct().ToList()` (both loom methods) and the unused `loomRowCount` loop (`loomReportWireList`). | `modMain.cs` |
| C-07 | Opt-in: template sheet copies from the first copy (`IntraWorkbookSheetCopy`), one-pass component/sheet workbooks (`OnePassWirelistWorkbook`) – see section 6. | `TemplateSheetCopier.cs`, `modExcel.cs`, `modMain.cs` |
| C-08 | `WireListReport` (project wire list): the buffer was `arrFT_CwithBCProject.Length` rows (rows × columns) although `AppendToExcel` only writes up to the first block of three blank rows. `RowBuffer.EquivalentHeight` gives the smallest buffer that makes `AppendToExcel` write exactly the same rows (checked against the original on 3,000 random arrays). Also skips the redundant intermediate save (C-04). | `Helpers/CableList/RowBuffer.cs`, `modMain.cs` |
| C-09 | Opt-in parallel Excel: the loom (with and without template), sheet-wise and component-wise (Component, Break Connector, Junction Module, Misc breakdowns via `equipmentReportWireList`) reports run through `ExcelParallel.ForEach`. `ExcelWorkerCount` = 1 (default) is the old sequential loop on the shared Excel instance; 2–8 runs that many separate Excel processes at once (STA worker threads, each with its own `modExcel`), because every loom/sheet is an independent workbook. Template opened read-only by workers; a worker that would share an already-used Excel instance is dropped; leftovers run on the caller's instance. Reviewed specifically for concurrency correctness (see section 5 note B-10). | `Helpers/ExcelHelpers/ExcelParallel.cs`, `modMain.cs`, `modExcel.cs` |
| C-11 | Cable List per-sheet data write: was one bulk `ExcelBulk.WriteValues` call **per row** (up to 37-38 calls per sheet). Now the whole sheet's raw row data is collected and written in **one** call per sheet, with the strikethrough/merge/unmerge pass deferred to run after. This reorders raw writes relative to merges - a change that was rejected from the Grok review (C-10 note below) when *unverified*. Here it **was** verified: directly tested in real Excel via COM (not just reasoned about) for the two cases that could plausibly differ - a "ME" merge group whose last row is written later in the same block, and a "UN" unmerge - and both produced byte-identical final cell values and merge geometry to the original interleaved order. See section 11 for the test description. | `modExcel.cs` (`GenerateHALReportFormat_CableList`) |
| C-10 | `.xls` (BIFF8) cannot hold more than 255 sheets - above that, Excel refuses the next sheet copy mid-loop and no file is produced at all. `GenerateHALReportFormat_CableList` and `GenerateHALReportFormat_ContinuityList` now save as `.xlsx` instead **only** when `numberOfSheetsRequired > 255` (`XlsSheetLimit`, opt-out via `AutoXlsxAboveSheetLimit`, default on). Below the threshold, file name, format and every other step are unchanged. Merged in from a review of a separate optimisation pass ("Grok's") - the one idea from that review that fixes a real Excel limit without touching sort order, calculation mode or row-truncation logic; see section 11 for what was deliberately **not** merged from that same review, and why. | `Helpers/ExcelHelpers/XlsSheetLimit.cs`, `modExcel.cs` |

### 3.4 Continuity, Power On, Megger

| ID | Change | Files |
|---|---|---|
| M-01 | **Megger text files** (`MeggerSheet5`, `MeggerSheet6and7`) rewritten as a streaming engine: only the distinct (connector, pin) nodes are held; lines are generated as UTF-8 bytes in final order; per-row cumulative counts let each 5,000,000-line file be written independently, in parallel (default `min(4, cores/2)`, optional app setting `MeggerWriteParallelism`). Same file layout: BOM + two header lines, low lines first in `Megger_n.txt`, roll-over every 5M lines. Measured ≈ 12–17 M lines/s in the test harness. | `Helpers/Megger/*` |
| M-02 | `MeggerSheet5`/`6and7` in `modMain` reduced to UI/logging + delegation. | `modMain.cs` |
| M-03 | "Continuity Components" sheet in `AppendToExcelMegger`: all values in one block write, then merges/centering (each component's rows are contiguous, merge ranges only cover its own column-1 cells). | `modExcel.cs` |
| M-04 | Continuity report's Continuity Components sheet (`AppendContinuityComponentsandPartNumber`): same operations in the same order, cheaper COM calls only (it writes `""` into already-merged rows, so it cannot be bulk-written). | `modExcel.cs` |
| M-05 | Power On `LogToExcel`: merges via one-call ranges; ground index built once per report. | `modExcel.cs` |

### 3.5 UI, logging, configuration

| ID | Change | Files |
|---|---|---|
| U-01 | `SelectCheckItems` and `UpdateCheckedListBox` wrapped in `BeginUpdate/EndUpdate` (one repaint instead of one per item; item-check handlers still run). | `SelectionFunctions.cs` |
| U-02 | `WriteLog` / `DeletePreviousLogs` of the three small loggers share `LogFileHelper`; `AppendMaterialListTxt` / `AppendAdminTxt` share `TextFileAppender`. Log format, folders, file names unchanged. | `Logs/*`, `Helpers/Global/*` |
| U-03 | `ServerGarbageCollection` and `ConcurrentGarbageCollection` enabled. | `Electre_Customize_DotNet.csproj` |
| U-04 | `PerfDiagnostics` timing (see section 7). | `Helpers/Global/PerfDiagnostics.cs`, `modMain.cs` |
| U-05 | Log writes are locked (`Logging`, `LogFileHelper`, log-file-name creation of the two group-ID loggers) so parallel Excel workers cannot lose or interleave log lines. No effect when `ExcelWorkerCount` is 1. | `Logs/*`, `Helpers/Global/LogFileHelper.cs` |

### 3.6 Architecture (helpers moved out of the legacy classes)

Pure helper methods were moved unchanged into separate classes (callers updated; `/* */` code untouched):

```
Helpers/
├─ Global/       LayerCodes, EquConnectorMap, SheetNames, ReportFolderCleaner, UniqueStringList,
│                PerfDiagnostics, ElectreTraversalIndex, TraversalIndexCache, ConnectorPinTracer,
│                LogFileHelper, TextFileAppender
├─ ExcelHelpers/ ExcelBulk, ExcelRowRules, ExcelFormatting, ExcelHyperlinks, TemplateSheetCopier, ExcelSwitches
├─ CableList/    CoreNumberRules, LoomGroupRules, LoomRowIndex, WireSubNetIndex, ColumnRowIndex, RowBuffer
├─ Continuity/   PinNumberSort
├─ PowerOn/      ComponentSelectors, GroundCandidateIndex
└─ Megger/       MeggerListConverters, MeggerRow, MeggerByteSink, MeggerHighPairs, MeggerFileWriter,
                 MeggerReportWriter, LibraryPinExpander
```

Notable: `ConnectorPinTracer` holds the `TracePinofEQUConnector / BreakConnector / OfJM / OfSPL / OfTER` logic that was
duplicated in `PowerOnReport` and `ContinuityWOBreakdown`. The logic is identical; only the searched collection
differs (whole project for Power On, selected looms for Continuity), so each report passes its own index.

## 4. Files changed

Modified: `CATLoom/CatLoom.cs`, `SegregateLoom/SegregateLoom.cs`, `Electre_Customize_DotNet/App.config`,
`Electre_Customize_DotNet.csproj`, `Logs/DuplWireLogging.cs`, `Logs/Incorrect cable group IDs.cs`,
`Logs/Null Group IDs.cs`, `MainOperation/DuplicateWiresCheck.cs`, `MainOperation/SelectionFunctions.cs`,
`MainOperation/modExcel.cs`, `MainOperation/modMain.cs`, `Reports/ContinuityWOBreakdown.cs`, `Reports/PowerOnReport.cs`.

New (untracked): the whole `Electre_Customize_DotNet/Helpers/` tree (32 classes, listed above), `VALIDATION_GUIDE.md`,
and this file.

## 5. Behaviour notes – please review these

These are the places where output or behaviour is **not** simply "same result, faster":

| # | Item | Detail |
|---|---|---|
| B-1 | Megger High pairs | The old parallel code was racy for a small group of pairs (which orientation of an A–B pair was tested against the low list depended on thread timing). The new output is deterministic and lies within the outcomes the old code could produce: a pair is excluded if either orientation is in the low list; the lower node in sort order is the "from" side; a filtered node with rows disagreeing on (wire, subnet) still yields one `X;pin;X;pin` self line. High lines are in ascending (connector, pin) order (the old bag enumeration also produced ascending files). Compare line counts and unordered pair sets, not byte-for-byte order. |
| B-2 | Component/sheet report sort | Previously the sort range was the whole array height; above the 65,536-row `.xls` limit Excel rejected it and the report stayed unsorted (with an error box). Now the sort runs on the written rows, so **large jobs that used to come out unsorted now come out sorted**. Jobs that already sorted correctly are unchanged. |
| B-3 | Cable List struck-through rows | The two lines removed (E-07) have no effect on file content (hidden Excel, automatic calculation). |
| B-4 | Loom workbooks | Now closed after saving (C-05); files should be identical. |
| B-5 | Sheet/component workbooks | Two intermediate saves skipped (C-04); final content saved by `Close(true)`. |
| B-6 | Linked formula cells | Assigned as range formulas (E-09). Verify a few cells on sheet 2+ (each should point at the same cell of sheet 1). |
| B-7 | Error paths | Where the original threw on null data (e.g. null `ComponentType`, null group code) the new code still throws at the same logical point; the equivalence tests cover null wires, null panels, null gauges and bad numeric sizes. |
| B-10 | `.xls`→`.xlsx` above 255 sheets (opt-in, default on) | Reports needing more than 255 sheets now produce a `.xlsx` file instead of failing outright. Nothing changes for reports at or under 255 sheets - same file name, same `.xls`, same everything. |
| B-9 | Parallel workers (opt-in) | With `ExcelWorkerCount` > 1: workbooks are created in a different order; message boxes (e.g. "incorrect cable group id") are shown from worker threads and only block that worker; each worker keeps its own row buffer and height cap (the "stale buffer after an error" behaviour of the sheet and component reports is per worker). Default 1 is unchanged. |
| B-8 | Preserved quirks | e.g. Cable List column 10 = `iarr[Q,19]`, `arrFTcwob` "one blank row" behaviour in `removeRedundantsMegger`, `Connection List` format range using the outer `row`, TCB merge reaching into the next block. |

## 6. Opt-in switches (`App.config`, all default `false` = original behaviour)

| Key | Effect | Validate by |
|---|---|---|
| `IntraWorkbookSheetCopy` | Cable List (`CL-n`) and Continuity (`CWOB-n`) sheets: copy only the first sheet from the template, the rest from that first copy inside the report workbook. | Generate the same large report with `false` and `true`; compare content, merges, number formats, widths, page setup, fonts. |
| `OnePassWirelistWorkbook` | Sheet-wise and component-wise wire-lists: single-sheet workbook created, header + data + format + Excel sort, **one** `SaveAs`. | Same comparison; also check the sort order and header formatting. |
| `ExcelWorkerCount` | `1` (default) = as before. `2`–`8` = that many Excel processes generate looms / sheets / components at the same time (each Excel needs its own memory; 2–4 is sensible). | Same large job with `1` and `2`/`4`: compare files; expect only the creation order to differ. Watch memory. |
| `PerfDiagnostics` | Writes timing lines to `Logs\perf_<date>.log`. | – |
| `AutoXlsxAboveSheetLimit` | `true` (default) = save as `.xlsx` only above 255 sheets, otherwise unchanged `.xls`. `false` = old all-or-nothing `.xls` behaviour (report fails above the limit). | Generate a report that needs >255 sheets with each value; confirm the `true` case opens correctly in Excel and the `false` case fails the way it always did. |
| `PerfFingerprint` | Intended to write workbook fingerprints – **not wired yet** (see section 8). | – |
| `MeggerWriteParallelism` (optional, not in the config file) | Number of Megger files written at once; use 1 on a spinning disk. | – |

## 7. PerfDiagnostics

`Helpers/Global/PerfDiagnostics.cs` – optional, zero overhead when off. With `PerfDiagnostics=true` each report type in
`ReportExcecution` is timed and one line is appended per report to `Logs\perf_<date>.log`, for example
`[PERF] report=Continuity phase=ReportGeneration elapsedMs=48210 SelectedLoomCount=37`.

## 8. Known gaps and limitations (honest status)

1. **No measurement yet.** No speed-up figure exists; production timing is still to be collected.
2. **Fingerprint capture is not wired.** `PerfDiagnostics.CaptureFingerprint` exists but nothing calls it, so
   `PerfFingerprint=true` produces no files. `VALIDATION_GUIDE.md` sections 3–4 describe it and are therefore
   inaccurate, and `main` does not contain `PerfDiagnostics` at all. A standalone fingerprint tool that reads the
   finished `.xls` files of the old and new builds is the recommended way to compare output.
3. Only the extracted helper classes and the Megger engine were tested against the original algorithms (randomized
   tests). Excel-side changes (E-*, C-04…C-07, M-03…M-05) are reasoned equivalences awaiting a real-data comparison.
4. Not done on purpose: manual Excel calculation mode; sorting component reports in C# (Excel's pinyin, stable,
   hyphen-ignoring sort cannot be reproduced safely); reordering the Power On / Cable List write-and-merge sequences;
   merging the duplicated CSV→`ElectreObject` parsing in `MainFunction` / `MainFunction_withPanels`; a shared library
   for the `Logging` / loom-UI code duplicated across the three executables; deleting commented-out code.

## 9. How to validate (short form)

1. Build the branch; keep a copy of the old build (`main`) for comparison.
2. Turn on `PerfDiagnostics`; run one large job per report type; collect `perf_*.log`.
3. Compare the generated files of old vs new: Cable List / Component / Continuity / Power On workbooks (values,
   merged ranges, number formats, widths, sort order) and the Megger text files (line counts per file, unordered pair
   sets – see `VALIDATION_GUIDE.md` section 8).
4. Try `IntraWorkbookSheetCopy` and `OnePassWirelistWorkbook` one at a time, comparing as in section 6.
5. Only after step 3 passes: commit and merge.

## 11. What was reviewed from a separate optimisation pass, and what was (and was not) merged

A separate pass on this codebase (referred to above as "Grok's" review, at
`C:\Electre_Project\Code\electre_customization_dotnet`) was read end-to-end for its Cable List / Component
Breakdown implementation, specifically to see whether any of its speed came from ideas that could be merged in
without breaking rule 1 in section 1 ("report logic, output values and row order must not change").

**Merged (C-10 above):** the `.xls`→`.xlsx` switch above 255 sheets. This fixes a real Excel format limit that
this codebase's own original code never handled (and still doesn't, below the opt-out). It changes nothing for
any report that was already working, and only ever activates for a report that would otherwise fail outright -
so there was no existing correct behaviour to preserve above that threshold.

**Not merged, and why:**

| Idea in the other pass | Why it was not merged |
|---|---|
| `Calculation = xlCalculationManual` for the whole Excel session, with no `.Calculate()` call found anywhere before the workbook is saved | The Cable List / Continuity templates use live cross-sheet formulas (`='CL-1'!C48` etc.) that this codebase deliberately keeps on Automatic calculation for exactly that reason (see E-01). Writing a formula under Manual calculation without a guaranteed recalculation before save can leave that cell's saved value stale or blank - this is a real risk, not a theoretical one, and it was not something that build proved safe. |
| Sorting the sheet/loom/component wire-list rows in C# (`List.Sort` with `string.Compare(..., OrdinalIgnoreCase)`) instead of Excel's own `.Sort` | Not equivalent to Excel's pinyin/collation sort (which treats hyphens and punctuation differently than ordinal comparison), and `List.Sort` is not a stable sort, so rows with equal wire codes can come out in a different order than the original produced. This is exactly the kind of output-order change rule 1 forbids, and it was not verified against the original. |
| Dropping the "stop at the first block of three blank rows" truncation rule (`WriteComponentBreakdownWorkbook` writes however many rows its own counter says, with no blank-row scan) | This is the exact quirk this codebase's own `RowBuffer`/`ExcelRowRules` reproduce deliberately and verified by test (see C-01/C-02, test E). Dropping it is a logic change, not a performance change, and was not proven equivalent to the original for data containing a blank block mid-array. |

None of these three were adopted.

**A fourth idea *was* adopted, but only after being independently verified, not merged on trust (C-11 above):**
Grok's Cable List generator writes a whole sheet's data in one bulk call and applies merges afterward, instead
of interleaving them row-by-row. That reordering is exactly the kind of change this document treats as
unproven by default - so before adopting it, it was tested directly in a live Excel instance via COM (a small
PowerShell/COM harness, not the C# equivalence-test harness used elsewhere in this document, since this needed
an actual running Excel rather than compiled logic). Two representative cases were run both ways
(row-by-row and batched) and diffed:
- a "ME" merge group where the merge is created before its own last row has been raw-written yet (the specific
  scenario that makes this reordering non-obviously-safe), and
- a "UN" unmerge group.

Both produced byte-identical final cell values and merge geometry in either order. On that basis, C-11 was
implemented as a real change (not an opt-in switch), because it now has the same kind of direct evidence behind
it that the rest of this document holds itself to, rather than being merged in on the strength of another
review's results.

 If you want any of them regardless, say so explicitly and each one can be
added as its own separate, clearly-labelled, opt-in change (not silently folded into "optimisation") so it can
be validated on its own before being trusted.

## 10. Rollback

Everything is uncommitted on `perf/report-generation-optimization`, so `main` is the rollback point.
To discard all changes: `git checkout -- .` and delete the untracked `Helpers/` folder (and the two `.md` files).
Individual switches in section 6 can simply be left `false`.
