# Electre Customization — Complete Change Document

This file lists **all performance and structure work** from the start of the optimization effort through Continuity HAL fill, one-SaveAs wire lists, Megger streaming, helper folders, and `FindGroundPin`.

**Constraint:** Production-tested report **logic was not rewritten** as new business rules, except the items in §9 (you asked for those, or they were required so huge Megger files would not run out of memory).

**Times:** The only measured production baseline is **~45 minutes for a 100-sheet HAL cable list**. Other times are engineered estimates. Official proof is `{ELECTRE start}\Logs\{yyyy-MM-dd}.log`.

---

## 1. Summary

| Area | Before | After |
|---|---|---|
| CSV / pairing | Nested scans, CSV read twice | One CSV load, wire/subnet index, 2-end fast path |
| Continuity / Power On hops | Full-list scan per hop | `ElectreTraceIndex` + shared `ConnectorTraceHops` |
| Power On grounds | LINQ over the panel **per SCB** | `GroundPinIndex` built once |
| HAL Excel (cable list, Continuity) | New workbook, sequential `.xls` copy, cell-by-cell | File.Copy template, **xlsx-before-copy**, copy sheet 1, bulk `Value2` |
| Small wire-list workbooks | Empty SaveAs + Excel sort + extra Save | In-memory sort + **one SaveAs** |
| Megger `.txt` | All pairs in RAM + sort | Stream to disk, `i < j`, no bag |
| Code layout | Helpers inside `modMain` / `modExcel` | `Helpers/Global`, `CableList`, `Continuity`, `Megger`, `PowerOn` |

---

## 2. Files

### Modified
| File | Role |
|---|---|
| `MainOperation/modMain.cs` | CSV/pairing orchestration, report entry, thin wrappers |
| `MainOperation/modExcel.cs` | HAL generate, one-SaveAs wire lists, Power On flush |
| `Reports/ContinuityWOBreakdown.cs` | Trace wrappers, `ConnectionDeduper` |
| `Reports/PowerOnReport.cs` | Hop wrappers, `GroundPinIndex`, collection filter |
| `ReportUI/PanelDrawingWindow.cs` | Panel filter helper |
| `ReportUI/PanelDrawingSchedulesWindow.cs` | Sheet filter helper |
| `Objects/loomSort.cs` | `WireParts` (parse once for HAL) |
| `Forms/LoadingForm.cs`, `ExportToPDF.cs`, `MainOperation/SelectionFunctions.cs`, `MainOperation/DuplicateWiresCheck.cs`, `ReportUI/ECM_VariantsList.cs` | `OrdinalIgnoreCase` |

### New — Objects (constants, same behaviour)
| File | Role |
|---|---|
| `Objects/CableTypes.cs` | Cable-type families; empty type still matches `"__"` |
| `Objects/EquConnectorMap.cs` | J1↔a … J24↔z |
| `Objects/LayerCodes.cs` | NERD |
| `Objects/WireListColumns.cs` | Column indexes 0–6 |

### New — Helpers (project root, sibling of `MainOperation`)
```
Helpers/
  Global/
    UniqueNameList.cs
    ExtractionCsvLoader.cs
    WirePairIndex.cs
    WireSubNetOrdinalComparer.cs
    PairWireCode.cs
    ElectreListSort.cs
    ExcelRangeHelper.cs
    ExcelSheetOps.cs
    ElectreTraceIndex.cs
    ConnectorTraceHops.cs
  CableList/
    CableListHeaderFormulas.cs
    CableListStrikethrough.cs
    CableListDataBlock.cs
    CableListMerges.cs
    LoomGroupSort.cs
    WireListRowSort.cs
  Continuity/
    ContinuityHeaderFormulas.cs
    ContinuityDataBlock.cs
    ConnectionDeduper.cs
  Megger/
    MeggerEndpoint.cs
    MeggerSplitWriter.cs
    MeggerTextReport.cs
  PowerOn/
    PowerOnExcelBlock.cs
    ElectreCollectionFilter.cs
    GroundPinIndex.cs
```

---

## 3. Preprocessing (`MainFunction` / `MainFunction_withPanels`)

| Step | Before | After |
|---|---|---|
| CSV | Each method read the extraction file | One `ExtractionCsvLoader.Load` (64 KB sequential) |
| Row split | Empty col 28 → general; else panel | Unchanged |
| Uniques | `List.Contains` / `AddIfPresent` | `HashSet` then `UniqueNameList.ToSortedList` |
| Pairing | Nested loop, `ToUpper()` every compare | `WirePairIndex`; 2-end fast path; inner loop only for 3+ ends |
| Wire code | Duplicated inline | `PairWireCode.Build` — **project vs breakdown extra-slash still different** |
| Sort | Connector then pin | Same keys; **stable** (`ElectreListSort`) |
| Panel uniques | `SearchAndAppend` | HashSet; **SCB/TCB only** for Power On breaker lists (requested) |

---

## 4. Excel freeze (all reports)

Screen, events, calculation, print communication, interactive, animations **off** for the Excel session.

---

## 5. Report-by-report Excel / text

### 5.1 HAL cable list (`CL-n`) — former ~45 min / 100 sheets

| | Before | After |
|---|---|---|
| Workbook | `Workbooks.Add` + copy sheet from another workbook | `File.Copy` template |
| Copy mode | Sequential / geometric in **`.xls` BIFF8** | Stamp CL-1, **SaveAs `.xlsx`**, copy **CL-1 only**, name CL-2… |
| Data | ~41,800 `Cells.Value` | One `Value2` block C–L (`CableListDataBlock`) |
| ME merge | Twice per group, Merge even on 1 row, rewrite D/K | Dedupe; skip 1-row Merge; batch `Range("B9:B12,B20:B24").Merge`; D/K already in block |
| Strike | Per D-row + Calculate | Union batches of 20 |
| Headers | Per-sheet formula cells | Arrays on CL-2, grouped copy |
| >250 sheets | `.xls` fails at 255 | `.xlsx` before sheet 256 |
| Last sheet | Old Q loop could **repeat last data row** | No duplicate |
| Log | None | `Cable list '{name}': copy= fill= save= total= sheets=N` |

| Workload | Before | After (estimate) |
|---|---|---|
| **100 sheets** | **~45 min (measured)** | **~2–5 min** (copy still largest Excel piece) |
| 300 sheets | Fail at 256 | Minutes as `.xlsx` |

**Unchanged:** From/To, ME/UN/LAST look, `SHEET R OF N`, report-name split F51/K52, year B47, 38 rows/sheet, HAL template.

---

### 5.2 HAL Continuity (`CWOB-n`)

| | Before | After |
|---|---|---|
| Copy | Sequential in `.xls` | Same as cable: **xlsx-before-copy**, copy CWOB-1 |
| Data | 48 × 7 cell writes | **5 bulk writes**: C, E:F, H:I, K, N (spacers untouched) |
| Align after fill | Group-select all sheets | **Skipped** — aligned on CWOB-1 before copy; no data merges |
| `L4` | `Cells[4,12].Value` | `L4.Value2` |
| Hops | Full-list scan | `ElectreTraceIndex` |
| Dedup From/To | Local `removeRedundants` | `ConnectionDeduper` (7 cols, trim, skip `####`) |

---

### 5.3 Continuity Components

Standalone `.xls` and Megger sheet: bulk A/B + same A-column merges. Header on already-open file.

---

### 5.4 Power On

| | Before | After |
|---|---|---|
| Header | Open again; center `A2:F{Rows.Count}` | Reuse open file; center A:F once |
| Body | COM **inside** recursive trace | Buffer; flush in original merge order (TCB overlap kept) |
| Grounds | LINQ per SCB | `GroundPinIndex` once per collection |

---

### 5.5 Wire lists without HAL (project, sheet, breakdown, loom)

| Report | Before | After |
|---|---|---|
| Project `WireList.xls` | Empty SaveAs + Excel sort | Used rows, in-memory wire-code sort, **one SaveAs** |
| Sheet `{s}_SHEET_Wirelist.xls` | Same per sheet | Index **selected** sheets only, one SaveAs |
| EQU/BRK/JM/MISC | Empty SaveAs + Excel sort per name | Selected-name index, sort indexes, unrolled copy, one SaveAs |
| Loom without HAL | Empty SaveAs + 11-col append | `WriteLoomWithoutHalWorkbook` — **NERD** (no space) + Group Number, A:K, one SaveAs |

Group-id filter/sort for looms: `LoomGroupSort` (one parse + `List.Sort`). Incorrect/null group logging and MessageBox unchanged.

---

### 5.6 Loom **with HAL** (data prep)

Cannot use the 10-column SaveAs writer (must keep HAL template). Excel was already one SaveAs after copy+fill.

Data prep: index **only looms that appear**; compact FromConn+WireCode so fill/merge/LAST share `loomRowCount`; `WireCode.Split` once (`loomSort.WireParts`).

Log: `HAL loom '{name}': rows= sheets= prep+excel=`

---

### 5.7 Megger Scheduler + `.txt`

| | Before | After |
|---|---|---|
| Workbook | Add + 7 extra sheets + Open twice | One sheet + six adds; reuse open file |
| Continuity Components sheet | Cell-by-cell | Bulk + merge |
| High pairs | `i≠j` + dict of every key + **all lines in RAM** + sort + write High **twice** | Unique ends, **`i < j`**, stream temps, 8 MB I/O, no concat key, raw merge copy |
| Low files | 64 KB writer | Same splitter, 5 million lines/file |

**Same:** unique undirected High pairs; Lows from continuity; headers; `Megger_` / `HighMegger_` / `LowMegger_` names.

**Different (required for crores of lines):** High From/To is **stable** (`i < j`), not a parallel race; **no global sort** by FromConnector.

Log: `MeggerSheet5: ends= combined= high= elapsed=`

---

### 5.8 Material list HAL (`ML-n`)

Bulk OOTB read, geometric/xlsx copy, bulk B/D/E/G/H/K. **Not called** from current Panel Drawing UI (live path is `.txt`).

---

## 6. Shared helpers (behaviour)

| Helper | Replaces |
|---|---|
| `UniqueNameList` | `SearchAndAppend` / `ToSortedList` |
| `WirePairIndex` | Nested mate scan |
| `PairWireCode` | Duplicated wire-code builders |
| `ElectreListSort` | Connector+pin sort |
| `ExcelSheetOps` | Sheet copy, `WriteBlock`, FindByName |
| `ConnectorTraceHops` | Duplicate EQU/DIS/TBK/SPL/TER hops in Continuity + Power On |
| `ConnectionDeduper` | `removeRedundants` (7 cols) vs Megger (4 cols, no trim) |
| `ElectreCollectionFilter` | Panel/sheet/loom HashSet filters |
| `GroundPinIndex` | LINQ `FindGroundPin` per SCB |
| `LoomGroupSort` | Per-loom LINQ group sort |
| `CableListMerges` | Per-group Merge B+D+K |

`modMain` / `modExcel` still **orchestrate**; `SearchAndAppend` and `MeggerSheet5` remain thin public wrappers.

---

## 7. String compares

`ToLower() ==` / `ToUpper() ==` → `StringComparison.OrdinalIgnoreCase` (sheets, SCB/TCB, filters, `LOC-`, `MS`/`ML-`, SDS, `Linked`).

Pairing dictionary uses `WireSubNetOrdinalComparer` (no `ToUpper()` keys).

**Left:** `switch (ComponentType)` `"SCB"` / `"TCB"` (CSV already that casing). CableTypes joined-string `"__"` for empty type.

---

## 8. Cell maps (must still match golden files)

### HAL cable list — rows 9–46
C–K = `iarr` 1–9; L = layer 19; B = ME/UN/LAST serial; D = length; L51:L52 = `SHEET R OF N`; F51/K52 name split on CL-1 only.

### HAL Continuity — rows 6–53
C,E,F,H,I,K,N from `arrFTcwob` 0–6. D,G,J,L,M not written. L4 = `{R} OF {N}`.

### Wire list / breakdown A–J
FROM CONN … ` NERD` (leading space).

Loom **without HAL** adds K = Group Number; `NERD` has **no** leading space.

---

## 9. Intentional behaviour differences

| Item | What changed |
|---|---|
| Power On from **panels** | **SCB and TCB only** (requested). SWT/ERM/IND from panel objects not added. Panel EQU pins not appended in `_withPanels` (general EQU from `MainFunction` unchanged). |
| Megger High From/To | Stable `i < j`, not “whichever parallel worker won”. No global sort. |
| Cable list last sheet | No duplicate of the previous last data row. |
| `OrdinalIgnoreCase` vs `ToLower()` | Same for ASCII Electre names; can differ only for culture-specific casing. |

---

## 10. Templates / env

`ELECTRE_CUSTOMIZE\system\`:

- `CABLE_ReportFormat.xls`
- `CONTINUITY_ReportFormat.xls`
- `MATERIALLIST_ReportFormat.xls` (if HAL material list is used)

Do not replace these templates as part of this drop.

---

## 11. Production send / run / log

**Send:** whole folder `bin\Debug\net8.0-windows\` (VS MSBuild; do not `dotnet build` on the no-internet PC).

**Run:** new `Electre_Customize_DotNet.exe`. Keep the old shortcut until sign-off.

**Log:** `{ELECTRE start}\Logs\{yyyy-MM-dd}.log`

| Line | Meaning |
|---|---|
| `MainFunction: total=` | CSV + lists |
| `MainFunction_withPanels: csv+sort= projectPair= breakdownPair=` | Pairing |
| `ReportExcecution: all reports total=` | Whole Generate click vs old ~45 min |
| `Cable list '…': copy= fill= save=` | HAL cable write |
| `Continuity '…': copy= fill= save=` | HAL Continuity write |
| `HAL loom '…': rows= sheets= prep+excel=` | Loom HAL prep + Excel |
| `MeggerSheet5: ends= high= elapsed=` | Pair gen + `.txt` |
| `Power On '…': total=` | Power On write |
| `Component breakdown '…': total=` | One EQU/BRK file |
| `Project wire list` / `Sheet wire list` / `Loom wire list` | One-SaveAs paths |

**Pass:** spot-check From/To, merges, Continuity cells, Power On SCB/TCB list; cable list 100 sheets `total=` in minutes not ~45 min.  
**Rollback:** old exe. No data migration.

---

## 12. What is **not** pending (do not change without a new log)

- HAL `copy=` — one formatted page per extra sheet  
- Cable HAL UN/LAST/strike  
- Megger crores of High lines — disk  
- One `.xls` per component/sheet/loom  
- Commented dead methods  
- Material HAL Excel (not in current UI)  

Until a new log names a stage, **do not** add more Excel COM, HashSet, or Megger pair-logic edits.
