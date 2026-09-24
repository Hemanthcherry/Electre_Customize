const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
        ShadingType, VerticalAlign, PageNumber, PageBreak, LevelFormat } = require("docx");
const fs = require("fs");
const path = require("path");

const PAGE_W = 11906;
const PAGE_H = 16838;
const MARGIN = 1080;
const CONTENT_W = PAGE_W - MARGIN * 2; // 9746
const NAVY = "1F4E79";
const HEADER_FILL = "1F4E79";
const ALT_FILL = "F2F2F2";
const border = { style: BorderStyle.SINGLE, size: 4, color: "BDD7EE" };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMargins = { top: 60, bottom: 60, left: 80, right: 80 };

function p(text, opts = {}) {
  return new Paragraph({
    spacing: { after: opts.after ?? 120, before: opts.before ?? 0 },
    children: [new TextRun({ text, font: "Arial", size: opts.size ?? 22, bold: opts.bold, color: opts.color })],
  });
}

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 280, after: 140 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: NAVY, space: 4 } },
    children: [new TextRun({ text, font: "Arial", size: 28, bold: true, color: NAVY })],
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 220, after: 100 },
    children: [new TextRun({ text, font: "Arial", size: 24, bold: true, color: "2E75B6" })],
  });
}

function cell(text, width, opts = {}) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: opts.fill || "FFFFFF", type: ShadingType.CLEAR },
    margins: cellMargins,
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      children: [new TextRun({
        text: text == null ? "" : String(text),
        font: "Arial",
        size: opts.size || 18,
        bold: !!opts.bold,
        color: opts.color || "000000",
      })],
    })],
  });
}

function table(headers, rows, colWidths) {
  const widths = colWidths || equalCols(headers.length);
  const headerRow = new TableRow({
    children: headers.map((h, i) => cell(h, widths[i], { fill: HEADER_FILL, bold: true, color: "FFFFFF", size: 18 })),
  });
  const dataRows = rows.map((r, ri) => new TableRow({
    children: r.map((c, i) => cell(c, widths[i], { fill: ri % 2 ? ALT_FILL : "FFFFFF" })),
  }));
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: widths,
    rows: [headerRow, ...dataRows],
  });
}

function equalCols(n) {
  const w = Math.floor(CONTENT_W / n);
  const cols = Array(n).fill(w);
  cols[n - 1] = CONTENT_W - w * (n - 1);
  return cols;
}

function codeBlock(label, code, kind) {
  const fill = kind === "before" ? "FDEDEC" : "E8F8F5";
  const color = kind === "before" ? "922B21" : "196F3D";
  const lines = String(code).replace(/\t/g, "  ").split(/\r?\n/);
  const paras = [
    new Paragraph({
      spacing: { after: 80 },
      children: [new TextRun({ text: label, font: "Arial", size: 16, bold: true, color })],
    }),
    ...lines.map((line) => new Paragraph({
      spacing: { after: 0, before: 0 },
      children: [new TextRun({ text: line.length ? line : " ", font: "Consolas", size: 16, color: "1B4F72" })],
    })),
  ];
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W],
    rows: [new TableRow({
      children: [new TableCell({
        borders,
        width: { size: CONTENT_W, type: WidthType.DXA },
        shading: { fill, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 100, left: 120, right: 120 },
        children: paras,
      })],
    })],
  });
}

function spacer() {
  return new Paragraph({ spacing: { after: 140 }, children: [] });
}

const children = [];

children.push(new Paragraph({
  spacing: { after: 80 },
  children: [new TextRun({ text: "ELECTRE CUSTOMIZATION", font: "Arial", size: 20, color: "5B9BD5", bold: true })],
}));
children.push(new Paragraph({
  spacing: { after: 80 },
  children: [new TextRun({ text: "Performance Optimization — Complete Change Document", font: "Arial", size: 36, bold: true, color: NAVY })],
}));
children.push(new Paragraph({
  spacing: { after: 200 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 16, color: NAVY, space: 8 } },
  children: [new TextRun({ text: "All changes from the start of the optimization work through HAL copy/fill, one-SaveAs wire lists, Megger streaming, helper architecture, and FindGroundPin.", font: "Arial", size: 22, color: "595959" })],
}));
children.push(p("Date: 20 September 2026", { size: 20, color: "595959" }));
children.push(p("Audience: Production / engineering sign-off", { size: 20, color: "595959", after: 200 }));

children.push(h1("1. Purpose and constraint"));
children.push(p("This document lists every performance and structure change. Production-tested report logic was not rewritten as new business rules, except the items in section 9 (requested by the owner, or required so huge Megger files would not run out of memory)."));
children.push(p("The only measured production baseline is about 45 minutes for a 100-sheet HAL cable list. Other times are engineered estimates. Official proof is the dated log under {ELECTRE start}\\Logs\\{yyyy-MM-dd}.log."));

children.push(h1("2. Summary"));
children.push(table(
  ["Area", "Before", "After"],
  [
    ["CSV / pairing", "Nested scans; CSV read twice", "One CSV load; wire/subnet index; 2-end fast path"],
    ["Continuity / Power On hops", "Full-list scan per hop", "ElectreTraceIndex + ConnectorTraceHops"],
    ["Power On grounds", "LINQ over the panel per SCB", "GroundPinIndex built once"],
    ["HAL Excel (cable, Continuity)", "New workbook; sequential .xls copy; cell-by-cell", "File.Copy template; xlsx-before-copy; copy sheet 1; bulk Value2"],
    ["Small wire-list workbooks", "Empty SaveAs + Excel sort + extra Save", "In-memory sort + one SaveAs"],
    ["Megger .txt", "All pairs in RAM + sort", "Stream to disk; i < j; no bag"],
    ["Code layout", "Helpers inside modMain / modExcel", "Helpers/Global, CableList, Continuity, Megger, PowerOn"],
  ],
  [2200, 3773, 3773]
));

children.push(h1("3. Files"));
children.push(h2("3.1 Modified"));
children.push(table(
  ["File", "Role"],
  [
    ["MainOperation/modMain.cs", "CSV/pairing orchestration, report entry, thin wrappers"],
    ["MainOperation/modExcel.cs", "HAL generate, one-SaveAs wire lists, Power On flush"],
    ["Reports/ContinuityWOBreakdown.cs", "Trace wrappers, ConnectionDeduper"],
    ["Reports/PowerOnReport.cs", "Hop wrappers, GroundPinIndex, collection filter"],
    ["ReportUI/PanelDrawingWindow.cs", "Panel filter helper"],
    ["ReportUI/PanelDrawingSchedulesWindow.cs", "Sheet filter helper"],
    ["Objects/loomSort.cs", "WireParts (parse once for HAL)"],
    ["Forms, SelectionFunctions, DuplicateWiresCheck, ECM_VariantsList", "OrdinalIgnoreCase"],
  ],
  [4200, 5546]
));

children.push(h2("3.2 New — Objects (same behaviour)"));
children.push(table(
  ["File", "Role"],
  [
    ["Objects/CableTypes.cs", "Cable-type families; empty type still matches \"__\""],
    ["Objects/EquConnectorMap.cs", "J1↔a … J24↔z"],
    ["Objects/LayerCodes.cs", "NERD"],
    ["Objects/WireListColumns.cs", "Column indexes 0–6"],
  ],
  [4200, 5546]
));

children.push(h2("3.3 New — Helpers (project root, sibling of MainOperation)"));
children.push(p("Helpers/Global — UniqueNameList, ExtractionCsvLoader, WirePairIndex, WireSubNetOrdinalComparer, PairWireCode, ElectreListSort, ExcelRangeHelper, ExcelSheetOps, ElectreTraceIndex, ConnectorTraceHops.", { after: 80 }));
children.push(p("Helpers/CableList — CableListHeaderFormulas, CableListStrikethrough, CableListDataBlock, CableListMerges, LoomGroupSort, WireListRowSort.", { after: 80 }));
children.push(p("Helpers/Continuity — ContinuityHeaderFormulas, ContinuityDataBlock, ConnectionDeduper.", { after: 80 }));
children.push(p("Helpers/Megger — MeggerEndpoint, MeggerSplitWriter, MeggerTextReport.", { after: 80 }));
children.push(p("Helpers/PowerOn — PowerOnExcelBlock, ElectreCollectionFilter, GroundPinIndex."));

children.push(h1("4. Preprocessing"));
children.push(table(
  ["Step", "Before", "After"],
  [
    ["CSV", "Each method read the extraction file", "One ExtractionCsvLoader.Load (64 KB sequential)"],
    ["Row split", "Empty col 28 → general; else panel", "Unchanged"],
    ["Uniques", "List.Contains / AddIfPresent", "HashSet then UniqueNameList.ToSortedList"],
    ["Pairing", "Nested loop, ToUpper every compare", "WirePairIndex; 2-end fast path; inner loop only for 3+ ends"],
    ["Wire code", "Duplicated inline", "PairWireCode.Build — project vs breakdown extra-slash still different"],
    ["Sort", "Connector then pin", "Same keys; stable (ElectreListSort)"],
    ["Panel uniques", "SearchAndAppend", "HashSet; SCB/TCB only for Power On breaker lists (requested)"],
  ],
  [2000, 3873, 3873]
));
children.push(spacer());
children.push(codeBlock("Before — nested pairing (concept)", `for (int i = 0; i < ElecCollection.Count; i++)
{
    var E1 = ElecCollection[i];
    for (int j = 0; j < ElecCollection.Count; j++)
    {
        if (i == j) continue;
        var E2 = ElecCollection[j];
        if (E1.WireNumber.ToUpper() == E2.WireNumber.ToUpper()
            && E1.SubNet.ToUpper() == E2.SubNet.ToUpper())
        {
            // write From/To row
        }
    }
}`, "before"));
children.push(spacer());
children.push(codeBlock("After — WirePairIndex (Helpers/Global/WirePairIndex.cs)", `public static WirePairIndex Build(List<ElectreObject> collection)
{
    var index = new Dictionary<(string, string), List<int>>(
        n, WireSubNetOrdinalComparer.Instance);
    for (int i = 0; i < n; i++)
    {
        var e = collection[i];
        if (string.IsNullOrEmpty(e.WireNumber)
            || string.Equals(e.ComponentType, "SDS", StringComparison.OrdinalIgnoreCase))
            continue;
        var key = (e.WireNumber, e.SubNet ?? "");
        if (!index.TryGetValue(key, out var mates))
        {
            mates = new List<int>(2);
            index[key] = mates;
        }
        mates.Add(i);
    }
    return new WirePairIndex(index, keys);
}

// Two-end fast path: write the other mate, no nested scan.
if (mates.Count == 2)
    WriteOtherEnd(mates[0] == i ? mates[1] : mates[0]);`, "after"));

children.push(h1("5. Excel freeze"));
children.push(p("Screen, events, calculation, print communication, interactive, and animations are off for the Excel session."));

children.push(h1("6. Report-by-report changes"));

children.push(h2("6.1 HAL cable list (CL-n) — former ~45 min / 100 sheets"));
children.push(table(
  ["Item", "Before", "After"],
  [
    ["Workbook", "Workbooks.Add + copy from another workbook", "File.Copy template"],
    ["Copy mode", "Sequential / geometric in .xls BIFF8", "Stamp CL-1, SaveAs .xlsx, copy CL-1 only, name CL-2…"],
    ["Data", "~41,800 Cells.Value", "One Value2 block C–L (CableListDataBlock)"],
    ["ME merge", "Twice per group; Merge on 1 row; rewrite D/K", "Dedupe; skip 1-row Merge; batch Range.Merge; D/K already in block"],
    ["Strike", "Per D-row + Calculate", "Union batches of 20"],
    ["Headers", "Per-sheet formula cells", "Arrays on CL-2, grouped copy"],
    [">250 sheets", ".xls fails at 255", ".xlsx before sheet 256"],
    ["Last sheet", "Old Q loop could repeat last data row", "No duplicate"],
    ["Log", "None", "Cable list '{name}': copy= fill= save= total= sheets=N"],
  ],
  [2000, 3873, 3873]
));
children.push(new Paragraph({ spacing: { before: 160, after: 80 }, children: [new TextRun({ text: "Time (100 sheets)", font: "Arial", size: 20, bold: true })] }));
children.push(table(
  ["Workload", "Before", "After (estimate)"],
  [
    ["100 sheets", "~45 min (measured)", "~2–5 min (copy still largest Excel piece)"],
    ["300 sheets", "Fail at 256", "Minutes as .xlsx"],
  ],
  [2400, 3673, 3673]
));
children.push(p("Unchanged: From/To, ME/UN/LAST look, SHEET R OF N, report-name split F51/K52, year B47, 38 rows/sheet, HAL template.", { before: 120 }));
children.push(spacer());
children.push(codeBlock("Before — HAL copy in .xls + cell-by-cell fill", `ReportWB = ExcelApp.Workbooks.Add();
TempWB = ExcelApp.Workbooks.Open(templatePath);
for (int l = 1; l <= numberOfSheetsRequired; l++)
{
    continuityOrCableSheet.Copy(After: ReportWB.Sheets[Count]);
    ReportWB.Sheets[Count].Name = "CL-" + l;
}
for (int O = 1; O <= 38; O++)
{
    ws1.Cells[row, col].Value = iarr[Q, P];  // ~11 cells × 38 rows × N sheets
    if (iarr[Q, 19] == "D") { ScreenUpdating = true; Calculate(); strike; }
}`, "before"));
children.push(spacer());
children.push(codeBlock("After — xlsx-before-copy + bulk fill (ExcelSheetOps + CableListDataBlock)", `// Stamp CL-1, convert to .xlsx, then copy the original HAL page only:
template.Copy(After: workbook.Sheets[workbook.Sheets.Count]);
((Worksheet)workbook.Sheets[Count]).Name = "CL-" + i;

object[,] block = CableListDataBlock.Build(iarr, Q, rowsThisSheet);
WriteBlock(ws1, startExcelRow, LoomInitialCol + 1, block); // one Value2

// ME: dedupe + batch merge (D/K already in the block)
meOps.Add(CableListMerges.BuildMe(...));
CableListMerges.ApplyMeBatch(ws1, meOps);
ws1.Range["L51:L52"].Value2 = sheetLabel;`, "after"));

children.push(h2("6.2 HAL Continuity (CWOB-n)"));
children.push(table(
  ["Item", "Before", "After"],
  [
    ["Copy", "Sequential in .xls", "Same as cable: xlsx-before-copy, copy CWOB-1"],
    ["Data", "48 × 7 cell writes", "5 bulk writes: C, E:F, H:I, K, N (spacers untouched)"],
    ["Align after fill", "Group-select all sheets", "Skipped — aligned on CWOB-1 before copy; no data merges"],
    ["L4", "Cells[4,12].Value", "L4.Value2"],
    ["Hops", "Full-list scan", "ElectreTraceIndex"],
    ["Dedup From/To", "Local removeRedundants", "ConnectionDeduper (7 cols, trim, skip ####)"],
  ],
  [2200, 3773, 3773]
));
children.push(spacer());
children.push(codeBlock("Before — Continuity cell writes + align every sheet", `ws1.Cells[4, 12].Value = $"{R} OF {N}";
for (int O = 6; O <= 53; O++)
{
    ws1.Cells[O, 3].Value = arr[Q, 0];  // C
    ws1.Cells[O, 5].Value = arr[Q, 1];  // E
    // ... F, H, I, K, N
    alignCellsXl(6, 53); // was even inside the row loop originally
}`, "before"));
children.push(spacer());
children.push(codeBlock("After — ContinuityDataBlock (Helpers/Continuity)", `ws1.Range["L4"].Value2 = $"{R} OF {numberOfSheetsRequired}";
// 5 bulk writes; spacers D/G/J/L/M untouched
ExcelSheetOps.WriteBlock(ws, startRow, 3, colC);   // C
ExcelSheetOps.WriteBlock(ws, startRow, 5, colEF);  // E:F
ExcelSheetOps.WriteBlock(ws, startRow, 8, colHI);  // H:I
ExcelSheetOps.WriteBlock(ws, startRow, 11, colK);  // K
ExcelSheetOps.WriteBlock(ws, startRow, 14, colN);  // N
// Align already stamped on CWOB-1 before copy — no grouped align after fill.`, "after"));

children.push(h2("6.3 Continuity Components"));
children.push(p("Standalone .xls and Megger sheet: bulk A/B + same A-column merges. Header on the already-open file."));

children.push(h2("6.4 Power On"));
children.push(table(
  ["Item", "Before", "After"],
  [
    ["Header", "Open again; center A2:F{Rows.Count}", "Reuse open file; center A:F once"],
    ["Body", "COM inside recursive trace", "Buffer; flush in original merge order (TCB overlap kept)"],
    ["Grounds", "LINQ per SCB", "GroundPinIndex once per collection"],
  ],
  [2000, 3873, 3873]
));
children.push(spacer());
children.push(codeBlock("Before — FindGroundPin LINQ per SCB", `var gndObjects = elecCollection.Where(e =>
    (e.ComponentType.Contains("GROUND") ||
     e.ComponentType.Contains("TER") ||
     e.ComponentType.Contains("TBK")) &&
    e.Panel == source.Panel &&
    e.ConnectorName.StartsWith(source.ConnectorName + "_RTN")).ToList();`, "before"));
children.push(spacer());
children.push(codeBlock("After — GroundPinIndex (Helpers/PowerOn/GroundPinIndex.cs)", `if (!ReferenceEquals(_groundIndexSource, elecCollection) || _groundIndex == null)
{
    _groundIndex = GroundPinIndex.Build(elecCollection);
    _groundIndexSource = elecCollection;
}
var gndObjects = _groundIndex.FindForSource(source);
// Build() keeps GROUND/TER/TBK objects grouped by Panel (collection order).
// FindForSource: panel lookup, then StartsWith(connector + "_RTN").`, "after"));

children.push(h2("6.5 Wire lists without HAL"));
children.push(table(
  ["Report", "Before", "After"],
  [
    ["Project WireList.xls", "Empty SaveAs + Excel sort", "Used rows, in-memory wire-code sort, one SaveAs"],
    ["Sheet {s}_SHEET_Wirelist.xls", "Same per sheet", "Index selected sheets only, one SaveAs"],
    ["EQU/BRK/JM/MISC", "Empty SaveAs + Excel sort per name", "Selected-name index, sort indexes, unrolled copy, one SaveAs"],
    ["Loom without HAL", "Empty SaveAs + 11-col append", "WriteLoomWithoutHalWorkbook — NERD (no space) + Group Number, A:K, one SaveAs"],
  ],
  [2600, 3573, 3573]
));
children.push(p("Loom group-id filter/sort: LoomGroupSort. Incorrect/null group logging and MessageBox unchanged."));
children.push(spacer());
children.push(codeBlock("Before — one file: empty SaveAs + Excel sort", `CreateNewWorkbook(name, sheet, folder);          // SaveAs empty .xls
CreateWirelistReportHeader(path);
AppendToExcel(arrCustomLoom);                    // Save
SortWirelistWireNumber(2, arr.GetUpperBound(0)+1); // Excel COM sort + Save + Close`, "before"));
children.push(spacer());
children.push(codeBlock("After — equipmentReportWireList + WriteComponentBreakdownWorkbook", `var wanted = new HashSet<string>(componentList.Where(c => c != null));
// One scan of used breakdown rows; index only selected names.
matchedRows.Sort((a, b) => string.Compare(
    arrFT_CwithBC[a, 4]?.ToString(),
    arrFT_CwithBC[b, 4]?.ToString(),
    StringComparison.OrdinalIgnoreCase));
// One compact copy, then a single SaveAs (header + data + format):
modExcelInst.WriteComponentBreakdownWorkbook(
    $"{s}_{fileSuffixName}", s, FolderName, arrCustomLoom, z);`, "after"));

children.push(h2("6.6 Loom with HAL (data prep)"));
children.push(p("Cannot use the 10-column SaveAs writer (must keep HAL template). Excel was already one SaveAs after copy+fill. Data prep: index only looms that appear; compact FromConn+WireCode so fill/merge/LAST share loomRowCount; WireCode.Split once (loomSort.WireParts). Log: HAL loom '{name}': rows= sheets= prep+excel="));

children.push(h2("6.7 Megger Scheduler and text files"));
children.push(table(
  ["Item", "Before", "After"],
  [
    ["Workbook", "Add + 7 extra sheets + Open twice", "One sheet + six adds; reuse open file"],
    ["Continuity Components sheet", "Cell-by-cell", "Bulk + merge"],
    ["High pairs", "i≠j + dict of every key + all lines in RAM + sort + write High twice", "Unique ends, i < j, stream temps, 8 MB I/O, no concat key, raw merge copy"],
    ["Low files", "64 KB writer", "Same splitter, 5 million lines/file"],
  ],
  [2400, 3673, 3673]
));
children.push(p("Same: unique undirected High pairs; Lows from continuity; headers; Megger_ / HighMegger_ / LowMegger_ names."));
children.push(p("Different (required for crores of lines): High From/To is stable (i < j), not a parallel race; no global sort by FromConnector. Log: MeggerSheet5: ends= combined= high= elapsed="));
children.push(spacer());
children.push(codeBlock("Before — MeggerSheet5 (RAM bag + i≠j + sort)", `for (int i = range.Item1; i < range.Item2; i++)
    for (int j = 0; j < count; j++)
    {
        if (i == j) continue;
        string key = string.Compare(partA, partB) < 0 ? partA+"##"+partB : partB+"##"+partA;
        if (!redundancySet.TryAdd(key, 0)) continue;
        highMeggerData.Add(connectionLine + ";" + status); // all lines in RAM
    }
highMeggerData = new ConcurrentBag(highMeggerData.OrderByDescending(...));
// then write Megger_*.txt AND rewrite High again in Sheet6`, "before"));
children.push(spacer());
children.push(codeBlock("After — MeggerTextReport (Helpers/Megger)", `for (int i = worker; i < count; i += dop)
    for (int j = i + 1; j < count; j++)   // unique undirected pairs
    {
        if (sameWireAndSubnet) continue;
        if (hasLowKeys && lowMeggerKeys.Contains((e1.Conn, e1.Pin, n1.Conn, n1.Pin)))
            continue;
        sb.Append(e1.Conn).Append(';') /* ... */ .Append(";High Megger").AppendLine();
    }
meggerWriter.AppendRawFile(tempPath);  // byte copy, 5M-line split
highWriter.AppendRawFile(tempPath);
// Sheet6 High skipped — files already written`, "after"));

children.push(h2("6.8 Material list HAL (ML-n)"));
children.push(p("Bulk OOTB read, geometric/xlsx copy, bulk B/D/E/G/H/K. Not called from current Panel Drawing UI (live path is .txt)."));

children.push(h1("7. Shared helpers"));
children.push(table(
  ["Helper", "Replaces"],
  [
    ["UniqueNameList", "SearchAndAppend / ToSortedList"],
    ["WirePairIndex", "Nested mate scan"],
    ["PairWireCode", "Duplicated wire-code builders"],
    ["ElectreListSort", "Connector+pin sort"],
    ["ExcelSheetOps", "Sheet copy, WriteBlock, FindByName"],
    ["ConnectorTraceHops", "Duplicate EQU/DIS/TBK/SPL/TER hops"],
    ["ConnectionDeduper", "removeRedundants (7 cols) vs Megger (4 cols, no trim)"],
    ["ElectreCollectionFilter", "Panel/sheet/loom HashSet filters"],
    ["GroundPinIndex", "LINQ FindGroundPin per SCB"],
    ["LoomGroupSort", "Per-loom LINQ group sort"],
    ["CableListMerges", "Per-group Merge B+D+K"],
  ],
  [3600, 6146]
));
children.push(p("modMain / modExcel still orchestrate. SearchAndAppend and MeggerSheet5 remain thin public wrappers."));
children.push(spacer());
children.push(codeBlock("After — shared hops (Helpers/Global/ConnectorTraceHops.cs)", `public static ElectreObject TracePinofEQUConnector(
    ElectreObject eleObj, ElectreTraceIndex index, List<ElectreObject> fallback)
{
    string last = ElectreTraceIndex.MapEquConnectorName(eleObj.ConnectorName);
    if (index != null)
        return index.FindByConnectorPin(last, eleObj.PinNumber);
    return fallback?.FirstOrDefault(w =>
        string.Equals(w.ConnectorName, last, StringComparison.OrdinalIgnoreCase)
        && w.PinNumber == eleObj.PinNumber);
}
// Continuity passes loom index; Power On passes AllIndex / ElecCollection_All.`, "after"));

children.push(h1("8. String compares"));
children.push(p("ToLower() == / ToUpper() == became StringComparison.OrdinalIgnoreCase (sheets, SCB/TCB, filters, LOC-, MS/ML-, SDS, Linked). Pairing dictionary uses WireSubNetOrdinalComparer (no ToUpper keys)."));
children.push(p("Left: switch (ComponentType) \"SCB\" / \"TCB\" (CSV already that casing). CableTypes joined-string \"__\" for empty type."));

children.push(h1("9. Cell maps (golden-file checks)"));
children.push(h2("HAL cable list — rows 9–46"));
children.push(p("C–K = iarr 1–9; L = layer 19; B = ME/UN/LAST serial; D = length; L51:L52 = SHEET R OF N; F51/K52 name split on CL-1 only."));
children.push(h2("HAL Continuity — rows 6–53"));
children.push(p("C, E, F, H, I, K, N from arrFTcwob 0–6. D, G, J, L, M not written. L4 = {R} OF {N}."));
children.push(h2("Wire list / breakdown A–J"));
children.push(p("FROM CONN … NERD (leading space). Loom without HAL adds K = Group Number; NERD has no leading space."));

children.push(h1("10. Intentional behaviour differences"));
children.push(table(
  ["Item", "What changed"],
  [
    ["Power On from panels", "SCB and TCB only (requested). SWT/ERM/IND from panel objects not added. Panel EQU pins not appended in MainFunction_withPanels (general EQU from MainFunction unchanged)."],
    ["Megger High From/To", "Stable i < j, not whichever parallel worker won. No global sort."],
    ["Cable list last sheet", "No duplicate of the previous last data row."],
    ["OrdinalIgnoreCase vs ToLower()", "Same for ASCII Electre names; can differ only for culture-specific casing."],
  ],
  [2800, 6946]
));

children.push(h1("11. Templates and environment"));
children.push(p("ELECTRE_CUSTOMIZE\\system\\ must still contain CABLE_ReportFormat.xls, CONTINUITY_ReportFormat.xls, and MATERIALLIST_ReportFormat.xls (if HAL material list is used). Do not replace these templates as part of this drop."));

children.push(h1("12. Production send, run, and log"));
children.push(p("Send the whole folder bin\\Debug\\net8.0-windows\\ (build with Visual Studio MSBuild; do not use dotnet build on the no-internet PC)."));
children.push(p("Run the new Electre_Customize_DotNet.exe. Keep the old shortcut until sign-off."));
children.push(p("Log: {ELECTRE start}\\Logs\\{yyyy-MM-dd}.log"));
children.push(table(
  ["Log line", "Meaning"],
  [
    ["MainFunction: total=", "CSV + lists"],
    ["MainFunction_withPanels: csv+sort= projectPair= breakdownPair=", "Pairing"],
    ["ReportExcecution: all reports total=", "Whole Generate click vs old ~45 min"],
    ["Cable list: copy= fill= save=", "HAL cable write"],
    ["Continuity: copy= fill= save=", "HAL Continuity write"],
    ["HAL loom: rows= sheets= prep+excel=", "Loom HAL prep + Excel"],
    ["MeggerSheet5: ends= high= elapsed=", "Pair gen + .txt"],
    ["Power On: total=", "Power On write"],
    ["Component breakdown: total=", "One EQU/BRK file"],
    ["Project / Sheet / Loom wire list", "One-SaveAs paths"],
  ],
  [5200, 4546]
));
children.push(p("Pass: spot-check From/To, merges, Continuity cells, Power On SCB/TCB list; cable list 100 sheets total= in minutes not ~45 min. Rollback: old exe. No data migration."));

children.push(h1("13. What is not pending"));
children.push(p("Do not change these without a new log: HAL copy= (one formatted page per extra sheet); cable HAL UN/LAST/strike; Megger crores of High lines (disk); one .xls per component/sheet/loom; commented dead methods; Material HAL Excel (not in current UI)."));
children.push(p("Until a new log names a stage, do not add more Excel COM, HashSet, or Megger pair-logic edits."));

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Arial", color: NAVY },
        paragraph: { spacing: { before: 280, after: 140 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "Arial", color: "2E75B6" },
        paragraph: { spacing: { before: 220, after: 100 }, outlineLevel: 1 } },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: PAGE_W, height: PAGE_H },
        margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN },
      },
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: NAVY, space: 6 } },
          spacing: { after: 120 },
          children: [
            new TextRun({ text: "Electre Customization  |  Optimization Change Document", font: "Arial", size: 16, color: NAVY }),
          ],
        })],
      }),
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          border: { top: { style: BorderStyle.SINGLE, size: 6, color: "BDD7EE", space: 6 } },
          spacing: { before: 80 },
          children: [
            new TextRun({ text: "Confidential  ·  Page ", font: "Arial", size: 16, color: "808080" }),
            new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 16, color: "808080" }),
          ],
        })],
      }),
    },
    children,
  }],
});

const outPath = path.join(__dirname, "..", "Electre_Optimization_Changes.docx");
Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync(outPath, buffer);
  console.log("Wrote " + outPath);
});
