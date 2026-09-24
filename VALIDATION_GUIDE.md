# Validating the report-generation performance changes

This file is a working checklist for testing the uncommitted performance changes against
real production data, before you merge anything into `main`. It's not meant to be a
permanent doc in the repo — delete it once you're done, or keep it if useful.

## 0. What you're validating

All changes are currently **uncommitted, in your working tree** (see `git status`). No commits,
no branch merge. Everything is gated so the app behaves identically to today unless you
explicitly turn diagnostics on — the report logic itself was not touched, only *how* the code
reaches the same output faster (bulk Excel writes instead of per-cell, lookup indexes instead
of repeated full-list scans, a few HashSet/memoization swaps).

## 1. Baseline first, changes second

Before touching anything, capture a baseline with **today's code** (i.e. `git stash` your
changes, or check out `main`):

1. Run each report type you care about once (Continuity, Power-On, Megger, Cable List, at
   minimum) against a real job.
2. Keep those output files somewhere separate (copy the report folder to e.g.
   `Baseline_Output\`) — you'll eyeball-diff or fingerprint-diff against these later.
3. Note how long each report took (a stopwatch or just the loading form's visible timing is
   fine for baseline — the new code adds proper logging, see below).

Then restore the changes (`git stash pop`, or re-apply the working tree) before continuing.

## 2. Turn on diagnostics

Edit `Electre_Customize_DotNet\App.config` and set both to `true`:

```xml
<add key="PerfDiagnostics" value="true" />
<add key="PerfFingerprint" value="true" />
```

Both default to `false` — leaving them off is the safe/normal state for actual production use.
Only enable them for this validation pass, then set them back to `false` before anyone runs the
app for real work (the fingerprint capture in particular reads whole workbooks and adds
overhead — not something you want on by default).

Rebuild after changing `App.config` (it's copied to the output folder at build time).

## 3. Run the same job with diagnostics on

Run the SAME report(s), against the SAME input data, with the changed code and diagnostics
enabled. Two new kinds of files appear under `<StrtCmd>\Logs\`:

- **`perf_YYYY-MM-DD.log`** — one line per report-generation phase, e.g.:
  ```
  [2026-...] [PERF] report=Continuity phase=ReportGeneration elapsedMs=842 SelectedLoomCount=12 InputCount=48213
  ```
  This is your timing evidence. Compare `elapsedMs` across runs/job sizes.

- **`fingerprint_<Label>_<timestamp>.json`** — one per generated worksheet, capturing sheet
  names, used-range dimensions, every cell's `Value2`, merged-cell ranges, and column widths, in
  a diffable JSON structure.

## 4. Compare fingerprints: old vs new

This is the actual correctness check. You need a fingerprint from **both** the old code and the
new code, run against **identical input data**, then diff them.

Simplest approach given no internet / no external tooling available:

1. With `main` (old code) + diagnostics on: run a report, note the fingerprint file(s)
   produced, rename/move them to a folder like `Fingerprints_Before\`.
2. With the changed code + diagnostics on: run the *same* report against the *same* input,
   move the new fingerprint file(s) to `Fingerprints_After\`.
3. Diff them. `fc` (Windows' built-in file compare) works fine on the JSON text directly:
   ```powershell
   fc Fingerprints_Before\fingerprint_Continuity_....json Fingerprints_After\fingerprint_Continuity_....json
   ```
   Or open both in Notepad++/VS Code and use its compare feature if available offline.
4. **Expect zero differences** in `Values`, `MergedRanges`. `ColumnWidths` may legitimately
   differ only for the reports where I consolidated per-row `AutoFit()` calls — check the
   commit/change notes for which those are (Cable List / Continuity List sheet loops) and treat
   a `ColumnWidths` difference there as something to eyeball in Excel, not necessarily a bug.

## 5. Specific things worth extra scrutiny

Given the size of this change, prioritize checking these (highest risk first):

1. **Continuity report + Continuity Components sheet** — the traversal-index rewrite
   (`ElectreTraversalIndex`) is the highest-risk change in this whole set. Test with:
   - A drawing/loom with duplicate wire numbers on the same subnet
   - A component with multiple EQU or DIS counterparts
   - At least one full production job (300-500 drawings), not just a small sample
2. **Power-On report** — same traversal index, but against the full project collection. Test a
   source with no ground/TER/TBK candidate (the "missing ground pin" case).
3. **Megger "Connection List" sheet** — check specifically for a stray blank/empty-looking row.
   I intentionally preserved an existing quirk (one blank sentinel row leaking into the Megger
   output) rather than "fixing" it — if you see exactly one blank row in Connection List, that's
   expected and present in the old code too; if you see it in the OLD output but not the NEW
   output (or vice versa), that's a real regression to report back.
4. **Cable List / Continuity List sheets** — visually check column widths look right (see the
   AutoFit note in step 4).

## 6. Reading the timing numbers

Once correctness checks pass, `perf_*.log` gives you the actual speedup evidence. Useful
comparisons:
- Same job, old code vs new code, same `report=` name → direct elapsedMs comparison.
- Look at `InputCount`/`SelectedLoomCount` alongside elapsedMs to see how the numbers scale as
  job size grows — the whole point of this work was making large jobs (300-500 drawings)
  scale much better than before, so the improvement should be more dramatic on your biggest
  jobs than on small test ones.

## 7. Once you're satisfied

Turn `PerfDiagnostics` and `PerfFingerprint` back to `false` in `App.config`. Let me know the
results and I'll commit the changes (or fix whatever the fingerprint diff turns up) before
anything gets merged into `main`.

## 8. Megger text files (MeggerSheet5 / MeggerSheet6and7) - how to validate

These two methods were redesigned (streaming writer, see the `Helpers/Megger/` folder), so validate
them differently from the other reports:

- Do NOT expect a byte-for-byte diff against the old output. The old code was non-deterministic: it ran a
  parallel loop where whichever thread claimed a pair first decided its orientation (A;..;B vs B;..;A) and,
  for pairs whose reverse orientation was in the Low list, whether it appeared at all. Two runs of the OLD
  code on the same data could already differ. The new output is deterministic.
- Compare instead: (1) total line counts of Megger_*.txt, HighMegger_*.txt, LowMegger_*.txt;
  (2) the set of unordered pairs (treat A;pa;B;pb and B;pb;A;pa as the same pair). These should match
  except for the small race-dependent group described above.
- Low lines come first in Megger_1.txt, then the High lines; files roll every 5,000,000 lines; every file
  starts with a UTF-8 BOM and the two header lines - unchanged.
- The High lines are now written in ASCENDING (connector, pin) order. The old code sorted descending but
  then rebuilt a ConcurrentBag, which enumerates in reverse, so the old files were ascending too.
- A filtered connector/pin that appears on 2+ rows with different WireNumber/SubNet still produces one
  "X;pin;X;pin" self line, as before.
- Optional app setting `MeggerWriteParallelism` (default = min(4, cores/2)) controls how many 5M-line
  files are written at once. Use 1 on a spinning disk.


## 9. Optional: faster template sheet copies (`IntraWorkbookSheetCopy`)

Cable List (`CL-n`) and Continuity (`CWOB-n`) reports create one sheet per ~40 rows by copying the template
sheet from the template workbook. With many sheets that is usually the slowest single step.

* `App.config` key `IntraWorkbookSheetCopy` is `false` by default = exactly the old behaviour.
* Set it to `true` to copy only the FIRST sheet from the template and the rest from that first copy inside the
  report workbook (much cheaper). At that moment no data has been written, so the copies should be identical.
* To adopt it: generate the same large Cable List and Continuity report with the key `false` and then `true`,
  and compare them with the fingerprint tool (values, merged ranges, number formats, column widths) plus a
  visual check of a few sheets (page setup, fonts, borders). Keep it `false` if anything differs.


## 10. Optional: one-pass workbooks for sheet / component wire-lists (`OnePassWirelistWorkbook`)

Sheet-wise (`*_SHEET_Wirelist`) and component-wise (Component Specific Breakdown) reports create one small
workbook per sheet / component. `App.config` key `OnePassWirelistWorkbook` (default `false` = original sequence).

* `true`: each workbook is created with a single sheet, gets the same header row / column widths, the same
  data block, the same formatting and the same Excel sort (so the same row order), and is saved ONCE with
  `SaveAs` (the original does SaveAs, re-Open, several Saves and a final Close(true)).
* To adopt it: generate the same component breakdown and sheet wire-list reports with the key `false` and
  `true`, and compare them with the fingerprint tool (values, merged ranges, number formats, column widths)
  plus a look at a few files (header bold, widths, wrap/centering, sort order). Keep `false` if anything differs.

## 11. Other Cable List changes to spot-check

* Loom (non-template) wire-list workbooks are now closed after they are saved (they used to stay open until
  Excel quit). The files should be identical.
* On sheets 2..n of the Cable List and Continuity templates, the linked cells C48:C52, D48:D52, K51:K52 (Cable
  List) and C56:C57 (Continuity) are now assigned with one range formula each instead of one write per cell.
  Compare a few of those cells on a later sheet: each should still point at the same cell on sheet 1
  (for example C49 = 'CL-1'!C49).


## 12. Optional: parallel Excel instances (`ExcelWorkerCount`)

Loom (with and without template), sheet-wise and component-wise (Component / Break Connector / Junction Module / Misc) wire-list reports create one independent workbook per loom / sheet.
`App.config` key `ExcelWorkerCount` (default `1` = old sequential behaviour) lets 2-8 separate Excel processes
generate them at the same time - the only way to use more than one core for the Excel part of a huge job.

* Try `2`, then `4`, on the same large job and compare the generated files with the `1` run (content, merges,
  formats). Only the creation order should differ.
* Each Excel process needs its own memory; watch RAM. Message boxes (e.g. incorrect cable group id) appear from the
  worker that hit them and only block that worker.
* If a worker cannot start its own Excel, the remaining looms/sheets are finished on the normal instance.


## 13. New: `.xls`→`.xlsx` above 255 sheets (`AutoXlsxAboveSheetLimit`)

Cable List / Continuity reports that need more than 255 sheets now save as `.xlsx` instead of failing outright
(default on). To check: generate a report you know needs >255 sheets (e.g. a very large project-wise Cable
List) and confirm a `.xlsx` file is produced and opens correctly in Excel. Reports at or under 255 sheets are
completely unaffected - same `.xls`, same everything.
