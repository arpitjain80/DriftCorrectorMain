# PROJECT MEMORY — DriftCorrector
> Auto-maintained by Claude. Last updated: 2026-03-16
> [2026-03-16] — Updated: Section 6 (NormSearch nuance) because PATH 3 false-substring-match fix added
> PURPOSE: This file is Claude's persistent memory. It captures architecture,
> logic flows, and non-obvious nuances so future sessions never start from scratch.

---

## 1. TECH STACK & ENTRY POINTS

**Runtime & Framework**
- .NET 6+ console app (top-level statements in `Program.cs`)
- C# 10+

**Key Libraries**
- `Aspose.Words` (dual-version strategy: V14 and V23) — Word document loading, manipulation, and PDF export
- `iText7` (`iText.Kernel`, `iText.Forms`) — PDF text extraction, field extraction, annotation reading/writing
- `Docnet.Core` — referenced in `PDFProcessor.cs` imports (low-level PDF page rendering)
- `System.Text.Json` — JSON serialization for diff reports and file-state caching
- `System.Drawing` — rectangle math and image operations for diff report rendering

**Entry Point & Bootstrap Sequence**
1. `DriftCorrector/Program.cs` — top-level statements; no `Main()` method
2. `DocumentService.SetLicense(licenseKeyPath)` — MUST be first; sets Aspose.Words license (thread-safe double-lock)
3. `EnsureEncodingProviderRegistered()` — called inside `SetLicense`; registers `CodePagesEncodingProvider` required by .NET Core
4. `ComparePDF()` — primary workflow function defined as a local static function in Program.cs

**Active entry-point functions (Program.cs)**
- `ComparePDF()` — full pipeline: compare → correct → re-render → compare again *(currently active)*
- `ComparePDF_Raw()` — comparison without word template context *(commented out)*
- `ConvertWordToPDFV23()` — standalone Word→PDF batch converter *(commented out)*
- `CompareOnly()` — raw compare with no correction loop *(commented out)*

---

## 2. FOLDER STRUCTURE & MODULE RESPONSIBILITIES

```
DriftCorrector/
├── Program.cs              ← Entry point; hardcoded paths; orchestration calls
├── PDFProcessor.cs         ← PdfComparer class: PDF extraction, comparison, HTML/JSON report generation
├── WordDriftCorrector.cs   ← WordDriftCorrector class: reads JSON diffs, applies corrections to .docx
├── DocumentService.cs      ← DocumentService class (partial): Word→PDF conversion, file management, workflow orchestration
└── Logger.cs               ← CustomLogger: simple append-to-file logger
```

**PDFProcessor.cs — `PdfComparer`**
- Owns: PDF text/field extraction, text/field diffing, JSON diff report generation, HTML report generation, folder-recursive comparison, diff statistics, "WhiteMasked" and "Ignore" annotation logic
- Does NOT own: Word document manipulation, file-state caching, process execution

**WordDriftCorrector.cs — `WordDriftCorrector`**
- Owns: loading `*_diff.json` reports, building correction plans, applying X/Y drift corrections to `.docx` files via Aspose.Words (table indent, cell padding, column widths, row heights, paragraph indent, SpaceBefore, header/footer corrections)
- Does NOT own: PDF extraction, report rendering, Word→PDF conversion

**DocumentService.cs — `DocumentService` (partial)**
- Owns: `ConvertDocToPdf`, `ConvertWordFolderToPDF` (with incremental state), `ConvertWordFolderToPDF_X` (simple), `ConvertFolderWordToPdf_V23`, `ExecuteDocumentComparisonWorkflow` (full pipeline as a reusable method), `ExecuteRobocopy`, `RunExternalExeAsync`, `CopyFiles`, `ApplyShifts` (legacy XML-based correction), `DiagnoseAsposeVersion`, `SetLicense`, `EnsureEncodingProviderRegistered`
- Does NOT own: PDF diff logic, Word correction strategy selection

**Logger.cs — `CustomLogger`**
- Owns: timestamped `Log(message)` and `Clear()` to a single log file path
- Silent on failure (exception swallowed in catch block)

---

## 3. CORE DATA FLOWS

### Flow A — Full Drift Correction Pipeline (`ComparePDF()`)

```
1. DocumentService.SetLicense()
         ↓
2. new PdfComparer(flags, logPath)
         ↓
3. PdfComparer.CompareFolders(V14Folder, V23Folder, ...)
   ├─ ExtractFields() + ExtractText() for each PDF pair
   ├─ CompareFields() + CompareText() → List<FieldDiff>, List<TextDiff>
   ├─ GenerateJsonDiffReport() → writes *_diff.json to JSONDiff subfolder
   │    Classifies each word diff as PhraseGroup with IsCleanDiff / IsXIgnore
   └─ GenerateHtmlReport() → per-file HTML + DiffRectangleInfo list
         ↓ returns (ComparisonDetails: List<FileComparisonResult>, ...)
4. new WordDriftCorrector(logPath)
         ↓
5. WordDriftCorrector.ApplyDriftCorrections(wordTemplateFolder, jsonDiffFolder, modifiedFolder)
   ├─ Loads each *_diff.json
   ├─ PATH 1: docxContext table corrections (uses DocxContext.TableIndex)
   ├─ PATH 2: JSON tableGroups fallback (uses TableDiffGroup from JSON)
   ├─ PATH 3: Paragraph corrections (PhraseGroups with LayoutContext=PARAGRAPH/LIST)
   └─ PATH 4: Header/footer corrections
         ↓ writes corrected .docx to modifiedFolder
6. DocumentService.ConvertWordFolderToPDF(modifiedFolder, modifiedV23PDFFolder)
         ↓
7. new PdfComparer(flags, logPath) — second instance with IsModifiedComparison=true
         ↓
8. PdfComparer.CompareFolders(V14Folder, modifiedV23PDFFolder, ..., priorDiffRects)
   └─ ComputeStats() compares priorDiffRects vs new rects → FileDiffStats per file
         ↓
9. Final HTML report written to modifiedReportPath
```

### Flow B — JSON Diff Report Generation (`GenerateJsonDiffReport`)

```
text1 (V14 words) + text2 (V23 words) [pre-filtered: whitespace + ignore regions removed]
         ↓
Build t2LineGroups: bucket validText2 by (page, round(Y/5)*5) for fast lookup
         ↓
For each V14 word: FindBestT2Match() → MATCH_Y_TOL=5pt, MATCH_X_TOL=60pt, score=|ΔY|*10+|ΔX|
  → POSITION_SHIFT if |dx|>0.5 or |dy|>0.5
  → Reflow check: if dx < -12pt AND pubX <= pageMinX+36pt AND baseline was >15pt right of margin
      → suppress POSITION_SHIFT, add to cleanWordPositions
  → FONT_SIZE_DIFF if |Δfont| >= 0.25pt; FONT_DIFF if font name changed
  → Unmatched V14 words → MISSING entry; Unmatched V23 words → EXTRA entry
         ↓
BuildPhraseGroups() → groups POSITION_SHIFT words by same-page, ΔY<3pt, adjacent X<30pt, ΔX within 1pt
         ↓
TagInlineRuns() → phrase groups on same Y-line with DISTINCT X-deltas → mark non-TABLE_CELL as INLINE_RUN
         ↓
DetectTableGroups() → PDF heuristic: Y-cluster rows (6pt) → X-cluster columns (8pt) → grid assignment
         ↓
BackFillTableInfoToWords() → propagates TableGroupId/Column/Row/LayoutContext to WordDiffEntry
         ↓
AugmentTableGroupsFromDocxContext() [if docxElements present] — validates/creates table groups from Word DOM
         ↓
Stamp IsCleanDiff on each PhraseDiffGroup (4 gates):
   Gate 1: Reflow (dx < -12pt AND pubX ≤ pageMinX+36pt)  → IsCleanDiff=true
   Gate 2 (partial): INLINE_RUN context               → IsCleanDiff=true
   Gate 3: IsModifiedComparison && !TableGroupId      → IsCleanDiff=true (non-table groups in modified comparison)
   Gate 4: |dy| > 8pt (Y-reflow / vertical overflow)  → IsCleanDiff=true
   Note: TABLE_CELL groups deliberately deferred to Phase 3 so XI detection can still see them
         ↓
XI reconciliation (same 3-option pivot logic as GenerateHtmlReport):
   Option A: mode of free-text X (words not in any X-drift phrase group)
   Option B: mode of green (low-drift) phrase group XStart values
   Option C: removed (no XDiff-only pivot)
   Leftmost X-drift groups where bx < pivotX - 3pt → IsCleanDiff=true, IsXIgnore=true
         ↓
Phase 3: Remaining table-affiliated (TABLE_CELL/PARAGRAPH/LIST_ITEM/MERGE_FIELD) phrase groups
   (skip IsXIgnore groups) → IsCleanDiff=true
         ↓
Serialize to *_diff.json: ReportMetadata, Summary, TableGroups[], PhraseGroups[], WordLevelDifferences[]
Returns (phraseGroups, cleanWordPositions) for use by GenerateHtmlReport
```

### Flow C — Incremental Word→PDF Conversion (`ConvertWordFolderToPDF`)

```
jsonStateFilePath (optional) → load Dictionary<absPath, FileState{Size, LastModified}>
         ↓
For each .doc/.docx in sourceRoot (recursive):
   Compare Size + LastWriteTimeUtc vs saved state
   If changed → ConvertDocToPdf() → update state
         ↓
Phase 2: Delete orphaned PDFs in destRoot (no corresponding source file)
         ↓
Phase 3: Remove stale JSON state keys (source files no longer present) → save JSON
```

---

## 4. KEY ARCHITECTURAL PATTERNS & DECISIONS

**Dual-pass comparison**
The pipeline runs twice: once against the original V23 PDF (finding drifts), once against the corrected V23 PDF (verifying fixes). The `priorDiffRects` dictionary bridges the two runs, enabling `ComputeStats()` to measure rectification percentage.

**`IsModifiedComparison` flag on `PdfComparer`**
Controls Gate 3 behavior. `false` (default, original comparison) — Gate 3 is disabled for table groups so real correctable diffs stay red. `true` (modified comparison) — Gate 3 marks unfixed diffs as acceptable/green because the corrector couldn't handle them. Always set this flag before the second `CompareFolders` call.

**4-PATH correction hierarchy in `WordDriftCorrector`**
- PATH 1: Uses `DocxContext.TableIndex` from per-word JSON metadata (most precise)
- PATH 2: Uses `TableDiffGroup` from JSON (fallback when DocxContext matching failed)
- PATH 3: Body paragraph `LeftIndent` / `SpaceBefore` (non-table, non-header text)
- PATH 4: Header/footer `LeftIndent` / `SpaceBefore`
A phrase group handled by an earlier PATH is tracked in `handledPhraseGroupIds` and skipped by later PATHs.

**Strategy enum for table correction**
`CellPadding | TableIndent | ColumnWidths | TableIndentFallback`
Selected by `DeriveStrategy()` heuristically, or overridden by `CorrectionStrategy` field in JSON. Strategy choice is determined once per table plan and cannot change mid-application.

**Nested table ancestry correction**
`BuildTableAncestryMap()` builds a parent→child tree. When a nested table is corrected, `GetAncestorAppliedXPt()` sums all ancestor corrections already applied, and `effectiveDriftX = plan.UniformDriftXPt - ancestorX` prevents double-correcting nested content. Tables are processed outermost-first (ordered by ancestry depth).

**Recursive folder mirroring**
Both `CompareFolders` (PDFProcessor) and `ApplyDriftCorrections` (WordDriftCorrector) recurse through subfolders, mirroring the source directory tree into report/output directories. The recursion is driven by the source (Word template or V14 PDF folder), and subfolders with no counterpart in the destination are silently skipped.

**XI (XIgnore) gate — table skip logic**
If any phrase group inside a Word table is marked `IsXIgnore` (already at the correct horizontal position), the entire table is skipped from correction. Moving the table would shift the already-correct content away from its reference position.

**Thread-safe license and encoding initialization**
Both `_licenseSet` and `_encodingProviderRegistered` use double-checked locking. `EnsureEncodingProviderRegistered()` must be called before any Aspose.Words operation on .NET Core.

---

## 5. INTER-MODULE DEPENDENCY MAP

```
Program.cs
  └─► DocumentService.SetLicense()
  └─► PdfComparer (PDFProcessor.cs)
        └─► CustomLogger (Logger.cs)
        └─► WordLevelExtractionStrategy (PDFProcessor.cs — inner class)
  └─► WordDriftCorrector (WordDriftCorrector.cs)
        └─► CustomLogger (Logger.cs)
        └─► Aspose.Words (external)
  └─► DocumentService (DocumentService.cs)
        └─► Aspose.Words (external)
        └─► PdfComparer (via ExecuteDocumentComparisonWorkflow)
        └─► WordDriftCorrector (via ExecuteDocumentComparisonWorkflow)

DocumentService.ExecuteDocumentComparisonWorkflow()
  ← encapsulates the same pipeline as ComparePDF() in Program.cs
  ← called by DriftCorrectorWinForm (separate project, not analyzed here)
```

**Shared data contracts (JSON model classes in PDFProcessor.cs or WordDriftCorrector.cs)**
- `WordDiffReport` — root: `ReportMetadata (DiffReportMeta)`, `Summary (DiffReportSummary)`, `TableGroups[]`, `PhraseGroups[]`, `WordLevelDifferences[]`
- `WordDiffEntry` — per-word diff: `DifferenceId`, `PrimaryType`, `Issues[]`, `Text`, `Page`, `LayoutContext`, `Baseline/Publish (WordPosition)`, `Delta (DeltaValues)`, `NegativeCorrection`, `PhraseGroupId`, `TableGroupId`, `TableColumn`, `TableRow`, `DocxContext`
- `PhraseDiffGroup` — spatial phrase cluster: `GroupId`, `Page`, `PhraseText`, `WordCount`, `Issues[]`, `LayoutContext`, `BaselineRegion/PublishRegion (PhraseRegion)`, `SharedDelta`, `NegativeCorrection`, `WordIds[]`, `TableGroupId`, `TableColumn`, `TableRow`, `DocxTableIndex`, `IsCleanDiff`, `IsXIgnore`
- `TableDiffGroup` — PDF heuristic table: `CorrectionStrategy`, `TableId`, `Page`, `RowCount`, `ColumnCount`, `ColumnDrifts[] (TableColumnDrift)`, `RowDrifts[] (TableRowDrift)`, `BaselineBounds/PublishBounds (PhraseRegion)`, `WordCorrectionHint`, `PhraseGroupIds[]`
- `DocxContext` — Word element context attached to each `WordDiffEntry`: all `DocxElementInfo` fields projected to JSON, plus `CorrectionTargets[] (DocxCorrectionTarget)`
- `DocxCorrectionTarget` — per-property fix target: `Property`, `Axis (X|Y)`, `CurrentValuePt`, `NewValuePt`, `CurrentValueTwips`, `NewValueTwips`, `CorrectionEmu (nullable)`, `Description`
- `PdfComparisonConfig` — config bag defined in `DocumentService.cs`, used by `ExecuteDocumentComparisonWorkflow()`
- `FileState` — defined in `DocumentService.cs`, serialized to JSON for incremental conversion state

---

## 6. TRICKY LOGIC & NON-OBVIOUS NUANCES

### BUG #6 — Zero-indent layout tables (FIXED)
**File:** `WordDriftCorrector.cs`, `DeriveStrategy()`

Old code: `if (absTableInd > 0 || tableInd == 0)` — the `tableInd == 0` branch was always true, routing ALL zero-indent tables to `TableIndent`, which then set `table.LeftIndent = -36pt`, clipping content into the left margin.

**Current fix:** Zero-indent tables where structural properties don't explain the drift now use `CellParagraphIndent` strategy (paragraphs inside cells get `LeftIndent` adjusted instead of the table itself).

**What "layout tables" are:** Borderless Word tables with `LeftIndent = 0` used to create hanging-indent or columnar text layouts. The cell paragraph indent carries the positional meaning, not the table indent.

### BUG #7 — Two-pass fallback in `ApplyParagraphCorrection` (FIXED)
**File:** `WordDriftCorrector.cs`, `ApplyParagraphCorrection()`

PATH 3 skips cell paragraphs with `if (para.GetAncestor(NodeType.Cell) != null) continue`. When `MatchWordToDocxElement` fails for a word inside a table cell, that word gets `LayoutContext="PARAGRAPH"` with no `DocxContext.TableIndex` and lands in PATH 3. The guard made it unreachable.

**Current fix:** Two-pass approach — Pass 1 tries body paragraphs; Pass 2 falls back to cell paragraphs if not found in body. No X-drift cap applied to cell paragraph fallback (hanging indents can be large).

### BUG #8 — `ApplyColumnWidths` and gridSpan cells (FIXED)
**File:** `WordDriftCorrector.cs`, `ApplyColumnWidths()`

Setting `CellFormat.Width` on a merged cell (gridSpan > 1) causes Aspose.Words to remove `gridSpan`, rebuild `tblGrid`, and corrupt the entire table layout.

**Two-check detection (either fires to protect the table):**
- Check 1: Width variance — `Cell[prevIdx].CellFormat.Width` differs across rows by > 1pt → merged cells present
- Check 2: Row cell-count — any row has fewer cells than `expectedColCount` → spanning structure

**When either check fires:** Width modification is SKIPPED, but `LeftIndent` correction still applies.

### BUG #9 — `SpaceBefore` correction was inverted (FIXED)
**File:** `WordDriftCorrector.cs`, `ApplyRowHeights()`, `ApplyParagraphCorrection()`, `ApplyHeaderFooterCorrection()`

PDF Y increases upward. SpaceBefore pushes content DOWN (decreases Y). If `driftY < 0` (V23 content is lower than V14), excess SpaceBefore must be REDUCED: `ns = sb + driftY`.

Old (buggy): `ns = sb - driftY` — this increased SpaceBefore when driftY < 0, pushing V23 content further down.

**Guard:** Only reduce existing SpaceBefore (`sb >= MIN_DRIFT`), never add it from zero. When `driftY > 0` and `sb = 0`, the formula would add SpaceBefore from nothing, creating blank gaps.

### Cumulative list-spacing sweep (`SweepListSpaceBefore`)
**File:** `WordDriftCorrector.cs`

V23 applies SpaceBefore for every list item; V14 applies it only for the first. Drift accumulates: row N has drift ≈ N × SpaceBefore. When detected drift > single SpaceBefore unit and new SpaceBefore ≈ 0, `SweepListSpaceBefore()` zeroes SpaceBefore on ALL list-paragraph rows in the table (except row 0 and the already-fixed row).

### `WhiteMasked` annotation type
**File:** `PDFProcessor.cs`, `ExtractText()`

PDF annotations with subject "WhiteMasked" mark regions where text is visually covered by a white box. These are used by `RemoveText()` to overlay white rectangles. During text extraction, `WordLevelExtractionStrategy` receives these as `whiteMaskAreas` rectangles. Per-character bounding boxes are tested against all mask rects (`Intersects()`); matching characters flush the current word buffer and are skipped entirely. The final `SaveCurrentWord()` also checks the accumulated word rect before adding to `Words`. This is how selective text masking is implemented without modifying the underlying PDF content stream.

### `Ignore` annotation type
PDF annotations with subject "Ignore" loaded via `GetAnnotations(file1Path, "Ignore")` and merged into `lstIgnoreAnnotationRegions`. Any text in these regions is excluded from comparison. This is the mechanism for suppressing known acceptable differences.

### `WordLevelExtractionStrategy` — iText7 word extraction details
**File:** `PDFProcessor.cs`, `WordLevelExtractionStrategy` class

iText7 `IEventListener` that processes `RENDER_TEXT` events per character:
- **Word boundaries**: `isNewLine` when yDistance > fontSize×0.5; `isNewWord` when xSpacing > fontSize×0.3
- **White mask**: per-character bbox checked against `whiteMaskAreas`; hit → flush + skip; `SaveCurrentWord()` double-checks final word rect before adding
- **Red-bracket mode** (`ConsiderRedTextInBracketAsOneWord=true`): text colored R>200/G<80/B<80 starting with `( [ <` triggers accumulation until matching `)  ] >` — emitted as single `WordInfo` with `IsRedTextInBracketWord=true`; masked red-bracket words are discarded after close bracket
- **`GetSupportedEvents()`**: returns `null` = listens to all event types (filtered internally to `RENDER_TEXT` only)
- **Bounding box**: computed as union of `info.GetAscentLine().GetBoundingRectangle()` and `info.GetDescentLine().GetBoundingRectangle()` per character

### `NormSearch()` text normalization (PATH 3 matching)
**File:** `WordDriftCorrector.cs`

Phrase text from JSON is normalized before matching document paragraph text. Minimum 5 characters (`TEXT_MATCH_MIN = 5`). PATH 3 matches the FIRST paragraph in the document body (not cell, not header/footer) that contains the normalized key — this means if the same phrase appears in multiple places, only the first match is corrected.

**False substring match guard:** If the paragraph text is more than 4× the key length AND the paragraph does NOT start with the key, the match is suppressed. Without this, a short phrase like "Prior Policy" (12 chars) intended to match a table header cell (in a cell, skipped by Pass 1) would instead match a long body paragraph that merely contains "Prior Policy" as a substring (e.g. "...the Prior Policy are more favorable..."), applying the correction to the wrong paragraph. The same guard applies in `ApplyHeaderFooterCorrection`.

### MAX_PARAGRAPH_X_DRIFT_PT = 10pt cap (PATH 3 only)
**File:** `WordDriftCorrector.cs`, `ApplyParagraphCorrection()`

X drift > 10pt is skipped for body paragraph corrections. This prevents large accidental shifts when phrase matching is ambiguous. No equivalent cap for cell paragraph fallback or table corrections.

### `TagInlineRuns` — INLINE_RUN classification
**File:** `PDFProcessor.cs`, `TagInlineRuns()`

Groups phrase groups by (page, Y within 3pt). Lines with 2+ phrase groups having **distinct** X-deltas (rounded to 1dp) are inline runs (separate tab segments on one line). All non-TABLE_CELL phrase groups on such lines are marked `LayoutContext = "INLINE_RUN"`. INLINE_RUN groups are cleaned early in Gate 2 (before XI detection) because they have already been purged from any table group by `AugmentTableGroupsFromDocxContext` and cannot be XI candidates.

### `GenerateJsonDiffReport` — 4-gate IsCleanDiff ordering rationale
**File:** `PDFProcessor.cs`, ~line 5230

Gate 2 (TABLE_CELL) is deliberately applied in **Phase 3** (after XI detection) rather than inline with Gates 1/3/4. Reason: if TABLE_CELL groups were marked `IsCleanDiff=true` before the XI pass, the XI check's `!pg.IsCleanDiff` filter would exclude them, meaning `xDriftPGs` would have no table phrase groups, XI would never fire, and Pass-2 sibling propagation in `GenerateHtmlReport` would find nothing to propagate. The deferred Phase 3 ensures XI sees uncorrected-table phrase groups in both original and modified comparisons.

### `CompareFolders` — subfolder results appended AFTER `GenerateSummaryReport`
**File:** `PDFProcessor.cs`, `CompareFolders()`

The root-level summary report is generated from only root-level files. Subfolder results are appended to `fileResults` AFTER the report is written. The returned `List<FileComparisonResult>` is flat (all depths), but the report only shows root-level files plus subfolder links. This is intentional to preserve visual hierarchy in HTML.

### Orphaned PDF cleanup in `ConvertWordFolderToPDF`
**File:** `DocumentService.cs`

Phase 2 deletes PDFs in `destRoot` that have no corresponding source `.doc/.docx`. JSON state cleanup (Phase 3) is scoped: only keys starting with `sourceRootWithSlash` are removed. This prevents a single folder's state update from corrupting entries for other source roots stored in the same JSON state file.

### `priorDiffRects` key is the FULL baseline path (not filename)
**File:** `PDFProcessor.cs`, `CompareFolders()` and `Program.cs`

The dictionary key is the absolute path of the V14 PDF (`r.PdfPath`). This avoids collisions when identically-named files exist in different subfolders. Both original and modified comparisons use the same V14 baseline folder, so `file1Path` is identical across both runs.

### `CELL_PADDING` safety gate: negative table indent
**File:** `WordDriftCorrector.cs`, `ApplyTableCorrection()` and `ApplyCellPadding()`

If the table has a negative `LeftIndent` and the strategy is `CELL_PADDING`, the strategy is silently overridden to `TABLE_INDENT`. Reason: negative-indent tables are already partially into the left margin; applying cell padding instead of moving the table would push content into the clipped zone.

---

## 7. STATE & DATA MANAGEMENT

**In-memory state (per run)**
- `PdfComparer.HighlightedRegions` (HashSet) — tracks which diff regions were already highlighted to avoid duplication in HTML report
- `correctedTableListIndices`, `correctedDocxTables`, `handledPhraseGroupIds` — correction-run guard sets in `WordDriftCorrector.ApplyCorrectionsToDocument()`
- `appliedXPerTableListIdx` — tracks drift applied per table to enable nested-table ancestor correction

**Persistent state**
- `*_diff.json` files in `JSONDiff/` subfolders — main data handoff between PdfComparer and WordDriftCorrector
- `fileConversionState.json` (optional, path passed to `ConvertWordFolderToPDF`) — incremental conversion cache; keyed by absolute source path; stores `Size` and `LastModified` (UTC)

**No database, no sessions, no external cache** — all state is file-based.

---

## 8. EXTERNAL INTEGRATIONS

**Aspose.Words (licensed)**
- Module owner: `DocumentService.cs` (license + encoding setup), `WordDriftCorrector.cs` (document manipulation)
- License file path is hardcoded in `Program.cs`: `C:\Personal\Project\DriftCorrector\Files\Key\Aspose.Words.NET.lic`
- Critical quirk: Setting `CellFormat.Width` on a gridSpan cell removes gridSpan and corrupts `tblGrid` (BUG #8 workaround)

**iText7 (`iText.Kernel`, `iText.Forms`)**
- Module owner: `PDFProcessor.cs`
- Used for: text extraction via `PdfCanvasProcessor + WordLevelExtractionStrategy`, form field extraction, annotation reading (`GetAnnotations`), white-rectangle overlays (`RemoveText`), annotation writing (`WhiteMasked` / `Ignore`)

**Docnet.Core**
- Used in `PDFProcessor.cs` → `RenderPdfPageToBitmap()`: opens page via `IDocLib.LoadDocument()`, renders to BGRA byte array at a computed DPI scale, then converts to `System.Drawing.Bitmap` via `BitmapData.Scan0` copy. Used to produce page images for HTML diff reports.

**robocopy.exe (Windows system)**
- Called via `DocumentService.ExecuteRobocopy()` with `/E /Z /R:5 /W:5 /MT:16 /NP` flags
- Exit codes 0–7 = success; 8+ = failure
- Trailing slashes are stripped from paths before passing to avoid escaping the closing quote

---

## 9. DEAD CODE & LANDMINES

**`DocumentService.ApplyShifts()`** — Legacy XML-based correction that adjusts shapes, paragraphs, and table cells based on `<textShift deltaX="" deltaY="">` XML nodes. This is the old correction approach, now fully superseded by `WordDriftCorrector`. Do not use.

**`DocumentService.ConvertWordToPdf_V23()` (commented out)** — Replaced by `ConvertDocToPdf()` + `ConvertWordFolderToPDF()`. The commented code is still present as reference.

**`ConvertFolderWordToPdf_V23()`** — Non-recursive flat-folder converter. Superseded by `ConvertWordFolderToPDF()` which handles recursive folder mirroring and incremental state. Use `ConvertWordFolderToPDF()` instead.

**`ComparePDF_Raw()` in Program.cs** — Variation that doesn't pass a word template directory to `CompareFolders`. No docx context → no XI detection → no table-level correction grouping. Produces less precise results. Currently commented out.

**`CompareOnly()` in Program.cs** — Ad-hoc triple comparison (V23P vs V23A, V14 vs V23P, V14 vs V23A). No correction loop. Currently commented out.

**`ComparePDFCore()` in Program.cs** — Empty stub with only a console write. Dead code.

**`DeriveStrategy()` — residual `tableInd == 0` branch**
The condition at line ~780: `if (absTableInd > 0 || tableInd == 0)` — this second OR clause `tableInd == 0` is always true (any non-negative table will match), making the preceding `CellPadding` check the only real gate. This is the original BUG #6 condition and is now the intentional fallback for the non-zero-indent case after the zero-indent fix was applied via the CellParagraphIndent strategy. Be careful when modifying this logic.

**`GetWin32LongPath()` / `RemoveLongPathPrefix()`** — Utility methods in `DocumentService.cs` for `\\?\` long-path prefix handling. Neither is called from other visible code in the analyzed files; may be legacy or used in a partial class extension not analyzed here.

**`DiagnoseAsposeVersion()`** — Diagnostic tool that writes Aspose assembly metadata to a file. Not called in the normal pipeline. Use only for debugging version mismatches.

---

## 10. QUICK ORIENTATION FOR BUG RESOLUTION

- **If PDF comparison produces no diffs when there should be some** → check `lstIgnoreAnnotationRegions` is not consuming the expected regions; verify `Ignore` annotations are not present in the baseline PDF; check `WhiteMasked` annotations aren't covering the text
- **If corrections are applied but drift gets WORSE** → check SpaceBefore direction (BUG #9 pattern); verify `IsModifiedComparison` flag is set correctly on the second comparer instance
- **If a table is shifted too far left after correction** → check if it's a zero-indent layout table (BUG #6 pattern); verify `DeriveStrategy()` is not returning `TableIndent` for a zero-indent table; should be `CellParagraphIndent`
- **If table text wraps unexpectedly after column-width correction** → check for gridSpan cells (BUG #8); look for width variance across rows in the target column; verify both Check 1 and Check 2 in `ApplyColumnWidths()`
- **If nested table is double-corrected** → check `BuildTableAncestryMap()` produced the correct ancestry; verify `GetAncestorAppliedXPt()` is accumulating parent drift correctly
- **If a table is skipped with "contains XI phrase group(s)"** → the XI detection in `GenerateJsonDiffReport` classified the phrase group as already-at-correct-position; check `ComputePivotX()` is computing the right reference X; may need to disable XI for that specific table group
- **If PATH 3 can't find a paragraph to correct** → phrase is likely inside a table cell; check two-pass fallback logic in `ApplyParagraphCorrection()`; verify `TEXT_MATCH_MIN` (5 chars) isn't filtering the phrase
- **If Aspose license error on startup** → `SetLicense()` must be called before any `Document()` constructor; check file path is correct; check `EnsureEncodingProviderRegistered()` was called first
- **If converted PDFs are stale (not re-converted after Word changes)** → check `ConvertWordFolderToPDF` state JSON file; clear or delete the state file to force full re-conversion
- **If orphaned PDFs remain in modified folder** → Phase 2 cleanup in `ConvertWordFolderToPDF` deletes orphans; if it isn't running, verify `destRoot` exists at the time of the cleanup scan
- **If `CompareFolders` report shows wrong stats for modified comparison** → verify `priorDiffRects` was built from `ComparisonDetails` of the FIRST (original) comparison run, and keyed by full V14 path (not filename)
- **If subfolder results are missing from the returned list** → they are appended to `fileResults` AFTER `GenerateSummaryReport`; if using the return value to build `priorDiffRects`, this is correct behavior
- **If robocopy fails silently** → `ExecuteRobocopy` reads both stdout and stderr asynchronously; check exit code — codes 0–7 are success; code 8+ is failure; check that paths don't have trailing backslashes
- **If reflow word appears as red diff when it should be green** → check `REFLOW_MIN_LEFTWARD_DRIFT_PT=12pt`, `REFLOW_MAX_PUBLISH_FROM_MARGIN=36pt`, `REFLOW_MIN_BASELINE_OFFSET=15pt`; all 3 thresholds must be met; the word must be added to `cleanWordPositions` in GJDR for `IsPhraseGroupClean` to work
- **If INLINE_RUN groups are not being suppressed** → verify `TagInlineRuns` runs after `BuildPhraseGroups` and before `AugmentTableGroupsFromDocxContext`; check that multiple phrase groups on the same line actually have distinct (rounded to 1dp) X-deltas
- **If IsCleanDiff is not propagating to TABLE_CELL groups** → Phase 3 gate is intentionally deferred after XI detection; check `pg.IsXIgnore` is not preventing Phase 3 from running; verify the groups have `TableGroupId.HasValue = true`

---

## FILES READ STATUS

- **`PDFProcessor.cs`** — **FULLY READ** (314KB, ~6262 lines). All regions documented above.
- **`WordDriftCorrector.cs`** — **MOSTLY READ** (~1300 of ~1900 lines). Unread utility sections:
  - `FindTemplate()` — locates .docx/.doc by base name in template folder
  - `LoadDiffReport()` — JSON deserialization entry point
  - `PrintConsoleSummary()` — terminal output after correction run
  - `FindJsonStrategy()` — extracts strategy string from JSON for a given table
  - `IsParaType()`, `NormSearch()`, `GetParaText()`, `Trunc()` — utility helpers
  - `GenerateCorrectionReport()` — HTML correction report generator
- **`DocumentService.cs`** — **FULLY READ** (762 lines)
- **`Logger.cs`** — **FULLY READ** (47 lines)
- **`Program.cs`** — **FULLY READ** (239 lines)
