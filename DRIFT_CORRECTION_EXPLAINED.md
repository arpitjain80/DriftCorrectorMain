# DriftCorrector — Complete Guide in Plain English
### For Management & Non-Technical Stakeholders

---

## Table of Contents

1. [The Core Problem — What is "Drift"?](#1-the-core-problem--what-is-drift)
2. [How We Detect Drift — The Comparison Engine](#2-how-we-detect-drift--the-comparison-engine)
3. [How We Fix Drift — The 4 Correction Paths](#3-how-we-fix-drift--the-4-correction-paths)
4. [Correction Strategy 1 — TABLE INDENT (Slide the Whole Table)](#4-correction-strategy-1--table-indent-slide-the-whole-table)
5. [Correction Strategy 2 — CELL PADDING (Shrink the Cushion)](#5-correction-strategy-2--cell-padding-shrink-the-cushion)
6. [Correction Strategy 3 — COLUMN WIDTHS (Resize Individual Columns)](#6-correction-strategy-3--column-widths-resize-individual-columns)
7. [Correction Strategy 4 — CELL PARAGRAPH INDENT (Fix Text Inside Zero-Border Tables)](#7-correction-strategy-4--cell-paragraph-indent-fix-text-inside-zero-border-tables)
8. [Correction Strategy 5 — FLOATING TABLE SHIFT (Move Absolutely-Positioned Tables)](#8-correction-strategy-5--floating-table-shift-move-absolutely-positioned-tables)
9. [Correction Strategy 6 — ROW HEIGHT (Fix Vertical Drift in Fixed-Height Rows)](#9-correction-strategy-6--row-height-fix-vertical-drift-in-fixed-height-rows)
10. [Correction Strategy 7 — SPACE BEFORE (Fix Vertical Gaps in Auto-Height Rows)](#10-correction-strategy-7--space-before-fix-vertical-gaps-in-auto-height-rows)
11. [Correction Strategy 8 — LIST SPACING SWEEP (Fix Cascading Bullet-List Drift)](#11-correction-strategy-8--list-spacing-sweep-fix-cascading-bullet-list-drift)
12. [Correction Strategy 9 — PARAGRAPH INDENT (Fix Standalone Body Text)](#12-correction-strategy-9--paragraph-indent-fix-standalone-body-text)
13. [Correction Strategy 10 — HEADER & FOOTER CORRECTION](#13-correction-strategy-10--header--footer-correction)
14. [The Smart Guard — XI / XIgnore (Don't Fix What Isn't Broken)](#14-the-smart-guard--xi--xignore-dont-fix-what-isnt-broken)
15. [The Smart Guard — Nested Table Ancestry (Don't Double-Fix)](#15-the-smart-guard--nested-table-ancestry-dont-double-fix)
16. [Bugs Found & Fixed — A Summary](#16-bugs-found--fixed--a-summary)
17. [The Full Pipeline at a Glance](#17-the-full-pipeline-at-a-glance)

---

## 1. The Core Problem — What is "Drift"?

### The Analogy

Imagine you have a perfectly formatted insurance policy document in **Microsoft Word**. You print it to a PDF using a printer from 2014 — let's call that the "reference print" **(V14)**. A few years later, the printing software is upgraded to a 2023 version **(V23)**.

When you print the exact same Word file with the new software, **the text and tables on the page shift slightly** — some a few millimetres to the right, some downward. The document *looks* almost identical to the human eye, but when you overlay the two PDFs and measure precisely, words and table borders are in slightly different positions.

This shift is called **"Drift."**

```
  V14 PDF (Reference - Correct)        V23 PDF (Drifted - Problem)
  ┌──────────────────────────┐          ┌──────────────────────────┐
  │                          │          │                          │
  │  ┌──────────────────┐    │          │      ┌──────────────────┐│
  │  │  Coverage Table  │    │          │      │  Coverage Table  ││
  │  │  $1,000,000      │    │          │      │  $1,000,000      ││
  │  └──────────────────┘    │          │      └──────────────────┘│
  │                          │          │                          │
  │  Premium: $500/year      │          │      Premium: $500/year  │
  └──────────────────────────┘          └──────────────────────────┘
        Table at X=36pt                       Table at X=72pt  ← DRIFTED 36pt RIGHT
        Text at X=36pt                        Text at X=72pt   ← DRIFTED 36pt RIGHT
```

### Why Does Drift Happen?

Microsoft Word documents use a complex layout engine. When Aspose (the software library we use to convert Word → PDF) was upgraded from version 14 to version 23, some small internal rendering decisions changed:

- How much **padding inside table cells** is interpreted
- How **table position** is calculated relative to the page margin
- How **paragraph spacing** (blank space above a paragraph) accumulates

None of these are "bugs" in the traditional sense — they are legitimate engine improvements — but they cause visual displacement of content between the two versions.

### What the DriftCorrector Does

DriftCorrector is an automated system that:
1. **Compares** the V14 and V23 PDFs word-by-word, position-by-position
2. **Measures** exactly how far each piece of text has drifted (in typographic points, where 1pt ≈ 0.35mm)
3. **Edits** the original Word (.docx) file to compensate — making V23 produce output that matches V14
4. **Verifies** the correction by re-converting and comparing again

---

## 2. How We Detect Drift — The Comparison Engine

### Like a Document Diff Tool, but for Positions

The comparison engine reads every word from both the V14 and V23 PDFs — including its exact X (horizontal) and Y (vertical) position on the page. It then matches words between the two documents and calculates the delta (difference).

```
  Word "Coverage"
  ─────────────────────────────────────────────────────
  V14 Position:  X = 36.0pt,  Y = 412.5pt
  V23 Position:  X = 72.4pt,  Y = 412.5pt
                 ──────────────────────────
  DRIFT:         ΔX = +36.4pt (shifted RIGHT by 36pt)
                 ΔY =   0.0pt (vertical position unchanged)
```

### Grouping Words Into "Phrase Groups"

Individual word drifts are grouped into **Phrase Groups** — clusters of words on the same line that drift by the same amount. This allows the system to understand "this entire sentence shifted right by 36pt" rather than treating each word separately.

```
  "The insured shall pay"  ← all words drift +36pt right → ONE phrase group
  "a premium of $500"      ← all words drift +36pt right → ONE phrase group

  vs.

  "Policy Number: 12345"   ← first 2 words drift +5pt right (column 1)
  "              ABCDE"    ← last word stays in place (column 2)
  → TWO separate phrase groups (non-uniform) → COLUMN WIDTHS strategy
```

### Table Detection

The system automatically detects which words belong to tables by analyzing the grid-like patterns of X and Y positions:

```
  Words clustered into rows (Y ± 6pt) and columns (X ± 8pt):

  ┌─────────────┬───────────────┬─────────────┐
  │ Row 0, Col 0│ Row 0, Col 1  │ Row 0, Col 2│  ← Y ≈ 412pt
  ├─────────────┼───────────────┼─────────────┤
  │ Row 1, Col 0│ Row 1, Col 1  │ Row 1, Col 2│  ← Y ≈ 425pt
  └─────────────┴───────────────┴─────────────┘
    X ≈ 36pt      X ≈ 180pt       X ≈ 300pt
```

---

## 3. How We Fix Drift — The 4 Correction Paths

The corrector uses 4 paths, applied in order from most-precise to least-precise:

```
  ┌─────────────────────────────────────────────────────────────────┐
  │                    CORRECTION PATHS (in order)                  │
  ├─────────────────────────────────────────────────────────────────┤
  │ PATH 1: Word DOM table corrections (most precise)               │
  │         Uses the Word document structure itself to identify     │
  │         which table and cell the drifted text belongs to.       │
  ├─────────────────────────────────────────────────────────────────┤
  │ PATH 2: PDF heuristic table corrections (fallback)              │
  │         Uses the grid patterns detected in the PDF when the     │
  │         Word DOM path didn't yield a clear match.               │
  ├─────────────────────────────────────────────────────────────────┤
  │ PATH 3: Standalone paragraph corrections                        │
  │         For text outside tables — body text, numbered           │
  │         paragraphs, standalone sentences.                       │
  ├─────────────────────────────────────────────────────────────────┤
  │ PATH 4: Header & Footer corrections                             │
  │         For text in the top/bottom repeating areas of pages.    │
  └─────────────────────────────────────────────────────────────────┘

  RULE: Once a phrase group is handled by an earlier path,
        later paths skip it. No double-corrections.
```

Each path selects the best **Strategy** for how to actually apply the fix in the Word file. The following sections explain each strategy.

---

## 4. Correction Strategy 1 — TABLE INDENT (Slide the Whole Table)

### What Is It?

The simplest fix: slide the entire table left to undo the rightward drift. This is like pushing a picture frame back to its correct position on a wall.

### When Is It Used?

When **all columns in the table drift by the same amount**. This is the "uniform shift" — the whole table moved as one block.

```
  BEFORE (V23 - drifted):
  ┌──────────────────────────────────────────┐ Page
  │                                          │
  │         ┌───────────────────────────┐    │
  │         │ Policy  │  $500  │  Annual│    │  ← Table starts at X=72pt
  │         └───────────────────────────┘    │
  │                                          │
  └──────────────────────────────────────────┘

  AFTER (Fixed - matches V14):
  ┌──────────────────────────────────────────┐ Page
  │                                          │
  │    ┌───────────────────────────┐         │
  │    │ Policy  │  $500  │  Annual│         │  ← Table moved back to X=36pt
  │    └───────────────────────────┘         │
  │                                          │
  └──────────────────────────────────────────┘

  Correction: Table.LeftIndent  72pt → 36pt  (shifted left by 36pt)
```

### Real-World Example

An insurance policy has a "Premium Summary" table. In V14 it sits at the left margin. In V23 it drifts 36pt to the right. The TABLE_INDENT strategy simply reduces the table's `LeftIndent` property by 36pt, sliding it back to the left.

### Safety Check

The system ensures the table never gets pushed into the **physical left margin** of the page — it clamps the indent to a safe minimum. It also handles pre-existing negative indents (tables that intentionally start before the normal text margin) carefully.

---

## 5. Correction Strategy 2 — CELL PADDING (Shrink the Cushion)

### What Is It?

Every table cell has an internal "cushion" — empty space between the cell wall and the text inside. This is called **left padding**. If V23 applies more padding than V14, the text inside the cell shifts right even though the table itself hasn't moved.

Think of it like a picture frame that has gotten thicker matting — the picture inside is pushed inward even though the outer frame is in the same position.

### When Is It Used?

When the **table's position hasn't changed** but the **text inside the cells** has shifted right, and the measured drift matches the current cell padding value.

```
  V14 Cell (correct):                V23 Cell (drifted):
  ┌─────────────────────┐            ┌─────────────────────┐
  │←3pt→ "Coverage"     │            │←──9pt──→ "Coverage" │
  └─────────────────────┘            └─────────────────────┘
    padding=3pt                        padding=9pt  ← too much!
    text starts at X=39pt              text starts at X=45pt (drift=+6pt)

  Fix: Reduce all cell left padding from 9pt to 3pt
       → text moves back to X=39pt ✓
```

### Real-World Example

An employment practices liability form has a table where each cell's internal padding was increased by 6pt in V23. The corrector reduces all cells' left padding by 6pt, pulling the text back to its correct position without touching the table frame at all.

### Safety Check

If the current cell padding is **less than** the drift amount, the system cannot reduce it below zero (that would be nonsensical). In that case, it caps the correction and logs a warning.

Also, if the table itself has a **negative left indent** (it already extends to the left of the normal page margin), cell padding is NOT the right fix — the system automatically switches to TABLE_INDENT instead.

---

## 6. Correction Strategy 3 — COLUMN WIDTHS (Resize Individual Columns)

### What Is It?

When different **columns drift by different amounts**, it means the table columns have changed width between V14 and V23. The fix is to adjust individual column widths to absorb the difference.

### When Is It Used?

When columns shift by **non-uniform amounts** — column 1 might drift 5pt right, but column 2 drifts 10pt right. This tells us column 1 has gotten wider.

```
  V14 (correct):                     V23 (drifted):
  ┌──────────┬──────────┬──────────┐  ┌─────────────┬──────────┬──────────┐
  │ Coverage │ Limit    │ Premium  │  │ Coverage    │ Limit    │ Premium  │
  │ Type     │          │          │  │ Type        │          │          │
  ├──────────┼──────────┼──────────┤  ├─────────────┼──────────┼──────────┤
  │ Bodily   │ $500K    │ $1,200   │  │ Bodily      │ $500K    │ $1,200   │
  │ Injury   │          │          │  │ Injury      │          │          │
  └──────────┴──────────┴──────────┘  └─────────────┴──────────┴──────────┘
    Col 1: X=36pt                       Col 1: X=36pt  (no drift)
    Col 2: X=144pt                      Col 2: X=180pt (drift=+36pt → col 1 got wider!)
    Col 3: X=252pt                      Col 3: X=288pt (drift=+36pt → same, col 2 unchanged)

  Fix: Shrink Column 1 width by 36pt
       → Column 2 and 3 slide back left ✓
```

### How the Math Works

The trick here is: each column's shift equals the sum of width changes in all columns to its left. By calculating the *difference* in drift between adjacent columns, we can determine exactly how much each column's width changed.

```
  Column drift:
    Col 0: 0pt (no drift)
    Col 1: +36pt (drifted right)
    Col 2: +36pt (same drift as col 1)

  Width change = drift[col N] - drift[col N-1]:
    Col 0 width: unchanged
    Col 1 width: drifted 36pt - 0pt = needs to shrink by 36pt
    Col 2 width: drifted 36pt - 36pt = unchanged
```

### Critical Safety Check — Merged Cells (BUG #8)

> ⚠️ **This is one of the most important safeguards in the system.**

Some tables use **merged cells** — where two adjacent cells are combined into one wide cell. Word tracks this internally as "gridSpan."

```
  Table with merged header (common in forms):
  ┌──────────────────────────────────────┐
  │         COVERAGE SUMMARY             │  ← ONE cell spanning all 3 grid columns
  ├───────────────┬──────────┬───────────┤
  │ Coverage Type │ Limit    │ Premium   │  ← THREE normal cells
  ├───────────────┼──────────┼───────────┤
  │ Bodily Injury │ $500K    │ $1,200    │
  └───────────────┴──────────┴───────────┘

  Row 0: 1 cell (spans 3 columns)
  Row 1: 3 cells
  Row 2: 3 cells
```

**The danger:** If you try to resize a merged cell by setting its Width property in the Aspose library, it **removes the merge** and corrupts the entire table. Text that was spread across the header row collapses, wraps unexpectedly, and the whole table layout breaks.

**The fix (two independent checks):**

- **Check 1 (Width Variance):** Before resizing any column, measure that column's cell width across ALL rows. If some rows have a different width than others (because a merged cell makes the "cell" appear bigger), merged cells are present → SKIP width changes, only adjust the table's left indent.

- **Check 2 (Row Cell Count):** If any row has fewer visible cells than expected (because some cells are merged), same conclusion → SKIP width changes.

---

## 7. Correction Strategy 4 — CELL PARAGRAPH INDENT (Fix Text Inside Zero-Border Tables)

### What Is It?

This strategy was created to fix a specific class of document layout — **"layout tables."**

### What Are Layout Tables?

Many document authors use Word tables as an invisible layout tool. The table has no visible borders and sits right at the page margin (left indent = 0). The actual positioning of text is done through the **paragraph indent inside each cell**, not through the table's own position.

A classic example: hanging-indent numbered paragraphs:

```
  APPEARS IN THE DOCUMENT AS:

  A.   The insured shall pay
       a premium of $500 per year.

  B.   Coverage applies to
       all listed vehicles.

  BUT INTERNALLY IT IS A TABLE:
  ┌──────────────────────────────────────────────┐
  │┌────┐┌─────────────────────────────────────┐ │
  ││ A. ││  The insured shall pay a premium... │ │  ← Row 1
  │└────┘└─────────────────────────────────────┘ │
  │┌────┐┌─────────────────────────────────────┐ │
  ││ B. ││  Coverage applies to all listed...  │ │  ← Row 2
  │└────┘└─────────────────────────────────────┘ │
  └──────────────────────────────────────────────┘
  Table LeftIndent = 0pt (sits at page margin)
  Cell paragraph LeftIndent = 36pt (creates the "A." hanging effect)
```

### The Original Bug (BUG #6)

The old system saw a drift on a zero-indent table and said: *"The drift is +36pt, and this table has LeftIndent=0. I'll fix it by setting LeftIndent = 0 - 36 = -36pt."*

```
  What happened BEFORE the fix:

  ┌──────────────────────────────────────────────┐ Page boundary
  │                                              │
  │────────────────────────────────────── Margin │
  │◄─ 36pt ─►                                    │
  │           ┌──────────────────────────────┐   │
  │           │ A.  The insured shall pay... │   │  ← Table at its correct position
  │           └──────────────────────────────┘   │
  └──────────────────────────────────────────────┘

  "Fix" applied: Table.LeftIndent = -36pt

  ┌──────────────────────────────────────────────┐ Page boundary
  │                                              │
  │────────────────────────────────────── Margin │
  │◄ -36pt►                                      │
  │ ┌──────────────────────────────┐             │
  │ │. The insured shall pay...    │  ← TABLE PUSHED INTO LEFT MARGIN! CLIPPED!
  │ └──────────────────────────────┘
  └──────────────────────────────────────────────┘

  The "A" and the first letter of text is now hidden behind the page margin. WORSE!
```

### The Fix

The new strategy recognizes that this table type should NOT be moved. Instead, the **paragraph indent inside each cell** is what needs adjusting — because that's what controls where the text visually appears.

```
  CORRECT FIX — CELL_PARAGRAPH_INDENT:

  Before: cell paragraph LeftIndent = 36pt (causing +36pt visual drift in V23)
  After:  cell paragraph LeftIndent = 0pt  (text now aligns with V14 reference)

  Table itself: untouched, still at LeftIndent = 0pt ✓
  Cell padding: untouched ✓
  Only the paragraph inside the cell is adjusted ✓
```

**The mathematical proof:**
> Total text position = Page Margin + Table Indent + Cell Padding + **Paragraph Indent**
>
> If the table, padding, and margin haven't changed, then the drift came from the Paragraph Indent.
> Therefore: `new Paragraph Indent = current Paragraph Indent - drift`

---

## 8. Correction Strategy 5 — FLOATING TABLE SHIFT (Move Absolutely-Positioned Tables)

### What Is It?

Most Word tables are "inline" — they flow with the text, positioned relative to the left margin. But some tables are **floating** — they have a fixed, absolute position on the page (like a text box). These are specified with a special XML tag called `tblpPr` in the Word file.

### The Problem

For floating tables, setting `LeftIndent` (which works fine for normal tables) is **silently ignored** by both Word and Aspose V23. The table stays exactly where its absolute coordinate says it should be, no matter what LeftIndent you set.

```
  FLOATING TABLE in the Word XML:
  <w:tblPr>
    <w:tblpPr w:horzAnchor="margin" w:tblpX="277" w:tblpY="600"/>
    <!-- ↑ This means: "place me 277 twips from the left margin,
                         600 twips from the top margin" -->
  </w:tblPr>

  When V23 renders this, it places the table at X=5.65pt MORE to the right than V14.
  Setting LeftIndent = -5.65pt has NO EFFECT because tblpX overrides it.
```

### The Fix — TABLE_FLOAT_SHIFT

Instead of adjusting `LeftIndent`, the system adjusts `AbsoluteHorizontalDistance` — which directly maps to the `tblpX` value in the XML.

```
  BEFORE: AbsoluteHorizontalDistance = 19.27pt (tblpX in XML)

  V14 renders table at:  X = 19.27pt ✓
  V23 renders table at:  X = 24.92pt (drift = +5.65pt)

  FIX: AbsoluteHorizontalDistance = 19.27 - 5.65 = 13.62pt

  AFTER: V23 now renders at X ≈ 19.27pt ✓ (matches V14)
```

### Real-World Example (DO_ERP_Header)

A form has a header area with a floating table containing the company logo and policy number. In V23 the entire header table drifted 5.65pt to the right. Previous fix attempts using `LeftIndent` had zero effect because it was a floating table. The TABLE_FLOAT_SHIFT strategy correctly modifies the absolute anchor position.

---

## 9. Correction Strategy 6 — ROW HEIGHT (Fix Vertical Drift in Fixed-Height Rows)

### What Is It?

While most of our drift corrections are horizontal (left-right), vertical (up-down) drift also occurs. If a table row has an explicitly set height (e.g., "this row must be exactly 20pt tall"), and V23 renders content that overflows that height slightly, the rows below shift downward.

### The Fix

For rows with an explicit height set, the system adjusts the `Row Height` property.

```
  V14 (correct):                     V23 (drifted):
  ┌───────────────────────┐           ┌───────────────────────┐
  │ Row 1: "Insured Name" │ 20pt tall │ Row 1: "Insured Name" │ 20pt tall  (ok)
  ├───────────────────────┤           ├───────────────────────┤
  │ Row 2: "Policy #"     │ 18pt tall │ Row 2: "Policy #"     │ 20pt tall  ← too tall!
  ├───────────────────────┤           ├───────────────────────┤
  │ Row 3: "Premium"      │ 18pt tall │ Row 3: "Premium"      │ 20pt tall  ← pushed down!
  └───────────────────────┘           └───────────────────────┘

  Row 3 appears 2pt lower in V23 than V14.
  Fix: Set Row 2 height = 18pt → Row 3 slides back up 2pt ✓
```

---

## 10. Correction Strategy 7 — SPACE BEFORE (Fix Vertical Gaps in Auto-Height Rows)

### What Is It?

When a table row does NOT have a fixed height (it automatically expands to fit content), the vertical drift is caused by something different: **SpaceBefore** — an invisible blank gap that Word can add above a paragraph.

If V23 applies more SpaceBefore than V14, the paragraph appears lower on the page.

### The Direction Fix (BUG #9)

> ⚠️ **This was a critical bug that made drift WORSE before it was fixed.**

Understanding the coordinate system is crucial here:

```
  PDF Y-coordinate system:

  ┌──────────────────────────┐
  │  Y = 800pt (top of page) │
  │                          │
  │  Y = 412pt  "Policy #"   │  ← V14 position (correct)
  │  Y = 400pt  "Policy #"   │  ← V23 position (too LOW, drift = -12pt)
  │                          │
  │  Y = 0pt   (bottom)      │
  └──────────────────────────┘

  In PDFs, Y increases UPWARD (opposite to screen coordinates).

  drift = V23.Y - V14.Y = 400 - 412 = -12pt (V23 is lower than V14)
```

**SpaceBefore pushes content DOWN (decreases Y).** So if V23 has too much SpaceBefore, content appears lower (smaller Y).

To fix: we need to **REDUCE** SpaceBefore. The correct formula:
```
  new SpaceBefore = current SpaceBefore + drift
                  = 12pt + (-12pt) = 0pt   ✓  (drift < 0, so this reduces spacing)
```

The old buggy formula was `current SpaceBefore - drift`:
```
  (WRONG) new SpaceBefore = 12pt - (-12pt) = 24pt  ✗  (made drift TWICE AS BAD!)
```

### Safety Guard

The system **only reduces existing SpaceBefore** — it never adds SpaceBefore where there was none. Adding SpaceBefore from zero would create visible blank gaps that weren't in the original document.

---

## 11. Correction Strategy 8 — LIST SPACING SWEEP (Fix Cascading Bullet-List Drift)

### What Is It?

This is a special fix for a very specific pattern: **bullet-point or numbered lists inside table cells**.

### The Cascading Drift Problem

In V14, SpaceBefore was applied only to the **first list item** in a cell. In V23, SpaceBefore is applied to **every list item**. This creates an accumulating problem:

```
  INSIDE A TABLE CELL — BULLET LIST:

  V14 (correct):                     V23 (drifted):
  Y=400  ● Item 1 (SpaceBefore=6pt)  Y=400  ● Item 1 (SpaceBefore=6pt) ← same
  Y=386  ● Item 2 (SpaceBefore=0pt)  Y=380  ● Item 2 (SpaceBefore=6pt) ← 6pt gap added!
  Y=372  ● Item 3 (SpaceBefore=0pt)  Y=354  ● Item 3 (SpaceBefore=6pt) ← 12pt total!
  Y=358  ● Item 4 (SpaceBefore=0pt)  Y=328  ● Item 4 (SpaceBefore=6pt) ← 18pt total!
         ↑ Uniform spacing                   ↑ Growing gaps - cascading drift
```

The drift ACCUMULATES. By the 4th item, the content is 18pt lower than it should be. By the 10th item, it could be 60pt off — nearly an inch!

### What the System Detects

The system notices:
- The last row's drift is much larger than one SpaceBefore unit (e.g., 18pt drift vs 6pt SpaceBefore)
- This means multiple rows must have SpaceBefore that needs to be zeroed

### The LIST_SPACING_SWEEP Fix

Instead of only fixing the one row that was detected as drifted, the system **sweeps through ALL rows** in the table and zeros the SpaceBefore on every list-paragraph row (except the very first one, which both V14 and V23 agree on).

```
  AFTER SWEEP:

  V14 (reference):                   V23 (fixed):
  Y=400  ● Item 1 (SpaceBefore=6pt)  Y=400  ● Item 1 (SpaceBefore=6pt) ← preserved
  Y=386  ● Item 2 (SpaceBefore=0pt)  Y=386  ● Item 2 (SpaceBefore=0pt) ← zeroed ✓
  Y=372  ● Item 3 (SpaceBefore=0pt)  Y=372  ● Item 3 (SpaceBefore=0pt) ← zeroed ✓
  Y=358  ● Item 4 (SpaceBefore=0pt)  Y=358  ● Item 4 (SpaceBefore=0pt) ← zeroed ✓

  All list items now match V14 positions ✓
```

### Real-World Example

An invoice document had a table cell containing a 6-item bullet list. By item 6, the content was 30pt lower in V23 than V14. The single-row fix only corrected item 6; items 2-5 remained drifted. After the sweep, all items snapped back to their correct positions.

---

## 12. Correction Strategy 9 — PARAGRAPH INDENT (Fix Standalone Body Text)

### What Is It?

Not all text is inside a table. Some documents have plain body paragraphs — numbered clauses, standalone sentences, legal text — that drift because V23 changed how paragraph indentation is calculated.

### When Is It Used?

For text that is:
- Not inside any table
- Not in a header or footer
- Simply drifted to the right compared to V14

```
  BEFORE:
  ┌──────────────────────────────────────────────────────────────┐
  │                                                              │
  │           1. The insured agrees to maintain coverage...      │ ← 46pt from margin
  │                                                              │
  │           2. The policy shall not apply to...                │
  │                                                              │
  └──────────────────────────────────────────────────────────────┘

  V14 correct position: X = 36pt
  V23 drifted position: X = 46pt (drift = +10pt)

  AFTER FIX — PARAGRAPH_INDENT:
  ┌──────────────────────────────────────────────────────────────┐
  │                                                              │
  │     1. The insured agrees to maintain coverage...            │ ← 36pt from margin
  │                                                              │
  │     2. The policy shall not apply to...                      │
  │                                                              │
  └──────────────────────────────────────────────────────────────┘

  Fix: ParagraphFormat.LeftIndent reduced by 10pt ✓
```

### Safety Cap: 10pt Maximum

To prevent accidental large shifts (e.g., if the text-matching logic mistakenly identifies the wrong paragraph), paragraph X drift corrections are **capped at 10pt**. Larger drifts are logged as warnings and skipped.

### Text Matching

The system finds the correct paragraph by searching for a normalized (whitespace-stripped) version of the phrase text. To prevent false matches, if a paragraph is more than 4× longer than the search phrase AND doesn't start with the phrase, the match is rejected.

```
  Search phrase: "Prior Policy" (12 chars)

  ✓ Matches: "Prior Policy benefits shall..."
  ✗ Rejected: "...the Prior Policy, if any, are more favorable than the..."
                    (This is 4x longer and the paragraph doesn't START with "Prior Policy")
```

---

## 13. Correction Strategy 10 — HEADER & FOOTER CORRECTION

### What Is It?

Headers and footers are special areas of a Word document that repeat on every page — usually containing page numbers, document title, company name, or policy number.

These use the same correction techniques as body paragraphs (adjust LeftIndent for horizontal drift, adjust SpaceBefore for vertical drift), but require searching in the header/footer sections rather than the main document body.

```
  ┌──────────────────────────────────────────────────────────┐
  │  HEADER:  ACME Insurance Co.            Policy #: 12345  │ ← drifted 5pt right
  ├──────────────────────────────────────────────────────────┤
  │                                                          │
  │  (main document content)                                 │
  │                                                          │
  ├──────────────────────────────────────────────────────────┤
  │  FOOTER:  Confidential           Page 1 of 4            │ ← drifted 5pt right
  └──────────────────────────────────────────────────────────┘

  Fix: ParagraphFormat.LeftIndent in the Header/Footer section reduced by 5pt ✓
```

---

## 14. The Smart Guard — XI / XIgnore (Don't Fix What Isn't Broken)

### The Problem This Solves

Imagine a table where Column 1 has drifted 36pt to the right, but Column 2 is already perfectly aligned with where it should be. If we "fix" the whole table by sliding it 36pt to the left, Column 2 (which was fine!) will now be 36pt too far LEFT.

The XI (X-Ignore) system detects this situation and prevents that mistake.

### How It Works — Finding the "True Reference Position"

The system identifies a **pivot X** — the left-margin reference point that text should align to — by looking at text that is clearly NOT drifted:

```
  STRATEGY: Find text that stayed in the right place

  Option A: Free-text words (not inside any table) that didn't drift
            These are like "anchors" — stable reference points on the page

  Option B: Table cells that have NO drift (green cells in the report)
            These confirm where the margin actually is

  ┌────────────────────────────────────────────────────────────────┐
  │                                                                │
  │ NOTICE: This form must be completed...  ← stable text X=36pt  │ ← PIVOT = 36pt
  │                                                                │
  │         ┌────────────────────┬──────────────────────┐         │
  │         │ Coverage (drifted) │ Limit (also drifted) │         │
  │         │ X=72pt, drift=+36pt│ X=216pt, drift=+36pt │         │
  │         ├────────────────────┼──────────────────────┤         │
  │         │ Bodily Injury      │ $500,000             │         │
  │         └────────────────────┴──────────────────────┘         │
  │                                                                │
  └────────────────────────────────────────────────────────────────┘

  Wait — Column 0 (Coverage) is ALREADY at X=72pt which is 72pt.
  Is that the correct position? NO.
  The free text "NOTICE" is at X=36pt. Column 0's baseline was at X=36pt in V14.
  So Column 0 DID drift right (it should be at 36pt, not 72pt). FIX IT.
```

### The XI Decision: Already in the Right Place

XI fires when a column's current position already **matches the pivot X** — meaning it hasn't drifted and the system should leave it alone.

```
  Example: Multi-column table where only some columns drifted

  V14:  Col A at X=36pt    Col B at X=200pt    Col C at X=300pt
  V23:  Col A at X=36pt    Col B at X=236pt    Col C at X=336pt

  Col A: No drift → it IS the pivot (X=36pt) → XIGNORE (already correct)
  Col B: Drifted +36pt to the right → needs correction
  Col C: Drifted +36pt to the right → needs correction

  System: "Col A is already at the correct position.
           If I slide the whole table left, Col A will end up at X=0pt — WRONG!
           Instead, Col A gets XI (ignored), and Cols B & C are corrected individually."
```

### The Single-Object Guard

A special edge case: what if ALL the text on the page is inside one single table? Then there's no "stable reference text" outside the table to use as the pivot. In this case, the XI system correctly detects "no external reference available" and does NOT fire — the whole table gets corrected normally.

---

## 15. The Smart Guard — Nested Table Ancestry (Don't Double-Fix)

### The Problem

Word documents can have **tables within tables** — an outer table that contains an inner table inside one of its cells. This is common in complex forms.

```
  OUTER TABLE:
  ┌──────────────────────────────────────────────────────┐
  │ Section A: Personal Information                      │
  ├──────────────────────────────────────────────────────┤
  │ ┌────────────────────────────────────────────────┐   │
  │ │ INNER TABLE: Name, DOB, Address fields         │   │
  │ │  ┌─────────────┬──────────────────────────┐   │   │
  │ │  │ First Name  │ [___________________]     │   │   │
  │ │  ├─────────────┼──────────────────────────┤   │   │
  │ │  │ Last Name   │ [___________________]     │   │   │
  │ │  └─────────────┴──────────────────────────┘   │   │
  │ └────────────────────────────────────────────────┘   │
  └──────────────────────────────────────────────────────┘
```

If both the outer table AND the inner table drift by the same 36pt, naively fixing both would result in:
- Outer table: moved 36pt left ✓
- Inner table: ALSO moved 36pt left (but it already moved left with the outer table!) → ends up 36pt too far left!

### The Solution — Ancestry-Aware Correction

The system builds a **"family tree"** of tables: which tables contain which other tables. When fixing a nested (child) table, it first checks how much its parent table was already corrected, and subtracts that amount.

```
  Outer table drift: +36pt → corrected: moved -36pt → net effect on content: 0pt shift
  Inner table drift: +36pt

  Ancestor already applied: 36pt (outer table's correction already moved inner table by 36pt)

  Effective inner table drift = 36pt - 36pt = 0pt → no additional correction needed ✓
```

Tables are always processed **outermost first** to ensure the ancestry calculation is accurate.

---

## 16. Bugs Found & Fixed — A Summary

This section explains the significant bugs that were found and corrected during development.

### BUG #6 — Layout Tables Pushed Into Margin

| | |
|---|---|
| **Problem** | Zero-indent tables (layout tables) were incorrectly corrected by setting negative `LeftIndent`, pushing content into the page margin |
| **Symptom** | Text partially hidden or clipped at left edge of page after correction |
| **Root Cause** | The strategy selection code had a logic flaw: `if (tableInd == 0)` was always true, sending ALL zero-indent tables to TABLE_INDENT strategy |
| **Fix** | New `CELL_PARAGRAPH_INDENT` strategy that adjusts text inside cells rather than moving the table |
| **Impact** | All borderless "layout tables" (numbered clause sections, A./B./C. formatted text) now corrected safely |

### BUG #7 — Cell Paragraphs Unreachable in PATH 3

| | |
|---|---|
| **Problem** | Some text inside table cells couldn't be found by the fallback paragraph correction (PATH 3) |
| **Symptom** | Drift reported and visible, but no correction applied to those words |
| **Root Cause** | PATH 3 explicitly skipped cell paragraphs — but sometimes text inside cells ends up in PATH 3 when the more precise matching (PATH 1/2) doesn't find it |
| **Fix** | Two-pass approach: try body paragraphs first, then try cell paragraphs as a fallback |
| **Impact** | Text inside cells that "fell through" the precise matching now gets corrected |

### BUG #8 — Merged-Cell Width Corruption

| | |
|---|---|
| **Problem** | Setting column widths on tables with merged (spanning) cells caused the merge to be destroyed, corrupting the entire table layout |
| **Symptom** | After correction, long text in merged header cells would wrap to multiple lines; table columns would change size dramatically |
| **Root Cause** | Aspose.Words removes the `gridSpan` attribute when you set `CellFormat.Width` on a merged cell |
| **Fix** | Two-check detection before any width changes: (1) if cell widths vary across rows, merges exist; (2) if any row has fewer cells than expected, merges exist. When detected, skip width changes, only apply table indent |
| **Impact** | All tables with merged header rows or full-width title cells are now safely handled |

### BUG #9 — SpaceBefore Direction Inversion (Made Drift Worse!)

| | |
|---|---|
| **Problem** | The vertical (Y) drift correction formula was backwards — it was *increasing* the gap instead of reducing it |
| **Symptom** | After "correction," text was further from its target position than before; drift reports showed worse scores on second run |
| **Root Cause** | PDF coordinates increase upward, SpaceBefore pushes down. The formula `sb - driftY` was wrong; the correct formula is `sb + driftY` |
| **Fix** | Corrected in three places: table row corrections, body paragraph corrections, and header/footer corrections |
| **Impact** | All vertical drift corrections now move content in the right direction |

### BUG — Cumulative List Spacing

| | |
|---|---|
| **Problem** | Bullet-list items inside cells accumulated drift — the further down the list, the bigger the drift — but only the last (worst) item was detected and corrected |
| **Symptom** | After correction, the last bullet item was fixed but all intermediate items still had drift |
| **Root Cause** | The comparison engine only detected the largest drift (last row). Intermediate rows had smaller drifts that fell below the detection threshold or were classified differently |
| **Fix** | LIST_SPACING_SWEEP: after detecting this pattern, zero SpaceBefore on ALL list items in the table (except the first, which is consistent) |
| **Impact** | Uniform list spacing restored in all bullet/numbered lists inside table cells |

### BUG — Floating Table Ignored LeftIndent

| | |
|---|---|
| **Problem** | Correction appeared to run successfully but had zero visible effect on floating tables |
| **Symptom** | Second-pass comparison showed identical drift to first pass; correction was technically applied but had no effect |
| **Root Cause** | Tables with `tblpPr` (absolute positioning) ignore `LeftIndent` entirely — the absolute anchor position overrides it |
| **Fix** | Detect `TextWrapping.Around` (floating table), and use `AbsoluteHorizontalDistance` instead of `LeftIndent` |
| **Impact** | All floating/anchored tables now correctly repositioned |

---

## 17. The Full Pipeline at a Glance

```
  ╔══════════════════════════════════════════════════════════════════════╗
  ║                    DRIFTCORRECTOR PIPELINE                          ║
  ╠══════════════════════════════════════════════════════════════════════╣
  ║                                                                      ║
  ║  STEP 1: COMPARE (V14 PDF vs V23 PDF)                               ║
  ║  ─────────────────────────────────────────────────────────────────  ║
  ║  • Extract all words + positions from both PDFs                     ║
  ║  • Match words, measure drift (ΔX, ΔY) for each word               ║
  ║  • Group words into Phrase Groups (same-line, same-drift)           ║
  ║  • Detect table grid structure from PDF position patterns           ║
  ║  • Cross-reference with Word document structure (more precise)      ║
  ║  • Apply XI detection (identify already-correct columns)            ║
  ║  • Write *_diff.json report for each document                       ║
  ║  • Write HTML visual report (red = drifted, green = correct)        ║
  ║                                                                      ║
  ║  STEP 2: CORRECT (Edit Word .docx files)                            ║
  ║  ─────────────────────────────────────────────────────────────────  ║
  ║  • For each document, read its _diff.json report                    ║
  ║  • PATH 1: Fix tables using precise Word DOM table indices           ║
  ║    → Pick strategy: CELL_PADDING / TABLE_INDENT / COLUMN_WIDTHS /  ║
  ║                     CELL_PARAGRAPH_INDENT / TABLE_FLOAT_SHIFT       ║
  ║    → Apply ROW_HEIGHT or SPACE_BEFORE for vertical drift            ║
  ║    → Apply LIST_SPACING_SWEEP if cumulative list drift detected     ║
  ║    → Skip tables with XI phrase groups (already correct)            ║
  ║    → Adjust nested tables by ancestor-correction amount             ║
  ║  • PATH 2: Fix remaining tables using PDF heuristic table groups    ║
  ║  • PATH 3: Fix standalone body paragraphs (PARAGRAPH_INDENT)        ║
  ║  • PATH 4: Fix header/footer text                                   ║
  ║  • Save corrected .docx to output folder                            ║
  ║                                                                      ║
  ║  STEP 3: RECONVERT (Word → PDF using V23 engine)                    ║
  ║  ─────────────────────────────────────────────────────────────────  ║
  ║  • Convert all corrected .docx files back to PDF using V23          ║
  ║  • These are the "Modified PDFs"                                    ║
  ║                                                                      ║
  ║  STEP 4: VERIFY (V14 PDF vs Modified PDF)                           ║
  ║  ─────────────────────────────────────────────────────────────────  ║
  ║  • Run the same comparison again                                    ║
  ║  • Compare drift counts: before vs after correction                 ║
  ║  • Generate final HTML report showing improvement %                 ║
  ║  • Remaining red areas = drifts the system couldn't fix             ║
  ║    (shown as acceptable in the modified comparison)                 ║
  ╚══════════════════════════════════════════════════════════════════════╝
```

### What the Before/After Looks Like in the HTML Report

```
  BEFORE CORRECTION (V14 vs V23):
  ┌──────────────────────────────────────────────────────────┐
  │                                                          │
  │  [RED BOX] "Coverage Type"     ← drifted 36pt right     │
  │  [RED BOX] "Bodily Injury"     ← drifted 36pt right     │
  │  [RED BOX] "$500,000"          ← drifted 36pt right     │
  │                                                          │
  └──────────────────────────────────────────────────────────┘

  AFTER CORRECTION (V14 vs Modified V23):
  ┌──────────────────────────────────────────────────────────┐
  │                                                          │
  │  [GREEN] "Coverage Type"       ← back in correct position│
  │  [GREEN] "Bodily Injury"       ← back in correct position│
  │  [GREEN] "$500,000"            ← back in correct position│
  │                                                          │
  └──────────────────────────────────────────────────────────┘
```

---

## Quick Reference — Strategy Selection Decision Tree

```
  Text has drifted. What strategy do we use?

  Is the text inside a TABLE?
  ├── YES
  │   Is the drift UNIFORM (all columns moved the same amount)?
  │   ├── YES, uniform drift
  │   │   Is it a FLOATING TABLE (absolute position)?
  │   │   ├── YES → TABLE_FLOAT_SHIFT (adjust AbsoluteHorizontalDistance)
  │   │   └── NO → inline table
  │   │           Does table LeftIndent explain the drift?
  │   │           ├── YES → TABLE_INDENT (adjust Table.LeftIndent)
  │   │           └── NO → Does cell padding explain the drift?
  │   │                   ├── YES → CELL_PADDING (reduce all cell left padding)
  │   │                   └── NO → Is table at zero indent (layout table)?
  │   │                           ├── YES → CELL_PARAGRAPH_INDENT (adjust cell para indent)
  │   │                           └── NO  → TABLE_INDENT_FALLBACK (best guess)
  │   │
  │   └── NO, non-uniform drift (different columns moved differently)
  │       → COLUMN_WIDTHS (resize individual column widths)
  │       → (also applies TABLE_INDENT for the overall table shift component)
  │       → SKIP columns with merged cells (BUG #8 protection)
  │
  │   Also: Is there vertical drift in the table?
  │   ├── Fixed-height row → ROW_HEIGHT
  │   └── Auto-height row → SPACE_BEFORE
  │       └── Is it a bullet list with cascading drift? → LIST_SPACING_SWEEP
  │
  └── NO, standalone text
      Is it in a HEADER or FOOTER?
      ├── YES → HEADER/FOOTER PARAGRAPH_INDENT / PARAGRAPH_SPACING
      └── NO  → body text → PARAGRAPH_INDENT (capped at 10pt X drift)
                             PARAGRAPH_SPACING (SpaceBefore adjustment)
```

---

## 18. Executive Summary — All 15 Correction Patterns at a Glance

> **For Management: Everything the system does, in one page.**
>
> Think of a Word document as a precisely arranged stage set. When the printing software was upgraded (V14 → V23), some furniture shifted slightly. This system is the stagehand that measures every shift and puts everything back in its exact original position — automatically, across hundreds of documents.

---

### The 15 Ways We Fix Drift

---

**1. Slide the Whole Table Left** *(TABLE_INDENT)*

The entire table slid to the right. We push it back left by the exact measured amount. Simple as sliding a box on a shelf.

---

**2. Shrink the Cell Cushion** *(CELL_PADDING)*

The table frame didn't move, but the text inside each cell shifted right because the internal padding got bigger. We reduce that cushion so the text sits back in its original spot.

---

**3. Resize Individual Columns** *(COLUMN_WIDTHS)*

Different columns drifted by different amounts, meaning one column got too wide. We shrink that column so all subsequent columns slide back into place — like squeezing one item in a row so everything to its right slides back.

> **Built-in safety:** If the table has merged cells (two cells joined into one, common in table headers), resizing those cells would destroy the merge and break the table layout. The system detects this automatically and skips width changes — only sliding the whole table instead.

---

**4. Fix Text Inside Invisible Layout Tables** *(CELL_PARAGRAPH_INDENT)*

Some tables are used purely as layout tools — no visible borders, sitting right at the page margin. These are used to create formatted lists like "A. Claim shall..." or "1. The insured agrees...". If we slide these tables, we push them *behind* the page margin and text disappears. Instead, we adjust the text indent *inside* each cell, leaving the invisible table frame completely untouched.

> **This fixed a serious bug** where the old system was accidentally hiding text off the left edge of the page.

---

**5. Move an Anchored Floating Table** *(TABLE_FLOAT_SHIFT)*

Some tables are "pinned" to an exact spot on the page (like a text box) rather than flowing with the text. For these, the normal slide method has no effect — the pin overrides it. We instead move the pin itself to the corrected position.

> **This fixed a silent failure** where corrections appeared to run successfully but had zero visible effect.

---

**6. Adjust Row Height** *(ROW_HEIGHT)*

When a table row has a fixed height (e.g., "this row must be exactly 20 points tall") and V23 renders it slightly taller, all rows below shift downward. We reduce that row's height so everything below slides back up.

---

**7. Reduce the Gap Above a Paragraph** *(SPACE_BEFORE)*

Word allows an invisible blank gap above any paragraph. If V23 adds more gap than V14, the paragraph appears lower on the page. We reduce that gap to pull the paragraph back up to its correct position.

> **This fixed a critical bug** where the correction formula was backwards — it was *increasing* the gap instead of reducing it, making things twice as bad. The fix was to reverse the math.

---

**8. Zero Out Cascading List Gaps** *(LIST_SPACING_SWEEP)*

For bullet-point or numbered lists inside table cells: V23 adds a small gap above *every* list item, while V14 only added it above the *first* item. By item 6, the content is 6× that gap lower than it should be. We zero out the gap on all list items (keeping just the first one, which both versions agree on), snapping the entire list back to the correct position.

---

**9. Fix a Standalone Paragraph's Indent** *(PARAGRAPH_INDENT)*

For regular body text outside any table — numbered clauses, legal paragraphs — that drifted to the right, we reduce the paragraph's left indent by the measured drift amount.

> **Built-in safety:** Corrections larger than 10 points are skipped with a warning. This prevents accidentally shifting the wrong paragraph if the text-matching finds an imprecise match.

---

**10. Fix a Standalone Paragraph's Vertical Position** *(PARAGRAPH_SPACING)*

Same as above, but for vertical (up/down) drift in body text. We adjust the blank gap above the paragraph to move it up or down to the correct position.

---

**11. Fix Header Text** *(HEADER_INDENT / HEADER_SPACING)*

The repeating header at the top of each page (company name, policy number, document title) can drift independently of the main content. We apply the same indent and spacing adjustments specifically within the header section of the document.

---

**12. Fix Footer Text** *(FOOTER_INDENT / FOOTER_SPACING)*

Same as above for the repeating footer at the bottom of each page (page numbers, confidentiality notices, etc.).

---

**13. Recognize Already-Correct Columns** *(XI / X-IGNORE Guard)*

Before fixing a table, the system checks whether any of its columns are *already in the right position*. If a column is correct, sliding the whole table would move it to the *wrong* position. Those correct columns are marked "X-Ignore" — the system works around them rather than disturbing them.

> **Plain English:** "Don't fix what isn't broken." If Column A is fine but Column B drifted, we don't slide the whole table — we only fix Column B.

---

**14. Avoid Double-Fixing Nested Tables** *(ANCESTRY CORRECTION)*

When a table contains another table inside it, fixing the outer table automatically moves the inner table too. The system tracks this and subtracts the outer table's correction from the inner table's correction — so the inner table gets moved exactly the right amount, not twice.

> **Plain English:** If you slide a desk 3 feet left, and there's a lamp on the desk, the lamp also moved 3 feet left. You don't need to move the lamp separately.

---

**15. Smart Text Matching Guard** *(FALSE MATCH PREVENTION)*

When searching for a paragraph to correct, the system could accidentally match the wrong paragraph if the target phrase appears as a short fragment inside a much longer sentence. For example, "Prior Policy" might correctly appear as a paragraph header — or incorrectly appear buried inside a 200-word legal clause. The system only accepts a match if the paragraph is a close match in length or starts with the search phrase. Too-short phrases inside too-long paragraphs are rejected.

---

### One-Line Summary Table

| # | Fix Name | What It Does | What Would Go Wrong Without It |
|---|---|---|---|
| 1 | Slide the Whole Table | Moves the entire table back left | Table stays shifted to the right |
| 2 | Shrink the Cell Cushion | Reduces internal cell padding | Text inside cells stays too far right |
| 3 | Resize Individual Columns | Narrows an overwide column | Subsequent columns stay shifted right |
| 4 | Fix Invisible Layout Tables | Adjusts text indent inside invisible layout tables | Text gets pushed behind the left page margin |
| 5 | Move Anchored Floating Table | Moves the table's anchor pin | Correction runs but has zero visible effect |
| 6 | Adjust Row Height | Reduces an oversized row's height | All rows below stay pushed down |
| 7 | Reduce Paragraph Gap | Shrinks the blank gap above a paragraph | Paragraph stays too low (or gets pushed further down) |
| 8 | Zero Cascading List Gaps | Removes per-item spacing on list items 2+ | Bullet list items drift further down with each item |
| 9 | Fix Body Text Indent | Reduces left indent of standalone paragraph | Body text paragraph stays shifted right |
| 10 | Fix Body Text Vertical | Adjusts gap above standalone paragraph | Paragraph stays too high or too low |
| 11 | Fix Header Text | Corrects drift in repeating page headers | Header text stays shifted; looks misaligned on every page |
| 12 | Fix Footer Text | Corrects drift in repeating page footers | Footer text stays shifted; looks misaligned on every page |
| 13 | Recognize Correct Columns | Skips columns that are already correct | Already-correct columns get moved to the wrong position |
| 14 | Avoid Double-Fixing Nested Tables | Deducts parent's correction from child table | Inner tables get moved twice as far as needed |
| 15 | Smart Text Matching | Rejects ambiguous paragraph matches | Wrong paragraph in the document gets shifted |

---

*Document generated 2026-03-25. Reflects all corrections implemented through the TABLE_FLOAT_SHIFT and cumulative LIST_SPACING_SWEEP patterns.*
