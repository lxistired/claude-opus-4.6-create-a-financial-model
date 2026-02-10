---
name: ib-financial-model
description: Build investment bank-style financial models in Excel using Python (openpyxl). Covers three-statement linkage (P&L + BS + CF), segment-driven DCF, scenario switching, and IB formatting. Use when the user asks to build a financial model, DCF model, valuation model, or create an Excel-based financial analysis from annual reports or 10-K filings.
---

# Investment Bank Financial Model Builder

## Overview

Build programmatic IB-grade Excel financial models via Python `openpyxl`. The model follows a **segment-driven, three-statement linked DCF** architecture with scenario switching.

## Architecture: 11-Tab Workbook

```
Cover → Key_Summary → Assumptions → Segment_Revenue → Segment_PL
→ Consolidated_PL → BS → Cash_Flow → DCF → Sensitivity → Ratio_Analysis
```

### Data Flow (critical linkage order)

```
Assumptions (scenarios) ──→ Segment_Revenue (bottom-up)
                              ↓
                         Segment_PL (by division)
                              ↓
                         Consolidated_PL (by cost type)
                              ↓                    ↓
                           BS (balance sheet)    DCF
                              ↓                    ↓
                         Cash_Flow            Sensitivity
                              ↓
                         Ratio_Analysis
```

**Key rule**: Build tabs in dependency order. Store row references in a `dict R = {}` for cross-tab formulas.

## Step 1: Extract Data from Source Documents

### From 10-K / Annual Report PDF

Target these sections (use `Read` with offset/limit for large PDFs):

| Data | Location | Priority |
|------|----------|----------|
| Revenue by segment | Note on Segment Information | P0 |
| Segment operating income | Note on Segment Information | P0 |
| Consolidated P&L | Consolidated Statements of Income | P0 |
| Balance Sheet | Consolidated Balance Sheets | P0 |
| Cash Flow | Consolidated Statements of Cash Flows | P0 |
| Segment cost breakdown | Segment note (emp comp / other) | P1 |
| Share count / EPS | Per-share data in P&L | P1 |
| Equity rollforward | Statements of Stockholders' Equity | P2 |

### Searching PDFs

Use `Grep` on the PDF for keywords like "Total assets", "Operating income", "Segment", then `Read` with the right offset. PDF pages ≈ 50-60 lines each.

## Step 2: Model Structure

### Assumptions Tab — The Control Center

```
Row layout: [Parameter | 2023A | 2024A | 2025A | (spacer) | Base | Bull | Bear | Selected | (spacer) | Notes]
```

**CHOOSE/MATCH formula** for scenario switching:
```
=CHOOSE(MATCH($B$2,{"Base","Bull","Bear"},0), F{row}, G{row}, H{row})
```

Every assumption MUST show **historical anchors** (3 years of actuals) so reviewers can judge reasonableness.

### Segment Revenue — Bottom-Up Build

For each business line:
1. Historical revenue (hardcoded, blue font)
2. Forecast: `=Prior × (1 + growth_rate)` where growth links to Assumptions via CHOOSE/MATCH
3. YoY % row below each segment
4. Revenue Mix % section at bottom

### Dual-Layer P&L with Cross-Validation

**Segment_PL**: costs by division (Employee Comp + Other Costs per segment)
**Consolidated_PL**: costs by type (COGS/R&D/S&M/G&A)

Critical: **EBIT in Consolidated_PL = Sum of Segment OI from Segment_PL**

```python
# Consolidated_PL EBIT links to Segment_PL (segment is source of truth)
fval(ws, ebit_row, c, f"=Segment_PL!{CL(c)}{sum_seg_row}")

# Segment_PL cross-check row
fval(ws, diff_row, c, f"={CL(c)}{sum_seg}-{CL(c)}{consol_ebit_link}")
# → must always = 0
```

### Balance Sheet — Cash as Plug

BS forecast logic:

| Line Item | Forecast Method |
|-----------|----------------|
| **Cash** | **PLUG** = Total L&E − NCA − Other CA (ensures A = L+E) |
| Marketable Securities | Grow at stable rate |
| Accounts Receivable | Revenue × AR Days / 365 |
| PP&E | Prior + CapEx − D&A |
| Accounts Payable | COGS × AP Days / 365 |
| Accrued Comp | Revenue × AccComp% |
| Long-term Debt | From Assumptions (absolute) |
| Equity | Prior + NI − Buyback − Dividends + SBC(net) |

**Balance Check row**: `=Total Assets − Total Liabilities − Total Equity` → must = 0

### Cash Flow — WC from BS Delta

```python
# Each WC line links to BS change
# AR change (increase = cash outflow)
f"=-(BS!{CL(c)}{ar_row}-BS!{prev}{ar_row})"
# AP change (increase = cash inflow)
f"=BS!{CL(c)}{ap_row}-BS!{prev}{ap_row}"
```

### DCF: UFCF → Enterprise Value → Equity Value

```
UFCF = NOPAT + D&A − CapEx − ΔNWC
TV = UFCF_terminal × (1+g) / (WACC−g)    [Gordon Growth]
EV = Σ PV(UFCF) + PV(TV)
Equity = EV + Cash − Debt
Price = Equity / Diluted Shares
```

## Step 3: IB Formatting Standards

### Color Coding (must implement + document in legend)

| Style | Meaning | Implementation |
|-------|---------|----------------|
| Blue font (FH) | Historical hardcoded data | `Font(color="305496")` |
| Green font (FL) | Cross-sheet link | `Font(color="548235")` |
| Blue fill + dark font (FI) | Editable assumption | `PatternFill("solid", fgColor="DDEBF7")` |
| Yellow fill (KEY) | Key output row | `PatternFill("solid", fgColor="FFF2CC")` |
| Black font (FN) | Formula / calculated | Default |
| Grey italic (FSM) | Percentage / memo row | `Font(size=9, italic=True, color="808080")` |

### Every Tab Must Have

1. **Title row** (A1, large bold font)
2. **Unit row** (A2, e.g. "USD Million")
3. **Year header** with dark fill + white font
4. **Notes column** (last column) — prediction logic for every row
5. **Data legend** at bottom — color coding explanation
6. **Data sources** at bottom — specific page/note references

### Key Summary Tab (Executive Dashboard)

Links to all other tabs. Contains:
- Core KPIs (Revenue/GP/EBIT/NI + margins + YoY)
- Revenue by segment with mix %
- BS highlights (Total Assets/Equity/Cash)
- ROE / ROA
- Investment highlights + risk warnings (bullet points)

### Ratio Analysis Tab

- Profitability: GPM / OPM / NPM / ROE / ROA
- Efficiency: AR Days / AP Days / CapEx Intensity / CapEx÷D&A
- Leverage: D/A / Current Ratio / Net Cash
- Cash Flow Quality: OCF÷NI / FCF÷Revenue
- Growth: Revenue / NI / EPS YoY

## Critical openpyxl Pitfalls

### 1. Notes column text starting with `=`

**Problem**: openpyxl treats any cell value starting with `=` as a formula. Text like `"=Revenue×TAC Rate"` becomes an invalid formula → Excel repair deletes it.

**Fix**:
```python
def note(ws, r, text):
    if text and text.startswith('='):
        text = text[1:]  # strip leading '='
    ws.cell(row=r, column=NCOL, value=text)
```

### 2. Notes column overlapping data columns

**Problem**: If Assumptions has 5 forecast years (2026-2030) in cols F-J, and NCOL=10 (J), Notes overwrites 2030E data.

**Fix**: Set `NCOL = 11` (column K) globally. Leave col J free for data.

### 3. CHOOSE/MATCH with array constants

The formula `CHOOSE(MATCH($B$2,{"Base","Bull","Bear"},0),...)` uses inline array `{}`. This works in modern Excel but may warn in older versions. NOT a CSE formula — do NOT wrap in extra braces.

### 4. Cross-sheet formula timing

Build tabs in dependency order. For circular references (e.g., Segment_PL cross-check needs Consolidated_PL EBIT which needs Segment_PL OI), use **two-pass**:
1. First pass: build Segment_PL, leave cross-check link row empty
2. Build Consolidated_PL, store EBIT row number
3. Second pass: go back and fill the cross-check link

```python
# After building Consolidated_PL
ws_seg = wb["Segment_PL"]
for c in range(2, 10):
    ws_seg.cell(row=cross_check_row, column=c,
                value=f"=Consolidated_PL!{CL(c)}{ebit_row}")
```

### 5. Cash plug formula — subtraction logic

**Wrong**: `f"=...−{CL(c)}{a}+{CL(c)}{b}+{CL(c)}{c_}"` (minus only applies to first term)
**Right**: `f"=...−{CL(c)}{a}−{CL(c)}{b}−{CL(c)}{c_}"` (explicit minus for each)

## Helper Function Pattern

```python
def hval(ws, r, col, val, fmt="#,##0"):
    """Historical value — blue font"""
    c = ws.cell(row=r, column=col, value=val)
    c.number_format = fmt; c.font = Font(color="305496"); c.border = BD

def fval(ws, r, col, formula, fmt="#,##0"):
    """Formula value — black font"""
    c = ws.cell(row=r, column=col, value=formula)
    c.number_format = fmt; c.border = BD

def lval(ws, r, col, formula, fmt="#,##0"):
    """Link from other sheet — green font"""
    c = ws.cell(row=r, column=col, value=formula)
    c.number_format = fmt; c.font = Font(color="548235"); c.border = BD

def arow(ws, r, label, h1, h2, h3, base, bull, bear, fmt, note_text):
    """Assumption row: 3 historical + 3 scenarios + CHOOSE/MATCH selected"""
    # Historical in cols 2-4 (grey fill, blue font)
    # Scenarios in cols 6-8 (blue fill, editable)
    # Selected in col 9 (CHOOSE formula)
    # Note in col NCOL
```

## Verification Checklist

After generating the workbook, verify programmatically:

```python
import openpyxl
wb = openpyxl.load_workbook("output.xlsx")

# 1. Segment_PL cross-check = 0 for historical years
# 2. BS balance check = 0 for historical years
# 3. EBIT link correct: Consolidated_PL!EBIT → Segment_PL!SumOI
# 4. No formula text in Notes column (check XML for <f> tags in Notes col)
# 5. All year columns have data (no overwrites from Notes)
```

## CapEx Handling

When management provides CapEx guidance (e.g., "$75B per quarter"), use **absolute value inputs** in Assumptions, NOT percentage of revenue. This is more accurate for capital-intensive periods.

```python
# Each forecast year gets its own row with absolute CapEx
arow(ws, 72, "2026E CapEx", hist_23, hist_24, hist_25,
     180000, 185000, 175000, "#,##0", "Q4'25 Earnings: $75B/Q")
```

## Reference Files

- For a complete working example, see [example-structure.md](example-structure.md)
