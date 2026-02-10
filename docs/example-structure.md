# Example: Complete Model Tab Structure

## Tab-by-Tab Row Reference Map

This shows the typical row layout and cross-tab references for a segment-driven 3-statement DCF model.

## 1. Cover
```
R1: Company Name + Ticker
R2: Model description + version + date
R4-R16: Usage instructions (edit blue cells, check cross-validation, etc.)
```

## 2. Key_Summary (Executive Dashboard)
```
R4: Header (KPIs | 2023A | 2024A | 2025A | 2026E | 2027E | 2028E | Notes)
R5: Total Revenue          ← links Consolidated_PL
R6:   YoY %
R7: Gross Profit           ← links Consolidated_PL
R8:   GPM %                ← =R7/R5
R9: EBIT                   ← links Consolidated_PL
R10:  OPM %
R11: Net Income            ← links Consolidated_PL
R12:  NPM %
R14: EPS                   ← links Consolidated_PL

R16: Revenue by Segment header
R17-R22: Segment revenue + mix %  ← links Segment_Revenue

R24: BS Highlights header
R25-R29: Total Assets / Equity / Cash / ROE / ROA  ← links BS + PL

R32: Investment Highlights (bullet points + risk warnings)
```

## 3. Assumptions (Control Center)
```
R2: Scenario selector (B2 = "Base" with data validation)
R4: Main header (Parameter | 2023A | 2024A | 2025A | _ | Base | Bull | Bear | Selected | _ | Notes)

R6-13: WACC section (Rf, Beta, ERP, Kd, ETR, D/TC, g)
R15-18: Derived WACC (Ke, After-tax Kd, WACC)
R20-23: Valuation Bridge (Cash, Debt, Shares)

R25: Revenue Growth header (separate sub-header with 2026E-2030E in cols F-J)
R27-32: Base case growth by segment (6 rows)
R34-40: Bull case growth
R42-48: Bear case growth

R50-57: Segment cost structure (Emp Comp%, Other Costs%, OI Margin)
R59-68: Consolidated cost rates (TAC, COGS, R&D, S&M, G&A, D&A, SBC, NWC)
R70-76: CapEx (absolute values per year, NOT % of revenue)
R78-80: Capital return (Buyback, DPS)
R82-91: BS assumptions (AR Days, AP Days, AccComp%, AccExp%, etc.)
```

## 4. Segment_Revenue (Bottom-Up)
```
For each segment:
  Revenue row (historical hardcoded + forecast formula)
  YoY % row

Sections:
- Google Advertising (Search + YouTube + Network → subtotal)
- Subscriptions/Platforms/Devices
- Google Services Total (= Advertising + Subs)
- Google Cloud
- Other Bets
- Hedging gains/losses
- TOTAL Revenue (= Services + Cloud + OtherBets + Hedge)
- Revenue Mix % (each segment / total)
```

## 5. Segment_PL (by Division)
```
Google Services:
  Revenue            ← links Segment_Revenue
  (-) Emp Comp       = Revenue × rate from Assumptions
  (-) Other Costs    = Revenue × rate from Assumptions
  Total Costs
  Operating Income   = Revenue − Costs   ★ KEY OUTPUT
  Op Margin %

Google Cloud:
  (same structure)

Other Bets:
  Revenue            ← links Segment_Revenue
  OI                 = Revenue × OI margin from Assumptions

Alphabet-level:
  Unallocated Costs  = -Assumptions!$I$xx

CROSS-CHECK:
  Sum of Segment OI  = Svc_OI + Cloud_OI + OB_OI + Alpha_Costs
  Consolidated EBIT  ← links Consolidated_PL (filled in second pass)
  Difference          = Sum − Consol  ★ MUST = 0
```

## 6. Consolidated_PL (by Cost Type)
```
Revenue              ← links Segment_Revenue
  Ad Revenue (memo)  ← for TAC calculation

COGS:
  TAC                = Ad Revenue × TAC rate
  Other COGS         = Revenue × rate
  Total COGS
Gross Profit         = Revenue − COGS
  GPM %

OpEx:
  R&D                = Revenue × rate
  S&M                = Revenue × rate
  G&A                = Revenue × rate
  Total OpEx (estimated)

EBIT                 ← links Segment_PL!Sum_OI  ★ SEGMENT-DRIVEN
  EBIT Margin %
  Implied OpEx       = GP − EBIT (transparency row)

OI&E
Pre-tax Income       = EBIT + OI&E
Tax                  = Pre-tax × ETR
Net Income           = Pre-tax − Tax
  NI Margin %

Memo: D&A, SBC
Diluted Shares
Diluted EPS          = NI / Shares
```

## 7. BS (Balance Sheet)
```
Current Assets:
  Cash ← PLUG (= Total L&E − NCA − other CA)
  Marketable Securities (grow at rate)
  Total Cash+Securities
  AR (= Revenue × AR Days / 365)
  Other Current
  Total Current Assets

Non-Current Assets:
  Non-marketable Securities
  Deferred Tax
  PP&E (= Prior + CapEx − D&A)  ← links Assumptions + PL
  Operating Lease
  Goodwill
  Other
  Total Non-Current

TOTAL ASSETS = CA + NCA

Current Liabilities:
  AP (= COGS × AP Days / 365)
  Accrued Comp (= Revenue × %)
  Accrued Expenses (= Revenue × %)
  Revenue Share (= Ad Revenue × %)
  Deferred Revenue (growth)
  Total CL

Non-Current Liabilities:
  LT Debt (from Assumptions)
  Tax NC, OpLease Liab, Other
  Total NCL

TOTAL LIABILITIES = CL + NCL

Equity = Prior + NI − Buyback − Dividends + SBC(net)

TOTAL L&E = Liab + Equity
Balance Check = Assets − L&E  ★ MUST = 0
```

## 8. Cash_Flow
```
Operating:
  Net Income         ← PL
  (+) D&A            ← PL
  (+) SBC            ← PL
  Deferred Tax / Securities GL / Other (hardcoded or zero for forecast)

  WC Changes (each links to BS delta):
    Chg AR = -(BS_AR_current − BS_AR_prior)
    Chg AP = BS_AP_current − BS_AP_prior
    Chg AccExp, RevShare, DefRev (same pattern)

  CFO = sum of all above

Investing:
  CapEx              ← Assumptions (negative)
  CapEx / Revenue %

  FCF = CFO + CapEx (CapEx is negative)
  FCF Margin %

Financing:
  Buyback            ← Assumptions (negative)
  Dividends          = -DPS × Shares
```

## 9. DCF
```
EBIT                 ← Consolidated_PL
(-) Tax on EBIT      = EBIT × ETR
NOPAT                = EBIT − Tax
(+) D&A              ← Consolidated_PL
(-) CapEx            ← Assumptions
(-) ΔNWC             = ΔRevenue × NWC%
UFCF                 = NOPAT + D&A − CapEx − ΔNWC

Discount Factor      = 1/(1+WACC)^n
PV of UFCF

Terminal Value       = UFCF_last × (1+g) / (WACC−g)
PV of TV             = TV × DF_last
TV as % of EV

EV Bridge:
  Sum PV UFCF
  PV of TV
  EV = UFCF_PV + TV_PV
  (+) Cash
  (-) Debt
  Equity Value = EV + Cash − Debt
  Diluted Shares
  Implied Price = Equity / Shares  ★ KEY OUTPUT
```

## 10. Sensitivity
```
2D data table: WACC (columns) × Terminal g (rows)
Each cell recalculates implied share price with different WACC and g
```

## 11. Ratio_Analysis
```
Profitability: GPM / OPM / NPM / ROE / ROA
Efficiency: AR Days / AP Days / CapEx Intensity / CapEx÷D&A
Leverage: D/A / Current Ratio / Net Cash
Cash Flow: OCF÷NI / FCF÷Revenue
Growth: Revenue YoY / NI YoY / EPS YoY
Investment Highlights (bullet points)
```

## Cross-Tab Reference Tracking

Use a Python dict to track row numbers across tabs:

```python
R = {}
# Segment_Revenue
R["total_rev"] = total_rev_row
R["ad_total"] = ad_total_row
R["cloud_rev"] = cloud_rev_row

# Segment_PL
R["sum_seg"] = sum_seg_oi_row

# Consolidated_PL
R["rev_pl"] = revenue_row
R["cogs_pl"] = cogs_row
R["ebit_pl"] = ebit_row
R["ni_pl"] = ni_row
R["da_pl"] = da_row
R["sbc_pl"] = sbc_row
R["shares_pl"] = shares_row
R["eps_pl"] = eps_row

# BS
R["total_assets"] = total_assets_row
R["equity"] = equity_row

# Cash_Flow
R["cfo"] = cfo_row
R["capex_cf"] = capex_row
R["fcf"] = fcf_row
```
