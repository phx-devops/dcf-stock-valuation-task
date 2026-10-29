# Nike DCF Valuation Task

## Overview

Build a formula-driven discounted cash flow (DCF) valuation of Nike, Inc. (NKE)
from the supplied SEC filings. The workflow combines PDF financial-statement
extraction, spreadsheet construction, financial modeling, and numeric
self-verification. The purpose of pinning the modeling assumptions is to make
the result deterministic and the evaluation fair: judgment calls are specified
in the input files rather than left open-ended.

**Category:** data-processing / finance  
**Difficulty:** hard  
**Tags:** pdf-extraction, xlsx, financial-modeling, dcf  
**Suggested time:** 30–45 minutes of agent time

## Provided files

- `/app/filings/` — Nike's three most recent 10-K PDFs and latest 10-Q PDF.
- `/app/formulae/` — PNG references for the required financial formulas. Use
  these when implementing the model.
- `/app/assumptions.json` — pinned WACC, forecast, terminal value, and valuation
  assumptions. Read this file and follow it; do not substitute your own
  assumptions.
- `/app/market_data.json` — fixed NKE share price and currency. No internet
  lookup is needed or expected.

## Work to complete
Execute the command to extract the inputs and build the DCF workbook.
```sh
make -C verifier_src extract-inputs
make -C verifier_src build-workbook
```
1. **Extract filing data.** Read the three 10-Ks and latest 10-Q. Extract
   revenue, EBIT, cash and equivalents, and total debt for each requested
   reporting period. Retain source filing and page references for the extracted
   values. Include other filing inputs needed by the model, such as diluted
   shares, interest expense, capex, D&A, taxes, and NWC components, where
   applicable.
2. **Create the workbook.** Save
   `/app/output/nke_dcf.xlsx` with these sheets: `Assumptions`, `Inputs`,
   `WACC`, `Historicals`, and `DCF`.
   - In `Inputs`, present extracted values, reporting periods, units, and source
     references.
   - In `WACC`, calculate cost of equity, cost of debt, capital weights, and
     WACC using the supplied assumptions and the formula references.
   - In `Historicals`, calculate revenue growth, EBIT, EBIT margin, effective
     tax rate, and net investment (capex − D&A + change in NWC) where data is
     available.
   - In `DCF`, forecast revenue and EBIT for ten years, then calculate free cash
     flow, discount factors, present values, Gordon-growth terminal value,
     enterprise value, equity value, and value per diluted share.
3. **Apply the pinned assumptions.** Use the specified 3-year historical CAGR
   as the starting revenue growth rate and fade it in equal steps to terminal
   growth by year 10. Use 3-year averages for EBIT margin, tax rate, D&A,
   capex, and NWC. Follow the specified base period, capital structure source,
   debt treatment, terminal-value timing, cash/debt bridge, and diluted-share
   source in `/app/assumptions.json`.
4. **Make the model auditable.** All calculated workbook cells must be Excel
   formulas that reference the `Inputs`, `WACC`, or other relevant model cells;
   do not hardcode calculated results. Keep extracted filing data distinct from
   calculated values.
5. **Write the summary.** Save `/app/output/summary.json` with numeric values
   for `wacc`, `enterprise_value`, `equity_value`, `value_per_share`,
   `market_price`, and `upside_pct`. Use the fixed price in
   `/app/market_data.json`, and calculate upside as
   `value_per_share / market_price - 1`.

## Deliverables

- `/app/output/nke_dcf.xlsx`
- `/app/output/summary.json`

## Verification

The verifier checks that both deliverables exist; the workbook contains the
required sheets; extracted filing values, source citations, and assumptions are
correct; and the workbook recalculates without spreadsheet errors. It compares
model outputs against a reference implementation, checks that the JSON summary
matches recalculated workbook cells, and checks key model relationships,
including decreasing discount factors, a terminal value that is 50–90% of
enterprise value, the equity-value bridge, and use of diluted shares.

The expected tolerances are exact extracted values at the filing's reporting
precision, WACC within 10 basis points, and model outputs within 1% of the
reference. The fixed market price is read from `market_data.json`; the reported
upside is informational and is not itself a pass/fail valuation target.

In the provided verifier environment, run:

```sh
make -C verifier_src test-outputs
```

The verifier uses LibreOffice headless to recalculate the workbook. Install the
dependencies first if they are not already available:

```sh
make -C verifier_src install-dependencies
```

## Solution steps

The agent gets Nike's last three 10-Ks and latest 10-Q as PDFs. It must extract
the financials, build a formula-driven DCF workbook, and report an implied value
per share. It is a realistic analyst workflow that combines PDF parsing,
spreadsheet generation, financial modeling, and numeric self-verification.

**Verifier:** Check for the existence of the DCF model workbook and the stock
valuation. The fixed market value is provided in `/app/market_data.json`.

**Why this task is difficult:** It requires accurate PDF extraction, sound
financial modeling, and construction of a formula-driven spreadsheet whose
results can be independently recalculated and verified.

**Terminal Bench 3.0 criteria:** verifiable, solvable, realistic, valuable, and
interesting.
