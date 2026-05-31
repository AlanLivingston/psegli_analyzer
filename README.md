# PSEG Long Island Consumption Analyzer

A command-line tool that reads a downloaded PSEG Long Island interval-usage CSV,
splits it into billing cycles, and produces an Excel workbook that estimates the
full bill — every line item that appears on a real PSEG bill — under four
residential rate plans: **580**, **180**, **194**, and **195**.

Rates are **date-matched**: the tool keeps a monthly rate schedule and, for each
billing cycle, blends the rates day-by-day across the cycle. A cycle that
straddles a rate change (such as the January 1 increase) or a season boundary
gets the correct blended rate — the way PSEG actually bills.

---

## Requirements

- Python 3.8 or newer
- Python packages: `pandas`, `openpyxl`
- (Optional) Internet access, for refreshing historical rates from the PSEG LI website

Install the dependencies:

```bash
pip install pandas openpyxl
```

---

## Quick start

```bash
python pseg_analyzer.py --cons_file Usage.csv
```

This analyzes the entire data period as a single billing cycle and writes
`Usage_analysis.xlsx` in the current folder.

To break the data into your actual billing cycles and match the demand figure on
your bill:

```bash
python pseg_analyzer.py --cons_file Usage.csv \
    --split 2025-12-15 2026-01-16 2026-02-13 2026-03-16 2026-04-16 \
    --demand 16
```

---

## Command-line options

| Option | Required | Description |
|---|---|---|
| `--cons_file <file.csv>` | Yes | The consumption CSV exported from PSEG LI. |
| `--split <date> [<date> ...]` | No | Billing-cycle **end** dates in `YYYY-MM-DD` form. See "Splitting into billing cycles" below. If omitted, the whole period is one cycle. |
| `--out <file.xlsx>` | No | Output filename. Defaults to `<cons_file>_analysis.xlsx`. |
| `--demand <kW>` | No | Override the peak demand used for the Customer Benefit Contribution. PSEG bills the CBC on a fixed kW (commonly the solar interconnection size), so set this to the kW shown on your bill — typically `16`. If omitted, the tool estimates demand from the data. |
| `--refresh` | No | Force a re-scrape of any month whose rates are still flagged as estimated (requires internet). |

---

## Input CSV format

The tool expects the standard PSEG LI interval export with three columns:

| Column | Meaning |
|---|---|
| `Start` | Interval start timestamp (e.g. `05/30/2025 12:15:00 AM`) |
| `Meter` | Meter label, e.g. `Meter #95175892` |
| `kWh` | Energy for the interval |

**Solar / net metering.** If your export includes a generation meter — a meter
whose number ends in `g` (e.g. `Meter #95175892g`) — the tool detects it
automatically. Consumption (import) and generation (export) are netted per
interval, so all billing is computed on **net** consumption, matching PSEG's
net-metering bills. Any `- Off-Peak` / `- On-Peak` suffixes on the meter label
are ignored; time-of-use buckets are derived from the timestamps.

---

## Splitting into billing cycles

Each `--split` date is the **inclusive last day** of a billing cycle; the next
cycle begins the following day. `N` split dates therefore produce `N + 1`
cycles. The first cycle starts at the first day of data; the last ends at the
final day.

**Example.** Data spanning Jan 1 – Jun 30 with:

```bash
--split 2025-02-15 2025-03-31 2025-05-15
```

produces four cycles:

1. Jan 1 – Feb 15
2. Feb 16 – Mar 31
3. Apr 1 – May 15
4. May 16 – Jun 30

> Note: PSEG's printed bills share the read date between consecutive cycles
> (e.g. "Nov 15 – Dec 15" then "Dec 15 – Jan 16"). This tool uses
> non-overlapping cycles, so a single day's usage may shift between adjacent
> cycles versus the printed bill. The effect on cycle totals is small.

---

## Output workbook

The workbook contains three tabs.

### `Dashboard`
The main view. For each billing cycle, plus a **TOTAL** column:

- **Usage** — days, net consumption, the time-of-use buckets (super-off-peak
  10 p.m.–6 a.m., off-peak, on-peak weekday 3–7 p.m.), and peak demand.
- **Per-rate bill breakdown** — for each of Rates 580, 180, 194, and 195, every
  bill line item: Basic Service, Customer Benefit Contribution, Merchant
  Function Charge, Delivery (tiered or time-of-use), Delivery & System subtotal,
  Power Supply, DER Charge, Delivery Service Adjustment, Revenue Decoupling
  Adjustment, NY State Assessment, Revenue-Based PILOTs, Suffolk Property Tax
  Adjustment, Sales Tax, and Total Charges.
- **Rate comparison** — total bill per plan side by side, with a
  **Lowest-cost plan** row that flags the cheapest rate for each cycle.

All bill figures are **live Excel formulas** that reference the other two tabs,
so editing a rate or a usage value reflows the whole dashboard.

### `EffectiveRates`
The date-weighted rate each cycle actually used — the day-by-day blend of the
monthly schedule across that cycle, with rate changes and seasons already mixed
in. This is what the Dashboard formulas reference.

### `Rates`
The underlying **monthly rate schedule** (the source data): supply charge,
delivery rates, and rider rates for every month. A `flag` column shows the
provenance of each month's rates:

| Flag | Meaning |
|---|---|
| `confirmed` | Verified against actual bills |
| `scraped` | Pulled from the PSEG LI website |
| `estimate` | Supply charge may be exact, but delivery/rider rates are estimated — refresh when online |
| `stale` | Month predates PSEG's published history (before June 2025); earliest known rates used |
| `carried` | Could not be obtained; carried forward from the most recent month |

Cells highlighted in yellow are estimated or stale and worth verifying.

---

## How rates are kept current

The tool maintains a local cache, **`pseg_rates_cache.json`**, alongside the
script. On each run it:

1. Determines which months your data spans.
2. If any of those months are missing from the cache, it scrapes them from the
   PSEG LI rate-information page:
   `https://www.psegliny.com/aboutpseglongisland/ratesandtariffs/rateinformation`
3. Falls back to the most recent known rates (with a printed warning) if the
   site is unreachable or a month cannot be parsed.
4. Saves the updated cache for next time.

Historical rates are published back to **June 2025**. For any data before that,
the tool uses the earliest known rates and prints a `stale` warning. Use
`--refresh` to force re-fetching months whose rates are still estimated.

Because website layouts change over time, the scraper is best-effort: if PSEG
restructures the page, the auto-refresh may stop finding values, but your cached
and seeded rates keep the analysis working. You can also edit
`pseg_rates_cache.json` (or the `Rates` tab) directly to correct any value.

---

## Console warnings

A typical run prints messages such as:

```
WARNING: Dates in/before 2025-05 predate published history (back to 2025-06); using earliest known rates.
WARNING: Delivery/rider rates for 2025-07 are estimated (PSC may be exact); run with --refresh once online to scrape exact values.
```

These tell you which cycles rest on stale or estimated rates so you can judge how
much to trust those numbers.

---

## Notes and limitations

- **Estimator, not a bill.** Output is an estimate. Validated against real bills
  it tracks within roughly 1% over a season, but individual cycles can differ by
  a small amount due to the cycle-boundary convention above and because a few
  riders (PILOTs, Suffolk property tax, sales tax) are modeled as steady
  percentages rather than their exact monthly values.
- **Demand for the CBC.** Pass `--demand` to match the kW on your bill; otherwise
  the CBC line is based on the metered peak, which is usually higher.
- **Suffolk County.** The Suffolk Property Tax Adjustment line applies to Suffolk
  County customers only. Set its rate to `0` on the `Rates` tab if it does not
  apply to you.
- **Recalculation.** Opening the workbook in Excel or LibreOffice calculates the
  formulas automatically. If you build pipelines around it, note that openpyxl
  writes formulas as text; a spreadsheet application (or a headless LibreOffice
  recalculation) is needed to populate the computed values.

---

## Example session

```bash
$ python pseg_analyzer.py --cons_file Usage.csv \
      --split 2025-11-14 2025-12-15 2026-01-16 2026-02-13 2026-03-16 2026-04-16 2026-05-15 \
      --demand 16

Loading consumption data...
  35,036 intervals, 2025-05-30 to 2026-05-29
Loading & refreshing rate schedule...
  Rate schedule already covers the data period; no scrape needed.
  8 billing cycle(s).
  WARNING: Dates in/before 2025-05 predate published history (back to 2025-06); using earliest known rates.
  WARNING: Delivery/rider rates for 2025-06 are estimated ...
Wrote Usage_analysis.xlsx
```

Open `Usage_analysis.xlsx` and start on the **Dashboard** tab.
