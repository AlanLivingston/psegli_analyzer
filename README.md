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
your bill, pass the meter-read dates from your bills:

```bash
python pseg_analyzer.py --cons_file Usage.csv \
    --read-dates 2025-11-15 2025-12-15 2026-01-16 2026-02-13 2026-03-16 2026-04-16 2026-05-15 \
    --demand 16
```

---

## Command-line options

| Option | Required | Description |
|---|---|---|
| `--cons_file <file.csv>` | Yes | The consumption CSV exported from PSEG LI. |
| `--read-dates <date> [<date> ...]` | No | Meter-read boundary dates in `YYYY-MM-DD` form. Pass each bill's **Service From** date plus the last bill's **Service To** date. See "Matching bills with `--read-dates`" below. If omitted, the whole period is analyzed as one cycle. |
| `--out <file.xlsx>` | No | Output filename. Defaults to `<cons_file>_analysis.xlsx`. |
| `--demand <kW>` | No | Override the peak demand used for the Customer Benefit Contribution. PSEG bills the CBC on a fixed kW (commonly the solar interconnection size), so set this to the kW shown on your bill — typically `16`. If omitted, the tool estimates demand from the data. |
| `--suffolk` / `--nassau` / `--rockaways` | No | Service-area tax treatment (mutually exclusive). Selects the local sales-tax rate on residential electricity. **`--suffolk` is the default.** See "Service area" below. |
| `--sales-tax <frac>` | No | Override the local sales-tax fraction directly (e.g. `0.03`), regardless of county — useful for the Nassau school-district exceptions. |
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

## Defining billing cycles with `--read-dates`

`--read-dates` is how you split the data into your actual billing cycles. You
pass the **meter-read boundary dates** straight off your bills and the tool
builds each cycle as the half-open interval `[read, next_read)` — i.e. a read
date's own usage rolls into the cycle that *starts* on it, which is how PSEG's
reads actually fall (taken at around midnight on the read date). If omitted, the
entire data period is analyzed as a single cycle.

**Which dates to pass.** Consecutive bills share a boundary: one bill's
*Service To* equals the next bill's *Service From*. So the full ordered list is
**every bill's Service From date, plus the last bill's Service To date**.
`N` billing cycles therefore need `N + 1` dates.

**Example.** Six bills with service periods Nov 15→Dec 15, Dec 15→Jan 16,
Jan 16→Feb 13, Feb 13→Mar 16, Mar 16→Apr 16, Apr 16→May 15:

```bash
python pseg_analyzer.py --cons_file Usage.csv \
    --read-dates 2025-11-15 2025-12-15 2026-01-16 2026-02-13 2026-03-16 2026-04-16 2026-05-15 \
    --demand 16
```

produces six cycles — `[Nov 15, Dec 15)`, `[Dec 15, Jan 16)`, … — labeled with
their true service periods. Any consumption data lying outside the first and
last dates is excluded (it belongs to bills you didn't list).

Because the read date's usage is attributed to the cycle that starts on it,
each cycle's net-kWh typically matches the bill to within a single
meter-register tick (the meter reports in ~40 kWh steps), with no boundary day
misattributed.

---

## Service area (`--suffolk` / `--nassau` / `--rockaways`)

PSEG Long Island operates one LIPA tariff across its whole territory — Nassau
County, Suffolk County, and the Rockaway Peninsula in Queens — so the rate
*schedules* (580, 180, 194, 195), the delivery tiers, the Power Supply Charge,
and the PILOT/property-tax recovery are identical everywhere. The single
component that genuinely varies by location is the **local sales tax on
residential electricity**, and the county switch selects it:

- `--suffolk` (default): ~2.5%, bill-verified for this account.
- `--nassau`: 0% — Nassau exempts residential energy from its local sales tax
  (since 2010). The Glen Cove and Long Beach school districts are the exception
  at ~3%; use `--sales-tax 0.03` there.
- `--rockaways`: ~4.5%, New York City's local rate (estimated — confirm against
  a Rockaways bill, or set `--sales-tax`).

Suffolk is the default. (It is, of course, the best county — though if your goal
is strictly the lowest sales tax on electricity, Nassau's 0% wins that line.)
Because sales tax applies roughly proportionally to every rate plan, the county
choice barely moves *which* rate is cheapest; it mainly affects the absolute
all-in totals. Use `--sales-tax <frac>` to override the rate directly for any
edge case.

---

## Energy-bank simulation

A naive analysis nets each time-of-day bucket *within* a cycle: a month where
afternoon solar pushes a bucket net-export shows up as an in-cycle credit against
that same bill. PSEG doesn't work that way. On a Time-of-Day rate it keeps
**separate credit banks per period** — three for Rate 195 (peak, off-peak,
super-off-peak; 1:2:4) and two for Rate 194 (peak, off-peak; 1:2). A net-export
period deposits kWh into its own bank, and moving credits between banks requires
the Excess Generation Exchange form, at up to two transfer statements per
request filed at least 10 business days before a meter read.

The tool models this on **every run** — there is no flag to enable it. Both TOD
rates are simulated, their banked billable kWh feed the Dashboard (so all four
plans are compared on the same realistic basis, see the **BILLABLE AFTER
BANKING** rows), and a `Banking194` and a `Banking195` tab are always written
with the full detail. The console prints a one-line comparison naming the cheaper
TOD plan.

Why this is safe to combine with the live, editable rate cells: banking decisions
depend only on kWh and the tariff's fixed value ratios (1:2:4 for 195, 1:2 for
194), not on the dollar rates, and the transfers are value-neutral. So editing a
price on the `Rates` tab updates every plan's bill correctly while the banked kWh
stay valid. (Re-banking would only matter if you edited the rate *structure* so
drastically that the value ordering of the periods changed.)

The Dashboard's "Lowest-cost plan" row is the authority on which plan is cheapest
overall, because it also weighs the flat/tiered plans (580, 180) that have no
per-period banks; the two Banking tabs are the drill-down on the TOD mechanics
behind those same numbers.

The simulation models this properly. For each cycle it deposits net-export
surplus into the period's bank, draws each period's consumption against its own
carried-in bank first, then runs a **profile-agnostic greedy transfer step** to
cover any remaining deficits from the surplus banks. It covers the costliest
deficit first and drains the lowest-value surplus bank first, using whole-kWh
source increments and value-neutral exchanges (1:2:4 on Rate 195, 1:2 on Rate
194). This adapts to any usage profile: an overnight-heavy heating/EV customer
ends up moving peak/off surplus into super-off, while a daytime-heavy household
with overnight export would move super-off into peak/off instead. With so few
banks, at most two transfers are ever needed to rebalance, so the form's
two-statement cap never truncates a needed move. Deposits made in a cycle are
only usable from the next cycle, mirroring the form's timing. Banked credits
roll forward; per PSEG's operation letter, residential (non-demand) customers
keep unused credits for up to 20 years, with only a balance still unused after
the twentieth year forfeited — there is no annual buyback. The balance shown at
the end of the window simply carries into the next period.

The `Banking194`/`Banking195` tab shows, per cycle, the net by period, bank
balances, the credit transfers made, billable kWh after banking, and the all-in
charge. It also includes an **optimality check** with two reference points:

- **Value-neutral optimum** — the best bill if credits were a single fungible
  pool rebalanced across buckets at the 1:2:4 weights. The gap from the greedy
  total is the cost of real transfer frictions (whole-kWh rounding, the two-per-
  cycle cap) and is typically under a dollar a year, because the exchange ratios
  are value-neutral.

There is a separate, real (if small) gain available from *choosing the transfer
destination* by season rather than rebalancing — see "Transfer-ratio arbitrage"
below — but it is a deliberate manual strategy, not something the greedy
simulation scores.

A **usage-profile** block reports each period's annual net (consumer vs producer)
and flags forfeiture exposure: a household that is a net *producer* in value
terms banks more than it consumes, so some credits never find a deficit to offset
and would accumulate toward the 20-year forfeiture; a net consumer's banks drain
through ordinary use and nothing is exposed. For each TOD rate the console prints
the banked total, the optimality gap, and the profile, and names the cheaper of
the two.

### Transfer-ratio arbitrage (hoard-and-convert)

The greedy simulation lets each period's banked credit offset its *own* period's
usage (off-peak credit covers off-peak heating, etc.). There is a small, real gain
available from instead steering off-peak credit into **super-off** during the
winter PSC season. It is pure ratio arbitrage, not calendar timing.

**Why it works.** Transfers are a fixed kWh swap — 1 off-peak credit always buys 2
super-off kWh (and 1 peak buys 4). But the *true* dollar ratio between periods
moves with the season:

| credit → super-off | true price ratio | form pays | verdict |
|---|---|---|---|
| off-peak (summer) | ≈ 2.20 | 2 | form **underpays** — don't convert |
| off-peak (winter) | ≈ 1.81 | 2 | form **overpays** — convert |
| peak (summer) | ≈ 4.29 | 4 | form underpays — keep on peak |
| peak (winter) | ≈ 4.04 | 4 | form underpays — keep on peak |

In the winter PSC season the off-peak rate falls (~\$0.22) while super-off stays
near \$0.12, so the true off/SO ratio drops to ~1.81 — below the flat 2 the form
pays. Converting one off-peak credit then yields 2 × \$0.122 ≈ \$0.245 of super-off
offset versus the ~\$0.22 it would save on off-peak heating: a ~\$0.025/kWh edge.
Across the banked off-peak surplus this is worth roughly **\$50–55/yr** on the
reference profile (~1.5% of the bill). The peak rows never cross 4, so peak credit
is always worth more left on peak — **never convert peak to super-off**.

**The routine.**

1. **Hoard through the summer PSC season (≈ Jun–Sep).** File no off→super-off
   transfers; let net export bank in the off-peak (and peak) buckets. The off/SO
   ratio is above 2 then, so converting would lose value, and summer super-off is
   the cheapest of the year.
2. **Convert once the winter PSC season starts (≈ October).** File an Excess
   Generation Exchange transfer to move the banked off-peak surplus into super-off,
   where the fixed 1:2 ratio now overpays. The credit you transfer is the *summer
   surplus that is actually sitting in the bank* — no need to make anything survive
   to late winter.
3. **Resume hoarding when summer returns.** Stop converting once the off/SO ratio
   climbs back above 2.

**Two honest caveats.**

- *Timing within the conversion window doesn't add value.* Because the reachable
  super-off rate is essentially flat (~\$0.12 from summer through December — the
  off-peak bank is emptied by heating load before the high-rate February cycles),
  converting in summer versus the first winter month lands on the same rate. The
  gain comes entirely from the **destination choice** (super-off instead of
  off-peak heating) during the winter ratio window, not from *when* you file. There
  is no extra prize from trying to reach the high-PSC late-winter super-off cycles —
  no banked credit survives to them.
- *It is profile- and effort-dependent.* The edge requires a summer off-peak
  surplus to convert; a net off-peak consumer has nothing to steer. And it means
  filing transfers every winter for a few tens of dollars, so the default banking
  is a perfectly reasonable do-nothing baseline. There is no downside to the
  hoarding itself — residential credits roll forward 20 years.

The simulation is most meaningful over a **full year** of read dates, because the
carry-forward is the whole point: a high-solar summer builds the banks the winter
conversion works on. With only winter read dates the banks start empty in November
and the picture is understated.

---

## Output workbook

The workbook contains three tabs.

### `Dashboard`
The main view. For each billing cycle, plus a **TOTAL** column:

- **Usage** — days, net consumption, the time-of-use buckets (super-off-peak
  10 p.m.–6 a.m., off-peak, on-peak weekday 3–7 p.m.), and peak demand.
- **Billable after carry-forward banking** — the kWh each plan is actually billed
  on once net-export credits are banked and carried forward (a single volumetric
  bank for 580/180, the per-period banks for 194/195). Every per-rate bill below
  is computed from these quantities, so all four plans are compared on the same
  realistic basis and the comparison can't favor a Time-of-Day rate that real
  banking would make more expensive.
- **Per-rate bill breakdown** — for each of Rates 580, 180, 194, and 195, every
  bill line item: Basic Service, Customer Benefit Contribution, Merchant
  Function Charge, Delivery (tiered or time-of-use), Delivery & System subtotal,
  Power Supply, DER Charge, Delivery Service Adjustment, Revenue Decoupling
  Adjustment, NY State Assessment, Revenue-Based PILOTs, Property Tax
  Adjustment, Sales Tax, and Total Charges.
- **Rate comparison** — total bill per plan side by side, with a
  **Lowest-cost plan** row that flags the cheapest rate for each cycle.

All bill figures are **live Excel formulas** that reference the other two tabs,
so editing a rate reflows the whole dashboard. The billable-after-banking
quantities are computed in Python (banking depends on kWh and fixed value ratios,
not on the dollar rates) and feed those formulas as inputs, so price edits stay
fully live while the comparison remains consistent.

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
  riders (PILOTs, property tax, NYSA) are modeled as steady percentages rather
  than their exact monthly values.
- **Demand for the CBC.** Pass `--demand` to match the kW on your bill; otherwise
  the CBC line is based on the metered peak, which is usually higher.
- **Service area.** Use `--suffolk` / `--nassau` / `--rockaways` to set the local
  sales tax (or `--sales-tax` to override). The supply/delivery schedules and the
  PILOT/property-tax recovery are LIPA-wide and do not vary by county; only the
  sales-tax line does. The Suffolk rate is bill-verified; Nassau (0%) is verified
  from the state exemption; the Rockaways/NYC rate is an estimate worth confirming
  against a local bill.
- **Recalculation.** Opening the workbook in Excel or LibreOffice calculates the
  formulas automatically. If you build pipelines around it, note that openpyxl
  writes formulas as text; a spreadsheet application (or a headless LibreOffice
  recalculation) is needed to populate the computed values.

---

## Example session

```bash
$ python pseg_analyzer.py --cons_file Usage.csv \
      --read-dates 2025-11-15 2025-12-15 2026-01-16 2026-02-13 2026-03-16 2026-04-16 2026-05-15 \
      --demand 16

Loading consumption data...
  35,036 intervals, 2025-05-30 to 2026-05-29
Loading & refreshing rate schedule...
  Rate schedule already covers the data period; no scrape needed.
  Service area: Suffolk -- local sales tax 2.50% (bill-verified).
  6 billing cycle(s).

  Time-of-Day banking: 194 $4,109.66 vs 195 $3,184.75  ->  Rate 195 is the cheaper TOD plan.
    Rate 194: banked $4,109.66  (optimum $4,109.42, gap $0.24); profile PEAK +411, OFF +15,232 -> net consumer
    Rate 195: banked $3,184.75  (optimum $3,184.35, gap $0.40); profile PEAK +411, OFF +2,506, SO +12,726 -> net consumer  <- recommended
Wrote Usage_analysis.xlsx
```

Open `Usage_analysis.xlsx` and start on the **Dashboard** tab.
