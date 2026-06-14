# RADET Dashboard Aggregation Engine

A modular, YAML-driven aggregation engine that reads RADET/HTS Excel workbooks and outputs a normalized dashboard dataset compatible with React, Recharts, AG Grid, Power BI, Apache ECharts, and Chart.js.

---

## Table of Contents

1. [Quick Start](#quick-start)
2. [Folder Structure](#folder-structure)
3. [How It Works](#how-it-works)
4. [Configuration](#configuration)
   - [indicators.yaml](#indicatorsyaml)
   - [ageBands.yaml](#agebandsyaml)
   - [mappings.yaml](#mappingsyaml)
5. [Date Modes](#date-modes)
6. [Adding a New Indicator](#adding-a-new-indicator)
7. [Adding a New Worksheet](#adding-a-new-worksheet)
8. [Output Schema](#output-schema)
9. [Environment Variables](#environment-variables)
10. [Scripts Reference](#scripts-reference)
11. [Architecture](#architecture)
12. [Troubleshooting](#troubleshooting)

---

## Quick Start

### Prerequisites

- **Node.js** v18 or later — [download here](https://nodejs.org)
- **npm** v9 or later (bundled with Node.js)

> **No virtual environment needed.** This is a Node.js project. All dependencies are installed locally into `node_modules/` by npm — no Python venv, conda, or any other environment manager is required.

### 1. Install dependencies

Run this once after cloning the repo (or whenever `package.json` changes):

```bash
npm install
```Remove-Item Env:RADET_FILE

### 2. Place your workbook

Copy your `RADET.xlsx` into the `input/` folder:

```
input/RADET.xlsx
```

**Don't have a real file yet?** Generate sample data for testing:

```bash
npx ts-node scripts/generate-sample.ts
```

This creates `input/RADET.xlsx` with 500 RADET rows and 300 HTS rows.

### 3. Run the engine

**Development** (recommended — runs TypeScript directly, no build step):

```bash
npm run dev
```

**Production** (compile first, then run the compiled output):

```bash
# PowerShell (Windows)
npm run build; npm start

# Bash / Git Bash / Linux / macOS
npm run build && npm start
```

**Or run directly with ts-node:**

```bash
npx ts-node index.ts
```

> **Note (Windows PowerShell):** `&&` is not supported in Windows PowerShell 5.1. Use `;` to chain commands instead.

### 4. Check outputs

All outputs are written to the `outputs/` folder:

| File | Description |
|------|-------------|
| `DashboardSummary.xlsx` | Main dashboard (normalized rows + indicator pivot sheet) |
| `DashboardSummary.csv` | Same data as CSV for Power BI / Python |
| `ValidationReport.xlsx` | Any data quality issues found |
| `process.log` | Full processing log with timings |

---

## Folder Structure

```
radet-dashboard/
│
├── config/
│   ├── indicators.yaml      ← Define all indicators here
│   ├── ageBands.yaml        ← PEPFAR MER age band definitions
│   └── mappings.yaml        ← Value standardisation maps
│
├── services/
│   ├── AggregationEngine.ts ← Main orchestrator
│   ├── FilterEngine.ts      ← Row filtering logic
│   ├── GroupEngine.ts       ← Grouping & aggregation
│   ├── FormulaEngine.ts     ← Calculated indicators
│   └── OutputWriter.ts      ← Excel & CSV output
│
├── helpers/
│   ├── DateHelper.ts        ← Date parsing & range logic
│   ├── AgeBandHelper.ts     ← PEPFAR age band assignment
│   └── MappingHelper.ts     ← Value mapping
│
├── scripts/
│   └── generate-sample.ts  ← Test data generator
│
├── input/                   ← Place RADET.xlsx here
├── outputs/                 ← Generated files appear here
├── index.ts                 ← Entry point
├── package.json
└── tsconfig.json
```

---

## How It Works

The engine follows this processing pipeline:

```
Load Workbook
     ↓
Load YAML Configurations
     ↓
For each indicator definition:
  ├── Find source worksheet
  ├── Read all rows (header row → column map)
  ├── Apply value mappings (e.g. "51" → "Health Facility")
  ├── Enrich each row with PEPFAR AgeBand
  ├── Apply filters (date range, value equals, contains, etc.)
  ├── Group rows by configured keys
  └── Aggregate (COUNT / SUM / AVG / etc.)
     ↓
Calculate formula indicators (LINKAGE_RATE, TB_TX_GAP, etc.)
     ↓
Write DashboardSummary.xlsx + .csv
     ↓
Write ValidationReport.xlsx (if issues found)
     ↓
Write process.log
```

---

## Configuration

### indicators.yaml

Each entry in the `indicators` array defines one indicator. No engine code changes are needed.

```yaml
indicators:
  - name: HTS_TST_POS                  # Unique indicator name
    description: "HIV positive testers"
    source: CombineHTS                  # Exact worksheet name in workbook
    requiredColumns:                    # Validated before processing
      - finalHIVTestResult
      - dateOfHIVTesting
      - Sex
      - Age
      - Facility
      - DATIMCode
    filters:                            # All filters must pass (AND logic)
      - column: finalHIVTestResult
        operator: equals
        value: "Positive"
      - column: dateOfHIVTesting
        operator: dateMode              # Uses current DATE_MODE
        value: DAILY
    groupBy:                            # Columns to disaggregate by
      - Facility
      - DATIMCode
      - State
      - LGA
      - Sex
      - AgeBand
    aggregation: COUNT                  # COUNT | SUM | AVG | MIN | MAX
    disaggregation: Sex                 # Sets the Disaggregation column
```

#### Filter Operators

| Operator | Description | Example value |
|----------|-------------|---------------|
| `equals` | Exact match | `"Positive"` |
| `notEquals` | Does not match | `"Transfer Out"` |
| `contains` | String contains | `"Active"` |
| `notContains` | String does not contain | `"transfer"` |
| `inList` | Value is in list | `["Presumptive TB", "TB suspect"]` |
| `notInList` | Value is not in list | `["Dead", "LTFU"]` |
| `dateMode` | Date falls in current period | `DAILY` |
| `greaterThan` | Numeric > value | `0` |
| `lessThan` | Numeric < value | `100` |
| `isNull` | Cell is empty | *(no value needed)* |
| `isNotNull` | Cell has a value | *(no value needed)* |

Add `caseSensitive: false` (default) or `caseSensitive: true` to string operators.

#### Aggregation Methods

| Method | Description |
|--------|-------------|
| `COUNT` | Count matching rows |
| `SUM` | Sum a numeric column |
| `AVG` | Average of a numeric column |
| `MIN` | Minimum value |
| `MAX` | Maximum value |
| `COUNTDISTINCT` | Count unique values in a column |

#### Formula Indicators

```yaml
formulaIndicators:
  - name: LINKAGE_RATE
    numerator: TX_NEW
    denominator: HTS_TST_POS
    formula: "numerator / denominator"   # or: - + *
    outputType: percentage               # percentage | count | ratio
    groupBy:
      - Facility
      - DATIMCode
      - State
      - LGA
```

---

### ageBands.yaml

Defines PEPFAR MER age bands. Modify or add bands without touching engine code.

```yaml
ageBands:
  - label: "<1"
    min: 0
    max: 0
  - label: "1-4"
    min: 1
    max: 4
  # ... etc.
  - label: "50+"
    min: 50
    max: 999
defaultBand: "Unknown"
```

---

### mappings.yaml

Maps raw data values to standardized forms. Useful for harmonizing inconsistent data entry.

```yaml
mappings:
  ARTEnrollmentSetting:
    "51": "Health Facility"
    "52": "Community"
    "Facility": "Health Facility"
    "COMMUNITY": "Community"

  Sex:
    "M": "Male"
    "F": "Female"
    "1": "Male"
    "2": "Female"
```

---

## Date Modes

Set the date mode using the `DATE_MODE` environment variable.

| Mode | Rows matched | Use case |
|------|-------------|----------|
| `TODAY` | Today's date only | Daily real-time dashboard |
| `DAILY` | Today's date only | Synonym for TODAY |
| `WEEKLY` | Current ISO week | Weekly reporting |
| `MONTHLY` | Current month | Monthly MER submission |
| `QUARTERLY` | Current quarter | PEPFAR quarterly targets |
| `YEARLY` | Current year | Annual COP targets |
| `CUSTOM` | Specified range | Ad-hoc analysis |

```bash
# Monthly run
DATE_MODE=MONTHLY npm run dev

# Custom range
DATE_MODE=CUSTOM CUSTOM_START=2026-01-01 CUSTOM_END=2026-03-31 npm run dev

# Use a reference date instead of today (useful for backfilling)
DATE_MODE=DAILY REFERENCE_DATE=2026-05-15 npm run dev
```

---

## Adding a New Indicator

**No engine code changes needed.** Edit `config/indicators.yaml`:

```yaml
indicators:
  # ... existing indicators ...

  - name: PMTCT_STAT
    description: "Pregnant women with known HIV status"
    source: CombinedPMTCT            # ← new sheet name
    requiredColumns:
      - HIVStatusKnown
      - AntenatalDate
      - Sex
      - Age
      - Facility
      - DATIMCode
      - State
      - LGA
    filters:
      - column: HIVStatusKnown
        operator: equals
        value: "Known"
      - column: AntenatalDate
        operator: dateMode
        value: MONTHLY
    groupBy:
      - Facility
      - DATIMCode
      - State
      - LGA
      - AgeBand
    aggregation: COUNT
    disaggregation: AgeBand
```

Save and re-run. Done.

---

## Adding a New Worksheet

1. Ensure the sheet exists in `RADET.xlsx` (the engine auto-discovers all sheets).
2. Add your indicator definition(s) in `indicators.yaml` pointing to `source: YourSheetName`.
3. Add any value mappings needed in `mappings.yaml`.
4. Run the engine.

---

## Output Schema

One row per unique combination of: `Facility × Indicator × Disaggregation × AgeBand × Sex × Period`

| Column | Type | Example |
|--------|------|---------|
| Period | string | `2026-06-13` / `2026-06` / `2026-Q2` |
| State | string | `Taraba` |
| LGA | string | `Yorro` |
| Facility | string | `General Hospital Wukari` |
| DATIMCode | string | `JPBcTpp6XUu` |
| Indicator | string | `HTS_TST_POS` |
| Disaggregation | string | `Sex` |
| Category | string | `Male` |
| Sex | string | `Male` |
| AgeBand | string | `25-29` |
| Value | number | `14` |
| Numerator | number | `14` |
| Denominator | number | `18` |
| Target | number | `20` |
| AchievementPct | string | `70%` |

---

## Environment Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `DATE_MODE` | `DAILY` | Date filter mode |
| `REFERENCE_DATE` | Today | Override today's date (`YYYY-MM-DD`) |
| `CUSTOM_START` | — | Start date for CUSTOM mode |
| `CUSTOM_END` | — | End date for CUSTOM mode |
| `RADET_FILE` | `input/RADET.xlsx` | Path to workbook |

---

## Scripts Reference

| Command | Description |
|---------|-------------|
| `npm run dev` | Run with ts-node (development) |
| `npm run build` | Compile TypeScript to `dist/` |
| `npm start` | Run compiled output |
| `npm run clean` | Remove `dist/` folder |
| `npx ts-node scripts/generate-sample.ts` | Generate sample RADET.xlsx |

---

## Architecture

The engine is built on **SOLID principles** with clean separation of concerns:

| Service | Responsibility |
|---------|---------------|
| `AggregationEngine` | Orchestrates the full pipeline; loads config, workbook, delegates to other services |
| `FilterEngine` | Stateless row filter; evaluates one `FilterDef` at a time |
| `GroupEngine` | Groups filtered rows by keys; computes COUNT/SUM/AVG/etc. |
| `FormulaEngine` | Post-processing; calculates derived indicators from existing dashboard rows |
| `OutputWriter` | Renders `DashboardSummary.xlsx`, `.csv`, and `ValidationReport.xlsx` |
| `DateHelper` | Parses dates (multiple formats); determines current date range |
| `AgeBandHelper` | Assigns PEPFAR age band from a numeric age |
| `MappingHelper` | Looks up standardised values from the mappings config |

Adding a new disaggregation (e.g. pregnant status) requires:
1. A column in the source sheet.
2. An entry in `groupBy` in `indicators.yaml`.
3. Optionally, a mapping in `mappings.yaml`.

---

## Troubleshooting

### "Workbook not found" error even though the file is in `input/`
The `RADET_FILE` environment variable from a previous session may be overriding the default path. Clear it:

```powershell
# PowerShell
Remove-Item Env:RADET_FILE
npm run dev
```

```bash
# Bash
unset RADET_FILE
npm run dev
```

### "Worksheet not found" warning
Check that the sheet name in `indicators.yaml → source` matches exactly (case-sensitive) the tab name in your Excel file.

### All indicators show 0 rows
Check your `DATE_MODE`. If you're testing with historical data, either:
- Set `DATE_MODE=MONTHLY` or `DATE_MODE=YEARLY`, or
- Use `REFERENCE_DATE=YYYY-MM-DD` to match a date in your data.

### Missing columns warning
The engine logs which columns are expected but absent. Either:
- Rename the column in your sheet to match the config, or
- Update `requiredColumns` and filter definitions in `indicators.yaml`.

### Formula indicators return 0
Ensure both the numerator and denominator indicators are producing rows. Check `process.log` for each indicator's row count.

---

## Extending to Power BI

Load `DashboardSummary.csv` directly:
1. **Get Data** → **Text/CSV** → select the file
2. Use **Indicator** column as a slicer/filter
3. Use **Value** as your measure
4. Pivot on **AgeBand** or **Sex** for disaggregation charts

No additional transformation is needed — the schema is already normalized.
