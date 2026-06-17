import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import yaml from 'js-yaml';
import { Logger } from 'winston';

import { FilterEngine, FilterDef } from './FilterEngine';
import { AggregationMethod } from './GroupEngine';
import { FormulaEngine, FormulaIndicatorDef } from './FormulaEngine';
import { AgeBandHelper } from '../helpers/AgeBandHelper';
import { MappingHelper } from '../helpers/MappingHelper';
import { OrgUnitHelper } from '../helpers/OrgUnitHelper';
import { DateHelper, DateModeConfig } from '../helpers/DateHelper';

// ─── Public Types ────────────────────────────────────────────────────────────

export interface DashboardRow {
  Period: string;
  State: string;
  Facility: string;
  DATIMCode: string;
  Indicator: string;
  Disaggregation: string;
  Category: string;
  Sex: string;
  AgeBand: string;
  Value: number | null;
  Numerator: number | null;
  Denominator: number | null;
  Target: number | null;
  AchievementPct: number | null;
}

export interface ValidationIssue {
  sheet: string;
  row: number | string;
  column: string;
  severity: 'ERROR' | 'WARNING' | 'INFO';
  message: string;
}

export interface ComputedColumnDef {
  name: string;
  formula: 'dateAdd';
  sourceColumn: string;
  addColumn: string;
  unit: 'day' | 'week' | 'month' | 'year';
}

export interface IndicatorDef {
  name: string;
  description?: string;
  source: string;
  requiredColumns: string[];
  computedColumns?: ComputedColumnDef[];
  filters: FilterDef[];
  groupBy: string[];
  aggregation: AggregationMethod;
  disaggregation?: string;
  valueColumn?: string;
  /** Column whose date value becomes the Period label in group-by-date modes. */
  periodColumn?: string;
}

export interface IndicatorsConfig {
  indicators: IndicatorDef[];
  formulaIndicators: FormulaIndicatorDef[];
}

export interface AggregationEngineOptions {
  workbookPath: string;
  configDir: string;
  dateModeConfig: DateModeConfig;
  logger: Logger;
}

export interface ProcessingStats {
  totalRows: number;
  filteredRows: number;
  aggregatedRows: number;
  indicatorsProcessed: string[];
  warnings: string[];
  durationMs: number;
}

// ─── Internal types ───────────────────────────────────────────────────────────

interface StreamAcc {
  count: number;
  sum: number;
  min: number;
  max: number;
  distinctSet: Set<string>;
}

/** Raw group entry: stores suffix-normalised field values before dynamic mapping */
interface RawGroupEntry {
  rawComponents: Record<string, string>;
  acc: StreamAcc;
}

// ─── Engine ──────────────────────────────────────────────────────────────────

export class AggregationEngine {
  private options: AggregationEngineOptions;
  private logger: Logger;

  private filterEngine!: FilterEngine;
  private formulaEngine!: FormulaEngine;
  private ageBandHelper!: AgeBandHelper;
  private mappingHelper!: MappingHelper;
  private orgUnitHelper!: OrgUnitHelper;
  private dateHelper!: DateHelper;

  private indicatorsConfig!: IndicatorsConfig;
  private sheetHeaders!: Record<string, string[]>;
  private skipFirstRowSheets: Set<string> = new Set();
  /** Maps logical source name (used in indicators.yaml) → physical sheet name in the workbook */
  private sheetAliases: Record<string, string> = {};
  private categoryRanksConfig!: { categoryRanks: Record<string, string[]> };

  // Built from Sheet1 during streaming; used to enrich HTS rows.
  // Persisted across runs via datim-cache.json so coverage grows over time.
  private datimToFacility = new Map<string, { Facility: string; State: string }>();
  private datimCachePath = '';

  public validationIssues: ValidationIssue[] = [];
  public stats: ProcessingStats = {
    totalRows: 0,
    filteredRows: 0,
    aggregatedRows: 0,
    indicatorsProcessed: [],
    warnings: [],
    durationMs: 0,
  };

  constructor(options: AggregationEngineOptions) {
    this.options = options;
    this.logger = options.logger;
  }

  // ─── Initialization ─────────────────────────────────────────────────────

  async init(): Promise<void> {
    this.logger.info('AggregationEngine: initialising…');

    this.ageBandHelper = new AgeBandHelper(path.join(this.options.configDir, 'ageBands.yaml'));
    this.mappingHelper = new MappingHelper(path.join(this.options.configDir, 'mappings.yaml'));
    this.dateHelper    = new DateHelper(this.options.dateModeConfig);

    this.filterEngine  = new FilterEngine({ dateModeConfig: this.options.dateModeConfig, logger: this.logger });
    this.formulaEngine = new FormulaEngine({ logger: this.logger });

    const indicatorsYaml = fs.readFileSync(path.join(this.options.configDir, 'indicators.yaml'), 'utf8');
    this.indicatorsConfig = yaml.load(indicatorsYaml) as IndicatorsConfig;
    this.logger.info(
      `AggregationEngine: loaded ${this.indicatorsConfig.indicators.length} indicators, ` +
      `${this.indicatorsConfig.formulaIndicators?.length ?? 0} formula indicators`,
    );

    const headersYaml = fs.readFileSync(path.join(this.options.configDir, 'sheetHeaders.yaml'), 'utf8');
    const headersRaw = yaml.load(headersYaml) as Record<string, unknown>;
    // Extract sheetAliases first, then process the rest as sheet headers
    this.sheetAliases = {};
    if (headersRaw['sheetAliases'] && typeof headersRaw['sheetAliases'] === 'object') {
      this.sheetAliases = headersRaw['sheetAliases'] as Record<string, string>;
      this.logger.info(`AggregationEngine: sheet aliases: ${JSON.stringify(this.sheetAliases)}`);
    }
    // Normalise: support both plain array and {columns, skipFirstRow} object forms
    this.sheetHeaders = {};
    for (const [sheet, val] of Object.entries(headersRaw)) {
      if (sheet === 'sheetAliases') continue;
      if (Array.isArray(val)) {
        this.sheetHeaders[sheet] = val as string[];
      } else if (val && typeof val === 'object' && 'columns' in val) {
        this.sheetHeaders[sheet] = (val as { columns: string[] }).columns;
        if ((val as { skipFirstRow?: boolean }).skipFirstRow) {
          this.skipFirstRowSheets.add(sheet);
        }
      }
    }
    this.logger.info(
      `AggregationEngine: loaded positional headers for: [${Object.keys(this.sheetHeaders).join(', ')}]`,
    );

    const orgUnitsPath = path.join(this.options.configDir, 'orgUnits.yaml');
    this.orgUnitHelper = new OrgUnitHelper(orgUnitsPath);
    if (this.orgUnitHelper.hasAnyMappings) {
      this.logger.info('AggregationEngine: loaded org-unit mappings from orgUnits.yaml');
    }

    this.datimCachePath = path.join(this.options.configDir, 'datim-cache.json');
    this.loadDatimCache();

    const ranksPath = path.join(this.options.configDir, 'categoryRanks.yaml');
    if (fs.existsSync(ranksPath)) {
      this.categoryRanksConfig = yaml.load(fs.readFileSync(ranksPath, 'utf8')) as {
        categoryRanks: Record<string, string[]>;
      };
    } else {
      this.categoryRanksConfig = { categoryRanks: {} };
      this.logger.warn('AggregationEngine: categoryRanks.yaml not found — no dynamic mapping');
    }
  }

  // ─── Group-by-date helpers ───────────────────────────────────────────────

  private isGroupByDateMode(): boolean {
    const m = this.options.dateModeConfig.mode;
    return m === 'DAILY' || m === 'WEEKLY' || m === 'MONTHLY' || m === 'QUARTERLY';
  }

  private formatRowPeriod(rawDate: unknown): string {
    const d = this.dateHelper.parse(rawDate);
    if (!d) return 'Unknown';
    const m = this.options.dateModeConfig.mode;
    if (m === 'DAILY')     return d.format('DD/MM/YYYY');
    if (m === 'WEEKLY')    return `W${String(d.week()).padStart(2, '0')}/${d.format('YYYY')}`;
    if (m === 'MONTHLY')   return d.format('MMM YYYY');
    if (m === 'QUARTERLY') return `Q${d.quarter()} ${d.format('YYYY')}`;
    return d.format('YYYY');
  }

  // ─── Compound UID helpers ────────────────────────────────────────────────

  /**
   * DHIS2 RADET exports qualify categorical values as "{facilityDATIM}_{optionUID}"
   * where the facility DATIM code is always exactly 11 alphanumeric chars.
   * Returns the 11-char DATIM prefix if the value matches this format, else null.
   */
  private extractDatimFromCompound(val: string): string | null {
    if (!val || val.length < 13) return null;
    const idx = val.indexOf('_');
    if (idx !== 11) return null;
    const prefix = val.slice(0, 11);
    return /^[A-Za-z][A-Za-z0-9]{10}$/.test(prefix) ? prefix : null;
  }

  /**
   * Find the facility DATIM code by scanning multiple columns of a row.
   *
   * For RADET: scans all columns except LGA and LGAOfResidence (those can embed
   * OTHER facilities' DATIM codes via compound values and would cause false matches).
   * DATIMCode (col5) is checked first as it is the authoritative source; all other
   * columns are scanned as fallbacks to handle ExcelJS shared-string cache misses
   * that can return unexpected values for any given column on any given run.
   *
   * For other sheets (HTS): scans all columns.
   *
   * Each candidate is validated against orgUnits.yaml so false positives are
   * effectively impossible (only 74 known codes).
   */
  private findFacilityDatim(record: Record<string, unknown>, sheetName: string): string | null {
    const tryVal = (str: string): string | null => {
      const datim = this.extractDatimFromCompound(str);
      if (datim && this.orgUnitHelper.lookupByDATIM(datim)) return datim;
      if (str.length === 11 && /^[A-Za-z][A-Za-z0-9]{10}$/.test(str)) {
        if (this.orgUnitHelper.lookupByDATIM(str)) return str;
      }
      return null;
    };

    if (sheetName === 'Sheet1') {
      // LGA columns are excluded: they embed compound UIDs from OTHER facilities
      // (e.g. the LGA a patient transferred from) and would produce wrong matches.
      const skipCols = new Set(['LGA', 'LGAOfResidence']);

      // Check col5 (DATIMCode) first — it is the authoritative facility identifier.
      const col5 = tryVal(String(record['DATIMCode'] ?? '').trim());
      if (col5) return col5;

      // Scan all remaining columns as fallbacks (handles ExcelJS shared-string
      // cache misses where col5 resolves to a non-DATIM value on a given run).
      for (const [col, val] of Object.entries(record)) {
        if (col === 'DATIMCode' || skipCols.has(col)) continue;
        const found = tryVal(String(val ?? '').trim());
        if (found) return found;
      }
    } else {
      for (const rawVal of Object.values(record)) {
        const found = tryVal(String(rawVal ?? '').trim());
        if (found) return found;
      }
    }
    return null;
  }

  /**
   * Normalise a categorical option value so that values from different facilities
   * that reference the SAME DHIS2 option compare equal.
   *
   * For compound values "{facilityDATIM}_{optionUID}", returns just the optionUID
   * suffix.  Plain values are returned unchanged.
   */
  private normalizeOptionValue(val: string): string {
    const datim = this.extractDatimFromCompound(val);
    return datim ? val.slice(12) : val;  // 11 DATIM chars + 1 underscore = skip 12
  }

  // ─── Main Entry Point ────────────────────────────────────────────────────

  async process(): Promise<DashboardRow[]> {
    const startTime = Date.now();
    const globalPeriod  = this.dateHelper.getPeriodLabel();
    const groupByDate   = this.isGroupByDateMode();
    const groupDateRange = groupByDate ? this.dateHelper.getGroupByDateRange() : null;

    if (groupDateRange) {
      this.logger.info(
        `AggregationEngine: group-by-date range → ` +
        `${groupDateRange.start.format('DD/MM/YYYY')} to ${groupDateRange.end.format('DD/MM/YYYY')}`,
      );
    }

    // ── Setup ────────────────────────────────────────────────────────────────

    const targetCols = new Set(Object.keys(this.categoryRanksConfig.categoryRanks));

    // frequency map: normalised value → count (suffix-based so all facilities share)
    const freq = new Map<string, Map<string, number>>();
    for (const col of targetCols) freq.set(col, new Map());

    // Group indicators by source sheet (resolve aliases: logical → physical)
    const bySheet = new Map<string, IndicatorDef[]>();
    // Track logical→physical mapping for warning messages
    const physicalToLogical = new Map<string, string>();
    for (const ind of this.indicatorsConfig.indicators) {
      const physical = this.sheetAliases[ind.source] ?? ind.source;
      if (!bySheet.has(physical)) bySheet.set(physical, []);
      bySheet.get(physical)!.push(ind);
      physicalToLogical.set(physical, ind.source);
    }

    // Determine all fields needed per indicator (groupBy + categorical filter cols)
    const indFilterCols = new Map<string, Set<string>>();
    for (const ind of this.indicatorsConfig.indicators) {
      const catFilters = new Set(
        ind.filters
          .filter(f => f.operator !== 'dateMode')
          .map(f => f.column),
      );
      indFilterCols.set(ind.name, catFilters);
    }

    // rawAccMap: indName → rawKeyStr → RawGroupEntry
    const rawAccMap = new Map<string, Map<string, RawGroupEntry>>();
    for (const ind of this.indicatorsConfig.indicators) rawAccMap.set(ind.name, new Map());

    // Org-unit raw data: col1-prefix → { stateRaw, lgaRaw, facilityRaw }
    const orgUnitSeen = new Map<string, { stateRaw: string; lgaRaw: string; facilityRaw: string }>();

    const seenSheets = new Set<string>();

    // ── SINGLE STREAMING PASS ─────────────────────────────────────────────────

    this.logger.info('AggregationEngine: single streaming pass (calibration + processing)…');

    const reader = new ExcelJS.stream.xlsx.WorkbookReader(
      this.options.workbookPath,
      { sharedStrings: 'cache', styles: 'ignore', hyperlinks: 'ignore', worksheets: 'emit' },
    );
    reader.read();

    for await (const worksheetReader of reader) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const sheetName: string = (worksheetReader as any).name;
      seenSheets.add(sheetName);

      const indicators = bySheet.get(sheetName);
      if (!indicators) {
        for await (const _ of worksheetReader) { /* drain */ }
        continue;
      }

      const predefined = this.sheetHeaders[sheetName] as string[] | undefined;
      const headers: string[] = predefined ? ['', ...predefined] : [];
      // isFirstRow: true when we need to read headers from row 1 (no predefined),
      // OR when the sheet has predefined headers but also a header row to skip.
      let isFirstRow = !predefined || this.skipFirstRowSheets.has(sheetName);

      this.logger.info(
        `  Streaming sheet: ${sheetName} (${indicators.map(i => i.name).join(', ')})` +
        (predefined ? ` [${predefined.length} positional cols]` : ' [first-row headers]'),
      );

      const skipRow1WithPredefined = this.skipFirstRowSheets.has(sheetName);

      for await (const row of worksheetReader) {
        if (isFirstRow) {
          isFirstRow = false;
          if (!skipRow1WithPredefined) {
            // No predefined headers: read column names from this row
            (row as ExcelJS.Row).eachCell((cell, col) => {
              headers[col] = String(cell.value ?? '').trim();
            });
          }
          // skipRow1WithPredefined: predefined headers are already set — just skip this row
          continue;
        }

        // ── Parse row ────────────────────────────────────────────────────────
        const record: Record<string, unknown> = {};
        (row as ExcelJS.Row).eachCell({ includeEmpty: true }, (cell, col) => {
          const header = headers[col];
          if (header) {
            record[header] =
              cell.value instanceof Object && 'result' in (cell.value as object)
                ? (cell.value as ExcelJS.CellFormulaValue).result
                : cell.value;
          }
        });

        this.stats.totalRows++;

        // ── Apply computed columns for all indicators on this sheet ───────────
        for (const ind of indicators) {
          if (ind.computedColumns?.length) {
            this.applyComputedColumns(record, ind.computedColumns);
          }
        }

        // ── Build frequency maps (suffix-normalised) ─────────────────────────
        for (const col of targetCols) {
          if (!(col in record)) continue;
          const raw = String(record[col] ?? '').trim();
          if (!raw) continue;
          const norm = this.normalizeOptionValue(raw);
          const colFreq = freq.get(col)!;
          colFreq.set(norm, (colFreq.get(norm) ?? 0) + 1);
        }

        // ── Extract facility DATIM (scans col1/col4/col5 for RADET, all cols for HTS) ─
        const rowDatim = this.findFacilityDatim(record, sheetName);

        // Populate RADET cross-sheet lookup (used for HTS facility info)
        if (sheetName === 'Sheet1' && rowDatim && !this.datimToFacility.has(rowDatim)) {
          const ou = this.orgUnitHelper.lookupByDATIM(rowDatim);
          this.datimToFacility.set(rowDatim, {
            Facility: ou?.facility ?? '',
            State:    ou?.state    ?? '',
          });
        }

        // Collect raw org-unit data for the CSV diagnostic output
        if (sheetName === 'Sheet1' && rowDatim && !orgUnitSeen.has(rowDatim)) {
          orgUnitSeen.set(rowDatim, {
            stateRaw:    String(record['State']    ?? '').trim(),
            lgaRaw:      String(record['LGA']      ?? '').trim(),
            facilityRaw: String(record['Facility'] ?? '').trim(),
          });
        }

        // ── Per-indicator accumulation ─────────────────────────────────────────
        for (const ind of indicators) {
          // ── Date-range gate (replaces dateMode filter in group-by-date modes) ──
          if (groupDateRange && ind.periodColumn) {
            const rowDate = this.dateHelper.parse(record[ind.periodColumn]);
            if (!rowDate || !rowDate.isBetween(groupDateRange.start, groupDateRange.end, 'day', '[]')) {
              continue;
            }
          } else if (!groupByDate) {
            // Standard (non-group-by-date) mode: apply dateMode filter directly
            const dmFilter = ind.filters.find(f => f.operator === 'dateMode');
            if (dmFilter && !this.dateHelper.isInRange(record[dmFilter.column])) continue;
          }

          // ── Row period label ──────────────────────────────────────────────────
          const rowPeriod = groupByDate && ind.periodColumn
            ? this.formatRowPeriod(record[ind.periodColumn])
            : globalPeriod;

          // ── Build raw (pre-mapping) group key components ──────────────────────
          // Includes all groupBy fields + AgeBand + Sex + categorical filter cols.
          // Values are suffix-normalised so that the same DHIS2 option has the
          // same key regardless of which facility reported it.
          const catFilterCols = indFilterCols.get(ind.name)!;
          const allTrackCols  = new Set([...ind.groupBy, 'AgeBand', 'Sex', ...catFilterCols]);

          const rawComponents: Record<string, string> = {
            __period__: rowPeriod,
            __datim__:  rowDatim ?? '',
            // Non-RADET sheets (e.g. HTS): the DATIMCode column IS the facility
            // identifier.  Store the raw value so post-processing can display it
            // even when it cannot be resolved to an orgUnit name.
            __rawDatimCode__: sheetName !== 'Sheet1'
              ? this.normalizeOptionValue(String(record['DATIMCode'] ?? '').trim())
              : '',
          };

          for (const field of allTrackCols) {
            if (field === 'AgeBand') {
              rawComponents['AgeBand'] = this.ageBandHelper.getBand(record['Age']);
            } else {
              const raw  = String(record[field] ?? '').trim();
              const norm = this.normalizeOptionValue(raw);
              rawComponents[field] = norm;
            }
          }

          // __datim__ is already set from findFacilityDatim() above (valid for all sheets)

          const rawKeyStr = JSON.stringify(rawComponents);
          const indRaw = rawAccMap.get(ind.name)!;

          if (!indRaw.has(rawKeyStr)) {
            indRaw.set(rawKeyStr, {
              rawComponents,
              acc: { count: 0, sum: 0, min: Infinity, max: -Infinity, distinctSet: new Set() },
            });
          }

          const entry = indRaw.get(rawKeyStr)!;
          entry.acc.count++;

          if (ind.valueColumn) {
            const v = record[ind.valueColumn];
            const n = typeof v === 'number' ? v : parseFloat(String(v));
            if (!isNaN(n)) {
              entry.acc.sum += n;
              if (n < entry.acc.min) entry.acc.min = n;
              if (n > entry.acc.max) entry.acc.max = n;
              entry.acc.distinctSet.add(String(v));
            }
          }
        }
      }
    }

    // ── Build dynamic mappings from frequency data (same stream = same UIDs) ──

    const dynamicMappings = new Map<string, Map<string, string>>();
    for (const [col, labels] of Object.entries(this.categoryRanksConfig.categoryRanks)) {
      const colFreq = freq.get(col);
      if (!colFreq || colFreq.size === 0) {
        this.logger.warn(`  Calibration: no values found for column '${col}'`);
        continue;
      }
      const sorted = [...colFreq.entries()].sort((a, b) => b[1] - a[1]);
      const colMap = new Map<string, string>();
      for (let i = 0; i < Math.min(labels.length, sorted.length); i++) {
        const [normVal, count] = sorted[i];
        colMap.set(normVal, labels[i]);
        this.logger.info(
          `  [${col}] rank ${i + 1}: "${normVal}" (n=${count.toLocaleString()}) → "${labels[i]}"`,
        );
      }
      dynamicMappings.set(col, colMap);
    }

    // Write diagnostic CSV so operators can verify / fill in missing org-unit names
    this.writeOrgUnitsRawCsv(orgUnitSeen);
    this.saveDatimCache();

    // ── Post-process: map raw keys → labels, apply categorical filters ────────

    this.logger.info('AggregationEngine: post-processing accumulators…');

    for (const [sheetName] of bySheet) {
      if (!seenSheets.has(sheetName)) {
        const msg = `Worksheet '${sheetName}' not found in workbook`;
        this.logger.warn(msg);
        this.addValidationIssue(sheetName, 'N/A', 'Sheet', 'WARNING', msg);
      }
    }

    const allRows: DashboardRow[] = [];

    for (const ind of this.indicatorsConfig.indicators) {
      this.logger.info(`\n── Finalising indicator: ${ind.name} ──`);
      const indRaw = rawAccMap.get(ind.name)!;

      if (indRaw.size === 0) {
        this.logger.warn(`  No rows accumulated for ${ind.name} (all filtered by date range)`);
        continue;
      }

      // Categorical filters that apply to mapped values (exclude dateMode)
      const catFilters = ind.filters.filter(f => f.operator !== 'dateMode');

      // Final accumulator after mapping + re-grouping
      const finalAcc  = new Map<string, StreamAcc>();
      const finalComps = new Map<string, Record<string, string>>();

      let keptGroups = 0;

      for (const [, { rawComponents, acc }] of indRaw) {
        // ── Map each raw component to a human-readable label ─────────────────
        const mapped: Record<string, string> = { __period__: rawComponents['__period__'] };

        for (const [field, rawNorm] of Object.entries(rawComponents)) {
          if (field.startsWith('__')) continue;

          // Dynamic mapping (calibrated in this same streaming pass)
          if (dynamicMappings.has(field)) {
            const label = dynamicMappings.get(field)!.get(rawNorm);
            if (label !== undefined) { mapped[field] = label; continue; }
          }
          // Static mappings.yaml fallback
          if (this.mappingHelper.hasMapping(field)) {
            mapped[field] = String(this.mappingHelper.map(field, rawNorm));
          } else {
            mapped[field] = rawNorm;
          }
        }

        // ── Resolve org-unit names from facility DATIM code ────────────────────
        const datim = rawComponents['__datim__'];
        if (datim) {
          const ou = this.orgUnitHelper.lookupByDATIM(datim);
          if (ou) {
            if (ou.facility) mapped['Facility'] = ou.facility;
            if (ou.state)    mapped['State']    = ou.state;
            if (ou.lga)      mapped['LGA']      = ou.lga;
          } else {
            // Fallback: look in the RADET cross-sheet map
            const rInfo = this.datimToFacility.get(datim);
            if (rInfo) {
              mapped['Facility'] = rInfo.Facility;
              mapped['State']    = rInfo.State;
            }
          }
          mapped['DATIMCode'] = datim;
        } else {
          // No valid facility DATIM resolved via orgUnits.yaml.
          mapped['State']    = '';
          mapped['Facility'] = '';
          mapped['LGA']      = '';
          // For non-RADET sheets the DATIMCode column is the facility identifier
          // (not a patient ID), so show it as-is.  For RADET, col5 is a patient
          // enrollment number — hide it.
          const rawDatimCode = rawComponents['__rawDatimCode__'];
          mapped['DATIMCode'] = rawDatimCode || '';
          // If the raw DATIMCode resolves via the RADET cross-sheet map, pick up names
          if (rawDatimCode) {
            const rInfo = this.datimToFacility.get(rawDatimCode);
            if (rInfo) {
              mapped['Facility'] = rInfo.Facility;
              mapped['State']    = rInfo.State;
            }
          }
        }

        // ── Apply categorical filters on mapped values ─────────────────────────
        if (catFilters.length > 0) {
          const passedStr: Record<string, unknown> = {};
          for (const [k, v] of Object.entries(mapped)) passedStr[k] = v;
          if (!this.filterEngine.passesAll(passedStr, catFilters)) continue;
        }

        keptGroups++;
        this.stats.filteredRows += acc.count;

        // ── Re-group by final (mapped) dimensions only ─────────────────────────
        const finalGroupFields = [...new Set([...ind.groupBy, 'AgeBand', 'Sex'])];
        const finalKeyComps: Record<string, string> = { __period__: mapped['__period__'] };
        for (const f of finalGroupFields) finalKeyComps[f] = mapped[f] ?? '';
        const finalKeyStr = JSON.stringify(finalKeyComps);

        if (!finalAcc.has(finalKeyStr)) {
          finalAcc.set(finalKeyStr, { count: 0, sum: 0, min: Infinity, max: -Infinity, distinctSet: new Set() });
          finalComps.set(finalKeyStr, finalKeyComps);
        }

        const fa = finalAcc.get(finalKeyStr)!;
        fa.count += acc.count;
        fa.sum   += acc.sum;
        if (acc.min < fa.min) fa.min = acc.min;
        if (acc.max > fa.max) fa.max = acc.max;
        acc.distinctSet.forEach(v => fa.distinctSet.add(v));
      }

      if (finalAcc.size === 0) {
        this.logger.warn(`  No rows passed filters for ${ind.name}`);
        continue;
      }

      // ── Emit DashboardRows ─────────────────────────────────────────────────
      let rowCount = 0;
      for (const [keyStr, acc] of finalAcc) {
        const kc    = finalComps.get(keyStr)!;
        const value = this.resolveAccValue(acc, ind);
        allRows.push({
          Period:        kc['__period__'] ?? globalPeriod,
          State:         kc['State']      ?? '',
          Facility:      kc['Facility']   ?? '',
          DATIMCode:     kc['DATIMCode']  ?? '',
          Indicator:     ind.name,
          Disaggregation: ind.disaggregation ?? 'Total',
          Category:      kc[ind.disaggregation ?? ''] ?? 'Total',
          Sex:           kc['Sex']        ?? '',
          AgeBand:       kc['AgeBand']    ?? '',
          Value:         value,
          Numerator:     value,
          Denominator:   null,
          Target:        null,
          AchievementPct: null,
        });
        rowCount++;
      }

      this.stats.aggregatedRows += rowCount;
      this.stats.indicatorsProcessed.push(ind.name);
      this.logger.info(`  Generated ${rowCount} dashboard rows for ${ind.name}`);
    }

    // ── Formula indicators ────────────────────────────────────────────────────
    if (this.indicatorsConfig.formulaIndicators?.length > 0) {
      this.logger.info('\n── Processing formula indicators ──');
      const formulaRows = this.formulaEngine.calculate(
        this.indicatorsConfig.formulaIndicators,
        allRows,
        globalPeriod,
      );
      allRows.push(...formulaRows);
      this.stats.aggregatedRows += formulaRows.length;
      formulaRows.forEach(r => {
        if (!this.stats.indicatorsProcessed.includes(r.Indicator)) {
          this.stats.indicatorsProcessed.push(r.Indicator);
        }
      });
    }

    this.stats.durationMs = Date.now() - startTime;
    this.logger.info(`\nAggregationEngine: processing complete in ${this.stats.durationMs}ms`);
    this.logger.info(`  Total rows read:        ${this.stats.totalRows}`);
    this.logger.info(`  Rows after filtering:   ${this.stats.filteredRows}`);
    this.logger.info(`  Dashboard rows output:  ${this.stats.aggregatedRows}`);
    this.logger.info(`  Indicators processed:   ${this.stats.indicatorsProcessed.join(', ')}`);

    return allRows;
  }

  // ─── Private Helpers ─────────────────────────────────────────────────────

  /**
   * Applies computedColumns definitions to a record, adding derived fields.
   * Currently supports formula: 'dateAdd' — adds a numeric column (months/days/etc.)
   * to a date column and stores the result as a JS Date on the record.
   */
  private applyComputedColumns(
    record: Record<string, unknown>,
    defs: ComputedColumnDef[],
  ): void {
    for (const def of defs) {
      if (def.formula !== 'dateAdd') continue;

      const baseDate = this.dateHelper.parse(record[def.sourceColumn]);
      if (!baseDate) continue;

      const addRaw = record[def.addColumn];
      const addNum = typeof addRaw === 'number' ? addRaw : parseFloat(String(addRaw ?? ''));
      if (isNaN(addNum) || addNum <= 0) continue;

      record[def.name] = baseDate.add(addNum, def.unit).toDate();
    }
  }

  private resolveAccValue(acc: StreamAcc, ind: IndicatorDef): number {
    switch (ind.aggregation) {
      case 'COUNT':         return acc.count;
      case 'SUM':           return acc.sum;
      case 'AVG':           return acc.count > 0 ? acc.sum / acc.count : 0;
      case 'MIN':           return acc.min === Infinity ? 0 : acc.min;
      case 'MAX':           return acc.max === -Infinity ? 0 : acc.max;
      case 'COUNTDISTINCT': return acc.distinctSet.size;
      default:              return acc.count;
    }
  }

  private writeOrgUnitsRawCsv(
    seen: Map<string, { stateRaw: string; lgaRaw: string; facilityRaw: string }>,
  ): void {
    if (seen.size === 0) return;
    const rawCsvPath = path.join(this.options.configDir, 'orgUnits-raw.csv');
    const lines = [
      'DATIMCode,State_raw,LGA_raw,Facility_raw,StateName,LGAName,FacilityName',
      ...[...seen.entries()]
        .sort((a, b) => a[0].localeCompare(b[0]))
        .map(([datim, ou]) =>
          `"${datim}","${ou.stateRaw}","${ou.lgaRaw}","${ou.facilityRaw}","","",""`,
        ),
    ];
    fs.writeFileSync(rawCsvPath, lines.join('\n'), 'utf8');
    this.logger.info(
      `\n  Org-unit raw values written to config/orgUnits-raw.csv (${seen.size} facilities)`,
    );
  }

  private addValidationIssue(
    sheet: string,
    row: number | string,
    column: string,
    severity: ValidationIssue['severity'],
    message: string,
  ): void {
    this.validationIssues.push({ sheet, row, column, severity, message });
    if (severity === 'ERROR') this.stats.warnings.push(message);
  }

  // ─── DATIM discovery cache ────────────────────────────────────────────────
  // ExcelJS streaming resolves shared strings inconsistently across runs:
  // compound facility UIDs (e.g. "W387n0fbylW_...") sometimes resolve to plain
  // patient codes/UUIDs, hiding the DATIM prefix.  We persist every DATIM we DO
  // discover so that coverage grows with each run.

  private loadDatimCache(): void {
    try {
      const raw = fs.readFileSync(this.datimCachePath, 'utf8');
      const cache = JSON.parse(raw) as Record<string, { Facility: string; State: string }>;
      let loaded = 0;
      for (const [datim, info] of Object.entries(cache)) {
        if (!this.datimToFacility.has(datim)) {
          this.datimToFacility.set(datim, info);
          loaded++;
        }
      }
      if (loaded > 0) {
        this.logger.info(
          `AggregationEngine: loaded ${loaded} DATIM entries from cache (${Object.keys(cache).length} total in cache)`,
        );
      }
    } catch {
      // Cache not yet created — first run
    }
  }

  private saveDatimCache(): void {
    let existing: Record<string, { Facility: string; State: string }> = {};
    try {
      existing = JSON.parse(fs.readFileSync(this.datimCachePath, 'utf8'));
    } catch { /* new file */ }

    let added = 0;
    for (const [datim, info] of this.datimToFacility.entries()) {
      if (info.Facility && !existing[datim]) {
        existing[datim] = info;
        added++;
      }
    }

    fs.writeFileSync(this.datimCachePath, JSON.stringify(existing, null, 2), 'utf8');
    if (added > 0) {
      this.logger.info(
        `AggregationEngine: saved ${added} new DATIM entries to cache (${Object.keys(existing).length} total)`,
      );
    }
  }
}
